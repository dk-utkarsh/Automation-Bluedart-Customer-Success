"""Zoho Mail outgoing-webhook receiver: a doorbell for process_replies().

  python webhook_server.py            # listen (config comes from .env)
  python webhook_server.py --selftest # signature + debounce checks, no socket

WHY A DOORBELL AND NOT A MAILMAN
--------------------------------
Zoho's Mail webhook payload carries `summary, subject, messageId, sender, html,
fromAddress, receivedTime, ...` and NOT the RFC headers - there is no Message-ID,
no In-Reply-To, no References. Those three are exactly what find_thread_replies()
uses to bind a reply to the report it answers; without them every reply falls to
the ambiguous subject path, which is what caused one reply to be commented onto
three different reports on 27-Aug.

So this server deliberately ignores the payload body. A POST means only "something
arrived, go look" - the mailbox is still read over IMAP, where the real headers
are. Consequences, all of them wanted:

  * the undocumented payload schema cannot break matching, because it is not parsed
  * a POST that Zoho never sends, or that arrives while we are restarting, costs
    nothing: --watch and the daily run re-read the same mailbox anyway
  * duplicate POSTs are harmless - they coalesce into one scan

WHAT IT DOES NOT DO
-------------------
run_followups() is NOT called. Follow-ups send mail to MAIL_TO, and a mail-triggered
loop that can itself send mail is a foot-gun; leave those to --followups on a timer.
"""

import hashlib
import hmac
import json
import os
import pathlib
import sys
import threading
import time
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BASE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

import main as M                                              # noqa: E402

# --- tunables, all overridable from .env so nothing is hard-coded here --------
DEFAULT_BIND = "127.0.0.1"      # nginx terminates TLS and proxies to us
DEFAULT_PORT = 8787
DEFAULT_PATH = "/zoho-mail-hook"
DEFAULT_DEBOUNCE = 5.0          # seconds to coalesce a burst of POSTs into one scan
MAX_BODY = 1 << 20              # 1 MiB; Zoho payloads are a few KB

SECRET_FILE = BASE / ".webhook-secret"   # x-hook-secret is sent ONCE, so persist it

# Cross-process exclusion lives in main.py (threads_lock), not here: --watch and the
# daily run touch the same threads.json, so the lock has to be held by whoever mutates
# the file, not by this listener. process_replies() takes it for us.


def log(msg):
    print("[{:%Y-%m-%d %H:%M:%S}] {}".format(datetime.now(M.IST), msg), flush=True)


def cfg(key, default):
    v = (M.ENV.get(key) or "").strip()
    return v if v and v != M.PLACEHOLDER else default


# ---------------------------------------------------------------------------
# Signature
# ---------------------------------------------------------------------------

def load_secret():
    """The shared secret, from .env or from the one-shot handshake file.

    Zoho delivers x-hook-secret in the header of the FIRST request only. If it is
    missed the webhook has to be deleted and recreated, so it is written to disk the
    moment it is seen. WEBHOOK_SECRET in .env takes precedence when set."""
    v = cfg("WEBHOOK_SECRET", "")
    if v:
        return v
    if SECRET_FILE.exists():
        return SECRET_FILE.read_text(encoding="utf-8").strip()
    return ""


def remember_secret(secret):
    try:
        SECRET_FILE.write_text(secret, encoding="utf-8")
        try:
            os.chmod(SECRET_FILE, 0o600)
        except OSError:
            pass                                    # best effort; Windows has no mode
        log("stored x-hook-secret from the handshake request")
    except OSError as e:
        log("WARNING could not persist x-hook-secret: {}".format(e))


def signature_ok(secret, body, header):
    """HMAC-SHA256 of the raw body, compared in constant time.

    Accepts hex or base64 because Zoho's encoding is not documented for this hook,
    and an unrecognised-but-correct digest must not be read as a forgery."""
    if not secret or not header:
        return False
    mac = hmac.new(secret.encode("utf-8"), body, hashlib.sha256)
    want_hex = mac.hexdigest()
    import base64
    want_b64 = base64.b64encode(mac.digest()).decode()
    got = header.strip()
    for candidate in (want_hex, want_b64):
        if hmac.compare_digest(got, candidate):
            return True
    return False


# ---------------------------------------------------------------------------
# The scan worker
# ---------------------------------------------------------------------------

class Scanner:
    """Coalesces POSTs into at most one process_replies() at a time.

    A courier answering several reports at once produces a burst of POSTs; running a
    scan per POST would open several IMAP sessions and race on threads.json, which is
    read-modify-written whole. One trailing scan sees every message anyway."""

    def __init__(self, debounce=DEFAULT_DEBOUNCE):
        self.debounce = debounce
        self.wake = threading.Event()
        self.lock = threading.Lock()
        self.pending = 0
        self.runs = 0
        self.last_run = None
        self.last_error = None

    def poke(self):
        with self.lock:
            self.pending += 1
        self.wake.set()

    def run_once(self):
        """One scan. process_replies() takes main.py's cross-process threads_lock, so
        a --watch or daily run happening at the same moment waits rather than racing -
        without that, both would load threads.json, both would see the same reply as
        unhandled, and the ticket would get two identical comments."""
        try:
            made = M.process_replies(verbose=True)
            self.last_run = datetime.now(M.IST).isoformat()
            self.runs += 1
            log("scan complete: {} comment(s) added".format(made))
        except Exception as e:                       # never let the worker die
            self.last_error = "{}: {}".format(e.__class__.__name__, e)
            log("scan FAILED {}".format(self.last_error))

    def loop(self):
        log("scanner ready (debounce {}s)".format(self.debounce))
        while True:
            self.wake.wait()
            time.sleep(self.debounce)                # let a burst accumulate
            self.wake.clear()
            with self.lock:
                n, self.pending = self.pending, 0
            log("scanning (coalesced {} event(s))".format(n))
            self.run_once()


SCANNER = Scanner()


# ---------------------------------------------------------------------------
# HTTP
# ---------------------------------------------------------------------------

class Handler(BaseHTTPRequestHandler):
    server_version = "ZohoDeskHook/1.0"

    def log_message(self, fmt, *args):
        pass                                          # we do our own logging

    def _reply(self, code, body=b"", ctype="text/plain"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if body:
            self.wfile.write(body)

    def do_GET(self):
        # Health check for nginx / uptime monitoring. Deliberately says nothing
        # about the mailbox, the tickets, or the config.
        if self.path.split("?")[0] in ("/healthz", "/"):
            payload = json.dumps({
                "ok": True,
                "scans": SCANNER.runs,
                "last_run_ist": SCANNER.last_run,
                "last_error": SCANNER.last_error,
            }).encode()
            return self._reply(200, payload, "application/json")
        return self._reply(404)

    def do_POST(self):
        path = self.path.split("?")[0]
        if path != cfg("WEBHOOK_PATH", DEFAULT_PATH):
            return self._reply(404)

        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            return self._reply(400, b"bad length")
        if length > MAX_BODY:
            return self._reply(413, b"too large")
        body = self.rfile.read(length) if length else b""

        # Zoho's one-time handshake: the secret arrives on the first request only.
        handshake = self.headers.get("x-hook-secret")
        if handshake:
            remember_secret(handshake.strip())
            # Echo it back - that is how the subscription is confirmed.
            self.send_response(200)
            self.send_header("x-hook-secret", handshake.strip())
            self.send_header("Content-Length", "0")
            self.end_headers()
            log("handshake accepted from {}".format(self.client_address[0]))
            return

        secret = load_secret()
        sig = self.headers.get("x-hook-signature") or self.headers.get("X-Hook-Signature")
        if secret:
            if not signature_ok(secret, body, sig):
                log("REJECTED unsigned/bad-signature POST from {}".format(
                    self.client_address[0]))
                return self._reply(401, b"bad signature")
        else:
            # Fail loudly rather than silently trusting the internet. An unsigned
            # endpoint lets anyone trigger scans; that is only tolerable while the
            # secret has genuinely not been issued yet.
            log("WARNING accepting UNSIGNED POST - no secret known yet. "
                "Set WEBHOOK_SECRET in .env or recreate the webhook to get one.")

        SCANNER.poke()
        # 200 immediately: Zoho auto-disables an endpoint that is slow or errors,
        # and the scan takes seconds (IMAP + Desk API). Never do work inline.
        self._reply(200, b"queued")


def serve():
    M.ENV = M.load_env()
    M.require_env(M.ENV, M.MAIL_ENV,
                  "The webhook only wakes the IMAP reader; those credentials are "
                  "still what actually fetches the reply.")
    bind = cfg("WEBHOOK_BIND", DEFAULT_BIND)
    port = int(cfg("WEBHOOK_PORT", str(DEFAULT_PORT)))
    path = cfg("WEBHOOK_PATH", DEFAULT_PATH)
    SCANNER.debounce = float(cfg("WEBHOOK_DEBOUNCE", str(DEFAULT_DEBOUNCE)))

    threading.Thread(target=SCANNER.loop, daemon=True).start()

    log("listening on http://{}:{}{}".format(bind, port, path))
    log("mailbox    : {}".format(M.ENV.get("OTP_EMAIL")))
    log("signature  : {}".format("enforced" if load_secret() else
                                 "NOT YET - awaiting Zoho handshake"))
    # A scan on boot, so anything that landed while we were down is picked up. This
    # is the property a webhook alone does not have.
    SCANNER.poke()
    ThreadingHTTPServer((bind, port), Handler).serve_forever()


# ---------------------------------------------------------------------------

def run_selftest():
    failed = 0

    def check(desc, got, want):
        nonlocal failed
        ok = got == want
        failed += not ok
        print("  [{}] {}: got {!r}, expected {!r}".format(
            "PASS" if ok else "FAIL", desc, got, want))

    body = b'{"subject":"Re: Customer Escalation"}'
    secret = "s3cr3t"
    good_hex = hmac.new(secret.encode(), body, hashlib.sha256).hexdigest()
    import base64
    good_b64 = base64.b64encode(
        hmac.new(secret.encode(), body, hashlib.sha256).digest()).decode()

    check("valid hex signature", signature_ok(secret, body, good_hex), True)
    check("valid base64 signature", signature_ok(secret, body, good_b64), True)
    check("wrong signature", signature_ok(secret, body, "deadbeef"), False)
    check("tampered body", signature_ok(secret, body + b"x", good_hex), False)
    check("no secret", signature_ok("", body, good_hex), False)
    check("no header", signature_ok(secret, body, None), False)

    s = Scanner(debounce=0.05)
    calls = []
    s.run_once = lambda: calls.append(1)
    threading.Thread(target=s.loop, daemon=True).start()
    for _ in range(5):
        s.poke()
    time.sleep(0.6)
    check("5 rapid POSTs coalesce into 1 scan", len(calls), 1)
    s.poke()
    time.sleep(0.4)
    check("a later POST triggers another scan", len(calls), 2)

    print("\n{}".format("ALL PASSED" if not failed else "{} FAILED".format(failed)))
    return failed


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        raise SystemExit(1 if run_selftest() else 0)
    try:
        serve()
    except KeyboardInterrupt:
        log("stopped")
