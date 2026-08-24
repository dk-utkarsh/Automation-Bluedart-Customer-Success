"""Fail-safe proof for the watermark state machine.  Run: python tests/test_failsafe.py

No ticket may ever be skipped because a run died, and no mail may ever be sent twice
because a run died. Four scenarios drive the REAL start_pending / commit_run and the
real resume logic in main(); everything external (Desk, ClickPost, SMTP) is stubbed,
so this needs no network and touches none of the project's own state files.

  1. mail fails       -> watermark must NOT move; the export is kept for the retry
  2. retry            -> reuses the SAME export instead of re-fetching, then commits
  3. crash AFTER SMTP accepted, before the watermark committed
  4. next run         -> must NOT re-send, and must close out the pending watermark

Scenario 3/4 covers a window that was open until 2026-08-24: `mailer_sent` was
initialised to False and never set, so a crash in that window made the next run mail
the courier a second copy of the same escalation.
"""
import contextlib
import importlib.util
import io
import json
import pathlib
import shutil
import sys
import tempfile
from datetime import datetime, timezone

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.argv = ["main.py"]
spec = importlib.util.spec_from_file_location("m", ROOT / "main.py")
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

TMP = pathlib.Path(tempfile.mkdtemp(prefix="failsafe-"))
m.STATE_FILE = TMP / "state.json"
m.OUT = TMP / "output"
m.OUT.mkdir()
export = m.OUT / "Zoho_Desk_Logistics_TEST.xlsx"
for name in (export.name, "cp.csv", "merged.xlsx", "Mapping.xlsx"):
    (m.OUT / name).write_bytes(b"placeholder")

m.ENV = {}          # nothing here reaches code that reads .env
m.save_state({"last_ticket_number": "100",
              "last_created_utc": "2026-08-24T00:00:00+00:00"})

fetch_calls, send_calls = [], []
next_watermark = {"ticketNumber": "200",
                  "created_utc": datetime(2026, 8, 24, 11, 0, tzinfo=timezone.utc)}


def fake_desk_fetch(state, force_today=False, force_now=False):
    if next_watermark is None:
        return None
    fetch_calls.append(1)
    m.start_pending(state, export, next_watermark)
    return {"path": export, "watermark": next_watermark}


def ok_send(path, rows):
    send_calls.append(1)
    return "<ok@localhost>"


def failed_send(path, rows):
    send_calls.append(1)
    return None


m.run_desk_fetch = fake_desk_fetch
m.awb_count = lambda p: 3
m.clickpost_session = lambda: object()
m.run_bulk_report = lambda d, t: True
m.run_download = lambda d: m.OUT / "cp.csv"
m.run_merge = lambda t, c: m.OUT / "merged.xlsx"
m.run_mapping = lambda s, asof=None: (m.OUT / "Mapping.xlsx", 5)
m.thread_ticket_map = lambda p: {}
m.record_thread = lambda *a, **k: None
m.cleanup_run_files = lambda *a, **k: None


def run():
    """Call main() with its console output swallowed; return (rc, exception)."""
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            return m.main(), None
    except Exception as e:                          # a crash mid-run is a scenario
        return None, e


def show(label):
    st = json.loads(m.STATE_FILE.read_text(encoding="utf-8"))
    p = st.get("pending")
    print("  {:26} watermark=#{:<5} pending={}".format(
        label, st.get("last_ticket_number"),
        "yes (mailer_sent={})".format(p.get("mailer_sent")) if p else "no"))
    return st


failed = 0
try:
    print("1. mail fails - the watermark must not move")
    m.send_mapping_mail = failed_send
    rc, exc = run()
    st = show("after failed run:")
    assert rc == 1, "expected exit 1, got {}".format(rc)
    assert st["last_ticket_number"] == "100", "WATERMARK MOVED after a failed mail"
    assert st["pending"]["watermark_ticket"] == "200"
    assert st["pending"]["mailer_sent"] is False

    print("\n2. retry - must reuse the downloaded export")
    before = len(fetch_calls)
    m.send_mapping_mail = ok_send
    rc, exc = run()
    st = show("after retry:")
    assert rc == 0, "retry should succeed, got {}".format(rc)
    assert len(fetch_calls) == before, "re-fetched from Desk instead of resuming"
    assert st["last_ticket_number"] == "200" and not st.get("pending")

    print("\n3. crash after SMTP accepted, before the watermark committed")
    next_watermark = {"ticketNumber": "300",
                      "created_utc": datetime(2026, 8, 24, 12, 0, tzinfo=timezone.utc)}

    def boom(*a, **k):
        raise RuntimeError("simulated crash straight after SMTP accepted")

    m.record_thread = boom
    rc, exc = run()
    st = show("after the crash:")
    assert exc is not None, "the simulated crash did not happen"
    assert st["last_ticket_number"] == "200", "watermark moved despite the crash"
    assert st["pending"]["mailer_sent"] is True, \
        "send NOT recorded - the next run would mail a duplicate"

    print("\n4. next run - must not re-send, must commit the pending watermark")
    m.record_thread = lambda *a, **k: None
    next_watermark = None                      # nothing new today
    send_calls.clear()
    rc, exc = run()
    st = show("after recovery run:")
    assert exc is None, "recovery run raised: {}".format(exc)
    assert not send_calls, "IT RE-SENT THE MAIL - duplicate escalation"
    assert st["last_ticket_number"] == "300", "pending watermark was not committed"
    assert not st.get("pending"), "pending block should be cleared"
except AssertionError as e:
    failed = 1
    print("\nFAILED: {}".format(e))
finally:
    shutil.rmtree(TMP, ignore_errors=True)

print("\n{}".format("ALL FAIL-SAFE ASSERTIONS PASSED" if not failed else "FAILED"))
sys.exit(failed)
