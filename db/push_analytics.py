#!/usr/bin/env python3
"""Push ticket journeys into Zoho Analytics as they change.

  python3 db/push_analytics.py --dry-run   # show what would go, send nothing
  python3 db/push_analytics.py             # push everything that has moved
  python3 db/push_analytics.py --all       # push every ticket, moved or not

Analytics is an APPEND-ONLY SNAPSHOT LOG, by request: no key, no unique id, no
upsert, no matching. A ticket gets a fresh row each time its journey moves and
nothing already there is ever modified or replaced. The newest row for a ticket
number is its current state.

That only works if this end knows what it already sent, which is what
analytics_pushes records. ticket_events is the change signal: it is append-only,
so a ticket has moved exactly when it carries an event newer than the one last
pushed. Re-sending an unchanged ticket would add a duplicate row that nothing
would ever clean up.

Rows are read from the ticket_journey view - the same source db/export_xlsx.py
uses - so there is one definition of a journey and this cannot drift from it.

Config comes from .env (BLUEDART_ENV overrides the path):
  DATABASE_URL     the Postgres to read from
  ZA_CLIENT_ID / ZA_CLIENT_SECRET / ZA_REFRESH_TOKEN
  ZA_ORG_ID        sent as the ZANALYTICS-ORGID header
  ZA_WORKSPACE_ID / ZA_VIEW_ID    from the Analytics URL:
      analytics.zoho.in/workspace/<ZA_WORKSPACE_ID>/view/<ZA_VIEW_ID>
  ZA_API_DOMAIN    https://analyticsapi.zoho.in   (.com for the US DC)
  ZA_ACCOUNTS_URL  https://accounts.zoho.in       (must match that DC)

The OAuth client needs ZohoAnalytics.fullaccess.all. A Desk-scoped token
cannot import, whatever else it can do.
"""
import csv
import io
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT_JSON = ROOT / "output" / "ticket_journey_upload.json"
TOKEN_CACHE = ROOT / "output" / ".za_token.json"

VIEW_NAME = "ticket_journey"

# Sent to Analytics and used to render the values, so the two cannot disagree.
DATE_FORMAT_ZOHO = "dd-MM-yyyy HH:mm:ss"

# Postgres renders timestamps as text through --csv; this spots them so they can
# be restated. Anchored, so a status remark is never mangled by accident.
TS_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})[ T]"
                   r"(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?"
                   r"(?:[+-]\d{2}(?::?\d{2})?)?$")

NUMERIC = ("integer", "bigint", "smallint", "numeric", "double", "real")

# Tickets whose journey has moved since they were last pushed, plus the event id
# that proves it - recorded only after Analytics has accepted the row, so a
# failed push is retried rather than lost.
PENDING_SQL = """
SELECT j.*, e.max_event_id
  FROM {view} j
  JOIN (SELECT ticket_number, max(id) AS max_event_id
          FROM ticket_events GROUP BY ticket_number) e
    ON e.ticket_number = j.ticket_number
  LEFT JOIN analytics_pushes p ON p.ticket_number = j.ticket_number
 WHERE e.max_event_id > coalesce(p.last_event_id, 0)
 ORDER BY j.ticket_number
""".format(view=VIEW_NAME)

ALL_SQL = """
SELECT j.*, coalesce(e.max_event_id, 0) AS max_event_id
  FROM {view} j
  LEFT JOIN (SELECT ticket_number, max(id) AS max_event_id
               FROM ticket_events GROUP BY ticket_number) e
    ON e.ticket_number = j.ticket_number
 ORDER BY j.ticket_number
""".format(view=VIEW_NAME)


def import_config():
    """The CONFIG Analytics is sent.

    append, with no matchingColumns: the table is a snapshot log and must never
    be updated in place. JSON rather than CSV - the CSV path needs two CONFIG
    keys the published API table does not mention (delimiter, quoted) and then
    still fails UNABLE_TO_PARSE_DATA_TYPE (8516) for even a single numeric
    column, while JSON is accepted as-is."""
    return {
        "importType": "append",
        "fileType": "json",
        "autoIdentify": "true",
        # One unparseable value empties that cell instead of losing the row.
        "onError": "setcolumnempty",
        "dateFormat": DATE_FORMAT_ZOHO,
    }


# ----------------------------------------------------------------- config
_ENV = None


def env():
    """Parsed .env, read once. Lazy so importing this module never fails."""
    global _ENV
    if _ENV is None:
        path = Path(os.environ.get("BLUEDART_ENV") or ROOT / ".env")
        if not path.exists():
            raise SystemExit("No .env at {}".format(path))
        _ENV = {}
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                _ENV[k.strip()] = v.strip()
    return _ENV


def need(key):
    v = (env().get(key) or "").strip()
    if not v:
        raise SystemExit("{} is not set in .env".format(key))
    return v


# --------------------------------------------------------------- Postgres
def dsn():
    m = re.match(r"postgresql://([^:]+):([^@]+)@([^:]+):(\d+)/(.+)",
                 need("DATABASE_URL"))
    if not m:
        raise SystemExit("DATABASE_URL is not postgresql://user:pw@host:port/db")
    return m.groups()


def psql(sql, quiet=False):
    user, pw, host, port, db = dsn()
    r = subprocess.run(
        ["psql", "-h", host, "-p", port, "-U", user, "-d", db, "--csv",
         "-c", sql],
        capture_output=True, text=True,
        # PGTZ, not a SET statement: psql prints "SET" for that too and --csv
        # reads it as the header row. Everything here happened in IST.
        env={**os.environ, "PGPASSWORD": pw, "PGTZ": "Asia/Kolkata"})
    if r.returncode:
        if quiet:
            return [], []
        raise SystemExit("psql failed: " + r.stderr.strip())
    rows = list(csv.reader(io.StringIO(r.stdout)))
    return (rows[0], rows[1:]) if rows else ([], [])


def column_types():
    _, rows = psql("SELECT column_name, data_type FROM information_schema.columns"
                   " WHERE table_name='{}'".format(VIEW_NAME))
    return {c: t for c, t in rows}


def restate(v):
    """ISO timestamp -> the format CONFIG declares. Anything else untouched."""
    m = TS_RE.match(v or "")
    if not m:
        return v
    y, mo, d, hh, mi, ss = m.groups()
    return "{}-{}-{} {}:{}:{}".format(d, mo, y, hh, mi, ss)


def build_rows(header, rows, types):
    """(records, {ticket_number: max_event_id}) for the payload.

    max_event_id rides along for bookkeeping and is stripped from what is sent -
    Analytics has no such column and would reject it.

    Numbers are sent as numbers using the column's DECLARED type. Guessing from
    the digits would turn ticket_number and awb into numbers and lose any
    leading zero."""
    records, marks = [], {}
    for row in rows:
        rec = {}
        for col, v in zip(header, row):
            if col == "max_event_id":
                continue
            if v == "":
                rec[col] = None          # an absent date, not an empty string
            elif str(types.get(col, "")).startswith(NUMERIC):
                try:
                    rec[col] = int(v) if "." not in v else float(v)
                except ValueError:
                    rec[col] = v
            else:
                rec[col] = restate(v)
        records.append(rec)
        pair = dict(zip(header, row))
        marks[pair["ticket_number"]] = int(pair.get("max_event_id") or 0)
    return records, marks


def remember(marks):
    """Record what was pushed. Only ever called after Analytics accepted it."""
    if not marks:
        return
    values = ", ".join(
        "('{}', {})".format(str(t).replace("'", "''"), int(e))
        for t, e in marks.items())
    psql(
        "INSERT INTO analytics_pushes (ticket_number, last_event_id) "
        "VALUES {} ON CONFLICT (ticket_number) DO UPDATE SET "
        "  last_event_id = EXCLUDED.last_event_id, "
        "  push_count    = analytics_pushes.push_count + 1, "
        "  pushed_at     = now()".format(values))


# ------------------------------------------------------------------- Zoho
def token():
    """A valid access token, refreshed only when the cached one is stale.

    Zoho rate-limits the refresh endpoint hard - refresh on every call and it
    answers "You have made too many requests continuously", which then blocks
    the real work until it cools down. Tokens last an hour.

    The cache sits in output/, which is gitignored: it holds a live credential."""
    try:
        c = json.loads(TOKEN_CACHE.read_text(encoding="utf-8"))
        if c.get("expires_at", 0) > time.time() + 60:
            return c["access_token"]
    except (OSError, ValueError, KeyError):
        pass
    tok = _refresh()
    try:
        TOKEN_CACHE.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_CACHE.write_text(json.dumps(
            {"access_token": tok, "expires_at": time.time() + 3500}),
            encoding="utf-8")
        os.chmod(TOKEN_CACHE, 0o600)
    except OSError:
        pass                      # a cache that cannot be written is not fatal
    return tok


def _refresh():
    body = urllib.parse.urlencode({
        "refresh_token": need("ZA_REFRESH_TOKEN"),
        "client_id": need("ZA_CLIENT_ID"),
        "client_secret": need("ZA_CLIENT_SECRET"),
        "grant_type": "refresh_token",
    }).encode()
    url = (env().get("ZA_ACCOUNTS_URL") or "https://accounts.zoho.in") \
        + "/oauth/v2/token"
    with urllib.request.urlopen(urllib.request.Request(url, data=body),
                                timeout=30) as r:
        d = json.load(r)
    if "access_token" not in d:
        # The token is never printed; the reason for the failure is.
        raise SystemExit("Zoho token refresh failed: {}".format(
            {k: v for k, v in d.items() if "token" not in k}))
    return d["access_token"]


def multipart(field, filename, content, ctype="application/json"):
    """One-file multipart/form-data body. urllib has no equivalent."""
    boundary = "----bluedart" + uuid.uuid4().hex
    body = (
        "--{b}\r\n"
        'Content-Disposition: form-data; name="{f}"; filename="{n}"\r\n'
        "Content-Type: {c}\r\n\r\n"
    ).format(b=boundary, f=field, n=filename, c=ctype).encode() \
        + content.encode("utf-8") \
        + "\r\n--{}--\r\n".format(boundary).encode()
    return body, "multipart/form-data; boundary=" + boundary


def push(payload):
    url = "{}/restapi/v2/workspaces/{}/views/{}/data?{}".format(
        (env().get("ZA_API_DOMAIN") or "https://analyticsapi.zoho.in").rstrip("/"),
        need("ZA_WORKSPACE_ID"), need("ZA_VIEW_ID"),
        urllib.parse.urlencode({"CONFIG": json.dumps(import_config())}))
    body, ctype = multipart("FILE", "ticket_journey.json", payload)
    req = urllib.request.Request(url, data=body, headers={
        "Authorization": "Zoho-oauthtoken " + token(),
        "ZANALYTICS-ORGID": need("ZA_ORG_ID"),
        "Content-Type": ctype,
    })
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            return r.status, json.load(r)
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            return e.code, json.loads(raw)
        except ValueError:
            return e.code, raw[:800]


def run(dry=False, everything=False, verbose=True):
    """Push whatever has moved. Returns the number of rows sent."""
    header, rows = psql(ALL_SQL if everything else PENDING_SQL)
    if not rows:
        if verbose:
            print("Analytics: nothing new to push.")
        return 0
    records, marks = build_rows(header, rows, column_types())
    payload = json.dumps(records, indent=2)
    try:
        OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
        OUT_JSON.write_text(payload, encoding="utf-8")
    except OSError:
        pass

    if dry:
        print("Analytics: {} row(s) would be appended (--dry-run, nothing "
              "sent).".format(len(records)))
        print("  " + json.dumps(records[0], indent=2).replace("\n", "\n  "))
        return 0

    code, body = push(payload)
    ok = code == 200 and isinstance(body, dict) and body.get("status") == "success"
    if not ok:
        print("Analytics push FAILED (HTTP {}): {}".format(
            code, json.dumps(body)[:400] if isinstance(body, dict) else body))
        print("  Nothing recorded as pushed, so the next run retries these.")
        return 0
    remember(marks)
    s = (body.get("data") or {}).get("importSummary") or {}
    if verbose:
        print("Analytics: appended {} of {} row(s) for ticket(s) {}".format(
            s.get("successRowCount"), s.get("totalRowCount"),
            ", ".join(sorted(marks)[:8]) + ("..." if len(marks) > 8 else "")))
    errs = (body.get("data") or {}).get("importErrors")
    if errs and verbose:
        print("  warnings: {}".format(errs))
    return int(s.get("successRowCount") or 0)


def main():
    return 0 if run(dry="--dry-run" in sys.argv,
                    everything="--all" in sys.argv) >= 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
