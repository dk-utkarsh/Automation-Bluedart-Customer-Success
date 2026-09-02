#!/usr/bin/env python3
"""Push the ticket_journey view into a Zoho Analytics table.

  python3 db/push_analytics.py --dry-run    # write the CSV, send nothing
  python3 db/push_analytics.py              # write it and upload

Reads the view through psql, exactly as db/export_xlsx.py does, so there is one
definition of what a ticket's journey is and this script cannot drift from it.

The import is `updateadd` matched on ticket_number, so re-running is safe: a
ticket already in Analytics is UPDATED rather than duplicated. That matters
because a journey is not finished when it is first pushed - delivered_at fills
in days later, and the row has to change rather than a second one appear.

Nothing here creates or alters the target table. Its columns already match the
view, and autoIdentify=false keeps the types the table was built with.

Config comes from .env (BLUEDART_ENV overrides the path):
  DATABASE_URL     the Postgres to read from
  ZA_CLIENT_ID / ZA_CLIENT_SECRET / ZA_REFRESH_TOKEN
  ZA_ORG_ID        sent as the ZANALYTICS-ORGID header
  ZA_WORKSPACE_ID / ZA_VIEW_ID    from the Analytics URL:
      analytics.zoho.in/workspace/<ZA_WORKSPACE_ID>/view/<ZA_VIEW_ID>
  ZA_API_DOMAIN    https://analyticsapi.zoho.in   (.com for the US DC)
  ZA_ACCOUNTS_URL  https://accounts.zoho.in       (must match the DC above)
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
ENV_FILE = Path(os.environ.get("BLUEDART_ENV") or ROOT / ".env")
DRY = "--dry-run" in sys.argv
OUT_JSON = ROOT / "output" / "ticket_journey_upload.json"

# The view is the contract. Named here rather than SELECT *ed blindly so a
# rename in the database is a loud failure instead of a silent column shift.
VIEW_NAME = "ticket_journey"
MATCH_ON = "ticket_number"

# What the CSV says its dates look like, and what CONFIG tells Analytics to
# expect. One constant, so the two can never disagree.
DATE_FORMAT_CSV = "%d-%m-%Y %H:%M:%S"
DATE_FORMAT_ZOHO = "dd-MM-yyyy HH:mm:ss"

# Postgres renders timestamps as text through --csv; this spots them so they can
# be restated in the format above. Anchored, so a status remark is never mangled.
TS_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})[ T]"
                   r"(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?"
                   r"(?:[+-]\d{2}(?::?\d{2})?)?$")


def env():
    if not ENV_FILE.exists():
        raise SystemExit("No .env at {}".format(ENV_FILE))
    out = {}
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            out[k.strip()] = v.strip()
    return out


ENV = env()


def need(key):
    v = (ENV.get(key) or "").strip()
    if not v:
        raise SystemExit("{} is not set in {}".format(key, ENV_FILE))
    return v


# ------------------------------------------------------------------ Postgres
def psql(dsn_parts, sql):
    user, pw, host, port, db = dsn_parts
    r = subprocess.run(
        ["psql", "-h", host, "-p", port, "-U", user, "-d", db, "--csv",
         "-c", sql],
        capture_output=True, text=True,
        # PGTZ, not a SET statement: psql prints "SET" for that and --csv reads
        # it as the header row. Everything in this database happened in IST.
        env={**os.environ, "PGPASSWORD": pw, "PGTZ": "Asia/Kolkata"})
    if r.returncode:
        raise SystemExit("psql failed: " + r.stderr.strip())
    rows = list(csv.reader(io.StringIO(r.stdout)))
    return (rows[0], rows[1:]) if rows else ([], [])


def read_view():
    """(header, rows, types) from the view, with timestamps already restated."""
    m = re.match(r"postgresql://([^:]+):([^@]+)@([^:]+):(\d+)/(.+)",
                 need("DATABASE_URL"))
    if not m:
        raise SystemExit("DATABASE_URL is not a postgresql://user:pw@host:port/db URL")
    dsn = m.groups()
    header, rows = psql(dsn, "SELECT * FROM {} ORDER BY {}".format(
        VIEW_NAME, MATCH_ON))
    if not header:
        raise SystemExit("{} returned nothing at all".format(VIEW_NAME))
    # Declared types, so a count is sent as a number and an identifier is not.
    # Guessing from the digits would turn ticket_number and awb into numbers and
    # lose any leading zero.
    _, trows = psql(dsn, "SELECT column_name, data_type FROM "
                         "information_schema.columns WHERE table_name='{}'"
                         .format(VIEW_NAME))
    return header, rows, {c: t for c, t in trows}


def restate(v):
    """ISO timestamp -> the format CONFIG declares. Anything else untouched."""
    m = TS_RE.match(v or "")
    if not m:
        return v
    y, mo, d, hh, mi, ss = m.groups()
    return "{}-{}-{} {}:{}:{}".format(d, mo, y, hh, mi, ss)


NUMERIC = ("integer", "bigint", "smallint", "numeric", "double", "real")


def build_rows(header, rows, types):
    """The payload Analytics receives: one JSON object per ticket.

    JSON rather than CSV, deliberately. The CSV path needs two undocumented
    CONFIG keys (delimiter, quoted) and still fails UNABLE_TO_PARSE_DATA_TYPE
    (8516) - even for a file of one numeric column - while the JSON path is
    accepted as-is. It also removes every quoting question from remarks that
    happen to contain a comma.

    Numbers are sent as numbers, using the column's DECLARED type. Guessing
    from the digits would turn ticket_number and awb into numbers and lose any
    leading zero."""
    out = []
    for row in rows:
        rec = {}
        for col, v in zip(header, row):
            if v == "":
                rec[col] = None            # an absent date, not an empty string
            elif str(types.get(col, "")).startswith(NUMERIC):
                try:
                    rec[col] = int(v) if "." not in v else float(v)
                except ValueError:
                    rec[col] = v
            else:
                rec[col] = restate(v)
        out.append(rec)
    return out


# ------------------------------------------------------------------- Zoho
TOKEN_CACHE = ROOT / "output" / ".za_token.json"


def token():
    """A valid access token, refreshed only when the cached one is stale.

    Zoho rate-limits the refresh endpoint hard - refresh on every call and it
    starts answering "You have made too many requests continuously", which then
    blocks the real work until it cools down. Tokens last an hour; this reuses
    one until a minute before it expires.

    The cache file lives in output/, which is gitignored - it holds a live
    credential and must never be committed."""
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
    url = (ENV.get("ZA_ACCOUNTS_URL") or "https://accounts.zoho.in") \
        + "/oauth/v2/token"
    with urllib.request.urlopen(urllib.request.Request(url, data=body),
                                timeout=30) as r:
        d = json.load(r)
    if "access_token" not in d:
        # The token itself is never printed; the failure reason is.
        raise SystemExit("Zoho token refresh failed: {}".format(
            {k: v for k, v in d.items() if "token" not in k}))
    return d["access_token"]


def multipart(field, filename, content):
    """One-file multipart/form-data body. urllib has no equivalent."""
    boundary = "----bluedart" + uuid.uuid4().hex
    body = (
        "--{b}\r\n"
        'Content-Disposition: form-data; name="{f}"; filename="{n}"\r\n'
        "Content-Type: text/csv\r\n\r\n"
    ).format(b=boundary, f=field, n=filename).encode() \
        + content.encode("utf-8") \
        + "\r\n--{}--\r\n".format(boundary).encode()
    return body, "multipart/form-data; boundary=" + boundary


def push(payload):
    config = {
        "importType": "updateadd",     # upsert: re-running updates, never dupes
        "fileType": "json",
        "autoIdentify": "true",
        "matchingColumns": [MATCH_ON],
        # An unparseable value empties that one cell rather than aborting the
        # whole file. A journey row is worth having with one field missing.
        "onError": "setcolumnempty",
        "dateFormat": DATE_FORMAT_ZOHO,
    }
    url = "{}/restapi/v2/workspaces/{}/views/{}/data?{}".format(
        (ENV.get("ZA_API_DOMAIN") or "https://analyticsapi.zoho.in").rstrip("/"),
        need("ZA_WORKSPACE_ID"), need("ZA_VIEW_ID"),
        urllib.parse.urlencode({"CONFIG": json.dumps(config)}))
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


def main():
    header, rows, types = read_view()
    print("Source : {} ({} row(s), {} column(s))".format(
        VIEW_NAME, len(rows), len(header)))
    records = build_rows(header, rows, types)
    payload = json.dumps(records, indent=2)
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(payload, encoding="utf-8")
    print("Payload: {}".format(OUT_JSON))

    if DRY:
        print("\n--dry-run: nothing sent. First record:\n")
        print("  " + json.dumps(records[0], indent=2).replace("\n", "\n  "))
        return 0

    print("Target : workspace {} / view {}".format(
        need("ZA_WORKSPACE_ID"), need("ZA_VIEW_ID")))
    print("Import : updateadd on {}\n".format(MATCH_ON))
    code, body = push(payload)
    if code == 200 and isinstance(body, dict) and             body.get("status") == "success":
        s = (body.get("data") or {}).get("importSummary") or {}
        print("Imported: {} of {} row(s), {} column(s), operation {}".format(
            s.get("successRowCount"), s.get("totalRowCount"),
            s.get("selectedColumnCount"), s.get("importOperation")))
        errs = (body.get("data") or {}).get("importErrors")
        if errs:
            print("Warnings: {}".format(errs))
        return 0
    print("FAILED (HTTP {}): {}".format(code, json.dumps(body)[:800]
                                        if isinstance(body, dict) else body))
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
