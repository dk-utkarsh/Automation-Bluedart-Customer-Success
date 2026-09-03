"""Late-edit sweep.  Run: python tests/test_late_edits.py

A ticket created BEFORE the watermark and rerouted to Logistics afterwards used to be
invisible forever: the creation-time watermark had already passed it, so no future run
looked at it again. run_desk_fetch now runs a second sweep by modifiedTime.

Four scenarios, all driving the real run_desk_fetch. Desk is stubbed, so no network.

  1. a late-rerouted ticket is picked up even though it predates the watermark
  2. a ticket already mailed is NOT picked up again when someone touches it
  3. late edits never drag the created watermark backwards
  4. a ticket older than the lookback bound is not dragged back in
  5. the modified watermark advances only as far as what was actually swept
"""
import contextlib
import importlib.util
import io
import json
import pathlib
import shutil
import sys
import tempfile
from datetime import datetime, timedelta, timezone

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.argv = ["main.py"]
spec = importlib.util.spec_from_file_location("m", ROOT / "main.py")
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

TMP = pathlib.Path(tempfile.mkdtemp(prefix="lateedit-"))
m.STATE_FILE = TMP / "state.json"

# Run state now lives in Postgres. These tests exercise the pending/watermark
# STATE MACHINE, not where it is stored, so they keep the old file-backed
# storage and leave the database backend to tests/test_state_in_db.py.
m.load_state = lambda: (json.loads(m.STATE_FILE.read_text(encoding="utf-8"))
                        if m.STATE_FILE.exists() else {})
m.save_state = lambda st: m.STATE_FILE.write_text(
    json.dumps(st, indent=2), encoding="utf-8")

m.OUT = TMP / "output"
m.OUT.mkdir()
m.ENV = {}
m.TOKEN = "stub"
m.require_env = lambda *a, **k: None
m.get_token = lambda: "stub"

NOW = datetime(2026, 8, 24, 12, 0, tzinfo=timezone.utc)


# The fixtures are pinned to NOW, so the clock main.py reads must be pinned too.
# It was not: resolve_window called the real datetime.now(), so the 14-day
# late-edit lookback floor drifted forward every day while the fixtures stayed
# put. On 03 Sep 2026 the floor passed the LATE ticket's creation date and this
# test began failing on a change to the calendar rather than a change to code.
class _Frozen(datetime):
    @classmethod
    def now(cls, tz=None):
        return NOW.astimezone(tz) if tz else NOW.replace(tzinfo=None)


m.datetime = _Frozen


def z(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S.000Z")


# #50 created 5 days ago - long before the watermark - but rerouted 10 minutes ago.
LATE = {"id": "id50", "ticketNumber": "50",
        "createdTime": z(NOW - timedelta(days=5)),
        "modifiedTime": z(NOW - timedelta(minutes=10))}
# #9 created 40 days ago and touched today - past the lookback, must be skipped.
ANCIENT = {"id": "id9", "ticketNumber": "9",
           "createdTime": z(NOW - timedelta(days=40)),
           "modifiedTime": z(NOW - timedelta(minutes=5))}
# #101 created inside today's window; its modifiedTime equals its createdTime.
NEW = {"id": "id101", "ticketNumber": "101",
       "createdTime": z(NOW - timedelta(hours=2)),
       "modifiedTime": z(NOW - timedelta(hours=2))}

captured = {}


def make_api_get(created_rows, modified_rows):
    def api_get(path):
        if "from=1&" not in path and "from=1" not in path:
            return {"data": []}                       # only ever one page here
        if "sortBy=-modifiedTime" in path:
            return {"data": modified_rows}
        return {"data": created_rows}
    return api_get


def fake_fetch_detail(t):
    return {"ticketId": t["id"], "ticketNumber": t["ticketNumber"],
            "created_utc": m.parse_zoho_time(t["createdTime"]),
            "created": "", "subject": "", "status": "Open", "statusType": "Open",
            "owner": "", "department": "", "priority": "",
            m.PENDING_FIELD: "Logistics Team",
            "Courier Partner": "Bluedart",
            "AWB Number": "AWB" + t["ticketNumber"],
            "Logistics Classification": "Out TAT Intransit",
            "Vinculum Shipment EDD": "", "States": ""}


def fake_write(rows, dropped, cutoff_ist, path, n_closed=0, n_courier=0):
    captured["kept"] = [r["ticketNumber"] for r in rows]
    path.write_bytes(b"placeholder")


m.fetch_detail = fake_fetch_detail
m.write_desk_excel = fake_write
m.save_registry = lambda reg: None


def seed(last_ticket, created_days_ago, modified=None):
    st = {"last_ticket_number": last_ticket,
          "last_created_utc": (NOW - timedelta(days=created_days_ago)).isoformat()}
    if modified:
        st["last_modified_utc"] = modified.isoformat()
    m.save_state(st)
    return m.load_state()


def run_fetch(state, threads=(), registry=None):
    m.load_threads = lambda: list(threads)
    m.load_registry = lambda: dict(registry or {})
    captured.clear()
    with contextlib.redirect_stdout(io.StringIO()) as buf:
        res = m.run_desk_fetch(state, force_today=False, force_now=False)
    return res, buf.getvalue()


failed = 0
try:
    print("1. a ticket rerouted to Logistics after the watermark is picked up")
    m.api_get = make_api_get([NEW], [LATE, NEW])
    st = seed("100", created_days_ago=1)
    res, out = run_fetch(st)
    print("   kept:", captured.get("kept"))
    assert res is not None, "run returned nothing"
    assert "50" in captured["kept"], "LATE EDIT MISSED - #50 was not picked up"
    assert "101" in captured["kept"], "normal ticket #101 missing"

    print("\n2. a ticket already mailed is not re-reported when touched")
    m.api_get = make_api_get([NEW], [LATE, NEW])
    st = seed("100", created_days_ago=1)
    already = [{"tickets": {"50": {"ticketId": "id50", "awb": "AWB50"}}}]
    res, out = run_fetch(st, threads=already)
    print("   kept:", captured.get("kept"))
    assert "50" not in captured["kept"], "re-reported a ticket already mailed"
    assert "101" in captured["kept"]

    print("\n3. late edits do not drag the created watermark backwards")
    m.api_get = make_api_get([], [LATE])          # ONLY a late edit, nothing new
    st = seed("100", created_days_ago=1)
    res, out = run_fetch(st)
    print("   kept:", captured.get("kept"), " watermark:", res["watermark"]["ticketNumber"])
    assert "50" in captured["kept"], "late edit missed when it was the only ticket"
    assert res["watermark"]["ticketNumber"] == "100", \
        "created watermark moved to #{} - it must stay at #100".format(
            res["watermark"]["ticketNumber"])

    print("\n4. a ticket older than the lookback is not dragged back in")
    m.api_get = make_api_get([NEW], [ANCIENT, LATE, NEW])
    st = seed("100", created_days_ago=1)
    res2, out2 = run_fetch(st)
    print("   kept:", captured.get("kept"))
    assert "9" not in captured["kept"], \
        "a 40-day-old ticket was pulled in despite the lookback bound"
    assert "50" in captured["kept"], "the in-range late edit was lost"
    assert "older than the" in out2, "the skip was not reported on the console"

    print("\n5. the modified watermark advances to what was swept")
    assert res["modified_utc"], "no modified watermark returned"
    got = datetime.fromisoformat(res["modified_utc"])
    assert got == m.parse_zoho_time(LATE["modifiedTime"]), \
        "modified watermark is {}, expected the swept max".format(got)
    print("   modified watermark:", got.isoformat())
except AssertionError as e:
    failed = 1
    print("\nFAILED: {}".format(e))
finally:
    shutil.rmtree(TMP, ignore_errors=True)

print("\n{}".format("ALL LATE-EDIT ASSERTIONS PASSED" if not failed else "FAILED"))
sys.exit(failed)
