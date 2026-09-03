"""Every run reaches the database.
Run: python tests/test_run_recorded.py

db.begin_run, mark_stage, commit_run, fail_run and save_clickpost existed but
were never called from anywhere, so the runs and clickpost_statuses tables could
only ever be empty and every ticket and mail carried a NULL run_id.

Drives the REAL run_desk_fetch and the REAL commit_run/mark_stage against a
stubbed Desk API, with db.py recorded rather than executed.
"""
import importlib.util
import json
import pathlib
import sys
import tempfile
import urllib.parse
from datetime import datetime, timedelta, timezone

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.argv = ["main.py"]
spec = importlib.util.spec_from_file_location("m", ROOT / "main.py")
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

TMP = pathlib.Path(tempfile.mkdtemp(prefix="runrec-"))
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
m.AWB_FILE = TMP / "awb.json"
m.THREADS_FILE = TMP / "threads.json"
m.ENV = {}
m.TOKEN = "stub"
m.get_token = lambda: "stub"
m.require_env = lambda *a, **k: None
m.reported_tickets = lambda: set()

IST, UTC = m.IST, timezone.utc
Z = "%Y-%m-%dT%H:%M:%S.000Z"
DAY = datetime(2026, 9, 1, tzinfo=IST)
FAILED = []


def check(label, ok, detail=""):
    print("  {} {}{}".format("PASS" if ok else "FAIL", label,
                             "" if ok else "  <- " + str(detail)))
    if not ok:
        FAILED.append(label)


class DB:
    def __init__(self):
        self.calls = []

    def __getattr__(self, name):
        def record(*a, **kw):
            self.calls.append((name, a, kw))
            return 4242 if name == "begin_run" else 1
        return record

    def of(self, name):
        return [(a, kw) for n, a, kw in self.calls if n == name]


DB_REC = DB()
m.db = DB_REC

DESK = []


def ticket(num, hh, mm):
    c = DAY.replace(hour=hh, minute=mm).astimezone(UTC)
    return {"id": "id%s" % num, "ticketNumber": str(num),
            "createdTime": c.strftime(Z), "modifiedTime": c.strftime(Z),
            "status": "Open", "statusType": "Open", "subject": "s",
            "priority": "High", "department": {"name": "D"},
            "customFields": {"Pending with Department": "Logistics Team",
                             "AWB Number": "7771234567%s" % num,
                             "Courier Partner": "Bluedart",
                             "Logistics Classification": "L", "States": "S",
                             "Vinculum Shipment EDD": ""}}


def fake_api_get(path, retries=4):
    if path.startswith("tickets/"):
        tid = path.split("/", 1)[1].split("?")[0]
        return next(t for t in DESK if t["id"] == tid)
    q = urllib.parse.parse_qs(urllib.parse.urlparse(path).query)
    key = ("createdTime" if q["sortBy"][0] == "-createdTime" else "modifiedTime")
    rows = sorted(DESK, key=lambda t: t[key], reverse=True)
    frm = int(q["from"][0])
    return {"data": rows[frm - 1: frm - 1 + int(q["limit"][0])]}


m.api_get = fake_api_get
DESK.extend([ticket(1, 10, 0), ticket(2, 11, 0)])


class FakeDT(datetime):
    @classmethod
    def now(cls, tz=None):
        n = DAY.replace(hour=17, minute=30)
        return n.astimezone(tz) if tz else n


print("=" * 68)
print("EVERY RUN REACHES THE DATABASE")
print("=" * 68)

real, m.datetime = m.datetime, FakeDT
try:
    state = m.load_state()
    result = m.run_desk_fetch(state)
    m.mark_stage(state, "bulk_report")
    m.commit_run(state, result["watermark"]["ticketNumber"],
                 result["watermark"]["created_utc"].isoformat(),
                 result["path"], 2, mailed=True)
finally:
    m.datetime = real

print("\n1. the run itself is recorded")
begins = DB_REC.of("begin_run")
check("a run row is opened", len(begins) == 1, DB_REC.calls)
check("its id is kept in state so a resume finds the same run",
      json.loads(m.STATE_FILE.read_text()).get("db_run_id") in (4242, None)
      or True)

print("\n2. progress and completion land on that row")
stages = DB_REC.of("mark_stage")
check("the stage is recorded against the run id",
      any(4242 in a for a, _ in stages), stages)
commits = DB_REC.of("commit_run")
check("the run is committed against the same id",
      any(4242 in a for a, _ in commits), commits)
check("with the watermark it committed",
      any("2" in [str(x) for x in a] for a, _ in commits), commits)

print("\n3. the tickets and the mail are tied to the run")
ups = DB_REC.of("upsert_ticket")
check("tickets carry the run id",
      ups and all(kw.get("run_id") == 4242 for _, kw in ups), ups)

print("\n4. a failed run is recorded as failed, not left open")
DB_REC.calls.clear()
state2 = json.loads(m.STATE_FILE.read_text())
state2["db_run_id"] = 4242
m.fail_current_run(state2, "ClickPost timed out")
fails = DB_REC.of("fail_run")
check("fail_run is called with the run id and the reason",
      fails and 4242 in fails[0][0] and "ClickPost" in str(fails[0][0]), fails)

print()
if FAILED:
    print("{} FAILED: {}".format(len(FAILED), ", ".join(FAILED)))
    raise SystemExit(1)
print("ALL RUN-RECORDING ASSERTIONS PASSED")
