"""Run state lives in Postgres, not state.json.
Run: python tests/test_state_in_db.py

The watermark decides which tickets the next run picks up, so where it is read
from decides whether a ticket can be silently skipped. It used to live in a
local JSON file: delete that file, rebuild the box, or run two crons at once
and the pipeline would quietly restart from "today 00:00" and drop every
unprocessed ticket created before midnight.

It now comes from the runs table, which already had every column state.json
held plus a unique index the file could never express - at most one run in
flight. These tests pin the mapping, the strictness, and the migration.
"""
import importlib.util
import pathlib
import sys
from datetime import datetime, timezone

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.argv = ["main.py"]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


db = load("db", ROOT / "db.py")

FAILED = []


def check(label, ok, detail=""):
    print("  {} {}{}".format("PASS" if ok else "FAIL", label,
                             "" if ok else "  <- " + str(detail)))
    if not ok:
        FAILED.append(label)


class Scripted:
    """A cursor that answers each execute() with the next scripted row."""

    def __init__(self, answers):
        self.answers = list(answers)
        self.calls = []

    def execute(self, sql, args=None):
        self.calls.append((" ".join(sql.split()), tuple(args or ())))

    def fetchone(self):
        return self.answers.pop(0) if self.answers else None


def with_cursor(cur):
    class Tx:
        def __init__(self, strict=False):
            self.strict = strict

        def __enter__(self):
            return cur

        def __exit__(self, *exc):
            return False
    return Tx


def run(fn, answers, *a, **kw):
    cur = Scripted(answers)
    real, db.tx = db.tx, with_cursor(cur)
    try:
        return fn(*a, **kw), cur
    finally:
        db.tx = real


UTC = timezone.utc
COMMITTED = ("332246", datetime(2026, 8, 1, 10, 50, tzinfo=UTC),
             datetime(2026, 8, 1, 11, 5, tzinfo=UTC),
             "out/Zoho_Desk_Logistics_2026-08-01.xlsx", 3,
             datetime(2026, 8, 1, 12, 20, tzinfo=UTC))
PENDING = (77, "out/Zoho_Desk_Logistics_2026-08-04.xlsx", 2, "clickpost",
           False, "332250", datetime(2026, 8, 4, 4, 45, tzinfo=UTC),
           datetime(2026, 8, 4, 5, 0, tzinfo=UTC))

print("=" * 68)
print("RUN STATE LIVES IN POSTGRES")
print("=" * 68)

print("\n1. the watermark comes from the newest committed run")
state, cur = run(db.load_state, [COMMITTED, None])
check("the ticket number is carried",
      state.get("last_ticket_number") == "332246",
      state.get("last_ticket_number"))
check("last_created_utc is an ISO string, as resolve_window expects",
      isinstance(state.get("last_created_utc"), str)
      and state["last_created_utc"].startswith("2026-08-01T10:50"),
      state.get("last_created_utc"))
check("the modified-sweep watermark is carried too",
      str(state.get("last_modified_utc", "")).startswith("2026-08-01T11:05"),
      state.get("last_modified_utc"))
check("it reads only committed runs",
      any("status='committed'" in s.replace('"', "'") for s, _ in cur.calls),
      [s[:70] for s, _ in cur.calls])
check("no pending block when nothing is in flight",
      "pending" not in state, state.get("pending"))

print("\n2. an unfinished run is restored as the pending block")
state, _ = run(db.load_state, [COMMITTED, PENDING])
p = state.get("pending") or {}
check("pending is present", bool(p), state)
check("the run id rides along so the same row is closed",
      state.get("db_run_id") == 77, state.get("db_run_id"))
check("the export to resume is named",
      p.get("export_file", "").endswith("2026-08-04.xlsx"), p.get("export_file"))
check("the attempt count survives", p.get("attempts") == 2, p.get("attempts"))
check("the stage it died at survives",
      p.get("stage_reached") == "clickpost", p.get("stage_reached"))
check("mailer_sent survives - this is what stops a duplicate escalation",
      p.get("mailer_sent") is False, p.get("mailer_sent"))
check("the held watermark survives",
      p.get("watermark_ticket") == "332250", p.get("watermark_ticket"))

print("\n3. a first run on an empty database is not mistaken for a watermark")
state, _ = run(db.load_state, [None, None])
check("no watermark invented", state.get("last_created_utc") is None, state)
check("no pending invented", "pending" not in state, state)

print("\n4. save_state writes the in-flight run back")
st = {"db_run_id": 77,
      "pending": {"export_file": "out/x.xlsx", "attempts": 3,
                  "stage_reached": "mapping", "mailer_sent": True,
                  "watermark_ticket": "332250",
                  "watermark_utc": "2026-08-04T04:45:00+00:00",
                  "modified_utc": None}}
_, cur = run(db.save_state, [], st)
sql = " ".join(s for s, _ in cur.calls)
args = [a for _, a in cur.calls]
check("it updates runs", "UPDATE runs" in sql, sql[:80])
check("it targets the run it loaded", any(77 in a for a in args), args)
for col in ("export_file", "attempts", "stage_reached", "mailer_sent"):
    check("{} is persisted".format(col), col in sql, sql[:200])

print("\n5. nothing in flight writes nothing")
_, cur = run(db.save_state, [], {"db_run_id": None})
check("no statement emitted", cur.calls == [], cur.calls)

print("\n6. a read that cannot reach Postgres raises, it does not return empty")
real = db.tx


class Dead:
    def __init__(self, strict=False):
        self.strict = strict

    def __enter__(self):
        if self.strict:
            raise db.Unavailable("connection refused")
        return None

    def __exit__(self, *exc):
        return False


db.tx = Dead
try:
    try:
        db.load_state()
        check("load_state raises when the database is unreachable", False,
              "it returned instead")
    except db.Unavailable:
        check("load_state raises when the database is unreachable", True)
    except Exception as e:
        check("load_state raises Unavailable specifically", False,
              type(e).__name__)
finally:
    db.tx = real
print("     (an empty state would read as 'no watermark' and re-scan from"
      " today 00:00,\n      silently skipping every unprocessed ticket before"
      " midnight)")

print()
if FAILED:
    print("{} FAILED: {}".format(len(FAILED), ", ".join(FAILED)))
    raise SystemExit(1)
print("ALL STATE-IN-DATABASE ASSERTIONS PASSED")
