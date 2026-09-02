"""Persistence-layer proof.  Run: python tests/test_persistence.py

db.py talks to Postgres, which is not available here, so these tests drive the
REAL db.* functions against a cursor that records the SQL instead of executing
it. That is enough to prove the two things that actually went wrong before:

  1. db.py wrote ten columns that db/schema.sql never created. Every statement
     failed, and _Conn swallowed the error, so the tables stayed empty in
     silence. Test 1 executes every writer and checks each column it touches
     against the schema, so the two files can never drift apart again.
  2. set_followup_state claimed to write "only the arguments given" but its
     defaults were None rather than the _UNSET sentinel, so every call nulled
     the five columns the caller had not passed.

The remaining tests cover the history this layer now has to keep: one row per
chase, one row per remark, and a Delivered Date that is written once and never
moved.
"""
import importlib.util
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.argv = ["main.py"]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


db = load("db", ROOT / "db.py")

# ---------------------------------------------------------------------------
# Reading the schema


def schema_columns():
    """{table: {column, ...}} from db/schema.sql, including later ALTERs."""
    sql = (ROOT / "db" / "schema.sql").read_text(encoding="utf-8")
    tables = {}
    for m in re.finditer(r"CREATE TABLE IF NOT EXISTS (\w+)\s*\((.*?)\n\);",
                         sql, re.S):
        name, body = m.group(1), m.group(2)
        cols = set()
        for line in body.splitlines():
            line = line.strip()
            if not line or line.startswith("--"):
                continue
            word = re.match(r"(\w+)", line)
            if word and word.group(1).upper() not in (
                    "PRIMARY", "UNIQUE", "CHECK", "FOREIGN", "CONSTRAINT"):
                cols.add(word.group(1))
        tables[name] = cols
    for t, c in re.findall(
            r"ALTER TABLE (\w+)\s+ADD COLUMN IF NOT EXISTS (\w+)", sql):
        tables.setdefault(t, set()).add(c)
    return tables


# ---------------------------------------------------------------------------
# A cursor that records instead of executing


class Rec:
    """Captures every statement db.py emits.

    rowcount is settable because some writes are guarded in SQL - a second
    Delivered matches no row - and the code has to react to that."""

    def __init__(self, rowcount=1):
        self.calls = []
        self.rowcount = rowcount

    def execute(self, sql, args=None):
        self.calls.append((" ".join(sql.split()), tuple(args or ())))

    def fetchone(self):
        return (1,)


def capture(fn, *a, **kw):
    """Run one db.* call and return the statements it emitted."""
    rec = Rec(rowcount=kw.pop("_rowcount", 1))

    class Tx:
        def __enter__(self):
            return rec

        def __exit__(self, *exc):
            return False

    real, db.tx = db.tx, Tx
    try:
        fn(*a, **kw)
    finally:
        db.tx = real
    return rec.calls


# Columns a statement writes: INSERT INTO t (a, b) / UPDATE t SET a=%s, b=%s
INSERT_RE = re.compile(r"INSERT INTO (\w+)\s*\(([^)]*)\)", re.I)
UPDATE_RE = re.compile(r"UPDATE (\w+)(?:\s+\w+)?\s+SET\s+(.*?)\s+(?:WHERE|RETURNING)",
                       re.I)


def written(sql):
    """[(table, [column, ...]), ...] for one statement."""
    out = []
    for m in INSERT_RE.finditer(sql):
        out.append((m.group(1),
                    [c.strip() for c in m.group(2).split(",") if c.strip()]))
    for m in UPDATE_RE.finditer(sql):
        cols = re.findall(r"(\w+)\s*=", m.group(2))
        out.append((m.group(1), cols))
    return out


# ---------------------------------------------------------------------------
# Every writer in db.py, called the way main.py calls it.


def every_writer():
    """[(label, [statement, ...]), ...] covering every write path."""
    now = "2026-08-01T15:00:00+05:30"
    return [
        ("begin_run", capture(db.begin_run, "332244", now, now, "x.xlsx")),
        ("mark_stage", capture(db.mark_stage, 1, "desk")),
        ("commit_run", capture(db.commit_run, 1, "332244", now, now, 5, True)),
        ("fail_run", capture(db.fail_run, 1, "boom")),
        ("upsert_ticket", capture(db.upsert_ticket, "332244", ticket_id="id1",
                                  awb="AWB1", created_utc=now, run_id=1,
                                  extra={"courier": "Bluedart"})),
        ("register_awb", capture(db.register_awb, "AWB1", "332244", now)),
        ("save_clickpost", capture(db.save_clickpost, 1,
                                   [("AWB1", "OFD", None, {"a": 1})])),
        ("record_email", capture(db.record_email, "<m@x>", "subj", now,
                                 {"332244": {"awb": "AWB1", "row": [1]}},
                                 columns=["Ticket"], run_id=1)),
        ("record_reply", capture(db.record_reply, 1, "<r@x>", received_at=now)),
        ("record_status", capture(db.record_status, 1, "332244", "OFD",
                                  awb="AWB1")),
        ("apply_latest", capture(db.apply_latest, 1, "332244", "Delivered",
                                 reply_id=1, is_final=True,
                                 final_reason="delivered")),
        ("set_followup_state", capture(db.set_followup_state, 1,
                                       last_reply_at=now)),
        ("rollup_email", capture(db.rollup_email, 1)),
    ]


# ---------------------------------------------------------------------------

FAILED = []


def check(desc, ok, detail=""):
    print("  {} {}{}".format("PASS" if ok else "FAIL", desc,
                             "" if ok else "  <- " + str(detail)))
    if not ok:
        FAILED.append(desc)


def test_schema_covers_every_column_written():
    """Any column db.py writes must exist in db/schema.sql.

    This is the test that was missing: without it, db.py grew ten columns the
    schema never got, every affected statement failed, and _Conn hid it."""
    print("\n1. schema.sql covers every column db.py writes")
    schema = schema_columns()
    missing = []
    for label, calls in every_writer():
        for sql, _ in calls:
            for table, cols in written(sql):
                known = schema.get(table)
                if known is None:
                    missing.append("{}: table {} not in schema".format(
                        label, table))
                    continue
                for col in cols:
                    if col not in known:
                        missing.append("{}: {}.{}".format(label, table, col))
    check("no column is written that the schema lacks", not missing,
          "; ".join(sorted(set(missing))))


def test_set_followup_state_writes_only_what_it_was_given():
    """A caller that names one column must not blank the other five.

    main.py:3317 calls this with followup_sent_at alone. With None defaults that
    call also nulled last_reply_at, next_followup_at, followup_round,
    awaiting_status_since and stalled_at - erasing the chase state of a thread
    whose only news was that a follow-up had gone out."""
    print("\n2. set_followup_state writes only the columns it was passed")
    now = "2026-08-03T11:00:00+05:30"

    calls = capture(db.set_followup_state, 1, followup_sent_at=now)
    cols = [c for sql, _ in calls for _, cs in written(sql) for c in cs]
    check("followup_sent_at alone touches only that column",
          cols == ["followup_sent_at"], "wrote {}".format(cols))

    calls = capture(db.set_followup_state, 1, last_reply_at=now,
                    next_followup_at=None, awaiting_status_since=None)
    cols = [c for sql, _ in calls for _, cs in written(sql) for c in cs]
    check("an explicit None is still written",
          sorted(cols) == ["awaiting_status_since", "last_reply_at",
                           "next_followup_at"], "wrote {}".format(sorted(cols)))

    check("nothing passed emits no statement",
          capture(db.set_followup_state, 1) == [])


def test_delivered_rule():
    """Which remarks set the Delivered Date.

    is_final_status answers "stop chasing", which is true of plain RTO as well -
    the shipment is heading back, so there is nothing to chase. That is NOT the
    same question as "was it delivered". "RTO Delivered" was delivered, back to
    origin, and counts. "RTOInTransit" has not been delivered to anyone yet.
    Negations are rejected on exactly is_final_status's terms, so "Undelivered"
    can never stamp a Delivered Date."""
    print("\n3. is_delivered_status - what earns a Delivered Date")
    m = load("m", ROOT / "main.py")
    for remark, want in [("Delivered", True),
                         ("DELIVERED", True),
                         ("Shipment delivered to consignee", True),
                         ("RTO Delivered", True),
                         ("RTODelivered", True),
                         ("rto delivered", True),
                         ("RTO", False),
                         ("RTOInTransit", False),
                         ("RTO In Transit", False),
                         ("OFD", False),
                         ("Out for delivery", False),
                         ("Undelivered", False),
                         ("Un-delivered", False),
                         ("Not delivered", False),
                         ("could not be delivered", False),
                         ("not yet delivered", False),
                         ("", False)]:
        got = m.is_delivered_status(remark)
        check("{!r} -> {}".format(remark, want), got is want,
              "got {}".format(got))

    check("every delivered remark is also final",
          all(m.is_final_status(r) for r in
              ["Delivered", "RTO Delivered", "RTODelivered"]))
    check("RTO stays final even though it is not delivered",
          m.is_final_status("RTOInTransit")
          and not m.is_delivered_status("RTOInTransit"))


def apply_update(row, sql, args):
    """Apply one `UPDATE t SET ...` to a dict, the way Postgres would.

    Understands the three right-hand sides this layer emits: `%s` (overwrite),
    `coalesce(col, %s)` (keep what is already there), and `now()`. Small, but it
    means write-once is proved by running the real statement rather than by
    reading it and hoping."""
    m = UPDATE_RE.search(sql + " WHERE")
    assert m, sql
    args = list(args)
    for assign in re.split(r",\s*(?=\w+\s*=)", m.group(2)):
        col, expr = [p.strip() for p in assign.split("=", 1)]
        if expr == "now()":
            row[col] = "NOW"
            continue
        if re.fullmatch(r"coalesce\(\s*{}\s*,\s*now\(\)\s*\)".format(col), expr):
            if row.get(col) is None:
                row[col] = "NOW"
            continue
        val = args.pop(0)
        if re.fullmatch(r"coalesce\(\s*{}\s*,\s*%s\s*\)".format(col), expr):
            if row.get(col) is None:
                row[col] = val
        elif expr == "%s":
            row[col] = val
        else:
            raise AssertionError("unhandled right-hand side: {!r}".format(expr))
    return row


def test_delivered_date_is_written_once_and_never_moved():
    """The Delivered Date is a permanent fact about the ticket.

    Bluedart can answer the same thread again after delivery - a courtesy note,
    a correction, a second Delivered on a re-used AWB. None of that may move the
    date the shipment was actually delivered, so the column is written only when
    it is still empty."""
    print("\n4. Delivered Date is set once and never moved")
    first = "2026-08-03T11:00:00+05:30"
    later = "2026-08-09T09:30:00+05:30"

    def ticket_update(calls):
        """The UPDATE tickets statement, of the two this emits."""
        hits = [c for c in calls if c[0].upper().startswith("UPDATE TICKETS")]
        assert len(hits) == 1, calls
        return hits[0]

    row = {}
    calls = capture(db.mark_delivered, "332244", first,
                    status_text="Delivered", awb="AWB1", reply_id=7)
    check("it stamps the ticket and appends one timeline entry",
          tables_of(calls) == ["tickets", "ticket_events"],
          str(tables_of(calls)))
    apply_update(row, *ticket_update(calls))
    check("delivered_at recorded", row.get("delivered_at") == first,
          repr(row.get("delivered_at")))
    check("the remark is kept", row.get("delivered_status_text") == "Delivered")

    calls = capture(db.mark_delivered, "332244", later,
                    status_text="Delivered again", awb="AWB1", reply_id=9)
    apply_update(row, *ticket_update(calls))
    check("a later Delivered does NOT move the date",
          row["delivered_at"] == first, repr(row["delivered_at"]))
    check("nor the remark that earned it",
          row["delivered_status_text"] == "Delivered")

    blocked = capture(db.mark_delivered, "332244", later,
                      status_text="Delivered again", _rowcount=0)
    check("a blocked stamp appends no second delivered event",
          "ticket_events" not in tables_of(blocked), tables_of(blocked))

    check("the ticket is the target",
          "WHERE ticket_number=%s" in ticket_update(calls)[0],
          ticket_update(calls)[0])
    check("nothing is written without a ticket number",
          capture(db.mark_delivered, "", first) == [])


def tables_of(calls):
    return [t for sql, _ in calls for t, _ in written(sql)]


def test_rto_gets_its_own_permanent_timestamp():
    """RTO needs the same treatment Delivered gets.

    is_final_status is true for plain RTO, so the chase stops - but nothing was
    delivered, so delivered_at must stay empty. The moment the shipment turned
    RTO is still a fact worth keeping, and email_tickets.finalised_at is NOT it:
    that is now(), when the run happened to process the reply. This is the
    reply's own received time, written once."""
    print("\n6. RTO is stamped permanently too, and separately from Delivered")
    first = "2026-08-02T15:00:00+05:30"
    later = "2026-08-06T09:00:00+05:30"

    calls = capture(db.mark_rto, "332246", first, status_text="RTO In Transit",
                    awb="77712345680", reply_id=3)
    check("it stamps the ticket and appends one timeline entry",
          tables_of(calls) == ["tickets", "ticket_events"], tables_of(calls))

    upd = [c for c in calls if c[0].upper().startswith("UPDATE TICKETS")][0]
    row = apply_update({}, *upd)
    check("rto_at recorded", row.get("rto_at") == first, row.get("rto_at"))
    check("the remark is kept", row.get("rto_status_text") == "RTO In Transit")
    check("it does NOT touch delivered_at", "delivered_at" not in row, row)

    apply_update(row, *[c for c in capture(
        db.mark_rto, "332246", later, status_text="RTO Delivered")
        if c[0].upper().startswith("UPDATE TICKETS")][0])
    check("a later RTO does NOT move the date", row["rto_at"] == first,
          row["rto_at"])

    blocked = capture(db.mark_rto, "332246", later, _rowcount=0)
    check("a blocked stamp appends no second rto event",
          "ticket_events" not in tables_of(blocked), tables_of(blocked))
    check("nothing is written without a ticket number",
          capture(db.mark_rto, "", first) == [])


def test_the_line_keeps_its_desk_comment():
    """email_tickets.desk_comment_id / desk_posted_at must not stay empty.

    Both columns have existed since the first schema, but apply_latest never
    wrote them, so the one table people read to see a ticket's current state
    could not say whether the comment had reached Desk."""
    print("\n7. the ticket line records its Desk comment")
    cols = [c for sql, _ in capture(
        db.apply_latest, 1, "332244", "Delivered", reply_id=2,
        desk_comment_id="cmt-9", desk_posted_at="2026-08-03T11:02:00+05:30")
        for _, cs in written(sql) for c in cs]
    check("desk_comment_id is written", "desk_comment_id" in cols, cols)
    check("desk_posted_at is written", "desk_posted_at" in cols, cols)


def test_every_followup_is_its_own_row():
    """Each chase is a separate, permanent row.

    threads.json only ever held the LATEST followup_sent_ist, so the third chase
    overwrote the second and the history the report needs never existed. Nothing
    here may collapse two sends into one row - not even two sends recorded with
    the same round number, which the 15:00 no-reply net can produce alongside the
    reply-anchored chase."""
    print("\n5. every follow-up send is appended, never collapsed")
    a = "2026-08-02T18:00:00+05:30"
    b = "2026-08-03T09:00:00+05:30"

    first = capture(db.record_followup, 1, round=1, kind="reply_anchored",
                    sent_at=a, ticket_numbers=["332244"], subject="Re: x")
    check("a row lands in email_followups",
          "email_followups" in tables_of(first), str(tables_of(first)))
    check("the ticket gets a timeline entry",
          "ticket_events" in tables_of(first), str(tables_of(first)))

    second = capture(db.record_followup, 1, round=1, kind="no_reply",
                     sent_at=b, ticket_numbers=["332244"], subject="Re: x")
    check("the same round again still appends",
          "email_followups" in tables_of(second), str(tables_of(second)))
    ins = [sql for sql, _ in second if "INSERT INTO email_followups" in sql]
    check("no ON CONFLICT can swallow a chase",
          ins and "ON CONFLICT" not in ins[0], ins[0] if ins else "no insert")

    args = [a for sql, a in first if "INSERT INTO email_followups" in sql][0]
    check("its own sent_at is stored", a in args, str(args))

    many = capture(db.record_followup, 1, round=2, kind="reply_anchored",
                   sent_at=b, ticket_numbers=["332244", "998877"])
    check("one timeline entry per ticket named in the chase",
          tables_of(many).count("ticket_events") == 2, str(tables_of(many)))
    check("nothing is written for an unknown thread",
          capture(db.record_followup, None, round=1, sent_at=a) == [])


def test_full_timeline_for_one_ticket():
    """The worked example, replayed through the real writers.

    Ticket 332244: escalated 1 Aug 15:00, answered OFD on 2 Aug 15:00, chased
    that evening, answered Delivered on 3 Aug 11:00 and commented on Desk. Every
    one of those has to survive as its own row - OFD in particular, which the old
    layer overwrote the moment Delivered arrived."""
    print("\n6. the complete timeline of ticket 332244")
    T = "332244"
    sent = "2026-08-01T15:00:00+05:30"
    ofd = "2026-08-02T15:00:00+05:30"
    chase = "2026-08-02T18:00:00+05:30"
    deliv = "2026-08-03T11:00:00+05:30"

    events, statements = [], []

    def collect(calls):
        statements.extend(sql for sql, _ in calls)
        for sql, args in calls:
            for table, cols in written(sql):
                if table == "ticket_events":
                    events.append(dict(zip(cols, args)))

    collect(capture(db.record_email, "<m@x>", "Escalation", sent,
                    {T: {"awb": "AWB1", "row": [1]}}, columns=["Ticket"],
                    run_id=1))
    collect(capture(db.record_status, 1, T, "OFD", awb="AWB1",
                    received_at=ofd))
    collect(capture(db.record_followup, 1, round=1, kind="reply_anchored",
                    sent_at=chase, ticket_numbers=[T]))
    collect(capture(db.record_status, 2, T, "Delivered", awb="AWB1",
                    received_at=deliv, desk_comment_id="c99"))
    collect(capture(db.mark_delivered, T, deliv, status_text="Delivered",
                    awb="AWB1", reply_id=2))

    got = [(e["event_type"], e["occurred_at"]) for e in events]
    want = [("email_sent", sent),
            ("status_recorded", ofd),
            ("followup_sent", chase),
            ("status_recorded", deliv),
            ("desk_comment_posted", deliv),
            ("delivered", deliv)]
    check("the timeline reads in order, with nothing missing", got == want,
          "got {}".format(got))

    remarks = [e["status_text"] for e in events if e["event_type"] == "status_recorded"]
    check("OFD survives the arrival of Delivered", remarks == ["OFD", "Delivered"],
          str(remarks))
    check("every event names the ticket",
          all(e["ticket_number"] == T for e in events))
    check("the Desk comment id is kept",
          any(e.get("desk_comment_id") == "c99" for e in events))
    check("the delivered event carries the reply's received time",
          [e["occurred_at"] for e in events
           if e["event_type"] == "delivered"] == [deliv])

    touched = [s for s in statements
               if re.search(r"(UPDATE|DELETE FROM)\s+ticket_events", s, re.I)]
    check("the timeline is append-only - nothing updates or deletes it",
          not touched, str(touched))


def test_analytics_push_is_append_only():
    """Zoho Analytics is an append-only snapshot log, by request.

    No key, no unique id, no upsert: a ticket gets a FRESH row each time its
    journey moves, and nothing already in Analytics is ever modified. So the
    import must be a plain append - matchingColumns or updateadd would silently
    reintroduce the matching the user ruled out."""
    print("\n8. the Analytics import appends, and never matches")
    push = load("push_analytics", ROOT / "db" / "push_analytics.py")
    cfg = push.import_config()

    check("importType is append", cfg.get("importType") == "append",
          cfg.get("importType"))
    check("no matching columns are sent", "matchingColumns" not in cfg, cfg)
    check("nothing asks Analytics to update or truncate",
          cfg.get("importType") not in ("updateadd", "truncateadd"))
    check("sent as JSON", cfg.get("fileType") == "json", cfg.get("fileType"))
    check("dates are declared, not guessed",
          cfg.get("dateFormat") == "dd-MM-yyyy HH:mm:ss", cfg.get("dateFormat"))


def test_only_changed_tickets_are_selected():
    """The push must send what has MOVED, not everything it can see.

    With append and no matching, re-sending an unchanged ticket adds a
    duplicate row that nothing will ever clean up. ticket_events is the signal:
    append-only, so a ticket has moved exactly when it has an event newer than
    the one last pushed."""
    print("\n9. only tickets whose journey moved are pushed")
    push = load("push_analytics", ROOT / "db" / "push_analytics.py")
    sql = " ".join(push.PENDING_SQL.split())

    check("it reads the journey view", "ticket_journey" in sql, sql[:120])
    check("it compares against the last pushed event",
          "analytics_pushes" in sql and "last_event_id" in sql, sql[:200])
    check("an unpushed ticket still qualifies", "coalesce" in sql.lower(),
          sql[:200])
    check("it is driven by ticket_events", "ticket_events" in sql, sql[:200])


if __name__ == "__main__":
    print("=" * 68)
    print("PERSISTENCE LAYER")
    print("=" * 68)
    test_schema_covers_every_column_written()
    test_set_followup_state_writes_only_what_it_was_given()
    test_delivered_rule()
    test_delivered_date_is_written_once_and_never_moved()
    test_rto_gets_its_own_permanent_timestamp()
    test_the_line_keeps_its_desk_comment()
    test_every_followup_is_its_own_row()
    test_full_timeline_for_one_ticket()
    test_analytics_push_is_append_only()
    test_only_changed_tickets_are_selected()
    print()
    if FAILED:
        print("{} FAILED: {}".format(len(FAILED), ", ".join(FAILED)))
        raise SystemExit(1)
    print("ALL PERSISTENCE ASSERTIONS PASSED")
