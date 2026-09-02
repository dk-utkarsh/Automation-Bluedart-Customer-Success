"""The email lifecycle reaches the database.
Run: python tests/test_lifecycle_recorded.py

test_persistence.py proves db.py emits the right SQL. This proves main.py
actually CALLS it, at the right moment and with the right timestamps - the half
that was missing before, when five db functions were written but never wired and
nothing recorded a follow-up send or a delivery at all.

The worked example, driven through the REAL process_replies and run_followups:

    1 Aug 15:00   escalation goes out for ticket 332244
    2 Aug 15:00   Bluedart replies  Status = OFD
    2 Aug 18:00   the 15-hour cycle chases
    3 Aug 11:00   Bluedart replies  Status = Delivered  -> Delivered Date

Only the network edges are stubbed (IMAP, Desk, SMTP) plus db.py itself, which
is recorded rather than executed. Table parsing, ticket resolution, finality and
the chase cycle are all the real code, so a business-logic change would break
these tests rather than slip past them.
"""
import email.utils
import importlib.util
import pathlib
import sys
import tempfile
from datetime import timedelta

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.argv = ["main.py"]
spec = importlib.util.spec_from_file_location("m", ROOT / "main.py")
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

TMP = pathlib.Path(tempfile.mkdtemp(prefix="lifecycle-"))
m.THREADS_FILE = TMP / "threads.json"
m.LOCK_FILE = TMP / ".threads.lock"

SENT = m.datetime(2026, 8, 1, 15, 0, tzinfo=m.IST)
OFD = m.datetime(2026, 8, 2, 15, 0, tzinfo=m.IST)
CHASE = m.datetime(2026, 8, 2, 18, 0, tzinfo=m.IST)
DELIVERED = m.datetime(2026, 8, 3, 11, 0, tzinfo=m.IST)

FAILED = []


def check(label, ok, detail=""):
    print("  {} {}{}".format("PASS" if ok else "FAIL", label,
                             "" if ok else "  <- " + str(detail)))
    if not ok:
        FAILED.append(label)


# ------------------------------------------------------------------ stubs
class DB:
    """Records every db.* call instead of touching Postgres."""

    def __init__(self):
        self.calls = []

    def __getattr__(self, name):
        def record(*a, **kw):
            self.calls.append((name, a, kw))
            return 1                      # every id-returning writer yields 1
        return record

    def of(self, name):
        return [(a, kw) for n, a, kw in self.calls if n == name]


DB_REC = DB()
m.db = DB_REC

INBOX = []
COMMENTS = []


def make_reply(mid, when, rows):
    body = ("<table><tr><th>Ticket Number</th><th>AWB Number</th>"
            "<th>Status</th></tr>")
    for tkt, awb, status in rows:
        body += "<tr><td>{}</td><td>{}</td><td>{}</td></tr>".format(
            tkt, awb, status)
    body += "</table>"
    raw = ("From: rahul@bluedart.com\r\n"
           "To: escalations@dentalkart.com\r\n"
           "Subject: Re: Bluedart escalation 01 Aug\r\n"
           "Message-ID: {}\r\n"
           "In-Reply-To: <orig@dentalkart>\r\n"
           "Date: {}\r\n"
           "Content-Type: text/html; charset=utf-8\r\n\r\n{}"
           ).format(mid, email.utils.format_datetime(when), body)
    return email.message_from_string(raw)


class FakeIMAP:
    def logout(self):
        pass


m.TOKEN = "stub"
m.get_token = lambda: "stub"
m.load_registry = lambda: {}
m.save_registry = lambda reg: None
m.imap_connect = lambda quiet=False: FakeIMAP()
m.find_thread_replies = lambda C, thread, siblings=(): [uid for uid, _ in INBOX]
m.fetch_message = lambda C, uid: dict(INBOX)[uid]
m.comment_on_ticket = (lambda ticket_id, ticket_number, status, when:
                       (COMMENTS.append((str(ticket_number), status)),
                        {"id": "cmt-{}".format(len(COMMENTS))})[1])
m.send_followup = lambda thread: True
m.ENV = {"MAIL_TO": "bluedart@example.com", "MAIL_FROM": "esc@dentalkart.com"}

m.save_threads([{
    "message_id": "<orig@dentalkart>",
    "subject": "Bluedart escalation 01 Aug",
    "sent_ist": SENT.isoformat(),
    "db_email_id": 1,
    "tickets": {"332244": {"ticketId": "id332244", "awb": "77712345678"}},
    "status": "awaiting_reply",
    "followup_sent_ist": None,
    "processed": {},
    "seen_replies": [],
}])

print("=" * 68)
print("EMAIL LIFECYCLE REACHES THE DATABASE")
print("=" * 68)

# ============================================ 2 Aug 15:00 - Bluedart says OFD
INBOX.append((b"1", make_reply("<r1@bd>", OFD,
                               [("332244", "77712345678", "OFD")])))
m.process_replies(verbose=False)

print("\n1. the OFD reply is recorded with the time it arrived")
statuses = DB_REC.of("record_status")
check("a remark was recorded", len(statuses) == 1, DB_REC.calls)
if statuses:
    a, kw = statuses[0]
    check("it is the OFD remark", "OFD" in a, a)
    check("stamped with the reply's own Date header",
          kw.get("received_at") == OFD, kw.get("received_at"))
    check("the Desk comment id is carried",
          kw.get("desk_comment_id") == "cmt-1", kw.get("desk_comment_id"))
check("OFD does NOT stamp a Delivered Date",
      DB_REC.of("mark_delivered") == [], DB_REC.of("mark_delivered"))

# ============================================ 2 Aug 18:00 - the cycle chases
t = m.load_threads()[0]
t["next_followup_ist"] = CHASE.isoformat()
m.save_threads([t])
DB_REC.calls.clear()
m.run_followups(verbose=False, now=CHASE)

print("\n2. the chase is recorded as its own row")
ups = DB_REC.of("record_followup")
check("one follow-up recorded", len(ups) == 1, DB_REC.calls)
if ups:
    a, kw = ups[0]
    check("against the right thread", 1 in a, a)
    check("with the moment it went out", kw.get("sent_at") == CHASE,
          kw.get("sent_at"))
    check("naming the ticket it chased",
          list(kw.get("ticket_numbers") or []) == ["332244"],
          kw.get("ticket_numbers"))
    check("and which round it was", kw.get("round") == 1, kw.get("round"))

# ======================================= 3 Aug 11:00 - Bluedart says Delivered
DB_REC.calls.clear()
INBOX.append((b"2", make_reply("<r2@bd>", DELIVERED,
                               [("332244", "77712345678", "Delivered")])))
m.process_replies(verbose=False)

print("\n3. Delivered stamps the permanent Delivered Date")
dels = DB_REC.of("mark_delivered")
check("the delivery was recorded", len(dels) == 1, DB_REC.calls)
if dels:
    a, kw = dels[0]
    check("for this ticket", "332244" in a, a)
    check("dated when Bluedart said it was delivered",
          DELIVERED in a or kw.get("delivered_at") == DELIVERED, (a, kw))
    check("keeping the remark that earned it",
          kw.get("status_text") == "Delivered", kw.get("status_text"))

print("\n4. the earlier remark was not disturbed")
check("the Delivered remark is also appended to the history",
      any("Delivered" in a for a, _ in DB_REC.of("record_status")),
      DB_REC.of("record_status"))
check("both remarks reached the Desk ticket",
      COMMENTS == [("332244", "OFD"), ("332244", "Delivered")], COMMENTS)

print()
if FAILED:
    print("{} FAILED: {}".format(len(FAILED), ", ".join(FAILED)))
    raise SystemExit(1)
print("ALL LIFECYCLE ASSERTIONS PASSED")
