"""The 15-hour reply-driven follow-up cycle, end to end.
Run: python tests/test_reply_cycle.py

Drives the REAL process_replies against the user's own worked example:

    reply 1   123445 = Delivered      -> final
              2233445 = Delivered     -> final
              998736 = OFD            -> still pending, chase in 15h
    reply 2   998736 = Under follow up -> commented AGAIN, still pending, chase in 15h
    reply 3   998736 = Delivered      -> commented, final, thread closed, no more mail

Only the network edges are stubbed (IMAP, Desk, the clock source is the reply's own
Date header). Table parsing, ticket resolution and completion are the real code.

Two behaviours this pins down, both of which the pre-2026-09-01 code got wrong:
  * a SECOND remark on a ticket already commented was silently dropped, so the
    "Under follow up" -> "Delivered" progression never reached the ticket;
  * a ticket answered "OFD" counted as done, so the thread closed and the AWB was
    never chased again.
"""
import email.utils
import importlib.util
import json
import pathlib
import sys
import tempfile
from datetime import timedelta

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.argv = ["main.py"]
spec = importlib.util.spec_from_file_location("m", ROOT / "main.py")
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

TMP = pathlib.Path(tempfile.mkdtemp(prefix="replycycle-"))
m.THREADS_FILE = TMP / "threads.json"

T0 = m.datetime(2026, 9, 1, 10, 0, tzinfo=m.IST)
SENT = T0 - timedelta(hours=4)

FAILED = []


def check(label, got, want):
    if got != want:
        FAILED.append(label)
        print("  FAIL {}: got {!r}, want {!r}".format(label, got, want))
    else:
        print("  ok   {}".format(label))


# ------------------------------------------------------------------ stubs
COMMENTS = []          # (ticket_number, status) in the order they were posted
INBOX = []             # (uid, message) - grows as Bluedart replies


def make_reply(mid, when, rows):
    body = ("<table><tr><th>Ticket Number</th><th>AWB Number</th>"
            "<th>Status</th></tr>")
    for tkt, awb, status in rows:
        body += "<tr><td>{}</td><td>{}</td><td>{}</td></tr>".format(tkt, awb, status)
    body += "</table>"
    raw = ("From: rahul@bluedart.com\r\n"
           "To: escalations@dentalkart.com\r\n"
           "Subject: Re: Bluedart escalation 01 Sep\r\n"
           "Message-ID: {}\r\n"
           "In-Reply-To: <orig@dentalkart>\r\n"
           "Date: {}\r\n"
           "Content-Type: text/html; charset=utf-8\r\n\r\n{}"
           ).format(mid, email.utils.format_datetime(when), body)
    return email.message_from_string(raw)


class FakeIMAP:
    def logout(self):
        pass


m.TOKEN = "stub-token"
m.get_token = lambda: "stub-token"
m.load_registry = lambda: {}
m.save_registry = lambda reg: None
m.imap_connect = lambda quiet=False: FakeIMAP()
m.find_thread_replies = lambda C, thread: [uid for uid, _ in INBOX]
m.fetch_message = lambda C, uid: dict(INBOX)[uid]
m.comment_on_ticket = (lambda ticket_id, ticket_number, status, when:
                       COMMENTS.append((str(ticket_number), status)))

# --------------------------------------------------------------- the thread
m.save_threads([{
    "message_id": "<orig@dentalkart>",
    "subject": "Bluedart escalation 01 Sep",
    "sent_ist": SENT.isoformat(),
    "tickets": {"1001": {"ticketId": "id1001", "awb": "123445"},
                "1002": {"ticketId": "id1002", "awb": "2233445"},
                "1003": {"ticketId": "id1003", "awb": "998736"}},
    "status": "awaiting_reply",
    "followup_sent_ist": None,
    "processed": {},
    "seen_replies": [],
}])


def only():
    return m.load_threads()[0]


# ============================================================ reply 1 at T0
INBOX.append((b"1", make_reply("<r1@bd>", T0, [
    ("1001", "123445", "Delivered"),
    ("1002", "2233445", "Delivered"),
    ("1003", "998736", "OFD"),
])))
m.process_replies(verbose=False)
t = only()

print("reply 1 - two delivered, one OFD:")
check("all three remarks reached their tickets", sorted(COMMENTS),
      [("1001", "Delivered"), ("1002", "Delivered"), ("1003", "OFD")])
check("the two delivered AWBs are final", sorted(t.get("final") or {}),
      ["1001", "1002"])
check("the OFD AWB is still pending", sorted(m.pending_targets(t)), ["1003"])
check("thread stays open", t["status"], "awaiting_reply")
check("next chase armed 15h after the reply landed",
      t.get("next_followup_ist"), (T0 + timedelta(hours=15)).isoformat())
check("the 15:00 no-reply net has stepped aside",
      m.followup_due(t, now=T0 + timedelta(days=2)), False)
check("the 15h chase is not due yet",
      m.reply_followup_due(t, now=T0 + timedelta(hours=14)), False)
check("the 15h chase is due at +15h",
      m.reply_followup_due(t, now=T0 + timedelta(hours=15)), True)

# ================================================== reply 2 at T0+16h
T1 = T0 + timedelta(hours=16)
INBOX.append((b"2", make_reply("<r2@bd>", T1, [
    ("1003", "998736", "Under follow up"),
])))
m.process_replies(verbose=False)
t = only()

print("reply 2 - the same AWB answered a second time:")
check("the new remark was commented, not swallowed as a duplicate",
      COMMENTS[-1], ("1003", "Under follow up"))
check("re-reading reply 1 added nothing", len(COMMENTS), 4)
check("still pending", sorted(m.pending_targets(t)), ["1003"])
check("thread still open", t["status"], "awaiting_reply")
check("chase re-armed from the NEW reply",
      t.get("next_followup_ist"), (T1 + timedelta(hours=15)).isoformat())

# ================================================== reply 3 at T0+40h
T2 = T0 + timedelta(hours=40)
INBOX.append((b"3", make_reply("<r3@bd>", T2, [
    ("1003", "998736", "Delivered"),
])))
m.process_replies(verbose=False)
t = only()

print("reply 3 - finally delivered:")
check("the final remark reached the ticket too",
      COMMENTS[-1], ("1003", "Delivered"))
check("every AWB is now final", sorted(t.get("final") or {}),
      ["1001", "1002", "1003"])
check("nothing left to chase", m.pending_targets(t), {})
check("thread closed", t["status"], "completed")
check("no further mail is due, ever",
      m.reply_followup_due(t, now=T2 + timedelta(days=30)), False)

# ================================================== the sending loop
SENT = []
m.send_followup = lambda t: (SENT.append(t["subject"]) or True)


def armed(round_=0, last=T0, **kw):
    t = {
        "message_id": "<t2@dk>",
        "subject": "Bluedart escalation 02 Sep",
        "sent_ist": (T0 - timedelta(hours=4)).isoformat(),
        "tickets": {"2001": {"ticketId": "id2001", "awb": "555111"}},
        "status": "awaiting_reply",
        "followup_sent_ist": None,
        "processed": {"2001": {"status": "OFD"}},
        "final": {},
        "last_reply_ist": last.isoformat() if last else None,
        "next_followup_ist": ((last + timedelta(hours=15)).isoformat()
                              if last else None),
        "followup_round": round_,
        "seen_replies": [],
    }
    t.update(kw)
    return t


print("run_followups - the 15h cycle:")
m.save_threads([armed()])
check("nothing goes out before the 15h mark",
      m.run_followups(verbose=False, now=T0 + timedelta(hours=14)), 0)

DUE = T0 + timedelta(hours=15)
m.save_threads([armed()])
check("one chase goes out at the 15h mark",
      m.run_followups(verbose=False, now=DUE), 1)
t = only()
check("the round is counted", t["followup_round"], 1)
check("and re-armed 15h out, so silence does not end the cadence",
      t["next_followup_ist"], (DUE + timedelta(hours=15)).isoformat())

m.save_threads([armed(round_=m.FOLLOWUP_MAX_ROUNDS - 1)])
m.run_followups(verbose=False, now=DUE)
t = only()
check("the last permitted round disarms the thread",
      t["next_followup_ist"], None)
check("and it is not mailed again",
      m.run_followups(verbose=False, now=DUE + timedelta(days=5)), 0)

print("run_followups - the 15:00 no-reply net, unchanged:")
m.save_threads([armed(last=None)])
check("a thread Bluedart never answered is still chased at 15:00",
      m.run_followups(verbose=False,
                      now=T0.replace(hour=15, minute=1) + timedelta(days=1)), 1)
check("and only once", m.run_followups(verbose=False,
                                       now=T0 + timedelta(days=3)), 0)


# ============================== one ticket listed twice in the SAME reply
# Dropping the per-ticket guard is what lets a LATER reply update a ticket
# again. It must not also let a single reply comment the same ticket twice -
# the guard that replaced it is per reply, not per thread for all time.
del COMMENTS[:]
del INBOX[:]
m.save_threads([{
    "message_id": "<dup@dk>",
    "subject": "Bluedart escalation 03 Sep",
    "sent_ist": (T0 - timedelta(hours=4)).isoformat(),
    "tickets": {"3001": {"ticketId": "id3001", "awb": "777222"}},
    "status": "awaiting_reply",
    "followup_sent_ist": None,
    "processed": {},
    "seen_replies": [],
}])
INBOX.append((b"9", make_reply("<dup1@bd>", T0, [
    ("3001", "777222", "OFD"),
    ("3001", "777222", "Under follow up"),
])))
m.process_replies(verbose=False)

print("a ticket listed twice in one reply:")
check("commented once, not twice", len(COMMENTS), 1)
check("and it is still pending", sorted(m.pending_targets(only())), ["3001"])


# ========================= a conversational reply PAUSES the chase
# Bluedart answering "we will update you" is an answer, even though it carries
# no status. Chasing them 15h later regardless reads as if we ignored them.
# The cycle must stand down and wait for a reply that actually carries a table.
from email.message import EmailMessage


def plain_reply(mid, when, text):
    msg = EmailMessage()
    msg["From"] = "rahul@bluedart.com"
    msg["To"] = "escalations@dentalkart.com"
    msg["Subject"] = "Re: Bluedart escalation 06 Sep"
    msg["Message-ID"] = mid
    msg["In-Reply-To"] = "<pause@dk>"
    msg["Date"] = email.utils.format_datetime(when)
    msg.set_content("<p>" + text + "</p>", subtype="html")
    return msg


del COMMENTS[:]
del INBOX[:]
del SENT[:]
m.save_threads([{
    "message_id": "<pause@dk>", "subject": "Bluedart escalation 06 Sep",
    "sent_ist": (T0 - timedelta(hours=4)).isoformat(),
    "tickets": {"7001": {"ticketId": "id7001", "awb": "313131"}},
    "status": "awaiting_reply", "followup_sent_ist": None,
    "processed": {}, "seen_replies": [],
}])

INBOX.append((b"70", make_reply("<p1@bd>", T0, [("7001", "313131", "OFD")])))
m.process_replies(verbose=False)
t = only()
print("status reply arms the cycle:")
check("armed 15h out", t["next_followup_ist"],
      (T0 + timedelta(hours=15)).isoformat())

T_CHAT = T0 + timedelta(hours=16)
INBOX.append((b"71", plain_reply("<p2@bd>", T_CHAT, "Hi Manish, we will update you.")))
before = len(COMMENTS)
m.process_replies(verbose=False)
t = only()
print("a conversational reply pauses it:")
check("nothing was commented", len(COMMENTS), before)
check("the chase is disarmed", t.get("next_followup_ist"), None)
check("we record that we are waiting for a status",
      t.get("awaiting_status_since"), T_CHAT.isoformat())
check("last_reply_ist moves to their message", t.get("last_reply_ist"),
      T_CHAT.isoformat())
check("no chase is due", m.reply_followup_due(t, now=T_CHAT + timedelta(days=5)),
      False)
check("run_followups sends nothing",
      m.run_followups(verbose=False, now=T_CHAT + timedelta(days=5)), 0)
check("the 15:00 net stays down too",
      m.followup_due(t, now=T_CHAT + timedelta(days=5)), False)
check("the chatty mail is marked seen", "<p2@bd>" in t.get("seen_replies"), True)

T_REAL = T_CHAT + timedelta(hours=6)
INBOX.append((b"72", make_reply("<p3@bd>", T_REAL,
                                [("7001", "313131", "Under follow up")])))
m.process_replies(verbose=False)
t = only()
print("the real status reply restarts it:")
check("the remark reached the ticket", COMMENTS[-1], ("7001", "Under follow up"))
check("re-armed 15h from THAT reply", t["next_followup_ist"],
      (T_REAL + timedelta(hours=15)).isoformat())
check("no longer waiting for a status", t.get("awaiting_status_since"), None)
check("still pending", sorted(m.pending_targets(t)), ["7001"])
check("chase fires at the new +15h",
      m.reply_followup_due(t, now=T_REAL + timedelta(hours=15)), True)


print()
if FAILED:
    print("{} FAILURE(S)".format(len(FAILED)))
    sys.exit(1)
print("all green")
