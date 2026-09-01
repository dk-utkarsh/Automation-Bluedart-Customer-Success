"""Reply-driven 15-hour follow-up cycle.  Run: python tests/test_followups.py

Bluedart answers an escalation with a status per AWB. The ones that came back
"Delivered" are finished; the rest must be chased again 15 hours after the reply
landed, on the same mail trail, and again after that, until they finish.

The finality test is the load-bearing part. "Undelivered" and "Not delivered"
both contain the substring "delivered", and a naive contains-check would mark a
live shipment final and abandon it forever. Negations are therefore checked
BEFORE the positive match, never after.
"""
import importlib.util
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.argv = ["main.py"]
spec = importlib.util.spec_from_file_location("m", ROOT / "main.py")
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

FAILED = []


def check(label, got, want):
    if got != want:
        FAILED.append("{}: got {!r}, want {!r}".format(label, got, want))
        print("  FAIL {}: got {!r}, want {!r}".format(label, got, want))
    else:
        print("  ok   {}".format(label))


# ---------------------------------------------------------------- finality
FINAL = [
    ("Delivered",                      "the base case"),
    ("DELIVERED",                      "upper case"),
    ("  delivered  ",                  "surrounding whitespace"),
    ("RTO Delivered",                  "the user's own example"),
    ("Delivered 02 Sep, POD attached",  "free text after"),
    ("Shipment delivered to customer",  "free text around"),
    ("RTOInTransit",                   "RTO prefix, already heading back"),
]

PENDING = [
    ("Undelivered",                    "the trap: contains 'delivered'"),
    ("Not delivered",                  "negation"),
    ("Not yet delivered",              "negation with a gap"),
    ("Could not be delivered",         "negation with a longer gap"),
    ("Delivery attempted, undelivered", "negation later in the string"),
    ("OFD",                            "real vocabulary"),
    ("Under follow up",                "real vocabulary"),
    ("Attempted today",                "real vocabulary"),
    ("Not traceable",                  "real vocabulary"),
    ("Check with B2B team",            "real vocabulary"),
    ("",                               "no information is not finality"),
    (None,                             "missing value is not finality"),
]

print("is_final_status - final:")
for text, why in FINAL:
    check("{!r} ({})".format(text, why), m.is_final_status(text), True)

print("is_final_status - pending:")
for text, why in PENDING:
    check("{!r} ({})".format(text, why), m.is_final_status(text), False)


# ------------------------------------------------- pending / due to chase
from datetime import timedelta

T0 = m.datetime(2026, 9, 1, 10, 0, tzinfo=m.IST)
# the user's own worked example
AWBS = {"1001": "123445", "1002": "2233445", "1003": "998736"}


def mkthread(statuses, last_reply=T0, round_=0, **kw):
    """A thread whose tickets carry the given remarks. None = never answered."""
    t = {
        "message_id": "<orig@dentalkart>",
        "subject": "Bluedart escalation 01 Sep",
        "sent_ist": (T0 - timedelta(hours=20)).isoformat(),
        "tickets": {num: {"ticketId": "id" + num,
                          "awb": AWBS.get(num, "9" + num)}
                    for num in statuses},
        "status": "awaiting_reply",
        "followup_sent_ist": None,
        "processed": {num: {"status": st} for num, st in statuses.items()
                      if st},
        "final": {num: {"status": s} for num, s in statuses.items()
                  if s and m.is_final_status(s)},
        "last_reply_ist": last_reply.isoformat() if last_reply else None,
        "next_followup_ist": ((last_reply + timedelta(hours=15)).isoformat()
                              if last_reply else None),
        "followup_round": round_,
        "seen_replies": [],
    }
    t.update(kw)
    return t


MIXED = {"1001": "Delivered", "1002": "Delivered", "1003": "OFD"}

print("pending_targets:")
check("two delivered, one OFD -> only the OFD is pending",
      sorted(m.pending_targets(mkthread(MIXED))), ["1003"])
check("nothing answered yet -> all pending",
      sorted(m.pending_targets(mkthread({"1001": None, "1002": None}))),
      ["1001", "1002"])
check("all delivered -> nothing pending",
      sorted(m.pending_targets(mkthread({"1001": "Delivered",
                                         "1002": "RTO Delivered"}))), [])
check("undelivered stays pending",
      sorted(m.pending_targets(mkthread({"1001": "Undelivered"}))), ["1001"])

print("reply_followup_due:")
check("no reply yet -> not due (the 15:00 net owns it)",
      m.reply_followup_due(mkthread(MIXED, last_reply=None),
                           now=T0 + timedelta(days=3)), False)
check("14h after the reply -> not yet",
      m.reply_followup_due(mkthread(MIXED), now=T0 + timedelta(hours=14)), False)
check("exactly 15h after the reply -> due",
      m.reply_followup_due(mkthread(MIXED), now=T0 + timedelta(hours=15)), True)
check("15h but every AWB delivered -> nothing to chase",
      m.reply_followup_due(mkthread({"1001": "Delivered"}),
                           now=T0 + timedelta(hours=15)), False)
check("cap reached -> stop mailing",
      m.reply_followup_due(mkthread(MIXED, round_=m.FOLLOWUP_MAX_ROUNDS),
                           now=T0 + timedelta(hours=15)), False)
check("suppressed by hand -> never",
      m.reply_followup_due(mkthread(MIXED, followup_suppressed=True),
                           now=T0 + timedelta(hours=15)), False)

print("followup_due (the existing 15:00 no-reply net):")
check("a thread that has had a reply is no longer the net's business",
      m.followup_due(mkthread(MIXED), now=T0 + timedelta(days=2)), False)
check("a thread with no reply at all is still chased at 15:00",
      m.followup_due(mkthread({"1001": None}, last_reply=None),
                     now=T0 + timedelta(days=2)), True)


# ------------------------------------------------------- the follow-up mail
text, html = m.followup_body(mkthread(MIXED))

print("followup_body - only what is still outstanding:")
check("the pending AWB is named", "998736" in html, True)
check("its ticket number is named", "1003" in html, True)
check("a delivered AWB is not chased again", "123445" in html, False)
check("nor the other one", "2233445" in html, False)
check("nor their ticket numbers", "1001" in html or "1002" in html, False)
check("the last remark is shown back", "OFD" in html, True)
check("ticket number is the first column",
      html.lower().index("ticket") < html.lower().index("awb"), True)
check("the plain-text part names the pending AWB too", "998736" in text, True)
check("the plain-text part omits the delivered ones", "123445" in text, False)

never = m.followup_body(mkthread({"1001": None, "1002": None}))[1]
check("a thread with no reply yet chases everything",
      "123445" in never and "2233445" in never, True)


print()
if FAILED:
    print("{} FAILURE(S)".format(len(FAILED)))
    sys.exit(1)
print("all green")
