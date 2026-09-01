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


# --------------------------------------------- the interval is configurable
# 15 hours is the production cadence. A test run needs to chase immediately
# rather than wait it out, so .env may override it - but ONLY .env, and the
# default must stay 15 so production cannot drift by accident.
print("followup_interval_hours:")
m.ENV = {}
check("with nothing set, production's 15h stands",
      m.followup_interval_hours(), 15)
m.ENV = {"REPLY_FOLLOWUP_HOURS": "0"}
check("0 means chase immediately", m.followup_interval_hours(), 0)
m.ENV = {"REPLY_FOLLOWUP_HOURS": "0.5"}
check("fractional hours allowed", m.followup_interval_hours(), 0.5)
m.ENV = {"REPLY_FOLLOWUP_HOURS": ""}
check("blank falls back to 15", m.followup_interval_hours(), 15)
m.ENV = {"REPLY_FOLLOWUP_HOURS": "not-a-number"}
check("garbage falls back to 15, never crashes", m.followup_interval_hours(), 15)
m.ENV = {"REPLY_FOLLOWUP_HOURS": "-3"}
check("negative is clamped to 0, never sends into the past",
      m.followup_interval_hours(), 0)

print("arm_followup honours it:")
m.ENV = {}
t = mkthread(MIXED)
m.arm_followup(t, T0)
check("default arms 15h after the reply", t["next_followup_ist"],
      (T0 + timedelta(hours=15)).isoformat())
m.ENV = {"REPLY_FOLLOWUP_HOURS": "0"}
t = mkthread(MIXED)
m.arm_followup(t, T0)
check("override arms at the reply time itself", t["next_followup_ist"],
      T0.isoformat())
check("and is therefore due at once", m.reply_followup_due(t, now=T0), True)
m.ENV = {}


# ------------------------------- our own mail must never look like a reply
# Bluedart quotes our follow-up when they answer. If our own table parses as a
# status table, a bare "will update you" is read as a fresh status, the pause
# never engages, and the chase re-arms off our own words - a loop that mails
# them every cycle forever.
sent_html = m.followup_body(mkthread(MIXED))[1]
check("our follow-up table is NOT a status table",
      m.parse_status_table(sent_html), [])

# but the moment Bluedart appends a status to it, it must parse
answered = sent_html.replace("<th>AWB Number</th>",
                             "<th>AWB Number</th><th>Status</th>")
answered = answered.replace("<td>998736</td>", "<td>998736</td><td>Delivered</td>")
rows = m.parse_status_table(answered)
check("the same table WITH a status column does parse", len(rows), 1)
check("and carries their remark", rows[0]["status"] if rows else None, "Delivered")
check("keyed on the ticket number", rows[0]["ticket"] if rows else None, "1003")


# ------------------- the chase reuses the FULL mapping table, minus delivered
FULL_COLS = ["Ticket Number", "AWB Number", "Vinc Shipment EDD", "Delay Days",
             "Concern Type", "Courier Partner", "State"]


def full_thread(statuses, cols=None, extra_cell=None):
    """A thread that persisted the whole mapping row, as a real send now does."""
    t = mkthread(statuses)
    t["columns"] = list(cols or FULL_COLS)
    for num, meta in t["tickets"].items():
        row = [num, meta["awb"], "2026-08-23", "2", "Fake NDR", "Bluedart",
               "Gujarat"]
        if extra_cell is not None:
            row.append(extra_cell)
        meta["row"] = row
    return t


ht = m.followup_body(full_thread(MIXED))[1]
print("followup_body - full mapping format:")
for col in FULL_COLS:
    check("column {!r} is kept".format(col), col in ht, True)
check("pending row's values are there",
      "Fake NDR" in ht and "Bluedart" in ht and "Gujarat" in ht, True)
check("delay days kept", ">2<" in ht.replace(" ", ""), True)
check("the EDD is kept", "2026-08-23" in ht, True)
check("the pending AWB is named", "998736" in ht, True)
check("the DELIVERED rows are gone", "123445" in ht or "2233445" in ht, False)
check("our own chase still never parses as a status table",
      m.parse_status_table(ht), [])

# a Status column in the source must never be echoed back - that is the loop
loop = m.followup_body(full_thread(MIXED, cols=FULL_COLS + ["Status"],
                                   extra_cell="OFD"))[1]
print("a Status column in the mapping is dropped, not echoed:")
check("no Status header in what we send", "<th>Status</th>" in loop, False)
check("and it still does not parse as a reply", m.parse_status_table(loop), [])

print("older threads without the stored row still work:")
old = m.followup_body(mkthread(MIXED))[1]
check("falls back to ticket + AWB", "998736" in old and "1003" in old, True)
check("and is still parser-safe", m.parse_status_table(old), [])


print()
if FAILED:
    print("{} FAILURE(S)".format(len(FAILED)))
    sys.exit(1)
print("all green")
