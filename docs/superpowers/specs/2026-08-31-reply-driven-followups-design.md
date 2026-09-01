# Reply-driven 15-hour follow-ups

**Date:** 2026-08-31
**Status:** Implemented 2026-09-01 on the existing JSON state (`threads.json`).
**Depends on:** nothing. Decision #4 (DB cutover first) was dropped: §7 assumed the
pending-AWB table needed persisted ticket rows, but `thread["tickets"]` already
carries `{ticketNumber: {ticketId, awb}}` from send time, which is all the table
needs. The Postgres design stands on its own as a later refactor.

**Deviations from this document, as built:**
- §9 cron entry not added — `watch()` already calls `run_followups()` on every
  wake (≤29 min), which is ample resolution for a 15-hour timer.
- At the cap the thread is disarmed and marked `stalled_ist`; its lines are left
  *pending* rather than marked final, so they still read as outstanding.
- Follow-up table columns are Ticket Number / AWB / last remark. The wider
  mapping columns are not persisted on a thread, so they cannot be rebuilt
  until the DB cutover lands.

## 1. Purpose

Today a thread gets **one** chase: 15:00 IST the day after the report, only if nobody
replied. After that, an AWB that Bluedart answered with `OFD` is never chased again —
the thread is left open and forgotten.

This adds a repeating, reply-anchored cycle. When Bluedart replies, each AWB is tested
for finality. AWBs that are done drop out; the rest are chased again 15 hours later on
the same email trail, and again, until they finish or a cap is reached.

The 17:30 daily run is unchanged and stays independent.

## 2. Decisions

Confirmed with the user before this document was written:

| # | Decision | Choice |
|---|---|---|
| 1 | What counts as final | Contains "delivered", **minus negations**; plus RTO-prefixed |
| 2 | Existing 15:00 next-day chase | **Kept**, as the no-reply safety net |
| 3 | Stop condition | Cap at **10 rounds** (~6 days), then mark `stalled` |
| 4 | Sequencing | **Database cutover first**, then this feature |
| 5 | Ticket identification | **Ticket number only**; AWB is a guarded last resort, always logged |
| 6 | Unanswered follow-up | **Re-arm** 15h after sending, so the cadence continues |
| 7 | Send window | **No clamping** — send at exactly +15h, any hour |
| 8 | Cap scope | Per thread, not per AWB |

**Accepted trade-off on #7.** Reply-anchored timing drifts: a reply at 14:00 produces a
follow-up at 05:00 the next morning. The user chose exact +15h over deferring to
business hours. Follow-ups will regularly arrive overnight. Revisit if Rahul reports
that chases are being missed.

## 3. The finality test

One function decides finality everywhere. Negations are checked **first**, so no
negated phrase can ever be read as final.

```python
_UNDELIVERED = re.compile(r"\bun-?delivered\b")
_NEGATED     = re.compile(r"\b(not|never|no|cannot|couldn'?t|could not|unable|fail\w*)"
                          r"[\w\s,'-]{0,24}\bdelivered\b")

def is_final_status(text):
    """True when this remark means the AWB needs no further chasing."""
    t = " ".join(str(text or "").split()).lower()
    if not t:                                  return False
    if _UNDELIVERED.search(t):                 return False
    if _NEGATED.search(t):                     return False
    if "delivered" in t:                       return True
    if t.replace(" ", "").startswith("rto"):   return True   # already heading back
    return False
```

The RTO clause matches the existing ClickPost filter at `main.py:1861-1863`, whose
comment reads: *"anything RTO-prefixed is already on its way back. Neither needs
chasing."*

### Required test vectors

| Input | Expected | Why |
|---|---|---|
| `Delivered` | final | the base case |
| `DELIVERED` / `delivered ` | final | case and whitespace insensitive |
| `RTO Delivered` | final | user's own example |
| `Delivered 02 Sep, POD attached` | final | free text around the word |
| `Shipment delivered to customer` | final | free text around the word |
| `RTOInTransit` | final | RTO prefix |
| `Undelivered` | **pending** | the trap |
| `Not delivered` | **pending** | the trap |
| `Not yet delivered` | **pending** | negation with a gap |
| `Could not be delivered` | **pending** | negation with a gap |
| `Delivery attempted, undelivered` | **pending** | negation later in the string |
| `OFD` | pending | real vocabulary |
| `Under follow up` | pending | real vocabulary |
| `Attempted today` | pending | real vocabulary |
| `Not traceable` | pending | real vocabulary |
| `Check with B2B team` | pending | real vocabulary |
| `` (empty) | pending | no information is not finality |

The six "real vocabulary" values are the actual distinct remarks in
`reply_ticket_statuses` as of migration.

## 4. Ticket identification rule

Carried in from the user's stated rule; already how `locate_target` behaves.

> **Ticket Number identifies the ticket. AWB is for shipment lookup only.**

Resolution order for a reply row:

1. Ticket number present → resolve on it alone; the AWB is never consulted.
2. Registry scanned **for that ticket number** (the registry is AWB-keyed, but the
   match is on ticket number).
3. Desk API lookup by ticket number.
4. Only if the row carries **no ticket number at all**: the AWB may resolve it, and
   only when it maps to exactly one ticket across open threads. Two or more candidates
   is reported and skipped, never guessed.

Every step-4 resolution is written to `reply_ticket_statuses.matched_via` so the
frequency can be audited. Step 4 exists because Bluedart frequently reply on their own
template, which carries an AWB column and no ticket column; requiring a ticket number
there previously dropped every row of every such reply.

## 5. Schema changes

Additive only. All eight existing tables keep their current shape.

```sql
ALTER TABLE email_tickets
    ADD COLUMN IF NOT EXISTS is_final      boolean NOT NULL DEFAULT false,
    ADD COLUMN IF NOT EXISTS final_reason  text,
    ADD COLUMN IF NOT EXISTS finalised_at  timestamptz;

ALTER TABLE email_tickets
    ADD CONSTRAINT email_tickets_final_reason_ck
    CHECK (final_reason IS NULL OR final_reason IN ('delivered','rto','stalled'));

ALTER TABLE emails_sent
    ADD COLUMN IF NOT EXISTS last_reply_at    timestamptz,
    ADD COLUMN IF NOT EXISTS next_followup_at timestamptz,
    ADD COLUMN IF NOT EXISTS followup_round   int NOT NULL DEFAULT 0;

CREATE INDEX IF NOT EXISTS et_pending ON email_tickets (email_id) WHERE NOT is_final;
CREATE INDEX IF NOT EXISTS es_due     ON emails_sent (next_followup_at)
    WHERE next_followup_at IS NOT NULL;
```

`line_status` and `is_final` are **orthogonal and both required**:

| line_status | is_final | Meaning |
|---|---|---|
| `awaiting_reply` | false | never answered — chase it |
| `answered` | false | answered `OFD` — **still chase it** |
| `answered` | true | answered `Delivered` — done |
| `awaiting_reply` | true | cap reached, `final_reason='stalled'` |

Collapsing these into one column is the mistake this feature exists to avoid.

## 6. The cycle

### On every reply that yields at least one parseable status

1. Comment on each ticket (unchanged behaviour).
2. For each affected line, evaluate `is_final_status(status_text)`. On true, set
   `is_final`, `final_reason` (`delivered` or `rto`), `finalised_at`.
   **Finality is never reversed** — a later non-final remark updates
   `latest_status_text` but does not un-finalise the line.
3. `emails_sent.last_reply_at = reply.received_at`
4. If any line remains pending and `followup_round < 10`:
   `next_followup_at = reply.received_at + interval '15 hours'`
5. Roll up `emails_sent.status`: `completed` when every line is final, else
   `partially_answered`.

`received_at` is the reply's `Date` header — "received in our mailbox", per the
requirement. When that header is missing or unparseable, fall back to `fetched_at`.

### The scheduler tick

A cron entry every 15 minutes runs `main.py --followups`:

```sql
SELECT * FROM emails_sent
WHERE status <> 'completed'
  AND NOT followup_suppressed
  AND next_followup_at IS NOT NULL
  AND next_followup_at <= now()
  AND followup_round < 10;
```

For each row:

1. Load pending lines: `SELECT ... FROM email_tickets WHERE email_id=$1 AND NOT is_final`.
2. If none, set `status='completed'`, `next_followup_at=NULL`, and skip.
3. Build and send the follow-up (section 7).
4. On success: `followup_round += 1`,
   `next_followup_at = now() + interval '15 hours'` (decision #6 — the cadence
   continues whether or not a reply arrives).
5. If `followup_round` has reached 10: mark every still-pending line
   `is_final=true, final_reason='stalled'`, set `next_followup_at=NULL`, and roll the
   thread to `completed`. Nothing further is mailed for that thread.

### The no-reply safety net, unchanged

The existing 15:00 IST next-day chase still fires, gated on `last_reply_at IS NULL` —
that is, only for a thread Bluedart has never answered. Once the first reply lands, the
15-hour cycle owns the thread and the 15:00 rule no longer applies to it.

`followup_suppressed` blocks **both** mechanisms. Both threads currently in the database
have `followup_suppressed = true`, so neither will begin chasing on deploy.

## 7. The follow-up mail

Today's follow-up is a bare line: *"Hi, please revert as soon as possible."* It names no
AWBs, so Rahul has to scroll the trail to see what is outstanding.

The new follow-up carries a table of **only the pending AWBs**, rebuilt from `tickets`
joined to `email_tickets`. This is only possible because the ticket rows are now
persisted; previously `Mapping.xlsx` was deleted when the run committed.

- Columns: the same set as the original mapping table, so the trail reads consistently.
- **Ticket Number is the first column**, so replies are more likely to carry it back and
  take the primary resolution path in section 4.
- Threading headers unchanged: `In-Reply-To` and `References` both set to the original
  `emails_sent.message_id`, and the subject is `Re: <original subject>`. This is what
  keeps every chase on one trail.
- A line already answered but not final shows its latest remark, so Rahul can see what
  he last told us.

## 8. Independence from the 17:30 run

No change to the daily run. The two processes cannot collide because `awb_registry`
already guarantees an AWB is escalated exactly once: a pending AWB from an existing
trail can never be picked up into a fresh 17:30 thread. It is chased only by the
15-hour cycle on its original trail, which is the required behaviour.

## 9. Scheduling

```cron
*/15 * * * * /usr/bin/python3 /home/ubuntu/scripts/utkarsh/Automation-Bluedart-Customer-Success/main.py --followups >> .../cron.log 2>&1
```

Fifteen-minute resolution on a fifteen-hour timer is ample. This matches the existing
cron conventions on the host. `flock` is not needed — the `next_followup_at <= now()`
read and the `followup_round` increment happen in one transaction, so a second
overlapping tick cannot double-send.

## 10. Error handling

- **Send fails.** `followup_round` is not incremented and `next_followup_at` is left in
  the past, so the next tick retries. A permanently failing thread retries every 15
  minutes; log loudly rather than silently give up.
- **Reply with no parseable status table.** Comments nothing, and does **not** re-arm
  the timer — an unparseable mail is not evidence of progress. Logged, thread left open.
- **Ambiguous AWB with no ticket number.** Reported and skipped, per section 4.
- **Desk comment fails.** Recorded in `reply_ticket_statuses.post_error`; the line is
  still finality-tested, because Bluedart's answer is a fact independent of whether our
  comment reached Desk.

## 11. Testing

TDD, following the existing `tests/` pattern.

- `is_final_status` against every vector in section 3. This is the highest-value test in
  the feature — the negation cases are the whole point.
- Cycle tests against a throwaway schema: reply arrives → correct lines finalise →
  `next_followup_at` set; tick sends → round increments and re-arms; cap reached →
  remaining lines marked `stalled` and mailing stops.
- Finality is not reversed by a later non-final remark.
- The 15:00 no-reply chase fires only while `last_reply_at IS NULL`.
- A thread with `followup_suppressed` is never mailed by either mechanism.
- The two existing failsafe tests must still pass unchanged.

## 12. Out of scope

- Any change to the 17:30 daily run.
- Business-hours clamping of send times (decision #7 declined it).
- Per-AWB caps — the cap is per thread.
- Reading ticket state back from Desk to stop chasing on closure.

## 13. Open

- **The DB cutover must land first.** `main.py` still reads and writes the JSON files;
  none of this can be built until `db.py` exists and the call sites are switched.
- **Uncommitted server code.** `main.py` on the server carries 210 uncommitted
  insertions, plus untracked `webhook_server.py`, `WEBHOOK.md` and two deploy files.
  That work should be committed before either feature starts, so there is a rollback
  point.
