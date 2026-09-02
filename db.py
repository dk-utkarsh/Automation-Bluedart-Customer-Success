"""PostgreSQL persistence for the Bluedart escalation pipeline.

main.py keeps its logic and its JSON files; this module mirrors every durable
fact into Postgres as it happens, so the database is a complete record rather
than something a migration script has to catch up on afterwards.

WRITES ARE BEST-EFFORT, DELIBERATELY. If DATABASE_URL is unset, psycopg is not
installed, or the database is unreachable, every call here becomes a no-op and
logs once. The mail pipeline must never fail because a logging database is
down - an escalation that does not go out costs a customer, a row that does not
land costs a report. Once the data is trusted the reads can move over too, and
that is the point at which a failure should become fatal.

Connection details come from DATABASE_URL in .env, matching db/schema.sql.
"""
import json
import os

try:                                    # psycopg 3 preferred, psycopg2 accepted
    import psycopg
    _V = 3
except ImportError:                     # pragma: no cover - environment dependent
    try:
        import psycopg2 as psycopg
        _V = 2
    except ImportError:
        psycopg = None
        _V = 0

_URL = None
_WARNED = False
_DISABLED = False


def configure(url):
    """Point the module at a database. None or empty disables every write."""
    global _URL, _DISABLED, _WARNED
    _URL = (url or "").strip() or None
    _DISABLED = _URL is None
    _WARNED = False
    # Said out loud every start. Silence used to be indistinguishable from
    # working: with no DATABASE_URL every write became a no-op and printed
    # nothing, so the tables stayed empty and looked healthy.
    if _URL is None:
        print("  db: persistence OFF - DATABASE_URL is not set. Running on JSON.")
    elif psycopg is None:
        print("  db: persistence OFF - no psycopg driver installed. "
              "Running on JSON.")
    else:
        print("  db: persistence ON (psycopg {}).".format(_V))


def enabled():
    return bool(_URL) and psycopg is not None and not _DISABLED


def _warn(exc):
    """Say why persistence is off - once, not once per row."""
    global _WARNED
    if not _WARNED:
        _WARNED = True
        print("  db: persistence OFF ({}: {}). The pipeline continues on JSON."
              .format(exc.__class__.__name__, exc))


class _Conn:
    """Context manager yielding a cursor, committing on success.

    Swallows every database error by design - see the module docstring. Returns
    None as the cursor when persistence is off, so callers guard with `if cur`.
    """

    def __enter__(self):
        self.conn = None
        if not enabled():
            return None
        try:
            self.conn = psycopg.connect(_URL)
            return self.conn.cursor()
        except Exception as e:
            _warn(e)
            self.conn = None
            return None

    def __exit__(self, exc_type, exc, tb):
        if self.conn is None:
            return exc_type is not None and issubclass(exc_type, Exception)
        try:
            if exc_type is None:
                self.conn.commit()
            else:
                self.conn.rollback()
                _warn(exc)
        except Exception as e:
            _warn(e)
        finally:
            try:
                self.conn.close()
            except Exception:
                pass
        # Suppress: a persistence failure must not propagate into the pipeline.
        return exc_type is not None and issubclass(exc_type, Exception)


def tx():
    return _Conn()


# ---------------------------------------------------------------- runs
def begin_run(watermark_ticket=None, watermark_utc=None, modified_utc=None,
              export_file=None):
    """Open a run row. Returns its id, or None when persistence is off."""
    with tx() as cur:
        if not cur:
            return None
        # runs_one_pending allows a single pending row. A run that died without
        # committing would otherwise block every future begin_run, silently, so
        # starting a new run closes out whatever was left open.
        cur.execute(
            "UPDATE runs SET status='failed', finished_at=now(), "
            "error=coalesce(error, 'superseded by a later run') "
            "WHERE status='pending'")
        cur.execute(
            "INSERT INTO runs (started_at, status, watermark_ticket, "
            "watermark_utc, modified_utc, export_file) "
            "VALUES (now(), 'pending', %s, %s, %s, %s) RETURNING id",
            (watermark_ticket, watermark_utc, modified_utc, export_file))
        row = cur.fetchone()
        return row[0] if row else None
    return None


def mark_stage(run_id, stage):
    if run_id is None:
        return
    with tx() as cur:
        if cur:
            cur.execute("UPDATE runs SET stage_reached=%s WHERE id=%s",
                        (stage, run_id))


def commit_run(run_id, watermark_ticket=None, watermark_utc=None,
               modified_utc=None, rows_count=None, mailer_sent=False):
    if run_id is None:
        return
    with tx() as cur:
        if cur:
            cur.execute(
                "UPDATE runs SET status='committed', finished_at=now(), "
                "watermark_ticket=coalesce(%s, watermark_ticket), "
                "watermark_utc=coalesce(%s, watermark_utc), "
                "modified_utc=coalesce(%s, modified_utc), "
                "rows_count=%s, mailer_sent=%s WHERE id=%s",
                (watermark_ticket, watermark_utc, modified_utc, rows_count,
                 bool(mailer_sent), run_id))


def fail_run(run_id, error):
    if run_id is None:
        return
    with tx() as cur:
        if cur:
            cur.execute("UPDATE runs SET status='failed', finished_at=now(), "
                        "error=%s WHERE id=%s", (str(error)[:4000], run_id))


# ------------------------------------------------------------- tickets
def upsert_ticket(ticket_number, ticket_id=None, awb=None, created_utc=None,
                  modified_utc=None, extra=None, run_id=None):
    """One ticket. Re-seen tickets refresh last_seen_at and any new detail.

    Nothing is overwritten with NULL: a later sighting that lacks a field keeps
    what the earlier one knew."""
    if not ticket_number:
        return
    extra = extra or {}
    with tx() as cur:
        if not cur:
            return
        cur.execute(
            "INSERT INTO tickets (ticket_number, ticket_id, awb, created_utc, "
            "  modified_utc, courier, vinculum_edd, logistics_class, states, "
            "  first_run_id) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
            "ON CONFLICT (ticket_number) DO UPDATE SET "
            "  ticket_id       = coalesce(EXCLUDED.ticket_id, tickets.ticket_id), "
            "  awb             = coalesce(EXCLUDED.awb, tickets.awb), "
            "  modified_utc    = coalesce(EXCLUDED.modified_utc, tickets.modified_utc), "
            "  courier         = coalesce(EXCLUDED.courier, tickets.courier), "
            "  vinculum_edd    = coalesce(EXCLUDED.vinculum_edd, tickets.vinculum_edd), "
            "  logistics_class = coalesce(EXCLUDED.logistics_class, tickets.logistics_class), "
            "  states          = coalesce(EXCLUDED.states, tickets.states), "
            "  last_seen_at    = now()",
            (str(ticket_number), ticket_id, awb, created_utc, modified_utc,
             extra.get("courier"), extra.get("vinculum_edd"),
             extra.get("logistics_class"), extra.get("states"), run_id))


def mark_delivered(ticket_number, delivered_at, status_text=None, awb=None,
                   reply_id=None):
    """Stamp the Delivered Date on the ticket. Written once, never moved.

    Every column uses coalesce(column, %s), so the FIRST delivered remark wins
    and everything after it is a no-op. Bluedart answering the same thread again
    after delivery - a courtesy note, a correction, a second Delivered on a
    re-used AWB - must not change when the shipment was actually delivered.

    What counts as delivered is is_delivered_status() in main.py, never decided
    here: this module stores facts, it does not classify them."""
    if not ticket_number or not delivered_at:
        return
    with tx() as cur:
        if cur:
            cur.execute(
                "UPDATE tickets SET "
                "  delivered_at          = coalesce(delivered_at, %s), "
                "  delivered_status_text = coalesce(delivered_status_text, %s), "
                "  delivered_awb         = coalesce(delivered_awb, %s), "
                "  delivered_reply_id    = coalesce(delivered_reply_id, %s), "
                "  delivered_recorded_at = coalesce(delivered_recorded_at, now()) "
                # Belt and braces with the coalesce above: this also makes a
                # repeat stamp match NO row, which is how the timeline below
                # knows not to log a second delivery of the same shipment.
                "WHERE ticket_number=%s AND delivered_at IS NULL",
                (delivered_at, status_text, awb, reply_id, str(ticket_number)))
            if cur.rowcount:
                _event(cur, ticket_number, "delivered", delivered_at,
                       status_text=status_text, awb=awb, reply_id=reply_id)


def mark_rto(ticket_number, rto_at, status_text=None, awb=None, reply_id=None):
    """Stamp when the shipment turned RTO. Written once, never moved.

    The mirror of mark_delivered, and deliberately a SEPARATE column: a plain
    RTO ends the chase but delivers nothing, so it must never fill
    delivered_at. "RTO Delivered" is the other case and goes to mark_delivered
    instead - is_delivered_status() in main.py decides which, never this module.

    rto_at is the reply's own received time. email_tickets.finalised_at is not a
    substitute: that is now(), when the run happened to read the mail."""
    if not ticket_number or not rto_at:
        return
    with tx() as cur:
        if cur:
            cur.execute(
                "UPDATE tickets SET "
                "  rto_at          = coalesce(rto_at, %s), "
                "  rto_status_text = coalesce(rto_status_text, %s), "
                "  rto_reply_id    = coalesce(rto_reply_id, %s), "
                "  rto_recorded_at = coalesce(rto_recorded_at, now()) "
                "WHERE ticket_number=%s AND rto_at IS NULL",
                (rto_at, status_text, reply_id, str(ticket_number)))
            if cur.rowcount:
                _event(cur, ticket_number, "rto", rto_at,
                       status_text=status_text, awb=awb, reply_id=reply_id)


def register_awb(awb, ticket_number, created_utc=None):
    """The dedup registry. An AWB already here is never escalated again."""
    if not awb or not ticket_number:
        return
    with tx() as cur:
        if cur:
            cur.execute(
                "INSERT INTO awb_registry (awb, ticket_number, created_utc) "
                "VALUES (%s,%s,coalesce(%s, now())) ON CONFLICT (awb) DO NOTHING",
                (str(awb), str(ticket_number), created_utc))


def save_clickpost(run_id, rows):
    """rows: iterable of (awb, status, edd, raw_dict)."""
    if run_id is None:
        return 0
    n = 0
    with tx() as cur:
        if not cur:
            return 0
        for awb, status, edd, raw in rows:
            cur.execute(
                "INSERT INTO clickpost_statuses (run_id, awb, status, edd, raw) "
                "VALUES (%s,%s,%s,%s,%s) ON CONFLICT (run_id, awb) DO UPDATE SET "
                "  status=EXCLUDED.status, edd=EXCLUDED.edd, raw=EXCLUDED.raw",
                (run_id, str(awb), status, edd,
                 json.dumps(raw) if raw is not None else None))
            n += 1
    return n


# ------------------------------------------------------------- timeline
def _event(cur, ticket_number, event_type, occurred_at, **kw):
    """Append one timeline row on an OPEN cursor.

    Takes the cursor rather than opening its own so an event lands in the same
    transaction as the thing it describes - a chase and its timeline entries
    commit together or not at all."""
    if not cur or not ticket_number or not occurred_at:
        return
    cur.execute(
        "INSERT INTO ticket_events (ticket_number, event_type, occurred_at, "
        "  status_text, awb, email_id, reply_id, followup_id, run_id, "
        "  desk_comment_id, detail) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        (str(ticket_number), event_type, occurred_at, kw.get("status_text"),
         kw.get("awb"), kw.get("email_id"), kw.get("reply_id"),
         kw.get("followup_id"), kw.get("run_id"), kw.get("desk_comment_id"),
         json.dumps(kw["detail"]) if kw.get("detail") is not None else None))


def record_event(ticket_number, event_type, occurred_at, **kw):
    """Append one entry to a ticket's timeline. Never updates, never deletes."""
    with tx() as cur:
        if cur:
            _event(cur, ticket_number, event_type, occurred_at, **kw)


def record_followup(email_id, round=None, kind=None, sent_at=None,
                    to_addrs=None, cc_addrs=None, subject=None,
                    body_text=None, body_html=None, ticket_numbers=None):
    """One chase that actually went out. Returns email_followups.id.

    Called only after SMTP accepted, so the table records what was sent rather
    than what was attempted. Each call appends; nothing collapses two sends into
    one row."""
    if email_id is None or not sent_at:
        return None
    nums = [str(n) for n in (ticket_numbers or [])]
    with tx() as cur:
        if not cur:
            return None
        cur.execute(
            "INSERT INTO email_followups (email_id, round, kind, sent_at, "
            "  to_addrs, cc_addrs, subject, body_text, body_html, "
            "  ticket_numbers) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id",
            (email_id, round, kind, sent_at, to_addrs, cc_addrs, subject,
             body_text, body_html, nums))
        row = cur.fetchone()
        followup_id = row[0] if row else None
        for num in nums:
            _event(cur, num, "followup_sent", sent_at, email_id=email_id,
                   followup_id=followup_id,
                   detail={"round": round, "kind": kind})
        return followup_id
    return None


# -------------------------------------------------------- outbound mail
def record_email(message_id, subject, sent_at, ticket_map, columns=None,
                 kind="escalation", run_id=None, from_addr=None, to_addrs=None,
                 cc_addrs=None, body_html=None, body_text=None):
    """One sent mail plus its per-ticket lines. Returns emails_sent.id.

    The mapping row and header are stored on the lines and the mail, because
    Mapping.xlsx is deleted when the run commits and a later chase has to
    rebuild the same table it escalated with."""
    with tx() as cur:
        if not cur:
            return None
        cur.execute(
            "INSERT INTO emails_sent (message_id, run_id, kind, subject, "
            "  from_addr, to_addrs, cc_addrs, body_html, body_text, "
            "  ticket_count, sent_at, status, mapping_columns) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,'awaiting_reply',%s) "
            "ON CONFLICT (message_id) DO UPDATE SET subject=EXCLUDED.subject "
            "RETURNING id",
            (message_id, run_id, kind, subject, from_addr, to_addrs, cc_addrs,
             body_html, body_text, len(ticket_map or {}), sent_at,
             json.dumps(list(columns or []))))
        row = cur.fetchone()
        email_id = row[0] if row else None
        if email_id is None:
            return None
        for num, meta in (ticket_map or {}).items():
            meta = meta or {}
            cur.execute(
                "INSERT INTO email_tickets (email_id, ticket_number, awb, "
                "  line_status, mapping_row) "
                "VALUES (%s,%s,%s,'awaiting_reply',%s) "
                "ON CONFLICT (email_id, ticket_number) DO UPDATE SET "
                "  mapping_row = coalesce(EXCLUDED.mapping_row, "
                "                         email_tickets.mapping_row)",
                (email_id, str(num), meta.get("awb"),
                 json.dumps(meta.get("row")) if meta.get("row") else None))
            # The first entry on this ticket's timeline: the escalation itself.
            _event(cur, num, "email_sent", sent_at, awb=meta.get("awb"),
                   email_id=email_id, run_id=run_id,
                   detail={"kind": kind, "subject": subject})
        return email_id
    return None


def email_id_for(message_id):
    with tx() as cur:
        if not cur:
            return None
        cur.execute("SELECT id FROM emails_sent WHERE message_id=%s",
                    (message_id,))
        row = cur.fetchone()
        return row[0] if row else None
    return None


# --------------------------------------------------------- inbound mail
def record_reply(email_id, message_id, in_reply_to=None, from_addr=None,
                 subject=None, received_at=None, body_html=None,
                 body_text=None, imap_uid=None, match_method=None,
                 status_rows_found=0):
    """One reply. message_id is the idempotency key; a repeat returns its id."""
    with tx() as cur:
        if not cur:
            return None
        cur.execute(
            "INSERT INTO email_replies (email_id, message_id, in_reply_to, "
            "  imap_uid, from_addr, subject, body_html, body_text, "
            "  received_at, match_method, status_rows_found) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
            "ON CONFLICT (message_id) DO UPDATE SET "
            "  status_rows_found = EXCLUDED.status_rows_found "
            "RETURNING id",
            (email_id, message_id, in_reply_to, imap_uid, from_addr, subject,
             body_html, body_text, received_at, match_method,
             int(status_rows_found or 0)))
        row = cur.fetchone()
        return row[0] if row else None
    return None


def record_status(reply_id, ticket_number, status_text, awb=None,
                  matched_via=None, desk_comment_id=None, post_error=None,
                  received_at=None, desk_posted_at=None):
    """Append one remark to the history. Never replaces an earlier one.

    received_at is when Bluedart's reply landed - the Date header, not when this
    run read it - and becomes the remark's place on the ticket's timeline. Without
    it the row is still stored; only the timeline entry is skipped, since an event
    with no time on it would sort arbitrarily."""
    if reply_id is None or not ticket_number:
        return None
    with tx() as cur:
        if not cur:
            return None
        cur.execute(
            "INSERT INTO reply_ticket_statuses (reply_id, ticket_number, awb, "
            "  status_text, matched_via, desk_comment_id, desk_posted_at, "
            "  post_error) "
            "VALUES (%s,%s,%s,%s,%s,%s,now(),%s) "
            "ON CONFLICT (reply_id, ticket_number) DO UPDATE SET "
            "  status_text=EXCLUDED.status_text, post_error=EXCLUDED.post_error "
            "RETURNING id",
            (reply_id, str(ticket_number), awb, status_text, matched_via,
             desk_comment_id, post_error))
        row = cur.fetchone()
        _event(cur, ticket_number, "status_recorded", received_at,
               status_text=status_text, awb=awb, reply_id=reply_id,
               detail={"matched_via": matched_via} if matched_via else None)
        if desk_comment_id:
            # A separate entry: the remark arriving and the comment reaching Desk
            # are two different moments, and the report asks for both.
            _event(cur, ticket_number, "desk_comment_posted",
                   desk_posted_at or received_at, status_text=status_text,
                   awb=awb, reply_id=reply_id, desk_comment_id=desk_comment_id)
        return row[0] if row else None
    return None


def apply_latest(email_id, ticket_number, status_text, reply_id=None,
                 is_final=False, final_reason=None, desk_comment_id=None,
                 desk_posted_at=None):
    """Cache the newest remark on the line, and finalise it if it is done.

    Finality is set once and never lifted: a later non-final remark updates the
    text but must not drag a delivered AWB back into the chase."""
    if email_id is None or not ticket_number:
        return
    with tx() as cur:
        if cur:
            cur.execute(
                "UPDATE email_tickets SET line_status='answered', "
                "  latest_status_text=%s, latest_reply_id=%s, answered_at=now(), "
                "  desk_comment_id = coalesce(%s, email_tickets.desk_comment_id), "
                "  desk_posted_at  = coalesce(%s, email_tickets.desk_posted_at), "
                "  is_final = email_tickets.is_final OR %s, "
                "  final_reason = coalesce(email_tickets.final_reason, %s), "
                "  finalised_at = CASE WHEN email_tickets.is_final THEN "
                "                      email_tickets.finalised_at "
                "                 WHEN %s THEN now() ELSE NULL END "
                "WHERE email_id=%s AND ticket_number=%s",
                (status_text, reply_id, desk_comment_id, desk_posted_at,
                 bool(is_final),
                 final_reason if is_final else None, bool(is_final),
                 email_id, str(ticket_number)))


class _Unset:
    def __repr__(self):
        return "<unset>"


# Defined above set_followup_state because it is that function's default. With
# None as the default there was no way to tell "leave this column alone" from
# "set this column to NULL", so every call wrote all six columns and a caller
# naming one of them silently erased the other five.
_UNSET = _Unset()


def set_followup_state(email_id, last_reply_at=_UNSET, next_followup_at=_UNSET,
                       followup_round=_UNSET, awaiting_status_since=_UNSET,
                       stalled_at=_UNSET, followup_sent_at=_UNSET):
    """The 15-hour cycle's state. Only the arguments given are written.

    Passing None is still a write - it clears that column, which is how a reply
    disarms next_followup_at. Omitting an argument leaves the column untouched."""
    if email_id is None:
        return
    sets, args = [], []
    for col, val in (("last_reply_at", last_reply_at),
                     ("next_followup_at", next_followup_at),
                     ("followup_round", followup_round),
                     ("awaiting_status_since", awaiting_status_since),
                     ("stalled_at", stalled_at),
                     ("followup_sent_at", followup_sent_at)):
        if val is not _UNSET:
            sets.append("{}=%s".format(col))
            args.append(val)
    if not sets:
        return
    args.append(email_id)
    with tx() as cur:
        if cur:
            cur.execute("UPDATE emails_sent SET {} WHERE id=%s"
                        .format(", ".join(sets)), tuple(args))


def rollup_email(email_id):
    """Set the mail's status from its lines. Returns the new status."""
    if email_id is None:
        return None
    with tx() as cur:
        if not cur:
            return None
        cur.execute(
            "UPDATE emails_sent e SET status = CASE "
            "    WHEN NOT EXISTS (SELECT 1 FROM email_tickets t "
            "                     WHERE t.email_id=e.id AND NOT t.is_final) "
            "    THEN 'completed' "
            "    WHEN EXISTS (SELECT 1 FROM email_tickets t "
            "                 WHERE t.email_id=e.id AND t.line_status='answered') "
            "    THEN 'partially_answered' ELSE 'awaiting_reply' END, "
            "  completed_at = CASE "
            "    WHEN NOT EXISTS (SELECT 1 FROM email_tickets t "
            "                     WHERE t.email_id=e.id AND NOT t.is_final) "
            "    THEN coalesce(e.completed_at, now()) ELSE NULL END "
            "WHERE e.id=%s RETURNING status", (email_id,))
        row = cur.fetchone()
        return row[0] if row else None
    return None
