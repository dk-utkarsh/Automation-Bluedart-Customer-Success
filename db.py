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
                  matched_via=None, desk_comment_id=None, post_error=None):
    """Append one remark to the history. Never replaces an earlier one."""
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
        return row[0] if row else None
    return None


def apply_latest(email_id, ticket_number, status_text, reply_id=None,
                 is_final=False, final_reason=None):
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
                "  is_final = email_tickets.is_final OR %s, "
                "  final_reason = coalesce(email_tickets.final_reason, %s), "
                "  finalised_at = CASE WHEN email_tickets.is_final THEN "
                "                      email_tickets.finalised_at "
                "                 WHEN %s THEN now() ELSE NULL END "
                "WHERE email_id=%s AND ticket_number=%s",
                (status_text, reply_id, bool(is_final),
                 final_reason if is_final else None, bool(is_final),
                 email_id, str(ticket_number)))


def set_followup_state(email_id, last_reply_at=None, next_followup_at=None,
                       followup_round=None, awaiting_status_since=None,
                       stalled_at=None, followup_sent_at=None):
    """The 15-hour cycle's state. Only the arguments given are written."""
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


class _Unset:
    def __repr__(self):
        return "<unset>"


_UNSET = _Unset()


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
