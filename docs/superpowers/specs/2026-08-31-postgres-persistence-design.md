# Postgres persistence for the Bluedart escalation automation

**Date:** 2026-08-31
**Status:** Design — approved decisions recorded below, implementation not started
**Diagram:** https://claude.ai/code/artifact/c578ff05-a4f3-4cea-9076-1972cd9a7ba2

## 1. Purpose

Replace the three JSON state files with a PostgreSQL database, and persist the
ticket, courier and reply data the pipeline currently deletes after every run.

Today `main.py` keeps all durable state in three gitignored files next to the
script:

| File | Shape | Role |
|---|---|---|
| `state.json` | one object | watermark + in-flight `pending` block |
| `awb_registry.json` | `{awb: {ticketNumber, ticketId, created_utc}}` | AWB dedup, last-resort ticket lookup |
| `threads.json` | array of thread objects | sent mail awaiting reply, and what it already commented |

Everything else is transient. `cleanup_run_files()` deletes the Desk export, the
ClickPost export and `Mapping.xlsx` once a run commits, so the ticket rows, the
courier statuses and the content of the mail that went out all disappear. Rahul's
remark is posted to Zoho Desk as a comment and then dropped from our side entirely.

## 2. Decisions

Settled with the user before this document was written:

| Decision | Choice |
|---|---|
| Scope | Automation state **and** full ticket history |
| Engine | PostgreSQL 16 |
| Host | `209.38.120.154`, listening on localhost only |
| Where `main.py` runs | Already on that server — no remote DB access needed |
| Cutover | Import existing JSON, then run DB-only; keep the files as a rollback point |
| Retention | **Keep everything.** No purge job, no TTL |

Consequences of the retention choice, accepted knowingly:

- Growth is dominated by message bodies rather than row count. Postgres compresses
  text values over roughly 2 KB out of line automatically, so HTML mail bodies land
  on disk well below their nominal size.
- Backups permanently carry customer PII. `pg_dump` output must be written somewhere
  readable only by the automation user, and never to a location that syncs off the box.

## 3. Architecture

A new module `db.py` sits between `main.py` and Postgres. `main.py` keeps all of its
logic; only the call sites that load or save state change.

The boundary matters because `main.py` is 3,138 lines with persistence scattered
through it — `load_state` (:190), `load_registry` (:350), `load_threads` (:2332),
plus inline mutation inside `dedupe_by_awb`, `commit_run`, `thread_ticket_map` and
`process_replies`. Pulling persistence behind one module makes it independently
testable and stops `main.py` growing further.

Connection details come from `DATABASE_URL` in `.env` (`chmod 600`, matching the
other secrets). One new dependency: `psycopg[binary]`.

## 4. The grain problem, and how the schema answers it

The user's brief asked for `ticket_id` as a column on the email table, keyed into
the reply table. That models one email to one ticket, which is not what this
pipeline does:

- **One escalation email covers many tickets.** The mapping report is a row per AWB;
  `thread_ticket_map()` already records `{ticket_number: {...}}` per sent mail.
- **One reply carries many statuses.** `parse_status_table()` extracts a row per
  ticket from Rahul's HTML table.

A `ticket_id` column on `emails_sent` would force one row per ticket per email,
duplicating subject, body and recipients across every line. On `email_replies` it
would keep the first status and discard the rest.

Two junction tables carry the ticket grain instead:

- `email_tickets` — one row per ticket per email, holding the **latest** status.
- `reply_ticket_statuses` — one row per ticket per reply, **append-only**.

The requirement is met in full: the relationship is keyed on ticket number, the
latest status is cached for fast reads, and the history is complete.

## 5. Schema

Create in this order; the foreign keys depend on it.

```sql
CREATE TABLE runs (
    id               bigserial PRIMARY KEY,
    started_at       timestamptz NOT NULL,
    finished_at      timestamptz,
    status           text NOT NULL CHECK (status IN ('pending','committed','failed')),
    stage_reached    text,
    attempts         int NOT NULL DEFAULT 1,
    watermark_ticket text,
    watermark_utc    timestamptz,
    modified_utc     timestamptz,
    rows_count       int,
    mailer_sent      boolean NOT NULL DEFAULT false,
    error            text
);
-- At most one run in flight. state.json cannot express this.
CREATE UNIQUE INDEX runs_one_pending ON runs (status) WHERE status = 'pending';
CREATE INDEX runs_committed_at ON runs (finished_at DESC) WHERE status = 'committed';

CREATE TABLE tickets (
    ticket_number     text PRIMARY KEY,
    ticket_id         text UNIQUE,
    created_utc       timestamptz NOT NULL,
    modified_utc      timestamptz,
    subject           text,
    status            text,
    status_type       text,
    owner             text,
    department        text,
    priority          text,
    awb               text,
    courier           text,
    logistics_class   text,
    vinculum_edd      date,
    states            text,
    pending_with_dept text,
    custom_fields     jsonb,
    first_seen_at     timestamptz NOT NULL DEFAULT now(),
    last_seen_at      timestamptz NOT NULL DEFAULT now(),
    first_run_id      bigint REFERENCES runs(id)
);
CREATE INDEX tickets_awb      ON tickets (awb);
CREATE INDEX tickets_created  ON tickets (created_utc);
CREATE INDEX tickets_modified ON tickets (modified_utc);

CREATE TABLE awb_registry (
    awb           text PRIMARY KEY,
    ticket_number text NOT NULL REFERENCES tickets(ticket_number),
    created_utc   timestamptz NOT NULL,
    registered_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE clickpost_statuses (
    id         bigserial PRIMARY KEY,
    run_id     bigint NOT NULL REFERENCES runs(id),
    awb        text NOT NULL,
    status     text,
    edd        date,
    raw        jsonb,
    fetched_at timestamptz NOT NULL DEFAULT now(),
    UNIQUE (run_id, awb)
);
CREATE INDEX clickpost_awb ON clickpost_statuses (awb);

CREATE TABLE emails_sent (
    id               bigserial PRIMARY KEY,
    message_id       text NOT NULL UNIQUE,
    run_id           bigint REFERENCES runs(id),
    kind             text NOT NULL CHECK (kind IN ('escalation','followup')),
    subject          text NOT NULL,
    from_addr        text NOT NULL,
    to_addrs         text[] NOT NULL,
    cc_addrs         text[],
    body_html        text,
    body_text        text,
    ticket_count     int NOT NULL DEFAULT 0,
    sent_at          timestamptz NOT NULL,
    status           text NOT NULL DEFAULT 'awaiting_reply'
                     CHECK (status IN ('awaiting_reply','partially_answered','completed')),
    followup_sent_at timestamptz,
    completed_at     timestamptz
);
CREATE INDEX emails_open ON emails_sent (status) WHERE status <> 'completed';
CREATE INDEX emails_sent_at ON emails_sent (sent_at DESC);

CREATE TABLE email_replies (
    id                bigserial PRIMARY KEY,
    email_id          bigint REFERENCES emails_sent(id),
    message_id        text NOT NULL UNIQUE,   -- idempotency key
    in_reply_to       text,
    imap_uid          bigint,
    from_addr         text NOT NULL,
    to_addrs          text[],
    subject           text,
    body_html         text,
    body_text         text,
    received_at       timestamptz,
    fetched_at        timestamptz NOT NULL DEFAULT now(),
    match_method      text CHECK (match_method IN ('in_reply_to','references','subject_and_table')),
    status_rows_found int NOT NULL DEFAULT 0
);
CREATE INDEX replies_email ON email_replies (email_id);
CREATE INDEX replies_uid   ON email_replies (imap_uid DESC);

CREATE TABLE email_tickets (
    email_id           bigint NOT NULL REFERENCES emails_sent(id) ON DELETE CASCADE,
    ticket_number      text   NOT NULL REFERENCES tickets(ticket_number),
    awb                text,
    line_status        text NOT NULL DEFAULT 'awaiting_reply'
                       CHECK (line_status IN ('awaiting_reply','answered')),
    latest_status_text text,
    latest_reply_id    bigint REFERENCES email_replies(id),
    answered_at        timestamptz,
    desk_comment_id    text,
    desk_posted_at     timestamptz,
    PRIMARY KEY (email_id, ticket_number)
);
CREATE INDEX et_ticket ON email_tickets (ticket_number);
CREATE INDEX et_open   ON email_tickets (line_status) WHERE line_status = 'awaiting_reply';

CREATE TABLE reply_ticket_statuses (
    id              bigserial PRIMARY KEY,
    reply_id        bigint NOT NULL REFERENCES email_replies(id) ON DELETE CASCADE,
    ticket_number   text NOT NULL REFERENCES tickets(ticket_number),
    awb             text,
    status_text     text NOT NULL,
    matched_via     text,
    desk_comment_id text,
    desk_posted_at  timestamptz,
    post_error      text,
    created_at      timestamptz NOT NULL DEFAULT now(),
    UNIQUE (reply_id, ticket_number)
);
CREATE INDEX rts_ticket_recent ON reply_ticket_statuses (ticket_number, created_at DESC);
```

### What the constraints replace

| Guard | Today | After |
|---|---|---|
| No duplicate comment on a ticket | `if key in book` against an in-memory dict | composite PK on `email_tickets` |
| No reprocessing the same reply | `seen_replies[]` scanned linearly | `UNIQUE (message_id)` on `email_replies` |
| No two runs in flight | nothing — a second run overwrites `pending` | partial unique index on `runs` |
| History survives an overwrite | later reply replaces the earlier remark | append to `reply_ticket_statuses` |
| Late replies not dropped | thread marked `completed` is skipped entirely | reply still matches and appends regardless of status |

## 6. Access layer — `db.py`

```python
# connection
def connect()                              # from DATABASE_URL
def tx()                                   # context manager, commit/rollback

# runs — replaces load_state / save_state / start_pending / mark_stage / commit_run
def current_watermark()                    -> dict | None
def pending_run()                          -> dict | None
def begin_run(watermark, modified_utc)     -> int
def mark_stage(run_id, stage)              -> None
def commit_run(run_id, ticket, created_utc, modified_utc, rows, mailed) -> None
def fail_run(run_id, error)                -> None

# tickets
def upsert_tickets(run_id, rows)           -> int
def ticket_id_for(ticket_number)           -> str | None

# awb registry — replaces load_registry / save_registry
def registry_lookup(awb)                   -> dict | None
def register_awb(awb, ticket_number, created_utc) -> None

# clickpost
def save_clickpost(run_id, rows)           -> int

# outbound mail — replaces record_thread
def record_email(run_id, kind, message_id, subject, from_addr,
                 to_addrs, cc_addrs, body_html, body_text, ticket_map) -> int
def open_emails()                          -> list[dict]
def followups_due(hour)                    -> list[dict]
def mark_followup_sent(email_id)           -> None

# inbound mail — replaces the threads.json half of process_replies
def max_imap_uid()                         -> int | None
def email_by_message_id(mid)               -> dict | None
def record_reply(**fields)                 -> int | None   # None if already seen
def record_statuses(reply_id, rows)        -> list[int]
def mark_posted(status_id, comment_id)     -> None
def mark_post_error(status_id, error)      -> None
def apply_latest(email_id, ticket_number, status_text, reply_id) -> None
def rollup_email(email_id)                 -> str          # new email status
```

### `main.py` call sites that change

| Function | Line | Change |
|---|---|---|
| `load_state` / `save_state` | 190, 196 | deleted; callers use `db.current_watermark()` / `db.pending_run()` |
| `start_pending` | 200 | `db.begin_run()` |
| `mark_stage` | 217 | `db.mark_stage()` |
| `commit_run` | 264 | `db.commit_run()` |
| `load_registry` / `save_registry` | 350, 356 | deleted |
| `dedupe_by_awb` | 361 | takes a lookup callable instead of a dict |
| `reported_tickets` | 557 | query instead of file read |
| `run_desk_fetch` | 659 | add `db.upsert_tickets()` |
| `run_merge` | 1914 | add `db.save_clickpost()` |
| `send_mapping_mail` | 2191 | returns the parts needed by `record_email` |
| `thread_ticket_map` | 2276 | uses `db.ticket_id_for()`; `resolve_ticket_id` becomes a fallback only |
| `record_thread` | 2316 | `db.record_email()` |
| `load_threads` / `save_threads` | 2332, 2338 | deleted |
| `find_thread_replies` | 2485 | `SEARCH UID <max_imap_uid+1>:*` instead of `SINCE <date>` |
| `process_replies` | 2664 | writes through `db.*`; in-memory dedup guards removed |
| `followup_due` / `run_followups` | 2653, 2778 | `db.followups_due()` |

## 7. Migration

A one-shot `migrate_json_to_db.py`, idempotent, `--dry` by default.

1. Read `state.json`, `awb_registry.json`, `threads.json` from the server's
   `/opt/zoho-desk`.
2. Synthesise one `runs` row per distinct historical run that can be inferred —
   at minimum a single `committed` row carrying the current watermark, so the next
   run does not re-sweep history. Carry a live `pending` block across if one exists.
3. Insert `tickets` stubs for every ticket number referenced by the registry or by
   a thread. Most fields will be null: the ticket rows were never stored. The
   `ticket_id` from the registry is preserved, which is the field that matters —
   it is what the comment API needs.
4. Insert `awb_registry` rows.
5. Insert `emails_sent` from each thread (`message_id`, `subject`, `sent_ist`,
   `status`, `followup_sent_ist`). Body and recipients are unavailable for historical
   threads and will be null; `MAIL_TO` from `.env` is **not** backfilled as a guess.
6. Insert `email_tickets` from `thread["tickets"]`, and set the latest-status fields
   from `thread["processed"]` where present.
7. Insert a synthetic `email_replies` row per entry in `seen_replies` so the
   idempotency key is populated and already-processed replies are not re-commented.
   These rows carry `message_id` only; body and headers are gone.
8. Print a reconciliation table: counts read vs counts written, per file.

Run `--dry` first, compare, then run for real.

## 8. Cutover and rollback

1. Install PostgreSQL 16, create role and database, bind to localhost.
2. Apply the schema.
3. Stop `zoho-desk-watch.service`.
4. Copy `state.json`, `awb_registry.json`, `threads.json` to `*.pre-db.bak`.
5. Run the migration `--dry`, review, then for real.
6. Deploy the `db.py` branch, `pip install psycopg[binary]`.
7. `main.py --process-replies --dry` — should report the same outstanding work as
   before the cutover. This is the check that the import was faithful.
8. Start the watcher, watch `journalctl -u zoho-desk-watch -f` through one reply.

**Rollback:** stop the service, `git checkout` the previous revision, restore the
`.bak` files. The JSON files are never deleted by this work, so rollback is always
available. Anything written to Postgres after cutover would need re-applying by hand
— acceptable, because the window between cutover and confidence is one reply cycle.

## 9. Backups

Nightly `pg_dump` via cron, owned by the automation user, mode 600, retained
locally. Given the keep-everything decision, the dump contains customer PII
permanently and must not be written to a synced or world-readable path.

## 10. Testing

TDD, following the existing `tests/` pattern (`test_failsafe.py`,
`test_late_edits.py` are the current examples).

- `db.py` unit tests against a throwaway schema in a test database.
- Constraint tests that assert the guards actually fire: a second `pending` run is
  rejected; the same `message_id` cannot be inserted twice; the same
  `(email_id, ticket_number)` cannot be commented twice.
- A migration test with fixture JSON files, asserting the reconciliation counts.
- The two existing failsafe tests must still pass unchanged.

## 11. Explicitly out of scope

- The Zoho Mail incoming webhook. `ZOHO_MAIL_INCOMING_WEBHOOK` is in `.env` but read
  nowhere in `main.py`. This schema is designed to support it — concurrent writers are
  why the guards are constraints rather than in-memory checks — but wiring it is
  separate work.
- Collapsing `awb_registry` into a `DISTINCT ON (awb)` view over `tickets`. Correct,
  and it would delete a chunk of `dedupe_by_awb`, but not during a migration.
- Any reporting UI over the new history.

## 12. Open

- **SSH access to `209.38.120.154`.** No key or password supplied yet. Nothing has
  been provisioned; PostgreSQL is not installed and no table exists.
- **The server's Python version.** `--watch` uses `imaplib.IMAP4.idle()`, which needs
  3.14+. The local Windows box has 3.14.3, but the watcher runs on the server, and a
  stock Ubuntu 24.04 image ships 3.12 — in which case the watcher is on the 60-second
  polling fallback. To be confirmed on first login; it does not affect this design.
