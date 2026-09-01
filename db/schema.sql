-- Bluedart escalation automation - persistence schema
-- Target: PostgreSQL 16, database bluedart_escalation, owner bluedart_app
-- Idempotent: safe to re-run.

BEGIN;

-- ================================================================ runs
-- Replaces state.json. The watermark is the newest committed row.
CREATE TABLE IF NOT EXISTS runs (
    id               bigserial PRIMARY KEY,
    started_at       timestamptz NOT NULL DEFAULT now(),
    finished_at      timestamptz,
    status           text NOT NULL CHECK (status IN ('pending','committed','failed')),
    stage_reached    text,
    attempts         int  NOT NULL DEFAULT 1,
    watermark_ticket text,
    watermark_utc    timestamptz,
    modified_utc     timestamptz,
    export_file      text,
    rows_count       int,
    mailer_sent      boolean NOT NULL DEFAULT false,
    error            text
);
-- At most one run in flight. state.json could not express this.
CREATE UNIQUE INDEX IF NOT EXISTS runs_one_pending ON runs (status) WHERE status = 'pending';
CREATE INDEX IF NOT EXISTS runs_committed ON runs (finished_at DESC) WHERE status = 'committed';

-- ================================================================ tickets
-- Mirror of the Desk ticket. ticket_number is the business key everything joins on.
CREATE TABLE IF NOT EXISTS tickets (
    ticket_number     text PRIMARY KEY,
    ticket_id         text UNIQUE,
    created_utc       timestamptz,
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
CREATE INDEX IF NOT EXISTS tickets_awb      ON tickets (awb);
CREATE INDEX IF NOT EXISTS tickets_created  ON tickets (created_utc);
CREATE INDEX IF NOT EXISTS tickets_modified ON tickets (modified_utc);

-- ================================================================ awb_registry
-- Replaces awb_registry.json. One keeper ticket per AWB; earliest-created wins.
CREATE TABLE IF NOT EXISTS awb_registry (
    awb           text PRIMARY KEY,
    ticket_number text NOT NULL REFERENCES tickets(ticket_number),
    created_utc   timestamptz NOT NULL,
    registered_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS awb_registry_ticket ON awb_registry (ticket_number);

-- ================================================================ clickpost_statuses
-- Courier snapshot per run per AWB. Today this is deleted after every run.
CREATE TABLE IF NOT EXISTS clickpost_statuses (
    id         bigserial PRIMARY KEY,
    run_id     bigint NOT NULL REFERENCES runs(id),
    awb        text NOT NULL,
    status     text,
    edd        date,
    raw        jsonb,
    fetched_at timestamptz NOT NULL DEFAULT now(),
    UNIQUE (run_id, awb)
);
CREATE INDEX IF NOT EXISTS clickpost_awb ON clickpost_statuses (awb);

-- ================================================================ emails_sent   [MAIN TABLE]
-- One row per outbound escalation or follow-up mail.
CREATE TABLE IF NOT EXISTS emails_sent (
    id                         bigserial PRIMARY KEY,
    message_id                 text NOT NULL UNIQUE,
    run_id                     bigint REFERENCES runs(id),
    kind                       text NOT NULL DEFAULT 'escalation'
                               CHECK (kind IN ('escalation','followup')),
    subject                    text NOT NULL,
    from_addr                  text,
    to_addrs                   text[],
    cc_addrs                   text[],
    body_html                  text,
    body_text                  text,
    ticket_count               int NOT NULL DEFAULT 0,
    sent_at                    timestamptz NOT NULL,
    status                     text NOT NULL DEFAULT 'awaiting_reply'
                               CHECK (status IN ('awaiting_reply','partially_answered','completed')),
    followup_sent_at           timestamptz,
    followup_suppressed        boolean NOT NULL DEFAULT false,
    followup_suppressed_reason text,
    completed_at               timestamptz
);
CREATE INDEX IF NOT EXISTS emails_open    ON emails_sent (status) WHERE status <> 'completed';
CREATE INDEX IF NOT EXISTS emails_sent_at ON emails_sent (sent_at DESC);

-- ================================================================ email_replies   [REPLY TABLE, append-only]
-- One row per inbound reply message. message_id is the idempotency key.
CREATE TABLE IF NOT EXISTS email_replies (
    id                bigserial PRIMARY KEY,
    email_id          bigint REFERENCES emails_sent(id),
    message_id        text NOT NULL UNIQUE,
    in_reply_to       text,
    imap_uid          bigint,
    from_addr         text,
    to_addrs          text[],
    subject           text,
    body_html         text,
    body_text         text,
    received_at       timestamptz,
    fetched_at        timestamptz NOT NULL DEFAULT now(),
    match_method      text CHECK (match_method IN ('in_reply_to','references','subject_and_table')),
    status_rows_found int NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS replies_email ON email_replies (email_id);
CREATE INDEX IF NOT EXISTS replies_uid   ON email_replies (imap_uid DESC);

-- ================================================================ email_tickets   [line items + latest status]
-- The composite primary key is what makes double-commenting impossible.
CREATE TABLE IF NOT EXISTS email_tickets (
    email_id           bigint NOT NULL REFERENCES emails_sent(id) ON DELETE CASCADE,
    ticket_number      text   NOT NULL REFERENCES tickets(ticket_number),
    awb                text,
    -- Copy of emails_sent.sent_at. Denormalised deliberately: a sent mail's
    -- timestamp never changes, so this cannot drift, and it saves a join on the
    -- one table people actually read.
    email_sent_at      timestamptz,
    line_status        text NOT NULL DEFAULT 'awaiting_reply'
                       CHECK (line_status IN ('awaiting_reply','answered')),
    latest_status_text text,
    latest_reply_id    bigint REFERENCES email_replies(id),
    answered_at        timestamptz,
    desk_comment_id    text,
    desk_posted_at     timestamptz,
    PRIMARY KEY (email_id, ticket_number)
);
ALTER TABLE email_tickets ADD COLUMN IF NOT EXISTS email_sent_at timestamptz;
CREATE INDEX IF NOT EXISTS et_ticket  ON email_tickets (ticket_number);
CREATE INDEX IF NOT EXISTS et_open    ON email_tickets (line_status) WHERE line_status = 'awaiting_reply';
CREATE INDEX IF NOT EXISTS et_sent_at ON email_tickets (email_sent_at DESC);

-- Fill email_sent_at on insert so a caller can never leave it NULL.
CREATE OR REPLACE FUNCTION email_tickets_fill_sent_at() RETURNS trigger
LANGUAGE plpgsql AS $fn$
BEGIN
    IF NEW.email_sent_at IS NULL THEN
        SELECT sent_at INTO NEW.email_sent_at FROM emails_sent WHERE id = NEW.email_id;
    END IF;
    RETURN NEW;
END
$fn$;

DROP TRIGGER IF EXISTS trg_email_tickets_sent_at ON email_tickets;
CREATE TRIGGER trg_email_tickets_sent_at
    BEFORE INSERT ON email_tickets
    FOR EACH ROW EXECUTE FUNCTION email_tickets_fill_sent_at();

-- ================================================================ reply_ticket_statuses   [history, append-only]
-- Every remark ever received. A later reply appends a row, it never replaces one.
CREATE TABLE IF NOT EXISTS reply_ticket_statuses (
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
CREATE INDEX IF NOT EXISTS rts_ticket_recent ON reply_ticket_statuses (ticket_number, created_at DESC);

COMMIT;
