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
    first_run_id      bigint REFERENCES runs(id),
    -- Delivered Date. Written once and never moved: a courtesy note or a second
    -- Delivered on a re-used AWB must not change when the shipment was actually
    -- delivered. is_delivered_status() in main.py decides what qualifies -
    -- "RTO Delivered" does, plain RTO / RTOInTransit does not.
    -- delivered_at is when Bluedart's reply said so; delivered_recorded_at is
    -- when this pipeline processed it. They differ whenever a run is delayed.
    delivered_at          timestamptz,
    delivered_status_text text,
    delivered_awb         text,
    -- No inline REFERENCES: email_replies is created further down this file.
    -- The constraint is added at the end, once both tables exist.
    delivered_reply_id    bigint,
    delivered_recorded_at timestamptz,
    -- The RTO equivalent, kept separately because the two are NOT the same
    -- outcome. A plain RTO stops the chase but delivers nothing, so it must
    -- never fill delivered_at; "RTO Delivered" fills delivered_at instead.
    -- Like the delivered columns, written once and never moved.
    rto_at                timestamptz,
    rto_status_text       text,
    rto_reply_id          bigint,
    rto_recorded_at       timestamptz
);
ALTER TABLE tickets ADD COLUMN IF NOT EXISTS delivered_at          timestamptz;
ALTER TABLE tickets ADD COLUMN IF NOT EXISTS delivered_status_text text;
ALTER TABLE tickets ADD COLUMN IF NOT EXISTS delivered_awb         text;
ALTER TABLE tickets ADD COLUMN IF NOT EXISTS delivered_reply_id    bigint;
ALTER TABLE tickets ADD COLUMN IF NOT EXISTS delivered_recorded_at timestamptz;
ALTER TABLE tickets ADD COLUMN IF NOT EXISTS rto_at                timestamptz;
ALTER TABLE tickets ADD COLUMN IF NOT EXISTS rto_status_text       text;
ALTER TABLE tickets ADD COLUMN IF NOT EXISTS rto_reply_id          bigint;
ALTER TABLE tickets ADD COLUMN IF NOT EXISTS rto_recorded_at       timestamptz;
CREATE INDEX IF NOT EXISTS tickets_delivered ON tickets (delivered_at DESC)
    WHERE delivered_at IS NOT NULL;
CREATE INDEX IF NOT EXISTS tickets_rto ON tickets (rto_at DESC)
    WHERE rto_at IS NOT NULL;
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
    -- The mapping table's header, kept so a chase can rebuild the exact table it
    -- escalated with after Mapping.xlsx has been cleaned up.
    mapping_columns            jsonb,
    -- The 15-hour cycle's state, mirrored from threads.json.
    last_reply_at              timestamptz,
    next_followup_at           timestamptz,
    followup_round             int NOT NULL DEFAULT 0,
    awaiting_status_since      timestamptz,
    stalled_at                 timestamptz,
    followup_sent_at           timestamptz,
    followup_suppressed        boolean NOT NULL DEFAULT false,
    followup_suppressed_reason text,
    completed_at               timestamptz
);
-- Databases created before these columns existed. db.py has always written them.
ALTER TABLE emails_sent ADD COLUMN IF NOT EXISTS mapping_columns       jsonb;
ALTER TABLE emails_sent ADD COLUMN IF NOT EXISTS last_reply_at         timestamptz;
ALTER TABLE emails_sent ADD COLUMN IF NOT EXISTS next_followup_at      timestamptz;
ALTER TABLE emails_sent ADD COLUMN IF NOT EXISTS followup_round        int NOT NULL DEFAULT 0;
ALTER TABLE emails_sent ADD COLUMN IF NOT EXISTS awaiting_status_since timestamptz;
ALTER TABLE emails_sent ADD COLUMN IF NOT EXISTS stalled_at            timestamptz;
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
    -- The line's own copy of the mapping row, for the same reason as
    -- emails_sent.mapping_columns.
    mapping_row        jsonb,
    -- Finality is set once and never lifted, so a later non-final remark cannot
    -- drag a delivered AWB back into the chase.
    is_final           boolean NOT NULL DEFAULT false,
    final_reason       text,
    finalised_at       timestamptz,
    PRIMARY KEY (email_id, ticket_number)
);
ALTER TABLE email_tickets ADD COLUMN IF NOT EXISTS email_sent_at timestamptz;
ALTER TABLE email_tickets ADD COLUMN IF NOT EXISTS mapping_row   jsonb;
ALTER TABLE email_tickets ADD COLUMN IF NOT EXISTS is_final      boolean NOT NULL DEFAULT false;
ALTER TABLE email_tickets ADD COLUMN IF NOT EXISTS final_reason  text;
ALTER TABLE email_tickets ADD COLUMN IF NOT EXISTS finalised_at  timestamptz;
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

-- ================================================================ email_followups   [history, append-only]
-- One row per chase actually sent, with its own timestamp.
--
-- threads.json could only ever hold the LATEST followup_sent_ist, so the third
-- chase overwrote the second and the per-chase history never existed anywhere.
--
-- Deliberately NO unique constraint. The reply-anchored chase and the 15:00
-- no-reply net can both fire against one thread and can carry the same round
-- number; a UNIQUE here would silently swallow one of them, which is the exact
-- class of loss this table exists to end.
CREATE TABLE IF NOT EXISTS email_followups (
    id             bigserial PRIMARY KEY,
    email_id       bigint NOT NULL REFERENCES emails_sent(id) ON DELETE CASCADE,
    round          int,
    kind           text CHECK (kind IN ('reply_anchored','no_reply')),
    sent_at        timestamptz NOT NULL,
    to_addrs       text[],
    cc_addrs       text[],
    subject        text,
    body_text      text,
    body_html      text,
    ticket_numbers text[],
    recorded_at    timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS followups_email ON email_followups (email_id, sent_at);

-- ================================================================ ticket_events   [timeline, append-only]
-- The whole life of a ticket in one chronological stream: the escalation going
-- out, every chase, every reply, every remark, every Desk comment, and the
-- delivery. Rows are only ever INSERTed - nothing here is updated or deleted,
-- so an earlier status can never be lost behind a later one.
--
--   SELECT occurred_at, event_type, status_text FROM ticket_events
--    WHERE ticket_number = '332244' ORDER BY occurred_at;
--
-- occurred_at is when the thing happened (the mail's Date header, the reply's
-- received time), NOT when this pipeline noticed it - recorded_at is that. A run
-- delayed by an outage therefore cannot distort the timeline.
CREATE TABLE IF NOT EXISTS ticket_events (
    id              bigserial PRIMARY KEY,
    ticket_number   text NOT NULL REFERENCES tickets(ticket_number),
    event_type      text NOT NULL CHECK (event_type IN
                      ('email_sent','followup_sent','reply_received',
                       'status_recorded','desk_comment_posted','delivered',
                       'rto')),
    occurred_at     timestamptz NOT NULL,
    status_text     text,
    awb             text,
    email_id        bigint REFERENCES emails_sent(id),
    reply_id        bigint REFERENCES email_replies(id),
    followup_id     bigint REFERENCES email_followups(id),
    run_id          bigint REFERENCES runs(id),
    desk_comment_id text,
    detail          jsonb,
    recorded_at     timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS ticket_events_timeline ON ticket_events (ticket_number, occurred_at);
CREATE INDEX IF NOT EXISTS ticket_events_type     ON ticket_events (event_type, occurred_at DESC);

-- tickets.delivered_reply_id could not carry its REFERENCES inline: tickets is
-- created before email_replies. Added here, once both exist, and guarded so a
-- re-run is still a no-op.
DO $fk$ BEGIN
    ALTER TABLE tickets ADD CONSTRAINT tickets_delivered_reply_fk
        FOREIGN KEY (delivered_reply_id) REFERENCES email_replies(id);
EXCEPTION WHEN duplicate_object THEN NULL;
END $fk$;

DO $fk2$ BEGIN
    ALTER TABLE tickets ADD CONSTRAINT tickets_rto_reply_fk
        FOREIGN KEY (rto_reply_id) REFERENCES email_replies(id);
EXCEPTION WHEN duplicate_object THEN NULL;
END $fk2$;

-- An existing ticket_events table was created before 'rto' was a valid type.
DO $ev$ BEGIN
    ALTER TABLE ticket_events DROP CONSTRAINT IF EXISTS ticket_events_event_type_check;
    ALTER TABLE ticket_events ADD CONSTRAINT ticket_events_event_type_check
        CHECK (event_type IN ('email_sent','followup_sent','reply_received',
                              'status_recorded','desk_comment_posted',
                              'delivered','rto'));
END $ev$;

-- ================================================================ ticket_journey
-- THE REPORTING SOURCE. One row per ticket, the whole journey flattened.
--
-- No table can be this on its own: ticket_events holds the complete history but
-- many rows per ticket, while tickets holds one row per ticket without the mail
-- and chase timings. This joins the two so a report can just SELECT * FROM it.
--
-- Everything except the outcome stamps is derived from ticket_events, which is
-- append-only, so this view can never disagree with the history it summarises.
CREATE OR REPLACE VIEW ticket_journey AS
SELECT
    t.ticket_number,
    t.awb,
    t.courier,
    t.logistics_class,
    t.states,
    t.created_utc                                   AS ticket_created_at,
    ev.first_email_sent_at,
    ev.first_reply_at,
    ev.last_reply_at,
    ev.followups_sent,
    ev.first_followup_at,
    ev.last_followup_at,
    ev.last_status_text,
    ev.desk_comments_posted,
    ev.last_desk_comment_at,
    t.delivered_at,
    t.delivered_status_text,
    t.rto_at,
    t.rto_status_text,
    CASE WHEN t.delivered_at IS NOT NULL THEN 'delivered'
         WHEN t.rto_at       IS NOT NULL THEN 'rto'
         WHEN ev.last_reply_at IS NOT NULL THEN 'answered, still open'
         WHEN ev.first_email_sent_at IS NOT NULL THEN 'awaiting first reply'
         ELSE 'not escalated' END                   AS outcome,
    -- How long the whole chase took, first escalation to the final answer.
    coalesce(t.delivered_at, t.rto_at) - ev.first_email_sent_at
                                                    AS time_to_outcome
FROM tickets t
LEFT JOIN LATERAL (
    SELECT
        min(occurred_at) FILTER (WHERE event_type = 'email_sent')
            AS first_email_sent_at,
        min(occurred_at) FILTER (WHERE event_type = 'status_recorded')
            AS first_reply_at,
        max(occurred_at) FILTER (WHERE event_type = 'status_recorded')
            AS last_reply_at,
        count(*) FILTER (WHERE event_type = 'followup_sent')
            AS followups_sent,
        min(occurred_at) FILTER (WHERE event_type = 'followup_sent')
            AS first_followup_at,
        max(occurred_at) FILTER (WHERE event_type = 'followup_sent')
            AS last_followup_at,
        count(*) FILTER (WHERE event_type = 'desk_comment_posted')
            AS desk_comments_posted,
        max(occurred_at) FILTER (WHERE event_type = 'desk_comment_posted')
            AS last_desk_comment_at,
        (array_agg(status_text ORDER BY occurred_at DESC, id DESC)
            FILTER (WHERE event_type = 'status_recorded'))[1]
            AS last_status_text
    FROM ticket_events te
    WHERE te.ticket_number = t.ticket_number
) ev ON true;

COMMIT;
