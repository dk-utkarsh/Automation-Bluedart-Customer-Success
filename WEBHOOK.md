# Zoho Mail outgoing webhook — instant reply pickup

Today a reply is found by IMAP: `--watch` idles on the mailbox (Python 3.14+) or polls
every 60s below that. This adds a push channel so a reply is picked up the moment it
lands, without waiting for a poll.

## Read this first: the two webhook types are not interchangeable

Zoho names webhooks from **Zoho's** point of view, which is the opposite of what the
mail is doing:

| Zoho's name | Direction | What it does |
|---|---|---|
| **Incoming** webhook | you → Zoho | You POST to it and Zoho **creates** an email in the mailbox |
| **Outgoing** webhook | Zoho → you | Mail arrives, Zoho **POSTs to your server** ← **this is the one we want** |

`ZOHO_MAIL_INCOMING_WEBHOOK` in `.env` is the **first** kind. It cannot deliver a reply
to this server — it is an ingestion endpoint, and anyone holding that URL can inject
mail into the mailbox. It is unused by any code (`grep INCOMING_WEBHOOK main.py` → 0).
Rotate or delete it.

## Why the receiver ignores the payload

Zoho's Mail webhook payload is roughly:

```
summary, sentDateInGMT, subject, messageId, toAddress, folderId,
zuid, ccAddress, size, sender, receivedTime, fromAddress, html, IntegIdList
```

There is **no `Message-ID`, no `In-Reply-To`, no `References`** — `messageId` is Zoho's
internal numeric id, not the RFC header. Those three headers are exactly how
`find_thread_replies()` binds a reply to the report it answers. Matching on the payload
would force every reply onto the ambiguous subject-only path — the same path that
caused one reply to be commented onto three different reports on 27-Aug.

So `webhook_server.py` treats a POST as a **doorbell**: it means "something arrived, go
look", and the mailbox is still read over IMAP where the real headers are. That means:

- the undocumented payload schema cannot break matching, because it is never parsed
- a POST Zoho never sends, or one that lands while the service is restarting, costs
  nothing — the existing IMAP loop re-reads the same mailbox anyway
- duplicate POSTs coalesce into a single scan

**IMAP is not removed and cannot be** — `clickpost_login()` reads the ClickPost OTP over
the same `imap_connect()`. Drop IMAP and Part 2 of the daily pipeline stops working.

## Prerequisite: a domain name

**A bare IP will not work.** Let's Encrypt does not issue certificates for IP
addresses, and Zoho needs an HTTPS URL it can validate.

Point a hostname at the server:

```
cs.dentalkart.com.   A   209.38.120.154
```

No domain available? Use a Cloudflare Tunnel instead — it gives you an HTTPS hostname
with no inbound port open at all, which is the safer option for a box holding customer
PII:

```bash
cloudflared tunnel --url http://127.0.0.1:8787
```

## Install

```bash
ssh root@209.38.120.154
cd /home/ubuntu/scripts/utkarsh/Automation-Bluedart-Customer-Success
git pull                      # webhook_server.py + deploy/ files

sudo apt install -y nginx certbot python3-certbot-nginx

sudo cp deploy/nginx-zoho-webhook.conf /etc/nginx/sites-available/zoho-webhook
sudo sed -i 's/cs.dentalkart.com/cs.dentalkart.com/g' \
     /etc/nginx/sites-available/zoho-webhook
sudo ln -s /etc/nginx/sites-available/zoho-webhook /etc/nginx/sites-enabled/
sudo certbot --nginx -d cs.dentalkart.com
sudo nginx -t && sudo systemctl reload nginx

sudo cp deploy/zoho-desk-webhook.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now zoho-desk-webhook
sudo systemctl status zoho-desk-webhook
```

Check it answers before telling Zoho about it — Zoho requires a 200 when the webhook is
saved and will refuse the configuration otherwise:

```bash
curl -s https://cs.dentalkart.com/healthz
# {"ok": true, "scans": 1, "last_run_ist": "...", "last_error": null}
```

## Configure the Zoho side

There is no API for this; it must be done in the UI.

1. Zoho Mail → **Settings → Integrations → Developer Space → Outgoing Webhooks**
2. **Add new configuration**
3. Entity: **Mail**
4. URL: `https://cs.dentalkart.com/zoho-mail-hook`
5. Add filter conditions so it fires only on what matters — e.g. *To* contains
   `akshay.s@dentalkart.com`, or *Subject* contains `Customer Escalation`. Without a
   filter every message in the mailbox wakes a scan.
6. Leave **Limited Data List** *unchecked*. Checking it strips the payload down to
   Subject/From/To — harmless for this design, but it also tends to indicate a
   narrower event set.
7. Save.

On the very first request Zoho sends an `x-hook-secret` header. The server stores it in
`.webhook-secret` (mode 600) and echoes it back to confirm the subscription. **That
header is sent once only** — if it is lost, delete the webhook and recreate it. From
then on every POST is verified by HMAC-SHA256 and unsigned requests are rejected.

## Verify end to end

```bash
sudo journalctl -u zoho-desk-webhook -f
```

Reply to an escalation mail. Expect:

```
handshake accepted from 136.143.x.x          (first request only)
scanning (coalesced 1 event(s))
  reply from ... - 1 status row(s)
    #267321 <- 'Delivered today'  (AWB 80144885251)
scan complete: 1 comment(s) added
```

## Operating notes

- **Keep `zoho-desk-watch.service` running too.** The webhook is a latency
  optimisation, not a replacement. Zoho publishes no retry policy, no ordering
  guarantee and no replay API — a missed POST is gone forever, whereas `--watch`
  re-reads the mailbox on every loop and self-heals after any outage.
- **Zoho silently auto-disables** an endpoint that stops answering "for an extended
  period", with no notification. `--watch` running alongside is what stops that being
  an outage. Watch `scans` in `/healthz` going stale as the signal.
- **Follow-ups are deliberately not triggered here.** `run_followups()` sends mail to
  `MAIL_TO`; a mail-triggered loop that can itself send mail is a foot-gun. Leave those
  to `--followups` on a timer.
- **Config** (`.env`, all optional — the code has defaults):
  `WEBHOOK_BIND`, `WEBHOOK_PORT`, `WEBHOOK_PATH`, `WEBHOOK_DEBOUNCE`, `WEBHOOK_SECRET`.
- **Same host, one state file.** `threads.json`, `state.json` and `awb_registry.json`
  are machine-local and gitignored. Whichever host sends the report must be the host
  that watches for the reply — split them and replies match nothing.
