# Deploying on Linux

The pipeline runs unchanged on Linux and Windows — every platform difference is
behind `IS_WINDOWS` in `main.py`. This is the server side.

## 1. Python

**Python 3.14 or newer is strongly preferred.**

`--watch` uses `imaplib.IMAP4.idle()`, which only entered the standard library in
Python 3.14. On an older interpreter the watcher still runs — it detects the missing
method and falls back to polling every 60s — but replies are then picked up on the
poll rather than the moment they land.

Ubuntu 24.04 ships 3.12 and 22.04 ships 3.10, so on a stock image you get the polling
fallback unless you install 3.14 (deadsnakes PPA, pyenv, or a source build).

Check which you have:

```bash
python3 -c "import imaplib; print('IDLE:', hasattr(imaplib.IMAP4, 'idle'))"
```

The four-part daily pipeline and `--process-replies` work identically on any
supported Python — this only affects `--watch`.

## 2. System packages

```bash
sudo apt update
sudo apt install -y python3-venv google-chrome-stable procps
```

- **`google-chrome-stable`** — Chromium also works, but chrome-for-testing does not
  publish a matching chromedriver for every Chromium build, so Chrome is the safe pick.
  No X server is needed; the pipeline is headless by default.
- **`procps`** — provides `pgrep`, which `close_stale_chrome()` uses to reap a crashed
  run's Chrome before it locks the profile. Without it the pipeline still runs but
  prints a warning and cannot self-heal a stuck profile.

If Chrome is somewhere unusual, set `CHROME_BINARY` in `.env`; otherwise it is found
automatically.

## 3. Application

```bash
sudo useradd --system --create-home --home-dir /opt/zoho-desk zohodesk
sudo -u zohodesk -H bash
cd /opt/zoho-desk
git clone <your-remote> .
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## 4. Configuration

```bash
cp .env.example .env
chmod 600 .env
$EDITOR .env
```

Then carry over the runtime state from the Windows box — **these are gitignored, so a
fresh clone has none of them**:

| File | Why it matters if missing |
|---|---|
| `state.json` | the watermark; without it the first run re-scans and re-mails history |
| `awb_registry.json` | the AWB dedup registry; without it duplicates are re-reported |
| `threads.json` | sent threads; without it replies match nothing and follow-ups never fire |

```bash
chmod 600 state.json awb_registry.json threads.json
```

## 5. First run

Run it by hand once. ClickPost needs an OTP login, which works headless because the
code reads the code over IMAP rather than from the screen:

```bash
sudo -u zohodesk /opt/zoho-desk/.venv/bin/python main.py --today --keep-files
```

The Chrome profile in `.chrome-profile/` keeps that session, so later runs usually
skip the OTP entirely.

Useful checks that touch nothing:

```bash
.venv/bin/python main.py --selftest              # OTP parser
.venv/bin/python main.py --process-replies --dry # shows what would be commented
```

## 6. Scheduling

**Daily report, 17:30 IST.** Cron runs in the server's timezone, so either set the box
to `Asia/Kolkata` (`sudo timedatectl set-timezone Asia/Kolkata`) and use the first
line, or leave it on UTC and use the second:

```cron
# Asia/Kolkata
30 17 * * * cd /opt/zoho-desk && .venv/bin/python -u main.py >> /var/log/zoho-desk/daily.log 2>&1

# UTC (17:30 IST == 12:00 UTC)
0 12 * * * cd /opt/zoho-desk && .venv/bin/python -u main.py >> /var/log/zoho-desk/daily.log 2>&1
```

`main.py` decides its own window and is safe to re-run: a repeated day replaces that
day's report rather than duplicating it, and the AWB registry stops a ticket being
chased twice.

**Reply watcher.** Run this as a service rather than from cron — it is a long-lived
loop, not a periodic job. It also sends both kinds of follow-up, so no extra cron
entry is needed: each wake (IDLE caps below 29 minutes) re-checks what is due.

Two chase rules, and a thread is only ever eligible for one of them:

| Thread | Rule |
|---|---|
| Bluedart has never replied | one nudge at 15:00 IST the day after the report |
| Bluedart replied, but some AWBs are not yet delivered | chased every 15h on the same trail, from when each reply landed, up to 10 times |

An AWB leaves the chase when its remark is final — "delivered" anywhere in the text
(so `RTO Delivered` counts), except where it is negated: `Undelivered`,
`Not delivered` and `Could not be delivered` all stay pending. Set
`"followup_suppressed": true` on a thread in `threads.json` to stop both rules for
it by hand.

```bash
sudo cp deploy/zoho-desk-watch.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now zoho-desk-watch
journalctl -u zoho-desk-watch -f
```

**Note before first start:** any thread still `awaiting_reply` whose follow-up is
already overdue fires the moment the watcher comes up — the rule is `now >= due`, with
no upper bound. This now applies to the 15-hour chase too, so a thread carrying an old
reply can send immediately. Check `threads.json` first if that would send unwanted
mail, and suppress anything you do not want chased.

## 7. Log directory

```bash
sudo mkdir -p /var/log/zoho-desk && sudo chown zohodesk:zohodesk /var/log/zoho-desk
```

Add logrotate if the daily log matters to you; the pipeline does not rotate it.

## Notes

- **A successful run leaves `output/` empty.** Once the mail is accepted and the
  watermark commits, the run deletes its own working files. That is intended — pass
  `--keep-files` to inspect them.
- **Timezone is safe.** Every timestamp goes through an explicit IST constant, so the
  17:30 window and the 15:00 follow-up are correct regardless of the server clock's
  zone. Only cron needs to know about it.
- **`output/` holds customer PII** and is gitignored. Keep it that way.
