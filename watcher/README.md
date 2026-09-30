# Mail watcher (GitHub Actions)

Checks every new email in the Gmail inbox every 30 minutes and saves the ones that matter for invoicing,
with their attachments, into `inbox/pending/`. It uses no Claude usage and costs nothing on GitHub's free tier for private repos.
It opens the mailbox **read-only** and never changes, moves or deletes any email.

## One-time setup

1. On the Google account **sunorientray@gmail.com**, turn on 2-Step Verification (if it isn't already), then create an
   **App password** at https://myaccount.google.com/apppasswords (name it "Mail watcher"). Copy the 16-character password.
2. In this repo on GitHub: **Settings → Secrets and variables → Actions → New repository secret**. Add both:
   - `GMAIL_USER` = `sunorientray@gmail.com`
   - `GMAIL_APP_PASSWORD` = the app password from step 1
3. **Actions** tab → **Mail watcher** → **Run workflow**. The first run looks back 3 days. It should finish green and print
   `checked N new emails, saved M for invoicing`.

To stop it: Actions → Mail watcher → "…" → Disable workflow. To revoke access: delete the app password in your Google account.

## What gets saved

See `rules.json`. The first rule that matches wins. IMS OTP emails and everything unrelated (shipping schedules, etc.) are
checked but not saved. Edit the rules there if a new kind of email should be picked up.
