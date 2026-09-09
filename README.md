# Cartrack Hourly Agent Performance Detail Report

Automatically pulls Cartrack's custom "Agent Performance Detail" report
from ViciDial every hour, builds it into an Excel file, and emails it — no
server or manual process required. It runs entirely on GitHub's free
Actions infrastructure.

## How it works

- `vici_report.py` requests the report directly from its real page:
  `AST_agent_performance_detail.php`, using `file_download=1` (the same
  request your browser makes when you click the Download button) — for a
  rolling window (default: the last hour).
- That page is protected by HTTP Basic Authentication (the plain
  username/password popup your browser shows), so the script sends those
  credentials with the request — no login form or session cookies needed.
- The report comes back as an unusually double-encoded CSV (each row is
  one big quoted field containing a second, normal comma-separated row).
  The script unwraps this automatically.
- It converts the result into an `.xlsx` file with `pandas` + `openpyxl`.
- It emails that file as an attachment via SMTP (configured for Outlook /
  Microsoft 365).
- `.github/workflows/hourly_report.yml` runs this script every hour
  automatically using GitHub Actions' cron scheduler — GitHub's own
  servers run it, not yours, so nothing needs to stay on or be watched.

## One-time setup

### 1. Confirm your ViciDial credentials
You already have these — they're the username and password from the
browser sign-in popup when you open:
`https://vicidial-debt-web-ndf.cartrack.com/vicidial/AST_agent_performance_detail.php`

### 2. Push this repo to GitHub
Push these files to a repository (private is recommended since this touches
company systems).

### 3. Add GitHub Secrets
In your repo: **Settings → Secrets and variables → Actions → New repository secret**.
Add each of the following:

| Secret name       | Example value                                    |
|--------------------|---------------------------------------------------|
| `VICI_SERVER`      | `https://vicidial-debt-web-ndf.cartrack.com`       |
| `VICI_API_USER`    | `9952` (your Basic Auth username)                  |
| `VICI_API_PASS`    | your Basic Auth password                           |
| `SMTP_HOST`        | `smtp.office365.com`                               |
| `SMTP_PORT`        | `587`                                              |
| `SMTP_USER`        | `kgotso.maila@cartrack.com`                        |
| `SMTP_PASS`        | your Outlook/Microsoft 365 app password (see note below) |
| `EMAIL_FROM`       | `kgotso.maila@cartrack.com`                        |
| `EMAIL_TO`         | `boss@cartrack.com,team@cartrack.com`              |

Never commit these values into the code — Secrets keep them encrypted and
out of your repo history, even though the repo itself might be public.

### 4. Test it
Go to the **Actions** tab → "Hourly Agent Performance Detail Report" →
**Run workflow** (this is the `workflow_dispatch` trigger) to fire it
manually and confirm you receive the email before waiting for the schedule.

### 5. Let it run
Once secrets are set, the `cron: "5 * * * *"` schedule fires automatically
every hour, all on GitHub's infrastructure. You don't need to keep any
device on or watch it.

## Notes & things to double check for your environment

- **Timezone**: the script builds its report window using the
  `Africa/Johannesburg` timezone by default (matching your ViciDial
  server), regardless of the fact that GitHub Actions itself runs in UTC.
  If your ViciDial server actually runs on a different timezone, set the
  optional `VICI_TIMEZONE` environment variable (see the workflow file) to
  the correct IANA timezone name.
- **All campaigns/groups included**: the script requests `--ALL--` for
  campaign groups, user groups, and agents, so you don't have to maintain
  a list of campaign names that will drift as campaigns are added or
  removed. If you ever want to restrict it to specific campaigns, that
  would need to be added back into the `params` dict in
  `download_report_text()`.
- **Outlook/Microsoft 365 SMTP**: since this is sending from a work
  account (`cartrack.com`), Cartrack's IT/Microsoft 365 admin may have
  **Basic Auth (SMTP AUTH) disabled by default** for security — common in
  corporate tenants. If your test run fails with an authentication error,
  ask IT to either enable SMTP AUTH for your mailbox, confirm an app
  password is allowed, or (if SMTP AUTH is blocked entirely) let your
  developer know you may need OAuth2 / Microsoft Graph API instead.
- **Free tier limits**: GitHub Actions gives 2,000 free minutes/month on
  private repos (unlimited on public repos). This job takes well under a
  minute per run, so an hourly schedule (~720 runs/month) uses a small
  fraction of that.
- **Scheduling precision**: GitHub's cron scheduler is not guaranteed to
  the minute — under heavy load, runs can be delayed by several minutes.
  For a periodic business report this is normally not an issue.
