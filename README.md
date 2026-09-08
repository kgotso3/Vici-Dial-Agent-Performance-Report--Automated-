# ViciDial Hourly Agent Performance Detail Report

Automatically pulls the "Agent Performance Detail" data from ViciDial every
hour, builds it into an Excel file, and emails it — no server or manual
process required. It runs entirely on GitHub's free Actions infrastructure.

## How it works

- `vici_report.py` calls ViciDial's built-in API function `agent_stats_export`
  (this is the same data source behind the "Agent Performance Detail" report
  under Agent Reports in the admin panel), for a rolling window (default: the
  last hour).
- It converts the response into an `.xlsx` file with `pandas` + `openpyxl`.
- It emails that file as an attachment via SMTP.
- `.github/workflows/hourly_report.yml` runs this script every hour
  automatically using GitHub Actions' cron scheduler — GitHub's own servers
  run it, not yours, so nothing needs to stay on or be watched.

## One-time setup

### 1. Create a ViciDial API user
In ViciDial admin, create (or reuse) a user with:
- `user_level` of 7 or higher
- "View Reports" permission enabled
- API access enabled

### 2. Push this repo to GitHub
Push these files to a repository (public or private — a private repo works
fine, GitHub gives free Actions minutes each month; a public repo has
unlimited free minutes).

### 3. Add GitHub Secrets
In your repo: **Settings → Secrets and variables → Actions → New repository secret**.
Add each of the following:

| Secret name       | Example value                          |
|--------------------|-----------------------------------------|
| `VICI_SERVER`      | `https://vici.mycompany.com`            |
| `VICI_API_USER`    | `6666`                                  |
| `VICI_API_PASS`    | `yourpassword`                          |
| `SMTP_HOST`        | `smtp.gmail.com`                        |
| `SMTP_PORT`        | `587`                                   |
| `SMTP_USER`        | `you@gmail.com`                         |
| `SMTP_PASS`        | an app password (not your login password)|
| `EMAIL_FROM`       | `you@gmail.com`                         |
| `EMAIL_TO`         | `boss@company.com,team@company.com`     |

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

- **Column layout**: `agent_stats_export`'s output columns are set from
  ViciDial's documentation but can vary slightly by version. After your
  first test run, open the resulting Excel file and confirm the columns
  line up with what you see in the web "Agent Performance Detail" report.
  Adjust the `COLUMNS` list in `vici_report.py` if needed.
- **Time zone**: the script builds its time window in UTC. If your ViciDial
  server's local time differs, adjust the `dt_start`/`dt_end` calculation in
  `fetch_agent_performance()` accordingly, or the hourly window may be
  offset from what you expect.
- **Gmail/Office365 SMTP**: most providers require an "app password" rather
  than your normal login password when sending mail via SMTP from a script.
- **Free tier limits**: GitHub Actions gives 2,000 free minutes/month on
  private repos (unlimited on public repos). This job takes well under a
  minute per run, so an hourly schedule (~720 runs/month) uses a small
  fraction of that.
- **Scheduling precision**: GitHub's cron scheduler is not guaranteed to the
  minute — under heavy load, runs can be delayed by several minutes. If you
  need exact-hour precision, you'd need a self-hosted runner instead, but for
  a periodic business report this is normally not an issue.
