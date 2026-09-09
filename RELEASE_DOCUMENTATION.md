# Agent Performance Detail — Hourly Automated Report

**Release Documentation v1.0**
**Project:** ViciDial Agent Performance Detail — Automated Hourly Email Report
**Owner:** Kgotso Maila (kgotso.maila@cartrack.com)
**Last updated:** 2026-09-09

---

## 1. What this project does

Every hour, this project automatically:

1. Logs into Cartrack's ViciDial reporting page
2. Downloads the "Agent Performance Detail" report, covering the last hour, across **all campaigns and agents**
3. Converts it into a formatted Excel (`.xlsx`) file
4. Emails that file as an attachment to a chosen list of recipients

It does this **without any server, computer, or person needing to be running or watching it**. It runs entirely on GitHub's own infrastructure using a feature called **GitHub Actions**, on a repeating schedule.

### 1.1 Why this approach

The stated constraint for this project was: *"I can only host the script on my personal GitHub, but I need it automated so I don't have to sit and monitor it."*

GitHub Actions solves this directly — it's a free, built-in feature of every GitHub repository that can run code on a timer ("cron schedule"), on GitHub's own cloud servers. Nothing needs to run on Kgotso's laptop, and no separate server needs to be paid for or maintained.

---

## 2. High-level architecture

```
┌─────────────────────────┐
│   GitHub Actions cron    │   Fires automatically every hour
│   (runs on GitHub's own  │   (no local machine involved)
│    servers, not yours)   │
└────────────┬─────────────┘
             │ triggers
             ▼
┌─────────────────────────┐
│  vici_report.py          │   1. Authenticates to ViciDial
│  (Python script)         │   2. Downloads the report (CSV)
│                           │   3. Parses it into a table
│                           │   4. Builds an Excel file
│                           │   5. Sends it by email
└────────────┬─────────────┘
             │
    ┌────────┴────────┐
    ▼                 ▼
┌─────────┐      ┌───────────┐
│ ViciDial │      │  Outlook   │
│ (Cartrack│      │  (SMTP,    │
│  server) │      │  Cartrack  │
│          │      │  mailbox)  │
└─────────┘      └───────────┘
```

All credentials (ViciDial login, email login) are stored as **GitHub Secrets** — encrypted values attached to the repository that the script reads at runtime. They are never written into the code itself, so the code can safely live in a GitHub repository (even a public one) without exposing passwords.

---

## 3. Repository structure

```
vici-hourly-report/
├── vici_report.py                       # The main script — does all the work
├── requirements.txt                     # List of Python libraries the script needs
├── README.md                            # Quick setup instructions
├── RELEASE_DOCUMENTATION.md             # This document
├── .gitignore                           # Prevents accidentally committing local secrets
└── .github/
    └── workflows/
        └── hourly_report.yml            # Tells GitHub Actions WHEN and HOW to run the script
```

---

## 4. How GitHub Actions automation works

### 4.1 The workflow file

`.github/workflows/hourly_report.yml` is a configuration file, written in YAML, that GitHub reads automatically the moment it's present in a repository under that exact folder path. It defines:

- **`on: schedule`** — a cron expression (`5 * * * *`) telling GitHub "run this at 5 minutes past every hour, every day." Cron syntax reads as: `minute hour day-of-month month day-of-week`, where `*` means "every."
- **`on: workflow_dispatch`** — adds a manual "Run workflow" button in GitHub's UI, so the automation can be tested on demand instead of only waiting for the schedule.
- **`jobs: send-report`** — the actual sequence of steps GitHub runs each time:
  1. **Check out repository** — downloads a fresh copy of the repo's code onto a temporary virtual machine GitHub spins up just for this run.
  2. **Set up Python** — installs Python 3.12 on that temporary machine.
  3. **Install dependencies** — runs `pip install -r requirements.txt` to install the Python libraries the script needs (`requests`, `pandas`, `openpyxl`).
  4. **Run report script** — executes `python vici_report.py`, passing in all the secret credentials as environment variables.

After the job finishes, GitHub throws away that temporary virtual machine entirely. Nothing persists between runs except what's stored in GitHub Secrets and in the repository's code — which is exactly why the script has no memory of its own and re-authenticates fresh every single hour.

### 4.2 Why GitHub Secrets matter

Secrets (`Settings → Secrets and variables → Actions`) are encrypted at rest by GitHub and injected into the running job only as environment variables — they are never visible in logs (GitHub automatically masks them, replacing any accidental print of a secret value with `***`), never stored in the repository's code or history, and only accessible to workflows running in that specific repository.

This is what allows the automation to safely live on a **personal** GitHub account: the code itself contains zero passwords, so even if the repository were public, nobody could extract credentials from it.

### 4.3 Why the schedule runs at 5 minutes past the hour, not exactly on the hour

GitHub's cron scheduler is a shared resource used by millions of repositories. Requests scheduled for exactly `:00` experience the heaviest congestion and the largest delays. Offsetting to `:05` avoids the worst of that queue.

---

## 5. How the Python script works (`vici_report.py`)

The script is organized into four sequential steps, each implemented as its own function, called in order from `main()`.

### 5.1 Configuration loading

At the top of the script, `get_required_env()` reads each required credential from the environment (which GitHub Actions populates from Secrets). If any are missing, the script stops immediately with a clear error message naming exactly which one is missing, rather than failing confusingly later.

Optional settings (`REPORT_HOURS_BACK`, `VICI_TIMEZONE`) have sensible defaults but can be overridden without touching the code, by adding them to the workflow file's `env:` block.

### 5.2 Step 1 — `download_report_text()`

This function:

1. Calculates the report's time window: "now" minus `REPORT_HOURS_BACK` hours (default 1), in the `Africa/Johannesburg` timezone (configurable), so the report window lines up with actual South African business hours regardless of the fact that GitHub's servers run on UTC time internally.
2. Builds the exact same URL parameters that the browser sends when a person manually opens the "Agent Performance Detail" report and clicks Download — including `file_download=1`, the date/time range, and `--ALL--` for campaigns/groups/agents (so it always captures everyone, without needing to hardcode and maintain a list of campaign names).
3. Sends that request to ViciDial, authenticating with the site's login credentials.
4. Checks the response: if authentication was rejected, it raises a clear error rather than silently continuing.
5. Prints the first few lines of whatever came back into the GitHub Actions log — this is a deliberate debugging aid. If something ever goes wrong again, that log immediately shows what the server actually sent, instead of requiring guesswork.

### 5.3 Step 2 — `parse_report(text)`

Cartrack's "Agent Performance Detail" report has an unusual export format: rather than a normal CSV file, each row is written as **one single CSV field, wrapped in quotes, whose content is itself a second, ordinary comma-separated row.** This looks like:

```
"USER NAME,""ID"",""CURRENT USER GROUP"",..."
```

Reading this requires unwrapping it twice:

1. Parse the line as CSV → yields exactly one field, containing the "real" row as a string (with doubled `""` marks where the original had literal `"` characters).
2. Parse *that string* as CSV again → yields the actual list of column values.

The function `_unwrap_row()` performs exactly these two passes for every line. `parse_report()` uses it first on whichever line contains `"USER NAME"` (auto-detecting the header, rather than assuming a fixed line number, so the parser keeps working even if the report ever adds extra header/footer lines above the data), then on every following data line.

**Safety behavior:** if the header row can never be found *and* the response text looks like an HTML page (e.g. `<html>` appears in it), the function assumes authentication failed and raises an error — rather than silently treating a login page as "zero agents worked this hour" and emailing an empty spreadsheet. This was a real bug found and fixed during testing (see Section 8, Known Issues & Fixes).

### 5.4 Step 3 — `build_excel(df)`

Takes the parsed data (a `pandas` DataFrame — essentially an in-memory spreadsheet table) and writes it into a real `.xlsx` file, entirely in memory (no temporary files touch disk), using the `openpyxl` engine. It also auto-sizes each column's width based on its longest value, so the resulting spreadsheet is readable immediately without manual resizing.

### 5.5 Step 4 — `send_email(...)`

Builds a standard email message with:
- A subject line stating the exact time window covered
- A short plain-text body summarizing row count
- The Excel file attached, named with the report's start time for easy filing

It connects to the SMTP server (Outlook: `smtp.office365.com:587`), upgrades the connection to an encrypted TLS channel (`starttls()`), authenticates, and sends.

### 5.6 `main()`

Ties all four steps together in order, with `print()` statements at each stage so the GitHub Actions run log reads as a clear timeline of what happened, in what order, and how long each part took.

---

## 6. Authentication mechanisms explained

### 6.1 ViciDial side

The report page is protected by a login prompt. Two authentication styles look similar from a screenshot but are handled completely differently by code:

| Style | What it looks like | How the script must authenticate |
|---|---|---|
| **True HTTP Basic Auth** | Plain, unstyled OS-native popup (grey, no custom fonts/colors) | Send credentials directly in the request's `Authorization` header on every request — no session needed |
| **Form-based / SSO login** | Styled HTML page (custom colors, fonts, rounded corners) even if made to resemble a native dialog | POST the username/password to a login endpoint first, receive a session cookie, then include that cookie on the actual report request |

*As of this document's writing, this is the open item being resolved — see Section 8.*

### 6.2 Email (SMTP) side

Microsoft 365 (Outlook) mailboxes authenticate over SMTP using a username and password sent after a `STARTTLS` handshake upgrades the connection to encrypted. Two things commonly block this for corporate accounts:

- **Legacy SMTP AUTH may be disabled tenant-wide.** Microsoft disabled this by default across all Microsoft 365 tenants in 2022 as a security hardening measure; it must be explicitly re-enabled per mailbox by an admin.
- **App passwords require MFA to be enabled first**, and are themselves being phased out in favor of OAuth2 in many tenants.

If SMTP AUTH cannot be enabled by Cartrack IT, the long-term correct fix is to send mail via the **Microsoft Graph API** using OAuth2 (an Azure AD "app registration" with `Mail.Send` permission) instead of SMTP entirely. This requires IT/admin involvement to create the app registration, but is Microsoft's currently-supported path and isn't subject to the legacy-protocol restrictions.

---

## 7. Step-by-step implementation guide

### 7.1 Prerequisites
- A GitHub account (free tier is sufficient)
- ViciDial report login credentials
- A Cartrack Outlook mailbox to send from, and confirmation from IT that SMTP AUTH is enabled for it (or a plan to use Graph API instead)

### 7.2 Setup steps

1. **Create a GitHub repository** — private is recommended for a project touching internal systems, though either works.
2. **Upload the project files** listed in Section 3, preserving the exact folder path for the workflow file (`.github/workflows/hourly_report.yml`).
3. **Add GitHub Secrets** under Settings → Secrets and variables → Actions:

   | Secret | Value |
   |---|---|
   | `VICI_SERVER` | `https://vicidial-debt-web-ndf.cartrack.com` |
   | `VICI_API_USER` | Your ViciDial reporting username |
   | `VICI_API_PASS` | Your ViciDial reporting password |
   | `SMTP_HOST` | `smtp.office365.com` |
   | `SMTP_PORT` | `587` |
   | `SMTP_USER` | `kgotso.maila@cartrack.com` |
   | `SMTP_PASS` | Mailbox password or app password |
   | `EMAIL_FROM` | `kgotso.maila@cartrack.com` |
   | `EMAIL_TO` | Comma-separated recipient list |

4. **Test manually**: Actions tab → "Hourly Agent Performance Detail Report" → Run workflow. Watch the live log.
5. **Verify the email and data**: Confirm the Excel attachment arrives and its numbers match what you see on the live ViciDial report page for the same time window.
6. **Let the schedule take over**: once a manual test succeeds, no further action is needed — it runs hourly indefinitely.

### 7.3 Ongoing maintenance

- If Cartrack ever rotates the ViciDial or email password, update the corresponding GitHub Secret — no code change needed.
- If the report's column layout changes on ViciDial's side, the parser (Section 5.3) adapts automatically since it reads column names dynamically rather than assuming fixed positions — but it's worth spot-checking a report after any known ViciDial system upgrade.
- Past run history and logs are always available under the repository's **Actions** tab for auditing or troubleshooting.

---

## 8. Known issues & fix history

| Date | Issue | Root cause | Resolution |
|---|---|---|---|
| 2026-09-08 | `AssertionError: 1 columns passed, passed data had 3 columns` | Script assumed the wrong data source entirely (`agent_stats_export` API function) instead of the real custom report page | Rewrote to call the real report endpoint (`AST_agent_performance_detail.php`) directly |
| 2026-09-08 | Parser failed on real report format | Report is double CSV-encoded (a quoted field containing a second CSV row) — not simple pipe- or comma-delimited | Implemented `_unwrap_row()` two-pass CSV parsing, verified against real sample data |
| 2026-09-09 | Script silently emailed an empty report on auth failure | "No header row found" was treated as "no agents worked this hour" in all cases, including when the response was actually an HTML login page | Added a check: if no header is found *and* the response looks like an HTML page, the script now raises an error instead of continuing |
| 2026-09-09 | ViciDial login failing (HTML login page returned instead of report) | Login is not true HTTP Basic Auth as initially assumed — appears to be a styled/form-based login | **Open — pending network trace of the real login request from the browser** |
| 2026-09-09 | SMTP `535 Authentication unsuccessful` | Likely Microsoft 365's default-disabled legacy SMTP AUTH policy | **Open — pending confirmation from Cartrack IT**, or migration to Graph API/OAuth2 |

---

## 9. Security considerations

- All credentials are stored exclusively as encrypted GitHub Secrets, never in code or version history.
- The repository should be set to **Private** given it touches internal collections/legal data (agent performance across debt-collection campaigns).
- Consider creating a **dedicated, scoped ViciDial reporting account** for this automation (rather than reusing a personal admin login), so its access can be independently revoked or audited without affecting Kgotso's own account.
- Email recipients should be limited to people with a legitimate need to see agent performance data, given its contents include per-agent disposition and productivity metrics.

---

## 10. Glossary

- **GitHub Actions** — GitHub's built-in automation feature; runs code on GitHub's servers in response to events or schedules.
- **Cron schedule** — a standard syntax for expressing "run this at these recurring times."
- **GitHub Secret** — an encrypted value attached to a repository, injected into workflow runs as an environment variable.
- **SMTP** — the standard protocol used to send email.
- **HTTP Basic Auth** — a simple web authentication method sending a username/password with each request.
- **Session cookie** — a token a website gives a browser after login, proving it's already authenticated for further requests.
- **DataFrame** — an in-memory table structure (rows and columns) used by the `pandas` Python library.
