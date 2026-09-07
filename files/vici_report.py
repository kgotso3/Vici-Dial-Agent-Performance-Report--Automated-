#!/usr/bin/env python3
"""
vici_report.py

Pulls agent performance data from ViciDial (the same data shown on the
"Agent Performance Detail" report under Agent Reports) via the built-in
non_agent_api.php 'agent_stats_export' function, saves it as an .xlsx
file, and emails it as an attachment.

Designed to run unattended on a schedule (see .github/workflows/hourly_report.yml).
All secrets are read from environment variables — never hardcode credentials.

Required environment variables:
    VICI_SERVER        e.g. "https://your-vicidial-server.example.com"
    VICI_API_USER       ViciDial user with 'view reports' permission + API access
    VICI_API_PASS
    SMTP_HOST           e.g. "smtp.gmail.com"
    SMTP_PORT           e.g. "587"
    SMTP_USER
    SMTP_PASS
    EMAIL_FROM
    EMAIL_TO             comma-separated list of recipients

Optional environment variables:
    REPORT_HOURS_BACK   how many hours back the report window should cover (default: 1)
    VICI_CAMPAIGN_ID    restrict to one campaign (optional)
    VICI_AGENT_USER     restrict to one agent (optional)
"""

import os
import sys
import io
import smtplib
import ssl
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage

import requests
import pandas as pd


# --------------------------------------------------------------------------
# Configuration (from environment / GitHub Actions secrets)
# --------------------------------------------------------------------------

def get_required_env(name: str) -> str:
    val = os.environ.get(name)
    if not val:
        print(f"ERROR: missing required environment variable: {name}", file=sys.stderr)
        sys.exit(1)
    return val


VICI_SERVER = get_required_env("VICI_SERVER").rstrip("/")
VICI_API_USER = get_required_env("VICI_API_USER")
VICI_API_PASS = get_required_env("VICI_API_PASS")

SMTP_HOST = get_required_env("SMTP_HOST")
SMTP_PORT = int(get_required_env("SMTP_PORT"))
SMTP_USER = get_required_env("SMTP_USER")
SMTP_PASS = get_required_env("SMTP_PASS")
EMAIL_FROM = get_required_env("EMAIL_FROM")
EMAIL_TO = [addr.strip() for addr in get_required_env("EMAIL_TO").split(",")]

REPORT_HOURS_BACK = int(os.environ.get("REPORT_HOURS_BACK", "1"))
VICI_CAMPAIGN_ID = os.environ.get("VICI_CAMPAIGN_ID", "")
VICI_AGENT_USER = os.environ.get("VICI_AGENT_USER", "")

# Columns returned by agent_stats_export, in order (per ViciDial's
# NON-AGENT_API.txt documentation). Adjust here if your ViciDial version's
# output differs (check the pipe-delimited output your server returns).
COLUMNS = [
    "user", "full_name", "user_group", "calls", "login_time",
    "total_talk_time", "avg_talk_time", "avg_wait_time", "pct_of_queue",
    "pause_time", "sessions", "avg_session", "pauses", "avg_pause_time",
    "pause_pct", "pauses_per_session", "wait_time", "talk_time",
    "dispo_time", "dead_time",
]


# --------------------------------------------------------------------------
# Step 1: Pull data from the ViciDial API
# --------------------------------------------------------------------------

def fetch_agent_performance():
    now = datetime.now(timezone.utc)
    start = now - timedelta(hours=REPORT_HOURS_BACK)

    # ViciDial expects "YYYY-MM-DD+HH:MM:SS" in the site's local server time.
    # If your ViciDial server is not on UTC, adjust here accordingly (e.g.
    # convert to the server's timezone before formatting).
    dt_start = start.strftime("%Y-%m-%d+%H:%M:%S")
    dt_end = now.strftime("%Y-%m-%d+%H:%M:%S")

    params = {
        "source": "hourly_report",
        "user": VICI_API_USER,
        "pass": VICI_API_PASS,
        "function": "agent_stats_export",
        "datetime_start": dt_start,
        "datetime_end": dt_end,
        "stage": "pipe",
        "header": "NO",
        "time_format": "M",  # minutes, easy to sort/format later
    }
    if VICI_CAMPAIGN_ID:
        params["campaign_id"] = VICI_CAMPAIGN_ID
    if VICI_AGENT_USER:
        params["agent_user"] = VICI_AGENT_USER

    url = f"{VICI_SERVER}/vicidial/non_agent_api.php"
    resp = requests.get(url, params=params, timeout=60)
    resp.raise_for_status()
    text = resp.text.strip()

    if text.startswith("ERROR"):
        raise RuntimeError(f"ViciDial API error: {text}")

    if not text:
        # No agent activity in that window is a valid (if boring) outcome.
        return pd.DataFrame(columns=COLUMNS), dt_start, dt_end

    rows = [line.split("|") for line in text.splitlines() if line.strip()]
    df = pd.DataFrame(rows, columns=COLUMNS[: len(rows[0])])
    return df, dt_start, dt_end


# --------------------------------------------------------------------------
# Step 2: Save to Excel
# --------------------------------------------------------------------------

def build_excel(df: pd.DataFrame, dt_start: str, dt_end: str) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Agent Performance Detail")
        ws = writer.sheets["Agent Performance Detail"]
        # Auto-width columns roughly based on content length
        for i, col in enumerate(df.columns, start=1):
            max_len = max([len(str(col))] + [len(str(v)) for v in df[col]]) if len(df) else len(str(col))
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(max_len + 2, 40)
    buffer.seek(0)
    return buffer.read()


# --------------------------------------------------------------------------
# Step 3: Email the file
# --------------------------------------------------------------------------

def send_email(xlsx_bytes: bytes, dt_start: str, dt_end: str, row_count: int):
    msg = EmailMessage()
    msg["Subject"] = f"Agent Performance Detail Report ({dt_start} to {dt_end})"
    msg["From"] = EMAIL_FROM
    msg["To"] = ", ".join(EMAIL_TO)
    msg.set_content(
        f"Attached: Agent Performance Detail report for {dt_start} to {dt_end}.\n"
        f"{row_count} agent record(s) included.\n\n"
        f"This is an automated message."
    )

    filename = f"agent_performance_detail_{dt_start.replace(':', '').replace('+', '_')}.xlsx"
    msg.add_attachment(
        xlsx_bytes,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )

    context = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls(context=context)
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    print(f"[{datetime.now(timezone.utc).isoformat()}] Fetching agent performance data...")
    df, dt_start, dt_end = fetch_agent_performance()
    print(f"Fetched {len(df)} row(s) for window {dt_start} -> {dt_end}")

    xlsx_bytes = build_excel(df, dt_start, dt_end)
    print("Excel file built, sending email...")

    send_email(xlsx_bytes, dt_start, dt_end, len(df))
    print("Email sent successfully.")


if __name__ == "__main__":
    main()
