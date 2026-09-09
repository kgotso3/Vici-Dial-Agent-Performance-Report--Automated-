#!/usr/bin/env python3
"""
vici_report.py

Pulls Cartrack's custom "Agent Performance Detail" report directly from
its ViciDial admin page:

    https://vicidial-debt-web-ndf.cartrack.com/vicidial/AST_agent_performance_detail.php

This page is protected by HTTP Basic Authentication (the native browser
sign-in popup), so the script authenticates the same way — sending a
username/password with the request, no login form or session cookies
needed.

The report itself is returned as an unusually-encoded CSV: each row is a
single outer-quoted CSV field, which itself contains a second, normal
comma-separated row. The script unwraps this twice to get clean data.

Once parsed, the data is saved as an .xlsx file and emailed as an
attachment. Designed to run unattended on a schedule (see
.github/workflows/hourly_report.yml). All secrets are read from
environment variables — never hardcode credentials.

Required environment variables:
    VICI_SERVER          e.g. "https://vicidial-debt-web-ndf.cartrack.com"
    VICI_API_USER        the HTTP Basic Auth username (same one used in the
                          browser sign-in popup, e.g. "9952")
    VICI_API_PASS         the matching HTTP Basic Auth password
    SMTP_HOST            e.g. "smtp.office365.com"
    SMTP_PORT            e.g. "587"
    SMTP_USER
    SMTP_PASS
    EMAIL_FROM
    EMAIL_TO              comma-separated list of recipients

Optional environment variables:
    REPORT_HOURS_BACK    how many hours back the report window should cover
                          (default: 1)
    VICI_TIMEZONE        IANA timezone name for the report's date/time
                          window (default: "Africa/Johannesburg"). This
                          should match the timezone ViciDial itself uses,
                          NOT the timezone GitHub Actions runs in (which is
                          always UTC).
"""

import os
import sys
import io
import csv
import smtplib
import ssl
from datetime import datetime, timedelta
from email.message import EmailMessage
from zoneinfo import ZoneInfo

import requests
from requests.auth import HTTPBasicAuth
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
VICI_TZ = ZoneInfo(os.environ.get("VICI_TIMEZONE", "Africa/Johannesburg"))

REPORT_PATH = "/vicidial/AST_agent_performance_detail.php"


# --------------------------------------------------------------------------
# Step 1: Download the report from ViciDial
# --------------------------------------------------------------------------

def download_report_text() -> tuple[str, str, str]:
    """Downloads the raw report CSV text for the last REPORT_HOURS_BACK hours."""
    now_local = datetime.now(VICI_TZ)
    start_local = now_local - timedelta(hours=REPORT_HOURS_BACK)

    query_date = start_local.strftime("%Y-%m-%d")
    query_time = start_local.strftime("%H:%M:%S")
    end_date = now_local.strftime("%Y-%m-%d")
    end_time = now_local.strftime("%H:%M:%S")

    # "--ALL--" tells ViciDial to include every campaign/group/agent,
    # rather than us having to hardcode and maintain a list that will
    # drift out of date as campaigns are added or removed.
    params = {
        "DB": "0",
        "query_date": query_date,
        "query_time": query_time,
        "end_date": end_date,
        "end_time": end_time,
        "group[]": "--ALL--",
        "user_group[]": "--ALL--",
        "users[]": "--ALL--",
        "report_display_type": "TEXT",
        "shift": "--",
        "stage": "",
        "show_percentages": "",
        "live_agents": "",
        "time_in_sec": "",
        "search_archived_data": "",
        "show_defunct_users": "",
        "breakdown_by_date": "",
        "file_download": "1",
        "SUBMIT": "SUBMIT",
    }

    url = f"{VICI_SERVER}{REPORT_PATH}"
    resp = requests.get(
        url,
        params=params,
        auth=HTTPBasicAuth(VICI_API_USER, VICI_API_PASS),
        timeout=90,
    )

    if resp.status_code == 401:
        raise RuntimeError(
            "ViciDial rejected the credentials (HTTP 401). Double check "
            "VICI_API_USER / VICI_API_PASS match the sign-in popup exactly."
        )
    resp.raise_for_status()

    text = resp.text
    window_label = f"{query_date} {query_time} -> {end_date} {end_time} ({VICI_TZ.key})"

    # Debug logging: first few lines, so any format surprise is visible in
    # the GitHub Actions run log without guesswork.
    preview_lines = text.splitlines()[:6]
    print("---- Raw report response preview (first 6 lines) ----")
    for line in preview_lines:
        print(repr(line[:200]))
    print("---- End preview ----")

    return text, query_date + "_" + query_time, end_date + "_" + end_time


# --------------------------------------------------------------------------
# Step 2: Parse the double-encoded CSV into a DataFrame
# --------------------------------------------------------------------------

def _unwrap_row(line: str) -> list[str]:
    """Undoes the double CSV-encoding used by this report: each raw line is
    one big quoted CSV field, whose content is itself a normal CSV row."""
    outer = next(csv.reader([line]), [])
    if not outer:
        return []
    return next(csv.reader([outer[0]]), [])


def parse_report(text: str) -> pd.DataFrame:
    lines = text.splitlines()

    header_idx = None
    for i, line in enumerate(lines):
        if "USER NAME" in line:
            header_idx = i
            break

    if header_idx is None:
        # No agents matched the filters/time window — a legitimate, if
        # uneventful, outcome (e.g. an overnight hour with no logged-in
        # agents). Return an empty frame rather than failing.
        print("No 'USER NAME' header row found in the report — treating as no data.")
        return pd.DataFrame()

    header = _unwrap_row(lines[header_idx])

    rows = []
    skipped = 0
    for line in lines[header_idx + 1:]:
        if not line.strip():
            continue
        row = _unwrap_row(line)
        if len(row) != len(header):
            skipped += 1
            continue
        rows.append(row)

    if skipped:
        print(f"WARNING: skipped {skipped} malformed row(s) that didn't match the header's column count.")

    return pd.DataFrame(rows, columns=header)


# --------------------------------------------------------------------------
# Step 3: Save to Excel
# --------------------------------------------------------------------------

def build_excel(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        sheet_name = "Agent Performance Detail"
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        for i, col in enumerate(df.columns, start=1):
            max_len = max([len(str(col))] + [len(str(v)) for v in df[col]]) if len(df) else len(str(col))
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(max_len + 2, 40)
    buffer.seek(0)
    return buffer.read()


# --------------------------------------------------------------------------
# Step 4: Email the file
# --------------------------------------------------------------------------

def send_email(xlsx_bytes: bytes, start_label: str, end_label: str, row_count: int):
    msg = EmailMessage()
    msg["Subject"] = f"Agent Performance Detail Report ({start_label} to {end_label})"
    msg["From"] = EMAIL_FROM
    msg["To"] = ", ".join(EMAIL_TO)
    msg.set_content(
        f"Attached: Agent Performance Detail report for {start_label} to {end_label}.\n"
        f"{row_count} agent record(s) included.\n\n"
        f"This is an automated message."
    )

    safe_name = f"agent_performance_detail_{start_label}".replace(":", "").replace(" ", "_")
    filename = f"{safe_name}.xlsx"
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
    print(f"[{datetime.now(VICI_TZ).isoformat()}] Downloading Agent Performance Detail report...")
    text, start_label, end_label = download_report_text()

    df = parse_report(text)
    print(f"Parsed {len(df)} agent row(s), {len(df.columns) if len(df.columns) else 0} column(s).")

    xlsx_bytes = build_excel(df)
    print("Excel file built, sending email...")

    send_email(xlsx_bytes, start_label, end_label, len(df))
    print("Email sent successfully.")


if __name__ == "__main__":
    main()
