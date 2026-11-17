#!/usr/bin/env python3
"""
NPD/EPD Launch Tracker — Weekly TAT Breach Reporter
Finds items where Given Date has passed but Actual Date is still empty
(not yet delivered = currently overdue) and writes them to an 'automation'
tab in the same sheet for n8n to pick up.

Usage:
    python3 breach_reporter.py
    python3 breach_reporter.py --dry-run   # prints to console, no sheet write
"""

import os
import re
import sys
from datetime import datetime, date
from typing import Optional, List, Dict, Any

from google.oauth2 import service_account
from googleapiclient.discovery import build

# ─── CONFIGURATION ────────────────────────────────────────────────────────────

SPREADSHEET_ID       = "17Zs-ZI6sBFpxszmyiwt4oFzp1MTiqhEE1VtWyxsunv0"
AUTOMATION_TAB       = "automation"

# Service account: env var takes priority (GitHub Actions), falls back to local file
SERVICE_ACCOUNT_FILE = os.environ.get(
    "SERVICE_ACCOUNT_FILE",
    "/Users/sahil/Documents/NPD:EPD Launch Tracker/npd-epd-launch-tracker-d56a8052a380.json",
)

# ─── SHEET COLUMN MAPPING (0-based; data rows start after 4 header rows) ─────

COL_BRAND        = 0   # A
COL_NPD_OWNER    = 1   # B — NPD owner
COL_CONTENT_OWN  = 2   # C — Content team owner
COL_DESIGNER     = 3   # D — Designer
COL_PRODUCT_NAME = 5   # F
COL_PRODUCT_CODE = 6   # G
COL_OVERALL_PLAN = 27  # AB — overall plan date
COL_OVERALL_ACT  = 28  # AC — overall actual date

# (team_label, plan_col, given_col, actual_col, owner_col) — 0-based
TEAMS = [
    ("NPD Team",           11, 12, 13, COL_NPD_OWNER),    # L, M, N → col B
    ("Content Team",       14, 15, 16, COL_CONTENT_OWN),  # O, P, Q → col C
    ("Legal",              17, 18, 19, COL_NPD_OWNER),    # R, S, T → col B
    ("Design Team",        20, 21, 22, COL_DESIGNER),     # U, V, W → col D
    ("Ready Date (China)", 23, 24, 25, COL_NPD_OWNER),    # X, Y, Z → col B
]

# ─── AUTOMATION TAB HEADERS ───────────────────────────────────────────────────

AUTO_HEADERS = [
    "Type", "Product", "Product Code", "Brand", "Owner", "Team",
    "Plan Date", "Given Date", "Actual Date",
    "Days Overdue (Given vs Today)", "Plan vs Actual (days)",
    "Report Date",
]

# ──────────────────────────────────────────────────────────────────────────────


def parse_date(val: str) -> Optional[date]:
    if not val or not val.strip():
        return None
    val = val.strip()
    val = re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", val, flags=re.IGNORECASE)
    val = val.replace("/", "-").replace(" ", "-")
    val = re.sub(r"-+", "-", val)

    current_year = datetime.now().year
    for fmt in ("%d-%b", "%d-%B", "%d-%b-%Y", "%d-%B-%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            parsed = datetime.strptime(val, fmt)
            if parsed.year == 1900:
                parsed = parsed.replace(year=current_year)
            return parsed.date()
        except ValueError:
            continue
    return None


def get_cell(row: list, idx: int) -> str:
    return str(row[idx]).strip() if idx < len(row) else ""


def ensure_automation_tab(service, spreadsheet_id: str) -> None:
    """Create the 'automation' tab if it doesn't exist."""
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    existing = [s["properties"]["title"] for s in meta.get("sheets", [])]
    if AUTOMATION_TAB not in existing:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": AUTOMATION_TAB}}}]},
        ).execute()
        print(f"Created tab: {AUTOMATION_TAB}")


def write_automation_tab(service, spreadsheet_id: str, rows: List[List[str]]) -> None:
    """Clear the automation tab and write fresh breach rows."""
    # Clear existing content
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"{AUTOMATION_TAB}!A:Z",
    ).execute()

    # Write header + data
    all_rows = [AUTO_HEADERS] + rows
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{AUTOMATION_TAB}!A1",
        valueInputOption="RAW",
        body={"values": all_rows},
    ).execute()
    print(f"Written {len(rows)} breach rows to '{AUTOMATION_TAB}' tab.")


def main(dry_run: bool = False) -> None:
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],  # read + write
    )
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    # ── Read source data ──────────────────────────────────────────────────────
    result = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range="Planned vs Actual!A4:AH300",
        )
        .execute()
    )
    rows: List[list] = result.get("values", [])
    today = date.today()

    # ── Find overdue + delivered-late items ───────────────────────────────────
    breach_rows: List[List[str]] = []

    for row in rows:
        brand   = get_cell(row, COL_BRAND)
        product = get_cell(row, COL_PRODUCT_NAME)

        if not product:
            continue

        product_code = get_cell(row, COL_PRODUCT_CODE)
        npd_owner    = get_cell(row, COL_NPD_OWNER) or "Unknown"

        # ── Pending: check each team's given date independently ──────────
        for team_label, plan_col, given_col, actual_col, owner_col in TEAMS:
            plan_date   = parse_date(get_cell(row, plan_col))
            given_date  = parse_date(get_cell(row, given_col))
            actual_date = parse_date(get_cell(row, actual_col))
            team_owner  = get_cell(row, owner_col) or "Unknown"

            if given_date and not actual_date and given_date < today:
                days_overdue = (today - given_date).days
                breach_rows.append([
                    "Pending (Not Delivered)",
                    product, product_code, brand, team_owner, team_label,
                    plan_date.strftime("%d-%b") if plan_date else "",
                    given_date.strftime("%d-%b"),
                    "",
                    str(days_overdue),
                    "",
                    today.strftime("%d-%b-%Y"),
                ])

        # ── Delivered Late: overall AB (plan) vs AC (actual) ─────────────
        overall_plan = parse_date(get_cell(row, COL_OVERALL_PLAN))
        overall_act  = parse_date(get_cell(row, COL_OVERALL_ACT))
        if overall_plan and overall_act and overall_act > overall_plan:
            days_late = (overall_act - overall_plan).days
            breach_rows.append([
                "Delivered Late",
                product, product_code, brand, npd_owner, "Overall",
                overall_plan.strftime("%d-%b"),
                "",
                overall_act.strftime("%d-%b"),
                str(days_late),
                str(days_late),
                today.strftime("%d-%b-%Y"),
            ])

    # ── Deduplicate: one entry per (owner, product_code) — keep least overdue ──
    seen: dict = {}
    for r in breach_rows:
        key = (r[4], r[2])  # owner, product_code
        existing = seen.get(key)
        if existing is None:
            seen[key] = r
        else:
            curr_days = int(r[9])        if r[9].lstrip("-").isdigit()        else 0
            prev_days = int(existing[9]) if existing[9].lstrip("-").isdigit() else 0
            if curr_days < prev_days:
                seen[key] = r
    breach_rows = list(seen.values())

    # ── Output ────────────────────────────────────────────────────────────────
    if dry_run:
        pending  = [r for r in breach_rows if r[0] == "Pending (Not Delivered)"]
        late     = [r for r in breach_rows if r[0] == "Delivered Late"]
        print(f"\n{'='*110}")
        print(f"PENDING (not delivered, given date passed): {len(pending)}")
        print(f"{'='*110}")
        print(f"{'Product':<30} {'Code':<12} {'Owner':<12} {'Team':<22} {'Given':<10} {'Days Overdue (vs Today)'}")
        print("-" * 110)
        for r in pending:
            # r: Type, Product, ProductCode, Brand, Owner, Team, PlanDate, GivenDate, _, DaysOverdue, PlanVsActual, ReportDate
            print(f"{r[1]:<30} {r[2]:<12} {r[4]:<12} {r[5]:<22} {r[7]:<10} {r[9]}")
        print(f"\n{'='*110}")
        print(f"DELIVERED LATE (actual > given commitment): {len(late)}")
        print(f"{'='*110}")
        print(f"{'Product':<30} {'Code':<12} {'Owner':<12} {'Team':<22} {'Given':<10} {'Actual':<10} {'Given→Actual':<15} {'Plan→Actual'}")
        print("-" * 110)
        for r in late:
            print(f"{r[1]:<30} {r[2]:<12} {r[4]:<12} {r[5]:<22} {r[7]:<10} {r[8]:<10} {r[9]:<15} {r[10]}")
        return

    if not breach_rows:
        print("No overdue items found. Clearing automation tab.")
        ensure_automation_tab(service, SPREADSHEET_ID)
        write_automation_tab(service, SPREADSHEET_ID, [])
        return

    ensure_automation_tab(service, SPREADSHEET_ID)
    write_automation_tab(service, SPREADSHEET_ID, breach_rows)


if __name__ == "__main__":
    dry_run = "--dry-run" in sys.argv
    main(dry_run=dry_run)
