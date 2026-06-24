#!/usr/bin/env python3
"""
Llamaya Cohort Analysis
Weekly cohorts, 3x28-day recharge windows.
'рано' only when window has not started yet; otherwise shows live counts.
"""

import os
import json
import gspread
import pandas as pd
from datetime import datetime
from google.oauth2.service_account import Credentials

ORDERS_SPREADSHEET_ID = "159XaSuCaBBb-9d_J93ujHgA5DpintxW37QRdeOQCOtE"
ORDERS_SHEET_NAME     = "orders"
COHORT_SPREADSHEET_ID = "1sM00OKAvedi4GlNav3wtN-efEBl2fxUHUhebkr37xcA"
COHORT_SHEET_NAME     = "Cohort_Llamaya"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

WINDOWS = [
    ("p1",  1,  28),
    ("p2", 29,  56),
    ("p3", 57,  84),
]

def get_client():
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not creds_json:
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON environment variable not set")
    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
    return gspread.authorize(creds)

def read_orders(client):
    print("Reading orders...")
    sheet = client.open_by_key(ORDERS_SPREADSHEET_ID).worksheet(ORDERS_SHEET_NAME)
    rows = sheet.get_all_values()
    if len(rows) < 2:
        raise ValueError("Orders sheet is empty")
    df = pd.DataFrame(rows[1:], columns=rows[0])
    print(f"  Total rows: {len(df)}")
    return df

def compute_cohorts(df):
    cols = df.columns
    col_date     = cols[1]
    col_email    = cols[2]
    col_stripe   = cols[8]
    col_recharge = cols[14]
    col_operator = cols[19]

    mask = (
        (df[col_operator].str.strip() == "Llamaya") &
        df[col_stripe].notna() & (df[col_stripe] != "")
    )
    llamaya = df[mask].copy()
    print(f"  Llamaya rows with stripe_id: {len(llamaya)}")

    llamaya[col_date] = pd.to_datetime(llamaya[col_date], dayfirst=True, errors="coerce")
    llamaya = llamaya.dropna(subset=[col_date])

    first     = llamaya[llamaya[col_recharge].str.strip() == ""].copy()
    recharged = llamaya[llamaya[col_recharge].str.strip() == "yes"].copy()
    print(f"  First purchases: {len(first)}, Recharges: {len(recharged)}")

    cohorts = (
        first.groupby(col_email)[col_date]
        .min()
        .reset_index()
        .rename(columns={col_email: "email", col_date: "cohort_date"})
        .sort_values("cohort_date")
        .reset_index(drop=True)
    )
    print(f"  Unique customers: {len(cohorts)}")

    recharge_map = recharged.groupby(col_email)[col_date].apply(list).to_dict()

    for col_name, days_from, days_to in WINDOWS:
        def flag(row, d0=days_from, d1=days_to):
            dates = recharge_map.get(row["email"], [])
            lo = row["cohort_date"] + pd.Timedelta(days=d0)
            hi = row["cohort_date"] + pd.Timedelta(days=d1)
            return 1 if any(lo <= d <= hi for d in dates) else 0
        cohorts[col_name] = cohorts.apply(flag, axis=1)

    # Normalize to midnight, then find Monday of that week
    cohort_day = cohorts["cohort_date"].dt.normalize()
    cohorts["cohort_week_start"] = cohort_day - pd.to_timedelta(
        cohort_day.dt.dayofweek, unit="D"
    )

    today = pd.Timestamp.now().normalize()

    records = []
    for week_start, group in cohorts.groupby("cohort_week_start"):
        total = len(group)
        row = {
            "cohort_week": week_start.strftime("%d.%m.%Y"),
            "customers":   str(total),   # str to prevent Sheets date auto-format
        }
        for col_name, days_from, days_to in WINDOWS:
            days_since = (today - week_start).days
            if days_since < days_from:
                # Window hasn't started for anyone yet
                row[f"{col_name}_count"] = "рано"
                row[f"{col_name}_pct"]   = ""
            else:
                count = int(group[col_name].sum())
                pct   = f"{count / total * 100:.1f}%"
                if days_since < days_to:
                    # Window in progress — show live count with note
                    row[f"{col_name}_count"] = str(count)
                    row[f"{col_name}_pct"]   = pct + " *"
                else:
                    # Window fully closed — final numbers
                    row[f"{col_name}_count"] = str(count)
                    row[f"{col_name}_pct"]   = pct
        records.append(row)

    summary = pd.DataFrame(records)
    print(f"  Cohort weeks: {len(summary)}")
    return summary

def write_cohorts(client, summary):
    print("\nWriting to Cohort_Llamaya...")
    spreadsheet = client.open_by_key(COHORT_SPREADSHEET_ID)
    try:
        sheet = spreadsheet.worksheet(COHORT_SHEET_NAME)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(COHORT_SHEET_NAME, rows=500, cols=10)
        print("  Created new sheet Cohort_Llamaya")

    # Clear values then reset all cell formatting (removes leftover date formats)
    sheet.clear()
    spreadsheet.batch_update({"requests": [{
        "repeatCell": {
            "range": {
                "sheetId": sheet.id,
                "startRowIndex": 0,
                "startColumnIndex": 0,
                "endColumnIndex": 8
            },
            "cell": {"userEnteredFormat": {}},
            "fields": "userEnteredFormat"
        }
    }]})

    headers = ["cohort_week", "customers",
               "p1 (d1-28)", "p1 %",
               "p2 (d29-56)", "p2 %",
               "p3 (d57-84)", "p3 %"]
    data_cols = ["cohort_week", "customers",
                 "p1_count", "p1_pct",
                 "p2_count", "p2_pct",
                 "p3_count", "p3_pct"]

    rows = summary[data_cols].values.tolist()
    # Use RAW so strings stay strings, no date auto-parsing
    sheet.update([headers] + rows, value_input_option="RAW")
    print(f"  Written {len(rows)} cohort weeks")
    print(f"  Updated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")

def main():
    client = get_client()
    df = read_orders(client)
    summary = compute_cohorts(df)
    write_cohorts(client, summary)
    print("\nDone!")

if __name__ == "__main__":
    main()
