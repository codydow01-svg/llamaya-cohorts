#!/usr/bin/env python3
"""
Llamaya Cohort Analysis
Reads orders directly from the orders spreadsheet,
computes cohort recharge metrics in 3x28-day windows,
writes summary + detail to Cohort_Llamaya sheet daily.
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

# Three 28-day windows after first purchase
WINDOWS = [
    ("recharged_1",  1,  28),
    ("recharged_2", 29,  56),
    ("recharged_3", 57,  84),
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

    # Per-user cohort date (earliest first purchase)
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

    cohorts["cohort_month"] = cohorts["cohort_date"].dt.to_period("M")

    # Summary: one row per cohort month
    summary = (
        cohorts.groupby("cohort_month")
        .agg(
            total=("email", "count"),
            recharged_1=("recharged_1", "sum"),
            recharged_2=("recharged_2", "sum"),
            recharged_3=("recharged_3", "sum"),
        )
        .reset_index()
        .sort_values("cohort_month")
    )
    summary["cohort_month"] = summary["cohort_month"].astype(str)
    summary["rate_1"] = (summary["recharged_1"] / summary["total"] * 100).round(1).astype(str) + "%"
    summary["rate_2"] = (summary["recharged_2"] / summary["total"] * 100).round(1).astype(str) + "%"
    summary["rate_3"] = (summary["recharged_3"] / summary["total"] * 100).round(1).astype(str) + "%"

    cohorts["cohort_date"]  = cohorts["cohort_date"].dt.strftime("%d.%m.%Y")
    cohorts["cohort_month"] = cohorts["cohort_month"].astype(str)

    return summary, cohorts

def write_cohorts(client, summary, cohorts):
    print("\nWriting to Cohort_Llamaya...")
    spreadsheet = client.open_by_key(COHORT_SPREADSHEET_ID)
    try:
        sheet = spreadsheet.worksheet(COHORT_SHEET_NAME)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(COHORT_SHEET_NAME, rows=10000, cols=10)
        print("  Created new sheet Cohort_Llamaya")

    summary_headers = ["cohort_month", "total", "recharged_1", "recharged_2", "recharged_3", "rate_1", "rate_2", "rate_3"]
    detail_headers  = ["email", "cohort_date", "cohort_month", "recharged_1", "recharged_2", "recharged_3"]

    summary_rows = summary[summary_headers].values.tolist()
    detail_rows  = cohorts[detail_headers].values.tolist()

    all_rows = (
        [summary_headers] + summary_rows +
        [[]] +
        [detail_headers] + detail_rows
    )

    sheet.clear()
    sheet.update(all_rows, value_input_option="USER_ENTERED")
    print(f"  Summary: {len(summary_rows)} cohort months")
    print(f"  Detail: {len(detail_rows)} users")
    print(f"  Updated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")

def main():
    client = get_client()
    df = read_orders(client)
    summary, cohorts = compute_cohorts(df)
    write_cohorts(client, summary, cohorts)
    print("\nDone!")

if __name__ == "__main__":
    main()
