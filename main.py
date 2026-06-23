#!/usr/bin/env python3
"""
Llamaya Cohort Analysis
Reads orders from Google Sheets, computes cohort recharge metrics,
writes results to Cohort_Llamaya sheet daily.
"""

import os
import json
import gspread
import pandas as pd
from datetime import datetime
from google.oauth2.service_account import Credentials

# Orders data is read via IMPORTRANGE mirror in the cohort spreadsheet (Лист1).
# This avoids needing direct access to the read-only orders spreadsheet.
# Лист1 formula: =IMPORTRANGE("159XaSuCaBBb-9d_J93ujHgA5DpintxW37QRdeOQCOtE","orders!A1:T35000")
ORDERS_SPREADSHEET_ID = "1sM00OKAvedi4GlNav3wtN-efEBl2fxUHUhebkr37xcA"  # cohort spreadsheet (has mirror)
COHORT_SPREADSHEET_ID = "1sM00OKAvedi4GlNav3wtN-efEBl2fxUHUhebkr37xcA"
ORDERS_SHEET_NAME     = "Лист1"   # IMPORTRANGE tab: orders!A1:T35000
COHORT_SHEET_NAME     = "Cohort_Llamaya"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

WINDOWS = [
    ("recharged_1", 14, 42),
    ("recharged_2", 43, 71),
    ("recharged_3", 72, 100),
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
        (df[col_stripe].str.strip() != "") &
        (df[col_stripe].notna())
    )
    llamaya = df[mask].copy()
    print(f"  Llamaya rows with stripe_id: {len(llamaya)}")

    llamaya[col_date] = pd.to_datetime(llamaya[col_date], dayfirst=True, errors="coerce")
    llamaya = llamaya.dropna(subset=[col_date])

    first = llamaya[llamaya[col_recharge].str.strip() == ""].copy()
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

    cohorts["cohort_date"] = cohorts["cohort_date"].dt.strftime("%d.%m.%Y")
    return cohorts


def write_cohorts(client, cohorts):
    print("\nWriting to Cohort_Llamaya...")
    spreadsheet = client.open_by_key(COHORT_SPREADSHEET_ID)

    try:
        sheet = spreadsheet.worksheet(COHORT_SHEET_NAME)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(COHORT_SHEET_NAME, rows=10000, cols=5)
        print("  Created new sheet Cohort_Llamaya")

    headers = ["email", "cohort_date", "recharged_1", "recharged_2", "recharged_3"]
    rows = cohorts[headers].values.tolist()

    sheet.clear()
    sheet.update([headers] + rows, value_input_option="USER_ENTERED")

    print(f"  Written {len(rows)} rows")
    print(f"  Updated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")


def main():
    client = get_client()
    df = read_orders(client)
    cohorts = compute_cohorts(df)
    write_cohorts(client, cohorts)
    print("\nDone!")


if __name__ == "__main__":
    main()
