#!/usr/bin/env python3
"""
Llamaya Cohort Analysis - with Google Sheets formatting
Weekly cohorts, 3x28-day recharge windows, green gradient on % columns.
"""

import os
import json
import gspread
import pandas as pd
from datetime import datetime
from google.oauth2.service_account import Credentials

ORDERS_SPREADSHEET_ID = "1eJau3HSsP_qYA7Sy2C9AF17uA47B3QlrnDmzDM59yrg"
ORDERS_SHEET_NAME     = "Main_Sheet-2-1"
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

# Column indices (0-based).
COL_DATE     = 1   # B — order date
COL_EMAIL    = 2   # C — customer email
COL_STRIPE   = 8   # I — stripe_id
COL_RECHARGE = 25  # Z — recharge status
COL_OPERATOR = 9   # J — operator


def get_client():
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not creds_json:
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON environment variable not set")
    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
    client = gspread.authorize(creds)
    print(f"  Service account: {creds.service_account_email}")
    return client


def read_orders(client):
    print("Reading orders...")
    sheet = client.open_by_key(ORDERS_SPREADSHEET_ID).worksheet(ORDERS_SHEET_NAME)
    rows = sheet.get_all_values()
    if len(rows) < 2:
        raise ValueError("Orders sheet is empty")

    headers = rows[0]
    print(f"  Total columns: {len(headers)}")
    print("  Headers (index: name):")
    for i, h in enumerate(headers):
        if h:
            print(f"    {i}: {h}")

    print(f"\n  Key columns used:")
    print(f"    COL_DATE     [{COL_DATE}]  = '{headers[COL_DATE] if COL_DATE < len(headers) else 'OUT OF RANGE'}'")
    print(f"    COL_EMAIL    [{COL_EMAIL}]  = '{headers[COL_EMAIL] if COL_EMAIL < len(headers) else 'OUT OF RANGE'}'")
    print(f"    COL_STRIPE   [{COL_STRIPE}]  = '{headers[COL_STRIPE] if COL_STRIPE < len(headers) else 'OUT OF RANGE'}'")
    print(f"    COL_RECHARGE [{COL_RECHARGE}] = '{headers[COL_RECHARGE] if COL_RECHARGE < len(headers) else 'OUT OF RANGE'}'")
    print(f"    COL_OPERATOR [{COL_OPERATOR}]  = '{headers[COL_OPERATOR] if COL_OPERATOR < len(headers) else 'OUT OF RANGE'}'")

    df = pd.DataFrame(rows[1:], columns=headers)
    print(f"\n  Total rows: {len(df)}")
    return df


def compute_cohorts(df):
    cols = df.columns
    col_date     = cols[COL_DATE]
    col_email    = cols[COL_EMAIL]
    col_stripe   = cols[COL_STRIPE]
    col_recharge = cols[COL_RECHARGE]
    col_operator = cols[COL_OPERATOR]

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
            "customers":   total,
        }
        for col_name, days_from, days_to in WINDOWS:
            days_since = (today - week_start).days
            if days_since < days_from:
                row[f"{col_name}_count"] = "рано"
                row[f"{col_name}_pct"]   = ""
            else:
                count = int(group[col_name].sum())
                row[f"{col_name}_count"] = count
                row[f"{col_name}_pct"]   = round(count / total * 100, 1)
        records.append(row)

    summary = pd.DataFrame(records)
    print(f"  Cohort weeks: {len(summary)}")
    return summary


def apply_formatting(spreadsheet, sheet, num_rows):
    sid = sheet.id
    end_row = num_rows + 1
    requests = []

    requests.append({"repeatCell": {"range": {"sheetId": sid}, "cell": {"userEnteredFormat": {}}, "fields": "userEnteredFormat"}})

    requests.append({"repeatCell": {"range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 8}, "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 10}, "backgroundColor": {"red": 0.851, "green": 0.851, "blue": 0.851}, "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat(textFormat,backgroundColor,verticalAlignment)"}})

    requests.append({"updateSheetProperties": {"properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1}}, "fields": "gridProperties.frozenRowCount"}})

    requests.append({"repeatCell": {"range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": end_row, "startColumnIndex": 1, "endColumnIndex": 8}, "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}}, "fields": "userEnteredFormat.horizontalAlignment"}})

    for col in [3, 5, 7]:
        requests.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sid, "startRowIndex": 1, "endRowIndex": end_row, "startColumnIndex": col, "endColumnIndex": col + 1}], "gradientRule": {"minpoint": {"colorStyle": {"rgbColor": {"red": 1.0, "green": 1.0, "blue": 1.0}}, "type": "NUMBER", "value": "0"}, "maxpoint": {"colorStyle": {"rgbColor": {"red": 0.204, "green": 0.659, "blue": 0.325}}, "type": "NUMBER", "value": "30"}}}, "index": 0}})

    for col_start, col_end in [(2, 4), (4, 6), (6, 8)]:
        requests.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sid, "startRowIndex": 1, "endRowIndex": end_row, "startColumnIndex": col_start, "endColumnIndex": col_end}], "booleanRule": {"condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "рано"}]}, "format": {"backgroundColor": {"red": 0.941, "green": 0.941, "blue": 0.941}, "textFormat": {"foregroundColor": {"red": 0.6, "green": 0.6, "blue": 0.6}, "italic": True}}}}, "index": 0}})

    for i, w in enumerate([130, 90, 100, 80, 110, 80, 110, 80]):
        requests.append({"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1}, "properties": {"pixelSize": w}, "fields": "pixelSize"}})

    requests.append({"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 32}, "fields": "pixelSize"}})

    spreadsheet.batch_update({"requests": requests})
    print("  Formatting applied")


def write_cohorts(client, summary):
    print("\nWriting to Cohort_Llamaya...")
    spreadsheet = client.open_by_key(COHORT_SPREADSHEET_ID)
    try:
        sheet = spreadsheet.worksheet(COHORT_SHEET_NAME)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(COHORT_SHEET_NAME, rows=500, cols=10)
        print("  Created new sheet Cohort_Llamaya")

    sheet.clear()

    headers = ["cohort_week", "customers", "p1 (d1-28)", "p1 %", "p2 (d29-56)", "p2 %", "p3 (d57-84)", "p3 %"]
    data_cols = ["cohort_week", "customers", "p1_count", "p1_pct", "p2_count", "p2_pct", "p3_count", "p3_pct"]

    rows = summary[data_cols].values.tolist()
    sheet.update([headers] + rows, value_input_option="RAW")
    apply_formatting(spreadsheet, sheet, len(rows))

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
