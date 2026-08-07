#!/usr/bin/env python3
"""
Llamaya Cohort Analysis
=======================
Reads all orders from Main_Sheet-2-1, filters Operator = "Llamaya",
computes weekly retention cohorts using MSISDN as the customer identifier,
then writes results to Cohort_Llamaya tab.

Cohort windows (relative to each customer first purchase date):
p1 d1-28 p2 d29-56 p3 d57-84

p1/p2/p3 counts only customers who have a Recharge="yes" transaction
in the respective window (column Z of the orders sheet).
"""

import os
import json
import datetime
from collections import defaultdict

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

ORDERS_SHEET_ID = "1eJau3HSsP_qYA7Sy2C9AF17uA47B3QlrnDmzDM59yrg"
ORDERS_TAB = "Main_Sheet-2-1"
COHORT_SHEET_ID = "1sM00OKAvedi4GlNav3wtN-efEBl2fxUHUhebkr37xcA"
COHORT_TAB = "Cohort_Llamaya"

COL_PAID     = 1   # column B: "Paid"
COL_OPERATOR = 9   # column J: "Operator"
COL_MSISDN   = 24  # column Y: "MSISDN"
COL_RECHARGE = 25  # column Z: "Recharge" - "yes" means customer recharged

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

HEADER = [
    "cohort_week", "customers",
    "p1 (d1-28)", "p1 %",
    "p2 (d29-56)", "p2 %",
    "p3 (d57-84)", "p3 %",
    "total (ever)", "total %",
]

GREEN = {"red": 0.204, "green": 0.659, "blue": 0.325}
WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}

def get_creds():
    info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    return Credentials.from_service_account_info(info, scopes=SCOPES)

def week_monday(d):
    return d - datetime.timedelta(days=d.weekday())

def parse_date(s):
    if not s:
        return None
    try:
        return datetime.datetime.strptime(s[:10], "%Y-%m-%d").date()
    except ValueError:
        return None

def build_cohorts(rows):
    msisdn_entries = defaultdict(list)  # MSISDN -> list of (date, is_recharge)
    for row in rows:
        if len(row) <= COL_MSISDN:
            continue
        if (row[COL_OPERATOR] or "").strip() != "Llamaya":
            continue
        msisdn = (row[COL_MSISDN] or "").strip()
        if not msisdn:
            continue
        d = parse_date(row[COL_PAID])
        if not d:
            continue
        is_recharge = (len(row) > COL_RECHARGE and
                       (row[COL_RECHARGE] or "").strip().lower() == "yes")
        msisdn_entries[msisdn].append((d, is_recharge))

    cohort_customers = defaultdict(set)
    cohort_p1 = defaultdict(set)
    cohort_p2 = defaultdict(set)
    cohort_p3 = defaultdict(set)

    for msisdn, entries in msisdn_entries.items():
        first = min(d for d, _ in entries)
        week = week_monday(first)
        cohort_customers[week].add(msisdn)
        for d, is_recharge in entries:
            if not is_recharge:
                continue
            delta = (d - first).days
            if  1 <= delta <= 28: cohort_p1[week].add(msisdn)
            if 29 <= delta <= 56: cohort_p2[week].add(msisdn)
            if 57 <= delta <= 84: cohort_p3[week].add(msisdn)

    result = []
    for week in sorted(cohort_customers):
        n = len(cohort_customers[week])
        p1 = len(cohort_p1[week])
        p2 = len(cohort_p2[week])
        p3 = len(cohort_p3[week])
        total = len(cohort_p1[week] | cohort_p2[week] | cohort_p3[week])
        result.append({
            "cohort_week": week.strftime("%Y-%m-%d"),
            "customers": n,
            "p1": p1, "p1_pct": p1 / n if n else 0,
            "p2": p2, "p2_pct": p2 / n if n else 0,
            "p3": p3, "p3_pct": p3 / n if n else 0,
            "total": total, "total_pct": total / n if n else 0,
        })
    return result

def cohort_to_row(c):
    return [c["cohort_week"], c["customers"],
            c["p1"], c["p1_pct"],
            c["p2"], c["p2_pct"],
            c["p3"], c["p3_pct"],
            c["total"], c["total_pct"]]

def apply_formatting(service, spreadsheet_id, sheet_id, n_data):
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties.sheetId,conditionalFormats)",
    ).execute()
    n_rules = 0
    for s in meta.get("sheets", []):
        if s["properties"]["sheetId"] == sheet_id:
            n_rules = len(s.get("conditionalFormats", []))
            break

    requests = [{"deleteConditionalFormatRule": {"sheetId": sheet_id, "index": i}}
                for i in range(n_rules - 1, -1, -1)]

    end_row = n_data + 1
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": True},
                "backgroundColor": {"red": 0.851, "green": 0.851, "blue": 0.851},
            }},
            "fields": "userEnteredFormat(textFormat.bold,backgroundColor)",
        }
    })
    for col in [3, 5, 7, 9]:
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id,
                          "startRowIndex": 1, "endRowIndex": end_row,
                          "startColumnIndex": col, "endColumnIndex": col + 1},
                "cell": {"userEnteredFormat": {
                    "numberFormat": {"type": "PERCENT", "pattern": "0.0%"}
                }},
                "fields": "userEnteredFormat.numberFormat",
            }
        })
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet_id,
                                "startRowIndex": 1, "endRowIndex": end_row,
                                "startColumnIndex": col, "endColumnIndex": col + 1}],
                    "gradientRule": {
                        "minpoint": {"color": WHITE, "type": "MIN"},
                        "maxpoint": {"color": GREEN, "type": "MAX"},
                    },
                },
                "index": 0,
            }
        })
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()

def main():
    creds = get_creds()
    gc = gspread.authorize(creds)
    service = build("sheets", "v4", credentials=creds)

    print("Reading orders...")
    orders_ws = gc.open_by_key(ORDERS_SHEET_ID).worksheet(ORDERS_TAB)
    data_rows = orders_ws.get_all_values()[1:]
    print(f"  {len(data_rows)} rows")

    cohorts = build_cohorts(data_rows)
    print(f"  {len(cohorts)} cohort weeks")
    for c in cohorts[-5:]:
        print(f"  {c['cohort_week']}: {c['customers']} customers "
              f"p1={c['p1_pct']:.1%} p2={c['p2_pct']:.1%} p3={c['p3_pct']:.1%} total={c['total_pct']:.1%}")

    print("Writing Cohort_Llamaya...")
    cohort_wb = gc.open_by_key(COHORT_SHEET_ID)
    cohort_ws = cohort_wb.worksheet(COHORT_TAB)
    sheet_rows = [HEADER] + [cohort_to_row(c) for c in cohorts]
    cohort_ws.clear()
    cohort_ws.update("A1", sheet_rows, value_input_option="USER_ENTERED")
    print(f"  {len(sheet_rows)} rows written")

    print("Formatting...")
    apply_formatting(service, COHORT_SHEET_ID, cohort_ws.id, len(cohorts))
    print("Done!")

if __name__ == "__main__":
    main()
