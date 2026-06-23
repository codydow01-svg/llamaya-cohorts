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


# ── Config ─────────────────────────────────────────────────────────────────────────────────
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


# Recharge windows (days after first purchase)
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
    # Column positions (0-indexed): B=1 date, C=2 email, I=8 stripe_id, O=14 recharge, T=19 operator
