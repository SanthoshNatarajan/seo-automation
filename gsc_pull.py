# ============================================
# Week 1 — GSC API Auto Pull
# Built by Santhosh Natarajan | SEO Automation
# ============================================

import os
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta

# ── CONFIG ──────────────────────────────────
CREDENTIALS_FILE = r"C:\seo-automation\credentials.json"
TOKEN_FILE       = r"C:\seo-automation\token.json"
OUTPUT_FILE      = r"C:\seo-automation\gsc_data.xlsx"
SITE_URL         = "sc-domain:systechgroup.in"
SCOPES           = ["https://www.googleapis.com/auth/webmasters.readonly"]

# Date range — last 3 months
END_DATE   = datetime.today().strftime("%Y-%m-%d")
START_DATE = (datetime.today() - timedelta(days=90)).strftime("%Y-%m-%d")

# ── AUTHENTICATE ─────────────────────────────
def authenticate():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
    return creds

# ── PULL DATA FROM GSC ───────────────────────
def pull_gsc_data(service):
    all_rows = []
    start_row = 0
    batch_size = 25000

    print(f"Pulling data from {START_DATE} to {END_DATE}...")
    print("Please wait...")

    while True:
        request = {
            "startDate": START_DATE,
            "endDate":   END_DATE,
            "dimensions": ["query"],
            "rowLimit":   batch_size,
            "startRow":   start_row,
        }
        response = service.searchanalytics().query(
            siteUrl=SITE_URL, body=request
        ).execute()

        rows = response.get("rows", [])
        if not rows:
            break

        for row in rows:
            all_rows.append({
                "Query":       row["keys"][0],
                "Clicks":      row["clicks"],
                "Impressions": row["impressions"],
                "CTR":         round(row["ctr"] * 100, 2),
                "Position":    round(row["position"], 1),
            })

        start_row += batch_size
        print(f"  Fetched {len(all_rows):,} rows so far...")

        if len(rows) < batch_size:
            break

    return all_rows

# ── SAVE TO EXCEL ────────────────────────────
def save_to_excel(data):
    df = pd.DataFrame(data)
    df = df.sort_values("Impressions", ascending=False).reset_index(drop=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="GSC Data")

    print(f"\n✅ Done! {len(df):,} rows saved to:")
    print(f"   {OUTPUT_FILE}")
    print(f"\nTop 10 queries by impressions:")
    print(df[["Query","Impressions","Clicks","CTR","Position"]].head(10).to_string(index=False))

# ── MAIN ─────────────────────────────────────
if __name__ == "__main__":
    print("=" * 50)
    print("  GSC Auto Pull — Santhosh Natarajan")
    print("=" * 50)

    creds   = authenticate()
    service = build("webmasters", "v3", credentials=creds)
    data    = pull_gsc_data(service)
    save_to_excel(data)

    print("\n🚀 Script complete. No manual download needed.")
