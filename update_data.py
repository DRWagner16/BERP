import os
import json
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# --- CONFIGURATION ---
SHEET_NAME = "BERP AR Tracking"

# --- 1. AUTHENTICATE ---
raw_creds = os.environ.get("GCP_SERVICE_ACCOUNT")
if not raw_creds:
    print("Error: GCP_SERVICE_ACCOUNT secret is missing.")
    exit(1)

scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = json.loads(raw_creds)
creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
client = gspread.authorize(creds)

# --- 2. LOAD DATA ---
try:
    print(f"Opening Google Sheet: '{SHEET_NAME}'...")
    sheet = client.open(SHEET_NAME).sheet1
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    df.columns = df.columns.astype(str).str.strip()
    print("Data loaded successfully.")
except Exception as e:
    print(f"Error loading sheet: {e}")
    exit(1)

# --- 3. SANITIZE (PRIVACY) ---
# CRITICAL: We drop 'Company' before saving to JSON.
if 'Company' in df.columns:
    df = df.drop(columns=['Company'])

# --- 4. CLEAN & CALCULATE ---
numeric_cols = ['Gas Savings (MMBtu/yr)', 'Electric Savings (kWh/yr)', 'Total Cost Savings', 'Implementation Costs', 'Percent Progress']
for col in numeric_cols:
    if col not in df.columns: df[col] = 0
    df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$,%]', '', regex=True), errors='coerce').fillna(0)

# Emission Factors
FACTORS = { 'Gas_CO2': 117.0, 'Elec_CO2': 1.5, 'Gas_NOx': 0.092, 'Elec_NOx': 0.001 }

# Calculations
df['Gas_CO2_lb'] = df['Gas Savings (MMBtu/yr)'] * FACTORS['Gas_CO2']
df['Elec_CO2_lb'] = df['Electric Savings (kWh/yr)'] * FACTORS['Elec_CO2']
df['Total_CO2_Tons'] = (df['Gas_CO2_lb'] + df['Elec_CO2_lb']) / 2000.0

# --- 5. EXPORT TO JSON ---
# We save this file to the repository. The website will read this file.
json_output = df.to_json(orient='records')

with open('site_data.json', 'w') as f:
    f.write(json_output)
    print("Success: site_data.json saved.")
