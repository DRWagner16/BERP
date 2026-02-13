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

# --- 3. SANITIZE ---
if 'Company' in df.columns:
    df = df.drop(columns=['Company'])

# --- 4. CLEAN & CALCULATE ---
numeric_cols = ['Gas Savings (MMBtu/yr)', 'Electric Savings (kWh/yr)', 'Total Cost Savings', 'Implementation Costs', 'Percent Progress']
for col in numeric_cols:
    if col not in df.columns: df[col] = 0
    df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$,%]', '', regex=True), errors='coerce').fillna(0)

# --- NEW: EXTRACT YEAR ---
if 'Date of Assessment' in df.columns:
    # Convert to datetime objects, handling errors
    df['Date_Obj'] = pd.to_datetime(df['Date of Assessment'], errors='coerce')
    # Extract the Year (e.g., 2025)
    df['Year'] = df['Date_Obj'].dt.year.fillna(0).astype(int)
    # Drop the temporary object column to keep JSON clean
    df = df.drop(columns=['Date_Obj'])
else:
    df['Year'] = 2025 # Default fallback

# Clean FIPS
if 'FIPS' in df.columns:
    df['FIPS'] = pd.to_numeric(df['FIPS'], errors='coerce').fillna(0).astype(int).astype(str)
    df['FIPS'] = df['FIPS'].str.zfill(5)

# Emissions Factors
FACTORS = { 'Gas_CO2_lb': 117.0, 'Elec_CO2_lb': 1.5, 'Gas_NOx_lb': 0.092, 'Elec_NOx_lb': 0.001 }

# Calculate
df['Gas_CO2_lb'] = df['Gas Savings (MMBtu/yr)'] * FACTORS['Gas_CO2_lb']
df['Elec_CO2_lb'] = df['Electric Savings (kWh/yr)'] * FACTORS['Elec_CO2_lb']
df['Total_CO2_Tons'] = (df['Gas_CO2_lb'] + df['Elec_CO2_lb']) / 2000.0

# --- 5. EXPORT ---
json_output = df.to_json(orient='records')
with open('site_data.json', 'w') as f:
    f.write(json_output)
    print("Success: site_data.json saved.")
