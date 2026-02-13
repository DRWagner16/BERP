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
    
    # Clean headers (remove spaces)
    df.columns = df.columns.astype(str).str.strip()
    print("Data loaded successfully.")
except Exception as e:
    print(f"Error loading sheet: {e}")
    exit(1)

# --- 3. SANITIZE (PRIVACY) ---
# CRITICAL: Drop 'Company' so it is never saved to the public JSON file
if 'Company' in df.columns:
    df = df.drop(columns=['Company'])
    print("Privacy Check: 'Company' column removed.")

# --- 4. CLEAN & CALCULATE ---
# Define columns that need to be numbers
numeric_cols = [
    'Gas Savings (MMBtu/yr)', 
    'Electric Savings (kWh/yr)', 
    'Total Cost Savings', 
    'Implementation Costs', 
    'Percent Progress'
]

# Clean numeric columns (remove $, %, commas)
for col in numeric_cols:
    if col not in df.columns:
        df[col] = 0
    df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$,%]', '', regex=True), errors='coerce').fillna(0)

# Clean FIPS Column (Critical for Map)
# We ensure it is a 5-digit string (e.g., "49035")
if 'FIPS' in df.columns:
    # Convert to number first to remove decimals like 49035.0, then to int, then string
    df['FIPS'] = pd.to_numeric(df['FIPS'], errors='coerce').fillna(0).astype(int).astype(str)
    # Pad with leading zeros (e.g., "1001" -> "01001")
    df['FIPS'] = df['FIPS'].str.zfill(5)
else:
    print("Warning: 'FIPS' column missing. Map may not colorize correctly.")

# Emission Factors (Estimates)
FACTORS = {
    'Gas_CO2_lb': 117.0,   # lb CO2 per MMBtu
    'Elec_CO2_lb': 1.5,    # lb CO2 per kWh (High estimate for UT)
    'Gas_NOx_lb': 0.092,
    'Elec_NOx_lb': 0.001
}

# Perform Calculations
df['Gas_CO2_lb'] = df['Gas Savings (MMBtu/yr)'] * FACTORS['Gas_CO2_lb']
df['Elec_CO2_lb'] = df['Electric Savings (kWh/yr)'] * FACTORS['Elec_CO2_lb']
df['Total_CO2_Tons'] = (df['Gas_CO2_lb'] + df['Elec_CO2_lb']) / 2000.0

df['Total_NOx_lb'] = (
    (df['Gas Savings (MMBtu/yr)'] * FACTORS['Gas_NOx_lb']) + 
    (df['Electric Savings (kWh/yr)'] * FACTORS['Elec_NOx_lb'])
)

# --- 5. EXPORT TO JSON ---
# Save the clean, calculated data to a JSON file
json_output = df.to_json(orient='records')

with open('site_data.json', 'w') as f:
    f.write(json_output)
    print("Success: site_data.json saved.")
