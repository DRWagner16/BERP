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
    # Strip whitespace from headers to avoid "Column not found" errors
    df.columns = df.columns.astype(str).str.strip()
    print("Data loaded successfully.")
except Exception as e:
    print(f"Error loading sheet: {e}")
    exit(1)

# --- 3. SANITIZE ---
if 'Company' in df.columns:
    df = df.drop(columns=['Company'])

# --- 4. CLEAN DATA ---
# List of ALL columns we need to turn into numbers
# (We include both LOW and HIGH columns here)
target_cols = [
    'Gas Savings (MMBtu/yr)', 
    'Electric Savings (kWh/yr)', 
    'Total Cost Savings', 
    'Implementation Costs', 
    'Percent Progress',
    # CO2 Columns
    'Gas Equivalent CO2 Savings (lb/yr)',
    'Electricity Equivalent CO2 Savings - LOW (lb/year)',
    'Electricity Equivalent CO2 Savings - HIGH (lb/year)',
    # NOx Columns
    'Gas NOx Savings (lb/yr)',
    'Electricity NOx Savings LOW (lb/yr)',
    'Electricity NOx Savings HIGH (lb/yr)',
    # SO2 Columns
    'Gas So2 Savings',
    'Electricity SO2 Savings LOW',
    'Electricity SO2 Savings HIGH (lb/yr)',
    # PM2.5 Columns
    'Gas PM2.5 Savings',
    'Electricity PM2.5 Savings LOW (lb/yr)',
    'Electricity PM2.5 Savings HIGH (lb/yr)'
]

for col in target_cols:
    if col not in df.columns:
        # Create missing columns with 0 to prevent crashes
        df[col] = 0
    # Convert to numeric, forcing errors to NaN then filling with 0
    df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$,%]', '', regex=True), errors='coerce').fillna(0)

# --- 5. CALCULATE AVERAGES & TOTALS ---

# --- CO2 (Tons) ---
# Formula: (Gas + ((Elec_Low + Elec_High) / 2)) / 2000
df['Elec_CO2_Avg'] = (df['Electricity Equivalent CO2 Savings - LOW (lb/year)'] + 
                      df['Electricity Equivalent CO2 Savings - HIGH (lb/year)']) / 2
df['Total_CO2_Tons'] = (df['Gas Equivalent CO2 Savings (lb/yr)'] + df['Elec_CO2_Avg']) / 2000.0

# --- NOx (lbs) ---
df['Elec_NOx_Avg'] = (df['Electricity NOx Savings LOW (lb/yr)'] + 
                      df['Electricity NOx Savings HIGH (lb/yr)']) / 2
df['Total_NOx_lb'] = df['Gas NOx Savings (lb/yr)'] + df['Elec_NOx_Avg']

# --- SO2 (lbs) ---
df['Elec_SO2_Avg'] = (df['Electricity SO2 Savings LOW'] + 
                      df['Electricity SO2 Savings HIGH (lb/yr)']) / 2
df['Total_SO2_lb'] = df['Gas So2 Savings'] + df['Elec_SO2_Avg']

# --- PM2.5 (lbs) ---
df['Elec_PM25_Avg'] = (df['Electricity PM2.5 Savings LOW (lb/yr)'] + 
                       df['Electricity PM2.5 Savings HIGH (lb/yr)']) / 2
df['Total_PM25_lb'] = df['Gas PM2.5 Savings'] + df['Elec_PM25_Avg']

# --- Equivalency (Cars) ---
# 5.07 US Tons CO2 per vehicle per year
df['Cars_Equivalent'] = df['Total_CO2_Tons'] / 5.07

# --- DATE HANDLING (For Year Filter) ---
if 'Date of Assessment' in df.columns:
    date_col = df['Date of Assessment'].astype(str)
    df['Date_Obj'] = pd.to_datetime(date_col, errors='coerce')
    df['Year'] = df['Date_Obj'].dt.year.fillna(0).astype(int)
    
    mask_zero = (df['Year'] == 0)
    if mask_zero.any():
        df.loc[mask_zero, 'Year'] = pd.to_numeric(
            date_col[mask_zero].str.extract(r'(\d{4})')[0], 
            errors='coerce'
        ).fillna(0).astype(int)
    
    if 'Date_Obj' in df.columns:
        df = df.drop(columns=['Date_Obj'])
else:
    df['Year'] = 0

# --- FIPS HANDLING (For Map) ---
if 'FIPS' in df.columns:
    df['FIPS'] = pd.to_numeric(df['FIPS'], errors='coerce').fillna(0).astype(int).astype(str)
    df['FIPS'] = df['FIPS'].str.zfill(5)

# --- 6. EXPORT ---
json_output = df.to_json(orient='records')
with open('site_data.json', 'w') as f:
    f.write(json_output)
    print("Success: site_data.json saved.")
