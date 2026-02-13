import os
import json
import gspread
import pandas as pd
import plotly.express as px
from google.oauth2.service_account import Credentials
from datetime import datetime

# --- CONFIGURATION ---
# REPLACE THIS with the exact name of your Google Sheet
SHEET_NAME = "BERP AR Tracking" 

# --- 1. AUTHENTICATE ---
# We get the secret "key" from GitHub settings
raw_creds = os.environ.get("GCP_SERVICE_ACCOUNT")

if not raw_creds:
    print("Error: GCP_SERVICE_ACCOUNT secret is missing.")
    exit(1)

# Set up the connection
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = json.loads(raw_creds)
creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
client = gspread.authorize(creds)

# --- 2. LOAD DATA ---
try:
    print(f"Opening Google Sheet: '{SHEET_NAME}'...")
    sheet = client.open(SHEET_NAME).sheet1 # Opens the first tab
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    print("Data loaded successfully.")
    
except Exception as e:
    print(f"Error loading sheet: {e}")
    exit(1)

# --- 3. CALCULATIONS ---
# Clean up numbers (remove commas if someone typed "1,000")
if 'Energy_kWh' in df.columns:
    df['Energy_kWh'] = pd.to_numeric(df['Energy_kWh'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

# Define Pollution Factors (kg CO2 per kWh)
factors = {
    'Diesel': 0.26, 
    'Natural Gas': 0.18, 
    'Grid': 0.40, 
    'Solar': 0.0
}

# Apply Math
if 'Fuel_Type' in df.columns and 'Energy_kWh' in df.columns:
    df['Emission_Factor'] = df['Fuel_Type'].map(factors).fillna(0)
    df['CO2_Emissions_kg'] = df['Energy_kWh'] * df['Emission_Factor']

# --- 4. BUILD WEBSITE ---
# Generate Chart
if 'Date' in df.columns and 'CO2_Emissions_kg' in df.columns:
    fig = px.line(df, x='Date', y='CO2_Emissions_kg', color='Location', markers=True, 
                  title='Pollution Dispersion Over Time')
    chart_html = fig.to_html(full_html=False, include_plotlyjs='cdn')
else:
    chart_html = "<p>Waiting for data... (Ensure columns 'Date' and 'CO2_Emissions_kg' exist)</p>"

# Create HTML File
html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Environmental Report</title>
    <style>
        body {{ font-family: sans-serif; max-width: 1000px; margin: auto; padding: 20px; background: #f4f6f8; }}
        .card {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }}
        h1 {{ color: #2c3e50; }}
    </style>
</head>
<body>
    <div class="card">
        <h1>Project Environmental Dashboard</h1>
        <p>Live Data from Google Sheets | Updated: {datetime.now().strftime("%Y-%m-%d %H:%M UTC")}</p>
    </div>
    <div class="card">
        {chart_html}
    </div>
    <div class="card">
        <h3>Recent Data Log</h3>
        {df.tail(5).to_html(classes='table', border=0, index=False)}
    </div>
</body>
</html>
"""

with open('index.html', 'w') as f:
    f.write(html_content)
    print("Success: index.html generated.")
