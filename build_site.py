import os
import json
import gspread
import pandas as pd
import plotly.express as px
from google.oauth2.service_account import Credentials
from datetime import datetime

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
    df.columns = df.columns.astype(str).str.strip() # Clean headers
    print("Data loaded successfully.")
except Exception as e:
    print(f"Error loading sheet: {e}")
    exit(1)

# --- 3. SANITIZE & CLEAN INPUTS ---
if 'Company' in df.columns:
    df = df.drop(columns=['Company'])

# Define the INPUT columns we need for math
input_cols = ['Gas Savings (MMBtu/yr)', 'Electric Savings (kWh/yr)', 'Total Cost Savings', 'Implementation Costs', 'Percent Progress']

for col in input_cols:
    if col not in df.columns:
        df[col] = 0
    # Convert to numeric
    df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$,%]', '', regex=True), errors='coerce').fillna(0)

# --- 4. PERFORM CALCULATIONS (THE SCIENCE) ---
# Emission Factors (You can update these specific values!)
FACTORS = {
    'Gas_CO2_lb_per_MMBtu': 117.0,
    'Elec_CO2_lb_per_kWh': 1.5,   # Adjust based on eGRID subregion
    'Gas_NOx_lb_per_MMBtu': 0.092,
    'Elec_NOx_lb_per_kWh': 0.001
}

# Calculate Gas CO2 (lb/yr)
df['Gas Equivalent CO2 Savings (lb/yr)'] = df['Gas Savings (MMBtu/yr)'] * FACTORS['Gas_CO2_lb_per_MMBtu']

# Calculate Electric CO2 (lb/yr)
df['Electricity Equivalent CO2 Savings - HIGH (lb/year)'] = df['Electric Savings (kWh/yr)'] * FACTORS['Elec_CO2_lb_per_kWh']

# Calculate TOTAL CO2 in TONS (The main metric)
# Formula: (Gas lbs + Elec lbs) / 2000
df['Equivalent CO2 (ton/yr)'] = (
    df['Gas Equivalent CO2 Savings (lb/yr)'] + 
    df['Electricity Equivalent CO2 Savings - HIGH (lb/year)']
) / 2000.0

# Calculate NOx (Example of Criteria Pollutant)
df['Gas NOx Savings (lb/yr)'] = df['Gas Savings (MMBtu/yr)'] * FACTORS['Gas_NOx_lb_per_MMBtu']
df['Electricity NOx Savings HIGH (lb/yr)'] = df['Electric Savings (kWh/yr)'] * FACTORS['Elec_NOx_lb_per_kWh']

# --- 5. PREPARE REPORT DATA ---
if 'Implemented? Yes/No/In Progress' in df.columns:
    implemented_df = df[df['Implemented? Yes/No/In Progress'] == 'Yes']
    progress_df = df[df['Implemented? Yes/No/In Progress'] == 'In Progress']
else:
    implemented_df = df
    progress_df = pd.DataFrame()

# --- 6. VISUALIZATIONS ---

# Chart A: CO2 Savings by Utility Type (Pie)
# We sum up the calculated columns
gas_co2_total = implemented_df['Gas Equivalent CO2 Savings (lb/yr)'].sum()
elec_co2_total = implemented_df['Electricity Equivalent CO2 Savings - HIGH (lb/year)'].sum()

fig_source = px.pie(
    names=['Natural Gas Savings', 'Electricity Savings'], 
    values=[gas_co2_total, elec_co2_total],
    title='Sources of CO2 Reduction (lb)',
    color_discrete_sequence=['#e67e22', '#3498db'] # Orange for Gas, Blue for Elec
)
chart_source_html = fig_source.to_html(full_html=False, include_plotlyjs='cdn')

# Chart B: CO2 by County
if 'County' in implemented_df.columns:
    county_group = implemented_df.groupby('County')['Equivalent CO2 (ton/yr)'].sum().reset_index()
    fig_county = px.bar(county_group, x='County', y='Equivalent CO2 (ton/yr)', 
                        title='Total CO2 Reduction by County (Tons)',
                        color='Equivalent CO2 (ton/yr)', color_continuous_scale='Teal')
    chart_county_html = fig_county.to_html(full_html=False, include_plotlyjs='cdn')
else:
    chart_county_html = "<p>Data missing</p>"

# Chart C: Pipeline
if not progress_df.empty and 'Percent Progress' in progress_df.columns:
    progress_df = progress_df.sort_values('Percent Progress', ascending=True)
    fig_progress = px.bar(progress_df, x='Percent Progress', y='Recommendation Name', 
                          orientation='h', title='Project Pipeline (In Progress)',
                          color='Percent Progress', color_continuous_scale='Bluyl')
    fig_progress.update_layout(yaxis={'categoryorder':'total ascending'})
    chart_progress_html = fig_progress.to_html(full_html=False, include_plotlyjs='cdn')
else:
    chart_progress_html = "<p>No active projects.</p>"

# --- 7. GENERATE HTML ---
total_co2 = implemented_df['Equivalent CO2 (ton/yr)'].sum()
total_dollars = implemented_df['Total Cost Savings'].sum()
total_nox = (implemented_df['Gas NOx Savings (lb/yr)'].sum() + implemented_df['Electricity NOx Savings HIGH (lb/yr)'].sum())

html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>BERP AR Tracking Dashboard</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: #f8f9fa; }}
        .header {{ background-color: #1a5276; color: white; padding: 20px; text-align: center; }}
        .container {{ max-width: 1200px; margin: 20px auto; padding: 20px; }}
        
        /* Metric Cards */
        .metrics-row {{ display: flex; justify-content: space-between; margin-bottom: 30px; }}
        .metric-card {{ background: white; padding: 20px; border-radius: 8px; width: 23%; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; border-bottom: 4px solid #1a5276; }}
        .metric-value {{ font-size: 2em; font-weight: bold; color: #1a5276; }}
        .metric-label {{ color: #7f8c8d; font-size: 0.9em; }}
        
        .chart-container {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 30px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 12px; border-bottom: 1px solid #ddd; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
    </style>
</head>
<body>

<div class="header">
    <h1>BERP AR Tracking Dashboard</h1>
    <p>Snapshot Date: {datetime.now().strftime("%Y-%m-%d")}</p>
</div>

<div class="container">
    <div class="metrics-row">
        <div class="metric-card">
            <div class="metric-value">{len(implemented_df)}</div>
            <div class="metric-label">Completed Projects</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{total_co2:,.1f}</div>
            <div class="metric-label">Tons CO2 Reduced</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{total_nox:,.1f}</div>
            <div class="metric-label">Lbs NOx Reduced</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">${total_dollars:,.0f}</div>
            <div class="metric-label">Annual Savings</div>
        </div>
    </div>

    <div style="display: flex; gap: 20px;">
        <div class="chart-container" style="flex: 1;">
            {chart_source_html}
        </div>
        <div class="chart-container" style="flex: 1;">
            {chart_county_html}
        </div>
    </div>

    <div class="chart-container">
        {chart_progress_html}
    </div>

    <div class="chart-container">
        <h3>Top Realized Savings (Calculated)</h3>
        {implemented_df.sort_values('Equivalent CO2 (ton/yr)', ascending=False).head(10)[['Report Number', 'Recommendation Name', 'Equivalent CO2 (ton/yr)', 'Total Cost Savings']].to_html(classes='table', index=False)}
    </div>
</div>

</body>
</html>
"""

with open('index.html', 'w') as f:
    f.write(html_content)
    print("Report generated successfully.")
