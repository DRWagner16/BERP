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
    
    # Force headers to strings and strip whitespace
    df.columns = df.columns.astype(str).str.strip()
    print("Data loaded successfully.")
    
except Exception as e:
    print(f"Error loading sheet: {e}")
    exit(1)

# --- 3. SANITIZE DATA (PRIVACY PROTECTION) ---
# CRITICAL: Drop the Company Name so it never reaches the public HTML file
if 'Company' in df.columns:
    df = df.drop(columns=['Company'])
    print("Privacy Check: 'Company' column removed.")

# --- 4. DATA PREPARATION ---
numeric_cols = [
    'Total Cost Savings', 
    'Gas Savings (MMBtu/yr)', 
    'Electric Savings (kWh/yr)',
    'Equivalent CO2 (ton/yr)',
    'Implementation Costs'
]

# Ensure columns exist and are numeric
for col in numeric_cols:
    if col not in df.columns:
        df[col] = 0
    # Clean currency/comma strings -> numbers
    df[col] = pd.to_numeric(df[col].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce').fillna(0)

# Filter for Implemented projects
if 'Implemented? Yes/No/In Progress' in df.columns:
    implemented_df = df[df['Implemented? Yes/No/In Progress'].isin(['Yes', 'In Progress'])]
else:
    implemented_df = df

# --- 5. VISUALIZATIONS ---

# Chart A: Implementation Status
if 'Implemented? Yes/No/In Progress' in df.columns:
    status_counts = df['Implemented? Yes/No/In Progress'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    fig_status = px.pie(status_counts, values='Count', names='Status', 
                        title='Implementation Status',
                        color_discrete_sequence=px.colors.qualitative.Set2)
    chart_status_html = fig_status.to_html(full_html=False, include_plotlyjs='cdn')
else:
    chart_status_html = "<p>Data missing</p>"

# Chart B: CO2 by County
if 'County' in implemented_df.columns:
    county_group = implemented_df.groupby('County')['Equivalent CO2 (ton/yr)'].sum().reset_index()
    fig_county = px.bar(county_group, x='County', y='Equivalent CO2 (ton/yr)', 
                        title='CO2 Reduction by County',
                        color='Equivalent CO2 (ton/yr)', color_continuous_scale='Teal')
    chart_county_html = fig_county.to_html(full_html=False, include_plotlyjs='cdn')
else:
    chart_county_html = "<p>Data missing</p>"

# Chart C: Cost vs Savings
if 'Total Cost Savings' in implemented_df.columns:
    fig_cost = px.scatter(implemented_df, x='Implementation Costs', y='Total Cost Savings', 
                          size='Equivalent CO2 (ton/yr)', 
                          hover_name='Recommendation Name',
                          title='Cost vs. Savings (Bubble size = CO2 Impact)')
    chart_cost_html = fig_cost.to_html(full_html=False, include_plotlyjs='cdn')
else:
    chart_cost_html = "<p>Data missing</p>"

# --- 6. BUILD REPORT ---
total_co2 = implemented_df['Equivalent CO2 (ton/yr)'].sum()
total_dollars = implemented_df['Total Cost Savings'].sum()
total_projects = len(implemented_df)

# Prepare the Top 10 Table (Using Report Number instead of Company)
table_cols = ['Report Number', 'Recommendation Name', 'Total Cost Savings', 'Equivalent CO2 (ton/yr)']
# Ensure these columns actually exist before trying to display them
valid_cols = [c for c in table_cols if c in implemented_df.columns]

html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Program Impact Report</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: #f8f9fa; }}
        .header {{ background-color: #2c3e50; color: white; padding: 20px; text-align: center; }}
        .container {{ max-width: 1200px; margin: 20px auto; padding: 20px; }}
        .metrics-row {{ display: flex; justify-content: space-between; margin-bottom: 30px; }}
        .metric-card {{ background: white; padding: 20px; border-radius: 8px; width: 30%; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; }}
        .metric-value {{ font-size: 2.5em; font-weight: bold; color: #27ae60; }}
        .metric-label {{ color: #7f8c8d; font-size: 1.1em; }}
        .chart-container {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 30px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        th, td {{ padding: 12px; border-bottom: 1px solid #ddd; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
    </style>
</head>
<body>

<div class="header">
    <h1>Energy & Emissions Program Dashboard</h1>
    <p>Snapshot Date: {datetime.now().strftime("%Y-%m-%d")}</p>
</div>

<div class="container">
    <div class="metrics-row">
        <div class="metric-card">
            <div class="metric-value">{total_projects}</div>
            <div class="metric-label">Active Projects</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{total_co2:,.1f}</div>
            <div class="metric-label">Tons CO2 Reduced</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">${total_dollars:,.0f}</div>
            <div class="metric-label">Annual Cost Savings</div>
        </div>
    </div>

    <div style="display: flex; gap: 20px;">
        <div class="chart-container" style="flex: 1;">
            {chart_status_html}
        </div>
        <div class="chart-container" style="flex: 1;">
            {chart_county_html}
        </div>
    </div>

    <div class="chart-container">
        {chart_cost_html}
    </div>

    <div class="chart-container">
        <h3>Top High-Impact Projects</h3>
        {implemented_df.sort_values('Equivalent CO2 (ton/yr)', ascending=False).head(10)[valid_cols].to_html(classes='table', index=False)}
    </div>
</div>

</body>
</html>
"""

with open('index.html', 'w') as f:
    f.write(html_content)
    print("Report generated successfully.")
