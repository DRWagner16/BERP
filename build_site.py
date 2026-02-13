import os
import json
import gspread
import pandas as pd
import plotly.express as px
from google.oauth2.service_account import Credentials
from datetime import datetime

# --- CONFIGURATION ---
SHEET_NAME = "BERP AR Tracking"  # <--- UPDATED!

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
# Drop Company Name immediately
if 'Company' in df.columns:
    df = df.drop(columns=['Company'])
    print("Privacy Check: 'Company' column removed.")

# --- 4. DATA PREPARATION ---
numeric_cols = [
    'Total Cost Savings', 
    'Gas Savings (MMBtu/yr)', 
    'Electric Savings (kWh/yr)',
    'Equivalent CO2 (ton/yr)',
    'Implementation Costs',
    'Percent Progress' # Added this for the new chart
]

# Ensure columns exist and are numeric
for col in numeric_cols:
    if col not in df.columns:
        df[col] = 0
    # Clean currency/comma/% strings -> numbers
    df[col] = pd.to_numeric(df[col].astype(str).str.replace('$', '').str.replace(',', '').str.replace('%', ''), errors='coerce').fillna(0)

# Create two dataframes: 
# 1. Implemented (Completed) for the main stats
# 2. In Progress (Active) for the pipeline chart
if 'Implemented? Yes/No/In Progress' in df.columns:
    implemented_df = df[df['Implemented? Yes/No/In Progress'] == 'Yes']
    progress_df = df[df['Implemented? Yes/No/In Progress'] == 'In Progress']
else:
    implemented_df = df
    progress_df = pd.DataFrame()

# --- 5. VISUALIZATIONS ---

# Chart A: Implementation Status (Overview)
if 'Implemented? Yes/No/In Progress' in df.columns:
    status_counts = df['Implemented? Yes/No/In Progress'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    fig_status = px.pie(status_counts, values='Count', names='Status', 
                        title='Program Overview: Implementation Status',
                        color_discrete_sequence=px.colors.qualitative.Pastel)
    chart_status_html = fig_status.to_html(full_html=False, include_plotlyjs='cdn')
else:
    chart_status_html = "<p>Data missing</p>"

# Chart B: CO2 by County (Implemented Only)
if 'County' in implemented_df.columns:
    county_group = implemented_df.groupby('County')['Equivalent CO2 (ton/yr)'].sum().reset_index()
    fig_county = px.bar(county_group, x='County', y='Equivalent CO2 (ton/yr)', 
                        title='Realized CO2 Reduction by County (Completed Projects)',
                        color='Equivalent CO2 (ton/yr)', color_continuous_scale='Teal')
    chart_county_html = fig_county.to_html(full_html=False, include_plotlyjs='cdn')
else:
    chart_county_html = "<p>Data missing</p>"

# Chart C: The Pipeline (In Progress items)
if not progress_df.empty and 'Percent Progress' in progress_df.columns:
    # Sort by progress so the bar chart looks organized
    progress_df = progress_df.sort_values('Percent Progress', ascending=True)
    
    fig_progress = px.bar(progress_df, x='Percent Progress', y='Recommendation Name', 
                          orientation='h', # Horizontal bars
                          title='Project Pipeline: Recommendations In Progress',
                          hover_data=['Report Number', 'Equivalent CO2 (ton/yr)'],
                          color='Percent Progress', color_continuous_scale='Bluyl')
    
    # Clean up the y-axis labels if they are too long
    fig_progress.update_layout(yaxis={'categoryorder':'total ascending'})
    chart_progress_html = fig_progress.to_html(full_html=False, include_plotlyjs='cdn')
else:
    chart_progress_html = "<p>No active projects currently in progress.</p>"


# --- 6. BUILD REPORT ---
total_co2 = implemented_df['Equivalent CO2 (ton/yr)'].sum()
total_dollars = implemented_df['Total Cost Savings'].sum()
total_projects = len(implemented_df)

# Top 10 Table columns
table_cols = ['Report Number', 'Recommendation Name', 'Total Cost Savings', 'Equivalent CO2 (ton/yr)']
valid_cols = [c for c in table_cols if c in implemented_df.columns]

html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>BERP AR Tracking Dashboard</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: #f8f9fa; }}
        .header {{ background-color: #1a5276; color: white; padding: 20px; text-align: center; }}
        .container {{ max-width: 1200px; margin: 20px auto; padding: 20px; }}
        .metrics-row {{ display: flex; justify-content: space-between; margin-bottom: 30px; }}
        .metric-card {{ background: white; padding: 20px; border-radius: 8px; width: 30%; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; border-bottom: 4px solid #1a5276; }}
        .metric-value {{ font-size: 2.5em; font-weight: bold; color: #1a5276; }}
        .metric-label {{ color: #7f8c8d; font-size: 1.1em; }}
        .chart-container {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 30px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
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
            <div class="metric-value">{total_projects}</div>
            <div class="metric-label">Completed Projects</div>
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
        {chart_progress_html}
    </div>

    <div class="chart-container">
        <h3>Top Realized Savings (Completed)</h3>
        {implemented_df.sort_values('Equivalent CO2 (ton/yr)', ascending=False).head(10)[valid_cols].to_html(classes='table', index=False)}
    </div>
</div>

</body>
</html>
"""

with open('index.html', 'w') as f:
    f.write(html_content)
    print("Report generated successfully.")
