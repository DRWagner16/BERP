import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
from datetime import datetime

# --- 1. SECURE CONFIGURATION ---
# We retrieve the URL from the environment variable we set in GitHub Secrets
BOX_SHARED_LINK = os.environ.get("BOX_URL")

if not BOX_SHARED_LINK:
    raise ValueError("Error: BOX_URL environment variable is missing!")

# Function to convert Box preview link to direct download link
def get_direct_box_link(shared_url):
    # Split the URL to get the hash (the part after /s/)
    if "/s/" in shared_url:
        file_hash = shared_url.split("/s/")[-1].split('/')[0]
    else:
        # Fallback if URL format is different
        file_hash = shared_url.split('/')[-1]
        
    return f"https://app.box.com/shared/static/{file_hash}.xlsx"

# --- 2. LOAD DATA ---
try:
    print("Attempting to download data from secure Box link...")
    direct_url = get_direct_box_link(BOX_SHARED_LINK)
    
    # Read Excel directly from the URL
    df = pd.read_excel(direct_url)
    print("Data loaded successfully.")

except Exception as e:
    print(f"CRITICAL ERROR: Could not load data. {e}")
    # In a real automated system, we might want to fail here so you get an email alert
    exit(1) 

# --- 3. CALCULATIONS (Pollution & Energy) ---
# Example Logic: Emissions = Energy * Factor
factors = {
    'Diesel': 0.26, 
    'Natural Gas': 0.18, 
    'Grid': 0.40, 
    'Solar': 0.0
}

# Ensure columns exist (basic validation)
required_cols = ['Fuel_Type', 'Energy_kWh', 'Date', 'Location']
if not all(col in df.columns for col in required_cols):
    raise ValueError(f"Excel file is missing required columns: {required_cols}")

df['Emission_Factor'] = df['Fuel_Type'].map(factors).fillna(0)
df['CO2_Emissions_kg'] = df['Energy_kWh'] * df['Emission_Factor']

# --- 4. GENERATE HTML REPORT ---
# Create Figure 1: Emissions
fig = px.line(
    df, x='Date', y='CO2_Emissions_kg', color='Location', markers=True,
    title='CO2 Emissions Over Time (kg)'
)
chart_html = fig.to_html(full_html=False, include_plotlyjs='cdn')

# Create Figure 2: Energy Mix
fig_pie = px.pie(df, names='Fuel_Type', values='Energy_kWh', title='Energy Mix')
chart_pie = fig_pie.to_html(full_html=False, include_plotlyjs='cdn')

# HTML Template
html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Environmental Dashboard</title>
    <style>
        body {{ font-family: sans-serif; margin: 40px; background: #f0f2f5; }}
        .card {{ background: white; padding: 20px; margin-bottom: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        h1 {{ color: #333; }}
        .footer {{ font-size: 0.8em; color: #666; margin-top: 20px; }}
    </style>
</head>
<body>
    <div class="card">
        <h1>Project Status Report</h1>
        <p><b>Data Source:</b> Box (Secure Link)</p>
        <p><b>Last Updated:</b> {datetime.now().strftime("%Y-%m-%d %H:%M")} UTC</p>
    </div>
    
    <div class="card">
        {chart_html}
    </div>

    <div class="card">
        {chart_pie}
    </div>
    
    <div class="footer">
        Generated automatically by GitHub Actions
    </div>
</body>
</html>
"""

with open('index.html', 'w') as f:
    f.write(html_content)
    print("Report generated successfully: index.html")
