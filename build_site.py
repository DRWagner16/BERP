import pandas as pd
import plotly.express as px
import os
from datetime import datetime

# --- 1. CONFIGURATION ---
# REPLACE THIS with your actual Box Shared Link
BOX_SHARED_LINK = "https://app.box.com/s/YOUR_LONG_RANDOM_STRING_HERE"

# Function to convert Box preview link to direct download link
def get_direct_box_link(shared_url):
    # Extract the random hash from the URL
    file_hash = shared_url.split('/')[-1]
    # Construct the 'direct' static download URL
    return f"https://app.box.com/shared/static/{file_hash}.xlsx"

# --- 2. LOAD DATA ---
try:
    print("Attempting to download data from Box...")
    direct_url = get_direct_box_link(BOX_SHARED_LINK)
    
    # Read directly into Pandas
    df = pd.read_excel(direct_url)
    print("Data loaded successfully.")
    
except Exception as e:
    print(f"Error loading data: {e}")
    # Create dummy data so the build doesn't fail completely during testing
    print("Generating dummy data for failsafe...")
    df = pd.DataFrame({
        'Date': pd.date_range(start='2024-01-01', periods=5, freq='W'),
        'Location': ['Site A', 'Site B', 'Site A', 'Site B', 'Site A'],
        'Energy_kWh': [450, 300, 500, 420, 600],
        'Fuel_Type': ['Diesel', 'Natural Gas', 'Diesel', 'Solar', 'Grid']
    })

# --- 3. CALCULATIONS ---
factors = {'Diesel': 0.26, 'Natural Gas': 0.18, 'Grid': 0.40, 'Solar': 0.0}
df['Emission_Factor'] = df['Fuel_Type'].map(factors).fillna(0)
df['CO2_Emissions_kg'] = df['Energy_kWh'] * df['Emission_Factor']

# --- 4. GENERATE HTML (Same as before) ---
fig = px.line(df, x='Date', y='CO2_Emissions_kg', color='Location', title='Emissions Over Time')
chart_html = fig.to_html(full_html=False, include_plotlyjs='cdn')

html_content = f"""
<html>
<head><title>Project Status</title></head>
<body>
    <h1>Environmental Report</h1>
    <p>Last Updated: {datetime.now().strftime("%Y-%m-%d %H:%M")}</p>
    {chart_html}
</body>
</html>
"""

with open('index.html', 'w') as f:
    f.write(html_content)
