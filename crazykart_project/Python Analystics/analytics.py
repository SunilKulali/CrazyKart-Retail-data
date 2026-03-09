"""
CrazyKart — Python ML Analytics Pipeline
FILE: analytics.py

WHERE TO RUN:
  Option A (Local): Install Python 3.9+, then:
    pip install pandas prophet scikit-learn openpyxl gspread google-auth matplotlib seaborn
    python analytics.py

  Option B (Google Colab — FREE, no setup):
    1. Go to https://colab.research.google.com
    2. Upload your CrazyKart_live_Sales.xlsx
    3. Paste this file into a notebook cell
    4. Run it

  Option C (Automate — GitHub Actions / PythonAnywhere):
    Push this file to GitHub and set up a workflow that runs daily.
"""

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# ════════════════════════════════════════════════════════
# CONFIG — Update paths and flags as needed
# ════════════════════════════════════════════════════════
EXCEL_FILE   = "CrazyKart_live_Sales.xlsx"  # Path to your Excel file
OUTPUT_FILE  = "CrazyKart_Analytics_Output.xlsx"
RUN_FORECAST = True   # Feature 07: Sales forecasting
RUN_ANOMALY  = True   # Feature 08: Anomaly detection
RUN_SCORING  = True   # Feature 09: Category scoring
FORECAST_DAYS = 30    # How many days ahead to forecast
# ════════════════════════════════════════════════════════


def load_data(filepath):
    """Load and clean the CrazyKart sales data."""
    df = pd.read_excel(filepath, sheet_name="Sheet1")
    df.columns = [c.strip().upper().replace(" ","_") for c in df.columns]

    # Parse date
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    df = df.dropna(subset=['DATE'])

    # Numeric types
    df['MRP']       = pd.to_numeric(df['MRP'],       errors='coerce').fillna(0)
    df['QTY']       = pd.to_numeric(df['QTY'],        errors='coerce').fillna(1)
    df['MRP_VALUE'] = pd.to_numeric(df['MRP_VALUE'],  errors='coerce').fillna(0)
    df['BILL_VAL']  = pd.to_numeric(df['BILL_VAL'],   errors='coerce').fillna(0)

    # Clean text
    df['RETAILER_NAME'] = df['RETAILER_NAME'].str.strip().str.title()
    df['BRAND']         = df['BRAND'].str.strip().str.upper()

    print(f"✅ Loaded {len(df):,} rows from {filepath}")
    print(f"   Date range: {df['DATE'].min().date()} → {df['DATE'].max().date()}")
    return df


# ════════════════════════════════════════════════════════
# FEATURE 07: SALES FORECASTING (Facebook Prophet)
# ════════════════════════════════════════════════════════
def run_sales_forecast(df, days_ahead=30):
    """
    Uses Facebook Prophet to forecast total daily revenue.
    Returns a DataFrame with ds (date), yhat (forecast),
    yhat_lower, yhat_upper columns.
    """
    try:
        from prophet import Prophet
    except ImportError:
        print("⚠️  Prophet not installed. Run: pip install prophet")
        return None

    print("\n🔮 Running Sales Forecast (Prophet)...")

    # Aggregate to daily total revenue
    daily = (df.groupby('DATE')['BILL_VAL']
               .sum()
               .reset_index()
               .rename(columns={'DATE':'ds', 'BILL_VAL':'y'}))

    daily = daily[daily['y'] > 0]  # Remove zero-sale days

    # Train model
    model = Prophet(
        yearly_seasonality=True,
        weekly_seasonality=True,
        daily_seasonality=False,
        changepoint_prior_scale=0.05,
        seasonality_prior_scale=10,
    )
    model.fit(daily)

    # Forecast
    future   = model.make_future_dataframe(periods=days_ahead)
    forecast = model.predict(future)

    result = forecast[['ds','yhat','yhat_lower','yhat_upper']].copy()
    result.columns = ['Date','Forecast_Revenue','Lower_Bound','Upper_Bound']
    result['Forecast_Revenue'] = result['Forecast_Revenue'].clip(lower=0).round(2)
    result['Lower_Bound']      = result['Lower_Bound'].clip(lower=0).round(2)
    result['Upper_Bound']      = result['Upper_Bound'].round(2)

    print(f"   ✅ Forecast generated for {days_ahead} days ahead")
    print(f"   Next 7 days avg forecast: ₹{result.tail(days_ahead).head(7)['Forecast_Revenue'].mean():,.0f}/day")
    return result


# ════════════════════════════════════════════════════════
# FEATURE 08: ANOMALY DETECTION
# ════════════════════════════════════════════════════════
def run_anomaly_detection(df):
    """
    Flags daily store revenue that is more than 2 standard deviations
    from the store's rolling 7-day average.
    Returns a DataFrame of flagged anomalies.
    """
    print("\n🚨 Running Anomaly Detection...")

    # Daily revenue per store
    store_daily = (df.groupby(['DATE','RETAILER_NAME','STORE_CODE'])['BILL_VAL']
                     .sum()
                     .reset_index())
    store_daily = store_daily.sort_values(['RETAILER_NAME','STORE_CODE','DATE'])

    # Rolling stats per store
    store_daily['rolling_mean'] = (
        store_daily.groupby(['RETAILER_NAME','STORE_CODE'])['BILL_VAL']
                   .transform(lambda x: x.rolling(7, min_periods=2).mean())
    )
    store_daily['rolling_std'] = (
        store_daily.groupby(['RETAILER_NAME','STORE_CODE'])['BILL_VAL']
                   .transform(lambda x: x.rolling(7, min_periods=2).std().fillna(1))
    )

    store_daily['z_score'] = (
        (store_daily['BILL_VAL'] - store_daily['rolling_mean'])
        / store_daily['rolling_std']
    )

    # Flag anomalies (z > 2 = spike, z < -2 = drop)
    anomalies = store_daily[store_daily['z_score'].abs() > 2.0].copy()
    anomalies['anomaly_type'] = anomalies['z_score'].apply(
        lambda z: 'SPIKE 📈' if z > 0 else 'DROP 📉'
    )
    anomalies['deviation_pct'] = (
        (anomalies['BILL_VAL'] - anomalies['rolling_mean'])
        / anomalies['rolling_mean'] * 100
    ).round(1)

    result = anomalies[['DATE','RETAILER_NAME','STORE_CODE','BILL_VAL',
                         'rolling_mean','z_score','anomaly_type','deviation_pct']].copy()
    result.columns = ['Date','Retailer','Store_Code','Actual_Revenue',
                       '7Day_Avg','Z_Score','Type','Deviation_%']
    result = result.sort_values('Date', ascending=False)

    print(f"   ✅ Found {len(result)} anomalous days across all stores")
    if len(result) > 0:
        print(f"   Spikes: {(result['Type'].str.contains('SPIKE')).sum()}  |  Drops: {(result['Type'].str.contains('DROP')).sum()}")
    return result


# ════════════════════════════════════════════════════════
# FEATURE 09: CATEGORY PERFORMANCE SCORING
# ════════════════════════════════════════════════════════
def run_category_scoring(df):
    """
    Scores each brand/category combination on:
      - Revenue contribution (40%)
      - Sell-through volume (30%)
      - Week-on-week growth (30%)
    Returns a ranked DataFrame with composite scores.
    """
    print("\n⭐ Running Category Performance Scoring...")

    # Revenue contribution per brand
    brand_rev = df.groupby('BRAND')['BILL_VAL'].sum()
    brand_qty = df.groupby('BRAND')['QTY'].sum()

    # Week-on-week growth (last 2 complete weeks)
    df['week_start'] = df['DATE'] - pd.to_timedelta(df['DATE'].dt.dayofweek, unit='d')
    recent_weeks = df['week_start'].sort_values().unique()[-3:]  # Last 3 weeks

    if len(recent_weeks) >= 2:
        w1 = df[df['week_start'] == recent_weeks[-2]].groupby('BRAND')['BILL_VAL'].sum()
        w2 = df[df['week_start'] == recent_weeks[-1]].groupby('BRAND')['BILL_VAL'].sum()
        wow_growth = ((w2 - w1) / w1.replace(0, np.nan) * 100).fillna(0)
    else:
        wow_growth = pd.Series(0, index=brand_rev.index)

    # Build scoring table
    scoring = pd.DataFrame({
        'Total_Revenue':    brand_rev,
        'Total_Units':      brand_qty,
        'WoW_Growth_%':     wow_growth,
    }).fillna(0)

    # Normalize each metric to 0–100
    def normalize(series):
        rng = series.max() - series.min()
        if rng == 0: return pd.Series(50, index=series.index)
        return ((series - series.min()) / rng * 100).round(1)

    scoring['Revenue_Score'] = normalize(scoring['Total_Revenue'])
    scoring['Volume_Score']  = normalize(scoring['Total_Units'])
    scoring['Growth_Score']  = normalize(scoring['WoW_Growth_%'])

    # Weighted composite score
    scoring['Composite_Score'] = (
        scoring['Revenue_Score'] * 0.40 +
        scoring['Volume_Score']  * 0.30 +
        scoring['Growth_Score']  * 0.30
    ).round(1)

    # Tier classification
    scoring['Tier'] = pd.cut(
        scoring['Composite_Score'],
        bins=[-1, 33, 66, 101],
        labels=['🔴 Underperforming', '🟡 Average', '🟢 Top Performer']
    )

    # Rank
    scoring['Rank'] = scoring['Composite_Score'].rank(ascending=False).astype(int)
    scoring = scoring.sort_values('Rank')
    scoring['Total_Revenue'] = scoring['Total_Revenue'].round(2)

    result = scoring.reset_index().rename(columns={'BRAND':'Brand'})
    result = result[['Rank','Brand','Total_Revenue','Total_Units',
                      'WoW_Growth_%','Composite_Score','Tier']]

    print(f"   ✅ Scored {len(result)} brands")
    print("\n   Top Performers:")
    for _, r in result[result['Tier'] == '🟢 Top Performer'].iterrows():
        print(f"      #{r['Rank']} {r['Brand']}: Score {r['Composite_Score']}, Rev ₹{r['Total_Revenue']:,.0f}")
    return result


# ════════════════════════════════════════════════════════
# FEATURE 15: STORE PERFORMANCE RANKING
# ════════════════════════════════════════════════════════
def run_store_ranking(df):
    """Ranks all stores by key KPIs."""
    print("\n🏆 Running Store Performance Ranking...")

    store_stats = df.groupby(['RETAILER_NAME','STORE_CODE']).agg(
        Total_Revenue   = ('BILL_VAL',  'sum'),
        Total_Units     = ('QTY',       'sum'),
        Avg_Bill_Value  = ('BILL_VAL',  'mean'),
        Total_Bills     = ('BILL_NO',   'count'),
    ).reset_index()

    store_stats['Avg_Bill_Value'] = store_stats['Avg_Bill_Value'].round(2)
    store_stats['Revenue_Rank']   = store_stats['Total_Revenue'].rank(ascending=False).astype(int)
    store_stats = store_stats.sort_values('Revenue_Rank')

    # Performance tier
    top33  = store_stats['Total_Revenue'].quantile(0.67)
    bot33  = store_stats['Total_Revenue'].quantile(0.33)
    store_stats['Performance'] = store_stats['Total_Revenue'].apply(
        lambda v: '🟢 Top Tier' if v >= top33 else ('🔴 Needs Attention' if v <= bot33 else '🟡 Mid Tier')
    )

    print(f"   ✅ Ranked {len(store_stats)} stores")
    return store_stats


# ════════════════════════════════════════════════════════
# FEATURE 16: INVENTORY DEPLETION ESTIMATE
# ════════════════════════════════════════════════════════
def run_inventory_forecast(df, opening_stock_per_sku=100):
    """
    Estimates days of stock remaining per style code.
    Uses last 7 days' average daily units sold.
    """
    print("\n📦 Running Inventory Depletion Forecast...")

    cutoff = df['DATE'].max() - pd.Timedelta(days=7)
    recent = df[df['DATE'] >= cutoff]

    sku_daily = (recent.groupby('STYLE_CODE')['QTY']
                        .sum()
                        .div(7)   # Average daily units
                        .reset_index())
    sku_daily.columns = ['Style_Code','Avg_Daily_Units']
    sku_daily['Avg_Daily_Units'] = sku_daily['Avg_Daily_Units'].round(2)

    # Join with total sold to estimate opening stock remaining
    total_sold = df.groupby('STYLE_CODE')['QTY'].sum().reset_index()
    total_sold.columns = ['Style_Code','Total_Units_Sold']

    inv = sku_daily.merge(total_sold, on='Style_Code')
    inv['Remaining_Stock']  = (opening_stock_per_sku - inv['Total_Units_Sold']).clip(lower=0)
    inv['Days_Remaining']   = (inv['Remaining_Stock'] / inv['Avg_Daily_Units'].replace(0, 0.01)).round(1)
    inv['Stock_Status']     = inv['Days_Remaining'].apply(
        lambda d: '🔴 Critical (<7 days)' if d < 7 else ('🟡 Low (7-14 days)' if d < 14 else '🟢 Adequate')
    )
    inv = inv.sort_values('Days_Remaining')

    print(f"   ✅ Inventory forecast for {len(inv)} SKUs")
    critical = inv[inv['Days_Remaining'] < 7]
    if len(critical) > 0:
        print(f"   ⚠️  {len(critical)} SKUs need immediate reorder!")
    return inv


# ════════════════════════════════════════════════════════
# SAVE ALL RESULTS TO EXCEL
# ════════════════════════════════════════════════════════
def save_results(results_dict, output_file):
    """Saves all analytics results to a multi-sheet Excel file."""
    print(f"\n💾 Saving results to {output_file}...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, df in results_dict.items():
            if df is not None and len(df) > 0:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"   ✅ Sheet '{sheet_name}': {len(df)} rows")
    print(f"\n🎉 Analytics complete! Open {output_file} to view results.")
    print("   Connect this file to Power BI: Home → Get Data → Excel Workbook")


# ════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 55)
    print("  CrazyKart Analytics Pipeline")
    print("=" * 55)

    df = load_data(EXCEL_FILE)

    results = {}

    if RUN_FORECAST:
        results['07_Sales_Forecast']   = run_sales_forecast(df, FORECAST_DAYS)

    if RUN_ANOMALY:
        results['08_Anomalies']        = run_anomaly_detection(df)

    if RUN_SCORING:
        results['09_Brand_Scoring']    = run_category_scoring(df)

    results['15_Store_Ranking']        = run_store_ranking(df)
    results['16_Inventory_Forecast']   = run_inventory_forecast(df)

    save_results(results, OUTPUT_FILE)
