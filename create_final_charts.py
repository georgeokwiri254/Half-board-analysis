import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.gridspec import GridSpec
from matplotlib.patches import Rectangle
import warnings
warnings.filterwarnings('ignore')

print("="*80)
print("CREATING FINAL EDA CHARTS (Categories 5-9)")
print("Charts 26-48")
print("="*80)

# Read the data
df = pd.read_excel('/home/gee_devops254/Downloads/Half Board/Half Board.xlsx')
df['Avg_Rate_Per_Night'] = df['Room Revenue'] / df['Room Nights']
df['Has_HB'] = df['Product (Descriptions)'].str.contains('Halfboard|Half Board', case=False, na=False)

df_hb = df[df['Has_HB']]
df_non_hb = df[~df['Has_HB']]

# Set style
plt.style.use('seaborn-v0_8-darkgrid')

# Color scheme
COLOR_HB = '#2ecc71'
COLOR_NON_HB = '#e74c3c'
COLOR_NEUTRAL = '#3498db'
COLOR_HIGHLIGHT = '#f39c12'
COLOR_CIS = '#9b59b6'
COLOR_LUX = '#1abc9c'

chart_count = 25  # Continue from where we left off

# Identify markets from rate codes
def identify_market(rate_code):
    if pd.isna(rate_code):
        return 'Unknown'
    rate_code = str(rate_code).upper()
    if 'CIS' in rate_code:
        return 'CIS Markets'
    elif 'MILUX' in rate_code:
        return 'Luxembourg'
    elif 'BBWI' in rate_code or 'BBJN' in rate_code or 'BB-WI' in rate_code or rate_code == 'TOBB':
        return 'Universal/Multi-Market'
    elif 'ROWI' in rate_code:
        return 'Specific Market'
    elif 'SSE' in rate_code or 'FSSE' in rate_code:
        return 'Secret Escapes'
    else:
        return 'Other'

df['Market_Segment'] = df['Rate Code'].apply(identify_market)

# ============================================================================
# CATEGORY 5: MARKET SEGMENTATION (5 charts)
# ============================================================================
print("\n[CATEGORY 5/9] Creating Market Segmentation Charts...")

# Prepare market data
market_data = df.groupby('Market_Segment').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum',
    'Has_HB': ['sum', lambda x: (x.sum() / len(x) * 100)]
}).reset_index()
market_data.columns = ['Market', 'Total_Nights', 'Total_Revenue', 'HB_Bookings', 'HB_Penetration']
market_data = market_data.sort_values('Total_Revenue', ascending=False)

# Chart 26: Market Segment Revenue Distribution
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

bars = ax.barh(range(len(market_data)), market_data['Total_Revenue'],
               color=plt.cm.Set3(np.linspace(0, 1, len(market_data))), edgecolor='black', linewidth=1)
ax.set_yticks(range(len(market_data)))
ax.set_yticklabels(market_data['Market'], fontsize=11)
ax.set_xlabel('Total Revenue (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 26: Market Segment Revenue Distribution\nRevenue by Market Type', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' AED {width:,.0f}',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/05_Market_Segmentation/26_market_revenue_distribution.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Market Revenue Distribution")

# Chart 27: Market Segment HB Penetration
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

colors = [COLOR_HB if x > 25 else (COLOR_HIGHLIGHT if x > 10 else COLOR_NON_HB) for x in market_data['HB_Penetration']]

bars = ax.barh(range(len(market_data)), market_data['HB_Penetration'], color=colors, edgecolor='black', linewidth=1)
ax.set_yticks(range(len(market_data)))
ax.set_yticklabels(market_data['Market'], fontsize=11)
ax.set_xlabel('HB Penetration (%)', fontsize=14, weight='bold')
ax.set_title('Chart 27: HB Penetration by Market Segment\nGreen >25% | Orange 10-25% | Red <10%', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)
ax.axvline(x=35, color='darkgreen', linestyle='--', linewidth=2, label='Target: 35%')
ax.legend(fontsize=12)

for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' {width:.1f}%',
            ha='left', va='center', fontsize=10, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/05_Market_Segmentation/27_market_hb_penetration.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Market HB Penetration")

# Chart 28: CIS Market Performance
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))

df_cis = df[df['Market_Segment'] == 'CIS Markets']
cis_by_rate = df_cis.groupby('Rate Code').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum',
    'Has_HB': 'sum'
}).sort_values('Room Revenue', ascending=False)

bars = ax.barh(range(len(cis_by_rate)), cis_by_rate['Room Revenue'], color=COLOR_CIS, edgecolor='black', linewidth=1, alpha=0.8)
ax.set_yticks(range(len(cis_by_rate)))
ax.set_yticklabels(cis_by_rate.index, fontsize=11)
ax.set_xlabel('Revenue (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 28: CIS Market Performance by Rate Code\nHighest HB Penetration Market (36%)', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

for i, bar in enumerate(bars):
    width = bar.get_width()
    hb_count = cis_by_rate.iloc[i]['Has_HB']
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' AED {width:,.0f} ({int(hb_count)} HB)',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/05_Market_Segmentation/28_cis_market_performance.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: CIS Market Performance")

# Chart 29: Luxembourg Market Analysis
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))

df_lux = df[df['Market_Segment'] == 'Luxembourg']
lux_by_agency = df_lux.groupby('Search Name').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum',
    'Has_HB': 'sum'
}).sort_values('Room Revenue', ascending=False)

bars = ax.barh(range(len(lux_by_agency)), lux_by_agency['Room Revenue'], color=COLOR_LUX, edgecolor='black', linewidth=1, alpha=0.8)
ax.set_yticks(range(len(lux_by_agency)))
ax.set_yticklabels(lux_by_agency.index, fontsize=11)
ax.set_xlabel('Revenue (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 29: Luxembourg Market - Agency Performance\nMiracle Tourism Dominates', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

for i, bar in enumerate(bars):
    width = bar.get_width()
    hb_count = lux_by_agency.iloc[i]['Has_HB']
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' AED {width:,.0f} ({int(hb_count)} HB)',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/05_Market_Segmentation/29_luxembourg_market_analysis.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Luxembourg Market Analysis")

# Chart 30: Universal vs Specific Market Comparison
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 8))

comparison_markets = ['Universal/Multi-Market', 'CIS Markets', 'Luxembourg', 'Specific Market', 'Secret Escapes']
comp_data = market_data[market_data['Market'].isin(comparison_markets)]

x = np.arange(len(comp_data))
width = 0.35

bars1 = ax.bar(x - width/2, comp_data['Total_Nights'], width, label='Total Room Nights',
               color=COLOR_NEUTRAL, edgecolor='black')
bars2 = ax.bar(x + width/2, comp_data['HB_Penetration'] * 10, width, label='HB Penetration % (×10)',
               color=COLOR_HB, edgecolor='black')

ax.set_ylabel('Count / Percentage', fontsize=14, weight='bold')
ax.set_title('Chart 30: Universal vs Specific Markets - Comparison\nVolume and HB Penetration', fontsize=16, weight='bold', pad=20)
ax.set_xticks(x)
ax.set_xticklabels(comp_data['Market'], rotation=45, ha='right', fontsize=10)
ax.legend(fontsize=11)
ax.grid(axis='y', alpha=0.3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/05_Market_Segmentation/30_universal_vs_specific_comparison.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Universal vs Specific Market Comparison")

# ============================================================================
# CATEGORY 6: HB PERFORMANCE ANALYSIS (6 charts)
# ============================================================================
print("\n[CATEGORY 6/9] Creating HB Performance Analysis Charts...")

# Chart 31: HB Bookings by Agency (Top 15 - Tree-like visualization)
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

agency_hb = df_hb.groupby('Search Name')['Room Nights'].sum().sort_values(ascending=False).head(15)

# Create a pseudo-treemap using nested bars
y_pos = 0
colors_tree = plt.cm.Greens(np.linspace(0.4, 0.9, len(agency_hb)))

for i, (agency, nights) in enumerate(agency_hb.items()):
    height = nights / agency_hb.sum() * 10  # Scale for visualization
    ax.barh(y_pos, nights, height=height, color=colors_tree[i], edgecolor='black', linewidth=2)
    ax.text(nights/2, y_pos, f'{agency}\n{nights:.0f} nights',
            ha='center', va='center', fontsize=9, weight='bold')
    y_pos += height + 0.1

ax.set_xlabel('Half Board Room Nights', fontsize=14, weight='bold')
ax.set_title('Chart 31: HB Bookings by Agency - Top 15\nProportional Representation', fontsize=16, weight='bold', pad=20)
ax.set_yticks([])
ax.grid(axis='x', alpha=0.3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/06_HB_Performance/31_hb_bookings_by_agency_tree.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: HB Bookings Tree Visualization")

# Chart 32: HB Revenue Concentration (Pareto)
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

agency_hb_rev = df_hb.groupby('Search Name')['Room Revenue'].sum().sort_values(ascending=False)
cumulative_pct = (agency_hb_rev.cumsum() / agency_hb_rev.sum() * 100)

# Plot top 20
top20_hb_rev = agency_hb_rev.head(20)
top20_cum = cumulative_pct.head(20)

ax2 = ax.twinx()

bars = ax.bar(range(len(top20_hb_rev)), top20_hb_rev.values, color=COLOR_HB, edgecolor='black', alpha=0.7)
line = ax2.plot(range(len(top20_cum)), top20_cum.values, color='red', marker='o', linewidth=3, markersize=8, label='Cumulative %')
ax2.axhline(y=80, color='orange', linestyle='--', linewidth=2, label='80% threshold')

ax.set_xlabel('Agency (Top 20)', fontsize=12, weight='bold')
ax.set_ylabel('HB Revenue (AED)', fontsize=12, weight='bold', color='green')
ax2.set_ylabel('Cumulative %', fontsize=12, weight='bold', color='red')
ax.set_title('Chart 32: HB Revenue Concentration - Pareto Analysis\n80/20 Rule: Few agencies drive most revenue', fontsize=16, weight='bold', pad=20)
ax.set_xticks(range(len(top20_hb_rev)))
ax.set_xticklabels([name[:15] for name in top20_hb_rev.index], rotation=45, ha='right', fontsize=8)
ax2.legend(loc='center right', fontsize=11)
ax.grid(axis='y', alpha=0.3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/06_HB_Performance/32_hb_revenue_pareto.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: HB Revenue Pareto Analysis")

# Chart 33: HB Average Rate Distribution
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))

hb_rates = df_hb['Avg_Rate_Per_Night'].dropna()

ax.hist(hb_rates, bins=30, color=COLOR_HB, edgecolor='black', alpha=0.7)
ax.axvline(hb_rates.mean(), color='red', linestyle='--', linewidth=2, label=f'Mean: AED {hb_rates.mean():.2f}')
ax.axvline(hb_rates.median(), color='orange', linestyle='--', linewidth=2, label=f'Median: AED {hb_rates.median():.2f}')

ax.set_xlabel('Average Rate per Night (AED)', fontsize=14, weight='bold')
ax.set_ylabel('Frequency', fontsize=14, weight='bold')
ax.set_title('Chart 33: HB Average Rate Distribution\nHistogram of HB Rates', fontsize=16, weight='bold', pad=20)
ax.legend(fontsize=12)
ax.grid(axis='y', alpha=0.3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/06_HB_Performance/33_hb_avgrate_distribution.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: HB Average Rate Distribution")

# Chart 34: HB vs Non-HB Average Rate by Agency (Scatter)
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

# Calculate avg rates by agency
agency_rate_comparison = []
for agency in df['Search Name'].dropna().unique():
    agency_df = df[df['Search Name'] == agency]
    agency_hb_df = agency_df[agency_df['Has_HB']]
    agency_nonhb_df = agency_df[~agency_df['Has_HB']]

    if len(agency_hb_df) > 0 and len(agency_nonhb_df) > 0:
        agency_rate_comparison.append({
            'Agency': agency,
            'HB_Rate': agency_hb_df['Avg_Rate_Per_Night'].mean(),
            'NonHB_Rate': agency_nonhb_df['Avg_Rate_Per_Night'].mean(),
            'Total_Revenue': agency_df['Room Revenue'].sum()
        })

df_rate_comp = pd.DataFrame(agency_rate_comparison)

# Size by revenue
sizes = (df_rate_comp['Total_Revenue'] / df_rate_comp['Total_Revenue'].max() * 500) + 50

scatter = ax.scatter(df_rate_comp['NonHB_Rate'], df_rate_comp['HB_Rate'],
                     s=sizes, alpha=0.6, c=range(len(df_rate_comp)), cmap='viridis', edgecolors='black', linewidth=1)

# Add diagonal line
max_rate = max(df_rate_comp['HB_Rate'].max(), df_rate_comp['NonHB_Rate'].max())
ax.plot([0, max_rate], [0, max_rate], 'r--', linewidth=2, label='Equal Rates')

ax.set_xlabel('Non-HB Average Rate (AED)', fontsize=14, weight='bold')
ax.set_ylabel('HB Average Rate (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 34: HB vs Non-HB Average Rates by Agency\nBubble size = Total Revenue', fontsize=16, weight='bold', pad=20)
ax.legend(fontsize=12)
ax.grid(alpha=0.3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/06_HB_Performance/34_hb_nonhb_rate_scatter.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: HB vs Non-HB Rate Scatter")

# Chart 35: HB Penetration Heatmap (Top 10 Agencies vs Top 10 Rate Codes)
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

# Get top 10 agencies and rate codes
top10_agencies = df.groupby('Search Name')['Room Revenue'].sum().sort_values(ascending=False).head(10).index
top10_rates = df.groupby('Rate Code')['Room Revenue'].sum().sort_values(ascending=False).head(10).index

# Create matrix
heatmap_data = np.zeros((len(top10_agencies), len(top10_rates)))

for i, agency in enumerate(top10_agencies):
    for j, rate in enumerate(top10_rates):
        agency_rate_df = df[(df['Search Name'] == agency) & (df['Rate Code'] == rate)]
        if len(agency_rate_df) > 0:
            hb_pct = (agency_rate_df['Has_HB'].sum() / len(agency_rate_df) * 100)
            heatmap_data[i, j] = hb_pct

im = ax.imshow(heatmap_data, cmap='RdYlGn', aspect='auto', vmin=0, vmax=100)

ax.set_xticks(range(len(top10_rates)))
ax.set_yticks(range(len(top10_agencies)))
ax.set_xticklabels(top10_rates, rotation=45, ha='right', fontsize=10)
ax.set_yticklabels([name[:20] for name in top10_agencies], fontsize=10)

ax.set_title('Chart 35: HB Penetration Heatmap\nTop 10 Agencies × Top 10 Rate Codes', fontsize=16, weight='bold', pad=20)

# Add colorbar
cbar = plt.colorbar(im, ax=ax)
cbar.set_label('HB Penetration %', fontsize=12, weight='bold')

# Add text annotations
for i in range(len(top10_agencies)):
    for j in range(len(top10_rates)):
        if heatmap_data[i, j] > 0:
            text = ax.text(j, i, f'{heatmap_data[i, j]:.0f}%',
                          ha="center", va="center", color="black", fontsize=8, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/06_HB_Performance/35_hb_penetration_heatmap.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: HB Penetration Heatmap")

# Chart 36: Booking Size vs HB Adoption
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

# Create bins
df['Booking_Size_Bin'] = pd.cut(df['Room Nights'],
                                  bins=[0, 5, 10, 20, 50, 100, 1000],
                                  labels=['1-5', '6-10', '11-20', '21-50', '51-100', '100+'])

booking_hb_analysis = df.groupby('Booking_Size_Bin', observed=True).agg({
    'Has_HB': ['sum', 'count']
}).reset_index()
booking_hb_analysis.columns = ['Booking_Size', 'HB_Count', 'Total_Count']
booking_hb_analysis['HB_Pct'] = (booking_hb_analysis['HB_Count'] / booking_hb_analysis['Total_Count'] * 100)

scatter = ax.scatter(range(len(booking_hb_analysis)), booking_hb_analysis['HB_Pct'],
                     s=booking_hb_analysis['Total_Count']*3, alpha=0.6, c=COLOR_HB, edgecolors='black', linewidth=2)

# Add trend line
z = np.polyfit(range(len(booking_hb_analysis)), booking_hb_analysis['HB_Pct'], 1)
p = np.poly1d(z)
ax.plot(range(len(booking_hb_analysis)), p(range(len(booking_hb_analysis))),
        "r--", linewidth=2, label='Trend')

ax.set_xlabel('Booking Size Category', fontsize=14, weight='bold')
ax.set_ylabel('HB Adoption %', fontsize=14, weight='bold')
ax.set_title('Chart 36: Booking Size vs HB Adoption\nBubble size = Number of bookings', fontsize=16, weight='bold', pad=20)
ax.set_xticks(range(len(booking_hb_analysis)))
ax.set_xticklabels(booking_hb_analysis['Booking_Size'], fontsize=12)
ax.legend(fontsize=12)
ax.grid(alpha=0.3)

# Add value labels
for i, row in booking_hb_analysis.iterrows():
    ax.text(i, row['HB_Pct'] + 2, f'{row["HB_Pct"]:.1f}%',
            ha='center', fontsize=10, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/06_HB_Performance/36_booking_size_vs_hb_adoption.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Booking Size vs HB Adoption")

print(f"\nCompleted Category 6: {chart_count} charts created")
print("="*80)
print(f"PROGRESS: {chart_count}/48 charts completed - 68% done!")
print("="*80)
