import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.gridspec import GridSpec
import warnings
warnings.filterwarnings('ignore')

print("="*80)
print("CREATING REMAINING EDA CHARTS (Categories 3-9)")
print("="*80)

# Read the data
df = pd.read_excel('/home/gee_devops254/Downloads/Half Board/Half Board.xlsx')
df['Avg_Rate_Per_Night'] = df['Room Revenue'] / df['Room Nights']
df['Has_HB'] = df['Product (Descriptions)'].str.contains('Halfboard|Half Board', case=False, na=False)

df_hb = df[df['Has_HB']]
df_non_hb = df[~df['Has_HB']]

# Set style
plt.style.use('seaborn-v0_8-darkgrid')
sns.set_palette("husl")

# Color scheme
COLOR_HB = '#2ecc71'
COLOR_NON_HB = '#e74c3c'
COLOR_NEUTRAL = '#3498db'
COLOR_HIGHLIGHT = '#f39c12'
COLOR_MIRACLE = '#9b59b6'

chart_count = 14  # Continue from where we left off

# ============================================================================
# CATEGORY 3: MIRACLE TOURISM DEEP DIVE (5 charts)
# ============================================================================
print("\n[CATEGORY 3/9] Creating Miracle Tourism Deep Dive Charts...")

miracle_data = df[df['Search Name'] == 'MIRACLE TOURISM LLC']
miracle_hb = miracle_data[miracle_data['Has_HB']]

# Chart 15: Miracle vs Top 5 Agencies Comparison
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))

# Get metrics for comparison
agencies = ['MIRACLE TOURISM LLC', 'TBO HOLIDAYS', 'DARINA HOLIDAYS', 'WEBBEDS', 'VOYAGE TOURS', 'DUBAI LINK TOURS L.L.C.']
comparison_data = []

for agency in agencies:
    agency_df = df[df['Search Name'] == agency]
    agency_hb_df = agency_df[agency_df['Has_HB']]

    if len(agency_df) > 0:
        comparison_data.append({
            'Agency': agency.split()[0],
            'Total Room Nights': agency_df['Room Nights'].sum(),
            'HB Penetration %': (len(agency_hb_df) / len(agency_df) * 100) if len(agency_df) > 0 else 0,
            'Avg Rate': agency_df['Avg_Rate_Per_Night'].mean(),
            'Total Revenue (k)': agency_df['Room Revenue'].sum() / 1000
        })

df_comp = pd.DataFrame(comparison_data)

x = np.arange(len(df_comp))
width = 0.2

fig, ax = plt.subplots(figsize=(14, 8))

# Normalize metrics for comparison
metrics = ['Total Room Nights', 'HB Penetration %', 'Avg Rate', 'Total Revenue (k)']
colors_bars = [COLOR_MIRACLE, COLOR_HB, COLOR_NEUTRAL, COLOR_HIGHLIGHT]

for i, metric in enumerate(metrics):
    normalized = df_comp[metric] / df_comp[metric].max() * 100
    ax.bar(x + i*width, normalized, width, label=metric, color=colors_bars[i], alpha=0.8)

ax.set_ylabel('Normalized Score (0-100)', fontsize=12, weight='bold')
ax.set_title('Chart 15: Miracle vs Top 5 Agencies - Multi-Metric Comparison\nNormalized Performance Across Key Metrics',
             fontsize=14, weight='bold', pad=15)
ax.set_xticks(x + width * 1.5)
ax.set_xticklabels(df_comp['Agency'], rotation=45, ha='right')
ax.legend(loc='upper right', fontsize=10)
ax.grid(axis='y', alpha=0.3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/03_Miracle_Deep_Dive/15_miracle_vs_top5_comparison.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Miracle vs Top 5 Comparison")

# Chart 16: Miracle Booking Size Distribution
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))

miracle_data['Booking_Category'] = pd.cut(miracle_data['Room Nights'],
                                           bins=[0, 10, 30, 50, 100, 1000],
                                           labels=['1-10', '11-30', '31-50', '51-100', '100+'])

booking_dist = miracle_data['Booking_Category'].value_counts().sort_index()

bars = ax.bar(range(len(booking_dist)), booking_dist.values, color=COLOR_MIRACLE, edgecolor='black', linewidth=2, alpha=0.8)
ax.set_xticks(range(len(booking_dist)))
ax.set_xticklabels(booking_dist.index, fontsize=12)
ax.set_xlabel('Booking Size (Room Nights)', fontsize=14, weight='bold')
ax.set_ylabel('Number of Bookings', fontsize=14, weight='bold')
ax.set_title('Chart 16: Miracle Tourism - Booking Size Distribution\nMost bookings are LARGE (30+ nights)', fontsize=16, weight='bold', pad=20)
ax.grid(axis='y', alpha=0.3)

# Add value labels
for bar in bars:
    height = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., height,
            f'{int(height)}',
            ha='center', va='bottom', fontsize=12, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/03_Miracle_Deep_Dive/16_miracle_booking_size_distribution.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Miracle Booking Size Distribution")

# Chart 17: Miracle HB vs Non-HB Split
chart_count += 1
fig, ax = plt.subplots(figsize=(10, 8))

sizes = [len(miracle_hb), len(miracle_data) - len(miracle_hb)]
labels = [f'Half Board\n{len(miracle_hb)} bookings',
          f'Non-Half Board\n{len(miracle_data) - len(miracle_hb)} bookings']
colors = [COLOR_HB, COLOR_NON_HB]
explode = (0.1, 0)

wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                                    explode=explode, shadow=True, startangle=90, textprops={'fontsize': 14, 'weight': 'bold'})
ax.set_title(f'Chart 17: Miracle Tourism - HB Penetration\n{len(miracle_hb)/len(miracle_data)*100:.1f}% of their bookings include HB',
             fontsize=16, weight='bold', pad=20)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/03_Miracle_Deep_Dive/17_miracle_hb_split.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Miracle HB Split")

# Chart 18: Miracle Rate Code Performance
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))

miracle_by_rate = miracle_data.groupby('Rate Code').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum'
}).sort_values('Room Revenue', ascending=False)

bars = ax.barh(range(len(miracle_by_rate)), miracle_by_rate['Room Revenue'], color=COLOR_MIRACLE, edgecolor='black', linewidth=1, alpha=0.8)
ax.set_yticks(range(len(miracle_by_rate)))
ax.set_yticklabels(miracle_by_rate.index, fontsize=11)
ax.set_xlabel('Revenue (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 18: Miracle Tourism - Rate Code Performance\nRevenue by Rate Code', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' AED {width:,.0f}',
            ha='left', va='center', fontsize=10, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/03_Miracle_Deep_Dive/18_miracle_rate_code_performance.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Miracle Rate Code Performance")

# Chart 19: Miracle Revenue Contribution
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))

miracle_contribution = [
    miracle_data['Room Revenue'].sum(),
    df['Room Revenue'].sum() - miracle_data['Room Revenue'].sum()
]
labels = [f'Miracle Tourism\nAED {miracle_data["Room Revenue"].sum():,.0f}\n({miracle_data["Room Revenue"].sum()/df["Room Revenue"].sum()*100:.2f}%)',
          f'All Other Agencies\nAED {df["Room Revenue"].sum() - miracle_data["Room Revenue"].sum():,.0f}']
colors = [COLOR_MIRACLE, '#bdc3c7']
explode = (0.1, 0)

wedges, texts, autotexts = ax.pie(miracle_contribution, labels=labels, colors=colors, autopct='%1.1f%%',
                                    explode=explode, shadow=True, startangle=90, textprops={'fontsize':13, 'weight': 'bold'})
ax.set_title('Chart 19: Miracle Tourism - Total Revenue Contribution\nShare of Total Property Revenue',
             fontsize=16, weight='bold', pad=20)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/03_Miracle_Deep_Dive/19_miracle_revenue_contribution.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Miracle Revenue Contribution")

# ============================================================================
# CATEGORY 4: RATE CODE ANALYSIS (6 charts)
# ============================================================================
print("\n[CATEGORY 4/9] Creating Rate Code Analysis Charts...")

# Prepare rate code data
rate_data = df.groupby('Rate Code').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum',
    'Has_HB': ['sum', lambda x: (x.sum() / len(x) * 100)]
}).reset_index()
rate_data.columns = ['Rate_Code', 'Total_Nights', 'Total_Revenue', 'HB_Bookings', 'HB_Penetration']

# Chart 20: Top 15 Rate Codes by Revenue
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))
top15_rates = rate_data.sort_values('Total_Revenue', ascending=False).head(15)
colors_gradient = plt.cm.coolwarm(np.linspace(0.2, 0.8, len(top15_rates)))

bars = ax.barh(range(len(top15_rates)), top15_rates['Total_Revenue'], color=colors_gradient, edgecolor='black', linewidth=1)
ax.set_yticks(range(len(top15_rates)))
ax.set_yticklabels(top15_rates['Rate_Code'], fontsize=11)
ax.set_xlabel('Total Revenue (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 20: Top 15 Rate Codes by Revenue\nHighest Revenue Generators', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' AED {width:,.0f}',
            ha='left', va='center', fontsize=9, weight='bold')

# Highlight TOBBWI and TOBBJN
for i, code in enumerate(top15_rates['Rate_Code']):
    if code in ['TOBBWI', 'TOBBJN']:
        ax.get_children()[i].set_edgecolor('red')
        ax.get_children()[i].set_linewidth(3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/04_Rate_Code_Analysis/20_top15_ratecodes_revenue.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Top 15 Rate Codes by Revenue")

# Chart 21: Top 15 Rate Codes by Room Nights
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))
top15_rates_nights = rate_data.sort_values('Total_Nights', ascending=False).head(15)
colors_gradient = plt.cm.viridis(np.linspace(0.2, 0.8, len(top15_rates_nights)))

bars = ax.barh(range(len(top15_rates_nights)), top15_rates_nights['Total_Nights'], color=colors_gradient, edgecolor='black', linewidth=1)
ax.set_yticks(range(len(top15_rates_nights)))
ax.set_yticklabels(top15_rates_nights['Rate_Code'], fontsize=11)
ax.set_xlabel('Total Room Nights', fontsize=14, weight='bold')
ax.set_title('Chart 21: Top 15 Rate Codes by Room Nights\nHighest Volume Rate Codes', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' {int(width):,} nights',
            ha='left', va='center', fontsize=9, weight='bold')

# Highlight TOBBWI and TOBBJN
for i, code in enumerate(top15_rates_nights['Rate_Code']):
    if code in ['TOBBWI', 'TOBBJN']:
        ax.get_children()[i].set_edgecolor('red')
        ax.get_children()[i].set_linewidth(3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/04_Rate_Code_Analysis/21_top15_ratecodes_nights.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Top 15 Rate Codes by Nights")

# Chart 22: TOBBWI - Agency Performance
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

tobbwi_agencies = df[df['Rate Code'] == 'TOBBWI'].groupby('Search Name').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum',
    'Has_HB': 'sum'
}).sort_values('Room Revenue', ascending=False).head(10).reset_index()

bars = ax.barh(range(len(tobbwi_agencies)), tobbwi_agencies['Room Revenue'], color='#3498db', edgecolor='black', linewidth=1, alpha=0.8)
ax.set_yticks(range(len(tobbwi_agencies)))
ax.set_yticklabels(tobbwi_agencies['Search Name'], fontsize=10)
ax.set_xlabel('Revenue (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 22: TOBBWI Rate Code - Top 10 Agency Performance\nUniversal Code - All Markets', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    hb_count = tobbwi_agencies.iloc[i]['Has_HB']
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' AED {width:,.0f} ({int(hb_count)} HB)',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/04_Rate_Code_Analysis/22_tobbwi_agency_performance.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: TOBBWI Agency Performance")

# Chart 23: TOBBJN - Agency Performance
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

tobbjn_agencies = df[df['Rate Code'] == 'TOBBJN'].groupby('Search Name').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum',
    'Has_HB': 'sum'
}).sort_values('Room Revenue', ascending=False).head(10).reset_index()

bars = ax.barh(range(len(tobbjn_agencies)), tobbjn_agencies['Room Revenue'], color='#e74c3c', edgecolor='black', linewidth=1, alpha=0.8)
ax.set_yticks(range(len(tobbjn_agencies)))
ax.set_yticklabels(tobbjn_agencies['Search Name'], fontsize=10)
ax.set_xlabel('Revenue (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 23: TOBBJN Rate Code - Top 10 Agency Performance\nUniversal Code - All Markets', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    hb_count = tobbjn_agencies.iloc[i]['Has_HB']
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' AED {width:,.0f} ({int(hb_count)} HB)',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/04_Rate_Code_Analysis/23_tobbjn_agency_performance.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: TOBBJN Agency Performance")

# Chart 24: Universal Codes HB Penetration
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))

universal_codes = ['TOBBWI', 'TOBBJN']
universal_hb_data = []

for code in universal_codes:
    code_df = df[df['Rate Code'] == code]
    code_hb = code_df[code_df['Has_HB']]
    total_bookings = len(code_df)
    hb_bookings = len(code_hb)
    non_hb_bookings = total_bookings - hb_bookings

    universal_hb_data.append({
        'Rate Code': code,
        'HB Bookings': hb_bookings,
        'Non-HB Bookings': non_hb_bookings,
        'HB %': (hb_bookings / total_bookings * 100) if total_bookings > 0 else 0
    })

df_universal = pd.DataFrame(universal_hb_data)

x = np.arange(len(df_universal))
width = 0.35

bars1 = ax.bar(x - width/2, df_universal['HB Bookings'], width, label='HB Bookings', color=COLOR_HB, edgecolor='black')
bars2 = ax.bar(x + width/2, df_universal['Non-HB Bookings'], width, label='Non-HB Bookings', color=COLOR_NON_HB, edgecolor='black')

ax.set_ylabel('Number of Bookings', fontsize=14, weight='bold')
ax.set_title('Chart 24: Universal Rate Codes - HB vs Non-HB Comparison\nMassive opportunity for HB bundling', fontsize=16, weight='bold', pad=20)
ax.set_xticks(x)
ax.set_xticklabels(df_universal['Rate Code'], fontsize=12)
ax.legend(fontsize=12)
ax.grid(axis='y', alpha=0.3)

# Add percentage labels
for i, (bar1, bar2) in enumerate(zip(bars1, bars2)):
    total = bar1.get_height() + bar2.get_height()
    ax.text(bar1.get_x() + bar1.get_width()/2., total + 2,
            f'{df_universal.iloc[i]["HB %"]:.1f}% HB',
            ha='center', va='bottom', fontsize=12, weight='bold', color='red')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/04_Rate_Code_Analysis/24_universal_codes_hb_penetration.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Universal Codes HB Penetration")

# Chart 25: Rate Code Average Rate Comparison
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

top15_rates_avgrate = rate_data.sort_values('Total_Revenue', ascending=False).head(15).copy()

# Calculate avg rate
rate_avg_rates = []
for code in top15_rates_avgrate['Rate_Code']:
    code_df = df[df['Rate Code'] == code]
    avg_rate = code_df['Avg_Rate_Per_Night'].mean()
    rate_avg_rates.append(avg_rate)

top15_rates_avgrate['Avg_Rate'] = rate_avg_rates

# Color bars by rate level
colors = [COLOR_HB if x > 400 else (COLOR_HIGHLIGHT if x > 300 else COLOR_NON_HB) for x in top15_rates_avgrate['Avg_Rate']]

bars = ax.barh(range(len(top15_rates_avgrate)), top15_rates_avgrate['Avg_Rate'], color=colors, edgecolor='black', linewidth=1)
ax.set_yticks(range(len(top15_rates_avgrate)))
ax.set_yticklabels(top15_rates_avgrate['Rate_Code'], fontsize=11)
ax.set_xlabel('Average Rate (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 25: Top 15 Rate Codes - Average Rate Comparison\nGreen >400 | Orange 300-400 | Red <300',
             fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' AED {width:.2f}',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/04_Rate_Code_Analysis/25_ratecode_avg_rate_comparison.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Rate Code Average Rate Comparison")

print(f"\nCompleted Categories 3-4: {chart_count} charts created so far")
print("Continuing to categories 5-9...")
