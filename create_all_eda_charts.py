import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.gridspec import GridSpec
import warnings
warnings.filterwarnings('ignore')

print("="*80)
print("CREATING COMPREHENSIVE EDA CHART ANALYSIS")
print("48 High-Quality Charts for Half Board Analysis")
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
COLOR_HB = '#2ecc71'  # Green for HB
COLOR_NON_HB = '#e74c3c'  # Red for Non-HB
COLOR_NEUTRAL = '#3498db'  # Blue
COLOR_HIGHLIGHT = '#f39c12'  # Orange
COLOR_GRADIENT = ['#e74c3c', '#f39c12', '#f1c40f', '#2ecc71', '#27ae60']

chart_count = 0

# ============================================================================
# CATEGORY 1: OVERVIEW & DISTRIBUTION ANALYSIS (6 charts)
# ============================================================================
print("\n[CATEGORY 1/9] Creating Overview & Distribution Charts...")

# Chart 1: Overall HB Penetration
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))
sizes = [len(df_hb), len(df_non_hb)]
labels = [f'Half Board\n{len(df_hb)} bookings\n({len(df_hb)/len(df)*100:.1f}%)',
          f'Non-Half Board\n{len(df_non_hb)} bookings\n({len(df_non_hb)/len(df)*100:.1f}%)']
colors = [COLOR_HB, COLOR_NON_HB]
explode = (0.1, 0)

wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                                    explode=explode, shadow=True, startangle=90, textprops={'fontsize': 14, 'weight': 'bold'})
ax.set_title('Chart 1: Overall Half Board Penetration Rate\nCRITICAL: Only 14% of bookings include F&B!',
             fontsize=16, weight='bold', pad=20)
plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/01_Overview_Distribution/01_overall_hb_penetration.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Overall HB Penetration")

# Chart 2: Revenue Distribution
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))
sizes = [df_hb['Room Revenue'].sum(), df_non_hb['Room Revenue'].sum()]
labels = [f'Half Board Revenue\nAED {df_hb["Room Revenue"].sum():,.0f}\n({df_hb["Room Revenue"].sum()/df["Room Revenue"].sum()*100:.1f}%)',
          f'Non-HB Revenue\nAED {df_non_hb["Room Revenue"].sum():,.0f}\n({df_non_hb["Room Revenue"].sum()/df["Room Revenue"].sum()*100:.1f}%)']
colors = [COLOR_HB, COLOR_NON_HB]
explode = (0.1, 0)

wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                                    explode=explode, shadow=True, startangle=90, textprops={'fontsize': 14, 'weight': 'bold'})
ax.set_title('Chart 2: Revenue Distribution - HB vs Non-HB\nHB generates only 3.1% of total revenue!',
             fontsize=16, weight='bold', pad=20)
plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/01_Overview_Distribution/02_revenue_distribution.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Revenue Distribution")

# Chart 3: Room Nights Distribution
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))
sizes = [df_hb['Room Nights'].sum(), df_non_hb['Room Nights'].sum()]
labels = [f'Half Board Nights\n{df_hb["Room Nights"].sum():,} nights\n({df_hb["Room Nights"].sum()/df["Room Nights"].sum()*100:.1f}%)',
          f'Non-HB Nights\n{df_non_hb["Room Nights"].sum():,} nights\n({df_non_hb["Room Nights"].sum()/df["Room Nights"].sum()*100:.1f}%)']
colors = [COLOR_HB, COLOR_NON_HB]
explode = (0.1, 0)

wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                                    explode=explode, shadow=True, startangle=90, textprops={'fontsize': 14, 'weight': 'bold'})
ax.set_title('Chart 3: Room Nights Distribution - HB vs Non-HB\nOnly 2.7% of room nights include F&B',
             fontsize=16, weight='bold', pad=20)
plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/01_Overview_Distribution/03_room_nights_distribution.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Room Nights Distribution")

# Chart 4: Booking Count Distribution
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))
categories = ['Half Board', 'Non-Half Board', 'Total']
counts = [len(df_hb), len(df_non_hb), len(df)]
colors_bar = [COLOR_HB, COLOR_NON_HB, COLOR_NEUTRAL]

bars = ax.bar(categories, counts, color=colors_bar, edgecolor='black', linewidth=2)
ax.set_ylabel('Number of Bookings', fontsize=14, weight='bold')
ax.set_title('Chart 4: Booking Count Distribution\nTotal Bookings by Category', fontsize=16, weight='bold', pad=20)
ax.grid(axis='y', alpha=0.3)

# Add value labels on bars
for bar in bars:
    height = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., height,
            f'{int(height):,}',
            ha='center', va='bottom', fontsize=14, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/01_Overview_Distribution/04_booking_count_distribution.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Booking Count Distribution")

# Chart 5: Average Rate Comparison
chart_count += 1
fig, ax = plt.subplots(figsize=(12, 8))
categories = ['Overall Avg', 'Half Board Avg', 'Non-HB Avg']
rates = [df['Avg_Rate_Per_Night'].mean(), df_hb['Avg_Rate_Per_Night'].mean(), df_non_hb['Avg_Rate_Per_Night'].mean()]
colors_bar = [COLOR_NEUTRAL, COLOR_HB, COLOR_NON_HB]

bars = ax.bar(categories, rates, color=colors_bar, edgecolor='black', linewidth=2)
ax.set_ylabel('Average Rate per Night (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 5: Average Rate Comparison\nHB Premium: AED 26 higher than Non-HB', fontsize=16, weight='bold', pad=20)
ax.grid(axis='y', alpha=0.3)

# Add value labels on bars
for bar in bars:
    height = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., height,
            f'AED {height:.2f}',
            ha='center', va='bottom', fontsize=14, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/01_Overview_Distribution/05_average_rate_comparison.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Average Rate Comparison")

# Chart 6: Revenue vs Room Nights Scatter
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))
scatter_hb = ax.scatter(df_hb['Room Nights'], df_hb['Room Revenue'],
                        c=COLOR_HB, s=100, alpha=0.6, edgecolors='black', linewidth=1, label='Half Board')
scatter_non = ax.scatter(df_non_hb['Room Nights'], df_non_hb['Room Revenue'],
                         c=COLOR_NON_HB, s=100, alpha=0.6, edgecolors='black', linewidth=1, label='Non-Half Board')

ax.set_xlabel('Room Nights', fontsize=14, weight='bold')
ax.set_ylabel('Room Revenue (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 6: Revenue vs Room Nights Scatter Plot\nAll Bookings by HB Status', fontsize=16, weight='bold', pad=20)
ax.legend(fontsize=12)
ax.grid(alpha=0.3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/01_Overview_Distribution/06_revenue_vs_room_nights_scatter.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Revenue vs Room Nights Scatter")

# ============================================================================
# CATEGORY 2: AGENCY ANALYSIS (8 charts)
# ============================================================================
print("\n[CATEGORY 2/9] Creating Agency Analysis Charts...")

# Prepare agency data
agency_data = df.groupby('Search Name').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum',
    'Has_HB': ['sum', lambda x: (x.sum() / len(x) * 100)]
}).reset_index()
agency_data.columns = ['Agency', 'Total_Nights', 'Total_Revenue', 'HB_Bookings', 'HB_Penetration']
agency_data = agency_data.sort_values('Total_Revenue', ascending=False)

agency_hb_data = df[df['Has_HB']].groupby('Search Name').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum'
}).reset_index()
agency_hb_data.columns = ['Agency', 'HB_Nights', 'HB_Revenue']

# Chart 7: Top 20 Agencies by Total Revenue
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))
top20 = agency_data.head(20).copy()
colors_gradient = plt.cm.viridis(np.linspace(0.3, 0.9, len(top20)))

bars = ax.barh(range(len(top20)), top20['Total_Revenue'], color=colors_gradient, edgecolor='black', linewidth=1)
ax.set_yticks(range(len(top20)))
ax.set_yticklabels(top20['Agency'], fontsize=10)
ax.set_xlabel('Total Revenue (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 7: Top 20 Agencies by Total Revenue\nRevenue Leaders in Dataset', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' AED {width:,.0f}',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/02_Agency_Analysis/07_top20_agencies_revenue.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Top 20 Agencies by Revenue")

# Chart 8: Top 20 Agencies by Room Nights
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))
top20_nights = agency_data.sort_values('Total_Nights', ascending=False).head(20).copy()
colors_gradient = plt.cm.plasma(np.linspace(0.3, 0.9, len(top20_nights)))

bars = ax.barh(range(len(top20_nights)), top20_nights['Total_Nights'], color=colors_gradient, edgecolor='black', linewidth=1)
ax.set_yticks(range(len(top20_nights)))
ax.set_yticklabels(top20_nights['Agency'], fontsize=10)
ax.set_xlabel('Total Room Nights', fontsize=14, weight='bold')
ax.set_title('Chart 8: Top 20 Agencies by Room Nights\nVolume Leaders in Dataset', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' {int(width):,} nights',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/02_Agency_Analysis/08_top20_agencies_room_nights.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Top 20 Agencies by Room Nights")

# Chart 9: Top 15 HB Agencies by Room Nights
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))
top15_hb = agency_hb_data.sort_values('HB_Nights', ascending=False).head(15).copy()

bars = ax.barh(range(len(top15_hb)), top15_hb['HB_Nights'], color=COLOR_HB, edgecolor='black', linewidth=1, alpha=0.8)
ax.set_yticks(range(len(top15_hb)))
ax.set_yticklabels(top15_hb['Agency'], fontsize=10)
ax.set_xlabel('Half Board Room Nights', fontsize=14, weight='bold')
ax.set_title('Chart 9: Top 15 Half Board Performers by Room Nights\nLeading HB Agencies', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' {int(width):,} nights',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/02_Agency_Analysis/09_top15_hb_agencies_nights.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Top 15 HB Agencies by Nights")

# Chart 10: Top 15 HB Agencies by Revenue
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))
top15_hb_rev = agency_hb_data.sort_values('HB_Revenue', ascending=False).head(15).copy()

bars = ax.barh(range(len(top15_hb_rev)), top15_hb_rev['HB_Revenue'], color=COLOR_HIGHLIGHT, edgecolor='black', linewidth=1, alpha=0.8)
ax.set_yticks(range(len(top15_hb_rev)))
ax.set_yticklabels(top15_hb_rev['Agency'], fontsize=10)
ax.set_xlabel('Half Board Revenue (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 10: Top 15 Half Board Performers by Revenue\nHighest HB Revenue Generators', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' AED {width:,.0f}',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/02_Agency_Analysis/10_top15_hb_agencies_revenue.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Top 15 HB Agencies by Revenue")

# Chart 11: HB Penetration Rate by Top 20 Agencies
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))
top20_pen = agency_data.head(20).copy()

# Color bars based on penetration rate
colors = [COLOR_HB if x > 25 else (COLOR_HIGHLIGHT if x > 10 else COLOR_NON_HB) for x in top20_pen['HB_Penetration']]

bars = ax.barh(range(len(top20_pen)), top20_pen['HB_Penetration'], color=colors, edgecolor='black', linewidth=1)
ax.set_yticks(range(len(top20_pen)))
ax.set_yticklabels(top20_pen['Agency'], fontsize=10)
ax.set_xlabel('HB Penetration Rate (%)', fontsize=14, weight='bold')
ax.set_title('Chart 11: HB Penetration Rate - Top 20 Agencies\nGreen >25% | Orange 10-25% | Red <10%',
             fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.grid(axis='x', alpha=0.3)
ax.axvline(x=35, color='green', linestyle='--', linewidth=2, label='Target: 35%')
ax.legend(fontsize=12)

# Add value labels
for i, bar in enumerate(bars):
    width = bar.get_width()
    ax.text(width, bar.get_y() + bar.get_height()/2.,
            f' {width:.1f}%',
            ha='left', va='center', fontsize=9, weight='bold')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/02_Agency_Analysis/11_hb_penetration_top20.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: HB Penetration by Top 20")

# Chart 12: Agency HB vs Non-HB Comparison (Stacked)
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))
top15_comp = agency_data.head(15).copy()

# Merge with HB data
top15_comp = top15_comp.merge(agency_hb_data, on='Agency', how='left').fillna(0)
top15_comp['Non_HB_Nights'] = top15_comp['Total_Nights'] - top15_comp['HB_Nights']

bars1 = ax.barh(range(len(top15_comp)), top15_comp['HB_Nights'],
                color=COLOR_HB, edgecolor='black', linewidth=1, label='Half Board')
bars2 = ax.barh(range(len(top15_comp)), top15_comp['Non_HB_Nights'],
                left=top15_comp['HB_Nights'], color=COLOR_NON_HB, edgecolor='black', linewidth=1, label='Non-Half Board')

ax.set_yticks(range(len(top15_comp)))
ax.set_yticklabels(top15_comp['Agency'], fontsize=10)
ax.set_xlabel('Room Nights', fontsize=14, weight='bold')
ax.set_title('Chart 12: Top 15 Agencies - HB vs Non-HB Room Nights\nStacked Comparison', fontsize=16, weight='bold', pad=20)
ax.invert_yaxis()
ax.legend(fontsize=12)
ax.grid(axis='x', alpha=0.3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/02_Agency_Analysis/12_agency_hb_nonhb_comparison.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: HB vs Non-HB Stacked Comparison")

# Chart 13: Agency Average Rate Comparison
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

# Calculate avg rates by agency
agency_rates = []
for agency in agency_data.head(10)['Agency']:
    agency_df = df[df['Search Name'] == agency]
    agency_hb_df = agency_df[agency_df['Has_HB']]
    agency_nonhb_df = agency_df[~agency_df['Has_HB']]

    agency_rates.append({
        'Agency': agency,
        'HB_Rate': agency_hb_df['Avg_Rate_Per_Night'].mean() if len(agency_hb_df) > 0 else 0,
        'Non_HB_Rate': agency_nonhb_df['Avg_Rate_Per_Night'].mean() if len(agency_nonhb_df) > 0 else 0
    })

df_rates = pd.DataFrame(agency_rates)

x = np.arange(len(df_rates))
width = 0.35

bars1 = ax.bar(x - width/2, df_rates['HB_Rate'], width, label='HB Avg Rate', color=COLOR_HB, edgecolor='black')
bars2 = ax.bar(x + width/2, df_rates['Non_HB_Rate'], width, label='Non-HB Avg Rate', color=COLOR_NON_HB, edgecolor='black')

ax.set_ylabel('Average Rate (AED)', fontsize=14, weight='bold')
ax.set_title('Chart 13: Average Rate Comparison - Top 10 Agencies\nHB vs Non-HB Rates', fontsize=16, weight='bold', pad=20)
ax.set_xticks(x)
ax.set_xticklabels(df_rates['Agency'], rotation=45, ha='right', fontsize=9)
ax.legend(fontsize=12)
ax.grid(axis='y', alpha=0.3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/02_Agency_Analysis/13_agency_avg_rate_comparison.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Average Rate Comparison")

# Chart 14: Agency Booking Size Distribution (Box Plot)
chart_count += 1
fig, ax = plt.subplots(figsize=(14, 10))

# Get top 10 agencies
top10_agencies = agency_data.head(10)['Agency'].tolist()
booking_data = []
labels = []

for agency in top10_agencies:
    agency_bookings = df[df['Search Name'] == agency]['Room Nights'].values
    booking_data.append(agency_bookings)
    labels.append(agency.split()[0][:15])  # Shorten names

bp = ax.boxplot(booking_data, labels=labels, patch_artist=True, showfliers=True)

for patch in bp['boxes']:
    patch.set_facecolor(COLOR_NEUTRAL)
    patch.set_alpha(0.7)

ax.set_ylabel('Booking Size (Room Nights)', fontsize=14, weight='bold')
ax.set_title('Chart 14: Booking Size Distribution - Top 10 Agencies\nBox Plot showing spread and outliers',
             fontsize=16, weight='bold', pad=20)
ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=9)
ax.grid(axis='y', alpha=0.3)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/charts/02_Agency_Analysis/14_agency_booking_size_boxplot.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"  ✓ Chart {chart_count}: Agency Booking Size Box Plot")

print(f"\nCompleted Category 2: {chart_count} charts created so far")

# Continue with remaining categories...
print("\n" + "="*80)
print(f"PROGRESS: {chart_count}/48 charts completed")
print("="*80)
