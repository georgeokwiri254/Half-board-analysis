import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
warnings.filterwarnings('ignore')

# Set style
sns.set_style("whitegrid")
sns.set_palette("husl")

# Read the data
df = pd.read_excel('/home/gee_devops254/Downloads/Half Board/Half Board.xlsx')
df['Avg_Rate_Per_Night'] = df['Room Revenue'] / df['Room Nights']
df['Is_HalfBoard'] = df['Product (Descriptions)'].str.contains('Halfboard|Half Board', case=False, na=False)
df_hb = df[df['Is_HalfBoard']]
df_cis = df[df['Rate Code'].str.contains('CIS', case=False, na=False)]
df_cis_hb = df_hb[df_hb['Rate Code'].str.contains('CIS', case=False, na=False)]

print("Creating visualizations...")

# ============================================================================
# 1. UNIVARIATE ANALYSIS VISUALIZATIONS
# ============================================================================
fig, axes = plt.subplots(2, 3, figsize=(18, 12))
fig.suptitle('Univariate Analysis - Distribution of Key Metrics', fontsize=16, fontweight='bold')

# Room Nights Distribution
axes[0, 0].hist(df['Room Nights'], bins=30, edgecolor='black', alpha=0.7)
axes[0, 0].set_xlabel('Room Nights')
axes[0, 0].set_ylabel('Frequency')
axes[0, 0].set_title('Distribution of Room Nights')
axes[0, 0].axvline(df['Room Nights'].mean(), color='red', linestyle='--', label=f'Mean: {df["Room Nights"].mean():.1f}')
axes[0, 0].axvline(df['Room Nights'].median(), color='green', linestyle='--', label=f'Median: {df["Room Nights"].median():.1f}')
axes[0, 0].legend()

# Room Revenue Distribution
axes[0, 1].hist(df['Room Revenue'], bins=30, edgecolor='black', alpha=0.7)
axes[0, 1].set_xlabel('Room Revenue (AED)')
axes[0, 1].set_ylabel('Frequency')
axes[0, 1].set_title('Distribution of Room Revenue')
axes[0, 1].axvline(df['Room Revenue'].mean(), color='red', linestyle='--', label=f'Mean: ${df["Room Revenue"].mean():,.0f}')
axes[0, 1].axvline(df['Room Revenue'].median(), color='green', linestyle='--', label=f'Median: ${df["Room Revenue"].median():,.0f}')
axes[0, 1].legend()

# Average Rate Distribution
axes[0, 2].hist(df['Avg_Rate_Per_Night'], bins=30, edgecolor='black', alpha=0.7)
axes[0, 2].set_xlabel('Average Rate per Night (AED)')
axes[0, 2].set_ylabel('Frequency')
axes[0, 2].set_title('Distribution of Average Rate per Night')
axes[0, 2].axvline(df['Avg_Rate_Per_Night'].mean(), color='red', linestyle='--', label=f'Mean: ${df["Avg_Rate_Per_Night"].mean():.2f}')
axes[0, 2].axvline(df['Avg_Rate_Per_Night'].median(), color='green', linestyle='--', label=f'Median: ${df["Avg_Rate_Per_Night"].median():.2f}')
axes[0, 2].legend()

# Box plots for outlier detection
axes[1, 0].boxplot(df['Room Nights'], vert=True)
axes[1, 0].set_ylabel('Room Nights')
axes[1, 0].set_title('Room Nights Box Plot')

axes[1, 1].boxplot(df['Room Revenue'], vert=True)
axes[1, 1].set_ylabel('Room Revenue (AED)')
axes[1, 1].set_title('Room Revenue Box Plot')

axes[1, 2].boxplot(df['Avg_Rate_Per_Night'], vert=True)
axes[1, 2].set_ylabel('Average Rate per Night (AED)')
axes[1, 2].set_title('Average Rate Box Plot')

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/1_univariate_analysis.png', dpi=300, bbox_inches='tight')
print("✓ Saved: 1_univariate_analysis.png")
plt.close()

# ============================================================================
# 2. HALF BOARD ANALYSIS VISUALIZATIONS
# ============================================================================
fig, axes = plt.subplots(2, 2, figsize=(16, 12))
fig.suptitle('Half Board Analysis', fontsize=16, fontweight='bold')

# HB vs Non-HB comparison
hb_comparison = pd.DataFrame({
    'Category': ['Half Board', 'Non-Half Board'],
    'Room Nights': [df_hb['Room Nights'].sum(), df[~df['Is_HalfBoard']]['Room Nights'].sum()],
    'Revenue': [df_hb['Room Revenue'].sum(), df[~df['Is_HalfBoard']]['Room Revenue'].sum()]
})

x = np.arange(len(hb_comparison['Category']))
width = 0.35

ax1 = axes[0, 0].twinx()
bars1 = axes[0, 0].bar(x - width/2, hb_comparison['Room Nights'], width, label='Room Nights', alpha=0.8)
bars2 = ax1.bar(x + width/2, hb_comparison['Revenue'], width, label='Revenue', alpha=0.8, color='orange')

axes[0, 0].set_xlabel('Category')
axes[0, 0].set_ylabel('Room Nights', color='tab:blue')
ax1.set_ylabel('Revenue (AED)', color='tab:orange')
axes[0, 0].set_title('Half Board vs Non-Half Board: Room Nights & Revenue')
axes[0, 0].set_xticks(x)
axes[0, 0].set_xticklabels(hb_comparison['Category'])
axes[0, 0].legend(loc='upper left')
ax1.legend(loc='upper right')

# Top 10 Agencies by HB Room Nights
hb_by_agency = df_hb.groupby('Search Name')['Room Nights'].sum().sort_values(ascending=False).head(10)
axes[0, 1].barh(range(len(hb_by_agency)), hb_by_agency.values, alpha=0.8)
axes[0, 1].set_yticks(range(len(hb_by_agency)))
axes[0, 1].set_yticklabels(hb_by_agency.index, fontsize=9)
axes[0, 1].set_xlabel('Room Nights')
axes[0, 1].set_title('Top 10 Agencies by Half Board Room Nights')
axes[0, 1].invert_yaxis()

# Top 10 Agencies by HB Revenue
hb_by_agency_rev = df_hb.groupby('Search Name')['Room Revenue'].sum().sort_values(ascending=False).head(10)
axes[1, 0].barh(range(len(hb_by_agency_rev)), hb_by_agency_rev.values, alpha=0.8, color='green')
axes[1, 0].set_yticks(range(len(hb_by_agency_rev)))
axes[1, 0].set_yticklabels(hb_by_agency_rev.index, fontsize=9)
axes[1, 0].set_xlabel('Revenue (AED)')
axes[1, 0].set_title('Top 10 Agencies by Half Board Revenue')
axes[1, 0].invert_yaxis()

# HB Rate Codes Distribution
hb_by_rate = df_hb.groupby('Rate Code')['Room Revenue'].sum().sort_values(ascending=False).head(10)
axes[1, 1].barh(range(len(hb_by_rate)), hb_by_rate.values, alpha=0.8, color='purple')
axes[1, 1].set_yticks(range(len(hb_by_rate)))
axes[1, 1].set_yticklabels(hb_by_rate.index, fontsize=9)
axes[1, 1].set_xlabel('Revenue (AED)')
axes[1, 1].set_title('Top 10 Rate Codes by Half Board Revenue')
axes[1, 1].invert_yaxis()

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/2_halfboard_analysis.png', dpi=300, bbox_inches='tight')
print("✓ Saved: 2_halfboard_analysis.png")
plt.close()

# ============================================================================
# 3. MULTIVARIATE ANALYSIS - AGENCY & RATE CODE
# ============================================================================
fig, axes = plt.subplots(2, 2, figsize=(16, 12))
fig.suptitle('Multivariate Analysis - Travel Agency & Rate Code Performance', fontsize=16, fontweight='bold')

# Top 15 Agencies by Revenue
agency_stats = df.groupby('Search Name').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum'
}).sort_values('Room Revenue', ascending=False).head(15)

axes[0, 0].barh(range(len(agency_stats)), agency_stats['Room Revenue'].values, alpha=0.8)
axes[0, 0].set_yticks(range(len(agency_stats)))
axes[0, 0].set_yticklabels(agency_stats.index, fontsize=8)
axes[0, 0].set_xlabel('Total Revenue (AED)')
axes[0, 0].set_title('Top 15 Travel Agencies by Total Revenue')
axes[0, 0].invert_yaxis()

# Top 15 Rate Codes by Revenue
rate_stats = df.groupby('Rate Code').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum'
}).sort_values('Room Revenue', ascending=False).head(15)

axes[0, 1].barh(range(len(rate_stats)), rate_stats['Room Revenue'].values, alpha=0.8, color='coral')
axes[0, 1].set_yticks(range(len(rate_stats)))
axes[0, 1].set_yticklabels(rate_stats.index, fontsize=9)
axes[0, 1].set_xlabel('Total Revenue (AED)')
axes[0, 1].set_title('Top 15 Rate Codes by Total Revenue')
axes[0, 1].invert_yaxis()

# Scatter: Room Nights vs Revenue by Agency
agency_scatter = df.groupby('Search Name').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum'
})
axes[1, 0].scatter(agency_scatter['Room Nights'], agency_scatter['Room Revenue'], alpha=0.6, s=100)
axes[1, 0].set_xlabel('Total Room Nights')
axes[1, 0].set_ylabel('Total Revenue (AED)')
axes[1, 0].set_title('Agency Performance: Room Nights vs Revenue')
axes[1, 0].grid(alpha=0.3)

# Top Agency-Rate Code Combinations
agency_rate = df.groupby(['Search Name', 'Rate Code'])['Room Revenue'].sum().sort_values(ascending=False).head(10)
combo_labels = [f"{agency}\n{rate}" for agency, rate in agency_rate.index]
axes[1, 1].barh(range(len(agency_rate)), agency_rate.values, alpha=0.8, color='teal')
axes[1, 1].set_yticks(range(len(agency_rate)))
axes[1, 1].set_yticklabels(combo_labels, fontsize=7)
axes[1, 1].set_xlabel('Revenue (AED)')
axes[1, 1].set_title('Top 10 Agency-Rate Code Combinations')
axes[1, 1].invert_yaxis()

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/3_multivariate_analysis.png', dpi=300, bbox_inches='tight')
print("✓ Saved: 3_multivariate_analysis.png")
plt.close()

# ============================================================================
# 4. CIS MARKET ANALYSIS
# ============================================================================
fig, axes = plt.subplots(2, 2, figsize=(16, 12))
fig.suptitle('CIS Market Analysis - Half Board Performance', fontsize=16, fontweight='bold')

# CIS Rate Codes Performance
if len(df_cis_hb) > 0:
    cis_rate_perf = df_cis_hb.groupby('Rate Code').agg({
        'Room Nights': 'sum',
        'Room Revenue': 'sum'
    }).sort_values('Room Revenue', ascending=False)

    axes[0, 0].bar(range(len(cis_rate_perf)), cis_rate_perf['Room Revenue'].values, alpha=0.8)
    axes[0, 0].set_xticks(range(len(cis_rate_perf)))
    axes[0, 0].set_xticklabels(cis_rate_perf.index, rotation=45, ha='right')
    axes[0, 0].set_ylabel('Revenue (AED)')
    axes[0, 0].set_title('CIS Rate Codes Revenue (Half Board)')

    axes[0, 1].bar(range(len(cis_rate_perf)), cis_rate_perf['Room Nights'].values, alpha=0.8, color='orange')
    axes[0, 1].set_xticks(range(len(cis_rate_perf)))
    axes[0, 1].set_xticklabels(cis_rate_perf.index, rotation=45, ha='right')
    axes[0, 1].set_ylabel('Room Nights')
    axes[0, 1].set_title('CIS Rate Codes Room Nights (Half Board)')

    # CIS Agencies Performance
    cis_agency_perf = df_cis_hb.groupby('Search Name').agg({
        'Room Nights': 'sum',
        'Room Revenue': 'sum'
    }).sort_values('Room Revenue', ascending=False)

    axes[1, 0].barh(range(len(cis_agency_perf)), cis_agency_perf['Room Revenue'].values, alpha=0.8, color='green')
    axes[1, 0].set_yticks(range(len(cis_agency_perf)))
    axes[1, 0].set_yticklabels(cis_agency_perf.index, fontsize=9)
    axes[1, 0].set_xlabel('Revenue (AED)')
    axes[1, 0].set_title('CIS Agencies Revenue (Half Board)')
    axes[1, 0].invert_yaxis()

    # Avg Rate Comparison
    cis_avg_rates = df_cis_hb.groupby('Rate Code').apply(
        lambda x: (x['Room Revenue'].sum() / x['Room Nights'].sum())
    ).sort_values(ascending=False)

    axes[1, 1].bar(range(len(cis_avg_rates)), cis_avg_rates.values, alpha=0.8, color='purple')
    axes[1, 1].set_xticks(range(len(cis_avg_rates)))
    axes[1, 1].set_xticklabels(cis_avg_rates.index, rotation=45, ha='right')
    axes[1, 1].set_ylabel('Average Rate (AED)')
    axes[1, 1].set_title('CIS Average Rates by Rate Code (Half Board)')
else:
    for ax in axes.flat:
        ax.text(0.5, 0.5, 'No CIS Half Board Data Available',
                ha='center', va='center', fontsize=14)
        ax.set_xlim([0, 1])
        ax.set_ylim([0, 1])

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/4_cis_market_analysis.png', dpi=300, bbox_inches='tight')
print("✓ Saved: 4_cis_market_analysis.png")
plt.close()

# ============================================================================
# 5. CORRELATION HEATMAP
# ============================================================================
fig, ax = plt.subplots(1, 1, figsize=(10, 8))
numeric_cols = ['Room Nights', 'Room Revenue', 'Avg_Rate_Per_Night']
correlation_matrix = df[numeric_cols].corr()

sns.heatmap(correlation_matrix, annot=True, fmt='.3f', cmap='coolwarm',
            square=True, linewidths=1, cbar_kws={"shrink": 0.8}, ax=ax)
ax.set_title('Correlation Matrix - Key Metrics', fontsize=14, fontweight='bold', pad=20)

plt.tight_layout()
plt.savefig('/home/gee_devops254/Downloads/Half Board/5_correlation_heatmap.png', dpi=300, bbox_inches='tight')
print("✓ Saved: 5_correlation_heatmap.png")
plt.close()

# ============================================================================
# 6. COMPREHENSIVE DASHBOARD
# ============================================================================
fig = plt.figure(figsize=(20, 12))
gs = fig.add_gridspec(3, 3, hspace=0.3, wspace=0.3)
fig.suptitle('Half Board Performance Dashboard', fontsize=18, fontweight='bold')

# Overall metrics
ax1 = fig.add_subplot(gs[0, 0])
metrics_data = {
    'Total\nBookings': len(df),
    'HB\nBookings': len(df_hb),
    'Total\nRevenue\n(k AED)': df['Room Revenue'].sum()/1000,
    'HB\nRevenue\n(k AED)': df_hb['Room Revenue'].sum()/1000
}
ax1.bar(range(len(metrics_data)), list(metrics_data.values()), alpha=0.8, color=['blue', 'green', 'orange', 'red'])
ax1.set_xticks(range(len(metrics_data)))
ax1.set_xticklabels(list(metrics_data.keys()), fontsize=9)
ax1.set_title('Key Metrics Overview')
ax1.set_ylabel('Count / Value')

# HB Penetration
ax2 = fig.add_subplot(gs[0, 1])
hb_penetration = [len(df_hb)/len(df)*100, (1-len(df_hb)/len(df))*100]
colors = ['#ff9999', '#66b3ff']
ax2.pie(hb_penetration, labels=['Half Board', 'Non-HB'], autopct='%1.1f%%',
        startangle=90, colors=colors)
ax2.set_title('Half Board Booking Penetration')

# Revenue Distribution
ax3 = fig.add_subplot(gs[0, 2])
revenue_dist = [df_hb['Room Revenue'].sum()/df['Room Revenue'].sum()*100,
                (1-df_hb['Room Revenue'].sum()/df['Room Revenue'].sum())*100]
ax3.pie(revenue_dist, labels=['Half Board', 'Non-HB'], autopct='%1.1f%%',
        startangle=90, colors=['#99ff99', '#ffcc99'])
ax3.set_title('Half Board Revenue Contribution')

# Top Agencies - Room Nights
ax4 = fig.add_subplot(gs[1, :])
top_agencies = df_hb.groupby('Search Name')['Room Nights'].sum().sort_values(ascending=False).head(15)
ax4.barh(range(len(top_agencies)), top_agencies.values, alpha=0.8)
ax4.set_yticks(range(len(top_agencies)))
ax4.set_yticklabels(top_agencies.index, fontsize=9)
ax4.set_xlabel('Room Nights')
ax4.set_title('Top 15 Agencies by Half Board Room Nights')
ax4.invert_yaxis()

# CIS Performance
ax5 = fig.add_subplot(gs[2, 0])
if len(df_cis_hb) > 0:
    cis_summary = df_cis_hb.groupby('Rate Code')['Room Revenue'].sum().sort_values(ascending=False)
    ax5.bar(range(len(cis_summary)), cis_summary.values, alpha=0.8, color='purple')
    ax5.set_xticks(range(len(cis_summary)))
    ax5.set_xticklabels(cis_summary.index, rotation=45, ha='right', fontsize=9)
    ax5.set_ylabel('Revenue (AED)')
    ax5.set_title('CIS Market Revenue by Rate Code (HB)')
else:
    ax5.text(0.5, 0.5, 'No CIS HB Data', ha='center', va='center', fontsize=12)

# Rate Code Performance
ax6 = fig.add_subplot(gs[2, 1])
top_rates_hb = df_hb.groupby('Rate Code')['Room Revenue'].sum().sort_values(ascending=False).head(10)
ax6.bar(range(len(top_rates_hb)), top_rates_hb.values, alpha=0.8, color='teal')
ax6.set_xticks(range(len(top_rates_hb)))
ax6.set_xticklabels(top_rates_hb.index, rotation=45, ha='right', fontsize=8)
ax6.set_ylabel('Revenue (AED)')
ax6.set_title('Top 10 Rate Codes (HB Revenue)')

# Average Rate Comparison
ax7 = fig.add_subplot(gs[2, 2])
avg_rate_comp = pd.DataFrame({
    'Category': ['Overall', 'Half Board', 'CIS HB'],
    'Avg Rate': [
        df['Avg_Rate_Per_Night'].mean(),
        df_hb['Avg_Rate_Per_Night'].mean(),
        df_cis_hb['Avg_Rate_Per_Night'].mean() if len(df_cis_hb) > 0 else 0
    ]
})
bars = ax7.bar(range(len(avg_rate_comp)), avg_rate_comp['Avg Rate'].values,
               alpha=0.8, color=['blue', 'green', 'purple'])
ax7.set_xticks(range(len(avg_rate_comp)))
ax7.set_xticklabels(avg_rate_comp['Category'], rotation=15)
ax7.set_ylabel('Average Rate (AED)')
ax7.set_title('Average Rate Comparison')
# Add value labels on bars
for i, bar in enumerate(bars):
    height = bar.get_height()
    ax7.text(bar.get_x() + bar.get_width()/2., height,
             f'${height:.0f}', ha='center', va='bottom', fontsize=9)

plt.savefig('/home/gee_devops254/Downloads/Half Board/6_comprehensive_dashboard.png', dpi=300, bbox_inches='tight')
print("✓ Saved: 6_comprehensive_dashboard.png")
plt.close()

print("\n" + "="*80)
print("All visualizations created successfully!")
print("="*80)
print("\nGenerated files:")
print("  1. 1_univariate_analysis.png - Distribution analysis of all key metrics")
print("  2. 2_halfboard_analysis.png - Half Board specific performance")
print("  3. 3_multivariate_analysis.png - Agency & Rate Code relationships")
print("  4. 4_cis_market_analysis.png - CIS market performance")
print("  5. 5_correlation_heatmap.png - Correlation between metrics")
print("  6. 6_comprehensive_dashboard.png - Complete overview dashboard")
print("\nCSV Reports:")
print("  - hb_agency_analysis.csv")
print("  - agency_performance.csv")
print("  - ratecode_performance.csv")
print("  - cis_market_analysis.csv")
print("  - cis_agency_analysis.csv")
print("="*80)
