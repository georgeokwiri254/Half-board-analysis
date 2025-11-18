import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
import warnings
warnings.filterwarnings('ignore')

# Set style
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (12, 8)

# Read the data
df = pd.read_excel('/home/gee_devops254/Downloads/Half Board/Half Board.xlsx')

print("="*80)
print("HALF BOARD STATISTICAL ANALYSIS")
print("="*80)

# ============================================================================
# DATA PREPARATION
# ============================================================================
print("\n1. DATA OVERVIEW")
print("-"*80)
print(f"Total Records: {len(df)}")
print(f"Columns: {', '.join(df.columns)}")
print(f"\nMissing Values:\n{df.isnull().sum()}")

# Filter Half Board records (case insensitive) - looking for "Halfboard" or "Half Board"
df_hb = df[df['Product (Descriptions)'].str.contains('Halfboard|Half Board', case=False, na=False)]
print(f"\nTotal Half Board Records: {len(df_hb)} out of {len(df)} ({len(df_hb)/len(df)*100:.1f}%)")

# Also show what we found
print(f"\nHalf Board Product Types Found:")
print(df_hb['Product (Descriptions)'].value_counts())

# ============================================================================
# UNIVARIATE ANALYSIS
# ============================================================================
print("\n\n" + "="*80)
print("2. UNIVARIATE ANALYSIS - ALL DATA")
print("="*80)

print("\n2.1 ROOM NIGHTS STATISTICS")
print("-"*80)
print(df['Room Nights'].describe())
print(f"\nSkewness: {df['Room Nights'].skew():.3f}")
print(f"Kurtosis: {df['Room Nights'].kurtosis():.3f}")
print(f"Variance: {df['Room Nights'].var():.3f}")
print(f"Coefficient of Variation: {(df['Room Nights'].std()/df['Room Nights'].mean())*100:.2f}%")

print("\n2.2 ROOM REVENUE STATISTICS")
print("-"*80)
print(df['Room Revenue'].describe())
print(f"\nSkewness: {df['Room Revenue'].skew():.3f}")
print(f"Kurtosis: {df['Room Revenue'].kurtosis():.3f}")
print(f"Variance: {df['Room Revenue'].var():.3f}")
print(f"Coefficient of Variation: {(df['Room Revenue'].std()/df['Room Revenue'].mean())*100:.2f}%")

print("\n2.3 AVERAGE RATE PER NIGHT")
print("-"*80)
df['Avg_Rate_Per_Night'] = df['Room Revenue'] / df['Room Nights']
print(df['Avg_Rate_Per_Night'].describe())
print(f"\nSkewness: {df['Avg_Rate_Per_Night'].skew():.3f}")
print(f"Kurtosis: {df['Avg_Rate_Per_Night'].kurtosis():.3f}")

print("\n2.4 CATEGORICAL VARIABLES FREQUENCY")
print("-"*80)
print(f"\nTotal Unique Travel Agencies: {df['Search Name'].nunique()}")
print(f"Total Unique Rate Codes: {df['Rate Code'].nunique()}")
print(f"Total Unique Products: {df['Product (Descriptions)'].nunique()}")

print("\n\nTop 10 Travel Agencies by Frequency:")
print(df['Search Name'].value_counts().head(10))

print("\n\nTop 10 Rate Codes by Frequency:")
print(df['Rate Code'].value_counts().head(10))

# ============================================================================
# HALF BOARD SPECIFIC ANALYSIS
# ============================================================================
print("\n\n" + "="*80)
print("3. HALF BOARD SPECIFIC ANALYSIS")
print("="*80)

print("\n3.1 HALF BOARD ROOM NIGHTS STATISTICS")
print("-"*80)
print(df_hb['Room Nights'].describe())
print(f"\nTotal Half Board Room Nights: {df_hb['Room Nights'].sum():,}")
print(f"Percentage of Total Room Nights: {df_hb['Room Nights'].sum()/df['Room Nights'].sum()*100:.2f}%")

print("\n3.2 HALF BOARD REVENUE STATISTICS")
print("-"*80)
print(df_hb['Room Revenue'].describe())
print(f"\nTotal Half Board Revenue: ${df_hb['Room Revenue'].sum():,.2f}")
print(f"Percentage of Total Revenue: {df_hb['Room Revenue'].sum()/df['Room Revenue'].sum()*100:.2f}%")

print("\n3.3 HALF BOARD AVERAGE RATE")
print("-"*80)
df_hb['Avg_Rate_Per_Night'] = df_hb['Room Revenue'] / df_hb['Room Nights']
print(df_hb['Avg_Rate_Per_Night'].describe())

print("\n3.4 TOP 10 AGENCIES BY HALF BOARD ROOM NIGHTS")
print("-"*80)
hb_by_agency = df_hb.groupby('Search Name').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum'
}).sort_values('Room Nights', ascending=False)
hb_by_agency['Avg_Rate'] = hb_by_agency['Room Revenue'] / hb_by_agency['Room Nights']
print(hb_by_agency.head(10))

print("\n3.5 TOP 10 AGENCIES BY HALF BOARD REVENUE")
print("-"*80)
print(hb_by_agency.sort_values('Room Revenue', ascending=False).head(10))

# ============================================================================
# MULTIVARIATE ANALYSIS - TRAVEL AGENCY AND RATE CODE
# ============================================================================
print("\n\n" + "="*80)
print("4. MULTIVARIATE ANALYSIS - TRAVEL AGENCY & RATE CODE")
print("="*80)

print("\n4.1 CORRELATION ANALYSIS")
print("-"*80)
numeric_cols = ['Room Nights', 'Room Revenue', 'Avg_Rate_Per_Night']
correlation_matrix = df[numeric_cols].corr()
print(correlation_matrix)

print("\n4.2 TRAVEL AGENCY PERFORMANCE")
print("-"*80)
agency_stats = df.groupby('Search Name').agg({
    'Room Nights': ['sum', 'mean', 'count'],
    'Room Revenue': ['sum', 'mean'],
}).round(2)
agency_stats.columns = ['Total_Nights', 'Avg_Nights', 'Transactions', 'Total_Revenue', 'Avg_Revenue']
agency_stats['Avg_Rate'] = (agency_stats['Total_Revenue'] / agency_stats['Total_Nights']).round(2)
agency_stats = agency_stats.sort_values('Total_Revenue', ascending=False)
print("\nTop 15 Agencies by Total Revenue:")
print(agency_stats.head(15))

print("\n4.3 RATE CODE PERFORMANCE")
print("-"*80)
ratecode_stats = df.groupby('Rate Code').agg({
    'Room Nights': ['sum', 'mean', 'count'],
    'Room Revenue': ['sum', 'mean'],
}).round(2)
ratecode_stats.columns = ['Total_Nights', 'Avg_Nights', 'Transactions', 'Total_Revenue', 'Avg_Revenue']
ratecode_stats['Avg_Rate'] = (ratecode_stats['Total_Revenue'] / ratecode_stats['Total_Nights']).round(2)
ratecode_stats = ratecode_stats.sort_values('Total_Revenue', ascending=False)
print("\nTop 15 Rate Codes by Total Revenue:")
print(ratecode_stats.head(15))

print("\n4.4 AGENCY-RATE CODE COMBINATION ANALYSIS")
print("-"*80)
agency_rate_combo = df.groupby(['Search Name', 'Rate Code']).agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum'
}).round(2)
agency_rate_combo['Avg_Rate'] = (agency_rate_combo['Room Revenue'] / agency_rate_combo['Room Nights']).round(2)
agency_rate_combo = agency_rate_combo.sort_values('Room Revenue', ascending=False)
print("\nTop 15 Agency-Rate Code Combinations by Revenue:")
print(agency_rate_combo.head(15))

# ============================================================================
# CIS MARKET ANALYSIS
# ============================================================================
print("\n\n" + "="*80)
print("5. CIS MARKET ANALYSIS (HALF BOARD)")
print("="*80)

# Filter CIS rate codes
df_cis = df[df['Rate Code'].str.contains('CIS', case=False, na=False)]
df_cis_hb = df_hb[df_hb['Rate Code'].str.contains('CIS', case=False, na=False)]

print(f"\n5.1 CIS MARKET OVERVIEW")
print("-"*80)
print(f"Total CIS Records: {len(df_cis)}")
print(f"CIS Half Board Records: {len(df_cis_hb)}")
print(f"CIS Half Board Penetration: {len(df_cis_hb)/len(df_cis)*100:.2f}%")

print("\n5.2 CIS RATE CODES IN HALF BOARD")
print("-"*80)
cis_rate_analysis = df_cis_hb.groupby('Rate Code').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum',
    'Search Name': 'count'
}).round(2)
cis_rate_analysis.columns = ['Total_Room_Nights', 'Total_Revenue', 'Num_Bookings']
cis_rate_analysis['Avg_Rate'] = (cis_rate_analysis['Total_Revenue'] / cis_rate_analysis['Total_Room_Nights']).round(2)
cis_rate_analysis['Avg_Nights_Per_Booking'] = (cis_rate_analysis['Total_Room_Nights'] / cis_rate_analysis['Num_Bookings']).round(2)
cis_rate_analysis = cis_rate_analysis.sort_values('Total_Revenue', ascending=False)
print(cis_rate_analysis)

print("\n5.3 CIS MARKET BY TRAVEL AGENCY (HALF BOARD)")
print("-"*80)
cis_agency_analysis = df_cis_hb.groupby('Search Name').agg({
    'Room Nights': 'sum',
    'Room Revenue': 'sum',
    'Rate Code': 'count'
}).round(2)
cis_agency_analysis.columns = ['Total_Room_Nights', 'Total_Revenue', 'Num_Bookings']
cis_agency_analysis['Avg_Rate'] = (cis_agency_analysis['Total_Revenue'] / cis_agency_analysis['Total_Room_Nights']).round(2)
cis_agency_analysis = cis_agency_analysis.sort_values('Total_Revenue', ascending=False)
print("\nTop Agencies in CIS Half Board Market:")
print(cis_agency_analysis)

print("\n5.4 BEST PERFORMING CIS MARKET")
print("-"*80)
if len(cis_rate_analysis) > 0:
    best_cis_rate = cis_rate_analysis.sort_values('Total_Revenue', ascending=False).iloc[0]
    print(f"Best CIS Rate Code by Revenue: {best_cis_rate.name}")
    print(f"  Total Revenue: ${best_cis_rate['Total_Revenue']:,.2f}")
    print(f"  Total Room Nights: {best_cis_rate['Total_Room_Nights']:,.0f}")
    print(f"  Number of Bookings: {best_cis_rate['Num_Bookings']:.0f}")
    print(f"  Average Rate: ${best_cis_rate['Avg_Rate']:.2f}")
    print(f"  Avg Nights per Booking: {best_cis_rate['Avg_Nights_Per_Booking']:.2f}")

    best_cis_nights = cis_rate_analysis.sort_values('Total_Room_Nights', ascending=False).iloc[0]
    if best_cis_nights.name != best_cis_rate.name:
        print(f"\nBest CIS Rate Code by Room Nights: {best_cis_nights.name}")
        print(f"  Total Room Nights: {best_cis_nights['Total_Room_Nights']:,.0f}")
        print(f"  Total Revenue: ${best_cis_nights['Total_Revenue']:,.2f}")
else:
    print("No CIS Half Board bookings found in the dataset.")

# ============================================================================
# STATISTICAL TESTS
# ============================================================================
print("\n\n" + "="*80)
print("6. STATISTICAL TESTS")
print("="*80)

print("\n6.1 HALF BOARD vs NON-HALF BOARD COMPARISON")
print("-"*80)
df['Is_HalfBoard'] = df['Product (Descriptions)'].str.contains('Halfboard|Half Board', case=False, na=False)
df_non_hb = df[~df['Is_HalfBoard']]

print("\nComparison of Average Rates:")
print(f"Half Board Avg Rate: ${df_hb['Avg_Rate_Per_Night'].mean():.2f}")
print(f"Non-Half Board Avg Rate: ${df_non_hb['Avg_Rate_Per_Night'].mean():.2f}")

# T-test
t_stat, p_value = stats.ttest_ind(df_hb['Avg_Rate_Per_Night'].dropna(),
                                   df_non_hb['Avg_Rate_Per_Night'].dropna())
print(f"\nIndependent T-Test Results:")
print(f"  T-statistic: {t_stat:.4f}")
print(f"  P-value: {p_value:.4f}")
print(f"  Significant difference: {'Yes' if p_value < 0.05 else 'No'} (α=0.05)")

# Mann-Whitney U test (non-parametric alternative)
u_stat, p_value_u = stats.mannwhitneyu(df_hb['Avg_Rate_Per_Night'].dropna(),
                                        df_non_hb['Avg_Rate_Per_Night'].dropna())
print(f"\nMann-Whitney U Test Results:")
print(f"  U-statistic: {u_stat:.4f}")
print(f"  P-value: {p_value_u:.4f}")

print("\n6.2 ANOVA - RATE CODE COMPARISON (TOP 10)")
print("-"*80)
top_rate_codes = df['Rate Code'].value_counts().head(10).index
df_top_rates = df[df['Rate Code'].isin(top_rate_codes)]
rate_groups = [group['Avg_Rate_Per_Night'].dropna() for name, group in df_top_rates.groupby('Rate Code')]
f_stat, p_value_anova = stats.f_oneway(*rate_groups)
print(f"F-statistic: {f_stat:.4f}")
print(f"P-value: {p_value_anova:.4f}")
print(f"Significant difference across rate codes: {'Yes' if p_value_anova < 0.05 else 'No'} (α=0.05)")

# ============================================================================
# SUMMARY INSIGHTS
# ============================================================================
print("\n\n" + "="*80)
print("7. KEY INSIGHTS & SUMMARY")
print("="*80)

total_revenue = df['Room Revenue'].sum()
total_nights = df['Room Nights'].sum()
hb_revenue = df_hb['Room Revenue'].sum()
hb_nights = df_hb['Room Nights'].sum()

print(f"\n• Half Board represents {len(df_hb)/len(df)*100:.1f}% of all bookings")
print(f"• Half Board generates {hb_revenue/total_revenue*100:.1f}% of total revenue (${hb_revenue:,.2f})")
print(f"• Half Board accounts for {hb_nights/total_nights*100:.1f}% of room nights ({hb_nights:,} nights)")
print(f"• Average Half Board rate: ${df_hb['Avg_Rate_Per_Night'].mean():.2f} vs Overall: ${df['Avg_Rate_Per_Night'].mean():.2f}")
print(f"• Top HB Agency: {hb_by_agency.index[0]} with {hb_by_agency.iloc[0]['Room Nights']:.0f} room nights")
print(f"• Top HB Rate Code: {df_hb.groupby('Rate Code')['Room Revenue'].sum().idxmax()}")

if len(df_cis_hb) > 0:
    print(f"\n• CIS Market in Half Board: {len(df_cis_hb)} bookings, ${df_cis_hb['Room Revenue'].sum():,.2f} revenue")
    if len(cis_rate_analysis) > 0:
        best_cis_rate_name = cis_rate_analysis.sort_values('Total_Revenue', ascending=False).index[0]
        best_cis_rate_rev = cis_rate_analysis.sort_values('Total_Revenue', ascending=False).iloc[0]['Total_Revenue']
        print(f"• Best CIS Rate Code: {best_cis_rate_name} with ${best_cis_rate_rev:,.2f} revenue")
    print(f"• CIS Average Rate: ${df_cis_hb['Avg_Rate_Per_Night'].mean():.2f}")
else:
    print(f"\n• No CIS bookings found with Half Board in this dataset")

print("\n" + "="*80)
print("Analysis complete. Creating visualizations...")
print("="*80)

# Save detailed reports
print("\nSaving detailed CSV reports...")
hb_by_agency.to_csv('/home/gee_devops254/Downloads/Half Board/hb_agency_analysis.csv')
agency_stats.to_csv('/home/gee_devops254/Downloads/Half Board/agency_performance.csv')
ratecode_stats.to_csv('/home/gee_devops254/Downloads/Half Board/ratecode_performance.csv')
if len(cis_rate_analysis) > 0:
    cis_rate_analysis.to_csv('/home/gee_devops254/Downloads/Half Board/cis_market_analysis.csv')
    cis_agency_analysis.to_csv('/home/gee_devops254/Downloads/Half Board/cis_agency_analysis.csv')

print("CSV reports saved successfully!")
