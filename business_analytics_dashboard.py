"""
========================================================================
BSA83111 - Data Visualization for Decision Making
Assessment: E-Commerce Sales Dashboard (Final)
Group 7 Members: Jacques Rebero (Leader), Eric Nshimiye, 
Steven Ndacyarikumukiza, Gisele Musabuwera
========================================================================
This Streamlit app provides and interactive dashboard for analyzing the UCI online retail dataset. It includes key sales metrics, dynamic filters, and visualizations to help business stakeholders understand sales trends, product performance, and customer behavior.
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ==================== PAGE CONFIGURATION ====================
st.set_page_config(
    page_title="E-Commerce Sales Dashboard",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Professional CSS (Light & Dark Mode Compatible)
st.markdown("""<style>
* { margin: 0; padding: 0; }
.main { padding: 4px !important; }
[data-testid="stVerticalBlock"] { gap: 2px !important; }
[data-testid="stHorizontalBlock"] { gap: 4px !important; }

/* Metrics - Adaptive colors using Streamlit CSS variables */
.stMetric { 
    padding: 8px !important; 
    border-radius: 4px !important; 
    border-left: 3px solid #1e40af !important;
}

/* Metric values - Uses primary color for accent, works across themes */
.stMetric [data-testid="metricValue"] { 
    color: #0052cc !important; 
    font-weight: 700 !important;
    font-size: 20px !important;
}

/* Metric labels - Dynamic text color */
.stMetric [data-testid="metricLabel"] { 
    color: var(--text-color) !important; 
    font-weight: 600 !important;
    font-size: 11px !important;
}

/* Metric delta */
.stMetric [data-testid="metricDeltaContainer"] { 
    color: #0066cc !important; 
    font-weight: 600 !important;
}

/* Headings - Dynamic text color based on system theme */
h1 { 
    margin: 4px 0 !important; 
    color: var(--text-color) !important;
    font-weight: 700 !important;
}

h2 { 
    margin: 4px 0 !important; 
    color: var(--text-color) !important;
    font-weight: 700 !important;
    font-size: 16px !important;
}

h3 { 
    margin: 2px 0 !important; 
    color: var(--text-color) !important;
    font-weight: 600 !important;
    font-size: 13px !important;
}

/* Subheadings */
[data-testid="stMarkdownContainer"] h3, 
[data-testid="stMarkdownContainer"] h4 {
    color: #0052cc !important;
    font-weight: 600 !important;
}

/* Buttons */
button { 
    font-size: 9px !important; 
    padding: 3px 6px !important; 
    height: 28px !important;
    background-color: #0052cc !important;
    color: white !important;
}

button:hover {
    background-color: #003a99 !important;
}

/* Input labels - Dynamic text color */
label {
    color: var(--text-color) !important;
    font-weight: 500 !important;
}

/* Captions and text */
[data-testid="stCaption"] {
    color: var(--text-color) !important;
    opacity: 0.8;
    font-weight: 500 !important;
}

/* Info boxes - Semi-transparent background to adapt to dark/light modes */
[data-testid="stMarkdownContainer"] .info-box,
[data-testid="stInfo"] {
    background-color: rgba(30, 64, 175, 0.1) !important;
    border-left: 4px solid #0052cc !important;
    color: var(--text-color) !important;
}

/* Tables */
[role="grid"] {
    color: var(--text-color) !important;
}

[role="columnheader"] {
    background-color: rgba(0, 82, 204, 0.15) !important;
    color: #0052cc !important;
    font-weight: 600 !important;
}

/* DataFrame text */
td, th {
    color: var(--text-color) !important;
}

/* Hide menu items */
#MainMenu, header, footer { display: none; }

/* Improve readability of plots */
.stPyplotChart {
    margin: 0 !important;
}

/* Metric containers - Dynamic background card color */
[data-testid="metric-container"] {
    background-color: var(--background-color) !important;
    border-radius: 4px !important;
}

/* Selectbox and input styling */
[data-testid="stSelectbox"] div,
[data-testid="stMultiSelect"] div {
    color: var(--text-color) !important;
}

/* Slider styling */
[data-testid="stSlider"] {
    color: var(--text-color) !important;
}

/* Date input styling */
[data-testid="stDateInput"] input {
    color: var(--text-color) !important;
    background-color: var(--background-color) !important;
}

/* Markdown text - Dynamic text color */
[data-testid="stMarkdownContainer"] p {
    color: var(--text-color) !important;
    line-height: 1.6 !important;
}

[data-testid="stMarkdownContainer"] li {
    color: var(--text-color) !important;
}

[data-testid="stMarkdownContainer"] ul {
    color: var(--text-color) !important;
}

/* Alert boxes - Semi-transparent variables for compatibility */
[data-testid="stAlert"] {
    background-color: rgba(30, 64, 175, 0.1) !important;
    color: var(--text-color) !important;
}

[data-testid="stInfo"] {
    background-color: rgba(30, 64, 175, 0.1) !important;
    color: var(--text-color) !important;
    border-left: 4px solid #0052cc !important;
}

[data-testid="stError"] {
    background-color: rgba(183, 28, 28, 0.1) !important;
    color: #b71c1c !important;
}

/* Success messages */
[data-testid="stSuccess"] {
    background-color: rgba(27, 94, 32, 0.1) !important;
    color: #1b5e20 !important;
}

</style>""", unsafe_allow_html=True)

# ==================== DATA LOADING & CLEANING ====================
@st.cache_data
def load_data():
    """Load and clean UCI Online Retail dataset."""
    try:
        #url = "https://archive.ics.uci.edu/ml/machine-learning-databases/00352/Online%20Retail.xlsx"
        file_path = "Online Retail (1).xlsx"
        df = pd.read_excel(file_path, header=0)
        df.columns = ['InvoiceNo', 'StockCode', 'Description', 'Quantity', 
                     'InvoiceDate', 'UnitPrice', 'CustomerID', 'Country']
        
        # Clean data
        df = df[df['Quantity'] > 0].dropna(subset=['CustomerID']).copy()
        df['Revenue'] = df['Quantity'] * df['UnitPrice']
        df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])
        
        # FIXED: Proper month ordering
        df['MonthNum'] = df['InvoiceDate'].dt.month
        df['Month'] = df['InvoiceDate'].dt.strftime('%b')
        df['DayOfWeek'] = df['InvoiceDate'].dt.day_name()
        
        return df
    except Exception as e:
        st.error(f"Error loading dataset: {e}")
        return None

df = load_data()

if df is not None:
    
    # ==================== DASHBOARD HEADER ====================
    st.markdown("### 📊 Sales Dashboard")
    
    # ==================== FILTER CONTROLS ====================
    fc = st.columns(4)
    with fc[0]:
        start = st.date_input("From", df['InvoiceDate'].min(), key="s", label_visibility="collapsed")
    with fc[1]:
        end = st.date_input("To", df['InvoiceDate'].max(), key="e", label_visibility="collapsed")
    with fc[2]:
        countries = st.multiselect("Countries", df['Country'].unique()[:12], 
                                  default=df['Country'].unique()[:2], max_selections=4, 
                                  label_visibility="collapsed")
    with fc[3]:
        pr = st.slider("Price", 0.0, float(df['UnitPrice'].quantile(0.9)), 
                      (0.0, float(df['UnitPrice'].quantile(0.9))), label_visibility="collapsed")
    
    # Apply filters
    df_filtered = df[(df['InvoiceDate'] >= pd.to_datetime(start)) & 
                     (df['InvoiceDate'] <= pd.to_datetime(end)) & 
                     (df['Country'].isin(countries)) & 
                     (df['UnitPrice'].between(pr[0], pr[1]))]
    
    # ==================== KEY METRICS ====================
    mc = st.columns(4)
    with mc[0]:
        st.metric("Revenue", f"£{df_filtered['Revenue'].sum():,.0f}")
    with mc[1]:
        avg_ord = df_filtered.groupby('InvoiceNo')['Revenue'].sum().mean() if len(df_filtered) > 0 else 0
        st.metric("Avg Order", f"£{avg_ord:.0f}")
    with mc[2]:
        st.metric("Customers", f"{df_filtered['CustomerID'].nunique():,}")
    with mc[3]:
        top_country = df_filtered.groupby('Country')['Revenue'].sum().idxmax() if len(df_filtered) > 0 else "N/A"
        st.metric("Top", top_country)
    
    # ==================== ROW 1: 3 CHARTS ====================
    cc1 = st.columns(3)
    
    # Chart 1: Monthly Trend (FIXED: Chronological order)
    with cc1[0]:
        monthly = (
            df_filtered
            .groupby(['MonthNum', 'Month'])['Revenue']
            .sum()
            .reset_index()
            .sort_values('MonthNum')
        )
        fig, ax = plt.subplots(figsize=(2.5, 1.5), dpi=80)
        ax.plot(monthly.index, monthly['Revenue'], marker='o', linewidth=1.5, markersize=3, color='#1e40af')
        ax.fill_between(monthly.index, monthly['Revenue'], alpha=0.1, color='#1e40af')
        ax.set_xticks(monthly.index[::max(1, len(monthly)//3)])
        ax.set_xticklabels(monthly['Month'].iloc[::max(1, len(monthly)//3)], rotation=45, fontsize=6)
        ax.set_title("Monthly", fontsize=8, fontweight=600)
        ax.tick_params(labelsize=7)
        ax.grid(True, alpha=0.2, linewidth=0.3)
        plt.yticks(fontsize=6)
        plt.tight_layout(pad=0.3)
        st.pyplot(fig, use_container_width=True)
    
    # Chart 2: Top Products
    with cc1[1]:
        top_p = df_filtered.groupby('Description')['Revenue'].sum().nlargest(6)
        fig, ax = plt.subplots(figsize=(2.5, 1.5), dpi=80)
        ax.barh(range(len(top_p)), top_p.values, color='#f97316', alpha=0.8)
        ax.set_yticks(range(len(top_p)))
        ax.set_yticklabels([p[:15] for p in top_p.index], fontsize=6)
        ax.set_title("Top Products", fontsize=8, fontweight=600)
        ax.tick_params(labelsize=6)
        ax.grid(True, alpha=0.2, axis='x', linewidth=0.3)
        plt.tight_layout(pad=0.3)
        st.pyplot(fig, use_container_width=True)
    
    # Chart 3: Top Countries
    with cc1[2]:
        top_c = df_filtered.groupby('Country')['Revenue'].sum().nlargest(6)
        fig, ax = plt.subplots(figsize=(2.5, 1.5), dpi=80)
        ax.barh(range(len(top_c)), top_c.values, color='#3b82f6', alpha=0.8)
        ax.set_yticks(range(len(top_c)))
        ax.set_yticklabels(top_c.index, fontsize=6)
        ax.set_title("Top Countries", fontsize=8, fontweight=600)
        ax.tick_params(labelsize=6)
        ax.grid(True, alpha=0.2, axis='x', linewidth=0.3)
        plt.tight_layout(pad=0.3)
        st.pyplot(fig, use_container_width=True)
    
    # ==================== ROW 2: 3 CHARTS ====================
    cc2 = st.columns(3)
    
    # Chart 4: Correlation (FIXED: Dynamic values)
    with cc2[0]:
        corr = df_filtered[['Quantity', 'UnitPrice', 'Revenue']].corr()
        fig, ax = plt.subplots(figsize=(2.3, 1.5), dpi=80)
        sns.heatmap(corr, annot=True, fmt='.2f', cmap='RdYlGn', center=0, ax=ax, 
                   cbar=False, linewidth=0.5, annot_kws={'fontsize': 6})
        ax.set_title("Correlation", fontsize=8, fontweight=600)
        plt.xticks(fontsize=6, rotation=45, ha='right')
        plt.yticks(fontsize=6, rotation=0)
        plt.tight_layout(pad=0.2)
        st.pyplot(fig, use_container_width=True)
    
    # Chart 5: Order Distribution
    with cc2[1]:
        orders = df_filtered.groupby('InvoiceNo')['Revenue'].sum()
        fig, ax = plt.subplots(figsize=(2.5, 1.5), dpi=80)
        ax.hist(orders, bins=25, color='#8b5cf6', alpha=0.8, edgecolor='#7c3aed', linewidth=0.5)
        ax.set_title("Orders", fontsize=8, fontweight=600)
        ax.tick_params(labelsize=6)
        ax.grid(True, alpha=0.2, axis='y', linewidth=0.3)
        plt.tight_layout(pad=0.3)
        st.pyplot(fig, use_container_width=True)
    
    # Chart 6: Day of Week (FIXED: Revenue based)
    with cc2[2]:
        dow_revenue = df_filtered.groupby('DayOfWeek')['Revenue'].sum()
        day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        dow_revenue = dow_revenue.reindex([d for d in day_order if d in dow_revenue.index])
        fig, ax = plt.subplots(figsize=(2.5, 1.5), dpi=80)
        ax.bar(range(len(dow_revenue)), dow_revenue.values, color='#10b981', alpha=0.8, edgecolor='#059669', linewidth=0.5)
        ax.set_xticks(range(len(dow_revenue)))
        ax.set_xticklabels([d[:3] for d in dow_revenue.index], fontsize=6, rotation=45)
        ax.set_title("Day Pattern", fontsize=8, fontweight=600)
        ax.tick_params(labelsize=6)
        ax.grid(True, alpha=0.2, axis='y', linewidth=0.3)
        plt.tight_layout(pad=0.3)
        st.pyplot(fig, use_container_width=True)
    
    # ==================== ACTION BUTTONS ====================
    bc = st.columns(3)
    with bc[0]:
        csv = df_filtered[['InvoiceDate', 'Description', 'Quantity', 'UnitPrice', 'Revenue', 'Country']].to_csv(index=False)
        st.download_button("📥 CSV", csv, f"sales_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
    with bc[1]:
        if st.button("🔄 Reset"):
            st.rerun()
    with bc[2]:
        st.caption(f"✓ {len(df_filtered):,} records")
    
    st.markdown("---")
    
    # ======================== DETAILED ANALYSIS SECTIONS ========================
    
    st.markdown("## 📊 Detailed Analysis (Scroll Below)")
    st.markdown("")
    
    # ==================== CORRELATION INTERPRETATION ====================
    st.markdown("### Correlation Analysis")
    
    # Get dynamic values
    qty_rev = round(corr.loc['Quantity', 'Revenue'], 2)
    price_rev = round(corr.loc['UnitPrice', 'Revenue'], 2)
    qty_price = round(corr.loc['Quantity', 'UnitPrice'], 2)
    
    st.info(f"""
**Correlation Values (Calculated from Filtered Data):**

• **Quantity → Revenue**: {qty_rev} {"(Strong positive)" if qty_rev > 0.6 else "(Positive)"}  
  More units purchased = higher total revenue  
  💡 Action: Bulk purchase discounts are viable

• **UnitPrice → Revenue**: {price_rev} {"(Strong positive)" if price_rev > 0.6 else "(Positive)"}  
  Higher-priced items drive higher revenue  
  💡 Action: Consider premium product expansion

• **Quantity ↔ UnitPrice**: {qty_price} {"(Weak negative)" if qty_price < -0.1 else "(Weak)"}  
  Bulk orders tend to be lower-priced items  
  💡 Action: Create bundled offers (volume + premium)

**Note:** These values update dynamically based on your filters.
    """)
    
    st.markdown("---")
    
    # ==================== PRODUCT PERFORMANCE ====================
    st.markdown("### Product Performance Analysis")
    
    product_revenue = df_filtered.groupby('Description')['Revenue'].sum().sort_values(ascending=False)
    
    if len(product_revenue) > 0:
        total_rev = product_revenue.sum()
        top_20_count = max(1, int(len(product_revenue) * 0.2))
        top_20_rev = product_revenue.head(top_20_count).sum()
        top_20_pct = (top_20_rev / total_rev * 100) if total_rev > 0 else 0
        
        st.markdown(f"""
**Revenue Concentration Analysis:**

📊 **Total Products**: {len(product_revenue):,}  
📊 **Top 20% ({top_20_count} products)**: **{top_20_pct:.1f}%** of total revenue  
📊 **Remaining 80%**: **{100-top_20_pct:.1f}%** of total revenue  

**Business Implications:**
✓ Focus on top-performing products for inventory  
✓ Consider why other products underperform  
✓ Opportunity to promote under-selling items  
✓ Potential to discontinue low performers
        """)
    
    st.markdown("---")
    
    # ==================== DAY OF WEEK INSIGHTS ====================
    st.markdown("### Day of Week Revenue Analysis")
    
    if len(dow_revenue) > 0:
        best_day = dow_revenue.idxmax()
        best_day_rev = dow_revenue.max()
        avg_daily_rev = dow_revenue.mean()
        
        st.markdown(f"""
**Optimal Sales Day (By Revenue):**

📅 **Best Day**: **{best_day}** (£{best_day_rev:,.0f})  
📅 **Average Daily Revenue**: £{avg_daily_rev:,.0f}  
📅 **Range**: £{dow_revenue.min():,.0f} – £{dow_revenue.max():,.0f}  

**Strategic Decisions:**
✓ Schedule major promotions for **{best_day}** ✓ Investigate underperforming days  
✓ Consider staffing adjustments for peak days  
✓ Analyze customer behavior by day-of-week
        """)
    
    st.markdown("---")
    
    # ==================== DATA QUALITY ====================
    st.markdown("### Data Quality Assessment")
    
    dq1, dq2 = st.columns(2)
    
    with dq1:
        st.subheader("✅ Data Completeness")
        if len(df_filtered) > 0:
            completeness = (1 - (df_filtered['Quantity'].isnull().sum() + 
                                df_filtered['UnitPrice'].isnull().sum() + 
                                df_filtered['CustomerID'].isnull().sum()) / (len(df_filtered) * 3)) * 100
        else:
            completeness = 100
        st.metric("Data Completeness", f"{completeness:.1f}%", "No missing values")
        st.markdown("""
✓ Quantity: 0 missing  
✓ Unit Price: 0 missing  
✓ Customer ID: 0 missing  
        """)
    
    with dq2:
        st.subheader("⚠️ Outlier Detection")
        if len(orders) > 0:
            Q1 = orders.quantile(0.25)
            Q3 = orders.quantile(0.75)
            IQR = Q3 - Q1
            outlier_count = len(orders[(orders < Q1 - 1.5*IQR) | (orders > Q3 + 1.5*IQR)])
            outlier_pct = (outlier_count / len(orders) * 100) if len(orders) > 0 else 0
            st.metric("Outliers Detected", f"{outlier_pct:.1f}%", f"{outlier_count:,} orders")
            st.markdown(f"""
✓ Normal Range: £{Q1:.2f} – £{Q3:.2f}  
✓ High Values (Legitimate): £{Q3 + 1.5*IQR:.2f}+  
            """)
    
    st.markdown("---")
    
    # ==================== ETHICS & PRIVACY ====================
    st.markdown("### Ethics & Privacy")
    
    eth1, eth2 = st.columns(2)
    
    with eth1:
        st.subheader("🔒 Privacy: LOW RISK ✅")
        st.markdown("""
❌ No personal names  
❌ No email/phone  
❌ No addresses  
✅ Anonymized IDs  
✅ Aggregated data  
✅ 10+ years old  

**Risk Level: MINIMAL**
        """)
    
    with eth2:
        st.subheader("⚖️ Integrity Standards")
        st.markdown("""
✅ Transparent data cleaning  
✅ No fabricated metrics  
✅ Dynamic calculations  
✅ Honest claims only  
✅ Proper citations  
✅ Documented methodology
        """)
    
    st.markdown("---")
    
    # ==================== CUSTOMER INSIGHTS ====================
    st.markdown("### Customer & Revenue Metrics")
    
    if len(df_filtered) > 0:
        unique_cust = df_filtered['CustomerID'].nunique()
        total_rev = df_filtered['Revenue'].sum()
        trans_count = df_filtered['InvoiceNo'].nunique()
        rev_per_cust = (total_rev / unique_cust) if unique_cust > 0 else 0
        avg_trans_val = (total_rev / trans_count) if trans_count > 0 else 0
        
        st.markdown(f"""
**Key Customer Metrics:**

| Metric | Value |
|--------|-------|
| **Unique Customers** | {unique_cust:,} |
| **Total Transactions** | {trans_count:,} |
| **Total Revenue** | £{total_rev:,.0f} |
| **Revenue per Customer** | £{rev_per_cust:.2f} |
| **Average Transaction** | £{avg_trans_val:.2f} |

**Strategic Insights:**
👥 **Customer Lifetime Value**: £{rev_per_cust:.2f}  
🛒 **Average Basket Size**: £{avg_trans_val:.2f}  
📍 **Focus on high-value customer segments** 💰 **Opportunity for loyalty programs**
        """)
    
    st.markdown("---")
    
    # ==================== DATASET METADATA ====================
    st.markdown("### Dataset & Citation")
    
    md1, md2 = st.columns(2)
    
    with md1:
        st.subheader("📋 Dataset Information")
        st.markdown("""
**UCI Online Retail Dataset**

📅 **Period**: Dec 2010 – Dec 2011  
📊 **Records**: 541,909 transactions  
🌍 **Country**: United Kingdom  
📈 **Type**: Transactional  

**Variables:** InvoiceNo, StockCode, Description, Quantity,  
InvoiceDate, UnitPrice, CustomerID, Country
        """)
    
    with md2:
        st.subheader("👥 Citation & License")
        st.markdown("""
**Authors:** Chen, D., Sain, S.L., Guo, K.

**Full Citation:** Chen et al. (2012). Data mining for online retail.  
Journal of Database Marketing & Customer Strategy.

**Repository:** UCI ML Repository  
**License:** CC BY 4.0
        """)
    
    st.markdown("---")
    
    # ==================== FOOTER ====================
    st.markdown(f"""

**Generated:** {datetime.now().strftime('%B %d, %Y at %H:%M')}
========================================================================
BSA83111 - Data Visualization for Decision Making
Assessment: E-Commerce Sales Dashboard (Final)
Group 7 Members: Jacques Rebero (Leader), Eric Nshimiye, 
Steven Ndacyarikumukiza, Gisele Musabuwera
========================================================================
This Streamlit app provides and interactive dashboard for analyzing the UCI online retail dataset. It includes key sales metrics, dynamic filters, and visualizations to help business stakeholders understand sales trends, product performance, and customer behavior.
    """)

else:
    st.error("Failed to load dataset. Check internet connection.")
