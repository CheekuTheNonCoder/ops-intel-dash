"""
app.py — Enterprise Operations Intelligence Platform v4.0 (Final UI Layer)
Fixed: Autodetected months, full dynamic period calculations, advanced cohorts & aging reports.
Fixed: Converted Python set index unions to sorted lists to prevent ValueError: index cannot be a set [1].
Fixed: Added registry and f_tick/f_del objects to generate_excel_report parameters safely [D].
Fixed: Completely removed all references to Brand/Product Health and Risk Scores, Rankings, and Grades [D].
Fixed: Removed trailing 'Brand Risk Score' string interpolation to prevent KeyError [D].
Fixed: Imported raw_esc from engine_analytics to resolve NameError [D].
Fixed: Safely unpacked all cohort validation metrics (orig, final_c, val_ok) to prevent NameError [D].
Added: Brand Filter and Escalation Delta sorting options dynamically inside Month Comparison Tab [D, 1].
Fixed: Imported compute_subcat_summary and computed subcategory summary dynamically to prevent KeyError: 'subcat_final' [D].
Fixed: Re-defined handle_ai_error to safely resolve the NameError [D].
Fixed: Upgraded default Google Gemini model path to gemini-2.5-flash to prevent 404 deprecation errors [1.2.9, 1.3.3].
Added: Global Pre/Post/Combined analysis universe radio selector, terms cleanup, and validation panel restoration [D].
"""
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import json
import os

st.set_page_config(
    page_title="Ops Intelligence Platform",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
/* Reset & Base styling */
html, body, [data-testid="stAppViewContainer"] { background: #0D1117 !important; }
[data-testid="stAppViewContainer"] > .main { background: #0D1117; }
.main .block-container { padding: 1rem 2rem 2rem 2rem !important; max-width: 100% !important; }
[data-testid="stSidebar"] { background: #161B26 !important; border-right: 1px solid #21262D !important; min-width: 260px !important; }
[data-testid="stSidebar"] * { color: #8B949E !important; }
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2 { color: #E6EDF3 !important; }
.stTabs [data-baseweb="tab-list"] { background: #161B26; border-radius: 8px; padding: 4px; border: 1px solid #21262D; }
.stTabs [data-baseweb="tab"] { color: #6E7681 !important; padding: 5px 14px !important; font-size: 12px !important; font-weight: 500 !important; }
.stTabs [aria-selected="true"] { background: #21262D !important; color: #E6EDF3 !important; }
.kpi { background: #161B26; border: 1px solid #21262D; border-radius: 8px; padding: 12px 14px; margin-bottom: 6px; min-height: 80px; }
.kpi.red { border-left: 3px solid #F85149; }
.kpi.amber { border-left: 3px solid #D29922; }
.kpi.green { border-left: 3px solid #3FB950; }
.kpi.blue { border-left: 3px solid #58A6FF; }
.kpi-lbl { font-size: 10px; font-weight: 600; color: #6E7681; text-transform: uppercase; margin: 0 0 4px; }
.kpi-val { font-size: 20px; font-weight: 700; color: #E6EDF3; margin: 0; }
.kpi-sub { font-size: 10px; color: #484F58; margin: 2px 0 0; }
.brow { background: #161B26; border: 1px solid #21262D; border-radius: 6px; padding: 8px; margin-bottom: 5px; font-size: 12px; }
.shdr { font-size: 11px; font-weight: 600; color: #6E7681; text-transform: uppercase; border-bottom: 1px solid #21262D; padding-bottom: 5px; margin: 16px 0 10px; }
</style>
""", unsafe_allow_html=True)

def kpi(label, value, sub="", color="blue"):
    st.markdown(f"""<div class="kpi {color}"><p class="kpi-lbl">{label}</p><p class="kpi-val">{value}</p>{'<p class="kpi-sub">'+sub+'</p>' if sub else ''}</div>""", unsafe_allow_html=True)

def badge(level):
    css   = {"CRITICAL":"ic-c","HIGH":"ic-h","MEDIUM":"ic-m","LOW":"ic-l"}
    lbl_color = {"CRITICAL":"#F85149","HIGH":"#D29922","MEDIUM":"#58A6FF","LOW":"#3FB950"}
    cls  = css.get(level,"ic-l")
    col  = lbl_color.get(level,"#3FB950")
    return f'<span class="ic {cls}" style="color:{col}">{level}</span>'

def handle_ai_error(e):
    """Safely intercepts standard network and provider server errors to show friendly UI advice [D]."""
    err_msg = str(e)
    if "getaddrinfo failed" in err_msg or "11001" in err_msg:
        st.error("🔌 **Network Connection Error:** The system is unable to connect to the internet. Please verify that your computer has active internet access and is not blocked by a firewall, corporate proxy, or VPN [D].")
    elif "503" in err_msg or "Service Unavailable" in err_msg:
        st.error("⏳ **AI Service Temporary Downtime (503):** The AI provider's server (Google Gemini or Anthropic) is currently overloaded or undergoing temporary maintenance. Please wait 1-2 minutes and try clicking the button again [D].")
    elif "401" in err_msg or "Unauthorized" in err_msg:
        st.error("🔑 **Authentication Error (401):** The API Key provided is invalid or has expired. Please verify and enter a valid API Key in the left sidebar [D].")
    elif "404" in err_msg or "Not Found" in err_msg:
        st.error("🔍 **AI Model Resolution Error (404):** The requested model is not enabled for your account region or project. Please check if your Google/Anthropic developer account is active [D].")
    else:
        st.error(f"⚠️ **AI Processing Error:** {err_msg}")


@st.cache_resource(show_spinner=False)
def run_pipeline(del_bytes, tick_bytes, cd, ce, ct, hd, he, md, me):
    from engine_loader import process_pipeline
    return process_pipeline(del_bytes, tick_bytes)

# ── SIDEBAR CONFIGURATION ─────────────────────────────────────
with st.sidebar:
    st.markdown("## 📦 Ops Intel Platform")
    st.caption("v4.0 • Pre + Post Cohort Edition")
    st.divider()

    st.markdown("**Upload Operational Files**")
    del_file = st.file_uploader("Delivered Orders", type=["xlsx", "xls"], key="del")
    tick_file = st.file_uploader("Tickets Dump", type=["xlsx", "xls"], key="tik")
    st.divider()

    # Restoration of Threshold Configuration Side Controls
    st.markdown("**Threshold Configuration**")
    with st.expander("Configure Thresholds"):
        crit_del = st.number_input("Critical Min Deliveries", value=300, step=50)
        crit_esc = st.number_input("Critical Min Esc %", value=7.0, step=0.5)
        crit_tix = st.number_input("Critical Min Tickets", value=25, step=5)
        high_del = st.number_input("High Min Deliveries", value=200, step=50)
        high_esc = st.number_input("High Min Esc %", value=5.0, step=0.5)
        med_del  = st.number_input("Medium Min Deliveries", value=100, step=25)
        med_esc  = st.number_input("Medium Min Esc %", value=3.0, step=0.5)

    st.divider()
    ai_on = st.toggle("Enable AI Analysis", value=False)
    api_key = ""
    if ai_on:
        try:
            api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
        except:
            api_key = ""
        if not api_key:
            api_key = st.text_input("AI API Key", type="password", help="Enter Google Gemini API Key")

    if st.button("🔄 Clear Cache & Reload"):
        st.cache_resource.clear()
        st.rerun()

# ── GATEWAY CHECK ─────────────────────────────────────────────
if not del_file or not tick_file:
    st.markdown("## 📦 Operations Intelligence Platform")
    st.caption("Pre-Delivery + Post-Delivery Integrated Cohort Attribution")
    st.divider()
    c1, c2, c3 = st.columns(3)
    with c1: kpi("1. Unified Architecture", "Pre & Post Delivery Logic", "Analyzes processing order issues alongside post-delivery quality defects.", "blue")
    with c2: kpi("2. Dynamic Cohorts", "Exact Attribution Analysis", "Matches tickets with Delivered Order ID cohorts dynamically.", "blue")
    with c3: kpi("3. 24-Sheet Export", "Complete Audit Logs", "Comprehensive multi-category summaries in corporate-styled formatting.", "blue")
    st.stop()

# Run Loader pipeline
try:
    D = run_pipeline(
        del_file.read(), tick_file.read(),
        int(crit_del), float(crit_esc), int(crit_tix),
        int(high_del), float(high_esc), int(med_del), float(med_esc)
    )
except Exception as e:
    st.error(f"❌ Error: {e}")
    st.stop()

# ── TIME INTELLIGENCE GLOBAL FILTERS ──────────────────────────
del_df = D["del_df"]
tick_df = D["tick_df"]
registry = D["registry"]
redist_sum = D["redist_summary"]

# Unpack cohort metrics for Validation tab
orig = D.get("original_ticket_count", 0)
final_c = D.get("final_ticket_count", 0)
val_ok = D.get("validation_ok", False)

# Retrieve available chronological months dynamically
available_months = sorted(del_df["Delivery Month Sort"].dropna().unique())
month_options = ["All Months"] + [m.strftime("%B %Y") for m in available_months]

st.markdown("### 📊 Global Month & Period Selector")
t1, t2 = st.columns([2, 4])
with t1:
    period_type = st.selectbox("Period Filter Level", ["All Data", "Year", "Quarter", "Month", "Week"])
with t2:
    if period_type == "All Data":
        selected_period = "All Data"
        st.info("Displaying aggregated metrics for all historical cohorts.")
    elif period_type == "Year":
        selected_period = st.selectbox("Select Year", sorted(del_df["Delivery Year"].unique()))
    elif period_type == "Quarter":
        selected_period = st.selectbox("Select Quarter", sorted(del_df["Delivery Quarter"].unique()))
    elif period_type == "Month":
        selected_period = st.selectbox("Select Month", month_options)
        if selected_period == "All Months" and len(month_options) == 2:
            selected_period = month_options[1] # Auto-select single available month
    elif period_type == "Week":
        selected_period = st.selectbox("Select Week", sorted(del_df["Delivery Week"].unique(), reverse=True))

# Filter both dataframes dynamically
if period_type == "All Data" or selected_period == "All Data":
    f_del = del_df.copy()
    f_tick = tick_df.copy()
else:
    col_map = {"Year": "Delivery Year", "Quarter": "Delivery Quarter", "Month": "Delivery Month", "Week": "Delivery Week"}
    col_name = col_map[period_type]
    f_del = del_df[del_df[col_name] == selected_period].copy()
    f_tick = tick_df[tick_df[col_name] == selected_period].copy()

# ── GLOBAL ANALYSIS MODE SELECTOR (V4.0 REQUIREMENT) ──────────
st.markdown("### 🔍 Global Analysis Universe Selector")
analysis_mode = st.radio(
    "Select Active Analysis Universe",
    ["Post Delivery", "Pre Delivery", "Combined"],
    horizontal=True,
    help="Post Delivery uses Delivered orders denominator. Pre Delivery uses All Active orders. Combined merges both universes [D]."
)

# Apply dynamic universe segmentations strictly based on Selected Mode [D]
if analysis_mode == "Post Delivery":
    f_del_universe = f_del[f_del["is_delivered"] == True].copy()
    f_tick_universe = f_tick[f_tick["ticket_category"] == "POST_DELIVERY"].copy()
elif analysis_mode == "Pre Delivery":
    f_del_universe = f_del.copy() # All active orders
    f_tick_universe = f_tick[f_tick["ticket_category"] == "PRE_DELIVERY"].copy()
else:
    f_del_universe = f_del.copy()
    f_tick_universe = f_tick.copy()

# Compute Summaries dynamically on the filtered active universe
from engine_analytics import (
    compute_brand_summary, compute_product_summary,
    compute_cohort_report, compute_weekly_trends, top_kpis, raw_esc,
    compute_subcat_summary
)
brand_sum = compute_brand_summary(f_del_universe, f_tick_universe, crit_del, crit_esc, crit_tix, high_del, high_esc, med_del, med_esc)
prod_sum = compute_product_summary(f_del_universe, f_tick_universe, crit_del, crit_esc, crit_tix, high_del, high_esc, med_del, med_esc)
cohort_report = compute_cohort_report(f_del_universe, f_tick_universe)
weeks_list = sorted(f_del_universe["Delivery Week"].unique())
weekly_trends = compute_weekly_trends(f_del_universe, f_tick_universe, weeks_list)
subcat_sum = compute_subcat_summary(f_tick_universe)

# Global metrics for KPI Blocks
kpis = top_kpis(brand_sum, prod_sum, subcat_sum, f_tick_universe, f_del_universe, weeks_list)

# ── COMPUTE COHORT vs OPERATIONAL ESCALATION (Bug #2 Fix) ──
if period_type == "Month" and selected_period != "All Months":
    op_tickets = tick_df[tick_df["Ticket Month"] == selected_period]
    op_delivered = del_df[del_df["Delivery Month"] == selected_period]
    operational_esc_rate = round((len(op_tickets) / max(len(op_delivered), 1)) * 100, 2)
    cohort_tickets = tick_df[tick_df["Delivery Month"] == selected_period]
    cohort_esc_rate = round((len(cohort_tickets) / max(len(op_delivered), 1)) * 100, 2)
else:
    operational_esc_rate = kpis['overall_esc']
    cohort_esc_rate = kpis['overall_esc']

# ── COMPUTE HISTORICAL MOVEMENT METRICS (MoM) ──────────────────
comp_df_brand = pd.DataFrame()
comp_df_prod = pd.DataFrame()
has_comparison = len(available_months) >= 2

if has_comparison:
    m_names = [m.strftime("%B %Y") for m in available_months]
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        month_a = st.selectbox("Compare Month A", m_names, index=len(m_names)-2)
    with col_m2:
        month_b = st.selectbox("Compare Month B (Latest)", m_names, index=len(m_names)-1)
        
    # Slices for Month comparison
    del_a = del_df[del_df["Delivery Month"] == month_a]
    tick_a = tick_df[tick_df["Delivery Month"] == month_a]
    del_b = del_df[del_df["Delivery Month"] == month_b]
    tick_b = tick_df[tick_df["Delivery Month"] == month_b]
    
    brand_a = compute_brand_summary(del_a, tick_a, crit_del, crit_esc, crit_tix, high_del, high_esc, med_del, med_esc).set_index("brand")
    brand_b = compute_brand_summary(del_b, tick_b, crit_del, crit_esc, crit_tix, high_del, high_esc, med_del, med_esc).set_index("brand")
    
    comp_df_brand = pd.DataFrame(index=sorted(list(set(brand_a.index) | set(brand_b.index))))
    comp_df_brand["Month A Esc %"] = comp_df_brand.index.map(brand_a["esc_pct"]).fillna(0.0)
    comp_df_brand["Month B Esc %"] = comp_df_brand.index.map(brand_b["esc_pct"]).fillna(0.0)
    comp_df_brand["Esc % Difference"] = (comp_df_brand["Month B Esc %"] - comp_df_brand["Month A Esc %"]).round(2)
    comp_df_brand["Esc Movement Status"] = comp_df_brand["Esc % Difference"].apply(
        lambda x: "🚨 INCREASE" if x > 1.0 else "✅ DECREASE" if x < -1.0 else "→ STABLE"
    )
    comp_df_brand = comp_df_brand.reset_index().rename(columns={"index": "Brand"})

    prod_a = compute_product_summary(del_a, tick_a, crit_del, crit_esc, crit_tix, high_del, high_esc, med_del, med_esc).set_index("brand_product")
    prod_b = compute_product_summary(del_b, tick_b, crit_del, crit_esc, crit_tix, high_del, high_esc, med_del, med_esc).set_index("brand_product")
    
    comp_df_prod = pd.DataFrame(index=sorted(list(set(prod_a.index) | set(prod_b.index))))
    comp_df_prod["Month A Esc %"] = comp_df_prod.index.map(prod_a["esc_pct"]).fillna(0.0)
    comp_df_prod["Month B Esc %"] = comp_df_prod.index.map(prod_b["esc_pct"]).fillna(0.0)
    comp_df_prod["Esc % Difference"] = (comp_df_prod["Month B Esc %"] - comp_df_prod["Month A Esc %"]).round(2)
    comp_df_prod["Esc Movement Status"] = comp_df_prod["Esc % Difference"].apply(
        lambda x: "🚨 INCREASE" if x > 1.0 else "✅ DECREASE" if x < -1.0 else "→ STABLE"
    )
    
    # Safely split brand_product back into Brand and Product columns to maintain consistent schema layouts [1]
    comp_df_prod = comp_df_prod.reset_index().rename(columns={"index": "brand_product"})
    comp_df_prod["Brand"] = comp_df_prod["brand_product"].apply(lambda x: x.split(" | ")[0])
    comp_df_prod["Product"] = comp_df_prod["brand_product"].apply(lambda x: x.split(" | ")[1])
    comp_df_prod = comp_df_prod[["Brand", "Product", "Month A Esc %", "Month B Esc %", "Esc % Difference", "Esc Movement Status"]]

# ── EXPORT AND DOWNLOAD BUTTON ──
from engine_export import generate_excel_report
xl_data = generate_excel_report(
    kpis, brand_sum, prod_sum, subcat_sum,
    weekly_trends, D["redist_summary"], cohort_report, comp_df_brand, comp_df_prod,
    registry, f_tick_universe, f_del_universe, # Strictly pass both resolved registry and active universe dataframes here [D]
    orig_tickets=orig, final_tickets=final_c, val_ok=val_ok, period=str(selected_period)
)

st.sidebar.download_button(
    "⬇️ Export Enterprise Excel Report", data=xl_data,
    file_name=f"OpsIntel_CohortReport_{datetime.now().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.divider()

# ── NEW EXECUTIVE KPI CARDS (V4.0 REQUIREMENT WITH CLEAN LABELS) ──
st.markdown("### 📊 Active Universe Executive KPIs")
c_k1, c_k2, c_k3, c_k4, c_k5 = st.columns(5)
with c_k1: kpi("Delivered Orders", f"{kpis['total_del']:,}", "Filtered Segment base count.", "blue")
with c_k2: kpi("Tickets", f"{kpis['total_tick']:,}", "All active ticket complaints.", "red")
with c_k3: kpi("Escalation %", f"{operational_esc_rate}%", "Overall Target <3.0%", "amber" if kpis['overall_esc'] >= 3.0 else "green")
with c_k4: kpi("Defect %", f"{kpis['overall_defect']}%", "Quality Defect target <1.5%", "red" if kpis['overall_defect'] >= 1.5 else "green")
with c_k5: kpi("Peak Week", str(kpis['spike_week']), "Highest reported point.", "purple")

# Pre/Post Split cards below (Phase 4 Additions)
st.markdown("**Pre/Post Escalation Breakdowns**")
s_c1, s_c2, s_c3, s_c4, s_c5 = st.columns(5)
pre_tix_v = len(tick_df[tick_df["ticket_category"] == "PRE_DELIVERY"])
post_tix_v = len(tick_df[tick_df["ticket_category"] == "POST_DELIVERY"])
with s_c1: kpi("Pre Delivery Tickets", f"{pre_tix_v:,}", "Total Pre-delivery complaints.", "blue")
with s_c2: kpi("Post Delivery Tickets", f"{post_tix_v:,}", "Total Post-delivery quality defects.", "red")
with s_c3: kpi("Combined Tickets", f"{len(tick_df):,}", "Total ticket base sum.", "purple")
with s_c4: kpi("Pre Escalation %", f"{round((pre_tix_v / max(len(del_df), 1)) * 100, 2)}%", "Pre tickets / Total Orders", "blue")
with s_c5: kpi("Post Escalation %", f"{round((post_tix_v / max(len(del_df[del_df['is_delivered'] == True]), 1)) * 100, 2)}%", "Post tickets / Delivered Orders", "red")

st.divider()

# ── EXECUTIVE DASHBOARD SUB-ROW ──
c_left, c_right = st.columns(2)
with c_left:
    st.markdown('<p class="shdr">Top Escalation Risk Brand Profiles</p>', unsafe_allow_html=True)
    for _, row in brand_sum.head(3).iterrows():
        # Bug Fix: Completely removed "Brand Risk Score" interpolation to resolve KeyError [D]
        st.markdown(f"""<div class="brow"><b style="color:#F85149">{row['brand']}</b><span style="float:right;color:#E6EDF3"><b>{row['esc_pct']:.2f}% Esc %</b> ({int(row['tickets']):,} tickets)</span><br><small style="color:#8B949E">Primary issue: {row['Top Escalation Driver']} | Defect: {row['defect_rate']:.2f}%</small></div>""", unsafe_allow_html=True)
with c_right:
    st.markdown('<p class="shdr">Top Escalation Defect Subcategories</p>', unsafe_allow_html=True)
    if not subcat_sum.empty:
        for _, row in subcat_sum.head(3).iterrows():
            st.markdown(f"""<div class="brow"><b style="color:#58A6FF">{row['subcat_final']}</b><span style="float:right;color:#E6EDF3"><b>{row['count']:,} complaints</b> ({row['pct']:.1f}%)</span></div>""", unsafe_allow_html=True)

st.divider()

# ── TAB SYSTEM ──
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "🏷️ Brand Intel", "📦 Product Intel", "📅 Weekly Trends",
    "📊 Issue Breakdown", "📈 Month Comparison", "📋 Validation Panel", "🗺️ Redistribution Audit", "🤖 AI Insights"
])

# ── TAB 1: BRAND INTEL (RESTORED FEATURES PRESERVED) ──
with tab1:
    st.markdown('<p class="shdr">All Normalised Brand Performance Matrix</p>', unsafe_allow_html=True)
    
    b_fa, b_fb, b_fc = st.columns(3)
    with b_fa:
        b_imp_f = st.multiselect("Impact Filter", ["CRITICAL", "HIGH", "MEDIUM", "LOW"], default=["CRITICAL", "HIGH", "MEDIUM", "LOW"], key="b_imp_tab")
    with b_fb:
        b_sort_choice = st.selectbox("Sort By", [
            "Highest Tickets", "Lowest Tickets", "Highest Esc %", "Lowest Esc %",
            "Highest Delivered", "Lowest Delivered", "A → Z", "Z → A", "Highest Defect Rate", "Lowest Defect Rate"
        ], key="b_sort_tab")
    with b_fc:
        b_min_del = st.number_input("Min Delivered Orders", value=0, step=50, key="b_min_tab")
        
    disp_b = brand_sum[brand_sum["impact"].isin(b_imp_f)].copy()
    if b_min_del > 0:
        disp_b = disp_b[disp_b["delivered"] >= b_min_del]
        
    # Strictly handle all requested sorting operations
    if b_sort_choice == "Highest Tickets":
        disp_b = disp_b.sort_values("tickets", ascending=False)
    elif b_sort_choice == "Lowest Tickets":
        disp_b = disp_b.sort_values("tickets", ascending=True)
    elif b_sort_choice == "Highest Esc %":
        disp_b = disp_b.sort_values("esc_pct", ascending=False)
    elif b_sort_choice == "Lowest Esc %":
        disp_b = disp_b.sort_values("esc_pct", ascending=True)
    elif b_sort_choice == "Highest Delivered":
        disp_b = disp_b.sort_values("delivered", ascending=False)
    elif b_sort_choice == "Lowest Delivered":
        disp_b = disp_b.sort_values("delivered", ascending=True)
    elif b_sort_choice == "A → Z":
        disp_b = disp_b.sort_values("brand", ascending=True)
    elif b_sort_choice == "Z → A":
        disp_b = disp_b.sort_values("brand", ascending=False)
    elif b_sort_choice == "Highest Defect Rate":
        disp_b = disp_b.sort_values("defect_rate", ascending=False)
    elif b_sort_choice == "Lowest Defect Rate":
        disp_b = disp_b.sort_values("defect_rate", ascending=True)

    st.dataframe(disp_b[["brand", "delivered", "tickets", "esc_pct", "defect_rate", "weighted_esc", "confidence", "Top Escalation Driver"]], use_container_width=True)

    # Brand Drilldown (Phase 2 Restore)
    st.markdown('<p class="shdr">Brand Drilldown Analysis Profile</p>', unsafe_allow_html=True)
    sel_b = st.selectbox("Select Brand for Drilldown", sorted(brand_sum["brand"].unique()), key="drill_brand")
    b_row = brand_sum[brand_sum["brand"] == sel_b].iloc[0]
    
    bd1, bd2, bd3, bd4, bd5, bd6 = st.columns(6)
    with bd1: kpi("Delivered", f"{int(b_row['delivered']):,}", color="blue")
    with bd2: kpi("Tickets", f"{int(b_row['tickets']):,}", color="red")
    with bd3: kpi("Esc %", f"{b_row['esc_pct']:.2f}%", color="amber")
    with bd4: kpi("Weighted Esc %", f"{b_row['weighted_esc']:.2f}%", color="purple")
    with bd5: kpi("Confidence %", f"{int(b_row['confidence'])}%", color="green")
    with bd6: kpi("Defect %", f"{b_row['defect_rate']:.2f}%", color="red" if b_row['defect_rate'] >= 1.5 else "green")
    
    b_left, b_right = st.columns(2)
    with b_left:
        st.markdown("**Top Products**")
        bp = prod_sum[prod_sum["brand"] == sel_b].head(10)[["canonical_product", "delivered", "tickets", "esc_pct", "impact"]].copy()
        st.dataframe(bp, use_container_width=True)
    with b_right:
        st.markdown("**Issue Breakdown**")
        bi = f_tick_universe[f_tick_universe["brand"] == sel_b].groupby("subcat_final").size().reset_index(name="Tickets").sort_values("Tickets", ascending=False)
        st.dataframe(bi, use_container_width=True)
        
    st.markdown("**Weekly Trend**")
    if sel_b in weekly_trends.columns or True:
        b_del_w = f_del_universe[f_del_universe["brand"] == sel_b].groupby("Delivery Week").size().reindex(weeks_list, fill_value=0)
        b_tick_w = f_tick_universe[f_tick_universe["brand"] == sel_b].groupby("Delivery Week").size().reindex(weeks_list, fill_value=0)
        b_wdf = pd.DataFrame({"Week": weeks_list, "Delivered": b_del_w.values, "Tickets": b_tick_w.values})
        b_wdf["Esc %"] = b_wdf.apply(lambda r: raw_esc(r["Tickets"], r["Delivered"]), axis=1)
        st.dataframe(b_wdf, use_container_width=True)

# ── TAB 2: PRODUCT INTEL ──
with tab2:
    st.markdown('<p class="shdr">Product Normalised Intelligence & Lifecycle Aging</p>', unsafe_allow_html=True)
    
    p_fa, p_fb, p_fc = st.columns(3)
    with p_fa:
        p_brand_f = st.multiselect("Brand Filter", sorted(prod_sum["brand"].unique()), key="p_brand_tab")
    with p_fb:
        p_imp_f = st.multiselect("Impact Filter", ["CRITICAL", "HIGH", "MEDIUM", "LOW"], default=["CRITICAL", "HIGH", "MEDIUM", "LOW"], key="p_imp_tab")
    with p_fc:
        p_min_del = st.number_input("Min Deliveries", value=0, step=50, key="p_min_tab")
        
    disp_p = prod_sum[prod_sum["impact"].isin(p_imp_f)].copy()
    if p_brand_f:
        disp_p = disp_p[disp_p["brand"].isin(p_brand_f)]
    if p_min_del > 0:
        disp_p = disp_p[disp_p["delivered"] >= p_min_del]
        
    st.dataframe(disp_p[["brand", "canonical_product", "delivered", "tickets", "esc_pct", "Primary Ticket Source Month", "Same Month Tickets", "Previous Month Tickets", "Older Tickets", "Ticket Aging Category", "impact"]], use_container_width=True)

    # Product Drilldown (Phase 2 Restore)
    st.markdown('<p class="shdr">Product Drilldown Analysis Profile</p>', unsafe_allow_html=True)
    pd_b = st.selectbox("Select Brand for Product Drilldown", sorted(prod_sum["brand"].unique()), key="p_drill_brand")
    pd_p_opts = sorted(prod_sum[prod_sum["brand"] == pd_b]["canonical_product"].unique())
    
    if pd_p_opts:
        pd_p = st.selectbox("Select Product for Drilldown", pd_p_opts, key="p_drill_product")
        p_row = prod_sum[(prod_sum["brand"] == pd_b) & (prod_sum["canonical_product"] == pd_p)].iloc[0]
        
        pd1, pd2, pd3, pd4 = st.columns(4)
        with pd1: kpi("Delivered", f"{int(p_row['delivered']):,}", color="blue")
        with pd2: kpi("Tickets", f"{int(p_row['tickets']):,}", color="red")
        with pd3: kpi("Esc %", f"{p_row['esc_pct']:.2f}%", color="amber")
        with pd4: kpi("Confidence %", f"{int(p_row['confidence'])}%", color="green")
        
        st.markdown("**Product Issue Breakdown**")
        p_bi = f_tick_universe[(f_tick_universe["brand"] == pd_b) & (f_tick_universe["canonical_product"] == pd_p)].groupby("subcat_final").size().reset_index(name="Tickets").sort_values("Tickets", ascending=False)
        st.dataframe(p_bi, use_container_width=True)

    # ── ADVANCED NORMALIZATION DEBUG LOG ──
    st.markdown('<p class="shdr">🛠️ Normalization Engine Debug Log</p>', unsafe_allow_html=True)
    with st.expander("Show/Hide Product Matching Debug Log", expanded=False):
        if hasattr(registry, "debug_log") and registry.debug_log:
            db_df = pd.DataFrame(registry.debug_log)
            st.dataframe(db_df.style.hide(axis="index"), use_container_width=True, height=350)
        else:
            st.info("No normalization logs recorded in this session.")

# ── TAB 3: WEEKLY TRENDS ──
with tab3:
    st.markdown('<p class="shdr">Weekly WoW Escalation Trends & Alerts</p>', unsafe_allow_html=True)
    st.dataframe(weekly_trends, use_container_width=True)

# ── TAB 4: ISSUE BREAKDOWN ──
with tab4:
    st.markdown('<p class="shdr">Attribute Subcategory Distribution</p>', unsafe_allow_html=True)
    st.dataframe(subcat_sum, use_container_width=True)

# ── TAB 5: MONTH COMPARISON ──
with tab5:
    st.markdown('<p class="shdr">Attributed Delivery Cohort Metrics</p>', unsafe_allow_html=True)
    st.dataframe(cohort_report, use_container_width=True)
    
    if not has_comparison:
        st.info("⚠️ Only one historical month detected in the datasets. Comparisons require at least 2 active months.")
    else:
        # Dynamic comparison Brand/Product filters and Sorting option [1]
        st.markdown('<p class="shdr">Month-over-Month Comparison Filters</p>', unsafe_allow_html=True)
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            comp_brand_filter = st.multiselect(
                "Filter Comparison Tables by Brand", 
                sorted(list(set(comp_df_brand["Brand"].unique()) | set(comp_df_prod["Brand"].unique()))),
                key="comp_brand_filter"
            )
        with col_c2:
            comp_sort_choice = st.selectbox(
                "Sort Comparison Tables",
                ["Brand / Product (A → Z)", "Brand / Product (Z → A)", "Highest Difference → Lowest", "Lowest Difference → Highest"],
                key="comp_sort_choice"
            )
            
        disp_comp_brand = comp_df_brand.copy()
        disp_comp_prod = comp_df_prod.copy()
        
        # Apply Multi-select Brand filters dynamically [1]
        if comp_brand_filter:
            disp_comp_brand = disp_comp_brand[disp_comp_brand["Brand"].isin(comp_brand_filter)]
            disp_comp_prod = disp_comp_prod[disp_comp_prod["Brand"].isin(comp_brand_filter)]
            
        # Apply custom selected sorting constraints [1]
        if comp_sort_choice == "Brand / Product (A → Z)":
            disp_comp_brand = disp_comp_brand.sort_values("Brand", ascending=True)
            disp_comp_prod = disp_comp_prod.sort_values(["Brand", "Product"], ascending=True)
        elif comp_sort_choice == "Brand / Product (Z → A)":
            disp_comp_brand = disp_comp_brand.sort_values("Brand", ascending=False)
            disp_comp_prod = disp_comp_prod.sort_values(["Brand", "Product"], ascending=False)
        elif comp_sort_choice == "Highest Difference → Lowest":
            disp_comp_brand = disp_comp_brand.sort_values("Esc % Difference", ascending=False)
            disp_comp_prod = disp_comp_prod.sort_values("Esc % Difference", ascending=False)
        elif comp_sort_choice == "Lowest Difference → Highest":
            disp_comp_brand = disp_comp_brand.sort_values("Esc % Difference", ascending=True)
            disp_comp_prod = disp_comp_prod.sort_values("Esc % Difference", ascending=True)

        st.markdown(f'<p class="shdr">Brand Movement Comparison ({month_a} vs {month_b})</p>', unsafe_allow_html=True)
        st.dataframe(disp_comp_brand, use_container_width=True)
        st.markdown(f'<p class="shdr">Product Movement Comparison ({month_a} vs {month_b})</p>', unsafe_allow_html=True)
        st.dataframe(disp_comp_prod, use_container_width=True)

# ── TAB 6: VALIDATION PANEL (Phase 2 Restore) ──
with tab6:
    st.markdown('<p class="shdr">Dataset Integrity Validation Panel</p>', unsafe_allow_html=True)
    
    # Bug #3 Verification check
    final_tickets_count = D["final_ticket_count"]
    validation_status = "PASS ✅" if D["validation_ok"] else "FAIL ❌"
    
    # Date Quality Reporting (Bug #1 Requirement)
    st.markdown("**Date Quality Report**")
    dq1, dq2 = st.columns(2)
    with dq1: kpi("Delivered Orders Invalid Dates Coerced", f"{D['invalid_del_dates']:,}", "Null/1970 dates cleanly coerced to Unknown Date.", "blue")
    with dq2: kpi("Tickets Invalid Dates Coerced", f"{D['invalid_tick_dates']:,}", "Null/1970 dates cleanly coerced to Unknown Date.", "red")
    
    st.markdown("**Ticket Integrity Audit**")
    v1, v2, v3, v4, v5, v6, v7 = st.columns(7)
    with v1: kpi("Original Tickets", f"{orig:,}", "Uploaded tickets count", "blue")
    with v2: kpi("Final Tickets", f"{final_c:,}", "Processed tickets count", "green" if D["validation_ok"] else "red")
    with v3: kpi("Brand Unmapped", f"{D['n_unmapped_brand']:,}", "Unmapped brand count", "purple")
    with v4: kpi("Need Details", f"{D['n_need_details']:,}", "Placeholder count", "purple")
    with v5: kpi("Not Found", f"{D['n_not_found_subcat']:,}", "Placeholder count", "purple")
    with v6: kpi("Integrity Check", str(validation_status), f"Original: {orig} == Final: {final_c}", "green" if D["validation_ok"] else "red")
    with v7: kpi("Ticket Mismatch", f"{abs(orig - final_c)}", "Should be exactly 0.", "green" if orig == final_c else "red")
    
    # Ticket Category Raw vs Normalized Audit (Bug #1 Diagnostic Panel) [D]
    st.markdown("**Ticket Category Normalization Audit**")
    ac1, ac2 = st.columns(2)
    with ac1:
        st.markdown("**Raw Ingested Categories**")
        st.write(D["raw_cat_counts"])
    with ac2:
        st.markdown("**Normalized Categories Used**")
        st.write(D["norm_cat_counts"])

# ── TAB 7: REDISTRIBUTION AUDIT (Phase 2 Restore) ──
with tab7:
    st.markdown('<p class="shdr">Ticket Apportionment Redistribution Audit</p>', unsafe_allow_html=True)
    st.dataframe(D["redist_summary"], use_container_width=True)

# ── TAB 8: AI INSIGHTS ──
with tab8:
    st.markdown('<p class="shdr">Cognitive Operational Insights & Recommendations</p>', unsafe_allow_html=True)
    if not ai_on:
        st.info("AI Analysis is disabled. Toggle 'Enable AI Analysis' in the left sidebar and provide an API Key.")
    elif not api_key:
        st.warning("Please enter a valid Google Gemini API Key in the left sidebar.")
    else:
        def call_gemini(prompt, key):
            import urllib.request
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={key.strip()}"
            body = json.dumps({"contents": [{"parts": [{"text": prompt}]}]}).encode()
            req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(req) as r:
                res_data = json.loads(r.read())
                return res_data["candidates"][0]["content"]["parts"][0]["text"]

        top10b = brand_sum.head(10)[["brand", "delivered", "tickets", "esc_pct"]].to_dict("records")
        top10p = prod_sum.head(10)[["brand", "canonical_product", "delivered", "tickets", "esc_pct"]].to_dict("records")
        top_i  = f_tick_universe.groupby("subcat_final").size().reset_index(name="count").sort_values("count", ascending=False).head(8).to_dict("records")

        ai1, ai2 = st.columns(2)
        with ai1:
            st.markdown("#### 📊 Executive Summary")
            if st.button("Generate Summary →", key="ai_exec"):
                with st.spinner("Analysing..."):
                    try:
                        out = call_gemini(f"""Senior ops analyst.
Active Analysis Universe Mode: {analysis_mode}
{kpis['total_del']:,} delivered, {kpis['total_tick']:,} tickets, {kpis['overall_esc']}% esc.
Brands: {json.dumps(top10b)} Products: {json.dumps(top10p)} Issues: {json.dumps(top_i)}
1)Executive Summary 2)Critical Brands 3)Root Causes 4)Product Spotlight 5)Top 5 Actions This Week.
Be specific with numbers. Format clearly. Include dedicated PRE vs POST delivery risk splits if Combined Mode is active.""", api_key)
                        st.markdown(f'<div class="ai-box">{out}</div>', unsafe_allow_html=True)
                    except Exception as e:
                        handle_ai_error(e)
        with ai2:
            st.markdown("#### 💬 Ask Anything")
            q = st.text_area("Ask:", placeholder="Why did BeatX escalation increase?",
                             height=90, key="ai_q")
            if st.button("Ask →", key="ai_ask"):
                if q.strip():
                    with st.spinner("Thinking..."):
                        try:
                            out = call_gemini(f"""Escalation analyst.
Active Analysis Universe Mode: {analysis_mode}
{kpis['total_del']:,} delivered, {kpis['total_tick']:,} tickets.
Brands:{json.dumps(top10b)} Products:{json.dumps(top10p)} Issues:{json.dumps(top_i)}
Q: {q}. Answer with operational cohort data and lifecycle attribution splits.""", api_key)
                            st.markdown(f'<div class="ai-box">{out}</div>', unsafe_allow_html=True)
                        except Exception as e:
                            handle_ai_error(e)

        st.divider()
        st.markdown("#### 🏷️ Brand Deep Dive")
        ai_b = st.selectbox("Brand", brand_sum["brand"].tolist(), key="ai_bd")
        if st.button("Generate Brand Report →", key="ai_bd_btn"):
            bd  = brand_sum[brand_sum["brand"]==ai_b].to_dict("records")
            bp2 = prod_sum[prod_sum["brand"]==ai_b].head(8)[
                  ["canonical_product","delivered","tickets","esc_pct"]].to_dict("records")
            bi3 = f_tick_universe[f_tick_universe["brand"]==ai_b]["subcat_final"].value_counts().head(6).to_dict()
            with st.spinner(f"Analysing {ai_b}..."):
                try:
                    out = call_gemini(f"""Brand analyst. {ai_b}: {json.dumps(bd)}
Active Analysis Universe Mode: {analysis_mode}
Products: {json.dumps(bp2)} Issues: {json.dumps(bi3)}
1)Defect Assessment 2)Problem Products 3)Issue Analysis 4)3 Recommendations 5)Priority Action This Week.""", api_key)
                    st.markdown(f'<div class="ai-box">{out}</div>', unsafe_allow_html=True)
                except Exception as e:
                    handle_ai_error(e)