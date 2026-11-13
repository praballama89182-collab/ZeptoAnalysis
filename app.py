"""
Zepto Ads Intelligence (Sponsored Products + Sponsored Brands)
--------------------------------------------------------------
Auto-identifies Campaign, Product, Category, and Keyword level Zepto reports upon upload
for both Sponsored Products (SP) and Sponsored Brands (SB - PCA/SBV), computes combined &
channel-level KPIs without double-counting, and generates actionable insights.
"""

import io
import pandas as pd
import streamlit as st

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Zepto Ads Intelligence",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CUSTOM CSS ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric {
        background-color: #ffffff;
        padding: 12px 16px;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    h1, h2, h3 { color: #1e293b; font-family: 'Inter', 'Helvetica Neue', sans-serif; }
    [data-testid="stDataFrame"] th, [data-testid="stDataFrame"] td {
        text-align: center !important;
    }
    .badge-sp {
        background-color: #ede9fe; color: #5b21b6; padding: 2px 7px;
        border-radius: 5px; font-weight: 600; font-size: 11px;
    }
    .badge-sb {
        background-color: #dbeafe; color: #1e40af; padding: 2px 7px;
        border-radius: 5px; font-weight: 600; font-size: 11px;
    }
    </style>
    """, unsafe_allow_html=True)


# --- FILE DETECTION & SCHEMA NORMALIZATION ---

def identify_zepto_report(df: pd.DataFrame) -> str:
    """
    Auto-detects Zepto report granularity based on exact column signatures across SP and SB:
    - Product Level: contains ProductID / ProductName
    - Keyword Level: contains KeywordName / KeywordMatchType
    - Category Level: contains Category (without ProductID)
    - Campaign Level: contains CampaignName + CampaignType / Daily_budget / Status
    """
    cols = {c.strip() for c in df.columns}
    if "ProductID" in cols or "ProductName" in cols:
        return "Product Level"
    if "KeywordName" in cols or "KeywordMatchType" in cols:
        return "Keyword Level"
    if "Category" in cols and "ProductID" not in cols:
        return "Category Level"
    if "CampaignName" in cols or "CampaignType" in cols or "Daily_budget" in cols:
        return "Campaign Level"
    return "Unknown Report"


def clean_numeric_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # Sponsored Brands files omit 'Cpc' column natively; handled dynamically via Spend/Clicks
    num_candidates = [
        "Atc", "Clicks", "Cpc", "Cpm", "Ctr", "Daily_budget",
        "Impressions", "New_to_brand_user_percentage", "Orders",
        "Other_skus", "Revenue", "Roas", "Robas", "Same_skus",
        "Spend", "Unique_considerations", "Unique_reach"
    ]
    for col in num_candidates:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Standardize Campaign Name column across Campaign-level ('CampaignName') and Entity-level ('Campaign_name')
    if "CampaignName" in df.columns and "Campaign_name" not in df.columns:
        df["Campaign_name"] = df["CampaignName"]
    elif "Campaign_name" in df.columns and "CampaignName" not in df.columns:
        df["CampaignName"] = df["Campaign_name"]

    # Parse Strategy / Format Tag (handles pipe-delimited 'SPA | ...' and underscore-delimited 'HM_..._PCA_...')
    if "CampaignName" in df.columns:
        def extract_tag(name: str) -> str:
            s = str(name).strip()
            if "|" in s:
                return s.split("|")[0].strip()
            for token in ["PCA", "SBV", "SPA", "SPCT", "SPM"]:
                if f"_{token}_" in s or s.endswith(f"_{token}") or s.startswith(f"{token}_"):
                    return token
            return "Standard"
        df["Strategy Tag"] = df["CampaignName"].apply(extract_tag)

    return df


def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Computes clean funnel & profitability metrics (including CPC for SB where Cpc column is absent)."""
    df = df.copy()
    spend = df["Spend"] if "Spend" in df.columns else 0
    rev = df["Revenue"] if "Revenue" in df.columns else 0
    orders = df["Orders"] if "Orders" in df.columns else 0
    clicks = df["Clicks"] if "Clicks" in df.columns else 0
    imps = df["Impressions"] if "Impressions" in df.columns else 0
    atc = df["Atc"] if "Atc" in df.columns else 0

    df["CTR %"] = (clicks / imps * 100).where(imps > 0, 0.0)
    df["CPC"] = (spend / clicks).where(clicks > 0, 0.0)
    df["CPM"] = (spend / imps * 1000).where(imps > 0, 0.0)
    df["ATC Rate %"] = (atc / clicks * 100).where(clicks > 0, 0.0)
    df["CVR %"] = (orders / clicks * 100).where(clicks > 0, 0.0)
    df["AOV"] = (rev / orders).where(orders > 0, 0.0)
    df["ROAS"] = (rev / spend).where(spend > 0, 0.0)
    df["ACOS %"] = (spend / rev * 100).where(rev > 0, 0.0)

    if "Same_skus" in df.columns and "Orders" in df.columns:
        df["Halo Share %"] = ((df["Orders"] - df["Same_skus"]) / df["Orders"] * 100).where(df["Orders"] > 0, 0.0)

    int_cols = ["Impressions", "Clicks", "Atc", "Orders", "Same_skus", "Other_skus", "Campaigns", "SKUs"]
    for c in int_cols:
        if c in df.columns:
            df[c] = df[c].fillna(0).round(0).astype(int)

    dec_cols = [
        "Spend", "Revenue", "ROAS", "Robas", "ACOS %", "CTR %",
        "CPC", "CPM", "ATC Rate %", "CVR %", "AOV", "Halo Share %", "Daily_budget"
    ]
    for c in dec_cols:
        if c in df.columns:
            df[c] = df[c].round(2)

    return df


def style_table(df: pd.DataFrame, row_styler=None):
    """Applies center alignment and strict max-2-decimal formatting."""
    dec_cols = [
        "Spend", "Revenue", "ROAS", "Robas", "ACOS %", "CTR %",
        "CPC", "CPM", "ATC Rate %", "CVR %", "AOV", "Halo Share %",
        "Daily_budget", "Spend Share %", "Sales Share %"
    ]
    fmt = {c: "{:.2f}" for c in dec_cols if c in df.columns}
    styler = df.style.set_properties(**{"text-align": "center"}).set_table_styles([
        {"selector": "th", "props": [("text-align", "center")]},
        {"selector": "td", "props": [("text-align", "center")]}
    ])
    if row_styler is not None:
        styler = styler.apply(row_styler, axis=1)
    if fmt:
        styler = styler.format(fmt)
    return styler


def get_authoritative_totals(channel_reports: dict) -> dict:
    """
    Prevents double-counting when multiple granularity files are uploaded for the same channel.
    Priority order for true channel totals:
    1. Campaign Level (exact account-wide total)
    2. Category Level (exact account-wide total)
    3. Product Level (exact account-wide total)
    4. Keyword Level (keyword-targeted subset only; used as fallback if no other file exists)
    """
    priority = ["Campaign Level", "Category Level", "Product Level", "Keyword Level"]
    for level in priority:
        if level in channel_reports and not channel_reports[level].empty:
            df = channel_reports[level]
            return {
                "source": level,
                "spend": float(df["Spend"].sum()),
                "revenue": float(df["Revenue"].sum()),
                "orders": int(df["Orders"].sum()),
                "clicks": int(df["Clicks"].sum()),
                "impressions": int(df["Impressions"].sum()),
                "atc": int(df["Atc"].sum()) if "Atc" in df.columns else 0,
            }
    return {"source": "None", "spend": 0.0, "revenue": 0.0, "orders": 0, "clicks": 0, "impressions": 0, "atc": 0}


def export_excel_workbook(sheets_dict: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt = wb.add_format({
            "bg_color": "#3B0764", "font_color": "#FFFFFF", "bold": True,
            "align": "center", "valign": "middle", "border": 1
        })
        cell_fmt = wb.add_format({"align": "center", "valign": "middle", "border": 1})
        num_fmt = wb.add_format({"align": "center", "valign": "middle", "border": 1, "num_format": "0.00"})

        dec_set = {
            "Spend", "Revenue", "ROAS", "Robas", "ACOS %", "CTR %",
            "CPC", "CPM", "ATC Rate %", "CVR %", "AOV", "Halo Share %",
            "Daily_budget", "Spend Share %", "Sales Share %"
        }

        for name, df in sheets_dict.items():
            if df is None or df.empty:
                continue
            sheet_title = name[:31]
            df.to_excel(writer, sheet_name=sheet_title, index=False)
            ws = writer.sheets[sheet_title]
            ws.freeze_panes(1, 1)

            for c_idx, col_name in enumerate(df.columns):
                ws.write(0, c_idx, col_name, hdr_fmt)

            for r_idx in range(len(df)):
                for c_idx, col_name in enumerate(df.columns):
                    val = df.iloc[r_idx, c_idx]
                    if pd.isna(val):
                        ws.write(r_idx + 1, c_idx, "", cell_fmt)
                    elif col_name in dec_set and isinstance(val, (int, float)):
                        ws.write_number(r_idx + 1, c_idx, round(float(val), 2), num_fmt)
                    else:
                        ws.write(r_idx + 1, c_idx, val, cell_fmt)

            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 4
                ws.set_column(i, i, min(max(max_len, 14), 38))

    return output.getvalue()


# --- FILE INGESTION HELPER ---

def parse_uploaded_files(uploaded_files, ad_format_label: str):
    """Reads multiple files, identifies their report level, and tags them with SP or SB."""
    reports = {}
    detected_log = []
    for f in uploaded_files:
        try:
            if f.name.lower().endswith(".csv"):
                df_raw = pd.read_csv(f)
            else:
                df_raw = pd.read_excel(f)
            level = identify_zepto_report(df_raw)
            df_clean = clean_numeric_cols(df_raw)
            df_clean["Ad Format"] = ad_format_label

            if level in reports:
                reports[level] = pd.concat([reports[level], df_clean], ignore_index=True)
            else:
                reports[level] = df_clean

            detected_log.append((f.name, level, len(df_clean)))
        except Exception as e:
            st.sidebar.error(f"Failed reading {f.name}: {e}")
    return reports, detected_log


# --- MAIN APPLICATION ---

def main():
    with st.sidebar:
        st.title("⚡ Zepto Ads Intelligence")
        st.caption("Sponsored Products & Sponsored Brands Intelligence")
        st.markdown("---")

        st.subheader("1. Sponsored Products (SP)")
        sp_files = st.file_uploader(
            "Upload SP Reports (Campaign, Product, Category, Keyword)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="sp_uploader"
        )

        st.subheader("2. Sponsored Brands (SB)")
        sb_files = st.file_uploader(
            "Upload SB Reports (Campaign, Product, Category, Keyword)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="sb_uploader"
        )

        st.markdown("---")
        st.subheader("3. Optimization Thresholds")
        target_roas = st.number_input("Target ROAS Benchmark", min_value=1.0, max_value=50.0, value=8.0, step=0.5)
        waste_spend_min = st.number_input("Zero-Order Waste Alert (₹)", min_value=100, max_value=20000, value=1000, step=250)

    sp_reports, sp_log = parse_uploaded_files(sp_files, "Sponsored Products") if sp_files else ({}, [])
    sb_reports, sb_log = parse_uploaded_files(sb_files, "Sponsored Brands") if sb_files else ({}, [])

    # Show auto-identified file badges in sidebar
    if sp_log or sb_log:
        with st.sidebar:
            st.markdown("---")
            st.subheader("📂 Auto-Detected Files")
            for fname, lvl, rows in sp_log:
                st.markdown(f"<span class='badge-sp'>SP</span> <b>{lvl}</b> — <small>{fname} ({rows} rows)</small>", unsafe_allow_html=True)
            for fname, lvl, rows in sb_log:
                st.markdown(f"<span class='badge-sb'>SB</span> <b>{lvl}</b> — <small>{fname} ({rows} rows)</small>", unsafe_allow_html=True)

    if not sp_reports and not sb_reports:
        st.title("⚡ Zepto Ads Intelligence")
        st.info("👈 Upload one or more **Sponsored Products (SP)** or **Sponsored Brands (SB)** Excel/CSV reports in the sidebar to begin. The app automatically detects Campaign, Product, Category, and Keyword files.")
        return

    # --- BRAND FILTER (IF MULTIPLE BRANDS EXIST) ---
    all_dfs = list(sp_reports.values()) + list(sb_reports.values())
    brands = sorted({b for d in all_dfs if "BrandName" in d.columns for b in d["BrandName"].dropna().unique()})
    if len(brands) > 1:
        selected_brands = st.sidebar.multiselect("Filter Brand", brands, default=brands)
        for k in sp_reports:
            if "BrandName" in sp_reports[k].columns:
                sp_reports[k] = sp_reports[k][sp_reports[k]["BrandName"].isin(selected_brands)]
        for k in sb_reports:
            if "BrandName" in sb_reports[k].columns:
                sb_reports[k] = sb_reports[k][sb_reports[k]["BrandName"].isin(selected_brands)]

    # --- MERGE REPORTS BY GRANULARITY LEVEL ---
    merged_levels = {}
    for lvl in ["Campaign Level", "Product Level", "Category Level", "Keyword Level"]:
        parts = []
        if lvl in sp_reports and not sp_reports[lvl].empty:
            parts.append(sp_reports[lvl])
        if lvl in sb_reports and not sb_reports[lvl].empty:
            parts.append(sb_reports[lvl])
        if parts:
            merged_levels[lvl] = pd.concat(parts, ignore_index=True)

    # --- COMPUTE DEDUPLICATED OVERVIEW TOTALS ---
    sp_tot = get_authoritative_totals(sp_reports)
    sb_tot = get_authoritative_totals(sb_reports)

    comb_spend = sp_tot["spend"] + sb_tot["spend"]
    comb_rev = sp_tot["revenue"] + sb_tot["revenue"]
    comb_orders = sp_tot["orders"] + sb_tot["orders"]
    comb_clicks = sp_tot["clicks"] + sb_tot["clicks"]
    comb_imps = sp_tot["impressions"] + sb_tot["impressions"]
    comb_roas = comb_rev / comb_spend if comb_spend > 0 else 0.0
    comb_acos = (comb_spend / comb_rev * 100) if comb_rev > 0 else 0.0
    comb_cvr = (comb_orders / comb_clicks * 100) if comb_clicks > 0 else 0.0
    comb_aov = comb_rev / comb_orders if comb_orders > 0 else 0.0

    # --- HEADER & COMPACT 2-ROW OVERVIEW ---
    st.title("⚡ Zepto Ads Intelligence")
    active_modes = []
    if sp_reports:
        active_modes.append(f"Sponsored Products ({len(sp_reports)} levels)")
    if sb_reports:
        active_modes.append(f"Sponsored Brands ({len(sb_reports)} levels)")
    st.caption(f"**Active Scope:** {' + '.join(active_modes)}  |  Deduplicated totals across Campaign, Product, Category & Keyword exports")

    # Row 1: Core Financial & Volume KPIs (Sales, Spends, Orders, Ad ROAS)
    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    r1c1.metric("Total Ad Sales", f"₹{comb_rev:,.2f}")
    r1c2.metric("Total Ad Spend", f"₹{comb_spend:,.2f}")
    r1c3.metric("Total Orders", f"{comb_orders:,}")
    r1c4.metric("Ad ROAS (Overall)", f"{comb_roas:.2f}x")

    # Row 2: Efficiency & Funnel KPIs (ACOS, CVR, AOV, Clicks / Imps)
    r2c1, r2c2, r2c3, r2c4 = st.columns(4)
    r2c1.metric("Overall ACOS %", f"{comb_acos:.2f}%")
    r2c2.metric("Conversion Rate (CVR)", f"{comb_cvr:.2f}%")
    r2c3.metric("Average Order Value", f"₹{comb_aov:,.2f}")
    r2c4.metric("Clicks / Impressions", f"{comb_clicks:,} / {comb_imps:,}")

    export_sheets = {}

    # If BOTH SP and SB are uploaded, show a clean side-by-side channel split banner
    if sp_reports and sb_reports:
        st.markdown("##### 🔀 Channel Split: Sponsored Products (SP) vs. Sponsored Brands (SB)")
        split_df = pd.DataFrame([
            {
                "Ad Format": "Sponsored Products (SP)",
                "Impressions": sp_tot["impressions"],
                "Clicks": sp_tot["clicks"],
                "CTR %": round(sp_tot["clicks"] / sp_tot["impressions"] * 100, 2) if sp_tot["impressions"] > 0 else 0.0,
                "Spend": round(sp_tot["spend"], 2),
                "Orders": sp_tot["orders"],
                "Revenue": round(sp_tot["revenue"], 2),
                "CPC": round(sp_tot["spend"] / sp_tot["clicks"], 2) if sp_tot["clicks"] > 0 else 0.0,
                "CPM": round(sp_tot["spend"] / sp_tot["impressions"] * 1000, 2) if sp_tot["impressions"] > 0 else 0.0,
                "ROAS": round(sp_tot["revenue"] / sp_tot["spend"], 2) if sp_tot["spend"] > 0 else 0.0,
                "ACOS %": round(sp_tot["spend"] / sp_tot["revenue"] * 100, 2) if sp_tot["revenue"] > 0 else 0.0,
                "Spend Share %": round(sp_tot["spend"] / comb_spend * 100, 2) if comb_spend > 0 else 0.0,
                "Sales Share %": round(sp_tot["revenue"] / comb_rev * 100, 2) if comb_rev > 0 else 0.0,
            },
            {
                "Ad Format": "Sponsored Brands (SB)",
                "Impressions": sb_tot["impressions"],
                "Clicks": sb_tot["clicks"],
                "CTR %": round(sb_tot["clicks"] / sb_tot["impressions"] * 100, 2) if sb_tot["impressions"] > 0 else 0.0,
                "Spend": round(sb_tot["spend"], 2),
                "Orders": sb_tot["orders"],
                "Revenue": round(sb_tot["revenue"], 2),
                "CPC": round(sb_tot["spend"] / sb_tot["clicks"], 2) if sb_tot["clicks"] > 0 else 0.0,
                "CPM": round(sb_tot["spend"] / sb_tot["impressions"] * 1000, 2) if sb_tot["impressions"] > 0 else 0.0,
                "ROAS": round(sb_tot["revenue"] / sb_tot["spend"], 2) if sb_tot["spend"] > 0 else 0.0,
                "ACOS %": round(sb_tot["spend"] / sb_tot["revenue"] * 100, 2) if sb_tot["revenue"] > 0 else 0.0,
                "Spend Share %": round(sb_tot["spend"] / comb_spend * 100, 2) if comb_spend > 0 else 0.0,
                "Sales Share %": round(sb_tot["revenue"] / comb_rev * 100, 2) if comb_rev > 0 else 0.0,
            }
        ])
        st.dataframe(style_table(split_df), use_container_width=True, hide_index=True)
        export_sheets["Channel_Split_SP_vs_SB"] = split_df

    st.markdown("---")

    # --- BUILD GRANULAR DATAFRAMES ---

    # 1. Campaign Level DF
    df_camp_view = pd.DataFrame()
    if "Campaign Level" in merged_levels:
        raw_c = merged_levels["Campaign Level"].copy()
        grp_cols = ["Ad Format", "CampaignName"]
        if "CampaignType" in raw_c.columns:
            grp_cols.append("CampaignType")
        if "Status" in raw_c.columns:
            grp_cols.append("Status")
        agg_dict = {
            "Impressions": "sum", "Clicks": "sum", "Atc": "sum",
            "Spend": "sum", "Orders": "sum", "Revenue": "sum",
            "Same_skus": "sum", "Other_skus": "sum"
        }
        agg_dict = {k: v for k, v in agg_dict.items() if k in raw_c.columns}
        df_camp_view = raw_c.groupby(grp_cols, as_index=False).agg(agg_dict)
        df_camp_view = compute_metrics(df_camp_view).sort_values(by="Spend", ascending=False).reset_index(drop=True)
        export_sheets["Campaign_Performance"] = df_camp_view

    # 2. Product Level DF
    df_prod_view = pd.DataFrame()
    if "Product Level" in merged_levels:
        raw_p = merged_levels["Product Level"].copy()
        grp_cols = ["Ad Format", "ProductName"]
        if "Category" in raw_p.columns:
            grp_cols.append("Category")
        agg_dict = {
            "Campaign_name": "nunique", "Impressions": "sum", "Clicks": "sum",
            "Atc": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum",
            "Same_skus": "sum", "Other_skus": "sum"
        }
        agg_dict = {k: v for k, v in agg_dict.items() if k in raw_p.columns}
        df_prod_view = raw_p.groupby(grp_cols, as_index=False).agg(agg_dict).rename(columns={"Campaign_name": "Campaigns"})
        df_prod_view = compute_metrics(df_prod_view).sort_values(by="Revenue", ascending=False).reset_index(drop=True)
        export_sheets["Product_Performance"] = df_prod_view

    # 3. Category Level DF
    df_cat_view = pd.DataFrame()
    if "Category Level" in merged_levels:
        raw_cat = merged_levels["Category Level"].copy()
        agg_dict = {
            "Campaign_name": "nunique", "Impressions": "sum", "Clicks": "sum",
            "Atc": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum",
            "Same_skus": "sum", "Other_skus": "sum"
        }
        agg_dict = {k: v for k, v in agg_dict.items() if k in raw_cat.columns}
        df_cat_view = raw_cat.groupby(["Ad Format", "Category"], as_index=False).agg(agg_dict).rename(columns={"Campaign_name": "Campaigns"})
        df_cat_view = compute_metrics(df_cat_view).sort_values(by="Spend", ascending=False).reset_index(drop=True)
        export_sheets["Category_Performance"] = df_cat_view

    # 4. Keyword Level DF
    df_kw_view = pd.DataFrame()
    if "Keyword Level" in merged_levels:
        raw_kw = merged_levels["Keyword Level"].copy()
        grp_cols = ["Ad Format", "KeywordName"]
        if "KeywordMatchType" in raw_kw.columns:
            grp_cols.append("KeywordMatchType")
        agg_dict = {
            "Campaign_name": "nunique", "Impressions": "sum", "Clicks": "sum",
            "Atc": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum",
            "Same_skus": "sum", "Other_skus": "sum"
        }
        agg_dict = {k: v for k, v in agg_dict.items() if k in raw_kw.columns}
        df_kw_view = raw_kw.groupby(grp_cols, as_index=False).agg(agg_dict).rename(columns={"Campaign_name": "Campaigns", "KeywordMatchType": "Match Type"})
        df_kw_view = compute_metrics(df_kw_view).sort_values(by="Spend", ascending=False).reset_index(drop=True)
        export_sheets["Keyword_Performance"] = df_kw_view

    # --- MULTI-TAB ANALYTICS ---
    tabs = st.tabs([
        "🎯 Campaign Level",
        "📦 Product (SKU) Level",
        "🗂️ Category Level",
        "🔑 Keyword Level",
        "💡 Actionable Insights & Waste"
    ])

    # =========================================================================
    # TAB 1: CAMPAIGN LEVEL (WITH CLICK-TO-DRILLDOWN INTO PRODUCTS & KEYWORDS)
    # =========================================================================
    with tabs[0]:
        st.subheader("🎯 Campaign-Wise Performance")
        if not df_camp_view.empty:
            c_f1, c_f2 = st.columns([1.2, 2])
            with c_f1:
                fmt_filter = st.radio("Ad Format", ["All", "Sponsored Products", "Sponsored Brands"], horizontal=True, key="camp_fmt")
            with c_f2:
                camp_search = st.text_input("Search Campaign Name", "", key="camp_search")

            disp_camp = df_camp_view.copy()
            if fmt_filter != "All":
                disp_camp = disp_camp[disp_camp["Ad Format"] == fmt_filter]
            if camp_search:
                disp_camp = disp_camp[disp_camp["CampaignName"].str.contains(camp_search, case=False, na=False)]

            # Campaign Type & Strategy Prefix Summary (AUCTION_UP_SELL vs PCA / SPA vs SPCT vs SPM vs SBV)
            raw_camp_all = merged_levels["Campaign Level"]
            grp_strat = ["Ad Format"]
            if "CampaignType" in raw_camp_all.columns:
                grp_strat.append("CampaignType")
            if "Strategy Tag" in raw_camp_all.columns:
                grp_strat.append("Strategy Tag")

            strat_df = raw_camp_all.groupby(grp_strat, as_index=False).agg({
                "CampaignName": "nunique", "Impressions": "sum", "Clicks": "sum",
                "Spend": "sum", "Orders": "sum", "Revenue": "sum"
            }).rename(columns={"CampaignName": "Campaigns"})
            strat_df = compute_metrics(strat_df).sort_values(by="Spend", ascending=False)
            with st.expander("📊 Campaign Format & Strategy Summary (AUCTION_UP_SELL / PCA / SPA / SPCT / SPM)", expanded=False):
                st.dataframe(style_table(strat_df), use_container_width=True, hide_index=True)

            st.caption("💡 Click any campaign row below to inspect its underlying Products, Categories, and Keywords.")
            camp_sel = st.dataframe(
                style_table(disp_camp),
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row",
                key="zepto_camp_select"
            )

            sel_rows = camp_sel.selection.rows if camp_sel and camp_sel.selection else []
            if sel_rows:
                chosen_camp = disp_camp.iloc[sel_rows[0]]["CampaignName"]
                chosen_fmt = disp_camp.iloc[sel_rows[0]]["Ad Format"]
                st.markdown(f"#### 🔎 Drill-Down: `{chosen_camp}` ({chosen_fmt})")

                if "Product Level" in merged_levels:
                    sub_p = merged_levels["Product Level"][
                        (merged_levels["Product Level"]["Campaign_name"] == chosen_camp) &
                        (merged_levels["Product Level"]["Ad Format"] == chosen_fmt)
                    ].copy()
                    if not sub_p.empty:
                        sub_p = compute_metrics(sub_p).sort_values(by="Spend", ascending=False)
                        keep_p = [c for c in ["ProductName", "Category", "Impressions", "Clicks", "CTR %", "Atc", "Spend", "Orders", "Revenue", "CPC", "ROAS", "CVR %"] if c in sub_p.columns]
                        with st.expander(f"📦 Products in this Campaign ({len(sub_p)})", expanded=True):
                            st.dataframe(style_table(sub_p[keep_p]), use_container_width=True, hide_index=True)

                if "Category Level" in merged_levels:
                    sub_cat = merged_levels["Category Level"][
                        (merged_levels["Category Level"]["Campaign_name"] == chosen_camp) &
                        (merged_levels["Category Level"]["Ad Format"] == chosen_fmt)
                    ].copy()
                    if not sub_cat.empty:
                        sub_cat = compute_metrics(sub_cat).sort_values(by="Spend", ascending=False)
                        keep_cat = [c for c in ["Category", "Impressions", "Clicks", "CTR %", "Atc", "Spend", "Orders", "Revenue", "CPC", "ROAS", "CVR %"] if c in sub_cat.columns]
                        with st.expander(f"🗂️ Categories in this Campaign ({len(sub_cat)})", expanded=True):
                            st.dataframe(style_table(sub_cat[keep_cat]), use_container_width=True, hide_index=True)

                if "Keyword Level" in merged_levels:
                    sub_k = merged_levels["Keyword Level"][
                        (merged_levels["Keyword Level"]["Campaign_name"] == chosen_camp) &
                        (merged_levels["Keyword Level"]["Ad Format"] == chosen_fmt)
                    ].copy()
                    if not sub_k.empty:
                        sub_k = compute_metrics(sub_k).sort_values(by="Spend", ascending=False)
                        keep_k = [c for c in ["KeywordName", "KeywordMatchType", "Impressions", "Clicks", "CTR %", "Atc", "Spend", "Orders", "Revenue", "CPC", "ROAS", "CVR %"] if c in sub_k.columns]
                        with st.expander(f"🔑 Keywords in this Campaign ({len(sub_k)})", expanded=True):
                            st.dataframe(style_table(sub_k[keep_k]), use_container_width=True, hide_index=True)
        else:
            st.info("Upload a **Campaign Level** report to view campaign-wise breakdown.")

    # =========================================================================
    # TAB 2: PRODUCT (SKU) LEVEL
    # =========================================================================
    with tabs[1]:
        st.subheader("📦 Product (SKU) Performance & Halo Effect")
        if not df_prod_view.empty:
            p_f1, p_f2 = st.columns([1.2, 2])
            with p_f1:
                p_fmt = st.radio("Ad Format", ["All", "Sponsored Products", "Sponsored Brands"], horizontal=True, key="prod_fmt")
            with p_f2:
                p_search = st.text_input("Filter by Product Name or Category", "", key="prod_search")

            disp_prod = df_prod_view.copy()
            if p_fmt != "All":
                disp_prod = disp_prod[disp_prod["Ad Format"] == p_fmt]
            if p_search:
                disp_prod = disp_prod[
                    disp_prod["ProductName"].str.contains(p_search, case=False, na=False) |
                    disp_prod.get("Category", pd.Series("", index=disp_prod.index)).str.contains(p_search, case=False, na=False)
                ]
            st.caption("ℹ️ **Halo Share %** shows the percentage of orders coming from *Other SKUs* in your catalog after clicking this product's ad.")
            st.dataframe(style_table(disp_prod), use_container_width=True, hide_index=True)
        else:
            st.info("Upload a **Product Level** report to view SKU performance and Same-SKU vs. Halo-SKU split.")

    # =========================================================================
    # TAB 3: CATEGORY LEVEL
    # =========================================================================
    with tabs[2]:
        st.subheader("🗂️ Category-Wise Performance")
        if not df_cat_view.empty:
            cat_fmt = st.radio("Ad Format", ["All", "Sponsored Products", "Sponsored Brands"], horizontal=True, key="cat_fmt")
            disp_cat = df_cat_view.copy()
            if cat_fmt != "All":
                disp_cat = disp_cat[disp_cat["Ad Format"] == cat_fmt]

            st.dataframe(style_table(disp_cat), use_container_width=True, hide_index=True)
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                st.markdown("##### Spend by Category")
                st.bar_chart(disp_cat.groupby("Category")["Spend"].sum())
            with col_c2:
                st.markdown("##### Revenue by Category")
                st.bar_chart(disp_cat.groupby("Category")["Revenue"].sum())
        else:
            st.info("Upload a **Category Level** report to view category-level spend, orders, and ROAS.")

    # =========================================================================
    # TAB 4: KEYWORD LEVEL
    # =========================================================================
    with tabs[3]:
        st.subheader("🔑 Keyword-Wise Performance")
        if not df_kw_view.empty:
            k_f1, k_f2, k_f3 = st.columns([1.3, 1.7, 1])
            with k_f1:
                kw_fmt = st.radio("Ad Format", ["All", "Sponsored Products", "Sponsored Brands"], horizontal=True, key="kw_fmt")
            with k_f2:
                kw_search = st.text_input("Search Keyword", "", key="kw_search")
            with k_f3:
                match_opts = ["All"] + sorted(df_kw_view["Match Type"].dropna().unique().tolist()) if "Match Type" in df_kw_view.columns else ["All"]
                match_sel = st.selectbox("Match Type", match_opts, key="kw_match")

            disp_kw = df_kw_view.copy()
            if kw_fmt != "All":
                disp_kw = disp_kw[disp_kw["Ad Format"] == kw_fmt]
            if kw_search:
                disp_kw = disp_kw[disp_kw["KeywordName"].str.contains(kw_search, case=False, na=False)]
            if match_sel != "All" and "Match Type" in disp_kw.columns:
                disp_kw = disp_kw[disp_kw["Match Type"] == match_sel]

            # Match Type summary
            if "Match Type" in df_kw_view.columns:
                mt_agg = merged_levels["Keyword Level"].groupby(["Ad Format", "KeywordMatchType"], as_index=False).agg({
                    "KeywordName": "nunique", "Impressions": "sum", "Clicks": "sum",
                    "Spend": "sum", "Orders": "sum", "Revenue": "sum"
                }).rename(columns={"KeywordMatchType": "Match Type", "KeywordName": "Unique Keywords"})
                mt_agg = compute_metrics(mt_agg).sort_values(by="Spend", ascending=False)
                with st.expander("📊 Match Type Efficiency Summary (EXACT vs. PHRASE vs. BROAD)", expanded=True):
                    st.dataframe(style_table(mt_agg), use_container_width=True, hide_index=True)

            st.dataframe(style_table(disp_kw), use_container_width=True, hide_index=True)
        else:
            st.info("Upload a **Keyword Level** report to view keyword & match-type performance.")

    # =========================================================================
    # TAB 5: ACTIONABLE INSIGHTS & WASTE DETECTOR
    # =========================================================================
    with tabs[4]:
        st.subheader("💡 Automated Optimization Insights")

        # 1. Zero-Order Wasted Spend Keywords
        if not df_kw_view.empty:
            waste_kw = df_kw_view[(df_kw_view["Orders"] == 0) & (df_kw_view["Spend"] >= waste_spend_min)].sort_values(by="Spend", ascending=False)
            st.markdown(f"#### 🛑 High-Spend Zero-Order Keywords (Spend ≥ ₹{waste_spend_min:,}) — `{len(waste_kw)} found`")
            if not waste_kw.empty:
                st.caption("Consider adding these as negative keywords or pausing bids to stop budget leakage.")
                st.dataframe(style_table(waste_kw), use_container_width=True, hide_index=True)
                export_sheets["Wasted_Keywords"] = waste_kw
            else:
                st.success("No keywords exceeded your zero-order spend threshold.")

        # 2. High-ROAS Scale Winners
        if not df_kw_view.empty:
            scale_kw = df_kw_view[(df_kw_view["ROAS"] >= target_roas) & (df_kw_view["Orders"] >= 3)].sort_values(by="Revenue", ascending=False)
            st.markdown(f"#### 🚀 Scale Winners (ROAS ≥ {target_roas:.1f}x & Orders ≥ 3) — `{len(scale_kw)} found`")
            if not scale_kw.empty:
                st.caption("Top converting keywords outperforming your target ROAS — prioritize Exact match & higher bid caps.")
                st.dataframe(style_table(scale_kw), use_container_width=True, hide_index=True)
                export_sheets["Scale_Winners"] = scale_kw

        # 3. High-ATC / Low-CVR Friction SKUs
        if not df_prod_view.empty and "Atc" in df_prod_view.columns:
            friction_skus = df_prod_view[(df_prod_view["Atc"] >= 15) & (df_prod_view["CVR %"] < 5.0)].sort_values(by="Atc", ascending=False)
            st.markdown(f"#### ⚠️ High Add-to-Cart but Low Conversion SKUs — `{len(friction_skus)} found`")
            if not friction_skus.empty:
                st.caption("Shoppers are adding these products to cart but dropping off before checkout — check dark-store stockouts, pricing, or delivery fee thresholds.")
                st.dataframe(style_table(friction_skus), use_container_width=True, hide_index=True)
                export_sheets["Cart_Dropoff_SKUs"] = friction_skus

    # --- MASTER EXCEL DOWNLOAD ---
    st.markdown("---")
    if export_sheets:
        excel_bytes = export_excel_workbook(export_sheets)
        st.download_button(
            label="📥 Download Complete Zepto Ads Analysis (Excel)",
            data=excel_bytes,
            file_name="Zepto_Ads_Master_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()
