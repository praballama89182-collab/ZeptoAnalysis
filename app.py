"""
Zepto Ads Intelligence (Sponsored Products + Sponsored Brands)
--------------------------------------------------------------
Auto-identifies Campaign, City, Product, Category, and Keyword level Zepto reports upon upload
for both Sponsored Products (SP) and Sponsored Brands (SB), computes combined &
channel-level KPIs without double-counting, and generates actionable insights.
"""

import io
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
        padding: 14px 18px;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    h1, h2, h3 { color: #1e293b; font-family: 'Inter', 'Helvetica Neue', sans-serif; }
    [data-testid="stDataFrame"] th, [data-testid="stDataFrame"] td {
        text-align: center !important;
        vertical-align: middle !important;
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


# --- NUMBER FORMATTING HELPERS ---

def format_smart_number(val):
    """Formats whole numbers without decimals and decimals to max 2 decimal places."""
    if pd.isna(val) or not isinstance(val, (int, float)):
        return val
    if float(val).is_integer():
        return f"{int(round(val)):,}"
    return f"{val:,.2f}"


# --- FILE DETECTION & SCHEMA NORMALIZATION ---

def identify_zepto_report(df: pd.DataFrame) -> str:
    cols = {c.strip().lower() for c in df.columns}
    if "city" in cols or "cityname" in cols:
        return "City Level"
    if "productid" in cols or "productname" in cols:
        return "Product Level"
    if "keywordname" in cols or "keywordmatchtype" in cols:
        return "Keyword Level"
    if "category" in cols and "productid" not in cols:
        return "Category Level"
    if "campaignname" in cols or "campaigntype" in cols or "daily_budget" in cols or "campaign_name" in cols:
        return "Campaign Level"
    return "Unknown Report"


def clean_numeric_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    num_candidates = [
        "Atc", "Clicks", "Cpc", "Cpm", "Ctr", "Daily_budget",
        "Impressions", "New_to_brand_user_percentage", "Orders",
        "Other_skus", "Revenue", "Roas", "Robas", "Same_skus",
        "Spend", "Unique_considerations", "Unique_reach"
    ]
    for col in num_candidates:
        if col in df.columns:
            s_col = df[col].astype(str).str.replace(",", "", regex=True)
            df[col] = pd.to_numeric(s_col, errors="coerce").fillna(0.0)

    # --- UNIFY CAMPAIGN NAME & ID FIELDS ---
    if "CampaignName" not in df.columns and "Campaign_name" in df.columns:
        df["CampaignName"] = df["Campaign_name"]
    elif "Campaign_name" not in df.columns and "CampaignName" in df.columns:
        df["Campaign_name"] = df["CampaignName"]
    elif "CampaignName" in df.columns and "Campaign_name" in df.columns:
        df["CampaignName"] = df["CampaignName"].fillna(df["Campaign_name"])
        df["Campaign_name"] = df["Campaign_name"].fillna(df["CampaignName"])

    if "CampaignID" not in df.columns and "Campaign_id" in df.columns:
        df["CampaignID"] = df["Campaign_id"]
    elif "Campaign_id" not in df.columns and "CampaignID" in df.columns:
        df["Campaign_id"] = df["CampaignID"]

    if "CityName" not in df.columns and "City" in df.columns:
        df["CityName"] = df["City"]
    elif "City" not in df.columns and "CityName" in df.columns:
        df["City"] = df["CityName"]

    return df


def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
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
        "CPC", "CPM", "ATC Rate %", "CVR %", "AOV", "Halo Share %", "Daily_budget", "Roas", "Robas"
    ]
    for c in dec_cols:
        if c in df.columns:
            df[c] = df[c].round(2)

    return df


def style_table(df: pd.DataFrame, row_styler=None):
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    fmt = {c: format_smart_number for c in num_cols}

    styler = df.style.set_properties(**{
        "text-align": "center",
        "vertical-align": "middle"
    }).set_table_styles([
        {"selector": "th", "props": [("text-align", "center"), ("vertical-align", "middle")]},
        {"selector": "td", "props": [("text-align", "center"), ("vertical-align", "middle")]}
    ])

    if len(df.columns) > 0:
        first_col = df.columns[0]
        styler = styler.set_properties(subset=[first_col], **{"text-align": "left", "vertical-align": "middle"})

    if row_styler is not None:
        styler = styler.apply(row_styler, axis=1)
    if fmt:
        styler = styler.format(fmt)
    return styler


def get_authoritative_totals(channel_reports: dict) -> dict:
    priority = ["Campaign Level", "City Level", "Category Level", "Product Level", "Keyword Level"]
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


def export_single_file_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
    return output.getvalue()


def export_excel_workbook(sheets_dict: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        zebra_fill = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
        header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        body_font = Font(name="Calibri", size=10, bold=False, color="1E293B")

        thin_side = Side(style="thin", color="CBD5E1")
        cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        first_col_align = Alignment(horizontal="left", vertical="center", wrap_text=False)
        metric_align = Alignment(horizontal="center", vertical="center", wrap_text=False)

        for name, df in sheets_dict.items():
            if df is None or df.empty:
                continue

            sheet_name = name[:31]
            df_export = df.copy()

            for col in df_export.select_dtypes(include=["float", "int"]).columns:
                df_export[col] = df_export[col].apply(
                    lambda x: int(round(x)) if pd.notna(x) and float(x).is_integer() else (round(float(x), 2) if pd.notna(x) else x)
                )

            df_export.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.row_dimensions[1].height = 26

            max_row = ws.max_row
            max_col = ws.max_column

            for col_idx in range(1, max_col + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align
                cell.border = cell_border

            for row_idx in range(2, max_row + 1):
                ws.row_dimensions[row_idx].height = 20
                is_even = (row_idx % 2 == 0)
                for col_idx in range(1, max_col + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.font = body_font
                    cell.border = cell_border
                    if is_even:
                        cell.fill = zebra_fill

                    if col_idx == 1:
                        cell.alignment = first_col_align
                    else:
                        cell.alignment = metric_align

                    val = cell.value
                    if isinstance(val, int) and not isinstance(val, bool):
                        cell.number_format = "#,##0"
                    elif isinstance(val, float):
                        if val.is_integer():
                            cell.value = int(round(val))
                            cell.number_format = "#,##0"
                        else:
                            cell.number_format = "#,##0.00"

            # Highly compact column width fitting tightly to content and header size
            for col_idx in range(1, max_col + 1):
                col_letter = get_column_letter(col_idx)
                max_len = 0
                for row_idx in range(1, max_row + 1):
                    val = ws.cell(row=row_idx, column=col_idx).value
                    if val is not None:
                        max_len = max(max_len, len(str(val)))
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 28)

    return output.getvalue()


def parse_uploaded_files(uploaded_files, ad_format_label: str):
    reports = {}
    raw_file_records = []
    detected_log = []
    for f in uploaded_files:
        try:
            if f.name.lower().endswith(".csv"):
                df_raw = pd.read_csv(f)
            else:
                df_raw = pd.read_excel(f, engine="openpyxl")
            
            level = identify_zepto_report(df_raw)
            df_cleaned_raw = clean_numeric_cols(df_raw)
            df_cleaned_raw["Ad Format"] = ad_format_label

            prefix = "SP" if ad_format_label == "Sponsored Products" else "SB"
            level_code = level.replace(" Level", "").replace(" ", "_")
            proper_filename = f"{prefix}_{level_code}_Report.xlsx"

            raw_file_records.append({
                "filename": proper_filename,
                "original_name": f.name,
                "level": level,
                "ad_format": ad_format_label,
                "df": df_cleaned_raw
            })

            if level in reports:
                reports[level] = pd.concat([reports[level], df_cleaned_raw], ignore_index=True)
            else:
                reports[level] = df_cleaned_raw

            detected_log.append((f.name, level, len(df_cleaned_raw)))
        except Exception as e:
            st.sidebar.error(f"Failed reading {f.name}: {e}")
    return reports, raw_file_records, detected_log


def main():
    with st.sidebar:
        st.title("⚡ Zepto Ads Intelligence")
        st.caption("Sponsored Products & Sponsored Brands Intelligence")
        st.markdown("---")

        st.subheader("1. Sponsored Products (SP)")
        sp_files = st.file_uploader(
            "Upload SP Reports (Campaign, City, Product, Category, Keyword)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="sp_uploader"
        )

        st.subheader("2. Sponsored Brands (SB)")
        sb_files = st.file_uploader(
            "Upload SB Reports (Campaign, City, Product, Category, Keyword)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="sb_uploader"
        )

        st.markdown("---")
        st.subheader("3. Consolidated Raw Settings")
        use_date_filter = st.checkbox("Add Duration Date / Range (2026)", value=False)
        duration_val = ""
        if use_date_filter:
            date_selection = st.date_input(
                "Select Date or Date Range (2026)",
                value=(pd.to_datetime("2026-01-01").date(), pd.to_datetime("2026-01-31").date()),
                min_value=pd.to_datetime("2026-01-01").date(),
                max_value=pd.to_datetime("2026-12-31").date(),
                key="duration_date_input"
            )
            if isinstance(date_selection, tuple) and len(date_selection) == 2:
                start_d, end_d = date_selection
                duration_val = f"{start_d} to {end_d}" if start_d != end_d else str(start_d)
            elif date_selection:
                duration_val = str(date_selection)

        st.markdown("---")
        st.subheader("4. Optimization Thresholds")
        target_roas = st.number_input("Target ROAS Benchmark", min_value=1.0, max_value=50.0, value=8.0, step=0.5)
        waste_spend_min = st.number_input("Zero-Order Waste Alert (₹)", min_value=100, max_value=20000, value=1000, step=250)

    sp_reports, sp_raw_records, sp_log = parse_uploaded_files(sp_files, "Sponsored Products") if sp_files else ({}, [], [])
    sb_reports, sb_raw_records, sb_log = parse_uploaded_files(sb_files, "Sponsored Brands") if sb_files else ({}, [], [])

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
        st.info("👈 Upload one or more **Sponsored Products (SP)** or **Sponsored Brands (SB)** Excel/CSV reports in the sidebar to begin. The app automatically detects Campaign, City, Product, Category, and Keyword files.")
        return

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

    merged_levels = {}
    for lvl in ["Campaign Level", "City Level", "Product Level", "Category Level", "Keyword Level"]:
        parts = []
        if lvl in sp_reports and not sp_reports[lvl].empty:
            parts.append(sp_reports[lvl])
        if lvl in sb_reports and not sb_reports[lvl].empty:
            parts.append(sb_reports[lvl])
        if parts:
            merged_levels[lvl] = pd.concat(parts, ignore_index=True)

    sp_tot = get_authoritative_totals(sp_reports)
    sb_tot = get_authoritative_totals(sb_reports)

    comb_spend = sp_tot["spend"] + sb_tot["spend"]
    comb_rev = sp_tot["revenue"] + sb_tot["revenue"]
    comb_orders = sp_tot["orders"] + sb_tot["orders"]
    comb_clicks = sp_tot["clicks"] + sb_tot["clicks"]
    comb_imps = sp_tot["impressions"] + sb_tot["impressions"]
    comb_roas = comb_rev / comb_spend if comb_spend > 0 else 0.0
    comb_cvr = (comb_orders / comb_clicks * 100) if comb_clicks > 0 else 0.0
    comb_aov = comb_rev / comb_orders if comb_orders > 0 else 0.0

    st.title("⚡ Zepto Ads Intelligence")
    active_modes = []
    if sp_reports:
        active_modes.append(f"Sponsored Products ({len(sp_reports)} levels)")
    if sb_reports:
        active_modes.append(f"Sponsored Brands ({len(sb_reports)} levels)")
    st.caption(f"**Active Scope:** {' + '.join(active_modes)}  |  Deduplicated totals across Campaign, City, Product, Category & Keyword exports")

    # --- TOP 6 KEY METRICS ---
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Ad Sales", f"₹{format_smart_number(comb_rev)}")
    m2.metric("Total Ad Spend", f"₹{format_smart_number(comb_spend)}")
    m3.metric("Ad ROAS", f"{format_smart_number(comb_roas)}x")

    m4, m5, m6 = st.columns(3)
    m4.metric("Total Orders", f"{comb_orders:,}")
    m5.metric("Conversion Rate (CVR)", f"{format_smart_number(comb_cvr)}%")
    m6.metric("Average Order Value (AOV)", f"₹{format_smart_number(comb_aov)}")

    export_sheets = {}

    # --- CONSOLIDATED RAW: STRICTLY EXACT REQUESTED COLUMNS FROM CAMPAIGN REPORT ---
    if "Campaign Level" in merged_levels and not merged_levels["Campaign Level"].empty:
        raw_camp_df = merged_levels["Campaign Level"].copy()
        
        # Exact requested columns matching user screenshot
        requested_raw_cols = [
            "Duration", "CampaignName", "BrandID", "BrandName", "CampaignType",
            "Atc", "Clicks", "Cpc", "Cpm", "Daily_budget", "Impressions",
            "Orders", "Other_skus", "Revenue", "Roas", "Robas", "Same_skus",
            "Spend", "Status"
        ]
        
        # Ensure Duration is populated
        if "Duration" in raw_camp_df.columns:
            raw_camp_df["Duration"] = duration_val
        else:
            raw_camp_df.insert(0, "Duration", duration_val)
            
        # Keep only the requested columns that exist in the dataframe
        existing_req_cols = [c for c in requested_raw_cols if c in raw_camp_df.columns]
        consolidated_raw_df = raw_camp_df[existing_req_cols].copy()
        export_sheets["Consolidated_Raw"] = consolidated_raw_df

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
    df_camp_view = pd.DataFrame()
    if "Campaign Level" in merged_levels:
        raw_c = merged_levels["Campaign Level"].copy()
        grp_cols = ["CampaignName", "Ad Format"]
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

    df_camp_type_view = pd.DataFrame()
    if "Campaign Level" in merged_levels and "CampaignType" in merged_levels["Campaign Level"].columns:
        raw_ct = merged_levels["Campaign Level"].copy()
        raw_ct["CampaignType"] = raw_ct["CampaignType"].fillna("Unassigned / Other")
        agg_dict = {
            "CampaignName": "nunique", "Impressions": "sum", "Clicks": "sum",
            "Atc": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum",
            "Same_skus": "sum", "Other_skus": "sum"
        }
        agg_dict = {k: v for k, v in agg_dict.items() if k in raw_ct.columns}
        df_camp_type_view = raw_ct.groupby(["CampaignType", "Ad Format"], as_index=False).agg(agg_dict).rename(columns={"CampaignName": "Campaigns"})
        df_camp_type_view = compute_metrics(df_camp_type_view).sort_values(by="Spend", ascending=False).reset_index(drop=True)
        export_sheets["Campaign_Type_Performance"] = df_camp_type_view

    df_city_view = pd.DataFrame()
    if "City Level" in merged_levels:
        raw_city = merged_levels["City Level"].copy()
        city_col = next((c for c in raw_city.columns if c.strip().lower() in ["city", "cityname"]), "City")
        agg_dict = {
            "Impressions": "sum", "Clicks": "sum",
            "Atc": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum",
            "Same_skus": "sum", "Other_skus": "sum"
        }
        agg_dict = {k: v for k, v in agg_dict.items() if k in raw_city.columns}
        df_city_view = raw_city.groupby([city_col, "Ad Format"], as_index=False).agg(agg_dict).rename(columns={city_col: "City"})
        df_city_view = compute_metrics(df_city_view).sort_values(by="Spend", ascending=False).reset_index(drop=True)
        export_sheets["City_Performance"] = df_city_view

    df_prod_view = pd.DataFrame()
    if "Product Level" in merged_levels:
        raw_p = merged_levels["Product Level"].copy()
        grp_cols = ["ProductName", "Ad Format"]
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

    df_cat_view = pd.DataFrame()
    if "Category Level" in merged_levels:
        raw_cat = merged_levels["Category Level"].copy()
        agg_dict = {
            "Campaign_name": "nunique", "Impressions": "sum", "Clicks": "sum",
            "Atc": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum",
            "Same_skus": "sum", "Other_skus": "sum"
        }
        agg_dict = {k: v for k, v in agg_dict.items() if k in raw_cat.columns}
        df_cat_view = raw_cat.groupby(["Category", "Ad Format"], as_index=False).agg(agg_dict).rename(columns={"Campaign_name": "Campaigns"})
        df_cat_view = compute_metrics(df_cat_view).sort_values(by="Spend", ascending=False).reset_index(drop=True)
        export_sheets["Category_Performance"] = df_cat_view

    df_kw_view = pd.DataFrame()
    if "Keyword Level" in merged_levels:
        raw_kw = merged_levels["Keyword Level"].copy()
        grp_cols = ["KeywordName", "Ad Format"]
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
        "🏷️ Campaign Type Level",
        "🏙️ City Level",
        "📦 Product (SKU) Level",
        "🗂️ Category Level",
        "🔑 Keyword Level",
        "💡 Actionable Insights & Waste",
        "📥 Raw Files Download"
    ])

    # TAB 1: CAMPAIGN LEVEL
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

            st.dataframe(style_table(disp_camp), use_container_width=True, hide_index=True)
        else:
            st.info("Upload a **Campaign Level** report to view campaign-wise breakdown.")

    # TAB 2: CAMPAIGN TYPE LEVEL
    with tabs[1]:
        st.subheader("🏷️ Campaign Type Performance Breakdown")
        if not df_camp_type_view.empty:
            ct_fmt = st.radio("Ad Format", ["All", "Sponsored Products", "Sponsored Brands"], horizontal=True, key="ct_fmt")
            disp_ct = df_camp_type_view.copy()
            if ct_fmt != "All":
                disp_ct = disp_ct[disp_ct["Ad Format"] == ct_fmt]

            st.dataframe(style_table(disp_ct), use_container_width=True, hide_index=True)
        else:
            st.info("Upload a **Campaign Level** report containing `CampaignType` to view campaign type aggregation.")

    # TAB 3: CITY LEVEL
    with tabs[2]:
        st.subheader("🏙️ City-Wise Performance")
        if not df_city_view.empty:
            city_fmt = st.radio("Ad Format", ["All", "Sponsored Products", "Sponsored Brands"], horizontal=True, key="city_fmt")
            disp_city = df_city_view.copy()
            if city_fmt != "All":
                disp_city = disp_city[disp_city["Ad Format"] == city_fmt]

            st.dataframe(style_table(disp_city), use_container_width=True, hide_index=True)
        else:
            st.info("Upload a **City Level** report to view geo-performance breakdown.")

    # TAB 4: PRODUCT LEVEL
    with tabs[3]:
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
            st.dataframe(style_table(disp_prod), use_container_width=True, hide_index=True)
        else:
            st.info("Upload a **Product Level** report to view SKU performance.")

    # TAB 5: CATEGORY LEVEL
    with tabs[4]:
        st.subheader("🗂️ Category-Wise Performance")
        if not df_cat_view.empty:
            cat_fmt = st.radio("Ad Format", ["All", "Sponsored Products", "Sponsored Brands"], horizontal=True, key="cat_fmt")
            disp_cat = df_cat_view.copy()
            if cat_fmt != "All":
                disp_cat = disp_cat[disp_cat["Ad Format"] == cat_fmt]

            st.dataframe(style_table(disp_cat), use_container_width=True, hide_index=True)
        else:
            st.info("Upload a **Category Level** report to view category-level spend and ROAS.")

    # TAB 6: KEYWORD LEVEL
    with tabs[5]:
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

            st.dataframe(style_table(disp_kw), use_container_width=True, hide_index=True)
        else:
            st.info("Upload a **Keyword Level** report to view keyword performance.")

    # TAB 7: ACTIONABLE INSIGHTS & WASTE DETECTOR
    with tabs[6]:
        st.subheader("💡 Automated Optimization Insights")

        if not df_kw_view.empty:
            waste_kw = df_kw_view[(df_kw_view["Orders"] == 0) & (df_kw_view["Spend"] >= waste_spend_min)].sort_values(by="Spend", ascending=False)
            st.markdown(f"#### 🛑 High-Spend Zero-Order Keywords (Spend ≥ ₹{waste_spend_min:,}) — `{len(waste_kw)} found`")
            if not waste_kw.empty:
                st.dataframe(style_table(waste_kw), use_container_width=True, hide_index=True)
                export_sheets["Wasted_Keywords"] = waste_kw
            else:
                st.success("No keywords exceeded your zero-order spend threshold.")

        if not df_kw_view.empty:
            scale_kw = df_kw_view[(df_kw_view["ROAS"] >= target_roas) & (df_kw_view["Orders"] >= 3)].sort_values(by="Revenue", ascending=False)
            st.markdown(f"#### 🚀 Scale Winners (ROAS ≥ {target_roas:.1f}x & Orders ≥ 3) — `{len(scale_kw)} found`")
            if not scale_kw.empty:
                st.dataframe(style_table(scale_kw), use_container_width=True, hide_index=True)
                export_sheets["Scale_Winners"] = scale_kw

    # TAB 8: INDIVIDUAL RAW FILES DOWNLOAD
    with tabs[7]:
        st.subheader("📥 Download Cleaned Individual Raw Reports")
        st.caption("Download each uploaded raw report separately with properly standardized filenames and numeric metrics fully cleaned.")
        
        if all_raw_records:
            col_d1, col_d2 = st.columns(2)
            for idx, rec in enumerate(all_raw_records):
                target_col = col_d1 if idx % 2 == 0 else col_d2
                with target_col:
                    file_bytes = export_single_file_excel(rec["df"])
                    badge_color = "badge-sp" if rec["ad_format"] == "Sponsored Products" else "badge-sb"
                    st.markdown(f"<span class='{badge_color}'>{rec['ad_format']}</span> <b>{rec['level']}</b>", unsafe_allow_html=True)
                    st.download_button(
                        label=f"⬇️ Download {rec['filename']}",
                        data=file_bytes,
                        file_name=rec["filename"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_raw_{idx}"
                    )
        else:
            st.info("No raw files uploaded yet.")

    # MASTER EXCEL DOWNLOAD
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
