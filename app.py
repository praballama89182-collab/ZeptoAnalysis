"""
Zepto Ads Intelligence (Sponsored Products + Sponsored Brands)
--------------------------------------------------------------
Auto-identifies Campaign, City, Product, Category, and Keyword level Zepto reports upon upload
for both Sponsored Products (SP) and Sponsored Brands (SB), computes combined &
channel-level KPIs without double-counting, and generates actionable insights.
"""

import io
import zipfile
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

# --- CUSTOM CSS: soft blue, compact, light tone ---
st.markdown("""
    <style>
    .main { background-color: #f7fafd; }
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
    h1 { color: #16324f; font-family: 'Inter', 'Helvetica Neue', sans-serif; }
    h2, h3 { color: #1e3a5f; font-family: 'Inter', 'Helvetica Neue', sans-serif; }
    [data-testid="stDataFrame"] th, [data-testid="stDataFrame"] td {
        text-align: center !important;
        vertical-align: middle !important;
    }
    .badge-sp {
        background-color: #e3edfb; color: #2c4a6b; padding: 2px 7px;
        border-radius: 5px; font-weight: 600; font-size: 11px;
    }
    .badge-sb {
        background-color: #dbeafe; color: #1e40af; padding: 2px 7px;
        border-radius: 5px; font-weight: 600; font-size: 11px;
    }
    .kpi-card {
        border-radius: 10px;
        padding: 10px 14px 9px 14px;
        border: 1px solid #c9dcf2;
        box-shadow: 0 1px 2px rgba(30,58,95,0.05);
    }
    .kpi-label {
        font-size: 11px; letter-spacing: .05em; text-transform: uppercase;
        color: #4a6a8a; font-weight: 600;
    }
    .kpi-value { font-size: 22px; font-weight: 700; color: #16324f; line-height: 1.25; }
    .kpi-sub { font-size: 11px; color: #5b7a99; }
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


# Soft blue shades for the overview KPI cards (light tone, low contrast).
KPI_SHADES = ["#eaf3fd", "#e2ecfa", "#dfe9f8", "#e8f0fb", "#e3eefa", "#eaf2fb"]


def kpi_card(label, value, sub="", shade="#eaf3fd"):
    """Compact overview metric card in a soft shade of blue."""
    st.markdown(
        f"""<div class="kpi-card" style="background:{shade};">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div></div>""",
        unsafe_allow_html=True,
    )


# --- FILE DETECTION & SCHEMA NORMALIZATION ---

def identify_zepto_report(df: pd.DataFrame) -> str:
    cols = {c.strip().lower() for c in df.columns}
    if "city" in cols or "cityname" in cols:
        return "City Level"
    if "productid" in cols or "productname" in cols or "product_name" in cols:
        return "Product Level"
    if "keywordname" in cols or "keywordmatchtype" in cols or "keyword_name" in cols:
        return "Keyword Level"
    if "category" in cols and "productid" not in cols and "productname" not in cols:
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

    # --- UNIFY & ENSURE CAMPAIGN & LOCATION FIELDS ---
    if "CampaignName" not in df.columns and "Campaign_name" in df.columns:
        df["CampaignName"] = df["Campaign_name"]
    elif "Campaign_name" not in df.columns and "CampaignName" in df.columns:
        df["Campaign_name"] = df["CampaignName"]
    elif "CampaignName" in df.columns and "Campaign_name" in df.columns:
        df["CampaignName"] = df["CampaignName"].fillna(df["Campaign_name"])
        df["Campaign_name"] = df["Campaign_name"].fillna(df["CampaignName"])
    elif "CampaignName" not in df.columns and "Campaign_name" not in df.columns:
        df["CampaignName"] = "Unknown Campaign"

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
    """Professional dashboard table: soft blue header, light zebra rows, compact."""
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    fmt = {c: format_smart_number for c in num_cols}

    styler = df.style.set_properties(**{
        "text-align": "center",
        "vertical-align": "middle",
        "font-size": "13px",
    }).set_table_styles([
        {"selector": "th", "props": [("text-align", "center"), ("vertical-align", "middle"), ("background-color", "#dbeafe"), ("color", "#1e3a5f"), ("font-size", "12px"), ("border", "1px solid #c9dcf2")]},
        {"selector": "td", "props": [("text-align", "center"), ("vertical-align", "middle"), ("border", "1px solid #e2ecf5")]},
        {"selector": "tr:nth-child(even)", "props": [("background-color", "#f2f7fd")]},
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
                "spend": float(df["Spend"].sum()) if "Spend" in df.columns else 0.0,
                "revenue": float(df["Revenue"].sum()) if "Revenue" in df.columns else 0.0,
                "orders": int(df["Orders"].sum()) if "Orders" in df.columns else 0,
                "clicks": int(df["Clicks"].sum()) if "Clicks" in df.columns else 0,
                "impressions": int(df["Impressions"].sum()) if "Impressions" in df.columns else 0,
                "atc": int(df["Atc"].sum()) if "Atc" in df.columns else 0,
            }
    return {"source": "None", "spend": 0.0, "revenue": 0.0, "orders": 0, "clicks": 0, "impressions": 0, "atc": 0}


def create_zip_archive(records: list) -> bytes:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for rec in records:
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                rec["df"].to_excel(writer, sheet_name="Sheet1", index=False)
            zip_file.writestr(rec["filename"], out.getvalue())
    return zip_buffer.getvalue()


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

            prefix = "SP" if ad_format_label == "Sponsored Products" else "SB"
            level_code = level.replace(" Level", "").replace(" ", "_")
            proper_filename = f"{prefix}_{level_code}.xlsx"

            # ZIP raw downloads stay exactly as uploaded (IDs included, no extra columns).
            raw_file_records.append({
                "filename": proper_filename,
                "original_name": f.name,
                "level": level,
                "ad_format": ad_format_label,
                "df": df_cleaned_raw
            })

            # Analysis copy only: Ad Format added, Brand ID and Campaign ID removed,
            # so every preview and the Excel export are ID-free while the ZIPs stay intact.
            df_analysis = df_cleaned_raw.copy()
            df_analysis["Ad Format"] = ad_format_label
            _id_cols = {c for c in df_analysis.columns if c.strip().lower() in ("brandid", "campaign_id", "campaignid")}
            df_analysis = df_analysis.drop(columns=list(_id_cols))

            if level in reports:
                reports[level] = pd.concat([reports[level], df_analysis], ignore_index=True)
            else:
                reports[level] = df_analysis

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

    # --- TOP 6 KEY METRICS: soft blue compact cards ---
    _acos = (comb_spend / comb_rev * 100) if comb_rev > 0 else 0.0
    kpi_defs = [
        ("Total Ad Sales", f"₹{format_smart_number(comb_rev)}", f"ROAS {format_smart_number(comb_roas)}x"),
        ("Total Ad Spend", f"₹{format_smart_number(comb_spend)}", f"ACOS {format_smart_number(_acos)}%"),
        ("Ad ROAS", f"{format_smart_number(comb_roas)}x", "Revenue ÷ Spend"),
        ("Total Orders", f"{comb_orders:,}", f"{comb_clicks:,} clicks"),
        ("Conversion Rate (CVR)", f"{format_smart_number(comb_cvr)}%", f"{comb_imps:,} impressions"),
        ("Average Order Value (AOV)", f"₹{format_smart_number(comb_aov)}", "Revenue ÷ Orders"),
    ]
    for _row in range(2):
        _cols = st.columns(3)
        for _j in range(3):
            _label, _value, _sub = kpi_defs[_row * 3 + _j]
            with _cols[_j]:
                kpi_card(_label, _value, _sub, KPI_SHADES[_row * 3 + _j])

    export_sheets = {}

    # --- CONSOLIDATED RAW: STRICTLY EXACT REQUESTED COLUMNS FROM CAMPAIGN REPORT ---
    if "Campaign Level" in merged_levels and not merged_levels["Campaign Level"].empty:
        raw_camp_df = merged_levels["Campaign Level"].copy()

        requested_raw_cols = [
            "Duration", "CampaignName", "BrandName", "CampaignType",
            "Atc", "Clicks", "Cpc", "Cpm", "Daily_budget", "Impressions",
            "Orders", "Other_skus", "Revenue", "Roas", "Robas", "Same_skus",
            "Spend", "Status"
        ]

        if "Duration" in raw_camp_df.columns:
            raw_camp_df["Duration"] = duration_val
        else:
            raw_camp_df.insert(0, "Duration", duration_val)

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

    # --- BUILD GRANULAR DATAFRAMES SAFELY ---
    df_camp_view = pd.DataFrame()
    if "Campaign Level" in merged_levels:
        raw_c = merged_levels["Campaign Level"].copy()

        grp_candidates = ["CampaignName", "Ad Format"]
        if "CampaignType" in raw_c.columns:
            grp_candidates.append("CampaignType")
        if "Status" in raw_c.columns:
            grp_candidates.append("Status")

        grp_cols = [c for c in grp_candidates if c in raw_c.columns]

        agg_dict = {
            "Impressions": "sum", "Clicks": "sum", "Atc": "sum",
            "Spend": "sum", "Orders": "sum", "Revenue": "sum",
            "Same_skus": "sum", "Other_skus": "sum"
        }
        agg_dict = {k: v for k, v in agg_dict.items() if k in raw_c.columns}

        if grp_cols:
            df_camp_view = raw_c.groupby(grp_cols, as_index=False).agg(agg_dict)
            df_camp_view = compute_metrics(df_camp_view).sort_values(by="Spend", ascending=False).reset_index(drop=True)
            export_sheets["Campaign_Performance"] = df_camp_view

    df_camp_type_view = pd.DataFrame()
    if "Campaign Level" in merged_levels:
        raw_ct = merged_levels["Campaign Level"].copy()
        if "CampaignType" not in raw_ct.columns:
            raw_ct["CampaignType"] = "Unassigned / Other"
        else:
            raw_ct["CampaignType"] = raw_ct["CampaignType"].fillna("Unassigned / Other")

        agg_dict = {
            "CampaignName": "nunique", "Impressions": "sum", "Clicks": "sum",
            "Atc": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum",
            "Same_skus": "sum", "Other_skus": "sum"
        }
        agg_dict = {k: v for k, v in agg_dict.items() if k in raw_ct.columns}

        group_keys = [k for k in ["CampaignType", "Ad Format"] if k in raw_ct.columns]
        if group_keys:
            df_camp_type_view = raw_ct.groupby(group_keys, as_index=False).agg(agg_dict)
            if "CampaignName" in df_camp_type_view.columns:
                df_camp_type_view = df_camp_type_view.rename(columns={"CampaignName": "Campaigns"})
            df_camp_type_view = compute_metrics(df_camp_type_view).sort_values(by="Spend", ascending=False).reset_index(drop=True)
            export_sheets["Campaign_Type_Performance"] = df_camp_type_view

    df_city_view = pd.DataFrame()
    if "City Level" in merged_levels:
        raw_city = merged_levels["City Level"].copy()
        city_col = next((c for c in raw_city.columns if c.strip().lower() in ["city", "cityname"]), None)
        if city_col and "Ad Format" in raw_city.columns:
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
        prod_col = next((c for c in raw_p.columns if c.strip().lower() in ["productname", "product_name"]), None)
        if prod_col:
            grp_cols = [prod_col, "Ad Format"]
            if "Category" in raw_p.columns:
                grp_cols.append("Category")
            grp_cols = [c for c in grp_cols if c in raw_p.columns]

            agg_dict = {
                "Campaign_name": "nunique" if "Campaign_name" in raw_p.columns else "CampaignName" if "CampaignName" in raw_p.columns else "Spend",
                "Impressions": "sum", "Clicks": "sum",
                "Atc": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum",
                "Same_skus": "sum", "Other_skus": "sum"
            }
            agg_dict = {k: v for k, v in agg_dict.items() if k in raw_p.columns}
            df_prod_view = raw_p.groupby(grp_cols, as_index=False).agg(agg_dict)
            if prod_col != "ProductName":
                df_prod_view = df_prod_view.rename(columns={prod_col: "ProductName"})
            if "Campaign_name" in df_prod_view.columns:
                df_prod_view = df_prod_view.rename(columns={"Campaign_name": "Campaigns"})
            df_prod_view = compute_metrics(df_prod_view).sort_values(by="Revenue", ascending=False).reset_index(drop=True)
            export_sheets["Product_Performance"] = df_prod_view

    df_cat_view = pd.DataFrame()
    if "Category Level" in merged_levels:
        raw_cat = merged_levels["Category Level"].copy()
        if "Category" in raw_cat.columns and "Ad Format" in raw_cat.columns:
            agg_dict = {
                "Impressions": "sum", "Clicks": "sum",
                "Atc": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum",
                "Same_skus": "sum", "Other_skus": "sum"
            }
            agg_dict = {k: v for k, v in agg_dict.items() if k in raw_cat.columns}
            df_cat_view = raw_cat.groupby(["Category", "Ad Format"], as_index=False).agg(agg_dict)
            df_cat_view = compute_metrics(df_cat_view).sort_values(by="Spend", ascending=False).reset_index(drop=True)
            export_sheets["Category_Performance"] = df_cat_view

    df_kw_view = pd.DataFrame()
    if "Keyword Level" in merged_levels:
        raw_kw = merged_levels["Keyword Level"].copy()
        kw_col = next((c for c in raw_kw.columns if c.strip().lower() in ["keywordname", "keyword_name"]), None)
        if kw_col:
            grp_cols = [kw_col, "Ad Format"]
            match_col = next((c for c in raw_kw.columns if c.strip().lower() in ["keywordmatchtype", "matchtype", "match_type"]), None)
            if match_col:
                grp_cols.append(match_col)
            grp_cols = [c for c in grp_cols if c in raw_kw.columns]

            agg_dict = {
                "Impressions": "sum", "Clicks": "sum",
                "Atc": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum",
                "Same_skus": "sum", "Other_skus": "sum"
            }
            agg_dict = {k: v for k, v in agg_dict.items() if k in raw_kw.columns}
            df_kw_view = raw_kw.groupby(grp_cols, as_index=False).agg(agg_dict)
            if kw_col != "KeywordName":
                df_kw_view = df_kw_view.rename(columns={kw_col: "KeywordName"})
            if match_col and match_col != "Match Type":
                df_kw_view = df_kw_view.rename(columns={match_col: "Match Type"})
            df_kw_view = compute_metrics(df_kw_view).sort_values(by="Spend", ascending=False).reset_index(drop=True)
            export_sheets["Keyword_Performance"] = df_kw_view

    # --- 7 CORE ANALYTICAL TABS ---
    tabs = st.tabs([
        "🎯 Campaign Level",
        "🏷️ Campaign Type Level",
        "🏙️ City Level",
        "📦 Product (SKU) Level",
        "🗂️ Category Level",
        "🔑 Keyword Level",
        "💡 Actionable Insights & Waste"
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
            if camp_search and "CampaignName" in disp_camp.columns:
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
                mask = disp_prod["ProductName"].str.contains(p_search, case=False, na=False)
                if "Category" in disp_prod.columns:
                    mask = mask | disp_prod["Category"].str.contains(p_search, case=False, na=False)
                disp_prod = disp_prod[mask]
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
                match_opts = ["All"]
                if "Match Type" in df_kw_view.columns:
                    match_opts += sorted(df_kw_view["Match Type"].dropna().unique().tolist())
                match_sel = st.selectbox("Match Type", match_opts, key="kw_match")

            disp_kw = df_kw_view.copy()
            if kw_fmt != "All":
                disp_kw = disp_kw[disp_kw["Ad Format"] == kw_fmt]
            if kw_search and "KeywordName" in disp_kw.columns:
                disp_kw = disp_kw[disp_kw["KeywordName"].str.contains(kw_search, case=False, na=False)]
            if match_sel != "All" and "Match Type" in disp_kw.columns:
                disp_kw = disp_kw[disp_kw["Match Type"] == match_sel]

            st.dataframe(style_table(disp_kw), use_container_width=True, hide_index=True)
        else:
            st.info("Upload a **Keyword Level** report to view keyword performance.")

    # TAB 7: ACTIONABLE INSIGHTS & WASTE DETECTOR
    with tabs[6]:
        st.subheader("💡 Automated Optimization Insights")

        if not df_kw_view.empty and "Orders" in df_kw_view.columns and "Spend" in df_kw_view.columns:
            waste_kw = df_kw_view[(df_kw_view["Orders"] == 0) & (df_kw_view["Spend"] >= waste_spend_min)].sort_values(by="Spend", ascending=False)
            st.markdown(f"#### 🛑 High-Spend Zero-Order Keywords (Spend ≥ ₹{waste_spend_min:,}) — `{len(waste_kw)} found`")
            if not waste_kw.empty:
                st.dataframe(style_table(waste_kw), use_container_width=True, hide_index=True)
                export_sheets["Wasted_Keywords"] = waste_kw
            else:
                st.success("No keywords exceeded your zero-order spend threshold.")

        if not df_kw_view.empty and "ROAS" in df_kw_view.columns and "Orders" in df_kw_view.columns:
            scale_kw = df_kw_view[(df_kw_view["ROAS"] >= target_roas) & (df_kw_view["Orders"] >= 3)].sort_values(by="Revenue", ascending=False)
            st.markdown(f"#### 🚀 Scale Winners (ROAS ≥ {target_roas:.1f}x & Orders ≥ 3) — `{len(scale_kw)} found`")
            if not scale_kw.empty:
                st.dataframe(style_table(scale_kw), use_container_width=True, hide_index=True)
                export_sheets["Scale_Winners"] = scale_kw

        # --- ADDITIONAL INSIGHTS ---
        st.markdown("---")
        st.markdown("#### 🔍 More Insights")

        # 1. EXACT vs PHRASE keyword match types
        if not df_kw_view.empty and "Match Type" in df_kw_view.columns:
            mt_df = df_kw_view.groupby("Match Type", as_index=False).agg(
                {"Impressions": "sum", "Clicks": "sum", "Spend": "sum", "Orders": "sum", "Revenue": "sum"}
            )
            mt_df = compute_metrics(mt_df).sort_values(by="Spend", ascending=False).reset_index(drop=True)
            st.markdown("##### 🔤 EXACT vs PHRASE Keywords")
            st.caption("EXACT usually converts better — compare ROAS, CTR and CVR before shifting budget between match types.")
            st.dataframe(style_table(mt_df), use_container_width=True, hide_index=True)
            export_sheets["Match_Type_Comparison"] = mt_df

        # 2. Halo champions: campaigns driving orders on other SKUs
        if not df_camp_view.empty and "Other_skus" in df_camp_view.columns and "Halo Share %" in df_camp_view.columns:
            halo_df = df_camp_view[df_camp_view["Orders"] >= 10].copy()
            halo_df = halo_df.sort_values(by="Other_skus", ascending=False).head(5)
            halo_cols = [c for c in ["CampaignName", "Orders", "Same_skus", "Other_skus", "Halo Share %", "Spend", "Revenue", "ROAS"] if c in halo_df.columns]
            st.markdown("##### 🌟 Halo Champions — campaigns driving orders on other SKUs")
            st.caption("High halo share means the campaign sells more than its own SKUs — good for cross-sell, not just direct ROAS.")
            st.dataframe(style_table(halo_df[halo_cols]), use_container_width=True, hide_index=True)
            export_sheets["Halo_Champions"] = halo_df[halo_cols]

        # 3. Fix-or-cut: meaningful spend, below target ROAS
        if not df_camp_view.empty and "ROAS" in df_camp_view.columns and "Spend" in df_camp_view.columns:
            fix_df = df_camp_view[(df_camp_view["Spend"] >= 1000) & (df_camp_view["ROAS"] < target_roas)].sort_values(by="Spend", ascending=False)
            st.markdown(f"##### 🔧 Fix or Cut — spend ≥ ₹1,000 and ROAS below {target_roas:.1f}x")
            if not fix_df.empty:
                st.caption(f"₹{format_smart_number(fix_df['Spend'].sum())} spend is below your ROAS target — review bids, keywords or creative here first.")
                fix_cols = [c for c in ["CampaignName", "Spend", "Revenue", "ROAS", "ACOS %", "Orders"] if c in fix_df.columns]
                st.dataframe(style_table(fix_df[fix_cols]), use_container_width=True, hide_index=True)
                export_sheets["Fix_or_Cut_Campaigns"] = fix_df[fix_cols]
            else:
                st.success("No campaign with spend ≥ ₹1,000 is below your ROAS target.")

        # 4. Zero-sale summary across campaigns, products and keywords (raw rows,
        # so spend is not hidden when a keyword/product also has revenue elsewhere)
        zero_rows = []
        for _lvl, _label in [("Campaign Level", "Campaigns"), ("Product Level", "Products"), ("Keyword Level", "Keywords")]:
            if _lvl in merged_levels:
                _raw = merged_levels[_lvl]
                if "Spend" in _raw.columns and "Revenue" in _raw.columns:
                    _z = _raw[(_raw["Spend"] > 0) & (_raw["Revenue"] == 0)]
                    zero_rows.append({"Level": _label, "Zero-sale rows": len(_z), "Wasted spend (₹)": round(float(_z["Spend"].sum()), 2)})
        if zero_rows:
            zero_df = pd.DataFrame(zero_rows)
            st.markdown("##### 🛑 Zero-Sale Summary — spend with no revenue")
            st.dataframe(style_table(zero_df), use_container_width=True, hide_index=True)
            export_sheets["Zero_Sale_Summary"] = zero_df

        # 5. Concentration + weakest category, as one-line takeaways
        if not df_camp_view.empty and "Spend" in df_camp_view.columns and df_camp_view["Spend"].sum() > 0:
            _top3 = df_camp_view.sort_values(by="Spend", ascending=False).head(3)
            _share = _top3["Spend"].sum() / df_camp_view["Spend"].sum() * 100
            st.markdown(f"**Concentration:** top 3 campaigns drive **{format_smart_number(_share)}%** of spend ({', '.join(_top3['CampaignName'].astype(str).tolist())}).")
        if not df_cat_view.empty and "ROAS" in df_cat_view.columns and "Spend" in df_cat_view.columns:
            _weak = df_cat_view[df_cat_view["Spend"] >= 1000].sort_values(by="ROAS").head(1)
            if not _weak.empty:
                _w = _weak.iloc[0]
                st.markdown(f"**Weakest category:** {_w['Category']} at **{format_smart_number(_w['ROAS'])}x ROAS** on ₹{format_smart_number(_w['Spend'])} spend.")

    # --- ADDITIONAL DOWNLOAD BUTTONS FOR RAW ZIP ARCHIVES (BELOW TABS) ---
    st.markdown("---")
    st.subheader("📥 Bulk Raw Report Downloads (ZIP Archives)")
    st.caption("Download all uploaded raw files grouped by ad format as ZIP archives, with metrics properly converted to numerical values and clean filenames.")

    dl_col1, dl_col2 = st.columns(2)

    with dl_col1:
        if sp_raw_records:
            sp_zip_bytes = create_zip_archive(sp_raw_records)
            st.download_button(
                label="📦 Download SP Raw Files (ZIP)",
                data=sp_zip_bytes,
                file_name="SP_Raw_Reports_Bundle.zip",
                mime="application/zip",
                key="download_sp_zip"
            )
        else:
            st.info("No Sponsored Products (SP) raw files uploaded.")

    with dl_col2:
        if sb_raw_records:
            sb_zip_bytes = create_zip_archive(sb_raw_records)
            st.download_button(
                label="📦 Download SB Raw Files (ZIP)",
                data=sb_zip_bytes,
                file_name="SB_Raw_Reports_Bundle.zip",
                mime="application/zip",
                key="download_sb_zip"
            )
        else:
            st.info("No Sponsored Brands (SB) raw files uploaded.")

    # --- MASTER EXCEL DOWNLOAD ---
    st.markdown("---")
    if export_sheets:
        excel_bytes = export_excel_workbook(export_sheets)
        st.download_button(
            label="📥 Download Complete Zepto Ads Analysis (Excel)",
            data=excel_bytes,
            file_name="Zepto_Ads_Master_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_master_excel"
        )


if __name__ == "__main__":
    main()
