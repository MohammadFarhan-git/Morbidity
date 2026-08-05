import io
import os
import re
import inspect
from datetime import datetime

# ==============================================================================
# Compatibility Patch for Starlette GZipResponder on Streamlit Cloud / ASGI
# ==============================================================================
try:
    import starlette.middleware.gzip as gz
    _orig_gzip_init = gz.GZipResponder.__init__
    _sig = inspect.signature(_orig_gzip_init)
    if 'thread_minimum_size' in _sig.parameters:
        def _patched_gzip_init(self, app, minimum_size=1024, compresslevel=9, *, thread_minimum_size=1024, **kwargs):
            return _orig_gzip_init(self, app, minimum_size=minimum_size, compresslevel=compresslevel, thread_minimum_size=thread_minimum_size, **kwargs)
        gz.GZipResponder.__init__ = _patched_gzip_init
except Exception:
    pass

import numpy as np
import openpyxl
import pandas as pd
import streamlit as st

# ==============================================================================
# Page Configuration & Styling
# ==============================================================================
st.set_page_config(
    page_title="Corneal Surgery Morbidity Analyzer",
    page_icon="👁️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1F4E78;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #555555;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8F9FA;
        border-radius: 8px;
        padding: 15px;
        border-left: 5px solid #1F4E78;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stDownloadButton button {
        background-color: #1F4E78 !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 6px !important;
        padding: 0.6rem 2rem !important;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# Configuration Constants
# ==============================================================================
CAMPUS_SHEETS = [
    'KAR Campus Data',
    'KVC Campus Data',
    'GMR Campus Data',
    'MTC Campus Data',
    'YSR Campus Data',
]

FOLLOW_UP_PERIODS = {
    "1D": (0, 3),
    "1W": (4, 10),
    "1M": (11, 45),
    "3M": (46, 135)
}

VA_SPECIAL_MAP = {
    "HM": 2.3, "HM+": 2.3,
    "CF CF": 1.9, "FFL": 1.9,
    "CF": 1.6, "CF 1M": 1.6, "CSM": 1.6,
    "CF 2M": 1.5,
    "NPL": 3.0,
    "PL": 2.7, "PL+": 2.7, "PL+ PR ACC": 2.7,
    "PL+ PR ACCURATE": 2.7, "PL+ PR INA": 2.7,
    "PL+ PR INACCURATE": 2.7,
}

REPEAT_KP_PROCS = {
    "THERAPUETIC PENETRATING KERATOPLASTY (TH PK)",
    "INTRAOCULAR ANTIBIOTIC INJECTION (IOAB),THERAPUETIC PENETRATING KERATOPLASTY (TH PK)",
    "DESCEMETS STRIPPING AUTOMATED ENDOTHELIAL KERATOPLASTY (DSAEK)",
    "ANTERIOR VITRECTOMY,THERAPUETIC PENETRATING KERATOPLASTY (TH PK)",
    "PENETRATING KERATOPLASTY (PK)",
    "DESCEMET MEMBRANE ENDOTHLIAL KERATOPLASTY (DMEK)",
    "ECCE + IOL,PENETRATING KERATOPLASTY (PK),TARSORRAPHY",
}

PREFIX_MAP = {"PN": "P", "NP": "N", "NPC": "CC", "PNC": "CC", "NPN": "N", "PNP": "P"}
AGE_BINS   = [0, 18, 40, 60, 80, np.inf]
AGE_LABELS = ["<18yr", "18-40yr", "41-60yr", "61-80yr", "80+yr"]

# ==============================================================================
# Helper Data Processing Functions
# ==============================================================================
def fix_date(df: pd.DataFrame, column: str) -> pd.DataFrame:
    if column not in df.columns:
        return df
    df = df.copy()
    parsed = pd.to_datetime(df[column], errors='coerce')
    numeric = pd.to_numeric(df[column], errors='coerce')
    valid_numeric = numeric.where(numeric.between(1, 100000))
    excel_dates = pd.Timestamp("1899-12-30") + pd.to_timedelta(valid_numeric, unit="D", errors='coerce')
    df[column] = parsed.fillna(excel_dates)
    return df

def clean_object_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.select_dtypes(include=["object", "string"]).columns:
        df[col] = df[col].apply(lambda x: x.strip().upper() if isinstance(x, str) else x)
    return df

def correct_mrno(mrno) -> str:
    if pd.isna(mrno):
        return ""
    mrno_str = str(mrno).strip()
    match = re.search(r"(?:.*-)?([A-Z]+)(\d+)$", mrno_str)
    if match:
        prefix, digits = match.groups()
        new_prefix = PREFIX_MAP.get(prefix, prefix)
        return re.sub(r"([A-Z]+)(\d+)$", new_prefix + digits, mrno_str)
    return mrno_str

def build_va_logmar_dict(df: pd.DataFrame, va_cols: list[str]) -> dict:
    unique_vals = set()
    for col in va_cols:
        if col in df.columns:
            unique_vals.update(df[col].dropna().unique())

    mapping = {}
    for va in unique_vals:
        va_str = str(va).strip().upper()
        if va_str in VA_SPECIAL_MAP:
            mapping[va_str] = VA_SPECIAL_MAP[va_str]
        elif "/" in va_str:
            try:
                cleaned = va_str.replace("P", "")
                num, den = cleaned.split("/")
                mapping[va_str] = round(-1 * np.log10(float(num) / float(den)), 4)
            except (ValueError, ZeroDivisionError):
                mapping[va_str] = np.nan
        else:
            mapping[va_str] = np.nan
    return mapping

def count_by(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col not in df.columns:
        return pd.DataFrame(columns=[col, 'Count'])
    return (
        df.groupby(col, dropna=False).size()
          .reset_index(name='Count')
          .sort_values('Count', ascending=False)
    )

def make_graft_summary(df: pd.DataFrame, group_cols: list[str] | None = None) -> pd.DataFrame:
    if group_cols is None or len(group_cols) == 0:
        summary = df.groupby('graft_health_text').size().reset_index(name='Count')
        total = len(df)
        summary['Total Performed'] = total
        summary['Percentage'] = (summary['Count'] / total * 100).round(2) if total > 0 else 0.0
        return summary

    summary = df.groupby(group_cols + ['graft_health_text']).size().reset_index(name='Count')
    totals = df.groupby(group_cols).size().reset_index(name='Total Performed')
    summary = summary.merge(totals, on=group_cols, how='left')
    summary['Percentage'] = (summary['Count'] / summary['Total Performed'] * 100).round(2)
    return summary

def combine_graft_periods(summary_dict: dict, group_cols: list[str], surgery_df: pd.DataFrame) -> pd.DataFrame:
    combined = None
    for period, sdf in summary_dict.items():
        temp = sdf[group_cols + ['graft_health_text', 'Count', 'Percentage']].copy()
        temp.rename(columns={'Count': f'{period} Count', 'Percentage': f'{period} %'}, inplace=True)
        if combined is None:
            combined = temp
        else:
            combined = combined.merge(temp, on=group_cols + ['graft_health_text'], how='outer')

    if len(group_cols) == 0:
        combined['Total Performed'] = len(surgery_df)
    else:
        totals = surgery_df.groupby(group_cols).size().reset_index(name='Total Performed')
        combined = combined.merge(totals, on=group_cols, how='left')

    for c in [c for c in combined.columns if c.endswith('Count')]:
        combined[c] = combined[c].fillna(0).astype(int)
    for c in [c for c in combined.columns if c.endswith('%')]:
        combined[c] = combined[c].fillna(0).round(2)

    ordered = group_cols + ['graft_health_text', 'Total Performed']
    for period in FOLLOW_UP_PERIODS:
        ordered.extend([f'{period} Count', f'{period} %'])
    return combined[ordered]

def adherence_summary(adherence_df: pd.DataFrame, group_col: str | None = None) -> dict:
    summaries = {}
    for period in FOLLOW_UP_PERIODS:
        if group_col is None:
            temp = (
                adherence_df[period]
                .value_counts()
                .reindex(["YES", "NO", "PENDING"], fill_value=0)
                .rename_axis("Status")
                .reset_index(name="Count")
            )
            total = temp["Count"].sum()
            temp["Percentage"] = (temp["Count"] / total * 100).round(1) if total > 0 else 0.0
        else:
            temp = (
                adherence_df
                .groupby([group_col, period]).size()
                .unstack(fill_value=0)
                .reindex(columns=["YES", "NO", "PENDING"], fill_value=0)
            )
            temp["Total"] = temp.sum(axis=1)
            for s in ["YES", "NO", "PENDING"]:
                temp[f"{s}_%"] = (temp[s] / temp["Total"] * 100).round(1)
            temp = temp.reset_index()
        summaries[period] = temp
    return summaries

def combine_adherence_periods(summary_dict: dict, adherence_df: pd.DataFrame, group_cols: list[str] | None = None) -> pd.DataFrame:
    if group_cols is None:
        group_cols = []

    combined = None
    for period, sdf in summary_dict.items():
        if len(group_cols) == 0:
            temp = sdf[['Status', 'Count', 'Percentage']].copy()
            temp.rename(columns={'Count': f'{period} Count', 'Percentage': f'{period} %'}, inplace=True)
            merge_cols = ['Status']
        else:
            temp = sdf[group_cols + ['YES', 'NO', 'PENDING', 'YES_%', 'NO_%', 'PENDING_%']].copy()
            count_long = temp.melt(
                id_vars=group_cols,
                value_vars=['YES', 'NO', 'PENDING'],
                var_name='Status', value_name=f'{period} Count'
            )
            pct_long = temp.melt(
                id_vars=group_cols,
                value_vars=['YES_%', 'NO_%', 'PENDING_%'],
                var_name='Status', value_name=f'{period} %'
            )
            pct_long['Status'] = pct_long['Status'].str.replace('_%', '', regex=False)
            temp = count_long.merge(pct_long, on=group_cols + ['Status'])
            merge_cols = group_cols + ['Status']

        combined = temp if combined is None else combined.merge(temp, on=merge_cols, how='outer')

    if len(group_cols) == 0:
        combined['Total Surgeries'] = len(adherence_df)
    else:
        totals = adherence_df.groupby(group_cols).size().reset_index(name='Total Surgeries')
        combined = combined.merge(totals, on=group_cols, how='left')

    for c in [c for c in combined.columns if c.endswith('Count')]:
        combined[c] = combined[c].fillna(0).astype(int)
    for c in [c for c in combined.columns if c.endswith('%')]:
        combined[c] = combined[c].fillna(0).round(2)

    ordered = group_cols + ['Status', 'Total Surgeries']
    for period in FOLLOW_UP_PERIODS:
        ordered.extend([f'{period} Count', f'{period} %'])
    return combined[ordered]

def categorize_repeat_surgery(x):
    if pd.isna(x):
        return "Others"
    x_str = str(x).strip()
    if x_str == "REBUBBLING":
        return "REBUBBLING"
    if x_str == "WOUND RESUTURING":
        return "WOUND_RESUTURING"
    if x_str in REPEAT_KP_PROCS:
        return "KP"
    return "Others"

def line_category(diff):
    if pd.isna(diff):
        return np.nan
    lines = int(abs(diff) // 0.1)
    if lines == 0:
        return "No Change (<1 Line)"
    elif lines <= 5:
        return f"{lines} Line{'s' if lines > 1 else ''}"
    else:
        return ">5 Lines"

def average_visits(df: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    return (
        df.groupby(group_cols + ["id"]).size()
          .reset_index(name="No_of_Visits")
          .groupby(group_cols)
          .agg(
              Total_Surgeries=("No_of_Visits", "count"),
              Total_Visits   =("No_of_Visits", "sum"),
              Avg_Visits     =("No_of_Visits", "mean"),
              Median_Visits  =("No_of_Visits", "median"),
              Min_Visits     =("No_of_Visits", "min"),
              Max_Visits     =("No_of_Visits", "max"),
              SD_Visits      =("No_of_Visits", "std"),
          )
          .round({"Avg_Visits": 2, "SD_Visits": 2})
          .reset_index()
    )

def beautify_workbook(wb: openpyxl.Workbook) -> openpyxl.Workbook:
    HEADER_FILL = openpyxl.styles.PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    HEADER_FONT = openpyxl.styles.Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    
    DATA_FONT = openpyxl.styles.Font(name="Segoe UI", size=10, bold=False, color="000000")
    BOLD_FONT = openpyxl.styles.Font(name="Segoe UI", size=10, bold=True, color="000000")
    TOTAL_ROW_FILL = openpyxl.styles.PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid")
    ALT_ROW_FILL = openpyxl.styles.PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    WHITE_ROW_FILL = openpyxl.styles.PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    THIN_BORDER_SIDE = openpyxl.styles.Side(border_style="thin", color="D9D9D9")
    MEDIUM_BOTTOM_SIDE = openpyxl.styles.Side(border_style="medium", color="1F4E78")
    DOUBLE_BOTTOM_SIDE = openpyxl.styles.Side(border_style="double", color="1F4E78")
    
    DATA_BORDER = openpyxl.styles.Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
    HEADER_BORDER = openpyxl.styles.Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=MEDIUM_BOTTOM_SIDE)
    TOTAL_BORDER = openpyxl.styles.Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=DOUBLE_BOTTOM_SIDE)
    
    TAB_COLORS = {
        "1": "2B579A",
        "2": "27AE60",
        "3": "D35400",
        "4": "C0392B",
        "5": "8E44AD",
        "6": "2980B9",
    }

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        prefix = sheetname.split("_")[0]
        if prefix in TAB_COLORS:
            ws.sheet_properties.tabColor = TAB_COLORS[prefix]
            
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = 'A2'
        
        max_row = ws.max_row
        max_col = ws.max_column
        if max_row == 0 or max_col == 0:
            continue
            
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 28
        
        headers = []
        for col in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col)
            header_val = str(cell.value) if cell.value is not None else ""
            headers.append(header_val)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = HEADER_BORDER

        for row in range(2, max_row + 1):
            ws.row_dimensions[row].height = 20
            first_cell_val = str(ws.cell(row=row, column=1).value or "").strip().lower()
            is_total_row = first_cell_val in ["overall", "total", "total primary surgeries"] or "total" in first_cell_val
            row_fill = TOTAL_ROW_FILL if is_total_row else (ALT_ROW_FILL if row % 2 == 0 else WHITE_ROW_FILL)
            row_font = BOLD_FONT if is_total_row else DATA_FONT
            row_border = TOTAL_BORDER if is_total_row else DATA_BORDER
            
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.font = row_font
                cell.fill = row_fill
                cell.border = row_border
                header_name = headers[col - 1] if col - 1 < len(headers) else ""
                val = cell.value
                header_lower = header_name.lower()
                
                if "date" in header_lower or header_lower == "dob":
                    cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
                    if val is not None:
                        cell.number_format = "yyyy-mm-dd"
                elif "%" in header_name or "pct" in header_lower or "percentage" in header_lower:
                    cell.alignment = openpyxl.styles.Alignment(horizontal="right", vertical="center")
                    if isinstance(val, (int, float)):
                        if abs(val) > 1.0 or val == 0:
                            cell.number_format = '0.0"%"'
                        else:
                            cell.number_format = '0.0%'
                elif any(k in header_name for k in ["Count", "Total", "Surgeries", "Visits", "Patients", "Lines"]) and not any(k in header_lower for k in ["avg", "sd", "median"]):
                    cell.alignment = openpyxl.styles.Alignment(horizontal="right", vertical="center")
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0'
                elif any(k in header_lower for k in ["logmar", "diff"]):
                    cell.alignment = openpyxl.styles.Alignment(horizontal="right", vertical="center")
                    if isinstance(val, (int, float)):
                        cell.number_format = '0.000'
                elif any(k in header_lower for k in ["avg", "sd", "median"]):
                    cell.alignment = openpyxl.styles.Alignment(horizontal="right", vertical="center")
                    if isinstance(val, (int, float)):
                        cell.number_format = '0.00'
                elif header_lower in ["sap_code", "gender", "status", "surg_eye", "age_category"]:
                    cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
                else:
                    if isinstance(val, (int, float)):
                        cell.alignment = openpyxl.styles.Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center")

        for col in range(1, max_col + 1):
            col_letter = openpyxl.utils.get_column_letter(col)
            max_len = 0
            for row in range(1, max_row + 1):
                val_str = str(ws.cell(row=row, column=col).value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            header_len = len(headers[col - 1]) if col - 1 < len(headers) else 0
            width = max(max_len + 4, header_len + 4, 12)
            ws.column_dimensions[col_letter].width = min(width, 48)
    return wb

# ==============================================================================
# Robust File Reader Engine
# ==============================================================================
def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    file_name = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    if file_name.endswith('.csv'):
        # For CSV, test header position
        sample = pd.read_csv(io.BytesIO(file_bytes), nrows=15, header=None)
        hdr_row = 0
        for idx, row in sample.iterrows():
            row_vals = [str(v).lower() for v in row.values]
            if any(k in row_vals for k in ['sap_code', 'tp_mrno', 'surg_date', 'graft_health_text']):
                hdr_row = idx
                break
        df = pd.read_csv(io.BytesIO(file_bytes), header=hdr_row)
        return df

    # Excel file processing (.xlsb, .xlsx, .xlsm, .xls)
    xl_engine = 'pyxlsb' if file_name.endswith('.xlsb') else None
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes), engine=xl_engine)

    sheet_names = excel_file.sheet_names
    matching_sheets = [s for s in sheet_names if s in CAMPUS_SHEETS or 'campus' in s.lower()]
    if not matching_sheets:
        matching_sheets = sheet_names

    campus_frames = []
    for sheet in matching_sheets:
        try:
            df_top = pd.read_excel(excel_file, sheet_name=sheet, header=None, nrows=15)
            hdr_row = 0
            for idx, r in df_top.iterrows():
                r_vals = [str(x).lower() for x in r.values]
                if any(k in r_vals for k in ['sap_code', 'tp_mrno', 'surg_date', 'graft_health_text']):
                    hdr_row = idx
                    break
            frame = pd.read_excel(excel_file, sheet_name=sheet, header=hdr_row)
            campus_frames.append(frame)
        except Exception as e:
            st.warning(f"Could not read sheet '{sheet}': {e}")

    if not campus_frames:
        raise ValueError("No valid data could be extracted from the uploaded Excel file.")

    df = pd.concat(campus_frames, ignore_index=True)
    return df

# ==============================================================================
# Full Analysis Pipeline Function
# ==============================================================================
def run_analysis_pipeline(df_raw: pd.DataFrame):
    df = df_raw.copy()

    # 1. Date Fixes
    DATE_COLS = ["visit_date", "surg_date", "dob", "fup_surg_date"]
    for col in DATE_COLS:
        if col in df.columns:
            df = fix_date(df, col)

    # 2. Text Column Standardisation
    df = clean_object_columns(df)

    # 3. MRNO Correction
    if "tp_mrno" in df.columns:
        df["tp_mrno_corrected"] = df["tp_mrno"].apply(correct_mrno)
    else:
        df["tp_mrno_corrected"] = df.get("id", "")

    # 4. Composite Surgery ID
    for required_col in ['surg_eye', 'surg_proc_group']:
        if required_col not in df.columns:
            df[required_col] = 'UNKNOWN'

    df['id'] = (
        df['tp_mrno_corrected'].astype(str) + '_' +
        df['surg_eye'].astype(str) + '_' +
        df['surg_proc_group'].astype(str) + '_' +
        df['surg_date'].dt.strftime('%Y-%m-%d').fillna('')
    )

    df = df.dropna(how='all').dropna(subset=['id']).sort_values(['id', 'surg_date', 'visit_date'])

    # 5. LogMAR Visual Acuity
    VA_COLS = ['ucva', 'pin_hole', 'bcva']
    va_map = build_va_logmar_dict(df, VA_COLS)
    for col in VA_COLS:
        if col in df.columns:
            df[f"{col}_logmar"] = df[col].map(va_map)

    existing_va_logmars = [f"{col}_logmar" for col in VA_COLS if f"{col}_logmar" in df.columns]
    if existing_va_logmars:
        df["best_va_logmar"] = df[existing_va_logmars].min(axis=1)
    else:
        df["best_va_logmar"] = np.nan

    # 6. Derived Columns
    if "visit_date" in df.columns and "dob" in df.columns:
        df['age'] = df["visit_date"].dt.year - df["dob"].dt.year
    else:
        df['age'] = np.nan

    if "surg_date" in df.columns and "dob" in df.columns:
        df['surg_age'] = df["surg_date"].dt.year - df["dob"].dt.year
    else:
        df['surg_age'] = np.nan

    if "visit_date" in df.columns and "surg_date" in df.columns:
        df['days_after_surgery'] = (df['visit_date'] - df['surg_date']).dt.days
    else:
        df['days_after_surgery'] = np.nan

    EXCEL_SHEETS = {}

    # 7. Visit-level & Primary Surgery Datasets
    VISIT_COLS = [
        'id', 'sap_code', 'tp_mrno_corrected', 'tp_mrno', 'dob', 'age', 'gender',
        'sub_category', 'surg_date', 'surg_eye', 'visit_date', 'days_after_surgery',
        'surg_age', 'best_va_logmar', 'surgeon_name', 'advise_surg',
        'surg_proc_group', 'graft_health_text', 'fup_surg_done', 'fup_surg_date',
        'fup_surg_doctor', 'fup_done_surg_proc',
    ]
    actual_visit_cols = [c for c in VISIT_COLS if c in df.columns]
    visit_df = df[actual_visit_cols].copy()

    for col_req in ['advise_surg', 'fup_surg_done', 'fup_done_surg_proc']:
        if col_req not in visit_df.columns:
            visit_df[col_req] = None

    repeat_mask = (
        (visit_df['visit_date'] > visit_df['surg_date']) &
        (visit_df['advise_surg'] == 'YES') &
        (visit_df['fup_surg_done'] == 'YES') &
        (visit_df['fup_done_surg_proc'].isin(REPEAT_KP_PROCS))
    )
    repeat_dates = (
        visit_df.loc[repeat_mask]
        .groupby('id', as_index=False)['fup_surg_date'].min()
        .rename(columns={'fup_surg_date': 'first_repeat_keratoplasty_date'})
    )
    visit_df = visit_df.merge(repeat_dates, on='id', how='left')

    visit_df_primary = visit_df[
        visit_df['first_repeat_keratoplasty_date'].isna() |
        (visit_df['visit_date'] < visit_df['first_repeat_keratoplasty_date'])
    ].copy()

    surgery_df = visit_df_primary.groupby('id', as_index=False).first()

    visit_df_primary["age_category"] = pd.cut(
        visit_df_primary["surg_age"], bins=AGE_BINS, labels=AGE_LABELS, right=True
    )

    # ── Analysis 1: Surgery Counts ──
    EXCEL_SHEETS["1_Total"]        = pd.DataFrame({"Metric": ["Total Primary Surgeries"], "Count": [len(surgery_df)]})
    EXCEL_SHEETS["1_By Gender"]    = count_by(surgery_df, 'gender')
    EXCEL_SHEETS["1_By Campus"]    = count_by(surgery_df, 'sap_code')
    EXCEL_SHEETS["1_By SurgType"]  = count_by(surgery_df, 'surg_proc_group')
    EXCEL_SHEETS["1_By Category"]  = count_by(surgery_df, 'sub_category')
    EXCEL_SHEETS["1_By Surgeon"]   = count_by(surgery_df, 'surgeon_name')

    # ── Analysis 2: Graft Health Trends ──
    graft_summaries = {name: {} for name in ['overall', 'campus', 'surgery', 'surgeon', 'surgeon_surgery', 'campus_surgery']}
    GROUP_SPECS = {
        'overall':         None,
        'campus':          ['sap_code'],
        'surgery':         ['surg_proc_group'],
        'surgeon':         ['surgeon_name'],
        'surgeon_surgery': ['surgeon_name', 'surg_proc_group'],
        'campus_surgery':  ['sap_code', 'surg_proc_group'],
    }
    for period, (start_day, end_day) in FOLLOW_UP_PERIODS.items():
        window_df = visit_df_primary[visit_df_primary['days_after_surgery'].between(start_day, end_day)].copy()
        if 'graft_health_text' in window_df.columns:
            window_graft = window_df.dropna(subset=['graft_health_text']).sort_values(['id', 'visit_date']).groupby('id', as_index=False).last()
        else:
            window_graft = pd.DataFrame(columns=['id', 'graft_health_text'])

        needed_cols = [c for c in ['id', 'sap_code', 'surg_proc_group', 'surgeon_name'] if c in surgery_df.columns]
        window_graft = surgery_df[needed_cols].merge(window_graft[['id', 'graft_health_text']], on='id', how='left')
        window_graft['graft_health_text'] = window_graft['graft_health_text'].fillna('NOT AVAILABLE')
        for name, cols in GROUP_SPECS.items():
            valid_cols = [c for c in (cols or []) if c in window_graft.columns]
            graft_summaries[name][period] = make_graft_summary(window_graft, valid_cols if valid_cols else None)

    EXCEL_SHEETS["2_Graft Overall"]      = combine_graft_periods(graft_summaries['overall'],        [], surgery_df)
    EXCEL_SHEETS["2_Graft by Campus"]    = combine_graft_periods(graft_summaries['campus'],         ['sap_code'], surgery_df)
    EXCEL_SHEETS["2_Graft by Surgery"]   = combine_graft_periods(graft_summaries['surgery'],        ['surg_proc_group'], surgery_df)
    EXCEL_SHEETS["2_Graft by Surgeon"]   = combine_graft_periods(graft_summaries['surgeon'],        ['surgeon_name'], surgery_df)
    EXCEL_SHEETS["2_Graft Srgn×Surg"]    = combine_graft_periods(graft_summaries['surgeon_surgery'],['surgeon_name', 'surg_proc_group'], surgery_df)
    EXCEL_SHEETS["2_Graft Camp×Surg"]    = combine_graft_periods(graft_summaries['campus_surgery'], ['sap_code', 'surg_proc_group'], surgery_df)

    # ── Analysis 3: Follow-Up Adherence ──
    today = pd.Timestamp.today().normalize()
    summary_rows = []
    for surgery_id, grp in visit_df_primary.groupby("id"):
        surg_date = grp["surg_date"].iloc[0]
        row = {
            "id": surgery_id,
            "gender": grp.get("gender", pd.Series([""])).iloc[0],
            "sub_category": grp.get("sub_category", pd.Series([""])).iloc[0],
            "surg_proc_group": grp.get("surg_proc_group", pd.Series([""])).iloc[0],
            "sap_code": grp.get("sap_code", pd.Series([""])).iloc[0],
            "age_category": grp.get("age_category", pd.Series([""])).iloc[0],
            "surg_age": grp.get("surg_age", pd.Series([np.nan])).iloc[0],
        }
        visits = grp["days_after_surgery"].dropna()
        for period, (start, end) in FOLLOW_UP_PERIODS.items():
            visit_count = visits.between(start, end).sum()
            row[f"{period}_Visits"] = int(visit_count)
            window_end = surg_date + pd.Timedelta(days=end) if pd.notna(surg_date) else pd.NaT
            if visit_count > 0:
                status = "YES"
            elif pd.notna(window_end) and today <= window_end:
                status = "PENDING"
            else:
                status = "NO"
            row[period] = status
        summary_rows.append(row)
    adherence_df = pd.DataFrame(summary_rows)

    adh_groups = {
        "Overall":      None,
        "Gender":        "gender",
        "Sub Category":  "sub_category",
        "Procedure":     "surg_proc_group",
        "SAP Code":      "sap_code",
        "Age Category":  "age_category",
    }

    for label, gcol in adh_groups.items():
        raw = adherence_summary(adherence_df, gcol if gcol in adherence_df.columns else None)
        gcols = [gcol] if gcol and gcol in adherence_df.columns else None
        combined = combine_adherence_periods(raw, adherence_df, group_cols=gcols)
        EXCEL_SHEETS[f"3_Adherence {label}"] = combined

    if "surg_proc_group" in visit_df_primary.columns:
        EXCEL_SHEETS["3_Avg Visits by Proc"] = average_visits(visit_df_primary, ["surg_proc_group"])

    # ── Analysis 4: Failed & Infection Cases ──
    def extract_graft_cases(visit_df_primary: pd.DataFrame, status: str) -> pd.DataFrame:
        if "graft_health_text" not in visit_df_primary.columns:
            return pd.DataFrame()
        cols = [c for c in ["tp_mrno_corrected", "tp_mrno", "id", "surg_proc_group", "sap_code", "surg_date", "visit_date", "days_after_surgery"] if c in visit_df_primary.columns]
        return (
            visit_df_primary[
                (visit_df_primary["days_after_surgery"] >= 0) &
                (visit_df_primary["graft_health_text"] == status)
            ]
            .drop_duplicates("tp_mrno_corrected")
            .sort_values("tp_mrno_corrected")
            [cols]
        )

    EXCEL_SHEETS["4_Failed Cases"]    = extract_graft_cases(visit_df_primary, "FAILED")
    EXCEL_SHEETS["4_Infection Cases"] = extract_graft_cases(visit_df_primary, "INFILTRATE/INFECTION")

    # ── Analysis 5: Repeat Surgery Counts ──
    repeat_df = visit_df_primary[
        (visit_df_primary["days_after_surgery"] > 0) &
        (visit_df_primary["advise_surg"] == "YES") &
        (visit_df_primary["fup_surg_done"] == "YES")
    ].copy()
    repeat_df["Repeat_Category"] = repeat_df["fup_done_surg_proc"].apply(categorize_repeat_surgery)

    overall_repeat = (
        repeat_df["Repeat_Category"].value_counts()
        .rename_axis("Repeat Surgery").reset_index(name="Count")
    )
    overall_repeat.loc[len(overall_repeat)] = ["Overall", len(repeat_df)]

    repeat_summary = (
        repeat_df.groupby(["surg_proc_group", "Repeat_Category"]).size()
        .unstack(fill_value=0)
    )
    repeat_summary["Overall"] = repeat_summary.sum(axis=1)
    repeat_summary = repeat_summary.reset_index()

    if not adherence_df.empty and "surg_proc_group" in adherence_df.columns:
        primary_counts = adherence_df.groupby("surg_proc_group").size()
        repeat_summary["Total Primary"] = repeat_summary["surg_proc_group"].map(primary_counts)
        repeat_summary["Repeat %"] = (repeat_summary["Overall"] / repeat_summary["Total Primary"] * 100).round(2)

    kp_details = (
        repeat_df[repeat_df["Repeat_Category"] == "KP"]
        .groupby(["surg_proc_group", "fup_done_surg_proc"]).size()
        .reset_index(name="Count")
        .sort_values(["surg_proc_group", "Count"], ascending=[True, False])
    )

    kp_cols = [c for c in ["tp_mrno_corrected", "id", "surg_proc_group", "fup_done_surg_proc", "days_after_surgery", "visit_date"] if c in repeat_df.columns]
    kp_patients = (
        repeat_df[repeat_df["Repeat_Category"] == "KP"]
        [kp_cols]
        .sort_values(["surg_proc_group", "fup_done_surg_proc", "tp_mrno_corrected"])
    )

    EXCEL_SHEETS["5_Repeat Overall"]   = overall_repeat
    EXCEL_SHEETS["5_Repeat by Surg"]   = repeat_summary
    EXCEL_SHEETS["5_KP Details"]       = kp_details
    EXCEL_SHEETS["5_KP Patients"]      = kp_patients

    # ── Analysis 6: Visual Acuity (VA) Change ──
    VA_df = visit_df_primary[visit_df_primary['surg_proc_group'] != 'THPK'].copy()

    def _latest_va(source_df, condition):
        if "best_va_logmar" not in source_df.columns:
            return pd.DataFrame(columns=["id", "preop_visit_date", "preop_logmar"])
        return (
            source_df[condition & source_df["best_va_logmar"].notna()]
            .sort_values(["id", "visit_date"])
            .groupby("id").tail(1)
            [["id", "visit_date", "best_va_logmar"]]
            .rename(columns={"visit_date": "preop_visit_date", "best_va_logmar": "preop_logmar"})
        )

    preop_before   = _latest_va(VA_df, VA_df["visit_date"] < VA_df["surg_date"])
    preop_same_day = _latest_va(VA_df, VA_df["visit_date"] == VA_df["surg_date"])
    preop = preop_before.combine_first(preop_same_day.set_index("id")).reset_index()

    base = VA_df.groupby("id").first().reset_index()
    keep = [c for c in ["id", "sap_code", "tp_mrno_corrected", "tp_mrno", "gender", "sub_category", "surg_date", "surg_eye", "surg_proc_group", "surgeon_name"] if c in base.columns]
    VA_summary = base[keep].merge(preop, on="id", how="left")

    for period, (start, end) in FOLLOW_UP_PERIODS.items():
        p = period.lower()
        if "best_va_logmar" in VA_df.columns:
            temp = (
                VA_df[
                    VA_df["days_after_surgery"].between(start, end) &
                    VA_df["best_va_logmar"].notna()
                ]
                .sort_values(["id", "visit_date"])
                .groupby("id").tail(1)
                [["id", "visit_date", "best_va_logmar"]]
                .rename(columns={"visit_date": f"{p}_visit_date", "best_va_logmar": f"{p}_logmar"})
            )
            VA_summary = VA_summary.merge(temp, on="id", how="left")
            VA_summary[f"{p}_diff"] = VA_summary["preop_logmar"] - VA_summary[f"{p}_logmar"]
        else:
            VA_summary[f"{p}_logmar"] = np.nan
            VA_summary[f"{p}_diff"] = np.nan

    for period in FOLLOW_UP_PERIODS:
        p = period.lower()
        VA_summary[f"{p}_line_cat"] = VA_summary[f"{p}_diff"].apply(line_category)

    diff_cols = [f"{p.lower()}_diff" for p in FOLLOW_UP_PERIODS]
    avg_change_proc = VA_summary.groupby("surg_proc_group")[diff_cols].mean().round(3).reset_index()
    avg_change_campus_surg = VA_summary.groupby(["sap_code", "surg_proc_group"])[diff_cols].mean().round(3).reset_index()

    line_summaries = {}
    for period in FOLLOW_UP_PERIODS:
        p = period.lower()
        line_summaries[period] = (
            VA_summary
            .groupby(["surg_proc_group", f"{p}_line_cat"]).size()
            .unstack(fill_value=0).reset_index()
        )

    EXCEL_SHEETS["6_VA Patient Detail"]    = VA_summary
    EXCEL_SHEETS["6_VA Avg by Proc"]       = avg_change_proc
    EXCEL_SHEETS["6_VA Avg Camp×Surg"]     = avg_change_campus_surg
    EXCEL_SHEETS["6_VA Lines 1D"]          = line_summaries["1D"]
    EXCEL_SHEETS["6_VA Lines 1W"]          = line_summaries["1W"]
    EXCEL_SHEETS["6_VA Lines 1M"]          = line_summaries["1M"]

    # 8. Create Formatted Excel Buffer
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        for sheet_name, frame in EXCEL_SHEETS.items():
            safe_name = sheet_name[:31]
            frame.to_excel(writer, sheet_name=safe_name, index=False)
        beautify_workbook(writer.book)
    output_buffer.seek(0)

    summary_stats = {
        "total_surgeries": len(surgery_df),
        "unique_patients": surgery_df['tp_mrno_corrected'].nunique() if 'tp_mrno_corrected' in surgery_df.columns else 0,
        "failed_cases": len(EXCEL_SHEETS["4_Failed Cases"]),
        "infection_cases": len(EXCEL_SHEETS["4_Infection Cases"]),
        "repeat_surgeries": len(repeat_df),
        "sheets_count": len(EXCEL_SHEETS)
    }

    return output_buffer, EXCEL_SHEETS, summary_stats

# ==============================================================================
# Main UI Layout
# ==============================================================================
st.markdown("<div class='main-header'>👁️ Corneal Transplant Surgery Morbidity Analyzer</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Upload your raw surgery & follow-up data (.xlsb, .xlsx, .xlsm, .csv) to generate executive analysis reports.</div>", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/ophthalmology.png", width=70)
    st.header("Upload Data File")
    uploaded_file = st.file_uploader(
        "Choose an Excel (.xlsb, .xlsx, .xlsm, .xls) or CSV file",
        type=["xlsb", "xlsx", "xlsm", "xls", "csv"],
        help="Upload the raw data file containing campus surgery & follow-up records."
    )
    
    st.markdown("---")
    st.markdown("### Analysis Pipeline Includes:")
    st.markdown("- 📊 **Surgery Counts & Demographics**")
    st.markdown("- 👁️ **Graft Health Trends (1D, 1W, 1M, 3M)**")
    st.markdown("- 📅 **Follow-Up Adherence & Visits**")
    st.markdown("- ⚠️ **Failed & Infection Case Register**")
    st.markdown("- 🔄 **Repeat Surgery Breakdown & KP Detail**")
    st.markdown("- 📈 **Visual Acuity (LogMAR) Line Changes**")

if uploaded_file is None:
    st.info("👆 Please upload a data file using the sidebar to begin analysis.")
    st.markdown("""
    ### Expected Data Structure
    The uploaded file can be:
    1. An **Excel Workbook** with campus sheets (`KAR Campus Data`, `KVC Campus Data`, etc.) or individual sheets.
    2. A **CSV File** with standardized column headers (`sap_code`, `tp_mrno`, `surg_date`, `graft_health_text`, `surg_proc_group`, etc.).
    
    *Header row auto-detection will locate column headers automatically even if title rows exist at the top.*
    """)
else:
    with st.spinner("Reading uploaded file and running analysis pipeline..."):
        try:
            df_raw = read_uploaded_file(uploaded_file)
            excel_bytes, excel_sheets, stats = run_analysis_pipeline(df_raw)
            st.success(f"✓ Analysis complete! Processed **{stats['total_surgeries']:,}** surgeries across **{stats['sheets_count']}** result tabs.")
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
            st.stop()

    # Metrics Display
    st.markdown("### Key Statistics Overview")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Primary Surgeries", f"{stats['total_surgeries']:,}")
    c2.metric("Unique Patients", f"{stats['unique_patients']:,}")
    c3.metric("Failed Graft Cases", f"{stats['failed_cases']:,}")
    c4.metric("Infection Cases", f"{stats['infection_cases']:,}")
    c5.metric("Repeat Surgeries", f"{stats['repeat_surgeries']:,}")

    st.markdown("---")
    
    # Download Section
    st.markdown("### 📥 Download Resultant Analysis Workbook")
    out_filename = f"Morbidity_Analysis_Results_{datetime.now().strftime('%b%Y')}.xlsx"
    st.download_button(
        label=f"⬇️ Download Executive Report ({out_filename})",
        data=excel_bytes,
        file_name=out_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download the complete executive report formatted with colors, filters, and double bottom borders."
    )

    st.markdown("---")

    # Preview Tabs
    st.markdown("### 📊 Interactive Data Previews")
    preview_tabs = st.tabs([
        "1. Surgery Counts",
        "2. Graft Health",
        "3. Adherence",
        "4. Failed & Infection",
        "5. Repeat Surgeries",
        "6. Visual Acuity Change"
    ])

    with preview_tabs[0]:
        st.subheader("1 · Surgery Counts & Demographics")
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("#### By Surgery Type")
            st.dataframe(excel_sheets["1_By SurgType"], use_container_width=True)
            st.markdown("#### By Campus")
            st.dataframe(excel_sheets["1_By Campus"], use_container_width=True)
        with col_b:
            st.markdown("#### By Gender")
            st.dataframe(excel_sheets["1_By Gender"], use_container_width=True)
            st.markdown("#### By Category")
            st.dataframe(excel_sheets["1_By Category"], use_container_width=True)

    with preview_tabs[1]:
        st.subheader("2 · Graft Health Trends Across Follow-Up Windows")
        st.markdown("#### Graft Health Overall")
        st.dataframe(excel_sheets["2_Graft Overall"], use_container_width=True)
        st.markdown("#### Graft Health by Surgery Type")
        st.dataframe(excel_sheets["2_Graft by Surgery"], use_container_width=True)

    with preview_tabs[2]:
        st.subheader("3 · Follow-Up Adherence & Visit Statistics")
        st.markdown("#### Overall Adherence Status")
        st.dataframe(excel_sheets["3_Adherence Overall"], use_container_width=True)
        st.markdown("#### Average Visits by Procedure")
        st.dataframe(excel_sheets["3_Avg Visits by Proc"], use_container_width=True)

    with preview_tabs[3]:
        st.subheader("4 · Failed & Infection Cases")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            st.markdown(f"#### Failed Graft Cases ({len(excel_sheets['4_Failed Cases'])})")
            st.dataframe(excel_sheets["4_Failed Cases"], use_container_width=True)
        with col_f2:
            st.markdown(f"#### Infection / Infiltrate Cases ({len(excel_sheets['4_Infection Cases'])})")
            st.dataframe(excel_sheets["4_Infection Cases"], use_container_width=True)

    with preview_tabs[4]:
        st.subheader("5 · Repeat Surgery Breakdown")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            st.markdown("#### Overall Repeat Categories")
            st.dataframe(excel_sheets["5_Repeat Overall"], use_container_width=True)
        with col_r2:
            st.markdown("#### Keratoplasty (KP) Repeat Details")
            st.dataframe(excel_sheets["5_KP Details"], use_container_width=True)

    with preview_tabs[5]:
        st.subheader("6 · Visual Acuity (LogMAR) Changes (non-THPK)")
        st.markdown("#### Average LogMAR Improvement by Procedure")
        st.dataframe(excel_sheets["6_VA Avg by Proc"], use_container_width=True)
        st.markdown("#### Line Change Distribution at 1 Month (1M)")
        st.dataframe(excel_sheets["6_VA Lines 1M"], use_container_width=True)
