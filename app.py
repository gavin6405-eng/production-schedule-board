import io
import base64
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st


# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(
    page_title="超慧科技｜生產排程反推看板",
    page_icon="🗓️",
    layout="wide",
)



# -----------------------------
# Helpers
# -----------------------------
def clean_text(value) -> str:
    """Normalize text for matching."""
    if pd.isna(value):
        return ""
    return (
        str(value)
        .replace("\n", "")
        .replace("\r", "")
        .replace("\t", "")
        .replace("　", "")
        .strip()
    )


def normalize_column_name(value) -> str:
    return clean_text(value).replace(" ", "")


def find_best_sheet_and_header(
    excel_bytes: bytes,
    max_scan_rows: int = 40,
) -> Tuple[str, int, pd.DataFrame]:
    """
    Find the sheet/header row with the highest keyword score.
    Returns: sheet_name, zero-based header row, preview dataframe.
    """
    keywords = [
        "製令",
        "製令號",
        "工單",
        "客戶入庫日",
        "入庫日",
        "組立地點",
        "組裝地點",
        "組立場所",
        "Category",
        "類別",
        "機型",
        "數量",
    ]

    xls = pd.ExcelFile(io.BytesIO(excel_bytes))
    best_score = -1
    best_sheet = xls.sheet_names[0]
    best_header = 0
    best_preview = pd.DataFrame()

    for sheet in xls.sheet_names:
        try:
            preview = pd.read_excel(
                io.BytesIO(excel_bytes),
                sheet_name=sheet,
                header=None,
                nrows=max_scan_rows,
                dtype=object,
            )
        except Exception:
            continue

        for idx, row in preview.iterrows():
            values = [normalize_column_name(v) for v in row.tolist()]
            joined = "|".join(values)
            score = sum(1 for kw in keywords if normalize_column_name(kw) in joined)

            # Favor rows with several non-empty cells.
            non_empty = sum(bool(v) for v in values)
            if non_empty >= 3:
                score += 0.25

            if score > best_score:
                best_score = score
                best_sheet = sheet
                best_header = int(idx)
                best_preview = preview

    return best_sheet, best_header, best_preview


def deduplicate_columns(columns: List[str]) -> List[str]:
    result = []
    counts: Dict[str, int] = {}
    for col in columns:
        base = normalize_column_name(col) or "未命名欄位"
        counts[base] = counts.get(base, 0) + 1
        result.append(base if counts[base] == 1 else f"{base}_{counts[base]}")
    return result


def find_column(columns: List[str], candidates: List[str]) -> Optional[str]:
    normalized = {normalize_column_name(c): c for c in columns}

    # Exact match first
    for candidate in candidates:
        key = normalize_column_name(candidate)
        if key in normalized:
            return normalized[key]

    # Partial match second
    for col in columns:
        ncol = normalize_column_name(col)
        for candidate in candidates:
            ncandidate = normalize_column_name(candidate)
            if ncandidate and (ncandidate in ncol or ncol in ncandidate):
                return col
    return None



def normalize_model(value) -> str:
    """將 Excel 的機型寫法統一成左側 Category 名稱。"""
    text = clean_text(value)
    key = text.lower().replace(" ", "")

    alias_map = {
        "efem": "EFEM",
        "sort": "sort",
        "sorter": "sort",
        "sorting": "sort",
        "骨包": "骨包",
        "bws": "BWS",
        "bwbs": "BWS",
        "ntb": "NTB",
        "other": "other",
        "其他": "other",
    }
    return alias_map.get(key, "other")


def parse_holidays(text: str) -> List[pd.Timestamp]:
    holidays: List[pd.Timestamp] = []
    for raw in text.replace("，", ",").replace("\n", ",").split(","):
        raw = raw.strip()
        if not raw:
            continue
        try:
            holidays.append(pd.Timestamp(raw).normalize())
        except Exception:
            pass
    return holidays


def workday_offset(
    start_date,
    offset_days: int,
    holidays: List[pd.Timestamp],
) -> pd.Timestamp:
    """
    Offset by business days. Negative means reverse scheduling.
    """
    if pd.isna(start_date):
        return pd.NaT

    ts = pd.Timestamp(start_date).normalize()
    holiday_dates = {h.date() for h in holidays}
    step = 1 if offset_days >= 0 else -1
    remaining = abs(int(offset_days))

    while remaining > 0:
        ts += pd.Timedelta(days=step)
        if ts.weekday() < 5 and ts.date() not in holiday_dates:
            remaining -= 1

    return ts


def safe_numeric(series: pd.Series, default: float = 0) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(default)


# -----------------------------
# Sidebar parameters
# -----------------------------

logo_base64 = "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"

st.markdown(
    f"""
    <style>
    .stApp {
        background:
            radial-gradient(circle at 15% 10%, rgba(239,68,68,0.15), transparent 28%),
            radial-gradient(circle at 85% 12%, rgba(14,165,233,0.12), transparent 25%),
            linear-gradient(135deg, #050914 0%, #0B1220 45%, #111827 100%);
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #171B28 0%, #0E1420 100%);
        border-right: 1px solid rgba(255,255,255,0.08);
    }
    .tech-cover {
        position: relative;
        overflow: hidden;
        padding: 34px 28px 30px;
        border-radius: 24px;
        border: 1px solid rgba(255,255,255,0.10);
        background:
            linear-gradient(135deg, rgba(255,255,255,0.07), rgba(255,255,255,0.02)),
            linear-gradient(135deg, rgba(239,68,68,0.12), rgba(14,165,233,0.07));
        box-shadow: 0 18px 60px rgba(0,0,0,0.35);
        margin-bottom: 24px;
    }
    .tech-cover::before {
        content: "";
        position: absolute;
        inset: 0;
        background-image:
            linear-gradient(rgba(255,255,255,0.025) 1px, transparent 1px),
            linear-gradient(90deg, rgba(255,255,255,0.025) 1px, transparent 1px);
        background-size: 28px 28px;
    }
    .tech-logo-wrap {
        position: relative;
        z-index: 2;
        display: flex;
        justify-content: center;
        margin-bottom: 16px;
    }
    .tech-logo {
        width: 320px;
        max-width: 80%;
        padding: 14px 24px;
        border-radius: 18px;
        background: rgba(255,255,255,0.97);
        box-shadow: 0 10px 30px rgba(0,0,0,0.28);
    }
    .tech-title {
        position: relative;
        z-index: 2;
        text-align: center;
        font-size: 42px;
        font-weight: 800;
        letter-spacing: 2px;
        color: #F8FAFC;
        margin-top: 6px;
    }
    .tech-subtitle {
        position: relative;
        z-index: 2;
        text-align: center;
        color: #B8C4D6;
        margin-top: 8px;
        font-size: 16px;
    }
    .tech-version {
        position: relative;
        z-index: 2;
        display: block;
        width: fit-content;
        margin: 16px auto 0;
        padding: 7px 14px;
        border-radius: 999px;
        color: #7DD3FC;
        border: 1px solid rgba(14,165,233,0.35);
        background: rgba(14,165,233,0.10);
        font-size: 13px;
    }
    h1, h2, h3 {
        color: #F8FAFC !important;
    }
    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.045);
        border: 1px solid rgba(255,255,255,0.08);
        padding: 14px;
        border-radius: 16px;
    }
    .stButton > button, .stDownloadButton > button {
        border-radius: 12px;
        border: 1px solid rgba(255,255,255,0.12);
        background: linear-gradient(135deg, #E11D48, #EF4444);
        color: white;
        font-weight: 700;
    }
    </style>

    <div class="tech-cover">
        <div class="tech-logo-wrap">
            <img class="tech-logo" src="data:image/png;base64,{logo_base64}">
        </div>
        <div class="tech-title">生產排程反推看板</div>
        <div class="tech-subtitle">客戶入庫日 × 組立地點緩衝 × Category 標準工期</div>
        <span class="tech-version">SUPER PLUS TECH｜2026-07-01-v8</span>
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.header("⚙️ 排程參數")

st.sidebar.subheader("組立地點緩衝工作日")
location_buffer: Dict[str, int] = {
    "竹東": st.sidebar.number_input("竹東", min_value=0, value=2, step=1),
    "模冠": st.sidebar.number_input("模冠", min_value=0, value=5, step=1),
    "御弘": st.sidebar.number_input("御弘", min_value=0, value=7, step=1),
    "宏田": st.sidebar.number_input("宏田", min_value=0, value=3, step=1),
}

unknown_location_buffer = st.sidebar.number_input(
    "未辨識地點預設緩衝",
    min_value=0,
    value=0,
    step=1,
    help="Excel 中出現其他地點或空白時使用，避免 KeyError。",
)

st.sidebar.subheader("Category 標準工期")
default_category_days: Dict[str, int] = {
    "EFEM": st.sidebar.number_input("EFEM", min_value=0, value=15, step=1),
    "sort": st.sidebar.number_input("sort", min_value=0, value=15, step=1),
    "骨包": st.sidebar.number_input("骨包", min_value=0, value=15, step=1),
    "BWS": st.sidebar.number_input("BWS", min_value=0, value=21, step=1),
    "NTB": st.sidebar.number_input("NTB", min_value=0, value=15, step=1),
    "other": st.sidebar.number_input("other", min_value=0, value=10, step=1),
}

st.sidebar.subheader("假日設定")
holiday_text = st.sidebar.text_area(
    "額外排除日期",
    placeholder="例如：2026-07-01, 2026-09-25",
    help="週六、週日會自動排除；公司休假日請以逗號分隔。",
)
holidays = parse_holidays(holiday_text)


# -----------------------------
# File upload
# -----------------------------
uploaded_file = st.file_uploader(
    "上傳生產排程檔案",
    type=["xlsx", "xlsm", "xls"],
)

if uploaded_file is None:
    st.info("請先上傳 Excel 檔案。")
    st.stop()

excel_bytes = uploaded_file.getvalue()

try:
    detected_sheet, detected_header, _ = find_best_sheet_and_header(excel_bytes)
except Exception as exc:
    st.error(f"無法讀取 Excel：{exc}")
    st.stop()

try:
    xls = pd.ExcelFile(io.BytesIO(excel_bytes))
except Exception as exc:
    st.error(f"Excel 格式無法解析：{exc}")
    st.stop()

col1, col2 = st.columns(2)
with col1:
    selected_sheet = st.selectbox(
        "工作表",
        options=xls.sheet_names,
        index=xls.sheet_names.index(detected_sheet),
    )
with col2:
    header_row_display = st.number_input(
        "標題列（Excel 列號）",
        min_value=1,
        value=detected_header + 1,
        step=1,
    )

header_row = int(header_row_display) - 1

try:
    df = pd.read_excel(
        io.BytesIO(excel_bytes),
        sheet_name=selected_sheet,
        header=header_row,
        dtype=object,
    )
except Exception as exc:
    st.error(f"讀取工作表失敗：{exc}")
    st.stop()

df.columns = deduplicate_columns(list(df.columns))
df = df.dropna(how="all").copy()

st.success(f"已讀取工作表：{selected_sheet}；標題列：第 {header_row + 1} 列")

with st.expander("目前讀到的欄位", expanded=False):
    st.write(list(df.columns))


# -----------------------------
# Column detection
# -----------------------------
columns = list(df.columns)

order_col = find_column(
    columns,
    ["製令", "製令號", "製造命令", "工單", "工單號", "MO"],
)
location_col = find_column(
    columns,
    ["組立地點", "組裝地點", "組立場所", "組裝場所", "生產地點"],
)
customer_date_col = find_column(
    columns,
    ["客戶入庫日", "客戶納入日", "客戶需求日", "入庫日", "交期", "出貨日"],
)
category_col = find_column(
    columns,
    ["Category", "類別", "分類", "製程類別", "機型類別"],
)
duration_col = find_column(
    columns,
    ["標準工期", "工期", "標準工作日", "生產工作日", "需求工時"],
)
quantity_col = find_column(
    columns,
    ["數量", "台數", "需求數量", "訂單數量"],
)

st.subheader("欄位對應")

def mapping_select(label: str, detected: Optional[str], allow_none: bool = True):
    options = ["（不使用）"] + columns if allow_none else columns
    if detected in columns:
        index = options.index(detected)
    else:
        index = 0
    return st.selectbox(label, options, index=index)

m1, m2, m3 = st.columns(3)
with m1:
    order_col = mapping_select("製令欄位", order_col)
    location_col = mapping_select("組立地點欄位", location_col)
with m2:
    customer_date_col = mapping_select("客戶入庫日欄位", customer_date_col)
    category_col = mapping_select("Category 欄位", category_col)
with m3:
    duration_col = mapping_select("標準工期欄位", duration_col)
    quantity_col = mapping_select("數量欄位", quantity_col)

def none_if_unused(value: str) -> Optional[str]:
    return None if value == "（不使用）" else value

order_col = none_if_unused(order_col)
location_col = none_if_unused(location_col)
customer_date_col = none_if_unused(customer_date_col)
category_col = none_if_unused(category_col)
duration_col = none_if_unused(duration_col)
quantity_col = none_if_unused(quantity_col)

if customer_date_col is None:
    st.error("請指定「客戶入庫日」欄位，才能反推排程。")
    st.stop()


# -----------------------------
# Data processing
# -----------------------------
result = df.copy()

# Clean key text columns
if location_col:
    result[location_col] = result[location_col].map(clean_text)
else:
    result["組立地點_系統"] = ""
    location_col = "組立地點_系統"

if category_col:
    result["Category_原始值"] = result[category_col].map(clean_text)
    result[category_col] = result[category_col].map(normalize_model)
else:
    result["Category_原始值"] = ""
    result["Category_系統"] = "other"
    category_col = "Category_系統"

# Parse date
result[customer_date_col] = pd.to_datetime(
    result[customer_date_col],
    errors="coerce",
)

# Location buffer: robust fallback, no KeyError
result["地點緩衝工作日"] = (
    result[location_col]
    .map(location_buffer)
    .fillna(unknown_location_buffer)
    .astype(int)
)

# Standard duration
# 一律依左側 Category 標準工期設定，不再採用 Excel 原有標工欄位。
result["標準工期_計算"] = (
    result[category_col]
    .map(default_category_days)
    .fillna(default_category_days["other"])
    .astype(int)
)

# Optional quantity; currently for display only
if quantity_col:
    result["數量_計算"] = safe_numeric(result[quantity_col], 1)
else:
    result["數量_計算"] = 1

# Reverse dates
# 客戶入庫日 - 組立地點緩衝工作日 = 排程入庫日
# 排程入庫日 - Category 標準工期 = 預計發料日
result["排程入庫日"] = result.apply(
    lambda r: workday_offset(
        r[customer_date_col],
        -int(r["地點緩衝工作日"]),
        holidays,
    ),
    axis=1,
)

result["預計發料日"] = result.apply(
    lambda r: workday_offset(
        r["排程入庫日"],
        -int(r["標準工期_計算"]),
        holidays,
    ),
    axis=1,
)

result["反推總工作日"] = (
    result["地點緩衝工作日"] + result["標準工期_計算"]
).astype(int)

today = pd.Timestamp(date.today())
result["排程狀態"] = np.select(
    [
        result[customer_date_col].isna(),
        result["排程入庫日"] < today,
        result["預計發料日"] <= today,
    ],
    [
        "缺少客戶入庫日",
        "已逾排程入庫日",
        "應進行中",
    ],
    default="未開始",
)


# -----------------------------
# Summary
# -----------------------------
st.divider()
st.info(
    "反推公式：客戶入庫日 − 組立地點緩衝工作日 ＝ 排程入庫日；"
    "排程入庫日 − Category 標準工期 ＝ 預計發料日。"
)
st.caption("標準工期完全依左側設定；Excel 原有標工欄位不參與計算。")
st.subheader("排程摘要")

total_rows = len(result)
valid_dates = int(result[customer_date_col].notna().sum())
overdue = int((result["排程狀態"] == "已逾排程入庫日").sum())
in_progress = int((result["排程狀態"] == "應進行中").sum())

k1, k2, k3, k4 = st.columns(4)
k1.metric("總筆數", f"{total_rows:,}")
k2.metric("有效入庫日", f"{valid_dates:,}")
k3.metric("已逾排程入庫日", f"{overdue:,}")
k4.metric("應進行中", f"{in_progress:,}")

display_cols = [
    c
    for c in [
        order_col,
        "Category_原始值",
        category_col,
        location_col,
        customer_date_col,
        "標準工期_計算",
        "地點緩衝工作日",
        "反推總工作日",
        "預計發料日",
        "排程入庫日",
        "排程狀態",
    ]
    if c is not None and c in result.columns
]

st.dataframe(
    result[display_cols],
    use_container_width=True,
    hide_index=True,
)


# -----------------------------
# Location summary
# -----------------------------
st.subheader("組立地點彙整")
location_summary = (
    result.groupby(location_col, dropna=False)
    .agg(
        筆數=(location_col, "size"),
        最早發料日=("預計發料日", "min"),
        最晚排程入庫日=("排程入庫日", "max"),
        已逾期=("排程狀態", lambda s: int((s == "已逾排程入庫日").sum())),
    )
    .reset_index()
)
st.dataframe(location_summary, use_container_width=True, hide_index=True)


# -----------------------------
# Download
# -----------------------------
output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    result.to_excel(writer, sheet_name="反推排程結果", index=False)
    location_summary.to_excel(writer, sheet_name="組立地點彙整", index=False)

output.seek(0)

st.download_button(
    "⬇️ 下載反推排程 Excel",
    data=output,
    file_name=f"生產排程反推結果_{datetime.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
