import io
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st


# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(
    page_title="生產排程反推看板",
    page_icon="🗓️",
    layout="wide",
)

st.title("🗓️ 生產排程反推看板")
st.caption("自動辨識 Excel 工作表、標題列及欄位，依組立地點與客戶入庫日反推生產排程。")


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
    "自製模組": st.sidebar.number_input("自製模組", min_value=0, value=15, step=1),
    "機構": st.sidebar.number_input("機構", min_value=0, value=10, step=1),
    "管路": st.sidebar.number_input("管路", min_value=0, value=10, step=1),
    "電控": st.sidebar.number_input("電控", min_value=0, value=10, step=1),
    "其他": st.sidebar.number_input("其他 Category", min_value=0, value=10, step=1),
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
    result[category_col] = result[category_col].map(clean_text)
else:
    result["Category_系統"] = "其他"
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
if duration_col:
    result["標準工期_計算"] = safe_numeric(result[duration_col], 0).astype(int)
    missing_duration = result["標準工期_計算"] <= 0
    result.loc[missing_duration, "標準工期_計算"] = (
        result.loc[missing_duration, category_col]
        .map(default_category_days)
        .fillna(default_category_days["其他"])
        .astype(int)
    )
else:
    result["標準工期_計算"] = (
        result[category_col]
        .map(default_category_days)
        .fillna(default_category_days["其他"])
        .astype(int)
    )

# Optional quantity; currently for display only
if quantity_col:
    result["數量_計算"] = safe_numeric(result[quantity_col], 1)
else:
    result["數量_計算"] = 1

# Reverse dates
# 反推邏輯：
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


# 將反推結果回寫到既有欄位，避免匯出後原欄位仍空白。
for target_name in ["發料日", "預計發料日"]:
    if target_name in result.columns:
        result[target_name] = result["預計發料日"]

for target_name in ["入庫日", "排程入庫日"]:
    if target_name in result.columns and target_name != customer_date_col:
        result[target_name] = result["排程入庫日"]

# -----------------------------
# Summary
# -----------------------------
st.divider()
st.info(
    "反推公式：客戶入庫日 − 組立地點緩衝工作日 ＝ 排程入庫日；"
    "排程入庫日 − Category 標準工期 ＝ 預計發料日"
)
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
        location_col,
        category_col,
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
