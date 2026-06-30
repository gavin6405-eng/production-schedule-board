
import io
from datetime import date
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(
    page_title="生產排程反推看板",
    page_icon="📅",
    layout="wide",
)

st.title("📅 生產排程反推看板")
st.caption("依組立地點、標準工期、發料日、入庫日與客戶入庫日，自動反推並判定排程風險。")

DEFAULT_LOCATION_BUFFER = {
    "竹東": 2,
    "竹南": 5,
    "外包": 7,
    "其他": 3,
}

DEFAULT_CATEGORY_DAYS = {
    "自製模組": 15,
    "Sorter": 25,
    "BWBS": 45,
    "PACKING PARTS COMP": 15,
    "其他": 20,
}

COLUMN_ALIASES = {
    "製令": ["製令", "工單", "MO", "Order"],
    "客戶": ["客戶", "Customer"],
    "P/N": ["P/N", "PN", "料號", "品號"],
    "Type": ["Type", "機型", "產品類型"],
    "Category": ["Category", "類別", "分類"],
    "組立地點": ["組立地點", "組裝地點", "地點"],
    "組立人員": ["組立人員", "組裝人員", "人員"],
    "組立進度": ["組立進度", "組裝進度", "進度"],
    "備註": ["備註", "Remark", "Remarks"],
    "發料日": ["發料日", "Release Date"],
    "入庫日": ["入庫日", "Warehouse Date"],
    "客戶入庫日": ["客戶入庫日", "Customer Due Date", "客戶交期"],
    "最晚到料日": ["最晚到料日", "到料日", "缺料到料日", "客供料到料日"],
    "標準工期": ["標準工期", "標準組裝工作日", "工期"],
}

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    existing = {str(c).strip(): c for c in df.columns}
    for target, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in existing:
                rename_map[existing[alias]] = target
                break
    return df.rename(columns=rename_map)

def parse_date_series(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

def business_day_offset(start, days, holidays):
    if pd.isna(start):
        return pd.NaT
    start_np = np.datetime64(pd.Timestamp(start).date(), "D")
    holiday_np = np.array(
        [np.datetime64(pd.Timestamp(h).date(), "D") for h in holidays],
        dtype="datetime64[D]"
    )
    result = np.busday_offset(
        start_np,
        int(days),
        roll="backward" if int(days) < 0 else "forward",
        holidays=holiday_np
    )
    return pd.Timestamp(result)

def business_days_between(start, end, holidays):
    if pd.isna(start) or pd.isna(end):
        return np.nan
    s = np.datetime64(pd.Timestamp(start).date(), "D")
    e = np.datetime64(pd.Timestamp(end).date(), "D")
    holiday_np = np.array(
        [np.datetime64(pd.Timestamp(h).date(), "D") for h in holidays],
        dtype="datetime64[D]"
    )
    return int(np.busday_count(s, e, holidays=holiday_np))

def read_uploaded_file(uploaded_file):
    suffix = uploaded_file.name.lower().split(".")[-1]
    if suffix == "csv":
        raw = uploaded_file.getvalue()
        for enc in ("utf-8-sig", "cp950", "big5"):
            try:
                return pd.read_csv(io.BytesIO(raw), encoding=enc)
            except Exception:
                pass
        raise ValueError("CSV 編碼無法辨識，請另存為 UTF-8 CSV 後再上傳。")
    if suffix in ("xlsx", "xls"):
        return pd.read_excel(uploaded_file)
    raise ValueError("只支援 CSV、XLSX、XLS。")

with st.sidebar:
    st.header("⚙️ 排程參數")

    st.subheader("組立地點緩衝工作日")
    location_buffer = {}
    for k, v in DEFAULT_LOCATION_BUFFER.items():
        location_buffer[k] = st.number_input(
            f"{k}緩衝日",
            min_value=0,
            max_value=60,
            value=v,
            step=1,
            key=f"loc_{k}"
        )

    st.subheader("Category標準工期")
    category_days = {}
    for k, v in DEFAULT_CATEGORY_DAYS.items():
        category_days[k] = st.number_input(
            f"{k}工期",
            min_value=1,
            max_value=180,
            value=v,
            step=1,
            key=f"cat_{k}"
        )

    st.subheader("假日設定")
    holiday_text = st.text_area(
        "每行輸入一個日期（YYYY-MM-DD）",
        placeholder="2026-01-01\n2026-02-16"
    )
    holidays = []
    for line in holiday_text.splitlines():
        line = line.strip()
        if line:
            try:
                holidays.append(pd.Timestamp(line))
            except Exception:
                st.warning(f"假日格式無法辨識：{line}")

uploaded_file = st.file_uploader(
    "上傳生產排程檔案",
    type=["xlsx", "xls", "csv"],
    help="需至少包含：組立地點、客戶入庫日；建議包含 Category、發料日、入庫日、最晚到料日。"
)

with st.expander("欄位需求與計算邏輯", expanded=False):
    st.markdown("""
**必要欄位**
- 組立地點
- 客戶入庫日

**建議欄位**
- 製令、客戶、P/N、Type、Category
- 發料日、入庫日、最晚到料日、標準工期

**計算方式**
1. 建議入庫日 = 客戶入庫日 − 組立地點緩衝工作日  
2. 建議發料日 = 建議入庫日 − 標準工期  
3. 實際可開工日 = 建議發料日與最晚到料日取較晚者  
4. 預估入庫日 = 實際可開工日 + 標準工期  
5. 預估入庫日晚於建議入庫日，判定為「排程需變更」
""")

if uploaded_file is None:
    st.info("請先上傳 Excel 或 CSV 排程檔案。")
    st.stop()

try:
    df = normalize_columns(read_uploaded_file(uploaded_file))
except Exception as e:
    st.error(f"檔案讀取失敗：{e}")
    st.stop()

required = ["組立地點", "客戶入庫日"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"缺少必要欄位：{', '.join(missing)}")
    st.write("目前讀到的欄位：", list(df.columns))
    st.stop()

for col in ["發料日", "入庫日", "客戶入庫日", "最晚到料日"]:
    if col in df.columns:
        df[col] = parse_date_series(df[col])

if "Category" not in df.columns:
    df["Category"] = "其他"
if "標準工期" not in df.columns:
    df["標準工期"] = np.nan
if "最晚到料日" not in df.columns:
    df["最晚到料日"] = pd.NaT

df["地點緩衝工作日"] = df["組立地點"].astype(str).map(location_buffer).fillna(location_buffer["其他"]).astype(int)

def get_standard_days(row):
    manual = pd.to_numeric(row.get("標準工期"), errors="coerce")
    if pd.notna(manual) and manual > 0:
        return int(manual)
    category = str(row.get("Category", "其他")).strip()
    return int(category_days.get(category, category_days["其他"]))

df["標準組裝工作日"] = df.apply(get_standard_days, axis=1)

df["建議入庫日"] = df.apply(
    lambda r: business_day_offset(r["客戶入庫日"], -r["地點緩衝工作日"], holidays),
    axis=1
)
df["建議發料日"] = df.apply(
    lambda r: business_day_offset(r["建議入庫日"], -r["標準組裝工作日"], holidays),
    axis=1
)

def calc_actual_start(row):
    suggested = row["建議發料日"]
    material = row["最晚到料日"]
    if pd.isna(material):
        return suggested
    if pd.isna(suggested):
        return material
    return max(suggested, material)

df["實際可開工日"] = df.apply(calc_actual_start, axis=1)
df["預估入庫日"] = df.apply(
    lambda r: business_day_offset(r["實際可開工日"], r["標準組裝工作日"], holidays),
    axis=1
)

df["入庫至客戶工作日"] = df.apply(
    lambda r: business_days_between(r["入庫日"], r["客戶入庫日"], holidays)
    if "入庫日" in df.columns else np.nan,
    axis=1
)

def judge(row):
    if pd.isna(row["客戶入庫日"]) or pd.isna(row["建議入庫日"]):
        return "資料不足"
    if pd.notna(row["預估入庫日"]) and row["預估入庫日"] > row["建議入庫日"]:
        return "排程需變更"
    if "入庫日" in row and pd.notna(row.get("入庫日")):
        gap = row.get("入庫至客戶工作日")
        if pd.notna(gap):
            if gap < 0:
                return "已逾期"
            if gap < row["地點緩衝工作日"]:
                return "緩衝不足"
    return "正常"

df["排程判定"] = df.apply(judge, axis=1)

def reason(row):
    if row["排程判定"] == "排程需變更":
        if pd.notna(row["最晚到料日"]) and row["最晚到料日"] > row["建議發料日"]:
            return "最晚到料日晚於建議發料日"
        return "預估入庫日晚於建議入庫日"
    if row["排程判定"] == "緩衝不足":
        return "原入庫日至客戶入庫日的工作日不足"
    if row["排程判定"] == "已逾期":
        return "原入庫日晚於客戶入庫日"
    if row["排程判定"] == "資料不足":
        return "客戶入庫日或必要日期缺漏"
    return ""

df["異常原因"] = df.apply(reason, axis=1)

total = len(df)
normal = int((df["排程判定"] == "正常").sum())
change = int((df["排程判定"] == "排程需變更").sum())
warning = int(df["排程判定"].isin(["緩衝不足", "已逾期", "資料不足"]).sum())

c1, c2, c3, c4 = st.columns(4)
c1.metric("總筆數", total)
c2.metric("正常", normal)
c3.metric("排程需變更", change)
c4.metric("其他異常", warning)

st.subheader("排程分析結果")

preferred_cols = [
    "製令", "客戶", "P/N", "Type", "Category", "組立地點",
    "組立進度", "備註", "發料日", "入庫日", "客戶入庫日",
    "最晚到料日", "地點緩衝工作日", "標準組裝工作日",
    "建議發料日", "建議入庫日", "實際可開工日",
    "預估入庫日", "排程判定", "異常原因"
]
show_cols = [c for c in preferred_cols if c in df.columns]

def highlight_status(row):
    status = row.get("排程判定", "")
    if status == "正常":
        return ["background-color: #d9ead3"] * len(row)
    if status in ("緩衝不足", "資料不足"):
        return ["background-color: #fff2cc"] * len(row)
    if status in ("排程需變更", "已逾期"):
        return ["background-color: #f4cccc"] * len(row)
    return [""] * len(row)

st.dataframe(
    df[show_cols].style.apply(highlight_status, axis=1),
    use_container_width=True,
    hide_index=True,
    height=520
)

st.subheader("依組立地點統計")
location_summary = (
    df.groupby(["組立地點", "排程判定"], dropna=False)
      .size()
      .reset_index(name="筆數")
)
st.dataframe(location_summary, use_container_width=True, hide_index=True)

output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="排程分析結果")
    location_summary.to_excel(writer, index=False, sheet_name="地點統計")
output.seek(0)

st.download_button(
    "⬇️ 下載排程分析 Excel",
    data=output,
    file_name="生產排程反推分析結果.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
