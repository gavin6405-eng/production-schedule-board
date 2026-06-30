
import io
import re
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="生產排程反推看板", page_icon="📅", layout="wide")
st.title("📅 生產排程反推看板")
st.caption("自動辨識 Excel 工作表、標題列及欄位名稱，再依組立地點與客戶入庫日反推排程。")

COLUMN_ALIASES = {
    "製令": ["製令", "工單", "mo", "order"],
    "客戶": ["客戶", "customer"],
    "P/N": ["p/n", "pn", "料號", "品號"],
    "Type": ["type", "機型", "產品類型"],
    "Category": ["category", "類別", "分類"],
    "組立地點": ["組立地點", "組裝地點", "地點"],
    "組立人員": ["組立人員", "組裝人員", "人員"],
    "組立進度": ["組立進度", "組裝進度", "進度"],
    "備註": ["備註", "remark", "remarks"],
    "發料日": ["發料日", "release date"],
    "入庫日": ["入庫日", "warehouse date"],
    "客戶入庫日": ["客戶入庫日", "客戶交期", "customer due date"],
    "最晚到料日": ["最晚到料日", "到料日", "缺料到料日", "客供料到料日"],
    "標準工期": ["標準工期", "標準組裝工作日", "工期"],
}

DEFAULT_LOCATION_BUFFER = {"竹東": 5, "模冠": 3, "御弘": 3, "宏田": 3}
DEFAULT_CATEGORY_DAYS = {"自製模組": 15, "Sorter": 25, "BWBS": 45, "PACKING PARTS COMP": 15, "其他": 20}

def clean_text(v):
    if pd.isna(v):
        return ""
    s = str(v).replace("\n", "").replace("\r", "").replace(" ", "").strip()
    return s.lower()

def score_header_row(values):
    cleaned = [clean_text(v) for v in values]
    score = 0
    for aliases in COLUMN_ALIASES.values():
        alias_clean = {clean_text(a) for a in aliases}
        if any(c in alias_clean for c in cleaned):
            score += 1
    return score

def detect_best_sheet_and_header(raw_bytes):
    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
    best = None
    diagnostics = []

    for sheet in xls.sheet_names:
        preview = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sheet, header=None, nrows=30)
        for row_idx in range(min(30, len(preview))):
            score = score_header_row(preview.iloc[row_idx].tolist())
            diagnostics.append((sheet, row_idx, score))
            if best is None or score > best[2]:
                best = (sheet, row_idx, score)

    if best is None or best[2] < 2:
        raise ValueError("找不到標題列。請確認 Excel 中至少有「組立地點」及「客戶入庫日」兩個欄位。")

    return best[0], best[1], diagnostics

def normalize_columns(df):
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all").copy()

    rename_map = {}
    for col in df.columns:
        c = clean_text(col)
        for target, aliases in COLUMN_ALIASES.items():
            if c in {clean_text(a) for a in aliases}:
                rename_map[col] = target
                break

    df = df.rename(columns=rename_map)
    return df

def read_file(uploaded):
    raw = uploaded.getvalue()
    name = uploaded.name.lower()

    if name.endswith(".csv"):
        for enc in ("utf-8-sig", "cp950", "big5"):
            try:
                df = pd.read_csv(io.BytesIO(raw), encoding=enc)
                return normalize_columns(df), "CSV", 0
            except Exception:
                pass
        raise ValueError("CSV 編碼無法辨識。")

    if name.endswith((".xlsx", ".xls")):
        sheet, header_row, _ = detect_best_sheet_and_header(raw)
        df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet, header=header_row)
        return normalize_columns(df), sheet, header_row + 1

    raise ValueError("只支援 xlsx、xls、csv。")

def business_day_offset(start, days, holidays):
    if pd.isna(start):
        return pd.NaT
    holiday_np = np.array(
        [np.datetime64(pd.Timestamp(h).date(), "D") for h in holidays],
        dtype="datetime64[D]"
    )
    result = np.busday_offset(
        np.datetime64(pd.Timestamp(start).date(), "D"),
        int(days),
        roll="backward" if int(days) < 0 else "forward",
        holidays=holiday_np
    )
    return pd.Timestamp(result)

with st.sidebar:
    st.header("⚙️ 排程參數")

    st.subheader("組立地點緩衝工作日")
    location_buffer = {
        k: st.number_input(f"{k}", min_value=0, max_value=60, value=v, step=1, key=f"loc_{k}")
        for k, v in DEFAULT_LOCATION_BUFFER.items()
    }

    st.subheader("Category 標準工期")
    category_days = {
        k: st.number_input(f"{k}", min_value=1, max_value=180, value=v, step=1, key=f"cat_{k}")
        for k, v in DEFAULT_CATEGORY_DAYS.items()
    }

    holiday_text = st.text_area("假日（每行一個 YYYY-MM-DD）")
    holidays = []
    for x in holiday_text.splitlines():
        x = x.strip()
        if x:
            try:
                holidays.append(pd.Timestamp(x))
            except Exception:
                st.warning(f"假日格式錯誤：{x}")

uploaded = st.file_uploader("上傳生產排程檔案", type=["xlsx", "xls", "csv"])

if uploaded is None:
    st.info("請上傳 Excel 或 CSV。")
    st.stop()

try:
    df, detected_sheet, detected_header = read_file(uploaded)
except Exception as e:
    st.error(f"檔案讀取失敗：{e}")
    st.stop()

st.success(f"已辨識工作表：{detected_sheet}；標題列：第 {detected_header} 列")
st.write("目前讀到的欄位：", list(df.columns))

required = ["組立地點", "客戶入庫日"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"仍缺少必要欄位：{', '.join(missing)}")
    st.info("可能原因：欄位有合併儲存格、特殊空白字元，或資料不在目前工作表。")
    st.stop()

for col in ["發料日", "入庫日", "客戶入庫日", "最晚到料日"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce")

if "Category" not in df.columns:
    df["Category"] = "其他"
if "最晚到料日" not in df.columns:
    df["最晚到料日"] = pd.NaT
if "標準工期" not in df.columns:
    df["標準工期"] = np.nan

df["地點緩衝工作日"] = (
    df["組立地點"].astype(str).str.strip()
      .map(location_buffer)
      .fillna(location_buffer["其他"])
      .astype(int)
)

def get_days(row):
    manual = pd.to_numeric(row.get("標準工期"), errors="coerce")
    if pd.notna(manual) and manual > 0:
        return int(manual)
    category = str(row.get("Category", "其他")).strip()
    return int(category_days.get(category, category_days["其他"]))

df["標準組裝工作日"] = df.apply(get_days, axis=1)
df["建議入庫日"] = df.apply(
    lambda r: business_day_offset(r["客戶入庫日"], -r["地點緩衝工作日"], holidays), axis=1
)
df["建議發料日"] = df.apply(
    lambda r: business_day_offset(r["建議入庫日"], -r["標準組裝工作日"], holidays), axis=1
)

def actual_start(row):
    a, b = row["建議發料日"], row["最晚到料日"]
    if pd.isna(b):
        return a
    if pd.isna(a):
        return b
    return max(a, b)

df["實際可開工日"] = df.apply(actual_start, axis=1)
df["預估入庫日"] = df.apply(
    lambda r: business_day_offset(r["實際可開工日"], r["標準組裝工作日"], holidays), axis=1
)

def judge(row):
    if pd.isna(row["客戶入庫日"]):
        return "資料不足"
    if pd.notna(row["預估入庫日"]) and row["預估入庫日"] > row["建議入庫日"]:
        return "排程需變更"
    if "入庫日" in df.columns and pd.notna(row.get("入庫日")):
        if row["入庫日"] > row["客戶入庫日"]:
            return "已逾期"
        if row["入庫日"] > row["建議入庫日"]:
            return "緩衝不足"
    return "正常"

df["排程判定"] = df.apply(judge, axis=1)

c1, c2, c3 = st.columns(3)
c1.metric("總筆數", len(df))
c2.metric("正常", int((df["排程判定"] == "正常").sum()))
c3.metric("需注意", int((df["排程判定"] != "正常").sum()))

show_cols = [c for c in [
    "製令", "客戶", "P/N", "Type", "Category", "組立地點", "組立進度", "備註",
    "發料日", "入庫日", "客戶入庫日", "最晚到料日",
    "地點緩衝工作日", "標準組裝工作日",
    "建議發料日", "建議入庫日", "實際可開工日", "預估入庫日", "排程判定"
] if c in df.columns]

st.dataframe(df[show_cols], use_container_width=True, hide_index=True, height=550)

output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="排程分析結果")
output.seek(0)

st.download_button(
    "⬇️ 下載分析結果",
    output,
    "生產排程反推分析結果.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
