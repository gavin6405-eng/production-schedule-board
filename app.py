
import io
import re
from datetime import datetime, date, timedelta

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="每日生產排程看板",
    page_icon="🏭",
    layout="wide"
)

FIXED_ALIASES = {
    "製令": ["製令", "製令號", "工單", "工單號"],
    "客戶": ["客戶", "Customer"],
    "P/N": ["P/N", "PN", "料號", "品號"],
    "Type": ["Type", "機型", "品名"],
    "Category": ["Category", "類別", "分類"],
    "組立地點": ["組立地點", "組裝地點", "地點"],
    "組立人員": ["組立人員", "組裝人員", "人員"],
    "組立進度": ["組立進度", "組裝進度", "進度"],
    "備註": ["備註", "Remark", "Remarks"],
    "發料日": ["發料日", "領料日"],
    "入庫日": ["入庫日", "公司入庫日"],
    "客戶入庫日": ["客戶入庫日", "客戶需求日", "交期"]
}

CODE_LABELS = {
    "組": "組立",
    "T": "測試",
    "Q": "品質確認",
    "R": "調整/重工",
    "B": "包裝",
    "W": "等待",
    "出": "出貨/入庫",
    "機": "機構",
    "配": "配線",
    "管": "管路",
    "-": "無排程",
    "": "無排程"
}

RISK_ORDER = ["逾期", "今日應完成", "7日內", "正常", "已完成"]
STATUS_DONE_WORDS = ["已完成", "完成", "入庫", "出貨"]

def clean_col_name(value):
    if value is None:
        return ""
    return str(value).strip().replace("\n", "").replace(" ", "")

def normalize_fixed_columns(df):
    rename = {}
    cleaned_map = {c: clean_col_name(c) for c in df.columns}
    for target, aliases in FIXED_ALIASES.items():
        alias_clean = [clean_col_name(x).lower() for x in aliases]
        for original, cleaned in cleaned_map.items():
            if cleaned.lower() in alias_clean:
                rename[original] = target
                break
    df = df.rename(columns=rename)

    required = ["製令", "客戶入庫日"]
    missing = [x for x in required if x not in df.columns]
    if missing:
        raise ValueError("缺少必要欄位：" + "、".join(missing))

    for col in FIXED_ALIASES:
        if col not in df.columns:
            df[col] = ""
    return df

def parse_date_header(col, default_year):
    if isinstance(col, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(col).normalize()

    text = str(col).strip()
    text = text.replace("年", "/").replace("月", "/").replace("日", "")
    text = re.sub(r"\s+", "", text)

    patterns = [
        (r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$", True),
        (r"^(\d{1,2})[/-](\d{1,2})$", False),
        (r"^(\d{8})$", True),
        (r"^(\d{4})$", False),
    ]

    for pattern, has_year in patterns:
        m = re.match(pattern, text)
        if not m:
            continue
        try:
            if pattern == r"^(\d{8})$":
                return pd.to_datetime(m.group(1), format="%Y%m%d").normalize()
            if pattern == r"^(\d{4})$":
                mmdd = m.group(1)
                return pd.Timestamp(default_year, int(mmdd[:2]), int(mmdd[2:]))
            if has_year:
                return pd.Timestamp(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return pd.Timestamp(default_year, int(m.group(1)), int(m.group(2)))
        except Exception:
            return None
    return None

def detect_timeline_columns(df):
    years = []
    for col in ["發料日", "入庫日", "客戶入庫日"]:
        if col in df.columns:
            vals = pd.to_datetime(df[col], errors="coerce").dropna()
            years.extend(vals.dt.year.tolist())
    default_year = int(pd.Series(years).mode().iloc[0]) if years else datetime.now().year

    timeline = []
    for col in df.columns:
        if col in FIXED_ALIASES:
            continue
        dt = parse_date_header(col, default_year)
        if dt is not None:
            timeline.append((col, dt))

    timeline.sort(key=lambda x: x[1])
    return timeline

def load_source(uploaded):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded, encoding="utf-8-sig")
    excel = pd.ExcelFile(uploaded)
    selected_sheet = st.sidebar.selectbox("選擇工作表", excel.sheet_names)
    return pd.read_excel(excel, sheet_name=selected_sheet)

def status_done(row):
    progress = str(row.get("組立進度", ""))
    return any(word in progress for word in STATUS_DONE_WORDS)

def calculate_risk(row):
    due = row.get("客戶入庫日")
    if pd.isna(due):
        return "正常"
    if status_done(row):
        return "已完成"
    today = pd.Timestamp.today().normalize()
    due = pd.Timestamp(due).normalize()
    days = (due - today).days
    if days < 0:
        return "逾期"
    if days == 0:
        return "今日應完成"
    if days <= 7:
        return "7日內"
    return "正常"

def prepare_data(df):
    df = normalize_fixed_columns(df.copy())
    for col in ["發料日", "入庫日", "客戶入庫日"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    df["風險"] = df.apply(calculate_risk, axis=1)
    df["距客戶入庫日"] = (df["客戶入庫日"] - pd.Timestamp.today().normalize()).dt.days
    return df

def timeline_long(df, timeline_cols):
    records = []
    for idx, row in df.iterrows():
        for original_col, schedule_date in timeline_cols:
            value = row.get(original_col, "")
            if pd.isna(value):
                value = ""
            code = str(value).strip()
            if not code or code == "-":
                continue
            records.append({
                "列號": idx,
                "製令": row["製令"],
                "客戶": row["客戶"],
                "P/N": row["P/N"],
                "Type": row["Type"],
                "Category": row["Category"],
                "組立地點": row["組立地點"],
                "組立人員": row["組立人員"],
                "組立進度": row["組立進度"],
                "日期": schedule_date,
                "代碼": code,
                "作業": CODE_LABELS.get(code, code),
                "風險": row["風險"]
            })
    return pd.DataFrame(records)

def create_export(df, timeline_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_main = df.copy()
        for c in ["發料日", "入庫日", "客戶入庫日"]:
            export_main[c] = export_main[c].dt.strftime("%Y-%m-%d")
        export_main.to_excel(writer, sheet_name="排程資料", index=False)
        timeline_df.to_excel(writer, sheet_name="每日作業明細", index=False)
        pd.DataFrame(
            [{"代碼": k, "作業說明": v} for k, v in CODE_LABELS.items() if k]
        ).to_excel(writer, sheet_name="代碼說明", index=False)
    output.seek(0)
    return output

st.markdown("""
<style>
.block-container {padding-top: 0.8rem; padding-bottom: 2rem;}
[data-testid="stMetric"] {
    border: 1px solid #d9e2ec;
    border-radius: 12px;
    padding: 12px;
    background: #ffffff;
}
[data-testid="stMetricValue"] {font-size: 1.9rem;}
div[data-testid="stDataFrame"] {border: 1px solid #d9e2ec; border-radius: 10px;}
</style>
""", unsafe_allow_html=True)

st.title("🏭 每日生產排程看板")
st.caption("專用於「固定資料欄＋右側每日日期欄」格式，自動辨識組、T、Q、R、B、W、出等排程代碼。")

with st.sidebar:
    st.header("📂 資料設定")
    uploaded = st.file_uploader("上傳 Excel 或 CSV", type=["xlsx", "xls", "csv"])
    st.markdown("---")
    st.markdown("**必要欄位**")
    st.code("製令、客戶入庫日")
    st.markdown("**建議欄位**")
    st.code("客戶、P/N、Type、Category、組立地點、組立進度、備註、發料日、入庫日")
    st.markdown("**每日欄位範例**")
    st.code("06/28、06/29、07/01\n或 2026/06/28")

if uploaded is None:
    st.info("請上傳原始排程表。ZIP 內附「排程範例.csv」可先測試。")
    st.stop()

try:
    raw = load_source(uploaded)
    data = prepare_data(raw)
    timeline_cols = detect_timeline_columns(data)
    daily = timeline_long(data, timeline_cols)
except Exception as exc:
    st.error(f"資料讀取失敗：{exc}")
    st.stop()

if not timeline_cols:
    st.warning("未找到右側每日日期欄。日期欄名稱請使用 06/28、2026/06/28 或 Excel 日期格式。")

# 篩選器
st.subheader("🔎 排程篩選")
f1, f2, f3, f4, f5 = st.columns(5)
customers = sorted([x for x in data["客戶"].dropna().astype(str).unique() if x])
categories = sorted([x for x in data["Category"].dropna().astype(str).unique() if x])
locations = sorted([x for x in data["組立地點"].dropna().astype(str).unique() if x])
risks = [x for x in RISK_ORDER if x in data["風險"].unique()]

sel_customer = f1.multiselect("客戶", customers)
sel_category = f2.multiselect("Category", categories)
sel_location = f3.multiselect("組立地點", locations)
sel_risk = f4.multiselect("風險", risks)
keyword = f5.text_input("搜尋製令 / P/N / Type")

filtered = data.copy()
if sel_customer:
    filtered = filtered[filtered["客戶"].astype(str).isin(sel_customer)]
if sel_category:
    filtered = filtered[filtered["Category"].astype(str).isin(sel_category)]
if sel_location:
    filtered = filtered[filtered["組立地點"].astype(str).isin(sel_location)]
if sel_risk:
    filtered = filtered[filtered["風險"].isin(sel_risk)]
if keyword:
    key = keyword.lower()
    mask = (
        filtered["製令"].astype(str).str.lower().str.contains(key, na=False)
        | filtered["P/N"].astype(str).str.lower().str.contains(key, na=False)
        | filtered["Type"].astype(str).str.lower().str.contains(key, na=False)
    )
    filtered = filtered[mask]

filtered_daily = daily[daily["列號"].isin(filtered.index)] if not daily.empty else daily

today = pd.Timestamp.today().normalize()
today_jobs = filtered_daily[filtered_daily["日期"] == today] if not filtered_daily.empty else filtered_daily
next7_jobs = filtered_daily[
    (filtered_daily["日期"] >= today) &
    (filtered_daily["日期"] <= today + pd.Timedelta(days=7))
] if not filtered_daily.empty else filtered_daily

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("製令筆數", f"{len(filtered):,}")
m2.metric("今日作業", f"{len(today_jobs):,}")
m3.metric("未來7日作業", f"{len(next7_jobs):,}")
m4.metric("逾期製令", f"{int((filtered['風險'] == '逾期').sum()):,}")
m5.metric("7日內交期", f"{int(filtered['風險'].isin(['今日應完成','7日內']).sum()):,}")

tabs = st.tabs(["📊 管理看板", "📅 每日排程", "📋 製令總表", "🚨 異常追蹤", "📖 代碼說明"])

with tabs[0]:
    c1, c2 = st.columns(2)

    with c1:
        risk_summary = (
            filtered["風險"].value_counts()
            .reindex(RISK_ORDER, fill_value=0)
            .reset_index()
        )
        risk_summary.columns = ["風險", "製令數"]
        fig = px.bar(
            risk_summary,
            x="風險",
            y="製令數",
            text_auto=True,
            title="交期風險分布",
            category_orders={"風險": RISK_ORDER}
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        category_summary = (
            filtered.groupby("Category", dropna=False)
            .size().reset_index(name="製令數")
            .sort_values("製令數", ascending=False)
        )
        fig = px.pie(
            category_summary,
            names="Category",
            values="製令數",
            hole=0.45,
            title="Category 製令占比"
        )
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        if filtered_daily.empty:
            st.info("無每日作業資料。")
        else:
            daily_count = (
                filtered_daily.groupby(["日期", "作業"])
                .size().reset_index(name="作業數")
            )
            fig = px.bar(
                daily_count,
                x="日期",
                y="作業數",
                color="作業",
                barmode="stack",
                title="每日作業負荷"
            )
            st.plotly_chart(fig, use_container_width=True)

    with c4:
        customer_summary = (
            filtered.groupby("客戶", dropna=False)
            .size().reset_index(name="製令數")
            .sort_values("製令數", ascending=False)
            .head(10)
        )
        fig = px.bar(
            customer_summary,
            x="製令數",
            y="客戶",
            orientation="h",
            text_auto=True,
            title="客戶製令 Top 10"
        )
        fig.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, use_container_width=True)

with tabs[1]:
    if filtered_daily.empty:
        st.warning("沒有可顯示的每日排程代碼。")
    else:
        date_min = filtered_daily["日期"].min().date()
        date_max = filtered_daily["日期"].max().date()
        selected_dates = st.date_input(
            "顯示日期範圍",
            value=(max(date_min, date.today() - timedelta(days=7)), date_max),
            min_value=date_min,
            max_value=date_max
        )

        date_filtered = filtered_daily.copy()
        if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
            start, end = selected_dates
            date_filtered = date_filtered[
                (date_filtered["日期"].dt.date >= start) &
                (date_filtered["日期"].dt.date <= end)
            ]

        pivot = date_filtered.pivot_table(
            index=["製令", "客戶", "P/N", "Type", "組立進度"],
            columns="日期",
            values="代碼",
            aggfunc=lambda x: "/".join(dict.fromkeys(str(v) for v in x)),
            fill_value=""
        ).reset_index()
        pivot.columns = [
            c.strftime("%m/%d") if isinstance(c, pd.Timestamp) else c
            for c in pivot.columns
        ]
        st.dataframe(pivot, use_container_width=True, hide_index=True, height=520)

        st.markdown("#### 每日作業明細")
        detail = date_filtered[
            ["日期", "代碼", "作業", "製令", "客戶", "P/N", "Type", "組立地點", "組立進度", "風險"]
        ].sort_values(["日期", "製令"])
        detail["日期"] = detail["日期"].dt.strftime("%Y-%m-%d")
        st.dataframe(detail, use_container_width=True, hide_index=True)

with tabs[2]:
    show_cols = [
        "風險", "製令", "客戶", "P/N", "Type", "Category",
        "組立地點", "組立人員", "組立進度", "備註",
        "發料日", "入庫日", "客戶入庫日", "距客戶入庫日"
    ]
    display = filtered[show_cols].copy()
    for col in ["發料日", "入庫日", "客戶入庫日"]:
        display[col] = display[col].dt.strftime("%Y-%m-%d")
    st.dataframe(display, use_container_width=True, hide_index=True, height=570)

with tabs[3]:
    abnormal = filtered[
        filtered["風險"].isin(["逾期", "今日應完成", "7日內"])
    ].copy().sort_values(["客戶入庫日", "製令"])
    abnormal_cols = [
        "風險", "製令", "客戶", "P/N", "Type", "Category",
        "組立進度", "備註", "入庫日", "客戶入庫日", "距客戶入庫日"
    ]
    for col in ["入庫日", "客戶入庫日"]:
        abnormal[col] = abnormal[col].dt.strftime("%Y-%m-%d")
    st.dataframe(abnormal[abnormal_cols], use_container_width=True, hide_index=True)

with tabs[4]:
    explanation = pd.DataFrame(
        [{"排程代碼": code, "作業說明": label} for code, label in CODE_LABELS.items() if code]
    )
    st.dataframe(explanation, use_container_width=True, hide_index=True)
    st.info("若公司有其他代碼，可在 app.py 的 CODE_LABELS 內自行增加。")

st.divider()
export_bytes = create_export(filtered, filtered_daily)
st.download_button(
    "⬇️ 下載篩選後排程與每日明細",
    data=export_bytes,
    file_name=f"每日生產排程看板匯出_{datetime.now():%Y%m%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.caption(f"已辨識每日日期欄：{len(timeline_cols)} 欄")
