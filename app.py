import streamlit as st
from typing import Dict

st.set_page_config(
    page_title="生產排程反推看板",
    page_icon="🗓️",
    layout="wide",
)

st.title("🗓️ 生產排程反推看板")
st.success("目前程式版本：2026-06-30-v2")

st.sidebar.header("⚙️ 排程參數")

st.sidebar.subheader("組立地點緩衝工作日")
location_buffer: Dict[str, int] = {
    "竹東": st.sidebar.number_input("竹東", min_value=0, value=2, step=1),
    "模冠": st.sidebar.number_input("模冠", min_value=0, value=5, step=1),
    "御弘": st.sidebar.number_input("御弘", min_value=0, value=7, step=1),
    "宏田": st.sidebar.number_input("宏田", min_value=0, value=3, step=1),
}

st.sidebar.subheader("Category 標準工期")
default_category_days: Dict[str, int] = {
    "EFEM": st.sidebar.number_input("EFEM", min_value=0, value=10, step=1),
    "sort": st.sidebar.number_input("sort", min_value=0, value=10, step=1),
    "骨包": st.sidebar.number_input("骨包", min_value=0, value=10, step=1),
    "BWS": st.sidebar.number_input("BWS", min_value=0, value=10, step=1),
    "NTB": st.sidebar.number_input("NTB", min_value=0, value=10, step=1),
    "other": st.sidebar.number_input("other", min_value=0, value=10, step=1),
}

st.write("請上傳 Excel 檔案後進行排程分析。")
