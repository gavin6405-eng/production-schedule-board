@echo off
chcp 65001 >nul
title 每日生產排程看板
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [錯誤] 找不到 Python。
    echo 請先安裝 Python 3.11 以上版本，安裝時勾選 Add Python to PATH。
    pause
    exit /b
)

if not exist ".venv" (
    echo 第一次啟動，正在建立執行環境...
    python -m venv .venv
)

call ".venv\Scripts\activate.bat"

python -c "import streamlit, pandas, plotly, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo 正在安裝必要套件，第一次約需數分鐘...
    python -m pip install --upgrade pip
    pip install -r requirements.txt
)

echo 正在開啟每日生產排程看板...
start "" http://localhost:8501
streamlit run app.py

pause
