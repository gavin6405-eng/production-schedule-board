@echo off
chcp 65001 >nul
title 安裝每日生產排程看板
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [錯誤] 找不到 Python。
    echo 請安裝 Python 3.11 以上版本，並勾選 Add Python to PATH。
    pause
    exit /b
)

python -m venv .venv
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo 安裝完成，之後直接雙擊「一鍵啟動看板.bat」。
pause
