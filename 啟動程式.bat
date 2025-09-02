@echo off
title 試劑入庫管理系統
color 0A
chcp 65001 >nul

echo.
echo  ========================================
echo  試劑入庫管理系統
echo  ========================================
echo.

REM 檢查Python是否已安裝
python --version >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 未找到Python，請先安裝Python 3.7+
    echo.
    pause
    exit /b 1
)

echo [✓] Python環境檢查完成

REM 檢查必要套件是否已安裝
python -c "import flask, flask_sqlalchemy, reportlab" >nul 2>&1
if errorlevel 1 (
    echo [安裝] 正在安裝必要套件...
    pip install -r requirements.txt
)

echo [✓] 套件檢查完成
echo [啟動] 正在啟動系統...
echo.

REM 直接在當前視窗執行Python應用
python simple_app.py
