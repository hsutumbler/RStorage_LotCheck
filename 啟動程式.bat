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
echo.

REM 檢查必要套件是否已安裝
python -c "import flask, flask_sqlalchemy, reportlab" >nul 2>&1
if errorlevel 1 (
    echo [安裝] 正在安裝必要套件...
    pip install -r requirements.txt
    echo.
)

echo [✓] 套件檢查完成
echo.

echo [啟動] 正在啟動Flask應用...
echo [等待] 請稍候3秒讓服務啟動...
echo.

REM 啟動Flask應用到背景
start /B python simple_app.py

REM 等待服務啟動
timeout /t 3 /nobreak >nul

echo [瀏覽器] 自動開啟瀏覽器...
start http://127.0.0.1:5000

echo.
echo ========================================
echo  ✓ 系統啟動完成！
echo  ✓ 瀏覽器已自動開啟
echo  ✓ 網址：http://127.0.0.1:5000
echo ========================================
echo.
echo [提示] 系統正在背景執行中
echo [提示] 關閉此視窗不會影響系統運行
echo [提示] 要停止系統，請使用 停止程式.bat
echo.

echo 按任意鍵關閉此視窗...
pause >nul
