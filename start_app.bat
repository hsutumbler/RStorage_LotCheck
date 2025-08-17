@echo off
chcp 65001 >nul
title 試劑批號確認系統 - 啟動器

echo.
echo ========================================
echo    試劑批號確認系統 - 啟動器
echo ========================================
echo.
echo 正在啟動系統，請稍候...
echo.

:: 設定環境變數隱藏 Flask 警告
set FLASK_ENV=development
set WERKZEUG_DISABLE_WARNINGS=1
set PYTHONWARNINGS=ignore

:: 檢查 Python 是否已安裝
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [錯誤] 未找到 Python，請先安裝 Python 3.7 或更新版本
    echo.
    echo 請前往 https://www.python.org/downloads/ 下載並安裝 Python
    echo 安裝時請勾選 "Add Python to PATH" 選項
    echo.
    pause
    exit /b 1
)

:: 檢查必要檔案是否存在
if not exist "app.py" (
    echo [錯誤] 找不到 app.py 檔案
    echo 請確保此批次檔案與 app.py 在同一目錄下
    echo.
    pause
    exit /b 1
)

if not exist "requirements.txt" (
    echo [警告] 找不到 requirements.txt 檔案
    echo 系統可能無法正常運作
    echo.
)

:: 檢查並安裝依賴套件
echo [資訊] 檢查並安裝必要套件...
pip install -r requirements.txt >nul 2>&1
if %errorlevel% neq 0 (
    echo [警告] 套件安裝失敗，嘗試繼續啟動...
    echo.
)

:: 啟動應用程式
echo [資訊] 啟動試劑批號確認系統...
echo.
echo 系統啟動成功後，請在瀏覽器中開啟：
echo http://localhost:5000
echo.
echo 按 Ctrl+C 可停止系統
echo.

python app.py

:: 如果程式異常結束，暫停讓使用者看到錯誤訊息
if %errorlevel% neq 0 (
    echo.
    echo [錯誤] 程式異常結束，錯誤代碼：%errorlevel%
    echo.
    pause
)
