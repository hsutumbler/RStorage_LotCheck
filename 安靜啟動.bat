@echo off
chcp 65001 >nul
title 試劑批號確認系統 - 安靜啟動

echo.
echo ========================================
echo    試劑批號確認系統 - 安靜啟動
echo ========================================
echo.
echo 正在啟動系統，請稍候...
echo.

:: 設定環境變數隱藏所有警告
set FLASK_ENV=development
set WERKZEUG_DISABLE_WARNINGS=1
set PYTHONWARNINGS=ignore
set PYTHONUNBUFFERED=1

:: 啟動應用程式（重定向錯誤輸出到空設備）
echo [資訊] 啟動試劑批號確認系統...
echo.
echo 系統啟動成功後，請在瀏覽器中開啟：
echo http://localhost:5000
echo.
echo 按 Ctrl+C 可停止系統
echo.

python app.py 2>nul

pause
