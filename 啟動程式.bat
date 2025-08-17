@echo off
chcp 65001 >nul
title 試劑批號確認系統

echo.
echo ========================================
echo    試劑批號確認系統
echo ========================================
echo.

:: 啟動應用程式
echo 正在啟動系統...
echo.
echo 啟動成功後，請在瀏覽器中開啟：
echo http://localhost:5000
echo.
echo 按 Ctrl+C 可停止系統
echo.

python app.py

pause
