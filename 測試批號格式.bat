@echo off
chcp 65001 >nul
title 測試批號格式修正

echo ========================================
echo 試劑入庫管理系統 - 批號格式修正測試
echo ========================================
echo.
echo 修正內容：
echo 1. 批號CEA002 修正為 批號：CEA002
echo 2. 所有列印標籤都已加上冒號
echo 3. 包含新增入庫和補印標籤功能
echo.
echo 測試方法：
echo 1. 新增一筆入庫記錄
echo 2. 檢查列印出來的標籤格式
echo 3. 確認批號前面有冒號
echo.
echo 正在啟動系統...
echo.

python simple_app.py

pause
