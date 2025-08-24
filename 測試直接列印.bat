@echo off
chcp 65001 >nul
title 測試直接列印功能

echo ========================================
echo 試劑入庫管理系統 - 直接列印測試版
echo ========================================
echo.
echo 修改內容：
echo 1. 新增入庫後自動列印標籤
echo 2. 使用者完全看不到PDF檢視器
echo 3. 直接發送到預設印表機
echo 4. 按鈕文字更新為「儲存並自動列印標籤」
echo 5. 移除成功對話視窗，靜默處理
echo.
echo 正在啟動系統...
echo.

python simple_app.py

pause
