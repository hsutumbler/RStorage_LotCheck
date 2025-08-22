@echo off
echo 試劑入庫管理系統啟動中...
echo.
echo 正在檢查Python環境...
python --version >nul 2>&1
if errorlevel 1 (
    echo 錯誤：未找到Python，請先安裝Python 3.7或更新版本
    pause
    exit /b 1
)

echo Python環境檢查完成
echo.
echo 正在安裝必要套件...
pip install -r requirements.txt

echo.
echo 套件安裝完成
echo.
echo 正在啟動系統...
echo 系統將自動開啟瀏覽器
echo.
echo 按 Ctrl+C 可停止程式
echo.

REM 啟動Flask應用並自動開啟瀏覽器
start /B python run_app.py

echo.
echo 系統已啟動！
echo 如果瀏覽器沒有自動開啟，請稍等片刻或查看控制台輸出獲取正確的URL
echo.
echo 系統正在背景執行中...
echo 要停止系統，請執行 stop_service.bat 或按任意鍵關閉此視窗
echo.

pause