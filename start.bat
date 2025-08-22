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
echo 系統將自動開啟瀏覽器到 http://127.0.0.1:5000
echo.
echo 按 Ctrl+C 可停止程式
echo.

REM 啟動Flask應用並等待3秒讓服務啟動
start /B python app.py
timeout /t 3 /nobreak >nul

REM 自動開啟預設瀏覽器到指定網址
start http://127.0.0.1:5000

echo.
echo 瀏覽器已自動開啟！
echo 如果瀏覽器沒有自動開啟，請手動前往：http://127.0.0.1:5000
echo.
echo 系統正在背景執行中...
echo 要停止系統，請關閉此視窗或按任意鍵
echo.

pause