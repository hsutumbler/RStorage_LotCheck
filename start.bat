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
echo 請在瀏覽器中開啟 http://localhost:5000
echo.
echo 按 Ctrl+C 可停止程式
echo.

python app.py

pause
