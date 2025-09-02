@echo off
title 停止試劑入庫管理系統
color 0C
chcp 65001 >nul

echo.
echo  ========================================
echo  停止試劑入庫管理系統
echo  ========================================
echo.

echo [停止] 正在停止Flask應用...
echo.

REM 停止所有Python進程
taskkill /f /im python.exe >nul 2>&1
if errorlevel 1 (
    echo [警告] 沒有找到正在運行的Python進程
) else (
    echo [✓] 已停止所有Python進程
)

echo.
echo [檢查] 檢查端口5000是否已釋放...
netstat -ano | findstr :5000 >nul 2>&1
if errorlevel 1 (
    echo [✓] 端口5000已釋放
) else (
    echo [警告] 端口5000仍在使用中
    echo 請手動關閉相關程式
)

echo.
echo ========================================
echo  ✓ 服務停止完成！
echo ========================================
echo.
echo 按任意鍵關閉此視窗...
pause >nul
