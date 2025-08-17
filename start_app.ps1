# 試劑批號確認系統 - PowerShell 啟動器
# 執行政策設定：Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

param(
    [switch]$Install,
    [switch]$Help
)

# 設定控制台編碼為 UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8

# 顯示說明
if ($Help) {
    Write-Host @"
試劑批號確認系統 - PowerShell 啟動器

使用方法：
    .\start_app.ps1              # 啟動系統
    .\start_app.ps1 -Install     # 安裝依賴套件
    .\start_app.ps1 -Help        # 顯示此說明

參數說明：
    -Install   安裝或更新必要的 Python 套件
    -Help      顯示此說明訊息

注意事項：
    1. 需要安裝 Python 3.7 或更新版本
    2. 首次使用建議先執行 -Install 參數
    3. 系統啟動後請在瀏覽器中開啟 http://localhost:5000
"@ -ForegroundColor Cyan
    exit 0
}

# 顯示標題
Write-Host @"

========================================
    試劑批號確認系統 - 啟動器
========================================

"@ -ForegroundColor Green

# 檢查 Python 是否已安裝
Write-Host "正在檢查 Python 環境..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Python 已安裝: $pythonVersion" -ForegroundColor Green
    } else {
        throw "Python 未正確安裝"
    }
} catch {
    Write-Host "✗ 錯誤：未找到 Python" -ForegroundColor Red
    Write-Host @"

請先安裝 Python 3.7 或更新版本：
1. 前往 https://www.python.org/downloads/
2. 下載最新版本的 Python
3. 安裝時請勾選 "Add Python to PATH" 選項
4. 重新啟動 PowerShell 或命令提示字元

"@ -ForegroundColor Yellow
    Read-Host "按 Enter 鍵退出"
    exit 1
}

# 檢查必要檔案
Write-Host "正在檢查必要檔案..." -ForegroundColor Yellow
$requiredFiles = @("app.py", "requirements.txt")
$missingFiles = @()

foreach ($file in $requiredFiles) {
    if (Test-Path $file) {
        Write-Host "✓ 找到 $file" -ForegroundColor Green
    } else {
        Write-Host "✗ 缺少 $file" -ForegroundColor Red
        $missingFiles += $file
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host @"

錯誤：缺少必要檔案，請確保以下檔案存在：
$($missingFiles -join ", ")

請確保此腳本與 app.py 在同一目錄下。
"@ -ForegroundColor Red
    Read-Host "按 Enter 鍵退出"
    exit 1
}

# 安裝依賴套件
if ($Install) {
    Write-Host "正在安裝/更新依賴套件..." -ForegroundColor Yellow
    try {
        $result = pip install -r requirements.txt 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ 套件安裝成功" -ForegroundColor Green
        } else {
            Write-Host "⚠ 套件安裝出現警告，但將繼續啟動" -ForegroundColor Yellow
            Write-Host $result -ForegroundColor Yellow
        }
    } catch {
        Write-Host "⚠ 套件安裝失敗，嘗試繼續啟動" -ForegroundColor Yellow
    }
    Write-Host ""
}

# 檢查套件是否已安裝
Write-Host "正在檢查套件狀態..." -ForegroundColor Yellow
try {
    $result = python -c "import flask; print('Flask 版本:', flask.__version__)" 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Flask 已安裝: $result" -ForegroundColor Green
    } else {
        Write-Host "⚠ Flask 未安裝，正在自動安裝..." -ForegroundColor Yellow
        pip install -r requirements.txt > $null 2>&1
    }
} catch {
    Write-Host "⚠ 無法檢查套件狀態" -ForegroundColor Yellow
}

Write-Host ""

# 啟動應用程式
Write-Host "正在啟動試劑批號確認系統..." -ForegroundColor Green
Write-Host @"

系統啟動成功後，請在瀏覽器中開啟：
    http://localhost:5000

操作說明：
    - 按 Ctrl+C 可停止系統
    - 系統會自動重新載入程式碼變更
    - 如遇問題請檢查終端機錯誤訊息

"@ -ForegroundColor Cyan

try {
    # 啟動 Flask 應用程式
    python app.py
} catch {
    Write-Host @"

[錯誤] 程式異常結束
錯誤訊息：$($_.Exception.Message)
錯誤代碼：$LASTEXITCODE

"@ -ForegroundColor Red
    
    # 提供故障排除建議
    Write-Host @"
故障排除建議：
1. 確認 Python 版本是否為 3.7 或更新版本
2. 執行 .\start_app.ps1 -Install 重新安裝套件
3. 檢查防火牆設定是否阻擋了端口 5000
4. 確認沒有其他程式佔用端口 5000

"@ -ForegroundColor Yellow
    
    Read-Host "按 Enter 鍵退出"
}
