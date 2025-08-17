# 試劑批號確認系統

## 🎯 系統概述

這是一個專為檢驗科設計的試劑批號確認系統，基於 Web 技術開發，使用 Flask 後端和 SQLite 資料庫。系統專注於確認試劑批號是否使用過，提供簡潔易用的介面，讓檢驗科人員能夠快速查詢和管理試劑批號資訊。

## ✨ 系統特色

- 🌐 **Web 介面**：基於瀏覽器的操作介面，無需安裝額外軟體
- 🚀 **輕量快速**：使用 Flask 框架，啟動速度快，資源消耗低
- 💾 **本地資料庫**：使用 SQLite 資料庫，無需額外資料庫伺服器
- 📱 **響應式設計**：支援各種螢幕尺寸，手機、平板、電腦都能使用
- 🎨 **現代化 UI**：使用 Bootstrap 5 和 Bootstrap Icons，介面美觀易用
- 🔍 **智能查詢**：支援多種篩選和搜尋功能
- ⚠️ **過期提醒**：自動識別過期和即將過期的批號

## 🏗️ 系統架構

### 前端技術
- **HTML5**：語義化標籤，提升可讀性和可維護性
- **CSS3**：現代化樣式，包含動畫效果和響應式設計
- **JavaScript (ES6+)**：前端邏輯處理，與後端 API 互動
- **Bootstrap 5**：UI 框架，提供美觀的元件和響應式網格系統
- **Bootstrap Icons**：豐富的圖示庫，提升使用者體驗

### 後端技術
- **Python 3.7+**：主要程式語言
- **Flask**：輕量級 Web 框架，適合快速開發
- **SQLite**：嵌入式資料庫，無需額外設定
- **Werkzeug**：WSGI 工具庫，Flask 的依賴套件

## 📋 功能模組

### 1. 新增批號 (`/`)
- 輸入檢驗項目名稱
- 輸入試劑批號（自動檢查重複）
- 設定有效期限（日期選擇器）
- 輸入數量（可選）
- 添加備註（可選）
- 顯示最近新增的批號列表

### 2. 查詢批號 (`/search`)
- 輸入批號進行精確查詢
- 顯示查詢結果（已使用/未使用）
- 過期批號特別標示
- 顯示詳細的儲存資訊
- 統計資訊（總數、有效、過期）

### 3. 批號列表 (`/list`)
- 顯示所有已儲存的批號
- 多種篩選條件（項目、狀態、關鍵字）
- 分頁顯示，提升效能
- 詳細資訊模態框
- 統計卡片（總數、有效、即將過期、已過期）

## 🗄️ 資料庫設計

### 資料表結構
```sql
CREATE TABLE used_reagents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    item TEXT NOT NULL,                    -- 檢驗項目名稱
    lot_number TEXT UNIQUE NOT NULL,       -- 試劑批號（唯一）
    expiration_date TEXT NOT NULL,         -- 有效期限
    quantity INTEGER DEFAULT 1,            -- 數量
    stored_date TEXT NOT NULL             -- 儲存時間
);
```

### 索引設計
- `idx_lot_number`：批號查詢索引
- `idx_item`：項目篩選索引
- `idx_stored_date`：時間排序索引

## 🚀 安裝與部署

### 系統需求
- **作業系統**：Windows 10/11、macOS 10.14+、Ubuntu 18.04+
- **Python**：3.7 或更新版本
- **記憶體**：至少 2GB RAM
- **硬碟空間**：至少 50MB 可用空間

### 安裝步驟

#### 1. 安裝 Python
確保系統已安裝 Python 3.7 或更新版本：
```bash
python --version
# 或
python3 --version
```

#### 2. 下載專案
```bash
git clone <repository-url>
cd reagent-batch-checker
```

#### 3. 建立虛擬環境（建議）
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

#### 4. 安裝依賴套件
```bash
pip install -r requirements.txt
```

#### 5. 啟動系統
```bash
python app.py
```

#### 6. 開啟瀏覽器
在瀏覽器中開啟：`http://localhost:5000`

### 快速啟動腳本

#### Windows
```batch
@echo off
echo 正在啟動試劑批號確認系統...
python app.py
pause
```

#### macOS/Linux
```bash
#!/bin/bash
echo "正在啟動試劑批號確認系統..."
python3 app.py
```

## 📖 使用說明

### 新增批號
1. 在首頁填寫試劑資訊
2. 點擊「儲存批號」按鈕
3. 系統會自動檢查批號是否重複
4. 儲存成功後會顯示確認訊息

### 查詢批號
1. 點擊「查詢批號」頁面
2. 輸入要查詢的批號
3. 點擊「查詢批號」按鈕
4. 查看查詢結果和詳細資訊

### 查看列表
1. 點擊「批號列表」頁面
2. 使用篩選條件縮小範圍
3. 使用關鍵字搜尋特定批號
4. 點擊「詳細」按鈕查看完整資訊

## 🔧 設定與自訂

### 修改資料庫路徑
在 `app.py` 中修改：
```python
DATABASE = 'your_custom_path/reagents.db'
```

### 修改伺服器設定
在 `app.py` 中修改：
```python
app.run(debug=False, host='0.0.0.0', port=8080)
```

### 自訂樣式
修改 `templates/base.html` 中的 CSS 變數：
```css
:root {
    --primary-color: #your-color;
    --secondary-color: #your-color;
    /* 其他顏色變數 */
}
```

## 📊 資料匯出與備份

### 資料庫備份
```bash
# 複製 SQLite 資料庫檔案
cp reagents.db reagents_backup_$(date +%Y%m%d).db
```

### 資料匯出
系統提供 API 端點用於資料匯出：
- `GET /api/get_all_reagents`：獲取所有試劑資料（JSON 格式）

## 🚨 故障排除

### 常見問題

#### 1. 無法啟動系統
**症狀**：執行 `python app.py` 時出現錯誤

**解決方法**：
- 確認 Python 版本是否為 3.7+
- 確認已安裝 Flask：`pip install Flask`
- 檢查防火牆設定，確保 5000 埠未被阻擋

#### 2. 資料庫錯誤
**症狀**：新增或查詢時出現資料庫錯誤

**解決方法**：
- 確認 `reagents.db` 檔案存在且可寫入
- 檢查檔案權限設定
- 刪除損壞的資料庫檔案，讓系統重新建立

#### 3. 頁面無法載入
**症狀**：瀏覽器顯示錯誤或空白頁面

**解決方法**：
- 確認 Flask 伺服器正在運行
- 檢查瀏覽器控制台是否有 JavaScript 錯誤
- 清除瀏覽器快取

#### 4. 效能問題
**症狀**：頁面載入緩慢或操作遲鈍

**解決方法**：
- 減少每頁顯示的項目數量
- 使用篩選條件縮小資料範圍
- 定期清理舊的資料記錄

### 錯誤日誌
系統會在控制台輸出詳細的錯誤訊息，包括：
- 資料庫操作狀態
- API 請求處理結果
- 錯誤詳細資訊

## 🔮 未來擴展

### 建議功能
1. **批次匯入**：支援 Excel/CSV 檔案批次匯入
2. **條碼掃描**：整合條碼掃描器，快速輸入批號
3. **報表功能**：生成統計報表和圖表
4. **使用者管理**：添加登入和權限控制
5. **備份同步**：自動備份和雲端同步
6. **行動應用**：開發手機 App 版本

### 技術升級
1. **資料庫升級**：從 SQLite 升級到 PostgreSQL/MySQL
2. **快取機制**：添加 Redis 快取提升效能
3. **API 版本化**：支援多版本 API
4. **微服務架構**：拆分為多個微服務

## 📄 授權條款

本系統採用 MIT 授權條款，詳見 LICENSE 檔案。

## 🤝 貢獻指南

歡迎提交 Issue 和 Pull Request 來改善系統：

1. Fork 專案
2. 建立功能分支
3. 提交變更
4. 發起 Pull Request

## 📞 聯絡資訊

如有問題或建議，請聯絡開發團隊：

- **專案維護者**：[您的姓名]
- **電子郵件**：[您的郵箱]
- **專案網址**：[專案 GitHub 頁面]

## 📝 更新日誌

### 版本 1.0.0 (2024-12-XX)
- ✨ 初始版本發布
- 🎯 實現核心功能：新增批號、查詢批號、批號列表
- 🎨 現代化 Web 介面設計
- 📱 響應式設計，支援多種裝置
- 🗄️ SQLite 資料庫整合
- 🔍 智能篩選和搜尋功能
- ⚠️ 過期批號提醒功能

---

**感謝使用試劑批號確認系統！** 🎉

如有任何問題或需要協助，請隨時聯絡我們。
