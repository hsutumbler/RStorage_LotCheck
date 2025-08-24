# 試劑入庫管理系統設定檔案

# 資料庫設定
DATABASE_URI = 'sqlite:///simple_inventory.db'

# 伺服器設定
HOST = '0.0.0.0'  # 監聽所有網路介面
PORT = 5000        # 服務埠號
DEBUG = True       # 除錯模式

# 標籤設定
LABEL_SETTINGS = {
    'page_size': 'A4',           # 頁面大小
    'margin_left': 20,           # 左邊距 (mm)
    'margin_top': 270,           # 上邊距 (mm)
    'font_size_title': 16,       # 標題字體大小
    'font_size_content': 14,     # 內容字體大小
    'line_spacing': 20,          # 行距 (mm)
}

# 系統設定
SYSTEM_SETTINGS = {
    'max_quantity': 1000,        # 最大入庫數量
    'auto_backup': True,         # 自動備份
    'backup_interval': 24,       # 備份間隔 (小時)
    'session_timeout': 3600,     # 會話超時 (秒)
}

# 標籤內容設定
LABEL_CONTENT = {
    'show_old_batch': True,      # 顯示舊批號
    'show_notes': True,          # 顯示備註
    'show_date': True,           # 顯示日期
    'show_quantity': True,       # 顯示數量
}

# 搜尋設定
SEARCH_SETTINGS = {
    'min_query_length': 1,       # 最小搜尋字元數
    'max_results': 100,          # 最大搜尋結果數
    'search_delay': 300,         # 搜尋延遲 (毫秒)
}

# 列印設定
PRINT_SETTINGS = {
    'default_printer': '',       # 預設印表機名稱
    'print_dialog': True,        # 顯示列印對話框
    'auto_print': False,         # 自動列印
    'print_format': 'PDF',       # 列印格式 (PDF/PRINT)
}
