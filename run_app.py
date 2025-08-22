"""
試劑入庫管理系統啟動腳本
"""
import os
import sys
import subprocess
import webbrowser
import time
import threading
import signal
import socket

def find_free_port():
    """找到一個可用的端口"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]

def open_browser(port):
    """延遲3秒後打開瀏覽器"""
    time.sleep(3)
    webbrowser.open(f'http://localhost:{port}')

def main():
    # 獲取應用程序所在的目錄
    if getattr(sys, 'frozen', False):
        # 如果是打包後的應用程序
        application_path = os.path.dirname(sys.executable)
    else:
        # 如果是開發環境
        application_path = os.path.dirname(os.path.abspath(__file__))
    
    # 切換到應用程序目錄
    os.chdir(application_path)
    
    # 找到一個可用的端口
    port = find_free_port()
    
    # 創建一個線程來打開瀏覽器
    browser_thread = threading.Thread(target=open_browser, args=(port,))
    browser_thread.daemon = True
    browser_thread.start()
    
    # 啟動 Flask 應用程序
    print(f"試劑入庫管理系統啟動中...")
    print(f"應用程序將在瀏覽器中自動打開，請稍候...")
    print(f"如果瀏覽器沒有自動打開，請手動訪問: http://localhost:{port}")
    print(f"按 Ctrl+C 可停止程序")
    
    # 導入 app.py 中的 app 對象
    sys.path.insert(0, application_path)
    from app import app
    
    # 啟動 Flask 應用程序
    app.run(host='0.0.0.0', port=port, debug=False)

if __name__ == "__main__":
    main()
