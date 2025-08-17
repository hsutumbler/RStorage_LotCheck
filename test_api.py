#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試試劑批號確認系統的 API 功能
"""

import requests
import json

def test_api():
    """測試 API 功能"""
    base_url = "http://localhost:5000"
    
    print("=" * 50)
    print("試劑批號確認系統 - API 測試")
    print("=" * 50)
    
    try:
        # 測試 1：獲取所有試劑
        print("\n1. 測試獲取所有試劑...")
        response = requests.get(f"{base_url}/api/get_all_reagents")
        print(f"狀態碼: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print(f"回應: {json.dumps(data, indent=2, ensure_ascii=False)}")
            
            if data.get('success'):
                reagents = data.get('data', [])
                print(f"✓ 成功獲取 {len(reagents)} 筆試劑記錄")
                
                for i, reagent in enumerate(reagents, 1):
                    print(f"  記錄 {i}:")
                    print(f"    - 項目: {reagent.get('item')}")
                    print(f"    - 批號: {reagent.get('lot_number')}")
                    print(f"    - 期限: {reagent.get('expiration_date')}")
                    print(f"    - 儲存時間: {reagent.get('stored_date')}")
                    print(f"    - 是否過期: {reagent.get('is_expired')}")
            else:
                print(f"✗ API 回應失敗: {data.get('message')}")
        else:
            print(f"✗ HTTP 錯誤: {response.status_code}")
            print(f"回應內容: {response.text}")
            
    except requests.exceptions.ConnectionError:
        print("✗ 無法連接到伺服器")
        print("請確保 Flask 應用程式正在運行")
        print("在瀏覽器中開啟: http://localhost:5000")
        
    except Exception as e:
        print(f"✗ 測試過程中發生錯誤: {e}")
    
    print("\n" + "=" * 50)
    print("測試完成")

if __name__ == "__main__":
    test_api()
