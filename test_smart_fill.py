#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能表單填充功能測試腳本
"""

import requests
import json
import time

# API基礎URL
BASE_URL = "http://127.0.0.1:5000"

def test_api_endpoint(endpoint, params=None):
    """測試API端點"""
    try:
        url = f"{BASE_URL}{endpoint}"
        response = requests.get(url, params=params, timeout=5)
        
        print(f"📡 測試: {endpoint}")
        print(f"   參數: {params}")
        print(f"   狀態: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print(f"   結果: {json.dumps(data, ensure_ascii=False, indent=2)}")
            print("   ✅ 成功")
        else:
            print(f"   ❌ 錯誤: {response.text}")
        
        print("-" * 50)
        return response.status_code == 200
        
    except Exception as e:
        print(f"   ❌ 異常: {str(e)}")
        print("-" * 50)
        return False

def main():
    """主測試函數"""
    print("🧪 智能表單填充功能測試")
    print("=" * 50)
    
    # 等待Flask應用啟動
    print("⏳ 等待Flask應用啟動...")
    time.sleep(2)
    
    # 測試試劑建議API
    print("\n1️⃣ 測試試劑名稱建議API")
    test_api_endpoint("/api/reagent-suggestions", {"q": "AF"})
    
    # 測試試劑供應商和單位API
    print("\n2️⃣ 測試試劑供應商和單位API")
    test_api_endpoint("/api/reagent-supplier", {"name": "AFP"})
    test_api_endpoint("/api/reagent-supplier", {"name": "不存在的試劑"})
    
    # 測試批號效期API
    print("\n3️⃣ 測試批號效期API")
    test_api_endpoint("/api/batch-expiry", {"name": "AFP", "batch": "AFP001"})
    test_api_endpoint("/api/batch-expiry", {"name": "AFP", "batch": "新批號999"})
    
    # 測試試劑批號列表API
    print("\n4️⃣ 測試試劑批號列表API")
    test_api_endpoint("/api/reagent-batches", {"name": "AFP"})
    test_api_endpoint("/api/reagent-batches", {"name": "不存在的試劑"})
    
    print("\n🎉 測試完成！")
    print("💡 請檢查上述結果，確認所有API都正常運作")

if __name__ == "__main__":
    main()
