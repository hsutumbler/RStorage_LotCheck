#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
檢查資料庫狀態的腳本
"""

import sqlite3
import os

def check_database():
    """檢查資料庫狀態"""
    db_path = 'reagents.db'
    
    if not os.path.exists(db_path):
        print(f"❌ 資料庫檔案 {db_path} 不存在")
        return
    
    print(f"✅ 資料庫檔案 {db_path} 存在")
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 檢查表格
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = cursor.fetchall()
        
        print(f"\n📋 資料庫中的表格：")
        for table in tables:
            print(f"  - {table[0]}")
        
        # 檢查 used_reagents 表格的資料
        if ('used_reagents',) in tables:
            print(f"\n🔍 檢查 used_reagents 表格：")
            
            # 檢查表格結構
            cursor.execute("PRAGMA table_info(used_reagents)")
            columns = cursor.fetchall()
            print(f"  表格欄位：")
            for col in columns:
                print(f"    {col[1]} ({col[2]})")
            
            # 檢查資料數量
            cursor.execute("SELECT COUNT(*) FROM used_reagents")
            count = cursor.fetchone()[0]
            print(f"  總記錄數：{count}")
            
            # 檢查最近的資料
            if count > 0:
                cursor.execute("SELECT * FROM used_reagents ORDER BY stored_date DESC LIMIT 5")
                rows = cursor.fetchall()
                print(f"  最近 5 筆記錄：")
                for i, row in enumerate(rows, 1):
                    print(f"    {i}. ID: {row[0]}, 項目: {row[1]}, 批號: {row[2]}, 儲存時間: {row[3]}, 有效期限: {row[4]}")
            else:
                print("  ⚠️  表格中沒有資料")
        else:
            print(f"\n❌ used_reagents 表格不存在")
        
        conn.close()
        
    except Exception as e:
        print(f"❌ 檢查資料庫時發生錯誤：{e}")

if __name__ == "__main__":
    check_database()
