#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
資料庫遷移腳本 - 為現有資料庫添加新欄位
"""

import sqlite3
import os

def migrate_database():
    """遷移資料庫，添加新欄位"""
    db_path = 'reagents.db'
    
    if not os.path.exists(db_path):
        print(f"❌ 資料庫檔案 {db_path} 不存在")
        return
    
    print(f"✅ 資料庫檔案 {db_path} 存在")
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 檢查現有表格結構
        cursor.execute("PRAGMA table_info(used_reagents)")
        columns = cursor.fetchall()
        existing_columns = [col[1] for col in columns]
        
        print(f"📋 現有欄位：{existing_columns}")
        
        # 需要添加的欄位
        new_columns = [
            ('quantity', 'INTEGER DEFAULT 1'),
            ('unit', 'TEXT'),
            ('supplier', 'TEXT')
        ]
        
        # 添加新欄位
        for column_name, column_type in new_columns:
            if column_name not in existing_columns:
                try:
                    cursor.execute(f'ALTER TABLE used_reagents ADD COLUMN {column_name} {column_type}')
                    print(f"✅ 已添加欄位：{column_name}")
                except Exception as e:
                    print(f"⚠️  添加欄位 {column_name} 時發生錯誤：{e}")
            else:
                print(f"ℹ️  欄位 {column_name} 已存在")
        
        # 檢查最終表格結構
        cursor.execute("PRAGMA table_info(used_reagents)")
        final_columns = cursor.fetchall()
        print(f"\n📋 最終欄位結構：")
        for col in final_columns:
            print(f"  {col[1]} ({col[2]})")
        
        conn.commit()
        conn.close()
        
        print("\n✅ 資料庫遷移完成！")
        
    except Exception as e:
        print(f"❌ 遷移資料庫時發生錯誤：{e}")

if __name__ == "__main__":
    migrate_database()
