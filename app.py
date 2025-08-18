#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
試劑批號確認系統 - Flask 後端
檢驗科專用，用於確認試劑批號是否使用過
"""

from flask import Flask, render_template, request, jsonify, redirect, url_for
import sqlite3
import datetime
import os
import warnings
from werkzeug.exceptions import BadRequest

# 隱藏 Flask 開發伺服器警告
warnings.filterwarnings("ignore", message="This is a development server")
warnings.filterwarnings("ignore", category=UserWarning, module="werkzeug")

app = Flask(__name__)

# 資料庫檔案路徑
DATABASE = 'reagents.db'

def init_db():
    """初始化資料庫，建立必要的表格"""
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    
    # 建立 used_reagents 表格
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS used_reagents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item TEXT NOT NULL,
            lot_number TEXT UNIQUE NOT NULL,
            expiration_date TEXT NOT NULL,
            quantity INTEGER DEFAULT 1,
            unit TEXT,
            supplier TEXT,
            stored_date TEXT NOT NULL
        )
    ''')
    
    # 建立索引以提升查詢效能
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_lot_number ON used_reagents(lot_number)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_item ON used_reagents(item)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_stored_date ON used_reagents(stored_date)')
    
    conn.commit()
    conn.close()
    
    print("資料庫初始化完成！")

def get_db_connection():
    """獲取資料庫連接"""
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row  # 讓查詢結果可以像字典一樣存取
    return conn

@app.route('/')
def index():
    """主頁面 - 顯示試劑資訊儲存介面"""
    return render_template('index.html')

@app.route('/search')
def search_page():
    """批號查詢頁面"""
    return render_template('search.html')

@app.route('/api/save_reagent', methods=['POST'])
def save_reagent():
    """儲存試劑資訊的 API 端點"""
    try:
        data = request.get_json()
        
        # 驗證必要欄位
        if not data:
            raise BadRequest("沒有接收到資料")
        
        item = data.get('item', '').strip()
        lot_number = data.get('lot_number', '').strip()
        expiration_date = data.get('expiration_date', '').strip()
        
        if not item:
            return jsonify({'success': False, 'message': '請輸入檢驗項目名稱'}), 400
        
        if not lot_number:
            return jsonify({'success': False, 'message': '請輸入批號'}), 400
        
        if not expiration_date:
            return jsonify({'success': False, 'message': '請選擇有效期限'}), 400
        
        # 檢查批號是否已存在
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT id, item, stored_date FROM used_reagents WHERE lot_number = ?', (lot_number,))
        existing = cursor.fetchone()
        
        if existing:
            conn.close()
            return jsonify({
                'success': False, 
                'message': f'批號 {lot_number} 已經存在於系統中',
                'details': {
                    'item': existing['item'],
                    'stored_date': existing['stored_date']
                }
            }), 400
        
        # 獲取當前時間
        current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # 獲取其他欄位
        quantity = data.get('quantity', 1)
        unit = data.get('unit', '').strip()
        supplier = data.get('supplier', '').strip()
        
        # 插入新記錄
        cursor.execute('''
            INSERT INTO used_reagents (item, lot_number, expiration_date, quantity, unit, supplier, stored_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (item, lot_number, expiration_date, quantity, unit, supplier, current_time))
        
        conn.commit()
        conn.close()
        
        return jsonify({
            'success': True,
            'message': f'試劑批號 {lot_number} 已成功儲存到系統中'
        })
        
    except BadRequest as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        print(f"儲存試劑時發生錯誤: {e}")
        return jsonify({'success': False, 'message': '儲存失敗，請稍後再試'}), 500

@app.route('/api/search_reagent', methods=['POST'])
def search_reagent():
    """查詢試劑批號的 API 端點"""
    try:
        data = request.get_json()
        
        if not data:
            raise BadRequest("沒有接收到資料")
        
        lot_number = data.get('lot_number', '').strip()
        
        if not lot_number:
            return jsonify({'success': False, 'message': '請輸入批號'}), 400
        
        # 查詢資料庫
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, item, lot_number, expiration_date, quantity, unit, supplier, stored_date
            FROM used_reagents 
            WHERE lot_number = ?
        ''', (lot_number,))
        
        result = cursor.fetchone()
        conn.close()
        
        if result:
            # 檢查是否過期
            try:
                exp_date = datetime.datetime.strptime(result['expiration_date'], '%Y-%m-%d').date()
                today = datetime.date.today()
                is_expired = exp_date < today
            except:
                is_expired = False
            
            return jsonify({
                'success': True,
                'found': True,
                'message': f'批號 {lot_number} 已使用過',
                'data': {
                    'id': result['id'],
                    'item': result['item'],
                    'lot_number': result['lot_number'],
                    'expiration_date': result['expiration_date'],
                    'quantity': result['quantity'],
                    'unit': result['unit'],
                    'supplier': result['supplier'],
                    'stored_date': result['stored_date'],
                    'is_expired': is_expired
                }
            })
        else:
            return jsonify({
                'success': True,
                'found': False,
                'message': f'批號 {lot_number} 尚未被使用'
            })
            
    except BadRequest as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        print(f"查詢試劑時發生錯誤: {e}")
        return jsonify({'success': False, 'message': '查詢失敗，請稍後再試'}), 500

@app.route('/api/get_all_reagents')
def get_all_reagents():
    """獲取所有試劑資訊的 API 端點（用於顯示列表）"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, item, lot_number, expiration_date, quantity, unit, supplier, stored_date
            FROM used_reagents 
            ORDER BY stored_date DESC
        ''')
        
        results = cursor.fetchall()
        conn.close()
        
        # 轉換為字典列表
        reagents = []
        for row in results:
            try:
                exp_date = datetime.datetime.strptime(row['expiration_date'], '%Y-%m-%d').date()
                today = datetime.date.today()
                is_expired = exp_date < today
            except:
                is_expired = False
            
            reagents.append({
                'id': row['id'],
                'item': row['item'],
                'lot_number': row['lot_number'],
                'expiration_date': row['expiration_date'],
                'quantity': row['quantity'],
                'unit': row['unit'],
                'supplier': row['supplier'],
                'stored_date': row['stored_date'],
                'is_expired': is_expired
            })
        
        return jsonify({'success': True, 'data': reagents})
        
    except Exception as e:
        print(f"獲取試劑列表時發生錯誤: {e}")
        return jsonify({'success': False, 'message': '獲取資料失敗，請稍後再試'}), 500

@app.route('/list')
def list_page():
    """試劑列表頁面"""
    return render_template('list.html')

@app.route('/debug')
def debug_page():
    """系統診斷頁面"""
    return render_template('debug.html')

if __name__ == '__main__':
    # 確保資料庫存在
    if not os.path.exists(DATABASE):
        init_db()
        print(f"已建立資料庫檔案: {DATABASE}")
    
    print("試劑批號確認系統啟動中...")
    print("請在瀏覽器中開啟: http://localhost:5000")
    
    # 啟動 Flask 應用程式（隱藏開發伺服器警告）
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=True)
