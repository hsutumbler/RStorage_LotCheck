from flask import Flask, render_template, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import tempfile
import uuid

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///reagent_inventory.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# 資料庫模型
class ReagentEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    reagent_name = db.Column(db.String(100), nullable=False)
    reagent_batch_number = db.Column(db.String(50), nullable=False)
    expiry_date = db.Column(db.Date, nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    unit = db.Column(db.String(20), nullable=False)
    supplier = db.Column(db.String(100), nullable=False)
    entry_date = db.Column(db.DateTime, default=datetime.utcnow)

# 建立資料庫表格
def init_db():
    try:
        with app.app_context():
            db.create_all()
            print("資料庫初始化成功")
    except Exception as e:
        print(f"資料庫初始化失敗: {e}")

init_db()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/entries', methods=['GET'])
def get_entries():
    try:
        entries = ReagentEntry.query.order_by(ReagentEntry.entry_date.desc()).all()
        return jsonify([{
            'id': entry.id,
            'reagent_name': entry.reagent_name,
            'reagent_batch_number': entry.reagent_batch_number,
            'expiry_date': entry.expiry_date.strftime('%Y-%m-%d'),
            'quantity': entry.quantity,
            'unit': entry.unit,
            'supplier': entry.supplier,
            'entry_date': entry.entry_date.strftime('%Y-%m-%d %H:%M:%S')
        } for entry in entries])
    except Exception as e:
        print(f"獲取記錄失敗: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/entries', methods=['POST'])
def add_entry():
    try:
        data = request.json
        
        # 檢查是否為新批號
        existing_entry = ReagentEntry.query.filter_by(
            reagent_name=data['reagent_name'],
            reagent_batch_number=data['reagent_batch_number']
        ).first()
        
        entry = ReagentEntry(
            reagent_name=data['reagent_name'],
            reagent_batch_number=data['reagent_batch_number'],
            expiry_date=datetime.strptime(data['expiry_date'], '%Y-%m-%d').date(),
            quantity=data['quantity'],
            unit=data['unit'],
            supplier=data['supplier']
        )
        
        db.session.add(entry)
        db.session.commit()
        
        # 生成並列印標籤
        labels_printed = generate_and_print_labels(entry)
        
        return jsonify({
            'success': True,
            'message': f'入庫成功！已列印 {labels_printed} 張標籤'
        })
    except Exception as e:
        print(f"新增記錄失敗: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/search', methods=['GET'])
def search_entries():
    try:
        query = request.args.get('q', '')
        if not query:
            return jsonify([])
        
        entries = ReagentEntry.query.filter(
            db.or_(
                ReagentEntry.reagent_name.contains(query),
                ReagentEntry.reagent_batch_number.contains(query),
                ReagentEntry.supplier.contains(query)
            )
        ).order_by(ReagentEntry.entry_date.desc()).all()
        
        return jsonify([{
            'id': entry.id,
            'reagent_name': entry.reagent_name,
            'reagent_batch_number': entry.reagent_batch_number,
            'expiry_date': entry.expiry_date.strftime('%Y-%m-%d'),
            'quantity': entry.quantity,
            'unit': entry.unit,
            'supplier': entry.supplier,
            'entry_date': entry.entry_date.strftime('%Y-%m-%d %H:%M:%S')
        } for entry in entries])
    except Exception as e:
        print(f"搜尋記錄失敗: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/print-labels/<int:entry_id>', methods=['POST'])
def reprint_labels(entry_id):
    try:
        entry = ReagentEntry.query.get_or_404(entry_id)
        data = request.json
        quantity = data.get('quantity', entry.quantity)
        
        labels_printed = generate_and_print_labels(entry, quantity)
        
        return jsonify({
            'success': True,
            'message': f'已重新列印 {labels_printed} 張標籤'
        })
    except Exception as e:
        print(f"補印標籤失敗: {e}")
        return jsonify({'error': str(e)}), 500

def generate_and_print_labels(entry, quantity=None):
    if quantity is None:
        quantity = entry.quantity
    
    # 建立標籤PDF
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    c = canvas.Canvas(temp_file.name, pagesize=A4)
    
    for i in range(quantity):
        # 標籤內容
        c.setFont("Helvetica-Bold", 16)
        c.drawString(20*mm, 270*mm, f"試劑名稱: {entry.reagent_name}")
        
        c.setFont("Helvetica", 14)
        c.drawString(20*mm, 250*mm, f"試劑批號: {entry.reagent_batch_number}")
        c.drawString(20*mm, 230*mm, f"穩定效期: {entry.expiry_date.strftime('%Y-%m-%d')}")
        c.drawString(20*mm, 210*mm, f"入庫數量: {entry.quantity} {entry.unit}")
        c.drawString(20*mm, 190*mm, f"供應商: {entry.supplier}")
        c.drawString(20*mm, 170*mm, f"入庫日期: {entry.entry_date.strftime('%Y-%m-%d')}")
        
        # 如果不是最後一頁，新增頁面
        if i < quantity - 1:
            c.showPage()
    
    c.save()
    
    # 這裡可以整合實際的列印功能
    # 目前只是生成PDF檔案
    print(f"已生成 {quantity} 張標籤PDF: {temp_file.name}")
    
    return quantity

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
