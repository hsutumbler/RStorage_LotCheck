from flask import Flask, render_template, jsonify, request, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date
import tempfile
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///simple_inventory.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# 簡化的資料庫模型
class ReagentEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    reagent_name = db.Column(db.String(100), nullable=False)
    reagent_batch_number = db.Column(db.String(50), nullable=False)
    expiry_date = db.Column(db.Date, nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    unit = db.Column(db.String(20), nullable=False)
    supplier = db.Column(db.String(100), nullable=False)
    entry_date = db.Column(db.DateTime, default=datetime.utcnow)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/entries', methods=['GET'])
def get_entries():
    try:
        # 先建立資料庫
        with app.app_context():
            db.create_all()
        
        # 查詢記錄
        entries = ReagentEntry.query.order_by(ReagentEntry.entry_date.desc()).all()
        print(f"找到 {len(entries)} 筆記錄")
        
        result = []
        for entry in entries:
            result.append({
                'id': entry.id,
                'reagent_name': entry.reagent_name,
                'reagent_batch_number': entry.reagent_batch_number,
                'expiry_date': entry.expiry_date.strftime('%Y-%m-%d'),
                'quantity': entry.quantity,
                'unit': entry.unit,
                'supplier': entry.supplier,
                'entry_date': entry.entry_date.strftime('%Y-%m-%d %H:%M:%S')
            })
        
        print(f"返回 {len(result)} 筆記錄")
        return jsonify(result)
        
    except Exception as e:
        print(f"錯誤: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/entries', methods=['POST'])
def add_entry():
    try:
        data = request.json
        print(f"收到新增請求: {data}")
        
        # 建立資料庫
        with app.app_context():
            db.create_all()
        
        # 檢查是否為新批號
        existing_entry = ReagentEntry.query.filter_by(
            reagent_name=data['reagent_name'],
            reagent_batch_number=data['reagent_batch_number']
        ).first()
        
        is_new_batch = existing_entry is None
        
        if is_new_batch:
            print(f"檢測到新批號: {data['reagent_name']} - {data['reagent_batch_number']}")
            return jsonify({
                'success': False,
                'is_new_batch': True,
                'message': f'檢測到新批號！\n試劑名稱: {data["reagent_name"]}\n批號: {data["reagent_batch_number"]}\n\n請確認是否要入庫？',
                'data': data
            })
        
        # 如果不是新批號，直接入庫
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
        
        print(f"成功新增記錄: {entry.id}")
        
        # 生成並列印標籤
        labels_printed = generate_and_print_labels(entry)
        
        return jsonify({
            'success': True,
            'is_new_batch': False,
            'message': f'入庫成功！已列印 {labels_printed} 張標籤'
        })
        
    except Exception as e:
        print(f"新增記錄失敗: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/confirm-entry', methods=['POST'])
def confirm_entry():
    try:
        data = request.json
        print(f"確認入庫請求: {data}")
        
        # 建立資料庫
        with app.app_context():
            db.create_all()
        
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
        
        print(f"成功新增新批號記錄: {entry.id}")
        
        # 生成並列印標籤
        labels_printed = generate_and_print_labels(entry)
        
        return jsonify({
            'success': True,
            'message': f'新批號入庫成功！已列印 {labels_printed} 張標籤'
        })
        
    except Exception as e:
        print(f"確認入庫失敗: {e}")
        import traceback
        traceback.print_exc()
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
        
        result = []
        for entry in entries:
            result.append({
                'id': entry.id,
                'reagent_name': entry.reagent_name,
                'reagent_batch_number': entry.reagent_batch_number,
                'expiry_date': entry.expiry_date.strftime('%Y-%m-%d'),
                'quantity': entry.quantity,
                'unit': entry.unit,
                'supplier': entry.supplier,
                'entry_date': entry.entry_date.strftime('%Y-%m-%d %H:%M:%S')
            })
        
        return jsonify(result)
        
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

def get_chinese_font():
    """獲取可用的中文字體"""
    try:
        # 嘗試註冊微軟正黑體
        font_path = "C:/Windows/Fonts/msjh.ttc"
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
            return 'ChineseFont'
    except:
        pass
    
    try:
        # 嘗試註冊微軟雅黑
        font_path = "C:/Windows/Fonts/msyh.ttc"
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
            return 'ChineseFont'
    except:
        pass
    
    # 使用預設字體
    return 'Helvetica'

def generate_and_print_labels(entry, quantity=None):
    if quantity is None:
        quantity = entry.quantity
    
    # 標籤尺寸：5cm x 3.5cm
    label_width = 50 * mm  # 5cm
    label_height = 35 * mm  # 3.5cm
    
    # 建立標籤PDF
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    c = canvas.Canvas(temp_file.name, pagesize=(label_width, label_height))
    
    # 獲取中文字體
    font_name = get_chinese_font()
    
    for i in range(quantity):
        # 標籤邊框
        c.rect(1*mm, 1*mm, label_width-2*mm, label_height-2*mm)
        
        # 標題
        c.setFont(font_name, 10)
        c.drawString(2*mm, 30*mm, "【入庫】")
        
        # 試劑名稱和批號
        c.setFont(font_name, 8)
        c.drawString(2*mm, 25*mm, f"名稱：{entry.reagent_name}")
        c.drawString(25*mm, 25*mm, f"批號{entry.reagent_batch_number}")
        
        # 穩定效期
        c.drawString(2*mm, 20*mm, f"穩定效期：{entry.expiry_date.strftime('%Y/%m/%d')}")
        
        # 入庫時間
        c.drawString(2*mm, 15*mm, f"入庫時間：{entry.entry_date.strftime('%Y/%m/%d')}")
        
        # 出庫標題 - 離上一行0.8cm (8mm)
        c.setFont(font_name, 10)
        c.drawString(2*mm, 7*mm, "【出庫】")
        
        # 出庫人員和日期（留白給蓋章用）- 離出庫標題0.5cm (5mm)
        c.setFont(font_name, 8)
        c.drawString(2*mm, 2*mm, "人員：")
        c.drawString(25*mm, 2*mm, "出庫日期：")
        
        # 如果不是最後一頁，新增頁面
        if i < quantity - 1:
            c.showPage()
    
    c.save()
    
    # 這裡可以整合實際的列印功能
    # 目前只是生成PDF檔案
    print(f"已生成 {quantity} 張標籤PDF: {temp_file.name}")
    
    return quantity

@app.route('/api/test')
def test():
    return jsonify({'message': 'API測試成功', 'status': 'ok'})

@app.route('/preview-label')
def preview_label():
    """預覽標籤版面"""
    try:
        # 創建一個示例入庫記錄
        
        class MockEntry:
            def __init__(self):
                self.reagent_name = "AFP"
                self.reagent_batch_number = "AFP001"
                self.expiry_date = date(2025, 8, 31)
                self.quantity = 1
                self.unit = "組"
                self.supplier = "亞培"
                self.entry_date = datetime(2025, 8, 20)
        
        mock_entry = MockEntry()
        
        # 生成標籤PDF
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        c = canvas.Canvas(temp_file.name, pagesize=(50*mm, 35*mm))
        
        # 獲取中文字體
        font_name = get_chinese_font()
        
        # 標籤邊框
        c.rect(1*mm, 1*mm, 48*mm, 33*mm)
        
        # 標題
        c.setFont(font_name, 10)
        c.drawString(2*mm, 30*mm, "【入庫】")
        
        # 試劑名稱和批號
        c.setFont(font_name, 8)
        c.drawString(2*mm, 25*mm, f"名稱：{mock_entry.reagent_name}")
        c.drawString(25*mm, 25*mm, f"批號{mock_entry.reagent_batch_number}")
        
        # 穩定效期
        c.drawString(2*mm, 20*mm, f"穩定效期：{mock_entry.expiry_date.strftime('%Y/%m/%d')}")
        
        # 入庫時間
        c.drawString(2*mm, 15*mm, f"入庫時間：{mock_entry.entry_date.strftime('%Y/%m/%d')}")
        
        # 出庫標題 - 離上一行0.8cm (8mm)
        c.setFont(font_name, 10)
        c.drawString(2*mm, 7*mm, "【出庫】")
        
        # 出庫人員和日期（留白給蓋章用）- 離出庫標題0.5cm (5mm)
        c.setFont(font_name, 8)
        c.drawString(2*mm, 2*mm, "人員：")
        c.drawString(25*mm, 2*mm, "出庫日期：")
        
        c.save()
        
        # 返回PDF文件
        return send_file(temp_file.name, as_attachment=True, download_name='label_preview.pdf')
        
    except Exception as e:
        print(f"預覽標籤失敗: {e}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
