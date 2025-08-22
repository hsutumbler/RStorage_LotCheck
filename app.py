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
import subprocess

# 嘗試導入Windows列印相關套件
try:
    import win32print
    import win32api
    WINDOWS_PRINT_AVAILABLE = True
except ImportError:
    WINDOWS_PRINT_AVAILABLE = False
    print("警告：win32print 套件未安裝，Windows列印功能將不可用")

app = Flask(__name__)

# 獲取當前工作目錄的絕對路徑
base_dir = os.path.abspath(os.path.dirname(__file__))
instance_dir = os.path.join(base_dir, 'instance')

# 確保 instance 目錄存在
try:
    if not os.path.exists(instance_dir):
        os.makedirs(instance_dir)
        print(f"已創建 instance 目錄: {instance_dir}")
except Exception as e:
    print(f"創建 instance 目錄失敗: {e}")

# 構建資料庫文件的絕對路徑
db_path = os.path.join(instance_dir, 'reagent_inventory.db')
db_uri = f'sqlite:///{db_path}'

print(f"資料庫路徑: {db_path}")
app.config['SQLALCHEMY_DATABASE_URI'] = db_uri
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# 資料庫模型 - 必須在 create_all 之前定義
class ReagentEntry(db.Model):
    __tablename__ = 'reagent_entry'  # 明確指定表名
    
    id = db.Column(db.Integer, primary_key=True)
    reagent_name = db.Column(db.String(100), nullable=False)
    old_batch_number = db.Column(db.String(50), nullable=True)
    new_batch_number = db.Column(db.String(50), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    entry_date = db.Column(db.DateTime, default=datetime.utcnow)
    entry_type = db.Column(db.String(20), nullable=True)
    expiry_date = db.Column(db.Date, nullable=True)  # 添加穩定效期欄位
    supplier = db.Column(db.String(50), nullable=True)  # 添加供應商欄位
    unit = db.Column(db.String(20), nullable=True)  # 添加單位欄位
    
    def __repr__(self):
        return f'<ReagentEntry {self.reagent_name} - {self.new_batch_number}>'

# 確保在應用程序啟動時創建所有表
def init_db():
    with app.app_context():
        try:
            # 刪除現有的資料庫文件（如果存在）
            if os.path.exists(db_path):
                os.remove(db_path)
                print("已刪除舊的資料庫文件")
            
            # 創建新的資料庫和表
            db.create_all()
            
            # 驗證表是否創建成功
            inspector = db.inspect(db.engine)
            tables = inspector.get_table_names()
            print(f"已創建的表: {tables}")
            
            if 'reagent_entry' in tables:
                # 檢查表結構
                columns = [col['name'] for col in inspector.get_columns('reagent_entry')]
                print(f"reagent_entry 表的欄位: {columns}")
                print("資料庫和表結構創建成功")
            else:
                print("警告：reagent_entry 表未創建")
                
        except Exception as e:
            print(f"資料庫操作失敗: {e}")
            print(f"當前工作目錄: {os.getcwd()}")
            print(f"目標目錄權限: {os.access(instance_dir, os.W_OK)}")
            raise e

# 初始化資料庫
init_db()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/entries', methods=['GET'])
def get_entries():
    try:
        # 查詢記錄
        entries = ReagentEntry.query.order_by(ReagentEntry.entry_date.desc()).all()
        print(f"找到 {len(entries)} 筆記錄")
        
        result = []
        for entry in entries:
            # 將資料庫欄位映射到前端所需的格式
            batch_number = entry.new_batch_number or ""
            old_batch = entry.old_batch_number or ""
            
            result.append({
                'id': entry.id,
                'reagent_name': entry.reagent_name,
                'reagent_batch_number': batch_number,  # 使用 new_batch_number 作為 reagent_batch_number
                'quantity': entry.quantity,
                'expiry_date': entry.expiry_date.strftime('%Y-%m-%d') if entry.expiry_date else "無資料",
                'entry_date': entry.entry_date.strftime('%Y-%m-%d %H:%M:%S') if entry.entry_date else "",
                'supplier': entry.supplier or "無資料",  # 添加供應商資訊
                'unit': entry.unit or "個"  # 添加單位資訊，默認為"個"
            })
        
        print(f"返回 {len(result)} 筆記錄")
        print(f"返回資料類型: {type(result)}")
        
        # 確保返回的是列表，即使是空列表
        if not result:
            result = []
            
        final_result = jsonify(result)
        print(f"JSON後資料類型: {type(final_result)}")
        return final_result
        
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
        
        # 檢查是否為新批號
        existing_entry = ReagentEntry.query.filter_by(
            reagent_name=data['reagent_name'],
            new_batch_number=data['reagent_batch_number']
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
        # 打印收到的數據，用於調試
        print("收到的數據:", data)
        
        try:
            expiry_date = datetime.strptime(data['expiry_date'], '%Y-%m-%d').date() if data.get('expiry_date') else None
        except Exception as e:
            print(f"解析穩定效期時出錯: {e}")
            expiry_date = None

        entry = ReagentEntry(
            reagent_name=data['reagent_name'],
            new_batch_number=data['reagent_batch_number'],
            old_batch_number=data.get('old_batch_number', ''),
            quantity=data['quantity'],
            entry_type='old',  # 標記為舊批號
            expiry_date=expiry_date,
            supplier=data.get('supplier', ''),  # 保存供應商資訊
            unit=data.get('unit', '')  # 保存單位資訊
        )
        
        db.session.add(entry)
        db.session.commit()
        
        print(f"成功新增記錄: {entry.id}")
        
        # 生成並列印標籤
        labels_printed = generate_and_print_labels(entry)
        
        return jsonify({
            'success': True,
            'is_new_batch': False
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
        
        # 打印收到的數據，用於調試
        print("收到的數據:", data)
        
        try:
            expiry_date = datetime.strptime(data['expiry_date'], '%Y-%m-%d').date() if data.get('expiry_date') else None
        except Exception as e:
            print(f"解析穩定效期時出錯: {e}")
            expiry_date = None

        entry = ReagentEntry(
            reagent_name=data['reagent_name'],
            new_batch_number=data['reagent_batch_number'],
            old_batch_number=data.get('old_batch_number', ''),
            quantity=data['quantity'],
            entry_type='new',  # 標記為新批號
            expiry_date=expiry_date,
            supplier=data.get('supplier', ''),  # 保存供應商資訊
            unit=data.get('unit', '')  # 保存單位資訊
        )
        
        db.session.add(entry)
        db.session.commit()
        
        print(f"成功新增新批號記錄: {entry.id}")
        
        # 生成並列印標籤
        labels_printed = generate_and_print_labels(entry)
        
        return jsonify({
            'success': True
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
                ReagentEntry.new_batch_number.contains(query)
            )
        ).order_by(ReagentEntry.entry_date.desc()).all()
        
        result = []
        for entry in entries:
            # 將資料庫欄位映射到前端所需的格式
            batch_number = entry.new_batch_number or ""
            old_batch = entry.old_batch_number or ""
            
            result.append({
                'id': entry.id,
                'reagent_name': entry.reagent_name,
                'reagent_batch_number': batch_number,
                'old_batch_number': old_batch,
                'quantity': entry.quantity,
                'entry_type': entry.entry_type or "",
                'notes': entry.notes or "",
                'entry_date': entry.entry_date.strftime('%Y-%m-%d %H:%M:%S') if entry.entry_date else ""
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
        c.drawString(25*mm, 25*mm, f"批號：{entry.new_batch_number}")
        
        # 穩定效期（如果有）
        if entry.expiry_date:
            c.drawString(2*mm, 20*mm, f"穩定效期：{entry.expiry_date.strftime('%Y/%m/%d')}")
        
        # 入庫時間
        c.drawString(2*mm, 15*mm, f"入庫時間：{entry.entry_date.strftime('%Y/%m/%d') if entry.entry_date else ''}")
        
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
    
    # 嘗試使用Windows預設印表機列印
    try:
        print(f"已生成 {quantity} 張標籤PDF: {temp_file.name}")
        
        if WINDOWS_PRINT_AVAILABLE:
            print("正在嘗試列印...")
            
            # 獲取預設印表機名稱
            default_printer = win32print.GetDefaultPrinter()
            print(f"使用預設印表機: {default_printer}")
            
            # 使用系統預設應用程式開啟PDF（通常是Adobe Reader或其他PDF閱讀器）
            # 這樣用戶可以選擇列印或儲存
            subprocess.Popen(['start', temp_file.name], shell=True)
            
            print(f"PDF已開啟，請在PDF閱讀器中選擇列印或儲存")
        else:
            print("Windows列印功能不可用，改為開啟PDF檔案")
            subprocess.Popen(['start', temp_file.name], shell=True)
            print(f"PDF已開啟，請手動選擇列印或儲存")
        
        return quantity
        
    except Exception as e:
        print(f"列印失敗: {e}")
        print("改為自動開啟PDF檔案")
        try:
            # 如果列印失敗，至少開啟PDF讓用戶手動處理
            subprocess.Popen(['start', temp_file.name], shell=True)
        except:
            pass
        return quantity

@app.route('/api/test')
def test():
    return jsonify({'message': 'API測試成功', 'status': 'ok'})

@app.route('/api/print-direct/<int:entry_id>', methods=['POST'])
def print_direct(entry_id):
    """直接列印到預設印表機"""
    try:
        entry = ReagentEntry.query.get_or_404(entry_id)
        data = request.json
        quantity = data.get('quantity', entry.quantity)
        
        print(f"直接列印請求: 記錄ID {entry_id}, 數量 {quantity}")
        
        # 生成標籤PDF
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        c = canvas.Canvas(temp_file.name, pagesize=(50*mm, 35*mm))
        
        # 獲取中文字體
        font_name = get_chinese_font()
        
        for i in range(quantity):
            # 標籤邊框
            c.rect(1*mm, 1*mm, 48*mm, 33*mm)
            
            # 標題
            c.setFont(font_name, 10)
            c.drawString(2*mm, 30*mm, "【入庫】")
            
            # 試劑名稱和批號
            c.setFont(font_name, 8)
            c.drawString(2*mm, 25*mm, f"名稱：{entry.reagent_name}")
            c.drawString(25*mm, 25*mm, f"批號：{entry.new_batch_number}")
            
            # 穩定效期（如果有）
            if entry.expiry_date:
                c.drawString(2*mm, 20*mm, f"穩定效期：{entry.expiry_date.strftime('%Y/%m/%d')}")
            
            # 入庫時間
            c.drawString(2*mm, 15*mm, f"入庫時間：{entry.entry_date.strftime('%Y/%m/%d') if entry.entry_date else ''}")
            
            # 出庫標題
            c.setFont(font_name, 10)
            c.drawString(2*mm, 7*mm, "【出庫】")
            
            # 出庫人員和日期
            c.setFont(font_name, 8)
            c.drawString(2*mm, 2*mm, "人員：")
            c.drawString(25*mm, 2*mm, "出庫日期：")
            
            if i < quantity - 1:
                c.showPage()
        
        c.save()
        
        # 嘗試直接列印到預設印表機
        try:
            if WINDOWS_PRINT_AVAILABLE:
                default_printer = win32print.GetDefaultPrinter()
                print(f"直接列印到: {default_printer}")
                
                # 使用系統命令列印PDF
                win32api.ShellExecute(0, "print", temp_file.name, None, ".", 0)
                
                return jsonify({
                    'success': True
                })
            else:
                # Windows列印功能不可用，開啟PDF讓用戶手動處理
                subprocess.Popen(['start', temp_file.name], shell=True)
                
                return jsonify({
                    'success': True
                })
            
        except Exception as e:
            print(f"直接列印失敗: {e}")
            # 如果直接列印失敗，開啟PDF讓用戶手動處理
            subprocess.Popen(['start', temp_file.name], shell=True)
            
            return jsonify({
                'success': True
            })
            
    except Exception as e:
        print(f"列印失敗: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/preview-label')
def preview_label():
    """預覽標籤版面"""
    try:
        # 創建一個示例入庫記錄
        
        class MockEntry:
            def __init__(self):
                self.reagent_name = "AFP"
                self.new_batch_number = "AFP001"
                self.old_batch_number = "AFP000"
                self.quantity = 1
                self.notes = "測試標籤"
                self.entry_type = "new"
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
        c.drawString(25*mm, 25*mm, f"批號：{mock_entry.new_batch_number}")
        
        # 舊批號
        c.drawString(2*mm, 20*mm, f"舊批號：{mock_entry.old_batch_number}")
        
        # 入庫時間
        c.drawString(2*mm, 15*mm, f"入庫時間：{mock_entry.entry_date.strftime('%Y/%m/%d')}")
        
        # 備註
        c.drawString(2*mm, 10*mm, f"備註：{mock_entry.notes}")
        
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
