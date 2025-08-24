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
        
        return jsonify({
            'success': True,
            'is_new_batch': False,
            'entry_id': entry.id
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
        
        return jsonify({
            'success': True,
            'entry_id': entry.id
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

@app.route('/api/import-csv', methods=['POST'])
def import_csv():
    """匯入CSV檔案到資料庫"""
    try:
        # 檢查是否有檔案上傳
        if 'file' not in request.files:
            return jsonify({'error': '沒有選擇檔案'}), 400
        
        file = request.files['file']
        
        # 檢查檔案名稱
        if file.filename == '':
            return jsonify({'error': '沒有選擇檔案'}), 400
        
        # 檢查檔案類型
        if not file.filename.endswith('.csv'):
            return jsonify({'error': '請選擇CSV檔案'}), 400
        
        # 讀取CSV內容
        import csv
        import io
        
        # 設定UTF-8編碼
        content = file.read().decode('utf-8-sig')  # 處理BOM
        csv_reader = csv.DictReader(io.StringIO(content))
        
        # 建立資料庫
        with app.app_context():
            db.create_all()
        
        success_count = 0
        error_count = 0
        errors = []
        
        # 處理每一行資料
        for row_num, row in enumerate(csv_reader, start=2):  # 從第2行開始（第1行是標題）
            try:
                # 檢查必要欄位
                required_fields = ['試劑名稱', '試劑批號', '穩定效期', '數量', '單位', '供應商', '入庫日期']
                for field in required_fields:
                    if not row.get(field) or not row[field].strip():
                        raise ValueError(f'欄位 "{field}" 不能為空')
                
                # 解析日期
                try:
                    expiry_date = datetime.strptime(row['穩定效期'].strip(), '%Y-%m-%d').date()
                    entry_date = datetime.strptime(row['入庫日期'].strip(), '%Y-%m-%d %H:%M:%S')
                except ValueError:
                    # 嘗試其他日期格式
                    try:
                        expiry_date = datetime.strptime(row['穩定效期'].strip(), '%Y/%m/%d').date()
                        entry_date = datetime.strptime(row['入庫日期'].strip(), '%Y/%m/%d').replace(hour=0, minute=0, second=0)
                    except ValueError:
                        raise ValueError(f'日期格式錯誤: 穩定效期="{row["穩定效期"]}", 入庫日期="{row["入庫日期"]}"')
                
                # 解析數量
                try:
                    quantity = int(row['數量'].strip())
                    if quantity <= 0:
                        raise ValueError('數量必須大於0')
                except ValueError:
                    raise ValueError(f'數量格式錯誤: "{row["數量"]}"')
                
                # 檢查是否已存在相同名稱和批號的記錄
                existing_entry = ReagentEntry.query.filter_by(
                    reagent_name=row['試劑名稱'].strip(),
                    reagent_batch_number=row['試劑批號'].strip()
                ).first()
                
                if existing_entry:
                    # 更新現有記錄
                    existing_entry.expiry_date = expiry_date
                    existing_entry.quantity = quantity
                    existing_entry.unit = row['單位'].strip()
                    existing_entry.supplier = row['供應商'].strip()
                    existing_entry.entry_date = entry_date
                    db.session.commit()
                    success_count += 1
                else:
                    # 新增新記錄
                    new_entry = ReagentEntry(
                        reagent_name=row['試劑名稱'].strip(),
                        reagent_batch_number=row['試劑批號'].strip(),
                        expiry_date=expiry_date,
                        quantity=quantity,
                        unit=row['單位'].strip(),
                        supplier=row['供應商'].strip(),
                        entry_date=entry_date
                    )
                    db.session.add(new_entry)
                    db.session.commit()
                    success_count += 1
                
            except Exception as e:
                error_count += 1
                errors.append(f'第{row_num}行: {str(e)}')
                # 繼續處理下一行，不中斷整個匯入過程
        
        # 提交所有變更
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': f'匯入完成！成功處理 {success_count} 筆記錄，失敗 {error_count} 筆',
            'success_count': success_count,
            'error_count': error_count,
            'errors': errors
        })
        
    except Exception as e:
        print(f"匯入CSV失敗: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'匯入失敗: {str(e)}'}), 500

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
        c.drawString(25*mm, 25*mm, f"批號：{entry.reagent_batch_number}")
        
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
            c.drawString(25*mm, 25*mm, f"批號：{entry.reagent_batch_number}")
            
            # 穩定效期
            c.drawString(2*mm, 20*mm, f"穩定效期：{entry.expiry_date.strftime('%Y/%m/%d')}")
            
            # 入庫時間
            c.drawString(2*mm, 15*mm, f"入庫時間：{entry.entry_date.strftime('%Y/%m/%d')}")
            
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
        c.drawString(25*mm, 25*mm, f"批號：{mock_entry.reagent_batch_number}")
        
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
    app.run(debug=True, host='127.0.0.1', port=5000)
