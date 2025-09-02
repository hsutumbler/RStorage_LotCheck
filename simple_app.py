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
import sys

# 嘗試導入Windows列印相關套件
try:
    import win32print
    import win32api
    WINDOWS_PRINT_AVAILABLE = True
except ImportError:
    WINDOWS_PRINT_AVAILABLE = False
    print("警告：win32print 套件未安裝，Windows列印功能將不可用")

# 取得執行檔所在目錄
def get_app_directory():
    """取得應用程式所在目錄"""
    if getattr(sys, 'frozen', False):
        # 如果是打包後的exe檔案
        app_dir = os.path.dirname(sys.executable)
    else:
        # 如果是開發環境
        app_dir = os.path.dirname(os.path.abspath(__file__))
    return app_dir

# 設定資料庫路徑（在exe檔案同目錄）
APP_DIR = get_app_directory()
DB_PATH = os.path.join(APP_DIR, 'simple_inventory.db')

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'
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
        
        # 獲取標籤機類型參數
        printer_type = data.get('printer_type', 'zpl')  # 預設為ZPL模式
        
        # 自動列印標籤
        if data.get('print_labels', True):  # 預設會列印
            labels_printed = generate_and_print_labels(entry, entry.quantity, False, printer_type)
            print(f"已列印 {labels_printed} 張標籤")
        
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
        
        # 獲取標籤機類型參數
        printer_type = data.get('printer_type', 'zpl')  # 預設為ZPL模式
        
        # 自動列印標籤
        if data.get('print_labels', True):  # 預設會列印
            labels_printed = generate_and_print_labels(entry, entry.quantity, True, printer_type)
            print(f"已列印 {labels_printed} 張新批號標籤")
        
        return jsonify({
            'success': True,
            'entry_id': entry.id,
            'is_new_batch': True
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

@app.route('/api/reagent-suggestions', methods=['GET'])
def get_reagent_suggestions():
    """獲取試劑名稱建議（僅從資料庫現有資料）"""
    try:
        query = request.args.get('q', '').strip()
        
        if not query or len(query) < 1:
            return jsonify([])
        
        # 獲取所有不重複的試劑名稱
        all_names = db.session.query(ReagentEntry.reagent_name).distinct().all()
        all_names = [name[0] for name in all_names]
        
        suggestions = []
        query_upper = query.upper()
        
        for name in all_names:
            name_upper = name.upper()
            
            # 計算匹配分數
            score = 0
            match_type = ''
            
            # 1. 完全匹配 (最高分)
            if name_upper == query_upper:
                score = 1000
                match_type = 'exact'
            # 2. 開頭匹配
            elif name_upper.startswith(query_upper):
                # 越短的名稱分數越高，越早匹配分數越高
                score = 900 - len(name) + (10 if len(query) > 1 else 0)
                match_type = 'prefix'
            # 3. 包含匹配
            elif query_upper in name_upper:
                # 包含位置越靠前分數越高
                position = name_upper.find(query_upper)
                score = 800 - (position * 10) - len(name)
                match_type = 'contains'
            # 4. 模糊匹配（字符順序匹配）
            else:
                # 檢查查詢字符是否按順序出現在名稱中
                query_chars = list(query_upper)
                name_chars = list(name_upper)
                
                matched_chars = 0
                name_index = 0
                
                for query_char in query_chars:
                    for i in range(name_index, len(name_chars)):
                        if name_chars[i] == query_char:
                            matched_chars += 1
                            name_index = i + 1
                            break
                
                # 如果至少匹配了70%的字符
                if matched_chars >= len(query_chars) * 0.7:
                    score = 700 - len(name) + (matched_chars * 10)
                    match_type = 'fuzzy'
            
            if score > 0:
                suggestions.append({
                    'name': name,
                    'score': score,
                    'match_type': match_type
                })
        
        # 按分數排序，取前8個
        suggestions.sort(key=lambda x: x['score'], reverse=True)
        suggestions = suggestions[:8]
        
        return jsonify([item['name'] for item in suggestions])
        
    except Exception as e:
        print(f"獲取建議失敗: {e}")
        return jsonify([])

@app.route('/api/reagent-supplier', methods=['GET'])
def get_reagent_supplier():
    """根據試劑名稱獲取常用供應商和單位"""
    try:
        reagent_name = request.args.get('name', '').strip()
        
        if not reagent_name:
            return jsonify({'found': False})
        
        # 查詢該試劑名稱最常用的供應商
        supplier_query = db.session.query(
            ReagentEntry.supplier,
            db.func.count(ReagentEntry.supplier).label('supplier_count')
        ).filter_by(
            reagent_name=reagent_name
        ).group_by(
            ReagentEntry.supplier
        ).order_by(
            db.func.count(ReagentEntry.supplier).desc()
        ).first()
        
        # 查詢該試劑名稱最常用的單位
        unit_query = db.session.query(
            ReagentEntry.unit,
            db.func.count(ReagentEntry.unit).label('unit_count')
        ).filter_by(
            reagent_name=reagent_name
        ).group_by(
            ReagentEntry.unit
        ).order_by(
            db.func.count(ReagentEntry.unit).desc()
        ).first()
        
        if not supplier_query and not unit_query:
            return jsonify({'found': False})
        
        result = {'found': True}
        
        if supplier_query:
            result['supplier'] = supplier_query[0]
            result['supplier_usage_count'] = supplier_query[1]
        
        if unit_query:
            result['unit'] = unit_query[0]
            result['unit_usage_count'] = unit_query[1]
        
        return jsonify(result)
        
    except Exception as e:
        print(f"獲取試劑供應商和單位失敗: {e}")
        return jsonify({'found': False})

@app.route('/api/batch-expiry', methods=['GET'])
def get_batch_expiry():
    """根據試劑名稱和批號獲取穩定效期"""
    try:
        reagent_name = request.args.get('name', '').strip()
        batch_number = request.args.get('batch', '').strip()
        
        if not reagent_name or not batch_number:
            return jsonify({'found': False})
        
        # 查詢特定批號的穩定效期
        entry = ReagentEntry.query.filter_by(
            reagent_name=reagent_name,
            reagent_batch_number=batch_number
        ).order_by(ReagentEntry.entry_date.desc()).first()
        
        if not entry:
            return jsonify({'found': False, 'is_new_batch': True})
        
        return jsonify({
            'found': True,
            'is_new_batch': False,
            'expiry_date': entry.expiry_date.strftime('%Y-%m-%d'),
            'last_entry_date': entry.entry_date.strftime('%Y-%m-%d'),
            'previous_quantity': entry.quantity,
            'previous_unit': entry.unit
        })
        
    except Exception as e:
        print(f"獲取批號效期失敗: {e}")
        return jsonify({'found': False})

@app.route('/api/reagent-batches', methods=['GET'])
def get_reagent_batches():
    """獲取指定試劑的所有歷史批號"""
    try:
        reagent_name = request.args.get('name', '').strip()
        
        if not reagent_name:
            return jsonify([])
        
        # 獲取該試劑的所有批號及其最新效期
        batches = db.session.query(
            ReagentEntry.reagent_batch_number,
            ReagentEntry.expiry_date,
            db.func.max(ReagentEntry.entry_date).label('latest_date')
        ).filter_by(
            reagent_name=reagent_name
        ).group_by(
            ReagentEntry.reagent_batch_number
        ).order_by(
            db.func.max(ReagentEntry.entry_date).desc()
        ).all()
        
        result = []
        for batch in batches:
            result.append({
                'batch_number': batch[0],
                'expiry_date': batch[1].strftime('%Y-%m-%d'),
                'latest_entry': batch[2].strftime('%Y-%m-%d')
            })
        
        return jsonify(result)
        
    except Exception as e:
        print(f"獲取試劑批號失敗: {e}")
        return jsonify([])

def get_chinese_font():
    """獲取可用的中文字體（粗體）"""
    try:
        # 嘗試註冊微軟正黑體粗體
        font_path = "C:/Windows/Fonts/msjhbd.ttc"
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('ChineseFontBold', font_path))
            return 'ChineseFontBold'
    except:
        pass
    
    try:
        # 嘗試註冊微軟正黑體（一般）
        font_path = "C:/Windows/Fonts/msjh.ttc"
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
            return 'ChineseFont'
    except:
        pass
    
    try:
        # 嘗試註冊微軟雅黑粗體
        font_path = "C:/Windows/Fonts/msyhbd.ttc"
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('ChineseFontBold', font_path))
            return 'ChineseFontBold'
    except:
        pass
    
    try:
        # 嘗試註冊微軟雅黑（一般）
        font_path = "C:/Windows/Fonts/msyh.ttc"
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
            return 'ChineseFont'
    except:
        pass
    
    # 使用預設粗體字體
    return 'Helvetica-Bold'

def generate_zpl_labels(entry, quantity=None, is_new_batch=False):
    """
    生成ZPL格式的標籤指令 (支援中文字型)
    - entry: 資料庫記錄
    - quantity: 列印數量
    - is_new_batch: 是否為新批號
    """
    if quantity is None:
        quantity = entry.quantity
    
    # ZPL配置參數
    dpi = 203  # 標籤機DPI
    label_width_mm = 50  # 標籤寬度 5cm
    label_height_mm = 35  # 標籤高度 3.5cm
    
    # 將mm轉換為點 (dots)
    label_width_dots = int(label_width_mm * dpi / 25.4)
    label_height_dots = int(label_height_mm * dpi / 25.4)
    
    # 文字位置 (Y軸位置，從上到下)
    title_y = int(29 * dpi / 25.4)  # 【入庫】
    name_y = int(25 * dpi / 25.4)   # 試劑名稱
    batch_y = int(21 * dpi / 25.4)  # 試劑批號
    expiry_y = int(17 * dpi / 25.4) # 穩定效期
    entry_y = int(13 * dpi / 25.4)  # 入庫時間
    out_title_y = int(8 * dpi / 25.4)  # 【出庫】
    out_info_y = int(4 * dpi / 25.4)   # 人員/出庫日期
    
    # X軸位置
    text_x = int(2 * dpi / 25.4)
    person_x = int(25 * dpi / 25.4)
    
    zpl_commands = []
    
    for i in range(quantity):
        zpl = f"^XA\n"  # 開始標籤
        
        # 設定標籤尺寸
        zpl += f"^PW{label_width_dots}\n"  # 設定列印寬度
        zpl += f"^LL{label_height_dots}\n"  # 設定標籤長度
        
        # 設定中文字型支援 (ZT410支援Unicode)
        # ^CI28 = 使用UTF-8編碼
        zpl += "^CI28\n"
        
        # 繪製邊框
        if is_new_batch:
            # 新批號：雙邊框
            # 外框 (粗線)
            outer_thickness = 6
            zpl += f"^FO10,10^GB{label_width_dots-20},{label_height_dots-20},{outer_thickness}^FS\n"
            # 內框 (細線)
            inner_thickness = 2
            inner_margin = 15
            zpl += f"^FO{inner_margin},{inner_margin}^GB{label_width_dots-2*inner_margin},{label_height_dots-2*inner_margin},{inner_thickness}^FS\n"
        else:
            # 舊批號：單邊框
            thickness = 2
            margin = 10
            zpl += f"^FO{margin},{margin}^GB{label_width_dots-2*margin},{label_height_dots-2*margin},{thickness}^FS\n"
        
        # 使用支援中文的字型
        # ^A@N = 使用Unicode字型, 字體大小
        # 標題【入庫】
        zpl += f"^FO{text_x},{title_y}^A@N,30,30^FD【入庫】^FS\n"
        
        # 試劑名稱
        zpl += f"^FO{text_x},{name_y}^A@N,24,24^FD試劑名稱：{entry.reagent_name}^FS\n"
        
        # 試劑批號
        if is_new_batch:
            batch_text = f"試劑批號：{entry.reagent_batch_number} >>新批號<<"
        else:
            batch_text = f"試劑批號：{entry.reagent_batch_number} (允收合格)"
        zpl += f"^FO{text_x},{batch_y}^A@N,24,24^FD{batch_text}^FS\n"
        
        # 穩定效期
        expiry_str = entry.expiry_date.strftime('%Y/%m/%d')
        zpl += f"^FO{text_x},{expiry_y}^A@N,24,24^FD穩定效期：{expiry_str}^FS\n"
        
        # 入庫時間
        entry_str = entry.entry_date.strftime('%Y/%m/%d')
        zpl += f"^FO{text_x},{entry_y}^A@N,24,24^FD入庫時間：{entry_str}^FS\n"
        
        # 出庫標題【出庫】
        zpl += f"^FO{text_x},{out_title_y}^A@N,30,30^FD【出庫】^FS\n"
        
        # 出庫資訊
        zpl += f"^FO{text_x},{out_info_y}^A@N,24,24^FD人員：^FS\n"
        zpl += f"^FO{person_x},{out_info_y}^A@N,24,24^FD出庫日期：^FS\n"
        
        zpl += "^XZ\n"  # 結束標籤
        
        zpl_commands.append(zpl)
    
    return zpl_commands

def send_zpl_to_printer(zpl_commands):
    """
    發送ZPL指令到印表機 (支援中文字型)
    """
    try:
        if not WINDOWS_PRINT_AVAILABLE:
            print("Windows列印功能不可用，無法發送ZPL指令")
            return False
        
        # 獲取預設印表機
        default_printer = win32print.GetDefaultPrinter()
        print(f"發送ZPL指令到: {default_printer}")
        
        # 開啟印表機
        printer_handle = win32print.OpenPrinter(default_printer)
        
        try:
            # 開始列印工作
            job_id = win32print.StartDocPrinter(printer_handle, 1, ("ZPL Label", None, "RAW"))
            win32print.StartPagePrinter(printer_handle)
            
            # 發送每個ZPL指令 (確保UTF-8編碼)
            for i, zpl in enumerate(zpl_commands):
                # 確保ZPL指令使用UTF-8編碼
                zpl_bytes = zpl.encode('utf-8')
                win32print.WritePrinter(printer_handle, zpl_bytes)
                print(f"已發送第 {i+1} 張標籤的ZPL指令")
            
            # 結束列印
            win32print.EndPagePrinter(printer_handle)
            win32print.EndDocPrinter(printer_handle)
            
            print(f"ZPL指令發送成功，共 {len(zpl_commands)} 張標籤")
            return True
            
        finally:
            win32print.ClosePrinter(printer_handle)
    
    except Exception as e:
        print(f"發送ZPL指令失敗: {e}")
        return False

def generate_and_print_labels(entry, quantity=None, is_new_batch=False, printer_type="pdf"):
    """
    統一的標籤生成和列印函數
    - entry: 資料庫記錄
    - quantity: 列印數量
    - is_new_batch: 是否為新批號
    - printer_type: 印表機類型 ("pdf" 或 "zpl")
    """
    if quantity is None:
        quantity = entry.quantity
    
    print(f"使用標籤機類型: {printer_type}")
    
    # 根據印表機類型選擇對應的生成方式
    if printer_type == "zpl":
        # 生成ZPL指令
        zpl_commands = generate_zpl_labels(entry, quantity, is_new_batch)
        
        # 嘗試發送到印表機
        if send_zpl_to_printer(zpl_commands):
            print(f"ZPL標籤列印成功: {quantity} 張")
            return quantity
        else:
            # ZPL列印失敗，保存為文字檔讓用戶手動處理或測試
            try:
                # 生成帶時間戳的ZPL文件名
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                zpl_filename = f"zpl_labels_{timestamp}.zpl"
                zpl_filepath = os.path.join(APP_DIR, zpl_filename)
                
                with open(zpl_filepath, 'w', encoding='utf-8') as f:
                    f.write(f"# ZPL標籤指令 - 生成時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"# 試劑名稱: {entry.reagent_name}\n")
                    f.write(f"# 批號: {entry.reagent_batch_number}\n")
                    f.write(f"# 新批號: {'是' if is_new_batch else '否'}\n")
                    f.write(f"# 數量: {quantity}\n")
                    f.write("# " + "="*50 + "\n\n")
                    
                    for i, zpl in enumerate(zpl_commands):
                        f.write(f"# 第 {i+1} 張標籤\n")
                        f.write(zpl)
                        f.write('\n\n')  # 分隔符
                
                print(f"ZPL指令已保存到: {zpl_filepath}")
                print("您可以使用ZPL Viewer開啟此檔案預覽標籤效果")
                
                # 嘗試用記事本開啟
                try:
                    subprocess.Popen(['notepad.exe', zpl_filepath])
                except:
                    print("無法自動開啟ZPL檔案，請手動開啟")
                
                return quantity
            except Exception as e:
                print(f"保存ZPL檔案失敗: {e}")
                return 0
    else:
        # 使用PDF模式 (原有邏輯)
        return generate_pdf_labels(entry, quantity, is_new_batch)

def print_pdf_direct(entry, quantity=None, is_new_batch=False):
    """
    PDF直接列印函數 - 恢復原本的列印邏輯
    """
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
        if is_new_batch:
            # 新批號標籤：雙重邊框
            # 外層邊框（粗邊框）
            c.setLineWidth(2)
            c.rect(0.5*mm, 0.5*mm, label_width-1*mm, label_height-1*mm)
            # 內層邊框（細邊框）
            c.setLineWidth(0.5)
            c.rect(1.5*mm, 1.5*mm, label_width-3*mm, label_height-3*mm)
        else:
            # 舊批號標籤：單層邊框
            c.setLineWidth(0.5)
            c.rect(1*mm, 1*mm, label_width-2*mm, label_height-2*mm)
        
        # 標題（粗體）
        c.setFont(font_name, 10)
        c.drawString(2*mm, 29*mm, "【入庫】")
        
        # 入庫資料（垂直排列，粗體）
        c.setFont(font_name, 8)
        c.drawString(2*mm, 25*mm, f"試劑名稱：{entry.reagent_name}")
        
        # 試劑批號（根據是否新批號顯示不同內容）
        if is_new_batch:
            batch_text = f"試劑批號：{entry.reagent_batch_number} >>新批號<<"
        else:
            batch_text = f"試劑批號：{entry.reagent_batch_number} (允收合格)"
        c.drawString(2*mm, 21*mm, batch_text)
        
        c.drawString(2*mm, 17*mm, f"穩定效期：{entry.expiry_date.strftime('%Y/%m/%d')}")
        c.drawString(2*mm, 13*mm, f"入庫時間：{entry.entry_date.strftime('%Y/%m/%d')}")
        
        # 出庫標題（粗體）
        c.setFont(font_name, 10)
        c.drawString(2*mm, 8*mm, "【出庫】")
        
        # 出庫人員和日期（粗體）
        c.setFont(font_name, 8)
        c.drawString(2*mm, 4*mm, "人員：")
        c.drawString(25*mm, 4*mm, "出庫日期：")
        
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
            
            return quantity
        else:
            # Windows列印功能不可用，開啟PDF讓用戶手動處理
            subprocess.Popen(['start', temp_file.name], shell=True)
            
            return quantity
        
    except Exception as e:
        print(f"直接列印失敗: {e}")
        # 如果直接列印失敗，開啟PDF讓用戶手動處理
        subprocess.Popen(['start', temp_file.name], shell=True)
        
        return quantity

def generate_pdf_labels(entry, quantity=None, is_new_batch=False):
    """
    生成PDF格式的標籤
    """
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
        if is_new_batch:
            # 新批號標籤：雙重邊框
            # 外層邊框（粗邊框）
            c.setLineWidth(2)
            c.rect(0.5*mm, 0.5*mm, label_width-1*mm, label_height-1*mm)
            # 內層邊框（細邊框）
            c.setLineWidth(0.5)
            c.rect(1.5*mm, 1.5*mm, label_width-3*mm, label_height-3*mm)
        else:
            # 舊批號標籤：單層邊框
            c.setLineWidth(0.5)
            c.rect(1*mm, 1*mm, label_width-2*mm, label_height-2*mm)
        
        # 標題（粗體）
        c.setFont(font_name, 10)
        c.drawString(2*mm, 29*mm, "【入庫】")
        
        # 入庫資料（垂直排列，粗體）
        c.setFont(font_name, 8)
        c.drawString(2*mm, 25*mm, f"試劑名稱：{entry.reagent_name}")
        
        # 試劑批號（根據是否新批號顯示不同內容）
        if is_new_batch:
            batch_text = f"試劑批號：{entry.reagent_batch_number} >>新批號<<"
        else:
            batch_text = f"試劑批號：{entry.reagent_batch_number} (允收合格)"
        c.drawString(2*mm, 21*mm, batch_text)
        
        c.drawString(2*mm, 17*mm, f"穩定效期：{entry.expiry_date.strftime('%Y/%m/%d')}")
        c.drawString(2*mm, 13*mm, f"入庫時間：{entry.entry_date.strftime('%Y/%m/%d')}")
        
        # 出庫標題（粗體）
        c.setFont(font_name, 10)
        c.drawString(2*mm, 8*mm, "【出庫】")
        
        # 出庫人員和日期（留白給蓋章用，粗體）
        c.setFont(font_name, 8)
        c.drawString(2*mm, 4*mm, "人員：")
        c.drawString(25*mm, 4*mm, "出庫日期：")
        
        # 如果不是最後一頁，新增頁面
        if i < quantity - 1:
            c.showPage()
    
    c.save()
    
    # 嘗試使用Windows預設印表機直接列印
    try:
        print(f"已生成 {quantity} 張標籤PDF: {temp_file.name}")
        
        if WINDOWS_PRINT_AVAILABLE:
            print("正在嘗試直接列印...")
            
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

@app.route('/api/csv-template')
def download_csv_template():
    """下載CSV範本檔案"""
    try:
        template_path = os.path.join(APP_DIR, '試劑入庫紀錄_CSV範本.csv')
        
        # 如果檔案不存在，創建一個
        if not os.path.exists(template_path):
            headers = ['試劑名稱', '試劑批號', '穩定效期', '數量', '單位', '供應商', '入庫日期']
            sample_data = [
                ['GOT', 'GOT001', '2025-12-31', '10', '組', '亞培', '2025-08-22 09:00:00'],
                ['GPT', 'GPT001', '2025-12-31', '5', '組', '羅氏', '2025-08-22 10:00:00']
            ]
            
            with open(template_path, 'w', encoding='utf-8-sig', newline='') as f:
                import csv
                writer = csv.writer(f)
                writer.writerow(headers)
                for row in sample_data:
                    writer.writerow(row)
        
        return send_file(template_path, as_attachment=True, download_name='試劑入庫紀錄_CSV範本.csv')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/print-direct/<int:entry_id>', methods=['POST'])
def print_direct(entry_id):
    """直接列印到預設印表機"""
    try:
        entry = ReagentEntry.query.get_or_404(entry_id)
        data = request.json
        quantity = data.get('quantity', entry.quantity)
        is_new_batch = data.get('is_new_batch', False)
        printer_type = data.get('printer_type', 'zpl')  # 預設為ZPL模式
        
        # 如果前端沒有明確指定is_new_batch，則自動檢查
        if not is_new_batch:
            # 檢查此批號是否為第一次入庫（即是否為新批號）
            earliest_entry = ReagentEntry.query.filter_by(
                reagent_name=entry.reagent_name,
                reagent_batch_number=entry.reagent_batch_number
            ).order_by(ReagentEntry.entry_date.asc()).first()
            
            # 如果當前記錄就是最早的記錄，說明是新批號
            if earliest_entry and earliest_entry.id == entry.id:
                is_new_batch = True
        
        print(f"直接列印請求: 記錄ID {entry_id}, 數量 {quantity}, 新批號: {is_new_batch}, 標籤機類型: {printer_type}")
        
        # 根據印表機類型選擇列印方式
        if printer_type == "pdf":
            # PDF模式：使用原本的直接列印邏輯
            labels_printed = print_pdf_direct(entry, quantity, is_new_batch)
        else:
            # ZPL模式：使用統一的列印函數
            labels_printed = generate_and_print_labels(entry, quantity, is_new_batch, printer_type)
        
        return jsonify({
            'success': True,
            'message': f'已列印 {labels_printed} 張標籤'
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
        
        # 標題（粗體）
        c.setFont(font_name, 10)
        c.drawString(2*mm, 29*mm, "【入庫】")
        
        # 入庫資料（垂直排列，粗體）
        c.setFont(font_name, 8)
        c.drawString(2*mm, 25*mm, f"試劑名稱：{mock_entry.reagent_name}")
        c.drawString(2*mm, 21*mm, f"試劑批號：{mock_entry.reagent_batch_number} (允收合格)")
        c.drawString(2*mm, 17*mm, f"穩定效期：{mock_entry.expiry_date.strftime('%Y/%m/%d')}")
        c.drawString(2*mm, 13*mm, f"入庫時間：{mock_entry.entry_date.strftime('%Y/%m/%d')}")
        
        # 出庫標題（粗體）
        c.setFont(font_name, 10)
        c.drawString(2*mm, 8*mm, "【出庫】")
        
        # 出庫人員和日期（留白給蓋章用，粗體）
        c.setFont(font_name, 8)
        c.drawString(2*mm, 4*mm, "人員：")
        c.drawString(25*mm, 4*mm, "出庫日期：")
        
        c.save()
        
        # 返回PDF文件
        return send_file(temp_file.name, as_attachment=True, download_name='label_preview.pdf')
        
    except Exception as e:
        print(f"預覽標籤失敗: {e}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    import threading
    import time
    import webbrowser
    
    # 建立資料庫
    with app.app_context():
        db.create_all()
        print(f"資料庫路徑: {DB_PATH}")
    
    def open_browser():
        """延遲3秒後開啟瀏覽器"""
        time.sleep(3)
        webbrowser.open('http://127.0.0.1:5000')
    
    # 在背景開啟瀏覽器
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()
    
    print("試劑入庫管理系統啟動中...")
    print("瀏覽器將自動開啟 http://127.0.0.1:5000")
    print("若瀏覽器未自動開啟，請手動開啟上述網址")
    print("按 Ctrl+C 停止系統")
    
    try:
        app.run(debug=False, host='127.0.0.1', port=5000, use_reloader=False)
    except KeyboardInterrupt:
        print("\n系統已停止")
