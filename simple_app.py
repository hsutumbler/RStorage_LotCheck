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

# 嘗試導入 PIL/ImageFont 用於 ZPL 圖形模式
try:
    from PIL import Image, ImageDraw, ImageFont
    IMAGE_AVAILABLE = True
except ImportError:
    IMAGE_AVAILABLE = False
    print("警告：PIL/Pillow 套件未安裝，ZPL 圖形模式將不可用")

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

def find_sumatra_pdf():
    """
    尋找 SumatraPDF.exe，優先順序：
    1. 應用程式目錄內的 SumatraPDF.exe
    2. PyInstaller 臨時解壓目錄（如果有）
    3. 系統常見安裝路徑
    """
    app_dir = get_app_directory()
    
    # 優先檢查應用程式目錄
    local_sumatra = os.path.join(app_dir, 'SumatraPDF.exe')
    if os.path.exists(local_sumatra):
        print(f"找到應用程式目錄內的 SumatraPDF: {local_sumatra}")
        return local_sumatra
    
    # 如果是打包版本，檢查 _internal 目錄（PyInstaller 的資料目錄）
    if getattr(sys, 'frozen', False):
        internal_dir = os.path.join(app_dir, '_internal')
        internal_sumatra = os.path.join(internal_dir, 'SumatraPDF.exe')
        if os.path.exists(internal_sumatra):
            print(f"找到打包版本內的 SumatraPDF: {internal_sumatra}")
            return internal_sumatra
    
    # 檢查系統常見安裝路徑
    system_paths = [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
        os.path.expanduser(r"~\AppData\Local\SumatraPDF\SumatraPDF.exe"),
    ]
    
    for path in system_paths:
        if os.path.exists(path):
            print(f"找到系統安裝的 SumatraPDF: {path}")
            return path
    
    return None

# 設定資料庫路徑
APP_DIR = get_app_directory()

# ZPL 圖形模式設定（全域變數）
ZPL_CHINESE_FONT_PATH = None
ZPL_CHINESE_FONT_SIZE = 22  # 點數，與 ZPL 字型大小對應
ZPL_FIXED_GRAPHICS = {}  # 儲存固定文字的預定義圖形

def _load_chinese_font_for_zpl():
    """載入微軟正黑體供 ZPL 圖形模式使用"""
    global ZPL_CHINESE_FONT_PATH
    if not IMAGE_AVAILABLE:
        return False
    
    try:
        # 優先使用標準微軟正黑體，如果找不到則嘗試粗體版本
        font_paths = [
            "C:/Windows/Fonts/msjh.ttc",   # 微軟正黑體（標準版）
            "C:/Windows/Fonts/msjhbd.ttc", # 微軟正黑體（粗體版，備用）
        ]
        
        for font_path in font_paths:
            if os.path.exists(font_path):
                try:
                    # 測試是否可以載入字型
                    test_font = ImageFont.truetype(font_path, ZPL_CHINESE_FONT_SIZE)
                    ZPL_CHINESE_FONT_PATH = font_path
                    print(f"ZPL 圖形模式使用字型: {font_path} (微軟正黑體)")
                    return True
                except Exception as e:
                    print(f"無法載入字型 {font_path}: {e}")
                    continue
        
        print("警告：未找到微軟正黑體字型，ZPL 圖形模式可能無法正確顯示文字")
        return False
                
    except Exception as e:
        print(f"載入微軟正黑體失敗: {e}")
        return False

def _text_to_zpl_graphic(text, item_name, bold=False):
    """將文字轉換為 ZPL 圖形格式（~DGR 指令）
    
    Args:
        text: 要轉換的文字
        item_name: ZPL 圖形項目名稱（如 ITEM_NAME）
        bold: 是否使用粗體字型（預設為 False）
    
    Returns:
        ZPL 圖形指令字串（包含 ~DGR 定義），如果失敗則返回 None
    """
    if not IMAGE_AVAILABLE or not ZPL_CHINESE_FONT_PATH:
        return None
    
    try:
        # 載入字型（如果要求粗體，優先使用粗體字型）
        if bold:
            # 嘗試載入粗體字型
            bold_font_paths = [
                "C:/Windows/Fonts/msjhbd.ttc",  # 微軟正黑體粗體
                "C:/Windows/Fonts/msyhbd.ttc",  # 微軟雅黑粗體（備用）
            ]
            font = None
            for font_path in bold_font_paths:
                if os.path.exists(font_path):
                    try:
                        font = ImageFont.truetype(font_path, ZPL_CHINESE_FONT_SIZE)
                        break
                    except:
                        continue
            # 如果粗體字型載入失敗，使用一般字型
            if font is None:
                font = ImageFont.truetype(ZPL_CHINESE_FONT_PATH, ZPL_CHINESE_FONT_SIZE)
        else:
            font = ImageFont.truetype(ZPL_CHINESE_FONT_PATH, ZPL_CHINESE_FONT_SIZE)
        
        # 計算文字尺寸
        # 使用較大的臨時圖片來測量文字大小（包含上升和下降部分）
        temp_img = Image.new('RGB', (1000, 200), (255, 255, 255))
        temp_draw = ImageDraw.Draw(temp_img)
        bbox = temp_draw.textbbox((0, 0), text, font=font)
        
        # 邊界框：left, top, right, bottom
        # top 可能是負數（上升部分，如大寫字母），bottom 是正數
        text_left = bbox[0]
        text_top = bbox[1]  # 可能是負數
        text_right = bbox[2]
        text_bottom = bbox[3]
        
        text_width = text_right - text_left
        text_height = text_bottom - text_top
        
        # 增加邊距（避免邊緣裁切）
        padding = 6
        img_width = text_width + padding * 2
        # 圖形高度需要包含完整的文字高度（包括上升和下降部分）
        img_height = text_height + padding * 2
        
        # 建立 1-bit 黑白圖片
        img = Image.new('1', (img_width, img_height), 1)  # 1 = 白色
        draw = ImageDraw.Draw(img)
        
        # 計算文字繪製的 Y 座標
        # 因為 text_top 可能是負數，需要調整 Y 座標
        text_y = padding - text_top  # 這樣可以確保文字完整顯示
        
        # 繪製文字（0 = 黑色）
        draw.text((padding, text_y), text, font=font, fill=0)
        
        # 轉換為 ZPL HEX 格式
        hex_data = _image_to_zpl_hex(img)
        
        if not hex_data:
            return None
        
        # 計算總位元組數和每行列數
        bytes_per_row = (img_width + 7) // 8  # 每行需要幾個位元組（8位=1位元組）
        total_bytes = bytes_per_row * img_height
        
        # 生成 ~DGR 指令
        # 格式：~DGR:名稱,總位元組數,每行列數,壓縮的HEX資料
        zpl_command = f"~DGR:{item_name},{total_bytes:05d},{bytes_per_row:03d},{hex_data}"
        
        return zpl_command
        
    except Exception as e:
        print(f"將文字轉換為 ZPL 圖形失敗: {text}, 錯誤: {e}")
        import traceback
        traceback.print_exc()
        return None

def _image_to_zpl_hex(img):
    """將 PIL 圖片轉換為 ZPL HEX 格式
    
    ZPL ~DGR 指令使用的格式：
    - 每個位元組代表 8 個像素（水平方向，從左到右）
    - 位元 7（最高位）對應最左側的像素
    - 1 = 黑色，0 = 白色
    
    PIL '1' 模式：
    - 0 = 黑色，1 = 白色
    - 需要反轉
    """
    try:
        width, height = img.size
        bytes_per_row = (width + 7) // 8
        
        # 取得圖片的像素數據
        pixels = img.load()
        
        # 轉換為位元組陣列
        hex_chars = []
        for y in range(height):
            for x_byte in range(bytes_per_row):
                byte_value = 0
                for bit in range(8):
                    pixel_x = x_byte * 8 + bit
                    if pixel_x < width:
                        # PIL 的 '1' 模式：0=黑色, 1=白色
                        # ZPL 需要：0=白色, 1=黑色
                        # 所以需要反轉：如果 PIL 是黑色(0)，ZPL 設為 1
                        pixel_value = pixels[pixel_x, y]
                        if pixel_value == 0:  # PIL 黑色像素
                            # ZPL 設為 1（黑色）
                            # 位元順序：最高位（bit 7）對應最左側像素
                            byte_value |= (1 << (7 - bit))
                        # 如果 pixel_value == 1（PIL 白色），ZPL 設為 0（白色），不需要操作
                
                # 轉換為 HEX（兩位數，大寫）
                hex_chars.append(f"{byte_value:02X}")
        
        # 合併所有 HEX 字元
        hex_string = ''.join(hex_chars)
        
        return hex_string
        
    except Exception as e:
        print(f"圖片轉換為 ZPL HEX 失敗: {e}")
        import traceback
        traceback.print_exc()
        return None

def _generate_fixed_graphics():
    """生成固定文字的預定義圖形（使用微軟正黑體）"""
    global ZPL_FIXED_GRAPHICS
    if not IMAGE_AVAILABLE or not ZPL_CHINESE_FONT_PATH:
        print("警告：無法生成固定圖形：缺少 PIL 或微軟正黑體字型")
        return
    
    try:
        # 固定文字的列表
        fixed_texts = {
            "IN": "【入庫】",
            "OUT": "【出庫】",
            "REAGENT_NAME": "試劑名稱:",
            "BATCH": "試劑批號:",
            "EXPIRY": "穩定效期:",
            "ENTRY_DATE": "入庫日期:",
            "PERSON": "人員:",
            "CHECKOUT_DATE": "出庫日期:",
            "NEW_BATCH": ">>新批號<<",
            "QUALIFIED": "(允收合格)"
        }
        
        for key, text in fixed_texts.items():
            # 新批號標記使用粗體
            is_bold = (key == "NEW_BATCH")
            zpl_graphic = _text_to_zpl_graphic(text, f"ITEM_{key}", bold=is_bold)
            if zpl_graphic:
                ZPL_FIXED_GRAPHICS[key] = zpl_graphic
        
        print(f"成功生成 {len(ZPL_FIXED_GRAPHICS)} 個固定圖形")
        
    except Exception as e:
        print(f"生成固定圖形失敗: {e}")
        import traceback
        traceback.print_exc()

# 初始化 ZPL 圖形模式
if IMAGE_AVAILABLE:
    _load_chinese_font_for_zpl()
    _generate_fixed_graphics()

# 在打包版本中，資料庫放在database資料夾中
if getattr(sys, 'frozen', False):
    # 打包版本：使用database資料夾
    db_dir = os.path.join(APP_DIR, 'database')
    if not os.path.exists(db_dir):
        os.makedirs(db_dir)
    DB_PATH = os.path.join(db_dir, 'simple_inventory.db')
else:
    # 開發版本：直接放在程式目錄
    DB_PATH = os.path.join(APP_DIR, 'simple_inventory.db')

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# 資料庫模型
class Supplier(db.Model):
    """供應商資料表"""
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)  # 供應商名稱
    created_at = db.Column(db.DateTime, default=datetime.utcnow)  # 建立時間

class ReagentEntry(db.Model):
    """試劑入庫記錄資料表"""
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

@app.route('/api/suppliers', methods=['GET'])
def get_suppliers():
    """獲取所有供應商列表"""
    try:
        suppliers = Supplier.query.order_by(Supplier.name).all()
        return jsonify([{'id': s.id, 'name': s.name} for s in suppliers])
    except Exception as e:
        print(f"獲取供應商列表失敗: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/suppliers', methods=['POST'])
def add_supplier():
    """新增供應商"""
    try:
        data = request.json
        supplier_name = data.get('name')
        
        # 檢查是否已存在
        existing = Supplier.query.filter_by(name=supplier_name).first()
        if existing:
            return jsonify({'success': True, 'message': '供應商已存在', 'supplier': {'id': existing.id, 'name': existing.name}})
        
        # 新增供應商
        new_supplier = Supplier(name=supplier_name)
        db.session.add(new_supplier)
        db.session.commit()
        
        return jsonify({'success': True, 'supplier': {'id': new_supplier.id, 'name': new_supplier.name}})
    except Exception as e:
        print(f"新增供應商失敗: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/entries', methods=['GET'])
def get_entries():
    try:
        # 先建立資料庫
        with app.app_context():
            db.create_all()
        
        # 查詢記錄（限制只返回最近50筆）
        entries = ReagentEntry.query.order_by(ReagentEntry.entry_date.desc()).limit(50).all()
        print(f"找到 {len(entries)} 筆記錄（最近50筆）")
        
        # 獲取總記錄數（用於前端顯示）
        total_count = ReagentEntry.query.count()
        
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
        
        print(f"返回 {len(result)} 筆記錄，資料庫共有 {total_count} 筆")
        return jsonify({
            'entries': result,
            'totalCount': total_count,
            'limitApplied': len(entries) < total_count
        })
        
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
        
        # 不在此處自動列印，由前端控制列印時機
        print(f"入庫完成，等待前端列印指令")
        
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
        
        # 不在此處自動列印，由前端控制列印時機
        print(f"新批號入庫完成，等待前端列印指令")
        
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
            return jsonify({'entries': [], 'totalCount': 0, 'isSearchResult': True})
        
        # 處理特殊查詢「all_records」- 用於日期篩選
        if query == 'all_records':
            print("執行全部記錄查詢（用於日期篩選）")
            entries = ReagentEntry.query.order_by(ReagentEntry.entry_date.desc()).all()
        else:
            # 一般搜尋 - 不限制筆數，可以找到所有符合條件的記錄
            entries = ReagentEntry.query.filter(
                db.or_(
                    ReagentEntry.reagent_name.contains(query),
                    ReagentEntry.reagent_batch_number.contains(query),
                    ReagentEntry.supplier.contains(query)
                )
            ).order_by(ReagentEntry.entry_date.desc()).all()
        
        # 計算總記錄數
        total_count = len(entries)
        print(f"搜尋 '{query}' 找到 {total_count} 筆記錄")
        
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
        
        return jsonify({
            'entries': result,
            'totalCount': total_count,
            'isSearchResult': True
        })
        
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
    生成ZPL格式的標籤指令 (使用圖形模式，確保中文正確顯示)
    - entry: 資料庫記錄
    - quantity: 列印數量
    - is_new_batch: 是否為新批號
    """
    if quantity is None:
        quantity = entry.quantity
    
    zpl_commands = []
    
    # 標籤尺寸：5cm x 3.5cm (203 DPI)
    # 50mm = 394 dots, 35mm = 276 dots
    label_width_dots = 394  # 50mm at 203 DPI (正確計算：50 * 203 / 25.4 ≈ 394)
    label_height_dots = 276  # 35mm at 203 DPI (正確計算：35 * 203 / 25.4 ≈ 276)
    
    # 準備動態文字的圖形
    dynamic_graphics = {}  # 儲存圖形定義的 ZPL 指令
    dynamic_graphic_names = {}  # 儲存圖形項目名稱
    
    # 生成動態文字的圖形
    reagent_name = entry.reagent_name
    batch_number = entry.reagent_batch_number
    expiry_str = entry.expiry_date.strftime('%Y/%m/%d')
    entry_str = entry.entry_date.strftime('%Y/%m/%d')
    
    # 如果固定圖形「試劑名稱:」失敗，為標籤也生成動態圖形
    if reagent_name and "REAGENT_NAME" not in ZPL_FIXED_GRAPHICS:
        label_item_name = "ITEM_REAGENT_NAME_LABEL_DYN"
        label_graphic = _text_to_zpl_graphic("試劑名稱:", label_item_name)
        if label_graphic:
            dynamic_graphics['reagent_name_label'] = label_graphic
            dynamic_graphic_names['reagent_name_label'] = label_item_name
    
    # 生成試劑名稱圖形
    if reagent_name:
        item_name = f"ITEM_REAGENT_NAME_DYN_{hash(reagent_name) % 10000}"
        name_graphic = _text_to_zpl_graphic(reagent_name, item_name)
        if name_graphic:
            dynamic_graphics['reagent_name'] = name_graphic
            dynamic_graphic_names['reagent_name'] = item_name
    
    # 生成批號圖形
    if batch_number:
        item_name = f"ITEM_BATCH_DYN_{hash(batch_number) % 10000}"
        batch_graphic = _text_to_zpl_graphic(batch_number, item_name)
        if batch_graphic:
            dynamic_graphics['batch_number'] = batch_graphic
            dynamic_graphic_names['batch_number'] = item_name
    
    # 生成日期圖形
    if expiry_str:
        item_name = f"ITEM_EXPIRY_DYN_{hash(expiry_str) % 10000}"
        expiry_graphic = _text_to_zpl_graphic(expiry_str, item_name)
        if expiry_graphic:
            dynamic_graphics['expiry'] = expiry_graphic
            dynamic_graphic_names['expiry'] = item_name
    
    if entry_str:
        item_name = f"ITEM_ENTRY_DYN_{hash(entry_str) % 10000}"
        entry_graphic = _text_to_zpl_graphic(entry_str, item_name)
        if entry_graphic:
            dynamic_graphics['entry_date'] = entry_graphic
            dynamic_graphic_names['entry_date'] = item_name
    
    for i in range(quantity):
        zpl = "^XA\n"  # 開始標籤
        
        # 設定標籤尺寸 (以點為單位，203 DPI)
        zpl += f"^PW{label_width_dots}\n"  # 設定列印寬度
        zpl += f"^LL{label_height_dots}\n"  # 設定標籤長度
        
        # 只有第一張標籤是新批號格式（當is_new_batch為True時）
        is_first_label = is_new_batch and i == 0
        
        # 繪製邊框
        border_thickness = 2
        if is_first_label:
            # 新批號：雙邊框
            # 外層邊框 (粗線)
            outer_x, outer_y = 5, 5
            outer_w = label_width_dots - 10
            outer_h = label_height_dots - 10
            zpl += f"^FO{outer_x},{outer_y}^GB{outer_w},{outer_h},{border_thickness * 3}^FS\n"
            # 內層邊框 (細線)
            inner_x, inner_y = 10, 10
            inner_w = label_width_dots - 20
            inner_h = label_height_dots - 20
            zpl += f"^FO{inner_x},{inner_y}^GB{inner_w},{inner_h},{border_thickness}^FS\n"
        else:
            # 一般標籤：單邊框
            border_x, border_y = 5, 5
            border_w = label_width_dots - 10
            border_h = label_height_dots - 10
            zpl += f"^FO{border_x},{border_y}^GB{border_w},{border_h},{border_thickness}^FS\n"
        
        # ========== 使用圖形模式顯示文字 ==========
        # 先定義固定圖形（如果可用）
        for key, graphic_def in ZPL_FIXED_GRAPHICS.items():
            zpl += graphic_def + "\n"
        
        # 定義動態圖形（每個標籤都需要定義一次）
        for key, graphic_def in dynamic_graphics.items():
            zpl += graphic_def + "\n"
        
        y_pos = 25  # 起始Y位置
        
        # 【入庫】標題
        if "IN" in ZPL_FIXED_GRAPHICS:
            zpl += f"^FO20,{y_pos}^XGITEM_IN^FS\n"
        else:
            # 備用方案：使用字型
            zpl += f"^FO20,{y_pos}^A0N,22,22^FD【入庫】^FS\n"
        y_pos += 30
        
        # 試劑名稱
        if "REAGENT_NAME" in ZPL_FIXED_GRAPHICS:
            zpl += f"^FO20,{y_pos}^XGITEM_REAGENT_NAME^FS\n"
            label_width = 80  # 估算「試劑名稱:」的寬度
            char_spacing = 22  # 一個中文字距離
            if 'reagent_name' in dynamic_graphic_names:
                zpl += f"^FO{20 + label_width + char_spacing},{y_pos}^XG{dynamic_graphic_names['reagent_name']}^FS\n"
            else:
                # 備用方案：使用字型
                zpl += f"^FO{20 + label_width + char_spacing},{y_pos}^A0N,22,22^FD{reagent_name}^FS\n"
        else:
            # 如果沒有固定圖形，使用動態圖形或字型
            # 先組合完整文字（比照試劑批號的方式）
            reagent_text = f"試劑名稱:{reagent_name}"
            if 'reagent_name' in dynamic_graphic_names:
                # 如果動態圖形成功，使用動態圖形顯示「試劑名稱:」標籤（如果有的話），再顯示試劑名稱圖形
                if 'reagent_name_label' in dynamic_graphic_names:
                    # 使用動態圖形顯示「試劑名稱:」標籤
                    zpl += f"^FO20,{y_pos}^XG{dynamic_graphic_names['reagent_name_label']}^FS\n"
                    label_width = 80  # 估算「試劑名稱:」的寬度
                    char_spacing = 22  # 一個中文字距離
                    zpl += f"^FO{20 + label_width + char_spacing},{y_pos}^XG{dynamic_graphic_names['reagent_name']}^FS\n"
                else:
                    # 如果標籤動態圖形也失敗，使用字型顯示標籤（可能無法顯示中文）
                    zpl += f"^FO20,{y_pos}^A0N,22,22^FD試劑名稱:^FS\n"
                    label_width = 80  # 估算「試劑名稱:」的寬度
                    char_spacing = 22  # 一個中文字距離
                    zpl += f"^FO{20 + label_width + char_spacing},{y_pos}^XG{dynamic_graphic_names['reagent_name']}^FS\n"
            else:
                # 如果動態圖形也失敗，顯示完整文字
                zpl += f"^FO20,{y_pos}^A0N,22,22^FD{reagent_text}^FS\n"
        y_pos += 30
        
        # 試劑批號
        if "BATCH" in ZPL_FIXED_GRAPHICS:
            zpl += f"^FO20,{y_pos}^XGITEM_BATCH^FS\n"
            label_width = 80  # 估算「試劑批號:」的寬度
            char_spacing = 22  # 一個中文字距離
            if 'batch_number' in dynamic_graphic_names:
                zpl += f"^FO{20 + label_width + char_spacing},{y_pos}^XG{dynamic_graphic_names['batch_number']}^FS\n"
            else:
                zpl += f"^FO{20 + label_width + char_spacing},{y_pos}^A0N,22,22^FD{batch_number}^FS\n"
            
            # 新批號標記或允收合格標記
            if is_first_label:
                if "NEW_BATCH" in ZPL_FIXED_GRAPHICS:
                    # 計算批號圖形的寬度（約100點），然後顯示新批號標記
                    zpl += f"^FO{20 + label_width + char_spacing + 100},{y_pos}^XGITEM_NEW_BATCH^FS\n"
                else:
                    # 使用粗體字型顯示新批號標記
                    zpl += f"^FO{20 + label_width + char_spacing + 100},{y_pos}^A0B,22,22^FD>>新批號<<^FS\n"
            else:
                if "QUALIFIED" in ZPL_FIXED_GRAPHICS:
                    zpl += f"^FO{20 + label_width + char_spacing + 100},{y_pos}^XGITEM_QUALIFIED^FS\n"
                else:
                    zpl += f"^FO{20 + label_width + char_spacing + 100},{y_pos}^A0N,22,22^FD(允收合格)^FS\n"
        else:
            # 如果沒有固定圖形，使用動態圖形或字型
            if is_first_label:
                batch_text = f"試劑批號:{batch_number} >>新批號<<"
                # 使用粗體字型顯示新批號標記
                if 'batch_number' in dynamic_graphic_names:
                    zpl += f"^FO20,{y_pos}^A0N,22,22^FD試劑批號:^FS\n"
                    char_spacing = 22  # 一個中文字距離
                    zpl += f"^FO{20 + 80 + char_spacing},{y_pos}^XG{dynamic_graphic_names['batch_number']}^FS\n"
                    zpl += f"^FO{20 + 80 + char_spacing + 100},{y_pos}^A0B,22,22^FD>>新批號<<^FS\n"
                else:
                    zpl += f"^FO20,{y_pos}^A0N,22,22^FD試劑批號:{batch_number} ^FS\n"
                    # 計算批號文字的寬度後顯示粗體新批號標記
                    zpl += f"^FO{20 + 80 + 22 + len(batch_number) * 11},{y_pos}^A0B,22,22^FD>>新批號<<^FS\n"
            else:
                batch_text = f"試劑批號:{batch_number} (允收合格)"
                if 'batch_number' in dynamic_graphic_names:
                    zpl += f"^FO20,{y_pos}^XG{dynamic_graphic_names['batch_number']}^FS\n"
                else:
                    zpl += f"^FO20,{y_pos}^A0N,22,22^FD{batch_text}^FS\n"
        y_pos += 30
        
        # 穩定效期
        if "EXPIRY" in ZPL_FIXED_GRAPHICS:
            zpl += f"^FO20,{y_pos}^XGITEM_EXPIRY^FS\n"
            label_width = 80
            char_spacing = 22  # 一個中文字距離
            if 'expiry' in dynamic_graphic_names:
                zpl += f"^FO{20 + label_width + char_spacing},{y_pos}^XG{dynamic_graphic_names['expiry']}^FS\n"
            else:
                zpl += f"^FO{20 + label_width + char_spacing},{y_pos}^A0N,22,22^FD{expiry_str}^FS\n"
        else:
            if 'expiry' in dynamic_graphic_names:
                zpl += f"^FO20,{y_pos}^A0N,22,22^FD穩定效期:^FS\n"
                char_spacing = 22  # 一個中文字距離
                zpl += f"^FO{20 + 80 + char_spacing},{y_pos}^XG{dynamic_graphic_names['expiry']}^FS\n"
            else:
                zpl += f"^FO20,{y_pos}^A0N,22,22^FD穩定效期:{expiry_str}^FS\n"
        y_pos += 30
        
        # 入庫日期
        if "ENTRY_DATE" in ZPL_FIXED_GRAPHICS:
            zpl += f"^FO20,{y_pos}^XGITEM_ENTRY_DATE^FS\n"
            label_width = 80
            char_spacing = 22  # 一個中文字距離
            if 'entry_date' in dynamic_graphic_names:
                zpl += f"^FO{20 + label_width + char_spacing},{y_pos}^XG{dynamic_graphic_names['entry_date']}^FS\n"
            else:
                zpl += f"^FO{20 + label_width + char_spacing},{y_pos}^A0N,22,22^FD{entry_str}^FS\n"
        else:
            if 'entry_date' in dynamic_graphic_names:
                zpl += f"^FO20,{y_pos}^A0N,22,22^FD入庫日期:^FS\n"
                char_spacing = 22  # 一個中文字距離
                zpl += f"^FO{20 + 80 + char_spacing},{y_pos}^XG{dynamic_graphic_names['entry_date']}^FS\n"
            else:
                zpl += f"^FO20,{y_pos}^A0N,22,22^FD入庫日期:{entry_str}^FS\n"
        y_pos += 30
        
        # 【出庫】標題
        if "OUT" in ZPL_FIXED_GRAPHICS:
            zpl += f"^FO20,{y_pos}^XGITEM_OUT^FS\n"
        else:
            zpl += f"^FO20,{y_pos}^A0N,22,22^FD【出庫】^FS\n"
        y_pos += 30
        
        # 出庫資訊
        if "PERSON" in ZPL_FIXED_GRAPHICS:
            zpl += f"^FO20,{y_pos}^XGITEM_PERSON^FS\n"
        else:
            zpl += f"^FO20,{y_pos}^A0N,22,22^FD人員^FS\n"
        
        if "CHECKOUT_DATE" in ZPL_FIXED_GRAPHICS:
            zpl += f"^FO190,{y_pos}^XGITEM_CHECKOUT_DATE^FS\n"
        else:
            zpl += f"^FO190,{y_pos}^A0N,22,22^FD出庫日期^FS\n"
        
        zpl += "^XZ\n"  # 結束標籤
        
        zpl_commands.append(zpl)
    
    return zpl_commands

def send_zpl_to_printer(zpl_commands):
    """
    發送ZPL指令到Zebra印表機 (支援UTF-8和Unicode中文字型)
    """
    try:
        if not WINDOWS_PRINT_AVAILABLE:
            print("Windows列印功能不可用，無法發送ZPL指令")
            return False
        
        # 獲取預設印表機
        default_printer = win32print.GetDefaultPrinter()
        print(f"發送ZPL指令到 Zebra 印表機: {default_printer}")
        
        # 開啟印表機
        printer_handle = win32print.OpenPrinter(default_printer)
        
        try:
            # 開始列印工作 (使用 RAW 模式直接發送ZPL指令)
            job_id = win32print.StartDocPrinter(printer_handle, 1, ("ZPL Label UTF-8", None, "RAW"))
            win32print.StartPagePrinter(printer_handle)
            
            # 發送每個ZPL指令 (確保UTF-8編碼)
            for i, zpl in enumerate(zpl_commands):
                # 確保ZPL指令使用UTF-8編碼
                # 這樣印表機可以正確解析 ^CI28 指令並使用 Unicode 尋找字形
                zpl_bytes = zpl.encode('utf-8')
                win32print.WritePrinter(printer_handle, zpl_bytes)
                print(f"已發送第 {i+1} 張標籤的ZPL指令 (UTF-8編碼)")
            
            # 結束列印
            win32print.EndPagePrinter(printer_handle)
            win32print.EndDocPrinter(printer_handle)
            
            print(f"ZPL指令發送成功，共 {len(zpl_commands)} 張標籤")
            print("注意：請確保 Zebra 印表機已正確設定並支援 UTF-8/Unicode 編碼")
            return True
            
        finally:
            win32print.ClosePrinter(printer_handle)
    
    except Exception as e:
        print(f"發送ZPL指令失敗: {e}")
        import traceback
        traceback.print_exc()
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
    # 為了避免列印時旋轉90度，將頁面設為直向（高度 > 寬度）
    label_width = 50 * mm  # 5cm（實際標籤寬度）
    label_height = 35 * mm  # 3.5cm（實際標籤高度）
    
    # 建立標籤PDF
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    # 創建PDF文檔（交換寬高，設為直向以避免列印時旋轉）
    c = canvas.Canvas(temp_file.name, pagesize=(label_height, label_width))
    
    # 獲取中文字體
    font_name = get_chinese_font()
    
    for i in range(quantity):
        # 旋轉畫布90度，讓內容正確顯示在直向頁面上
        # 先平移再旋轉，確保內容顯示在正確位置
        c.translate(label_height, 0)
        c.rotate(90)
        
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
        
        # 重置變換矩陣，為下一頁做準備
        c.resetTransforms()
        
        if i < quantity - 1:
            c.showPage()
    
    c.save()
    
    # 嘗試直接列印到預設印表機
    try:
        if WINDOWS_PRINT_AVAILABLE:
            default_printer = win32print.GetDefaultPrinter()
            print(f"直接列印到: {default_printer}")
            
            # 優先使用應用程式目錄內的 SumatraPDF（支援靜默列印）
            sumatra_path = find_sumatra_pdf()
            
            if sumatra_path:
                print(f"使用SumatraPDF靜默列印: {sumatra_path}")
                try:
                    # SumatraPDF 不支援直接設定列印方向，但我們可以確保 PDF 本身的方向正確
                    # 使用基本參數列印，PDF 的方向應該由 PDF 本身的頁面大小決定
                    subprocess.run([
                        sumatra_path, 
                        "-print-to-default", 
                        "-silent",
                        temp_file.name
                    ], check=True, timeout=30)
                    print("SumatraPDF 靜默列印命令已發送")
                    print("注意：如果列印方向不正確，請檢查印表機驅動的預設設定")
                    return quantity
                except Exception as e:
                    print(f"SumatraPDF 列印失敗: {e}")
            
            # 如果 SumatraPDF 不可用，使用系統命令列印PDF
            print("使用系統預設方式列印（可能不是靜默列印）")
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
    print(f"生成標籤 - 數量: {quantity}, 新批號: {is_new_batch}")
    
    if quantity is None:
        quantity = entry.quantity
    
    # 標籤尺寸：5cm x 3.5cm
    # 為了避免列印時旋轉90度，將頁面設為直向（高度 > 寬度）
    label_width = 50 * mm  # 5cm（實際標籤寬度）
    label_height = 35 * mm  # 3.5cm（實際標籤高度）
    
    # 建立標籤PDF
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    
    # 創建PDF文檔（交換寬高，設為直向以避免列印時旋轉）
    doc = canvas.Canvas(temp_file.name, pagesize=(label_height, label_width))
    
    # 獲取中文字體
    font_name = get_chinese_font()
    
    # 繪製每一頁標籤
    for i in range(quantity):
        print(f"正在生成第 {i+1} 張標籤...")
        
        # 只有第一張標籤是新批號格式（當is_new_batch為True時）
        is_first_label = is_new_batch and i == 0
        
        # 旋轉畫布90度，讓內容正確顯示在直向頁面上
        # 先平移再旋轉，確保內容顯示在正確位置
        doc.translate(label_height, 0)
        doc.rotate(90)
        
        # 繪製邊框
        if is_first_label:
            # 新批號標籤：雙重邊框
            # 外層邊框（粗邊框）
            doc.setLineWidth(2)
            doc.rect(0.5*mm, 0.5*mm, label_width-1*mm, label_height-1*mm)
            # 內層邊框（細邊框）
            doc.setLineWidth(0.5)
            doc.rect(1.5*mm, 1.5*mm, label_width-3*mm, label_height-3*mm)
        else:
            # 一般標籤：單層邊框
            doc.setLineWidth(0.5)
            doc.rect(1*mm, 1*mm, label_width-2*mm, label_height-2*mm)
        
        # 標題（粗體）
        doc.setFont(font_name, 10)
        doc.drawString(2*mm, 29*mm, "【入庫】")
        
        # 入庫資料（垂直排列，粗體）
        doc.setFont(font_name, 8)
        doc.drawString(2*mm, 25*mm, f"試劑名稱：{entry.reagent_name}")
        
        # 試劑批號（根據是否第一張決定顯示方式）
        if is_first_label:
            batch_text = f"試劑批號：{entry.reagent_batch_number} >>新批號<<"
        else:
            batch_text = f"試劑批號：{entry.reagent_batch_number} (允收合格)"
        doc.drawString(2*mm, 21*mm, batch_text)
        
        # 其他資訊
        doc.drawString(2*mm, 17*mm, f"穩定效期：{entry.expiry_date.strftime('%Y/%m/%d')}")
        doc.drawString(2*mm, 13*mm, f"入庫時間：{entry.entry_date.strftime('%Y/%m/%d')}")
        
        # 出庫標題（粗體）
        doc.setFont(font_name, 10)
        doc.drawString(2*mm, 8*mm, "【出庫】")
        
        # 出庫人員和日期（留白給蓋章用，粗體）
        doc.setFont(font_name, 8)
        doc.drawString(2*mm, 4*mm, "人員：")
        doc.drawString(25*mm, 4*mm, "出庫日期：")
        
        # 重置變換矩陣，為下一頁做準備
        doc.resetTransforms()
        
        # 如果不是最後一頁，則新增頁面
        if i < quantity - 1:
            doc.showPage()
    
    # 儲存PDF
    doc.save()
    print(f"標籤生成完成，檔案位置: {temp_file.name}")
    
    try:
        print(f"已生成 {quantity} 張標籤PDF: {temp_file.name}")
        
        if WINDOWS_PRINT_AVAILABLE:
            print("正在嘗試直接列印...")
            
            # 獲取預設印表機名稱
            default_printer = win32print.GetDefaultPrinter()
            print(f"使用預設印表機: {default_printer}")
            
            # 優先使用應用程式目錄內的 SumatraPDF（支援靜默列印）
            sumatra_path = find_sumatra_pdf()
            
            if sumatra_path:
                print(f"使用SumatraPDF靜默列印: {sumatra_path}")
                try:
                    # SumatraPDF 不支援直接設定列印方向
                    # PDF 的方向由 PDF 本身的頁面大小決定（寬度 > 高度 = 橫向）
                    # 如果列印方向不正確，可能是印表機驅動的自動旋轉功能造成的
                    subprocess.run([
                        sumatra_path, 
                        "-print-to-default", 
                        "-silent",
                        temp_file.name
                    ], check=True, timeout=30)
                    print("SumatraPDF 靜默列印命令已發送")
                    print("注意：如果列印方向不正確，請在印表機驅動設定中關閉自動旋轉功能")
                    return quantity
                except subprocess.TimeoutExpired:
                    print("警告：SumatraPDF 列印超時")
                except subprocess.CalledProcessError as e:
                    print(f"警告：SumatraPDF 列印失敗: {e}")
                except Exception as e:
                    print(f"警告：SumatraPDF 列印發生錯誤: {e}")
            
            # 如果 SumatraPDF 不可用，嘗試使用 Adobe Reader
            adobe_path = r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
            if os.path.exists(adobe_path):
                print("使用Adobe Reader列印...")
                try:
                    subprocess.run([adobe_path, "/T", temp_file.name, default_printer], 
                                  check=True, timeout=30)
                    print("Adobe Reader 列印命令已發送")
                    return quantity
                except Exception as e:
                    print(f"警告：Adobe Reader 列印失敗: {e}")
            
            # 如果都不可用，使用系統預設方式（可能不是靜默列印）
            print("找不到支援靜默列印的PDF閱讀器，使用系統預設程式...")
            print("注意：此方式可能會顯示列印對話框")
            win32api.ShellExecute(0, "print", temp_file.name, None, ".", 0)
            print("列印命令已發送")
        else:
            print("Windows列印功能不可用，改為開啟PDF檔案")
            # 使用系統預設PDF查看器開啟
            subprocess.Popen(['start', temp_file.name], shell=True)
            print(f"PDF已開啟，請手動選擇列印或儲存")
    except Exception as e:
        print(f"列印失敗: {e}")
        print("改為開啟PDF檔案供手動列印")
        try:
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
        printer_type = data.get('printer_type', 'pdf')  # 預設為PDF模式
        
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
        
        # 根據列印模式選擇對應的列印函數
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
