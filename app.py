import os
import re
import sqlite3
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo
from flask import Flask, render_template, request, jsonify, send_file, g
from werkzeug.utils import secure_filename
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font
import io

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# ---------- 数据库 ----------
DATABASE = 'scan_status.db'

def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DATABASE)
        db.row_factory = sqlite3.Row
    return db

def init_db():
    with app.app_context():
        db = get_db()
        db.execute('''
            CREATE TABLE IF NOT EXISTS scan_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                recipient TEXT NOT NULL,
                barcode TEXT NOT NULL,
                scanner_name TEXT,
                scanned_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(recipient, barcode)
            )
        ''')
        # 如果 scanner_name 列不存在则添加（兼容旧数据库）
        try:
            db.execute('ALTER TABLE scan_records ADD COLUMN scanner_name TEXT')
        except sqlite3.OperationalError:
            pass
        db.execute('CREATE INDEX IF NOT EXISTS idx_recipient ON scan_records (recipient)')
        db.commit()

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

# ---------- 辅助函数 ----------
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_custom_base_code(code):
    if not isinstance(code, str):
        return code
    match = re.search(r'^(.*)U(\d{6})$', code)
    if match:
        return match.group(1)
    return code

def generate_serial_list(base_code, count):
    return [f"{base_code}U{str(i+1).zfill(6)}" for i in range(count)]

def parse_excel(file_path):
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    sheet = wb['按箱出库'] if '按箱出库' in wb.sheetnames else wb.active

    header = [cell.value for cell in sheet[1]]
    try:
        recipient_col_idx = header.index('Recipient/收件人') + 1
        fba_col_idx = header.index('FBA货件ID/FBAShipmentID') + 1
        custom_col_idx = header.index('Custom box barcode/自定义箱条码') + 1
    except ValueError as e:
        raise Exception(f"未找到必要的列: {e}")

    data_rows = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            continue
        recipient = row[recipient_col_idx - 1] if len(row) >= recipient_col_idx else None
        fba_id = row[fba_col_idx - 1] if len(row) >= fba_col_idx else None
        custom_barcode = row[custom_col_idx - 1] if len(row) >= custom_col_idx else None
        data_rows.append({
            'recipient': recipient,
            'fba_id': fba_id,
            'custom_barcode': custom_barcode
        })

    recipients = set()
    for row in data_rows:
        if row['recipient']:
            recipients.add(str(row['recipient']).strip())
    recipients = list(recipients)

    result = {}
    for rec in recipients:
        sub_rows = [r for r in data_rows if r['recipient'] and str(r['recipient']).strip() == rec]

        # FBA
        fba_list = [r['fba_id'] for r in sub_rows if r['fba_id'] and str(r['fba_id']).strip()]
        fba_counts = defaultdict(int)
        for fba in fba_list:
            fba_counts[str(fba).strip()] += 1
        fba_items = []
        for base_code, count in fba_counts.items():
            serials = generate_serial_list(base_code, count)
            fba_items.append({
                'master_code': base_code,
                'type': 'FBA',
                'total_boxes': count,
                'serials': serials
            })

        # Custom (FBA为空)
        empty_fba_rows = [r for r in sub_rows if not r['fba_id'] or str(r['fba_id']).strip() == '']
        custom_list = [r['custom_barcode'] for r in empty_fba_rows if r['custom_barcode'] and str(r['custom_barcode']).strip()]
        custom_base_counts = defaultdict(int)
        for code in custom_list:
            base = extract_custom_base_code(str(code).strip())
            custom_base_counts[base] += 1
        custom_items = []
        for base_code, count in custom_base_counts.items():
            serials = generate_serial_list(base_code, count)
            custom_items.append({
                'master_code': base_code,
                'type': 'Custom',
                'total_boxes': count,
                'serials': serials
            })

        result[rec] = {
            'fba_items': fba_items,
            'custom_items': custom_items,
            'all_items': fba_items + custom_items
        }

    return result, recipients

# ---------- 全局缓存 ----------
latest_data = None

# ---------- 路由 ----------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    global latest_data
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if not allowed_file(file.filename):
        return jsonify({'error': 'File type not allowed'}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    try:
        data, recipients = parse_excel(filepath)
        latest_data = data
        return jsonify({
            'success': True,
            'recipients': recipients,
            'data': data
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/get_data', methods=['GET'])
def get_data():
    global latest_data
    if latest_data is None:
        return jsonify({'success': False, 'message': '暂无数据，请先上传出库表'}), 404
    recipients = list(latest_data.keys())
    return jsonify({
        'success': True,
        'recipients': recipients,
        'data': latest_data
    })

@app.route('/scan', methods=['POST'])
def scan_barcode():
    req = request.json
    recipient = req.get('recipient')
    barcode = req.get('barcode')
    scanner_name = req.get('scanner_name', '').strip()
    if not recipient or not barcode:
        return jsonify({'error': 'Missing parameters'}), 400
    if not scanner_name:
        return jsonify({'error': '请先输入扫描员姓名'}), 400

    global latest_data
    if latest_data is None:
        return jsonify({'error': 'No data uploaded'}), 400
    if recipient not in latest_data:
        return jsonify({'error': 'Recipient not found'}), 400

    all_items = latest_data[recipient].get('all_items', [])
    valid_codes = set()
    for item in all_items:
        for code in item['serials']:
            valid_codes.add(code)
    if barcode not in valid_codes:
        return jsonify({'match': False, 'message': '无效条码'}), 200

    db = get_db()
    cur = db.execute('SELECT 1 FROM scan_records WHERE recipient=? AND barcode=?', (recipient, barcode))
    if cur.fetchone():
        return jsonify({'match': True, 'already_scanned': True, 'message': '已扫过'}), 200

    db.execute('INSERT INTO scan_records (recipient, barcode, scanner_name) VALUES (?, ?, ?)',
               (recipient, barcode, scanner_name))
    db.commit()
    return jsonify({'match': True, 'already_scanned': False, 'message': '扫描成功'}), 200

@app.route('/status/<recipient>')
def get_status(recipient):
    db = get_db()
    cur = db.execute('SELECT barcode, scanner_name, scanned_at FROM scan_records WHERE recipient=?', (recipient,))
    rows = cur.fetchall()
    scanned = [{'barcode': row['barcode'], 'scanner': row['scanner_name'], 'time': row['scanned_at']} for row in rows]
    return jsonify({'scanned': scanned})

@app.route('/export', methods=['GET'])
def export_excel():
    recipient = request.args.get('recipient')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    if not recipient:
        return jsonify({'error': '缺少收件人参数'}), 400

    global latest_data
    if latest_data is None or recipient not in latest_data:
        return jsonify({'error': '收件人数据不存在'}), 404

    items = latest_data[recipient].get('all_items', [])
    if not items:
        return jsonify({'error': '无数据可导出'}), 400

    mexico_tz = ZoneInfo("America/Mexico_City")
    utc_start = None
    utc_end = None
    if start_date:
        local_dt = datetime.strptime(start_date, '%Y-%m-%d').replace(tzinfo=mexico_tz)
        utc_start = local_dt.astimezone(timezone.utc).replace(tzinfo=None)
    if end_date:
        local_dt = datetime.strptime(end_date, '%Y-%m-%d').replace(tzinfo=mexico_tz)
        next_day = local_dt + timedelta(days=1)
        utc_end = next_day.astimezone(timezone.utc).replace(tzinfo=None)

    query = 'SELECT barcode, scanner_name, scanned_at FROM scan_records WHERE recipient=?'
    params = [recipient]
    if utc_start:
        query += ' AND scanned_at >= ?'
        params.append(utc_start.strftime('%Y-%m-%d %H:%M:%S'))   # 关键修复：统一日期时间格式
    if utc_end:
        query += ' AND scanned_at < ?'
        params.append(utc_end.strftime('%Y-%m-%d %H:%M:%S'))

    db = get_db()
    cur = db.execute(query, params)
    rows = cur.fetchall()
    scan_info = {}
    for row in rows:
        scan_info[row['barcode']] = {
            'scanner': row['scanner_name'],
            'time': row['scanned_at']
        }

    wb = Workbook()
    ws = wb.active
    ws.title = '复核清单'
    headers = ['主码Master', '类型Tipo', '总箱数Veces', '条码Codigo', '扫描状态Estado de escaneo',
               '箱数Caja', '日期Fecha', '时间Hora', '扫描员Escaneador']
    ws.append(headers)
    for col in range(1, len(headers)+1):
        ws.cell(row=1, column=col).font = Font(bold=True)

    for item in items:
        for code in item['serials']:
            info = scan_info.get(code)
            if info:
                status = '已扫/Si'
                caja = 1
                utc_time = info['time']
                if isinstance(utc_time, str):
                    try:
                        utc_dt = datetime.strptime(utc_time, '%Y-%m-%d %H:%M:%S')
                    except ValueError:
                        utc_dt = datetime.fromisoformat(utc_time.replace('Z', '+00:00'))
                else:
                    utc_dt = utc_time
                if utc_dt.tzinfo is None:
                    utc_dt = utc_dt.replace(tzinfo=timezone.utc)
                mexico_dt = utc_dt.astimezone(mexico_tz)
                fecha = mexico_dt.strftime('%Y-%m-%d')
                hora = mexico_dt.strftime('%H:%M:%S')
                scanner = info['scanner'] or ''
            else:
                status = '未扫/No'
                caja = 0
                fecha = ''
                hora = ''
                scanner = ''

            ws.append([
                item['master_code'],
                item['type'],
                item['total_boxes'],
                code,
                status,
                caja,
                fecha,
                hora,
                scanner
            ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output,
                     as_attachment=True,
                     download_name=f'复核清单_{recipient}_{datetime.now().strftime("%Y%m%d")}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/reset/<recipient>')
def reset_recipient(recipient):
    db = get_db()
    db.execute('DELETE FROM scan_records WHERE recipient=?', (recipient,))
    db.commit()
    return jsonify({'success': True})

@app.route('/reset_master/<recipient>/<master_code>', methods=['POST'])
def reset_master(recipient, master_code):
    db = get_db()
    db.execute('DELETE FROM scan_records WHERE recipient=? AND barcode LIKE ?',
               (recipient, master_code + 'U%'))
    db.commit()
    return jsonify({'success': True})

if __name__ == '__main__':
    init_db()
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
