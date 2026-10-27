import os
import re
from io import BytesIO
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
import io
import psycopg2
from psycopg2.extras import RealDictCursor
import json

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# ---------- 数据库连接（Supabase Transaction Pooler） ----------
def get_db_connection():
    database_url = os.environ.get('DATABASE_URL')
    if not database_url:
        raise RuntimeError("DATABASE_URL environment variable not set")
    # Transaction Pooler 不支持预处理语句，禁用 prepare_threshold
    return psycopg2.connect(
        database_url,
        sslmode='require',
        options='-c prepare_threshold=0'
    )

def init_db():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS manifests (
            recipient TEXT PRIMARY KEY,
            data JSONB NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS scan_records (
            id SERIAL PRIMARY KEY,
            recipient TEXT NOT NULL,
            barcode TEXT NOT NULL,
            scanner_name TEXT,
            scanned_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(recipient, barcode)
        )
    ''')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_scan_recipient ON scan_records (recipient)')
    conn.commit()
    cur.close()
    conn.close()

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

def parse_excel_from_workbook(wb):
    sheet = wb['按箱出库'] if '按箱出库' in wb.sheetnames else wb.active
    header = [str(cell.value).strip() if cell.value else '' for cell in sheet[1]]

    # 1. 智能识别各列索引（支持新旧不同表头）
    recipient_idx = None
    for i, h in enumerate(header):
        if '收货地址' in h or '收件人' in h or 'Recipient' in h:
            recipient_idx = i
            break

    fba_idx = None
    for i, h in enumerate(header):
        if '货件号' in h or 'FBA货件ID' in h or 'FBAShipmentID' in h:
            fba_idx = i
            break

    custom_idx = None
    for i, h in enumerate(header):
        if '自定义箱条码' in h or 'Custom box barcode' in h or '箱唛-入仓号' in h:
            custom_idx = i
            break

    box_count_idx = None
    for i, h in enumerate(header):
        if '箱数' in h:
            box_count_idx = i
            break

    if fba_idx is None and custom_idx is None:
        raise Exception("未找到必要的列: 既没有找到 '货件号/FBA'，也没有找到 '自定义箱条码/箱唛-入仓号'")

    data_rows = []
    last_fba_code = None
    last_custom_code = None
    last_recipient = "默认收件人"

    for row in sheet.iter_rows(min_row=2, values_only=True):
        # 跳过完全空白的行
        if all(cell is None or str(cell).strip() == '' for cell in row):
            continue

        # 提取数据
        recipient = str(row[recipient_idx]).strip() if recipient_idx is not None and len(row) > recipient_idx and row[recipient_idx] else None
        fba_id = str(row[fba_idx]).strip() if fba_idx is not None and len(row) > fba_idx and row[fba_idx] else None
        custom_barcode = str(row[custom_idx]).strip() if custom_idx is not None and len(row) > custom_idx and row[custom_idx] else None
        
        # 提取箱数 (如果存在且合法)
        box_count = 1
        if box_count_idx is not None and len(row) > box_count_idx and row[box_count_idx] is not None:
            try:
                box_count = int(float(row[box_count_idx]))
            except ValueError:
                box_count = 1

        # 处理合并单元格的继承逻辑：如果为空，则沿用上一行的值
        if not recipient and last_recipient:
            recipient = last_recipient
        elif recipient:
            last_recipient = recipient

        if not fba_id and last_fba_code:
            fba_id = last_fba_code
        elif fba_id:
            last_fba_code = fba_id

        if not custom_barcode and last_custom_code:
            custom_barcode = last_custom_code
        elif custom_barcode:
            last_custom_code = custom_barcode

        # 判定主码：优先 FBA
        if fba_id:
            master_code = fba_id
            type_tag = 'FBA'
        elif custom_barcode:
            master_code = extract_custom_base_code(custom_barcode)
            type_tag = 'Custom'
        else:
            continue

        data_rows.append({
            'recipient': recipient,
            'master_code': master_code,
            'type': type_tag,
            'box_count': box_count
        })

    # 2. 聚合并累加箱数 (处理同一个货件分多行的情况)
    agg_data = defaultdict(lambda: defaultdict(lambda: {'count': 0, 'type': ''}))
    for r in data_rows:
        rec = r['recipient']
        master = r['master_code']
        agg_data[rec][master]['count'] += r['box_count']
        agg_data[rec][master]['type'] = r['type']

    # 3. 生成最终结果
    result = {}
    recipients = list(agg_data.keys())

    for rec in recipients:
        items = []
        for master, info in agg_data[rec].items():
            count = info['count']
            t = info['type']
            serials = generate_serial_list(master, count)
            items.append({
                'master_code': master,
                'type': t,
                'total_boxes': count,
                'serials': serials
            })
        result[rec] = {'all_items': items}

    return result, recipients

# ---------- 数据库操作 ----------
def save_manifest(recipient, data):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute('''
        INSERT INTO manifests (recipient, data, updated_at)
        VALUES (%s, %s, CURRENT_TIMESTAMP)
        ON CONFLICT (recipient) DO UPDATE
        SET data = EXCLUDED.data, updated_at = CURRENT_TIMESTAMP
    ''', (recipient, json.dumps(data)))
    conn.commit()
    cur.close()
    conn.close()

def load_manifest(recipient):
    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT data FROM manifests WHERE recipient=%s', (recipient,))
    row = cur.fetchone()
    cur.close()
    conn.close()
    if row:
        return row['data']
    return None

def load_all_recipients():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute('SELECT recipient FROM manifests')
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return [row[0] for row in rows]

# ---------- 路由 ----------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if not allowed_file(file.filename):
        return jsonify({'error': 'File type not allowed'}), 400

    try:
        # 内存读取，不落盘（Vercel 文件系统只读）
        file_bytes = file.read()
        wb = load_workbook(BytesIO(file_bytes), data_only=True)
        data, recipients = parse_excel_from_workbook(wb)
        for rec, rec_data in data.items():
            save_manifest(rec, rec_data)
        return jsonify({
            'success': True,
            'recipients': recipients,
            'data': data
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/get_data', methods=['GET'])
def get_data():
    recipients = load_all_recipients()
    if not recipients:
        return jsonify({'success': False, 'message': '暂无数据，请先上传出库表'}), 404
    data = {}
    for rec in recipients:
        manifest = load_manifest(rec)
        if manifest:
            data[rec] = manifest
    return jsonify({
        'success': True,
        'recipients': recipients,
        'data': data
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

    manifest = load_manifest(recipient)
    if not manifest:
        return jsonify({'error': 'Recipient not found'}), 400

    all_items = manifest.get('all_items', [])
    valid_codes = set()
    for item in all_items:
        for code in item['serials']:
            valid_codes.add(code)
    if barcode not in valid_codes:
        return jsonify({'match': False, 'message': '无效条码'}), 200

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute('SELECT 1 FROM scan_records WHERE recipient=%s AND barcode=%s', (recipient, barcode))
    if cur.fetchone():
        cur.close()
        conn.close()
        return jsonify({'match': True, 'already_scanned': True, 'message': '已扫过'}), 200

    cur.execute('INSERT INTO scan_records (recipient, barcode, scanner_name) VALUES (%s, %s, %s)',
                (recipient, barcode, scanner_name))
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({'match': True, 'already_scanned': False, 'message': '扫描成功'}), 200

@app.route('/status/<recipient>')
def get_status(recipient):
    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT barcode, scanner_name, scanned_at FROM scan_records WHERE recipient=%s', (recipient,))
    rows = cur.fetchall()
    cur.close()
    conn.close()
    scanned = [{'barcode': row['barcode'], 'scanner': row['scanner_name'], 'time': row['scanned_at']} for row in rows]
    return jsonify({'scanned': scanned})

# ========== 导出：复核清单（基于当前清单） ==========
@app.route('/export', methods=['GET'])
def export_excel():
    recipient = request.args.get('recipient')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    if not recipient:
        return jsonify({'error': '缺少收件人参数'}), 400

    manifest = load_manifest(recipient)
    if not manifest:
        return jsonify({'error': '收件人数据不存在'}), 404

    items = manifest.get('all_items', [])
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

    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    query = 'SELECT barcode, scanner_name, scanned_at FROM scan_records WHERE recipient=%s'
    params = [recipient]
    if utc_start:
        query += ' AND scanned_at >= %s'
        params.append(utc_start.strftime('%Y-%m-%d %H:%M:%S'))
    if utc_end:
        query += ' AND scanned_at < %s'
        params.append(utc_end.strftime('%Y-%m-%d %H:%M:%S'))
    cur.execute(query, params)
    rows = cur.fetchall()
    cur.close()
    conn.close()
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

# ========== 导出：历史记录（所有扫描记录，格式与复核清单一致） ==========
@app.route('/export_history', methods=['GET'])
def export_history():
    recipient = request.args.get('recipient')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    if not recipient:
        return jsonify({'error': '缺少收件人参数'}), 400

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

    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    query = 'SELECT barcode, scanner_name, scanned_at FROM scan_records WHERE recipient=%s'
    params = [recipient]
    if utc_start:
        query += ' AND scanned_at >= %s'
        params.append(utc_start.strftime('%Y-%m-%d %H:%M:%S'))
    if utc_end:
        query += ' AND scanned_at < %s'
        params.append(utc_end.strftime('%Y-%m-%d %H:%M:%S'))
    cur.execute(query, params)
    rows = cur.fetchall()
    cur.close()
    conn.close()

    if not rows:
        return jsonify({'error': '该时间段内无扫描记录'}), 404

    wb = Workbook()
    ws = wb.active
    ws.title = '历史记录'

    headers = ['主码Master', '类型Tipo', '总箱数Veces', '条码Codigo', '扫描状态Estado de escaneo',
               '箱数Caja', '日期Fecha', '时间Hora', '扫描员Escaneador']
    ws.append(headers)
    for col in range(1, len(headers)+1):
        ws.cell(row=1, column=col).font = Font(bold=True)

    for row in rows:
        barcode = row['barcode']
        master_code = extract_custom_base_code(barcode)
        if not master_code:
            master_code = barcode

        utc_time = row['scanned_at']
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
        scanner = row['scanner_name'] or ''

        ws.append([
            master_code,
            '历史',
            1,
            barcode,
            '已扫/Si',
            1,
            fecha,
            hora,
            scanner
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output,
                     as_attachment=True,
                     download_name=f'历史记录_{recipient}_{datetime.now().strftime("%Y%m%d")}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/reset/<recipient>')
def reset_recipient(recipient):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute('DELETE FROM scan_records WHERE recipient=%s', (recipient,))
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({'success': True})

@app.route('/reset_master/<recipient>/<master_code>', methods=['POST'])
def reset_master(recipient, master_code):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute('DELETE FROM scan_records WHERE recipient=%s AND barcode LIKE %s',
                (recipient, master_code + 'U%'))
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({'success': True})

# Vercel 需要这个入口（本地运行时也保留）
# ---- Vercel 启动时自动建表 ----
try:
    init_db()
    print("数据库表初始化成功")
except Exception as e:
    print(f"数据库表初始化失败，请检查 DATABASE_URL: {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)