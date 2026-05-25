from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, jsonify, Response)
import sqlite3, qrcode, io, base64, csv, os, json, smtplib, hashlib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, date
from functools import wraps
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = 'it-equipment-secret-key-2024'
DATABASE = 'equipment.db'
UPLOAD_FOLDER = os.path.join('static', 'uploads')
ALLOWED = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ── DB HELPERS ─────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        name TEXT NOT NULL,
        role TEXT DEFAULT 'user',
        department TEXT,
        email TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS equipment (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        category TEXT NOT NULL,
        serial_number TEXT UNIQUE,
        brand TEXT,
        model TEXT,
        status TEXT DEFAULT 'available',
        description TEXT,
        image TEXT,
        added_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS borrows (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        equipment_id INTEGER NOT NULL,
        user_id INTEGER NOT NULL,
        borrow_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        due_date DATE,
        return_date TIMESTAMP,
        status TEXT DEFAULT 'borrowed',
        notes TEXT,
        FOREIGN KEY (equipment_id) REFERENCES equipment(id),
        FOREIGN KEY (user_id) REFERENCES users(id)
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS email_settings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        smtp_host TEXT DEFAULT '',
        smtp_port INTEGER DEFAULT 587,
        smtp_user TEXT DEFAULT '',
        smtp_password TEXT DEFAULT '',
        smtp_from TEXT DEFAULT '',
        enabled INTEGER DEFAULT 0
    )''')

    for sql in ['ALTER TABLE users ADD COLUMN email TEXT',
                'ALTER TABLE equipment ADD COLUMN image TEXT']:
        try:
            c.execute(sql)
        except Exception:
            pass

    if not c.execute('SELECT COUNT(*) FROM email_settings').fetchone()[0]:
        c.execute('INSERT INTO email_settings (smtp_host, smtp_port) VALUES ("", 587)')

    admin_pw = hashlib.sha256('admin123'.encode()).hexdigest()
    try:
        c.execute("INSERT INTO users (username, password, name, role, department) VALUES (?,?,?,?,?)",
                  ('admin', admin_pw, 'ผู้ดูแลระบบ', 'admin', 'ฝ่าย IT'))
    except Exception:
        pass

    sample = [
        ('MacBook Pro 14"',  'Notebook', 'NB-001', 'Apple',  'MacBook Pro M2', 'มี 16GB RAM / 512GB SSD'),
        ('Dell Monitor 27"', 'Monitor',  'MN-001', 'Dell',   'U2722D 4K',      'จอ 4K IPS 27 นิ้ว'),
        ('iPhone 14 Pro',    'Mobile',   'MB-001', 'Apple',  'iPhone 14 Pro',  '256GB Deep Purple'),
        ('Canon EOS R5',     'Camera',   'CM-001', 'Canon',  'EOS R5',         'กล้อง Mirrorless 45MP'),
        ('iPad Pro 12.9"',   'Tablet',   'TB-001', 'Apple',  'iPad Pro M2',    '256GB WiFi+Cellular'),
        ('HP LaserJet Pro',  'Printer',  'PR-001', 'HP',     'LaserJet M404n', 'เครื่องพิมพ์ขาวดำ'),
    ]
    for s in sample:
        try:
            c.execute('''INSERT INTO equipment (name, category, serial_number, brand, model, description)
                         VALUES (?,?,?,?,?,?)''', s)
        except Exception:
            pass

    conn.commit()
    conn.close()


# ── AUTH DECORATORS ────────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('role') != 'admin':
            flash('ไม่มีสิทธิ์เข้าถึงหน้านี้', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated


# ── EMAIL HELPERS ──────────────────────────────────────────────────────────────

def get_email_cfg():
    conn = get_db()
    row = conn.execute('SELECT * FROM email_settings LIMIT 1').fetchone()
    conn.close()
    return dict(row) if row else {}


def send_email(to_email, subject, html_body):
    if not to_email:
        return False
    cfg = get_email_cfg()
    if not cfg.get('enabled') or not cfg.get('smtp_host'):
        return False
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From']    = cfg.get('smtp_from') or cfg.get('smtp_user', '')
        msg['To']      = to_email
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))
        with smtplib.SMTP(cfg['smtp_host'], int(cfg['smtp_port']), timeout=10) as s:
            s.ehlo()
            s.starttls()
            s.login(cfg['smtp_user'], cfg['smtp_password'])
            s.sendmail(msg['From'], [to_email], msg.as_string())
        return True
    except Exception as e:
        print(f'[Email Error] {e}')
        return False


def email_borrow(user_email, user_name, eq_name, due_date):
    html = f"""
    <div style="font-family:sans-serif;max-width:500px;margin:auto;padding:20px">
      <div style="background:#2563eb;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0">
        <h2 style="margin:0">✅ ยืมอุปกรณ์สำเร็จ</h2>
      </div>
      <div style="border:1px solid #e2e8f0;border-top:none;padding:20px;border-radius:0 0 8px 8px">
        <p>สวัสดีคุณ <strong>{user_name}</strong></p>
        <p>คุณได้ทำการยืม <strong>{eq_name}</strong> เรียบร้อยแล้ว</p>
        {'<p>📅 กำหนดคืน: <strong>' + str(due_date) + '</strong></p>' if due_date else ''}
        <p style="color:#64748b;font-size:0.85em">หากมีข้อสงสัยกรุณาติดต่อฝ่าย IT</p>
      </div>
    </div>"""
    return send_email(user_email, f'[IT Borrow] ยืมอุปกรณ์: {eq_name}', html)


def email_overdue(user_email, user_name, eq_name, due_date):
    html = f"""
    <div style="font-family:sans-serif;max-width:500px;margin:auto;padding:20px">
      <div style="background:#dc2626;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0">
        <h2 style="margin:0">⚠️ อุปกรณ์เกินกำหนดคืน</h2>
      </div>
      <div style="border:1px solid #fecaca;border-top:none;padding:20px;border-radius:0 0 8px 8px;background:#fff5f5">
        <p>สวัสดีคุณ <strong>{user_name}</strong></p>
        <p>อุปกรณ์ <strong>{eq_name}</strong> ที่คุณยืมไว้เกินกำหนดคืนแล้ว</p>
        <p>📅 กำหนดคืนเดิม: <strong style="color:#dc2626">{due_date}</strong></p>
        <p>กรุณาคืนอุปกรณ์โดยเร็วที่สุด</p>
      </div>
    </div>"""
    return send_email(user_email, f'[IT Borrow] ⚠️ เกินกำหนดคืน: {eq_name}', html)


# ── QR ─────────────────────────────────────────────────────────────────────────

def generate_qr(data):
    qr = qrcode.QRCode(version=1, box_size=8, border=4)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color='black', back_color='white')
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return base64.b64encode(buf.getvalue()).decode()


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED


# ── AUTH ───────────────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form['username']
        password = hashlib.sha256(request.form['password'].encode()).hexdigest()
        conn = get_db()
        user = conn.execute('SELECT * FROM users WHERE username=? AND password=?',
                            (username, password)).fetchone()
        conn.close()
        if user:
            session.update({'user_id': user['id'], 'username': user['username'],
                            'name': user['name'], 'role': user['role']})
            flash(f'ยินดีต้อนรับ คุณ{user["name"]}', 'success')
            return redirect(url_for('dashboard'))
        flash('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง', 'danger')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ── DASHBOARD ──────────────────────────────────────────────────────────────────

@app.route('/')
@login_required
def dashboard():
    conn  = get_db()
    today = date.today().isoformat()

    total     = conn.execute("SELECT COUNT(*) as c FROM equipment").fetchone()['c']
    available = conn.execute("SELECT COUNT(*) as c FROM equipment WHERE status='available'").fetchone()['c']
    borrowed  = conn.execute("SELECT COUNT(*) as c FROM equipment WHERE status='borrowed'").fetchone()['c']

    # Overdue borrows
    overdue = conn.execute('''
        SELECT b.*, e.name as eq_name, u.name as user_name, u.email as user_email
        FROM borrows b
        JOIN equipment e ON b.equipment_id=e.id
        JOIN users u ON b.user_id=u.id
        WHERE b.status='borrowed' AND b.due_date IS NOT NULL AND b.due_date < ?
        ORDER BY b.due_date ASC
    ''', (today,)).fetchall()

    recent = conn.execute('''
        SELECT b.*, e.name as eq_name, e.category, u.name as user_name
        FROM borrows b
        JOIN equipment e ON b.equipment_id=e.id
        JOIN users u ON b.user_id=u.id
        ORDER BY b.borrow_date DESC LIMIT 8
    ''').fetchall()

    my_borrows = conn.execute('''
        SELECT b.*, e.name as eq_name, e.category, e.brand, e.image
        FROM borrows b JOIN equipment e ON b.equipment_id=e.id
        WHERE b.user_id=? AND b.status='borrowed'
        ORDER BY b.borrow_date DESC
    ''', (session['user_id'],)).fetchall()

    # Chart: monthly borrows last 6 months
    monthly = conn.execute('''
        SELECT strftime('%Y-%m', borrow_date) as month, COUNT(*) as cnt
        FROM borrows
        WHERE borrow_date >= date('now', '-6 months')
        GROUP BY month ORDER BY month
    ''').fetchall()
    chart_labels  = [r['month'] for r in monthly]
    chart_borrows = [r['cnt']   for r in monthly]

    # Chart: category stats
    cat_stats = conn.execute('''
        SELECT category,
               COUNT(*) as total,
               SUM(CASE WHEN status='available' THEN 1 ELSE 0 END) as avail,
               SUM(CASE WHEN status='borrowed'  THEN 1 ELSE 0 END) as lent
        FROM equipment GROUP BY category
    ''').fetchall()

    conn.close()
    return render_template('dashboard.html',
                           total=total, available=available, borrowed=borrowed,
                           overdue=overdue, recent=recent, my_borrows=my_borrows,
                           chart_labels=json.dumps(chart_labels),
                           chart_borrows=json.dumps(chart_borrows),
                           cat_stats=cat_stats)


# ── EQUIPMENT ──────────────────────────────────────────────────────────────────

@app.route('/equipment')
@login_required
def equipment_list():
    search   = request.args.get('search', '')
    category = request.args.get('category', '')
    status   = request.args.get('status', '')
    conn     = get_db()
    query    = 'SELECT * FROM equipment WHERE 1=1'
    params   = []
    if search:
        query += ' AND (name LIKE ? OR serial_number LIKE ? OR brand LIKE ? OR model LIKE ?)'
        params.extend([f'%{search}%'] * 4)
    if category:
        query  += ' AND category=?'; params.append(category)
    if status:
        query  += ' AND status=?'; params.append(status)
    query += ' ORDER BY added_date DESC'
    equipment  = conn.execute(query, params).fetchall()
    categories = conn.execute('SELECT DISTINCT category FROM equipment ORDER BY category').fetchall()
    conn.close()
    return render_template('equipment_list.html',
                           equipment=equipment, categories=categories,
                           search=search, selected_category=category, selected_status=status)


@app.route('/equipment/add', methods=['GET', 'POST'])
@admin_required
def add_equipment():
    if request.method == 'POST':
        name = request.form['name']; category = request.form['category']
        serial = request.form['serial_number']; brand = request.form['brand']
        model  = request.form['model'];  desc  = request.form['description']
        conn   = get_db()
        try:
            conn.execute('''INSERT INTO equipment (name, category, serial_number, brand, model, description)
                            VALUES (?,?,?,?,?,?)''', (name, category, serial, brand, model, desc))
            eid = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
            conn.commit()
            # Handle image upload
            f = request.files.get('image')
            if f and f.filename and allowed_file(f.filename):
                ext  = secure_filename(f.filename).rsplit('.', 1)[1].lower()
                fname = f'eq_{eid}.{ext}'
                f.save(os.path.join(UPLOAD_FOLDER, fname))
                conn.execute('UPDATE equipment SET image=? WHERE id=?', (fname, eid))
                conn.commit()
            flash('เพิ่มอุปกรณ์สำเร็จ', 'success')
            return redirect(url_for('equipment_list'))
        except sqlite3.IntegrityError:
            flash('Serial Number นี้มีอยู่แล้วในระบบ', 'danger')
        finally:
            conn.close()
    return render_template('add_equipment.html')


@app.route('/equipment/<int:eid>')
@login_required
def equipment_detail(eid):
    conn = get_db()
    eq   = conn.execute('SELECT * FROM equipment WHERE id=?', (eid,)).fetchone()
    if not eq:
        flash('ไม่พบอุปกรณ์', 'danger')
        conn.close()
        return redirect(url_for('equipment_list'))
    history = conn.execute('''
        SELECT b.*, u.name as user_name, u.department
        FROM borrows b JOIN users u ON b.user_id=u.id
        WHERE b.equipment_id=? ORDER BY b.borrow_date DESC
    ''', (eid,)).fetchall()
    current_borrow = conn.execute('''
        SELECT b.*, u.name as user_name
        FROM borrows b JOIN users u ON b.user_id=u.id
        WHERE b.equipment_id=? AND b.status='borrowed'
    ''', (eid,)).fetchone()
    qr_data  = f"ID:{eid}|ชื่อ:{eq['name']}|Serial:{eq['serial_number']}|หมวด:{eq['category']}"
    qr_image = generate_qr(qr_data)
    conn.close()
    return render_template('equipment_detail.html',
                           eq=eq, history=history, current_borrow=current_borrow,
                           qr_image=qr_image)


@app.route('/equipment/<int:eid>/edit', methods=['GET', 'POST'])
@admin_required
def edit_equipment(eid):
    conn = get_db()
    eq   = conn.execute('SELECT * FROM equipment WHERE id=?', (eid,)).fetchone()
    if not eq:
        conn.close()
        return redirect(url_for('equipment_list'))
    if request.method == 'POST':
        name = request.form['name']; category = request.form['category']
        serial = request.form['serial_number']; brand = request.form['brand']
        model  = request.form['model'];  desc  = request.form['description']
        try:
            conn.execute('''UPDATE equipment SET name=?, category=?, serial_number=?,
                            brand=?, model=?, description=? WHERE id=?''',
                         (name, category, serial, brand, model, desc, eid))
            f = request.files.get('image')
            if f and f.filename and allowed_file(f.filename):
                ext   = secure_filename(f.filename).rsplit('.', 1)[1].lower()
                fname = f'eq_{eid}.{ext}'
                f.save(os.path.join(UPLOAD_FOLDER, fname))
                conn.execute('UPDATE equipment SET image=? WHERE id=?', (fname, eid))
            conn.commit()
            flash('แก้ไขข้อมูลอุปกรณ์สำเร็จ', 'success')
        except Exception:
            flash('เกิดข้อผิดพลาด', 'danger')
        finally:
            conn.close()
        return redirect(url_for('equipment_detail', eid=eid))
    conn.close()
    return render_template('edit_equipment.html', eq=eq)


@app.route('/equipment/<int:eid>/delete', methods=['POST'])
@admin_required
def delete_equipment(eid):
    conn = get_db()
    eq   = conn.execute('SELECT * FROM equipment WHERE id=?', (eid,)).fetchone()
    if eq and eq['status'] == 'available':
        if eq['image']:
            try:
                os.remove(os.path.join(UPLOAD_FOLDER, eq['image']))
            except Exception:
                pass
        conn.execute('DELETE FROM equipment WHERE id=?', (eid,))
        conn.commit()
        flash('ลบอุปกรณ์สำเร็จ', 'success')
    else:
        flash('ไม่สามารถลบอุปกรณ์ที่กำลังถูกยืมอยู่ได้', 'danger')
    conn.close()
    return redirect(url_for('equipment_list'))


# ── BORROW / RETURN ────────────────────────────────────────────────────────────

@app.route('/borrow/<int:eid>', methods=['POST'])
@login_required
def borrow(eid):
    conn = get_db()
    eq   = conn.execute('SELECT * FROM equipment WHERE id=?', (eid,)).fetchone()
    if not eq or eq['status'] != 'available':
        flash('อุปกรณ์ไม่พร้อมให้ยืม', 'danger')
        conn.close()
        return redirect(url_for('equipment_list'))
    due_date = request.form.get('due_date') or None
    notes    = request.form.get('notes', '')
    conn.execute('INSERT INTO borrows (equipment_id, user_id, due_date, notes) VALUES (?,?,?,?)',
                 (eid, session['user_id'], due_date, notes))
    conn.execute("UPDATE equipment SET status='borrowed' WHERE id=?", (eid,))
    conn.commit()
    # Send confirmation email
    user = conn.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()
    conn.close()
    if user and user['email']:
        email_borrow(user['email'], user['name'], eq['name'], due_date)
    flash(f'ยืม "{eq["name"]}" สำเร็จ', 'success')
    return redirect(url_for('dashboard'))


@app.route('/return/<int:bid>', methods=['POST'])
@login_required
def return_equipment(bid):
    conn   = get_db()
    borrow_row = conn.execute('''
        SELECT b.*, e.name as eq_name
        FROM borrows b JOIN equipment e ON b.equipment_id=e.id WHERE b.id=?
    ''', (bid,)).fetchone()
    if not borrow_row:
        flash('ไม่พบข้อมูลการยืม', 'danger')
        conn.close()
        return redirect(url_for('dashboard'))
    if borrow_row['user_id'] != session['user_id'] and session.get('role') != 'admin':
        flash('ไม่มีสิทธิ์คืนอุปกรณ์ชิ้นนี้', 'danger')
        conn.close()
        return redirect(url_for('dashboard'))
    conn.execute("UPDATE borrows SET return_date=CURRENT_TIMESTAMP, status='returned' WHERE id=?", (bid,))
    conn.execute("UPDATE equipment SET status='available' WHERE id=?", (borrow_row['equipment_id'],))
    conn.commit()
    conn.close()
    flash(f'คืน "{borrow_row["eq_name"]}" สำเร็จ', 'success')
    return redirect(url_for('dashboard'))


# ── HISTORY + EXPORT ───────────────────────────────────────────────────────────

@app.route('/history')
@login_required
def history():
    today = date.today().isoformat()
    conn  = get_db()
    if session.get('role') == 'admin':
        borrows = conn.execute('''
            SELECT b.*, e.name as eq_name, e.category, u.name as user_name, u.department
            FROM borrows b
            JOIN equipment e ON b.equipment_id=e.id
            JOIN users u ON b.user_id=u.id
            ORDER BY b.borrow_date DESC
        ''').fetchall()
    else:
        borrows = conn.execute('''
            SELECT b.*, e.name as eq_name, e.category, u.name as user_name, u.department
            FROM borrows b
            JOIN equipment e ON b.equipment_id=e.id
            JOIN users u ON b.user_id=u.id
            WHERE b.user_id=? ORDER BY b.borrow_date DESC
        ''', (session['user_id'],)).fetchall()
    conn.close()
    return render_template('history.html', borrows=borrows, today=today)


@app.route('/history/export')
@login_required
def export_history():
    fmt   = request.args.get('fmt', 'csv')
    today = date.today().isoformat()
    conn  = get_db()
    if session.get('role') == 'admin':
        rows = conn.execute('''
            SELECT b.id, e.name as eq_name, e.category, e.serial_number,
                   u.name as user_name, u.department,
                   b.borrow_date, b.due_date, b.return_date, b.status, b.notes
            FROM borrows b
            JOIN equipment e ON b.equipment_id=e.id
            JOIN users u ON b.user_id=u.id
            ORDER BY b.borrow_date DESC
        ''').fetchall()
    else:
        rows = conn.execute('''
            SELECT b.id, e.name as eq_name, e.category, e.serial_number,
                   u.name as user_name, u.department,
                   b.borrow_date, b.due_date, b.return_date, b.status, b.notes
            FROM borrows b
            JOIN equipment e ON b.equipment_id=e.id
            JOIN users u ON b.user_id=u.id
            WHERE b.user_id=? ORDER BY b.borrow_date DESC
        ''', (session['user_id'],)).fetchall()
    conn.close()

    headers_th = ['#', 'อุปกรณ์', 'หมวด', 'Serial', 'ผู้ยืม', 'แผนก',
                  'วันที่ยืม', 'ครบกำหนด', 'วันที่คืน', 'สถานะ', 'หมายเหตุ']

    if fmt == 'excel':
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'ประวัติการยืม'

        # Header style
        hdr_fill = PatternFill('solid', fgColor='2563EB')
        hdr_font = Font(bold=True, color='FFFFFF', size=11)
        for col, h in enumerate(headers_th, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        status_map = {'borrowed': 'ยืมอยู่', 'returned': 'คืนแล้ว'}
        for r, row in enumerate(rows, 2):
            vals = [row['id'], row['eq_name'], row['category'], row['serial_number'],
                    row['user_name'], row['department'],
                    row['borrow_date'][:16] if row['borrow_date'] else '',
                    row['due_date'] or '',
                    row['return_date'][:16] if row['return_date'] else '',
                    status_map.get(row['status'], row['status']),
                    row['notes'] or '']
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                if row['status'] == 'borrowed' and (row['due_date'] or '') < today and row['due_date']:
                    cell.fill = PatternFill('solid', fgColor='FEE2E2')

        # Column widths
        widths = [5, 25, 12, 14, 18, 14, 18, 12, 18, 10, 20]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return Response(buf.getvalue(),
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        headers={'Content-Disposition': f'attachment; filename=borrow_history_{today}.xlsx'})

    # CSV
    output  = io.StringIO()
    writer  = csv.writer(output)
    writer.writerow(headers_th)
    status_map = {'borrowed': 'ยืมอยู่', 'returned': 'คืนแล้ว'}
    for row in rows:
        writer.writerow([
            row['id'], row['eq_name'], row['category'], row['serial_number'],
            row['user_name'], row['department'],
            row['borrow_date'][:16] if row['borrow_date'] else '',
            row['due_date'] or '',
            row['return_date'][:16] if row['return_date'] else '',
            status_map.get(row['status'], row['status']),
            row['notes'] or ''
        ])
    output.seek(0)
    return Response('﻿' + output.getvalue(),
                    mimetype='text/csv; charset=utf-8-sig',
                    headers={'Content-Disposition': f'attachment; filename=borrow_history_{today}.csv'})


# ── USERS ──────────────────────────────────────────────────────────────────────

@app.route('/users')
@admin_required
def user_list():
    conn  = get_db()
    users = conn.execute('SELECT * FROM users ORDER BY created_at DESC').fetchall()
    conn.close()
    return render_template('users.html', users=users)


@app.route('/users/add', methods=['GET', 'POST'])
@admin_required
def add_user():
    if request.method == 'POST':
        username = request.form['username']
        password = hashlib.sha256(request.form['password'].encode()).hexdigest()
        name = request.form['name']; role = request.form['role']
        department = request.form['department']; email = request.form.get('email', '')
        conn = get_db()
        try:
            conn.execute('INSERT INTO users (username, password, name, role, department, email) VALUES (?,?,?,?,?,?)',
                         (username, password, name, role, department, email))
            conn.commit()
            flash('เพิ่มผู้ใช้สำเร็จ', 'success')
            return redirect(url_for('user_list'))
        except sqlite3.IntegrityError:
            flash('ชื่อผู้ใช้นี้มีอยู่แล้ว', 'danger')
        finally:
            conn.close()
    return render_template('add_user.html')


@app.route('/users/<int:uid>/delete', methods=['POST'])
@admin_required
def delete_user(uid):
    if uid == session['user_id']:
        flash('ไม่สามารถลบบัญชีของตัวเองได้', 'danger')
        return redirect(url_for('user_list'))
    conn = get_db()
    conn.execute('DELETE FROM users WHERE id=?', (uid,))
    conn.commit()
    conn.close()
    flash('ลบผู้ใช้สำเร็จ', 'success')
    return redirect(url_for('user_list'))


# ── PROFILE ────────────────────────────────────────────────────────────────────

@app.route('/profile', methods=['GET', 'POST'])
@login_required
def profile():
    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'update_info':
            name       = request.form['name']
            department = request.form['department']
            email      = request.form.get('email', '')
            conn.execute('UPDATE users SET name=?, department=?, email=? WHERE id=?',
                         (name, department, email, session['user_id']))
            conn.commit()
            session['name'] = name
            flash('อัพเดทข้อมูลสำเร็จ', 'success')

        elif action == 'change_password':
            old_pw  = hashlib.sha256(request.form['old_password'].encode()).hexdigest()
            new_pw  = hashlib.sha256(request.form['new_password'].encode()).hexdigest()
            confirm = hashlib.sha256(request.form['confirm_password'].encode()).hexdigest()
            if old_pw != user['password']:
                flash('รหัสผ่านเดิมไม่ถูกต้อง', 'danger')
            elif new_pw != confirm:
                flash('รหัสผ่านใหม่ไม่ตรงกัน', 'danger')
            else:
                conn.execute('UPDATE users SET password=? WHERE id=?', (new_pw, session['user_id']))
                conn.commit()
                flash('เปลี่ยนรหัสผ่านสำเร็จ', 'success')

        conn.close()
        return redirect(url_for('profile'))

    my_borrows = conn.execute('''
        SELECT b.*, e.name as eq_name, e.category
        FROM borrows b JOIN equipment e ON b.equipment_id=e.id
        WHERE b.user_id=? ORDER BY b.borrow_date DESC LIMIT 10
    ''', (session['user_id'],)).fetchall()
    conn.close()
    return render_template('profile.html', user=user, my_borrows=my_borrows)


# ── EMAIL SETTINGS ─────────────────────────────────────────────────────────────

@app.route('/admin/email-settings', methods=['GET', 'POST'])
@admin_required
def email_settings_page():
    conn = get_db()
    cfg  = conn.execute('SELECT * FROM email_settings LIMIT 1').fetchone()

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'save':
            host     = request.form['smtp_host']
            port     = int(request.form.get('smtp_port', 587))
            user     = request.form['smtp_user']
            password = request.form['smtp_password']
            from_    = request.form['smtp_from']
            enabled  = 1 if request.form.get('enabled') else 0
            conn.execute('''UPDATE email_settings SET smtp_host=?, smtp_port=?, smtp_user=?,
                            smtp_password=?, smtp_from=?, enabled=? WHERE id=?''',
                         (host, port, user, password, from_, enabled, cfg['id']))
            conn.commit()
            flash('บันทึกการตั้งค่า Email สำเร็จ', 'success')

        elif action == 'test':
            test_to = request.form.get('test_email', '')
            ok = send_email(test_to, '[IT Borrow] ทดสอบการส่ง Email',
                            '<p>ระบบส่ง Email ทำงานปกติ ✅</p>')
            flash('ส่ง Email ทดสอบสำเร็จ' if ok else 'ส่งไม่สำเร็จ กรุณาตรวจสอบการตั้งค่า', 'success' if ok else 'danger')

        elif action == 'send_overdue':
            today   = date.today().isoformat()
            overdue = conn.execute('''
                SELECT b.*, e.name as eq_name, u.name as user_name, u.email as user_email
                FROM borrows b
                JOIN equipment e ON b.equipment_id=e.id
                JOIN users u ON b.user_id=u.id
                WHERE b.status='borrowed' AND b.due_date IS NOT NULL AND b.due_date < ?
            ''', (today,)).fetchall()
            sent = sum(1 for b in overdue if email_overdue(b['user_email'], b['user_name'], b['eq_name'], b['due_date']))
            flash(f'ส่ง Email แจ้งเตือนเกินกำหนด {sent}/{len(overdue)} รายการ', 'success')

        cfg = conn.execute('SELECT * FROM email_settings LIMIT 1').fetchone()
        conn.close()
        return redirect(url_for('email_settings_page'))

    conn.close()
    return render_template('email_settings.html', cfg=cfg)


# ── SCAN QR ────────────────────────────────────────────────────────────────────

@app.route('/scan')
@login_required
def scan_qr():
    return render_template('scan_qr.html')


@app.route('/api/equipment/<int:eid>/qr')
@login_required
def api_qr(eid):
    conn = get_db()
    eq   = conn.execute('SELECT * FROM equipment WHERE id=?', (eid,)).fetchone()
    conn.close()
    if not eq:
        return jsonify({'error': 'not found'}), 404
    qr_data  = f"ID:{eid}|ชื่อ:{eq['name']}|Serial:{eq['serial_number']}|หมวด:{eq['category']}"
    return jsonify({'qr': generate_qr(qr_data), 'equipment': dict(eq)})


if __name__ == '__main__':
    init_db()
    app.run(debug=True, host='0.0.0.0', port=5000)
