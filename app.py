# -*- coding: utf-8 -*-
"""
Журнал регистрации приказов по видам деятельности
Веб-приложение (Flask + SQLite) с ролевым доступом, загрузкой PDF-сканов,
отчётами (Excel/Word/печать), архивированием и восстановлением БД.
"""
import os
import io
import json
import shutil
import socket
import sqlite3
import logging
import tempfile
import traceback
from logging.handlers import RotatingFileHandler
from datetime import date, datetime

from flask import (
    Flask, render_template, request, redirect, url_for, session, flash,
    send_file, abort, Response, g
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
except ImportError:
    Workbook = None

try:
    from docx import Document as DocxDocument
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT
except ImportError:
    DocxDocument = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'data', 'orders.db')
UPLOAD_DIR = os.path.join(BASE_DIR, 'data', 'scans')
ARCHIVE_DIR = os.path.join(BASE_DIR, 'data', 'archives')
VERSION_FILE = os.path.join(BASE_DIR, 'data', 'version.json')
LOG_DIR = os.path.join(BASE_DIR, 'data', 'logs')
LOG_FILE = os.path.join(LOG_DIR, 'app.log')

# Порт по умолчанию изменён на 8080, т.к. на сервере порт 5000 уже занят
# другой программой. Порт можно переопределить переменной окружения PORT
# или параметром командной строки: python3 app.py --port 8090
DEFAULT_PORT = 8080


def find_free_port(start):
    """Возвращает первый свободный TCP-порт, начиная с start (для 0.0.0.0)."""
    for port in range(start, start + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                s.bind(('0.0.0.0', port))
                return port
            except OSError:
                continue
    return None

APP_VERSION = '1.1.1'  # 1.1.0 — смена пароля сотрудником; 1.1.1 — исправлен NameError при сохранении вида приказа


def setup_logging():
    """Настройка журнала логирования: файл data/logs/app.log с ротацией
    (5 файлов по 2 МБ) + вывод в консоль."""
    os.makedirs(LOG_DIR, exist_ok=True)
    logger = logging.getLogger('orders_app')
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        fmt = logging.Formatter(
            '%(asctime)s | %(levelname)-8s | %(message)s', '%Y-%m-%d %H:%M:%S')
        fh = RotatingFileHandler(LOG_FILE, maxBytes=2_000_000,
                                 backupCount=5, encoding='utf-8')
        fh.setFormatter(fmt)
        logger.addHandler(fh)
        sh = logging.StreamHandler()
        sh.setFormatter(fmt)
        logger.addHandler(sh)
    return logger


log = setup_logging()


def audit(action, details=''):
    """Запись события аудита в журнал логирования (кто и что сделал)."""
    user = session.get('user_id')
    if user:
        try:
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            r = conn.execute('SELECT full_name, login FROM employees '
                             'WHERE id=?', (user,)).fetchone()
            conn.close()
            who = f'{r["full_name"]} ({r["login"]})' if r else f'id={user}'
        except Exception:
            who = f'id={user}'
    else:
        who = 'не авторизован'
    ip = request.remote_addr if has_request_context_safe() else '-'
    log.info(f'[АУДИТ] {action} | пользователь: {who} | IP: {ip} | {details}')


def has_request_context_safe():
    try:
        from flask import has_request_context
        return has_request_context()
    except Exception:
        return False

JOURNAL_TYPES = ['opeka', 'nuzhdaemost', 'profilaktika']
JOURNAL_TITLES = {
    'opeka': 'Опека',
    'nuzhdaemost': 'Нуждаемость',
    'profilaktika': 'Профилактика',
}
JOURNAL_SUFFIX = {
    'opeka': '-оп',
    'nuzhdaemost': '-ну',
    'profilaktika': '-пр',
}
ROLE_ACCESS = {
    'opeka': ('Администратор', 'Куратор', 'Опека'),
    'nuzhdaemost': ('Администратор', 'Куратор', 'Нуждаемость'),
    'profilaktika': ('Администратор', 'Куратор', 'Профилактика'),
}
ALL_ROLES = ['Администратор', 'Куратор', 'Опека', 'Нуждаемость', 'Профилактика']

app = Flask(__name__)


def _load_or_create_secret():
    p = os.path.join(BASE_DIR, 'data', 'secret_key.txt')
    os.makedirs(os.path.dirname(p), exist_ok=True)
    if os.path.exists(p):
        with open(p) as f:
            return f.read().strip()
    key = os.urandom(24).hex()
    with open(p, 'w') as f:
        f.write(key)
    return key


app.secret_key = _load_or_create_secret()


@app.after_request
def log_request(response):
    """Журнал логирования HTTP-запросов (каждый запрос)."""
    try:
        uid = session.get('user_id')
        status = response.status_code
        lvl = logging.WARNING if status >= 400 else logging.INFO
        # Ошибки доступа (403) и 404 пишем отдельной строкой для аудита
        log.log(lvl, f'[HTTP] {request.method} {request.path} -> {status} '
                     f'(IP: {request.remote_addr}, user_id: {uid or "-"})')
        if status == 403:
            audit('ОТКАЗ В ДОСТУПЕ (403)', f'{request.method} {request.path}')
        # Неудачные попытки входа пишем в журнал отдельно
        if (request.endpoint == 'login' and request.method == 'POST'
                and status == 200):
            log.warning(f'[БЕЗОПАСНОСТЬ] Неудачная попытка входа: логин '
                        f"'{request.form.get('login','')}' от IP "
                        f"{request.remote_addr}")
    except Exception:
        pass
    return response

# ---------------------------------------------------------------- DB helpers

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys = ON')
    return conn


SCHEMA = """
CREATE TABLE IF NOT EXISTS roles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL
);
CREATE TABLE IF NOT EXISTS employees (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    full_name TEXT NOT NULL,
    position TEXT NOT NULL DEFAULT '',
    login TEXT UNIQUE,
    password_hash TEXT,
    pwd_changed_at TEXT,           -- дата последней самостоятельной смены пароля
    role_id INTEGER REFERENCES roles(id)
);
CREATE TABLE IF NOT EXISTS order_types (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT NOT NULL,              -- вид приказа (название журнала)
    name TEXT NOT NULL,              -- наименование приказа
    UNIQUE(kind, name)
);
CREATE TABLE IF NOT EXISTS orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    journal TEXT NOT NULL,           -- opeka / nuzhdaemost / profilaktika
    order_number TEXT NOT NULL,      -- номер приказа, напр. "12-оп"
    order_date TEXT NOT NULL,        -- ГГГГ-ММ-ДД
    order_type_id INTEGER REFERENCES order_types(id),
    person_fio TEXT NOT NULL DEFAULT '',   -- ФИО в отношении кого
    specialist_fio TEXT NOT NULL DEFAULT '', -- ФИО специалиста
    scan_path TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS counters (
    journal TEXT PRIMARY KEY,
    value INTEGER NOT NULL DEFAULT 0
);
"""


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    conn = get_db()
    conn.executescript(SCHEMA)
    # Миграция: добавляем колонку pwd_changed_at в существующих базах
    cols = [c['name'] for c in conn.execute('PRAGMA table_info(employees)')]
    if 'pwd_changed_at' not in cols:
        conn.execute('ALTER TABLE employees ADD COLUMN pwd_changed_at TEXT')
    # seed roles
    for r in ALL_ROLES:
        conn.execute('INSERT OR IGNORE INTO roles(name) VALUES (?)', (r,))
    # seed admin if no users
    cnt = conn.execute('SELECT COUNT(*) c FROM employees').fetchone()['c']
    if cnt == 0:
        role = conn.execute("SELECT id FROM roles WHERE name='Администратор'").fetchone()
        conn.execute(
            'INSERT INTO employees(full_name, position, login, password_hash, role_id) '
            'VALUES (?,?,?,?,?)',
            ('Системный администратор', 'Администратор', 'admin',
             generate_password_hash('admin'), role['id']))
    # seed order types
    seeded = conn.execute("SELECT COUNT(*) c FROM order_types").fetchone()['c']
    if seeded == 0:
        demo = {
            'opeka': ['Назначение опекуна', 'Прекращение опеки',
                      'Помещение под опеку', 'Отстранение опекуна'],
            'nuzhdaemost': ['Признание нуждающимся в социальном обслуживании',
                            'Снятие с учёта нуждающихся',
                            'Продление статуса нуждающегося'],
            'profilaktika': ['Постановка на профилактический учёт',
                             'Снятие с профилактического учёта',
                             'Проведение профилактического мероприятия'],
        }
        for kind, names in demo.items():
            for n in names:
                conn.execute('INSERT OR IGNORE INTO order_types(kind,name) VALUES (?,?)',
                             (kind, n))
    for j in JOURNAL_TYPES:
        conn.execute('INSERT OR IGNORE INTO counters(journal,value) VALUES (?,0)', (j,))
    conn.commit()
    conn.close()


def write_version_file():
    os.makedirs(os.path.dirname(VERSION_FILE), exist_ok=True)
    with open(VERSION_FILE, 'w', encoding='utf-8') as f:
        json.dump({'version': APP_VERSION,
                   'updated': datetime.now().isoformat(timespec='seconds')}, f)


init_db()
write_version_file()

# ---------------------------------------------------------------- auth

def current_user():
    uid = session.get('user_id')
    if not uid:
        return None
    conn = get_db()
    u = conn.execute('SELECT e.*, r.name AS role_name FROM employees e '
                     'LEFT JOIN roles r ON r.id=e.role_id WHERE e.id=?', (uid,)).fetchone()
    conn.close()
    return u


def require_login():
    if not current_user():
        abort(redirect(url_for('login')))


def is_admin(user=None):
    user = user or current_user()
    return bool(user) and user['role_name'] == 'Администратор'


def can_access_journal(user, journal):
    if not user:
        return False
    if user['role_name'] == 'Администратор':
        return True
    return user['role_name'] in ROLE_ACCESS[journal]


@app.before_request
def _guard():
    open_exceptions = ('login', 'logout', 'static')
    if request.endpoint in open_exceptions:
        return
    if not session.get('user_id'):
        return redirect(url_for('login'))


@app.context_processor
def inject_globals():
    user = current_user()
    journals = [j for j in JOURNAL_TYPES if user and can_access_journal(user, j)]
    return dict(
        cu=user,
        accessible_journals=[(j, JOURNAL_TITLES[j]) for j in journals],
        all_journals=[(j, JOURNAL_TITLES[j]) for j in JOURNAL_TYPES],
        journal_titles=JOURNAL_TITLES,
        all_roles=ALL_ROLES,
        app_version=APP_VERSION,
    )

# ---------------------------------------------------------------- login

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        login_name = request.form.get('login', '').strip()
        password = request.form.get('password', '')
        conn = get_db()
        u = conn.execute('SELECT e.*, r.name role_name FROM employees e '
                         'LEFT JOIN roles r ON r.id=e.role_id '
                         'WHERE e.login=? AND e.password_hash IS NOT NULL',
                         (login_name,)).fetchone()
        conn.close()
        if u and check_password_hash(u['password_hash'], password):
            session['user_id'] = u['id']
            # При входе показываем подсказку о смене пароля, если сотрудник
            # ещё не менял пароль после его установки администратором.
            if not u['pwd_changed_at']:
                flash('Рекомендуем сменить пароль: перейдите в «Моя учётная запись».', 'warn')
            audit('ВХОД В СИСТЕМУ', f'роль: {u["role_name"]}')
            return redirect(url_for('index'))
        error = 'Неверный логин или пароль'
    if session.get('user_id'):
        return redirect(url_for('index'))
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    audit('ВЫХОД ИЗ СИСТЕМЫ')
    session.clear()
    return redirect(url_for('login'))

# ---------------------------------------------------------------- main page

@app.route('/')
def index():
    user = current_user()
    conn = get_db()
    stats = {}
    for j in JOURNAL_TYPES:
        if can_access_journal(user, j):
            stats[j] = conn.execute(
                'SELECT COUNT(*) c FROM orders WHERE journal=?', (j,)).fetchone()['c']
    conn.close()
    return render_template('index.html', stats=stats)

# ---------------------------------------------------------------- справочник сотрудников

@app.route('/employees')
def employees():
    require_login()
    conn = get_db()
    rows = conn.execute('SELECT e.*, r.name role_name FROM employees e '
                        'LEFT JOIN roles r ON r.id=e.role_id ORDER BY e.full_name').fetchall()
    conn.close()
    return render_template('employees.html', rows=rows)


@app.route('/employees/save', methods=['POST'])
def employee_save():
    user = current_user()
    if not is_admin(user):
        abort(403)
    eid = request.form.get('id', type=int)
    full_name = request.form.get('full_name', '').strip()
    position = request.form.get('position', '').strip()
    login_name = request.form.get('login', '').strip() or None
    role_name = request.form.get('role', 'Куратор')
    password = request.form.get('password', '')
    conn = get_db()
    role = conn.execute('SELECT id FROM roles WHERE name=?', (role_name,)).fetchone()
    role_id = role['id'] if role else None
    try:
        if eid:
            cur = conn.execute('SELECT * FROM employees WHERE id=?', (eid,)).fetchone()
            if not cur:
                abort(404)
            pw_hash = cur['password_hash']
            pwd_changed = cur['pwd_changed_at'] if 'pwd_changed_at' in cur.keys() else None
            if password:
                # администратор установил новый пароль — счётчик смены обнуляется,
                # сотруднику будет рекомендовано сменить его самостоятельно
                pw_hash = generate_password_hash(password)
                pwd_changed = None
            elif not login_name:
                pw_hash = None
            conn.execute(
                'UPDATE employees SET full_name=?, position=?, login=?, '
                'password_hash=?, pwd_changed_at=?, role_id=? WHERE id=?',
                (full_name, position, login_name, pw_hash, pwd_changed, role_id, eid))
            audit('ИЗМЕНЁН СОТРУДНИК', f'id={eid}, {full_name}')
        else:
            pw_hash = generate_password_hash(password) if (login_name and password) else None
            conn.execute(
                'INSERT INTO employees(full_name, position, login, password_hash, role_id) '
                'VALUES (?,?,?,?,?)',
                (full_name, position, login_name, pw_hash, role_id))
        conn.commit()
        audit('СОТРУДНИК СОХРАНЁН', full_name)
        flash('Сотрудник сохранён', 'ok')
    except sqlite3.IntegrityError:
        conn.rollback()
        audit('ОШИБКА СОХРАНЕНИЯ СОТРУДНИКА', f'занятый логин: {login_name}')
        flash('Ошибка: такой логин уже используется', 'err')
    finally:
        conn.close()
    return redirect(url_for('employees'))


@app.route('/employees/delete/<int:eid>', methods=['POST'])
def employee_delete(eid):
    if not is_admin():
        abort(403)
    conn = get_db()
    emp = conn.execute('SELECT * FROM employees WHERE id=?', (eid,)).fetchone()
    if emp and emp['id'] == session.get('user_id'):
        conn.close()
        flash('Нельзя удалить текущую учётную запись', 'err')
        return redirect(url_for('employees'))
    conn.execute('DELETE FROM employees WHERE id=?', (eid,))
    conn.commit()
    conn.close()
    audit('СОТРУДНИК УДАЛЁН', f'id={eid}')
    flash('Сотрудник удалён', 'ok')
    return redirect(url_for('employees'))

# ---------------------------------------------------------------- справочник видов приказов

@app.route('/order_types')
def order_types():
    require_login()
    conn = get_db()
    rows = conn.execute('SELECT * FROM order_types ORDER BY kind, name').fetchall()
    conn.close()
    return render_template('order_types.html', rows=rows)


@app.route('/order_types/save', methods=['POST'])
def order_type_save():
    user = current_user()
    if not is_admin(user):
        abort(403)
    tid = request.form.get('id', type=int)
    kind = request.form.get('kind', 'opeka')
    name = request.form.get('name', '').strip()
    if not name:
        flash('Укажите наименование приказа', 'err')
        return redirect(url_for('order_types'))
    conn = get_db()
    try:
        if tid:
            conn.execute('UPDATE order_types SET kind=?, name=? WHERE id=?', (kind, name, tid))
        else:
            conn.execute('INSERT INTO order_types(kind,name) VALUES (?,?)', (kind, name))
        conn.commit()
        # в этой функции переменная называется kind (вид приказа), а не journal —
        # исправлена ошибка NameError при аудите сохранения вида приказа
        audit('ВИД ПРИКАЗА СОХРАНЁН', f'{JOURNAL_TITLES.get(kind, kind)}: {name}')
        flash('Вид приказа сохранён', 'ok')
    except sqlite3.IntegrityError:
        conn.rollback()
        audit('ОШИБКА: ДУБЛЬ ВИДА ПРИКАЗА', f'{JOURNAL_TITLES.get(kind, kind)}: {name}')
        flash('Такой вид приказа уже существует', 'err')
    finally:
        conn.close()
    return redirect(url_for('order_types'))


@app.route('/order_types/delete/<int:tid>', methods=['POST'])
def order_type_delete(tid):
    if not is_admin():
        abort(403)
    conn = get_db()
    used = conn.execute('SELECT COUNT(*) c FROM orders WHERE order_type_id=?', (tid,)).fetchone()['c']
    if used:
        conn.close()
        flash('Нельзя удалить: вид приказа используется в документах', 'err')
    else:
        conn.execute('DELETE FROM order_types WHERE id=?', (tid,))
        conn.commit()
        conn.close()
        audit('ВИД ПРИКАЗА УДАЛЁН', f'id={tid}')
        flash('Удалено', 'ok')
    return redirect(url_for('order_types'))

# ---------------------------------------------------------------- журналы

def next_order_number(conn, journal):
    conn.execute('UPDATE counters SET value=value+1 WHERE journal=?', (journal,))
    row = conn.execute('SELECT value FROM counters WHERE journal=?', (journal,)).fetchone()
    if not row:
        conn.execute('INSERT INTO counters(journal,value) VALUES (?,1)', (journal,))
        row = conn.execute('SELECT value FROM counters WHERE journal=?', (journal,)).fetchone()
    return f"{row['value']}{JOURNAL_SUFFIX[journal]}"


@app.route('/journal/<journal>')
def journal_list(journal):
    if journal not in JOURNAL_TYPES:
        abort(404)
    user = current_user()
    if not can_access_journal(user, journal):
        abort(403)
    q = request.args.get('q', '').strip()
    conn = get_db()
    sql = ('SELECT o.*, t.name AS type_name FROM orders o '
           'LEFT JOIN order_types t ON t.id=o.order_type_id WHERE o.journal=?')
    params = [journal]
    if q:
        sql += ' AND (o.order_number LIKE ? OR o.person_fio LIKE ? OR o.specialist_fio LIKE ?)'
        like = f'%{q}%'
        params += [like, like, like]
    sql += (' ORDER BY o.order_date DESC, '
            "CAST(NULLIF(SUBSTR(o.order_number, 1, INSTR(o.order_number || '-', '-')), '') AS INTEGER) DESC")
    rows = conn.execute(sql, params).fetchall()
    types = conn.execute('SELECT * FROM order_types WHERE kind=? ORDER BY name',
                         (journal,)).fetchall()
    today = date.today().isoformat()
    conn.close()
    return render_template('journal.html', journal=journal, title=JOURNAL_TITLES[journal],
                           rows=rows, types=types, today=today, q=q)


@app.route('/journal/<journal>/api/next_number')
def api_next_number(journal):
    if journal not in JOURNAL_TYPES:
        abort(404)
    if not can_access_journal(current_user(), journal):
        abort(403)
    conn = get_db()
    num = next_order_number(conn, journal)
    conn.commit()
    conn.close()
    return {'number': num}


@app.route('/journal/<journal>/create', methods=['POST'])
def journal_create(journal):
    if journal not in JOURNAL_TYPES:
        abort(404)
    user = current_user()
    if not can_access_journal(user, journal):
        abort(403)
    conn = get_db()
    order_number = request.form.get('order_number', '').strip()
    if not order_number:
        order_number = next_order_number(conn, journal)
    else:
        # синхронизация счётчика, если пользователь ввёл больший номер
        base = order_number.split('-')[0]
        if base.isdigit():
            conn.execute('UPDATE counters SET value=MAX(value,?) WHERE journal=?',
                         (int(base), journal))
    order_date = request.form.get('order_date') or date.today().isoformat()
    type_id = request.form.get('order_type_id', type=int)
    person_fio = request.form.get('person_fio', '').strip()
    specialist_fio = request.form.get('specialist_fio', '').strip()
    if type_id:
        ok = conn.execute('SELECT 1 FROM order_types WHERE id=? AND kind=?',
                          (type_id, journal)).fetchone()
        if not ok:
            type_id = None
    scan_rel = None
    f = request.files.get('scan')
    if f and f.filename:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext != '.pdf':
            conn.close()
            audit('ОТКАЗ: СКАН НЕ В ФОРМАТЕ PDF', f.filename)
            flash('Скан должен быть в формате PDF', 'err')
            return redirect(url_for('journal_list', journal=journal))
        fname = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(f.filename)}"
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        f.save(os.path.join(UPLOAD_DIR, fname))
        scan_rel = fname
    conn.execute(
        'INSERT INTO orders(journal, order_number, order_date, order_type_id, '
        'person_fio, specialist_fio, scan_path) VALUES (?,?,?,?,?,?,?)',
        (journal, order_number, order_date, type_id, person_fio, specialist_fio, scan_rel))
    conn.commit()
    conn.close()
    audit('ПРИКАЗ СОЗДАН', f'журнал={JOURNAL_TITLES.get(journal)}, №={order_number}, по кому={person_fio}')
    flash('Приказ создан', 'ok')
    return redirect(url_for('journal_list', journal=journal))


@app.route('/order/<int:oid>/update', methods=['POST'])
def order_update(oid):
    conn = get_db()
    row = conn.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    if not row:
        conn.close()
        abort(404)
    journal = row['journal']
    if not can_access_journal(current_user(), journal):
        conn.close()
        abort(403)
    order_number = request.form.get('order_number', '').strip() or row['order_number']
    order_date = request.form.get('order_date') or row['order_date']
    type_id = request.form.get('order_type_id', type=int)
    person_fio = request.form.get('person_fio', '').strip()
    specialist_fio = request.form.get('specialist_fio', '').strip()
    scan_rel = row['scan_path']
    f = request.files.get('scan')
    if f and f.filename:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext != '.pdf':
            conn.close()
            audit('ОТКАЗ: СКАН НЕ В ФОРМАТЕ PDF', f.filename)
            flash('Скан должен быть в формате PDF', 'err')
            return redirect(url_for('journal_list', journal=journal))
        fname = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(f.filename)}"
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        f.save(os.path.join(UPLOAD_DIR, fname))
        if scan_rel and os.path.exists(os.path.join(UPLOAD_DIR, scan_rel)):
            os.remove(os.path.join(UPLOAD_DIR, scan_rel))
        scan_rel = fname
    conn.execute(
        'UPDATE orders SET order_number=?, order_date=?, order_type_id=?, '
        'person_fio=?, specialist_fio=?, scan_path=? WHERE id=?',
        (order_number, order_date, type_id, person_fio, specialist_fio, scan_rel, oid))
    conn.commit()
    conn.close()
    audit('ПРИКАЗ ИЗМЕНЁН', f'id={oid}')
    flash('Изменения сохранены', 'ok')
    return redirect(url_for('journal_list', journal=journal))


@app.route('/order/<int:oid>/delete', methods=['POST'])
def order_delete(oid):
    conn = get_db()
    row = conn.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    if not row:
        conn.close()
        abort(404)
    journal = row['journal']
    if not can_access_journal(current_user(), journal):
        conn.close()
        abort(403)
    if row['scan_path'] and os.path.exists(os.path.join(UPLOAD_DIR, row['scan_path'])):
        os.remove(os.path.join(UPLOAD_DIR, row['scan_path']))
    conn.execute('DELETE FROM orders WHERE id=?', (oid,))
    conn.commit()
    conn.close()
    audit('ПРИКАЗ УДАЛЁН', f'id={oid}')
    flash('Приказ удалён', 'ok')
    return redirect(url_for('journal_list', journal=journal))


@app.route('/order/<int:oid>/scan')
def order_scan(oid):
    conn = get_db()
    row = conn.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    conn.close()
    if not row or not row['scan_path']:
        abort(404)
    if not can_access_journal(current_user(), row['journal']):
        abort(403)
    path = os.path.join(UPLOAD_DIR, os.path.basename(row['scan_path']))
    if not os.path.exists(path):
        abort(404)
    return send_file(path, mimetype='application/pdf')

# ---------------------------------------------------------------- отчёты

def _report_period():
    d1 = request.args.get('d1') or date.today().replace(day=1).isoformat()
    d2 = request.args.get('d2') or date.today().isoformat()
    return d1, d2


def _query_report_rows(d1, d2, journal=None):
    conn = get_db()
    sql = ('SELECT o.*, t.name AS type_name FROM orders o '
           'LEFT JOIN order_types t ON t.id=o.order_type_id '
           'WHERE o.order_date>=? AND o.order_date<=?')
    params = [d1, d2]
    if journal:
        sql += ' AND o.journal=?'
        params.append(journal)
    sql += ' ORDER BY o.order_date, o.order_number'
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return rows


@app.route('/reports')
def reports():
    user = current_user()
    if not user or user['role_name'] not in ('Администратор', 'Куратор'):
        abort(403)
    return render_template('reports.html')


@app.route('/reports/by_period')
def report_by_period():
    user = current_user()
    if not user or user['role_name'] not in ('Администратор', 'Куратор'):
        abort(403)
    d1, d2 = _report_period()
    rows = _query_report_rows(d1, d2)
    agg = {}
    for r in rows:
        agg[r['journal']] = agg.get(r['journal'], 0) + 1
    data = [(JOURNAL_TITLES[j], agg.get(j, 0)) for j in JOURNAL_TYPES]
    fmt = request.args.get('fmt')
    if fmt == 'xlsx':
        return _export_xlsx_simple(d1, d2, data, 'Приказы за период')
    if fmt == 'docx':
        return _export_docx_simple(d1, d2, data, 'Приказы за период')
    return render_template('report_table.html', title='Приказы за период',
                           d1=d1, d2=d2, headers=('Вид приказа', 'Количество приказов'),
                           rows=data, total=sum(x[1] for x in data),
                           base_url=url_for('report_by_period'))


@app.route('/reports/by_period_detail')
def report_by_period_detail():
    user = current_user()
    if not user or user['role_name'] not in ('Администратор', 'Куратор'):
        abort(403)
    d1, d2 = _report_period()
    rows = _query_report_rows(d1, d2)
    matrix = {}
    col_names = set()
    for r in rows:
        kind = JOURNAL_TITLES[r['journal']]
        name = r['type_name'] or '(не указан)'
        col_names.add(name)
        matrix.setdefault(kind, {})
        matrix[kind][name] = matrix[kind].get(name, 0) + 1
    cols = sorted(col_names)
    table = []
    for kind in [JOURNAL_TITLES[j] for j in JOURNAL_TYPES]:
        if kind in matrix:
            table.append([kind] + [matrix[kind].get(c, 0) for c in cols])
    totals = [sum(row[i] for row in table) for i in range(1, len(cols) + 1)] if table else []
    fmt = request.args.get('fmt')
    if fmt == 'xlsx':
        return _export_xlsx_matrix(d1, d2, ('Вид приказа',) + tuple(cols), table, totals,
                                   'Приказы за период с расшифровкой')
    if fmt == 'docx':
        return _export_docx_matrix(d1, d2, ('Вид приказа',) + tuple(cols), table, totals,
                                   'Приказы за период с расшифровкой')
    return render_template('report_matrix.html', title='Приказы за период с расшифровкой',
                           d1=d1, d2=d2, cols=cols, table=table, totals=totals,
                           base_url=url_for('report_by_period_detail'))


@app.route('/reports/journal_print')
def report_journal_print():
    user = current_user()
    if not user or user['role_name'] not in ('Администратор', 'Куратор'):
        abort(403)
    journal = request.args.get('journal', 'opeka')
    if journal not in JOURNAL_TYPES:
        abort(404)
    d1, d2 = _report_period()
    rows = _query_report_rows(d1, d2, journal=journal)
    title = f'Журнал «{JOURNAL_TITLES[journal]}» за период'
    data = []
    for i, r in enumerate(rows, 1):
        name = r['type_name'] or ''
        disp_name = (f'{name} ({r["person_fio"]})' if journal == 'opeka' else name)
        data.append((i, r['order_number'], r['order_date'], r['person_fio'],
                     disp_name, r['specialist_fio']))
    headers = ('№ п/п', '№ приказа', 'Дата', 'Ф.И.О.', 'Наименование приказа', 'ФИО специалиста')
    fmt = request.args.get('fmt')
    if fmt == 'xlsx':
        return _export_xlsx_table(d1, d2, headers, data, title)
    if fmt == 'docx':
        return _export_docx_table(d1, d2, headers, data, title)
    return render_template('report_table.html', title=title, d1=d1, d2=d2,
                           headers=headers, rows=data, total=len(data),
                           base_url=url_for('report_journal_print'),
                           extra_args={'journal': journal})

# ---------------------------------------------------------------- export helpers

def _mem_download(stream, mimetype, filename):
    stream.seek(0)
    return send_file(stream, mimetype=mimetype, as_attachment=True, download_name=filename)


def _need_xlsx():
    if Workbook is None:
        abort(500, description='Не установлен модуль openpyxl (pip install openpyxl)')


def _need_docx():
    if DocxDocument is None:
        abort(500, description='Не установлен модуль python-docx (pip install python-docx)')


def _export_xlsx_simple(d1, d2, data, title):
    _need_xlsx()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Отчёт'
    ws.append([title])
    ws.append([f'Период: с {d1} по {d2}'])
    ws.append([])
    ws.append(['Вид приказа', 'Количество приказов'])
    for row in data:
        ws.append(list(row))
    ws.append(['Итого', sum(r[1] for r in data)])
    for c in ws[1][0], ws[5][0], ws[5][1]:
        c.font = Font(bold=True)
    buf = io.BytesIO()
    wb.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         'report_period.xlsx')


def _export_xlsx_matrix(d1, d2, headers, table, totals, title):
    _need_xlsx()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Отчёт'
    ws.append([title])
    ws.append([f'Период: с {d1} по {d2}'])
    ws.append([])
    ws.append(list(headers))
    for row in table:
        ws.append(list(row))
    if totals:
        ws.append(['Итого'] + list(totals))
    for cell in ws[len(table) + 4]:
        cell.font = Font(bold=True)
    buf = io.BytesIO()
    wb.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         'report_period_detail.xlsx')


def _export_xlsx_table(d1, d2, headers, data, title):
    _need_xlsx()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Журнал'
    ws.append([title])
    ws.append([f'Период: с {d1} по {d2}'])
    ws.append([])
    ws.append(list(headers))
    for row in data:
        ws.append(list(row))
    for cell in ws[4]:
        cell.font = Font(bold=True)
    widths = [6, 12, 12, 30, 45, 30]
    for i, w in enumerate(widths[:len(headers)], 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         'report_journal.xlsx')


def _export_docx_simple(d1, d2, data, title):
    _need_docx()
    doc = DocxDocument()
    doc.add_heading(title, level=1)
    doc.add_paragraph(f'Период: с {d1} по {d2}')
    t = doc.add_table(rows=1, cols=2)
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    hdr[0].text, hdr[1].text = 'Вид приказа', 'Количество приказов'
    for row in data:
        cells = t.add_row().cells
        cells[0].text, cells[1].text = str(row[0]), str(row[1])
    cells = t.add_row().cells
    cells[0].text, cells[1].text = 'Итого', str(sum(r[1] for r in data))
    buf = io.BytesIO()
    doc.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         'report_period.docx')


def _export_docx_matrix(d1, d2, headers, table, totals, title):
    _need_docx()
    doc = DocxDocument()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    doc.add_heading(title, level=1)
    doc.add_paragraph(f'Период: с {d1} по {d2}')
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    for i, h in enumerate(headers):
        t.rows[0].cells[i].text = str(h)
    for row in table:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
    if totals:
        cells = t.add_row().cells
        cells[0].text = 'Итого'
        for i, v in enumerate(totals):
            cells[i + 1].text = str(v)
    buf = io.BytesIO()
    doc.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         'report_period_detail.docx')


def _export_docx_table(d1, d2, headers, data, title):
    _need_docx()
    doc = DocxDocument()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    doc.add_heading(title, level=1)
    doc.add_paragraph(f'Период: с {d1} по {d2}')
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    for i, h in enumerate(headers):
        t.rows[0].cells[i].text = str(h)
    for row in data:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
    buf = io.BytesIO()
    doc.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         'report_journal.docx')

# ---------------------------------------------------------------- сервис (архив/восстановление/обновление)

@app.route('/service')
def service():
    if not is_admin():
        abort(403)
    archives = []
    if os.path.isdir(ARCHIVE_DIR):
        for fn in sorted(os.listdir(ARCHIVE_DIR), reverse=True):
            if fn.endswith('.zip'):
                p = os.path.join(ARCHIVE_DIR, fn)
                archives.append({'name': fn,
                                 'size': f"{os.path.getsize(p)/1024:.0f} КБ",
                                 'date': datetime.fromtimestamp(os.path.getmtime(p))
                                         .strftime('%d.%m.%Y %H:%M')})
    return render_template('service.html', archives=archives)


@app.route('/service/backup', methods=['POST'])
def service_backup():
    if not is_admin():
        abort(403)
    stamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    tmp = tempfile.mkdtemp(prefix='backup_')
    try:
        dbcopy = os.path.join(tmp, 'orders.db')
        src = get_db()
        dst = sqlite3.connect(dbcopy)
        src.backup(dst)
        src.close()
        dst.close()
        scans_dir = os.path.join(tmp, 'scans')
        if os.path.isdir(UPLOAD_DIR):
            shutil.copytree(UPLOAD_DIR, scans_dir)
        zipname = f'backup_{stamp}.zip'
        zpath = os.path.join(ARCHIVE_DIR, zipname)
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        shutil.make_archive(zpath[:-4], 'zip', tmp)
        audit('РЕЗЕРВНАЯ КОПИЯ СОЗДАНА', zipname)
        flash(f'Архив создан: {zipname}', 'ok')
    except Exception as e:
        flash(f'Ошибка создания архива: {e}', 'err')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return redirect(url_for('service'))


@app.route('/service/download/<name>')
def service_download(name):
    if not is_admin():
        abort(403)
    p = os.path.join(ARCHIVE_DIR, os.path.basename(name))
    if not os.path.exists(p):
        abort(404)
    return send_file(p, as_attachment=True, download_name=os.path.basename(p))


@app.route('/service/delete_archive/<name>', methods=['POST'])
def service_delete_archive(name):
    if not is_admin():
        abort(403)
    p = os.path.join(ARCHIVE_DIR, os.path.basename(name))
    if os.path.exists(p):
        os.remove(p)
        audit('АРХИВ УДАЛЁН', name)
        flash('Архив удалён', 'ok')
    return redirect(url_for('service'))


@app.route('/service/restore', methods=['POST'])
def service_restore():
    if not is_admin():
        abort(403)
    f = request.files.get('archive')
    if not f or not f.filename.lower().endswith('.zip'):
        flash('Выберите ZIP-архив для восстановления', 'err')
        return redirect(url_for('service'))
    tmp = tempfile.mkdtemp(prefix='restore_')
    try:
        apath = os.path.join(tmp, 'upload.zip')
        f.save(apath)
        extract = os.path.join(tmp, 'ex')
        shutil.unpack_archive(apath, extract, 'zip')
        db_in = os.path.join(extract, 'orders.db')
        if not os.path.exists(db_in):
            raise FileNotFoundError('В архиве отсутствует orders.db')
        # safety: make current backup first
        stamp = datetime.now().strftime('pre_restore_%Y-%m-%d_%H-%M-%S')
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        cur_zip = os.path.join(ARCHIVE_DIR, f'backup_{stamp}.zip')
        pre = tempfile.mkdtemp(prefix='pre_')
        try:
            s = get_db(); d = sqlite3.connect(os.path.join(pre, 'orders.db'))
            s.backup(d); s.close(); d.close()
            if os.path.isdir(UPLOAD_DIR):
                shutil.copytree(UPLOAD_DIR, os.path.join(pre, 'scans'))
            shutil.make_archive(cur_zip[:-4], 'zip', pre)
        finally:
            shutil.rmtree(pre, ignore_errors=True)
        # restore db
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        shutil.copy(db_in, DB_PATH)
        # restore scans
        scans_in = os.path.join(extract, 'scans')
        if os.path.isdir(scans_in):
            if os.path.isdir(UPLOAD_DIR):
                shutil.rmtree(UPLOAD_DIR)
            shutil.copytree(scans_in, UPLOAD_DIR)
        audit('ВОССТАНОВЛЕНИЕ ИЗ АРХИВА', f.filename)
        flash('База данных восстановлена из архива', 'ok')
    except Exception as e:
        flash(f'Ошибка восстановления: {e}', 'err')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return redirect(url_for('service'))


@app.route('/service/update', methods=['POST'])
def service_update():
    """Обновление приложения из ZIP-архива новой версии (только Администратор)."""
    if not is_admin():
        abort(403)
    f = request.files.get('package')
    if not f or not f.filename.lower().endswith('.zip'):
        flash('Загрузите ZIP-архив обновления', 'err')
        return redirect(url_for('service'))
    tmp = tempfile.mkdtemp(prefix='update_')
    try:
        upath = os.path.join(tmp, 'pkg.zip')
        f.save(upath)
        ex = os.path.join(tmp, 'ex')
        shutil.unpack_archive(upath, ex, 'zip')
        applied = 0
        for root, dirs, files in os.walk(ex):
            rel = os.path.relpath(root, ex)
            if rel.startswith('data'):
                continue  # не трогаем базу и сканы
            target = os.path.normpath(os.path.join(BASE_DIR, rel))
            os.makedirs(target, exist_ok=True)
            for fn in files:
                if fn in ('orders.db', 'secret_key.txt', 'version.json'):
                    continue
                srcf = os.path.join(root, fn)
                dstf = os.path.join(target, fn)
                if os.path.commonpath([dstf, BASE_DIR]) != BASE_DIR:
                    continue
                shutil.copy2(srcf, dstf)
                applied += 1
        write_version_file()
        audit('ОБНОВЛЕНИЕ ПРОГРАММЫ ПРИМЕНЕНО', f'файлов: {applied}')
        flash(f'Обновление применено: файлов — {applied}. Перезапустите сервер при необходимости.',
              'ok')
    except Exception as e:
        flash(f'Ошибка обновления: {e}', 'err')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return redirect(url_for('service'))

# ---------------------------------------------------------------- profile / password

@app.route('/profile', methods=['GET', 'POST'])
def profile():
    """Личная страница сотрудника: информация + самостоятельная смена пароля."""
    user = current_user()
    if request.method == 'POST':
        old = request.form.get('old_password', '')
        new = request.form.get('new_password', '')
        confirm = request.form.get('confirm_password', '')
        conn = get_db()
        u = conn.execute('SELECT * FROM employees WHERE id=?', (user['id'],)).fetchone()
        if not u['password_hash'] or not check_password_hash(u['password_hash'], old):
            conn.close()
            audit('ПАРОЛЬ НЕ ИЗМЕНЁН', 'неверно указан текущий пароль')
            flash('Текущий пароль указан неверно', 'err')
            return redirect(url_for('profile'))
        if len(new) < 4:
            conn.close()
            flash('Новый пароль слишком короткий (минимум 4 символа)', 'err')
            return redirect(url_for('profile'))
        if new != confirm:
            conn.close()
            flash('Новый пароль и подтверждение не совпадают', 'err')
            return redirect(url_for('profile'))
        if new == old:
            conn.close()
            flash('Новый пароль не должен совпадать с текущим', 'err')
            return redirect(url_for('profile'))
        conn.execute(
            'UPDATE employees SET password_hash=?, pwd_changed_at=? WHERE id=?',
            (generate_password_hash(new),
             datetime.now().strftime('%Y-%m-%d %H:%M:%S'), user['id']))
        conn.commit()
        conn.close()
        audit('ПАРОЛЬ ИЗМЕНЁН', 'собственный профиль (сотрудник сменил пароль сам)')
        flash('Пароль успешно изменён', 'ok')
        return redirect(url_for('profile'))
    return render_template('profile.html')


@app.errorhandler(403)
def err403(e):
    return render_template('error.html', code=403, msg='Доступ запрещён. Недостаточно прав.'), 403


@app.errorhandler(404)
def err404(e):
    return render_template('error.html', code=404, msg='Страница не найдена.'), 404


@app.errorhandler(Exception)
def err500(e):
    """Любое необработанное исключение пишется в журнал логирования."""
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return e  # 403/404 и т.п. отдаём штатными обработчиками
    log.error(f'НЕОБРАБОТАННАЯ ОШИБКА при {request.method} {request.path}: '
              f'{e}\n{traceback.format_exc()}')
    return render_template('error.html', code=500,
                           msg='Внутренняя ошибка сервера. Подробности — в журнале логов.'), 500

# ------------------------------------------------------------- просмотр логов

LOG_TAIL_MAX = 2000  # максимум строк, выводимых на страницу журнала


def _read_log_tail(n):
    """Читает последние n строк основного файла лога (с учётом ротации)."""
    try:
        with open(LOG_FILE, encoding='utf-8', errors='replace') as f:
            lines = f.readlines()[-n:]
        return [l.rstrip('\n') for l in lines]
    except FileNotFoundError:
        return []


@app.route('/logs')
def logs_view():
    """Просмотр журнала логирования — только Администратор."""
    if not is_admin():
        abort(403)
    try:
        n = min(int(request.args.get('n', 200)), LOG_TAIL_MAX)
    except ValueError:
        n = 200
    lines = _read_log_tail(n)
    files = sorted([fn for fn in os.listdir(LOG_DIR) if fn.startswith('app.log')]
                   ) if os.path.isdir(LOG_DIR) else []
    return render_template('logs.html', lines=lines, n=n, files=files,
                           log_file=LOG_FILE)


@app.route('/logs/download')
def logs_download():
    """Скачивание файла журнала — только Администратор."""
    if not is_admin():
        abort(403)
    name = os.path.basename(request.args.get('file', 'app.log'))
    p = os.path.join(LOG_DIR, name)
    if not os.path.isfile(p):
        abort(404)
    return send_file(p, as_attachment=True, download_name=name)


if __name__ == '__main__':
    import sys
    host = os.environ.get('HOST', '0.0.0.0')
    port = None
    # Порт можно задать: аргументом --port, переменной окружения PORT
    # или значением по умолчанию DEFAULT_PORT (8080, т.к. 5000 занят другой
    # программой на этом сервере). Если порт занят — автоматически
    # подбирается ближайший свободный.
    args = sys.argv[1:]
    if '--port' in args:
        try:
            port = int(args[args.index('--port') + 1])
        except (IndexError, ValueError):
            print('Неверное значение --port'); sys.exit(1)
    elif os.environ.get('PORT'):
        port = int(os.environ['PORT'])
    else:
        port = DEFAULT_PORT
    requested_port = port
    port = find_free_port(port)
    if port is None:
        print(f'Не найден свободный порт начиная с {requested_port}. '
              f'Запустите с другим портом: python3 app.py --port <порт>')
        sys.exit(1)
    if port != requested_port:
        print(f'Порт {requested_port} занят другой программой, '
              f'использую свободный порт {port}')
        log.warning(f'Запрошенный порт {requested_port} занят, '
                    f'приложение запущено на порту {port}')
    log.info(f'Сервер запущен: http://{host}:{port} '
             f'(версия {APP_VERSION}, логи: {LOG_FILE})')
    print(f'Сервер запущен: http://{host}:{port}  (логин: admin / admin)')
    print(f'Журнал логирования: {LOG_FILE}')
    app.run(host=host, port=port, debug=False)
