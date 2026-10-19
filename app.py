import calendar
import csv
import io
import logging
import os
import secrets
import uuid
from datetime import datetime, timedelta
from urllib.parse import urlparse
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session, make_response, abort, send_from_directory
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from flask_mail import Mail, Message
from apscheduler.schedulers.background import BackgroundScheduler
import xlsxwriter
import pandas as pd
from sqlalchemy import text
import i18n
from i18n import TRANSLATIONS, t as tr, flash_t, status_label, priority_label, localized_user_name, update_content, normalize_status, is_system_update

APP_VERSION = '3.0.0'
APP_NAME = 'Lotus Task Manager'
APP_PORT = 5000

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ALLOWED_EXTENSIONS = {
    'pdf', 'doc', 'docx', 'xls', 'xlsx', 'csv', 'ppt', 'pptx', 'txt',
    'png', 'jpg', 'jpeg', 'gif', 'webp', 'zip', 'rar', '7z', 'msg', 'eml',
}
MIN_PASSWORD_LENGTH = 4

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(name)s: %(message)s')
log = logging.getLogger('lotus')

app = Flask(__name__)


def _load_secret_key():
    # Prefer the environment; otherwise keep a random key in instance/ so sessions survive restarts.
    key = os.environ.get('SECRET_KEY')
    if key:
        return key
    os.makedirs(app.instance_path, exist_ok=True)
    key_file = os.path.join(app.instance_path, 'secret_key')
    if os.path.exists(key_file):
        with open(key_file) as fh:
            key = fh.read().strip()
        if key:
            return key
    key = secrets.token_hex(32)
    with open(key_file, 'w') as fh:
        fh.write(key)
    return key


app.config['SECRET_KEY'] = _load_secret_key()
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///taskmanager.db')
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'static', 'uploads')
app.config['MAX_CONTENT_LENGTH'] = int(os.environ.get('MAX_UPLOAD_MB', '25')) * 1024 * 1024
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=12)

# إعدادات البريد (من متغيرات البيئة)
app.config['MAIL_SERVER'] = os.environ.get('MAIL_SERVER', 'smtp.office365.com')
app.config['MAIL_PORT'] = int(os.environ.get('MAIL_PORT', '587'))
app.config['MAIL_USE_TLS'] = os.environ.get('MAIL_USE_TLS', '1') == '1'
app.config['MAIL_USERNAME'] = os.environ.get('MAIL_USERNAME', '')
app.config['MAIL_PASSWORD'] = os.environ.get('MAIL_PASSWORD', '')
app.config['MAIL_TIMEOUT'] = 10
mail = Mail(app)

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

AVAILABLE_ROLES = sorted(['CEO', 'Section Head', 'Manager', 'Second line manager', 'Employee', 'admin'])

DEFAULT_FEATURE_VISIBILITY = {
    'dashboard': AVAILABLE_ROLES,
    'create_task': AVAILABLE_ROLES,
    'tasks_list': AVAILABLE_ROLES,
    'reports': ['admin', 'CEO', 'Manager', 'Section Head'],
    'users_manage': ['admin', 'CEO'],
    'departments_manage': ['admin', 'CEO'],
    'hierarchy_manage': ['admin', 'CEO'],
    'import_users': ['admin', 'CEO'],
}

FEATURE_LABELS = {
    'dashboard': {'ar': 'لوحة التحكم', 'en': 'Dashboard'},
    'create_task': {'ar': 'إنشاء مهمة', 'en': 'Create Task'},
    'tasks_list': {'ar': 'قائمة المهام', 'en': 'Task List'},
    'reports': {'ar': 'التقارير والتصدير', 'en': 'Reports & Export'},
    'users_manage': {'ar': 'إدارة الموظفين', 'en': 'Manage Users'},
    'departments_manage': {'ar': 'إدارة الأقسام', 'en': 'Manage Departments'},
    'hierarchy_manage': {'ar': 'قواعد التسلسل الهرمي', 'en': 'Hierarchy Rules'},
    'import_users': {'ar': 'استيراد الموظفين', 'en': 'Import Users'},
}

# --- النماذج (Models) ---
class Department(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    head_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    head = db.relationship('User', foreign_keys=[head_id])
    users = db.relationship('User', backref='department', lazy=True, foreign_keys='User.department_id')

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(100))
    username = db.Column(db.String(50), unique=True)
    email = db.Column(db.String(120), nullable=True)
    password = db.Column(db.String(200))
    role = db.Column(db.String(50)) 
    is_active = db.Column(db.Boolean, default=True)
    department_id = db.Column(db.Integer, db.ForeignKey('department.id'), nullable=True)

class HierarchyRule(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    from_role = db.Column(db.String(50))
    to_role = db.Column(db.String(50))
    cc_role = db.Column(db.String(50), nullable=True)

class Task(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200))
    description = db.Column(db.Text)
    status = db.Column(db.String(50), default='In Progress')
    priority = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=datetime.now)
    opened_at = db.Column(db.DateTime, nullable=True)
    started_at = db.Column(db.DateTime, nullable=True)     
    completed_at = db.Column(db.DateTime, nullable=True)   
    deadline = db.Column(db.DateTime, nullable=True)
    attachment = db.Column(db.String(200), nullable=True)
    recurrence = db.Column(db.String(20), default='None')
    next_run = db.Column(db.DateTime, nullable=True)
    
    creator_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    assignee_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    head_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    
    creator = db.relationship('User', foreign_keys=[creator_id])
    assignee = db.relationship('User', foreign_keys=[assignee_id])
    head = db.relationship('User', foreign_keys=[head_id])
    updates = db.relationship('TaskUpdate', backref='task', lazy=True, cascade="all, delete-orphan", order_by='TaskUpdate.created_at')
    
    @property
    def is_overdue(self):
        if self.deadline and normalize_status(self.status) not in ['Completed', 'Canceled']:
            return datetime.now() > self.deadline
        return False

    @property
    def display_status(self):
        return normalize_status(self.status)

    @property
    def is_terminal(self):
        return normalize_status(self.status) in ['Completed', 'Canceled']

class TaskUpdate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    content = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.now)
    is_note = db.Column(db.Boolean, default=False)
    filename = db.Column(db.String(200), nullable=True)
    task_id = db.Column(db.Integer, db.ForeignKey('task.id'))
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    user = db.relationship('User')

class Notification(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    type = db.Column(db.String(50))
    title = db.Column(db.String(200))
    message = db.Column(db.Text)
    link = db.Column(db.String(200), nullable=True)
    task_id = db.Column(db.Integer, db.ForeignKey('task.id'), nullable=True)
    is_read = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.now)
    user = db.relationship('User', backref='notifications')

class FeatureVisibility(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    feature_key = db.Column(db.String(50), unique=True, nullable=False)
    allowed_roles = db.Column(db.Text)

class AppSetting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(50), unique=True, nullable=False)
    value = db.Column(db.String(200))

TERMINAL_STATUSES = ['Completed', 'Canceled', 'Closed', 'Closed_by_System', 'Overdue_Closed']
LEGACY_CANCELED = ['Canceled', 'Closed', 'Closed_by_System', 'Overdue_Closed']
ACTIVE_STATUSES = ['In Progress', 'New', 'Under Review']

def get_setting(key, default=''):
    row = AppSetting.query.filter_by(key=key).first()
    return row.value if row else default

def set_setting(key, value):
    row = AppSetting.query.filter_by(key=key).first()
    if row:
        row.value = str(value)
    else:
        db.session.add(AppSetting(key=key, value=str(value)))

def get_auto_cancel_minutes():
    try:
        return max(1, int(get_setting('auto_cancel_minutes', '10')))
    except ValueError:
        return 10

i18n.auto_cancel_minutes_provider = get_auto_cancel_minutes

def seed_app_settings():
    if not AppSetting.query.filter_by(key='auto_cancel_minutes').first():
        db.session.add(AppSetting(key='auto_cancel_minutes', value='10'))
    db.session.commit()

@login_manager.user_loader
def load_user(user_id):
    user = db.session.get(User, int(user_id))
    # Suspended users are logged out on their next request.
    return user if user and user.is_active else None

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or current_user.role not in ['admin', 'CEO']:
            flash(tr('access_denied'), 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def user_can_see(feature_key):
    if not current_user.is_authenticated:
        return False
    setting = FeatureVisibility.query.filter_by(feature_key=feature_key).first()
    if not setting:
        default_roles = DEFAULT_FEATURE_VISIBILITY.get(feature_key, AVAILABLE_ROLES)
        return current_user.role in default_roles
    roles = [r.strip() for r in setting.allowed_roles.split(',') if r.strip()]
    return current_user.role in roles

def create_notification(user_id, ntype, title, message, link=None, task_id=None):
    if not user_id:
        return
    db.session.add(Notification(
        user_id=user_id, type=ntype, title=title,
        message=message, link=link, task_id=task_id
    ))

def mail_enabled():
    return bool(app.config.get('MAIL_USERNAME') and app.config.get('MAIL_PASSWORD'))

def send_email(subject, recipients, body, cc=None):
    recipients = [r for r in (recipients or []) if r]
    if not recipients or not mail_enabled():
        return
    try:
        msg = Message(subject, sender=app.config['MAIL_USERNAME'], recipients=recipients, cc=[c for c in (cc or []) if c])
        msg.body = body
        mail.send(msg)
    except Exception:
        log.exception('Failed to send email: %s', subject)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def save_upload(file, prefix=''):
    """Save an uploaded file. Returns the stored name, None if no file, or False if rejected."""
    if not file or not file.filename:
        return None
    safe = secure_filename(file.filename)
    if not safe or not allowed_file(safe):
        return False
    fname = f"{prefix}{datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:6]}_{safe}"
    file.save(os.path.join(app.config['UPLOAD_FOLDER'], fname))
    return fname

def add_months(dt, months):
    month = dt.month - 1 + months
    year = dt.year + month // 12
    month = month % 12 + 1
    day = min(dt.day, calendar.monthrange(year, month)[1])
    return dt.replace(year=year, month=month, day=day)

def next_recurrence(dt, recurrence):
    if recurrence == 'Daily':
        return dt + timedelta(days=1)
    if recurrence == 'Weekly':
        return dt + timedelta(days=7)
    if recurrence == 'Monthly':
        return add_months(dt, 1)
    return None

def is_safe_redirect(target):
    if not target:
        return False
    parsed = urlparse(target)
    return not parsed.scheme and not parsed.netloc and target.startswith('/')

def back(default_endpoint='dashboard'):
    ref = request.referrer
    if ref:
        parsed = urlparse(ref)
        if parsed.netloc == request.host:
            return redirect(ref)
    return redirect(url_for(default_endpoint))

def seed_feature_visibility():
    for key, roles in DEFAULT_FEATURE_VISIBILITY.items():
        if not FeatureVisibility.query.filter_by(feature_key=key).first():
            db.session.add(FeatureVisibility(feature_key=key, allowed_roles=','.join(roles)))
    db.session.commit()

# --- Translations ---
@app.context_processor
def inject_translations():
    lang = session.get('lang', 'ar')
    tr_dict = TRANSLATIONS.get(lang, TRANSLATIONS['ar'])
    display_name = localized_user_name(current_user) if current_user.is_authenticated else ''
    return dict(
        lang=lang, t=tr_dict, user_can_see=user_can_see, display_name=display_name,
        app_version=APP_VERSION, app_name=APP_NAME,
        status_label=status_label, priority_label=priority_label, update_content=update_content,
        is_system_update=is_system_update,
    )

@app.route('/set_language/<lang_code>')
def set_language(lang_code):
    if lang_code in TRANSLATIONS:
        session['lang'] = lang_code
    return back()

# --- CSRF protection ---
def get_csrf_token():
    token = session.get('_csrf_token')
    if not token:
        token = secrets.token_urlsafe(32)
        session['_csrf_token'] = token
    return token

@app.context_processor
def inject_csrf():
    return dict(csrf_token=get_csrf_token)

@app.before_request
def security_checks():
    # Attachments are only served to logged-in users through /uploads/<name>.
    if request.endpoint == 'static' and (request.view_args or {}).get('filename', '').replace('\\', '/').startswith('uploads/'):
        abort(404)
    if request.method == 'POST' and not app.config.get('TESTING_SKIP_CSRF'):
        sent = request.form.get('csrf_token') or request.headers.get('X-CSRFToken')
        expected = session.get('_csrf_token')
        if not sent or not expected or not secrets.compare_digest(sent, expected):
            if request.path.startswith('/api/'):
                return jsonify({'ok': False, 'error': 'csrf'}), 400
            flash_t('session_expired', 'warning')
            return back('login')

@app.after_request
def security_headers(response):
    response.headers.setdefault('X-Content-Type-Options', 'nosniff')
    response.headers.setdefault('X-Frame-Options', 'SAMEORIGIN')
    response.headers.setdefault('Referrer-Policy', 'same-origin')
    return response

@app.route('/uploads/<path:filename>')
@login_required
def uploaded_file(filename):
    image_types = ('.png', '.jpg', '.jpeg', '.gif', '.webp', '.pdf')
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename,
                               as_attachment=not filename.lower().endswith(image_types))

@app.errorhandler(413)
def file_too_large(_e):
    flash_t('file_too_large', 'danger')
    return back()

# --- الجدولة التلقائية الدقيقة (كل دقيقة) ---
def automated_system_checks():
    with app.app_context():
        try:
            _run_system_checks()
        except Exception:
            db.session.rollback()
            log.exception('Automated system checks failed')

def _run_system_checks():
    now = datetime.now()

    # 1. Cancel tasks the assignee never opened
    minutes = get_auto_cancel_minutes()
    cutoff = now - timedelta(minutes=minutes)
    neglected_tasks = Task.query.filter(
        Task.opened_at.is_(None),
        Task.status.in_(ACTIVE_STATUSES),
        Task.created_at <= cutoff
    ).all()
    emails = []
    for task in neglected_tasks:
        task.status = 'Canceled'
        task.completed_at = now
        db.session.add(TaskUpdate(content=f"[SYS:auto_cancel:{minutes}]", is_note=True, task_id=task.id, user_id=task.creator_id))
        note = TRANSLATIONS['en']['system_note_auto_cancel'].format(minutes=minutes)
        if task.creator_id:
            create_notification(
                task.creator_id, 'task_update',
                f"{TRANSLATIONS['ar']['notif_auto_closed']} / {TRANSLATIONS['en']['notif_auto_closed']}",
                f"'{task.title}' — {note}",
                link=f'/tasks/{task.id}', task_id=task.id
            )
        if task.creator and task.creator.email:
            emails.append((
                f"Urgent: Auto-closed task #{task.id}", [task.creator.email],
                note + f"\nTask: {task.title}\nAssignee: {task.assignee.full_name if task.assignee else '--'}",
                [task.head.email] if (task.head and task.head.email) else [],
            ))

    # 2. Deadline alerts (notify only — status stays In Progress)
    overdue_tasks = Task.query.filter(
        Task.deadline < now,
        Task.status.in_(ACTIVE_STATUSES)
    ).all()
    for task in overdue_tasks:
        if TaskUpdate.query.filter_by(task_id=task.id, content='[SYS:deadline_notified]').first():
            continue
        db.session.add(TaskUpdate(content="[SYS:deadline_notified]", is_note=True, task_id=task.id, user_id=task.creator_id))
        title = f"{TRANSLATIONS['ar']['deadline_passed']} / {TRANSLATIONS['en']['deadline_passed']}"
        message = f"'{task.title}' — {TRANSLATIONS['en']['deadline_passed_msg']}"
        for uid in {task.creator_id, task.head_id, task.assignee_id} - {None}:
            create_notification(uid, 'deadline_passed', title, message, link=f'/tasks/{task.id}', task_id=task.id)
        if task.creator and task.creator.email:
            emails.append((
                f"Alert: Task deadline passed #{task.id}", [task.creator.email],
                TRANSLATIONS['en']['deadline_passed_msg'] + f"\nTask: {task.title}",
                [task.head.email] if (task.head and task.head.email) else [],
            ))

    # 3. المهام المتكررة
    recurring = Task.query.filter(Task.recurrence.in_(['Daily', 'Weekly', 'Monthly']), Task.next_run <= now).all()
    for rt in recurring:
        run_at = rt.next_run
        # Catch up past missed runs without spawning a copy for each one.
        while run_at and run_at <= now:
            run_at = next_recurrence(run_at, rt.recurrence)
        rt.next_run = run_at
        # The copy is a one-off task; only the original keeps the schedule.
        new_t = Task(title=rt.title, description=rt.description, priority=rt.priority, recurrence='None',
                     creator_id=rt.creator_id, assignee_id=rt.assignee_id, head_id=rt.head_id,
                     status='In Progress', deadline=run_at, attachment=rt.attachment)
        db.session.add(new_t)
        db.session.flush()
        if new_t.assignee_id:
            create_notification(
                new_t.assignee_id, 'new_task',
                f"{TRANSLATIONS['ar']['notif_recurring_task']} / {TRANSLATIONS['en']['notif_recurring_task']}",
                new_t.title, link=f'/tasks/{new_t.id}', task_id=new_t.id
            )

    db.session.commit()

    # Send mail after commit so a slow/failed SMTP server never blocks DB changes.
    for subject, recipients, body, cc in emails:
        send_email(subject, recipients, body, cc)

scheduler = BackgroundScheduler()
scheduler.add_job(func=automated_system_checks, trigger="interval", minutes=1, max_instances=1, coalesce=True)
if os.environ.get('LOTUS_DISABLE_SCHEDULER') != '1':
    scheduler.start()

# --- مسارات النظام الأساسية ---
@app.route('/')
def index(): return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated: return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        user = User.query.filter_by(username=username).first() if username else None
        if user and user.password and check_password_hash(user.password, password):
            if not user.is_active:
                flash_t('account_suspended', 'danger')
                return redirect(url_for('login'))
            session.permanent = True
            login_user(user)
            next_url = request.args.get('next')
            return redirect(next_url if is_safe_redirect(next_url) else url_for('dashboard'))
        flash_t('invalid_login', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout(): 
    logout_user()
    return redirect(url_for('login'))

@app.route('/change_my_password', methods=['POST'])
@login_required
def change_my_password():
    current_password = request.form.get('current_password') or ''
    new_password = request.form.get('new_password') or ''

    if not check_password_hash(current_user.password, current_password):
        flash_t('wrong_password', 'danger')
    elif len(new_password) < MIN_PASSWORD_LENGTH:
        flash_t('password_too_short', 'danger')
    else:
        current_user.password = generate_password_hash(new_password)
        db.session.commit()
        flash_t('password_changed', 'success')

    return back()

@app.route('/dashboard')
@login_required
def dashboard():
    if not user_can_see('dashboard'):
        return redirect(url_for('tasks'))
    stats = {
        'total_tasks': Task.query.count(),
        'in_progress': Task.query.filter(Task.status.in_(ACTIVE_STATUSES)).count(),
        'overdue': overdue_query().count(),
        'completed': Task.query.filter(Task.status == 'Completed').count(),
    }
    stats['completion_rate'] = round(stats['completed'] * 100 / stats['total_tasks']) if stats['total_tasks'] else 0
    recent = Task.query.order_by(Task.created_at.desc()).limit(8).all()
    my_tasks = Task.query.filter(Task.assignee_id == current_user.id, Task.status.in_(ACTIVE_STATUSES)) \
        .order_by(Task.deadline.is_(None), Task.deadline.asc(), Task.created_at.desc()).limit(10).all()
    assigned_by_me = Task.query.filter(Task.creator_id == current_user.id, Task.status.in_(ACTIVE_STATUSES)) \
        .order_by(Task.created_at.desc()).limit(10).all()

    # Per-department workload, grouped by the assignee's department.
    dept_stats = {}
    now = datetime.now()
    rows = db.session.query(Task.status, Task.deadline, User.department_id) \
        .outerjoin(User, Task.assignee_id == User.id).all()
    for status, deadline, dept_id in rows:
        d = dept_stats.setdefault(dept_id, {'open': 0, 'overdue': 0, 'completed': 0, 'total': 0})
        d['total'] += 1
        if status in ACTIVE_STATUSES:
            d['open'] += 1
            if deadline and deadline < now:
                d['overdue'] += 1
        elif status == 'Completed':
            d['completed'] += 1
    dept_names = {d.id: d.name for d in Department.query.all()}
    departments = sorted(
        ({'name': dept_names.get(k) or tr('no_department'), **v} for k, v in dept_stats.items()),
        key=lambda d: (-d['open'], d['name']))

    return render_template('dashboard.html', stats=stats, recent_tasks=recent, my_tasks=my_tasks,
                           assigned_by_me=assigned_by_me, departments=departments,
                           today_date=datetime.now().strftime('%Y-%m-%d'))

def overdue_query():
    return Task.query.filter(Task.deadline.isnot(None), Task.deadline < datetime.now(), Task.status.in_(ACTIVE_STATUSES))

# --- الإشعارات الحية ---
@app.route('/api/notifications')
@login_required
def get_notifications():
    since = request.args.get('since', 0, type=int)
    items = Notification.query.filter(
        Notification.user_id == current_user.id,
        Notification.id > since
    ).order_by(Notification.id.asc()).limit(30).all()
    latest = Notification.query.filter_by(user_id=current_user.id).order_by(Notification.id.desc()).first()
    return jsonify({
        'notifications': [{
            'id': n.id, 'type': n.type, 'title': n.title,
            'message': n.message, 'link': n.link or (f'/tasks/{n.task_id}' if n.task_id else '#'),
            'created_at': n.created_at.strftime('%Y-%m-%d %H:%M')
        } for n in items],
        'latest_id': latest.id if latest else since
    })

@app.route('/api/notifications/mark_read', methods=['POST'])
@login_required
def mark_notifications_read():
    Notification.query.filter_by(user_id=current_user.id, is_read=False).update({'is_read': True})
    db.session.commit()
    return jsonify({'ok': True})

@app.route('/api/task_notifications')
@login_required
def task_notifications():
    latest = Notification.query.filter_by(user_id=current_user.id).order_by(Notification.id.desc()).first()
    return jsonify({'latest_id': latest.id if latest else 0})

@app.route('/api/manager_notifications')
@login_required
def manager_notifications():
    latest = Notification.query.filter_by(user_id=current_user.id).order_by(Notification.id.desc()).first()
    return jsonify({'latest_id': latest.id if latest else 0})

# --- التقارير والتصدير ---
@app.route('/reports')
@login_required
def reports():
    if not user_can_see('reports'):
        flash_t('access_denied', 'danger')
        return redirect(url_for('dashboard'))
    return render_template('reports.html')

@app.route('/reports/export/<type>')
@login_required
def export_excel(type):
    if not user_can_see('reports'):
        abort(403)
    if type not in ('users', 'tasks', 'tasks_overdue'):
        abort(404)

    def cell(value):
        # Block spreadsheet formula injection from user-entered text.
        if isinstance(value, str) and value[:1] in ('=', '+', '-', '@'):
            return "'" + value
        return '' if value is None else value

    def fmt(dt):
        return dt.strftime('%Y-%m-%d %H:%M') if dt else ''

    output = io.StringIO()
    writer = csv.writer(output)
    if type == 'users':
        writer.writerow(['ID', 'Full Name', 'Username', 'Email', 'Role', 'Department', 'Active'])
        for u in User.query.order_by(User.full_name).all():
            dept_name = u.department.name if u.department else ''
            writer.writerow([cell(v) for v in (u.id, u.full_name, u.username, u.email, u.role, dept_name, 'Yes' if u.is_active else 'No')])
    else:
        writer.writerow(['ID', 'Title', 'Creator', 'Assignee', 'Status', 'Priority', 'Created', 'Started', 'Completed', 'Duration (Hrs)', 'Deadline', 'Overdue'])
        query = overdue_query() if type == 'tasks_overdue' else Task.query
        for t in query.order_by(Task.created_at.desc()).all():
            duration = round((t.completed_at - t.started_at).total_seconds() / 3600, 2) if (t.started_at and t.completed_at) else "N/A"
            writer.writerow([cell(v) for v in (
                t.id, t.title, t.creator.full_name if t.creator else '', t.assignee.full_name if t.assignee else '',
                t.display_status, t.priority, fmt(t.created_at), fmt(t.started_at), fmt(t.completed_at),
                duration, fmt(t.deadline), 'Yes' if t.is_overdue else 'No',
            )])

    response = make_response(output.getvalue().encode('utf-8-sig'))
    response.headers["Content-Disposition"] = f"attachment; filename={type}_{datetime.now().strftime('%Y%m%d')}.csv"
    response.headers["Content-type"] = "text/csv; charset=utf-8"
    return response

# --- إدارة الأقسام ---
@app.route('/admin/departments', methods=['GET'])
@login_required
@admin_required
def manage_departments():
    departments = Department.query.order_by(Department.name).all()
    heads = User.query.filter(User.role.in_(['Manager', 'Section Head', 'CEO', 'admin'])).order_by(User.full_name).all()
    return render_template('manage_departments.html', departments=departments, heads=heads)

@app.route('/admin/departments/add', methods=['POST'])
@login_required
@admin_required
def add_department():
    name = (request.form.get('name') or '').strip()
    head_id = request.form.get('head_id', type=int)
    if not name:
        flash_t('invalid_input', 'danger')
    elif Department.query.filter_by(name=name).first():
        flash_t('department_exists', 'danger')
    else:
        db.session.add(Department(name=name, head_id=head_id or None))
        db.session.commit()
        flash_t('department_added', 'success')
    return redirect(url_for('manage_departments'))

@app.route('/admin/departments/delete/<int:id>', methods=['POST'])
@login_required
@admin_required
def delete_department(id):
    dept = Department.query.get_or_404(id)
    for u in dept.users: u.department_id = None
    db.session.delete(dept)
    db.session.commit()
    flash_t('department_deleted', 'success')
    return redirect(url_for('manage_departments'))

# --- إدارة الموظفين ---
@app.route('/admin/users')
@login_required
@admin_required
def manage_users(): return render_template('manage_users.html', users=User.query.order_by(User.full_name).all(), departments=Department.query.order_by(Department.name).all(), roles=AVAILABLE_ROLES)

@app.route('/admin/add_user', methods=['POST'])
@login_required
@admin_required
def add_user():
    username = (request.form.get('username') or '').strip()
    full_name = (request.form.get('full_name') or '').strip()
    password = request.form.get('password') or ''
    role = request.form.get('role')
    if not username or not full_name or role not in AVAILABLE_ROLES:
        flash_t('invalid_input', 'danger')
    elif len(password) < MIN_PASSWORD_LENGTH:
        flash_t('password_too_short', 'danger')
    elif User.query.filter_by(username=username).first():
        flash_t('user_exists', 'danger')
    else:
        new_u = User(full_name=full_name, username=username, email=(request.form.get('email') or '').strip() or None,
                     password=generate_password_hash(password), role=role,
                     department_id=request.form.get('department_id', type=int) or None)
        db.session.add(new_u); db.session.commit()
        flash_t('user_added', 'success')
    return redirect(url_for('manage_users'))

@app.route('/admin/user/toggle_status/<int:id>', methods=['POST'])
@login_required
@admin_required
def toggle_user_status(id):
    if id == current_user.id:
        flash_t('cannot_suspend_self', 'danger')
        return redirect(url_for('manage_users'))
    user = User.query.get_or_404(id); user.is_active = not user.is_active; db.session.commit()
    return redirect(url_for('manage_users'))

@app.route('/admin/user/reset_password/<int:id>', methods=['POST'])
@login_required
@admin_required
def reset_password(id):
    user = User.query.get_or_404(id)
    new_pass = request.form.get('new_password') or ''
    if len(new_pass) < MIN_PASSWORD_LENGTH:
        flash_t('password_too_short', 'danger')
    else:
        user.password = generate_password_hash(new_pass); db.session.commit()
        flash_t('password_reset_done', 'success')
    return redirect(url_for('manage_users'))

@app.route('/admin/user/edit/<int:id>', methods=['POST'])
@login_required
@admin_required
def edit_user(id):
    user = User.query.get_or_404(id)
    new_username = request.form.get('username', '').strip()
    full_name = (request.form.get('full_name') or '').strip()
    role = request.form.get('role')
    if not new_username or not full_name:
        flash_t('invalid_input', 'danger')
        return redirect(url_for('manage_users'))
    existing = User.query.filter(User.username == new_username, User.id != id).first()
    if existing:
        flash_t('user_update_error', 'danger')
        return redirect(url_for('manage_users'))
    user.full_name = full_name
    user.username = new_username
    user.email = (request.form.get('email') or '').strip() or None
    if id != current_user.id and role in AVAILABLE_ROLES:
        user.role = role
    user.department_id = request.form.get('department_id', type=int) or None
    db.session.commit()
    flash_t('user_updated', 'success')
    return redirect(url_for('manage_users'))

@app.route('/admin/hierarchy', methods=['GET', 'POST'])
@login_required
@admin_required
def manage_hierarchy():
    if request.method == 'POST':
        from_role, to_role = request.form.get('from_role'), request.form.get('to_role')
        cc_role = request.form.get('cc_role') or None
        if from_role not in AVAILABLE_ROLES or to_role not in AVAILABLE_ROLES or (cc_role and cc_role not in AVAILABLE_ROLES):
            flash_t('invalid_input', 'danger')
        else:
            db.session.add(HierarchyRule(from_role=from_role, to_role=to_role, cc_role=cc_role)); db.session.commit()
            flash_t('rule_added', 'success')
        return redirect(url_for('manage_hierarchy'))
    return render_template('manage_hierarchy.html', rules=HierarchyRule.query.order_by(HierarchyRule.from_role, HierarchyRule.to_role).all(), roles=AVAILABLE_ROLES)

@app.route('/admin/hierarchy/delete/<int:id>', methods=['POST'])
@login_required
@admin_required
def delete_hierarchy_rule(id):
    rule = HierarchyRule.query.get_or_404(id)
    db.session.delete(rule)
    db.session.commit()
    flash_t('rule_deleted', 'success')
    return redirect(url_for('manage_hierarchy'))

@app.route('/admin/permissions', methods=['GET', 'POST'])
@login_required
@admin_required
def manage_permissions():
    if request.method == 'POST':
        for key in DEFAULT_FEATURE_VISIBILITY:
            roles = [r for r in request.form.getlist(f'feature_{key}') if r in AVAILABLE_ROLES]
            setting = FeatureVisibility.query.filter_by(feature_key=key).first()
            if setting:
                setting.allowed_roles = ','.join(roles)
            else:
                db.session.add(FeatureVisibility(feature_key=key, allowed_roles=','.join(roles)))
        db.session.commit()
        flash_t('permissions_saved', 'success')
        return redirect(url_for('manage_permissions'))
    seed_feature_visibility()
    settings = {s.feature_key: [r.strip() for r in s.allowed_roles.split(',') if r.strip()] for s in FeatureVisibility.query.all()}
    return render_template('manage_permissions.html', settings=settings, roles=AVAILABLE_ROLES, features=FEATURE_LABELS)

@app.route('/admin/settings', methods=['GET', 'POST'])
@login_required
@admin_required
def manage_settings():
    if request.method == 'POST':
        try:
            minutes = max(1, min(1440, int(request.form.get('auto_cancel_minutes', 10))))
            set_setting('auto_cancel_minutes', minutes)
            db.session.commit()
            flash_t('settings_saved', 'success')
        except ValueError:
            flash_t('settings_invalid', 'danger')
        return redirect(url_for('manage_settings'))
    seed_app_settings()
    return render_template('manage_settings.html', auto_cancel_minutes=get_auto_cancel_minutes())

# --- كود الرفع المعدل والذكي للأقسام والمستخدمين ---
@app.route('/admin/import_users', methods=['GET', 'POST'])
@login_required
@admin_required
def import_users():
    if request.method == 'POST':
        file = request.files.get('file')
        if file and file.filename.lower().endswith('.xlsx'):
            try:
                df = pd.read_excel(file, dtype=str)
                # تنظيف أسماء الأعمدة
                df.columns = df.columns.astype(str).str.strip().str.lower()
                missing = [c for c in ('full_name', 'username', 'password', 'role') if c not in df.columns]
                if missing:
                    flash(f"{tr('import_missing_columns')}: {', '.join(missing)}", 'danger')
                    return render_template('import_users.html')

                def val(row, col):
                    v = row.get(col)
                    return str(v).strip() if pd.notna(v) else ''

                skipped = 0
                for index, row in df.iterrows():
                    username = val(row, 'username')
                    if username:
                        role = val(row, 'role')
                        if role not in AVAILABLE_ROLES:
                            skipped += 1
                            continue
                        
                        # التأكد إن اليوزر مش موجود قبل كده
                        user = User.query.filter_by(username=username).first()
                        
                        dept_id = None
                        dept_name = val(row, 'department')
                        if dept_name:
                            # ابحث عن القسم، ولو مش موجود انشئه فوراً
                            dept_obj = Department.query.filter_by(name=dept_name).first()
                            if not dept_obj:
                                dept_obj = Department(name=dept_name)
                                db.session.add(dept_obj)
                                db.session.flush() # للحصول على الـ ID قبل الـ commit النهائي
                            dept_id = dept_obj.id
                        
                        email = val(row, 'email') or None
                        if not user:
                            password = val(row, 'password')
                            if len(password) < MIN_PASSWORD_LENGTH:
                                skipped += 1
                                continue
                            # إضافة مستخدم جديد
                            new_u = User(
                                full_name=val(row, 'full_name') or username,
                                username=username,
                                email=email,
                                password=generate_password_hash(password),
                                role=role,
                                department_id=dept_id
                            )
                            db.session.add(new_u)
                        else:
                            # تحديث بيانات المستخدم الحالي لو اترفع تاني
                            user.full_name = val(row, 'full_name') or user.full_name
                            user.department_id = dept_id
                            user.role = role
                            if email:
                                user.email = email

                db.session.commit()
                flash_t('import_success', 'success')
                if skipped:
                    flash(f"{tr('invalid_input')} ({skipped})", 'warning')
                return redirect(url_for('manage_users'))
            except Exception as e:
                db.session.rollback()
                flash(f"{tr('import_fail')}: {str(e)}", 'danger')
        else:
            flash_t('file_type_not_allowed', 'danger')
    return render_template('import_users.html')

@app.route('/admin/download_template')
@login_required
@admin_required
def download_template():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("Users Template")
    headers = ['full_name', 'username', 'password', 'role', 'department', 'email']
    for col_num, header in enumerate(headers): worksheet.write(0, col_num, header)
    worksheet.data_validation(1, 3, 500, 3, {'validate': 'list', 'source': AVAILABLE_ROLES, 'input_title': 'اختر الوظيفة', 'input_message': 'اختر من القائمة', 'error_title': 'خطأ', 'error_message': 'الوظيفة غير معتمدة'})
    departments = [d.name for d in Department.query.order_by(Department.name).all()]
    if departments: worksheet.data_validation(1, 4, 500, 4, {'validate': 'list', 'source': departments, 'input_title': 'اختر القسم', 'input_message': 'اختر القسم الخاص بالموظف', 'error_title': 'خطأ', 'error_message': 'القسم غير مسجل في النظام'})
    workbook.close(); output.seek(0)
    response = make_response(output.read())
    response.headers["Content-Disposition"] = "attachment; filename=users_template.xlsx"
    response.headers["Content-type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return response

# --- إدارة المهام ---
@app.route('/tasks')
@login_required
def tasks():
    if not user_can_see('tasks_list'):
        flash_t('access_denied', 'danger')
        return redirect(url_for('dashboard'))
    f = request.args.get('status')
    if f == 'In Progress':
        query = Task.query.filter(Task.status.in_(ACTIVE_STATUSES))
    elif f == 'Canceled':
        query = Task.query.filter(Task.status.in_(LEGACY_CANCELED))
    elif f == 'Overdue':
        query = overdue_query()
    elif f:
        query = Task.query.filter_by(status=f)
    else:
        query = Task.query
    if request.args.get('mine'):
        query = query.filter(db.or_(Task.assignee_id == current_user.id, Task.creator_id == current_user.id, Task.head_id == current_user.id))
    q = (request.args.get('q') or '').strip()
    if q:
        like = f'%{q}%'
        query = query.filter(db.or_(Task.title.ilike(like), Task.description.ilike(like)))
    return render_template('task_list.html', tasks=query.order_by(Task.created_at.desc()).all())

@app.route('/tasks/create', methods=['GET', 'POST'])
@login_required
def create_task():
    if not user_can_see('create_task'):
        flash_t('access_denied', 'danger')
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        assignee_id = request.form.get('assigned_to', type=int)
        assignee = db.session.get(User, assignee_id) if assignee_id else None
        title = (request.form.get('title') or '').strip()
        if not assignee or not assignee.is_active or not title:
            flash_t('invalid_input', 'danger')
            return redirect(url_for('create_task'))
        priority = request.form.get('priority')
        if priority not in ('Low', 'Medium', 'High', 'Urgent'):
            priority = 'Medium'
        recurrence = request.form.get('recurrence', 'None')
        if recurrence not in ('None', 'Daily', 'Weekly', 'Monthly'):
            recurrence = 'None'
        deadline = None
        dl = (request.form.get('deadline') or '').strip()
        if dl:
            for fmt in ('%Y-%m-%dT%H:%M', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
                try:
                    deadline = datetime.strptime(dl, fmt)
                    break
                except ValueError:
                    continue
            if deadline is None:
                flash_t('invalid_deadline', 'danger')
                return redirect(url_for('create_task'))
        fname = save_upload(request.files.get('attachment'))
        if fname is False:
            flash_t('file_type_not_allowed', 'danger')
            return redirect(url_for('create_task'))

        rule = HierarchyRule.query.filter_by(from_role=current_user.role, to_role=assignee.role).first()
        cc_user = None
        if rule and rule.cc_role:
            # Prefer a CC in the assignee's own department before falling back to anyone with that role.
            cc_query = User.query.filter_by(role=rule.cc_role, is_active=True)
            if assignee.department_id:
                cc_user = cc_query.filter_by(department_id=assignee.department_id).first()
            cc_user = cc_user or cc_query.first()
        if not cc_user and assignee.department and assignee.department.head:
            # No hierarchy rule matched: keep the assignee's department head in the loop.
            head = assignee.department.head
            if head.is_active and head.id not in (assignee.id, current_user.id):
                cc_user = head
        now = datetime.now()
        new_task = Task(title=title, description=request.form.get('description'), priority=priority,
                        recurrence=recurrence, next_run=next_recurrence(now, recurrence),
                        creator_id=current_user.id, assignee_id=assignee.id, head_id=cc_user.id if cc_user else None,
                        status='In Progress', created_at=now, deadline=deadline, attachment=fname or None)
        db.session.add(new_task); db.session.commit()
        create_notification(
            assignee.id, 'new_task',
            tr('notif_new_task'),
            f"{current_user.full_name}: {new_task.title}",
            link=f'/tasks/{new_task.id}', task_id=new_task.id
        )
        if cc_user and cc_user.id != assignee.id:
            create_notification(
                cc_user.id, 'new_task',
                tr('notif_new_task_cc'),
                f"{current_user.full_name}: {new_task.title}",
                link=f'/tasks/{new_task.id}', task_id=new_task.id
            )
        db.session.commit()
        send_email(
            f"مهمة جديدة: {new_task.title}", [assignee.email],
            f"مهمة جديدة.\nالعنوان: {new_task.title}\nبواسطة: {current_user.full_name}\n{request.host_url.rstrip('/')}/tasks/{new_task.id}",
            cc=[cc_user.email] if cc_user else None,
        )
        flash_t('task_sent', 'success')
        return redirect(url_for('tasks'))
    return render_template('create_task.html', users=User.query.filter_by(is_active=True).order_by(User.full_name).all())

@app.route('/tasks/<int:id>', methods=['GET', 'POST'])
@login_required
def task_detail(id):
    task = Task.query.get_or_404(id)
    chat_updates = TaskUpdate.query.filter_by(task_id=task.id).order_by(TaskUpdate.created_at.asc()).all()

    if request.method == 'GET':
        # Viewing a task clears its unread notifications for this user.
        Notification.query.filter_by(user_id=current_user.id, task_id=task.id, is_read=False).update({'is_read': True})
        db.session.commit()
        if current_user.id == task.assignee_id and not task.opened_at:
            task.opened_at = datetime.now()
            if not task.started_at:
                task.started_at = datetime.now()
            if task.status in ['New', 'Under Review']:
                task.status = 'In Progress'
            db.session.add(TaskUpdate(content="[SYS:task_opened]", is_note=True, task_id=task.id, user_id=current_user.id))
            db.session.commit()
            if task.creator_id and task.creator_id != current_user.id:
                create_notification(
                    task.creator_id, 'task_update',
                    tr('notif_task_opened'),
                    f"{current_user.full_name}: {task.title}",
                    link=f'/tasks/{task.id}', task_id=task.id
                )
                db.session.commit()
            chat_updates = TaskUpdate.query.filter_by(task_id=task.id).order_by(TaskUpdate.created_at.asc()).all()
        return render_template('task_detail.html', task=task, chat_updates=chat_updates)

    if task.is_terminal:
        flash_t('task_closed_msg', 'danger')
        return redirect(url_for('task_detail', id=task.id))

    participants = {task.creator_id, task.assignee_id, task.head_id}
    if current_user.id not in participants and current_user.role not in ['admin', 'CEO']:
        flash_t('access_denied', 'danger')
        return redirect(url_for('task_detail', id=task.id))

    ns = request.form.get('status') or task.status
    cnt = (request.form.get('content') or '').strip()

    if ns == 'Completed' and not cnt and TaskUpdate.query.filter_by(task_id=task.id, is_note=False).count() == 0:
        flash_t('update_required', 'danger')
        return redirect(url_for('task_detail', id=task.id))

    if ns == 'Canceled' and current_user.id not in [task.creator_id] and current_user.role not in ['admin', 'CEO']:
        flash_t('access_denied', 'danger')
        return redirect(url_for('task_detail', id=task.id))

    if ns not in ['In Progress', 'Completed', 'Canceled']:
        ns = 'In Progress'

    status_changed = normalize_status(ns) != normalize_status(task.status)
    if status_changed:
        if ns == 'In Progress' and not task.started_at:
            task.started_at = datetime.now()
        if ns in ['Completed', 'Canceled'] and not task.completed_at:
            task.completed_at = datetime.now()
        task.status = ns

    if not cnt and not status_changed:
        return redirect(url_for('task_detail', id=task.id))

    if cnt:
        fname = save_upload(request.files.get('file'), prefix='update_')
        if fname is False:
            db.session.rollback()
            flash_t('file_type_not_allowed', 'danger')
            return redirect(url_for('task_detail', id=task.id))
        is_note = bool(request.form.get('is_note')) and current_user.role in ['admin', 'CEO', 'Section Head', 'Manager']
        db.session.add(TaskUpdate(content=cnt, is_note=is_note, filename=fname or None, task_id=task.id, user_id=current_user.id))

    if status_changed and not cnt:
        db.session.add(TaskUpdate(content=f"[SYS:status_change:{ns}]", is_note=True, task_id=task.id, user_id=current_user.id))

    db.session.commit()

    if cnt or status_changed:
        notify_ids = set()
        if task.creator_id and task.creator_id != current_user.id:
            notify_ids.add(task.creator_id)
        if task.head_id and task.head_id != current_user.id:
            notify_ids.add(task.head_id)
        if task.assignee_id and task.assignee_id != current_user.id:
            notify_ids.add(task.assignee_id)
        msg_text = f"{current_user.full_name}: {task.title}"
        if status_changed:
            msg_text += f" → {status_label(ns)}"
        if cnt:
            msg_text += f" — {cnt[:80]}{'...' if len(cnt) > 80 else ''}"
        for uid in notify_ids:
            create_notification(
                uid, 'task_update', tr('notif_task_update'),
                msg_text, link=f'/tasks/{task.id}', task_id=task.id
            )
        db.session.commit()

    flash_t('update_success', 'success')
    return redirect(url_for('task_detail', id=task.id))

def run_migrations():
    for stmt in (
        'ALTER TABLE user ADD COLUMN department_id INTEGER REFERENCES department(id)',
        'ALTER TABLE task ADD COLUMN opened_at DATETIME',
    ):
        try:
            db.session.execute(text(stmt)); db.session.commit()
        except Exception:
            db.session.rollback()  # column already exists

_db_initialized = False

def init_db():
    """Create tables, run light migrations and seed defaults. Safe to call repeatedly."""
    global _db_initialized
    if _db_initialized:
        return
    with app.app_context():
        db.create_all()
        run_migrations()
        seed_feature_visibility()
        seed_app_settings()
        Task.query.filter(Task.status == 'New').update({'status': 'In Progress'})
        db.session.commit()

        admin_user = User.query.filter_by(username='admin').first()
        if not admin_user:
            # غيّر كلمة المرور الافتراضية فور أول تسجيل دخول
            db.session.add(User(full_name='المدير العام', username='admin', password=generate_password_hash(os.environ.get('ADMIN_PASSWORD', 'admin')), role='admin'))
            log.warning("Created default 'admin' user — change its password after first login.")
        db.session.commit()
        if not mail_enabled():
            log.info('MAIL_USERNAME/MAIL_PASSWORD not set — email notifications are disabled.')
    _db_initialized = True

# Runs on import too, so `gunicorn app:app` gets a ready database.
init_db()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('APP_PORT', APP_PORT)), debug=False, use_reloader=False)
