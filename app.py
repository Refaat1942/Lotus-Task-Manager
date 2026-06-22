import io
import os
from datetime import datetime, timedelta
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session, make_response, abort
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from flask_mail import Mail, Message
from apscheduler.schedulers.background import BackgroundScheduler
import xlsxwriter
import pandas as pd
from sqlalchemy import text
from i18n import TRANSLATIONS, t as tr, flash_t, status_label, priority_label, localized_user_name, update_content, normalize_status

APP_VERSION = '2.5.0'
APP_NAME = 'Lotus Task Manager'
APP_PORT = 5000

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_secret_key_here'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///taskmanager.db'
app.config['UPLOAD_FOLDER'] = 'static/uploads'

# إعدادات البريد
app.config['MAIL_SERVER'] = 'smtp.office365.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'your-email@yourdomain.com'
app.config['MAIL_PASSWORD'] = 'your-secret-password'
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

def seed_app_settings():
    if not AppSetting.query.filter_by(key='auto_cancel_minutes').first():
        db.session.add(AppSetting(key='auto_cancel_minutes', value='10'))
    db.session.commit()

@login_manager.user_loader
def load_user(user_id): return User.query.get(int(user_id))

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
    )

@app.route('/set_language/<lang_code>')
def set_language(lang_code):
    session['lang'] = lang_code
    return redirect(request.referrer or url_for('dashboard'))

# --- الجدولة التلقائية الدقيقة (كل دقيقة) ---
def automated_system_checks():
    with app.app_context():
        now = datetime.now()
        
        minutes = get_auto_cancel_minutes()
        cutoff = now - timedelta(minutes=minutes)
        neglected_tasks = Task.query.filter(
            Task.opened_at.is_(None),
            Task.status.in_(ACTIVE_STATUSES),
            Task.created_at <= cutoff
        ).all()
        for task in neglected_tasks:
            task.status = 'Canceled'
            task.completed_at = now
            db.session.add(TaskUpdate(content="[SYS:auto_cancel]", is_note=True, task_id=task.id, user_id=task.creator_id))
            if task.creator_id:
                create_notification(
                    task.creator_id, 'task_update',
                    f"{TRANSLATIONS['ar']['notif_auto_closed']} / {TRANSLATIONS['en']['notif_auto_closed']}",
                    TRANSLATIONS['en']['system_note_auto_cancel'].format(minutes=minutes),
                    link=f'/tasks/{task.id}', task_id=task.id
                )
            try:
                recipients = [task.creator.email] if task.creator.email else []
                cc_list = [task.head.email] if (task.head and task.head.email) else []
                if recipients:
                    msg = Message(f"Urgent: Auto-closed task #{task.id}", sender=app.config['MAIL_USERNAME'], recipients=recipients, cc=cc_list)
                    msg.body = TRANSLATIONS['en']['system_note_auto_cancel'].format(minutes=minutes) + f"\nTask: {task.title}\nAssignee: {task.assignee.full_name if task.assignee else '--'}"
                    mail.send(msg)
            except: pass

        # 2. Deadline alerts (notify only — status stays In Progress)
        overdue_tasks = Task.query.filter(
            Task.deadline < now,
            Task.status.in_(ACTIVE_STATUSES + ['In Progress'])
        ).all()
        for task in overdue_tasks:
            if normalize_status(task.status) in ['Completed', 'Canceled']:
                continue
            if TaskUpdate.query.filter_by(task_id=task.id, content='[SYS:deadline_notified]').first():
                continue
            db.session.add(TaskUpdate(content="[SYS:deadline_notified]", is_note=True, task_id=task.id, user_id=task.creator_id))
            if task.creator_id:
                create_notification(
                    task.creator_id, 'deadline_passed',
                    f"{TRANSLATIONS['ar']['deadline_passed']} / {TRANSLATIONS['en']['deadline_passed']}",
                    f"'{task.title}' — {TRANSLATIONS['en']['deadline_passed_msg']}",
                    link=f'/tasks/{task.id}', task_id=task.id
                )
            if task.head_id and task.head_id != task.creator_id:
                create_notification(
                    task.head_id, 'deadline_passed',
                    f"{TRANSLATIONS['ar']['deadline_passed']} / {TRANSLATIONS['en']['deadline_passed']}",
                    f"'{task.title}' — {TRANSLATIONS['en']['deadline_passed_msg']}",
                    link=f'/tasks/{task.id}', task_id=task.id
                )
            try:
                if task.creator and task.creator.email:
                    cc_list = [task.head.email] if (task.head and task.head.email) else []
                    msg = Message(f"Alert: Task deadline passed #{task.id}", sender=app.config['MAIL_USERNAME'], recipients=[task.creator.email], cc=cc_list)
                    msg.body = TRANSLATIONS['en']['deadline_passed_msg'] + f"\nTask: {task.title}"
                    mail.send(msg)
            except: pass
        
        # 3. المهام المتكررة
        recurring = Task.query.filter(Task.recurrence != 'None', Task.next_run <= now).all()
        for rt in recurring:
            new_t = Task(title=rt.title, description=rt.description, priority=rt.priority, recurrence=rt.recurrence, creator_id=rt.creator_id, assignee_id=rt.assignee_id, head_id=rt.head_id, status='In Progress')
            if rt.recurrence == 'Daily': rt.next_run = now + timedelta(days=1)
            elif rt.recurrence == 'Weekly': rt.next_run = now + timedelta(days=7)
            elif rt.recurrence == 'Monthly': rt.next_run = now + timedelta(days=30)
            new_t.deadline = rt.next_run
            db.session.add(new_t)
            
        db.session.commit()

scheduler = BackgroundScheduler()
scheduler.add_job(func=automated_system_checks, trigger="interval", minutes=1)
scheduler.start()

# --- مسارات النظام الأساسية ---
@app.route('/')
def index(): return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated: return redirect(url_for('dashboard'))
    if request.method == 'POST':
        user = User.query.filter_by(username=request.form.get('username')).first()
        if user and check_password_hash(user.password, request.form.get('password')):
            if not user.is_active:
                flash_t('account_suspended', 'danger')
                return redirect(url_for('login'))
            login_user(user)
            return redirect(url_for('dashboard'))
        flash_t('invalid_login', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout(): 
    logout_user()
    return redirect(url_for('login'))

@app.route('/change_my_password', methods=['POST'])
@login_required
def change_my_password():
    current_password = request.form.get('current_password')
    new_password = request.form.get('new_password')
    
    if check_password_hash(current_user.password, current_password):
        current_user.password = generate_password_hash(new_password)
        db.session.commit()
        flash_t('password_changed', 'success')
    else:
        flash_t('wrong_password', 'danger')
            
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/dashboard')
@login_required
def dashboard():
    if not user_can_see('dashboard'):
        return redirect(url_for('tasks'))
    stats = {
        'total_tasks': Task.query.count(),
        'in_progress': Task.query.filter(Task.status.in_(ACTIVE_STATUSES)).count(),
        'overdue': len([task for task in Task.query.all() if task.is_overdue]),
        'completed': Task.query.filter(Task.status.in_(['Completed'])).count(),
    }
    recent = Task.query.order_by(Task.created_at.desc()).limit(5).all()
    return render_template('dashboard.html', stats=stats, recent_tasks=recent, today_date=datetime.now().strftime('%Y-%m-%d'))

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
    output = io.StringIO()
    import csv
    writer = csv.writer(output)
    if type == 'users':
        writer.writerow(['ID', 'Full Name', 'Email', 'Role', 'Department'])
        for u in User.query.all(): 
            dept_name = u.department.name if u.department else ''
            writer.writerow([u.id, u.full_name, u.email, u.role, dept_name])
    else:
        writer.writerow(['ID', 'Title', 'Assignee', 'Status', 'Started', 'Completed', 'Duration (Hrs)', 'Deadline'])
        data = [t for t in Task.query.all() if t.is_overdue] if type == 'tasks_overdue' else Task.query.all()
        for t in data:
            duration = round((t.completed_at - t.started_at).total_seconds() / 3600, 2) if (t.started_at and t.completed_at) else "N/A"
            writer.writerow([t.id, t.title, t.assignee.full_name, t.status, t.started_at, t.completed_at, duration, t.deadline])
    
    response = make_response(output.getvalue().encode('utf-8-sig'))
    response.headers["Content-Disposition"] = f"attachment; filename={type}.csv"
    response.headers["Content-type"] = "text/csv; charset=utf-8-sig"
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
    name = request.form.get('name')
    head_id = request.form.get('head_id')
    if name and not Department.query.filter_by(name=name).first():
        dept = Department(name=name, head_id=head_id if head_id else None)
        db.session.add(dept)
        db.session.commit()
    return redirect(url_for('manage_departments'))

@app.route('/admin/departments/delete/<int:id>')
@login_required
@admin_required
def delete_department(id):
    dept = Department.query.get_or_404(id)
    for u in dept.users: u.department_id = None
    db.session.delete(dept)
    db.session.commit()
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
    username = request.form.get('username')
    if not User.query.filter_by(username=username).first():
        new_u = User(full_name=request.form.get('full_name'), username=username, email=request.form.get('email'), password=generate_password_hash(request.form.get('password')), role=request.form.get('role'), department_id=request.form.get('department_id') or None)
        db.session.add(new_u); db.session.commit()
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
    new_pass = request.form.get('new_password')
    if new_pass: user.password = generate_password_hash(new_pass); db.session.commit()
    return redirect(url_for('manage_users'))

@app.route('/admin/user/edit/<int:id>', methods=['POST'])
@login_required
@admin_required
def edit_user(id):
    user = User.query.get_or_404(id)
    new_username = request.form.get('username', '').strip()
    existing = User.query.filter(User.username == new_username, User.id != id).first()
    if existing:
        flash_t('user_update_error', 'danger')
        return redirect(url_for('manage_users'))
    user.full_name = request.form.get('full_name')
    user.username = new_username
    user.email = request.form.get('email')
    if id != current_user.id:
        user.role = request.form.get('role')
    user.department_id = request.form.get('department_id') or None
    db.session.commit()
    flash_t('user_updated', 'success')
    return redirect(url_for('manage_users'))

@app.route('/admin/hierarchy', methods=['GET', 'POST'])
@login_required
@admin_required
def manage_hierarchy():
    if request.method == 'POST':
        db.session.add(HierarchyRule(from_role=request.form.get('from_role'), to_role=request.form.get('to_role'), cc_role=request.form.get('cc_role'))); db.session.commit()
    return render_template('manage_hierarchy.html', rules=HierarchyRule.query.order_by(HierarchyRule.from_role, HierarchyRule.to_role).all(), roles=AVAILABLE_ROLES)

@app.route('/admin/hierarchy/delete/<int:id>')
@login_required
@admin_required
def delete_hierarchy_rule(id):
    rule = HierarchyRule.query.get_or_404(id)
    db.session.delete(rule)
    db.session.commit()
    return redirect(url_for('manage_hierarchy'))

@app.route('/admin/permissions', methods=['GET', 'POST'])
@login_required
@admin_required
def manage_permissions():
    if request.method == 'POST':
        for key in DEFAULT_FEATURE_VISIBILITY:
            roles = request.form.getlist(f'feature_{key}')
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
        if file and file.filename.endswith('.xlsx'):
            try:
                df = pd.read_excel(file)
                # تنظيف أسماء الأعمدة
                df.columns = df.columns.str.strip()
                
                for index, row in df.iterrows():
                    if pd.notna(row['username']):
                        username = str(row['username']).strip()
                        
                        # التأكد إن اليوزر مش موجود قبل كده
                        user = User.query.filter_by(username=username).first()
                        
                        dept_id = None
                        if 'department' in df.columns and pd.notna(row['department']):
                            dept_name = str(row['department']).strip()
                            # ابحث عن القسم، ولو مش موجود انشئه فوراً
                            dept_obj = Department.query.filter_by(name=dept_name).first()
                            if not dept_obj:
                                dept_obj = Department(name=dept_name)
                                db.session.add(dept_obj)
                                db.session.flush() # للحصول على الـ ID قبل الـ commit النهائي
                            dept_id = dept_obj.id
                        
                        if not user:
                            # إضافة مستخدم جديد
                            new_u = User(
                                full_name=str(row['full_name']).strip(),
                                username=username,
                                password=generate_password_hash(str(row['password']).strip()),
                                role=str(row['role']).strip(),
                                department_id=dept_id
                            )
                            db.session.add(new_u)
                        else:
                            # تحديث بيانات المستخدم الحالي لو اترفع تاني
                            user.full_name = str(row['full_name']).strip()
                            user.department_id = dept_id
                            user.role = str(row['role']).strip()

                db.session.commit()
                flash_t('import_success', 'success')
                return redirect(url_for('manage_users'))
            except Exception as e:
                db.session.rollback()
                flash(f"{tr('import_fail')}: {str(e)}", 'danger')
    return render_template('import_users.html')

@app.route('/admin/download_template')
@login_required
@admin_required
def download_template():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("Users Template")
    headers = ['full_name', 'username', 'password', 'role', 'department']
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
        all_t = Task.query.filter(Task.status.in_(ACTIVE_STATUSES)).order_by(Task.created_at.desc()).all()
    elif f == 'Canceled':
        all_t = Task.query.filter(Task.status.in_(LEGACY_CANCELED)).order_by(Task.created_at.desc()).all()
    elif f:
        all_t = Task.query.filter_by(status=f).order_by(Task.created_at.desc()).all()
    else:
        all_t = Task.query.order_by(Task.created_at.desc()).all()
    return render_template('task_list.html', tasks=all_t)

@app.route('/tasks/create', methods=['GET', 'POST'])
@login_required
def create_task():
    if not user_can_see('create_task'):
        flash_t('access_denied', 'danger')
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        assignee = User.query.get(request.form.get('assigned_to'))
        rule = HierarchyRule.query.filter_by(from_role=current_user.role, to_role=assignee.role).first()
        cc_user = User.query.filter_by(role=rule.cc_role, is_active=True).first() if rule and rule.cc_role else None
        new_task = Task(title=request.form.get('title'), description=request.form.get('description'), priority=request.form.get('priority'), recurrence=request.form.get('recurrence', 'None'), creator_id=current_user.id, assignee_id=assignee.id, head_id=cc_user.id if cc_user else None, status='In Progress')
        dl = request.form.get('deadline')
        if dl:
            try:
                new_task.deadline = datetime.strptime(dl, '%Y-%m-%dT%H:%M')
            except ValueError:
                new_task.deadline = datetime.strptime(dl, '%Y-%m-%d %H:%M')
        file = request.files.get('attachment')
        if file and file.filename != '':
            fname = f"{datetime.now().strftime('%Y%m%d%H%M')}_{secure_filename(file.filename)}"
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], fname))
            new_task.attachment = fname
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
        try:
            recipients = [assignee.email] if assignee.email else []
            cc_list = [cc_user.email] if (cc_user and cc_user.email) else []
            if recipients:
                msg = Message(f"مهمة جديدة: {new_task.title}", sender=app.config['MAIL_USERNAME'], recipients=recipients, cc=cc_list)
                msg.body = f"مهمة جديدة.\nالعنوان: {new_task.title}\nبواسطة: {current_user.full_name}"
                mail.send(msg)
        except: pass
        flash_t('task_sent', 'success')
        return redirect(url_for('tasks'))
    return render_template('create_task.html', users=User.query.filter_by(is_active=True).order_by(User.full_name).all())

@app.route('/tasks/<int:id>', methods=['GET', 'POST'])
@login_required
def task_detail(id):
    task = Task.query.get_or_404(id)
    chat_updates = TaskUpdate.query.filter_by(task_id=task.id).order_by(TaskUpdate.created_at.asc()).all()

    if request.method == 'GET':
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
        file = request.files.get('file')
        fname = None
        if file and file.filename != '':
            fname = f"update_{datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(file.filename)}"
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], fname))
        db.session.add(TaskUpdate(content=cnt, is_note=bool(request.form.get('is_note')), filename=fname, task_id=task.id, user_id=current_user.id))

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

if __name__ == '__main__':
    with app.app_context(): 
        db.create_all()
        try: db.session.execute(text('ALTER TABLE user ADD COLUMN department_id INTEGER REFERENCES department(id)')); db.session.commit()
        except: pass
        try: db.session.execute(text('ALTER TABLE task ADD COLUMN opened_at DATETIME')); db.session.commit()
        except: pass
        seed_feature_visibility()
        seed_app_settings()
        Task.query.filter(Task.status == 'New').update({'status': 'In Progress'})
        db.session.commit()
        
        # تحديث بيانات الأدمن تلقائياً لو موجود بدل مسح قاعدة البيانات
        admin_user = User.query.filter_by(username='admin').first()
        if not admin_user:
            db.session.add(User(full_name='CEO', username='admin', password=generate_password_hash('admin'), role='admin'))
        else:
            admin_user.full_name = 'المدير العام'  # تقدر تغير اسم المدير من هنا بعدين براحتك
        db.session.commit()
        
    app.run(host='0.0.0.0', port=APP_PORT, debug=False, use_reloader=False)