import io
import os
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session, make_response
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from flask_mail import Mail, Message
from apscheduler.schedulers.background import BackgroundScheduler
import xlsxwriter
import pandas as pd
from sqlalchemy import text

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

AVAILABLE_ROLES = ['CEO', 'Section Head', 'Manager', 'Second line manager', 'Employee', 'admin']

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
    status = db.Column(db.String(50), default='New')
    priority = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=datetime.now)
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
    updates = db.relationship('TaskUpdate', backref='task', lazy=True, cascade="all, delete-orphan")
    
    @property
    def is_overdue(self):
        if self.deadline and self.status not in ['Completed', 'Closed', 'Closed_by_System', 'Overdue_Closed']:
            return datetime.now() > self.deadline
        return False

class TaskUpdate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    content = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.now)
    is_note = db.Column(db.Boolean, default=False)
    filename = db.Column(db.String(200), nullable=True)
    task_id = db.Column(db.Integer, db.ForeignKey('task.id'))
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    user = db.relationship('User')

@login_manager.user_loader
def load_user(user_id): return User.query.get(int(user_id))

# --- الترجمة ---
@app.context_processor
def inject_translations():
    lang = session.get('lang', 'ar')
    if lang == 'en':
        t = {
            'dir': 'ltr', 'font': "'Arial', sans-serif", 'system_title': 'Lotus CRM', 'dashboard': 'Dashboard', 'new_task': 'Assign Task', 'tasks_list': 'Tasks List', 'users_manage': 'Manage Users', 'import_users': 'Import Users', 'hierarchy_manage': 'Hierarchy Rules', 'departments_manage': 'Manage Departments', 'department': 'Department', 'no_dept': '-- No Department --', 'alert_settings': 'Alert Settings', 'lang_name': 'عربي', 'logout': 'Logout', 'login_title': 'Lotus CRM', 'login_subtitle': 'Login to continue', 'username': 'Username', 'password': 'Password', 'login_btn': 'Login', 'copyright': '© 2026 Lotus CRM.', 'full_name': 'Full Name', 'email': 'Email', 'role': 'Role', 'actions': 'Actions', 'add_user': 'Add User', 'save': 'Save', 'cancel': 'Cancel', 'view': 'View', 'total_tasks': 'Total Tasks', 'in_progress': 'In Progress', 'overdue': 'Overdue', 'completed': 'Completed', 'task_title': 'Task Title', 'creator': 'Created By', 'assignee': 'Assignee', 'head_cc': 'Monitored By (CC)', 'deadline': 'Deadline', 'status': 'Status', 'priority': 'Priority', 'action': 'Action', 'description': 'Description', 'create_btn': 'Send Task', 'attachment': 'Attachment', 'status_New': 'New', 'status_In Progress': 'In Progress', 'status_Completed': 'Completed', 'status_Under Review': 'Under Review', 'status_Closed': 'Closed', 'status_Overdue_Closed': 'Closed (Overdue)', 'status_Closed_by_System': 'Closed (Ignored >10m)', 'select_emp': '-- Select Employee --', 'select_head': '-- Select CC / Head --', 'no_deadline': 'No Deadline', 'recurrence': 'Recurrence', 'rec_none': 'Once', 'rec_daily': 'Daily', 'rec_weekly': 'Weekly', 'rec_monthly': 'Monthly', 'priority_low': 'Low', 'priority_medium': 'Medium', 'priority_high': 'High', 'priority_urgent': 'Urgent', 'task_details_title': 'Task Details', 'history_updates': 'History & Updates', 'add_update': 'Add Update', 'admin_note': 'Head Note', 'active': 'Active', 'suspended': 'Suspended', 'status_col': 'Status', 'edit': 'Edit', 'block_unblock': 'Block/Unblock', 'reset_pass': 'Reset Password', 'import_excel': 'Import Excel', 'import_excel_title': 'Import from Excel', 'import_instructions': 'Instructions:', 'download_template': 'Download Template', 'import_desc': 'Ensure Excel format (.xlsx):', 'import_name_desc': 'Full Name', 'import_username_desc': 'Username', 'import_pass_desc': 'Password', 'import_role_desc': 'Role', 'import_dept_desc': 'Department', 'import_choose_file': 'Choose Excel file', 'start_import': 'Start Import', 'back_to_users': 'Back to Users', 'hierarchy_title': 'Hierarchy Matrix', 'add_rule': 'Add Rule', 'from_role': 'From Role', 'to_role': 'To Role', 'cc_role': 'CC Role', 'save_rule': 'Save', 'sender': 'Sender', 'receiver': 'Receiver', 'appears_to': 'Appears to', 'delete': 'Delete', 'confirm_delete': 'Are you sure?', 'no_cc': '-- No CC --', 'reports': 'Reports & Exports', 'add_dept': 'Add Department', 'dept_name': 'Department Name', 'dept_head': 'Department Head', 'emp_count': 'Employees Count', 'no_depts': 'No departments found',
            'knowledge_title': 'Company Library & Regulations', 'add_article': 'Add Regulation/Article', 'search_placeholder': 'Search regulations, policies, manuals...', 'search_btn': 'Search', 'read_details': 'Read Details', 'added_by': 'Added by:', 'no_documents': 'No documents found', 'add_new_doc': 'Add New Document', 'doc_title': 'Document Title', 'category': 'Category', 'cat_hr': 'HR Regulations', 'cat_it': 'IT Manuals', 'cat_general': 'General Policies', 'cat_forms': 'Business Forms', 'content': 'Content', 'save_publish': 'Save & Publish',
            'reports_title': 'Reports & Data Export', 'emp_report': 'Employees Report', 'emp_report_desc': 'Download all employees data, roles, and status.', 'download_excel': 'Download Excel', 'tasks_report': 'Comprehensive Tasks Report', 'tasks_report_desc': 'Download full task history (Completed, In-progress).', 'overdue_report': 'Overdue Tasks Report', 'overdue_report_desc': 'Report for tasks exceeding the deadline.', 'back_dashboard': 'Back to Dashboard',
            'admin_panel_title': 'Admin Dashboard', 'add_new_user': 'Add New User', 'can_reports': 'Reports Access', 'can_excel': 'Excel Access', 'add_btn': 'Add', 'current_users': 'Current Users',
            'add_new_cat': 'Add New Category', 'cat_name_placeholder': 'Category Name (e.g. Warehouse)', 'current_cats': 'Current Categories',
            'branch_stats_title': 'Complaints by Branch', 'complaints_count': 'Complaints Count', 'back_home': 'Back to Home',
            'current_status': '(Current)', 'attachment_label': 'Attachment:',
            'change_password': 'Change Password', 'current_password': 'Current Password', 'new_password': 'New Password'
        }
    else:
        t = {
            'dir': 'rtl', 'font': "'Cairo', sans-serif", 'system_title': 'لوتس لإدارة المهام', 'dashboard': 'لوحة المراقبة', 'new_task': 'إسناد مهمة', 'tasks_list': 'قائمة المهام', 'users_manage': 'إدارة الموظفين', 'import_users': 'استيراد الموظفين', 'hierarchy_manage': 'شجرة الصلاحيات', 'departments_manage': 'إدارة الأقسام', 'department': 'القسم', 'no_dept': '-- بدون قسم --', 'alert_settings': 'إعدادات التنبيهات', 'lang_name': 'English', 'logout': 'تسجيل الخروج', 'login_title': 'نظام إدارة المهام', 'login_subtitle': 'قم بتسجيل الدخول للمتابعة', 'username': 'اسم المستخدم', 'password': 'كلمة المرور', 'login_btn': 'دخول', 'copyright': '© 2026 لوتس CRM.', 'full_name': 'الاسم بالكامل', 'email': 'البريد الإلكتروني', 'role': 'الوظيفة', 'actions': 'الإجراءات', 'add_user': 'إضافة موظف', 'save': 'حفظ', 'cancel': 'إلغاء', 'view': 'عرض', 'total_tasks': 'إجمالي المهام', 'in_progress': 'قيد التنفيذ', 'overdue': 'متأخرة', 'completed': 'مكتملة', 'task_title': 'المهمة', 'creator': 'المرسل', 'assignee': 'المسؤول', 'head_cc': 'متابعة بواسطة (CC)', 'deadline': 'موعد التسليم', 'status': 'الحالة', 'priority': 'الأولوية', 'action': 'الإجراء', 'description': 'الوصف', 'create_btn': 'إرسال المهمة', 'attachment': 'مرفقات', 'status_New': 'جديدة (لم تفتح)', 'status_In Progress': 'جاري العمل', 'status_Completed': 'مكتملة', 'status_Under Review': 'قيد المراجعة', 'status_Closed': 'مغلقة', 'status_Overdue_Closed': 'إغلاق تلقائي', 'status_Closed_by_System': 'مغلقة بواسطة النظام', 'select_emp': '-- اختر الموظف --', 'select_head': '-- اختر المتابع --', 'no_deadline': 'بدون موعد', 'recurrence': 'تكرار المهمة', 'rec_none': 'مرة واحدة', 'rec_daily': 'يومياً', 'rec_weekly': 'أسبوعياً', 'rec_monthly': 'شهرياً', 'priority_low': 'منخفضة', 'priority_medium': 'متوسطة', 'priority_high': 'عالية', 'priority_urgent': 'عاجلة جداً', 'task_details_title': 'تفاصيل المهمة', 'history_updates': 'سجل المتابعة', 'add_update': 'إضافة تحديث', 'admin_note': 'ملحوظة إدارية', 'active': 'نشط', 'suspended': 'موقوف', 'status_col': 'الحالة', 'edit': 'تعديل', 'block_unblock': 'إيقاف/تفعيل', 'reset_pass': 'إعادة تعيين المرور', 'import_excel': 'استيراد شيت', 'import_excel_title': 'استيراد الموظفين', 'import_instructions': 'تعليمات التجهيز:', 'download_template': 'تحميل النموذج', 'import_desc': 'يجب أن يكون الملف Excel (.xlsx):', 'import_name_desc': 'الاسم بالكامل', 'import_username_desc': 'اسم المستخدم', 'import_pass_desc': 'كلمة المرور', 'import_role_desc': 'الوظيفة', 'import_dept_desc': 'القسم', 'import_choose_file': 'اختر ملف إكسيل', 'start_import': 'بدء الاستيراد', 'back_to_users': 'عودة', 'hierarchy_title': 'شجرة مسارات الصلاحيات', 'add_rule': 'إضافة قاعدة', 'from_role': 'المرسل', 'to_role': 'المستلم', 'cc_role': 'نسخة CC', 'save_rule': 'حفظ', 'sender': 'المرسل', 'receiver': 'المستلم', 'appears_to': 'يظهر لـ (CC)', 'delete': 'حذف', 'confirm_delete': 'تأكيد الحذف؟', 'no_cc': '-- بدون CC --', 'reports': 'التقارير الشاملة', 'add_dept': 'إضافة قسم', 'dept_name': 'اسم القسم', 'dept_head': 'رئيس القسم', 'emp_count': 'عدد الموظفين', 'no_depts': 'لا توجد أقسام مسجلة',
            'knowledge_title': 'مكتبة الشركة واللوائح الداخلية', 'add_article': 'إضافة لائحة/مقال', 'search_placeholder': 'ابحث في اللوائح، سياسات الإجازات، أدلة التشغيل...', 'search_btn': 'بحث', 'read_details': 'قراءة التفاصيل', 'added_by': 'أُضيف بواسطة:', 'no_documents': 'لا توجد مستندات', 'add_new_doc': 'إضافة مستند جديد', 'doc_title': 'عنوان المستند', 'category': 'التصنيف', 'cat_hr': 'لوائح HR', 'cat_it': 'أدلة تقنية (IT)', 'cat_general': 'سياسات عامة', 'cat_forms': 'نماذج عمل', 'content': 'المحتوى', 'save_publish': 'حفظ ونشر',
            'reports_title': 'التقارير وتصدير البيانات', 'emp_report': 'تقرير الموظفين', 'emp_report_desc': 'تحميل بيانات جميع الموظفين وحالة حساباتهم والأدوار الوظيفية.', 'download_excel': 'تحميل Excel', 'tasks_report': 'التقرير الشامل للمهام', 'tasks_report_desc': 'تحميل سجل المهام بالكامل (المكتملة، قيد التنفيذ، والملغاة).', 'overdue_report': 'المهام المتأخرة (Overdue)', 'overdue_report_desc': 'تقرير خاص بالمهام التي تخطت وقت التسليم المحدد (Deadline).', 'back_dashboard': 'العودة للوحة المراقبة',
            'admin_panel_title': 'لوحة تحكم المشرف', 'add_new_user': 'إضافة مستخدم جديد', 'can_reports': 'صلاحية التقارير', 'can_excel': 'صلاحية Excel', 'add_btn': 'إضافة', 'current_users': 'المستخدمين الحاليين',
            'add_new_cat': 'إضافة تصنيف جديد', 'cat_name_placeholder': 'اسم التصنيف (مثال: شكوى مخازن)', 'current_cats': 'التصنيفات الحالية',
            'branch_stats_title': 'إحصائيات الشكاوى حسب الفرع', 'complaints_count': 'عدد الشكاوى', 'back_home': 'عودة للرئيسية',
            'current_status': '(الحالية)', 'attachment_label': 'المرفق:',
            'change_password': 'تغيير كلمة المرور', 'current_password': 'كلمة المرور الحالية', 'new_password': 'كلمة المرور الجديدة'
        }
    return dict(lang=lang, t=t)

@app.route('/set_language/<lang_code>')
def set_language(lang_code):
    session['lang'] = lang_code
    return redirect(request.referrer or url_for('dashboard'))

# --- الجدولة التلقائية الدقيقة (كل دقيقة) ---
def automated_system_checks():
    with app.app_context():
        now = datetime.now()
        
        # 1. فحص الـ 10 دقائق (المهام الجديدة التي لم تفتح)
        ten_mins_ago = now - timedelta(minutes=10)
        neglected_tasks = Task.query.filter(Task.status == 'New', Task.created_at <= ten_mins_ago).all()
        for t in neglected_tasks:
            t.status = 'Closed_by_System'
            db.session.add(TaskUpdate(content="[إجراء تلقائي] تم الإغلاق بواسطة النظام: الموظف لم يفتح المهمة خلال 10 دقائق من إرسالها.", is_note=True, task_id=t.id, user_id=t.creator_id))
            try:
                recipients = [t.creator.email] if t.creator.email else []
                cc_list = [t.head.email] if (t.head and t.head.email) else []
                if recipients:
                    msg = Message(f"عاجل: إغلاق نظام لمهمة مهملة #{t.id}", sender=app.config['MAIL_USERNAME'], recipients=recipients, cc=cc_list)
                    msg.body = f"تم إغلاق المهمة '{t.title}' تلقائياً لأن الموظف ({t.assignee.full_name}) لم يقم بفتحها أو العمل عليها خلال 10 دقائق."
                    mail.send(msg)
            except: pass

        # 2. فحص الديدلاين (Overdue)
        overdue_tasks = Task.query.filter(Task.deadline < now, Task.status.notin_(['Completed', 'Closed', 'Closed_by_System', 'Overdue_Closed'])).all()
        for t in overdue_tasks:
            t.status = 'Overdue_Closed'
            t.completed_at = now
            db.session.add(TaskUpdate(content="[إجراء تلقائي] تم إغلاق المهمة إجبارياً لتجاوز الموعد المحدد (Deadline).", is_note=True, task_id=t.id, user_id=t.creator_id))
        
        # 3. المهام المتكررة
        recurring = Task.query.filter(Task.recurrence != 'None', Task.next_run <= now).all()
        for rt in recurring:
            new_t = Task(title=rt.title, description=rt.description, priority=rt.priority, recurrence=rt.recurrence, creator_id=rt.creator_id, assignee_id=rt.assignee_id, head_id=rt.head_id)
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
                flash('حسابك موقوف.', 'danger')
                return redirect(url_for('login'))
            login_user(user)
            return redirect(url_for('dashboard'))
        flash('خطأ في البيانات', 'danger')
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
        if session.get('lang', 'ar') == 'ar':
            flash('تم تغيير كلمة المرور بنجاح', 'success')
        else:
            flash('Password changed successfully!', 'success')
    else:
        if session.get('lang', 'ar') == 'ar':
            flash('كلمة المرور الحالية غير صحيحة', 'danger')
        else:
            flash('Incorrect current password!', 'danger')
            
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/dashboard')
@login_required
def dashboard():
    stats = {'total_tasks': Task.query.count(), 'in_progress': Task.query.filter_by(status='In Progress').count(), 'overdue': len([t for t in Task.query.all() if t.is_overdue]), 'completed': Task.query.filter_by(status='Completed').count()}
    recent = Task.query.order_by(Task.created_at.desc()).limit(5).all()
    return render_template('dashboard.html', stats=stats, recent_tasks=recent, today_date=datetime.now().strftime('%Y-%m-%d'))

# --- الإشعارات الحية ---
@app.route('/api/task_notifications')
@login_required
def task_notifications():
    latest = Task.query.filter_by(assignee_id=current_user.id).order_by(Task.id.desc()).first()
    return jsonify({'latest_task_id': latest.id if latest else 0, 'task_title': latest.title if latest else ''})

@app.route('/api/manager_notifications')
@login_required
def manager_notifications():
    latest_update = TaskUpdate.query.join(Task).filter(
        (Task.creator_id == current_user.id) | (Task.head_id == current_user.id)
    ).order_by(TaskUpdate.id.desc()).first()
    
    if latest_update and latest_update.user_id != current_user.id:
        return jsonify({
            'update_id': latest_update.id,
            'task_id': latest_update.task_id,
            'task_title': latest_update.task.title,
            'status': latest_update.task.status,
            'by': latest_update.user.full_name
        })
    return jsonify({'update_id': 0})

# --- التقارير والتصدير ---
@app.route('/reports')
@login_required
def reports():
    if current_user.role not in ['admin', 'CEO', 'Manager', 'Section Head']:
        flash('غير مصرح لك بدخول هذه الصفحة', 'danger')
        return redirect(url_for('dashboard'))
    return render_template('reports.html')

@app.route('/reports/export/<type>')
@login_required
def export_excel(type):
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
def manage_departments():
    departments = Department.query.all()
    heads = User.query.filter(User.role.in_(['Manager', 'Section Head', 'CEO', 'admin'])).all()
    return render_template('manage_departments.html', departments=departments, heads=heads)

@app.route('/admin/departments/add', methods=['POST'])
@login_required
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
def delete_department(id):
    dept = Department.query.get_or_404(id)
    for u in dept.users: u.department_id = None
    db.session.delete(dept)
    db.session.commit()
    return redirect(url_for('manage_departments'))

# --- إدارة الموظفين ---
@app.route('/admin/users')
@login_required
def manage_users(): return render_template('manage_users.html', users=User.query.all(), departments=Department.query.all())

@app.route('/admin/add_user', methods=['POST'])
@login_required
def add_user():
    username = request.form.get('username')
    if not User.query.filter_by(username=username).first():
        new_u = User(full_name=request.form.get('full_name'), username=username, email=request.form.get('email'), password=generate_password_hash(request.form.get('password')), role=request.form.get('role'), department_id=request.form.get('department_id') or None)
        db.session.add(new_u); db.session.commit()
    return redirect(url_for('manage_users'))

@app.route('/admin/user/toggle_status/<int:id>', methods=['POST'])
@login_required
def toggle_user_status(id):
    user = User.query.get_or_404(id); user.is_active = not user.is_active; db.session.commit()
    return redirect(url_for('manage_users'))

@app.route('/admin/user/reset_password/<int:id>', methods=['POST'])
@login_required
def reset_password(id):
    user = User.query.get_or_404(id)
    new_pass = request.form.get('new_password')
    if new_pass: user.password = generate_password_hash(new_pass); db.session.commit()
    return redirect(url_for('manage_users'))

@app.route('/admin/user/edit/<int:id>', methods=['POST'])
@login_required
def edit_user(id):
    user = User.query.get_or_404(id)
    user.full_name = request.form.get('full_name'); user.username = request.form.get('username'); user.email = request.form.get('email'); user.role = request.form.get('role'); user.department_id = request.form.get('department_id') or None
    db.session.commit()
    return redirect(url_for('manage_users'))

@app.route('/admin/hierarchy', methods=['GET', 'POST'])
@login_required
def manage_hierarchy():
    if request.method == 'POST':
        db.session.add(HierarchyRule(from_role=request.form.get('from_role'), to_role=request.form.get('to_role'), cc_role=request.form.get('cc_role'))); db.session.commit()
    return render_template('manage_hierarchy.html', rules=HierarchyRule.query.all(), roles=AVAILABLE_ROLES)

# --- كود الرفع المعدل والذكي للأقسام والمستخدمين ---
@app.route('/admin/import_users', methods=['GET', 'POST'])
@login_required
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
                msg = 'تم استيراد الموظفين والأقسام بنجاح!' if session.get('lang', 'ar') == 'ar' else 'Users and departments imported successfully!'
                flash(msg, 'success')
                return redirect(url_for('manage_users'))
            except Exception as e:
                db.session.rollback()
                msg = f'حدث خطأ أثناء الرفع: {str(e)}' if session.get('lang', 'ar') == 'ar' else f'Error during import: {str(e)}'
                flash(msg, 'danger')
    return render_template('import_users.html')

@app.route('/admin/download_template')
@login_required
def download_template():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("Users Template")
    headers = ['full_name', 'username', 'password', 'role', 'department']
    for col_num, header in enumerate(headers): worksheet.write(0, col_num, header)
    worksheet.data_validation(1, 3, 500, 3, {'validate': 'list', 'source': AVAILABLE_ROLES, 'input_title': 'اختر الوظيفة', 'input_message': 'اختر من القائمة', 'error_title': 'خطأ', 'error_message': 'الوظيفة غير معتمدة'})
    departments = [d.name for d in Department.query.all()]
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
    f = request.args.get('status')
    all_t = Task.query.filter_by(status=f).order_by(Task.created_at.desc()).all() if f else Task.query.order_by(Task.created_at.desc()).all()
    return render_template('task_list.html', tasks=all_t)

@app.route('/tasks/create', methods=['GET', 'POST'])
@login_required
def create_task():
    if request.method == 'POST':
        assignee = User.query.get(request.form.get('assigned_to'))
        rule = HierarchyRule.query.filter_by(from_role=current_user.role, to_role=assignee.role).first()
        cc_user = User.query.filter_by(role=rule.cc_role, is_active=True).first() if rule and rule.cc_role else None
        new_task = Task(title=request.form.get('title'), description=request.form.get('description'), priority=request.form.get('priority'), recurrence=request.form.get('recurrence', 'None'), creator_id=current_user.id, assignee_id=assignee.id, head_id=cc_user.id if cc_user else None)
        dl = request.form.get('deadline')
        if dl: new_task.deadline = datetime.strptime(dl, '%Y-%m-%dT%H:%M')
        file = request.files.get('attachment')
        if file and file.filename != '':
            fname = f"{datetime.now().strftime('%Y%m%d%H%M')}_{secure_filename(file.filename)}"
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], fname))
            new_task.attachment = fname
        db.session.add(new_task); db.session.commit()
        try:
            recipients = [assignee.email] if assignee.email else []
            cc_list = [cc_user.email] if (cc_user and cc_user.email) else []
            if recipients:
                msg = Message(f"مهمة جديدة: {new_task.title}", sender=app.config['MAIL_USERNAME'], recipients=recipients, cc=cc_list)
                msg.body = f"مهمة جديدة.\nالعنوان: {new_task.title}\nبواسطة: {current_user.full_name}"
                mail.send(msg)
        except: pass
        flash('تم إرسال المهمة بنجاح!', 'success')
        return redirect(url_for('tasks'))
    return render_template('create_task.html', users=User.query.filter_by(is_active=True).all())

@app.route('/tasks/<int:id>', methods=['GET', 'POST'])
@login_required
def task_detail(id):
    t = Task.query.get_or_404(id)
    if request.method == 'POST':
        ns = request.form.get('status'); cnt = request.form.get('content')
        if ns == 'Completed' and not cnt and TaskUpdate.query.filter_by(task_id=t.id, is_note=False).count() == 0:
            flash('مرفوض! يجب كتابة تحديث قبل الإكمال.', 'danger')
            return redirect(url_for('task_detail', id=t.id))
        
        # تحويل حالة المهمة من "جديدة" إلى "قيد التنفيذ" إذا تم إرسال تحديث
        if t.status == 'New' and cnt and ns == 'New': ns = 'In Progress'
        
        status_changed = False
        if ns and ns != t.status:
            if ns == 'In Progress' and not t.started_at: t.started_at = datetime.now()
            if ns in ['Completed', 'Closed'] and not t.completed_at: t.completed_at = datetime.now()
            t.status = ns
            status_changed = True
            
        if cnt: 
            # معالجة الملف المرفق في التحديث
            file = request.files.get('file')
            fname = None
            if file and file.filename != '':
                fname = f"update_{datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(file.filename)}"
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], fname))
                
            db.session.add(TaskUpdate(content=cnt, is_note=bool(request.form.get('is_note')), filename=fname, task_id=t.id, user_id=current_user.id))
        db.session.commit()
        
        # إرسال إيميل للمديرين عند تغيير الحالة
        if status_changed:
            try:
                recipients = [t.creator.email] if t.creator.email else []
                cc_list = [t.head.email] if (t.head and t.head.email) else []
                if recipients:
                    msg = Message(f"تحديث حالة مهمة: #{t.id}", sender=app.config['MAIL_USERNAME'], recipients=recipients, cc=cc_list)
                    msg.body = f"الموظف ({current_user.full_name}) قام بتحديث حالة المهمة '{t.title}' إلى: {ns}."
                    mail.send(msg)
            except: pass

        flash('تم التحديث بنجاح!', 'success')
        return redirect(url_for('task_detail', id=t.id))
    return render_template('task_detail.html', task=t)

if __name__ == '__main__':
    with app.app_context(): 
        db.create_all()
        try: db.session.execute(text('ALTER TABLE user ADD COLUMN department_id INTEGER REFERENCES department(id)')); db.session.commit()
        except: pass
        
        # تحديث بيانات الأدمن تلقائياً لو موجود بدل مسح قاعدة البيانات
        admin_user = User.query.filter_by(username='admin').first()
        if not admin_user:
            db.session.add(User(full_name='CEO', username='admin', password=generate_password_hash('admin'), role='admin'))
        else:
            admin_user.full_name = 'المدير العام'  # تقدر تغير اسم المدير من هنا بعدين براحتك
        db.session.commit()
        
    app.run(host='0.0.0.0', port=5000, debug=True, use_reloader=False)