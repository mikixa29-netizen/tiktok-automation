import os
import threading
import time
import random
import io
import pandas as pd
from docx import Document
from flask import Flask, render_template, render_template_string, request, redirect, url_for, flash, session
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from playwright.sync_api import sync_playwright

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-super-secret-key-change-this'

# Render ላይ ሲሰራ ዳታቤዙን ወደ /tmp ፎልደር እንዲያደርገው (Read-only ችግር እንዳይፈጠር)
if 'RENDER' in os.environ:
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:////tmp/tiktok_enterprise.db'
else:
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///tiktok_enterprise.db'

app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
# SQLite በባክግራውንድ ትሬድ (Background Thread) በሰላም ዳታ መጻፍ እንዲችል
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {"connect_args": {"check_same_thread": False}}

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# --- ቋንቋዎች (Translations Dictionary) ---
TRANSLATIONS = {
    'en': {
        'title': 'Enterprise TikTok Operations Portal',
        'live_status': 'System Operational',
        'user': 'User',
        'logout': 'Logout',
        'instructions_title': 'Enterprise Operational Guidelines',
        'inst_1': 'Upload verified account credentials via Excel (.xlsx) with <code>username</code> and <code>password</code> columns.',
        'inst_2': 'Configure anti-detection delays (min/max seconds) to mimic human interaction patterns safely.',
        'inst_3': 'Distribute target URLs and comment pools before initiating autonomous bot execution.',
        'settings': 'Automation & Anti-Detection Parameters',
        'likes': 'Likes Count',
        'follows': 'Follows Count',
        'comments': 'Comments Count',
        'min_delay': 'Min Delay (s)',
        'max_delay': 'Max Delay (s)',
        'save_settings': 'Save Settings',
        'accounts': 'Account Pool',
        'upload': 'Upload',
        'comment_pool': 'Comment Pool',
        'distribute': 'Distribute',
        'target_links': 'Target Links',
        'add_link': 'Add Link',
        'start_bot': 'Launch Stealth Automation',
        'admin_mgmt': 'Admin User & Role Management',
        'add_user': 'Add User',
        'reports': 'Execution Audit & Reports',
        'time': 'Timestamp',
        'account': 'Account Used',
        'link': 'Target Link',
        'action': 'Action Performed',
        'status': 'Status'
    },
    'am': {
        'title': 'የድርጅት ቲክቶክ ኦፕሬሽን ፖርታል',
        'live_status': 'ሥርዓቱ በሥራ ላይ ይገኛል',
        'user': 'ተጠቃሚ',
        'logout': 'ውጣ',
        'instructions_title': 'የሥራ መመሪያዎች እና ትዕዛዞች',
        'inst_1': 'የአካውንት መረጃዎችን በ Excel (.xlsx) <code>username</code> እና <code>password</code> በሚሉ ኬላዎች ጫን።',
        'inst_2': 'ሰው ሰራሽ እንዳልሆነ ለማስመሰል የቆይታ ጊዜ (Min/Max Delay) ማስተካከያዎችን አስተካክል።',
        'inst_3': 'ቦቱን ከማስጀመርዎ በፊት የ লক্ষ্য ሊንኮችን እና የኮሜንት ማከማቻዎችን ያስገቡ።',
        'settings': 'የአውቶሜሽን እና የጥንቃቄ ማስተካከያዎች',
        'likes': 'የላይክ ብዛት',
        'follows': 'የፎሎው ብዛት',
        'comments': 'የኮሜንት ብዛት',
        'min_delay': 'ዝቅተኛ መዘግየት (ሰከንድ)',
        'max_delay': 'ከፍተኛ መዘግየት (ሰከንድ)',
        'save_settings': 'ቅንብሮችን ቆጥብ',
        'accounts': 'የአካውንቶች ማከማቻ',
        'upload': 'ጫን',
        'comment_pool': 'የኮሜንት ማከማቻ',
        'distribute': 'አሰራጭ',
        'target_links': 'የታለመዱ ሊንኮች',
        'add_link': 'ሊንክ ጨምር',
        'start_bot': 'ቦቱን በድብቅ (Stealth) አስጀምር',
        'admin_mgmt': 'የአስተዳዳሪ እና የዩዘር አስተዳደር',
        'add_user': 'ዩዘር ጨምር',
        'reports': 'የክንውን ሪፖርቶች እና ታሪክ',
        'time': 'ሰዓት',
        'account': 'የተጠቀመበት አካውንት',
        'link': 'የታለመደ ሊንክ',
        'action': 'የተከናወነ ተግባር',
        'status': 'ሁኔታ'
    }
}

# --- ዳታቤዝ ሞዴሎች (Database Models) ---
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password = db.Column(db.String(150), nullable=False)
    role = db.Column(db.String(50), nullable=False, default='operator')

class Account(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), nullable=False)
    password = db.Column(db.String(150), nullable=False)

class Comment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    text = db.Column(db.Text, nullable=False)

class Setting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    likes_count = db.Column(db.Integer, default=1)
    follows_count = db.Column(db.Integer, default=1)
    comments_count = db.Column(db.Integer, default=1)
    min_delay = db.Column(db.Integer, default=5)
    max_delay = db.Column(db.Integer, default=12)

class Link(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    url = db.Column(db.String(500), nullable=False)

class Report(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    timestamp = db.Column(db.String(100), nullable=False)
    account_used = db.Column(db.String(150), nullable=False)
    target_link = db.Column(db.String(500), nullable=False)
    action_performed = db.Column(db.String(200), nullable=False)
    status = db.Column(db.String(50), default='Success')

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# --- አፕሊኬሽኑ ሲነሳ ዳታቤዙን እና ነባሪ አድሚን በራስሰር መፍጠር ---
def init_db():
    with app.app_context():
        db.create_all()
        if not User.query.filter_by(username='admin').first():
            hashed_pw = generate_password_hash('admin123')
            db.session.add(User(username='admin', password=hashed_pw, role='admin'))
            db.session.commit()
        if not Setting.query.first():
            db.session.add(Setting())
            db.session.commit()

init_db()

# --- ራውቶች (Routes) ---
@app.route('/')
@login_required
def index():
    lang = session.get('lang', 'am')
    t = TRANSLATIONS.get(lang, TRANSLATIONS['am'])
    
    accounts = Account.query.all()
    comments = Comment.query.all()
    links = Link.query.all()
    settings = Setting.query.first()
    users = User.query.all() if current_user.role == 'admin' else []
    reports = Report.query.order_by(Report.id.desc()).all()
    
    return render_template('dashboard.html', t=t, lang=lang, accounts=accounts, 
                           comments=comments, links=links, settings=settings, 
                           users=users, reports=reports)

@app.route('/set_language/<lang_code>')
def set_language(lang_code):
    if lang_code in ['en', 'am']:
        session['lang'] = lang_code
    return redirect(url_for('index'))

LOGIN_HTML = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Login - Enterprise TikTok Portal</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light d-flex align-items-center justify-content-center vh-100">
    <div class="card shadow p-4" style="width: 380px; border-radius: 10px;">
        <h4 class="text-center mb-4 fw-bold text-dark">TikTok Portal Login</h4>
        {% with messages = get_flashed_messages() %}
          {% if messages %}
            <div class="alert alert-danger py-2 small">{{ messages[0] }}</div>
          {% endif %}
        {% endwith %}
        <form method="POST">
            <div class="mb-3">
                <label class="form-label small fw-semibold text-muted">Username</label>
                <input type="text" name="username" class="form-control" required>
            </div>
            <div class="mb-3">
                <label class="form-label small fw-semibold text-muted">Password</label>
                <input type="password" name="password" class="form-control" required>
            </div>
            <button type="submit" class="btn btn-dark w-100 py-2">Login</button>
        </form>
    </div>
</body>
</html>
'''

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        if user and check_password_hash(user.password, password):
            login_user(user)
            return redirect(url_for('index'))
        flash('የተሳሳተ መግቢያ ስም ወይም የይለፍ ቃል!')
    return render_template_string(LOGIN_HTML)

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/update_settings', methods=['POST'])
@login_required
def update_settings():
    if current_user.role == 'viewer':
        return redirect(url_for('index'))
    
    settings = Setting.query.first()
    if not settings:
        settings = Setting()
        db.session.add(settings)
    
    settings.likes_count = int(request.form.get('likes_count', 1))
    settings.follows_count = int(request.form.get('follows_count', 1))
    settings.comments_count = int(request.form.get('comments_count', 1))
    settings.min_delay = int(request.form.get('min_delay', 5))
    settings.max_delay = int(request.form.get('max_delay', 12))
    db.session.commit()
    
    flash('ቅንብሮችዎ በተሳካ ሁኔታ ተዘምነዋል!')
    return redirect(url_for('index'))

@app.route('/upload_accounts', methods=['GET', 'POST'])
@login_required
def upload_accounts():
    if request.method == 'GET':
        return redirect(url_for('index'))
    if current_user.role == 'viewer':
        return redirect(url_for('index'))
    
    file = request.files.get('account_file')
    if file:
        try:
            file_stream = io.BytesIO(file.read())
            df = pd.read_excel(file_stream, engine='openpyxl')
            df.columns = df.columns.astype(str).str.strip().str.lower()
            if 'username' in df.columns and 'password' in df.columns:
                for _, row in df.iterrows():
                    uname = str(row['username']).strip()
                    pasw = str(row['password']).strip()
                    if uname and uname != 'nan':
                        db.session.add(Account(username=uname, password=pasw))
                db.session.commit()
                flash('አካውንቶች በተሳካ ሁኔታ ተጭነዋል!')
            else:
                flash('ስህተት፡ Excel ፋይሉ ላይ `username` እና `password` የሚሉ ኮለሞች መኖር አለባቸው!')
        except Exception as e:
            db.session.rollback()
            flash(f'ፋይሉን ማንበብ አልተቻለም፡ {str(e)}')
    return redirect(url_for('index'))

@app.route('/upload_comments', methods=['GET', 'POST'])
@login_required
def upload_comments():
    if request.method == 'GET':
        return redirect(url_for('index'))
    if current_user.role == 'viewer':
        return redirect(url_for('index'))
    
    file = request.files.get('comment_file')
    if file:
        try:
            filename = file.filename.lower()
            file_stream = io.BytesIO(file.read())
            if filename.endswith('.docx'):
                doc = Document(file_stream)
                for para in doc.paragraphs:
                    if para.text.strip():
                        db.session.add(Comment(text=para.text.strip()))
                db.session.commit()
                flash('ኮሜንቶች በተሳካ ሁኔታ ተጭነዋል!')
            elif filename.endswith('.xlsx'):
                df = pd.read_excel(file_stream, engine='openpyxl')
                for col in df.columns:
                    for val in df[col].dropna():
                        val_str = str(val).strip()
                        if val_str and val_str != 'nan':
                            db.session.add(Comment(text=val_str))
                db.session.commit()
                flash('ኮሜንቶች በተሳካ ሁኔታ ተጭነዋል!')
            else:
                flash('ስህተት፡ እባክዎን የ .xlsx ወይም .docx ፋይል ይጫኑ!')
        except Exception as e:
            db.session.rollback()
            flash(f'ፋይሉን ማንበብ አልተቻለም፡ {str(e)}')
    return redirect(url_for('index'))

@app.route('/add_link', methods=['GET', 'POST'])
@login_required
def add_link():
    if request.method == 'GET':
        return redirect(url_for('index'))
    if current_user.role == 'viewer':
        return redirect(url_for('index'))
    
    url = request.form.get('post_url')
    if url:
        try:
            db.session.add(Link(url=url.strip()))
            db.session.commit()
            flash('ሊንክ ተጨምሯል!')
        except Exception as e:
            db.session.rollback()
            flash(f'ሊንኩን መመዝገብ አልተቻለም፡ {str(e)}')
    return redirect(url_for('index'))

@app.route('/delete/account/<int:id>')
@login_required
def delete_account(id):
    if current_user.role != 'viewer':
        acc = Account.query.get(id)
        if acc:
            db.session.delete(acc)
            db.session.commit()
    return redirect(url_for('index'))

@app.route('/delete/comment/<int:id>')
@login_required
def delete_comment(id):
    if current_user.role != 'viewer':
        c = Comment.query.get(id)
        if c:
            db.session.delete(c)
            db.session.commit()
    return redirect(url_for('index'))

@app.route('/delete/link/<int:id>')
@login_required
def delete_link(id):
    if current_user.role != 'viewer':
        l = Link.query.get(id)
        if l:
            db.session.delete(l)
            db.session.commit()
    return redirect(url_for('index'))

@app.route('/add_user', methods=['POST'])
@login_required
def add_user():
    if current_user.role != 'admin':
        return redirect(url_for('index'))
    
    username = request.form.get('new_username')
    password = generate_password_hash(request.form.get('new_password'))
    role = request.form.get('new_role', 'operator')
    
    if not User.query.filter_by(username=username).first():
        db.session.add(User(username=username, password=password, role=role))
        db.session.commit()
        flash('አዲስ ዩዘር ተፈጥሯል!')
    return redirect(url_for('index'))

@app.route('/delete_user/<int:id>')
@login_required
def delete_user(id):
    if current_user.role == 'admin':
        u = User.query.get(id)
        if u and u.id != current_user.id:
            db.session.delete(u)
            db.session.commit()
    return redirect(url_for('index'))

# --- አውቶሜሽን ቦት ማስጀመሪያ ከ Like እና Comment Interaction ጋር ---
def run_tiktok_bot():
    with app.app_context():
        accounts = Account.query.all()
        comments = Comment.query.all()
        links = Link.query.all()
        settings = Setting.query.first()
        
        if not accounts or not links:
            return
        
        min_d = settings.min_delay if settings else 5
        max_d = settings.max_delay if settings else 12
        
        proxies = [
            # {"server": "http://your_proxy_ip:port", "username": "proxy_user", "password": "proxy_password"},
        ]
        
        with sync_playwright() as p:
            for link in links:
                for acc in accounts:
                    delay = random.randint(min_d, max_d)
                    time.sleep(delay)
                    
                    proxy_cfg = random.choice(proxies) if proxies else None
                    
                    browser = p.chromium.launch(
                        headless=True,
                        args=[
                            "--disable-blink-features=AutomationControlled",
                            "--disable-infobars",
                            "--no-sandbox",
                            "--disable-setuid-sandbox",
                            "--disable-dev-shm-usage"
                        ],
                        proxy=proxy_cfg
                    )
                    
                    context = browser.new_context(
                        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                        viewport={'width': 1280, 'height': 800}
                    )
                    page = context.new_page()
                    
                    status = 'Success'
                    actions_taken = []
                    try:
                        # ሊንኩን መክፈት
                        page.goto(link.url, timeout=60000, wait_until='domcontentloaded')
                        time.sleep(5)
                        
                        # የሰው ልጅ እንቅስቃሴዎችን ማስመሰል
                        page.mouse.move(random.randint(100, 400), random.randint(100, 400))
                        page.evaluate("window.scrollBy(0, 200)")
                        time.sleep(2)
                        
                        # 1. ላይክ ማድረግ (Like Action)
                        try:
                            like_button = page.locator('[data-e2e="like-icon"], span[class*="ButtonLike"]').first
                            if like_button.is_visible(timeout=5000):
                                like_button.click()
                                actions_taken.append("Liked")
                                time.sleep(2)
                            else:
                                actions_taken.append("Like button not visible")
                        except Exception as e:
                            actions_taken.append("Like skipped")
                        
                        # 2. ኮሜንት መጻፍ (Comment Action)
                        if comments:
                            chosen_comment = random.choice(comments).text
                            try:
                                comment_box = page.locator('[data-e2e="comment-input"], div[contenteditable="true"]').first
                                if comment_box.is_visible(timeout=5000):
                                    comment_box.click()
                                    time.sleep(1)
                                    comment_box.fill(chosen_comment)
                                    time.sleep(1)
                                    page.keyboard.press("Enter")
                                    actions_taken.append(f"Commented: '{chosen_comment}'")
                                    time.sleep(3)
                                else:
                                    actions_taken.append("Comment box not visible")
                            except Exception as e:
                                actions_taken.append("Comment skipped")
                        else:
                            actions_taken.append("No comments in pool")
                        
                        action_desc = f"Stealth | Acc: {acc.username} | Actions: {', '.join(actions_taken)}"
                    
                    except Exception as e:
                        action_desc = f"Failed for {acc.username}: {str(e)}"
                        status = 'Failed'
                    finally:
                        browser.close()
                    
                    report = Report(
                        timestamp=time.strftime('%Y-%m-%d %H:%M:%S'),
                        account_used=acc.username,
                        target_link=link.url,
                        action_performed=action_desc,
                        status=status
                    )
                    db.session.add(report)
                    db.session.commit()

@app.route('/start_automation', methods=['POST'])
@login_required
def start_automation():
    if current_user.role == 'viewer':
        return redirect(url_for('index'))
    
    thread = threading.Thread(target=run_tiktok_bot)
    thread.daemon = True
    thread.start()
    
    flash('ቲክቶክ ስቴልዝ አውቶሜሽን ቦቱ በደህንነትና በአይፒ ሮቴሽን መሥራት ጀምሯል!')
    return redirect(url_for('index'))

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)