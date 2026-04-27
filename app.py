from dotenv import load_dotenv
load_dotenv()

from flask import Flask, render_template, request, redirect, session, jsonify, send_file
from datetime import datetime, timedelta
import sqlite3
import io
import os
from openpyxl import Workbook

# ---------------- 確定メール送信 ----------------
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr

def format_time_range(time_str):
    t = time_str[:5]  # "09:30" に統一

    start = datetime.strptime(t, "%H:%M")
    end = start + timedelta(minutes=30)

    return f"{start.strftime('%H:%M')}～{end.strftime('%H:%M')}"

def send_admin_mail(subject, body):
    import os
    import smtplib
    from email.mime.text import MIMEText



    ADMIN_EMAIL = os.getenv("ADMIN_EMAIL")
    SMTP_USER = os.getenv("SMTP_USER")
    SMTP_PASS = os.getenv("SMTP_PASS")

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = SMTP_USER
    msg["To"] = ADMIN_EMAIL

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(SMTP_USER, SMTP_PASS)
    server.send_message(msg)
    server.quit()

def mail_new(date, time, name, phone):
    send_admin_mail(
        "【新規予約】",
        f"""
新規予約が入りました

--------------------
予約日付：{date}
予約時間：{format_time_range(time)}
氏名：{name}
電話：{phone}
--------------------
"""
    )


def mail_edit(date, time, name, phone):
    send_admin_mail(
        "【予約変更】",
        f"""
予約が変更されました

--------------------
予約日付：{date}
予約時間：{format_time_range(time)}
氏名：{name}
電話：{phone}
--------------------
"""
    )


def mail_delete(date, time, name, phone):
    send_admin_mail(
        "【予約削除】",
        f"""
予約が削除されました

--------------------
予約日付：{date}
予約時間：{format_time_range(time)}
氏名：{name}
電話：{phone}
--------------------
"""
    )

def send_mail(row):
    SMTP_USER = os.getenv("SMTP_USER")
    SMTP_PASS = os.getenv("SMTP_PASS")

    name, date, time, email = row

    print("RAW TIME =", time)

    try:
        t = time[:5]

        # 「09」みたいな場合の補正
        if ":" not in t:
            t = f"{t}:00"

        start = datetime.strptime(t, "%H:%M")
        end = start + timedelta(minutes=30)

        # ★ここ重要（必ず2桁表示）
        time_range = f"{start.strftime('%H:%M')}～{end.strftime('%H:%M')}"

    except:
        time_range = time or ""

    body = f"""{name} 様

ガス点検の予約が確定しました。

■日時
{date} {time_range}

上記の日時にお伺いいたします。なお、都合により時間が前後する場合もございますが、ご容赦ください。
また、このメール受信以降に予約の変更を希望される際は、お手数ですがお電話にてご相談ください。
"""

    msg = MIMEText(body, "plain", "utf-8")

    # 件名（UTF-8）
    msg["Subject"] = str(Header("予約確定のお知らせ", "utf-8"))

    # ★ここも重要（FromもUTF-8にする）
    SMTP_USER = os.getenv("SMTP_USER")

    msg["From"] = formataddr(
        (str(Header("ガス点検", "utf-8")), SMTP_USER)
    )

    msg["To"] = email

    smtp = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    smtp.login(SMTP_USER, SMTP_PASS)
    smtp.send_message(msg)
    smtp.quit()
# ---------------- 確定メール送信ここまで ----------------

from datetime import timezone
JST = timezone(timedelta(hours=9))

app = Flask(__name__)

app.secret_key = "supersecretkey"
app.permanent_session_lifetime = timedelta(days=7)

# ---------------- DB ----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "reservation.db")

print("DB PATH =", DB_FILE)

# ---------------- DB初期化 ----------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS settings (
        id INTEGER PRIMARY KEY,
        start_date TEXT,
        end_date TEXT
    )
    """)


    c.execute("""
    CREATE TABLE IF NOT EXISTS reservations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        time TEXT,
        consumer_code TEXT,
        name TEXT,
        phone TEXT,
        address TEXT,
        action TEXT,
        is_deleted INTEGER DEFAULT 0,
        created_at TEXT
        )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS blocked_times (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        start_time TEXT,
        end_time TEXT
    )
    """)

    conn.commit()
    conn.close()
init_db()

# ---------------- トップ ----------------
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        code = request.form.get('consumer_code')

        if not code:
            return "消費者コードを入力してください"

        session['code'] = code
        action = request.form.get('action')

        if action == "新規":
            return redirect('/new')
        elif action == "変更":
            return redirect('/edit')
        elif action == "削除":
            return redirect('/delete')
        elif action == "確認":   # ★ここに追加
            return redirect('/view')

    return render_template('index.html')

# ---------------- ログイン ----------------
import os

ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD")

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if request.form.get('password') == ADMIN_PASSWORD:
            session['login'] = True
            return redirect('/admin_menu')
        return render_template('login.html', error="パスワードが違います")

    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')

# ---------------- 管理者メイン画面 ----------------
@app.route('/admin_menu')
def admin_menu():
    if not session.get('login'):
        return redirect('/login')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("SELECT start_date, end_date FROM settings WHERE id=1")
    row = c.fetchone()
    conn.close()

    if row:
        start_date, end_date = row
    else:
        start_date, end_date = None, None

    return render_template(
        'admin_menu.html',
        start_date=start_date,
        end_date=end_date
    )

# ---------------- 予約可能期間設定 ----------------
@app.route('/admin_setting')
def admin_setting():
    if not session.get('login'):
        return redirect('/login')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("SELECT start_date, end_date FROM settings WHERE id=1")
    row = c.fetchone()
    conn.close()

    start = row[0] if row else ""
    end = row[1] if row else ""

    return render_template("admin_setting.html", start=start, end=end)


@app.route('/save_setting', methods=['POST'])
def save_setting():

    if not session.get('login'):
        return redirect('/login')

    start_date = request.form.get('start_date')
    end_date = request.form.get('end_date')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # 既存あれば更新、なければ作成
    c.execute("SELECT id FROM settings WHERE id=1")
    if c.fetchone():
        c.execute("""
            UPDATE settings
            SET start_date=?, end_date=?
            WHERE id=1
        """, (start_date, end_date))
    else:
        c.execute("""
            INSERT INTO settings (id, start_date, end_date)
            VALUES (1, ?, ?)
        """, (start_date, end_date))

    conn.commit()
    conn.close()

    return redirect('/admin_menu')

@app.route('/clear_setting', methods=['POST'])
def clear_setting():

    if not session.get('login'):
        return redirect('/login')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        UPDATE settings
        SET start_date=NULL, end_date=NULL
        WHERE id=1
    """)

    conn.commit()
    conn.close()

    return redirect('/admin_menu')

# ---------------- admin（★ここ改修） ----------------
@app.route('/admin')
def admin():
    if not session.get('login'):
        return redirect('/login')

    code = request.args.get('code', '')
    name = request.args.get('name', '')
    confirmed = request.args.get('confirmed', '')

    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    created_from = request.args.get('created_from', '')
    created_to = request.args.get('created_to', '')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    query = """
        SELECT id,date,time,consumer_code,name,phone,address,email,action,created_at,is_confirmed
        FROM reservations
        WHERE is_deleted = 0
    """

    params = []

    if confirmed != "":
        query += " AND is_confirmed = ?"
        params.append(confirmed)

    if code:
        query += " AND consumer_code LIKE ?"
        params.append(f"%{code}%")

    if name:
        query += " AND name LIKE ?"
        params.append(f"%{name}%")

    if date_from:
        query += " AND date >= ?"
        params.append(date_from)

    if date_to:
        query += " AND date <= ?"
        params.append(date_to)

    if created_from:
        query += " AND date(created_at) >= ?"
        params.append(created_from)

    if created_to:
        query += " AND date(created_at) <= ?"
        params.append(created_to)

    # ★ここが重要（並び順）
    query += """
        ORDER BY
            REPLACE(consumer_code,'-','') ASC,
            datetime(created_at) DESC
    """

    c.execute(query, params)
    rows = c.fetchall()
    conn.close()

    return render_template("admin.html", reservations=rows)

@app.route('/toggle_confirm', methods=['POST'])
def toggle_confirm():
    id = request.form['id']
    confirmed = 1 if request.form.get('confirmed') == 'on' else 0

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # 更新
    c.execute("UPDATE reservations SET is_confirmed=? WHERE id=?", (confirmed, id))

    # ★メール用にデータ取得
    c.execute("""
        SELECT name, date, time, email
        FROM reservations
        WHERE id=?
    """, (id,))
    row = c.fetchone()

    conn.commit()
    conn.close()

    # ★ここ重要：確定＋メールありのときだけ送信
    if confirmed == 1 and row and row[3]:
        send_mail(row)

    return redirect('/admin')

@app.route('/admin_delete', methods=['POST'])
def admin_delete():

    reservation_id = request.form.get('id')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # ★今の状態を保存してから削除
    c.execute("""
        UPDATE reservations
        SET is_deleted = 1,
            before_action = action
        WHERE id = ?
    """, (reservation_id,))

    conn.commit()
    conn.close()

    return redirect('/admin')

@app.route('/admin_edit/<int:id>')
def admin_edit(id):
    if not session.get('login'):
        return redirect('/login')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        SELECT id,date,time,name,phone,address,email
        FROM reservations
        WHERE id=?
    """, (id,))

    data = c.fetchone()
    conn.close()

    return render_template("admin_edit.html", data=data)

@app.route('/admin_edit_save', methods=['POST'])
def admin_edit_save():

    if not session.get('login'):
        return redirect('/login')

    data = request.form

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        UPDATE reservations
        SET date=?, time=?, name=?, phone=?, address=?, email=?
        WHERE id=?
    """, (
        data['date'],
        data['time'],
        data['name'],
        data['phone'],
        data['address'],
        data['email'],
        data['id']
    ))

    conn.commit()
    conn.close()

    return redirect('/admin')

@app.route('/admin_deleted')
def admin_deleted():
    if not session.get('login'):
        return redirect('/login')

    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    name = request.args.get('name', '')
    code = request.args.get('code', '')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    query = """
        SELECT id,date,time,consumer_code,name,phone,address,action,created_at
        FROM reservations
        WHERE is_deleted = 1
    """

    params = []

    if date_from:
        query += " AND date >= ?"
        params.append(date_from)

    if date_to:
        query += " AND date <= ?"
        params.append(date_to)

    if name:
        query += " AND name LIKE ?"
        params.append(f"%{name}%")

    if code:
        query += " AND consumer_code LIKE ?"
        params.append(f"%{code}%")

    query += " ORDER BY created_at DESC"

    c.execute(query, params)
    rows = c.fetchall()
    conn.close()

    # ★時間変換
    new_rows = []
    for r in rows:
        t = r[2][:5]
        h, m = map(int, t.split(":"))

        end_h = h
        end_m = m + 30
        if end_m >= 60:
            end_h += 1
            end_m -= 60

        time_range = f"{t}～{str(end_h).zfill(2)}:{str(end_m).zfill(2)}"

        new_rows.append((
            r[0], r[1], time_range,
            r[3], r[4], r[5], r[6], r[7], r[8]
        ))

    return render_template("admin_deleted.html", reservations=new_rows)

@app.route('/admin_restore', methods=['POST'])
def admin_restore():

    if not session.get('login'):
        return redirect('/login')

    reservation_id = request.form.get('id')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # ★保存してた状態を戻す
    c.execute("""
        UPDATE reservations
        SET is_deleted = 0,
            action = before_action
        WHERE id = ?
    """, (reservation_id,))

    conn.commit()
    conn.close()

    return redirect('/admin_deleted')

@app.route('/admin_restore_multi', methods=['POST'])
def admin_restore_multi():
    if not session.get('login'):
        return redirect('/login')

    ids = request.form.getlist('ids')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    for i in ids:
        c.execute("""
            UPDATE reservations
            SET is_deleted = 0
            WHERE id = ?
        """, (i,))

    conn.commit()
    conn.close()

    return redirect('/admin_deleted')


@app.route('/admin_bulk_delete', methods=['POST'])
def admin_bulk_delete():
    if not session.get('login'):
        return redirect('/login')

    ids = request.form.getlist('ids')

    if not ids:
        return redirect('/admin_deleted')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    for i in ids:
        c.execute("""
            DELETE FROM reservations
            WHERE id = ?
        """, (i,))

    conn.commit()
    conn.close()

    return redirect('/admin_deleted')

# ---------------- admin_block ----------------
@app.route('/admin_block')
def admin_block():
    if not session.get('login'):
        return redirect('/login')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        SELECT id, date, start_time, end_time
        FROM blocked_times
        ORDER BY date ASC, start_time
    """)
    rows = c.fetchall()
    conn.close()

    return render_template("admin_block.html", blocks=rows)

@app.route('/export_block_excel')
def export_block_excel():
    if not session.get('login'):
        return redirect('/login')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        SELECT date, start_time, end_time
        FROM blocked_times
        ORDER BY date ASC, start_time ASC
    """)

    rows = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "予約ブロック"

    ws.append(["日付", "開始", "終了"])

    for r in rows:
        ws.append(r)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        download_name="reservation_blocks.xlsx",
        as_attachment=True
    )

@app.route('/add_block', methods=['POST'])
def add_block():

    if not session.get('login'):
        return redirect('/login')

    date = request.form['date']
    start_time = request.form['start_time']
    end_time = request.form['end_time']

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        INSERT INTO blocked_times (date, start_time, end_time)
        VALUES (?, ?, ?)
    """, (date, start_time, end_time))

    conn.commit()
    conn.close()

    return redirect('/admin_block')


@app.route('/delete_block/<int:block_id>')
def delete_block(block_id):

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        DELETE FROM blocked_times
        WHERE id = ?
    """, (block_id,))

    conn.commit()
    conn.close()

    return redirect('/admin_block')

# ---------------- new ----------------
@app.route('/new', methods=['GET', 'POST'])
def new():
    if not session.get('code'):
        return redirect('/')

    # ★予約可能期間取得
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("SELECT start_date, end_date FROM settings WHERE id=1")
    setting = c.fetchone()
    conn.close()

    start_date = setting[0] if setting else ""
    end_date = setting[1] if setting else ""

    # ★デフォルト
    data = {
        "date": "",
        "time": "",
        "name": "",
        "phone": "",
        "address": "",
        "email": ""
    }

    # ★POST時は復元
    if request.method == 'POST':
        data = request.form

    return render_template(
        "new.html",
        data=data,
        start_date=start_date,
        end_date=end_date
    )

# ---------------- get_times ----------------
@app.route('/get_times')
def get_times():

    date = request.args.get('date')
    if not date:
        return jsonify([])

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # ★追加：予約可能期間チェック
    c.execute("SELECT start_date, end_date FROM settings WHERE id=1")
    setting = c.fetchone()

    if setting:
        start_date, end_date = setting

        if start_date and date < start_date:
            return jsonify([])

        if end_date and date > end_date:
            return jsonify([])

    # ①予約取得（削除済は完全に無視）
    c.execute("""
        SELECT time
        FROM reservations
        WHERE date=? AND is_deleted=0
    """, (date,))
    rows = c.fetchall()

    reserved = set()
    for (t,) in rows:
        reserved.add(t[:5])

    # ②ブロック取得
    c.execute("""
        SELECT start_time, end_time
        FROM blocked_times
        WHERE date=?
    """, (date,))
    blocks = c.fetchall()

    conn.close()

    # ★ここが本質（今から24時間）
    limit_time = datetime.now(JST) + timedelta(hours=24)

    # ④時間生成
    slots = []

    start = datetime.strptime("09:30", "%H:%M")
    end = datetime.strptime("16:30", "%H:%M")

    while start <= end:

        t = start.strftime("%H:%M")
        current_dt = datetime.strptime(
            date + " " + t,
            "%Y-%m-%d %H:%M"
        )
        current_dt = current_dt.replace(tzinfo=JST)

        # ①予約済みチェック
        if t in reserved:
            start += timedelta(minutes=30)
            continue

        # ②ブロックチェック
        blocked = False

        for b_start, b_end in blocks:
            bs = datetime.strptime(date + " " + b_start, "%Y-%m-%d %H:%M")
            bs = bs.replace(tzinfo=JST)

            be = datetime.strptime(date + " " + b_end, "%Y-%m-%d %H:%M")
            be = be.replace(tzinfo=JST)

            # 日跨ぎ対応
            if be <= bs:
                be += timedelta(days=1)

            if bs <= current_dt < be:
                blocked = True
                break

        if blocked:
            start += timedelta(minutes=30)
            continue

        # ★③24時間ルール（ここだけ）
        if current_dt < limit_time:
            start += timedelta(minutes=30)
            continue

        # OK
        slots.append(t)

        start += timedelta(minutes=30)

    return jsonify(slots)


@app.route('/confirm', methods=['POST'])
def confirm():

    data = {
        "date": request.form.get("date"),
        "time": request.form.get("time"),
        "name": request.form.get("name"),
        "phone": request.form.get("phone"),
        "address": request.form.get("address"),
        "email": request.form.get("email")
    }

    return render_template("confirm.html", data=data)

# ---------------- create ----------------
@app.route('/create_confirm', methods=['POST'])
def create_confirm():

    data = request.form
    code = session.get('code')

    if not code:
        return "ログイン情報なし"

    target_dt = datetime.strptime(
        data['date'] + " " + data['time'][:5],
        "%Y-%m-%d %H:%M"
    ).replace(tzinfo=JST)

    # ★ここが本質（今から24時間）
    limit_time = datetime.now(JST) + timedelta(hours=24)

    if target_dt < limit_time:
        return "24時間後以降の予約しかできません"

    # ★通常処理
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        INSERT INTO reservations
        (date,time,consumer_code,name,phone,address,email,action,created_at)
        VALUES (?,?,?,?,?,?,?,?,?)
    """, (
        data['date'],
        data['time'][:5],
        code,
        data['name'],
        data['phone'],
        data['address'],
        data['email'],
        "新規",
        datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
    ))

    conn.commit()
    conn.close()

    mail_new(
    data['date'],
    data['time'][:5],
    data['name'],
    data['phone']
)

    return render_template("complete.html", data=data)

# ---------------- check_day ----------------
@app.route('/check_day')
def check_day():

    date = request.args.get('date')
    if not date:
        return jsonify({"ok": False})

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        SELECT start_time, end_time
        FROM blocked_times
        WHERE date=?
    """, (date,))
    blocks = c.fetchall()

    conn.close()

    # 予約枠（固定）
    slots = []
    start = datetime.strptime("09:30", "%H:%M")
    end = datetime.strptime("16:30", "%H:%M")

    while start <= end:
        slots.append(start.strftime("%H:%M"))
        start += timedelta(minutes=30)

    # ブロック削除
    for b_start, b_end in blocks:
        bs = datetime.strptime(b_start, "%H:%M")
        be = datetime.strptime(b_end, "%H:%M")

        slots = [
            t for t in slots
            if not (bs <= datetime.strptime(t, "%H:%M") < be)
        ]

    return jsonify({
        "ok": len(slots) > 0,
        "message": "この日は予約できません" if len(slots) == 0 else ""
    })

# ---------------- edit ----------------
@app.route('/edit')
def edit():
    code = session.get('code')

    if not code:
        return redirect('/')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # ★予約データ取得
    c.execute("""
        SELECT id,date,time,name,phone,address,email
        FROM reservations
        WHERE consumer_code=?
        ORDER BY id DESC
        LIMIT 1
    """, (code,))
    data = c.fetchone()

    # ★期間取得
    c.execute("SELECT start_date, end_date FROM settings WHERE id=1")
    setting = c.fetchone()

    conn.close()

    if not data:
        return "予約データがありません"

    # ★ここ強化
    if setting and setting[0] and setting[1]:
        start_date = setting[0]
        end_date = setting[1]
    else:
        start_date = ""
        end_date = ""

    return render_template(
        "edit.html",
        data=data,
        start_date=start_date,
        end_date=end_date
    )

@app.route('/edit_confirm', methods=['POST'])
def edit_confirm():

    print(request.form)

    data = {
        "date": request.form.get("date"),
        "time": request.form.get("time"),
        "name": request.form.get("name"),
        "phone": request.form.get("phone"),
        "address": request.form.get("address"),
        "email": request.form.get("email")
    }

    return render_template("edit_confirm.html", data=data)

# ---------------- edit_save ----------------
@app.route('/edit_save', methods=['POST'])
def edit_save():
    data = request.form
    code = session.get('code')

    print(data)

    if not code:
        return redirect('/')

    email = request.form.get('email', '')  # ←★これ追加

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        INSERT INTO reservations
        (date,time,consumer_code,name,phone,address,email,action,created_at)
        VALUES (?,?,?,?,?,?,?,?,?)
    """, (
        data['date'],
        data['time'][:5],
        code,
        data['name'],
        data['phone'],
        data['address'],
        email,
        "変更",
        datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
    ))

    conn.commit()
    conn.close()

    mail_edit(
    data['date'],
    data['time'][:5],
    data['name'],
    data['phone']
)

    return render_template("edit_complete.html", data=data)   

# ---------------- delete ----------------
@app.route('/delete')
def delete():
    code = session.get('code')

    if not code:
        return redirect('/')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        SELECT id,date,time,name,phone,address,email,is_deleted
        FROM reservations
        WHERE consumer_code=? AND is_deleted=0
        ORDER BY id DESC
        LIMIT 1
    """, (code,))

    row = c.fetchone()
    conn.close()

    if not row:
        return render_template("delete.html", data=None)

    # ★ここから置き換え
    if not row or not row[2]:
        return render_template("delete.html", data=None)

    try:
        start = datetime.strptime(row[2][:5], "%H:%M")
        end = start + timedelta(minutes=30)
        time_range = f"{start.strftime('%H:%M')}～{end.strftime('%H:%M')}"
    except:
        time_range = row[2] or ""
# ★ここまで

    data = (row[0], row[1], time_range, row[3], row[4], row[5], row[6])

    return render_template("delete.html", data=data)

@app.route('/delete', methods=['POST'])
def delete_post():
    code = session.get('code')

    if not code:
        return redirect('/')

    date = request.form.get('date')
    time = request.form.get('time')
    name = request.form.get('name')
    phone = request.form.get('phone')
    address = request.form.get('address')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        INSERT INTO reservations
        (date,time,consumer_code,name,phone,address,action,created_at)
        VALUES (?,?,?,?,?,?,?,?)
    """, (
        date,
        time,
        code,
        name,
        phone,
        address,
        "削除",
        datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
    ))

    conn.commit()
    conn.close()

    mail_delete(
    date,
    time[:5],
    name,
    phone
)

    return render_template("delete_done.html")

# ---------------- 予約内容確認 ----------------
@app.route('/view')
def view():
    code = session.get('code')

    if not code:
        return redirect('/')

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # ① ここでrowを作る
    c.execute("""
        SELECT id,date,time,name,phone,address,email,is_confirmed
        FROM reservations
        WHERE consumer_code=? AND is_deleted=0
        ORDER BY id DESC
        LIMIT 1
    """, (code,))

    row = c.fetchone()
    print(row)
    conn.close()

    # ② rowがない場合
    if not row:
        return render_template("view.html", data=None)

    # ③ time整形
    try:
        start = datetime.strptime(row[2][:5], "%H:%M")
        end = start + timedelta(minutes=30)
        time_range = f"{start.strftime('%H:%M')}～{end.strftime('%H:%M')}"
    except:
        time_range = row[2] or ""

    # ④ ここでdata作る（←ここでrow使うのが正解）
    data = (
        row[0],  # id
        row[1],  # date
        time_range,
        row[3],  # name
        row[4],  # phone
        row[5],  # address
        row[6],  # email
        row[7]   # is_confirmed
    )

    return render_template("view.html", data=data)

# ---------------- 予約excel出力 ----------------
@app.route('/export_excel')
def export_excel():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        SELECT date,time,created_at,consumer_code,name,address,phone,action
        FROM reservations
        WHERE is_deleted = 0
        ORDER BY created_at DESC
    """)

    rows = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "予約一覧"

    ws.append([
        "予約日",
        "予約時間",
        "申込日時",
        "消費者コード",
        "氏名",
        "住所",
        "電話番号",
        "状態"
    ])

    def format_range(t):
        if not t:
            return ""

        t = t[:5]

        h, m = map(int, t.split(":"))
        end_h = h
        end_m = m + 30

        if end_m >= 60:
            end_h += 1
            end_m -= 60

        return f"{t}～{str(end_h).zfill(2)}:{str(end_m).zfill(2)}"

    for r in rows:
        date, time, created_at, code, name, address, phone, action = r

        ws.append([
            date,
            format_range(time),   # ★ここが重要
            created_at,
            code,
            name,
            address,
            phone,
            action
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        download_name="reservations.xlsx",
        as_attachment=True
    )

# ---------------- 予約削除excel出力 ----------------
@app.route('/export_deleted_excel')
def export_deleted_excel():

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        SELECT date,time,created_at,consumer_code,name,address,phone,action
        FROM reservations
        WHERE is_deleted = 1
        ORDER BY created_at DESC
    """)

    rows = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "削除一覧"

    ws.append([
        "予約日",
        "予約時間",
        "申込日時",
        "消費者コード",
        "氏名",
        "住所",
        "電話番号",
        "状態"
    ])

    def format_range(t):
        if not t:
            return ""

        t = t[:5]
        h, m = map(int, t.split(":"))

        end_h = h
        end_m = m + 30

        if end_m >= 60:
            end_h += 1
            end_m -= 60

        return f"{t}～{str(end_h).zfill(2)}:{str(end_m).zfill(2)}"

    for r in rows:
        date, time, created_at, code, name, address, phone, action = r

        ws.append([
            date,
            format_range(time),
            created_at,
            code,
            name,
            address,
            phone,
            action
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        download_name="deleted_reservations.xlsx",
        as_attachment=True
    )

# ---------------- 起動 ----------------
if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=10000, debug=True)