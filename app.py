# =========================
# 環境変数ロード
# =========================
from dotenv import load_dotenv
import os

load_dotenv()

# =========================
# 標準ライブラリ
# =========================
import io
from datetime import datetime, timedelta, timezone

# =========================
# サードパーティ
# =========================
from flask import Flask, render_template, request, redirect, session, jsonify, send_file
from supabase import create_client
from openpyxl import Workbook
import resend

# =========================
# Flask設定
# =========================
app = Flask(__name__)

app.secret_key = os.getenv("SECRET_KEY")
app.permanent_session_lifetime = timedelta(days=7)

# =========================
# 予約間隔の選択肢
# =========================
INTERVAL_OPTIONS = [10, 15, 30, 60]

# =========================
# 日本時間
# =========================
JST = timezone(timedelta(hours=9))

# =========================
# Supabase接続
# =========================
supabase = create_client(
    os.getenv("SUPABASE_URL"),
    os.getenv("SUPABASE_KEY")
)

# =========================
# Resend設定
# =========================
RESEND_API_KEY = os.getenv("RESEND_API_KEY")
MAIL_FROM = os.getenv("MAIL_FROM")
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL")

if RESEND_API_KEY:
    resend.api_key = RESEND_API_KEY


# =========================================================
# 共通処理
# =========================================================

def get_reservation_interval():
    """
    settingsテーブルから現在の予約間隔を取得する。
    取得できない場合は30分を使用。
    """
    try:
        res = supabase.table("settings") \
            .select("interval") \
            .eq("id", 1) \
            .limit(1) \
            .execute()

        setting = res.data[0] if res.data else None

        interval = int(setting.get("interval", 30)) if setting else 30

        if interval not in INTERVAL_OPTIONS:
            interval = 30

        return interval

    except Exception:
        return 30


def format_time_range(time_str):
    """
    09:30 → 09:30～09:40
    予約設定のintervalに合わせて終了時刻を計算。
    """
    try:
        t = (time_str or "")[:5]

        if ":" not in t:
            return time_str or ""

        start = datetime.strptime(t, "%H:%M")
        interval = get_reservation_interval()
        end = start + timedelta(minutes=interval)

        return f"{start.strftime('%H:%M')}～{end.strftime('%H:%M')}"

    except Exception:
        return time_str or ""


# =========================================================
# メール送信
# =========================================================

def send_resend_mail(to_email, subject, body):
    """
    Resend APIを使ってメール送信。
    送信失敗しても予約処理自体は止めない。
    """

    if not RESEND_API_KEY:
        print("メール送信失敗: RESEND_API_KEY が設定されていません")
        return False

    if not MAIL_FROM:
        print("メール送信失敗: MAIL_FROM が設定されていません")
        return False

    if not to_email:
        print("メール送信失敗: 宛先メールアドレスがありません")
        return False

    try:
        params = {
            "from": MAIL_FROM,
            "to": [to_email],
            "subject": subject,
            "text": body
        }

        result = resend.Emails.send(params)

        print("メール送信成功")

        return True

    except Exception as e:
        print("メール送信失敗:", type(e).__name__, str(e))

        return False


def send_admin_mail(subject, body):
    """
    管理者宛メール。
    """

    if not ADMIN_EMAIL:
        print("管理者メール送信失敗: ADMIN_EMAIL が設定されていません")
        return False

    return send_resend_mail(
        ADMIN_EMAIL,
        subject,
        body
    )


def mail_new(data, time, name, phone):
    send_admin_mail(
        "【新規予約】",
        f"""新規予約が入りました

--------------------
予約日付：{data}
予約時間：{format_time_range(time)}
氏名：{name}
電話：{phone}
--------------------
"""
    )


def mail_edit(data, time, name, phone):
    send_admin_mail(
        "【予約変更】",
        f"""予約が変更されました

--------------------
予約日付：{data}
予約時間：{format_time_range(time)}
氏名：{name}
電話：{phone}
--------------------
"""
    )


def mail_delete(data, time, name, phone):
    send_admin_mail(
        "【予約削除】",
        f"""予約が削除されました

--------------------
予約日付：{data}
予約時間：{format_time_range(time)}
氏名：{name}
電話：{phone}
--------------------
"""
    )


def send_mail(row):
    """
    予約確定時に予約者へ送信するメール。
    """

    try:
        name, data, time, email = row

        if not email:
            print("予約確定メール送信失敗: メールアドレスなし")
            return False

        time_range = format_time_range(time)

        body = f"""{name} 様

ガス点検の予約が確定しました。

■日時
{data} {time_range}

上記の日時にお伺いしますので、ご在宅をお願いいたします。なお、都合により時間が前後する場合もございますが、ご容赦ください。
また、このメール受信以降に予約の変更を希望される際は、お手数ですがお電話にてご相談ください。
"""

        return send_resend_mail(
            email,
            "予約確定のお知らせ",
            body
        )

    except Exception as e:
        print("予約確定メール処理失敗:", type(e).__name__, str(e))
        return False


# =========================================================
# トップ
# =========================================================

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

        elif action == "確認":
            return redirect('/view')

    return render_template('index.html')


# =========================================================
# ログイン
# =========================================================

ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD")


@app.route('/login', methods=['GET', 'POST'])
def login():

    if request.method == 'POST':

        pw = request.form.get('password')

        admin_password = os.getenv("ADMIN_PASSWORD")

        if pw == admin_password:
            session['login'] = True
            return redirect('/admin_menu')

        return render_template(
            'login.html',
            error="パスワードが違います"
        )

    return render_template('login.html')


@app.route('/logout')
def logout():

    session.clear()

    return redirect('/login')


# =========================================================
# 管理者メイン画面
# =========================================================

@app.route('/admin_menu')
def admin_menu():

    if not session.get('login'):
        return redirect('/login')

    # ① 予約ブロック一覧
    res = supabase.table("blocked_times") \
        .select("id,data,start_time,end_time") \
        .order("data", desc=False) \
        .execute()

    blocks = res.data or []

    # ② 予約可能期間
    setting_res = supabase.table("settings") \
        .select("start_data,end_data,capacity") \
        .eq("id", 1) \
        .limit(1) \
        .execute()

    setting = setting_res.data[0] if setting_res.data else None

    start_data = setting.get("start_data") if setting else None
    end_data = setting.get("end_data") if setting else None
    capacity = setting.get("capacity", 1) if setting else 1

    return render_template(
        'admin_menu.html',
        blocks=blocks,
        start_data=start_data,
        end_data=end_data,
        capacity=capacity
    )


# =========================================================
# 予約可能期間設定
# =========================================================

@app.route('/admin_setting')
def admin_setting():

    if not session.get('login'):
        return redirect('/login')

    res = supabase.table("settings") \
        .select("*") \
        .eq("id", 1) \
        .execute()

    setting = res.data[0] if res.data else None

    start = setting.get("start_data", "") if setting else ""
    end = setting.get("end_data", "") if setting else ""

    capacity = setting.get("capacity", 1) if setting else 1

    start_time = setting.get("start_time", "09:30") if setting else "09:30"
    end_time = setting.get("end_time", "17:00") if setting else "17:00"
    interval = setting.get("interval", 30) if setting else 30

    try:
        interval = int(interval)
    except Exception:
        interval = 30

    if interval not in INTERVAL_OPTIONS:
        interval = 30

    return render_template(
        "admin_setting.html",
        start=start,
        end=end,
        capacity=capacity,
        start_time=start_time,
        end_time=end_time,
        interval=interval,
        interval_options=INTERVAL_OPTIONS
    )


@app.route('/save_setting', methods=['POST'])
def save_setting():

    if not session.get('login'):
        return redirect('/login')

    start_data = request.form.get('start_data')
    end_data = request.form.get('end_data')
    capacity = request.form.get('capacity')

    start_time = request.form.get('start_time')
    end_time = request.form.get('end_time')
    interval = request.form.get('interval')

    if not start_data or not end_data or not capacity:
        return redirect('/admin_setting')

    if not start_time or not end_time or not interval:
        return redirect('/admin_setting')

    try:
        capacity = int(capacity)
        interval = int(interval)
    except ValueError:
        return redirect('/admin_setting')

    if capacity < 1:
        return redirect('/admin_setting')

    if interval not in INTERVAL_OPTIONS:
        return redirect('/admin_setting')

    supabase.table("settings").upsert({
        "id": 1,
        "start_data": start_data,
        "end_data": end_data,
        "capacity": capacity,
        "start_time": start_time,
        "end_time": end_time,
        "interval": interval
    }).execute()

    return redirect('/admin_menu')


@app.route('/clear_setting', methods=['POST'])
def clear_setting():

    if not session.get('login'):
        return redirect('/login')

    supabase.table("settings").upsert({
        "id": 1,
        "start_data": None,
        "end_data": None
    }).execute()

    return redirect('/admin_menu')


# =========================================================
# 管理予約一覧
# =========================================================

@app.route('/admin')
def admin():

    if not session.get('login'):
        return redirect('/login')

    code = request.args.get('code', '')
    name = request.args.get('name', '')
    confirmed = request.args.get('confirmed', '')

    data_from = request.args.get('data_from', '')
    data_to = request.args.get('data_to', '')

    created_from = request.args.get('created_from', '')
    created_to = request.args.get('created_to', '')

    query = supabase.table("reservations") \
        .select("*") \
        .eq("is_deleted", False)

    if confirmed != "":
        query = query.eq(
            "is_confirmed",
            confirmed == "1"
        )

    if code:
        query = query.ilike(
            "consumer_code",
            f"%{code}%"
        )

    if name:
        query = query.ilike(
            "name",
            f"%{name}%"
        )

    if data_from:
        query = query.gte(
            "data",
            data_from
        )

    if data_to:
        query = query.lte(
            "data",
            data_to
        )

    if created_from:
        query = query.gte(
            "created_at",
            created_from
        )

    if created_to:
        query = query.lte(
            "created_at",
            created_to
        )

    res = query \
        .order("consumer_code", desc=False) \
        .order("created_at", desc=True) \
        .execute()

    reservations = res.data or []

    for r in reservations:
        r["status"] = r.get("status") or "新規"

    return render_template(
        "admin.html",
        reservations=reservations
    )


# =========================================================
# 予約確定
# =========================================================

@app.route('/toggle_confirm', methods=['POST'])
def toggle_confirm():

    if not session.get('login'):
        return redirect('/login')

    reservation_id = request.form.get('id')
    confirmed = request.form.get('confirmed') == 'on'

    if not reservation_id or not reservation_id.isdigit():
        return redirect('/admin')

    reservation_id = int(reservation_id)

    supabase.table("reservations") \
        .update({
            "is_confirmed": confirmed
        }) \
        .eq("id", reservation_id) \
        .execute()

    res = supabase.table("reservations") \
        .select("name,data,time,email") \
        .eq("id", reservation_id) \
        .limit(1) \
        .execute()

    if not res.data:
        return redirect('/admin')

    r = res.data[0]

    row = (
        r.get("name"),
        r.get("data"),
        r.get("time"),
        r.get("email")
    )

    if confirmed and row[3]:
        send_mail(row)

    return redirect('/admin')


# =========================================================
# 管理者による削除
# =========================================================

@app.route('/admin_delete', methods=['POST'])
def admin_delete():

    if not session.get('login'):
        return redirect('/login')

    reservation_id = request.form.get('id')

    if not reservation_id:
        return redirect('/admin')

    r = supabase.table("reservations") \
        .select("*") \
        .eq("id", int(reservation_id)) \
        .limit(1) \
        .execute()

    if not r.data:
        return redirect('/admin')

    supabase.table("reservations") \
        .update({
            "is_deleted": True
        }) \
        .eq("id", int(reservation_id)) \
        .execute()

    return redirect('/admin')


# =========================================================
# 管理者編集画面
# =========================================================

@app.route('/admin_edit/<int:id>')
def admin_edit(id):

    if not session.get('login'):
        return redirect('/login')

    res = supabase.table("reservations") \
        .select("id,data,time,name,phone,address,email") \
        .eq("id", id) \
        .execute()

    data = None

    if res.data:

        r = res.data[0]

        data = (
            r.get("id"),
            r.get("data"),
            r.get("time"),
            r.get("name"),
            r.get("phone"),
            r.get("address"),
            r.get("email")
        )

    return render_template(
        "admin_edit.html",
        data=data
    )


# =========================================================
# 管理者による予約変更
# =========================================================

@app.route('/admin_edit_save', methods=['POST'])
def admin_edit_save():

    if not session.get('login'):
        return redirect('/login')

    reservation_id = request.form.get('id')

    if not reservation_id:
        return redirect('/admin')

    data = request.form.get('date')
    time = request.form.get('time')
    name = request.form.get('name')
    phone = request.form.get('phone')
    address = request.form.get('address')
    email = request.form.get('email')

    if not data or not time:
        return "日付と時間は必須です"

    time = time[:5]

    check = supabase.table("reservations") \
        .select("id") \
        .eq("data", data) \
        .eq("time", time) \
        .eq("is_deleted", False) \
        .neq("id", int(reservation_id)) \
        .execute()

    if check.data:
        return "この日時にはすでに別の予約があります"

    supabase.table("reservations") \
        .update({
            "data": data,
            "time": time,
            "name": name,
            "phone": phone,
            "address": address,
            "email": email,
            "is_confirmed": False
        }) \
        .eq("id", int(reservation_id)) \
        .execute()

    return redirect('/admin')


# =========================================================
# 利用者による予約変更
# =========================================================

@app.route('/edit_save', methods=['POST'])
def edit_save():

    data = request.form
    code = session.get('code')

    if not code:
        return redirect('/')

    check = supabase.table("reservations") \
        .select("id") \
        .eq("data", data['data']) \
        .eq("time", data['time'][:5]) \
        .eq("is_deleted", False) \
        .neq("consumer_code", code) \
        .execute()

    if check.data:
        return "この時間は予約できません"

    supabase.table("reservations").insert({
        "data": data['data'],
        "time": data['time'][:5],
        "consumer_code": code,
        "name": data['name'],
        "phone": data['phone'],
        "address": data['address'],
        "email": data['email'],
        "status": "変更",
        "is_deleted": False,
        "is_confirmed": False,
        "created_at": datetime.now(JST).isoformat()
    }).execute()

    mail_edit(
        data['data'],
        data['time'][:5],
        data['name'],
        data['phone']
    )

    return render_template(
        "edit_complete.html",
        data=data
    )


# =========================================================
# 削除済み予約一覧
# =========================================================

@app.route('/admin_deleted')
def admin_deleted():

    if not session.get('login'):
        return redirect('/login')

    data_from = request.args.get('data_from', '')
    data_to = request.args.get('data_to', '')
    name = request.args.get('name', '')
    code = request.args.get('code', '')

    query = supabase.table("reservations") \
        .select("*") \
        .eq("is_deleted", True)

    if data_from:
        query = query.gte("data", data_from)

    if data_to:
        query = query.lte("data", data_to)

    if name:
        query = query.ilike(
            "name",
            f"%{name}%"
        )

    if code:
        query = query.ilike(
            "consumer_code",
            f"%{code}%"
        )

    res = query \
        .order("created_at", desc=True) \
        .execute()

    rows = res.data or []

    new_rows = []

    for r in rows:

        time_range = format_time_range(
            r.get("time", "")
        )

        new_rows.append({
            "id": r.get("id"),
            "data": r.get("data"),
            "time": time_range,
            "created_at": r.get("created_at"),
            "consumer_code": r.get("consumer_code"),
            "name": r.get("name"),
            "phone": r.get("phone"),
            "address": r.get("address"),
            "email": r.get("email"),
            "status": r.get("status"),
            "is_deleted": True
        })

    return render_template(
        "admin_deleted.html",
        reservations=new_rows
    )


# =========================================================
# 削除予約復元
# =========================================================

@app.route('/admin_restore', methods=['POST'])
def admin_restore():

    if not session.get('login'):
        return redirect('/login')

    reservation_id = request.form.get('id')

    supabase.table("reservations") \
        .update({
            "is_deleted": False
        }) \
        .eq("id", int(reservation_id)) \
        .execute()

    return redirect('/admin_deleted')


@app.route('/admin_restore_multi', methods=['POST'])
def admin_restore_multi():

    if not session.get('login'):
        return redirect('/login')

    ids = request.form.getlist('ids')

    for i in ids:

        supabase.table("reservations") \
            .update({
                "is_deleted": False
            }) \
            .eq("id", int(i)) \
            .execute()

    return redirect('/admin_deleted')


# =========================================================
# 削除済み予約を完全削除
# =========================================================

@app.route('/admin_bulk_delete', methods=['POST'])
def admin_bulk_delete():

    if not session.get('login'):
        return redirect('/login')

    ids = request.form.getlist('ids')

    if not ids:
        return redirect('/admin_deleted')

    for i in ids:

        supabase.table("reservations") \
            .delete() \
            .eq("id", int(i)) \
            .execute()

    return redirect('/admin_deleted')


# =========================================================
# 予約ブロック
# =========================================================

@app.route('/admin_block')
def admin_block():

    if not session.get('login'):
        return redirect('/login')

    res = supabase.table("blocked_times") \
        .select("id,data,start_time,end_time") \
        .order("data", desc=False) \
        .execute()

    rows = res.data or []

    return render_template(
        "admin_block.html",
        blocks=rows
    )


@app.route('/export_block_excel')
def export_block_excel():

    if not session.get('login'):
        return redirect('/login')

    res = supabase.table("blocked_times") \
        .select("id,data,start_time,end_time") \
        .order("data", desc=False) \
        .execute()

    rows = res.data or []

    wb = Workbook()
    ws = wb.active
    ws.title = "予約ブロック"

    ws.append([
        "日付",
        "開始",
        "終了"
    ])

    for r in rows:

        ws.append([
            r.get("data"),
            r.get("start_time"),
            r.get("end_time")
        ])

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

    try:

        data = request.form.get('data')
        start_time = request.form.get('start_time')
        end_time = request.form.get('end_time')

        if not data or not start_time or not end_time:
            return redirect('/admin_block')

        start_time = start_time[:5]
        end_time = end_time[:5]

        start_dt = datetime.strptime(
            start_time,
            "%H:%M"
        )

        end_dt = datetime.strptime(
            end_time,
            "%H:%M"
        )

        if end_dt <= start_dt:
            return redirect('/admin_block')

        existing = supabase.table("blocked_times") \
            .select("id") \
            .eq("data", data) \
            .eq("start_time", start_time) \
            .eq("end_time", end_time) \
            .execute()

        if existing.data:
            return redirect('/admin_block')

        reservations = supabase.table("reservations") \
            .select("time") \
            .eq("data", data) \
            .eq("is_deleted", False) \
            .execute()

        reserved_times = set()

        for r in (reservations.data or []):

            if r.get("time"):
                reserved_times.add(
                    r["time"][:5]
                )

        interval = get_reservation_interval()

        current = start_dt

        while current < end_dt:

            t = current.strftime("%H:%M")

            if t in reserved_times:
                return redirect('/admin_block')

            current += timedelta(
                minutes=interval
            )

        result = supabase.table("blocked_times").insert({
            "data": data,
            "start_time": start_time,
            "end_time": end_time
        }).execute()

        if not result.data:
            return redirect('/admin_block')

        return redirect('/admin_block')

    except Exception as e:

        print(
            "予約ブロック追加エラー:",
            type(e).__name__,
            str(e)
        )

        return redirect('/admin_block')


@app.route('/delete_block/<int:block_id>')
def delete_block(block_id):

    if not session.get('login'):
        return redirect('/login')

    supabase.table("blocked_times") \
        .delete() \
        .eq("id", block_id) \
        .execute()

    return redirect('/admin_block')


# =========================================================
# 新規予約
# =========================================================

@app.route('/new', methods=['GET', 'POST'])
def new():

    if not session.get('code'):
        return redirect('/')

    res = supabase.table("settings") \
        .select("*") \
        .eq("id", 1) \
        .limit(1) \
        .execute()

    setting = res.data[0] if res.data else None

    start_data = (
        setting.get("start_data")
        if setting
        else ""
    )

    end_data = (
        setting.get("end_data")
        if setting
        else ""
    )

    start_time = (
        setting.get("start_time", "09:30")
        if setting
        else "09:30"
    )

    end_time = (
        setting.get("end_time", "17:00")
        if setting
        else "17:00"
    )

    interval = (
        setting.get("interval", 30)
        if setting
        else 30
    )

    try:
        interval = int(interval)
    except Exception:
        interval = 30

    if interval not in INTERVAL_OPTIONS:
        interval = 30

    data = {
        "data": "",
        "time": "",
        "name": "",
        "phone": "",
        "address": "",
        "email": ""
    }

    if request.method == 'POST':
        data = request.form

    return render_template(
        "new.html",
        data=data,
        start_data=start_data,
        end_data=end_data,
        start_time=start_time,
        end_time=end_time,
        interval=interval
    )


# =========================================================
# 予約時間取得
# =========================================================

@app.route('/get_times')
def get_times():

    data = request.args.get('data')

    if not data:
        return jsonify([])

    setting_res = supabase.table("settings") \
        .select(
            "start_data,end_data,capacity,"
            "start_time,end_time,interval"
        ) \
        .eq("id", 1) \
        .limit(1) \
        .execute()

    setting = (
        setting_res.data[0]
        if setting_res.data
        else None
    )

    if setting:

        start_data = setting.get("start_data")
        end_data = setting.get("end_data")

        if start_data and data < start_data:
            return jsonify([])

        if end_data and data > end_data:
            return jsonify([])

    capacity = (
        setting.get("capacity", 1)
        if setting
        else 1
    )

    start_time = (
        setting.get("start_time", "09:30")
        if setting
        else "09:30"
    )

    end_time = (
        setting.get("end_time", "17:00")
        if setting
        else "17:00"
    )

    start_time = start_time[:5]
    end_time = end_time[:5]

    interval = (
        setting.get("interval", 30)
        if setting
        else 30
    )

    try:
        interval = int(interval)
    except Exception:
        interval = 30

    if interval not in INTERVAL_OPTIONS:
        interval = 30

    try:
        capacity = int(capacity)
    except Exception:
        capacity = 1

    if capacity < 1:
        capacity = 1

    # 予約取得
    res = supabase.table("reservations") \
        .select("time") \
        .eq("data", data) \
        .eq("is_deleted", False) \
        .execute()

    reserved_count = {}

    for r in (res.data or []):

        if r.get("time"):

            t = r["time"][:5]

            reserved_count[t] = (
                reserved_count.get(t, 0) + 1
            )

    # ブロック取得
    block_res = supabase.table("blocked_times") \
        .select("start_time,end_time") \
        .eq("data", data) \
        .execute()

    blocks = []

    for r in (block_res.data or []):

        blocks.append(
            (
                r["start_time"],
                r["end_time"]
            )
        )

    # 時間生成
    slots = []

    start = datetime.strptime(
        start_time,
        "%H:%M"
    )

    end = datetime.strptime(
        end_time,
        "%H:%M"
    )

    while start <= end:

        t = start.strftime("%H:%M")

        # 定員チェック
        if reserved_count.get(t, 0) >= capacity:

            start += timedelta(
                minutes=interval
            )

            continue

        # ブロックチェック
        current_dt = datetime.strptime(
            data + " " + t,
            "%Y-%m-%d %H:%M"
        ).replace(tzinfo=JST)

        blocked = False

        for b_start, b_end in blocks:

            bs = datetime.strptime(
                data + " " + b_start[:5],
                "%Y-%m-%d %H:%M"
            ).replace(tzinfo=JST)

            be = datetime.strptime(
                data + " " + b_end[:5],
                "%Y-%m-%d %H:%M"
            ).replace(tzinfo=JST)

            if be <= bs:
                be += timedelta(days=1)

            if bs <= current_dt < be:

                blocked = True
                break

        if blocked:

            start += timedelta(
                minutes=interval
            )

            continue

        slots.append(t)

        start += timedelta(
            minutes=interval
        )

    return jsonify(slots)


# =========================================================
# 新規予約確認
# =========================================================

@app.route('/confirm', methods=['POST'])
def confirm():

    data = {
        "data": request.form.get("data"),
        "time": request.form.get("time"),
        "name": request.form.get("name"),
        "phone": request.form.get("phone"),
        "address": request.form.get("address"),
        "email": request.form.get("email")
    }

    return render_template(
        "confirm.html",
        data=data
    )


# =========================================================
# 新規予約登録
# =========================================================

@app.route('/create_confirm', methods=['POST'])
def create_confirm():

    data = request.form

    code = session.get('code')

    if not code:
        return "ログイン情報なし"

    # 予約時間チェック
    target_dt = datetime.strptime(
        data['data'] + " " + data['time'][:5],
        "%Y-%m-%d %H:%M"
    ).replace(tzinfo=JST)

    limit_time = (
        datetime.now(JST)
        + timedelta(hours=24)
    )

    if target_dt < limit_time:
        return "24時間後以降の予約しかできません"

    # 重複チェック
    check = supabase.table("reservations") \
        .select("id") \
        .eq("data", data.get("data")) \
        .eq("time", data.get("time")[:5]) \
        .eq("is_deleted", False) \
        .execute()

    if check.data:
        return "この時間は予約できません。"

    # Supabase保存
    try:

        supabase.table("reservations").insert({
            "data": data.get("data", ""),
            "time": data.get("time", "")[:5],
            "consumer_code": code,
            "name": data.get("name", ""),
            "phone": data.get("phone", ""),
            "address": data.get("address", ""),
            "email": data.get("email", ""),
            "status": "新規",
            "is_deleted": False,
            "is_confirmed": False,
            "created_at": datetime.now(JST).isoformat()
        }).execute()

    except Exception as e:

        print(
            "INSERT ERROR:",
            type(e).__name__,
            str(e)
        )

        return "この時間はすでに予約されています"

    # 管理者メール
    mail_new(
        data.get('data', ''),
        data.get('time', '')[:5],
        data.get('name', ''),
        data.get('phone', '')
    )

    return render_template(
        "complete.html",
        data=data
    )


# =========================================================
# 日付が予約可能か確認
# =========================================================

@app.route('/check_day')
def check_day():

    data = request.args.get('data')

    if not data:
        return jsonify({
            "ok": False
        })

    setting_res = supabase.table("settings") \
        .select(
            "start_data,end_data,"
            "start_time,end_time,interval"
        ) \
        .eq("id", 1) \
        .limit(1) \
        .execute()

    setting = (
        setting_res.data[0]
        if setting_res.data
        else None
    )

    if not setting:

        return jsonify({
            "ok": False,
            "message": "予約設定がありません"
        })

    start_data = setting.get("start_data")
    end_data = setting.get("end_data")

    if start_data and data < start_data:

        return jsonify({
            "ok": False,
            "message": "この日は予約できません"
        })

    if end_data and data > end_data:

        return jsonify({
            "ok": False,
            "message": "この日は予約できません"
        })

    start_time = setting.get(
        "start_time",
        "09:30"
    )

    end_time = setting.get(
        "end_time",
        "17:00"
    )

    start_time = start_time[:5]
    end_time = end_time[:5]

    try:
        interval = int(
            setting.get("interval", 30)
        )
    except Exception:
        interval = 30

    if interval not in INTERVAL_OPTIONS:
        interval = 30

    # ブロック取得
    res = supabase.table("blocked_times") \
        .select("start_time,end_time") \
        .eq("data", data) \
        .execute()

    blocks = res.data or []

    # 予約枠生成
    slots = []

    start = datetime.strptime(
        start_time,
        "%H:%M"
    )

    end = datetime.strptime(
        end_time,
        "%H:%M"
    )

    while start <= end:

        slots.append(
            start.strftime("%H:%M")
        )

        start += timedelta(
            minutes=interval
        )

    # ブロック削除
    for b in blocks:

        b_start = b.get("start_time")
        b_end = b.get("end_time")

        if not b_start or not b_end:
            continue

        bs = datetime.strptime(
            b_start[:5],
            "%H:%M"
        )

        be = datetime.strptime(
            b_end[:5],
            "%H:%M"
        )

        slots = [
            t
            for t in slots
            if not (
                bs <= datetime.strptime(
                    t,
                    "%H:%M"
                ) < be
            )
        ]

    return jsonify({
        "ok": len(slots) > 0,
        "message":
            "この日は予約できません"
            if len(slots) == 0
            else ""
    })


# =========================================================
# 予約変更画面
# =========================================================

@app.route('/edit')
def edit():

    code = session.get('code')

    if not code:
        return redirect('/')

    res = supabase.table("reservations") \
        .select("*") \
        .eq("consumer_code", code) \
        .eq("is_deleted", False) \
        .order("id", desc=True) \
        .limit(1) \
        .execute()

    data = (
        res.data[0]
        if res.data
        else None
    )

    if not data:
        return render_template(
            "no_reservation.html"
        )

    res = supabase.table("settings") \
        .select("*") \
        .eq("id", 1) \
        .limit(1) \
        .execute()

    setting = (
        res.data[0]
        if res.data
        else None
    )

    start_data = (
        setting.get("start_data")
        if setting
        else ""
    )

    end_data = (
        setting.get("end_data")
        if setting
        else ""
    )

    return render_template(
        "edit.html",
        data=(
            data["id"],
            data["data"],
            data["time"],
            data["name"],
            data["phone"],
            data["address"],
            data.get("email", "")
        ),
        start_data=start_data,
        end_data=end_data
    )


@app.route('/edit_confirm', methods=['POST'])
def edit_confirm():

    data = {
        "data": request.form.get("data"),
        "time": request.form.get("time"),
        "name": request.form.get("name"),
        "phone": request.form.get("phone"),
        "address": request.form.get("address"),
        "email": request.form.get("email")
    }

    return render_template(
        "edit_confirm.html",
        data=data
    )


# =========================================================
# 予約削除画面
# =========================================================

@app.route('/delete')
def delete():

    code = session.get('code')

    if not code:
        return redirect('/')

    res = supabase.table("reservations") \
        .select("*") \
        .eq("consumer_code", code) \
        .eq("is_deleted", False) \
        .order("id", desc=True) \
        .limit(1) \
        .execute()

    row = (
        res.data[0]
        if res.data
        else None
    )

    if not row:
        return render_template(
            "delete.html",
            data=None
        )

    time_range = format_time_range(
        row["time"]
    )

    data = (
        row["id"],
        row["data"],
        time_range,
        row["name"],
        row["phone"],
        row["address"],
        row.get("email", "")
    )

    return render_template(
        "delete.html",
        data=data
    )


# =========================================================
# 予約削除実行
# =========================================================

@app.route('/delete_confirm', methods=['POST'])
def delete_confirm():

    code = session.get('code')

    if not code:
        return redirect('/')

    res = supabase.table("reservations") \
        .select("*") \
        .eq("consumer_code", code) \
        .eq("is_deleted", False) \
        .order("created_at", desc=True) \
        .limit(1) \
        .execute()

    if not res.data:
        return redirect('/')

    r = res.data[0]

    # 削除履歴を追加
    supabase.table("reservations").insert({
        "data": r["data"],
        "time": r["time"],
        "consumer_code": r["consumer_code"],
        "name": r["name"],
        "phone": r["phone"],
        "address": r["address"],
        "email": r.get("email"),
        "status": "削除",
        "is_deleted": False,
        "is_confirmed": False,
        "created_at": datetime.now(JST).isoformat()
    }).execute()

    mail_delete(
        r["data"],
        r["time"][:5],
        r["name"],
        r["phone"]
    )

    return render_template(
        "delete_done.html"
    )


# =========================================================
# 予約内容確認
# =========================================================

@app.route('/view')
def view():

    code = session.get('code')

    if not code:
        return redirect('/')

    res = supabase.table("reservations") \
        .select("*") \
        .eq("consumer_code", code) \
        .eq("is_deleted", False) \
        .order("id", desc=True) \
        .limit(1) \
        .execute()

    row = (
        res.data[0]
        if res.data
        else None
    )

    if not row:
        return render_template(
            "view.html",
            data=None
        )

    time_range = format_time_range(
        row.get("time", "")
    )

    data = (
        row["id"],
        row["data"],
        time_range,
        row["name"],
        row["phone"],
        row["address"],
        row.get("email"),
        row.get("is_confirmed")
    )

    return render_template(
        "view.html",
        data=data
    )


# =========================================================
# 予約Excel出力
# =========================================================

@app.route('/export_excel')
def export_excel():

    if not session.get('login'):
        return redirect('/login')

    res = supabase.table("reservations") \
        .select("*") \
        .eq("is_deleted", False) \
        .order("created_at", desc=True) \
        .execute()

    rows = res.data or []

    wb = Workbook()
    ws = wb.active
    ws.title = "予約一覧"

    ws.append([
        "予約日",
        "時間",
        "申込日時",
        "コード",
        "氏名",
        "住所",
        "電話",
        "メール",
        "状態"
    ])

    for r in rows:

        time = (
            r.get("time") or ""
        )[:5]

        created = (
            r.get("created_at") or ""
        )

        if created:
            created = created[:19].replace(
                "T",
                " "
            )

        ws.append([
            r.get("data"),
            time,
            created,
            r.get("consumer_code"),
            r.get("name"),
            r.get("address"),
            r.get("phone"),
            r.get("email"),
            r.get("status") or "new"
        ])

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return send_file(
        output,
        download_name="reservations.xlsx",
        as_attachment=True
    )


# =========================================================
# 削除予約Excel出力
# =========================================================

@app.route('/export_deleted_excel')
def export_deleted_excel():

    if not session.get('login'):
        return redirect('/login')

    res = supabase.table("reservations") \
        .select("*") \
        .eq("is_deleted", True) \
        .order("created_at", desc=True) \
        .execute()

    rows = res.data or []

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

    for r in rows:

        time = (
            r.get("time") or ""
        )[:5]

        created = (
            r.get("created_at") or ""
        )

        if created:
            created = created[:19].replace(
                "T",
                " "
            )

        ws.append([
            r.get("data"),
            time,
            created,
            r.get("consumer_code"),
            r.get("name"),
            r.get("address"),
            r.get("phone"),
            r.get("status") or "new"
        ])

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return send_file(
        output,
        download_name="deleted_reservations.xlsx",
        as_attachment=True
    )


# =========================================================
# 起動
# =========================================================

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=10000,
        debug=True
    )