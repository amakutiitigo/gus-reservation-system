import os
import io
import json
import urllib.request
import urllib.error

from datetime import datetime, timedelta, timezone, time

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    jsonify,
    send_file
)

from dotenv import load_dotenv
from supabase import create_client
from openpyxl import Workbook


# =========================================================
# 環境変数
# =========================================================

load_dotenv()


# =========================================================
# Flask
# =========================================================

app = Flask(__name__)

app.secret_key = os.getenv("SECRET_KEY")

app.permanent_session_lifetime = timedelta(days=7)


# =========================================================
# 基本設定
# =========================================================

INTERVAL_OPTIONS = [10, 15, 30, 60]

JST = timezone(timedelta(hours=9))


# =========================================================
# Supabase
# =========================================================

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

supabase = create_client(
    SUPABASE_URL,
    SUPABASE_KEY
)


# =========================================================
# Brevo
# =========================================================

BREVO_API_KEY = os.getenv("BREVO_API_KEY")

MAIL_FROM = os.getenv("MAIL_FROM")
MAIL_FROM_NAME = os.getenv("MAIL_FROM_NAME", "")

ADMIN_EMAIL = os.getenv("ADMIN_EMAIL")


# =========================================================
# 予約時間間隔取得
# =========================================================

def get_reservation_interval():

    try:
        result = (
            supabase
            .table("settings")
            .select("interval")
            .eq("id", 1)
            .single()
            .execute()
        )

        if result.data and result.data.get("interval"):
            return int(result.data["interval"])

    except Exception as e:
        print("予約時間間隔取得エラー:", e)

    return 30


# =========================================================
# 時間表示
# =========================================================

def format_time_range(time_str):

    try:

        interval = get_reservation_interval()

        start = datetime.strptime(
            time_str,
            "%H:%M"
        )

        end = start + timedelta(minutes=interval)

        return (
            start.strftime("%H:%M")
            + "～"
            + end.strftime("%H:%M")
        )

    except Exception as e:

        print("時間表示変換エラー:", e)

        return time_str


# =========================================================
# Brevoメール送信
# =========================================================

def send_brevo_mail(
    to_email,
    subject,
    body
):

    if not BREVO_API_KEY:

        print("BREVO_API_KEY が設定されていません。")

        return False

    if not MAIL_FROM:

        print("MAIL_FROM が設定されていません。")

        return False

    if not to_email:

        print("送信先メールアドレスがありません。")

        return False


    # -----------------------------------------------------
    # sender
    # -----------------------------------------------------

    sender = {
        "email": MAIL_FROM
    }

    if MAIL_FROM_NAME:
        sender["name"] = MAIL_FROM_NAME


    # -----------------------------------------------------
    # Brevo API リクエスト
    # -----------------------------------------------------

    payload = {

        "sender": sender,

        "to": [
            {
                "email": to_email
            }
        ],

        "subject": subject,

        "textContent": body
    }


    data = json.dumps(
        payload,
        ensure_ascii=False
    ).encode("utf-8")


    req = urllib.request.Request(

        "https://api.brevo.com/v3/smtp/email",

        data=data,

        headers={
            "accept": "application/json",
            "api-key": BREVO_API_KEY,
            "content-type": "application/json"
        },

        method="POST"
    )


    try:

        with urllib.request.urlopen(
            req,
            timeout=30
        ) as response:

            response_body = response.read().decode(
                "utf-8"
            )

            print(
                "Brevoメール送信成功:",
                response.status,
                response_body
            )

            return True


    except urllib.error.HTTPError as e:

        error_body = ""

        try:
            error_body = e.read().decode("utf-8")

        except Exception:
            pass

        print(
            "Brevoメール送信HTTPエラー:",
            e.code,
            error_body
        )

        return False


    except Exception as e:

        print(
            "Brevoメール送信エラー:",
            e
        )

        return False


# =========================================================
# 管理者メール
# =========================================================

def send_admin_mail(
    subject,
    body
):

    if not ADMIN_EMAIL:

        print("ADMIN_EMAIL が設定されていません。")

        return False

    return send_brevo_mail(
        ADMIN_EMAIL,
        subject,
        body
    )


# =========================================================
# 新規予約メール
# =========================================================

def mail_new(row):

    data = row.get("data", "")
    time_str = row.get("time", "")
    name = row.get("name", "")
    phone = row.get("phone", "")
    address = row.get("address", "")
    email = row.get("email", "")
    consumer_code = row.get("consumer_code", "")


    time_display = format_time_range(
        time_str
    )


    subject = "新しい予約が入りました"


    body = f"""
新しい予約が入りました。

【予約情報】

お客様コード：
{consumer_code}

氏名：
{name}

住所：
{address}

電話番号：
{phone}

メールアドレス：
{email}

予約日：
{data}

予約時間：
{time_display}
"""


    send_admin_mail(
        subject,
        body
    )


# =========================================================
# 変更予約メール
# =========================================================

def mail_edit(row):

    data = row.get("data", "")
    time_str = row.get("time", "")
    name = row.get("name", "")
    phone = row.get("phone", "")
    address = row.get("address", "")
    email = row.get("email", "")
    consumer_code = row.get("consumer_code", "")


    time_display = format_time_range(
        time_str
    )


    subject = "予約が変更されました"


    body = f"""
予約が変更されました。

【予約情報】

お客様コード：
{consumer_code}

氏名：
{name}

住所：
{address}

電話番号：
{phone}

メールアドレス：
{email}

予約日：
{data}

予約時間：
{time_display}
"""


    send_admin_mail(
        subject,
        body
    )


# =========================================================
# 削除予約メール
# =========================================================

def mail_delete(row):

    data = row.get("data", "")
    time_str = row.get("time", "")
    name = row.get("name", "")
    phone = row.get("phone", "")
    address = row.get("address", "")
    email = row.get("email", "")
    consumer_code = row.get("consumer_code", "")


    time_display = format_time_range(
        time_str
    )


    subject = "予約がキャンセルされました"


    body = f"""
予約がキャンセルされました。

【予約情報】

お客様コード：
{consumer_code}

氏名：
{name}

住所：
{address}

電話番号：
{phone}

メールアドレス：
{email}

予約日：
{data}

予約時間：
{time_display}
"""


    send_admin_mail(
        subject,
        body
    )


# =========================================================
# 利用者への予約確定メール
# =========================================================

def send_mail(row):

    email = row.get("email", "")

    if not email:
        print(
            "予約者のメールアドレスがありません。"
        )
        return False


    # 予約内容
    data = row.get("data", "")
    time_str = row.get("time", "")
    name = row.get("name", "")
    phone = row.get("phone", "")
    address = row.get("address", "")
    consumer_code = row.get("consumer_code", "")


    # 予約時間を「開始～終了」にする
    time_display = format_time_range(
        time_str
    )


    subject = "予約確定のお知らせ"


    body = f"""
{name} 様

ガス点検の予約が確定しました。

【予約内容】

お客様コード：
{consumer_code}

お名前：
{name}

予約日：
{data}

予約時間：
{time_display}

電話番号：
{phone}

住所：
{address}

メールアドレス：
{email}


予約内容に変更・キャンセルがある場合は、
予約システムからお手続きをお願いします。

よろしくお願いいたします。
"""


    return send_brevo_mail(
        email,
        subject,
        body
    )


# =========================================================
# トップ
# =========================================================

@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        consumer_code = (
            request.form.get("consumer_code")
            or ""
        ).strip()

        action = (
            request.form.get("action")
            or ""
        ).strip()


        session["code"] = consumer_code

        if action == "新規":

            return redirect(
                url_for("new")
            )

        elif action == "変更":

            return redirect(
                url_for("edit")
            )

        elif action == "削除":

            return redirect(
                url_for("delete")
            )

        elif action == "確認":

            return redirect(
                url_for("view")
            )

        elif action == "手続きをやめる":

            session.pop(
                "code",
                None
            )

            return redirect(
                url_for("index")
            )


    return render_template(
        "index.html"
    )


# =========================================================
# 管理者ログイン
# =========================================================

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        password = (
            request.form.get("password")
            or ""
        )

        admin_password = (
            os.getenv("ADMIN_PASSWORD")
            or ""
        )


        if password == admin_password:

            session["admin_logged_in"] = True

            return redirect(
                url_for("admin_menu")
            )


        return render_template(
            "login.html",
            error="パスワードが違います。"
        )


    return render_template(
        "login.html"
    )


# =========================================================
# 管理者ログアウト
# =========================================================

@app.route("/logout")
def logout():

    session.pop(
        "admin_logged_in",
        None
    )

    return redirect(
        url_for("login")
    )


# =========================================================
# 管理者メニュー
# =========================================================

@app.route("/admin_menu")
def admin_menu():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    blocked = (
        supabase
        .table("blocked_times")
        .select("*")
        .order("data")
        .order("start_time")
        .execute()
    )


    setting = (
        supabase
        .table("settings")
        .select(
            "start_data,end_data,capacity"
        )
        .eq("id", 1)
        .single()
        .execute()
    )


    setting_data = setting.data or {}


    return render_template(
        "admin_menu.html",
        blocked_times=blocked.data or [],
        start_data=setting_data.get(
            "start_data"
        ),
        end_data=setting_data.get(
            "end_data"
        ),
        capacity=setting_data.get(
            "capacity"
        )
    )


# =========================================================
# 管理者設定
# =========================================================

@app.route("/admin_setting")
def admin_setting():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    result = (
        supabase
        .table("settings")
        .select(
            "start_data,"
            "end_data,"
            "capacity,"
            "start_time,"
            "end_time,"
            "interval"
        )
        .eq("id", 1)
        .single()
        .execute()
    )


    setting = result.data or {}


    return render_template(
        "admin_setting.html",
        setting=setting,
        interval_options=INTERVAL_OPTIONS
    )


# =========================================================
# 設定保存
# =========================================================

@app.route(
    "/save_setting",
    methods=["POST"]
)
def save_setting():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    start_data = (
        request.form.get("start_data")
        or ""
    ).strip()

    end_data = (
        request.form.get("end_data")
        or ""
    ).strip()

    capacity_str = (
        request.form.get("capacity")
        or ""
    ).strip()

    start_time = (
        request.form.get("start_time")
        or ""
    ).strip()

    end_time = (
        request.form.get("end_time")
        or ""
    ).strip()

    interval_str = (
        request.form.get("interval")
        or ""
    ).strip()


    try:

        capacity = int(
            capacity_str
        )

        interval = int(
            interval_str
        )


    except ValueError:

        return "設定値が不正です。", 400


    if capacity < 1:

        return "定員は1人以上にしてください。", 400


    if interval not in INTERVAL_OPTIONS:

        return "時間間隔が不正です。", 400


    if not start_data:
        start_data = None

    if not end_data:
        end_data = None

    if not start_time:
        start_time = None

    if not end_time:
        end_time = None


    supabase.table(
        "settings"
    ).upsert(
        {
            "id": 1,
            "start_data": start_data,
            "end_data": end_data,
            "capacity": capacity,
            "start_time": start_time,
            "end_time": end_time,
            "interval": interval
        }
    ).execute()


    return redirect(
        url_for("admin_setting")
    )


# =========================================================
# 予約期間クリア
# =========================================================

@app.route("/clear_setting", methods=["POST"])
def clear_setting():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    supabase.table(
        "settings"
    ).upsert(
        {
            "id": 1,
            "start_data": None,
            "end_data": None
        }
    ).execute()


    return redirect(
        url_for("admin_setting")
    )


# =========================================================
# 管理者予約一覧
# =========================================================

@app.route("/admin")
def admin():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    code = (
        request.args.get("code")
        or ""
    ).strip()

    name = (
        request.args.get("name")
        or ""
    ).strip()

    confirmed = (
        request.args.get("confirmed")
        or ""
    ).strip()

    start_date = (
        request.args.get("start_date")
        or ""
    ).strip()

    end_date = (
        request.args.get("end_date")
        or ""
    ).strip()

    created_start = (
        request.args.get("created_start")
        or ""
    ).strip()

    created_end = (
        request.args.get("created_end")
        or ""
    ).strip()


    query = (
        supabase
        .table("reservations")
        .select("*")
        .eq("is_deleted", False)
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


    if confirmed == "yes":

        query = query.eq(
            "is_confirmed",
            True
        )

    elif confirmed == "no":

        query = query.eq(
            "is_confirmed",
            False
        )


    if start_date:

        query = query.gte(
            "data",
            start_date
        )


    if end_date:

        query = query.lte(
            "data",
            end_date
        )


    if created_start:

        query = query.gte(
            "created_at",
            created_start + "T00:00:00"
        )


    if created_end:

        query = query.lte(
            "created_at",
            created_end + "T23:59:59"
        )


    result = (
        query
        .order("consumer_code")
        .order(
            "created_at",
            desc=True
        )
        .execute()
    )


    reservations = result.data or []


    for row in reservations:

        row["time_display"] = format_time_range(
            row.get("time", "")
        )


    return render_template(
        "admin.html",
        reservations=reservations
    )


# =========================================================
# 確定状態切り替え
# =========================================================

@app.route(
    "/toggle_confirm",
    methods=["POST"]
)
def toggle_confirm():

    if not session.get("admin_logged_in"):
        return redirect(
            url_for("login")
        )

    reservation_id = request.form.get("id")

    if not reservation_id or not reservation_id.isdigit():
        return redirect(
            request.referrer
            or url_for("admin")
        )

    reservation_id = int(reservation_id)

    result = (
        supabase
        .table("reservations")
        .select("is_confirmed")
        .eq("id", reservation_id)
        .single()
        .execute()
    )

    current = False

    if result.data:
        current = bool(
            result.data.get(
                "is_confirmed"
            )
        )

    new_value = not current

    supabase.table(
        "reservations"
    ).update(
        {
            "is_confirmed": new_value
        }
    ).eq(
        "id",
        reservation_id
    ).execute()

    # 「確定」に変更したときだけメール送信
    if new_value:

        row_result = (
            supabase
            .table("reservations")
            .select("*")
            .eq("id", reservation_id)
            .single()
            .execute()
        )

        row = row_result.data

        if row and row.get("email"):
            send_mail(row)

    return redirect(
        request.referrer
        or url_for("admin")
    )


# =========================================================
# 管理者削除
# =========================================================

@app.route(
    "/admin_delete/<int:id>",
    methods=["POST"]
)
def admin_delete(id):

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    supabase.table(
        "reservations"
    ).update(
        {
            "is_deleted": True
        }
    ).eq(
        "id",
        id
    ).execute()


    return redirect(
        request.referrer
        or url_for("admin")
    )


# =========================================================
# 管理者一括削除
# =========================================================

@app.route(
    "/admin_delete_selected",
    methods=["POST"]
)
def admin_delete_selected():

    if not session.get("admin_logged_in"):
        return redirect(
            url_for("login")
        )

    ids = request.form.getlist("ids")

    if not ids:
        return redirect(
            request.referrer
            or url_for("admin")
        )

    for reservation_id in ids:

        try:
            reservation_id = int(reservation_id)

            supabase.table(
                "reservations"
            ).update(
                {
                    "is_deleted": True
                }
            ).eq(
                "id",
                reservation_id
            ).execute()

        except (ValueError, TypeError):
            continue

    return redirect(
        request.referrer
        or url_for("admin")
    )

# =========================================================
# 管理者編集
# =========================================================

@app.route("/admin_edit/<int:id>")
def admin_edit(id):

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    result = (
        supabase
        .table("reservations")
        .select("*")
        .eq("id", id)
        .single()
        .execute()
    )


    row = result.data


    if not row:

        return "予約が見つかりません。", 404


    return render_template(
        "admin_edit.html",
        reservation=row
    )


# =========================================================
# 管理者編集保存
# =========================================================

@app.route(
    "/admin_edit_save",
    methods=["POST"]
)
def admin_edit_save():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    reservation_id = request.form.get(
        "id"
    )

    data = (
        request.form.get("data")
        or ""
    ).strip()

    time_str = (
        request.form.get("time")
        or ""
    ).strip()

    name = (
        request.form.get("name")
        or ""
    ).strip()

    phone = (
        request.form.get("phone")
        or ""
    ).strip()

    address = (
        request.form.get("address")
        or ""
    ).strip()

    email = (
        request.form.get("email")
        or ""
    ).strip()


    if not data or not time_str:

        return "予約日と時間は必須です。", 400


    duplicate = (
        supabase
        .table("reservations")
        .select("id")
        .eq("data", data)
        .eq("time", time_str)
        .eq("is_deleted", False)
        .neq("id", reservation_id)
        .execute()
    )


    if duplicate.data:

        return "その日時にはすでに予約があります。", 400


    supabase.table(
        "reservations"
    ).update(
        {
            "data": data,
            "time": time_str,
            "name": name,
            "phone": phone,
            "address": address,
            "email": email,
            "is_confirmed": False
        }
    ).eq(
        "id",
        reservation_id
    ).execute()


    return redirect(
        url_for("admin")
    )


# =========================================================
# ユーザー予約変更
# =========================================================

@app.route(
    "/edit_save",
    methods=["POST"]
)
def edit_save():

    consumer_code = session.get(
        "code"
    )


    if not consumer_code:

        return redirect(
            url_for("index")
        )


    data = (
        request.form.get("data")
        or ""
    ).strip()

    time_str = (
        request.form.get("time")
        or ""
    ).strip()

    name = (
        request.form.get("name")
        or ""
    ).strip()

    phone = (
        request.form.get("phone")
        or ""
    ).strip()

    address = (
        request.form.get("address")
        or ""
    ).strip()

    email = (
        request.form.get("email")
        or ""
    ).strip()


    duplicate = (
        supabase
        .table("reservations")
        .select("id")
        .eq("data", data)
        .eq("time", time_str)
        .eq("is_deleted", False)
        .neq(
            "consumer_code",
            consumer_code
        )
        .execute()
    )


    if duplicate.data:

        return "その日時にはすでに予約があります。", 400


    new_row = {

        "data": data,

        "time": time_str,

        "consumer_code": consumer_code,

        "name": name,

        "phone": phone,

        "address": address,

        "email": email,

        "status": "変更",

        "is_deleted": False,

        "is_confirmed": False,

        "created_at": datetime.now(
            JST
        ).isoformat()
    }


    result = (
        supabase
        .table("reservations")
        .insert(new_row)
        .execute()
    )


    if result.data:

        mail_edit(
            result.data[0]
        )


    return render_template(
        "edit_complete.html",
        data=new_row
    )


# =========================================================
# 削除済み一覧
# =========================================================

@app.route("/admin_deleted")
def admin_deleted():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    result = (
        supabase
        .table("reservations")
        .select("*")
        .eq("is_deleted", True)
        .order(
            "created_at",
            desc=True
        )
        .execute()
    )


    reservations = result.data or []


    for row in reservations:

        row["time_display"] = format_time_range(
            row.get("time", "")
        )


    return render_template(
        "admin_deleted.html",
        reservations=reservations
    )


# =========================================================
# 復元
# =========================================================

@app.route(
    "/admin_restore/<int:id>",
    methods=["POST"]
)
def admin_restore(id):

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    supabase.table(
        "reservations"
    ).update(
        {
            "is_deleted": False
        }
    ).eq(
        "id",
        id
    ).execute()


    return redirect(
        request.referrer
        or url_for("admin_deleted")
    )


# =========================================================
# 複数復元
# =========================================================

@app.route(
    "/admin_restore_multi",
    methods=["POST"]
)
def admin_restore_multi():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    ids = request.form.getlist(
        "ids"
    )


    for reservation_id in ids:

        supabase.table(
            "reservations"
        ).update(
            {
                "is_deleted": False
            }
        ).eq(
            "id",
            reservation_id
        ).execute()


    return redirect(
        url_for("admin_deleted")
    )


# =========================================================
# 完全削除
# =========================================================

@app.route(
    "/admin_bulk_delete",
    methods=["POST"]
)
def admin_bulk_delete():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    ids = request.form.getlist(
        "ids"
    )


    for reservation_id in ids:

        supabase.table(
            "reservations"
        ).delete().eq(
            "id",
            reservation_id
        ).execute()


    return redirect(
        url_for("admin_deleted")
    )


# =========================================================
# ブロック時間管理
# =========================================================

@app.route("/admin_block")
def admin_block():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    result = (
        supabase
        .table("blocked_times")
        .select("*")
        .order("data")
        .order("start_time")
        .execute()
    )


    return render_template(
        "admin_block.html",
        blocked_times=result.data or []
    )


# =========================================================
# ブロック時間Excel
# =========================================================

@app.route("/export_block_excel")
def export_block_excel():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    result = (
        supabase
        .table("blocked_times")
        .select("*")
        .order("data")
        .order("start_time")
        .execute()
    )


    rows = result.data or []


    wb = Workbook()

    ws = wb.active

    ws.title = "予約不可時間"


    ws.append(
        [
            "日付",
            "開始時間",
            "終了時間"
        ]
    )


    for row in rows:

        ws.append(
            [
                row.get("data", ""),
                row.get("start_time", ""),
                row.get("end_time", "")
            ]
        )


    output = io.BytesIO()

    wb.save(output)

    output.seek(0)


    return send_file(
        output,
        as_attachment=True,
        download_name="予約不可時間.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# =========================================================
# ブロック追加
# =========================================================

@app.route(
    "/add_block",
    methods=["POST"]
)
def add_block():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    data = (
        request.form.get("data")
        or ""
    ).strip()

    start_time = (
        request.form.get("start_time")
        or ""
    ).strip()

    end_time = (
        request.form.get("end_time")
        or ""
    ).strip()


    if not data or not start_time or not end_time:

        return "入力が不足しています。", 400


    if start_time >= end_time:

        return "終了時間は開始時間より後にしてください。", 400


    # -----------------------------------------------------
    # 同一ブロック確認
    # -----------------------------------------------------

    duplicate = (
        supabase
        .table("blocked_times")
        .select("id")
        .eq("data", data)
        .eq("start_time", start_time)
        .eq("end_time", end_time)
        .execute()
    )


    if duplicate.data:

        return "同じ予約不可時間がすでに登録されています。", 400


    # -----------------------------------------------------
    # 既存予約確認
    # -----------------------------------------------------

    reservations = (
        supabase
        .table("reservations")
        .select("time")
        .eq("data", data)
        .eq("is_deleted", False)
        .execute()
    )


    interval = get_reservation_interval()


    start_dt = datetime.strptime(
        start_time,
        "%H:%M"
    )

    end_dt = datetime.strptime(
        end_time,
        "%H:%M"
    )


    for reservation in reservations.data or []:

        reservation_time = reservation.get(
            "time"
        )

        if not reservation_time:
            continue


        reservation_dt = datetime.strptime(
            reservation_time,
            "%H:%M"
        )


        slot_end = (
            reservation_dt
            + timedelta(minutes=interval)
        )


        if (
            reservation_dt < end_dt
            and slot_end > start_dt
        ):

            return (
                "その時間帯には既に予約があります。"
            ), 400


    # -----------------------------------------------------
    # ブロック登録
    # -----------------------------------------------------

    supabase.table(
        "blocked_times"
    ).insert(
        {
            "data": data,
            "start_time": start_time,
            "end_time": end_time
        }
    ).execute()


    return redirect(
        url_for("admin_block")
    )


@app.route("/bulk_add_block", methods=["POST"])
def bulk_add_block():
    if not session.get("admin_logged_in"):
        return redirect(url_for("login"))

    start_date_str = (request.form.get("start_date") or "").strip()
    end_date_str = (request.form.get("end_date") or "").strip()

    if not start_date_str or not end_date_str:
        return "開始日と終了日を入力してください。", 400

    try:
        start_date = datetime.strptime(
            start_date_str, "%Y-%m-%d"
        ).date()

        end_date = datetime.strptime(
            end_date_str, "%Y-%m-%d"
        ).date()

    except ValueError:
        return "日付の形式が正しくありません。", 400

    if start_date > end_date:
        return "終了日は開始日以降にしてください。", 400

    # ==========================================
    # 対象曜日
    # ==========================================

    weekdays = request.form.getlist("weekdays")

    weekday_numbers = set()

    for value in weekdays:
        try:
            weekday_numbers.add(int(value))
        except (ValueError, TypeError):
            continue

    target_dates = []

    current_date = start_date

    while current_date <= end_date:

        if (
            not weekday_numbers
            or current_date.weekday() in weekday_numbers
        ):
            target_dates.append(
                current_date.strftime("%Y-%m-%d")
            )

        current_date += timedelta(days=1)

    # ==========================================
    # 個別日付
    # ==========================================

    individual_dates = request.form.getlist(
        "individual_dates"
    )

    for date_str in individual_dates:

        date_str = (date_str or "").strip()

        if not date_str:
            continue

        try:
            datetime.strptime(
                date_str, "%Y-%m-%d"
            )

            if date_str not in target_dates:
                target_dates.append(date_str)

        except ValueError:
            continue

    target_dates = sorted(set(target_dates))

    if not target_dates:
        return "登録対象の日付がありません。", 400

    # ==========================================
    # 時間帯
    # ==========================================

    start_times = request.form.getlist(
        "start_times"
    )

    end_times = request.form.getlist(
        "end_times"
    )

    time_ranges = []

    for start_time, end_time in zip(
        start_times,
        end_times
    ):

        start_time = (start_time or "").strip()
        end_time = (end_time or "").strip()

        if not start_time or not end_time:
            continue

        try:
            start_dt = datetime.strptime(
                start_time,
                "%H:%M"
            )

            end_dt = datetime.strptime(
                end_time,
                "%H:%M"
            )

        except ValueError:
            continue

        if start_dt >= end_dt:
            continue

        pair = (
            start_time,
            end_time
        )

        if pair not in time_ranges:
            time_ranges.append(pair)

    if not time_ranges:
        return "登録する時間帯がありません。", 400

    # ==========================================
    # 登録処理
    # ==========================================

    registered_count = 0
    duplicate_count = 0
    reservation_conflict_count = 0

    conflicts = []

    interval = get_reservation_interval()

    for data in target_dates:

        # --------------------------------------
        # その日の予約を取得
        # --------------------------------------

        reservation_result = (
            supabase
            .table("reservations")
            .select("time,name")
            .eq("data", data)
            .eq("is_deleted", False)
            .execute()
        )

        reservations = (
            reservation_result.data or []
        )

        # --------------------------------------
        # その日の既存ブロックを取得
        # --------------------------------------

        block_result = (
            supabase
            .table("blocked_times")
            .select("*")
            .eq("data", data)
            .execute()
        )

        existing_blocks = (
            block_result.data or []
        )

        # ======================================
        # 新しい時間帯を1つずつ処理
        # ======================================

        for start_time, end_time in time_ranges:

            new_start = datetime.strptime(
                start_time,
                "%H:%M"
            )

            new_end = datetime.strptime(
                end_time,
                "%H:%M"
            )

            # ==================================
            # 予約との重複チェック
            # ==================================

            reservation_exists = False

            for reservation in reservations:

                reservation_time = (
                    reservation.get("time")
                )

                if not reservation_time:
                    continue

                try:
                    reservation_dt = (
                        datetime.strptime(
                            reservation_time,
                            "%H:%M"
                        )
                    )

                except ValueError:
                    continue

                reservation_end_dt = (
                    reservation_dt
                    + timedelta(
                        minutes=interval
                    )
                )

                if (
                    reservation_dt < new_end
                    and reservation_end_dt > new_start
                ):

                    reservation_exists = True

                    reservation_name = (
                        reservation.get("name")
                        or ""
                    )

                    conflicts.append(
                        f"{data} "
                        f"{start_time}～{end_time}："
                        f"既に予約があります"
                        f"（{reservation_name}）"
                    )

                    break

            if reservation_exists:

                reservation_conflict_count += 1

                continue

            # ==================================
            # 既存ブロックと重なっているか確認
            # ==================================

            overlapping_blocks = []

            for block in existing_blocks:

                block_start_str = (
                    block.get("start_time")
                    or ""
                )

                block_end_str = (
                    block.get("end_time")
                    or ""
                )

                if (
                    not block_start_str
                    or not block_end_str
                ):
                    continue

                try:
                    block_start = (
                        datetime.strptime(
                            block_start_str,
                            "%H:%M"
                        )
                    )

                    block_end = (
                        datetime.strptime(
                            block_end_str,
                            "%H:%M"
                        )
                    )

                except ValueError:
                    continue

                # --------------------------------
                # 重なっているか
                # --------------------------------

                if (
                    new_start < block_end
                    and new_end > block_start
                ):

                    overlapping_blocks.append(
                        block
                    )

            # ==================================
            # 重なっているブロックがない
            # ==================================

            if not overlapping_blocks:

                supabase.table(
                    "blocked_times"
                ).insert(
                    {
                        "data": data,
                        "start_time": start_time,
                        "end_time": end_time
                    }
                ).execute()

                registered_count += 1

                existing_blocks.append(
                    {
                        "data": data,
                        "start_time": start_time,
                        "end_time": end_time
                    }
                )

                continue

            # ==================================
            # 重なっているブロックがある
            #
            # → 全部まとめて1つにする
            # ==================================

            merged_start = new_start
            merged_end = new_end

            for block in overlapping_blocks:

                block_start = datetime.strptime(
                    block["start_time"],
                    "%H:%M"
                )

                block_end = datetime.strptime(
                    block["end_time"],
                    "%H:%M"
                )

                if block_start < merged_start:
                    merged_start = block_start

                if block_end > merged_end:
                    merged_end = block_end

            # ==================================
            # さらに連続しているブロックも探す
            #
            # 例：
            # 10:00～11:00
            # 11:00～12:00
            # 新規 10:30～11:30
            #
            # → 10:00～12:00
            # ==================================

            changed = True

            while changed:

                changed = False

                for block in existing_blocks:

                    block_start = datetime.strptime(
                        block["start_time"],
                        "%H:%M"
                    )

                    block_end = datetime.strptime(
                        block["end_time"],
                        "%H:%M"
                    )

                    # --------------------------------
                    # 重なっている、または
                    # ぴったり接している
                    # --------------------------------

                    if (
                        block_start <= merged_end
                        and block_end >= merged_start
                    ):

                        new_merged_start = min(
                            merged_start,
                            block_start
                        )

                        new_merged_end = max(
                            merged_end,
                            block_end
                        )

                        if (
                            new_merged_start
                            != merged_start
                            or
                            new_merged_end
                            != merged_end
                        ):

                            merged_start = (
                                new_merged_start
                            )

                            merged_end = (
                                new_merged_end
                            )

                            changed = True

            # ==================================
            # 既存の重複ブロックを削除
            # ==================================

            for block in existing_blocks:

                block_start = datetime.strptime(
                    block["start_time"],
                    "%H:%M"
                )

                block_end = datetime.strptime(
                    block["end_time"],
                    "%H:%M"
                )

                if (
                    block_start <= merged_end
                    and block_end >= merged_start
                ):

                    supabase.table(
                        "blocked_times"
                    ).delete().eq(
                        "id",
                        block["id"]
                    ).execute()

            # ==================================
            # 統合した1つのブロックを登録
            # ==================================

            merged_start_str = (
                merged_start.strftime("%H:%M")
            )

            merged_end_str = (
                merged_end.strftime("%H:%M")
            )

            result = (
                supabase
                .table("blocked_times")
                .insert(
                    {
                        "data": data,
                        "start_time": merged_start_str,
                        "end_time": merged_end_str
                    }
                )
                .execute()
            )

            registered_count += 1

            # ==================================
            # DB上の現在のブロック一覧を更新
            # ==================================

            block_result = (
                supabase
                .table("blocked_times")
                .select("*")
                .eq("data", data)
                .execute()
            )

            existing_blocks = (
                block_result.data or []
            )

    # ==========================================
    # 最新の一覧を取得
    # ==========================================

    result = (
        supabase
        .table("blocked_times")
        .select("*")
        .order("data")
        .order("start_time")
        .execute()
    )

    return render_template(
        "admin_block.html",
        blocked_times=result.data or [],
        message=(
            f"一括登録完了："
            f"{registered_count}件処理、"
            f"{duplicate_count}件スキップ、"
            f"{reservation_conflict_count}件スキップ"
        ),
        conflicts=conflicts
    )


# =========================================================
# ブロック削除
# =========================================================

@app.route(
    "/delete_block/<int:id>",
    methods=["POST"]
)
def delete_block(id):

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    supabase.table(
        "blocked_times"
    ).delete().eq(
        "id",
        id
    ).execute()


    return redirect(
        url_for("admin_block")
    )


@app.route(
    "/bulk_delete_block",
    methods=["POST"]
)
def bulk_delete_block():

    if not session.get("admin_logged_in"):
        return redirect(
            url_for("login")
        )

    ids = request.form.getlist("ids")

    if not ids:
        return redirect(
            request.referrer
            or url_for("admin_block")
        )

    deleted_count = 0

    for block_id in ids:

        try:
            block_id = int(block_id)

        except (ValueError, TypeError):
            continue

        result = (
            supabase
            .table("blocked_times")
            .delete()
            .eq("id", block_id)
            .execute()
        )

        if result.data:
            deleted_count += len(
                result.data
            )

    return redirect(
        url_for("admin_block")
    )


# =========================================================
# 新規予約
# =========================================================

@app.route("/new")
def new():

    result = (
        supabase
        .table("settings")
        .select(
            "start_data,"
            "end_data,"
            "capacity,"
            "start_time,"
            "end_time,"
            "interval"
        )
        .eq("id", 1)
        .single()
        .execute()
    )


    setting = result.data or {}


    return render_template(
        "new.html",
        start_data=setting.get(
            "start_data"
        ),
        end_data=setting.get(
            "end_data"
        ),
        capacity=setting.get(
            "capacity"
        ),
        start_time=setting.get(
            "start_time"
        ),
        end_time=setting.get(
            "end_time"
        ),
        interval=setting.get(
            "interval"
        )
    )


# =========================================================
# 予約可能時間取得
# =========================================================

@app.route("/get_times")
def get_times():
    date_str = (
        request.args.get("date")
        or request.args.get("data")
        or ""
    ).strip()

    if not date_str:
        return jsonify([])

    setting_result = (
        supabase.table("settings")
        .select(
            "start_data,"
            "end_data,"
            "capacity,"
            "start_time,"
            "end_time,"
            "interval"
        )
        .eq("id", 1)
        .single()
        .execute()
    )

    setting = setting_result.data or {}

    start_data = setting.get("start_data")
    end_data = setting.get("end_data")
    capacity = setting.get("capacity")
    start_time = setting.get("start_time")
    end_time = setting.get("end_time")
    interval = setting.get("interval") or 30

    if start_data and date_str < str(start_data):
        return jsonify([])

    if end_data and date_str > str(end_data):
        return jsonify([])

    if not start_time or not end_time:
        return jsonify([])

    start_time = str(start_time)[:5]
    end_time = str(end_time)[:5]

    reservation_result = (
        supabase.table("reservations")
        .select("time")
        .eq("data", date_str)
        .eq("is_deleted", False)
        .execute()
    )

    reservations = reservation_result.data or []

    reservation_counts = {}

    for row in reservations:
        t = row.get("time")

        if t:
            t = str(t)[:5]
            reservation_counts[t] = (
                reservation_counts.get(t, 0) + 1
            )

    blocked_result = (
        supabase.table("blocked_times")
        .select("start_time,end_time")
        .eq("data", date_str)
        .execute()
    )

    blocked_times = blocked_result.data or []

    today = datetime.today().date()

    current = datetime.combine(
        today,
        datetime.strptime(
            start_time,
            "%H:%M"
        ).time()
    )

        # -----------------------------------------------------
    # 24時間前チェック
    # -----------------------------------------------------

    now = datetime.now(JST)

    limit_datetime = now + timedelta(hours=24)

    end = datetime.combine(
        today,
        datetime.strptime(
            end_time,
            "%H:%M"
        ).time()
    )

    slots = []

    while current < end:

        time_str = current.strftime("%H:%M")

        # -------------------------------------------------
        # 24時間以内の予約枠は表示しない
        # -------------------------------------------------

        slot_datetime = datetime.combine(
            datetime.strptime(
                date_str,
                "%Y-%m-%d"
            ).date(),
            current.time()
        ).replace(
            tzinfo=JST
        )

        if slot_datetime < limit_datetime:
            current += timedelta(
                minutes=int(interval)
            )
            continue

        # 定員チェック
        if (
            capacity is not None
            and reservation_counts.get(
                time_str,
                0
            ) >= int(capacity)
        ):
            current += timedelta(
                minutes=int(interval)
            )
            continue

        slot_start = current

        slot_end = (
            current
            + timedelta(
                minutes=int(interval)
            )
        )

        # 予約不可時間チェック
        blocked = False

        for block in blocked_times:

            block_start_str = str(
                block.get("start_time", "")
            )[:5]

            block_end_str = str(
                block.get("end_time", "")
            )[:5]

            if (
                not block_start_str
                or not block_end_str
            ):
                continue

            block_start = datetime.combine(
                today,
                datetime.strptime(
                    block_start_str,
                    "%H:%M"
                ).time()
            )

            block_end = datetime.combine(
                today,
                datetime.strptime(
                    block_end_str,
                    "%H:%M"
                ).time()
            )

            if (
                slot_start < block_end
                and slot_end > block_start
            ):
                blocked = True
                break

        if not blocked:
            slots.append(time_str)

        current += timedelta(
            minutes=int(interval)
        )

        print(
        "get_times:",
        date_str,
        "→",
        slots
    )

    return jsonify(slots)


# =========================================================
# 予約確認画面
# =========================================================

@app.route(
    "/confirm",
    methods=["POST"]
)
def confirm():

    data = request.form.to_dict()


    return render_template(
        "confirm.html",
        data=data
    )


# =========================================================
# 新規予約確定
# =========================================================

@app.route(
    "/create_confirm",
    methods=["POST"]
)
def create_confirm():

    consumer_code = session.get(
        "code"
    )


    if not consumer_code:

        return redirect(
            url_for("index")
        )


    data = (
        request.form.get("data")
        or ""
    ).strip()

    time_str = (
        request.form.get("time")
        or ""
    ).strip()

    name = (
        request.form.get("name")
        or ""
    ).strip()

    phone = (
        request.form.get("phone")
        or ""
    ).strip()

    address = (
        request.form.get("address")
        or ""
    ).strip()

    email = (
        request.form.get("email")
        or ""
    ).strip()


    # -----------------------------------------------------
    # 24時間前チェック
    # -----------------------------------------------------

    try:

        reservation_datetime = datetime.strptime(
            f"{data} {time_str}",
            "%Y-%m-%d %H:%M"
        ).replace(
            tzinfo=JST
        )


        now = datetime.now(JST)


        if reservation_datetime < (
            now + timedelta(hours=24)
        ):

            return (
                "予約は24時間前までにお願いします。"
            ), 400


    except Exception as e:

        print(
            "24時間前チェックエラー:",
            e
        )


    # -----------------------------------------------------
    # 重複チェック
    # -----------------------------------------------------

    duplicate = (
        supabase
        .table("reservations")
        .select("id")
        .eq("data", data)
        .eq("time", time_str)
        .eq("is_deleted", False)
        .execute()
    )


    if duplicate.data:

        return (
            "その日時にはすでに予約があります。"
        ), 400


    # -----------------------------------------------------
    # 予約登録
    # -----------------------------------------------------

    new_row = {

        "data": data,

        "time": time_str,

        "consumer_code": consumer_code,

        "name": name,

        "phone": phone,

        "address": address,

        "email": email,

        "status": "新規",

        "is_deleted": False,

        "is_confirmed": False,

        "created_at": datetime.now(
            JST
        ).isoformat()
    }


    result = (
        supabase
        .table("reservations")
        .insert(new_row)
        .execute()
    )


    if result.data:

        mail_new(
            result.data[0]
        )


    return render_template(
        "complete.html",
        data={
            "consumer_code": consumer_code,
            "data": data,
            "time": time_str,
            "name": name,
            "phone": phone,
            "address": address,
            "email": email
        }
    )


# =========================================================
# 日付チェック
# =========================================================

@app.route("/check_day")
def check_day():
    date_str = (request.args.get("data") or request.args.get("date") or "").strip()

    if not date_str:
        return jsonify({
            "ok": False,
            "message": "日付が指定されていません。"
        })

    try:
        setting_result = (
            supabase
            .table("settings")
            .select(
                "start_data,"
                "end_data,"
                "start_time,"
                "end_time,"
                "interval"
            )
            .eq("id", 1)
            .single()
            .execute()
        )

        setting = setting_result.data or {}

        start_data = setting.get("start_data")
        end_data = setting.get("end_data")
        start_time = setting.get("start_time")
        end_time = setting.get("end_time")
        interval = setting.get("interval") or 30

        # -----------------------------------------
        # 予約可能期間チェック
        # -----------------------------------------

        if start_data and date_str < str(start_data):
            return jsonify({
                "ok": False,
                "message": "予約可能期間外です。"
            })

        if end_data and date_str > str(end_data):
            return jsonify({
                "ok": False,
                "message": "予約可能期間外です。"
            })

        # -----------------------------------------
        # 予約時間設定チェック
        # -----------------------------------------

        if not start_time or not end_time:
            return jsonify({
                "ok": False,
                "message": "予約時間が設定されていません。"
            })

        # -----------------------------------------
        # 時間を HH:MM に統一
        # Supabaseのtime型は HH:MM:SS で返る場合がある
        # -----------------------------------------

        def normalize_time(value):
            value = str(value)

            if len(value) >= 5:
                return value[:5]

            return value

        start_time = normalize_time(start_time)
        end_time = normalize_time(end_time)

        # -----------------------------------------
        # 予約不可時間を取得
        # -----------------------------------------

        blocked_result = (
            supabase
            .table("blocked_times")
            .select("start_time,end_time")
            .eq("data", date_str)
            .execute()
        )

        blocked_times = blocked_result.data or []

        # -----------------------------------------
        # 時間帯作成
        # -----------------------------------------

        current = datetime.combine(
            datetime.today(),
            datetime.strptime(
                start_time,
                "%H:%M"
            ).time()
        )

        end = datetime.combine(
            datetime.today(),
            datetime.strptime(
                end_time,
                "%H:%M"
            ).time()
        )

        interval = int(interval)

        slots = []

        # 最終時間は「開始できる時間」
        # なので、その時間からinterval分後が
        # 終了時間を超えないものだけ作る
        while current + timedelta(minutes=interval) <= end:

            slot_start = current

            slot_end = current + timedelta(
                minutes=interval
            )

            blocked = False

            for block in blocked_times:

                block_start_str = normalize_time(
                    block.get("start_time")
                )

                block_end_str = normalize_time(
                    block.get("end_time")
                )

                if not block_start_str or not block_end_str:
                    continue

                block_start = datetime.combine(
                    datetime.today(),
                    datetime.strptime(
                        block_start_str,
                        "%H:%M"
                    ).time()
                )

                block_end = datetime.combine(
                    datetime.today(),
                    datetime.strptime(
                        block_end_str,
                        "%H:%M"
                    ).time()
                )

                # 時間帯が重なっているか確認
                if (
                    slot_start < block_end
                    and
                    slot_end > block_start
                ):
                    blocked = True
                    break

            if not blocked:
                slots.append(
                    current.strftime("%H:%M")
                )

            current += timedelta(
                minutes=interval
            )

        # -----------------------------------------
        # 予約可能時間がない場合
        # -----------------------------------------

        if not slots:
            return jsonify({
                "ok": False,
                "message": "この日は予約できる時間帯がありません。"
            })

        return jsonify({
            "ok": True,
            "message": "予約可能です。"
        })

    except Exception as e:

        print("check_day エラー:", repr(e))

        return jsonify({
            "ok": False,
            "message": "予約可能日の確認中にエラーが発生しました。"
        }), 500


# =========================================================
# 予約変更画面
# =========================================================

@app.route("/edit")
def edit():

    consumer_code = session.get(
        "code"
    )


    if not consumer_code:

        return redirect(
            url_for("index")
        )


    result = (
        supabase
        .table("reservations")
        .select(
            "id,"
            "data,"
            "time,"
            "name,"
            "phone,"
            "address,"
            "email"
        )
        .eq(
            "consumer_code",
            consumer_code
        )
        .eq(
            "is_deleted",
            False
        )
        .order(
            "created_at",
            desc=True
        )
        .limit(1)
        .execute()
    )


    if not result.data:

        return "予約が見つかりません。", 404


    reservation = result.data[0]


    setting_result = (
        supabase
        .table("settings")
        .select(
            "start_data,"
            "end_data"
        )
        .eq("id", 1)
        .single()
        .execute()
    )


    setting = setting_result.data or {}


    return render_template(
        "edit.html",
        reservation=reservation,
        start_data=setting.get(
            "start_data"
        ),
        end_data=setting.get(
            "end_data"
        )
    )


# =========================================================
# 変更確認
# =========================================================

@app.route(
    "/edit_confirm",
    methods=["POST"]
)
def edit_confirm():

    data = request.form.to_dict()


    return render_template(
        "edit_confirm.html",
        data=data
    )


# =========================================================
# 予約削除画面
# =========================================================

@app.route("/delete")
def delete():

    consumer_code = session.get(
        "code"
    )


    if not consumer_code:

        return redirect(
            url_for("index")
        )


    result = (
        supabase
        .table("reservations")
        .select("*")
        .eq(
            "consumer_code",
            consumer_code
        )
        .eq(
            "is_deleted",
            False
        )
        .order(
            "created_at",
            desc=True
        )
        .limit(1)
        .execute()
    )


    if not result.data:

        return "予約が見つかりません。", 404


    reservation = result.data[0]


    reservation["time_display"] = (
        format_time_range(
            reservation.get("time", "")
        )
    )


    return render_template(
        "delete.html",
        reservation=reservation
    )


# =========================================================
# 予約削除確定
# =========================================================

@app.route(
    "/delete_confirm",
    methods=["POST"]
)
def delete_confirm():

    consumer_code = session.get(
        "code"
    )


    if not consumer_code:

        return redirect(
            url_for("index")
        )


    result = (
        supabase
        .table("reservations")
        .select("*")
        .eq(
            "consumer_code",
            consumer_code
        )
        .eq(
            "is_deleted",
            False
        )
        .order(
            "created_at",
            desc=True
        )
        .limit(1)
        .execute()
    )


    if not result.data:

        return "予約が見つかりません。", 404


    original = result.data[0]


    delete_row = {

        "data": original.get("data"),

        "time": original.get("time"),

        "consumer_code": consumer_code,

        "name": original.get("name"),

        "phone": original.get("phone"),

        "address": original.get("address"),

        "email": original.get("email"),

        "status": "削除",

        "is_deleted": False,

        "is_confirmed": False,

        "created_at": datetime.now(
            JST
        ).isoformat()
    }


    delete_result = (
        supabase
        .table("reservations")
        .insert(delete_row)
        .execute()
    )


    if delete_result.data:

        mail_delete(
            delete_result.data[0]
        )


    return render_template(
        "delete_done.html"
    )


# =========================================================
# 予約確認
# =========================================================

@app.route('/view')
def view():
    code = session.get('code')

    if not code:
        return redirect('/')

    # 同じお客様コードの予約を、最新のものから取得
    res = supabase.table("reservations") \
        .select("*") \
        .eq("consumer_code", code) \
        .eq("is_deleted", False) \
        .order("id", desc=True) \
        .limit(1) \
        .execute()

    row = res.data[0] if res.data else None

    # 最新の予約がない場合
    if not row:
        return render_template("view.html", data=None)

    # 最新の予約が「削除」なら、現在の予約はない扱い
    if str(row.get("status") or "").strip() == "削除":
        return render_template("view.html", data=None)

    # 予約時間を「開始～終了」にする
    time_range = format_time_range(row.get("time", ""))

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

    return render_template("view.html", data=data)


# =========================================================
# 予約Excel
# =========================================================

@app.route("/export_excel")
def export_excel():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    result = (
        supabase
        .table("reservations")
        .select("*")
        .eq("is_deleted", False)
        .order("data")
        .order("time")
        .execute()
    )


    rows = result.data or []


    wb = Workbook()

    ws = wb.active

    ws.title = "予約一覧"


    ws.append(
        [
            "予約日",
            "予約時間",
            "申込日時",
            "お客様コード",
            "氏名",
            "住所",
            "電話番号",
            "メールアドレス",
            "ステータス"
        ]
    )


    for row in rows:

        time_str = row.get(
            "time",
            ""
        )


        created_at = row.get(
            "created_at",
            ""
        )


        if created_at:

            created_at = (
                created_at
                .replace("T", " ")
                [:19]
            )


        ws.append(
            [
                row.get("data", ""),
                time_str,
                created_at,
                row.get(
                    "consumer_code",
                    ""
                ),
                row.get("name", ""),
                row.get("address", ""),
                row.get("phone", ""),
                row.get("email", ""),
                row.get("status", "")
            ]
        )


    output = io.BytesIO()

    wb.save(output)

    output.seek(0)


    return send_file(
        output,
        as_attachment=True,
        download_name="予約一覧.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# =========================================================
# 削除済み予約Excel
# =========================================================

@app.route("/export_deleted_excel")
def export_deleted_excel():

    if not session.get("admin_logged_in"):

        return redirect(
            url_for("login")
        )


    result = (
        supabase
        .table("reservations")
        .select("*")
        .eq("is_deleted", True)
        .order("created_at")
        .execute()
    )


    rows = result.data or []


    wb = Workbook()

    ws = wb.active

    ws.title = "削除済み予約"


    ws.append(
        [
            "予約日",
            "予約時間",
            "申込日時",
            "お客様コード",
            "氏名",
            "住所",
            "電話番号",
            "ステータス"
        ]
    )


    for row in rows:

        created_at = row.get(
            "created_at",
            ""
        )


        if created_at:

            created_at = (
                created_at
                .replace("T", " ")
                [:19]
            )


        ws.append(
            [
                row.get("data", ""),
                row.get("time", ""),
                created_at,
                row.get(
                    "consumer_code",
                    ""
                ),
                row.get("name", ""),
                row.get("address", ""),
                row.get("phone", ""),
                row.get("status", "")
            ]
        )


    output = io.BytesIO()

    wb.save(output)

    output.seek(0)


    return send_file(
        output,
        as_attachment=True,
        download_name="削除済み予約.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
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