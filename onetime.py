import smtplib
from email.mime.text import MIMEText

ADMIN_EMAIL = "penseries2tensen+reservationreport@gmail.com"

def send_admin_mail(subject, body):

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = "system@example.com"
    msg["To"] = ADMIN_EMAIL

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()

    server.login(
        "penseries2tensen+reservationreport@gmail.com",
        "adtygxeaishohucc"
    )

    server.send_message(msg)
    server.quit()

# ★ここが重要（実行トリガー）
send_admin_mail(
    "テストメール",
    "これは予約システムのテストです"
)

print("送信処理完了")