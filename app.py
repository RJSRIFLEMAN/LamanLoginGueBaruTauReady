from flask import Flask, render_template, request, redirect
import pandas as pd
import smtplib, ssl
from email.mime.text import MIMEText
import os
from dotenv import load_dotenv
from datetime import datetime
from flask import send_file

load_dotenv()

app = Flask(__name__)
OTP_STORE = {}

# Kirim OTP via Gmail SMTP
def send_otp(email, otp):
    msg = MIMEText(f"Kode OTP kamu adalah: {otp}")
    msg["Subject"] = "Kode OTP Login"
    msg["From"] = os.getenv("EMAIL_SENDER")
    msg["To"] = email

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(os.getenv("EMAIL_SENDER"), os.getenv("EMAIL_PASS"))
        server.send_message(msg)

@app.route('/')
def home():
    return render_template('login.html')

@app.route('/verify', methods=['POST'])
def verify():
    email = request.form['email']
    otp = str(datetime.now().second * 42)[-4:]
    OTP_STORE[email] = otp
    send_otp(email, otp)
    return f'''
        <form method="POST" action="/login">
            <input type="hidden" name="email" value="{email}">
            <input type="text" name="otp" placeholder="Masukkan OTP">
            <button type="submit">Verifikasi OTP</button>
        </form>
    '''

@app.route('/login', methods=['POST'])
def login():
    email = request.form['email']
    otp_input = request.form['otp']

    if OTP_STORE.get(email) != otp_input:
        return "OTP Salah!", 403

    df = pd.DataFrame([[email, datetime.now()]], columns=["Email", "Waktu Login"])
    try:
        existing = pd.read_excel("data.xlsx")
        df = pd.concat([existing, df], ignore_index=True)
    except FileNotFoundError:
        pass
    df.to_excel("data.xlsx", index=False)

    return render_template("success.html")

if __name__ == '__main__':
    app.run(debug=True)

@app.route('/data')
def lihat_data():
    try:
        df = pd.read_excel("data.xlsx")
        return df.to_html(index=False)
    except:
        return "Belum ada data login atau file tidak ditemukan."

@app.route('/download')
def download_excel():
    try:
        return send_file ("data.xlsx", as_attachment=True)
    except:
        return "File data.xlsx tidak ditemukan."
