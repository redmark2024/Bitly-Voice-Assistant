import smtplib
import random
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Set your Gmail credentials here
GMAIL_USER = 'your_gmail@gmail.com'  # <-- CHANGE THIS
GMAIL_APP_PASSWORD = 'your_app_password'  # <-- CHANGE THIS (App Password, not your main password)

def generate_otp(length=6):
    return ''.join([str(random.randint(0, 9)) for _ in range(length)])

def send_otp_email(to_email, otp):
    subject = 'Your Bitly Voice Assistant Password Reset OTP'
    body = f'Your OTP for password reset is: {otp}\n\nIf you did not request this, please ignore this email.'

    msg = MIMEMultipart()
    msg['From'] = GMAIL_USER
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Error sending OTP email: {e}")
        return False 