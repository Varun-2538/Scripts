import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import configparser
import os

def send_emails():
    config = configparser.ConfigParser()
    config.read('config.ini')

    sender_email = config.get('EMAIL', 'sender_email')
    sender_password = config.get('EMAIL', 'app_password')
    smtp_server = config.get('EMAIL', 'smtp_server')
    smtp_port = config.getint('EMAIL', 'smtp_port')
    use_tls = config.get('EMAIL', 'use_tls').strip().lower() == 'yes'

    subject = "Interview Invitation for Crypto Domain at Blockchain Club SRM"
    content = """Dear {name},

Thank you for your interest in Crypto Domain within Blockchain Club SRM. We were impressed with your application and are excited to invite you to interview for the role.

We are scheduling interviews for shortlisted candidates starting from September 25th. To best accommodate your schedule, we've provided a Calendly link below where you can choose a convenient time for your interview:

https://calendly.com/d/cmfb-fjt-v8g/blockchain-club-srm-crypto-interview
Please note: Interviews will be held virtually via Google Meet.

To prepare for the interview, we kindly request you to have your Linkedin and Github readily available to share in the Calendly form. 

If you have any questions about the interview process or the position, please don't hesitate to reach out to the leads "Varun Singh" (varunsingh2538@gmail.com) or "Haaniya Iram" (hs5937@srmist.edu.in).

We look forward to meeting you soon!

Sincerely,

Blockchain Club SRM"""

    excel_file = "D:\Scripts\demo.xlsx"
    df = pd.read_excel(excel_file)

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        if use_tls:
            server.starttls()
        server.login(sender_email, sender_password)
    except Exception as e:
        print(f"Failed to connect to the SMTP server: {e}")
        return

    for index, row in df.iterrows():
        recipient_name = str(row['Name'])
        recipient_email = str(row['Email'])

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        personalized_content = content.replace("{name}", recipient_name)

        msg.attach(MIMEText(personalized_content, 'plain'))

        try:
            server.sendmail(sender_email, recipient_email, msg.as_string())
            print(f"Email sent to {recipient_name} at {recipient_email}")
        except Exception as e:
            print(f"Failed to send email to {recipient_email}: {e}")

    server.quit()

if __name__ == '__main__':
    send_emails()