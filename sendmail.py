import pandas as pd
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import os

# Load email addresses from Excel
df = pd.read_excel('emails.xlsx')
email_list = df['Email'].dropna().tolist()

# Gmail credentials
load_dotenv()
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_PASSWORD = os.getenv("GMAIL_PASSWORD")  # Use App Password if 2FA is enabled

# Email content
subject = 'Janitorial Quote- Follow Up'
body = '''Hello, Good Morning!

Are you interested in getting a cleaning service estimate for your premises?

Once we receive your response, we can arrange a walk-through to provide an accurate estimate.

Best Regards,
Belinda
Janitorial Services - New Jersey

'''

# Set up the SMTP server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(GMAIL_USER, GMAIL_PASSWORD)

# Send emails
for email in email_list:
    msg = EmailMessage()
    msg['From'] = GMAIL_USER
    msg['To'] = email
    msg['Subject'] = subject
    msg.set_content(body)

    try:
        server.send_message(msg)
        print(f'Email sent to {email}')
    except Exception as e:
        print(f'Failed to send to {email}: {e}')

server.quit()
