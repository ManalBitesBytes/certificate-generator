import pandas as pd
import time
import smtplib
from email.utils import formataddr, formatdate
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from datetime import datetime
from email_content import participant_html, volunteer_html, plain_text

LOG_FILE = 'failed_emails.log'

def log_failure(recipient_email, recipient_name, error_message):
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(f"[{datetime.now()}] {recipient_name} <{recipient_email}> - ERROR: {error_message}\n")


def send_certificates_by_email(csv_file, pdf_dir, sender_email, sender_password, smtp_server, smtp_port):
    """
    Reads recipient data from a CSV, attaches corresponding PDFs, and sends emails.

    Args:
        csv_file (str): Path to the CSV file containing recipient data.
        pdf_dir (str): Directory where generated PDF certificates are stored.
        sender_email (str): The email address from which to send emails.
        sender_password (str): The password for the sender_email (use an App Password for Gmail).
        smtp_server (str): The SMTP server address (e.g., 'smtp.gmail.com').
        smtp_port (int): The SMTP server port (e.g., 587 for TLS).
    """
    try:
        df = pd.read_csv(csv_file)
        print(f"Successfully read CSV file: {csv_file}")
    except FileNotFoundError:
        print(f"Error: CSV file not found at {csv_file}")
        return
    except Exception as e:
        print(f"Error reading CSV file: {e}")
        return
    #SMTP
    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=10)
        print(f"Successfully connected to SMTP server: {smtp_server}:{smtp_port}")
        server.login(sender_email, sender_password)
        print("Login successful.")
    except Exception as e:
        print(f"Error connecting to SMTP server or logging in: {e}")
        return


    for index, row in df.iterrows():
        recipient_name = row['Name']
        recipient_email = row['Email'].lower()



        pdf_filename = f"{recipient_name}.pdf"
        pdf_path = os.path.join(pdf_dir, pdf_filename)

        if not os.path.exists(pdf_path):
            print(f"Warning: PDF not found for {recipient_name} at {pdf_path}. Skipping email for this recipient.")
            continue

        print(f"\nPreparing email for: {recipient_name} <{recipient_email}>")

        #email body
        msg = MIMEMultipart('alternative')
        msg['From'] = formataddr(("فريق آياتثون", sender_email))
        msg['To'] = recipient_email
        msg['Subject'] = "شهادة شكر وتقدير | شكرًا لعطائكم في آياتثون" #"شهادة مشاركة | شكرًا لمشاركتكم في آياتثون"
        msg['Date'] = formatdate(localtime=True)
        msg.add_header('List-Unsubscribe', '<mailto:unsubscribe@parmg.sa>')
        msg.attach(MIMEText(plain_text, 'plain', 'utf-8'))
        msg.attach(MIMEText(volunteer_html, 'html', 'utf-8'))

        #attaching pdfs
        with open(pdf_path, 'rb') as f:
            attach = MIMEApplication(f.read(), _subtype="pdf")
            attach.add_header('Content-Disposition', 'attachment', filename=pdf_filename)
            msg.attach(attach)
            print(f"Attached PDF: {pdf_filename}")

        # Send the email
        try:
            server.send_message(msg)
            print(f"Email sent successfully to {recipient_email}")
        except Exception as e:
            error_msg = f"{e.__class__.__name__}: {e}"
            print(f"Failed to send email to {recipient_email}: {error_msg}")
            log_failure(recipient_email, recipient_name, error_msg)

        time.sleep(5)

    server.quit()  # Close the SMTP connection
    print("\nEmail sending process completed.")


if __name__ == "__main__":

    SENDER_EMAIL = ''
    SENDER_PASSWORD = ''


    SMTP_SERVER = 'smtp.dreamhost.com'
    SMTP_PORT = 465


    CSV_FILE_PATH = 'recipients/volunteers_data.csv'
    PDF_DIRECTORY = 'pdf_volunteers'


    send_certificates_by_email(
            CSV_FILE_PATH,
            PDF_DIRECTORY,
            SENDER_EMAIL,
            SENDER_PASSWORD,
            SMTP_SERVER,
            SMTP_PORT
        )
