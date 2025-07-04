import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os


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

    # Establish SMTP connection outside the loop for efficiency
    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)

        server.login(sender_email, sender_password)
        print(f"Successfully connected to SMTP server: {smtp_server}:{smtp_port}")
    except Exception as e:
        print(f"Error connecting to SMTP server or logging in: {e}")
        print("Please check your sender_email, sender_password, SMTP server, and port.")
        print("For Gmail, ensure 'Less secure app access' is enabled or use an App Password if 2FA is on.")
        return

    for index, row in df.iterrows():
        recipient_name = row['name']
        recipient_email = row['email']


        pdf_filename = f"{recipient_name}.pdf"
        pdf_path = os.path.join(pdf_dir, pdf_filename)

        if not os.path.exists(pdf_path):
            print(f"Warning: PDF not found for {recipient_name} at {pdf_path}. Skipping email for this recipient.")
            continue

        print(f"\nPreparing email for: {recipient_name} <{recipient_email}>")

        #email body
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = f"شهادة إتمام الدورة - {recipient_name}"

        body_html = f"""
                <html>
                    <body dir="rtl" style="font-family: 'Tajawal', 'Cairo', 'Amiri', sans-serif; line-height: 1.6; color: #333;">
                        <p style="text-align: right;">عزيزي/ـتي <strong>{recipient_name}</strong>،</p>
                        <p style="text-align: right;">تهانينا!</p>
                        <p style="text-align: right;">يسرنا أن نرفق لكم شهادة إتمام الدورة التدريبية التي أكملتموها بنجاح.</p>
                        <p style="text-align: right;">نتمنى لكم كل التوفيق في مسيرتكم المهنية.</p>
                        <p style="text-align: right;">مع خالص التقدير،</p>
                        <p style="text-align: right;">فريق التدريب</p>
                    </body>
                </html>
                """
        msg.attach(MIMEText(body_html, 'html', 'utf-8'))

        # Attach the PDF file
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
            print(f"Failed to send email to {recipient_email}: {e}")

    server.quit()  # Close the SMTP connection
    print("\nEmail sending process completed.")


# --- Main execution ---
if __name__ == "__main__":

    SENDER_EMAIL = os.getenv("SENDER_EMAIL")
    SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
    SMTP_SERVER = 'smtp.gmail.com'
    SMTP_PORT = 465  #ssl


    CSV_FILE_PATH = 'certified.csv'
    PDF_OUTPUT_DIRECTORY = 'generated_certificates'


    send_certificates_by_email(
            CSV_FILE_PATH,
            PDF_OUTPUT_DIRECTORY,
            SENDER_EMAIL,
            SENDER_PASSWORD,
            SMTP_SERVER,
            SMTP_PORT
        )
