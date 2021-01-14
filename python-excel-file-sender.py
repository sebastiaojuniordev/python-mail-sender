import smtplib
from email.message import EmailMessage

SMTP_PORT = 587
SMTP_HOST = "smtp.office365.com"
SENDER_EMAIL = "teste@hotmail.com.br"
SENDER_PASSWORD = "teste"

def send_mail_with_attachment(recipient_email, subject, content, excel_file):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient_email
    msg.set_content(content)

    with open(excel_file, 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application",
                       subtype="xlsx", filename=excel_file)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
        smtp.send_message(msg)
        smtp.quit()


if __name__ == '__main__':
    send_mail_with_attachment("test@hotmail.com.br", "Subject...", "Description...", "file.csv")
