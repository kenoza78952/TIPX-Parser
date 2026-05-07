# email_utils.py
import os
import smtplib
from email.message import EmailMessage
from debug_utils import debug

SMTP_USER = ""    
SMTP_PASS = ""       

def send_email_with_attachments(
    to=[""],  #
    subject="🔥 Automatic Report from TIPX",
    body="Hello,\n\nAttached is the latest Excel report for changes detected.\n\nBest regards,\nTIPX System",
    attachments=None,
    from_addr=SMTP_USER,
    smtp_host=,
    smtp_port=,
    username=None,
    password=None
):
    debug("Preparing email...", "DEBUG")

    username = SMTP_USER
    password = SMTP_PASS

    if not username or not password:
        debug("Missing SMTP credentials!", "ERROR")
        return

    if isinstance(to, str):
        to = [to] 

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = ", ".join(to)
    msg.set_content(body)

    if attachments:
        for file_path in attachments:
            if not os.path.isfile(file_path):
                debug(f"Attachment not found: {file_path}", "WARNING")
                continue

            with open(file_path, "rb") as f:
                file_data = f.read()
                file_name = os.path.basename(file_path)

            msg.add_attachment(
                file_data,
                maintype="application",
                subtype="octet-stream",
                filename=file_name
            )
            debug(f"📎 Attached file: {file_name}", "DEBUG")
    else:
        debug("📭 No attachments to add.", "DEBUG")

    debug("🔌 Connecting to SMTP...", "DEBUG")
    try:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()  # Secure connection
            server.login(username, password)
            debug("✅ Logged into SMTP server successfully.", "DEBUG")
            server.send_message(msg)
            debug(f"📤 Email sent to: {to}", "INFO")
    except Exception as e:
        debug(f"❌ Error sending email: {e}", "ERROR")
