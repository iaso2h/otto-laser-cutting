from config import cfg
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from pathlib import Path
from typing import Optional

sslPort       = cfg.email.sslPort
smtpServer    = cfg.email.smtpServer
receiverEmails = cfg.email.receiverAccounts
senderEmail   = cfg.email.senderAccount
password      = cfg.email.senderPassword

emailSubjectTemplateDict = {
    "finished01":        "切割完成",
    "finished02":        "切割完成",
    "pauseThenContinue": "切割头碰板",
    "alert":             "警报",
    "alertForceReturn":  "警报: 强制回原点",
}


def send(templateName: str, tubeProTitle: str, imagePath: Optional[Path] = None) -> None:
    """
    Sends an email notification using SMTP_SSL, optionally with an image attachment.

    Args:
        templateName (str): The email template name (e.g., "finished01").
        tubeProTitle (str): The email message content.
        imagePath (str, optional): Path to the image file to attach. Defaults to None.

    Note:
        Requires SMTP server configuration (smtp_server, sslPort) and email credentials
        (sender_email, password, receiver_email) to be properly set.
        If any credential is missing, the function will exit silently.
    """
    if not all((smtpServer, receiverEmails, senderEmail, password)):
        return print("Email notification hasn't configured. Please check configuration file.")

    # Create the email message
    msg = MIMEMultipart()
    msg['Subject'] = emailSubjectTemplateDict[templateName]
    msg['From'] = senderEmail
    msg['To'] = ", ".join(receiverEmails)

    # Attach the text message
    msg.attach(MIMEText(tubeProTitle, 'plain'))

    # Attach the image if provided
    if imagePath and imagePath.exists():
        with open(imagePath, 'rb') as imgFile:
            imgData = imgFile.read()
        image = MIMEImage(imgData, name=imagePath.name)
        msg.attach(image)

    # Send the email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtpServer, sslPort, context=context) as server:
        server.login(senderEmail, password)
        server.sendmail(senderEmail, receiverEmails, msg.as_string())
