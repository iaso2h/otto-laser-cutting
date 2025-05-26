from config import cfg
from util import pr
import smtplib, ssl

sslPort        = cfg.email.sslPort
smtp_server    = cfg.email.smtpServer
receiver_email = cfg.email.recieverAccount
sender_email   = cfg.email.senderAccount
password       = cfg.email.senderPassword


def send(message: str):
    """
    Sends an email notification using SMTP_SSL.

    Args:
        message (str): The email message content to be sent.

    Note:
        Requires SMTP server configuration (smtp_server, sslPort) and email credentials
        (sender_email, password, receiver_email) to be properly set.
        If any credential is missing, the function will exit silently.
    """
    if not all((smtp_server, receiver_email, sender_email, password)):
        return pr("Email notification hasn't configured. Please check configuration file.")
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, sslPort, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message)
    pr(f"Sent email to{receiver_email}")
