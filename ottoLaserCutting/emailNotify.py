from config import cfg
from console import print
import smtplib, ssl

sslPort        = cfg.email.sslPort
smtp_server    = cfg.email.smtpServer
receiver_email = cfg.email.recieverAccount
sender_email   = cfg.email.senderAccount
password       = cfg.email.senderPassword

def send(message: str):
    if not all((smtp_server, receiver_email, sender_email, password)):
        return
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, sslPort, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message)
    print(f"Sent email to{receiver_email}")
