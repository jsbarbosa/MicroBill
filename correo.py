import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from config import COTIZACION_MENSAJE, COTIZACION_SUBJECT, SERVER, PORT, REPORTE_MENSAJE, REPORTE_SUBJECT, REQUEST_SUBJECT, REQUEST_MENSAJE
from login import *

import constants

CORREO = None

def initCorreo():
    global CORREO
    CORREO = smtplib.SMTP(SERVER, PORT)
    CORREO.ehlo() # Hostname to send for this command defaults to the fully qualified domain name of the local host.
    CORREO.starttls() #Puts connection to SMTP server in TLS mode
    CORREO.ehlo()
    CORREO.login(FROM, PASSWORD)

def sendEmail(to, subject, text, attachments = []):
    global CORREO
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = FROM

    msg['To'] = to
    msg['Cc'] = FROM

    body = MIMEText(text, 'html')
    msg.attach(body)

    for item in attachments:
        name = os.path.join(constants.PDF_DIR, item + ".pdf")
        with open(name, "rb") as file:
            app = MIMEApplication(file.read())
            app.add_header('Content-Disposition', 'attachment', filename = item + ".pdf")
            msg.attach(app)

    to = [to, FROM]

    if CORREO != None:
        CORREO.sendmail(FROM, to, msg.as_string())
    else:
        initCorreo()
        if CORREO != None:
            CORREO.sendmail(FROM, to, msg.as_string())

def sendCotizacion(to, file_name):
    sendEmail(to, COTIZACION_SUBJECT + " - %s"%file_name, COTIZACION_MENSAJE, [file_name])

def sendRegistro(to, file_name):
    sendEmail(to, REPORTE_SUBJECT + " - %s"%file_name, REPORTE_MENSAJE, [file_name + "_Reporte"])

def sendRequest(to):
    sendEmail(to, REQUEST_SUBJECT, REQUEST_MENSAJE)
