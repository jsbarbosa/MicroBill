import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from config import SERVER, PORT, REPORTE_MENSAJE, REPORTE_SUBJECT, REQUEST_SUBJECT, REQUEST_MENSAJE, DEPENDENCIAS
from config import COTIZACION_MENSAJE_RECIBO, COTIZACION_MENSAJE_FACTURA, COTIZACION_MENSAJE_TRANSFERENCIA
from config import COTIZACION_SUBJECT_RECIBO, COTIZACION_SUBJECT_FACTURA, COTIZACION_SUBJECT_TRANSFERENCIA
from login import *

import constants

dependencias = ("\n" + "\n".join(DEPENDENCIAS)).title()

dependencias = dependencias.replace("De", "de")

COTIZACION_MENSAJE_RECIBO = (COTIZACION_MENSAJE_RECIBO + dependencias).replace("\n", "<br>")
COTIZACION_MENSAJE_FACTURA = (COTIZACION_MENSAJE_FACTURA + dependencias).replace("\n", "<br>")
COTIZACION_MENSAJE_TRANSFERENCIA = (COTIZACION_MENSAJE_TRANSFERENCIA + dependencias).replace("\n", "<br>")


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

def sendCotizacionRecibo(to, file_name):
    sendEmail(to, COTIZACION_SUBJECT_RECIBO + " - %s"%file_name, COTIZACION_MENSAJE_RECIBO, [file_name])

def sendCotizacionTransferencia(to, file_name):
    sendEmail(to, COTIZACION_SUBJECT_TRANSFERENCIA + " - %s"%file_name, COTIZACION_MENSAJE_TRANSFERENCIA, [file_name])

def sendCotizacionFactura(to, file_name):
    sendEmail(to, COTIZACION_SUBJECT_FACTURA + " - %s"%file_name, COTIZACION_MENSAJE_FACTURA, [file_name])

def sendRegistro(to, file_name):
    sendEmail(to, REPORTE_SUBJECT + " - %s"%file_name, REPORTE_MENSAJE, [file_name + "_Reporte"])

def sendRequest(to):
    sendEmail(to, REQUEST_SUBJECT, REQUEST_MENSAJE)
