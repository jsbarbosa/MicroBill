import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from config import SEND_SERVER, SEND_PORT, REPORTE_MENSAJE, REPORTE_SUBJECT, REQUEST_SUBJECT, REQUEST_MENSAJE, DEPENDENCIAS
from config import COTIZACION_MENSAJE_RECIBO, COTIZACION_MENSAJE_FACTURA, COTIZACION_MENSAJE_TRANSFERENCIA
from config import COTIZACION_SUBJECT_RECIBO, COTIZACION_SUBJECT_FACTURA, COTIZACION_SUBJECT_TRANSFERENCIA
from config import GESTOR_RECIBO_CORREO, GESTOR_RECIBO_SUBJECT, GESTOR_RECIBO_MENSAJE
from config import GESTOR_FACTURA_CORREO, GESTOR_FACTURA_SUBJECT, GESTOR_FACTURA_MENSAJE, SALUDO
from login import *
import constants

dependencias = ("\n" + "\n".join(DEPENDENCIAS)).title()

dependencias = dependencias.replace("De", "de")

SALUDO = SALUDO.replace("\n", "<br>")

COTIZACION_MENSAJE_RECIBO = (COTIZACION_MENSAJE_RECIBO + dependencias).replace("\n", "<br>")
COTIZACION_MENSAJE_FACTURA = (COTIZACION_MENSAJE_FACTURA + dependencias).replace("\n", "<br>")
COTIZACION_MENSAJE_TRANSFERENCIA = (COTIZACION_MENSAJE_TRANSFERENCIA + dependencias).replace("\n", "<br>")
REPORTE_MENSAJE = (REPORTE_MENSAJE + dependencias).replace("\n", "<br>")
REQUEST_MENSAJE = (REQUEST_MENSAJE + dependencias).replace("\n", "<br>")

GESTOR_RECIBO_MENSAJE = (GESTOR_RECIBO_MENSAJE + dependencias).replace("\n", "<br>")
GESTOR_FACTURA_MENSAJE = (GESTOR_FACTURA_MENSAJE + dependencias).replace("\n", "<br>")

CORREO = None

def initCorreo():
    global CORREO
    CORREO = smtplib.SMTP(SEND_SERVER, SEND_PORT, timeout = 30)
    CORREO.ehlo() # Hostname to send for this command defaults to the fully qualified domain name of the local host.
    CORREO.starttls() #Puts connection to SMTP server in TLS mode
    CORREO.ehlo()
    CORREO.login(FROM, PASSWORD)
    # except Exception as e:
    #     CORREO = None
    #     raise(e)

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
        item = os.path.basename(item + ".pdf")
        with open(name, "rb") as file:
            app = MIMEApplication(file.read())
            app.add_header('Content-Disposition', 'attachment', filename = item)
            msg.attach(app)

    to = [to, FROM]

    for i in range(5):
        try:
            initCorreo()
            CORREO.sendmail(FROM, to, msg.as_string())
            break
        except Exception as e:
            print(e)
            if i == 4:
                raise(Exception("Could not send email."))

def sendCotizacionRecibo(to, file_name, observaciones = ""):
    if observaciones != "": observaciones = observaciones.replace("\n", "<br>") + 2*"<br>"
    sendEmail(to, COTIZACION_SUBJECT_RECIBO + " - %s"%file_name, SALUDO + observaciones + COTIZACION_MENSAJE_RECIBO, [file_name])

def sendCotizacionTransferencia(to, file_name, observaciones = ""):
    if observaciones != "": observaciones = observaciones.replace("\n", "<br>") + 2*"<br>"
    sendEmail(to, COTIZACION_SUBJECT_TRANSFERENCIA + " - %s"%file_name,  SALUDO + observaciones + COTIZACION_MENSAJE_TRANSFERENCIA, [file_name])

def sendCotizacionFactura(to, file_name, observaciones = ""):
    if observaciones != "": observaciones = observaciones.replace("\n", "<br>") + 2*"<br>"
    sendEmail(to, COTIZACION_SUBJECT_FACTURA + " - %s"%file_name, SALUDO + observaciones + COTIZACION_MENSAJE_FACTURA, [file_name])

def sendRegistro(to, file_name):
    sendEmail(to, REPORTE_SUBJECT + " - %s"%file_name, SALUDO + REPORTE_MENSAJE, [file_name + "_Reporte"])

def sendRequest(to):
    sendEmail(to, REQUEST_SUBJECT, SALUDO + REQUEST_MENSAJE)

def sendGestorRecibo(file_name):
    sendEmail(GESTOR_RECIBO_CORREO, GESTOR_RECIBO_SUBJECT + " - %s"%file_name, GESTOR_RECIBO_MENSAJE, [file_name])

def sendGestorFactura(file_name, orden_name):
    orden_name = orden_name.split(".")[:-1]
    if type(orden_name) is list:
        orden_name = ".".join(orden_name)
    sendEmail(GESTOR_FACTURA_CORREO, GESTOR_FACTURA_SUBJECT + " - %s"%file_name, GESTOR_FACTURA_MENSAJE, [file_name, orden_name])
