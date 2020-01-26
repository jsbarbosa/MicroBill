import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from .config import SEND_SERVER, SEND_PORT, REPORTE_MENSAJE, REPORTE_SUBJECT, REQUEST_SUBJECT, REQUEST_MENSAJE
from .config import COTIZACION_MENSAJE_RECIBO, COTIZACION_MENSAJE_FACTURA, COTIZACION_MENSAJE_TRANSFERENCIA
from .config import COTIZACION_SUBJECT_RECIBO, DEPENDENCIAS
from .config import GESTOR_RECIBO_CORREO, GESTOR_RECIBO_SUBJECT, GESTOR_RECIBO_MENSAJE
from .config import GESTOR_FACTURA_CORREO, GESTOR_FACTURA_SUBJECT, GESTOR_FACTURA_MENSAJE, SALUDO
from . import constants, config
from .utils import export, print_log
from .objects import Cotizacion
from typing import Callable, Iterable

__all__ = ["dependencias", "SALUDO", "COTIZACION_MENSAJE_RECIBO", "COTIZACION_MENSAJE_FACTURA",
           "COTIZACION_MENSAJE_TRANSFERENCIA", "REPORTE_MENSAJE", "REQUEST_MENSAJE", "GESTOR_RECIBO_MENSAJE",
           "GESTOR_FACTURA_MENSAJE", "CORREO"]

dependencias: str = ("\n" + "\n".join(DEPENDENCIAS)).title()

#: almacena la variable DEPENDENCIAS de la config como un string con formato de saltos de línea (enter)
dependencias = dependencias.replace("De", "de")

#: almacena la constante SALUDO de la config remplazando los saltos de línea por <br>
SALUDO = SALUDO.replace("\n", "<br>")

#: almacena la constante COTIZACION_MENSAJE_RECIBO de config con saltos de línea dados por <br>
COTIZACION_MENSAJE_RECIBO = COTIZACION_MENSAJE_RECIBO.replace("\n", "<br>")

#: almacena la constante COTIZACION_MENSAJE_FACTURA de config con saltos de línea dados por <br>
COTIZACION_MENSAJE_FACTURA = COTIZACION_MENSAJE_FACTURA.replace("\n", "<br>")

#: almacena la constante COTIZACION_MENSAJE_TRANSFERENCIA de config con saltos de línea dados por <br>
COTIZACION_MENSAJE_TRANSFERENCIA = COTIZACION_MENSAJE_TRANSFERENCIA.replace("\n", "<br>")

#: almacena la constante REPORTE_MENSAJE de config con saltos de línea dados por <br>
REPORTE_MENSAJE = REPORTE_MENSAJE.replace("\n", "<br>")

#: almacena la constante REQUEST_MENSAJE de config con saltos de línea dados por <br>
REQUEST_MENSAJE = REQUEST_MENSAJE.replace("\n", "<br>")

#: almacena la constante GESTOR_RECIBO_MENSAJE de config con saltos de línea dados por <br>
GESTOR_RECIBO_MENSAJE = GESTOR_RECIBO_MENSAJE.replace("\n", "<br>")

#: almacena la constante GESTOR_FACTURA_MENSAJE de config con saltos de línea dados por <br>
GESTOR_FACTURA_MENSAJE = GESTOR_FACTURA_MENSAJE.replace("\n", "<br>")

CORREO = smtplib.SMTP(SEND_SERVER,
                      SEND_PORT,
                      timeout=10)  #: instancia de SMTP

@export
def initCorreo():
    """ Inicializa la comunicación con el servidor SMTP, y autentica a el usuario
    """

    global CORREO
    try:
        CORREO.starttls()  # Puts connection to SMTP server in TLS mode
        CORREO.ehlo()  # Hostname to send, command defaults to the fully qualified domain name of the local host.
        CORREO.login(config.FROM, config.PASSWORD)
        print_log("[INFO] initCorreo() successful")
    except smtplib.SMTPException:
        print_log("[EXCEPT] smtplib.SMTPException in initCorreo()")

@export
def sendEmail(to: str, subject: str, text: str, attachments: Iterable = []):
    """ Función encargada de constuir un correo electrónico y enviarlo

    Parameters
    ----------
    to : str
        dirección de correo electrónico del destinatario
    subject : str
        asunto del correo electrónico
    text : str
        contenido escrito del correo electrónico
    attachments : Iterable
        nombres de los archivos a adjuntar

    Raises
    ------
    Exception
        en caso que luego de 5 intentos, no se pueda enviar el correo electrónico
    """

    global CORREO
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = config.FROM

    msg['To'] = to
    msg['Cc'] = config.FROM

    body = MIMEText(text, 'html')
    msg.attach(body)

    for item in attachments:
        if item != constants.REPORTE_INTERNO:
            name = os.path.join(constants.PDF_DIR, item + ".pdf")
            item = os.path.basename(item + ".pdf")
        else:
            name = item
        with open(name, "rb") as file:
            app = MIMEApplication(file.read())
            app.add_header('Content-Disposition', 'attachment', filename=item)
            msg.attach(app)

    to = [to, config.FROM]

    for i in range(5):
        try:
            CORREO.sendmail(config.FROM, to, msg.as_string())
            print_log("[INFO] Email {to} sent successfully".format(to=to))
            break
        except Exception as e:
            initCorreo()
            print_log("[EXCEPT] Email {to}:", e)
            if i == 4:
                raise(Exception("Could not send email."))


@export
def sendCotizacionRecibo(to: str, file_name: (str, Iterable), observaciones: str = ""):
    """ Función que envía por correo electrónico una cotización con forma de pago de recibo

    Parameters
    ----------
    to: str
        dirección de correo electrónico del destinatario
    file_name: str, Iterable
        nombre del archivo PDF de la cotización a enviar o nombres de los archivos PDF de las cotizaciones a enviar
    observaciones: str
        observaciones al correo electrónico

    Raises
    -------
    Exception
        en caso que luego de 5 intentos, no se pueda enviar el correo electrónico
    """

    if observaciones != "":
        observaciones = observaciones.replace("\n", "<br>") + 2 * "<br>"
    subject = COTIZACION_SUBJECT_RECIBO + ' - '
    subject += ' - '.join(file_name)
    sendEmail(to, subject, SALUDO + observaciones + COTIZACION_MENSAJE_RECIBO, file_name)


@export
def sendCotizacionTransferencia(to: str, file_name: (str, Iterable), observaciones: str = ""):
    """ Función que envía por correo electrónico una cotización con forma de pago de transferencia interna

    Parameters
    ----------
    to: str
        dirección de correo electrónico del destinatario
    file_name: str, Iterable
        nombre del archivo PDF de la cotización a enviar o nombres de los archivos PDF de las cotizaciones a enviar
    observaciones: str
        observaciones al correo electrónico

    Raises
    -------
    Exception
        en caso que luego de 5 intentos, no se pueda enviar el correo electrónico
    """

    if observaciones != "":
        observaciones = observaciones.replace("\n", "<br>") + 2 * "<br>"
    subject = COTIZACION_SUBJECT_RECIBO + ' - '
    subject += ' - '.join(file_name)
    sendEmail(to, subject,  SALUDO + observaciones + COTIZACION_MENSAJE_TRANSFERENCIA, file_name)


@export
def sendCotizacionFactura(to: str, file_name: (str, Iterable), observaciones: str = ""):
    """ Función que envía por correo electrónico una cotización con forma de pago de factura

    Parameters
    ----------
    to: str
        dirección de correo electrónico del destinatario
    file_name: str, Iterable
        nombre del archivo PDF de la cotización a enviar o nombres de los archivos PDF de las cotizaciones a enviar
    observaciones: str
        observaciones al correo electrónico

    Raises
    -------
    Exception
        en caso que luego de 5 intentos, no se pueda enviar el correo electrónico
    """

    if observaciones != "":
        observaciones = observaciones.replace("\n", "<br>") + 2 * "<br>"
    subject = COTIZACION_SUBJECT_RECIBO + ' - '
    subject += ' - '.join(file_name)
    sendEmail(to, subject, SALUDO + observaciones + COTIZACION_MENSAJE_FACTURA, file_name)


@export
def sendRegistro(to: str, file_name: str):
    """ Función que envía un correo electrónico con un reporte de los servicios usados hasta el momento

    Parameters
    ----------
    to : str
         dirección de correo electrónico del destinatario
    file_name : str
         nombre del archivo PDF del registro a enviar

    Raises
    -------
    Exception
        en caso que luego de 5 intentos, no se pueda enviar el correo electrónico
    """

    sendEmail(to, REPORTE_SUBJECT + " - %s" % file_name, SALUDO + REPORTE_MENSAJE, [file_name + "_Reporte"])


@export
def sendRequest(to: str):
    """ Función que envía un correo electrónico de solicitud de información

    Parameters
    ----------
    to : str
         dirección de correo electrónico del destinatario

    Raises
    -------
    Exception
        en caso que luego de 5 intentos, no se pueda enviar el correo electrónico
    """

    sendEmail(to, REQUEST_SUBJECT, SALUDO + REQUEST_MENSAJE)


@export
def sendGestorRecibo(file_name: str):
    """ Función que envía al Gestor de Recibos una solicitud para realizar un recibo

    Parameters
    ----------
    file_name : str
        nombre del archivo PDF de la cotización asociada a la solicitud de factura

    Raises
    -------
    Exception
        en caso que luego de 5 intentos, no se pueda enviar el correo electrónico
    """

    sendEmail(GESTOR_RECIBO_CORREO, GESTOR_RECIBO_SUBJECT + " - %s" % file_name, GESTOR_RECIBO_MENSAJE, [file_name])


@export
def sendGestorFactura(file_name: str, orden_name: str):
    """ Función que envía al Gestor de Facturas una solicitud para realizar una factura

    Parameters
    ----------
    file_name : str
        nombre del archivo PDF de la cotización asociada a la solicitud de factura
    orden_name : str
        nombre del archivo PDF de la orden de servicios con la que se solicita la factura

    Raises
    -------
    Exception
        en caso que luego de 5 intentos, no se pueda enviar el correo electrónico
    """

    orden_name = orden_name.split(".")[:-1]
    if type(orden_name) is list:
        orden_name = ".".join(orden_name)
    sendEmail(GESTOR_FACTURA_CORREO, GESTOR_FACTURA_SUBJECT + " - %s" % file_name,
              GESTOR_FACTURA_MENSAJE, [file_name, orden_name])


@export
def sendReporteExcel(to: str):
    """ Función encargada de enviar al destinatario el archivo de Excel asociado a un reporte

    Parameters
    ----------
    to : str
        dirección de correo electrónico del destinatario

    Raises
    -------
    Exception
        en caso que luego de 5 intentos, no se pueda enviar el correo electrónico
    """

    sendEmail(to, "Reporte semanal", "Microbill envia el reporte de cotizaciones", [constants.REPORTE_INTERNO])


@export
def correoTargetArgs(cotizaciones: Iterable, observaciones: str) -> (Callable, int):
    """ Función que se encarga de generar los parámetros para las funciones sendCotizacionTransferencia,
    sendCotizacionFactura, sendCotizacioinRecibo dependiendo de la información contenida en la forma de pago de las
    cotizaciones

    Parameters
    ----------
    cotizaciones : Cotizacion
        parámetro que contiene las cotizaciones a enviar
    observaciones : str
        observaciones a las cotizaciones

    Returns
    -------
        Callable: función de envío de correo electrónico
        tuple: parámetros para la función retornada en la primera posición
    """

    to = cotizaciones[0].getUsuario().getCorreo()
    tipo_pago = cotizaciones[0].getUsuario().getPago()
    pdfs = [cot.getFileName() for cot in cotizaciones]
    args = (to, pdfs, observaciones)
    if tipo_pago == "Transferencia interna":
        target = sendCotizacionTransferencia
    elif tipo_pago == "Factura":
        target = sendCotizacionFactura
    elif tipo_pago == "Recibo":
        target = sendCotizacionRecibo

    return target, args
