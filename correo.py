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

### Email reading
import imaplib
from time import sleep
import pandas as pd
from html2text import html2text #  lxml
from email import message_from_bytes
from config import READ_SERVER, READ_PORT, \
                READ_EMAIL_FOLDER, READ_REQUEST_DETAILS, \
                READ_LAST, READ_FIELDS, READ_FIELDS_DICT, \
                READ_NUMERIC_FIELDS, READ_SEARCH_FOR, \
                READ_EVERY_S

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
    # try:
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
            CORREO.sendmail(FROM, to, msg.as_string())
            break
        except:
            try: initCorreo()
            except Exception as e: print(e)
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

class EmailReader(object):
    def __init__(self):
        self.SERVER = None
        self.signIn()

    def getTextBlock(self, msg):
        type = msg.get_content_maintype()
        if type == 'multipart':
            for part in msg.get_payload():
                if part.get_content_maintype() == 'text':
                    return part.get_payload()
        elif type == 'text':
            return msg.get_payload()

    def cleanSpaces(self, data):
        for key in data.keys():
            if key != "Table": data[key] = data[key].lstrip().rstrip()

        for key in READ_NUMERIC_FIELDS:
            try: data[key] = data[key].replace(" ", "")
            except KeyError: pass
        return data

    def parseEmail(self, msg):
        msg = msg.replace("=\r\n", "").replace(":", "")#.replace("=", "")
        txt = html2text(msg).replace("*", "")
        table = pd.read_html(msg)[0]

        txt_1 = txt.find(READ_REQUEST_DETAILS)
        txt_2 = txt.find(READ_FIELDS)
        end = txt.find(READ_LAST)

        txt1 = txt[:txt_1]
        txt2 = txt[txt_2 + len(READ_FIELDS): end]

        txt = txt1 + txt2

        lines = txt.split("\n")
        lines = [line for line in lines if line != ""]

        data = {'Table': table}
        # data = {}
        for line in lines:
            for key in READ_FIELDS_DICT.keys():
                field = READ_FIELDS_DICT[key]
                if field in line:
                    index = line.find(field) + len(field)
                    if key == "nombre":
                        index2 = line.find("(")
                        data[key] = line[index : index2 -1]
                        data["correo"] = line[index2 + 1 : line.find(")")]
                    else:
                        data[key] = line[index : -1].replace(".", "")
        data = self.cleanSpaces(data)
        return data

    def getEmails(self, ids):
        emails = []
        for i in ids:
            typ, data = self.SERVER.fetch(i, '(RFC822)')
            data = data[0][1]
            txt = message_from_bytes(data)
            msg = self.getTextBlock(txt)
            emails.append(msg)
        return emails

    def signIn(self):
        self.SERVER = imaplib.IMAP4_SSL(READ_SERVER, READ_PORT)
        self.SERVER.login(FROM, PASSWORD)
        self.SERVER.select(READ_EMAIL_FOLDER)

    def getIds(self):
        ans, ids = self.SERVER.search(None, READ_SEARCH_FOR, '(UNSEEN)')
        ids = ids[0].split()
        return ids

    def markAsRead(self, id):
        self.SERVER.store(id, '+FLAGS', '\Seen')

    def retrieveNewRequests(self):
        self.signIn()
        ids = self.getIds()
        emails = self.getEmails(ids)

        return ids, [self.parseEmail(email) for email in emails]

    def makeUsuario(self, dicc):
        temp = dicc.copy()
        try:
            del temp["Table"]
            del temp['tipo']
            del temp['id']
        except KeyError: pass

        # pago = temp['interno']
        # if pago == "Universidad de los Andes":
        #     temp['pago'] = "Transferencia interna"
        # elif pago == "External Institution":
        #     temp['pago'] = "Factura"
        # elif pago == "Personal funds":
        #     temp['pago'] = "Recibo"

        # if temp['interno'] == "Universidad de los Andes":
        #     temp['interno'] = "Interno"
        # else:
        #     temp['interno'] = "Externo"
        # temp["institucion"] = "Universidad de los Andes"
        return Usuario(**temp)

    def makeServicios(self, dicc):
        table = dicc['Table'][['Subclass', 'Item', 'Units', 'Unit price (COP)']]

        servicios = {}
        for servicio in constants.EQUIPOS:
            servicios[servicio] = []

        for i in range(len(table)):
            row = table.loc[i]
            sheet = row['Subclass']
            df = constants.DAEMON_DF[sheet]
            df = df[df['Item'] == row['Item']]
            if len(df):
                codigo = df['Código'].values[0]
                equipo = df['Equipo'].values[0]
                cantidad = float(row['Units'])
                interno = "Interno"
                if int(row['Unit price (COP)']) != int(df['Interno']):
                    interno = "Externo"
                servicio = Servicio(equipo, str(codigo), interno, cantidad)
                servicios[equipo].append(servicio)
        return servicios

    def makeCotizaciones(self, usuario, servicios, tipo):
        for key in servicios.keys():
            servicios_ = servicios[key]
            if len(servicios_):
                l = key[0]
                equipo = servicios_[0].getEquipo()
                number = getNumeroCotizacion(equipo)
                usuario.setInterno(servicios_[0].getInterno())
                cot = Cotizacion(number, usuario, servicios_, tipo)
                cot.save()

                pago = usuario.getPago()
                if pago == "Transferencia interna":
                    sendCotizacionTransferencia(usuario.getCorreo(), number)
                elif pago == "Factura":
                    sendCotizacionFactura(usuario.getCorreo(), number)
                elif pago == "Recibo":
                    sendCotizacionRecibo(usuario.getCorreo(), number)

def runDaemon():
    er = EmailReader()

    i = 1000
    while True:
        try:
            ids, new = er.retrieveNewRequests()
            if len(new):
                for (id, request) in zip(ids, new):
                    usuario = er.makeUsuario(request)
                    servicios = er.makeServicios(request)
                    er.makeCotizaciones(usuario, servicios, request['tipo'])
                    er.markAsRead(id)
        except Exception as e:
            print(e)
        sleep(READ_EVERY_S)

if __name__ == '__main__':
    import sys
    from io import StringIO
    from threading import Thread
    from objects import Cotizacion, Usuario, Servicio, getNumeroCotizacion
    sys.stdout = StringIO()

    thread = Thread(target = runDaemon, daemon = True)
    thread.start()
    while True:
        value = sys.stdout.getvalue()
        with open("test.txt", "a") as file:
            file.write(value)
        sleep(READ_EVERY_S)
