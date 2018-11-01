### Email reading
import os
import sys
sys.path.append(os.path.dirname(sys.executable))

import re
import codecs
import imaplib
import pandas as pd
from correo import *
from time import sleep
from io import StringIO
from threading import Thread
from datetime import datetime
from html2text import html2text #  lxml
from email import message_from_bytes

from objects import Cotizacion, Usuario, Servicio, getNumeroCotizacion

from config import READ_SERVER, READ_PORT, \
                READ_EMAIL_FOLDER, READ_REQUEST_DETAILS, \
                READ_LAST, READ_FIELDS, READ_FIELDS_DICT, \
                READ_NUMERIC_FIELDS, READ_SEARCH_FOR, \
                READ_EVERY_S

from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import QLabel, QWidget, QMainWindow, QHBoxLayout,\
        QGroupBox, QFormLayout, QSystemTrayIcon, QApplication, QMenu, \
        QStyleFactory, QMessageBox

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
            msg = self.toUTF8(msg)
            emails.append(msg)
        return emails

    def toUTF8(self, text):
        matches = re.findall(r"=\S\S=\S\S", text)
        replaces = [codecs.decode(match.replace("=", ""), "hex").decode('utf-8')\
                    for match in matches]
        for (match, replace) in zip(matches, replaces):
            text = text.replace(match, replace)
        return text

    def signIn(self):
        self.SERVER = imaplib.IMAP4_SSL(READ_SERVER, READ_PORT)
        self.SERVER.login(FROM, PASSWORD)
        self.SERVER.select(READ_EMAIL_FOLDER)

    def getIds(self):
        # ans, ids = self.SERVER.search(None, READ_SEARCH_FOR, '(UNSEEN)')
        ans, ids = self.SERVER.search(None, READ_SEARCH_FOR)
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
        cots = []
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
                cots.append(number)
        return cots

def runDaemon():
    er = EmailReader()
    while True:
        try:
            ids, new = er.retrieveNewRequests()
            if len(new):
                for (id, request) in zip(ids, new):
                    usuario = er.makeUsuario(request)
                    servicios = er.makeServicios(request)
                    cots = er.makeCotizaciones(usuario, servicios, request['tipo'])
                    er.markAsRead(id)
                    cots = ", ".join(cots)
                    print("Email sent to: %s(%s) with quote(s): %s"%(usuario.getNombre(), usuario.getCorreo(), cots))
        except Exception as e:
            print(e)
        sleep(READ_EVERY_S)

class MainWindow(QMainWindow):
    def __init__(self):
        super(QMainWindow, self).__init__()
        self.setWindowTitle("Microbill daemon")
        widget = QtWidgets.QWidget()
        widget.setAutoFillBackground(True)
        self.setCentralWidget(widget)
        p = widget.palette()
        p.setColor(widget.backgroundRole(), QtCore.Qt.black)
        widget.setPalette(p)

        self.main_layout = QtWidgets.QHBoxLayout(widget)
        self.main_layout.setContentsMargins(11, 11, 11, 11)
        self.main_layout.setSpacing(6)

        self.holder = QtWidgets.QPlainTextEdit()
        self.holder.setReadOnly(True)
        # self.holder.setEnabled(False)
        self.holder.setAutoFillBackground(True)
        self.format = '%s'
        self.holder.setStyleSheet("background-color: rgb(0, 0, 0);color: #FFFFFF;")

        self.main_layout.addWidget(self.holder)
        self.timer = QtCore.QTimer()
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.update)
        self.timer.start()

        self.time_format = "(%H:%M %d/%m/%Y)"
        print("Microbill daemon started")
        self.resize(600, 400)

    def update(self):
        current = sys.stdout.getvalue()
        sys.stdout.truncate(0)
        if current != "":
            date = datetime.now().strftime(self.time_format) + " "
            value = date + current
            self.holder.insertPlainText(value)

    def closeEvent(self, event):
        event.ignore()
        self.hide()

class SystemTrayIcon(QSystemTrayIcon):
    def __init__(self, icon, parent = None):
        QSystemTrayIcon.__init__(self, icon, parent)
        menu = QMenu(parent)
        openAction = menu.addAction("Open")
        exitAction = menu.addAction("Exit")
        self.setContextMenu(menu)

        openAction.triggered.connect(self.openMain)
        exitAction.triggered.connect(self.exit)
        self.activated.connect(self.systemIcon)

        self.main_window = MainWindow()
        self.openMain()

    def exit(self):
        QtCore.QCoreApplication.exit()

    def systemIcon(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.openMain()

    def openMain(self):
        self.main_window.show()

if __name__ == '__main__':
    sys.stdout = StringIO()

    app = QApplication(sys.argv)
    QApplication.setStyle(QStyleFactory.create('Fusion')) # <- Choose the style

    icon = QtGui.QIcon('icon.ico')
    app.setWindowIcon(icon)
    app.processEvents()

    thread = Thread(target = runDaemon)
    thread.setDaemon(True)
    thread.start()

    w = QWidget()
    trayIcon = SystemTrayIcon(icon, w)

    trayIcon.show()

    app.exec_()
