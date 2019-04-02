### Email reading
import os
import sys
import traceback
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

import constants

import objects
from objects import Cotizacion, Usuario, Servicio, getNumeroCotizacion

from config import READ_SERVER, READ_PORT, CORREOS_CONVENIOS, \
                READ_EMAIL_FOLDER, \
                READ_FIELDS_DICT, \
                READ_SEARCH_FOR, \
                READ_EVERY_S # READ_LAST, READ_REQUEST_DETAILS, READ_FIELDS_DICT

from PyQt5.QtWidgets import QLabel, QWidget, QMainWindow, QHBoxLayout,\
        QGroupBox, QFormLayout, QSystemTrayIcon, QApplication, QMenu, \
        QStyleFactory, QMessageBox, QWidget, QPlainTextEdit, QTextEdit

from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QTimer, Qt, QCoreApplication

REDIRECT_STDOUT = True
constants.DEBUG = True
LECTOR_DE_CORREOS = None

def imprimirExcepcion(e):
    if constants.DEBUG:
        traceback.print_exc()

class LectorCorreos(object):
    IMAP_EXCEPTIONS = (imaplib.IMAP4.error, imaplib.IMAP4.abort, imaplib.IMAP4.readonly)
    def __init__(self):
        self.servidor = None #: atributo para objeto de la clase IMAP4_SSL
        self.identificadores = [] #: atributo contenedor de los identificadores numericos de los correos
        self.contenido_correos = [] #: atributo contenedor del contenido de los correos asociados al identificador en la misma posicion
        self.registrarse()

    def registrarse(self):
        while True:
            try:
                self.servidor = imaplib.IMAP4_SSL(READ_SERVER, READ_PORT)
                self.servidor.login(FROM, PASSWORD)
                self.servidor.select(READ_EMAIL_FOLDER)
                break
            except self.IMAP_EXCEPTIONS as e:
                imprimirExcepcion(e)
                sleep(1)

    def darBloqueTexto(self, mensaje):
        type = mensaje.get_content_maintype()
        if type == 'multipart':
            for parte in mensaje.get_payload():
                if parte.get_content_maintype() == 'text':
                    return parte.get_payload()
        elif type == 'text':
            return mensaje.get_payload()

    def darCorreos(self, identificadores):
        correos = []
        for i in identificadores:
            typ, data = self.servidor.fetch(i, '(RFC822)')
            self.marcarNoLeido(i) # fetch marca como leido
            txt = message_from_bytes(data[0][1])
            msg = self.darBloqueTexto(txt)
            msg = self.decodificar(msg)
            correos.append(msg)
        self.contenido_correos += correos
        self.contenido_correos = list(set(self.contenido_correos))
        return self.contenido_correos

    def decodificar(self, texto):
        # decodificar acentos
        buscar = re.findall(r"=\S\S=\S\S", texto)
        remplazos = [codecs.decode(patron.replace("=", ""), "hex").decode('utf-8')\
                    for patron in buscar]
        for (patron, cambio) in zip(buscar, remplazos):
            texto = texto.replace(patron, cambio)

        return texto.replace('=\r\n', '')

    def darIdentificadores(self):
        for identificador in READ_SEARCH_FOR:
            ans, ids = self.servidor.search(None, identificador, '(UNSEEN)')
            self.identificadores += ids[0].split()
        self.identificadores = list(set(self.identificadores))
        return self.identificadores

    def marcarLeido(self, identificador):
        self.servidor.store(identificador, '+FLAGS', '\Seen')
        pos = self.identificadores.index(identificador)
        del self.identificadores[pos]
        del self.contenido_correos[pos]

    def marcarNoLeido(self, identificador):
        self.servidor.store(identificador, '-FLAGS', '\Seen')

    def darNuevasSolicitudes(self):
        self.registrarse()
        identificadores = self.darIdentificadores()
        correos = self.darCorreos(identificadores)
        return identificadores, correos

class CorreoAgendo(object):
    PATRON_CORREO = re.compile('\((.*?)\)')
    PATRON_OBSERVACIONES = re.compile('(?s)The following comment was left by the requesting user  \n  \n(.*?)-----------------------------')
    REMPLAZOS_NORMALIZACION = ['*', "=\r\n", ":", "= "] #: caracteres innecesarios a remover en contenido_texto
    def __init__(self, identificador, contenido):
        self.tabla = None
        self.identificador = identificador
        self.contenido_completo = contenido
        self.contenido_texto = self.normalizar(contenido)
        for atributo in READ_FIELDS_DICT.keys():
            setattr(self, atributo, None)
            self.darAtributo(atributo)

        self.correo = re.findall(self.PATRON_CORREO, self.contenido_texto)[0]
        self.darAtributo('nombre')
        self.nombre = self.nombre.replace(" (%s)"%self.correo, '')
        try:
            self.observaciones = re.findall(self.PATRON_OBSERVACIONES, self.contenido_texto)[0]
            self.observaciones = self.observaciones.rstrip('\\').rstrip(' ').rstrip('\n')
        except IndexError:
            self.observaciones = ''

    def __darServicios__(self):
        if self.tabla is None:
            self.tabla = pd.read_html(self.contenido_completo)[0]
        return self.tabla

    def darDiccionario(self):
        return self.__dict__

    def darAtributo(self, atributo):
        try:
            if atributo == "tabla":
                return self.__darServicios__()

            variable = getattr(self, atributo)
            if variable != None:
                return variable

            buscar = READ_FIELDS_DICT[atributo]
            p_inicial = self.contenido_texto.find(buscar)
            p_final = self.contenido_texto[p_inicial:].find('\n')
            if (p_inicial >= 0) | (p_final >= 0):
                fragmento = self.contenido_texto[p_inicial + len(buscar): p_inicial + p_final]
                fragmento = fragmento.rstrip(' ').rstrip('.').lstrip(' ')
            else:
                fragmento = ''

            setattr(self, atributo, fragmento)
            return fragmento
        except AttributeError as e:
            imprimirExcepcion(e)

    def normalizar(self, texto):
        texto = html2text(texto)
        for remplazo in self.REMPLAZOS_NORMALIZACION:
            texto = texto.replace(remplazo, '')
        return texto.lstrip().rstrip()

    def __repr__(self):
        txt = ['Correo: %s (%d)'%(self.correo, int(self.identificador))]
        for atributo in READ_FIELDS_DICT.keys():
            txt += ['\t%s: %s'%(atributo, str(getattr(self, atributo)))]
        txt = '\n'.join(txt)
        return txt

def obtenerTipoUsuario(correo):
    convenios = sum([1 for convenio in CORREOS_CONVENIOS if convenio in correo])
    if convenios > 0:
        suma = sum(constants.INDEPENDIENTES_DF['Correo'] == correo)
        if suma:
            return 'Independiente'
        return 'Interno'
    elif '.edu.' in correo: return 'Académico'
    else: return 'Industria'

def crearUsuario(correo_agendo):
    dict = correo_agendo.darDiccionario()
    remitente = correo_agendo.darAtributo('correo')
    dict['interno'] = obtenerTipoUsuario(remitente)
    return Usuario(**dict)

def crearServicios(tabla_agendo, interno):
    campos = tabla_agendo[['Subclass', 'Item', 'Units']]

    servicios = {}
    for servicio in constants.EQUIPOS:
        servicios[servicio] = []
    for i in range(len(campos)):
        fila = campos.loc[i]
        hoja_excel = fila['Subclass']
        precios_df = constants.DAEMON_DF[hoja_excel]
        iguales = precios_df[precios_df['Item'] == fila['Item']]
        if len(iguales):
            codigo = str(iguales['Código'].values[0])
            equipo = iguales['Equipo'].values[0]
            cantidad = float(fila['Units'])
            servicio = Servicio(equipo.split("_")[1] + codigo, interno, cantidad)
            servicios[equipo].append(servicio)
    return servicios

def solicitudMicrobill(correo_agendo):
    global LECTOR_DE_CORREOS
    usuario = crearUsuario(correo_agendo)
    try: muestra = correo_agendo.darAtributo('muestra')
    except AttributeError: muestra = ""
    servicios = crearServicios(correo_agendo.darAtributo('tabla'), usuario.getInterno())
    cotizaciones = []
    numeros = []
    for equipo in servicios.keys():
        objetos = servicios[equipo]
        if len(objetos):
            numero = getNumeroCotizacion(equipo)
            cot = Cotizacion(numero, usuario, objetos, muestra)
            cot.setObservacionPDF(correo_agendo.darAtributo('observaciones'))
            cotizaciones.append(cot)
            numeros.append(numero)

    return cotizaciones, numeros

def enviarCorreos(lector_correos, correo_agendo, cotizaciones, numeros):
    pago = correo_agendo.darAtributo('pago')
    identificador = correo_agendo.darAtributo('identificador')
    correo = correo_agendo.darAtributo('correo')
    try:
        for cotizacion in cotizaciones:
            cotizacion.save(to_cotizacion = False)
        if pago == "Transferencia interna":
            sendCotizacionTransferencia(correo, numeros)
        elif pago == "Factura":
            sendCotizacionFactura(correo, numeros)
        elif pago == "Recibo":
            sendCotizacionRecibo(correo, numeros)

        for cotizacion in cotizaciones:
            cotizacion.save(to_pdf = False)
        lector_correos.marcarLeido(identificador)
        return numeros
    except Exception as e:
        imprimirExcepcion(e)
        return []

class MainWindow(QMainWindow):
    def __init__(self):
        super(QMainWindow, self).__init__()
        self.setWindowTitle("Microbill daemon")
        widget = QWidget()
        widget.setAutoFillBackground(True)
        self.setCentralWidget(widget)
        p = widget.palette()
        p.setColor(widget.backgroundRole(), Qt.black)
        widget.setPalette(p)

        self.main_layout = QHBoxLayout(widget)
        self.main_layout.setContentsMargins(11, 11, 11, 11)
        self.main_layout.setSpacing(6)

        # self.holder = QPlainTextEdit()
        self.holder = QTextEdit()
        self.holder.setReadOnly(True)
        # self.holder.setEnabled(False)
        self.holder.setAutoFillBackground(True)
        self.holder.setStyleSheet("background-color: rgb(0, 0, 0);color: #FFFFFF;")

        self.main_layout.addWidget(self.holder)
        self.timer = QTimer()
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.update)
        self.timer.start()

        self.time_format = "(%H:%M %d/%m/%Y)"
        print("Microbill daemon started")
        self.resize(600, 400)

    def update(self):
        if REDIRECT_STDOUT:
            current = sys.stdout.getvalue()
            sys.stdout.truncate(0)
            if current != "":
                date = datetime.now().strftime(self.time_format) + " "
                current = current.replace('\n', '<br>')
                if 'Traceback' in current:
                    current += '<br>' * 4
                    color = 'red'
                else:
                    color = 'green'
                self.holder.insertHtml('%s<font color="%s">%s</font>'%(date, color, current))

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
        QCoreApplication.exit()

    def systemIcon(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.openMain()

    def openMain(self):
        self.main_window.show()

def runDaemon():
    lector = None
    i = 0
    while True:
        cli, reg = objects.readDataFrames()
        if (cli is None):
            cli = objects.CLIENTES_DATAFRAME
        if (reg is None):
            reg = objects.REGISTRO_DATAFRAME
        if (not cli.equals(objects.CLIENTES_DATAFRAME)) | (not reg.equals(objects.REGISTRO_DATAFRAME)):
            objects.CLIENTES_DATAFRAME = cli
            objects.REGISTRO_DATAFRAME = reg
        if i%10 == 0:
            try:
                if lector == None:
                    lector = LectorCorreos()
                ids, new = lector.darNuevasSolicitudes()
                if len(new):
                    for (id, request) in zip(ids, new):
                        correo = CorreoAgendo(id, request)
                        cotizaciones, numeros = solicitudMicrobill(correo)
                        numeros = enviarCorreos(lector, correo, cotizaciones, numeros)
                        nombre = correo.darAtributo('nombre')
                        correo = correo.darAtributo('correo')
                        if numeros:
                            numeros = ', '.join(numeros)
                            print("Correo enviado a: %s (%s) con las cotizaciones: %s"%(nombre, correo, numeros))
            except Exception as e:
                imprimirExcepcion(e)
        i += 1
        sleep(READ_EVERY_S / 10)

if __name__ == '__main__':
    thread = Thread(target = runDaemon)
    if REDIRECT_STDOUT:
        thread.setDaemon(True)
    thread.start()

    if REDIRECT_STDOUT:
        sys.stdout = StringIO()
        sys.stderr = sys.stdout

        app = QApplication(sys.argv)
        QApplication.setStyle(QStyleFactory.create('Fusion')) # <- Choose the style

        icon = QIcon('icon.ico')
        app.setWindowIcon(icon)
        app.processEvents()

        w = QWidget()
        trayIcon = SystemTrayIcon(icon, w)

        trayIcon.show()

        app.exec_()
