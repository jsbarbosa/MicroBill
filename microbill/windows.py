import os
import sys
from . import config
import traceback
import numpy as np
import pandas as pd
from copy import copy
from time import sleep
from datetime import datetime
from PyQt5 import QtCore, QtWidgets, QtGui

from . import correo, objects, constants
from .exceptions import *

import psutil
from subprocess import Popen
from threading import Thread

# from PyQt5.QtWebEngine import QtWebEngine
# from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings
from PyQt5.QtCore import QUrl

config.ADMINS = [""] + config.ADMINS

class SubWindow(QtWidgets.QMdiSubWindow):
    def __init__(self, parent = None):
        super(SubWindow, self).__init__(parent)
        self.parent = parent
        self.is_closed = True
        # self.setWindowIcon(constants.ICON)
        self.setMinimumSize(400, 400)
        # self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

    def closeEvent(self, evnt):
        evnt.accept()
        self.hide()
        self.is_closed = True

    def show(self):
        self.is_closed = False
        QtWidgets.QMdiSubWindow.show(self)

class Table(QtWidgets.QTableWidget):
    HEADER = ['Código', 'Descripción', 'Cantidad', 'Valor Unitario', 'Valor Total']
    def __init__(self, parent, rows = 25, cols = 5):
        super(Table, self).__init__(rows, cols)

        self.parent = parent
        self.n_rows = rows
        self.n_cols = cols
        self.setHorizontalHeaderLabels(self.HEADER)
        self.resizeRowsToContents()

        header = self.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)
        self.clean()

        self.cellChanged.connect(self.handler)

    def agregarServicio(self, codigo):
        for i in range(self.n_rows):
            if self.item(i, 0).text() == "": break

        self.item(i, 0).setText(codigo)

    def removeServicio(self):
        table_codigos = self.getCodigos()
        regis_codigos = self.parent.getCodigosPrefix()
        for (i, codigo) in enumerate(regis_codigos):
            if not codigo in table_codigos: self.parent.removeServicio(i)

    def handler(self, row, col):
        self.blockSignals(True)
        item = self.item(row, col)
        try:
            cod = self.item(row, 0).text()
            if col == 0:
                self.removeServicio()
                if cod == "":
                    self.item(row, 1).setText("")
                    self.item(row, 2).setText("")
                    self.item(row, 3).setText("")
                    self.item(row, 4).setText("")

                else:
                    interno = self.parent.getInterno()
                    try:
                        n = round(float(self.item(row, 2).text()), 1)
                        total = 0
                    except:
                        try:
                            total = int(self.item(row, 4).text().replace(",", ""))
                            n = 1
                        except:
                            n = 1
                            total = 0
                            self.item(row, 2).setText("1.0")

                    try:
                        servicio = objects.Servicio(codigo = cod, interno = interno, cantidad = n)
                        self.parent.addServicio(servicio)
                    except Exception as e:
                        self.item(row, 0).setText("")
                        self.item(row, 1).setText("")
                        self.item(row, 2).setText("")
                        self.item(row, 3).setText("")
                        self.item(row, 4).setText("")
                        raise(Exception("Código inválido."))

                    desc = servicio.getDescripcion()
                    valor = servicio.getValorUnitario()

                    if (total != 0) and (n == 1):
                        n = total / servicio.getValorUnitario()
                        n = np.ceil(10 * n) / 10
                        servicio.setCantidad(n)
                        servicio.setValorTotal(total)
                        self.item(row, 2).setText("%.1f"%n)
                    else:
                        total = servicio.getValorTotal()

                    self.item(row, 1).setText(desc)
                    self.item(row, 3).setText("{:,}".format(valor))
                    self.item(row, 4).setText("{:,}".format(total))

            if col == 2:
                try: n = round(float(self.item(row, 2).text()), 1)
                except: raise(Exception("Cantidad inválida.")); self.item(row, 2).setText("")

                self.item(row, 2).setText("%.1f"%n)

                if cod != "":
                    servicio = self.parent.getServicio(cod)
                    servicio.setCantidad(n)
                    total = servicio.getValorTotal()

                    self.item(row, 4).setText("{:,}".format(total))

            if col == 4:
                try: total = int(self.item(row, 4).text())
                except: raise(Exception("Valor total inválido.")); self.item(row, 4).setText("")

                self.item(row, 4).setText("{:,}".format(total))

                if cod != "":
                    servicio = self.parent.getServicio(cod)
                    n = total / servicio.getValorUnitario()

                    n = np.ceil(10 * n) / 10

                    servicio.setCantidad(n)
                    servicio.setValorTotal(total)

                    self.item(row, 2).setText("%.1f"%n)

            self.parent.setTotal()
        except Exception as e:
            self.parent.errorWindow(e)
        self.blockSignals(False)

    def updateInterno(self):
        self.blockSignals(True)

        cods = self.getCodigos()
        pos = np.where(np.array(cods) != "")[0]

        for i in pos:
            cod = cods[i]
            servicio = self.parent.getServicio(cod)
            self.item(i, 2).setText("%.1f"%servicio.getCantidad())
            self.item(i, 3).setText("{:,}".format(servicio.getValorUnitario()))
            self.item(i, 4).setText("{:,}".format(servicio.getValorTotal()))

        self.blockSignals(False)

    def setFromCotizacion(self):
        servicios = self.parent.getServicios()
        self.blockSignals(True)
        for (row, servicio) in enumerate(servicios):
            if not servicio.isAgregado():
                try:
                    self.item(row, 0).setText(servicio.getCodigoPrefix())
                except (IncompatibleError, AttributeError):
                    self.item(row, 0).setText(servicio.getCodigo())
                    raise(IncompatibleError)
                self.item(row, 1).setText(servicio.getDescripcion())
                self.item(row, 2).setText("%.1f"%servicio.getCantidad())
                self.item(row, 3).setText("{:,}".format(servicio.getValorUnitario()))
                self.item(row, 4).setText("{:,}".format(servicio.getValorTotal()))
        self.blockSignals(False)

    def getCodigos(self):
        return [self.item(i, 0).text() for i in range(self.n_rows)]

    def clean(self):
        self.blockSignals(True)
        for r in range(self.n_rows):
            for c in range(self.n_cols):
                item = QtWidgets.QTableWidgetItem("")
                self.setItem(r, c, item)
        self.readOnly()
        self.blockSignals(False)

    def readOnly(self):
        flags = QtCore.Qt.ItemIsEditable
        for r in range(self.n_rows):
            for c in [1, 3]:
                item = QtWidgets.QTableWidgetItem("")
                item.setFlags(flags)
                self.setItem(r, c, item)

class AutoLineEdit(QtWidgets.QLineEdit):
    AUTOCOMPLETE = ["Nombre", "Correo", "Documento", "Teléfono", "Cotización"]
    def __init__(self, target, parent, autochange = True):
        super(AutoLineEdit, self).__init__()
        self.target = target
        self.parent = parent
        self.model = QtCore.QStringListModel()
        completer = QtWidgets.QCompleter()
        completer.setCaseSensitivity(False)
        completer.setModelSorting(0)
        completer.setModel(self.model)
        self.setCompleter(completer)
        self.update()

    def event(self, event):
        if event.type() == QtCore.QEvent.KeyPress and event.key() == QtCore.Qt.Key_Tab:
            try:
                self.parent.changeAutocompletar()
            except:
                pass
            return False
        return QtWidgets.QWidget.event(self, event)

    def update(self):
        if type(self.parent) is CotizacionWindow:
            dataframe = objects.CLIENTES_DATAFRAME
            order = 1
        else:
            dataframe = objects.REGISTRO_DATAFRAME
            order = -1
        data = list(set(dataframe[self.target].values.astype('str')))
        data = sorted(data)[::order]
        self.model.setStringList(data)

class ChangeCotizacion(QtWidgets.QDialog):
    FIELDS = ["Cotización", "Fecha", "Nombre", "Correo", "Equipo", "Valor"]
    WIDGETS = ["cotizacion", "fecha", "nombre", "correo", "equipo", "valor"]

    def __init__(self, parent):
        super(ChangeCotizacion, self).__init__()
        self.setWindowTitle("Modificar cotización")
        self.parent = parent
        self.setModal(True)

        self.layout = QtWidgets.QVBoxLayout()
        self.setLayout(self.layout)

        self.form = QtWidgets.QFrame()
        self.form_layout = QtWidgets.QFormLayout(self.form)

        cotizacion_label = QtWidgets.QLabel("Cotización:")
        self.cotizacion_widget = AutoLineEdit("Cotización", self)

        fecha_label = QtWidgets.QLabel("Fecha:")
        self.fecha_widget = QtWidgets.QLabel("")

        nombre_label = QtWidgets.QLabel("Nombre:")
        self.nombre_widget = QtWidgets.QLabel("")

        correo_label = QtWidgets.QLabel("Correo:")
        self.correo_widget = QtWidgets.QLabel("")

        equipo_label = QtWidgets.QLabel("Equipo:")
        self.equipo_widget = QtWidgets.QLabel("")

        valor_label = QtWidgets.QLabel("Valor:")
        self.valor_widget = QtWidgets.QLabel("")

        self.form_layout.addRow(fecha_label, self.fecha_widget)
        self.form_layout.addRow(nombre_label, self.nombre_widget)
        self.form_layout.addRow(correo_label, self.correo_widget)
        self.form_layout.addRow(equipo_label, self.equipo_widget)
        self.form_layout.addRow(valor_label, self.valor_widget)
        self.form_layout.addRow(cotizacion_label, self.cotizacion_widget)

        self.layout.addWidget(self.form)

        self.buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel,
            QtCore.Qt.Horizontal, self)

        self.layout.addWidget(self.buttons)
        self.buttons.accepted.connect(self.accept2)
        self.buttons.rejected.connect(self.reject)

        self.cotizacion_widget.textChanged.connect(self.autoCompletar)

    def autoCompletar(self, text):
        if text != "":
            registro = objects.REGISTRO_DATAFRAME[objects.REGISTRO_DATAFRAME["Cotización"] == text]
            if len(registro):
                for (field, widgetT) in zip(self.FIELDS, self.WIDGETS):
                        val = str(registro[field].values[0])
                        if val == "nan": val = ""
                        widget = eval("self.%s_widget"%widgetT)
                        widget.setText(val)

    def accept2(self):
        self.parent.loadCotizacion(self.cotizacion_widget.text())
        self.accept()

class CorreoDialog(QtWidgets.QDialog):
    def __init__(self, args, target):
        super(CorreoDialog, self).__init__()
        self.setWindowTitle("Enviando correo...")

        self.args = args

        self.layout = QtWidgets.QVBoxLayout()
        self.setLayout(self.layout)
        label = "Enviando correo a: \t<b>%s</b>"%args[0]

        self.layout.addWidget(QtWidgets.QLabel(label))
        self.resize(200, 100)

        self.setModal(True)

        self.thread = Thread(target = self.sendCorreo, args=(target,))
        self.thread.setDaemon(True)

        self.finished = False
        self.exception = None

    def closeEvent(self, event):
        if self.finished:
            self.thread = None
            # sleep(2.5)
            event.accept()
        else:
            event.ignore()

    def sendCorreo(self, func):
        try:
            func(*self.args)
        except Exception as e:
            self.exception = Exception("Email error: " + str(e))
            print("Error:", self.exception)
        self.finished = True
        self.close()

    def start(self):
        self.thread.start()

class CodigosDialog(QtWidgets.QDialog):
    def __init__(self, parent):
        super(CodigosDialog, self).__init__()

        self.parent = parent
        self.setWindowTitle("Ver códigos")
        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(0)
        self.setLayout(layout)

        self.tabs = QtWidgets.QTabWidget(self)
        for equipo in constants.EQUIPOS:
            tab = QtWidgets.QWidget()
            lout = QtWidgets.QVBoxLayout(tab)
            lout.setContentsMargins(0, 6, 0, 0)
            lout.setSpacing(0)
            table = QtWidgets.QTableView()
            table.doubleClicked.connect(self.doubleClick)

            lout.addWidget(table)

            self.tabs.addTab(tab, equipo)
            setattr(self, "tab_%s" % equipo, tab)
            setattr(self, "table_%s" % equipo, table)

            self.setModel(equipo)

        layout.addWidget(self.tabs)
        self.resize(600, 400)

    def doubleClick(self, modelIndex):
        row = modelIndex.row()
        tab = self.tabs.currentIndex()
        equipo = constants.EQUIPOS[tab]
        table = getattr(self, "table_%s" % equipo)

        df = table.model().dataframe
        codigo = df.iloc[row]['Código']
        codigo = equipo.split("_")[-1] + codigo
        self.parent.agregarDesdeCodigo(codigo)

    def setModel(self, equipo):
        table = getattr(self, "table_%s" % equipo)
        df = getattr(constants, equipo)
        table.setModel(PandasModel(df, checkbox = False))

        table.setSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)

        table.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        for i in range(2, len(df.keys())):
            header.setSectionResizeMode(i, QtWidgets.QHeaderView.ResizeToContents)

class CotizacionWindow(SubWindow):
    IGNORE = ["proyecto", "codigo"]
    FIELDS = ["Nombre", "Correo", "Teléfono", "Institución", "Documento", "Dirección", "Ciudad", "Interno", "Responsable", "Proyecto", "Código", "Muestra"]
    WIDGETS = ["nombre", "correo", "telefono", "institucion", "documento", "direccion", "ciudad", "interno", "responsable", "proyecto", "codigo", "muestra"]

    AUTOCOMPLETE_FIELDS = ["Nombre", "Correo", "Documento", "Teléfono"]
    AUTOCOMPLETE_WIDGETS = ["nombre", "correo", "documento", "telefono"]
    def __init__(self, parent = None):
        super(CotizacionWindow, self).__init__(parent)
        self.setWindowTitle("Cotización")

        wid = QtWidgets.QWidget(self)
        self.setWidget(wid)

        self.ver_dialog = CodigosDialog(self)
        self.verticalLayout = QtWidgets.QVBoxLayout(wid)

        self.verticalLayout.setContentsMargins(6, 0, 6, 0)
        self.verticalLayout.setSpacing(4)

        self.cotizacion_frame = QtWidgets.QFrame()
        self.form_frame = QtWidgets.QFrame()
        self.button_frame = QtWidgets.QFrame()
        self.total_frame = QtWidgets.QFrame()
        self.observaciones_frame = QtWidgets.QGroupBox("Observaciones")

        self.form_frame_layout = QtWidgets.QGridLayout(self.form_frame)
        self.form_frame_layout.setContentsMargins(0, 0, 0, 0)
        self.form_frame_layout.setSpacing(2)

        self.autocompletar_widget = QtWidgets.QCheckBox("Autocompletar")
        self.autocompletar_widget.setChecked(True)

        nombre_label = QtWidgets.QLabel("Nombre:")
        self.nombre_widget = AutoLineEdit("Nombre", self)
        correo_label = QtWidgets.QLabel("Correo:")
        self.correo_widget = AutoLineEdit("Correo", self)

        institucion_label = QtWidgets.QLabel("Institución/Empresa:")
        self.institucion_widget = AutoLineEdit("Institución", self)
        documento_label = QtWidgets.QLabel("Nit/CC:")
        self.documento_widget = AutoLineEdit("Documento", self)

        direccion_label = QtWidgets.QLabel("Dirección:")
        self.direccion_widget = AutoLineEdit("Dirección", self)
        ciudad_label = QtWidgets.QLabel("Ciudad:")
        self.ciudad_widget = AutoLineEdit("Ciudad", self)

        telefono_label = QtWidgets.QLabel("Teléfono:")
        self.telefono_widget = AutoLineEdit("Teléfono", self)

        interno_label = QtWidgets.QLabel("Interno:")
        self.interno_widget = QtWidgets.QComboBox()
        self.interno_widget.addItems(constants.PRICES_DIVISION)
        # QCheckBox("Interno")
        responsable_label = QtWidgets.QLabel("Responsable:")
        self.responsable_widget = AutoLineEdit("Responsable", self)

        proyecto_label = QtWidgets.QLabel("Nombre proyecto:")
        self.proyecto_widget = QtWidgets.QLineEdit()
        codigo_label = QtWidgets.QLabel("Código proyecto:")
        self.codigo_widget = QtWidgets.QLineEdit()

        muestra_label = QtWidgets.QLabel("Tipo de muestras:")
        self.muestra_widget = QtWidgets.QLineEdit()

        # equipo_label = QtWidgets.QLabel("Equipo:")
        # self.equipo_widget = QtWidgets.QComboBox()
        # self.equipo_widget.addItems(constants.EQUIPOS)

        end_label = QtWidgets.QLabel("Documento final")
        self.pago_widget = QtWidgets.QComboBox()
        self.pago_widget.addItems(constants.DOCUMENTOS_FINALES)

        self.form_frame_layout.addWidget(nombre_label, 0, 0)
        self.form_frame_layout.addWidget(self.nombre_widget, 0, 1)
        self.form_frame_layout.addWidget(correo_label, 0, 2)
        self.form_frame_layout.addWidget(self.correo_widget, 0, 3)

        self.form_frame_layout.addWidget(institucion_label, 1, 0)
        self.form_frame_layout.addWidget(self.institucion_widget, 1, 1)
        self.form_frame_layout.addWidget(documento_label, 1, 2)
        self.form_frame_layout.addWidget(self.documento_widget, 1, 3)

        self.form_frame_layout.addWidget(direccion_label, 2, 0)
        self.form_frame_layout.addWidget(self.direccion_widget, 2, 1)
        self.form_frame_layout.addWidget(ciudad_label, 2, 2)
        self.form_frame_layout.addWidget(self.ciudad_widget, 2, 3)

        self.form_frame_layout.addWidget(telefono_label, 3, 0)
        self.form_frame_layout.addWidget(self.telefono_widget, 3, 1)

        self.form_frame_layout.addWidget(responsable_label, 4, 0)
        self.form_frame_layout.addWidget(self.responsable_widget, 4, 1)
        self.form_frame_layout.addWidget(interno_label, 4, 2)
        self.form_frame_layout.addWidget(self.interno_widget, 4, 3)

        self.form_frame_layout.addWidget(proyecto_label, 5, 0)
        self.form_frame_layout.addWidget(self.proyecto_widget, 5, 1)
        self.form_frame_layout.addWidget(codigo_label, 5, 2)
        self.form_frame_layout.addWidget(self.codigo_widget, 5, 3)

        self.form_frame_layout.addWidget(muestra_label, 6, 0)
        self.form_frame_layout.addWidget(self.muestra_widget, 6, 1)
        # self.form_frame_layout.addWidget(equipo_label, 6, 2)
        # self.form_frame_layout.addWidget(self.equipo_widget, 6, 3)

        self.form_frame_layout.addWidget(end_label, 6, 2)
        self.form_frame_layout.addWidget(self.pago_widget, 6, 3)

        self.table = Table(self)

        self.elaborado_frame = QtWidgets.QFrame()
        self.elaborado_layout = QtWidgets.QHBoxLayout(self.elaborado_frame)

        self.elaborado_label = QtWidgets.QLabel("Elaborado por:")
        self.elaborado_widget = QtWidgets.QComboBox()
        self.elaborado_widget.addItems(config.ADMINS)
        self.elaborado_layout.addWidget(self.elaborado_label)
        self.elaborado_layout.addWidget(self.elaborado_widget)

        self.button_frame_layout = QtWidgets.QHBoxLayout(self.button_frame)
        self.notificar_widget = QtWidgets.QCheckBox("Notificar")
        self.notificar_widget.setCheckState(2)
        self.view_button = QtWidgets.QPushButton("Ver códigos")
        self.guardar_button = QtWidgets.QPushButton("Guardar")
        self.limpiar_button = QtWidgets.QPushButton("Limpiar")

        self.button_frame_layout.addWidget(self.notificar_widget)
        self.button_frame_layout.addWidget(self.guardar_button)
        self.button_frame_layout.addWidget(self.limpiar_button)
        self.button_frame_layout.addWidget(self.view_button)

        self.total_frame_layout = QtWidgets.QFormLayout(self.total_frame)
        subtotal_label = QtWidgets.QLabel("Subtotal:")
        self.subtotal_widget = QtWidgets.QLabel()
        descuento_label = QtWidgets.QLabel("Descuento:")
        self.descuento_widget = QtWidgets.QLabel()
        total_label = QtWidgets.QLabel("Total:")
        self.total_widget = QtWidgets.QLabel()

        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.cotizacion_frame.setSizePolicy(sizePolicy)
        self.total_frame.setSizePolicy(sizePolicy)
        self.button_frame.setSizePolicy(sizePolicy)
        self.elaborado_frame.setSizePolicy(sizePolicy)

        self.total_frame_layout.setAlignment(QtCore.Qt.AlignRight)
        self.button_frame_layout.setAlignment(QtCore.Qt.AlignRight)
        self.verticalLayout.setAlignment(QtCore.Qt.AlignRight)
        self.elaborado_layout.setAlignment(QtCore.Qt.AlignRight)

        self.observaciones_correo_widget = QtWidgets.QTextEdit()
        self.observaciones_correo_widget.setMaximumHeight(30)
        self.observaciones_pdf_widget = QtWidgets.QTextEdit()
        self.observaciones_pdf_widget.setMaximumHeight(30)
        self.descuento_text_widget = QtWidgets.QLineEdit()
        # self.descuento_text_widget.setMaximumHeight(20)

        self.observaciones_layout = QtWidgets.QFormLayout(self.observaciones_frame)
        self.observaciones_layout.addRow(QtWidgets.QLabel("Correo:"), self.observaciones_correo_widget)
        self.observaciones_layout.addRow(QtWidgets.QLabel("PDF:"), self.observaciones_pdf_widget)
        self.observaciones_layout.addRow(QtWidgets.QLabel("Descuento por:"), self.descuento_text_widget)

        self.total_frame_layout.addRow(subtotal_label, self.subtotal_widget)
        self.total_frame_layout.addRow(descuento_label, self.descuento_widget)
        self.total_frame_layout.addRow(total_label, self.total_widget)

        self.verticalLayout.addWidget(self.cotizacion_frame)
        self.verticalLayout.addWidget(self.autocompletar_widget)
        self.verticalLayout.addWidget(self.form_frame)
        self.verticalLayout.addWidget(self.table)
        self.verticalLayout.addWidget(self.total_frame)
        self.verticalLayout.addWidget(self.observaciones_frame)
        self.verticalLayout.addWidget(self.elaborado_frame)
        self.verticalLayout.addWidget(self.button_frame)

        self.setAutoCompletar()

        self.cotizacion = objects.Cotizacion()
        self.setLastCotizacion()

        self.resize(600, 620)

        self.interno_widget.currentIndexChanged.connect(self.changeInterno)
        self.limpiar_button.clicked.connect(self.limpiar)
        self.guardar_button.clicked.connect(self.guardar)
        self.view_button.clicked.connect(self.verCodigos)

    def agregarDesdeCodigo(self, codigo):
        self.table.agregarServicio(codigo)

    def setAutoCompletar(self):
        for item in self.AUTOCOMPLETE_WIDGETS:
            widget = eval("self.%s_widget"%item)
            widget.textChanged.connect(self.autoCompletar)
            widget.returnPressed.connect(self.changeAutocompletar)

    def changeAutocompletar(self):
        self.autocompletar_widget.setChecked(False)

    def updateAutoCompletar(self):
        for item in self.AUTOCOMPLETE_WIDGETS:
            exec("self.%s_widget.update()"%item)

    def setInternoWidget(self, value):
        index = self.interno_widget.findText(value)
        self.interno_widget.setCurrentIndex(index)

    def autoCompletar(self, text):
        if text != "":
            df = objects.CLIENTES_DATAFRAME[self.AUTOCOMPLETE_FIELDS]
            booleans = df.isin([text]).values.sum(axis = 1)
            pos = np.where(booleans)[0]
            cliente = objects.CLIENTES_DATAFRAME.iloc[pos]

            if len(pos):
                # if len(pos) > 1:
                #     print(len(pos))
                if self.autocompletar_widget.isChecked():
                    for (field, widgetT) in zip(self.FIELDS, self.WIDGETS):
                        if field in objects.CLIENTES_DATAFRAME.keys():
                            val = str(cliente[field].values[0])
                            if val == "nan": val = ""
                            widget = eval("self.%s_widget"%widgetT)

                            if field == "Interno":
                                self.setInternoWidget(val)
                                # if val == "Interno": widget.setCheckState(2)
                                # else: widget.setCheckState(0)
                            else:
                                widget.blockSignals(True)
                                widget.setText(val)
                                widget.blockSignals(False)

    def changeInterno(self, index):
        state = False
        division = self.getInterno()
        try:
            if division:
                if division == 'Interno': state = True
                self.responsable_widget.setEnabled(state)
                self.proyecto_widget.setEnabled(state)
                self.codigo_widget.setEnabled(state)

                self.cotizacion.setInterno(division)
                self.table.updateInterno()

                self.setTotal()

                if division == 'Industria':
                    self.descuento_text_widget.setText("")
                    self.descuento_text_widget.setEnabled(False)
                else: self.descuento_text_widget.setEnabled(True)
        except Exception as e:
            self.errorWindow(e)

    def changeEquipo(self, i):
        pass

    def limpiar(self):
        self.table.clean()
        for field in self.WIDGETS:
            widget = eval("self.%s_widget"%field)
            widget.blockSignals(True)
            if field != "interno":
                widget.setText("")
            widget.blockSignals(False)
        self.setLastCotizacion()
        self.setInternoWidget('Interno')
        # self.interno_widget.setCheckState(2)
        self.pago_widget.setCurrentIndex(0)
        self.elaborado_widget.setCurrentIndex(0)
        self.notificar_widget.setChecked(True)
        self.cotizacion.setServicios([])
        self.subtotal_widget.setText("")
        self.descuento_widget.setText("")
        self.total_widget.setText("")
        self.elaborado_label.setText("Elaborado por:")
        self.autocompletar_widget.setChecked(True)
        self.observaciones_pdf_widget.setText("")
        self.observaciones_correo_widget.setText("")
        self.descuento_text_widget.setText("")

    def verCodigos(self):
        self.ver_dialog.show()

    def sendCorreo(self, names):
        if self.notificar_widget.isChecked():
            to = self.cotizacion.getUsuario().getCorreo()
            pago = self.pago_widget.currentText()
            observaciones = self.observaciones_correo_widget.toPlainText()
            if pago == "Transferencia interna":
                self.dialog = CorreoDialog((to, names, observaciones), target = correo.sendCotizacionTransferencia)
            elif pago == "Factura":
                self.dialog = CorreoDialog((to, names, observaciones), target = correo.sendCotizacionFactura)
            elif pago == "Recibo":
                self.dialog = CorreoDialog((to, names, observaciones), target = correo.sendCotizacionRecibo)
            else: print("ERROR, not implemented")
            self.dialog.start()
            self.dialog.exec_()
            if self.dialog.exception != None:
                raise(self.dialog.exception)
        else:
            NoNotificacion().exec_()

    def closePDF(self, p1, old):
        new = [proc.pid for proc in psutil.process_iter()]
        try:
            new.remove(p1.pid)
        except:
            return None

        new = [proc for proc in new if proc not in old]

        current = psutil.Process(os.getpid()).name()
        caller = p1.pid
        caller = psutil.Process(caller).name()
        try:
            for proc in new:
                p = psutil.Process(proc)
                parent = p.parent()
                if (parent != None):
                    sparent = parent.parent()
                    if (sparent != None):
                        if (sparent.name() == current) or (sparent.name() == caller):
                            p.terminate()
                    if (parent.name() == current) or (parent.name() == caller):
                        p.terminate()
            p1.kill()
        except Exception as e:
            print("On closePDF:", e)

    def confirmGuardar(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText("¿Está seguro que desea guardar esta cotización?.\nVerifique los datos.")
        msg.setWindowTitle("Confirmar")
        msg.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.Cancel)
        ans = msg.exec_()
        if ans == QtWidgets.QMessageBox.Yes:
            return True
        return False

    def openPDF(self, file):
        path = os.path.dirname(sys.executable)
        path = os.path.join(path, file)
        if not os.path.exists(path):
            path = os.path.join(os.getcwd(), file)
        old = [proc.pid for proc in psutil.process_iter()]

        return path, old

    def guardar(self):
        try:
            dic = {}
            for key in self.WIDGETS:
                if key == "interno":
                    value = self.getInterno()
                    if not value:
                        raise(Exception('El tipo de usuario no es valido.'))
                else: value = eval("self.%s_widget.text()"%key)
                if ((value == "") and not (key in self.IGNORE)):
                    if (key == "responsable") and not (self.getInterno() in ['Interno', 'Campus']):
                        pass
                    else: raise(Exception("Existen campos sin llenar en la información del usuario."))
                dic[key] = value
            del dic["muestra"]
            dic["pago"] = self.pago_widget.currentText()

            if len(self.getServicios()) == 0:
                raise(Exception("No existen servicios cotizados."))

            if self.cotizacion.getElaborado() != "":
                self.cotizacion.setModificado(self.elaborado_widget.currentText())
            else:
                self.cotizacion.setElaborado(self.elaborado_widget.currentText())

            usuario = objects.Usuario(**dic)
            self.cotizacion.setUsuario(usuario)
            servicios = objects.sortServicios(self.getServicios())
            self.cotizacion.setMuestra(self.muestra_widget.text())
            self.cotizacion.setObservacionPDF(self.observaciones_pdf_widget.toPlainText())
            self.cotizacion.setObservacionCorreo(self.observaciones_correo_widget.toPlainText())

            numeros = []
            cotizaciones = []
            processes = []
            paths = []
            olds = []
            for key in servicios:
                cotizacion = copy(self.cotizacion)
                cotizacion.setDescuentoText(self.descuento_text_widget.text())
                n = objects.getNumeroCotizacion(key)
                numeros.append(n)
                cotizacion.setServicios(servicios[key])
                cotizacion.setNumero(n)
                cotizacion.makePDFCotizacion()
                cotizaciones.append(cotizacion)

                path, old = self.openPDF(cotizacion.getPath())
                paths.append(path)
                olds.append(old)
                p1 = Popen(path, shell = True)
                processes.append(p1)
            if self.confirmGuardar():
                for (i, cotizacion) in enumerate(cotizaciones):
                    self.closePDF(processes[i], olds[i])
                    cotizacion.save(to_pdf = False)
                self.updateAutoCompletar()
                self.sendCorreo(numeros)
                self.limpiar()
                self.setLastCotizacion()
            else:
                self.cotizacion.setUsuario(None)
                self.cotizacion.setMuestra(None)
                for (i, cotizacion) in enumerate(cotizaciones):
                    self.closePDF(processes[i], olds[i])
                    try: os.remove(paths[i])
                    except PermissionError: pass

            old_servicios = self.getServicios()
            self.cotizacion = objects.Cotizacion()
            self.cotizacion.setServicios(old_servicios)

        except Exception as e:
            self.errorWindow(e)

    def setLastCotizacion(self):
        pass

    def loadCotizacion(self, number):
        # try:
        try:
            cotizacion = self.cotizacion.load(number)
            self.limpiar()

            self.cotizacion = cotizacion
            user = self.cotizacion.getUsuario()

            for widgetT in self.WIDGETS:
                if widgetT != "interno":
                    text = widgetT.title()
                    widget = eval("self.%s_widget"%widgetT)
                    try:
                        val = str(eval("user.get%s()"%text))
                        widget.setText(val)
                    except: pass

            self.setInternoWidget(user.getInterno())
            self.pago_widget.setCurrentText(user.getPago())
            self.muestra_widget.setText(self.cotizacion.getMuestra())
            self.elaborado_label.setText("Modificado por:")
            self.elaborado_widget.setCurrentIndex(0)

            self.observaciones_pdf_widget.setText(self.cotizacion.getObservacionPDF())
            self.observaciones_correo_widget.setText(self.cotizacion.getObservacionCorreo())
            try:
                self.setTotal()
                self.table.setFromCotizacion()
            except Exception:
                e = Exception("Cotización incompatible")
                self.errorWindow(e)

        except FileNotFoundError:
            e = Exception("La cotización no se encuentra grabada.")
            self.errorWindow(e)

    def centerOnScreen(self):
        resolution = QtWidgets.QDesktopWidget().screenGeometry()
        self.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))

    def addServicio(self, servicio):
        self.cotizacion.addServicio(servicio)

    def getCodigos(self):
        return self.cotizacion.getCodigos()

    def getCodigosPrefix(self):
        return self.cotizacion.getCodigosPrefix()

    def getServicio(self, cod):
        return self.cotizacion.getServicio(cod)

    def getServicios(self):
        return self.cotizacion.getServicios()

    # def getEquipo(self):
    #     return self.equipo_widget.currentText()

    def getInterno(self):
        interno = self.interno_widget.currentText()
        return interno

    def removeServicio(self, index):
        self.cotizacion.removeServicio(index)

    def setTotal(self, total = None):
        if total == None:
            total = self.cotizacion.getTotal()
            subtotal = self.cotizacion.getSubtotal()
            descuento = self.cotizacion.getDescuentos()
        self.subtotal_widget.setText("{:,}".format(subtotal))
        self.descuento_widget.setText("{:,}".format(-descuento))
        self.total_widget.setText("{:,}".format(total))

    def errorWindow(self, exception):
        traceback.print_tb(exception.__traceback__)
        error_text = str(exception)
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setText(error_text)
        msg.setWindowTitle("Error")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

class NoNotificacion(QtWidgets.QMessageBox):
    def __init__(self):
        super(NoNotificacion, self).__init__()
        self.setIcon(QtWidgets.QMessageBox.Information)
        self.setText("El usuario no será notificado.")
        self.setWindowTitle("Warning")
        self.setStandardButtons(QtWidgets.QMessageBox.Ok)

class DescontarWindow(SubWindow):
    FIELDS = ["Cotización", "Fecha", "Nombre", "Correo", "Equipo", "Valor"]
    WIDGETS = ["cotizacion", "fecha", "nombre", "correo", "equipo", "valor"]

    def __init__(self, parent = None):
        super(DescontarWindow, self).__init__(parent)
        self.setWindowTitle("Descontar servicios")

        wid = QtWidgets.QWidget(self)
        self.setWidget(wid)

        self.layout = QtWidgets.QVBoxLayout(wid)
        self.form = QtWidgets.QFrame()
        self.form_layout = QtWidgets.QFormLayout(self.form)
        cotizacion_label = QtWidgets.QLabel("Cotización:")
        self.cotizacion_widget = AutoLineEdit("Cotización", self)
        fecha_label = QtWidgets.QLabel("Fecha:")
        self.fecha_widget = QtWidgets.QLabel("")
        nombre_label = QtWidgets.QLabel("Nombre:")
        self.nombre_widget = QtWidgets.QLabel("")
        correo_label = QtWidgets.QLabel("Correo:")
        self.correo_widget = QtWidgets.QLabel("")
        equipo_label = QtWidgets.QLabel("Equipo:")
        self.equipo_widget = QtWidgets.QLabel("")
        valor_label = QtWidgets.QLabel("Valor:")
        self.valor_widget = QtWidgets.QLabel("")

        self.form_layout.addRow(fecha_label, self.fecha_widget)
        self.form_layout.addRow(nombre_label, self.nombre_widget)
        self.form_layout.addRow(correo_label, self.correo_widget)
        self.form_layout.addRow(equipo_label, self.equipo_widget)
        self.form_layout.addRow(valor_label, self.valor_widget)
        self.form_layout.addRow(cotizacion_label, self.cotizacion_widget)

        self.layout.addWidget(self.form)

        self.item_frame = QtWidgets.QFrame()
        self.item_frame.setFrameStyle(1)
        self.item_layout = QtWidgets.QGridLayout(self.item_frame)
        cod = QtWidgets.QLabel("Código")
        des = QtWidgets.QLabel("Descripción")
        n = QtWidgets.QLabel("Cantidad")
        self.item_layout.addWidget(cod, 0, 0)
        self.item_layout.addWidget(des, 0, 1)
        self.item_layout.addWidget(n, 0, 2)

        self.layout.addWidget(self.item_frame)

        self.pago_frame = QtWidgets.QFrame()
        self.pago_layout = QtWidgets.QGridLayout(self.pago_frame)
        self.check_widget = None
        self.check_label = None
        self.referencia_widget = None

        self.layout.addWidget(self.pago_frame)

        self.buttons_frame = QtWidgets.QFrame()
        self.buttons_layout = QtWidgets.QHBoxLayout(self.buttons_frame)
        self.guardar_button = QtWidgets.QPushButton("Guardar")
        self.notificar_widget = QtWidgets.QCheckBox("Notificar")
        self.notificar_widget.setCheckState(2)

        self.otros_widget = QtWidgets.QPushButton("Otros")
        self.otros_widget.setEnabled(False)

        self.otros_frame = QtWidgets.QFrame()
        self.otros_layout = QtWidgets.QGridLayout(self.otros_frame)
        self.layout.addWidget(self.otros_frame)

        self.buttons_layout.addWidget(self.notificar_widget)
        self.buttons_layout.addWidget(self.guardar_button)
        self.buttons_layout.addWidget(self.otros_widget)

        self.layout.addWidget(self.buttons_frame)

        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.buttons_frame.setSizePolicy(sizePolicy)
        self.buttons_layout.setAlignment(QtCore.Qt.AlignRight)

        self.cotizacion_widget.textChanged.connect(self.cotizacionChanged)
        self.guardar_button.clicked.connect(self.guardarHandler)
        self.otros_widget.clicked.connect(self.otrosHandler)

        self.floats_labels = []
        self.floats_spins = []
        self.cotizacion = objects.Cotizacion()
        self.init_size = (600, 450)
        self.setFixedSize(*self.init_size)

    def removeOtros(self):
        if self.otros_widget.text() != "Otros":
            self.otros_widget.setText("Otros")
            self.otros_layout.removeWidget(self.otros_equipo_label)
            self.otros_layout.removeWidget(self.otros_codigo_label)
            self.otros_layout.removeWidget(self.otros_cantidad_label)
            self.otros_layout.removeWidget(self.otros_equipo_widget)
            self.otros_layout.removeWidget(self.otros_codigo_widget)
            self.otros_layout.removeWidget(self.otros_cantidad_widget)

            self.otros_equipo_label.deleteLater()
            self.otros_codigo_label.deleteLater()
            self.otros_cantidad_label.deleteLater()
            self.otros_equipo_widget.deleteLater()
            self.otros_codigo_widget.deleteLater()
            self.otros_cantidad_widget.deleteLater()

    def otrosHandler(self):
        if self.otros_widget.text() == "Otros":
            self.otros_widget.setText("Agregar")
            self.otros_equipo_label = QtWidgets.QLabel("Equipo:")
            self.otros_codigo_label = QtWidgets.QLabel("Código:")
            self.otros_cantidad_label = QtWidgets.QLabel("Cantidad:")

            self.otros_equipo_widget = QtWidgets.QComboBox()
            self.otros_codigo_widget = QtWidgets.QComboBox()
            self.otros_cantidad_widget = QtWidgets.QDoubleSpinBox()

            self.otros_equipo_widget.addItems(constants.EQUIPOS)

            self.otros_cantidad_widget.setMinimum(0)
            self.otros_cantidad_widget.setDecimals(1)
            # self.otros_cantidad_widget.setMaximum(rest)
            self.otros_cantidad_widget.setSingleStep(0.1)

            self.otros_equipo_widget.currentIndexChanged.connect(self.equipoChanged)

            self.otros_layout.addWidget(self.otros_equipo_label, 0, 0)
            self.otros_layout.addWidget(self.otros_codigo_label, 0, 1)
            self.otros_layout.addWidget(self.otros_cantidad_label, 0, 2)

            self.otros_layout.addWidget(self.otros_equipo_widget, 1, 0)
            self.otros_layout.addWidget(self.otros_codigo_widget, 1, 1)
            self.otros_layout.addWidget(self.otros_cantidad_widget, 1, 2)

            self.equipoChanged(0)
        else:
            equipo = self.otros_equipo_widget.currentText()
            codigo = self.otros_codigo_widget.currentText().split("-")[0]
            value = self.otros_cantidad_widget.value()

            if value:
                servicio = objects.Servicio(equipo, codigo, self.cotizacion.getInterno(),
                            cantidad = 0, agregado_posteriormente = True)
                try:
                    self.cotizacion.addServicio(servicio)
                except Exception as e:
                    for cod in self.cotizacion.getCodigos():
                        if cod == codigo:
                            servicio = self.cotizacion.getServicio(cod)
                servicio.descontar(value)
            self.removeOtros()
            self.guardarHandler()

    def equipoChanged(self, equipo):
        equipo = self.otros_equipo_widget.currentText()
        self.otros_codigo_widget.clear()
        if self.cotizacion != None:
            interno = "Externo"
            if self.cotizacion.getInterno():
                interno = "Interno"

            df = eval("constants.%s"%equipo)[["Código", "Descripción", interno]].values
            values = ["%s-%s (%s)"%(c, d, p) for (c, d, p) in df]
            self.otros_codigo_widget.addItems(values)

    def updateDataFrames(self):
        self.cotizacion_widget.update()

    def sendCorreo(self):
        if self.notificar_widget.isChecked():
            to = self.cotizacion.getUsuario().getCorreo()
            file_name = self.cotizacion.getNumero()
            self.dialog = CorreoDialog((to, file_name), target = correo.sendRegistro)
            self.dialog.start()
            self.dialog.exec_()
            if self.dialog.exception != None:
                raise(self.dialog.exception)
        else:
            NoNotificacion().exec_()

    def guardarHandler(self):
        servicios = self.cotizacion.getServicios()
        try:
            if len(servicios):
                if self.check_widget != None:
                    if self.check_widget.isChecked():
                        self.cotizacion.setPago(self.referencia_widget.text())
                        self.cotizacion.setAplicado(self.elaborado_widget.currentText())

                for i in range(len(self.floats_spins)):
                    val = self.floats_spins[i].value()
                    servicio = servicios[i]
                    servicio.descontar(val)
                self.cotizacion.save(to_cotizacion = False, to_reporte = True)
                self.cotizacion.toRegistro()
                self.sendCorreo()
            self.clean()
        except Exception as e:
            self.errorWindow(e)

    def clean(self):
        for widget in self.WIDGETS: exec("self.%s_widget.setText('')"%widget)
        self.cleanWidgets()

    def cotizacionChanged(self, text):
        if text != "":
            self.otros_widget.setEnabled(False)
            self.removeOtros()
            registro = objects.REGISTRO_DATAFRAME[objects.REGISTRO_DATAFRAME["Cotización"] == text]
            if len(registro):
                for (field, widgetT) in zip(self.FIELDS, self.WIDGETS):
                        val = str(registro[field].values[0])
                        widget = eval("self.%s_widget"%widgetT)
                        if val == "nan": val = ""
                        widget.setText(val)
            self.cleanWidgets()
            try:
                self.cotizacion = self.cotizacion.load(text)
                n = len(self.cotizacion.getServicios())
            except: n = 0

            i = -1
            for (i, servicio) in enumerate(self.cotizacion.getServicios()):
                self.otros_widget.setEnabled(True)
                fmt = "%s"
                if servicio.isAgregado():
                    fmt = "<font color='red'>%s</font>"

                cod = QtWidgets.QLabel(fmt % servicio.getCodigo())
                dec = QtWidgets.QLabel(fmt % servicio.getDescripcion())
                paid = servicio.getCantidad()
                rest = servicio.getRestantes()
                used = paid - rest
                spin = QtWidgets.QDoubleSpinBox()
                total = QtWidgets.QLabel(fmt % ("%.1f/%.1f"%(used, paid)))
                spin.setMinimum(0)
                spin.setDecimals(1)
                spin.setMaximum(rest)
                spin.setSingleStep(0.1)
                self.item_layout.addWidget(cod, i + 1, 0)
                self.item_layout.addWidget(dec, i + 1, 1)
                self.item_layout.addWidget(spin, i + 1, 2)
                self.item_layout.addWidget(total, i + 1, 3)
                self.floats_labels.append(cod)
                self.floats_labels.append(dec)
                self.floats_spins.append(spin)
                self.floats_labels.append(total)

            if i >= 0:
                dinero = "{:,}".format(self.cotizacion.getDineroUsado())
                total_d = "{:,}".format(self.cotizacion.getTotal())
                self.valor_widget.setText("%s/%s"%(dinero, total_d))

                self.check_widget = QtWidgets.QCheckBox("Aplicar pago")
                self.check_label = QtWidgets.QLabel("Referencia:")
                self.referencia_widget = QtWidgets.QLineEdit()
                self.elaborado_label = QtWidgets.QLabel("Aplicado por:")
                self.elaborado_widget = QtWidgets.QComboBox()

                self.elaborado_widget.addItems(config.ADMINS)

                self.pago_layout.addWidget(self.check_widget, 0, 0)
                self.pago_layout.addWidget(self.check_label, 1, 0)
                self.pago_layout.addWidget(self.referencia_widget, 1, 1)
                self.pago_layout.addWidget(self.elaborado_label, 2, 0)
                self.pago_layout.addWidget(self.elaborado_widget, 2, 1)

                self.check_widget.setTristate(False)
                self.referencia_widget.setText(self.cotizacion.getReferenciaPago())
                self.check_widget.setChecked(self.cotizacion.isPago())

                if not self.cotizacion.isPago():
                    self.check_widget.stateChanged.connect(self.checkHandler)
                    self.checkHandler(self.cotizacion.isPago())
                else:
                    self.check_widget.setEnabled(False)
                    self.referencia_widget.setEnabled(False)
                    self.elaborado_widget.setEnabled(False)

                h = self.init_size[1]
                h += 18*(len(self.cotizacion.getServicios()) + 4)
                self.setFixedHeight(h)

    def cleanWidgets(self):
        for item in self.floats_labels:
            self.item_layout.removeWidget(item)
            item.deleteLater()
        for item in self.floats_spins:
            self.item_layout.removeWidget(item)
            item.deleteLater()
        if self.check_widget != None:
            self.pago_layout.removeWidget(self.check_widget)
            self.pago_layout.removeWidget(self.check_label)
            self.pago_layout.removeWidget(self.referencia_widget)
            self.pago_layout.removeWidget(self.elaborado_label)
            self.pago_layout.removeWidget(self.elaborado_widget)
            self.check_widget.deleteLater()
            self.check_label.deleteLater()
            self.referencia_widget.deleteLater()
            self.elaborado_label.deleteLater()
            self.elaborado_widget.deleteLater()

            self.check_widget = None
            self.check_label = None
            self.referencia_widget = None
            self.elaborado_label = None
            self.elaborado_widget = None

        self.floats_labels = []
        self.floats_spins = []
        self.cotizacion = objects.Cotizacion()
        self.setFixedSize(*self.init_size)

    def checkHandler(self, state):
        if state:
            self.referencia_widget.setEnabled(True)
            self.elaborado_widget.setEnabled(True)
        else:
            self.referencia_widget.setText("")
            self.elaborado_widget.setCurrentIndex(0)
            self.referencia_widget.setEnabled(False)
            self.elaborado_widget.setEnabled(False)

    def errorWindow(self, exception):
        error_text = str(exception)
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setText(error_text)
        msg.setWindowTitle("Error")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

class PandasModel(QtCore.QAbstractTableModel):
    def __init__(self, data, parent = None, checkbox = True):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self.dataframe = data
        self._data = data.values
        self.checkbox = checkbox

        if self.checkbox:
            temp = []
            for i in range(self._data.shape[0]):
                c = QtWidgets.QCheckBox()
                c.setChecked(True)
                temp.append(c)

            new = np.zeros((self._data.shape[0], self._data.shape[1] + 1), dtype = object)
            new[:, 0] = temp
            new[:, 1:] = self._data

            self._data = new

            self.headerdata = ["Guardar"] + list(data.keys())
        else:
            self.headerdata = list(data.keys())

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role = QtCore.Qt.DisplayRole):
        if not index.isValid():
            return None
        if (index.column() == 0 and self.checkbox):
            value = self._data[index.row(), index.column()].text()
        else:
            value = self._data[index.row(), index.column()]
        if role == QtCore.Qt.EditRole:
            return value
        elif role == QtCore.Qt.DisplayRole:
            return value
        elif role == QtCore.Qt.CheckStateRole:
            if index.column() == 0 and self.checkbox:
                if self._data[index.row(), index.column()].isChecked():
                    return QtCore.Qt.Checked
                else:
                    return QtCore.Qt.Unchecked

    def flags(self, index):
        if not index.isValid():
            return None
        if index.column() == 0 and self.checkbox:
            return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable
        else:
            return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable

    def headerData(self, col, orientation, role):
        if orientation == QtCore.Qt.Horizontal and role == QtCore.Qt.DisplayRole:
            return QtCore.QVariant(self.headerdata[col])

        return QtCore.QVariant()

    def setData(self, index, value, role):
        if not index.isValid():
            return False
        if role == QtCore.Qt.CheckStateRole and index.column() == 0:
            if value == QtCore.Qt.Checked:
                self._data[index.row(), index.column()].setChecked(True)
            else:
                self._data[index.row(), index.column()].setChecked(False)

        self.dataChanged.emit(index, index)
        return True

    def whereIsChecked(self):
        return np.array([self._data[i, 0].isChecked() for i in range(self.rowCount())], dtype = bool)

class BuscarWindow(SubWindow):
    WIDGETS = ["equipo", "nombre", "correo", "institucion", "responsable", "cotizacion"]
    FIELDS = ["Equipo", "Nombre", "Correo", "Institución", "Responsable", "Cotización"]
    def __init__(self, parent = None):
        super(BuscarWindow, self).__init__(parent)
        self.setWindowTitle("Buscar")
        self.parent = parent

        wid = QtWidgets.QWidget(self)
        self.setWidget(wid)
        self.layout = QtWidgets.QVBoxLayout(wid)

        form = QtWidgets.QFrame()
        layout = QtWidgets.QHBoxLayout(form)

        self.form1 = QtWidgets.QFrame()
        self.form1_layout = QtWidgets.QFormLayout(self.form1)
        self.form2 = QtWidgets.QFrame()
        self.form2_layout = QtWidgets.QFormLayout(self.form2)
        self.buttons_frame = QtWidgets.QFrame()
        self.buttons_layout = QtWidgets.QHBoxLayout(self.buttons_frame)

        layout.addWidget(self.form1)
        layout.addWidget(self.form2)

        self.equipo_widget = AutoLineEdit('Equipo', self, False)
        self.nombre_widget = AutoLineEdit('Nombre', self, False)
        self.institucion_widget = AutoLineEdit("Institución", self, False)
        self.correo_widget = AutoLineEdit("Correo", self, False)
        self.responsable_widget = AutoLineEdit("Responsable", self, False)
        self.cotizacion_widget = AutoLineEdit('Cotización', self, False)

        self.form1_layout.addRow(QtWidgets.QLabel('Equipo'), self.equipo_widget)
        self.form1_layout.addRow(QtWidgets.QLabel('Nombre'), self.nombre_widget)
        self.form1_layout.addRow(QtWidgets.QLabel('Número Cotización'), self.cotizacion_widget)
        self.form2_layout.addRow(QtWidgets.QLabel('Institución'), self.institucion_widget)
        self.form2_layout.addRow(QtWidgets.QLabel('Correo'), self.correo_widget)
        self.form2_layout.addRow(QtWidgets.QLabel('Responsable'), self.responsable_widget)

        self.guardar_button = QtWidgets.QPushButton("Generar reportes")
        self.limpiar_button = QtWidgets.QPushButton("Limpiar")

        self.buttons_layout.addWidget(self.guardar_button)
        self.buttons_layout.addWidget(self.limpiar_button)

        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.buttons_frame.setSizePolicy(sizePolicy)
        self.buttons_layout.setAlignment(QtCore.Qt.AlignRight)

        self.equipo_widget.textChanged.connect(lambda: self.getChanges('Equipo'))
        self.nombre_widget.textChanged.connect(lambda: self.getChanges('Nombre'))
        self.institucion_widget.textChanged.connect(lambda: self.getChanges('Institución'))
        self.cotizacion_widget.textChanged.connect(lambda: self.getChanges('Cotización'))
        self.correo_widget.textChanged.connect(lambda: self.getChanges('Correo'))
        self.responsable_widget.textChanged.connect(lambda: self.getChanges('Responsable'))

        self.guardar_button.clicked.connect(self.guardar)
        self.limpiar_button.clicked.connect(self.limpiar)

        self.table = QtWidgets.QTableView()
        # self.table.setSortingEnabled(True)
        # self.table.doubleClicked.connect(self.doubleClick)
        self.table.viewport().installEventFilter(self)
        self.layout.addWidget(form)
        self.layout.addWidget(self.table)

        self.layout.addWidget(self.buttons_frame)

        model = PandasModel(objects.REGISTRO_DATAFRAME)

        self.table.setModel(model)

        self.bools = np.ones(objects.REGISTRO_DATAFRAME.shape[0], dtype = bool)

        self.resize(800, 600)

        self.table.resizeRowsToContents()
        self.table.resizeColumnsToContents()

    def eventFilter(self, source, event):
        if (event.type() == QtCore.QEvent.MouseButtonDblClick and
            source is self.table.viewport()):
            item = self.table.indexAt(event.pos())
            row = item.row()
            if(event.buttons() == QtCore.Qt.RightButton):
                self.doubleClick(row, False)
            else:
                self.doubleClick(row, True)
            return True
        return False

    def doubleClick(self, row, left):
        df = self.table.model().dataframe
        cotizacion = df.iloc[row]['Cotización']
        if left:
            self.parent.modificarCotizacion(cotizacion)
        else:
            self.parent.abrirPDFCotizacion(cotizacion)

    def getChanges(self, source):
        self.bools = np.ones(objects.REGISTRO_DATAFRAME.shape[0], dtype = bool)
        for i in range(len(self.WIDGETS)):
            source = self.FIELDS[i]
            widget = self.WIDGETS[i]
            value = eval("self.%s_widget"%widget).text()
            if value != "":
                pos = objects.REGISTRO_DATAFRAME[source].str.contains(value, case = False, na = False)
                self.bools *= pos
        self.update()

    def updateAutoCompletar(self):
        self.equipo_widget.update()
        self.nombre_widget.update()
        self.institucion_widget.update()
        self.responsable_widget.update()

    def update(self):
        if self.bools.shape[0] != objects.REGISTRO_DATAFRAME.shape[0]:
            self.bools = np.ones(objects.REGISTRO_DATAFRAME.shape[0], dtype = bool)
        old = self.table.model().dataframe
        df = objects.REGISTRO_DATAFRAME[self.bools]
        if not old.equals(df):
            model = PandasModel(df)
            self.table.setModel(model)
            #self.table.resizeRowsToContents()

    def limpiar(self):
        for widget in self.WIDGETS:
            widget = eval("self.%s_widget"%widget)
            widget.blockSignals(True)
            widget.setText("")
            widget.blockSignals(False)

        self.table.setModel(PandasModel(objects.REGISTRO_DATAFRAME))

    def guardar(self):
        if os.getcwd()[0] == "\\":
            quit_msg = "No es posible grabar desde un computador en red."
            QtWidgets.QMessageBox.warning(self, 'Error',
                             quit_msg, QtWidgets.QMessageBox.Ok)

        else:
            folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Directory")
            model = self.table.model()
            pos = np.where(model.whereIsChecked())[0]
            cotizacion = objects.Cotizacion()
            for i in pos:
                cot = model._data[i, 1]
                try:
                    cotizacion.load(cot).makePDFReporte()
                    old = os.path.join(constants.PDF_DIR, cot + "_Reporte.pdf")
                    new = os.path.join(folder, cot + "_Reporte.pdf")
                    os.rename(old, new)
                except Exception as e: pass

class GestorWindow(SubWindow):
    FIELDS = ["Cotización", "Fecha", "Nombre", "Correo", "Equipo", "Valor", "Tipo de Pago"]
    WIDGETS = ["cotizacion", "fecha", "nombre", "correo", "equipo", "valor", "tipo"]
    def __init__(self, parent = None):
        super(GestorWindow, self).__init__(parent)
        self.setWindowTitle("Enviar correo a Gestor")

        wid = QtWidgets.QWidget(self)
        self.setWidget(wid)
        self.layout = QtWidgets.QVBoxLayout(wid)

        self.form = QtWidgets.QFrame()
        self.form_layout = QtWidgets.QFormLayout(self.form)

        cotizacion_label = QtWidgets.QLabel("Cotización:")
        self.cotizacion_widget = AutoLineEdit("Cotización", self)

        fecha_label = QtWidgets.QLabel("Fecha:")
        self.fecha_widget = QtWidgets.QLabel("")

        nombre_label = QtWidgets.QLabel("Nombre:")
        self.nombre_widget = QtWidgets.QLabel("")

        correo_label = QtWidgets.QLabel("Correo:")
        self.correo_widget = QtWidgets.QLabel("")

        equipo_label = QtWidgets.QLabel("Equipo:")
        self.equipo_widget = QtWidgets.QLabel("")

        valor_label = QtWidgets.QLabel("Valor:")
        self.valor_widget = QtWidgets.QLabel("")

        tipo_label = QtWidgets.QLabel("Tipo de Pago:")
        self.tipo_widget = QtWidgets.QLabel("")

        self.form_layout.addRow(fecha_label, self.fecha_widget)
        self.form_layout.addRow(nombre_label, self.nombre_widget)
        self.form_layout.addRow(correo_label, self.correo_widget)
        self.form_layout.addRow(equipo_label, self.equipo_widget)
        self.form_layout.addRow(valor_label, self.valor_widget)
        self.form_layout.addRow(tipo_label, self.tipo_widget)
        self.form_layout.addRow(cotizacion_label, self.cotizacion_widget)

        self.buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok,
            QtCore.Qt.Horizontal, self)

        self.layout.addWidget(self.form)
        self.layout.addWidget(self.buttons)
        self.buttons.accepted.connect(self.accept)
        self.layout.addWidget(self.buttons)

        self.cotizacion_widget.textChanged.connect(self.autoCompletar)

        self.setFixedSize(self.layout.sizeHint())

    def updateAutoCompletar(self):
        self.cotizacion_widget.update()

    def autoCompletar(self, text):
        if text != "":
            registro = objects.REGISTRO_DATAFRAME[objects.REGISTRO_DATAFRAME["Cotización"] == text]
            if len(registro):
                for (field, widgetT) in zip(self.FIELDS, self.WIDGETS):
                        val = str(registro[field].values[0])
                        if val == "nan": val = ""
                        widget = eval("self.%s_widget"%widgetT)
                        widget.setText(val)

    def sendCorreo(self):
        file_name = self.cotizacion_widget.text()
        pago = self.tipo_widget.text()
        if pago == "Factura":
            fileName, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Orden de servicios", "", "Portable Document Format (*.pdf)")
            if fileName:
                cwd = os.getcwd()
                if cwd[0] == "\\":
                    quit_msg = "No es posible leer desde un computador en red."
                    QtWidgets.QMessageBox.warning(self, 'Error',
                                     quit_msg, QtWidgets.QMessageBox.Ok)
                else:
                    self.dialog = CorreoDialog((file_name, fileName), target = correo.sendGestorFactura)
                    self.dialog.start()
                    self.dialog.exec_()
                    if self.dialog.exception != None:
                        raise(self.dialog.exception)
                    else: self.clean()
        elif pago == "Recibo":
            self.dialog = CorreoDialog((file_name, ), target = correo.sendGestorRecibo)
            self.dialog.start()
            self.dialog.exec_()
            if self.dialog.exception != None:
                raise(self.dialog.exception)
            else: self.clean()
        else:
            raise(Exception("Los tipos válidos corresponden con: Recibo y Factura."))

    def clean(self):
        self.cotizacion_widget.blockSignals(True)
        for widgetT in self.WIDGETS:
            widget = eval("self.%s_widget"%widgetT)
            widget.setText("")
        self.cotizacion_widget.blockSignals(False)

    def accept(self):
        tipo = self.tipo_widget.text()
        try:
            if tipo != '':
                self.sendCorreo()
            else:
                raise(Exception("La cotización actual no tiene tipo."))
        except Exception as e:
            self.errorWindow(e)

    def errorWindow(self, exception):
        error_text = str(exception)
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setText(error_text)
        msg.setWindowTitle("Error")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, parent = None):
        super(QtWidgets.QMainWindow, self).__init__(parent)
        self.setWindowTitle(config.CENTRO)

        central_widget = QtWidgets.QWidget(self)
        self.setCentralWidget(central_widget)

        self.central_layout = QtWidgets.QHBoxLayout(central_widget)

        buttons_widget = QtWidgets.QWidget(central_widget)
        self.buttons_layout = QtWidgets.QVBoxLayout(buttons_widget)
        self.buttons_layout.setSpacing(0)
        self.buttons_layout.setContentsMargins(0,0,0,0)
        buttons_widget.setSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Expanding)

        self.central_layout.addWidget(buttons_widget)

        self.request_widget = QtWidgets.QPushButton("Solicitar información")
        self.cotizacion_widget = QtWidgets.QPushButton("Generar Cotización")
        self.descontar_widget = QtWidgets.QPushButton("Descontar")
        self.buscar_widget = QtWidgets.QPushButton("Buscar")
        self.open_widget = QtWidgets.QPushButton("Abrir PDFs")
        self.gestor_widget = QtWidgets.QPushButton("A Gestor...")
        self.reporte_widget = QtWidgets.QPushButton("Reportes")

        self.buttons_layout.addWidget(self.cotizacion_widget)
        self.buttons_layout.addWidget(self.descontar_widget)
        self.buttons_layout.addWidget(self.request_widget)
        self.buttons_layout.addWidget(self.buscar_widget)
        self.buttons_layout.addWidget(self.open_widget)
        self.buttons_layout.addWidget(self.gestor_widget)
        self.buttons_layout.addWidget(self.reporte_widget)

        self.request_widget.clicked.connect(self.requestHandler)
        self.cotizacion_widget.clicked.connect(self.cotizacionHandler)
        self.descontar_widget.clicked.connect(self.descontarHandler)
        self.buscar_widget.clicked.connect(self.buscarHandler)
        self.open_widget.clicked.connect(self.openHandler)
        self.gestor_widget.clicked.connect(self.gestorHandler)
        self.reporte_widget.clicked.connect(self.reporteHandler)

        self.request_window = RequestWindow(self)
        self.cotizacion_window = CotizacionWindow(self)
        self.descontar_window = DescontarWindow(self)
        self.buscar_window = BuscarWindow(self)
        self.gestor_window = GestorWindow(self)
        self.reporte_window = ReporteWindow(self)

        # self.cotizacion_window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)

        self.request_window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)

        self.mdi = QtWidgets.QMdiArea()
        self.mdi.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.central_layout.addWidget(self.mdi)

        self.mdi.addSubWindow(self.cotizacion_window)
        self.mdi.addSubWindow(self.request_window)
        self.mdi.addSubWindow(self.descontar_window)
        self.mdi.addSubWindow(self.buscar_window)
        self.mdi.addSubWindow(self.gestor_window)
        self.mdi.addSubWindow(self.reporte_window)

        self.cotizacion_window.hide()
        self.request_window.hide()
        self.descontar_window.hide()
        self.buscar_window.hide()
        self.gestor_window.hide()
        self.reporte_window.hide()

        self.update_timer = QtCore.QTimer()
        self.update_timer.setInterval(1000)
        self.update_timer.timeout.connect(self.updateDataFrames)
        self.update_timer.start()

        self.resize(1000, 800)
        self.centerOnScreen()

    def updateDataFrames(self):
        cli, reg = objects.readDataFrames()
        if (cli is None):
            cli = objects.CLIENTES_DATAFRAME
        if (reg is None):
            reg = objects.REGISTRO_DATAFRAME
        if (not cli.equals(objects.CLIENTES_DATAFRAME)) | (not reg.equals(objects.REGISTRO_DATAFRAME)):
            objects.CLIENTES_DATAFRAME = cli
            objects.REGISTRO_DATAFRAME = reg
            self.cotizacion_window.updateAutoCompletar()
            self.cotizacion_window.setLastCotizacion()
            self.descontar_window.updateDataFrames()
            self.gestor_window.updateAutoCompletar()
            self.buscar_window.updateAutoCompletar()
            self.buscar_window.update()
            self.buscar_window.limpiar()

    def centerOnScreen(self):
        resolution = QtWidgets.QDesktopWidget().screenGeometry()
        self.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))

    def cotizacionHandler(self):
        self.cotizacion_window.show()

    def descontarHandler(self):
        self.descontar_window.show()

    def requestHandler(self):
        self.request_window.show()

    def buscarHandler(self):
        self.buscar_window.table.resizeRowsToContents()
        self.buscar_window.show()

    def openHandler(self):
        path = os.path.dirname(sys.executable)
        path = os.path.join(path, constants.PDF_DIR)
        try:
            os.startfile(path)
        except FileNotFoundError as e:
            print(e)

    def gestorHandler(self):
        self.gestor_window.show()

    def reporteHandler(self):
        self.reporte_window.show()

    def modificarCotizacion(self, cot):
        self.cotizacionHandler()
        self.cotizacion_window.loadCotizacion(cot)

    def abrirPDFCotizacion(self, cot):
        file = os.path.join(constants.PDF_DIR, cot + ".pdf")
        path, old = self.cotizacion_window.openPDF(file)
        Popen('"%s"' % path, shell = True)

    def closeEvent(self, event):
        windows = [self.cotizacion_window, self.descontar_window, self.request_window, self.buscar_window, self.gestor_window]
        suma = sum([window.is_closed for window in windows])
        if suma == len(windows):
            event.accept()
        else:
            quit_msg = "Existen ventanas sin cerrar.\n¿Está seguro que desea cerrar el programa?"
            reply = QtWidgets.QMessageBox.question(self, 'Cerrar',
                             quit_msg, QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.Yes:
                for window in windows:
                    if not window.is_closed: window.close()
                event.accept()
            else:
                event.ignore()

class CalendarWidget(QtWidgets.QDateTimeEdit):
    def __init__(self, parent = None):
        now = datetime.now()
        super(CalendarWidget, self).__init__(now)
        self.setDisplayFormat("dd/MM/yyyy")
        self.setCalendarPopup(True)

        self.setMaximumDate(now)

class ReporteWindow(SubWindow):
    def __init__(self, parent = None):
        super(ReporteWindow, self).__init__(parent)

        self.setWindowTitle("Reporte Interno")

        self.parent = parent
        wid = QtWidgets.QWidget(self)
        self.setWidget(wid)

        self.layout = QtWidgets.QGridLayout(wid)

        self.layout.setSpacing(3)
        self.layout.setContentsMargins(6, 6, 6, 6)
        wid.setSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)

        from_ = QtWidgets.QLabel("From:")
        to_ = QtWidgets.QLabel("To:")
        send_to = QtWidgets.QLabel("Send to:")
        include = QtWidgets.QLabel("Include:")

        self.from_widget = CalendarWidget()
        self.from_widget.setDate(self.from_widget.date().addDays(-7))
        self.to_widget = CalendarWidget()

        self.send_to = QtWidgets.QLineEdit()
        self.send_widget = QtWidgets.QPushButton("Send")
        self.text_widget = QtWidgets.QLabel("0")

        self.layout.addWidget(from_, 0, 0)
        self.layout.addWidget(self.from_widget, 0, 1)
        self.layout.addWidget(to_, 1, 0)
        self.layout.addWidget(self.to_widget, 1, 1)
        self.layout.addWidget(send_to, 2, 0)
        self.layout.addWidget(self.send_to, 2, 1)
        self.layout.addWidget(self.text_widget, 3, 1)
        self.layout.addWidget(self.send_widget, 4, 1)

        self.send_widget.clicked.connect(self.send)
        self.from_widget.dateChanged.connect(self.changeDate)
        self.to_widget.dateChanged.connect(self.changeDate)

        self.excel = None
        self.changeDate()

        self.setFixedSize(400, 250)

    def changeDate(self):
        start_date = self.from_widget.date().toPyDate()
        end_date =  self.to_widget.date().toPyDate()

        start_date = datetime.combine(start_date, datetime.min.time())
        end_date = datetime.combine(end_date, datetime.max.time())

        dates = pd.to_datetime(objects.REGISTRO_DATAFRAME["Fecha"])
        mask = (dates > start_date) & (dates <= end_date)

        self.excel = objects.REGISTRO_DATAFRAME[mask]
        self.setText()

    def setText(self):
        n = len(self.excel)
        total = len(objects.REGISTRO_DATAFRAME)
        self.text_widget.setText("Cotizaciones seleccionadas: %d/%d" % (n, total))

    def send(self):
        e_to = self.send_to.text()
        try:
            if e_to != "":
                if not "@" in e_to: e_to += '@uniandes.edu.co'
                if len(self.excel):
                    self.excel.to_excel(config.REPORTE_INTERNO, index = False)
                    self.dialog = CorreoDialog([e_to], correo.sendReporteExcel)
                    self.dialog.start()
                    self.dialog.exec_()
                else:
                    raise(Exception("No existen cotizaciones realizadas en esas fechas."))
            else:
                raise(Exception("El correo no es válido."))
        except Exception as e:
            error_text = str(e)
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Warning)
            msg.setText(error_text)
            msg.setWindowTitle("Error")
            msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
            msg.exec_()

class RequestWindow(SubWindow):
    def __init__(self, parent = None):
        super(RequestWindow, self).__init__(parent)
        self.setWindowTitle("Solicitar información")

        wid = QtWidgets.QWidget(self)
        self.setWidget(wid)

        self.layout = QtWidgets.QHBoxLayout(wid)


        form1 = QtWidgets.QFrame()
        hlayout1 = QtWidgets.QHBoxLayout(form1)

        label = QtWidgets.QLabel("Correo:")
        self.correo_widget = QtWidgets.QLineEdit()
        self.enviar_button = QtWidgets.QPushButton("Enviar")
        self.correo_widget.setFixedWidth(250)

        hlayout1.addWidget(label)
        hlayout1.addWidget(self.correo_widget)
        hlayout1.addWidget(self.enviar_button)
        self.layout.addWidget(form1)

        self.correo_widget.returnPressed.connect(self.sendCorreo)
        self.enviar_button.clicked.connect(self.sendCorreo)

        self.setFixedSize(400, 100)

    def sendCorreo(self):
        text = self.correo_widget.text()
        try:
            if not "@" in text: text += '@uniandes.edu.co'
            self.dialog = CorreoDialog((text,), correo.sendRequest)
            self.dialog.start()
            self.dialog.exec_()
            if self.dialog.exception != None:
                raise(self.dialog.exception)
            else:
                self.close()
            # else: raise(Exception("Correo no válido."))
        except Exception as e:
            self.errorWindow(e)

    def errorWindow(self, exception):
        error_text = str(exception)
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setText(error_text)
        msg.setWindowTitle("Error")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    QtWidgets.QApplication.setStyle(QtWidgets.QStyleFactory.create('Fusion')) # <- Choose the style

    app.processEvents()
    main = MainWindow()

    main.show()

    app.exec_()
