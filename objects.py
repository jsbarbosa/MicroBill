import os
import pickle
import numpy as np
import pandas as pd
from datetime import datetime

import constants
from pdflib import PDFCotizacion, PDFReporte

from unidecode import unidecode

if os.path.isdir(constants.OLD_DIR): pass
else: os.makedirs(constants.OLD_DIR)

def readDataFrames():
    c = pd.read_excel(constants.CLIENTES_FILE).fillna("").astype(str)
    r = pd.read_excel(constants.REGISTRO_FILE).fillna("").astype(str)
    return c, r

CLIENTES_DATAFRAME, REGISTRO_DATAFRAME = readDataFrames()

class Cotizacion(object):
    def __init__(self, numero = None, usuario = None, servicios = [], muestra = None):
        self.numero = numero
        self.usuario = usuario
        self.muestra = muestra
        self.is_pago = False
        self.pdf_file_name = ""
        self.referencia_pago = ""
        self.setServicios(servicios)

    def getUsuario(self):
        return self.usuario

    def getInterno(self):
        return self.usuario.getInterno()

    def getCodigos(self):
        return [servicio.getCodigo() for servicio in self.servicios]

    def getNumero(self):
        return self.numero

    def getTotal(self):
        return sum([servicio.getValorTotal() for servicio in self.servicios])

    def getServicio(self, cod):
        i = self.getCodigos().index(cod)
        return self.servicios[i]

    def getServicios(self):
        return self.servicios

    def getMuestra(self):
        return self.muestra

    def isPago(self):
        return self.is_pago

    def isPagoStr(self):
        if self.isPago():
            return "Pago"
        else:
            return "Pendiente"

    def getReferenciaPago(self):
        return self.referencia_pago

    def getFileName(self):
        return self.pdf_file_name

    def getEstado(self):
        total_cotizadas = 0
        usadas = 0
        for servicio in self.servicios:
            usadas += servicio.getCantidad() - servicio.getRestantes()
            total_cotizadas += servicio.getCantidad()
        return int(np.floor(100*usadas/total_cotizadas))

    def setNumero(self, numero):
        self.numero = numero

    def setUsuario(self, usuario):
        self.usuario = usuario

    def setInterno(self, interno):
        try: self.usuario.setInterno(interno)
        except AttributeError: pass
        for servicio in self.servicios:
            servicio.setInterno(interno)

    def setServicios(self, servicios):
        codigos = [servicio.getCodigo() for servicio in servicios]
        if len(codigos) != len(set(codigos)):
            raise(Exception("Existe un código repetido."))
        else:
            self.servicios = servicios

    def setMuestra(self, muestra):
        self.muestra = muestra

    def setPago(self, ref):
        self.is_pago = True
        self.referencia_pago = ref

    def setFileName(self, name):
        self.pdf_file_name = name

    def removeServicio(self, index):
        del self.servicios[index]

    def addServicio(self, servicio):
        servicios = self.servicios + [servicio]
        self.setServicios(servicios)

    def addServicios(self, servicios):
        servicios = self.servicios + servicios
        self.setServicios(servicios)

    def makeCotizacionTable(self):
        table = []
        for servicio in self.servicios:
            row = servicio.makeCotizacionTable()
            table.append(row)
        return table

    def makeReporteTable(self):
        table = []
        for servicio in self.servicios:
            usos = servicio.makeReporteTable()
            table += usos
        if len(table):
            table = np.array(table, dtype = str)
            pos = np.argsort(table[:, 0])
            table = table[pos]
            return list(table)
        return table

    def makeResumenTable(self):
        table = []
        for servicio in self.servicios:
            row = servicio.makeResumenTable()
            table.append(row)
        return table

    def save(self, to_cotizacion = True):
        file = os.path.join(constants.OLD_DIR, self.numero + ".pkl")
        with open(file, 'wb') as output:
            pickle.dump(self, output, pickle.HIGHEST_PROTOCOL)
        if to_cotizacion:
            self.usuario.save()
            self.toRegistro()
            self.makePDFCotizacion()
        else:
            self.makePDFReporte()

    def makePDFReporte(self):
        PDFReporte(self).doAll()

    def makePDFCotizacion(self):
        PDFCotizacion(self).doAll()

    def toRegistro(self):
        global REGISTRO_DATAFRAME

        last = REGISTRO_DATAFRAME.shape[0]

        usuario = self.getUsuario()
        fields = [self.getNumero(), datetime.now().replace(microsecond = 0),
                    usuario.getNombre(), usuario.getCorreo(), usuario.getTelefono(),
                    usuario.getInstitucion(), usuario.getInterno(), usuario.getResponsable(),
                    self.getMuestra(), self.getServicios()[0].equipo, "%d %%"%self.getEstado(),
                    self.isPagoStr(), self.getReferenciaPago(), usuario.getPago(), "{:,}".format(self.getTotal())]

        REGISTRO_DATAFRAME.loc[last] = fields

        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.drop_duplicates("Cotización", "last")
        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.sort_values("Cotización", ascending = False)
        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.reset_index(drop = True)

        writer = pd.ExcelWriter(constants.REGISTRO_FILE, engine='xlsxwriter',
                    datetime_format= "dd/mm/yy hh:mm:ss")

        REGISTRO_DATAFRAME.to_excel(writer, index = False)

    def load(self, file):
        file = os.path.join(constants.OLD_DIR, file + ".pkl")
        with open(file, "rb") as data:
            return pickle.load(data)

class Usuario(object):
    def __init__(self, nombre = None, correo = None, institucion = None, documento = None,
                 direccion = None, ciudad = None, telefono = None, interno = None, responsable = None,
                 proyecto = None, codigo = None, pago = None):
        self.nombre = nombre
        self.correo = correo
        self.institucion = institucion
        self.documento = documento
        self.direccion = direccion
        self.ciudad = ciudad
        self.telefono = telefono
        self.interno = interno
        self.responsable = responsable
        self.proyecto = proyecto
        self.codigo = codigo
        self.pago = pago

    def getNombre(self):
        return self.nombre

    def getCorreo(self):
        return self.correo

    def getInstitucion(self):
        return self.institucion

    def getDocumento(self):
        return self.documento

    def getDireccion(self):
        return self.direccion

    def getCiudad(self):
        return self.ciudad

    def getTelefono(self):
        return self.telefono

    def getInterno(self):
        return self.interno

    def getResponsable(self):
        return self.responsable

    def getProyecto(self):
        return self.proyecto

    def getCodigo(self):
        return self.codigo

    def getPago(self):
        return self.pago

    def setNombre(self, nombre):
        self.nombre = nombre

    def setCorreo(self, correo):
        self.correo = correo

    def setInstitucion(self, institucion):
        self.institucion = institucion

    def setDocumento(self, documento):
        self.documento = documento

    def setDireccion(self, direccion):
        self.direccion = direccion

    def setCiudad(self, ciudad):
        self.ciudad = ciudad

    def setTelefono(self, telefono):
        self.telefono = telefono

    def setInterno(self, interno):
        self.interno = interno

    def setResponsable(self, responsable):
        self.responsable = responsable

    def setProyecto(self, proyecto):
        self.proyecto = proyecto

    def setCodigo(self, codigo):
        self.codigo = codigo

    def setPago(self, pago):
        self.pago = pago

    def save(self):
        global CLIENTES_DATAFRAME

        last = CLIENTES_DATAFRAME.shape[0]

        fields = []
        for key in CLIENTES_DATAFRAME.keys():
            key = unidecode(key)
            try:
                fields.append(eval("self.get%s()"%key))
            except: pass

        fields.append(self.getPago())

        CLIENTES_DATAFRAME.loc[last] = fields
        CLIENTES_DATAFRAME = CLIENTES_DATAFRAME.drop_duplicates("Nombre", "last")
        CLIENTES_DATAFRAME = CLIENTES_DATAFRAME.sort_values("Nombre")
        CLIENTES_DATAFRAME.to_excel(constants.CLIENTES_FILE, index = False, na_rep = '')

class Servicio(object):
    def __init__(self, equipo = None, codigo = None, interno = None, cantidad = None, usos = None):
        self.equipo = equipo
        self.codigo = codigo
        self.cantidad = cantidad
        if usos == None: self.usos = {}
        else: self.usos = usos

        self.interno = interno

        self.valor_unitario = None
        self.valor_total = None
        self.descripcion = None
        self.restantes = None

        self.setRestantes()
        self.setValorUnitario()
        self.setValorTotal()
        self.setDescripcion()

    def getEquipo(self):
        return self.equipo

    def getCodigo(self):
        return self.codigo

    def getCantidad(self):
        return self.cantidad

    def getUsos(self):
        return self.usos

    def getValorUnitario(self):
        return self.valor_unitario

    def getValorTotal(self):
        return self.valor_total

    def getDescripcion(self):
        return self.descripcion

    def getRestantes(self):
        return self.restantes

    def setEquipo(self, equipo):
        self.equipo = equipo
        self.setDescripcion()
        self.setValorUnitario()
        self.setValorTotal()

    def setCodigo(self, codigo):
        self.codigo = codigo
        self.setDescripcion()
        self.setValorUnitario()
        self.setValorTotal()

    def setCantidad(self, cantidad):
        self.cantidad = cantidad
        self.setRestantes()
        self.setValorTotal()

    def setUsos(self, usos):
        self.usos = usos

    def setValorUnitario(self, valor = None):
        if valor == None:
            equipo = eval("constants.%s"%self.equipo)
            df = equipo[equipo["Código"] == self.codigo]
            if len(df) == 0: raise(Exception("Código inválido."))
            self.valor_unitario = int(df[self.interno].values[0])
        else:
            self.valor_unitario = valor

    def setValorTotal(self, valor = None):
        if valor == None:
            self.valor_total = int(self.getValorUnitario() * self.getCantidad())
        else:
            self.valor_total = valor

    def setDescripcion(self, valor = None):
        if valor == None:
            equipo = eval("constants.%s"%self.equipo)
            df = equipo[equipo["Código"] == self.codigo]
            self.descripcion = df["Descripción"].values[0]
        else:
            self.descripcion = valor

    def setRestantes(self):
        self.restantes = self.cantidad - sum(self.usos.values())

    def setInterno(self, interno):
        self.interno = interno
        old_valor = self.getValorUnitario()
        old_total = self.getValorTotal()
        old_cantidad = self.getCantidad()

        self.setValorUnitario()
        if old_total == int(old_valor * old_cantidad):
            self.setValorTotal()
        else:
            cantidad = old_total / self.getValorUnitario()
            cantidad = np.ceil(10 * cantidad) / 10
            self.setCantidad(cantidad)
            self.setValorTotal(old_total)

    def descontar(self, n):
        if self.restantes >= n:
            if n > 0:
                today = datetime.strftime(datetime.now(), "%Y/%m/%d")
                if today in self.usos.keys():
                    self.usos[today] += n
                else:
                    self.usos[today] = n
                self.restantes -= n
        else:
            raise(Exception("No es posible descontar tanto."))

    def makeCotizacionTable(self):
        return [self.getCodigo(), self.getDescripcion(), "%.1f"%self.getCantidad(),
                "{:,}".format(self.getValorUnitario()), "{:,}".format(self.getValorTotal())]

    def makeReporteTable(self):
        fechas = sorted(list(self.usos.keys()))
        n = len(fechas)
        table = [None]*n
        cantidad = self.getCantidad()
        for i in range(n):
            fecha = fechas[i]
            usados = self.usos[fecha]
            restantes = cantidad - usados
            table[i] = [fecha, self.getCodigo(), self.getDescripcion(), "%.1f"%self.getCantidad(), "%.1f"%usados, "%.1f"%restantes]
            cantidad = restantes
        return table

    def makeResumenTable(self):
        return [self.getDescripcion(), "%.1f"%self.getCantidad(), "%.1f"%self.getRestantes()]
