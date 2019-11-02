import os
import sys
import time
import pickle
import numpy as np
import pandas as pd
from copy import copy
from datetime import datetime

from . import constants, config
from .utils import export
from .exceptions import *
from .pdflib import PDFCotizacion, PDFReporte

from unidecode import unidecode
from typing import Callable, Iterable

if os.path.isdir(constants.OLD_DIR):
    pass
else:
    os.makedirs(constants.OLD_DIR)

__all__ = ["LAST_MODIFICATION_CLIENTES", "LAST_MODIFICATION_REGISTRO", "CLIENTES_DATAFRAME", "REGISTRO_DATAFRAME"]

LAST_MODIFICATION_CLIENTES = 0  #: tiempo de la última modificación del archivo de clientes
LAST_MODIFICATION_REGISTRO = 0  #: tiempo de la última modificación del archivo de registro


@export
def readDataFrames():
    """ Función que revisa si existen modificaciones en los archivos de clientes y registro.

    Returns
    -------
    clientes: None si no hay cambios en el archivo de clientes, dataframe en caso contrario
    regirstro: None si no hay cambios en el archivo de registro, dataframe en caso contrario
    """

    global LAST_MODIFICATION_CLIENTES, LAST_MODIFICATION_REGISTRO
    t1 = os.path.getmtime(constants.CLIENTES_FILE)
    t2 = os.path.getmtime(constants.REGISTRO_FILE)
    c = None
    r = None
    if t1 > LAST_MODIFICATION_CLIENTES:
        c = pd.read_excel(constants.CLIENTES_FILE).fillna("").astype(str)
        LAST_MODIFICATION_CLIENTES = t1
    if t2 > LAST_MODIFICATION_REGISTRO:
        r = pd.read_excel(constants.REGISTRO_FILE).fillna("").astype(str)
        LAST_MODIFICATION_REGISTRO = t2
    return c, r


CLIENTES_DATAFRAME, REGISTRO_DATAFRAME = readDataFrames()


@export
def getNumeroCotizacion(equipo: str) -> str:
    """ Función que usando el nombre del equipo que entra por parámetro retorna el último número de la cotización + 1
    asociada a dicho equipo en formato ____

    Parameters
    ----------
    equipo: str
        equipo o servicio al que se le desea obtener un número nuevo de cotización

    Returns
    -------
    cod: str
        número de cotización nuevo
    """

    global REGISTRO_DATAFRAME
    year = str(datetime.now().year)[-2:]
    try:
        cot = REGISTRO_DATAFRAME[REGISTRO_DATAFRAME["Equipo"] == equipo]["Cotización"].values[0]
        cod, val = cot.split("-")
    except IndexError:
        cod = equipo[0] + year
        val = "%04d" % 0

    if "_" in equipo:
        cod = equipo.split('_')[1] + cod[-2:]

    if year != cod[-2:]:
        cod = cod[:-2] + year
        val = "%04d" % 1
    else:
        val = "%04d" % (int(val) + 1)

    cod = "%s%s-%s" % (config.PREFIJO, cod, val)
    return cod


@export
def getEquipoName(codigo: str) -> str:
    """ Función que retorna el nombre del equipo a partir de su código

    Parameters
    ----------
    codigo: str
        código del equipo

    Returns
    -------
    str: nombre del equipo, o None en caso de que no exista un nombre asociado a ese código
    """

    for name in constants.EQUIPOS:
        if codigo in name:
            return name
    return None


@export
def sortServicios(servicios: Iterable) -> dict:
    """ Función que usando los servicios que entran por parámetro, los ordena en un diccionario asociandolos por equipo

    Parameters
    ----------
    servicios: Iterable
        servicios que se desean agrupar por equipo

    Returns
    -------
    dict: servicios agrupados por equipo
    """

    dic = {}
    for servicio in servicios:
        equipo = servicio.getEquipo()
        if equipo in dic.keys():
            dic[equipo].append(servicio)
        else:
            dic[equipo] = [servicio]
    return dic


@export
def obtenerTipoUsuario(correo: str) -> str:
    """ Función que clasifica a un usuario como Interno, Independiente, Académico, o Industria a partir de su correo

    Parameters
    ----------
    correo: str
        correo del usuario

    Returns
    -------
    str: Interno o Independiente o Académico o Industria

    """

    from .config import CORREOS_CONVENIOS
    convenios = sum([1 for convenio in CORREOS_CONVENIOS if convenio in correo])
    if convenios > 0:
        suma = sum(constants.INDEPENDIENTES_DF['Correo'] == correo)
        if suma:
            return 'Independiente'
        return 'Interno'
    elif '.edu.' in correo:
        return 'Académico'
    else:
        return 'Industria'


@export
def crearCotizaciones(usuario, servicios: Iterable, muestra: str, observaciones_pdf: str = "",
                      observaciones_correo: str = "", descuento: str = ''):
    """ Función que crea las cotizaciones asociadas a varios servicios de un mismo usuario

    Parameters
    ----------
    usuario: objects.Usuario
        usuario al cual se le realizan las cotizaciones
    servicios: Iterable
        servicios a incluir en las cotizaciones
    muestra: str
        tipo de muestra
    observaciones_pdf: str
        observaciones a incluir en el PDF de las cotizaciones
    observaciones_correo: str
        observaciones a incluir en el correo de las cotizaciones
    descuento: str
        motivo por el cual se realiza el descuento

    Returns
    -------
    list: lista de cotizaciones (una por cada equipo)
    """

    cotizacion = Cotizacion()
    cotizacion.setUsuario(usuario)
    servicios = sortServicios(servicios)
    cotizacion.setMuestra(muestra)
    cotizacion.setObservacionPDF(observaciones_pdf)
    cotizacion.setObservacionCorreo(observaciones_correo)

    cotizaciones = []
    for key in servicios:
        cot = copy(cotizacion)
        cot.setDescuentoText(descuento)
        cot.setServicios(servicios[key])
        cot.setNumero(getNumeroCotizacion(key))
        cotizaciones.append(cot)

    return cotizaciones


@export
def generarPDFs(cotizaciones: Iterable):
    """ Función que genera los PDFs de las cotizaciones que entran por parámetro

    Parameters
    ----------
    cotizaciones: Iterable
        cotizaciones a las cuales se les busca general el pdf
    """

    for cot in cotizaciones:
        cot.makePDFCotizacion()


@export
def guardarCotizaciones(cotizaciones):
    """ Función que guarda las cotizaciones que entran por parámetro en el registro

    Parameters
    ----------
    cotizaciones: Iterable
        cotizaciones a las cuales se les busca guardar en el archivo de registro
    """

    for cot in cotizaciones:
        cot.save(to_pdf=False)


@export
class Cotizacion(object):
    """ Clase Cotización, la clase cotización cuenta con un usuario de la clase Usuario, y varios servicios asociados
    a un mismo equipo
    """
    def __init__(self, numero=None, usuario=None, servicios: Iterable = [], muestra: str = None):
        self.numero = numero
        self.usuario = usuario
        self.muestra = muestra
        self.is_pago = False
        self.servicios = None
        self.pdf_file_name = ""
        self.pdf_path = ""
        self.referencia_pago = ""
        self.elaborado_por = ""
        self.modificado_por = ""
        self.aplicado_por = ""
        self.observacion_correo = ""
        self.observacion_pdf = ""
        self.descuento_text = ""
        self.setServicios(servicios)

    def getUsuario(self):
        """ Método que retorna el usuario de la cotización

        Returns
        -------
            Usuario: el usuario de la cotización
        """

        return self.usuario

    def getInterno(self):
        """ Método que retorna el atributo interno del usuario de la cotización

        Returns
        -------
            str: categoría del usuario
        """

        return self.usuario.getInterno()

    def internoTreatment(self) -> bool:
        """ Método que retorna si la cotización está asociada a un usuario interno o independiente

        Returns
        -------
        bool:
            True si la cotización está asociada a un interno o independiente
        """

        interno = self.getInterno()
        if interno in ["Interno", "Independiente"]:
            return True
        return False

    def getPago(self) -> str:
        """ Método que retorna el tipo de pago asociado al usuario

        Returns
        -------
        str: tipo de pago asociado al usuario
        """

        return self.usuario.getPago()

    def getCodigos(self) -> list:
        """ Método que retorna los códigos de los servicios asociados a la cotización

        Returns
        -------
        list: códigos de los servicios asociados a la cotización
        """

        return [servicio.getCodigo() for servicio in self.servicios]

    def getCodigosPrefix(self) -> list:
        """ Método que retorna los prefijos códigos de los servicios asociados a la cotización

        Returns
        -------
        list: prefijos de los códigos de los servicios asociados a la cotización
        """

        return [servicio.getCodigoPrefix() for servicio in self.servicios]

    def getNumero(self) -> str:
        """ Método que retorna el número de la cotización

        Returns
        -------
        str: número de la cotización
        """

        return self.numero

    def getTotal(self) -> int:
        """ Método que retorna el total a pagar por la cotización

        Returns
        -------
        int: total a pagar por la cotización
        """

        return sum([servicio.getTotal() for servicio in self.servicios])

    def getSubtotal(self) -> int:
        """ Método que retorna el subtotal a pagar por la cotización

        Returns
        -------
        int: subtotal a pagar por la cotización
        """

        return sum([servicio.getValorTotal() for servicio in self.servicios])

    def getDescuentos(self) -> int:
        """ Método que retorna el total de descuentos de la cotización

        Returns
        -------
        int: total de descuentos de la cotización
        """

        return sum([servicio.getDescuentoTotal() for servicio in self.servicios])

    def getServicio(self, cod: str):
        """ Método que retorna el servicio con codigo cod

        Parameters
        ----------
        cod: str
            código del servicio a retornar

        Returns
        -------
        Servicio: servicio asociado al código
        """

        i = self.getCodigosPrefix().index(cod)
        return self.servicios[i]

    def getServicios(self) -> Iterable:
        """ Método que retorna todos los servicios asociados a la cotización

        Returns
        -------
        Iterable: todos los servicios asociados a la cotización
        """

        return self.servicios

    def getMuestra(self) -> str:
        """ Método que el tipo de muestra de la cotización

        Returns
        -------
        str: tipo de muestra de la cotización
        """

        return self.muestra

    def getObservacionPDF(self) -> str:
        """ Método que retorna la observación al PDF de la cotización

        Returns
        -------
        str: observación al PDF de la cotización
        """

        return self.observacion_pdf

    def getObservacionCorreo(self) -> str:
        """ Método que retorna la observación al correo de la cotización

        Returns
        -------
        str: observación al correo de la cotización
        """

        return self.observacion_correo

    def getValorRestante(self) -> int:
        """ Método que retorna el valor en dinero que aún no ha sido usado de la cotización

        Returns
        -------
        str: valor en dinero que aún no ha sido usado de la cotización
        """

        return self.getTotal() - self.getDineroUsado()

    def isPago(self) -> bool:
        """ Método que retorna si la cotización ha sido pagada

        Returns
        -------
        bool: True si la cotización ha sido pagada
        """

        return self.is_pago

    def isPagoStr(self) -> str:
        """ Método que retorna si la cotización ha sido paga como un string

        Returns
        -------
        str: Pagado si ha sido paga, Pendiente si no ha sido pagada
        """

        if self.isPago():
            return "Pago"
        else:
            return "Pendiente"

    def getReferenciaPago(self) -> str:
        """ Método que retorna la referencia de pago de la cotización

        Returns
        -------
        str: referencia de pago de la cotización
        """

        return self.referencia_pago

    def getFileName(self) -> str:
        """ Método que retorna el nombre del archivo pdf asociado a la cotización

        Returns
        -------
        str: nombre del archivo pdf asociado a la cotización
        """

        return self.pdf_file_name

    def getFilePath(self) -> str:
        """ Método que retorna la ruta del archivo pdf asociado a la cotización

        Returns
        -------
        str: ruta del archivo pdf asociado a la cotización
        """

        return self.pdf_path

    def getEstado(self) -> int:
        """ Método que el porcentaje de uso de la cotización

        Returns
        -------
        int: porcentaje de uso de la cotización
        """

        total_cotizadas = 0
        usadas = 0
        for servicio in self.servicios:
            usadas += servicio.getCantidad() - servicio.getRestantes()
            total_cotizadas += servicio.getCantidad()
        return int(np.floor(100 * usadas / total_cotizadas))

    def getAplicado(self) -> str:
        """ Método que retorna el nombre de la persona que aplicó el pago de la cotización

        Returns
        -------
        str: nombre de la persona que aplicó el pago de la cotización
        """

        return self.aplicado_por

    def getElaborado(self) -> str:
        """ Método que retorna el nombre de la persona que elaboró la cotización

        Returns
        -------
        str: nombre de la persona que elaboró la cotización
        """

        return self.elaborado_por

    def getModificado(self) -> str:
        """ Método que retorna el nombre de la persona que modificó la cotización

        Returns
        -------
        str: nombre de la persona que modificó la cotización
        """

        return self.modificado_por

    def getDineroUsado(self) -> int:
        """ Método que retorna el total de dinero usado de la cotización

        Returns
        -------
        int: total de dinero usado de la cotización
        """

        dinero = [servicio.getDineroUsado() for servicio in self.getServicios()]
        return sum(dinero)

    def getDescuentoText(self) -> str:
        """ Método que retorna el texto asociado al descuento de la cotización

        Returns
        -------
        str: texto asociado al descuento de la cotización
        """

        return self.descuento_text

    def setNumero(self, numero: str):
        """ Método que asigna el número de la cotización que entra por parámetro como el valor del atributo numero

        Parameters
        ----------
        numero: str
            número de la cotización
        """

        self.numero = numero

    def setUsuario(self, usuario):
        """ Método que asigna el usuario que entra por parámetro como el valor del atributo usuario

        Parameters
        ----------
        usuario: Usuario
            usuario asociado a la cotización
        """

        self.usuario = usuario

    def setInterno(self, interno: str):
        """ Método que asigna el tipo de usuario que entra por parámetro como el valor del atributo interno

        Parameters
        ----------
        interno: str
            tipo de usuario asociado a la cotización
        """

        try:
            self.usuario.setInterno(interno)
        except AttributeError:
            pass
        for servicio in self.servicios:
            servicio.setInterno(interno)

    def setServicios(self, servicios: Iterable):
        """ Método que asigna los servicios que entran por parámetro como el atributo servicios

        Parameters
        ----------
        servicios: Iterable
            servicios asociados a la cotización
        """

        codigos = [(servicio.getEquipo() + servicio.getCodigo()) for servicio in servicios]
        if len(codigos) != len(set(codigos)):
            raise Exception("Existe un código repetido.")
        else:
            self.servicios = servicios

    def setMuestra(self, muestra: str):
        """ Método que asigna el tipo de muestra que entra por parámetro como el valor del atributo muestra

        Parameters
        ----------
        muestra: str
            tipo de muestra asociada a la cotización
        """

        self.muestra = muestra

    def setPago(self, ref: str):
        """ Método que asigna el la referencia de pago que entra por parámetro como el valor del atributo
        referencia_pago y asigna el atributo is_pago a True

        Parameters
        ----------
        ref: str
            referencia de pago de la cotización
        """

        self.is_pago = True
        self.referencia_pago = ref

    def setFileName(self, name: str):
        """ Método que asigna el nombre del archivo PDF de la cotización que entra por parámetro como el valor del
        atributo pdf_file_name

        Parameters
        ----------
        name: str
            nombre del archivo PDF de la cotización
        """

        self.pdf_file_name = name

    def setPath(self, name: str):
        """ Método que asigna la ruta del archivo PDF de la cotización que entra por parámetro como el valor del
        atributo pdf_path y pdf_file_name

        Parameters
        ----------
        name: str
            ruta del archivo PDF de la cotización
        """

        self.pdf_path = name
        self.pdf_file_name = os.path.splitext(os.path.basename(self.pdf_path))[0]

    def setElaborado(self, name: str):
        """ Método que asigna el nombre de la persona que elaboró la cotización que entra por parámetro como el
        valor del atributo elaborado_por

        Parameters
        ----------
        name: str
            nombre de la persona que elaboró la cotización
        """

        if name != "":
            self.elaborado_por = name
        else:
            raise Exception("No se especifica quién realiza la cotización.")

    def setModificado(self, name: str):
        """ Método que asigna el nombre de la persona que modificó la cotización que entra por parámetro como el valor
        del atributo modificado_por

        Parameters
        ----------
        name: str
            nombre de la persona que modificó la cotización
        """

        if name != "":
            self.modificado_por = name
        else:
            raise Exception("No se especifica quién realiza la modificación.")

    def setAplicado(self, name: str):
        """ Método que asigna el nombre de la persona que aplica el pago de la cotización que entra por parámetro como
        el valor del atributo aplicado_por

        Parameters
        ----------
        name: str
            nombre de la persona que aplica el pago de la cotización
        """

        if name != "":
            self.aplicado_por = name
        else:
            raise Exception("No se especifica quién realiza la modificación.")

    def setObservacionPDF(self, text: str):
        """ Método que asigna el texto de observaciones al PDF de la cotización que entra por parámetro como el valor
        del atributo observacion_pdf

        Parameters
        ----------
        text: str
            texto de observaciones al PDF de la cotización
        """

        self.observacion_pdf = text

    def setObservacionCorreo(self, text: str):
        """ Método que asigna el texto de las observaciones al correo de la cotización que entra por parámetro
        como el valor del atributo observacion_correo

        Parameters
        ----------
        text: str
            texto de observaciones al correo de la cotización
        """

        self.observacion_correo = text

    def setDescuentoText(self, text: str):
        """ Método que asigna el texto de descuento de la cotización que entra por parámetro como el valor del atributo
         descuento_text

        Parameters
        ----------
        text: str
            texto de descuento
        """

        self.descuento_text = text
        for s in self.servicios:
            s.setDescuentoText(text)

    def removeServicio(self, index: int):
        """ Método que elimina el servicio asociado al índice que entra por parámetro

        Parameters
        ----------
        index: int
            índice del servicio a eliminar
        """

        del self.servicios[index]

    def addServicio(self, servicio):
        """ Método que agrega el servicio a la cotización

        Parameters
        ----------
        servicio: Servicio
            servicio a agregar en la cotización
        """

        servicios = self.servicios + [servicio]
        self.setServicios(servicios)

    def addServicios(self, servicios: list):
        """ Método que agrega los servicios que entran por parámetro a la cotización

        Parameters
        ----------
        servicios: list
            servicios que serán agregados a la cotización
        """

        servicios = self.servicios + servicios
        self.setServicios(servicios)

    def makeCotizacionTable(self):
        """ Método que genera la tabla que es usada por PDFCotizacion con la información de la cotización

        Returns
        -------
        list: lista en donde cada fila se encuentra la información de makeCotizacionTable por cada servicio de la
        cotización
        """

        table = []
        for servicio in self.servicios:
            table += servicio.makeCotizacionTable()
        return table

    def makeReporteTable(self):
        """ Método que genera la tabla que es usada por PDFReporte con la información de usos de los servicios
        de la cotización

        Returns
        -------
        list: lista en donde cada fila se encuentra la información de makeReporteTable por cada servicio de la
        cotización
        """

        table = []
        for servicio in self.servicios:
            usos = servicio.makeReporteTable()
            table += usos
        if len(table):
            table = np.array(table, dtype=str)
            pos = np.argsort(table[:, 0])
            table = table[pos]
            return list(table)
        return table

    def makeResumenTable(self):
        """ Método que genera la tabla que es usada por PDFReporte con la información de los servicios
        de la cotización

        Returns
        -------
        list: lista en donde cada fila se encuentra la información de makeResumenTable por cada servicio de la
        cotización
        """

        table = []
        for servicio in self.servicios:
            row = servicio.makeResumenTable()
            table.append(row)
        return table

    def save(self, to_cotizacion: bool = True, to_reporte: bool = False, to_pdf: bool = True):
        """ Método que guarda una cotización como pickle. Si to_cotizacion == True, la información se guarda en los
        archivos de registro (tanto el registro del usuario como el registro de cotizaciones); si to_reporte == to_pdf
        == True se genera el PDF del reporte; si solo to_pdf == True, se genera el PDF de la cotización
        """

        file = os.path.join(constants.OLD_DIR, self.numero + ".pkl")
        with open(file, 'wb') as output:
            pickle.dump(self, output, pickle.HIGHEST_PROTOCOL)
        if to_cotizacion:
            self.usuario.save()
            self.toRegistro()
        elif to_pdf and to_reporte:
            self.makePDFReporte()
        elif to_pdf:
            self.makePDFCotizacion()

    def makePDFReporte(self):
        """ Método que renderiza el PDF de reporte de una cotización """
        PDFReporte(self).doAll()

    def makePDFCotizacion(self):
        """ Método que renderiza el PDF de una cotización """
        PDFCotizacion(self).doAll()

    def toRegistro(self):
        """ Método que guarda una cotización en el archivo de Registro """

        global REGISTRO_DATAFRAME

        last = REGISTRO_DATAFRAME.shape[0]

        usuario = self.getUsuario()
        fields = [self.getNumero(), datetime.now().replace(microsecond=0),
                  usuario.getNombre(), usuario.getCorreo(), usuario.getTelefono(),
                  usuario.getInstitucion(), usuario.getInterno(), usuario.getResponsable(),
                  self.getMuestra(), self.getServicios()[0].equipo, self.getElaborado(), self.getModificado(),
                  "%d %%" % self.getEstado(), self.isPagoStr(), self.getReferenciaPago(), self.getAplicado(),
                  usuario.getPago(), "{:,}".format(self.getTotal())]

        REGISTRO_DATAFRAME.loc[last] = fields

        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.drop_duplicates("Cotización", "last")
        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.sort_values("Cotización", ascending=False)
        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.reset_index(drop=True)

        path = constants.REGISTRO_FILE

        for i in range(10):
            try:
                writer = pd.ExcelWriter(path, engine='xlsxwriter',
                                        datetime_format="dd/mm/yy hh:mm:ss")

                REGISTRO_DATAFRAME.to_excel(writer, index=False)
                writer.save()
                writer.close()
                break
            except Exception as e:
                if i == 9:
                    raise e

    def limpiar(self):
        """ Método que limpia completamente la cotización """

        self = Cotizacion()

    def load(self, file: str):
        """ Método que carga una cotización vieja en un objeto de la clase Cotizacion

        Parameters
        ----------
        file: str
            nombre del archivo de la cotización a cargar

        Returns
        -------
        Cotizacion: objeto de la clase cotización con la información del archivo que entra por parametro
        """

        file = os.path.join(constants.OLD_DIR, file + ".pkl")
        with open(file, "rb") as data:
            try:
                load = pickle.load(data)
            except ModuleNotFoundError:
                sys.modules['objects'] = sys.modules[__name__]
                load = pickle.load(data)
            return load


@export
class Usuario(object):
    """ Clase Usuario, la clase representa la información de un usuario """
    def __init__(self, nombre: str = None, correo: str = None, institucion: str = None, documento: str = None,
                 direccion: str = None, ciudad: str = None, telefono: str = None, interno: str = None,
                 responsable: str = None, proyecto: str = None, codigo: str = None, pago: str = None, **kwargs):
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

    def getNombre(self) -> str:
        """ Método que retorna el nombre asociado al usuario

        Returns
        -------
        str: nombre asociado al usuario
        """

        return self.nombre

    def getCorreo(self) -> str:
        """ Método que retorna el correo asociado al usuario

        Returns
        -------
        str: correo asociado al usuario
        """

        return self.correo

    def getInstitucion(self) -> str:
        """ Método que retorna de la institución asociada al usuario

        Returns
        -------
        str: nombre de la institución asociada al usuario
        """

        return self.institucion

    def getDocumento(self) -> str:
        """ Método que retorna el documento asociado al usuario

        Returns
        -------
        str: documento asociado al usuario
        """

        return self.documento

    def getDireccion(self) -> str:
        """ Método que retorna la dirección asociada al usuario

        Returns
        -------
        str: dirección asociada al usuario
        """

        return self.direccion

    def getCiudad(self) -> str:
        """ Método que retorna la ciudad de residencia asociada al usuario

        Returns
        -------
        str: ciudad de recidencia asociada al usuario
        """

        return self.ciudad

    def getTelefono(self) -> str:
        """ Método que retorna el teléfono asociado al usuario

        Returns
        -------
        str: teléfono asociado al usuario
        """

        return self.telefono

    def getInterno(self) -> str:
        """ Método que retorna el tipo de usuario

        Returns
        -------
        str: tipo de usuario
        """

        if self.interno is None:
            self.interno = obtenerTipoUsuario(self.correo)
        return self.interno

    def getResponsable(self) -> str:
        """ Método que retorna el nombre del responsable asociado al usuario

        Returns
        -------
        str: nombre del responsable asociado al usuario
        """

        return self.responsable

    def getProyecto(self) -> str:
        """ Método que retorna el nombre del proyecto asociado al usuario

        Returns
        -------
        str: nombre del proyecto asociado al usuario
        """

        return self.proyecto

    def getCodigo(self) -> str:
        """ Método que retorna el código del proyecto asociado al usuario

        Returns
        -------
        str: código del proyecto asociado al usuario
        """

        return self.codigo

    def getPago(self) -> str:
        """ Método que retorna el tipo de pago asociado al usuario

        Returns
        -------
        str: tipo de pago asociado al usuario
        """

        return self.pago

    def setNombre(self, nombre: str):
        """ Método que asigna el nombre que entra por parámetro como el valor del atributo nombre

        Parameters
        ----------
        nombre: str
            nombre al que será asociado el usuario
        """

        self.nombre = nombre

    def setCorreo(self, correo: str):
        """ Método que asigna el correo que entra por parámetro como el valor del atributo correo

        Parameters
        ----------
        correo: str
            correo al que será asociado el usuario
        """

        self.correo = correo

    def setInstitucion(self, institucion: str):
        """ Método que asigna el nombre de la institución que entra por parámetro como el valor del atributo institucion

        Parameters
        ----------
        institucion: str
            nombre de la institución al que será asociado el usuario
        """

        self.institucion = institucion

    def setDocumento(self, documento: str):
        """ Método que asigna el documento que entra por parámetro como el valor del atributo documento

        Parameters
        ----------
        documento: str
            documento al que será asociado el usuario
        """

        self.documento = documento

    def setDireccion(self, direccion: str):
        """ Método que asigna la dirección que entra por parámetro como el valor del atributo direccion

        Parameters
        ----------
        direccion: str
            dirección a la que será asociada el usuario
        """

        self.direccion = direccion

    def setCiudad(self, ciudad: str):
        """ Método que asigna el nombre de la ciudad que entra por parámetro como el valor del atributo ciudad

        Parameters
        ----------
        ciudad: str
            nombre de la ciudad al que será asociado el usuario
        """

        self.ciudad = ciudad

    def setTelefono(self, telefono: str):
        """ Método que asigna el teléfono que entra por parámetro como el valor del atributo telefono

        Parameters
        ----------
        telefono: str
            teléfono al que será asociado el usuario
        """

        self.telefono = telefono

    def setInterno(self, interno: str):
        """ Método que asigna el tipo de usuario que entra por parámetro como el valor del atributo interno

        Parameters
        ----------
        interno: str
            tipo de usuario al que será asociado el usuario
        """

        self.interno = interno

    def setResponsable(self, responsable: str):
        """ Método que asigna el nombre del responsable que entra por parámetro como el valor del atributo responsable

        Parameters
        ----------
        responsable: str
            nombre del responsable al que será asociado el usuario
        """

        self.responsable = responsable

    def setProyecto(self, proyecto: str):
        """ Método que asigna el nombre del proyecto que entra por parámetro como el valor del atributo proyecto

        Parameters
        ----------
        proyecto: str
            nombre del proyecto al que será asociado el usuario
        """

        self.proyecto = proyecto

    def setCodigo(self, codigo: str):
        """ Método que asigna el código del proyecto que entra por parámetro como el valor del atributo codigo

        Parameters
        ----------
        codigo: str
            código del proyecto al que será asociado el usuario
        """

        self.codigo = codigo

    def setPago(self, pago: str):
        """ Método que asigna el tipo de pago que entra por parámetro como el valor del atributo pago

        Parameters
        ----------
        pago: str
            tipo de pago al que será asociado el usuario
        """

        self.pago = pago

    def save(self):
        """ Método que guarda un usuario en el archivo de Clientes """
        global CLIENTES_DATAFRAME

        last = CLIENTES_DATAFRAME.shape[0]

        fields = []
        for key in CLIENTES_DATAFRAME.keys():
            key = unidecode(key)
            try:
                fields.append(eval("self.get%s()" % key))
            except:
                pass

        fields.append(self.getPago())

        CLIENTES_DATAFRAME.loc[last] = fields
        CLIENTES_DATAFRAME = CLIENTES_DATAFRAME.drop_duplicates("Nombre", "last")
        CLIENTES_DATAFRAME = CLIENTES_DATAFRAME.sort_values("Nombre")
        CLIENTES_DATAFRAME.to_excel(constants.CLIENTES_FILE, index=False, na_rep='')

    def __repr__(self):
        fields = ["Nombre", "Correo", "Institucion", "Documento", "Ciudad", "Telefono",
                  "Responsable", "Proyecto", "Codigo", "Pago"]
        data = []
        for text in fields:
            att = getattr(self, text.lower())
            if type(att) is str:
                text = text + ": " + att
            else:
                text = text + ": " + "None"
            data.append(text)
        return "\n".join(data)


@export
class Servicio(object):
    """ Clase Servicio, un servicio contiene el código del servicio, si el servicio es para un usuario interno,
    industria, etc, la cantidad solicitada de este servicio, el número de usos que ha tenido, y si fue agregado
    posteriormente
    """
    def __init__(self, codigo: str = None, interno: str = None, cantidad: (int, float) = None, usos: dict = None,
                 agregado_posteriormente: bool = False):
        if codigo is not None:
            self.codigo_prefix = codigo
            self.codigo = codigo[2:]
            self.equipo = codigo[:2]
            self.equipo = getEquipoName(self.equipo)
        else:
            self.codigo = None
            self.equipo = None
            self.codigo_prefix = None
        self.cantidad = cantidad
        if usos is None:
            self.usos = {}
        else:
            self.usos = usos

        self.interno = interno
        self.agregado_posteriormente = agregado_posteriormente

        self.valor_unitario = None
        self.valor_total = None
        self.descripcion = None
        self.restantes = None

        self.descuento_unitario = None
        self.descuento_total = None
        self.descripcion = None
        self.restantes = None

        self.descuento_text = ""

        self.setRestantes()
        self.setValorUnitario()
        self.setValorTotal()

        self.setDescuentoUnitario()
        self.setDescuentoTotal()
        self.setDescripcion()

    def getEquipo(self) -> str:
        """ Método que retorna el equipo asociado a este servicio

        Returns
        -------
        str: equipo asociado
        """

        return self.equipo

    def getCodigo(self) -> str:
        """ Método que retorna el código asociado a este servicio

        Returns
        -------
        str: código asociado
        """

        return self.codigo

    def getCodigoPrefix(self) -> str:
        """ Método que retorna el prefijo del código del servicio (código del equipo)

        Returns
        -------
        str: prefijo del código
        """

        return self.codigo_prefix

    def getInterno(self) -> str:
        """ Método que retorna la categoria del cliente y del servicio (Interno, Industria, etc)

        Returns
        -------
        str: Interno, Industria, etc
        """
        return self.interno

    def getCantidad(self) -> (int, float):
        """ Método que retorna la cantidad del servicio

        Returns
        -------
        int, float: cantidad del servicio
        """
        return self.cantidad

    def getUsos(self) -> dict:
        """ Método que retorna el diccionario de usos del servicio

        Returns
        -------
        dict: diccionario de usos del servicio
        """

        return self.usos

    def getValorUnitario(self) -> int:
        """ Método que retorna el valor unitario del servicio

        Returns
        -------
        int: valor unitario del servicio
        """

        return self.valor_unitario

    def getValorTotal(self) -> int:
        """ Método que retorna el valor total asociado al servicio (cantidad * valor_unitario)

        Returns
        -------
        int: valor total asociado al servicio
        """

        return self.valor_total

    def getTotal(self) -> int:
        """ Método que retorna el valor a pagar por el servicio

        Returns
        -------
        int: valor a pagar por el servicio
        """

        return self.getValorTotal() - self.getDescuentoTotal()

    def getDescripcion(self) -> str:
        """ Método que retorna la descripción del servicio

        Returns
        -------
        str: descripción del servicio
        """

        return self.descripcion

    def getDescuentoUnitario(self) -> int:
        """ Método que retorna el valor del descuento por unidad del servicio

        Returns
        -------
        int: descuento unitario del servicio
        """

        return self.descuento_unitario

    def getDescuentoTotal(self) -> int:
        """ Método que retorna el descuento total asociado al servicio

        Returns
        -------
        int: descuento total asociado al servicio
        """

        return self.descuento_total

    def getRestantes(self) -> (int, float):
        """ Método que retorna el número de usos restantes

        Returns
        -------
        int, float: numero de usos restantes
        """

        return self.restantes

    def getUsados(self) -> (int, float):
        """ Método que retorna el número de usos que ha tenido el servicio

        Returns
        -------
        int, float: número de usos que ha tenido el servicio
        """

        return sum(self.usos.values())

    def getDineroUsado(self) -> int:
        """ Método que retorna el valor en dinero asociado a las cantidades ya usadas

        Returns
        -------
        int: dinero asociado a las cantidades ya usadas
        """

        total = self.getCantidad()
        if total != 0:
            usados = total - self.getRestantes()
            return int(usados * self.getValorTotal() // total)
        else:
            return int(-self.getRestantes() * self.getValorUnitario())

    def isAgregado(self) -> bool:
        """ Método que retorna si el servicio fue agregado posterior a la realización de la cotización o no

        Returns
        -------
        bool: si el servicio fue agregado posterior a la realización de la cotización o no
        """

        return self.agregado_posteriormente

    def setEquipo(self, equipo: str):
        """ Método que asigna el equipo que entra por parámetro como el valor del atributo equipo

        Parameters
        ----------
        equipo: str
            equipo al que será asociado el servicio
        """

        self.equipo = equipo
        self.setDescripcion()
        self.setValorUnitario()
        self.setValorTotal()

    def setCodigo(self, codigo: str):
        """ Método que asigna el código que entra por parámetro como el valor del atributo codigo

        Parameters
        ----------
        codigo: str
            codigo al que será asociado el servicio
        """

        self.codigo = codigo
        self.setDescripcion()
        self.setValorUnitario()
        self.setValorTotal()

    def setCantidad(self, cantidad: (int, float)):
        """ Método que asigna la cantidad que entra por parámetro como el valor del atributo cantidad

        Parameters
        ----------
        cantidad: int, float
            cantidad al que será asociado el servicio
        """

        self.cantidad = cantidad
        self.setRestantes()
        self.setValorTotal()
        self.setDescuentoTotal()

    def setUsos(self, usos: dict):
        """ Método que asigna los usos que entran por parámetro como el valor del atributo usos

        Parameters
        ----------
        usos: dict
            usos a los que estará asociado el servicio
        """

        self.usos = usos

    def setValorUnitario(self, valor: int = None):
        """ Método que asigna el valor unitario que entra por parámetro como el valor del atributo valor_unitario

        Parameters
        ----------
        valor: int
            valor unitario al que será asociado el servicio
        """

        if valor is None:
            try:
                equipo = eval("constants.%s" % self.equipo)
            except AttributeError:
                raise IncompatibleError
            df = equipo[equipo["Código"] == self.codigo]
            if len(df) == 0:
                raise(Exception("Código inválido."))
            self.valor_unitario = int(df["Industria"].values[0])
        else:
            self.valor_unitario = valor

    def setValorTotal(self, valor: int = None):
        """ Método que asigna el valor total que entra por parámetro como el valor del atributo valor_total

        Parameters
        ----------
        valor: int
            valor total al que será asociado el servicio
        """

        if valor is None:
            self.valor_total = int(self.getValorUnitario() * self.getCantidad())
        else:
            self.valor_total = valor

    def setDescuentoUnitario(self, valor: int = None):
        """ Método que asigna el descuento unitario que entra por parámetro como el valor del atributo
        descuento_unitario

        Parameters
        ----------
        valor: int
            descuento unitario al que será asociado el servicio
        """

        if valor is None:
            try:
                equipo = eval("constants.%s" % self.equipo)
            except AttributeError:
                raise IncompatibleError
            df = equipo[equipo["Código"] == self.codigo]
            if len(df) == 0:
                raise Exception("Código inválido.")
            self.descuento_unitario = self.getValorUnitario() - int(df[self.interno].values[0])
        else:
            self.descuento_unitario = valor

    def setDescuentoTotal(self, valor: int = None):
        """ Método que asigna el descuento total que entra por parámetro como el valor del atributo descuento_total

        Parameters
        ----------
        valor: int
            descuento total al que será asociado el servicio
        """

        if valor is None:
            self.descuento_total = int(self.getDescuentoUnitario() * self.getCantidad())
        else:
            self.descuento_total = valor

    def setDescripcion(self, valor: str = None):
        """ Método que asigna la descripción que entra por parámetro como el valor del atributo descripcion

        Parameters
        ----------
        valor: str
            descripción a la que será asociada el servicio
        """

        if valor is None:
            equipo = eval("constants.%s" % self.equipo)
            df = equipo[equipo["Código"] == self.codigo]
            self.descripcion = df["Descripción"].values[0]
        else:
            self.descripcion = valor

    def setRestantes(self):
        """ Método que calcula el número de usos restantes del servicio """
        self.restantes = self.cantidad - self.getUsados()

    def setInterno(self, interno: str):
        """ Método que asigna el valor de interno que entra por parámetro como el valor del atributo interno

        Parameters
        ----------
        interno: str
            interno al que será asociado el servicio
        """

        self.interno = interno
        old_descuento = self.getDescuentoTotal()
        old_valor = self.getValorUnitario()
        old_total = self.getValorTotal()
        old_cantidad = self.getCantidad()

        self.setValorUnitario()
        self.setDescuentoUnitario()
        if old_total == int(old_valor * old_cantidad):
            self.setValorTotal()
            self.setDescuentoTotal()
        else:
            cantidad = old_total / self.getValorUnitario()
            cantidad = np.ceil(10 * cantidad) / 10
            self.setCantidad(cantidad)
            self.setValorTotal(old_total)
            self.setDescuentoTotal(old_descuento)

    def setDescuentoText(self, text: str):
        """ Método que asigna el texto asociado al descuento que entra por parámetro como el valor del atributo
        descuento_text

        Parameters
        ----------
        text: str
            texto de descuento al que será asociado el servicio
        """

        self.descuento_text = text

    def descontar(self, n: (int, float)):
        """ Método que descuenta en n la cantidad disponible del servicio, a la fecha actual

        Parameters
        ----------
        n: int, float
            cantidad a descontar del servicio
        """

        if n > 0:
            today = datetime.strftime(datetime.now(), "%Y/%m/%d")
            if today in self.usos.keys():
                self.usos[today] += n
            else:
                self.usos[today] = n
            self.restantes -= n

    def makeCotizacionTable(self) -> tuple:
        """ Método que genera la tabla que es usada por PDFCotizacion con la información del servicio

        Returns
        -------
        tuple:
            completo: la información del servicio, el código, la descripción, la cantidad, el valor unitario y el valor
            total
            descuento: el descuento asociado al servicio, la razón del descuento y el valor total del descuento
        """
        if self.getValorUnitario() == 0:
            completo = ["", self.getDescripcion(), "", ""]
            return completo, 5 * [""]
        else:
            completo = [self.getCodigo(), self.getDescripcion(), "%.1f" % self.getCantidad(),
                        "{:,}".format(self.getValorUnitario()), "{:,}".format(self.getValorTotal())]

            if (self.interno != "Industria") and (self.descuento_text == ""):
                descuento = ["", "Descuento por %s" % self.interno, "", "", "{:,}".format(-self.getDescuentoTotal())]
                return completo, descuento
            elif self.interno != "Industria":
                descuento = ["", "Descuento por %s" % self.descuento_text, "", "", "{:,}".format(-self.getDescuentoTotal())]
                return completo, descuento
            else:
                return completo,

    def makeReporteTable(self) -> list:
        """ Método que genera la tabla que es usada por PDFReporte con la información de usos del servicio

        Returns
        -------
        list:
            fecha de cada uso, código, descripción, cantidad, número de usos y número de restantes por fecha
        """

        fechas = sorted(list(self.usos.keys()))
        n = len(fechas)
        table = [None] * n
        cantidad = self.getCantidad()
        for i in range(n):
            fecha = fechas[i]
            usados = self.usos[fecha]
            restantes = cantidad - usados
            table[i] = [fecha, self.getCodigo(), self.getDescripcion(),
                        "%.1f" % self.getCantidad(), "%.1f" % usados, "%.1f" % restantes]
            cantidad = restantes
        return table

    def makeResumenTable(self) -> list:
        """ Método que genera la tabla que es usada por PDFReporte con la información resumida del servicio

        Returns
        -------
        list:
            descripción del servicio, cantidad, número de usos restantes, y su equivalente en dinero usado
        """
        return [self.getDescripcion(), "%.1f" % self.getCantidad(),
                "%.1f" % self.getRestantes(), "{:,}".format(self.getDineroUsado())]
