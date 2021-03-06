import os
import pandas as pd
from pandas import ExcelWriter

DEBUG: bool = True

__all__ = ['OLD_DIR', 'PDF_DIR', 'REGISTERS_DIR', 'CLIENTES_FILE', 'REGISTRO_FILE', 'PRECIOS_FILE',
           'BASE_DIR', 'EQUIPOS_KEYS', 'REGISTRO_KEYS', 'CLIENTES_KEYS', 'DOCUMENTOS_FINALES', 'EQUIPOS',
           'PRICES_DIVISION', 'REPORTE_INTERNO', 'DEFAULT_CONFIG']

__version__ = '1.0.1'

OLD_DIR: str = "Old"  #: carpeta donde se guardan las cotizaciones realizadas
PDF_DIR: str = "PDF"  #: carpeta donde se guardan los PDFs asociados a cotizaciones realizadas

#: carpeta donde se encuentran los archivos de registro de Microbill (Excel de los que se alimenta)
REGISTERS_DIR: str = "Registers"

CLIENTES_FILE: str = "Clientes.xlsx" #: archivo en el que se guarda el registro de clientes
REGISTRO_FILE: str = "Registro.xlsx" #: archivo en el que se guarda el resumen de la información de las cotizaciones
PRECIOS_FILE: str = "Precios.xlsx" #: archivo que contiene los precios de los servicios a usar
INDEPENDIENTES_FILE: str = "Independientes.xlsx"
PRECIOS_DAEMON_FILE: str = "Precios (daemon).xlsx"

#: directorio base desde el cual se está ejecutando el intérprete de Python
BASE_DIR: str = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if os.path.exists(os.path.join(BASE_DIR, 'microbill')):
    BASE_DIR = os.path.join(BASE_DIR, 'microbill')

OLD_DIR = os.path.join(BASE_DIR, OLD_DIR)
if not os.path.exists(OLD_DIR):
    os.mkdir(OLD_DIR)

PDF_DIR = os.path.join(BASE_DIR, PDF_DIR)
if not os.path.exists(PDF_DIR):
    os.mkdir(PDF_DIR)

REGISTERS_DIR = os.path.join(BASE_DIR, REGISTERS_DIR)
if not os.path.exists(REGISTERS_DIR):
    os.mkdir(REGISTERS_DIR)

CLIENTES_FILE = os.path.join(REGISTERS_DIR, CLIENTES_FILE)
REGISTRO_FILE = os.path.join(REGISTERS_DIR, REGISTRO_FILE)
PRECIOS_FILE = os.path.join(REGISTERS_DIR, PRECIOS_FILE)

PRECIOS_DAEMON_FILE = os.path.join(REGISTERS_DIR, PRECIOS_DAEMON_FILE)
INDEPENDIENTES_FILE = os.path.join(REGISTERS_DIR, INDEPENDIENTES_FILE)

#: columnas que deben tener el archivo de precios
EQUIPOS_KEYS: list = ['Código', 'Descripción', 'Interno', 'Académico', 'Industria', 'Independiente']

#: columnas que debe tener el archivo de registro
REGISTRO_KEYS: list = ['Cotización', 'Fecha', 'Nombre', 'Correo', 'Teléfono', 'Institución', 'Interno',
                       'Responsable', 'Muestra', 'Equipo', 'Elaboró', 'Modificó', 'Estado', 'Pago', 'Referencia',
                       'Aplicó', 'Tipo de Pago', 'Valor']

#: columnas que debe tener el archivo de clientes
CLIENTES_KEYS: list = ['Nombre', 'Correo', 'Teléfono', 'Institución', 'Documento',
                       'Dirección', 'Ciudad', 'Interno', 'Responsable', 'Proyecto', 'Código', 'Tipo de Pago']

#: columnas que debe tener el archivo de precios para microbill daemon
PRECIOS_DAEMON_KEYS: list = ['Equipo', 'Código', 'Item',
                             'Interno', 'Académico', 'Industria',
                             'Independiente']

#: columnas que debe tener el archivo de clientes independientes
INDEPENDIENTES_KEYS: list = ['Nombre', 'Correo']

#: posibles formas de pago de una cotizacion
DOCUMENTOS_FINALES: list = ["Transferencia interna", "Factura", "Recibo"]

if not os.path.exists(PRECIOS_FILE):
    df = pd.DataFrame(columns=EQUIPOS_KEYS)
    with ExcelWriter(PRECIOS_FILE) as writer:
        df.to_excel(writer, 'Equipo_prueba_01', index=False)

if not os.path.exists(CLIENTES_FILE):
    pd.DataFrame(columns=CLIENTES_KEYS).to_excel(CLIENTES_FILE, index=False)

if not os.path.exists(REGISTRO_FILE):
    pd.DataFrame(columns=REGISTRO_KEYS).to_excel(REGISTRO_FILE, index=False)

if not os.path.exists(PRECIOS_DAEMON_FILE):
    pd.DataFrame(columns=PRECIOS_DAEMON_KEYS).to_excel(PRECIOS_DAEMON_FILE, index=False)

if not os.path.exists(INDEPENDIENTES_FILE):
    pd.DataFrame(columns=INDEPENDIENTES_KEYS).to_excel(INDEPENDIENTES_FILE, index=False)

df = pd.read_excel(PRECIOS_FILE, sheet_name=None)
EQUIPOS: list = list(df.keys()) #: almacena el nombre de las hojas del archivo de precios

for item in EQUIPOS:
    data = df[item].astype(str)
    exec("%s = data" % item)

INDEPENDIENTES_DF = pd.read_excel(INDEPENDIENTES_FILE)

DAEMON_DF = pd.read_excel(PRECIOS_DAEMON_FILE, sheet_name=None)
DAEMON_SUBCLASSES = list(df.keys())

#: categorias que dividen los precios de los servicios
PRICES_DIVISION: list = ['Interno', 'Académico', 'Industria', 'Independiente']

REPORTE_INTERNO: str = "ReporteInterno.xlsx" #: nombre del archivo en donde se almacenan los reportes internos

#: configuracion por defecto
DEFAULT_CONFIG: str = '''﻿ADMINS = ["Humberto Ibarra", "Monica Lopez", "Juan Camilo Orozco", "Laura Sotelo"]

CORREOS_CONVENIOS = ["uniandes"]

"""
    PDF
"""
CODIGO_GESTION = "UA-FM-CM-001"
# CODIGO_PEP = "P%s.340121.0%s"
CODIGO_PEP = "340121001"

ANCHO_LOGO = 8.58
ALTO_LOGO = 1.5

LOGO_PATH = "logo.png"
SPLASH_LOGO_PATH = "logo.png"

TERMINOS_Y_CONDICIONES = ["Para usuarios internos favor realizar el traslado al PEP indicado en la parte superior en rojo", "La Universidad no cobra IVA por ser una Institución de educación superior sin ánimo de lucro (Art. 92 de la ley 30 de 1992).",]

CONFIDENCIALIDAD = "El contratista se obliga a respetar el carácter confidencial de este documento y de la información, condiciones y documentos relacionados con el mismo, de conformidad con las normas constitucionales y legales aplicables. En consecuencia, el contratista se compromete a no publicar, difundir, comentar o analizar frente a terceros, copiar, reproducir o hacer uso diferente al acordado, de la información que el Centro de Microscopía le entregue en este documento, ya sea de forma impresa, electrónica, verbal o de cualquier otra manera, o de aquella que el contratista llegue a conocer, teniendo en cuenta que dicha información tiene como finalidad exclusiva, permitir y facilitar la debida prestación de los servicios solicitados."
CENTRO = "Centro de Microscopía"
PREFIJO = "U"

DEPENDENCIAS = [CENTRO.upper(),
        "VICERRECTORÍA DE INVESTIGACIONES",
        "Carrera 1 N° 18A-10, Edificio B, Laboratorio 101.",
         "Tel: (57-1) 339 4949 Ext. 1444"]

"""
    Correo
"""

SALUDO = """Estimado usuario,

"""

COTIZACION_SUBJECT_RECIBO = "Solicitud de cotización"
COTIZACION_MENSAJE_RECIBO = """A continuación hacemos envío de la(s) cotización(es) correspondiente(s) a los servicios solicitados.

Tenga en cuenta las siguientes recomendaciones:
- Para el día de la sesión es necesario contar con el comprobante de pago.

Quedamos atentos a sus comentarios,

Cordial saludo,

CENTRO DE MICROSCOPÍA
VICERRECTORÍA DE INVESTIGACIONES
Carrera 1 N° 18A-10, Edificio B, Laboratorio 101.
Tel: (57-1) 339 4949 Ext. 1444
"""

COTIZACION_MENSAJE_TRANSFERENCIA = """A continuación hacemos envío de la(s) cotización(es) correspondiente(s) a los servicios solicitados.

Tenga en cuenta las siguientes recomendaciones:
- Para el día de la sesión es necesario contar con la aprobación del director de departamento y con la evidencia de translado presupuestal, con referencia a la cotización enviada.

Quedamos atentos a sus comentarios,

Cordial saludo,

CENTRO DE MICROSCOPÍA
VICERRECTORÍA DE INVESTIGACIONES
Carrera 1 N° 18A-10, Edificio B, Laboratorio 101.
Tel: (57-1) 339 4949 Ext. 1444
"""

COTIZACION_MENSAJE_FACTURA = """A continuación hacemos envío de la(s) cotización(es) correspondiente(s) a los servicios solicitados.

Tenga en cuenta las siguientes recomendaciones:
- Le solicitamos enviar la orden de servicios, para poder generar la factura.

Quedamos atentos a sus comentarios,

Cordial saludo,

CENTRO DE MICROSCOPÍA
VICERRECTORÍA DE INVESTIGACIONES
Carrera 1 N° 18A-10, Edificio B, Laboratorio 101.
Tel: (57-1) 339 4949 Ext. 1444
"""


REPORTE_SUBJECT = "Estado actual"
REPORTE_MENSAJE = """A continuación hacemos envío del estado actual de la cotización, en este documento usted podrá encontrar:
- Los servicios que fueron cotizados.
- Las fechas en que estos servicios fueron usados.
- Un resumen del número de usos, y la cantidad restante que aun puede utilizar.

Quedamos atentos a sus comentarios,

Cordial saludo,

CENTRO DE MICROSCOPÍA
VICERRECTORÍA DE INVESTIGACIONES
Carrera 1 N° 18A-10, Edificio B, Laboratorio 101.
Tel: (57-1) 339 4949 Ext. 1444
"""





REQUEST_SUBJECT = "Solicitud de información"
REQUEST_MENSAJE = """Para cotizar los ensayos solicitados, es necesario que nos proporcione la siguiente información, de tal manera que podamos asesorarlo de la mejor manera posible.

1. Tamaño y cantidad de muestras:
2. Objetivo del ensayo a realizar:
3. Tipo de muestra:
4. Nombre de la persona de contacto:
5. Institución:
6. Documento (C.C o NIT) de la persona o empresa que realiza el pago:
7. Teléfono:
8. Dirección:
9. Correo electrónico:

Si usted es usuario Uniandes y hace parte de un proyecto dentro de la Universidad le solicitamos nos indique:

- Facultad/departamento al que pertenece:
- Nombre del profesor responsable:
- Nombre/código del proyecto:

Quedamos atentos a sus comentarios,

Cordial saludo,

CENTRO DE MICROSCOPÍA
VICERRECTORÍA DE INVESTIGACIONES
Carrera 1 N° 18A-10, Edificio B, Laboratorio 101.
Tel: (57-1) 339 4949 Ext. 1444
"""



GESTOR_RECIBO_CORREO = "cjvergara@uniandes.edu.co"
GESTOR_FACTURA_CORREO = "carrodri@uniandes.edu.co"

GESTOR_RECIBO_SUBJECT = "Generación recibo de pago"
GESTOR_RECIBO_MENSAJE = """Estimado Carlos Julio,

Por este medio solicitamos la generación del recibo de pago correspondiente a la cotización adjunta.

Cordial saludo,

CENTRO DE MICROSCOPÍA
VICERRECTORÍA DE INVESTIGACIONES
Carrera 1 N° 18A-10, Edificio B, Laboratorio 101.
Tel: (57-1) 339 4949 Ext. 1444
"""

GESTOR_FACTURA_SUBJECT = "Generación factura"
GESTOR_FACTURA_MENSAJE = """Estimada Mercedes,

Por este medio solicitamos la generación de la factura correspondiente a la cotización y orden/aceptación de servicio adjuntas

Cordial saludo,

CENTRO DE MICROSCOPÍA
VICERRECTORÍA DE INVESTIGACIONES
Carrera 1 N° 18A-10, Edificio B, Laboratorio 101.
Tel: (57-1) 339 4949 Ext. 1444
"""

SEND_SERVER = "smtp-mail.outlook.com"
SEND_PORT = 587

READ_SERVER = "imap-mail.outlook.com"
READ_PORT = 993

READ_EVERY_S = 10

READ_EMAIL_FOLDER = "INBOX/AGENDO"
READ_REQUEST_DETAILS = "Request details"
READ_FIELDS = "Fields"
READ_LAST = "Please go to the link below to view this request.  "

READ_FIELDS_DICT = {'nombre': "A new request was submitted by ",
                "id":"Request id ",
                "documento": "Nit/CC ",
                "responsable": "Person in charge ",
                "direccion": "Address ",
                "ciudad": "City ",
                "telefono": "Telephone ",
                "institucion": "Institution ",
                "pago": "Paying method ",
                "tipo": "Sample type ",
                "proyecto": "Project name ",
                "muestra": "Sample type ",
                "codigo": "Project code ",
		"externo": "External funds ",
                }

READ_NUMERIC_FIELDS = ["Request id", "Id", "Telephone"]
READ_SEARCH_FOR = ['FROM "lists@cirklo.org" SUBJECT "[Agendo] New request (2-Quote request"']


FROM = "cmicroscopia@uniandes.edu.co"
PASSWORD = "fmPClsKmZcKpwppowqlhZQ=="
'''

EXIT_CODE_REBOOT: int = -14 #: código usado para señalizar que el programa busca reiniciar automaticamente
