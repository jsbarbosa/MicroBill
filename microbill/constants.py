import os
import pandas as pd

DEBUG = False

OLD_DIR = "Old"
PDF_DIR = "PDF"
REGISTERS_DIR = "Registers"

CLIENTES_FILE = "Clientes.xlsx"
REGISTRO_FILE = "Registro.xlsx"
PRECIOS_FILE = "Precios.xlsx"
INDEPENDIENTES_FILE = "Independientes.xlsx"
PRECIOS_DAEMON_FILE = "Precios (daemon).xlsx"

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

if os.path.exists(os.path.join(BASE_DIR, 'microbill')):
    BASE_DIR = os.path.join(BASE_DIR, 'microbill')

OLD_DIR = os.path.join(BASE_DIR, OLD_DIR)
PDF_DIR = os.path.join(BASE_DIR, PDF_DIR)
REGISTERS_DIR = os.path.join(BASE_DIR, REGISTERS_DIR)

CLIENTES_FILE = os.path.join(REGISTERS_DIR, CLIENTES_FILE)
REGISTRO_FILE = os.path.join(REGISTERS_DIR, REGISTRO_FILE)
PRECIOS_FILE = os.path.join(REGISTERS_DIR, PRECIOS_FILE)

PRECIOS_DAEMON_FILE = os.path.join(REGISTERS_DIR, PRECIOS_DAEMON_FILE)

INDEPENDIENTES_FILE = os.path.join(REGISTERS_DIR, INDEPENDIENTES_FILE)

EQUIPOS_KEYS = ['Código', 'Descripción', 'Interno', 'Académico', 'Industria', 'Independiente']
# DAEMON_KEYS = []
REGISTRO_KEYS = ['Cotización', 'Fecha', 'Nombre', 'Correo', 'Teléfono', 'Institución', 'Interno',
                  'Responsable', 'Muestra', 'Equipo', 'Elaboró', 'Modificó', 'Estado', 'Pago', 'Referencia', 'Aplicó', 'Tipo de Pago', 'Valor']
CLIENTES_KEYS = ['Nombre', 'Correo', 'Teléfono', 'Institución', 'Documento',
                 'Dirección', 'Ciudad', 'Interno', 'Responsable', 'Proyecto', 'Código', 'Tipo de Pago']

DOCUMENTOS_FINALES = ["Transferencia interna", "Factura", "Recibo"]

df = pd.read_excel(PRECIOS_FILE, sheet_name = None)
EQUIPOS = list(df.keys())

for item in EQUIPOS:
    data = df[item].astype(str)
    exec("%s = data"%item)

INDEPENDIENTES_DF = pd.read_excel(INDEPENDIENTES_FILE)

DAEMON_DF = pd.read_excel(PRECIOS_DAEMON_FILE, sheet_name = None)
DAEMON_SUBCLASSES = list(df.keys())

PRICES_DIVISION = ['Interno', 'Académico', 'Industria', 'Independiente']

REPORTE_INTERNO = "ReporteInterno.xlsx"

DEFAULT_CONFIG = '''﻿ADMINS = ["Humberto Ibarra", "Monica Lopez", "Juan Camilo Orozco", "Laura Sotelo"]


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

FROM = "cmicroscopia@uniandes.edu.co"
PASSWORD = "fmPClsKmZcKpwppowqlhZQ=="
'''

EXIT_CODE_REBOOT = -14
