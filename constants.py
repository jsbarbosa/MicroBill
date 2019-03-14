import os
import config
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
