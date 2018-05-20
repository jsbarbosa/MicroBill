"""
    PDF
"""
CODIGO_GESTION = "UA-FM-CM-001"

TERMINOS_Y_CONDICIONES = ["Para usuarios internos favor realizar el traslado al Centro de beneficios/costos: 340121001",
                        "La Universidad no cobra IVA por ser una Institución de educación superior sin ánimo de lucro (Art. 92 de la ley 30 de 1992).",
			             ]

CONFIDENCIALIDAD = "Para asegurar la confidencialidad de cada individuo se utilizan códigos especiales de identificación. Es decir en lugar de utilizar el nombre y apellidos reales, o incluso el registro de la institución, se asignan otros códigos para su identificación. Por otro lado, el número de personas con acceso a dicha información es limitado. Generalmente se utilizan contraseñas personales para poder acceder a las bases de datos. Algunas de las bases de datos computarizadas pueden registrar quienes accedieron a la base y que información obtuvieron. Por último, los registros de papel se mantienen en un lugar cerrado y protegido."

CENTRO = "Centro de Microscopía"

DEPENDENCIAS = [CENTRO.upper(),
        "VICERRECTORÍA DE INVESTIGACIONES",
        "Carrera 1 N° 18A-10, Edificio B, Laboratorio 101.",
         "Tel: (57-1) 339 4949 Ext. 1444"]

"""
    Correo
"""

COTIZACION_SUBJECT_RECIBO = "Solicitud de cotización"
COTIZACION_MENSAJE_RECIBO = """Estimado usuario,

A continuación hacemos envío de la cotización correspondiente a los servicios solicitados.

Tenga en cuenta las siguientes recomendaciones:
- Para el día de la sesión es necesario contar con el comprobante de pago.

Quedamos atentos a sus comentarios,
"""

COTIZACION_SUBJECT_TRANSFERENCIA = COTIZACION_SUBJECT_RECIBO
COTIZACION_MENSAJE_TRANSFERENCIA = "COTIZACION_MENSAJE_TRANSFERENCIA"

COTIZACION_SUBJECT_FACTURA = COTIZACION_SUBJECT_RECIBO
COTIZACION_MENSAJE_FACTURA = COTIZACION_MENSAJE_RECIBO

REPORTE_SUBJECT = "Estado actual"
REPORTE_MENSAJE = "REPORTE_MENSAJE"

REQUEST_SUBJECT = "Solicitud de información para cotización"
REQUEST_MENSAJE = "REQUEST_MENSAJE"

SERVER = "smtp-mail.outlook.com"
PORT = 587
