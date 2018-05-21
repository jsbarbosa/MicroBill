"""
    PDF
"""
CODIGO_GESTION = "UA-FM-CM-001"

TERMINOS_Y_CONDICIONES = ["Para usuarios internos favor realizar el traslado al Centro de beneficios/costos: 340121001",
                        "La Universidad no cobra IVA por ser una Institución de educación superior sin ánimo de lucro (Art. 92 de la ley 30 de 1992).",
			             ]

CONFIDENCIALIDAD = "El contratista se obliga a respetar el carácter confidencial de este documento y de la información, condiciones y documentos relacionados con el mismo, de conformidad con las normas constitucionales y legales aplicables. En consecuencia, el contratista se compromete a no publicar, difundir, comentar o analizar frente a terceros, copiar, reproducir o hacer uso diferente al acordado, de la información que el centro de microscopia le entregue en este documento, ya sea de forma impresa, electrónica, verbal o de cualquier otra manera, o de aquella que el contratista llegue a conocer, teniendo en cuenta que dicha información tiene como finalidad exclusiva, permitir y facilitar la debida prestación de los servicios solicitados."
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
