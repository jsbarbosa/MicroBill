ADMINS = ["Humberto Ibarra", "Monica Lopez", "Juan Camilo Orozco", "Laura Cruz", "Juan Barbosa"]

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

Cordial saludo,
"""





COTIZACION_SUBJECT_TRANSFERENCIA = COTIZACION_SUBJECT_RECIBO
COTIZACION_MENSAJE_TRANSFERENCIA = """Estimado usuario,

A continuación hacemos envío de la cotización correspondiente a los servicios solicitados.

Tenga en cuenta las siguientes recomendaciones:
- Para el día de la sesión es necesario contar con la aprobación del director de departamento y con la evidencia de translado presupuestal, con referencia a la cotización enviada.

Quedamos atentos a sus comentarios,

Cordial saludo,
"""




COTIZACION_SUBJECT_FACTURA = COTIZACION_SUBJECT_RECIBO
COTIZACION_MENSAJE_FACTURA = COTIZACION_MENSAJE_RECIBO

REPORTE_SUBJECT = "Estado actual"
REPORTE_MENSAJE = """Estimado usuario,

A continuación hacemos envío del estado actual de la cotización, en este documento usted podrá encontrar:
- Los servicios que fueron cotizados.
- Las fechas en que estos servicios fueron usados.
- Un resumen del número de usos, y la cantidad restante que aun puede utilizar.

Quedamos atentos a sus comentarios,

Cordial saludo,
"""





REQUEST_SUBJECT = "Solicitud de información"
REQUEST_MENSAJE = """Estimado usuario,

Para cotizar los ensayos solicitados, es necesario que nos proporcione la siguiente información, de tal manera que podamos asesorarlo de la mejor manera posible.

1. Tamaño y cantidad de muestras:
2. Objetivo del ensayo a realizar:
3. Nombre de la persona de contacto:
4. Institución:
5. Documento (C.C o NIT) de la persona o empresa que realiza el pago:
6. Teléfono:
7. Dirección:
8. Correo electrónico:

Si usted es usuario Uniandes y hace parte de un proyecto dentro de la Universidad le solicitamos nos indique:

- Facultad/departamento al que pertenece:
- Nombre del profesor responsable:
- Nombre/código del proyecto:

Quedamos atentos a sus comentarios,

Cordial saludo,
"""



GESTOR_RECIBO_CORREO = "js.barbosa10@uniandes.edu.co"
GESTOR_FACTURA_CORREO = "js.barbosa10@uniandes.edu.co"

GESTOR_RECIBO_SUBJECT = "Generación recibo de pago"
GESTOR_RECIBO_MENSAJE = """Estimado Carlos Julio,

Por este medio solicitamos la generación del recibo de pago correspondiente a la cotización adjunta.

Cordial saludo,
"""

GESTOR_FACTURA_SUBJECT = "Generación factura"
GESTOR_FACTURA_MENSAJE = """Estimada Mercedes,

Por este medio solicitamos la generación de la factura correspondiente a la cotización y orden/aceptación de servicio adjuntas

Cordial saludo,
"""


SERVER = "smtp-mail.outlook.com"
PORT = 587
