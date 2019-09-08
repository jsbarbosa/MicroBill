import os
from datetime import datetime, timedelta
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from .constants import BASE_DIR
from .utils import export
# from . import constants
# from . import config
# from .config import *
from . import constants, config

if os.path.isdir(constants.PDF_DIR):
    pass
else:
    os.makedirs(constants.PDF_DIR)


@export
class PDFBase(object):
    """
    Clase base para PDFCotización y PDFReporte. Inicializa el SimpleDocTemplate de ReportLab sobre el cual se construye
    el PDF.
    """

    def __init__(self, cotizacion=None, is_reporte: bool = False):
        self.cotizacion = cotizacion
        if is_reporte:
            self.file_name = self.cotizacion.getNumero() + "_Reporte.pdf"
        else:
            self.file_name = self.cotizacion.getNumero() + ".pdf"
        self.file_name = os.path.join(constants.PDF_DIR, self.file_name)

        self.cotizacion.setPath(self.file_name)

        self.doc = SimpleDocTemplate(self.file_name, pagesize=letter,
                                     rightMargin=cm, leftMargin=cm,
                                     topMargin=cm, bottomMargin=cm)

        self.styles = getSampleStyleSheet()
        self.styles.add(ParagraphStyle(name='Center', alignment=TA_CENTER))
        self.styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY))

        border = ParagraphStyle(name='Border', alignment=TA_JUSTIFY)
        border.borderWidth = 0.5
        border.borderColor = colors.black
        border.borderRadius = 2
        border.borderPadding = 3
        self.styles.add(border)

        self.story = []

    def makeInfo(self):
        """ Método que construye la parte superior del PDF en donde se muestra el nombre de la dependencia, junto con
        Universidad de los Andes y la información del usuario.
        """

        usuario = self.cotizacion.getUsuario()
        ptext = '<font size = 12><b>%s</b></font>' % "%s - UNIVERSIDAD DE LOS ANDES" % config.CENTRO.upper()
        self.story.append(Paragraph(ptext, self.styles["Center"]))

        self.story.append(Spacer(1, 12))

        h = 3
        c1 = 2.2*cm
        c2 = 9*cm
        c3 = 2.5*cm
        c4 = self.doc.width - (c1 + c2 + c3)
        data = [[Paragraph("<b>Nombre:</b>", self.styles["Normal"]),
                 Paragraph(usuario.getNombre(), self.styles["Border"]),
                 Paragraph("<b>Institución:</b>", self.styles["Normal"]),
                 Paragraph(usuario.getInstitucion(), self.styles["Border"])
                 ]]
        t = Table(data, [c1, c2, c3, c4], hAlign='LEFT')
        self.story.append(Spacer(1, h))
        self.story.append(t)

        data = [[Paragraph("<b>Nit/C.C.:</b>", self.styles["Normal"]),
                 Paragraph(usuario.getDocumento(), self.styles["Border"]),
                 Paragraph("<b>Teléfono:</b>", self.styles["Normal"]),
                 Paragraph(usuario.getTelefono(), self.styles["Border"])
                 ]]
        t = Table(data, [c1, c2, c3, c4], hAlign='LEFT')
        self.story.append(Spacer(1, h))
        self.story.append(t)

        data = [[Paragraph("<b>Dirección:</b>", self.styles["Normal"]),
                 Paragraph(usuario.getDireccion(), self.styles["Border"]),
                 Paragraph("<b>Ciudad:</b>", self.styles["Normal"]),
                 Paragraph(usuario.getCiudad(), self.styles["Border"])
                 ]]
        t = Table(data, [c1, c2, c3, c4], hAlign='LEFT')
        self.story.append(Spacer(1, h))
        self.story.append(t)

        data = [[Paragraph("<b>Correo:</b>", self.styles["Normal"]),
                 Paragraph(usuario.getCorreo(), self.styles["Border"])
                 ]]

        t = Table(data, [c1, c2, c3, c4], hAlign='LEFT')
        self.story.append(Spacer(1, h))
        self.story.append(t)

        data = [[Paragraph("<b>Muestra:</b>", self.styles["Normal"]),
                 Paragraph(self.cotizacion.getMuestra(), self.styles["Border"])
                 ]]
        t = Table(data, [c1, c2 + c3 + c4], hAlign='LEFT')
        self.story.append(Spacer(1, h))
        self.story.append(t)

    def makeEnd(self):
        """ Método que genera el pie de página con el listado de dependencias de la universida y el centro de servicios
        """

        for i in range(len(config.DEPENDENCIAS)):
            text = config.DEPENDENCIAS[i]
            if i < 2:
                ptext = '<font size = 10> <b> %s </b></font>' % text
            else:
                ptext = '<font size = 10>%s</font>' % text
            self.story.append(Paragraph(ptext, self.styles["Center"]))

    def drawPage(self, canvas, doc):
        """ Método necesario por reportlab al momento de generar el PageTemplate

        Parameters
        ----------
        canvas: reportlab.Canvas
        doc: reportlab.SimpleDocTemplate
        """

        canvas.setTitle(config.CENTRO)
        canvas.setSubject("Cotización")
        canvas.setAuthor("Juan Barbosa")
        canvas.setCreator("MicroBill")
        self.styles = getSampleStyleSheet()

        p = Paragraph("<font size = 8>Creado usando MicroBill - Juan Barbosa (github.com/jsbarbosa)</font>",
                      self.styles["Normal"])
        w, h = p.wrap(doc.width, doc.bottomMargin)
        p.drawOn(canvas, doc.leftMargin, h)

    def build(self, template: PageTemplate):
        """ Renderiza el documento PDF

        Parameters
        ----------
        template: el PageTemplate usado para renderizar el documento
        """

        self.doc.addPageTemplates([template])
        self.doc.build(self.story)


@export
class PDFCotizacion(PDFBase):
    """ Clase que representa el PDF de las cotizaciones. Hereda de la clase PDFBase """

    def __init__(self, cotizacion):
        super(PDFCotizacion, self).__init__(cotizacion)
        self.frame1 = None
        self.frame2 = None
        self.last_frame = None

    def makeFrames(self):
        """ Método que genera los marcos en los cuales se organiza la información al interior de un SimpleDocTemplate.
        Los dos primeros frames contienen la información personal del usuario (una columna por frame).
        El último marco contiene la tabla, el recuadro de observaciones y condiciones
        """

        # Two Columns
        h = 4*cm
        self.frame1 = Frame(self.doc.leftMargin, self.doc.bottomMargin + self.doc.height - h,
                            self.doc.width / 2 - 6, h, id='col1')
        self.frame2 = Frame(self.doc.leftMargin + self.doc.width / 2 + 2*cm,
                            self.doc.bottomMargin + self.doc.height - h, self.doc.width / 2 - 6, h, id='col2')

        last_h = self.doc.bottomMargin + self.doc.height - h
        self.last_frame = Frame(self.doc.leftMargin, self.doc.bottomMargin, self.doc.width, last_h - cm)

    def makeTop(self):
        """ Método que genera el encabezado del PDF, incluye el LOGO_PATH, el CODIGO_PEP, el número de la cotización,
        y la fecha de validez de la misma
        """

        logo = Image(os.path.join(BASE_DIR, config.LOGO_PATH), config.ANCHO_LOGO * cm,
                     config.ALTO_LOGO * cm, hAlign="CENTER")

        self.story.append(logo)
        self.story.append(Spacer(0, 12))
        data = [["CÓDIGO S.GESTIÓN", config.CODIGO_GESTION]]

        style = [('BOX', (0, 0), (1, 0), 0.25, colors.black),
                 ('FONTSIZE', (0, 0), (-1, -1), 8),
                 ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey)
                 ]

        t = Table(data, 90, 20)
        t.setStyle(TableStyle(style))
        self.story.append(t)

        if self.cotizacion.internoTreatment() and constants.DOCUMENTOS_FINALES[0]:
            # cod = self.cotizacion.getNumero()[:2]
            # year = str(datetime.now().year)[-2:]
            # data = [[CODIGO_PEP % (year, cod)]] # not ready
            data = [[config.CODIGO_PEP]]
            style = [('BOX', (0, 0), (0, 0), 3, colors.black),
                     ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                     ('VALIGN', (0, 0), (0, 0), 'TOP'),
                     ('FONTSIZE', (0, 0), (0, 0), 13),
                     ('TEXTCOLOR', (0, 0), (0, 0), colors.red)
                     ]

            t = Table(data, 160, 20)
            t.setStyle(TableStyle(style))
            self.story.append(Spacer(1, 5))
            self.story.append(t)

        self.story.append(FrameBreak())

        ptext = '<font size = 12><b>%s</b></font>' % "COTIZACIÓN DE SERVICIOS"
        self.story.append(Paragraph(ptext, self.styles["Normal"]))

        numero = self.cotizacion.getNumero()
        today = datetime.now()
        fecha = datetime.strftime(today, "%d/%m/%Y")
        until = datetime.strftime(today + timedelta(days=45), "%d/%m/%Y")

        data = [["COTIZACIÓN N°", "FECHA"], [numero, fecha], [numero, "VALIDA HASTA"], [numero, until]]

        t = Table(data, 80, 15, hAlign='LEFT')
        t.setStyle(TableStyle([('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                               ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                               ('FONTSIZE', (0, 0), (-1, -1), 8),
                               ('FONTSIZE', (0, 1), (0, 1), 12),
                               ('ALIGN', (0, 1), (0, 1), "CENTER"),
                               ('VALIGN', (0, 1), (0, 1), "MIDDLE"),
                               ('SPAN', (0, 1), (0, -1)),
                               ('BACKGROUND', (0, 0), (1, 0), colors.lightgrey),
                               ('BACKGROUND', (1, 2), (1, 2), colors.lightgrey)
                               ]))

        self.story.append(Spacer(1, 12))
        self.story.append(t)

    def makeTable(self):
        """ Función que genera la tabla en donde se enlistan los servicios, sus códigos, cantidad, precio unitatio,
        precio total, el subtotal, el descuento total y el total a pagar
        """

        table = self.cotizacion.makeCotizacionTable()
        subtotal = self.cotizacion.getSubtotal()
        descuentos = -self.cotizacion.getDescuentos()
        total = self.cotizacion.getTotal()
        table.insert(0, ["COD", "SERVICIO", "CANTIDAD", "PRECIO UNIDAD", "PRECIO TOTAL"])

        if self.cotizacion.getInterno() == "Industria":
            o_row = -1
            table.append(["", "", "", "TOTAL", "{:,}".format(total)])
        else:
            o_row = -3
            table.append(["", "", "", "SUBTOTAL", "{:,}".format(subtotal)])
            table.append(["", "", "", "D.TOTAL", "{:,}".format(descuentos)])
            table.append(["", "", "", "TOTAL", "{:,}".format(total)])

        self.story.append(Spacer(1, 24))

        t = Table(table, [40, 260, 70, 90, 90], 15, hAlign='CENTER')

        t.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                               ('ALIGN', (0, 0), (-1, 0), "CENTER"),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                               ('ALIGN', (-3, 0), (-1, -1), "RIGHT"),
                               ('SPAN', (0, -1), (-3, -1)),
                               ('BOX', (-2, o_row), (-1, -1), 0.5, colors.black),
                               ('BOX', (-2, -1), (-1, -1), 0.5, colors.red),
                               ('BACKGROUND', (-2, o_row), (-2, -1), colors.lightgrey),
                               ]))
        self.story.append(t)
        self.story.append(Spacer(1, 24))

    def makeObservaciones(self):
        """ Método que genera el contenido del recuadro de observaciones """

        text = "OBSERVACIONES"
        ptext = '<font size = 10> <b> %s </b></font>' % text

        self.story.append(Paragraph(ptext, self.styles["Center"]))
        self.story.append(Spacer(1, 6))

        usuario = self.cotizacion.getUsuario()

        text = [("Responsable", usuario.getResponsable()),
                ("Proyecto", usuario.getProyecto()),
                ("Código", usuario.getCodigo())]

        ptext = '<font size = 9>%s</font><br/>' % self.cotizacion.getObservacionPDF().replace("\n", "<br/>")
        if self.cotizacion.internoTreatment():
            for item in text:
                ptext += '<font size = 9> <b>%s:</b> <u>%s</u></font> <br/>' % item

        self.story.append(Paragraph(ptext, self.styles["Border"]))

        self.story.append(Spacer(1, 24))

    def makeTerminos(self):
        """ Método que genera el contenido del marco de términos y condiciones """

        text = "TÉRMINOS Y CONDICIONES"
        ptext = '<font size = 10> <b> %s </b></font>' % text
        self.story.append(Paragraph(ptext, self.styles["Center"]))
        self.story.append(Spacer(1, 6))

        ptext = ""

        starts = 1
        if self.cotizacion.internoTreatment():
            starts = 0

        for i in range(starts, len(config.TERMINOS_Y_CONDICIONES)):
            text = config.TERMINOS_Y_CONDICIONES[i]
            ptext += '<font size = 8>%d. %s</font><br/>' % (i + abs(starts - 1), text)
            # self.story.append(Paragraph(ptext, self.styles["Justify"]))

        ptext += '<font size = 8>%d.</font> <font size = 6>%s</font>' \
                 % (i + 1 + abs(starts - 1), config.CONFIDENCIALIDAD)
        self.story.append(Paragraph(ptext, self.styles["Border"]))
        self.story.append(Spacer(1, 24))

    def doAll(self):
        """ Método que renderiza una cotización """

        self.makeFrames()
        self.makeTop()
        self.makeInfo()
        self.makeTable()
        self.makeObservaciones()
        self.makeTerminos()
        self.makeEnd()
        temp = PageTemplate(frames=[self.frame1, self.frame2, self.last_frame], onPage=self.drawPage)
        self.build(temp)


@export
class PDFReporte(PDFBase):
    """ Clase que representa el PDF de los reportes. Hereda de la clase PDFBase """

    def __init__(self, cotizacion):
        super(PDFReporte, self).__init__(cotizacion, True)
        self.table = None
        self.resumen = None
        self.frame1 = None
        self.frame2 = None
        self.last_frame = None

    def makeFrames(self):
        """ Método que genera los marcos en los cuales se organiza la información al interior de un SimpleDocTemplate.
        Los dos primeros frames contienen la información personal del usuario (una columna por frame).
        El último marco contiene la tabla, el recuadro de observaciones y condiciones
        """

        self.table = self.cotizacion.makeReporteTable()
        self.resumen = self.cotizacion.makeResumenTable()

        self.resumen.insert(0, ["DESCRIPCIÓN", "COTIZADOS", "RESTANTES"])
        self.table.insert(0, ["FECHA", "COD", "DESCRIPCIÓN", "COTIZADOS", "USADOS", "RESTANTES"])

        # Two Columns
        h = 3 * cm
        self.frame1 = Frame(self.doc.leftMargin, self.doc.bottomMargin + self.doc.height - h,
                            self.doc.width / 2 - 6, h, id='col1')
        self.frame2 = Frame(self.doc.leftMargin + self.doc.width / 2 + 2*cm,
                            self.doc.bottomMargin + self.doc.height - h, self.doc.width / 2 - 6, h, id='col2')

        last_h = self.doc.bottomMargin + self.doc.height - h

        self.last_frame = Frame(self.doc.leftMargin, self.doc.bottomMargin, self.doc.width, last_h - cm)

    def makeTop(self):
        """ Método que genera el encabezado del PDF, incluye el LOGO_PATH, el CODIGO_PEP, el número de la cotización,
        y la fecha de validez de la misma
        """

        height = 1.5
        logo = Image("logo.png", height * 5.72 * cm, height * cm, hAlign="CENTER")

        self.story.append(logo)
        self.story.append(FrameBreak())

        ptext = '<font size = 12><b>%s</b></font>' % "REPORTE DE SERVICIOS"
        self.story.append(Paragraph(ptext, self.styles["Normal"]))

        numero = self.cotizacion.getNumero()
        today = datetime.now()
        fecha = datetime.strftime(today, "%d/%m/%Y")

        data = [["COTIZACIÓN N°", "FECHA"], [numero, fecha]]

        t = Table(data, 80, 20, hAlign='LEFT')
        t.setStyle(TableStyle([('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                               ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                               ('FONTSIZE', (0, 0), (-1, -1), 10),
                               ('FONTSIZE', (0, 1), (0, 1), 12),
                               ('ALIGN', (0, 0), (-1, -1), "CENTER"),
                               ('VALIGN', (0, 0), (-1, -1), "MIDDLE"),
                               ('BACKGROUND', (0, 0), (1, 0), colors.lightgrey),
                               ('BACKGROUND', (1, 2), (1, 2), colors.lightgrey)
                               ]))

        self.story.append(Spacer(1, 12))
        self.story.append(t)
        self.story.append(FrameBreak())

    def makeTable(self):
        """ Función que genera la tabla en donde se enlistan los servicios, sus códigos, cantidad, precio unitatio,
        precio total, el subtotal, el descuento total y el total a pagar
        """

        self.story.append(Spacer(1, 24))
        t = Table(self.table, [70, 40, 200, 80, 80, 80], 15, hAlign='CENTER')

        t.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                               ('ALIGN', (0, 0), (-1, 0), "CENTER"),
                               ('ALIGN', (3, 0), (-1, -1), "CENTER"),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                               ('ALIGN', (-3, 1), (-1, -1), "RIGHT"),
                               ]))
        self.story.append(t)
        self.story.append(Spacer(0, cm))

    def makeResumen(self):
        """ Función que genera la tabla en donde se enlistan los servicios ya usados """
        text = "RESUMEN"
        ptext = '<font size = 10> <b> %s </b></font>' % text
        self.story.append(Paragraph(ptext, self.styles["Center"]))
        self.story.append(Spacer(1, 6))

        t = Table(self.resumen, [200, 80, 80], 15, hAlign='CENTER')

        t.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                               ('ALIGN', (0, 0), (-1, 0), "CENTER"),
                               ('ALIGN', (0, 1), (0, -1), "LEFT"),
                               ('ALIGN', (-2, 1), (-1, -1), "RIGHT"),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                               ]))

        self.story.append(t)
        self.story.append(Spacer(0, cm))

    def doAll(self):
        """ Método que renderiza una cotización """

        self.makeFrames()
        self.makeTop()
        self.makeInfo()
        self.makeTable()
        self.makeResumen()
        self.makeEnd()
        temp = PageTemplate(frames=[self.frame1, self.frame2, self.last_frame],
                            onPage=self.drawPage)
        self.build(temp)
