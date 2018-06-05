import os
from datetime import datetime, timedelta
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

import constants
from config import *

if os.path.isdir(constants.PDF_DIR): pass
else: os.makedirs(constants.PDF_DIR)

class PDFBase():
    def __init__(self, cotizacion = None, is_reporte = False):
        self.cotizacion = cotizacion
        if is_reporte:
            self.file_name = self.cotizacion.getNumero() + "_Reporte.pdf"
        else:
            self.file_name = self.cotizacion.getNumero() + ".pdf"
        self.file_name = os.path.join(constants.PDF_DIR, self.file_name)

        self.cotizacion.setFileName(self.file_name)

        self.doc =  SimpleDocTemplate(self.file_name, pagesize = letter,
                                rightMargin = cm, leftMargin = cm,
                                topMargin = cm, bottomMargin = cm)

        self.styles = getSampleStyleSheet()
        self.styles.add(ParagraphStyle(name = 'Center', alignment = TA_CENTER))
        self.styles.add(ParagraphStyle(name = 'Justify', alignment = TA_JUSTIFY))

        border = ParagraphStyle(name = 'Border', alignment = TA_JUSTIFY)
        border.borderWidth = 0.5
        border.borderColor = colors.black
        border.borderRadius = 2
        border.borderPadding = 3
        self.styles.add(border)

        self.story = []

    def makeInfo(self):
        usuario = self.cotizacion.getUsuario()
        ptext = '<font size = 12><b>%s</b></font>' % "%s - UNIVERSIDAD DE LOS ANDES"%CENTRO.upper()
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
            Paragraph(usuario.getCiudad(), self.styles["Border"]),
            ]]
        t = Table(data, [c1, c2, c3, c4], hAlign='LEFT')
        self.story.append(Spacer(1, h))
        self.story.append(t)

        data = [[Paragraph("<b>Correo:</b>", self.styles["Normal"]),
            Paragraph(usuario.getCorreo(), self.styles["Border"]),
            Paragraph("<b>Muestra:</b>", self.styles["Normal"]),
            Paragraph(self.cotizacion.getMuestra(), self.styles["Border"]),
            ]]
        t = Table(data, [c1, c2, c3, c4], hAlign='LEFT')
        self.story.append(Spacer(1, h))
        self.story.append(t)

    def makeEnd(self):
        for i in range(len(DEPENDENCIAS)):
            text = DEPENDENCIAS[i]
            if i < 2:
                ptext = '<font size = 10> <b> %s </b></font>'%text
            else:
                ptext = '<font size = 10>%s</font>'%text
            self.story.append(Paragraph(ptext, self.styles["Center"]))

    def drawPage(self, canvas, doc):
        canvas.setTitle(CENTRO)
        canvas.setSubject("Cotización")
        canvas.setAuthor("Juan Barbosa")
        canvas.setCreator("MicroBill")
        self.styles = getSampleStyleSheet()

        P = Paragraph("<font size = 8>Creado usando MicroBill - Juan Barbosa (github.com/jsbarbosa)</font>",
                  self.styles["Normal"])
        w, h = P.wrap(doc.width, doc.bottomMargin)
        P.drawOn(canvas, doc.leftMargin, h)

    def build(self, temp):
        self.doc.addPageTemplates([temp])
        self.doc.build(self.story)

class PDFCotizacion(PDFBase):
    def __init__(self, cotizacion):
        super(PDFCotizacion, self).__init__(cotizacion)

    def makeFrames(self):
        #Two Columns
        h = 4*cm
        self.frame1 = Frame(self.doc.leftMargin, self.doc.bottomMargin + self.doc.height - h, self.doc.width / 2 - 6, h, id='col1')
        self.frame2 = Frame(self.doc.leftMargin + self.doc.width / 2 + 2*cm, self.doc.bottomMargin + self.doc.height - h, self.doc.width / 2 - 6, h, id='col2')

        space = 0.5*cm
        info_h = 4*cm
        table_h = 7*cm
        observaciones_h = 3*cm
        terminos_h = 2.5*cm

        last_h = self.doc.bottomMargin + self.doc.height - h
        self.last_frame = Frame(self.doc.leftMargin, self.doc.bottomMargin, self.doc.width, last_h - cm)

    def makeTop(self):
        height = 1.5

        logo = Image("logo.png", height*3.33*cm, height*cm, hAlign = "CENTER")

        self.story.append(logo)
        self.story.append(Spacer(0, 12))
        data = [["CÓDIGO S.GESTIÓN"], [CODIGO_GESTION]]
        t = Table(data, 90, 20)
        t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                               ('BOX', (0,0), (-1,-1), 0.25, colors.black),
                               ('FONTSIZE', (0, 0), (-1, -1), 8),
                               ]))

        self.story.append(t)
        self.story.append(FrameBreak())

        ptext = '<font size = 12><b>%s</b></font>' % "COTIZACIÓN DE SERVICIOS"
        self.story.append(Paragraph(ptext, self.styles["Normal"]))

        numero = self.cotizacion.getNumero()
        today = datetime.now()
        fecha = datetime.strftime(today, "%d/%m/%Y")
        until = datetime.strftime(today + timedelta(days = 45), "%d/%m/%Y")

        data = [["COTIZACIÓN N°", "FECHA"], [numero, fecha], [numero, "VALIDA HASTA"], [numero, until]]

        t = Table(data, 80, 15, hAlign='LEFT')
        t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                               ('BOX', (0,0), (-1,-1), 0.25, colors.black),
                               ('FONTSIZE', (0, 0), (-1, -1), 8),
                               ('FONTSIZE', (0, 1), (0, 1), 12),
                               ('ALIGN', (0, 1), (0, 1), "CENTER"),
                               ('VALIGN', (0, 1), (0, 1), "MIDDLE"),
                               ('SPAN',(0,1),(0,-1)),
                               ('BACKGROUND', (0, 0), (1, 0), colors.lightgrey),
                               ('BACKGROUND', (1, 2), (1, 2), colors.lightgrey)
                                ]))

        self.story.append(Spacer(1, 12))
        self.story.append(t)

    def makeTable(self):
        table = self.cotizacion.makeCotizacionTable()
        total = self.cotizacion.getTotal()
        table.insert(0, ["COD", "SERVICIO", "CANTIDAD", "PRECIO UNIDAD", "PRECIO TOTAL"])
        table.append(["", "", "", "TOTAL", "{:,}".format(total)])

        self.story.append(Spacer(1, 24))

        t = Table(table, [40, 260, 70, 90, 90], 15, hAlign='CENTER')

        t.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.25, colors.black),
                               ('ALIGN', (0, 0), (-1, 0), "CENTER"),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                               ('ALIGN', (-3, 0), (-1, -1), "RIGHT"),
                               ('SPAN',(0, -1),(-3, -1)),
                                ]))
        self.story.append(t)
        self.story.append(Spacer(1, 24))

    def makeObservaciones(self):
        text = "OBSERVACIONES"
        ptext = '<font size = 10> <b> %s </b></font>'%text

        self.story.append(Paragraph(ptext, self.styles["Center"]))
        self.story.append(Spacer(1, 6))

        usuario = self.cotizacion.getUsuario()

        text = [("Responsable", usuario.getResponsable()),
                ("Proyecto", usuario.getProyecto()),
                ("Código", usuario.getCodigo())]

        ptext = '<font size = 9>%s</font><br/>'%self.cotizacion.getObservacionPDF().replace("\n", "<br/>")
        if usuario.getInterno() == "Interno":
            for item in text:
                ptext += '<font size = 9> <b>%s:</b> <u>%s</u></font> <br/>'%item

        self.story.append(Paragraph(ptext, self.styles["Border"]))

        self.story.append(Spacer(1, 24))

    def makeTerminos(self):
        text = "TÉRMINOS Y CONDICIONES"
        ptext = '<font size = 10> <b> %s </b></font>'%text
        self.story.append(Paragraph(ptext, self.styles["Center"]))
        self.story.append(Spacer(1, 6))

        ptext = ""

        for i in range(len(TERMINOS_Y_CONDICIONES)):
            text = TERMINOS_Y_CONDICIONES[i]
            ptext += '<font size = 8>%d. %s</font><br/>'%(i + 1, text)
            # self.story.append(Paragraph(ptext, self.styles["Justify"]))

        ptext += '<font size = 8>%d.</font> <font size = 6>%s</font>'%(i + 2, CONFIDENCIALIDAD)
        self.story.append(Paragraph(ptext, self.styles["Border"]))
        self.story.append(Spacer(1, 24))

    def doAll(self):
        self.makeFrames()
        self.makeTop()
        self.makeInfo()
        self.makeTable()
        self.makeObservaciones()
        self.makeTerminos()
        # self.story.append(FrameBreak())
        self.makeEnd()
        # temp = PageTemplate(frames = [self.frame1, self.frame2, self.info_frame, self.table_frame,
        #                 self.observaciones_frame, self.terminos_frame, self.end_frame],
        #             onPage = self.drawPage)
        temp = PageTemplate(frames = [self.frame1, self.frame2, self.last_frame],
                    onPage = self.drawPage)
        self.build(temp)

class PDFReporte(PDFBase):
    def __init__(self, cotizacion):
        super(PDFReporte, self).__init__(cotizacion, True)

    def makeFrames(self):
        self.table = self.cotizacion.makeReporteTable()
        self.resumen = self.cotizacion.makeResumenTable()

        self.resumen.insert(0, ["DESCRIPCIÓN", "COTIZADOS", "RESTANTES"])
        self.table.insert(0, ["FECHA", "COD", "DESCRIPCIÓN", "COTIZADOS", "USADOS", "RESTANTES"])

        # Two Columns
        h = 3*cm
        self.frame1 = Frame(self.doc.leftMargin, self.doc.bottomMargin + self.doc.height - h, self.doc.width / 2 - 6, h, id='col1')
        self.frame2 = Frame(self.doc.leftMargin + self.doc.width / 2 + 2*cm, self.doc.bottomMargin + self.doc.height - h, self.doc.width / 2 - 6, h, id='col2')

        space = 0.5*cm
        info_h = 4*cm
        table_h = (len(self.table) + 1)*15
        observaciones_h = (len(self.resumen) + 2)*15
        terminos_h = 2.5*cm

        last_h = self.doc.bottomMargin + self.doc.height - h
        # self.info_frame = Frame(self.doc.leftMargin, last_h - info_h - space, self.doc.width, info_h, showBoundary = 1)
        # last_h += -info_h - space

        self.last_frame = Frame(self.doc.leftMargin, self.doc.bottomMargin, self.doc.width, last_h - cm)

        # self.table_frame = Frame(self.doc.leftMargin, self.doc.bottomMargin, self.doc.width, last_h - cm)

    def makeTop(self):
        height = 1.5
        logo = Image("logo.png", height*3.33*cm, height*cm, hAlign = "CENTER")

        self.story.append(logo)
        self.story.append(FrameBreak())

        ptext = '<font size = 12><b>%s</b></font>' % "REPORTE DE SERVICIOS"
        self.story.append(Paragraph(ptext, self.styles["Normal"]))

        numero = self.cotizacion.getNumero()
        today = datetime.now()
        fecha = datetime.strftime(today, "%d/%m/%Y")

        data = [["COTIZACIÓN N°", "FECHA"], [numero, fecha]]

        t = Table(data, 80, 20, hAlign='LEFT')
        t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                               ('BOX', (0,0), (-1,-1), 0.25, colors.black),
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
        self.story.append(Spacer(1, 24))
        t = Table(self.table, [70, 40, 200, 80, 80, 80], 15, hAlign='CENTER')

        t.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.25, colors.black),
                               ('ALIGN', (0, 0), (-1, 0), "CENTER"),
                               ('ALIGN', (3, 0), (-1, -1), "CENTER"),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                               ('ALIGN', (-3, 1), (-1, -1), "RIGHT"),
                                ]))
        self.story.append(t)
        self.story.append(Spacer(0, cm))

    def makeResumen(self):
        text = "RESUMEN"
        ptext = '<font size = 10> <b> %s </b></font>'%text
        self.story.append(Paragraph(ptext, self.styles["Center"]))
        self.story.append(Spacer(1, 6))

        t = Table(self.resumen, [200, 80, 80], 15, hAlign='CENTER')

        t.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.25, colors.black),
                               ('ALIGN', (0, 0), (-1, 0), "CENTER"),
                               ('ALIGN', (0, 1), (0, -1), "LEFT"),
                               ('ALIGN', (-2, 1), (-1, -1), "RIGHT"),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                                ]))

        self.story.append(t)
        self.story.append(Spacer(0, cm))

    def doAll(self):
        self.makeFrames()
        self.makeTop()
        self.makeInfo()
        self.makeTable()
        self.makeResumen()
        self.makeEnd()
        # temp = PageTemplate(frames = [self.frame1, self.frame2, self.info_frame, self.table_frame],
        #                 onPage = self.drawPage)
        temp = PageTemplate(frames = [self.frame1, self.frame2, self.last_frame],
                        onPage = self.drawPage)
        self.build(temp)
