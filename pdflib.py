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

        self.doc =  SimpleDocTemplate(self.file_name, pagesize = letter,
                                rightMargin = cm, leftMargin = cm,
                                topMargin = cm, bottomMargin = cm)

        self.styles = getSampleStyleSheet()
        self.styles.add(ParagraphStyle(name = 'Center', alignment = TA_CENTER))
        self.styles.add(ParagraphStyle(name = 'Justify', alignment = TA_JUSTIFY))

        self.story = []

    def makeInfo(self):
        usuario = self.cotizacion.getUsuario()

        self.story.append(FrameBreak())

        ptext = '<font size = 12><b>%s</b></font>' % "%s - UNIVERSIDAD DE LOS ANDES"%CENTRO.upper()
        self.story.append(Paragraph(ptext, self.styles["Center"]))

        self.story.append(Spacer(1, 12))

        ptext = '<font size = 10>'
        ptext += "<b>%s</b>" % "Nombre:"
        ptext += "&nbsp %s" % usuario.getNombre()

        ptext += "&nbsp &nbsp &nbsp &nbsp <b>%s</b>" % "Institución:"
        ptext += "&nbsp %s" %  usuario.getInstitucion()

        ptext += "&nbsp &nbsp &nbsp &nbsp <b>%s</b>" % "Documento (Nit/C.C.):"
        ptext += "&nbsp %s" % usuario.getDocumento()
        ptext += '</font>'

        self.story.append(Paragraph(ptext, self.styles["Justify"]))
        self.story.append(Spacer(1, 6))

        ptext = '<font size = 10>'
        ptext += "<b>%s</b>" % "Teléfono:"
        ptext += "&nbsp %s" % usuario.getTelefono()

        ptext += "&nbsp &nbsp &nbsp &nbsp <b>%s</b>" % "Dirección:"
        ptext += "&nbsp %s" % usuario.getDireccion()

        ptext += "&nbsp &nbsp &nbsp &nbsp <b>%s</b>" % "Ciudad:"
        ptext += "&nbsp %s" % usuario.getCiudad()
        ptext += '</font>'

        self.story.append(Paragraph(ptext, self.styles["Justify"]))

        self.story.append(Spacer(1, 6))

        ptext = '<font size = 10>'
        ptext += "<b>%s</b>" % "Correo:"
        ptext += "&nbsp %s" % usuario.getCorreo()

        ptext += "&nbsp &nbsp &nbsp &nbsp <b>%s</b>" % "Tipo de Muestra:"
        ptext += "&nbsp %s" % self.cotizacion.getMuestra()
        ptext += '</font>'

        self.story.append(Paragraph(ptext, self.styles["Justify"]))

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
        self.info_frame = Frame(self.doc.leftMargin, last_h - info_h - space, self.doc.width, info_h, showBoundary = 1)
        last_h += - info_h - space
        self.table_frame = Frame(self.doc.leftMargin, last_h - table_h - space, self.doc.width, table_h)
        last_h += - table_h - space
        self.observaciones_frame = Frame(self.doc.leftMargin, last_h - observaciones_h - space, self.doc.width, observaciones_h, showBoundary = 1)
        last_h += - observaciones_h - space
        self.terminos_frame = Frame(self.doc.leftMargin, last_h - terminos_h - space, self.doc.width, terminos_h, showBoundary = 1)
        last_h += - terminos_h - space
        self.end_frame = Frame(self.doc.leftMargin, last_h - 3*cm - space, self.doc.width, 3*cm)

    def makeTop(self):
        height = 1.5

        logo = Image("logo.png", height*3.33*cm, height*cm, hAlign = "CENTER")

        self.story.append(logo)
        self.story.append(Spacer(0, 12))
        data = [["CÓDIGO UNIANDES", "CÓDIGO S.GESTIÓN"], [CODIGO_UNIANDES, CODIGO_GESTION]]
        t = Table(data, 90, 20)
        t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                               ('BOX', (0,0), (-1,-1), 0.25, colors.black),
                               ('FONTSIZE', (0, 0), (1, 1), 8),
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
                               ('BACKGROUND', (0, 0), (1, 0), colors.grey),
                               ('BACKGROUND', (1, 2), (1, 2), colors.grey)
                                ]))

        self.story.append(Spacer(1, 12))
        self.story.append(t)

    def makeTable(self):
        table = self.cotizacion.makeCotizacionTable()
        total = self.cotizacion.getTotal()
        table.insert(0, ["COD", "SERVICIO", "CANTIDAD", "PRECIO UNIDAD", "PRECIO TOTAL"])
        table.append(["", "", "", "TOTAL", "{:,}".format(total)])

        self.story.append(FrameBreak())

        t = Table(table, [40, 260, 70, 90, 90], 15, hAlign='CENTER')

        t.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.25, colors.black),
                               ('ALIGN', (0, 0), (-1, 0), "CENTER"),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                               ('ALIGN', (-3, 0), (-1, -1), "RIGHT"),
                               ('SPAN',(0, -1),(-3, -1)),
                                ]))
        self.story.append(t)

    def makeObservaciones(self):
        self.story.append(FrameBreak())
        text = "OBSERVACIONES"
        ptext = '<font size = 10> <b> %s </b></font>'%text
        self.story.append(Paragraph(ptext, self.styles["Center"]))
        self.story.append(Spacer(1, 6))

        usuario = self.cotizacion.getUsuario()

        text = [("Responsable", usuario.getResponsable()),
                ("Proyecto", usuario.getProyecto()),
                ("Código", usuario.getCodigo())]

        if usuario.getInterno() == "Interno":
            for item in text:
                ptext = '<font size = 9> <b>%s:</b> <u>%s</u></font>'%item
                self.story.append(Paragraph(ptext, self.styles["Normal"]))

    def makeTerminos(self):
        self.story.append(FrameBreak())

        text = "TÉRMINOS Y CONDICIONES"
        ptext = '<font size = 10> <b> %s </b></font>'%text
        self.story.append(Paragraph(ptext, self.styles["Center"]))
        self.story.append(Spacer(1, 6))

        for i in range(len(TERMINOS_Y_CONDICIONES)):
            text = TERMINOS_Y_CONDICIONES[i]
            ptext = '<font size = 8>%d. %s</font>'%(i + 1, text)
            self.story.append(Paragraph(ptext, self.styles["Normal"]))

    def doAll(self):
        self.makeFrames()
        self.makeTop()
        self.makeInfo()
        self.makeTable()
        self.makeObservaciones()
        self.makeTerminos()
        self.story.append(FrameBreak())
        self.makeEnd()
        temp = PageTemplate(frames = [self.frame1, self.frame2, self.info_frame, self.table_frame,
                        self.observaciones_frame, self.terminos_frame, self.end_frame],
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
        self.info_frame = Frame(self.doc.leftMargin, last_h - info_h - space, self.doc.width, info_h, showBoundary = 1)
        last_h += -info_h - space

        self.table_frame = Frame(self.doc.leftMargin, self.doc.bottomMargin, self.doc.width, last_h - cm)

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
                               ('BACKGROUND', (0, 0), (1, 0), colors.grey),
                               ('BACKGROUND', (1, 2), (1, 2), colors.grey)
                                ]))

        self.story.append(Spacer(1, 12))
        self.story.append(t)

    def makeTable(self):
        self.story.append(FrameBreak())
        t = Table(self.table, [70, 40, 200, 80, 80, 80], 15, hAlign='CENTER')

        t.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.25, colors.black),
                               ('ALIGN', (0, 0), (-1, 0), "CENTER"),
                               ('ALIGN', (3, 0), (-1, -1), "CENTER"),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
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
                               ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
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
        temp = PageTemplate(frames = [self.frame1, self.frame2, self.info_frame, self.table_frame],
                        onPage = self.drawPage)
        self.build(temp)
