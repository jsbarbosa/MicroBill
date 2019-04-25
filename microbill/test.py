import objects
import pdflib

cot = objects.Cotizacion().load('0119-0001')
pdflib.PDFCotizacion(cot).doAll()
