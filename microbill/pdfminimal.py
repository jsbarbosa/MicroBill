import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings
from PyQt5.QtCore import QUrl

sys.argv += ["--disable-web-security", "--web-security=no"]
app = QtWidgets.QApplication(sys.argv)

QtWidgets.QApplication.setStyle(QtWidgets.QStyleFactory.create('Fusion')) # <- Choose the style

app.processEvents()

file = 'B18-0307.pdf'

HTML = '<object data="%s" type="application/pdf" width="100%%" height="100%%"><p>Alternative text - include a link <a href="%s">to the PDF!</a></p></object>'%(file, file)

view = QWebEngineView()
view.settings().setAttribute(QWebEngineSettings.PluginsEnabled, True)

url = QUrl('https://www.google.com')
# view.setUrl(url)
view.setHtml(HTML)
# self.urlview.load(url)
view.show()

app.exec_()
