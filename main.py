import os
import sys
import config
import constants
from time import sleep
from threading import Thread
from PyQt5 import QtCore, QtGui, QtWidgets

import correo
from windows import MainWindow

def threadedCorreo():
    try: correo.initCorreo()
    except Exception as e: print(e)

app = QtWidgets.QApplication(sys.argv)

icon = QtGui.QIcon('icon.ico')
app.setWindowIcon(icon)

splash_pix = QtGui.QPixmap('logo.png').scaledToWidth(600)
splash = QtWidgets.QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)
splash.show()
app.processEvents()

main = MainWindow()

sleep(2)

thread = Thread(target = threadedCorreo)
thread.setDaemon(True)
thread.start()

QtWidgets.QApplication.setStyle(QtWidgets.QStyleFactory.create('Fusion')) # <- Choose the style
# QtWidgets.QApplication.setStyle(QtWidgets.QStyleFactory.create('Windows')) # <- Choose the style

main.setWindowIcon(icon)
splash.close()
main.show()

sys.exit(app.exec_())
