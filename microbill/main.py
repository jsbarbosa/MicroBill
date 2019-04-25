import sys
from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWebEngineWidgets import QWebEngineView

def threadedCorreo():
    while True:
        try:
            correo.initCorreo()
            sleep(10 * 60)
        except Exception as e:
            sleep(1 * 60)
            print(e)

def testFiles():
    """
        Registro y clientes
    """
    if (list(REGISTRO_DATAFRAME.keys()) != constants.REGISTRO_KEYS):
        fields = ", ".join(constants.REGISTRO_KEYS)
        txt = "Registro no está bien configurado, las columnas deben ser: %s."%fields
        raise(Exception(txt))

    if (list(CLIENTES_DATAFRAME.keys()) != constants.CLIENTES_KEYS):
        fields = ", ".join(constants.CLIENTES_KEYS)
        txt = "Clientes no está bien configurado, las columnas deben ser: %s."%fields
        raise(Exception(txt))
    """
        Equipo
    """
    for item in constants.EQUIPOS:
        df = eval("constants.%s"%item)
        keys = list(df.keys())
        if (keys != constants.EQUIPOS_KEYS):
            fields = ", ".join(constants.EQUIPOS_KEYS)
            txt = "%s no está bien configurado, las columnas deben ser: %s."%(item, fields)
            raise(Exception(txt))

def run():
    sys.argv += ["--disable-web-security", "--web-security=no", "--allow-file-access-from-files"]
    app = QtWidgets.QApplication(sys.argv)

    icon = QtGui.QIcon('icon.ico')
    app.setWindowIcon(icon)

    splash_pix = QtGui.QPixmap('logo.png').scaledToWidth(600)
    splash = QtWidgets.QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()

    try:
        import os
        from . import correo, config, constants
        from time import sleep
        from threading import Thread
        from windows import MainWindow
        from .objects import REGISTRO_DATAFRAME, CLIENTES_DATAFRAME

        testFiles()

    except Exception as e:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setText("Error de configuración inicial:\n%s"%str(e))
        msg.setWindowTitle("Error")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()
        sys.exit()

    main = MainWindow()

    thread = Thread(target = threadedCorreo)
    thread.setDaemon(True)
    thread.start()

    QtWidgets.QApplication.setStyle(QtWidgets.QStyleFactory.create('Fusion'))

    main.setWindowIcon(icon)
    splash.close()
    main.show()

    sys.exit(app.exec_())

if __name__ == '__main__':
    run()
