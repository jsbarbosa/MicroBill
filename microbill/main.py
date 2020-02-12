import os
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from .constants import BASE_DIR
from .utils import export


@export
def testFiles():
    """ Función que revisa que los archivos de registro y clientes se encuentren configurados de manera correcta

    Raises
    -------
    Exception
        en caso que alguno de los archivos no se encuentre configurado de manera correcta
    """

    global REGISTRO_DATAFRAME, CLIENTES_DATAFRAME
    if list(REGISTRO_DATAFRAME.keys()) != constants.REGISTRO_KEYS:
        fields = ", ".join(constants.REGISTRO_KEYS)
        txt = "Registro no está bien configurado, las columnas deben ser: %s." % fields
        raise(Exception(txt))

    if list(CLIENTES_DATAFRAME.keys()) != constants.CLIENTES_KEYS:
        fields = ", ".join(constants.CLIENTES_KEYS)
        txt = "Clientes no está bien configurado, las columnas deben ser: %s." % fields
        raise(Exception(txt))
    """
        Equipo
    """
    for item in constants.EQUIPOS:
        df = eval("constants.%s"%item)
        keys = list(df.keys())
        if keys != constants.EQUIPOS_KEYS:
            fields = ", ".join(constants.EQUIPOS_KEYS)
            txt = "%s no está bien configurado, las columnas deben ser: %s." % (item, fields)
            raise(Exception(txt))


@export
def run() -> int:
    """ Función encargada de cargar la aplicación principal y sus dependencias. Antes de ejecutar la aplicación verifica
    la integridad de las dependencias incluyendo los archivos a los que el administrador tiene acceso

    Returns
    -------
    int: Código de salida de la aplicación al cerrarse
    """

    from . import config
    global correo, config, constants, sleep, Thread, MainWindow, REGISTRO_DATAFRAME, CLIENTES_DATAFRAME
    app = QtWidgets.QApplication(sys.argv)

    icon = QtGui.QIcon(os.path.join(BASE_DIR, 'icon.ico'))
    app.setWindowIcon(icon)

    splash_pix = QtGui.QPixmap(os.path.join(BASE_DIR, config.SPLASH_LOGO_PATH)).scaledToWidth(600)
    splash = QtWidgets.QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()

    try:
        from . import correo, config, constants
        from time import sleep
        from threading import Thread
        from .windows import MainWindow
        from .objects import REGISTRO_DATAFRAME, CLIENTES_DATAFRAME

        testFiles()

    except Exception as e:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setText("Error de configuración inicial:\n%s" % str(e))
        msg.setWindowTitle("Error")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()
        sys.exit()

    main = MainWindow()

    QtWidgets.QApplication.setStyle(QtWidgets.QStyleFactory.create('Fusion'))

    main.setWindowIcon(icon)
    splash.close()
    main.show()

    e = app.exec_()
    app = None
    return e
