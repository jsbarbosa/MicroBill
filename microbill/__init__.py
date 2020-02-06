try:  # intenta importar el modulo config
    from .objects import *
    from . import config
except ImportError:  # si el modulo no se puede importar, lo crea usando la informacion de DEFAULT_CONFIG
    import os
    import sys
    from .constants import DEFAULT_CONFIG
    try:
        import config
    except ModuleNotFoundError:
        if os.path.exists('microbill'):
            file = os.path.join('microbill', 'config.py')
        else:
            file = 'config.py'
        with open(file, 'w', encoding="utf8") as file:
            file.write(DEFAULT_CONFIG)
        from . import config
    sys.modules['microbill.config'] = config

import os
import sys
from . import correo, utils
from .main import run
import importlib

importlib.reload(config)  # en caso que el programa se reinicie, el modulo config se vuelve a importar completamente

config.PASSWORD = utils.decode(config.PASSWORD)  #: almacena la contraseña del correo desencriptada

stdout_path = os.path.join(os.path.dirname(sys.executable), 'stdout.log')

if 'python.exe' not in sys.executable:
    sys.stdout = open(stdout_path, 'a')
    sys.stderr = sys.stdout
