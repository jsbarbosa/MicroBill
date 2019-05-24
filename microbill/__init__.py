try:
    from .objects import *
    from . import config
except ImportError:
    import os
    import sys
    from .constants import DEFAULT_CONFIG
    try:
        import config
    except ModuleNotFoundError:
        if os.path.exists('microbill'): file = os.path.join('microbill', 'config.py')
        else: file = 'config.py'
        with open(file, 'w', encoding = "utf8") as file: file.write(DEFAULT_CONFIG)
        from . import config
    sys.modules['microbill.config'] = config

from . import correo
from .main import run
import importlib

importlib.reload(config)

def decode(enc):
    import base64
    dec = []
    key = '1234567890123456'
    enc = base64.urlsafe_b64decode(enc).decode()
    for i in range(len(enc)):
        key_c = key[i % len(key)]
        dec_c = chr((256 + ord(enc[i]) - ord(key_c)) % 256)
        dec.append(dec_c)
    return "".join(dec)

config.PASSWORD = decode(config.PASSWORD)
