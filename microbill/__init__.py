try:
    from .objects import *
except ImportError:
    import sys
    import config
    import login
    sys.modules['microbill.config'] = config
    sys.modules['microbill.login'] = login
    
from . import correo
from .main import run
