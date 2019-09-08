import sys
import base64


def decode(enc: str) -> str:
    """ Descifra la informacion de la variable enc

    Parameters
    ----------
    enc : str
        El parametro a descifrar

    Returns
    -------
    str
        parametro enc descifrado
    """

    dec = []
    key = '1234567890123456'
    enc = base64.urlsafe_b64decode(enc).decode()
    for i in range(len(enc)):
        key_c = key[i % len(key)]
        dec_c = chr((256 + ord(enc[i]) - ord(key_c)) % 256)
        dec.append(dec_c)
    return "".join(dec)


def export(function_or_class):
    """ Función que sirve de decorator para determinar que clases, funciones o variables deben ser exportados

    Parameters
    ----------
        function_or_class: clase, función o variable
    """

    mod = sys.modules[function_or_class.__module__]
    if hasattr(mod, '__all__'):
        mod.__all__.append(function_or_class.__name__)
    else:
        mod.__all__ = [function_or_class.__name__]
    return function_or_class
