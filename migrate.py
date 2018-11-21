import pickle
from glob import glob
from objects import Cotizacion

cot = glob("Old/*")

def attExist(object, att):
    try:
        eval("object.%s"%att)
        return True
    except AttributeError:
        return False

def assignAtt(object, att, value):
    if not attExist(object, att):
        if type(value) is str:
            exec("object.%s = '%s'"%(att, value))
        else:
            exec("object.%s = %s"%(att, value))
        return True
    return False

includes_cot = {'elaborado_por': '', 'modificado_por': '', 'aplicado_por': '',
            'observacion_correo': '', 'observacion_pdf': ''}

includes_ser = {'agregado_posteriormente': False}

for item in cot:
    i = 0
    with open(item, "rb") as fileObject:
        C = pickle.load(fileObject)
        for key in includes_cot.keys():
            assignAtt(C, key, includes_cot[key])
            for servicio in C.getServicios():
                for key in includes_ser.keys():
                    assignAtt(servicio, key, includes_ser[key])
            i += 1
    if i != 0:
        with open(item, "wb") as fileObject:
            pickle.dump(C, fileObject)