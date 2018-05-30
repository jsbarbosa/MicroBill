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

includes = {'elaborado_por': '', 'modificado_por': '', 'aprobado_por': '',
            'observacion_correo': '', 'observacion_pdf': ''}

for item in cot:
    with open(item, "rb") as fileObject:
        C = pickle.load(fileObject)
        for key in includes.keys():
            assignAtt(C, key, includes[key])
    with open(item, "wb") as fileObject:
        pickle.dump(C,fileObject)
