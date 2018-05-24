import pickle
from glob import glob
from objects import Cotizacion

cot = glob("Old/*")
for item in cot:
    with open(item, "rb") as fileObject:
        C = pickle.load(fileObject)
        C.setElaborado("Humberto Ibarra")

    with open(item, "wb") as fileObject:
        pickle.dump(C,fileObject)
