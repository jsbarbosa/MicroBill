import pickle
from glob import glob
from objects import Cotizacion, Servicio

cot = glob("Old/M*")

for item in cot:
    new_name = item.replace("M", "B")
    with open(item, "rb") as file:
        C = pickle.load(file)
        C.numero = C.numero.replace("M", "B")
        servicios = C.getServicios()
        for servicio in servicios:
            servicio.equipo = "BE"


    with open(new_name, "wb") as file:
        pickle.dump(C, file)