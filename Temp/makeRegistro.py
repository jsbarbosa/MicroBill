
import pandas as pd

FIELDS = ["Cotización", "Fecha", "Nombre", "Correo", "Teléfono", "Institución", "Interno", "Responsable", "Muestra", "Equipo", "Valor"]

df = pd.read_excel("Registro.xlsx")

df = df[FIELDS]

df.to_excel("Registro.xlsx", index = False)
