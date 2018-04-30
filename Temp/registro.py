import pandas as pd

FIELDS = ["Cotización", "Fecha", "Nombre", "Correo", "Teléfono", "Institución", "Interno", "Responsable", "Muestra", "Equipo", "Valor"]

df = pd.read_excel("temp.xlsx")

df = df[FIELDS]

t = df["Cotización"]

news = []
for i in t:
    news.append("M18-%04d"%i)

df["Cotización"] = news

df = df.drop_duplicates("Cotización")
df = df.sort_values("Cotización", ascending = False)

df.to_excel("Temp2.xlsx", index = False)
