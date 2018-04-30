import pandas as pd

FIELDS = ["Nombre", "Correo", "Teléfono", "Institución", "Documento", "Dirección", "Ciudad", "Interno", "Responsable"]

df = pd.read_excel("ClientesRAW.xlsx")

df = df[FIELDS]

for field in FIELDS:
    try:
        df[field] = df[field].str.strip()
    except:
        pass
df["Correo"] = df["Correo"].str.lower()

df = df.drop_duplicates("Correo", 'last')

df = df.sort_values("Nombre")

df.to_excel("Clientes.xlsx", index = False)



FIELDS = ["Cotización", "Fecha", "Nombre", "Correo", "Teléfono", "Institución", "Interno", "Responsable", "Muestra", "Equipo", "Valor"]

df = pd.read_excel("RegistroAFM.xlsx")

df = df[FIELDS]

t = df["Cotización"]

news = []
for i in t:
    news.append("A18-%04d"%i)

df["Cotización"] = news

df = df.drop_duplicates("Cotización")
df = df.sort_values("Cotización", ascending = False)

df.to_excel("Registro.xlsx", index = False)
