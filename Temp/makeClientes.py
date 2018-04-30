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
