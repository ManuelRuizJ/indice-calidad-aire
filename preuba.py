import pandas as pd
df = pd.read_excel("datos/datos_calidad_aire.xlsx", header=None)
# Buscar fila con "Fecha & Hora"
for i, row in df.iterrows():
    if isinstance(row[0], str) and "Fecha & Hora" in row[0]:
        start = i
        break
fechas = pd.to_datetime(df.iloc[start+4:, 0].astype(str), errors='coerce', dayfirst=True)
print(fechas.unique())