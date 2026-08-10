import pandas as pd
import csv

entrada = "Facebook_Marketplace.csv"
salida = "Facebook_Marketplace.xlsx"

columnas = [
    "nombre",
    "precio",
    "disponibilidad",
    "sku",
    "publicado",
    "listing_id"
]

filas = []

with open(entrada, "r", encoding="utf-8-sig", newline="") as archivo:

    for linea in archivo:
        linea = linea.strip()

        if not linea:
            continue

        # Separar por comas
        partes = linea.split(",")

        # Verificar que tenga exactamente 6 datos
        if len(partes) == 6:
            partes = [x.strip() for x in partes]
            filas.append(partes)

        else:
            print("Fila ignorada:", linea)
            print("Cantidad de columnas:", len(partes))


# Crear DataFrame
df = pd.DataFrame(filas, columns=columnas)

# Exportar a Excel
df.to_excel(salida, index=False)

print()
print("===================================")
print("Excel generado correctamente")
print("Archivo:", salida)
print("Registros:", len(df))
print("Columnas:", len(df.columns))
print("===================================")