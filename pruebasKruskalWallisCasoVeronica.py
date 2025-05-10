import pandas as pd
import numpy as np
from scipy.stats import kruskal, chi2
import warnings

# Ignorar advertencias de pandas por celdas vacías o tipos mixtos
warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

# Archivo y hojas a analizar
archivo = "AplicandoFiltrado-InterGrupo.xlsx"
hojas = [
    "KRUSKAL-WALLIS-BAJO",
    "KRUSKAL-WALLIS-MEDIO",
    "KRUSKAL-WALLIS-ALTO"
]

# Función para limpiar y preparar los datos
def limpiar_puntaje(df):
    # Eliminar filas con valores no numéricos en Puntaje_Postest
    df = df.copy()
    df["Puntaje_Postest"] = pd.to_numeric(df["Puntaje_Postest"], errors="coerce")
    df = df.dropna(subset=["Puntaje_Postest"])
    return df

# Función para ejecutar la prueba Kruskal-Wallis
def prueba_kruskal(df):
    # Agrupar por grupo experimental (Condicion_Experimental)
    grupos = []
    nombres_grupos = []
    for valor in sorted(df["Condicion_Experimental"].unique()):
        grupo = df[df["Condicion_Experimental"] == valor]["Puntaje_Postest"].values
        if len(grupo) > 0:
            grupos.append(grupo)
            nombres_grupos.append(valor)
    if len(grupos) < 2:
        return {
            "Estadístico": np.nan,
            "Valor crítico": np.nan,
            "Valor p": np.nan,
            "Resultado": "No hay suficientes grupos"
        }
    # Prueba de Kruskal-Wallis
    estadistico, p = kruskal(*grupos)
    gl = len(grupos) - 1
    valor_critico = chi2.ppf(0.95, df=gl)
    resultado = "Rechaza H0" if p < 0.05 else "No rechaza H0"
    return {
        "Estadístico": estadistico,
        "Valor crítico": valor_critico,
        "Valor p": p,
        "Resultado": resultado
    }

# Procesar cada hoja y guardar resultados
resultados = []
for hoja in hojas:
    df = pd.read_excel(archivo, sheet_name=hoja)
    df = limpiar_puntaje(df)
    res = prueba_kruskal(df)
    res["Muestra"] = hoja
    resultados.append(res)

# Crear DataFrame de resultados y exportar
df_resultados = pd.DataFrame(resultados)[["Muestra", "Estadístico", "Valor crítico", "Valor p", "Resultado"]]
df_resultados.to_excel("Reporte_KruskalWallis.xlsx", index=False)

print("Reporte generado: Reporte_KruskalWallis.xlsx")
