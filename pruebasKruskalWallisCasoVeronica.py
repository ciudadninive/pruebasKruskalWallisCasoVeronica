import pandas as pd
from scipy.stats import kruskal, chi2_contingency

# Cargar datos
archivo = "MuestrasAplicandoFiltradoInterIntraGrupo.xlsx"
df_control = pd.read_excel(archivo, sheet_name="KRUSKALL-WALLIS-CE1")
df_experimental = pd.read_excel(archivo, sheet_name="KRUSKALL-WALLIS-CE2")

# Limpiar columnas y valores
for df in [df_control, df_experimental]:
    df.columns = [col.strip() for col in df.columns]
    df['NivelLectura-Pretest'] = df['NivelLectura-Pretest'].astype(str).str.strip().str.capitalize()

# Niveles de lectura presentes en ambos grupos
niveles = sorted(list(set(df_control['NivelLectura-Pretest'].unique()) | set(df_experimental['NivelLectura-Pretest'].unique())))

# Valor crítico chi-cuadrado para gl=1, alfa=0.05
valor_critico = 3.841

resultados = []

for nivel in niveles:
    # Filtrar postest por nivel
    control = df_control[df_control['NivelLectura-Pretest'] == nivel]['Puntaje_Postest'].dropna()
    experimental = df_experimental[df_experimental['NivelLectura-Pretest'] == nivel]['Puntaje_Postest'].dropna()
    n_control = len(control)
    n_experimental = len(experimental)
    if n_control > 0 and n_experimental > 0:
        stat, p = kruskal(control, experimental)
        if stat >= valor_critico and p < 0.05:
            conclusion = "Rechazamos H0: diferencia significativa entre grupos."
        else:
            conclusion = "No se rechaza H0: no hay diferencia significativa entre grupos."
        resultados.append({
            "Nivel de Lectura": nivel,
            "N Control": n_control,
            "N Experimental": n_experimental,
            "Estadístico H": stat,
            "Valor crítico (tabla)": valor_critico,
            "Valor p": p,
            "Resultado": conclusion
        })

# Crear DataFrame y exportar a Excel
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel("Reporte_KruskalWallis_Resultados.xlsx", index=False)
print(df_resultados)
