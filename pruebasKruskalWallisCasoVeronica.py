import pandas as pd
import numpy as np
from scipy.stats import kruskal, chi2
import matplotlib.pyplot as plt

archivo = "AplicandoFiltrado-InterGrupo.xlsx"
hojas = [
    "KRUSKAL-WALLIS-BAJO",
    "KRUSKAL-WALLIS-MEDIO",
    "KRUSKAL-WALLIS-ALTO"
]

def limpiar_puntaje(df):
    df = df.copy()
    df["Puntaje_Postest"] = pd.to_numeric(df["Puntaje_Postest"], errors="coerce")
    df = df.dropna(subset=["Puntaje_Postest"])
    return df

def prueba_kruskal(df):
    grupos = []
    for valor in sorted(df["Condicion_Experimental"].unique()):
        grupo = df[df["Condicion_Experimental"] == valor]["Puntaje_Postest"].values
        if len(grupo) > 0:
            grupos.append(grupo)
    if len(grupos) < 2:
        return {
            "Estadístico": np.nan,
            "Valor crítico": np.nan,
            "Valor p": np.nan,
            "Resultado": "No hay suficientes grupos",
            "Grados de libertad": np.nan
        }
    estadistico, p = kruskal(*grupos)
    gl = len(grupos) - 1
    valor_critico = chi2.ppf(0.95, df=gl)
    resultado = "Rechaza H0" if p < 0.05 else "No rechaza H0"
    return {
        "Estadístico": estadistico,
        "Valor crítico": valor_critico,
        "Valor p": p,
        "Resultado": resultado,
        "Grados de libertad": gl
    }

def grafico_colores(resultado, nombre_archivo):
    gl = resultado['Grados de libertad']
    estadistico = resultado['Estadístico']
    valor_critico = resultado['Valor crítico']
    valor_p = resultado['Valor p']
    decision = resultado['Resultado']

    if not (np.isfinite(estadistico) and np.isfinite(valor_critico) and np.isfinite(gl) and gl > 0):
        print(f"Datos no válidos para graficar en {nombre_archivo}")
        return

    x_max = max(estadistico, valor_critico) * 1.5
    x_max = max(x_max, 10)
    x = np.linspace(0, x_max, 1000)
    y = chi2.pdf(x, gl)

    plt.figure(figsize=(11, 7), constrained_layout=True)

    # Zonas de aceptación y rechazo
    plt.fill_between(x, 0, y, where=(x < valor_critico), color='#A0C4FF', alpha=0.7)
    plt.fill_between(x, 0, y, where=(x >= valor_critico), color='#FF8FA3', alpha=0.8)

    # Curva chi-cuadrado
    plt.plot(x, y, color='#22223B', lw=2)

    # Líneas verticales
    plt.axvline(estadistico, color='#00B140', linestyle='--', lw=2)
    plt.axvline(valor_critico, color='#7C3AED', linestyle='-', lw=2)

    ymax = max(y)
    # Zona de aceptación
    plt.text(valor_critico*0.35, ymax*0.7, 'Zona de aceptación\n(H₀)', color='#22223B', fontsize=15, ha='center', va='center', bbox=dict(facecolor='#A0C4FF', alpha=0.4, edgecolor='none'))
    # Zona de rechazo
    plt.text(valor_critico + (x_max-valor_critico)/2, ymax*0.7, 'Zona de rechazo\n(H₁)', color='#22223B', fontsize=15, ha='center', va='center', bbox=dict(facecolor='#FF8FA3', alpha=0.4, edgecolor='none'))

    # Estadístico
    plt.text(estadistico, ymax*0.55, f'Estadístico\n{estadistico:.3f}', color='#00B140', fontsize=13, ha='center', va='bottom', weight='bold', bbox=dict(facecolor='white', alpha=0.7, edgecolor='none'))
    # Valor crítico
    plt.text(valor_critico, ymax*0.4, f'Valor crítico\n{valor_critico:.3f}', color='#7C3AED', fontsize=13, ha='center', va='bottom', weight='bold', bbox=dict(facecolor='white', alpha=0.7, edgecolor='none'))
    # p-valor
    plt.text(estadistico, ymax*0.20, f'p-valor = {valor_p:.4f}', color='black', fontsize=14, ha='center', va='bottom', bbox=dict(facecolor='#FFD6A5', alpha=0.85, edgecolor='black'))

    # Bloque resumen EXPLÍCITO fuera del área de la gráfica
    resumen = (
        f"Estadístico H = {estadistico:.3f}\n"
        f"Valor crítico = {valor_critico:.3f}\n"
        f"p-valor = {valor_p:.4f}\n"
        f"gl = {gl}\n"
        f"Decisión: {decision}"
    )
    plt.gcf().text(0.75, 0.85, resumen, fontsize=14, ha='left', va='top',
                   bbox=dict(facecolor='#f5f5f5', alpha=0.95, edgecolor='#22223B', boxstyle='round,pad=0.7'))

    # Títulos y leyendas
    plt.title(f'Prueba Kruskal-Wallis: {nombre_archivo}', fontsize=18, weight='bold')
    plt.xlabel('Valor del estadístico', fontsize=16)
    plt.ylabel('Densidad de probabilidad', fontsize=16)
    # Leyenda manual
    from matplotlib.lines import Line2D
    legend_elements = [
        Line2D([0], [0], color='#00B140', linestyle='--', lw=2, label='Estadístico'),
        Line2D([0], [0], color='#7C3AED', linestyle='-', lw=2, label='Valor crítico'),
        Line2D([0], [0], color='#A0C4FF', lw=8, label='Zona de aceptación (H₀)'),
        Line2D([0], [0], color='#FF8FA3', lw=8, label='Zona de rechazo (H₁)')
    ]
    plt.legend(handles=legend_elements, fontsize=13, loc='upper left')
    plt.grid(True, linestyle=':', alpha=0.6)
    plt.savefig(f"{nombre_archivo}.svg", format='svg')
    plt.close()

# Procesar cada hoja y graficar
resultados = []
for hoja in hojas:
    df = pd.read_excel(archivo, sheet_name=hoja)
    df = limpiar_puntaje(df)
    res = prueba_kruskal(df)
    res["Muestra"] = hoja
    resultados.append(res)
    if res["Resultado"] != "No hay suficientes grupos":
        grafico_colores(res, hoja)

# Crear DataFrame de resultados y exportar
df_resultados = pd.DataFrame(resultados)[["Muestra", "Estadístico", "Valor crítico", "Valor p", "Resultado"]]
df_resultados.to_excel("Reporte_KruskalWallis.xlsx", index=False)
print("¡SVGs generados con valores explícitos y anotaciones completas!")
