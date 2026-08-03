"""
Genera las 4 graficas del reporte, como archivos .png dentro de la carpeta
'reportes/' (junto a este script). Se pueden volver a correr las veces que
quieras; cada corrida sobreescribe las mismas 4 imagenes.
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from consultas import tendencia_diaria, promedio_por_dia_semana, promedio_por_categoria, promedio_por_dia_y_categoria

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARPETA_REPORTES = os.path.join(BASE_DIR, 'reportes')
os.makedirs(CARPETA_REPORTES, exist_ok=True)

AZUL = "#1F4E79"
AZUL_CLARO = "#2E75B6"
GRIS = "#8c8c8c"
COLORES_CAT = {"Estudiantes secundaria": "#1F4E79", "Estudiantes primaria": "#2E75B6", "Docentes": "#8FAADC"}

plt.rcParams.update({
    "font.size": 11,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.edgecolor": "#888888",
})


def ruta(nombre_archivo):
    return os.path.join(CARPETA_REPORTES, nombre_archivo)


def generar_todas():
    # --- Grafica 1: tendencia diaria general (linea) ---
    df_tendencia = tendencia_diaria()
    fig, ax = plt.subplots(figsize=(7.8, 3.2))
    ax.plot(df_tendencia['fecha'], df_tendencia['porcentaje'] * 100, color=AZUL, linewidth=2, marker='o', markersize=3)
    ax.axhline(df_tendencia['porcentaje'].mean() * 100, color=GRIS, linestyle='--', linewidth=1, label="Promedio del periodo")
    ax.set_ylabel("% de asistencia")
    ax.set_ylim(50, 100)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%b'))
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=0))
    fig.autofmt_xdate(rotation=0, ha='center')
    ax.legend(frameon=False, loc="lower left", fontsize=9)
    ax.set_title("Tendencia de asistencia general por día (todas las categorías)", color=AZUL, fontsize=12, fontweight='bold', loc='left')
    fig.tight_layout()
    fig.savefig(ruta("grafica_tendencia.png"), dpi=170)
    plt.close(fig)

    # --- Grafica 2: promedio por dia de la semana (barras) ---
    prom_dia = promedio_por_dia_semana()
    fig, ax = plt.subplots(figsize=(7.8, 3.2))
    barras = ax.bar(prom_dia.index, prom_dia.values * 100, color=AZUL_CLARO, width=0.55)
    for b in barras:
        ax.text(b.get_x() + b.get_width()/2, b.get_height() + 1, f"{b.get_height():.0f}%", ha='center', fontsize=10, color=AZUL)
    ax.set_ylabel("% de asistencia promedio")
    ax.set_ylim(0, 100)
    ax.set_title("Asistencia promedio por día de la semana", color=AZUL, fontsize=12, fontweight='bold', loc='left')
    fig.tight_layout()
    fig.savefig(ruta("grafica_dia_semana.png"), dpi=170)
    plt.close(fig)

    # --- Grafica 3: promedio por categoria (barras horizontales) ---
    prom_cat = promedio_por_categoria().sort_values()
    fig, ax = plt.subplots(figsize=(7.8, 2.6))
    colores = [COLORES_CAT.get(c, AZUL) for c in prom_cat.index]
    barras = ax.barh(prom_cat.index, prom_cat.values * 100, color=colores, height=0.5)
    for b in barras:
        ax.text(b.get_width() + 1, b.get_y() + b.get_height()/2, f"{b.get_width():.0f}%", va='center', fontsize=10, color=AZUL)
    ax.set_xlabel("% de asistencia promedio")
    ax.set_xlim(0, 100)
    ax.set_title("Asistencia promedio por categoría (todo el periodo)", color=AZUL, fontsize=12, fontweight='bold', loc='left')
    fig.tight_layout()
    fig.savefig(ruta("grafica_categoria.png"), dpi=170)
    plt.close(fig)

    # --- Grafica 4: dia de semana x categoria (barras agrupadas) ---
    tabla = promedio_por_dia_y_categoria()
    fig, ax = plt.subplots(figsize=(7.8, 3.4))
    x = range(len(tabla.index))
    ancho = 0.25
    for i, cat in enumerate(tabla.columns):
        offsets = [xi + (i - 1) * ancho for xi in x]
        ax.bar(offsets, tabla[cat].values * 100, width=ancho, label=cat, color=COLORES_CAT.get(cat, AZUL))
    ax.set_xticks(list(x))
    ax.set_xticklabels(tabla.index)
    ax.set_ylabel("% de asistencia promedio")
    ax.set_ylim(0, 100)
    ax.legend(frameon=False, fontsize=8, loc="lower center", bbox_to_anchor=(0.5, -0.32), ncol=3)
    ax.set_title("Asistencia por día de la semana, desglosada por categoría", color=AZUL, fontsize=12, fontweight='bold', loc='left')
    fig.tight_layout()
    fig.savefig(ruta("grafica_dia_categoria.png"), dpi=170)
    plt.close(fig)

    print(f"Gráficas guardadas en: {CARPETA_REPORTES}")


if __name__ == "__main__":
    generar_todas()