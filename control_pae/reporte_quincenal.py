"""
Genera el reporte de asistencia (PDF) a partir de la base de datos SQLite
real del sistema. Este archivo debe vivir en la MISMA carpeta que
entradas_pae26.py, que lo importa y lo llama automáticamente cada vez que
se cierra una QUINCENA DE CALENDARIO (días 1-15, y 16-fin de mes) — ver
gestionar_reporte_quincenal() en entradas_pae26.py.

También se puede correr manualmente en cualquier momento:
    python reporte_quincenal.py
(usará la carpeta 'datos/asistencia.db' relativa a esta misma ubicación, y
generará el reporte de la quincena de calendario actual)
"""
import os
import sqlite3
import calendar
import logging
from datetime import date, timedelta
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, HRFlowable
)
from reportlab.lib.enums import TA_LEFT

AZUL = "#1F4E79"
AZUL_CLARO = "#2E75B6"
GRIS = "#8c8c8c"
COLORES_CAT = {"Estudiantes secundaria": "#1F4E79", "Estudiantes primaria": "#2E75B6", "Docentes": "#8FAADC"}

MINIMO_DIAS_CON_DATOS = 5  # con menos días que esto en la quincena, el reporte no dice mucho


# ============================================================
# Quincenas de CALENDARIO (1-15 y 16-fin de mes)
# ============================================================

def identificar_quincena(fecha):
    """Dado un date, devuelve (identificador, fecha_inicio, fecha_fin) de la
    quincena de calendario a la que pertenece:
      - Días 1 a 15  -> primera quincena del mes.
      - Día 16 en adelante -> segunda quincena (hasta el último día del mes,
        sin importar si el mes tiene 28, 30 o 31 días).
    El identificador (ej. '2026-08-Q1') sirve para comparar fácilmente si
    ya cambiamos de quincena desde la última vez."""
    if fecha.day <= 15:
        inicio = fecha.replace(day=1)
        fin = fecha.replace(day=15)
        numero = 1
    else:
        inicio = fecha.replace(day=16)
        ultimo_dia_mes = calendar.monthrange(fecha.year, fecha.month)[1]
        fin = fecha.replace(day=ultimo_dia_mes)
        numero = 2
    identificador = f"{fecha.year}-{fecha.month:02d}-Q{numero}"
    return identificador, inicio, fin


def rango_desde_identificador(identificador):
    """Operación inversa: a partir de un identificador como '2026-08-Q1',
    reconstruye (fecha_inicio, fecha_fin) de esa quincena."""
    año_txt, mes_txt, q_txt = identificador.split("-")
    año, mes = int(año_txt), int(mes_txt)
    numero = int(q_txt.replace("Q", ""))
    if numero == 1:
        inicio = date(año, mes, 1)
        fin = date(año, mes, 15)
    else:
        inicio = date(año, mes, 16)
        ultimo_dia_mes = calendar.monthrange(año, mes)[1]
        fin = date(año, mes, ultimo_dia_mes)
    return inicio, fin


# ---------------- Consultas (todas aceptan un rango de fechas opcional) ----------------

def _conectar(db_path):
    return sqlite3.connect(db_path)


def _tendencia_diaria(db_path, rango_fechas):
    conexion = _conectar(db_path)
    df = pd.read_sql("""
        SELECT e.fecha,
               COUNT(DISTINCT e.persona_id) * 1.0 / (SELECT COUNT(*) FROM personas) AS porcentaje
        FROM entradas e
        WHERE e.fecha BETWEEN ? AND ?
        GROUP BY e.fecha
        ORDER BY e.fecha
    """, conexion, params=[rango_fechas[0].isoformat(), rango_fechas[1].isoformat()])
    conexion.close()
    df['fecha'] = pd.to_datetime(df['fecha'])
    return df


def _estadisticas_del_rango(db_path, rango_fechas):
    conexion = _conectar(db_path)
    df = pd.read_sql(
        "SELECT fecha, categoria, porcentaje FROM estadisticas_diarias WHERE fecha BETWEEN ? AND ?",
        conexion, params=[rango_fechas[0].isoformat(), rango_fechas[1].isoformat()]
    )
    conexion.close()
    df['fecha'] = pd.to_datetime(df['fecha'])
    return df


def _promedio_por_dia_semana(db_path, rango_fechas):
    df = _estadisticas_del_rango(db_path, rango_fechas)
    dias_es = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
    df['dia_semana'] = df['fecha'].dt.weekday.map(dict(enumerate(dias_es)))
    return df.groupby('dia_semana')['porcentaje'].mean().reindex(dias_es)


def _promedio_por_categoria(db_path, rango_fechas):
    df = _estadisticas_del_rango(db_path, rango_fechas)
    return df.groupby('categoria')['porcentaje'].mean()


def _total_por_categoria(db_path):
    conexion = _conectar(db_path)
    df = pd.read_sql("SELECT categoria, COUNT(*) as total FROM personas GROUP BY categoria", conexion)
    conexion.close()
    return df.set_index('categoria')['total']


def _promedio_por_dia_y_categoria(db_path, rango_fechas):
    df = _estadisticas_del_rango(db_path, rango_fechas)
    dias_es = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
    df['dia_semana'] = df['fecha'].dt.weekday.map(dict(enumerate(dias_es)))
    tabla = df.pivot_table(index='dia_semana', columns='categoria', values='porcentaje', aggfunc='mean')
    return tabla.reindex(dias_es)


def _dias_con_datos(db_path, rango_fechas):
    conexion = _conectar(db_path)
    total = conexion.execute(
        "SELECT COUNT(DISTINCT fecha) FROM estadisticas_diarias WHERE fecha BETWEEN ? AND ?",
        [rango_fechas[0].isoformat(), rango_fechas[1].isoformat()]
    ).fetchone()[0]
    conexion.close()
    return total


# ---------------- Gráficas ----------------

def _generar_graficas(db_path, carpeta_graficas, rango_fechas):
    plt.rcParams.update({
        "font.size": 11, "axes.spines.top": False, "axes.spines.right": False, "axes.edgecolor": "#888888",
    })

    df_tendencia = _tendencia_diaria(db_path, rango_fechas)
    fig, ax = plt.subplots(figsize=(7.8, 3.2))
    ax.plot(df_tendencia['fecha'], df_tendencia['porcentaje'] * 100, color=AZUL, linewidth=2, marker='o', markersize=4)
    ax.axhline(df_tendencia['porcentaje'].mean() * 100, color=GRIS, linestyle='--', linewidth=1, label="Promedio de la quincena")
    ax.set_ylabel("% de asistencia")
    ax.set_ylim(0, 100)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%b'))
    fig.autofmt_xdate(rotation=30, ha='right')
    ax.legend(frameon=False, loc="lower left", fontsize=9)
    ax.set_title("Tendencia de asistencia por día (días con clases/PAE)", color=AZUL, fontsize=12, fontweight='bold', loc='left')
    fig.tight_layout()
    fig.savefig(os.path.join(carpeta_graficas, "grafica_tendencia.png"), dpi=170)
    plt.close(fig)

    prom_dia = _promedio_por_dia_semana(db_path, rango_fechas)
    fig, ax = plt.subplots(figsize=(7.8, 3.2))
    barras = ax.bar(prom_dia.index, prom_dia.values * 100, color=AZUL_CLARO, width=0.55)
    for b in barras:
        altura = b.get_height()
        if pd.notna(altura):
            ax.text(b.get_x() + b.get_width()/2, altura + 1, f"{altura:.0f}%", ha='center', fontsize=10, color=AZUL)
    ax.set_ylabel("% de asistencia promedio")
    ax.set_ylim(0, 100)
    ax.set_title("Asistencia promedio por día de la semana (esta quincena)", color=AZUL, fontsize=12, fontweight='bold', loc='left')
    fig.tight_layout()
    fig.savefig(os.path.join(carpeta_graficas, "grafica_dia_semana.png"), dpi=170)
    plt.close(fig)

    prom_cat = _promedio_por_categoria(db_path, rango_fechas).sort_values()
    fig, ax = plt.subplots(figsize=(7.8, 2.6))
    colores = [COLORES_CAT.get(c, AZUL) for c in prom_cat.index]
    barras = ax.barh(prom_cat.index, prom_cat.values * 100, color=colores, height=0.5)
    for b in barras:
        ax.text(b.get_width() + 1, b.get_y() + b.get_height()/2, f"{b.get_width():.0f}%", va='center', fontsize=10, color=AZUL)
    ax.set_xlabel("% de asistencia promedio")
    ax.set_xlim(0, 100)
    ax.set_title("Asistencia promedio por categoría (esta quincena)", color=AZUL, fontsize=12, fontweight='bold', loc='left')
    fig.tight_layout()
    fig.savefig(os.path.join(carpeta_graficas, "grafica_categoria.png"), dpi=170)
    plt.close(fig)

    tabla = _promedio_por_dia_y_categoria(db_path, rango_fechas)
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
    fig.savefig(os.path.join(carpeta_graficas, "grafica_dia_categoria.png"), dpi=170)
    plt.close(fig)


# ---------------- Documento PDF ----------------

def _construir_pdf(db_path, carpeta_graficas, salida_pdf, etiqueta_periodo, rango_fechas):
    AZUL_C = colors.HexColor(AZUL)
    AZUL_CLARO_C = colors.HexColor(AZUL_CLARO)
    GRIS_C = colors.HexColor("#595959")
    FONDO_TABLA = colors.HexColor("#EEF3F8")

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloReporte", fontSize=22, textColor=AZUL_C, fontName="Helvetica-Bold", spaceAfter=6, leading=26))
    styles.add(ParagraphStyle(name="Subtitulo", fontSize=12, textColor=GRIS_C, fontName="Helvetica", spaceAfter=18))
    styles.add(ParagraphStyle(name="H1", fontSize=15, textColor=AZUL_C, fontName="Helvetica-Bold", spaceBefore=18, spaceAfter=8))
    styles.add(ParagraphStyle(name="Cuerpo", fontSize=10.3, textColor=colors.HexColor("#333333"), fontName="Helvetica", leading=15, spaceAfter=8, alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="Nota", fontSize=9, textColor=GRIS_C, fontName="Helvetica-Oblique", leading=13))

    def img(nombre, alto_relativo):
        return Image(os.path.join(carpeta_graficas, nombre), width=6.9*inch, height=6.9*inch*alto_relativo)

    story = []
    story.append(Paragraph("Control de Entradas PAE", ParagraphStyle(name="marca", fontSize=11, textColor=AZUL_CLARO_C, fontName="Helvetica-Bold")))
    story.append(Paragraph("Reporte quincenal de asistencia y estimación de porciones", styles["TituloReporte"]))
    story.append(Paragraph(f"Quincena: {etiqueta_periodo}", styles["Subtitulo"]))
    story.append(HRFlowable(width="100%", thickness=1, color=AZUL_CLARO_C))
    story.append(Spacer(1, 10))

    story.append(Paragraph("1. Tendencia de asistencia en la quincena", styles["H1"]))
    story.append(Paragraph(
        "Porcentaje de asistencia diario dentro de esta quincena (estudiantes de secundaria, primaria y "
        "docentes juntos), considerando únicamente días en los que hubo clases/PAE — los días sin ninguna "
        "entrada registrada (festivos u otros días sin actividad) se excluyen automáticamente para no "
        "distorsionar los promedios.", styles["Cuerpo"]))
    story.append(img("grafica_tendencia.png", 3.2/7.8))
    story.append(Spacer(1, 8))

    df_tend = _tendencia_diaria(db_path, rango_fechas)
    promedio_general = df_tend['porcentaje'].mean()
    minimo = df_tend.loc[df_tend['porcentaje'].idxmin()]
    maximo = df_tend.loc[df_tend['porcentaje'].idxmax()]
    story.append(Paragraph(
        f"En esta quincena, la asistencia general promedió <b>{promedio_general*100:.1f}%</b>. "
        f"El día más bajo fue el <b>{minimo['fecha'].strftime('%d de %B')}</b> ({minimo['porcentaje']*100:.1f}%) "
        f"y el más alto el <b>{maximo['fecha'].strftime('%d de %B')}</b> ({maximo['porcentaje']*100:.1f}%).",
        styles["Cuerpo"]))
    story.append(PageBreak())

    story.append(Paragraph("2. Asistencia promedio por día de la semana", styles["H1"]))
    story.append(Paragraph(
        "La gráfica más relevante para planear la cantidad de alimentos de la próxima quincena: qué días "
        "de la semana tuvieron más o menos asistencia en esta.", styles["Cuerpo"]))
    story.append(img("grafica_dia_semana.png", 3.2/7.8))
    story.append(Spacer(1, 8))

    prom_dia = _promedio_por_dia_semana(db_path, rango_fechas)
    dia_mas_alto = prom_dia.idxmax()
    dia_mas_bajo = prom_dia.idxmin()
    story.append(Paragraph(
        f"El día con más asistencia en promedio fue <b>{dia_mas_alto}</b> ({prom_dia[dia_mas_alto]*100:.1f}%) "
        f"y el más bajo <b>{dia_mas_bajo}</b> ({prom_dia[dia_mas_bajo]*100:.1f}%).",
        styles["Cuerpo"]))

    story.append(Spacer(1, 10))
    story.append(Paragraph("3. Asistencia por categoría", styles["H1"]))
    story.append(img("grafica_categoria.png", 2.6/7.8))
    story.append(PageBreak())

    story.append(Paragraph("4. Desglose: día de la semana por categoría", styles["H1"]))
    story.append(img("grafica_dia_categoria.png", 3.4/7.8))
    story.append(PageBreak())

    story.append(Paragraph("5. Estimación de porciones sugeridas por día", styles["H1"]))
    story.append(Paragraph(
        "Traduce el % de asistencia promedio de cada día (en esta quincena) en un número estimado de "
        "personas presentes, asumiendo una porción por persona presente — un punto de partida para la "
        "próxima quincena, junto con la entrega de recursos.", styles["Cuerpo"]))

    tabla_dia_cat = _promedio_por_dia_y_categoria(db_path, rango_fechas)
    totales_cat = _total_por_categoria(db_path)
    total_roster = totales_cat.sum()

    encabezados = ["Día", "Est. secundaria", "Est. primaria", "Docentes", "Total estimado"]
    filas = [encabezados]
    for dia in tabla_dia_cat.index:
        fila = [dia]
        total_dia = 0
        for cat in ["Estudiantes secundaria", "Estudiantes primaria", "Docentes"]:
            pct = tabla_dia_cat.loc[dia, cat] if cat in tabla_dia_cat.columns and pd.notna(tabla_dia_cat.loc[dia, cat]) else 0
            personas_est = round(pct * totales_cat.get(cat, 0))
            total_dia += personas_est
            fila.append(str(personas_est))
        fila.append(str(total_dia))
        filas.append(fila)

    fila_prom = ["Promedio quincena"]
    total_prom = 0
    for cat in ["Estudiantes secundaria", "Estudiantes primaria", "Docentes"]:
        pct_prom = tabla_dia_cat[cat].mean() if cat in tabla_dia_cat.columns else 0
        personas_est = round((pct_prom or 0) * totales_cat.get(cat, 0))
        total_prom += personas_est
        fila_prom.append(str(personas_est))
    fila_prom.append(str(total_prom))
    filas.append(fila_prom)

    tabla = Table(filas, colWidths=[1.3*inch, 1.4*inch, 1.3*inch, 1.1*inch, 1.4*inch])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL_C),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("BACKGROUND", (0, -1), (-1, -1), FONDO_TABLA),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F7FAFD")]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#C7D6E5")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(Spacer(1, 6))
    story.append(tabla)
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        f"Roster total del plantel: {int(total_roster)} personas "
        f"({int(totales_cat.get('Estudiantes secundaria', 0))} secundaria, "
        f"{int(totales_cat.get('Estudiantes primaria', 0))} primaria, "
        f"{int(totales_cat.get('Docentes', 0))} docentes).", styles["Nota"]))
    story.append(Paragraph(
        "Los días festivos o sin clases/PAE (0 entradas en todas las categorías) se excluyen "
        "automáticamente de todos los promedios de este reporte.", styles["Nota"]))

    doc = SimpleDocTemplate(
        salida_pdf, pagesize=letter,
        topMargin=0.75*inch, bottomMargin=0.75*inch, leftMargin=0.8*inch, rightMargin=0.8*inch,
    )
    doc.build(story)


# ---------------- Función pública ----------------

def generar_reporte_quincenal(db_path, carpeta_salida, rango_fechas, etiqueta_periodo="—"):
    """Genera el PDF del reporte para la quincena delimitada por
    'rango_fechas' = (fecha_inicio, fecha_fin), ambos objetos date.

    Devuelve la ruta del PDF generado, o None si no hay suficientes datos
    todavía dentro de ese rango (en cuyo caso no genera nada)."""
    if not os.path.exists(db_path):
        logging.warning(f"No se generó el reporte quincenal: no existe '{db_path}' todavía.")
        return None

    dias_con_datos = _dias_con_datos(db_path, rango_fechas)
    if dias_con_datos < MINIMO_DIAS_CON_DATOS:
        logging.info(
            f"No se generó el reporte de la quincena {etiqueta_periodo}: solo {dias_con_datos} día(s) "
            f"con datos (mínimo {MINIMO_DIAS_CON_DATOS}). No se genera reporte para este periodo."
        )
        return None

    os.makedirs(carpeta_salida, exist_ok=True)
    _generar_graficas(db_path, carpeta_salida, rango_fechas)

    nombre_archivo = f"reporte_asistencia_{rango_fechas[0].isoformat()}_a_{rango_fechas[1].isoformat()}.pdf"
    salida_pdf = os.path.join(carpeta_salida, nombre_archivo)
    _construir_pdf(db_path, carpeta_salida, salida_pdf, etiqueta_periodo, rango_fechas)
    return salida_pdf


if __name__ == "__main__":
    # Uso manual: genera el reporte de la quincena de calendario ACTUAL
    # (la que está corriendo hoy), útil para probar sin esperar a que cierre.
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    db_path_manual = os.path.join(BASE_DIR, 'datos', 'asistencia.db')
    carpeta_manual = os.path.join(BASE_DIR, 'reportes')
    _, inicio, fin = identificar_quincena(date.today())
    etiqueta = f"{inicio.strftime('%d %b %Y')} al {fin.strftime('%d %b %Y')}"
    ruta = generar_reporte_quincenal(db_path_manual, carpeta_manual, (inicio, fin), etiqueta)
    print(f"Reporte generado en: {ruta}" if ruta else "No se generó reporte (datos insuficientes).")
