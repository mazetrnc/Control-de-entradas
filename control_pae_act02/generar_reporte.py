"""
Genera el reporte de asistencia en PDF: llama primero a graficas.py (para
tener las imágenes al día) y arma el documento final con reportlab.

Uso normal:  python generar_reporte.py
El PDF queda en la carpeta 'reportes/', junto a este script.
"""
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, HRFlowable
)
from reportlab.lib.enums import TA_LEFT

import graficas
from consultas import (
    tendencia_diaria, promedio_por_dia_semana, promedio_por_categoria,
    total_por_categoria, promedio_por_dia_y_categoria, USAR_DATOS_SIMULADOS
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARPETA_REPORTES = os.path.join(BASE_DIR, 'reportes')
os.makedirs(CARPETA_REPORTES, exist_ok=True)


def ruta_imagen(nombre):
    return os.path.join(CARPETA_REPORTES, nombre)


def generar():
    # 1) Asegura que las gráficas estén al día antes de armar el PDF
    graficas.generar_todas()

    AZUL = colors.HexColor("#1F4E79")
    AZUL_CLARO = colors.HexColor("#2E75B6")
    GRIS = colors.HexColor("#595959")
    FONDO_TABLA = colors.HexColor("#EEF3F8")

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloReporte", fontSize=22, textColor=AZUL, fontName="Helvetica-Bold", spaceAfter=6, leading=26))
    styles.add(ParagraphStyle(name="Subtitulo", fontSize=12, textColor=GRIS, fontName="Helvetica", spaceAfter=18))
    styles.add(ParagraphStyle(name="H1", fontSize=15, textColor=AZUL, fontName="Helvetica-Bold", spaceBefore=18, spaceAfter=8))
    styles.add(ParagraphStyle(name="Cuerpo", fontSize=10.3, textColor=colors.HexColor("#333333"), fontName="Helvetica", leading=15, spaceAfter=8, alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="Nota", fontSize=9, textColor=GRIS, fontName="Helvetica-Oblique", leading=13))

    story = []

    story.append(Paragraph("Control de Entradas PAE", ParagraphStyle(name="marca", fontSize=11, textColor=AZUL_CLARO, fontName="Helvetica-Bold")))
    story.append(Paragraph("Reporte de asistencia y estimación de porciones", styles["TituloReporte"]))

    df_tend_preview = tendencia_diaria()
    rango = f"{df_tend_preview['fecha'].min().strftime('%d %b %Y')} al {df_tend_preview['fecha'].max().strftime('%d %b %Y')}"
    etiqueta_fuente = "datos SIMULADOS (prueba)" if USAR_DATOS_SIMULADOS else "datos reales"
    story.append(Paragraph(f"Periodo: {rango} — fuente: {etiqueta_fuente}", styles["Subtitulo"]))

    if USAR_DATOS_SIMULADOS:
        aviso = Table([[Paragraph(
            "<b>Este reporte usa datos simulados</b>, generados únicamente para probar el diseño y las "
            "consultas antes de la prueba real con cédulas. Los números aquí NO reflejan asistencia real "
            "y no deben usarse para tomar decisiones. Cambia USAR_DATOS_SIMULADOS a False en consultas.py "
            "cuando quieras generar este mismo reporte con datos reales.", styles["Cuerpo"])]],
            colWidths=[6.9*inch])
        aviso.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFF8E1")),
            ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#C9A227")),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ]))
        story.append(aviso)

    story.append(Spacer(1, 14))
    story.append(HRFlowable(width="100%", thickness=1, color=AZUL_CLARO))
    story.append(Spacer(1, 6))

    # ---------------- 1. Tendencia general ----------------
    story.append(Paragraph("1. Tendencia general de asistencia", styles["H1"]))
    story.append(Paragraph(
        "Porcentaje de asistencia diario, considerando las tres categorías juntas (estudiantes de "
        "secundaria, primaria y docentes). Sirve para detectar si la asistencia general está subiendo, "
        "bajando, o si hay días atípicos que valga la pena investigar.", styles["Cuerpo"]))
    story.append(Image(ruta_imagen("grafica_tendencia.png"), width=6.9*inch, height=6.9*inch*3.2/7.8))
    story.append(Spacer(1, 8))

    df_tend = tendencia_diaria()
    promedio_general = df_tend['porcentaje'].mean()
    minimo = df_tend.loc[df_tend['porcentaje'].idxmin()]
    maximo = df_tend.loc[df_tend['porcentaje'].idxmax()]
    story.append(Paragraph(
        f"En este periodo, la asistencia general promedió <b>{promedio_general*100:.1f}%</b>. "
        f"El día más bajo fue el <b>{minimo['fecha'].strftime('%d de %B')}</b> ({minimo['porcentaje']*100:.1f}%) "
        f"y el más alto el <b>{maximo['fecha'].strftime('%d de %B')}</b> ({maximo['porcentaje']*100:.1f}%).",
        styles["Cuerpo"]))

    story.append(PageBreak())

    # ---------------- 2. Por día de la semana ----------------
    story.append(Paragraph("2. Asistencia promedio por día de la semana", styles["H1"]))
    story.append(Paragraph(
        "Esta es la gráfica más relevante para planear la cantidad de alimentos: muestra qué días de la "
        "semana tienden a tener más o menos asistencia, promediando el periodo.", styles["Cuerpo"]))
    story.append(Image(ruta_imagen("grafica_dia_semana.png"), width=6.9*inch, height=6.9*inch*3.2/7.8))
    story.append(Spacer(1, 8))

    prom_dia = promedio_por_dia_semana()
    dia_mas_alto = prom_dia.idxmax()
    dia_mas_bajo = prom_dia.idxmin()
    story.append(Paragraph(
        f"El día con más asistencia en promedio es <b>{dia_mas_alto}</b> ({prom_dia[dia_mas_alto]*100:.1f}%) "
        f"y el más bajo es <b>{dia_mas_bajo}</b> ({prom_dia[dia_mas_bajo]*100:.1f}%). "
        f"Un patrón así sugiere ajustar la cantidad de comida preparada según el día para reducir sobrantes.",
        styles["Cuerpo"]))

    story.append(Spacer(1, 10))
    story.append(Paragraph("3. Asistencia por categoría", styles["H1"]))
    story.append(Paragraph(
        "Comparación del promedio de asistencia entre estudiantes de secundaria, primaria y docentes "
        "durante todo el periodo.", styles["Cuerpo"]))
    story.append(Image(ruta_imagen("grafica_categoria.png"), width=6.9*inch, height=6.9*inch*2.6/7.8))

    story.append(PageBreak())

    story.append(Paragraph("4. Desglose: día de la semana por categoría", styles["H1"]))
    story.append(Paragraph(
        "Combina las dos vistas anteriores: cómo varía la asistencia de cada categoría a lo largo de la "
        "semana. Es útil si el comedor prepara porciones diferenciadas por categoría.", styles["Cuerpo"]))
    story.append(Image(ruta_imagen("grafica_dia_categoria.png"), width=6.9*inch, height=6.9*inch*3.4/7.8))

    story.append(PageBreak())

    # ---------------- 5. Estimacion de porciones ----------------
    story.append(Paragraph("5. Estimación de porciones sugeridas por día", styles["H1"]))
    story.append(Paragraph(
        "Traduce el % de asistencia promedio de cada día en un número estimado de personas presentes, "
        "asumiendo <b>una porción por persona presente</b>. Es un punto de partida para ajustar cuánta "
        "comida preparar cada día — no un número exacto, ya que la asistencia real varía día a día.",
        styles["Cuerpo"]))

    tabla_dia_cat = promedio_por_dia_y_categoria()
    totales_cat = total_por_categoria()

    encabezados = ["Día", "Est. secundaria", "Est. primaria", "Docentes", "Total estimado"]
    filas = [encabezados]
    total_roster = totales_cat.sum()
    for dia in tabla_dia_cat.index:
        fila = [dia]
        total_dia = 0
        for cat in ["Estudiantes secundaria", "Estudiantes primaria", "Docentes"]:
            pct = tabla_dia_cat.loc[dia, cat] if cat in tabla_dia_cat.columns else 0
            personas_est = round(pct * totales_cat.get(cat, 0))
            total_dia += personas_est
            fila.append(str(personas_est))
        fila.append(str(total_dia))
        filas.append(fila)

    fila_prom = ["Promedio semanal"]
    total_prom = 0
    for cat in ["Estudiantes secundaria", "Estudiantes primaria", "Docentes"]:
        pct_prom = tabla_dia_cat[cat].mean() if cat in tabla_dia_cat.columns else 0
        personas_est = round(pct_prom * totales_cat.get(cat, 0))
        total_prom += personas_est
        fila_prom.append(str(personas_est))
    fila_prom.append(str(total_prom))
    filas.append(fila_prom)

    tabla = Table(filas, colWidths=[1.3*inch, 1.4*inch, 1.3*inch, 1.1*inch, 1.4*inch])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
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

    story.append(Spacer(1, 16))
    story.append(Paragraph("6. Metodología y siguiente paso", styles["H1"]))
    story.append(Paragraph(
        "Este reporte se genera automáticamente a partir de la base de datos SQLite (asistencia.db) que "
        "llena el sistema de control de entradas en cada escaneo. Para volver a generarlo con datos "
        "reales, cambia USAR_DATOS_SIMULADOS a False en consultas.py y vuelve a correr este script.",
        styles["Cuerpo"]))
    story.append(Paragraph(
        "Nota: la categoría \"Docentes\" suele tener una muestra más pequeña que las de estudiantes, por "
        "lo que sus porcentajes pueden variar más de un periodo a otro — esto es estadísticamente "
        "esperado, no un error del sistema.", styles["Nota"]))

    nombre_salida = "reporte_asistencia_SIMULACION.pdf" if USAR_DATOS_SIMULADOS else "reporte_asistencia.pdf"
    salida = os.path.join(CARPETA_REPORTES, nombre_salida)
    doc = SimpleDocTemplate(
        salida, pagesize=letter,
        topMargin=0.75*inch, bottomMargin=0.75*inch, leftMargin=0.8*inch, rightMargin=0.8*inch,
    )
    doc.build(story)
    print(f"Reporte generado en: {salida}")


if __name__ == "__main__":
    generar()