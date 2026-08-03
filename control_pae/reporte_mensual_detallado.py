"""
Genera el reporte MENSUAL DETALLADO POR ESTUDIANTE: a diferencia del
reporte quincenal (que resume tendencias generales para prever comida),
este reporte lista a CADA persona con su % de asistencia del mes y el
cambio contra el mes anterior -- pensado para identificar a quienes
podrían ceder su cupo de alimentación por asistencia sostenidamente baja.

Debe vivir en la MISMA carpeta que entradas_pae26.py.
"""
import os
import sqlite3
import calendar
import logging
from datetime import date
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable
)
from reportlab.lib.enums import TA_LEFT

AZUL = "#1F4E79"
AZUL_CLARO = "#2E75B6"
ROJO_CANDIDATO = "#C0392B"
FONDO_CANDIDATO = "#FDECEA"

# Debajo de este % de asistencia en el mes, una persona se marca como
# "candidata" a ceder su cupo de alimentación. Ajusta este número según el
# criterio real de la institución.
UMBRAL_CANDIDATO = 0.50

# Mínimo de días con clase/PAE en el mes para generar el reporte (si hay
# menos, probablemente el mes apenas empezó o hubo muy pocos días activos).
MINIMO_DIAS_CON_DATOS = 5

CATEGORIAS_ORDEN = ["Estudiantes secundaria", "Estudiantes primaria", "Docentes"]


# ============================================================
# Meses de calendario
# ============================================================

def identificar_mes(fecha):
    identificador = f"{fecha.year}-{fecha.month:02d}"
    inicio = fecha.replace(day=1)
    ultimo_dia = calendar.monthrange(fecha.year, fecha.month)[1]
    fin = fecha.replace(day=ultimo_dia)
    return identificador, inicio, fin


def rango_desde_identificador_mes(identificador):
    año, mes = (int(x) for x in identificador.split("-"))
    inicio = date(año, mes, 1)
    ultimo_dia = calendar.monthrange(año, mes)[1]
    fin = date(año, mes, ultimo_dia)
    return inicio, fin


def mes_anterior(identificador):
    año, mes = (int(x) for x in identificador.split("-"))
    if mes == 1:
        return f"{año - 1}-12"
    return f"{año}-{mes - 1:02d}"


# ============================================================
# Consultas
# ============================================================

def _asistencia_por_estudiante(db_path, rango_fechas):
    conexion = sqlite3.connect(db_path)
    df = pd.read_sql("""
        SELECT p.id, p.nombre, p.apellido, p.grado, p.categoria,
               COUNT(e.id) AS dias_asistidos
        FROM personas p
        LEFT JOIN entradas e
            ON e.persona_id = p.id AND e.fecha BETWEEN ? AND ?
        GROUP BY p.id
    """, conexion, params=[rango_fechas[0].isoformat(), rango_fechas[1].isoformat()])

    dias_por_categoria = pd.read_sql("""
        SELECT categoria, COUNT(DISTINCT fecha) AS dias_con_clase
        FROM estadisticas_diarias
        WHERE fecha BETWEEN ? AND ?
        GROUP BY categoria
    """, conexion, params=[rango_fechas[0].isoformat(), rango_fechas[1].isoformat()])
    conexion.close()

    df = df.merge(dias_por_categoria, on="categoria", how="left")
    df["dias_con_clase"] = df["dias_con_clase"].fillna(0)
    df["porcentaje"] = df.apply(
        lambda r: (r["dias_asistidos"] / r["dias_con_clase"]) if r["dias_con_clase"] else None, axis=1
    )
    return df


def _dias_con_datos_mes(db_path, rango_fechas):
    conexion = sqlite3.connect(db_path)
    total = conexion.execute(
        "SELECT COUNT(DISTINCT fecha) FROM estadisticas_diarias WHERE fecha BETWEEN ? AND ?",
        [rango_fechas[0].isoformat(), rango_fechas[1].isoformat()]
    ).fetchone()[0]
    conexion.close()
    return total


def _tabla_con_tendencia(db_path, id_mes_actual):
    inicio_actual, fin_actual = rango_desde_identificador_mes(id_mes_actual)
    actual = _asistencia_por_estudiante(db_path, (inicio_actual, fin_actual))

    id_anterior = mes_anterior(id_mes_actual)
    inicio_ant, fin_ant = rango_desde_identificador_mes(id_anterior)
    anterior = _asistencia_por_estudiante(db_path, (inicio_ant, fin_ant))
    anterior = anterior[["id", "porcentaje"]].rename(columns={"porcentaje": "porcentaje_anterior"})

    tabla = actual.merge(anterior, on="id", how="left")
    tabla["cambio"] = tabla["porcentaje"] - tabla["porcentaje_anterior"]
    tabla["candidato"] = tabla["porcentaje"].apply(
        lambda p: "Sí" if pd.notna(p) and p < UMBRAL_CANDIDATO else ""
    )
    return tabla


# ============================================================
# Documento PDF
# ============================================================

def _fmt_pct(valor):
    return f"{valor * 100:.0f}%" if pd.notna(valor) else "—"


def _fmt_cambio(valor):
    if pd.isna(valor):
        return "—"
    signo = "+" if valor >= 0 else ""
    return f"{signo}{valor * 100:.0f} pts"


def _tabla_estudiantes(filas, styles):
    encabezados = ["ID", "Nombre", "Apellido", "Grado", "% mes", "% mes ant.", "Cambio", "Candidato"]
    datos = [encabezados]
    estilos_extra = []
    for i, fila in enumerate(filas, start=1):
        datos.append([
            str(fila["id"]), fila["nombre"], fila["apellido"], fila["grado"],
            _fmt_pct(fila["porcentaje"]), _fmt_pct(fila["porcentaje_anterior"]),
            _fmt_cambio(fila["cambio"]), fila["candidato"],
        ])
        if fila["candidato"] == "Sí":
            estilos_extra.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor(FONDO_CANDIDATO)))

    anchos = [0.85*inch, 1.25*inch, 1.25*inch, 0.6*inch, 0.65*inch, 0.75*inch, 0.65*inch, 0.7*inch]
    tabla = Table(datos, colWidths=anchos, repeatRows=1)
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(AZUL)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (4, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#C7D6E5")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ] + estilos_extra
    tabla.setStyle(TableStyle(estilo))
    return tabla


def _construir_pdf(db_path, salida_pdf, etiqueta_periodo, id_mes):
    AZUL_C = colors.HexColor(AZUL)
    AZUL_CLARO_C = colors.HexColor(AZUL_CLARO)
    GRIS_C = colors.HexColor("#595959")

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloReporte", fontSize=20, textColor=AZUL_C, fontName="Helvetica-Bold", spaceAfter=6, leading=24))
    styles.add(ParagraphStyle(name="Subtitulo", fontSize=11, textColor=GRIS_C, fontName="Helvetica", spaceAfter=14))
    styles.add(ParagraphStyle(name="H1", fontSize=14, textColor=AZUL_C, fontName="Helvetica-Bold", spaceBefore=14, spaceAfter=6))
    styles.add(ParagraphStyle(name="Cuerpo", fontSize=9.7, textColor=colors.HexColor("#333333"), fontName="Helvetica", leading=14, spaceAfter=8))
    styles.add(ParagraphStyle(name="Nota", fontSize=8.5, textColor=GRIS_C, fontName="Helvetica-Oblique", leading=12))

    tabla_completa = _tabla_con_tendencia(db_path, id_mes)

    story = []
    story.append(Paragraph("Control de Entradas PAE", ParagraphStyle(name="marca", fontSize=10, textColor=AZUL_CLARO_C, fontName="Helvetica-Bold")))
    story.append(Paragraph("Reporte mensual detallado por estudiante", styles["TituloReporte"]))
    story.append(Paragraph(f"Mes: {etiqueta_periodo}", styles["Subtitulo"]))
    story.append(HRFlowable(width="100%", thickness=1, color=AZUL_CLARO_C))
    story.append(Spacer(1, 8))

    story.append(Paragraph(
        f"Este reporte muestra el % de asistencia de CADA persona durante el mes, comparado contra el mes "
        f"anterior. Quienes tuvieron menos de {int(UMBRAL_CANDIDATO*100)}% de asistencia quedan marcados "
        f"como <b>Candidato</b> en la columna final y resaltados en la tabla — son quienes, por asistencia "
        f"sostenidamente baja, podrían ceder su cupo de alimentación a alguien más.", styles["Cuerpo"]))

    # ---- Resumen: candidatos, todas las categorías juntas, para revisión rápida ----
    candidatos = tabla_completa[tabla_completa["candidato"] == "Sí"].sort_values("porcentaje")
    story.append(Paragraph(f"Resumen: {len(candidatos)} candidato(s) a liberar su cupo este mes", styles["H1"]))
    if candidatos.empty:
        story.append(Paragraph("Nadie estuvo por debajo del umbral este mes.", styles["Cuerpo"]))
    else:
        story.append(_tabla_estudiantes(candidatos.to_dict("records"), styles))
    story.append(PageBreak())

    # ---- Detalle completo, por categoría ----
    for categoria in CATEGORIAS_ORDEN:
        sub = tabla_completa[tabla_completa["categoria"] == categoria].sort_values(
            "porcentaje", na_position="last"
        )
        if sub.empty:
            continue
        story.append(Paragraph(f"Detalle completo — {categoria}", styles["H1"]))
        story.append(Paragraph(
            f"{len(sub)} persona(s), ordenadas de menor a mayor asistencia.", styles["Nota"]))
        story.append(Spacer(1, 4))
        story.append(_tabla_estudiantes(sub.to_dict("records"), styles))
        story.append(PageBreak())

    story.append(Paragraph(
        "Los días festivos o sin clases/PAE se excluyen automáticamente del cálculo (no cuentan ni como "
        "día con clase ni como inasistencia). 'Cambio' compara el % de este mes contra el % del mes "
        "anterior; '—' significa que no hay suficiente información del mes anterior para comparar.",
        styles["Nota"]))

    doc = SimpleDocTemplate(
        salida_pdf, pagesize=letter,
        topMargin=0.7*inch, bottomMargin=0.7*inch, leftMargin=0.7*inch, rightMargin=0.7*inch,
    )
    doc.build(story)


# ============================================================
# Función pública
# ============================================================

def generar_reporte_mensual_detallado(db_path, carpeta_salida, id_mes, etiqueta_periodo="—"):
    """Genera el PDF del reporte mensual detallado para el mes 'id_mes'
    (formato 'YYYY-MM'). Devuelve la ruta del PDF, o None si no hay
    suficientes datos todavía en ese mes."""
    if not os.path.exists(db_path):
        logging.warning(f"No se generó el reporte mensual detallado: no existe '{db_path}' todavía.")
        return None

    inicio, fin = rango_desde_identificador_mes(id_mes)
    dias_con_datos = _dias_con_datos_mes(db_path, (inicio, fin))
    if dias_con_datos < MINIMO_DIAS_CON_DATOS:
        logging.info(
            f"No se generó el reporte mensual detallado de {id_mes}: solo {dias_con_datos} día(s) con "
            f"datos (mínimo {MINIMO_DIAS_CON_DATOS})."
        )
        return None

    os.makedirs(carpeta_salida, exist_ok=True)
    nombre_archivo = f"reporte_mensual_detallado_{id_mes}.pdf"
    salida_pdf = os.path.join(carpeta_salida, nombre_archivo)
    _construir_pdf(db_path, salida_pdf, etiqueta_periodo, id_mes)
    return salida_pdf


if __name__ == "__main__":
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    db_path_manual = os.path.join(BASE_DIR, 'datos', 'asistencia.db')
    carpeta_manual = os.path.join(BASE_DIR, 'reportes')
    id_mes_actual, inicio, fin = identificar_mes(date.today())
    etiqueta = f"{inicio.strftime('%B %Y')}"
    ruta = generar_reporte_mensual_detallado(db_path_manual, carpeta_manual, id_mes_actual, etiqueta)
    print(f"Reporte generado en: {ruta}" if ruta else "No se generó reporte (datos insuficientes).")