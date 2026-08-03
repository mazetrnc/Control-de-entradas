"""
Genera datos de asistencia SIMULADOS (no reales) en una base de datos SQLite
aparte, para poder probar y previsualizar el reporte de estadisticas antes
de tener semanas de datos reales acumulados.

Este script asume que vive en la MISMA carpeta que entradas_pae26.py, junto
a la carpeta 'datos/' real (usa tus ids_secundaria.xlsx, ids_primaria.xlsx e
ids_docentes.xlsx REALES para la lista de personas, pero INVENTA la
asistencia — no toca ni lee tu asistencia.db real).

Solo necesitas correr este script cuando quieras generar (o regenerar)
datos de prueba. No es necesario para el uso diario del sistema.
"""
import os
import sqlite3
import random
from datetime import date, timedelta
import pandas as pd

random.seed()  # semilla fija: si se vuelve a correr, da los mismos resultados

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'datos')
CARPETA_SIMULACION = os.path.join(BASE_DIR, 'simulacion')
os.makedirs(CARPETA_SIMULACION, exist_ok=True)

DB_PATH = os.path.join(CARPETA_SIMULACION, 'asistencia_SIMULACION.db')
if os.path.exists(DB_PATH):
    os.remove(DB_PATH)

COLUMNAS_IDS = ['ID', 'Nombre', 'Apellido', 'Grado']


def leer_ids_categoria(ruta):
    hojas = pd.read_excel(ruta, sheet_name=None, dtype=str)
    partes = []
    for nombre_hoja, df in hojas.items():
        if df is None or df.empty:
            continue
        df = df.copy()
        if 'Grado' not in df.columns or df['Grado'].isna().all():
            df['Grado'] = nombre_hoja
        for col in COLUMNAS_IDS:
            if col not in df.columns:
                df[col] = None
        partes.append(df[COLUMNAS_IDS])
    todos = pd.concat(partes, ignore_index=True)
    todos = todos.dropna(subset=['ID'])
    todos['ID'] = todos['ID'].astype(str).str.strip()
    return todos


# Debe coincidir con los nombres de archivo reales dentro de tu carpeta 'datos/'
CATEGORIAS = [
    {"nombre": "Estudiantes secundaria", "ids_path": os.path.join(DATA_DIR, "ids_secundaria.xlsx")},
    {"nombre": "Estudiantes primaria",   "ids_path": os.path.join(DATA_DIR, "ids_primaria.xlsx")},
    {"nombre": "Docentes",               "ids_path": os.path.join(DATA_DIR, "ids_docentes.xlsx")},
]

# Multiplicador de asistencia por dia de la semana (0=lunes ... 4=viernes)
MULT_DIA_SEMANA = {0: 0.95, 1: 1.00, 2: 1.02, 3: 1.00, 4: 0.90}


def main():
    for cat in CATEGORIAS:
        if not os.path.exists(cat["ids_path"]):
            raise FileNotFoundError(
                f"No se encontró '{cat['ids_path']}'.\n"
                f"Este script debe estar guardado junto a 'entradas_pae26.py' y su "
                f"carpeta 'datos/' con tus archivos ids_*.xlsx reales."
            )

    conexion = sqlite3.connect(DB_PATH)
    conexion.execute("""CREATE TABLE personas (id TEXT PRIMARY KEY, nombre TEXT, apellido TEXT, grado TEXT, categoria TEXT)""")
    conexion.execute("""CREATE TABLE entradas (id INTEGER PRIMARY KEY AUTOINCREMENT, persona_id TEXT, fecha TEXT, hora TEXT,
                         FOREIGN KEY (persona_id) REFERENCES personas(id))""")
    conexion.execute("""CREATE TABLE estadisticas_diarias (fecha TEXT, categoria TEXT, porcentaje REAL, PRIMARY KEY (fecha, categoria))""")
    conexion.commit()

    personas_por_categoria = {}
    for cat in CATEGORIAS:
        df = leer_ids_categoria(cat["ids_path"])
        personas = []
        for _, p in df.iterrows():
            conexion.execute("INSERT INTO personas VALUES (?, ?, ?, ?, ?)",
                              (p["ID"], p["Nombre"], p["Apellido"], p["Grado"], cat["nombre"]))
            if random.random() < 0.85:
                prob_personal = random.uniform(0.88, 0.97)
            else:
                prob_personal = random.uniform(0.60, 0.80)
            personas.append({"id": p["ID"], "prob": prob_personal})
        personas_por_categoria[cat["nombre"]] = personas
    conexion.commit()

    dia = date(2026, 5, 11)  # lunes
    dias_habiles = []
    while len(dias_habiles) < 25:
        if dia.weekday() < 5:
            dias_habiles.append(dia)
        dia += timedelta(days=1)

    for fecha in dias_habiles:
        mult_dia = MULT_DIA_SEMANA[fecha.weekday()]
        for cat in CATEGORIAS:
            personas = personas_por_categoria[cat["nombre"]]
            entraron = 0
            for persona in personas:
                prob_hoy = min(persona["prob"] * mult_dia, 0.99)
                if random.random() < prob_hoy:
                    entraron += 1
                    minuto_base = random.randint(-15, 15)
                    hora_evento = (12 * 60 + minuto_base) + random.randint(0, 5)
                    hh, mm = divmod(hora_evento, 60)
                    hora_txt = f"{hh:02d}:{mm:02d}:{random.randint(0,59):02d}"
                    conexion.execute(
                        "INSERT INTO entradas (persona_id, fecha, hora) VALUES (?, ?, ?)",
                        (persona["id"], fecha.isoformat(), hora_txt)
                    )
            porcentaje = entraron / len(personas) if personas else 0
            conexion.execute(
                "INSERT INTO estadisticas_diarias (fecha, categoria, porcentaje) VALUES (?, ?, ?)",
                (fecha.isoformat(), cat["nombre"], porcentaje)
            )

    conexion.commit()

    total_entradas = conexion.execute("SELECT COUNT(*) FROM entradas").fetchone()[0]
    total_personas = conexion.execute("SELECT COUNT(*) FROM personas").fetchone()[0]
    total_dias = conexion.execute("SELECT COUNT(DISTINCT fecha) FROM entradas").fetchone()[0]
    print(f"Base de datos simulada creada en: {DB_PATH}")
    print(f"Personas: {total_personas}")
    print(f"Dias simulados: {total_dias}  (del {dias_habiles[0]} al {dias_habiles[-1]})")
    print(f"Entradas totales generadas: {total_entradas}")

    conexion.close()


if __name__ == "__main__":
    main()