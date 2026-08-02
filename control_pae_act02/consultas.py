"""
Consultas SQL para el reporte de asistencia.

Este script asume que vive en la MISMA carpeta que entradas_pae26.py,
junto a la carpeta 'datos/' que ya usa la aplicación principal.

Tiene un interruptor (USAR_DATOS_SIMULADOS) para alternar entre la base de
datos simulada (para pruebas) y la base de datos REAL, sin tener que andar
buscando qué línea editar.
"""
import os
import sqlite3
import pandas as pd

# Carpeta donde vive ESTE script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Carpeta 'datos/' real, la misma que usa entradas_pae26.py
DATA_DIR = os.path.join(BASE_DIR, 'datos')

# Carpeta aparte para no mezclar nunca los datos simulados con los reales
CARPETA_SIMULACION = os.path.join(BASE_DIR, 'simulacion')

# ------------------------------------------------------------------
# CAMBIA ESTO cuando quieras usar la base de datos real en vez de la
# simulada (normalmente, una vez que haya semanas de asistencia real
# acumulada tras la prueba piloto):
#
#   USAR_DATOS_SIMULADOS = False
#
# ------------------------------------------------------------------
USAR_DATOS_SIMULADOS = True

if USAR_DATOS_SIMULADOS:
    DB_PATH = os.path.join(CARPETA_SIMULACION, 'asistencia_SIMULACION.db')
else:
    DB_PATH = os.path.join(DATA_DIR, 'asistencia.db')


def conectar():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(
            f"No se encontró la base de datos en:\n  {DB_PATH}\n"
            f"Revisa que este script esté guardado junto a 'entradas_pae26.py' "
            f"y su carpeta 'datos/', o corre primero generar_simulacion.py si "
            f"estás probando con datos simulados."
        )
    return sqlite3.connect(DB_PATH)


# 1) Tendencia diaria general: % de asistencia global (todas las
#    categorias juntas) para cada dia, para ver si sube o baja con el tiempo.
def tendencia_diaria():
    conexion = conectar()
    df = pd.read_sql("""
        SELECT e.fecha,
               COUNT(DISTINCT e.persona_id) * 1.0 / (SELECT COUNT(*) FROM personas) AS porcentaje
        FROM entradas e
        GROUP BY e.fecha
        ORDER BY e.fecha
    """, conexion)
    conexion.close()
    df['fecha'] = pd.to_datetime(df['fecha'])
    return df


# 2) Asistencia PROMEDIO por dia de la semana (clave para prever comida)
def promedio_por_dia_semana():
    conexion = conectar()
    df = pd.read_sql("SELECT fecha, categoria, porcentaje FROM estadisticas_diarias", conexion)
    conexion.close()
    df['fecha'] = pd.to_datetime(df['fecha'])
    dias_es = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
    df['dia_semana'] = df['fecha'].dt.weekday.map(dict(enumerate(dias_es)))
    resumen = df.groupby('dia_semana')['porcentaje'].mean().reindex(dias_es)
    return resumen


# 3) Asistencia promedio por categoria, para todo el periodo
def promedio_por_categoria():
    conexion = conectar()
    df = pd.read_sql("SELECT categoria, porcentaje FROM estadisticas_diarias", conexion)
    conexion.close()
    return df.groupby('categoria')['porcentaje'].mean()


# 4) Total de personas por categoria
def total_por_categoria():
    conexion = conectar()
    df = pd.read_sql("SELECT categoria, COUNT(*) as total FROM personas GROUP BY categoria", conexion)
    conexion.close()
    return df.set_index('categoria')['total']


# 5) Asistencia promedio por dia de la semana Y por categoria
def promedio_por_dia_y_categoria():
    conexion = conectar()
    df = pd.read_sql("SELECT fecha, categoria, porcentaje FROM estadisticas_diarias", conexion)
    conexion.close()
    df['fecha'] = pd.to_datetime(df['fecha'])
    dias_es = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
    df['dia_semana'] = df['fecha'].dt.weekday.map(dict(enumerate(dias_es)))
    tabla = df.pivot_table(index='dia_semana', columns='categoria', values='porcentaje', aggfunc='mean')
    return tabla.reindex(dias_es)


if __name__ == "__main__":
    print(f"Usando base de datos: {DB_PATH}")
    print("\n=== Tendencia diaria (primeras filas) ===")
    print(tendencia_diaria().head())
    print("\n=== Promedio por dia de semana ===")
    print(promedio_por_dia_semana())
    print("\n=== Promedio por categoria ===")
    print(promedio_por_categoria())
    print("\n=== Total por categoria ===")
    print(total_por_categoria())