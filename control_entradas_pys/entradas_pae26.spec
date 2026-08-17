# -*- mode: python ; coding: utf-8 -*-
#
# Spec de PyInstaller para Control de Entradas PAE.
#
# Uso:
#   1. Coloca este archivo junto a entradas_pae26.py, reporte_quincenal.py
#      y reporte_mensual_detallado.py
#   2. En una terminal, parado en esa carpeta, corre:
#        python -m PyInstaller entradas_pae26.spec
#      (usar "python -m PyInstaller" en vez de solo "pyinstaller" asegura
#      que se use el MISMO Python donde instalaste pandas y las demás
#      librerías -- esta es la causa más común de "módulo no encontrado"
#      a pesar de haberlo instalado).
#   3. El resultado queda en dist/ControlEntradasPAE/
#
# Antes de compilar, instala todo en el MISMO entorno de Python:
#   python -m pip install pyinstaller pandas openpyxl playsound3 matplotlib reportlab

from PyInstaller.utils.hooks import collect_all
import os

datas = []
binaries = []
hiddenimports = []

# collect_all() (en vez de solo --hidden-import) trae TODOS los
# sub-módulos, binarios compilados y archivos de datos internos de cada
# paquete -- esto es justo lo que le faltaba a un --hidden-import simple
# para librerías grandes como pandas, matplotlib y reportlab.
PAQUETES_COMPLETOS = ['pandas', 'openpyxl', 'matplotlib', 'reportlab', 'playsound3']

for paquete in PAQUETES_COMPLETOS:
    d, b, h = collect_all(paquete)
    datas += d
    binaries += b
    hiddenimports += h

# reporte_quincenal.py y reporte_mensual_detallado.py se detectan solos (son
# imports locales normales), pero se dejan explícitos aquí como respaldo por
# si PyInstaller no los detectara.
hiddenimports += ['reporte_quincenal', 'reporte_mensual_detallado']

# Logo opcional: si existe 'recursos/logo.ico' junto a este .spec, se usa
# tanto para el ícono del .exe (lo que se ve en el Explorador de Windows)
# como para empaquetarlo dentro del programa (para el ícono de la ventana,
# que se aplica desde el propio código con ruta_recurso('logo.ico')). Si
# no existe todavía, compila igual, sin ícono personalizado -- en cuanto
# agregues el archivo, solo hay que volver a compilar.
RUTA_LOGO = os.path.join('recursos', 'logo.ico')
TIENE_LOGO = os.path.exists(RUTA_LOGO)
if TIENE_LOGO:
    datas += [(RUTA_LOGO, 'recursos')]

a = Analysis(
    ['entradas_pae26.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ControlEntradasPAE',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # sin ventana de consola negra de fondo
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=RUTA_LOGO if TIENE_LOGO else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ControlEntradas',
)
