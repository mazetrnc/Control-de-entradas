"""
Envía un correo de PRUEBA (sin reporte real adjunto) usando la misma
configuración que usará el programa principal, para confirmar que
correo_config.txt y directivos_correo.txt están bien completados antes de
esperar a que cierre una quincena real.

Debe correrse desde la MISMA carpeta que entradas_pae26.py (junto a la
carpeta 'datos/').

Uso:
    python probar_correo.py
"""
import os
import smtplib
from email.message import EmailMessage
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'datos')
CORREO_CONFIG_PATH = os.path.join(DATA_DIR, 'correo_config.txt')
DIRECTIVOS_CORREO_PATH = os.path.join(DATA_DIR, 'directivos_correo.txt')


def leer_config_correo():
    config = {}
    if os.path.exists(CORREO_CONFIG_PATH):
        with open(CORREO_CONFIG_PATH, "r", encoding="utf-8") as f:
            for linea in f:
                linea = linea.strip()
                if not linea or linea.startswith("#") or "=" not in linea:
                    continue
                clave, _, valor = linea.partition("=")
                config[clave.strip().lower()] = valor.strip()
    return config


def leer_directivos_correo():
    destinatarios = []
    if os.path.exists(DIRECTIVOS_CORREO_PATH):
        with open(DIRECTIVOS_CORREO_PATH, "r", encoding="utf-8") as f:
            for linea in f:
                linea = linea.strip()
                if linea and not linea.startswith("#"):
                    destinatarios.append(linea)
    return destinatarios


def main():
    print(f"Leyendo configuración desde: {CORREO_CONFIG_PATH}")
    config = leer_config_correo()
    destinatarios = leer_directivos_correo()

    faltantes = [c for c in ("servidor", "puerto", "remitente", "contraseña_app") if not config.get(c)]
    if faltantes:
        print(f"❌ Faltan datos en correo_config.txt: {', '.join(faltantes)}")
        return
    if not destinatarios:
        print("❌ No hay ningún correo en directivos_correo.txt (o solo tiene comentarios).")
        return

    print(f"Servidor: {config['servidor']}:{config['puerto']}  (seguridad={config.get('seguridad', 'ssl')})")
    print(f"Remitente: {config['remitente']}")
    print(f"Destinatarios: {', '.join(destinatarios)}")
    print("Intentando enviar correo de prueba...")

    try:
        mensaje = EmailMessage()
        mensaje["Subject"] = "Prueba de configuración — Control de Entradas PAE"
        mensaje["From"] = config["remitente"]
        mensaje["To"] = ", ".join(destinatarios)
        mensaje.set_content(
            f"Este es un correo de PRUEBA generado manualmente el "
            f"{datetime.now().strftime('%d/%m/%Y %H:%M')} para confirmar que la configuración de "
            f"envío automático de reportes está funcionando correctamente.\n\n"
            f"Si recibiste esto, la configuración quedó correcta y los reportes quincenales "
            f"reales llegarán a esta misma lista de correos."
        )

        modo_seguridad = config.get("seguridad", "ssl").strip().lower()
        if modo_seguridad == "starttls":
            with smtplib.SMTP(config["servidor"], int(config["puerto"])) as smtp:
                smtp.starttls()
                smtp.login(config["remitente"], config["contraseña_app"])
                smtp.send_message(mensaje)
        else:
            with smtplib.SMTP_SSL(config["servidor"], int(config["puerto"])) as smtp:
                smtp.login(config["remitente"], config["contraseña_app"])
                smtp.send_message(mensaje)

        print("✅ Correo de prueba enviado correctamente. Revisa la bandeja de entrada (y spam) de los destinatarios.")

    except smtplib.SMTPAuthenticationError:
        print("❌ Google rechazó el usuario/contraseña. Revisa que:")
        print("   - Copiaste bien la contraseña de aplicación de 16 caracteres (no tu contraseña normal).")
        print("   - El correo en 'remitente' es exactamente el mismo con el que generaste esa contraseña.")
    except Exception as e:
        print(f"❌ No se pudo enviar el correo: {e}")


if __name__ == "__main__":
    main()
