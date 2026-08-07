from pathlib import Path
import csv
import datetime as dt
import json
import re
import unicodedata

import win32com.client as win32

BASE_DIR = Path(__file__).resolve().parent
CARPETA_DESTINO = Path(r"C:\Users\gerardo.noeller\OneDrive - Secretaría de Seguridad y Protección Ciudadana\Escritorio\FESE")
ARCHIVO_CORREOS = BASE_DIR / "correos_comprobantes.txt"
CUENTA_REMITENTE = "admin.bnext.cni@sspc.gob.mx"
MODO = "BORRADOR"  # BORRADOR o ENVIAR

fecha_periodo = dt.date.today().replace(day=1) - dt.timedelta(days=1)
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
NOMBRE_MES = MESES[fecha_periodo.month - 1]
ANIO = fecha_periodo.year
CARPETA_COMPROBANTES = CARPETA_DESTINO / f"Comprobantes {NOMBRE_MES} {ANIO}"
REGISTRO_ENVIOS = CARPETA_COMPROBANTES / "registro_envios_comprobantes.json"


def normalizar_entidad(valor):
    clave = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode("ascii")
    clave = re.sub(r"\s+", " ", clave.strip().upper())
    equivalencias = {"COAHUILA DE ZARAGOZA": "COAHUILA", "ESTADO DE MEXICO": "MEXICO", "MICHOACAN DE OCAMPO": "MICHOACAN", "VERACRUZ DE IGNACIO DE LA LLAVE": "VERACRUZ", "CDMX": "CIUDAD DE MEXICO"}
    return equivalencias.get(clave, clave)


def leer_destinatarios():
    if not ARCHIVO_CORREOS.is_file():
        raise FileNotFoundError(f"No existe el archivo de correos: {ARCHIVO_CORREOS}")

    destinatarios = {}

    with open(ARCHIVO_CORREOS, "r", encoding="utf-8-sig", newline="") as f:
        for numero_linea, fila in enumerate(csv.reader(f, delimiter=";"), start=1):
            if not fila or not any(celda.strip() for celda in fila):
                continue

            entidad = fila[0].strip()

            if normalizar_entidad(entidad) == "ENTIDAD":
                continue

            correos = []

            for celda in fila[1:]:
                correos.extend([correo.strip() for correo in re.split(r"[,\s]+", celda) if correo.strip()])

            correos = list(dict.fromkeys(correos))
            invalidos = [correo for correo in correos if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", correo)]

            if not correos or invalidos:
                raise ValueError(f"Línea {numero_linea} inválida en {ARCHIVO_CORREOS.name}: {fila}")

            clave = normalizar_entidad(entidad)

            if clave not in destinatarios:
                destinatarios[clave] = {"entidad": entidad, "correos": []}

            destinatarios[clave]["correos"] = list(dict.fromkeys(destinatarios[clave]["correos"] + correos))

    if not destinatarios:
        raise ValueError(f"No se encontraron destinatarios en {ARCHIVO_CORREOS}")

    return destinatarios


def indexar_comprobantes():
    CARPETA_COMPROBANTES.mkdir(parents=True, exist_ok=True)

    prefijo = f"COMPROBANTE FESE {NOMBRE_MES.upper()} {ANIO} - "
    comprobantes = {}

    for archivo in CARPETA_COMPROBANTES.glob("*.pdf"):
        if archivo.stem.upper().startswith(prefijo):
            clave = normalizar_entidad(archivo.stem[len(prefijo):])

            if clave in comprobantes:
                raise ValueError(f"Hay dos comprobantes para {clave}: {comprobantes[clave].name} y {archivo.name}")

            comprobantes[clave] = archivo

    return comprobantes


def cargar_registro():
    if not REGISTRO_ENVIOS.is_file():
        return {}

    with open(REGISTRO_ENVIOS, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar_registro(registro):
    temporal = REGISTRO_ENVIOS.with_name(f"~TEMP_{REGISTRO_ENVIOS.name}")

    with open(temporal, "w", encoding="utf-8") as f:
        json.dump(registro, f, ensure_ascii=False, indent=2)

    temporal.replace(REGISTRO_ENVIOS)


def buscar_cuenta_outlook(sesion):
    for cuenta in sesion.Accounts:
        smtp = str(getattr(cuenta, "SmtpAddress", "") or "").strip().lower()

        if smtp == CUENTA_REMITENTE.lower():
            return cuenta

    disponibles = [str(getattr(cuenta, "SmtpAddress", "") or cuenta) for cuenta in sesion.Accounts]
    raise RuntimeError(f"La cuenta {CUENTA_REMITENTE} no está configurada en Outlook. Cuentas disponibles: {disponibles}")


def asignar_cuenta(correo, cuenta):
    try:
        correo._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta))
    except Exception:
        correo.SendUsingAccount = cuenta


def cuerpo_correo(entidad):
    return f"""<html><body style="font-family:Calibri,Arial,sans-serif;font-size:11pt">
<p>Buen día,</p>
<p>Por este medio se remite el comprobante correspondiente a la entrega del FESE del servicio de atención de emergencias 9-1-1 de <b>{NOMBRE_MES.lower()} de {ANIO}</b>, para la entidad de <b>{entidad}</b>.</p>
<p>Agradecemos su atención y colaboración.</p>
<p>Saludos.</p>
<p>Centro Nacional de Información y Desarrollo Tecnológico</p>
</body></html>"""


def ejecutar():
    modo = MODO.strip().upper()

    if modo not in {"BORRADOR", "ENVIAR"}:
        raise ValueError("MODO debe ser BORRADOR o ENVIAR.")

    destinatarios = leer_destinatarios()
    comprobantes = indexar_comprobantes()
    faltantes = sorted(set(destinatarios) - set(comprobantes))

    if faltantes:
        raise FileNotFoundError(f"Faltan comprobantes para estas entidades: {faltantes}")

    if modo == "ENVIAR" and input(f"Se enviarán {len(destinatarios)} correos reales desde {CUENTA_REMITENTE}. Escribe ENVIAR para continuar: ").strip() != "ENVIAR":
        print("Envío cancelado.")
        return

    outlook = win32.Dispatch("Outlook.Application")
    sesion = outlook.Session
    cuenta = buscar_cuenta_outlook(sesion)
    registro = cargar_registro()
    errores = []

    for clave, datos in destinatarios.items():
        llave_registro = f"{ANIO}-{fecha_periodo.month:02d}|{clave}"

        if modo == "ENVIAR" and llave_registro in registro:
            print(f"Ya enviado, se omite: {datos['entidad']}")
            continue

        try:
            correo = outlook.CreateItem(0)
            asignar_cuenta(correo, cuenta)
            correo.To = ";".join(datos["correos"])
            correo.Subject = f"Comprobante FESE 9-1-1 - {NOMBRE_MES} {ANIO} - {datos['entidad']}"
            correo.HTMLBody = cuerpo_correo(datos["entidad"])
            correo.Attachments.Add(str(comprobantes[clave].resolve()))

            if modo == "BORRADOR":
                correo.Save()
                print(f"Borrador creado: {datos['entidad']} -> {correo.To}")
            else:
                correo.Send()
                registro[llave_registro] = {"entidad": datos["entidad"], "correos": datos["correos"], "archivo": comprobantes[clave].name, "fecha_envio": dt.datetime.now().isoformat(timespec="seconds")}
                guardar_registro(registro)
                print(f"Enviado: {datos['entidad']} -> {';'.join(datos['correos'])}")

        except Exception as e:
            errores.append(f"{datos['entidad']}: {e}")
            print(f"ERROR: {datos['entidad']}: {e}")

    if errores:
        raise RuntimeError("Hubo errores:\n- " + "\n- ".join(errores))

    print("Borradores terminados." if modo == "BORRADOR" else "Envíos terminados.")


if __name__ == "__main__":
    ejecutar()