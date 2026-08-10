from pathlib import Path
from io import BytesIO
import datetime as dt
import os
import re
import unicodedata

import pandas as pd
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
TIPOS = [("MEDICO", "MÉDICO"), ("PROTECCION CIVIL", "PROTECCIÓN CIVIL"), ("SEGURIDAD", "SEGURIDAD"), ("SERVICIOS PUBLICOS", "SERVICIOS PÚBLICOS"), ("ASISTENCIA", "ASISTENCIA"), ("OTROS SERVICIOS", "OTROS SERVICIOS"), ("IMPROCEDENTES", "IMPROCEDENTES")]
NOMBRES_CORTOS = {"COAHUILA": "Coahuila", "MEXICO": "México", "MICHOACAN": "Michoacán", "VERACRUZ": "Veracruz"}
_LOGOS = {}


def sin_acentos(valor):
    return unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode("ascii")


def normalizar_entidad(valor):
    clave = re.sub(r"\s+", " ", sin_acentos(valor).strip().upper())
    equivalencias = {"COAHUILA DE ZARAGOZA": "COAHUILA", "ESTADO DE MEXICO": "MEXICO", "MICHOACAN DE OCAMPO": "MICHOACAN", "VERACRUZ DE IGNACIO DE LA LLAVE": "VERACRUZ", "CDMX": "CIUDAD DE MEXICO"}
    return equivalencias.get(clave, clave)


def normalizar_tipo(valor):
    texto = re.sub(r"\s+", " ", sin_acentos(valor).strip().upper())
    equivalencias = [("OTROS SERVICIOS", "OTROS SERVICIOS"), ("OTRO SERVICIO", "OTROS SERVICIOS"), ("SERVICIOS PUBLICOS", "SERVICIOS PUBLICOS"), ("SERVICIO PUBLICO", "SERVICIOS PUBLICOS"), ("PROTECCION CIVIL", "PROTECCION CIVIL"), ("SEGURIDAD", "SEGURIDAD"), ("ASISTENCIA", "ASISTENCIA"), ("MEDIC", "MEDICO")]
    for texto_buscado, clave in equivalencias:
        if texto_buscado in texto:
            return clave
    return None


def cargar_logo(plantilla):
    plantilla = Path(plantilla)
    if plantilla not in _LOGOS:
        imagenes = PdfReader(str(plantilla)).pages[0].images
        if not imagenes:
            raise ValueError(f"La plantilla no contiene la imagen institucional: {plantilla}")
        _LOGOS[plantilla] = imagenes[0].data
    return ImageReader(BytesIO(_LOGOS[plantilla]))


def formatear_numero(valor):
    valor = float(valor)
    if abs(valor - round(valor)) > 0.000001:
        raise ValueError(f"El comprobante recibió un total no entero: {valor}")
    return f"{int(round(valor)):,}"


def calcular_variacion(actual, anterior, decimales=2):
    return None if float(anterior) == 0 else round(((float(actual) - float(anterior)) / float(anterior)) * 100, decimales)


def fecha_texto(fecha):
    return f"{fecha.day} de {MESES[fecha.month - 1].lower()} de {fecha.year}, CDMX"


def nombre_seguro(valor):
    return re.sub(r"[^A-Z0-9]+", " ", sin_acentos(valor).upper()).strip()


def crear_pdf(entidad, actuales, anteriores, fecha_periodo, fecha_anterior, plantilla, archivo_salida, observaciones="Sin observaciones."):
    archivo_salida = Path(archivo_salida)
    temporal = archivo_salida.with_name(f"~TEMP_{archivo_salida.name}")
    pdf = canvas.Canvas(str(temporal), pagesize=letter)
    ancho, _ = letter
    mes_actual = MESES[fecha_periodo.month - 1]
    mes_anterior = MESES[fecha_anterior.month - 1]
    total_actual = sum(actuales)
    total_anterior = sum(anteriores)

    pdf.setTitle(f"Comprobante FESE {mes_actual} {fecha_periodo.year} - {entidad}")
    pdf.setAuthor("Centro Nacional de Información y Desarrollo Tecnológico")
    pdf.drawImage(cargar_logo(plantilla), 252.633, 711.337, width=283.471, height=45.525, preserveAspectRatio=True, mask="auto")
    pdf.setFillColor(colors.black)
    pdf.setFont("Times-Roman", 13)
    pdf.drawString(72, 704, "Comprobante")
    pdf.setLineWidth(0.45)
    pdf.line(72, 699, 540, 699)
    pdf.setFont("Times-Roman", 12)
    pdf.drawRightString(540, 653, fecha_texto(dt.date.today()))

    estilo_cuerpo = ParagraphStyle("cuerpo", fontName="Times-Roman", fontSize=11.7, leading=14.4, alignment=TA_JUSTIFY, textColor=colors.black)
    cuerpo = Paragraph(f"El Centro Nacional de Información y Desarrollo Tecnológico emite el presente documento que certifica la entrega de información por parte de la entidad federativa de <b>{entidad}</b>, para el <b>reporte estadístico mensual de la cantidad de llamadas recibidas en el servicio de atención de emergencias 9-1-1</b> correspondiente al mes de {mes_actual.lower()} de {fecha_periodo.year}.", estilo_cuerpo)
    _, alto_cuerpo = cuerpo.wrap(468, 110)
    cuerpo.drawOn(pdf, 72, 604 - alto_cuerpo)

    pdf.setFont("Times-Roman", 12)
    pdf.drawString(72, 490, "Comparativo contra el mes anterior:")

    datos_tabla = [["Tipos generales de incidentes", mes_actual, mes_anterior, "Variación %"]]
    variaciones = []

    for i, ((_, etiqueta), actual, anterior) in enumerate(zip(TIPOS, actuales, anteriores), start=1):
        variacion = calcular_variacion(actual, anterior)
        variaciones.append(variacion)
        datos_tabla.append([f"{i}.- {etiqueta}", formatear_numero(actual), formatear_numero(anterior), "N/D" if variacion is None else f"{variacion:.2f}"])

    variacion_total = calcular_variacion(total_actual, total_anterior, 1)
    datos_tabla.append(["TOTALES", formatear_numero(total_actual), formatear_numero(total_anterior), "N/D" if variacion_total is None else f"{variacion_total:.1f}"])

    tabla = Table(datos_tabla, colWidths=[260, 60, 60, 76], rowHeights=[22] + [17] * 8)
    estilos = [("FONTNAME", (0, 0), (-1, 0), "Times-Bold"), ("FONTSIZE", (0, 0), (-1, 0), 10.8), ("FONTNAME", (0, 1), (-1, -2), "Times-Roman"), ("FONTSIZE", (0, 1), (-1, -2), 10.7), ("FONTNAME", (0, -1), (-1, -1), "Times-Bold"), ("FONTSIZE", (0, -1), (-1, -1), 11.2), ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("ALIGN", (0, 0), (0, -1), "LEFT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5), ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1), ("LINEABOVE", (0, 0), (-1, 0), 0.8, colors.black), ("LINEBELOW", (0, 0), (-1, 0), 0.45, colors.black), ("LINEBELOW", (0, -1), (-1, -1), 0.8, colors.black)]

    for fila in [1, 3, 5, 7]:
        estilos.append(("BACKGROUND", (0, fila), (-1, fila), colors.HexColor("#F0F0F0")))

    for fila, variacion in enumerate(variaciones, start=1):
        if variacion is not None and abs(variacion) >= 10:
            estilos.append(("TEXTCOLOR", (3, fila), (3, fila), colors.blue))

    tabla.setStyle(TableStyle(estilos))
    _, alto_tabla = tabla.wrapOn(pdf, 456, 200)
    y_tabla = 466 - alto_tabla
    tabla.drawOn(pdf, 72, y_tabla)
    pdf.setLineWidth(0.55)
    pdf.line(72, y_tabla - 3, 528, y_tabla - 3)

    estilo_metodologia = ParagraphStyle("metodologia", fontName="Times-Roman", fontSize=9.5, leading=11.5, textColor=colors.black)
    nota_metodologica = Paragraph("<b>Nota metodológica:</b> Las variaciones porcentuales iguales o superiores al 10% en valor absoluto se resaltan en color azul para facilitar la identificación de los cambios más significativos respecto del mes anterior.", estilo_metodologia)
    _, alto_metodologia = nota_metodologica.wrap(468, 45)
    nota_metodologica.drawOn(pdf, 72, y_tabla - 24 - alto_metodologia)

    estilo_nota = ParagraphStyle("nota", fontName="Times-Roman", fontSize=11.5, leading=14)
    nota = Paragraph(f"<b>Observaciones:</b> {observaciones}", estilo_nota)
    nota.wrapOn(pdf, 468, 50)
    nota.drawOn(pdf, 72, 214)

    pdf.setFont("Times-Roman", 10.5)
    pdf.drawCentredString(ancho / 2, 28, "1")
    pdf.save()
    os.replace(temporal, archivo_salida)


def generar_comprobantes(salida, carpeta_destino, plantilla, fecha_periodo, fecha_anterior):
    plantilla = Path(plantilla)

    if not plantilla.is_file():
        raise FileNotFoundError(f"No existe la plantilla de comprobante: {plantilla}")

    requeridas = {"ao", "mes", "estado", "tipo", "procedencia", "total"}

    if not requeridas.issubset(salida.columns):
        raise ValueError(f"Faltan columnas para generar comprobantes: {sorted(requeridas - set(salida.columns))}")

    periodo = ((salida["ao"] == fecha_periodo.year) & (salida["mes"] == fecha_periodo.month)) | ((salida["ao"] == fecha_anterior.year) & (salida["mes"] == fecha_anterior.month))
    datos = salida.loc[periodo, ["ao", "mes", "estado", "tipo", "procedencia", "total"]].copy()
    datos["total"] = pd.to_numeric(datos["total"], errors="raise")
    datos["clave_entidad"] = datos["estado"].apply(normalizar_entidad)
    datos["procedencia_norm"] = datos["procedencia"].apply(lambda x: sin_acentos(x).strip().upper())
    datos["categoria"] = datos["tipo"].apply(normalizar_tipo)
    datos.loc[datos["procedencia_norm"] == "IMPROCEDENTES", "categoria"] = "IMPROCEDENTES"

    sin_categoria = sorted(datos[(datos["procedencia_norm"] != "IMPROCEDENTES") & datos["categoria"].isna() & (datos["total"] != 0)]["tipo"].dropna().astype(str).unique())

    if sin_categoria:
        raise ValueError(f"Hay tipos de incidente sin equivalencia para el comprobante: {sin_categoria}")

    entidades = datos[(datos["ao"] == fecha_periodo.year) & (datos["mes"] == fecha_periodo.month)].groupby("clave_entidad")["estado"].first().sort_index()

    if entidades.empty:
        raise ValueError(f"No hay entidades con datos para {MESES[fecha_periodo.month - 1]} {fecha_periodo.year}.")

    datos = datos[datos["categoria"].notna()].copy()
    resumen = datos.groupby(["clave_entidad", "ao", "mes", "categoria"])["total"].sum()
    carpeta = Path(carpeta_destino) / f"Comprobantes {MESES[fecha_periodo.month - 1]} {fecha_periodo.year}"
    carpeta.mkdir(parents=True, exist_ok=True)

    for clave_entidad, entidad in entidades.items():
        entidad = NOMBRES_CORTOS.get(clave_entidad, entidad)
        actuales = [resumen.get((clave_entidad, fecha_periodo.year, fecha_periodo.month, clave), 0) for clave, _ in TIPOS]
        anteriores = [resumen.get((clave_entidad, fecha_anterior.year, fecha_anterior.month, clave), 0) for clave, _ in TIPOS]
        archivo = carpeta / f"COMPROBANTE FESE {MESES[fecha_periodo.month - 1].upper()} {fecha_periodo.year} - {nombre_seguro(clave_entidad)}.pdf"
        crear_pdf(entidad, actuales, anteriores, fecha_periodo, fecha_anterior, plantilla, archivo)
        print(f"Comprobante generado: {archivo.name}")

    return carpeta