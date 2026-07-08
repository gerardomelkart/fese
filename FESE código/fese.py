from pathlib import Path
import datetime as dt
import locale
import shutil
import json
import re
import win32com.client as win32

import pandas as pd
import openpyxl
import numpy as np
import pyreadr
from unidecode import unidecode
from time import perf_counter

BASE_DIR = Path(__file__).resolve().parent
CARPETA_INSUMOS = BASE_DIR / "insumos"
CARPETA_DESTINO = Path(r"C:\Users\gerardo.noeller\OneDrive - Secretaría de Seguridad y Protección Ciudadana\Escritorio\FESE")
PLANTILLA_CNIEDT = CARPETA_DESTINO / "Formato CNIEDT plantilla.xlsx"


# REVISAR BIEN PUEBLA Y SINALOA que cuadren el total con loq eu mandan en el excel
# Hacer validación de que sean todos números enteros

# Configuramos la localización en español de México para que nos dé el nombre correcto del mes pasado
locale.setlocale(locale.LC_TIME, "es_MX.UTF-8")

# Conseguimos el nombre del mes pasado a través del día de hoy
fecha_mes_pasado = dt.date.today().replace(day=1) - dt.timedelta(days=1)
nombre_mes_pasado = fecha_mes_pasado.strftime("%B").title()
mes_pasado = int(fecha_mes_pasado.strftime("%#m"))
anio_mes_pasado = int(fecha_mes_pasado.strftime("%Y"))

# Conseguimos el nombre del mes anterior al pasado a través del día de hoy
fecha_mes_anterior = fecha_mes_pasado.replace(day=1) - dt.timedelta(days=1)
nombre_mes_anterior = fecha_mes_anterior.strftime("%B").title()
mes_anterior = int(fecha_mes_anterior.strftime("%#m"))
anio_mes_anterior = int(fecha_mes_anterior.strftime("%Y"))


mapa_estados = {
    "Aguascalientes": "Aguascalientes",
    "Baja California": "Baja California",
    "Baja California Sur": "Baja California Sur",
    "Campeche": "Campeche",
    "Coahuila": "Coahuila de Zaragoza",
    "Colima": "Colima",
    "Chiapas": "Chiapas",
    "Chihuahua": "Chihuahua",
    "Ciudad De Mexico": "Ciudad de México",
    "Durango": "Durango",
    "Guanajuato": "Guanajuato",
    "Guerrero": "Guerrero",
    "Hidalgo": "Hidalgo",
    "Jalisco": "Jalisco",
    "Estado De Mexico": "México",
    "Estado De México": "México",
    # "Michoacan": "Michoacán de Ocampo",
    "Michoacán": "Michoacán de Ocampo",
    "Morelos": "Morelos",
    "Nayarit": "Nayarit",
    "Nuevo Leon": "Nuevo León",
    "Oaxaca": "Oaxaca",
    "Puebla": "Puebla",
    "Queretaro": "Querétaro",
    "Quintana Roo": "Quintana Roo",
    "San Luis Potosi": "San Luis Potosí",
    "Sinaloa": "Sinaloa",
    "Sonora": "Sonora",
    "Tabasco": "Tabasco",
    "Tamaulipas": "Tamaulipas",
    "Tlaxcala": "Tlaxcala",
    "Veracruz": "Veracruz de Ignacio de la Llave",
    "Yucatan": "Yucatán",
    "Zacatecas": "Zacatecas",
}

columnas_long = [
    "codigo",
    "ao",
    "procedencia",
    "region",
    "estado",
    "centro",
    "tipo",
    "incidente",
    "total",
    "mes",
    "mes_largo",
    "fecha",
]


def proceso_fese():
    inicio_total = perf_counter()

    inicio = perf_counter()
    fese = cargar_y_limpiar(BASE_DIR / f"fese{anio_mes_anterior}-{mes_anterior}.rds")
    print(f"⏱ RDS histórico cargado en {perf_counter() - inicio:.1f} s")

    insumos = obtener_insumos(CARPETA_INSUMOS)
    print(f"📁 Insumos encontrados: {len(insumos)}")

    inicio = perf_counter()
    resultados, validaciones = leer_insumos(fese, insumos)
    print(f"⏱ Insumos procesados en {perf_counter() - inicio:.1f} s")

    inicio = perf_counter()
    generar_salidas(fese, columnas_long, resultados, validaciones)
    print(f"⏱ Salidas generadas en {perf_counter() - inicio:.1f} s")

    print(f"⏱ PROCESO TOTAL: {perf_counter() - inicio_total:.1f} s")


def cargar_y_limpiar(nombre_archivo="df_fese.rds"):
    # Cargamos el archivo histórico en un dataframe
    fese = pyreadr.read_r(nombre_archivo)[None]  # pd.read_excel("fese.xlsx")
    fese["codigo"] = fese["codigo"].astype(int)
    mapa_estados = {
        "CIUDAD DE MEXICO": "CIUDAD DE MÉXICO",
        "ESTADO DE MEXICO": "ESTADO DE MÉXICO",
        "NUEVO LEON": "NUEVO LEÓN",
        "QUERETARO": "QUERÉTARO",
        "SAN LUIS POTOSI": "SAN LUIS POTOSÍ",
        "YUCATAN": "YUCATÁN",
        "COAHUILA": "COAHUILA DE ZARAGOZA",
    }
    fese.estado = fese.estado.str.upper()
    fese.estado = fese.estado.replace(mapa_estados)
    return fese


def obtener_insumos(path_insumos="insumos"):
    # Obtenemos los nombres de todos los archivos de Excel insumo.
    # Esto se hace leyendo todos los objetos dentro del directorio e ignorando los que no son un archivo
    return [p for p in Path(path_insumos).iterdir() if p.is_file() and p.suffix.lower() in [".xlsx", ".xlsm", ".xls"] and not p.name.startswith("~$")]


def leer_insumos(fese, insumos):
    validaciones = []
    resultados = []

    print("Preparando catálogos históricos...")
    inicio_catalogos = perf_counter()

    codigos = set(fese["codigo"].unique())
    centros_historicos = set(fese["centro"].dropna().unique())
    codigo_incidente = fese.groupby("codigo")["incidente"].first().to_dict()

    estado_region_grupos = fese.groupby("region")["estado"].unique().to_dict()
    estado_a_region = {estado: region for region, estados in estado_region_grupos.items() for estado in estados}

    codigo_tipo_grupos = fese.groupby("tipo")["codigo"].unique().to_dict()
    codigo_a_tipo = {codigo: tipo for tipo, codigos_tipo in codigo_tipo_grupos.items() for codigo in codigos_tipo}

    print(f"⏱ Catálogos preparados en {perf_counter() - inicio_catalogos:.1f} s")

    codigos_improcedentes = {70101, 70102, 70103, 70104, 70105, 70106, 70107, 70108}

    for insumo in insumos:
        inicio_archivo = perf_counter()
        print(f"\nAbriendo archivo: {insumo.name}")

        with pd.ExcelFile(insumo) as xlsx:
            hojas_con_datos = [hoja for hoja in xlsx.sheet_names if hoja.startswith("C")]
            datos_hojas = pd.read_excel(xlsx, sheet_name=hojas_con_datos, usecols="B:D", skiprows=18)
            workbook = openpyxl.load_workbook(insumo, data_only=True, read_only=True)

            try:
                resumen_c10 = workbook["RESUMEN"]["C10"].value
                resumen_c11 = workbook["RESUMEN"]["C11"].value

                for hoja in hojas_con_datos:
                    hoja_centro = workbook[hoja]
                    nombre_centro = hoja_centro["D18"].value.strip()
                    nombre_entidad = hoja_centro["D12"].value

                    if nombre_entidad is None:
                        nombre_entidad = resumen_c10
                    if nombre_entidad is None:
                        nombre_entidad = resumen_c11

                    nombre_entidad = nombre_entidad.strip()
                    print(f"Trabajando: {nombre_entidad}, {nombre_centro}, hoja={hoja}, archivo={insumo}")

                    datos_centro = datos_hojas[hoja].copy()

                    datos_centro.rename(
                        columns={
                            datos_centro.columns[0]: "codigo",
                            datos_centro.columns[1]: "incidente",
                            datos_centro.columns[2]: "total",
                        },
                        inplace=True,
                    )

                    datos_centro["incidente"] = datos_centro["incidente"].str.title().str.strip()
                    datos_centro["total"] = pd.to_numeric(datos_centro["total"], errors="coerce").fillna(0)
                    datos_centro = datos_centro[datos_centro["codigo"].isin(codigos)].copy()

                    datos_centro = datos_centro.assign(
                        estado=nombre_entidad,
                        centro=nombre_centro,
                        ao=anio_mes_pasado,
                        mes=mes_pasado,
                        mes_largo=nombre_mes_pasado,
                        fecha=fecha_mes_pasado.replace(day=1),
                    )

                    mapeo_estado = {
                        "México": "MEXICO",
                        "Michoacán": "MICHOACAN DE OCAMPO",
                    }

                    datos_centro["estado"] = datos_centro["estado"].replace(mapeo_estado).str.upper()
                    datos_centro["procedencia"] = np.where(datos_centro["codigo"].isin(codigos_improcedentes), "Improcedentes", "Procedentes")
                    datos_centro["estado"] = datos_centro["estado"].str.title().replace(mapa_estados).str.upper()

                    datos_centro["incidente"] = datos_centro["codigo"].map(codigo_incidente)
                    datos_centro["region"] = datos_centro["estado"].map(estado_a_region).fillna("")
                    datos_centro["tipo"] = datos_centro["codigo"].map(codigo_a_tipo).fillna("")

                    datos_centro = datos_centro[columnas_long]

                    if nombre_centro not in centros_historicos:
                        validaciones.append(
                            {
                                "clave_error": "centro_no_existe",
                                "centro": nombre_centro,
                                "entidad": nombre_entidad,
                                "archivo": str(insumo),
                                "hoja": hoja,
                            }
                        )

                    if not (datos_centro["total"] >= 0).all():
                        filas_con_error = list(datos_centro[datos_centro["total"] < 0]["incidente"])
                        validaciones.append(
                            {
                                "clave_error": "numeros_negativos",
                                "centro": nombre_centro,
                                "entidad": nombre_entidad,
                                "archivo": str(insumo),
                                "hoja": hoja,
                                "incidentes": filas_con_error,
                            }
                        )

                    datos_mes_pasado = datos_centro[
                        (datos_centro["mes"] == mes_pasado)
                        & (datos_centro["ao"] == anio_mes_pasado)
                    ][columnas_long]

                    datos_mes_anterior = datos_centro[
                        (datos_centro["mes"] == mes_anterior)
                        & (datos_centro["ao"] == mes_anterior)
                    ][columnas_long]

                    datos_ambos_meses = datos_mes_pasado.join(
                        datos_mes_anterior,
                        "codigo",
                        lsuffix="_pasado",
                        rsuffix="_anterior",
                    )

                    if (
                        datos_ambos_meses["total_pasado"]
                        == datos_ambos_meses["total_anterior"]
                    ).all():
                        validaciones.append(
                            {
                                "clave_error": "mes_pasado_igual_anterior",
                                "centro": nombre_centro,
                                "entidad": nombre_entidad,
                                "archivo": str(insumo),
                                "hoja": hoja,
                            }
                        )

                    datos_anio_pasado = datos_centro[
                        (datos_centro["mes"] == mes_pasado)
                        & (datos_centro["ao"] == (anio_mes_pasado - 1))
                    ][columnas_long]

                    datos_ambos_meses = datos_mes_pasado.join(
                        datos_anio_pasado,
                        "codigo",
                        lsuffix="_pasado",
                        rsuffix="_anio_pasado",
                    )

                    if (
                        datos_ambos_meses["total_pasado"]
                        == datos_ambos_meses["total_anio_pasado"]
                    ).all():
                        validaciones.append(
                            {
                                "clave_error": "mes_pasado_igual_anio_pasado",
                                "centro": nombre_centro,
                                "entidad": nombre_entidad,
                                "archivo": str(insumo),
                                "hoja": hoja,
                            }
                        )

                    if datos_centro["total"].sum() == 0:
                        validaciones.append(
                            {
                                "clave_error": "suma_total_cero",
                                "centro": nombre_centro,
                                "entidad": nombre_entidad,
                                "archivo": str(insumo),
                                "hoja": hoja,
                            }
                        )

                    resultados.append(datos_centro)

            finally:
                workbook.close()

        print(f"⏱ Archivo terminado en {perf_counter() - inicio_archivo:.1f} s: {insumo.name}")

    return resultados, validaciones

def normalizar_clave_entidad(valor):
    clave = unidecode(str(valor)).strip().upper()
    equivalencias = {
        "COAHUILA DE ZARAGOZA": "COAHUILA",
        "ESTADO DE MEXICO": "MEXICO",
        "MICHOACAN DE OCAMPO": "MICHOACAN",
        "VERACRUZ DE IGNACIO DE LA LLAVE": "VERACRUZ",
    }
    return equivalencias.get(clave, clave)

def generar_formato_cniedt(salida):
    if not PLANTILLA_CNIEDT.is_file():
        raise FileNotFoundError(f"No existe la plantilla CNIEDT: {PLANTILLA_CNIEDT}")

    datos = salida[(salida["ao"] == anio_mes_pasado) & (salida["mes"] == mes_pasado) & (salida["procedencia"] == "Procedentes")].copy()

    if datos.empty:
        raise ValueError(f"No hay datos procedentes para {nombre_mes_pasado} {anio_mes_pasado}.")

    datos["clave_entidad"] = datos["estado"].apply(normalizar_clave_entidad)
    totales = datos.groupby("clave_entidad")["total"].sum().to_dict()

    archivo_salida = CARPETA_DESTINO / f"Formato CNIEDT {anio_mes_pasado}-{str(mes_pasado).zfill(2)}.xlsx"

    if archivo_salida.exists():
        raise FileExistsError(f"Ya existe el formato CNIEDT generado: {archivo_salida}")

    shutil.copy2(PLANTILLA_CNIEDT, archivo_salida)

    excel = None
    libro = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        libro = excel.Workbooks.Open(str(archivo_salida))
        hoja = libro.Worksheets("Llamadas procedentes 911")

        tabla = None

        for i in range(1, hoja.ListObjects.Count + 1):
            candidata = hoja.ListObjects(i)
            encabezados = [str(candidata.HeaderRowRange.Cells(1, j).Value).strip() for j in range(1, candidata.ListColumns.Count + 1)]

            if {"Entidad", "Mes", "Número", "TOTAL"}.issubset(set(encabezados)):
                tabla = candidata
                break

        if tabla is None:
            raise ValueError("No se encontró la tabla Entidad/Mes/Número/TOTAL en la pestaña 'Llamadas procedentes 911'.")

        encabezados = [str(tabla.HeaderRowRange.Cells(1, j).Value).strip() for j in range(1, tabla.ListColumns.Count + 1)]
        idx_entidad = encabezados.index("Entidad") + 1
        idx_mes = encabezados.index("Mes") + 1
        idx_numero = encabezados.index("Número") + 1
        idx_total = encabezados.index("TOTAL") + 1

        filas_objetivo = []
        entidades_formato = set()

        for fila in range(1, tabla.DataBodyRange.Rows.Count + 1):
            entidad = tabla.DataBodyRange.Cells(fila, idx_entidad).Value
            mes = tabla.DataBodyRange.Cells(fila, idx_mes).Value
            numero = tabla.DataBodyRange.Cells(fila, idx_numero).Value

            try:
                es_911 = int(float(numero)) == 911
            except (TypeError, ValueError):
                es_911 = False

            if str(mes).strip().lower() == nombre_mes_pasado.lower() and es_911:
                clave = normalizar_clave_entidad(entidad)
                filas_objetivo.append((fila, clave, entidad))
                entidades_formato.add(clave)

        if len(filas_objetivo) != 32:
            raise ValueError(f"Se esperaban 32 entidades para {nombre_mes_pasado} y 911, pero se encontraron {len(filas_objetivo)}.")

        faltantes = entidades_formato - set(totales.keys())
        extras = set(totales.keys()) - entidades_formato

        if faltantes:
            raise ValueError(f"Faltan entidades en los datos FESE: {sorted(faltantes)}")

        if extras:
            raise ValueError(f"Hay entidades FESE que no coinciden con el formato CNIEDT: {sorted(extras)}")

        for fila, clave, entidad in filas_objetivo:
            tabla.DataBodyRange.Cells(fila, idx_total).Value = float(totales[clave])

        excel.CalculateFull()
        libro.Save()

    finally:
        if libro is not None:
            libro.Close(SaveChanges=True)

        if excel is not None:
            excel.Quit()

    return archivo_salida

def generar_salidas(fese, columnas_long, resultados, validaciones):
    resultados = [fese[columnas_long]] + resultados
    salida = pd.concat(resultados, ignore_index=True)

    mapeo_centro = {
        "Izucar De Matamoros": "Izúcar De Matamoros",
        "C4 Gomez Palacio": "C4 Gómez Palacio",
        "C4 Obregon": "C4 Obregón",
        "Sub C4 Lazaro Cardenas": "Sub C4 Lázaro Cárdenas",
        "Sub C4 Patzcuaro": "Sub C4 Pátzcuaro",
        "Sub C4 Zitacuaro": "Sub C4 Zitácuaro",
        "El Marqués": "El Márques",
        "C4 Valparaiso": "C4 Valparaíso",
        "Subcentro C4 Cancun": "Subcentro C4 Cancún",
        "C4 Martinez de la torre": "C4 Martínez de la torre",
    }
    salida.centro = salida.centro.replace(mapeo_centro)
    wide = salida[columnas_long].pivot(
        index=[
            "ao",
            "region",
            "estado",
            "centro",
            "tipo",
            "codigo",
            "incidente",
            "procedencia",
        ],
        columns="mes_largo",
        values="total",
    )

    wide = wide.reset_index()

    wide.estado = wide.estado.str.title().replace(mapa_estados).str.upper()

    wide.rename(
        columns={
            "codigo": "Código",
            "ao": "Año",
            "procedencia": "Procedencia",
            "region": "Region",
            "estado": "Estado",
            "centro": "Centro",
            "tipo": "Tipo",
            "incidente": "Incidente",
        },
        inplace=True,
    )

    columnas_wide = [
        "Código",
        "Año",
        "Procedencia",
        "Region",
        "Estado",
        "Centro",
        "Tipo",
        "Incidente",
        "Enero",
        "Febrero",
        "Marzo",
        "Abril",
        "Mayo",
        "Junio",
        "Julio",
        "Agosto",
        "Septiembre",
        "Octubre",
        "Noviembre",
        "Diciembre",
    ]
    wide = wide[columnas_wide]

    meses = [
        "Enero",
        "Febrero",
        "Marzo",
        "Abril",
        "Mayo",
        "Junio",
        "Julio",
        "Agosto",
        "Septiembre",
        "Octubre",
        "Noviembre",
        "Diciembre",
    ]

    for mes in meses:
        wide[mes] = wide[mes].fillna(0)

    wide = wide.sort_values(by=["Año", "Código"], ascending=[True, True])

    archivo_excel = BASE_DIR / f"Rep_anual{anio_mes_pasado}-{str(mes_pasado).zfill(2)}.xlsx"
    wide.to_excel(archivo_excel, index=False)

    salida.estado = salida.estado.str.title()
    salida.estado = salida.estado.apply(unidecode)
    salida.estado = salida.estado.replace(mapa_estados)

    archivo_rds = BASE_DIR / f"fese{anio_mes_pasado}-{mes_pasado}.rds"
    pyreadr.write_rds(str(archivo_rds), salida)

    with open(BASE_DIR / "validaciones.json", "w") as f:
        json.dump(validaciones, f)

    copiar_resultados_y_limpiar(archivo_excel, archivo_rds, salida)


def copiar_resultados_y_limpiar(archivo_excel, archivo_rds, salida):
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
    carpeta_formatos = CARPETA_DESTINO / f"Formatos {nombre_mes_pasado} {anio_mes_pasado}"

    if carpeta_formatos.exists():
        print(f"La carpeta de formatos ya existe, se omite la copia: {carpeta_formatos}")
    else:
        shutil.copytree(CARPETA_INSUMOS, carpeta_formatos)
        print(f"Formatos copiados a: {carpeta_formatos}")
    destino_excel = CARPETA_DESTINO / archivo_excel.name
    destino_rds = CARPETA_DESTINO / archivo_rds.name
    shutil.copy2(archivo_excel, destino_excel)
    shutil.copy2(archivo_rds, destino_rds)

    archivo_cniedt = generar_formato_cniedt(salida)

    if not destino_excel.is_file() or not destino_rds.is_file() or not carpeta_formatos.is_dir() or not archivo_cniedt.is_file():
        raise RuntimeError("No se completó correctamente la generación y copia de resultados. No se limpiarán los archivos originales.")

    print(f"Excel copiado a: {destino_excel}")
    print(f"RDS copiado a: {destino_rds}")
    print(f"CNIEDT generado en: {archivo_cniedt}")

    for item in CARPETA_INSUMOS.iterdir():
        if item.is_dir():
            shutil.rmtree(item)
        else:
            item.unlink()

    meses_conservar = {(anio_mes_pasado, mes_pasado), (anio_mes_anterior, mes_anterior)}
    patron_rds = re.compile(r"^fese(\d{4})-(\d{1,2})\.rds$", re.IGNORECASE)
    patron_excel = re.compile(r"^Rep_anual(\d{4})-(\d{1,2})\.xlsx$", re.IGNORECASE)

    for archivo in BASE_DIR.iterdir():
        coincidencia = patron_rds.match(archivo.name) or patron_excel.match(archivo.name)

        if coincidencia and (int(coincidencia.group(1)), int(coincidencia.group(2))) not in meses_conservar:
            archivo.unlink()
            print(f"Resultado antiguo eliminado: {archivo.name}")

    print("Carpeta insumos vaciada correctamente.")


if __name__ == "__main__":
    proceso_fese()
