from pathlib import Path
import datetime as dt
import locale
import shutil
import json
import re
import os
import subprocess
import sys
import win32com.client as win32

import pandas as pd
import openpyxl
import numpy as np
import pyreadr
from unidecode import unidecode
from time import perf_counter
from modulos.comprobantes import generar_comprobantes

BASE_DIR = Path(__file__).resolve().parent
CARPETA_INSUMOS = BASE_DIR / "insumos"
CARPETA_DATOS = BASE_DIR / "datos"
CARPETA_SALIDAS = BASE_DIR / "salidas"
CARPETA_DESTINO = Path(r"C:\Users\gerardo.noeller\OneDrive - Secretaría de Seguridad y Protección Ciudadana\Escritorio\FESE")
CARPETA_PLANTILLAS = BASE_DIR / "plantillas"
PLANTILLA_CNIEDT = CARPETA_PLANTILLAS / "Formato CNIEDT plantilla.xlsx"
PLANTILLA_COMPROBANTE = CARPETA_PLANTILLAS / "COMPROBANTE FESE plantilla.pdf"
CARPETA_CACHE = Path(os.getenv("LOCALAPPDATA", str(BASE_DIR))) / "FESE" / "cache"

for carpeta in (CARPETA_INSUMOS, CARPETA_DATOS, CARPETA_SALIDAS):
    carpeta.mkdir(parents=True, exist_ok=True)

# REVISAR BIEN PUEBLA Y SINALOA que cuadren el total con loq eu mandan en el excel
# Hacer validación de que sean todos números enteros

locale.setlocale(locale.LC_TIME, "es_MX.UTF-8")

fecha_mes_pasado = dt.date.today().replace(day=1) - dt.timedelta(days=1)
nombre_mes_pasado = fecha_mes_pasado.strftime("%B").title()
mes_pasado = int(fecha_mes_pasado.strftime("%#m"))
anio_mes_pasado = int(fecha_mes_pasado.strftime("%Y"))

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

columnas_long = ["codigo", "ao", "procedencia", "region", "estado", "centro", "tipo", "incidente", "total", "mes", "mes_largo", "fecha"]

MAPA_ESTADOS_HISTORICO = {
    "CIUDAD DE MEXICO": "CIUDAD DE MÉXICO",
    "ESTADO DE MEXICO": "ESTADO DE MÉXICO",
    "NUEVO LEON": "NUEVO LEÓN",
    "QUERETARO": "QUERÉTARO",
    "SAN LUIS POTOSI": "SAN LUIS POTOSÍ",
    "YUCATAN": "YUCATÁN",
    "COAHUILA": "COAHUILA DE ZARAGOZA",
}


def proceso_fese():
    inicio_total = perf_counter()

    inicio = perf_counter()
    fese = cargar_y_limpiar(CARPETA_DATOS / f"fese{anio_mes_anterior}-{mes_anterior}.rds")
    print(f"⏱ Histórico cargado en {perf_counter() - inicio:.1f} s")

    insumos = obtener_insumos(CARPETA_INSUMOS)
    print(f"📁 Insumos encontrados: {len(insumos)}")

    inicio = perf_counter()
    resultados, validaciones = leer_insumos(fese, insumos)
    print(f"⏱ Insumos procesados en {perf_counter() - inicio:.1f} s")

    inicio = perf_counter()
    generar_salidas(fese, columnas_long, resultados, validaciones)
    print(f"⏱ Salidas generadas en {perf_counter() - inicio:.1f} s")
    print(f"⏱ PROCESO TOTAL: {perf_counter() - inicio_total:.1f} s")


def normalizar_historico(fese):
    fese["codigo"] = fese["codigo"].astype(int)
    fese["estado"] = fese["estado"].str.upper().replace(MAPA_ESTADOS_HISTORICO)
    return fese


def rutas_cache_historico(nombre_archivo):
    nombre_archivo = Path(nombre_archivo)
    return CARPETA_CACHE / f"{nombre_archivo.stem}.pkl", CARPETA_CACHE / f"{nombre_archivo.stem}.meta.json"


def huella_archivo(nombre_archivo):
    nombre_archivo = Path(nombre_archivo)
    stat = nombre_archivo.stat()
    return {"nombre": nombre_archivo.name, "tamano": stat.st_size, "mtime_ns": stat.st_mtime_ns}


def cache_historico_valida(nombre_archivo, archivo_cache, archivo_meta):
    if not archivo_cache.is_file() or not archivo_meta.is_file():
        return False
    try:
        with open(archivo_meta, "r", encoding="utf-8") as f:
            return json.load(f) == huella_archivo(nombre_archivo)
    except (OSError, ValueError, json.JSONDecodeError):
        return False


def guardar_cache_historico(fese, nombre_archivo):
    inicio = perf_counter()
    CARPETA_CACHE.mkdir(parents=True, exist_ok=True)
    archivo_cache, archivo_meta = rutas_cache_historico(nombre_archivo)
    temporal_cache = archivo_cache.with_name(f"~TEMP_{archivo_cache.name}")
    temporal_meta = archivo_meta.with_name(f"~TEMP_{archivo_meta.name}")

    for temporal in (temporal_cache, temporal_meta):
        if temporal.exists():
            temporal.unlink()

    fese.to_pickle(temporal_cache)
    with open(temporal_meta, "w", encoding="utf-8") as f:
        json.dump(huella_archivo(nombre_archivo), f)

    os.replace(temporal_cache, archivo_cache)
    os.replace(temporal_meta, archivo_meta)
    print(f"⚡ Cache histórico actualizado en {perf_counter() - inicio:.1f} s: {archivo_cache}")


def registrar_cache_desde_pickle(archivo_pickle, nombre_archivo):
    inicio = perf_counter()
    CARPETA_CACHE.mkdir(parents=True, exist_ok=True)
    archivo_cache, archivo_meta = rutas_cache_historico(nombre_archivo)
    temporal_meta = archivo_meta.with_name(f"~TEMP_{archivo_meta.name}")

    if temporal_meta.exists():
        temporal_meta.unlink()

    os.replace(archivo_pickle, archivo_cache)
    with open(temporal_meta, "w", encoding="utf-8") as f:
        json.dump(huella_archivo(nombre_archivo), f)
    os.replace(temporal_meta, archivo_meta)
    print(f"⚡ Cache histórico registrado en {perf_counter() - inicio:.1f} s: {archivo_cache}")


def limpiar_caches_antiguos():
    if not CARPETA_CACHE.is_dir():
        return
    conservar = {f"fese{anio_mes_pasado}-{mes_pasado}", f"fese{anio_mes_anterior}-{mes_anterior}"}
    for archivo in CARPETA_CACHE.iterdir():
        if archivo.is_file() and archivo.name.startswith("fese") and archivo.name.split(".", 1)[0] not in conservar:
            archivo.unlink()


def cargar_y_limpiar(nombre_archivo="df_fese.rds"):
    nombre_archivo = Path(nombre_archivo)
    archivo_cache, archivo_meta = rutas_cache_historico(nombre_archivo)

    if cache_historico_valida(nombre_archivo, archivo_cache, archivo_meta):
        inicio = perf_counter()
        try:
            fese = pd.read_pickle(archivo_cache)
            fese = normalizar_historico(fese)
            print(f"⚡ Cache histórico cargado en {perf_counter() - inicio:.1f} s: {archivo_cache}")
            return fese
        except Exception as e:
            print(f"⚠ No se pudo usar el cache histórico; se leerá el RDS: {e}")

    fese = pyreadr.read_r(str(nombre_archivo))[None]
    fese = normalizar_historico(fese)

    try:
        guardar_cache_historico(fese, nombre_archivo)
    except Exception as e:
        print(f"⚠ No se pudo guardar el cache histórico; el proceso continúa: {e}")

    return fese


def obtener_insumos(path_insumos="insumos"):
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
                    datos_centro.rename(columns={datos_centro.columns[0]: "codigo", datos_centro.columns[1]: "incidente", datos_centro.columns[2]: "total"}, inplace=True)

                    datos_centro["incidente"] = datos_centro["incidente"].str.title().str.strip()
                    datos_centro["total"] = pd.to_numeric(datos_centro["total"], errors="coerce").fillna(0)
                    datos_centro = datos_centro[datos_centro["codigo"].isin(codigos)].copy()

                    datos_centro = datos_centro.assign(estado=nombre_entidad, centro=nombre_centro, ao=anio_mes_pasado, mes=mes_pasado, mes_largo=nombre_mes_pasado, fecha=fecha_mes_pasado.replace(day=1))
                    mapeo_estado = {"México": "MEXICO", "Michoacán": "MICHOACAN DE OCAMPO"}
                    datos_centro["estado"] = datos_centro["estado"].replace(mapeo_estado).str.upper()
                    datos_centro["procedencia"] = np.where(datos_centro["codigo"].isin(codigos_improcedentes), "Improcedentes", "Procedentes")
                    datos_centro["estado"] = datos_centro["estado"].str.title().replace(mapa_estados).str.upper()
                    datos_centro["incidente"] = datos_centro["codigo"].map(codigo_incidente)
                    datos_centro["region"] = datos_centro["estado"].map(estado_a_region).fillna("")
                    datos_centro["tipo"] = datos_centro["codigo"].map(codigo_a_tipo).fillna("")
                    datos_centro = datos_centro[columnas_long]

                    if nombre_centro not in centros_historicos:
                        validaciones.append({"clave_error": "centro_no_existe", "centro": nombre_centro, "entidad": nombre_entidad, "archivo": str(insumo), "hoja": hoja})

                    if not (datos_centro["total"] >= 0).all():
                        filas_con_error = list(datos_centro[datos_centro["total"] < 0]["incidente"])
                        validaciones.append({"clave_error": "numeros_negativos", "centro": nombre_centro, "entidad": nombre_entidad, "archivo": str(insumo), "hoja": hoja, "incidentes": filas_con_error})

                    datos_mes_pasado = datos_centro[(datos_centro["mes"] == mes_pasado) & (datos_centro["ao"] == anio_mes_pasado)][columnas_long]
                    datos_mes_anterior = datos_centro[(datos_centro["mes"] == mes_anterior) & (datos_centro["ao"] == mes_anterior)][columnas_long]
                    datos_ambos_meses = datos_mes_pasado.join(datos_mes_anterior, "codigo", lsuffix="_pasado", rsuffix="_anterior")

                    if (datos_ambos_meses["total_pasado"] == datos_ambos_meses["total_anterior"]).all():
                        validaciones.append({"clave_error": "mes_pasado_igual_anterior", "centro": nombre_centro, "entidad": nombre_entidad, "archivo": str(insumo), "hoja": hoja})

                    datos_anio_pasado = datos_centro[(datos_centro["mes"] == mes_pasado) & (datos_centro["ao"] == (anio_mes_pasado - 1))][columnas_long]
                    datos_ambos_meses = datos_mes_pasado.join(datos_anio_pasado, "codigo", lsuffix="_pasado", rsuffix="_anio_pasado")

                    if (datos_ambos_meses["total_pasado"] == datos_ambos_meses["total_anio_pasado"]).all():
                        validaciones.append({"clave_error": "mes_pasado_igual_anio_pasado", "centro": nombre_centro, "entidad": nombre_entidad, "archivo": str(insumo), "hoja": hoja})

                    if datos_centro["total"].sum() == 0:
                        validaciones.append({"clave_error": "suma_total_cero", "centro": nombre_centro, "entidad": nombre_entidad, "archivo": str(insumo), "hoja": hoja})

                    resultados.append(datos_centro)
            finally:
                workbook.close()

        print(f"⏱ Archivo terminado en {perf_counter() - inicio_archivo:.1f} s: {insumo.name}")

    return resultados, validaciones


def normalizar_clave_entidad(valor):
    clave = unidecode(str(valor)).strip().upper()
    equivalencias = {"COAHUILA DE ZARAGOZA": "COAHUILA", "ESTADO DE MEXICO": "MEXICO", "MICHOACAN DE OCAMPO": "MICHOACAN", "VERACRUZ DE IGNACIO DE LA LLAVE": "VERACRUZ"}
    return equivalencias.get(clave, clave)


def generar_formato_cniedt(salida):
    inicio_cniedt = perf_counter()

    if not PLANTILLA_CNIEDT.is_file():
        raise FileNotFoundError(f"No existe la plantilla CNIEDT: {PLANTILLA_CNIEDT}")

    datos = salida[(salida["ao"] == anio_mes_pasado) & (salida["mes"] == mes_pasado) & (salida["procedencia"] == "Procedentes")].copy()

    if datos.empty:
        raise ValueError(f"No hay datos procedentes para {nombre_mes_pasado} {anio_mes_pasado}.")

    datos["clave_entidad"] = datos["estado"].apply(normalizar_clave_entidad)
    totales = datos.groupby("clave_entidad")["total"].sum().to_dict()

    archivo_salida = CARPETA_DESTINO / f"Formato CNIEDT {anio_mes_pasado}-{str(mes_pasado).zfill(2)}.xlsx"
    archivo_temporal = CARPETA_DESTINO / f"~CNIEDT_TEMP_{anio_mes_pasado}-{str(mes_pasado).zfill(2)}.xlsx"

    if archivo_temporal.exists():
        archivo_temporal.unlink()

    shutil.copy2(PLANTILLA_CNIEDT, archivo_temporal)

    excel = None
    libro = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        libro = excel.Workbooks.Open(str(archivo_temporal))

        try:
            hoja = libro.Worksheets("Llamadas 911")
        except Exception as e:
            raise ValueError("No se encontró la pestaña 'Llamadas 911' en el formato CNIEDT.") from e

        columna_mes = None

        for columna in range(2, 14):
            encabezado = str(hoja.Cells(2, columna).Value or "").strip().lower()

            if encabezado == nombre_mes_pasado.lower():
                columna_mes = columna
                break

        if columna_mes is None:
            raise ValueError(f"No se encontró la columna del mes '{nombre_mes_pasado}' en la pestaña 'Llamadas 911'.")

        filas_objetivo = []
        entidades_formato = set()

        for fila in range(3, 35):
            entidad = hoja.Cells(fila, 1).Value

            if entidad is None:
                raise ValueError(f"La fila {fila} de 'Llamadas 911' no contiene una entidad.")

            clave = normalizar_clave_entidad(entidad)
            filas_objetivo.append((fila, clave, str(entidad).strip()))
            entidades_formato.add(clave)

        if len(filas_objetivo) != 32:
            raise ValueError(f"Se esperaban 32 entidades en 'Llamadas 911', pero se encontraron {len(filas_objetivo)}.")

        faltantes = entidades_formato - set(totales.keys())
        extras = set(totales.keys()) - entidades_formato

        if faltantes:
            raise ValueError(f"Faltan entidades en los datos FESE: {sorted(faltantes)}")

        if extras:
            raise ValueError(f"Hay entidades FESE que no coinciden con el formato CNIEDT: {sorted(extras)}")

        for fila, clave, entidad in filas_objetivo:
            hoja.Cells(fila, columna_mes).Value = float(totales[clave])

        excel.CalculateFull()

        total_fese = float(sum(totales.values()))
        total_formato = hoja.Cells(35, columna_mes).Value

        try:
            total_formato = float(total_formato)
        except (TypeError, ValueError):
            raise ValueError(f"El total calculado en la fila 35 para {nombre_mes_pasado} no es numérico: {total_formato}")

        if abs(total_formato - total_fese) > 0.01:
            raise ValueError(f"El total del formato CNIEDT ({total_formato:,.0f}) no coincide con el total FESE ({total_fese:,.0f}).")

        libro.Save()

    finally:
        if libro is not None:
            libro.Close(SaveChanges=True)

        if excel is not None:
            excel.Quit()

    try:
        os.replace(archivo_temporal, archivo_salida)
    except PermissionError as e:
        raise PermissionError(f"No se pudo reemplazar el CNIEDT porque el archivo destino probablemente está abierto o bloqueado: {archivo_salida}") from e

    print(f"✅ CNIEDT actualizado: {nombre_mes_pasado} {anio_mes_pasado} = {total_fese:,.0f}")
    print(f"⏱ Generación CNIEDT: {perf_counter() - inicio_cniedt:.1f} s")

    return archivo_salida


def preparar_fuente_rds(salida, archivo_rds):
    inicio = perf_counter()
    CARPETA_CACHE.mkdir(parents=True, exist_ok=True)
    archivo_pickle = CARPETA_CACHE / f"~RDS_SOURCE_{archivo_rds.stem}.pkl"

    if archivo_pickle.exists():
        archivo_pickle.unlink()

    salida.to_pickle(archivo_pickle)
    print(f"⏱ Preparación fuente RDS/cache: {perf_counter() - inicio:.1f} s")
    return archivo_pickle


def iniciar_escritura_rds_paralela(archivo_pickle, archivo_rds):
    archivo_temporal = archivo_rds.with_name(f"~TEMP_{archivo_rds.name}")
    if archivo_temporal.exists():
        archivo_temporal.unlink()

    codigo_worker = "import sys,pandas as pd,pyreadr; df=pd.read_pickle(sys.argv[1]); pyreadr.write_rds(sys.argv[2],df)"
    inicio = perf_counter()
    proceso = subprocess.Popen([sys.executable, "-c", codigo_worker, str(archivo_pickle), str(archivo_temporal)], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    print("🚀 Escritura RDS iniciada en paralelo con Excel anual")
    return proceso, archivo_temporal, inicio


def finalizar_escritura_rds_paralela(proceso, archivo_temporal, archivo_rds, inicio):
    stdout, stderr = proceso.communicate()

    if proceso.returncode != 0:
        if archivo_temporal.exists():
            archivo_temporal.unlink()
        detalle = stderr.strip() or stdout.strip() or f"código de salida {proceso.returncode}"
        raise RuntimeError(f"Falló la escritura paralela del RDS: {detalle}")

    if not archivo_temporal.is_file():
        raise RuntimeError(f"La escritura paralela terminó sin generar el RDS temporal: {archivo_temporal}")

    os.replace(archivo_temporal, archivo_rds)
    print(f"⏱ Escritura RDS paralela total: {perf_counter() - inicio:.1f} s")


def cancelar_escritura_rds(proceso, archivo_temporal):
    if proceso is not None and proceso.poll() is None:
        proceso.terminate()
        try:
            proceso.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proceso.kill()
            proceso.wait()

    if archivo_temporal is not None and archivo_temporal.exists():
        archivo_temporal.unlink()


def generar_salidas(fese, columnas_long, resultados, validaciones):
    inicio_total = perf_counter()

    inicio = perf_counter()
    resultados = [fese[columnas_long]] + resultados
    salida = pd.concat(resultados, ignore_index=True)
    del resultados
    print(f"⏱ Concatenación histórica: {perf_counter() - inicio:.1f} s")

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

    inicio = perf_counter()
    wide = salida[columnas_long].pivot(index=["ao", "region", "estado", "centro", "tipo", "codigo", "incidente", "procedencia"], columns="mes_largo", values="total")
    wide = wide.reset_index()
    print(f"⏱ Pivot anual: {perf_counter() - inicio:.1f} s")

    wide.estado = wide.estado.str.title().replace(mapa_estados).str.upper()
    wide.rename(columns={"codigo": "Código", "ao": "Año", "procedencia": "Procedencia", "region": "Region", "estado": "Estado", "centro": "Centro", "tipo": "Tipo", "incidente": "Incidente"}, inplace=True)

    columnas_wide = ["Código", "Año", "Procedencia", "Region", "Estado", "Centro", "Tipo", "Incidente", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    wide = wide[columnas_wide]
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

    for mes in meses:
        wide[mes] = wide[mes].fillna(0)

    wide = wide.sort_values(by=["Año", "Código"], ascending=[True, True])
    archivo_excel = CARPETA_SALIDAS / f"Rep_anual{anio_mes_pasado}-{str(mes_pasado).zfill(2)}.xlsx"

    salida.estado = salida.estado.str.title()
    salida.estado = salida.estado.apply(unidecode)
    salida.estado = salida.estado.replace(mapa_estados)

    archivo_rds = CARPETA_DATOS / f"fese{anio_mes_pasado}-{mes_pasado}.rds"
    archivo_pickle = preparar_fuente_rds(salida, archivo_rds)
    proceso_rds = None
    temporal_rds = None

    try:
        archivo_cniedt = generar_formato_cniedt(salida)
        generar_comprobantes(salida, CARPETA_DESTINO, PLANTILLA_COMPROBANTE, fecha_mes_pasado, fecha_mes_anterior)
        del salida

        proceso_rds, temporal_rds, inicio_rds = iniciar_escritura_rds_paralela(archivo_pickle, archivo_rds)

        inicio = perf_counter()
        wide.to_excel(archivo_excel, index=False)
        print(f"⏱ Escritura Excel anual: {perf_counter() - inicio:.1f} s")

        with open(CARPETA_SALIDAS / "validaciones.json", "w") as f:
            json.dump(validaciones, f)

        finalizar_escritura_rds_paralela(proceso_rds, temporal_rds, archivo_rds, inicio_rds)
        proceso_rds = None
        temporal_rds = None

        try:
            registrar_cache_desde_pickle(archivo_pickle, archivo_rds)
            limpiar_caches_antiguos()
        except Exception as e:
            print(f"⚠ No se pudo registrar el cache del nuevo histórico; el proceso continúa: {e}")
            if archivo_pickle.exists():
                archivo_pickle.unlink()

        inicio = perf_counter()
        copiar_resultados_y_limpiar(archivo_excel, archivo_rds, archivo_cniedt)
        print(f"⏱ Copias finales: {perf_counter() - inicio:.1f} s")
        print(f"⏱ generar_salidas TOTAL: {perf_counter() - inicio_total:.1f} s")
    except Exception:
        cancelar_escritura_rds(proceso_rds, temporal_rds)
        if archivo_pickle.exists():
            archivo_pickle.unlink()
        raise


def copiar_reemplazo_seguro(origen, destino):
    inicio = perf_counter()
    temporal = destino.with_name(f"~TEMP_{destino.name}")

    if temporal.exists():
        temporal.unlink()
    shutil.copy2(origen, temporal)

    try:
        os.replace(temporal, destino)
    except PermissionError as e:
        raise PermissionError(f"No se pudo reemplazar el archivo porque probablemente está abierto o bloqueado: {destino}") from e

    print(f"⏱ Copia {destino.name}: {perf_counter() - inicio:.1f} s")


def copiar_resultados_y_limpiar(archivo_excel, archivo_rds, archivo_cniedt):
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
    carpeta_formatos = CARPETA_DESTINO / f"Formatos {nombre_mes_pasado} {anio_mes_pasado}"

    if carpeta_formatos.exists():
        print(f"La carpeta de formatos ya existe, se omite la copia: {carpeta_formatos}")
    else:
        shutil.copytree(CARPETA_INSUMOS, carpeta_formatos)
        print(f"Formatos copiados a: {carpeta_formatos}")

    destino_excel = CARPETA_DESTINO / archivo_excel.name
    destino_rds = CARPETA_DESTINO / archivo_rds.name
    copiar_reemplazo_seguro(archivo_excel, destino_excel)
    copiar_reemplazo_seguro(archivo_rds, destino_rds)

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

    for archivo in CARPETA_DATOS.glob("fese*.rds"):
        coincidencia = patron_rds.match(archivo.name)
        if coincidencia and (int(coincidencia.group(1)), int(coincidencia.group(2))) not in meses_conservar:
            archivo.unlink()
            print(f"Histórico antiguo eliminado: {archivo.name}")

    for archivo in CARPETA_SALIDAS.glob("Rep_anual*.xlsx"):
        coincidencia = patron_excel.match(archivo.name)
        if coincidencia and (int(coincidencia.group(1)), int(coincidencia.group(2))) not in meses_conservar:
            archivo.unlink()
            print(f"Resultado antiguo eliminado: {archivo.name}")

    print("Carpeta insumos vaciada correctamente.")


if __name__ == "__main__":
    proceso_fese()
