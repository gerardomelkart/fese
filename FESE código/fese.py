from pathlib import Path
import datetime as dt
import locale
import shutil
import json
import re

import pandas as pd
import openpyxl
import numpy as np
import pyreadr
from unidecode import unidecode

BASE_DIR = Path(__file__).resolve().parent
CARPETA_INSUMOS = BASE_DIR / "insumos"
CARPETA_DESTINO = Path(r"C:\Users\gerardo.noeller\OneDrive - Secretaría de Seguridad y Protección Ciudadana\Escritorio\FESE")


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
    fese = cargar_y_limpiar(BASE_DIR / f"fese{anio_mes_anterior}-{mes_anterior}.rds")
    insumos = obtener_insumos(CARPETA_INSUMOS)
    resultados, validaciones = leer_insumos(fese, insumos)
    generar_salidas(fese, columnas_long, resultados, validaciones)


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
    # Creamos una lista de validaciones que no pasaron. En caso de haber alguna validación fallida se agrega como diccionario a esta lista
    validaciones = []

    resultados = []
    # Iteramos sobre cada archivo insumo
    for insumo in insumos:
        # Abrimos el archivo excel como un archivo (NO UN DATAFRAME) para leer metadatos sobre las hojas
        with pd.ExcelFile(insumo) as xlsx:
            # Obtenemos los nombres de las hojas que empiezan con C0 (estas hojas contienen los datos)
            hojas_con_datos = [
                hoja for hoja in xlsx.sheet_names if hoja.startswith("C")
            ]
            # Iteramos sobre las hojas
            for hoja in hojas_con_datos:
                # Conseguimos el nombre del centro que se encuentra en la celda D18
                workbook = openpyxl.load_workbook(insumo, data_only=True)
                hoja_centro = workbook[hoja]
                # Agarra el nombre de centro y de entidad de las celdas hardcodeadas
                nombre_centro = hoja_centro["D18"].value.strip()
                nombre_entidad = hoja_centro["D12"].value
                # Si el nombre de entidad está vacío entonces intenta en otras celdas
                if nombre_entidad is None:
                    nombre_entidad = workbook["RESUMEN"]["C10"].value
                if nombre_entidad is None:
                    nombre_entidad = workbook["RESUMEN"]["C11"].value
                nombre_entidad = nombre_entidad.strip()
                print(f"Trabajando: {nombre_entidad}, {nombre_centro}, hoja={hoja}, archivo={insumo}")
                workbook.close()
                # Cargamos la hoja en un dataframe usando las columnas y filas correctas
                datos_centro = pd.read_excel(
                    xlsx, sheet_name=hoja, usecols="B:D", skiprows=18
                )
                # Lista de código históricos
                codigos = fese["codigo"].unique()
                # Renombramos la columna Cantidad por total y la columna ID por codigo
                datos_centro.rename(
                    columns={
                        datos_centro.columns[0]: "codigo",
                        datos_centro.columns[1]: "incidente",
                        datos_centro.columns[2]: "total",
                    },
                    inplace=True,
                )
                # Pasar la primera letra de cada centro a mayúscula y hacerle strip
                datos_centro["incidente"] = (
                    datos_centro["incidente"].str.title().str.strip()
                )
                datos_centro["incidente"] = datos_centro["incidente"].str.title().str.strip()
                datos_centro["total"] = pd.to_numeric(datos_centro["total"], errors="coerce").fillna(0)

                # Quitamos todas las filas que no correspondan a un código existente
                # Esto se debe a que a veces ponen filas de "Total"
                datos_centro = datos_centro.drop(datos_centro[~datos_centro["codigo"].isin(codigos)].index)
                # Agregamos columna con el nombre del centro y con el nombre de la entidad federativa
                datos_centro = datos_centro.assign(
                    estado=nombre_entidad.strip(),
                    centro=nombre_centro.strip(),
                    ao=anio_mes_pasado,
                    mes=mes_pasado,
                    mes_largo=nombre_mes_pasado,
                    fecha=fecha_mes_pasado.replace(day=1),
                )
                mapeo_estado = {
                    "México": "MEXICO",
                    "Michoacán": "MICHOACAN DE OCAMPO",
                }
                datos_centro.estado = datos_centro.estado.replace(mapeo_estado)
                datos_centro.estado = datos_centro.estado.str.upper()
                # Si tiene código improcedente entonces ponemos el valor "Improcedentes" en procedencia. De lo contrario ponemos "Procedentes"
                codigos_improcedentes = [
                    70101,
                    70102,
                    70103,
                    70104,
                    70105,
                    70106,
                    70107,
                    70108,
                ]
                datos_centro.loc[
                    datos_centro["codigo"].isin(codigos_improcedentes), "procedencia"
                ] = "Improcedentes"
                datos_centro.loc[
                    ~(datos_centro["codigo"].isin(codigos_improcedentes)), "procedencia"
                ] = "Procedentes"
                datos_centro["region"] = ""
                datos_centro["tipo"] = ""
                datos_centro.estado = (
                    datos_centro.estado.str.title().replace(mapa_estados).str.upper()
                )
                # Mapeamos los incidentes según su código. Esto se debe a que los incidentes están diferentes en el reporte y en los insumos.
                codigo_incidente = (
                    fese.groupby("codigo")["incidente"]
                    .agg(lambda x: list(x.unique())[0])
                    .to_dict()
                )
                datos_centro.incidente = datos_centro.codigo.replace(codigo_incidente)
                # Mapeamos los estados según su región
                estado_region = (
                    fese.groupby("region")["estado"]
                    .agg(lambda x: list(x.unique()))
                    .to_dict()
                )
                for key, value in estado_region.items():
                    datos_centro.loc[datos_centro["estado"].isin(value), "region"] = key
                # Asignamos los tipo según el código
                codigo_tipo = (
                    fese.groupby("tipo")["codigo"]
                    .agg(lambda x: list(x.unique()))
                    .to_dict()
                )
                for key, value in codigo_tipo.items():
                    datos_centro.loc[datos_centro["codigo"].isin(value), "tipo"] = key
                #
                # Obtenemos un dataframe sólo con las columnas útiles y en el orden correcto
                datos_centro = datos_centro[columnas_long]

                # Hacemos las validaciones y agregamos los errores a la lista de validaciones
                if nombre_centro not in fese["centro"].unique():
                    # Si el centro no existe en el reporte anual entonces agrega un error
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
                    # Si algún valor de llamadas es negativo entonces agrega un error
                    filas_con_error = list(
                        datos_centro[datos_centro["total"] < 0]["incidente"]
                    )
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
                # Hacemos un join de los datos del mes pasado y del anterior para comparar las cantidades
                columnas_validacion = ["codigo", "ao", "total", "mes"]
                datos_mes_pasado = datos_centro[
                    (datos_centro["mes"] == mes_pasado)
                    & (datos_centro["ao"] == anio_mes_pasado)
                ][columnas_long]
                datos_mes_anterior = datos_centro[
                    (datos_centro["mes"] == mes_anterior)
                    & (datos_centro["ao"] == mes_anterior)
                ][columnas_long]
                datos_ambos_meses = datos_mes_pasado.join(
                    datos_mes_anterior, "codigo", lsuffix="_pasado", rsuffix="_anterior"
                )
                if (
                    datos_ambos_meses["total_pasado"]
                    == datos_ambos_meses["total_anterior"]
                ).all():
                    # Si los valores del mes pasado son iguales a los del mes anterior entonces agrega un error
                    validaciones.append(
                        {
                            "clave_error": "mes_pasado_igual_anterior",
                            "centro": nombre_centro,
                            "entidad": nombre_entidad,
                            "archivo": str(insumo),
                            "hoja": hoja,
                        }
                    )
                # Hacemos un join de los datos del mes pasado con los del año anterior en ese mismo mes
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
                    # Si los valores del mes pasado son iguales a los del año pasado en el mismo mes entonces agrega un error
                    validaciones.append(
                        {
                            "clave_error": "mes_pasado_igual_anio_pasado",
                            "centro": nombre_centro,
                            "entidad": nombre_entidad,
                            "archivo": str(insumo),
                            "hoja": hoja,
                        }
                    )
                if datos_centro.total.sum() == 0:
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
    return resultados, validaciones


def generar_salidas(fese, columnas_long, resultados, validaciones):
    resultados = [fese[columnas_long]] + resultados

    salida = pd.concat(resultados)

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

    copiar_resultados_y_limpiar(archivo_excel, archivo_rds)


def copiar_resultados_y_limpiar(archivo_excel, archivo_rds):
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
    carpeta_formatos = CARPETA_DESTINO / f"Formatos {nombre_mes_pasado} {anio_mes_pasado}"

    if carpeta_formatos.exists():
        shutil.rmtree(carpeta_formatos)

    shutil.copytree(CARPETA_INSUMOS, carpeta_formatos)
    destino_excel = CARPETA_DESTINO / archivo_excel.name
    destino_rds = CARPETA_DESTINO / archivo_rds.name
    shutil.copy2(archivo_excel, destino_excel)
    shutil.copy2(archivo_rds, destino_rds)

    if not destino_excel.is_file() or not destino_rds.is_file() or not carpeta_formatos.is_dir():
        raise RuntimeError("No se completó correctamente la copia a OneDrive. No se limpiarán los archivos originales.")

    print(f"Excel copiado a: {destino_excel}")
    print(f"RDS copiado a: {destino_rds}")
    print(f"Formatos copiados a: {carpeta_formatos}")

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
