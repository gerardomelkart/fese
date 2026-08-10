from pathlib import Path
import calendar
import datetime as dt
import re

import pandas as pd

from fese import cargar_y_limpiar
from modulos.comprobantes import MESES, TIPOS, NOMBRES_CORTOS, crear_pdf, nombre_seguro, normalizar_entidad, normalizar_tipo, sin_acentos


BASE_DIR = Path(__file__).resolve().parent
CARPETA_DATOS = BASE_DIR / "datos"
CARPETA_SALIDA = BASE_DIR / "comprobantes_pendientes"
PLANTILLA = BASE_DIR / "plantillas" / "COMPROBANTE FESE plantilla.pdf"


def buscar_historico_mas_reciente():
    patron = re.compile(r"^fese(\d{4})-(\d{1,2})\.rds$", re.IGNORECASE)
    archivos = []

    for archivo in CARPETA_DATOS.glob("fese*.rds"):
        coincidencia = patron.match(archivo.name)
        if coincidencia:
            archivos.append(((int(coincidencia.group(1)), int(coincidencia.group(2))), archivo))

    if not archivos:
        raise FileNotFoundError(f"No se encontraron históricos RDS en: {CARPETA_DATOS}")

    return max(archivos, key=lambda x: x[0])[1]


def preparar_datos(archivo_rds):
    print(f"\nCargando histórico: {archivo_rds.name}")
    datos = cargar_y_limpiar(archivo_rds)

    requeridas = {"ao", "mes", "estado", "tipo", "procedencia", "total"}

    if not requeridas.issubset(datos.columns):
        raise ValueError(f"Faltan columnas en el histórico: {sorted(requeridas - set(datos.columns))}")

    datos = datos[["ao", "mes", "estado", "tipo", "procedencia", "total"]].copy()
    datos["ao"] = pd.to_numeric(datos["ao"], errors="raise").astype(int)
    datos["mes"] = pd.to_numeric(datos["mes"], errors="raise").astype(int)
    datos["total"] = pd.to_numeric(datos["total"], errors="raise")
    datos["clave_entidad"] = datos["estado"].apply(normalizar_entidad)
    datos["procedencia_norm"] = datos["procedencia"].apply(lambda x: sin_acentos(x).strip().upper())
    datos["categoria"] = datos["tipo"].apply(normalizar_tipo)
    datos.loc[datos["procedencia_norm"] == "IMPROCEDENTES", "categoria"] = "IMPROCEDENTES"

    sin_categoria = sorted(datos[(datos["procedencia_norm"] != "IMPROCEDENTES") & datos["categoria"].isna() & (datos["total"] != 0)]["tipo"].dropna().astype(str).unique())

    if sin_categoria:
        raise ValueError(f"Hay tipos de incidente sin equivalencia para el comprobante: {sin_categoria}")

    return datos[datos["categoria"].notna()].copy()


def elegir_entidad(datos):
    entidades = datos.groupby("clave_entidad")["estado"].first().sort_index()
    opciones = []

    print("\nENTIDADES\n")

    for numero, (clave, entidad) in enumerate(entidades.items(), start=1):
        nombre = NOMBRES_CORTOS.get(clave, str(entidad).title())
        opciones.append((clave, nombre))
        print(f"{numero:2}. {nombre}")

    while True:
        valor = input("\nSelecciona entidad: ").strip()

        if valor.isdigit() and 1 <= int(valor) <= len(opciones):
            return opciones[int(valor) - 1]

        print("Opción inválida.")


def elegir_anio(datos, clave_entidad):
    anios = sorted(datos.loc[datos["clave_entidad"] == clave_entidad, "ao"].unique())
    anio_default = int(max(anios))

    print("\nAños disponibles:", ", ".join(str(int(x)) for x in anios))

    while True:
        valor = input(f"Año [{anio_default}]: ").strip()

        if not valor:
            return anio_default

        if valor.isdigit() and int(valor) in anios:
            return int(valor)

        print("Año inválido.")


def parsear_meses(valor):
    meses = set()

    for parte in valor.split(","):
        parte = parte.strip()

        if not parte:
            continue

        if "-" in parte:
            inicio, fin = parte.split("-", 1)

            if not inicio.strip().isdigit() or not fin.strip().isdigit():
                raise ValueError

            inicio = int(inicio)
            fin = int(fin)

            if inicio > fin:
                inicio, fin = fin, inicio

            meses.update(range(inicio, fin + 1))
        else:
            if not parte.isdigit():
                raise ValueError

            meses.add(int(parte))

    if not meses or any(mes < 1 or mes > 12 for mes in meses):
        raise ValueError

    return sorted(meses)


def elegir_meses(datos, clave_entidad, anio):
    disponibles = sorted(datos.loc[(datos["clave_entidad"] == clave_entidad) & (datos["ao"] == anio), "mes"].unique())

    print("\nMESES")
    print(" | ".join(f"{i}={MESES[i - 1]}" for i in range(1, 13)))
    print("\nCon datos:", ", ".join(MESES[int(mes) - 1] for mes in disponibles))

    while True:
        valor = input("\nMeses a generar (ej. 3,4,5,6 o 3-6): ").strip()

        try:
            meses = parsear_meses(valor)
        except ValueError:
            print("Formato inválido.")
            continue

        faltantes = [mes for mes in meses if mes not in disponibles]

        if faltantes:
            print("No existen datos para:", ", ".join(MESES[mes - 1] for mes in faltantes))
            continue

        return meses


def generar_comprobante(datos, clave_entidad, entidad, anio, mes):
    ultimo_dia = calendar.monthrange(anio, mes)[1]
    fecha_periodo = dt.date(anio, mes, ultimo_dia)
    fecha_anterior = fecha_periodo.replace(day=1) - dt.timedelta(days=1)

    actuales_df = datos[(datos["clave_entidad"] == clave_entidad) & (datos["ao"] == fecha_periodo.year) & (datos["mes"] == fecha_periodo.month)]
    anteriores_df = datos[(datos["clave_entidad"] == clave_entidad) & (datos["ao"] == fecha_anterior.year) & (datos["mes"] == fecha_anterior.month)]

    if actuales_df.empty:
        raise ValueError(f"No hay datos de {entidad} para {MESES[mes - 1]} {anio}.")

    if anteriores_df.empty:
        raise ValueError(f"No hay datos del mes anterior ({MESES[fecha_anterior.month - 1]} {fecha_anterior.year}) para calcular el comparativo.")

    periodo = pd.concat([actuales_df, anteriores_df], ignore_index=True)
    resumen = periodo.groupby(["ao", "mes", "categoria"])["total"].sum()

    actuales = [resumen.get((fecha_periodo.year, fecha_periodo.month, clave), 0) for clave, _ in TIPOS]
    anteriores = [resumen.get((fecha_anterior.year, fecha_anterior.month, clave), 0) for clave, _ in TIPOS]

    carpeta = CARPETA_SALIDA / nombre_seguro(clave_entidad)
    carpeta.mkdir(parents=True, exist_ok=True)

    archivo = carpeta / f"COMPROBANTE FESE {MESES[mes - 1].upper()} {anio} - {nombre_seguro(clave_entidad)}.pdf"

    crear_pdf(entidad, actuales, anteriores, fecha_periodo, fecha_anterior, PLANTILLA, archivo)

    print(f"Generado: {archivo.name}")


def main():
    if not PLANTILLA.is_file():
        raise FileNotFoundError(f"No existe la plantilla: {PLANTILLA}")

    archivo_rds = buscar_historico_mas_reciente()
    datos = preparar_datos(archivo_rds)

    clave_entidad, entidad = elegir_entidad(datos)
    anio = elegir_anio(datos, clave_entidad)
    meses = elegir_meses(datos, clave_entidad, anio)

    print(f"\nEntidad: {entidad}")
    print(f"Año: {anio}")
    print("Meses:", ", ".join(MESES[mes - 1] for mes in meses))

    confirmar = input("\nEscribe GENERAR para continuar: ").strip().upper()

    if confirmar != "GENERAR":
        print("Operación cancelada.")
        return

    print()

    for mes in meses:
        generar_comprobante(datos, clave_entidad, entidad, anio, mes)

    carpeta = CARPETA_SALIDA / nombre_seguro(clave_entidad)

    print("\nProceso terminado.")
    print(f"Carpeta: {carpeta}")


if __name__ == "__main__":
    main()