"""
============================================================
 EXTRAER BASE · SEGUIMIENTO DE CAMPO
============================================================
Lee el Excel de preliquidación de comisiones
(PreLiquidacion Mercaderista_Supervisor_Lider_<Mes>.xlsx) y deja
limpias sus 3 hojas:

    Indicadores                      -> campo_indicadores
    Liquidacion por Rutas            -> campo_liquidacion_rutas
    Liquidacion Lideres y Superviso  -> campo_liquidacion_lideres

Qué hace:
  - Encuentra el archivo más reciente dentro de "BASE DE DATOS".
  - Detecta la fila de encabezados aunque haya títulos arriba.
  - Resuelve columnas repetidas ("$ Liquidar", "Categoria Efec"...).
  - Convierte porcentajes a escala 0-100 (0.85 -> 85).
  - Normaliza el mes (texto "Agosto", fecha, 2026-08...) a:
        mes        -> "Agosto 2026"   (etiqueta para mostrar)
        mes_orden  -> 202608          (para ordenar y comparar)
  - NO toma columnas de dinero: el dashboard trabaja solo con
    porcentajes de cumplimiento.

Uso:
  - Ejecutarlo directo genera CSV + archivo de control para revisar.
  - sincronizar_supabase.py lo importa para subir los datos.
============================================================
"""

import re
import sys
import unicodedata
from datetime import datetime, date
from pathlib import Path

import pandas as pd


# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA_BASE = Path(__file__).resolve().parent / "BASE DE DATOS"

# Si la carpeta no está junto al .py, pon la ruta completa:
# CARPETA_BASE = Path(r"C:\Users\Dylan\Desktop\BASE DE DATOS")

# Patrón del archivo de preliquidación (se toma el más reciente)
PATRON_ARCHIVO = "PreLiquidacion*.xls*"

# Año que se usa cuando la columna MES solo dice "Agosto"
ANIO_POR_DEFECTO = datetime.now().year

# Si el máximo de una columna de % es <= a este valor, se asume que
# viene como fracción (0.85) y se multiplica por 100.
UMBRAL_FRACCION = 5

CARPETA_CSV = CARPETA_BASE / "CSV_EXTRAIDOS"
ARCHIVO_CONTROL = CARPETA_BASE / "CONTROL_EXTRACCION.xlsx"


# ============================================================
# MESES
# ============================================================

MESES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
    "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
    "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10,
    "NOVIEMBRE": 11, "DICIEMBRE": 12,
}

NOMBRES_MES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


# ============================================================
# ESTRUCTURA DE LAS 3 HOJAS
# ============================================================
# Cada columna: (ENCABEZADO NORMALIZADO, columna_supabase, tipo)
#
# Encabezado normalizado = sin tildes, en mayúsculas y con un solo
# espacio. Si un encabezado se repite en la hoja, la 2da aparición
# se llama "NOMBRE #2", la 3ra "NOMBRE #3", etc.
#
# Tipos: "texto" | "pct" | "numero" | "mes"
# ============================================================

HOJAS = {

    "INDICADORES": {
        "hoja": "Indicadores",
        "tabla": "campo_indicadores",
        "clave": "RUTA",
        "columnas": [
            ("MES", "mes", "mes"),
            ("SUPERVISOR", "supervisor", "texto"),
            ("RUTA", "ruta", "texto"),
            ("% CAPTURA", "pct_captura", "pct"),
            ("OBJ COBERTURA", "obj_cobertura", "numero"),
            ("COBERTURA", "cobertura", "numero"),
            ("% COBERTURA", "pct_cobertura", "pct"),
            ("OBJ VISITAS", "obj_visitas", "numero"),
            ("VISITAS", "visitas", "numero"),
            ("% VISITAS", "pct_visitas", "pct"),
            ("FECHAS CORTAS %", "pct_fechas_cortas", "pct"),
            ("AGOTADOS %", "pct_agotados", "pct"),
            ("PARTICIPACION %", "pct_participacion", "pct"),
            ("ACTIVIDADES %", "pct_actividades", "pct"),
            ("PLANES COMERCIALES", "pct_planes_comerciales", "pct"),
            ("DEVOLUCIONES %", "pct_devoluciones", "pct"),
            ("PDV DEVOLUCIONES", "pdv_devoluciones", "numero"),
            ("% OBJETIVO DEVOLUCIONES", "pct_objetivo_devoluciones", "pct"),
        ],
    },

    "RUTAS": {
        "hoja": "Liquidacion por Rutas",
        "tabla": "campo_liquidacion_rutas",
        "clave": "RUTA",
        "columnas": [
            ("MES", "mes", "mes"),
            ("REGIONAL", "regional", "texto"),
            ("RUTA", "ruta", "texto"),
            ("LIDER ALQUERIA", "lider_alqueria", "texto"),
            ("LIDER", "lider", "texto"),
            ("SUPERVISOR", "supervisor", "texto"),
            ("% CUMPL. CAPTURA", "pct_captura", "pct"),
            ("% CUMPLIMIENTO COBERTURA", "pct_cobertura", "pct"),
            ("PROMEDIO LLAVES", "pct_promedio_llaves", "pct"),
            ("% CUMPLIMIENTO TP", "pct_tienda_perfecta", "pct"),
            ("% CUMPLIMIENTO VENTAS", "pct_ventas", "pct"),
            ("% CUMPLIMIENTO FC & DV", "pct_fc_dv", "pct"),
            ("CORRESPONDE COMISION", "corresponde_comision", "texto"),
            ("OBSERVACIONES", "observaciones", "texto"),
        ],
    },

    "LIDERES": {
        "hoja": "Liquidacion Lideres y Superviso",
        "tabla": "campo_liquidacion_lideres",
        "clave": "NOMBRE",
        "columnas": [
            ("MES", "mes", "mes"),
            ("CARGO", "cargo", "texto"),
            ("NOMBRE", "nombre", "texto"),
            ("LLAVE 1 CAPTURA", "pct_llave_captura", "pct"),
            ("LLAVE 2 COBERTURA", "pct_llave_cobertura", "pct"),
            ("PROM. LLAVE", "pct_promedio_llaves", "pct"),
            ("% CUMPL. TP", "pct_tienda_perfecta", "pct"),
            ("% VENTAS", "pct_ventas", "pct"),
            ("% FC & DV", "pct_fc_dv", "pct"),
            ("CUMPLE LLAVES", "cumple_llaves", "texto"),
            ("COMENTARIO", "comentario", "texto"),
        ],
    },
}


# ============================================================
# UTILIDADES DE TEXTO
# ============================================================

def normalizar(texto):
    """Quita tildes, pasa a mayúsculas y deja un solo espacio."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    texto = str(texto).replace("\n", " ")
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return " ".join(texto.upper().split())


def desduplicar(encabezados):
    """Renombra encabezados repetidos: X, X #2, X #3..."""
    vistos = {}
    salida = []
    for nombre in encabezados:
        base = nombre or "SIN NOMBRE"
        vistos[base] = vistos.get(base, 0) + 1
        salida.append(base if vistos[base] == 1 else f"{base} #{vistos[base]}")
    return salida


def limpiar_texto(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    texto = " ".join(str(valor).split())
    return texto or None


# ============================================================
# CONVERSIÓN DE NÚMEROS Y PORCENTAJES
# ============================================================

def convertir_numero(valor):
    """
    Convierte números con formato colombiano o inglés.
    Si el texto trae '%', devuelve la fracción (85% -> 0.85).
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    if texto in ("", "-", "#N/A", "#DIV/0!", "#VALUE!", "#REF!"):
        return None

    es_porcentaje = "%" in texto
    texto = texto.replace("%", "").replace("$", "").replace(" ", "")

    try:
        if "," in texto and "." in texto:
            if texto.rfind(",") > texto.rfind("."):
                texto = texto.replace(".", "").replace(",", ".")
            else:
                texto = texto.replace(",", "")
        elif "," in texto:
            texto = texto.replace(",", ".")

        numero = float(texto)
        return numero / 100 if es_porcentaje else numero
    except ValueError:
        return None


def columna_a_porcentaje(serie):
    """
    Deja toda la columna en escala 0-100.
    Excel entrega los % como fracción (0.85); si la columna ya
    viene en 0-100 (85) se respeta.
    """
    numeros = serie.apply(convertir_numero)
    maximo = numeros.abs().max(skipna=True)
    if pd.notna(maximo) and maximo <= UMBRAL_FRACCION:
        numeros = numeros * 100
    return numeros.round(2)


# ============================================================
# MES
# ============================================================

def mes_desde_nombre_archivo(archivo):
    """Busca 'Agosto' (u otro mes) en el nombre del archivo."""
    nombre = normalizar(archivo.stem)
    for texto_mes, numero in MESES.items():
        if texto_mes in nombre:
            anio = re.search(r"(20\d{2})", nombre)
            return (int(anio.group(1)) if anio else ANIO_POR_DEFECTO), numero
    return None


def interpretar_mes(valor, respaldo):
    """
    Devuelve (anio, mes) a partir de lo que venga en la columna MES.
    Acepta fechas, 'Agosto', 'AGOSTO 2026', '2026-08', '08/2026' u 8.
    """
    if isinstance(valor, (datetime, date, pd.Timestamp)):
        return valor.year, valor.month

    if isinstance(valor, (int, float)) and not pd.isna(valor):
        numero = int(valor)
        if 1 <= numero <= 12:
            return ANIO_POR_DEFECTO, numero
        if numero > 30000:  # número de serie de fecha de Excel
            fecha = pd.to_datetime("1899-12-30") + pd.to_timedelta(numero, unit="D")
            return fecha.year, fecha.month

    texto = normalizar(valor)
    if texto:
        for texto_mes, numero in MESES.items():
            if texto_mes in texto:
                anio = re.search(r"(20\d{2})", texto)
                return (int(anio.group(1)) if anio else ANIO_POR_DEFECTO), numero

        iso = re.match(r"^(20\d{2})[-/](\d{1,2})", texto)
        if iso:
            return int(iso.group(1)), int(iso.group(2))

        mes_anio = re.match(r"^(\d{1,2})[-/](20\d{2})$", texto)
        if mes_anio:
            return int(mes_anio.group(2)), int(mes_anio.group(1))

        fecha = pd.to_datetime(texto, errors="coerce", dayfirst=True)
        if pd.notna(fecha):
            return fecha.year, fecha.month

    return respaldo


# ============================================================
# ARCHIVO Y HOJAS
# ============================================================

def buscar_archivo():
    """Toma el archivo de preliquidación modificado más recientemente."""
    candidatos = [
        a for a in CARPETA_BASE.glob(PATRON_ARCHIVO)
        if a.is_file() and not a.name.startswith("~$")
    ]
    if not candidatos:
        return None
    return max(candidatos, key=lambda a: a.stat().st_mtime)


def buscar_hoja(nombres_hojas, nombre_esperado):
    """
    Encuentra la hoja aunque cambien tildes, mayúsculas o esté
    truncada (Excel corta los nombres a 31 caracteres).
    """
    esperado = normalizar(nombre_esperado)
    for hoja in nombres_hojas:
        actual = normalizar(hoja)
        if actual == esperado or actual.startswith(esperado) or esperado.startswith(actual):
            return hoja
    return None


def leer_hoja(archivo, nombre_hoja, clave):
    """
    Lee la hoja sin asumir que el encabezado está en la fila 1:
    busca en las primeras 15 filas la que contiene la columna clave.
    """
    crudo = pd.read_excel(archivo, sheet_name=nombre_hoja, header=None, dtype=object)

    fila_encabezado = None
    for i in range(min(15, len(crudo))):
        valores = [normalizar(v) for v in crudo.iloc[i].tolist()]
        if clave in valores:
            fila_encabezado = i
            break

    if fila_encabezado is None:
        raise ValueError(f'No se encontró la columna "{clave}" en las primeras 15 filas.')

    encabezados = desduplicar([normalizar(v) for v in crudo.iloc[fila_encabezado].tolist()])
    datos = crudo.iloc[fila_encabezado + 1:].copy()
    datos.columns = encabezados
    datos = datos.dropna(how="all").reset_index(drop=True)
    return datos


# ============================================================
# LIMPIEZA DE UNA HOJA
# ============================================================

def preparar_hoja(datos, config, mes_respaldo):
    """Selecciona, convierte y renombra las columnas de una hoja."""
    esperadas = [col for col, _, _ in config["columnas"]]
    faltantes = [col for col in esperadas if col not in datos.columns]
    avisos = []

    if faltantes:
        avisos.append("Columnas faltantes (quedan vacías): " + ", ".join(faltantes))

    salida = pd.DataFrame(index=datos.index)

    for encabezado, destino, tipo in config["columnas"]:
        serie = datos[encabezado] if encabezado in datos.columns else pd.Series([None] * len(datos), index=datos.index)

        if tipo == "pct":
            salida[destino] = columna_a_porcentaje(serie)
        elif tipo == "numero":
            salida[destino] = serie.apply(convertir_numero)
        elif tipo == "mes":
            partes = serie.apply(lambda v: interpretar_mes(v, mes_respaldo))
            salida["mes_orden"] = partes.apply(lambda p: p[0] * 100 + p[1] if p else None)
            salida["mes"] = partes.apply(lambda p: f"{NOMBRES_MES[p[1]]} {p[0]}" if p else None)
        else:
            salida[destino] = serie.apply(limpiar_texto)

    # Quitar filas sin clave y filas de totales
    clave_destino = next(d for c, d, _ in config["columnas"] if c == config["clave"])
    salida = salida[salida[clave_destino].notna()]
    salida = salida[~salida[clave_destino].str.upper().str.contains("TOTAL", na=False)]

    sin_mes = salida["mes_orden"].isna().sum()
    if sin_mes:
        avisos.append(f"{sin_mes} filas sin mes identificable (se descartan).")
        salida = salida[salida["mes_orden"].notna()]

    salida["mes_orden"] = salida["mes_orden"].astype(int)
    salida = salida.astype(object).where(pd.notna(salida), None)
    return salida.reset_index(drop=True), avisos


# ============================================================
# FUNCIÓN PRINCIPAL (la usa sincronizar_supabase.py)
# ============================================================

def cargar_preliquidacion(archivo=None, mostrar=True):
    """
    Devuelve (archivo, {"INDICADORES": df, "RUTAS": df, "LIDERES": df}).
    Los DataFrames ya tienen los nombres de columna de Supabase.
    """
    def log(msg=""):
        if mostrar:
            print(msg)

    if not CARPETA_BASE.exists():
        raise FileNotFoundError(f"No existe la carpeta: {CARPETA_BASE}")

    archivo = Path(archivo) if archivo else buscar_archivo()
    if archivo is None:
        raise FileNotFoundError(
            f'No hay ningún archivo "{PATRON_ARCHIVO}" en {CARPETA_BASE}'
        )

    log(f"\n📄 Archivo: {archivo.name}")
    mes_respaldo = mes_desde_nombre_archivo(archivo)
    if mes_respaldo:
        log(f"🗓  Mes de respaldo (nombre del archivo): {NOMBRES_MES[mes_respaldo[1]]} {mes_respaldo[0]}")

    nombres_hojas = pd.ExcelFile(archivo).sheet_names
    resultado = {}

    for clave, config in HOJAS.items():
        log("\n" + "=" * 70)
        log(f" HOJA: {config['hoja']}  ->  {config['tabla']}")
        log("=" * 70)

        hoja = buscar_hoja(nombres_hojas, config["hoja"])
        if hoja is None:
            log(f"❌ No se encontró la hoja. Hojas disponibles: {', '.join(nombres_hojas)}")
            continue

        try:
            datos = leer_hoja(archivo, hoja, config["clave"])
            df, avisos = preparar_hoja(datos, config, mes_respaldo)
        except Exception as error:
            log(f"❌ Error leyendo la hoja: {error}")
            continue

        for aviso in avisos:
            log(f"⚠️  {aviso}")

        meses = sorted(df["mes"].dropna().unique().tolist()) if len(df) else []
        log(f"✅ Registros: {len(df):,}   Meses: {', '.join(meses) or '—'}")
        resultado[clave] = df

    return archivo, resultado


# ============================================================
# EJECUCIÓN DIRECTA: genera CSV + control para revisar
# ============================================================

def main():
    print("\n" + "=" * 70)
    print(" EXTRACTOR · PRELIQUIDACIÓN DE CAMPO")
    print("=" * 70)
    print(f"\n📁 Carpeta: {CARPETA_BASE}")

    try:
        archivo, bases = cargar_preliquidacion()
    except Exception as error:
        print(f"\n❌ {error}")
        input("\nPresiona ENTER para salir...")
        sys.exit(1)

    CARPETA_CSV.mkdir(exist_ok=True)
    control = []

    print("\n" + "=" * 70)
    print(" GENERANDO CSV")
    print("=" * 70)

    for clave, config in HOJAS.items():
        if clave not in bases:
            control.append({"HOJA": config["hoja"], "TABLA": config["tabla"], "REGISTROS": 0, "ESTADO": "NO CARGADA"})
            print(f"❌ {config['tabla']:<28} NO CARGADA")
            continue

        df = bases[clave]
        ruta_csv = CARPETA_CSV / f"{config['tabla']}.csv"
        df.to_csv(ruta_csv, index=False, encoding="utf-8-sig")
        control.append({"HOJA": config["hoja"], "TABLA": config["tabla"], "REGISTROS": len(df), "ESTADO": "OK"})
        print(f"✅ {config['tabla']:<28} {len(df):>8,} registros")

    pd.DataFrame(control).to_excel(ARCHIVO_CONTROL, index=False, sheet_name="CONTROL")

    print("\n" + "=" * 70)
    print(" PROCESO TERMINADO")
    print("=" * 70)
    print(f"\n📁 CSV:     {CARPETA_CSV}")
    print(f"📋 Control: {ARCHIVO_CONTROL}\n")
    input("Presiona ENTER para cerrar...")


if __name__ == "__main__":
    main()