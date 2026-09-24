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

Además, los valores en dinero de las hojas de liquidación se separan en
dos tablas protegidas que solo ven los usuarios con login:

    Liquidacion por Rutas            -> campo_comision_rutas
    Liquidacion Lideres y Superviso  -> campo_comision_lideres

Qué hace:
  - Procesa TODOS los Excel de preliquidación que haya en "BASE DE DATOS"
    (cualquier .xlsx/.xls cuyo nombre contenga "PreLiquidacion").
  - Detecta la fila de encabezados aunque haya títulos arriba.
  - Resuelve columnas repetidas ("$ Liquidar", "Categoria Efec"...).
  - Convierte porcentajes a escala 0-100 (0.85 -> 85).
  - Normaliza el mes (texto "Agosto", fecha, 2026-08...) a:
        mes        -> "Agosto 2026"   (etiqueta para mostrar)
        mes_orden  -> 202608          (para ordenar y comparar)
  - Las tablas públicas solo llevan porcentajes; el dinero va aparte.
  - Guarda "fila_excel" en ambas tablas para unirlas fila a fila.

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

# Se procesan todos los Excel cuyo nombre contenga este texto
# (sin importar mayúsculas, tildes, espacios o guiones).
TEXTO_EN_NOMBRE = "PRELIQUIDACION"
EXTENSIONES = (".xlsx", ".xlsm", ".xls")

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
        "detectar_umbral": True,
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
        # Valores en dinero -> tabla protegida (solo usuarios con login)
        "dinero": {
            "tabla": "campo_comision_rutas",
            "identidad": ["regional", "ruta", "lider", "supervisor"],
            "columnas": [
                ("VALOR COMISION", "valor_comision", "numero"),
                ("VALOR COMISION TP", "base_tp", "numero"),
                ("$ LIQUIDAR", "liquidado_tp", "numero"),
                ("VALOR COMISION VENTAS", "base_ventas", "numero"),
                ("$ LIQUIDAR #2", "liquidado_ventas", "numero"),
                ("VALOR COMISION FC Y DEV", "base_fc_dv", "numero"),
                ("$ LIQUIDAR #3", "liquidado_fc_dv", "numero"),
                ("VALOR TOTAL A PAGAR REAL", "valor_pagar", "numero"),
            ],
        },
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
        # En esta hoja "VALOR COMISION" se repite: #1 = TP liquidado,
        # #2 = Ventas liquidado, #3 = FC & DV liquidado, #4 = base total.
        "dinero": {
            "tabla": "campo_comision_lideres",
            "identidad": ["cargo", "nombre"],
            "columnas": [
                ("VALOR COMISION #4", "valor_comision", "numero"),
                ("VALOR COMISION TP", "base_tp", "numero"),
                ("VALOR COMISION", "liquidado_tp", "numero"),
                ("VALOR COMISION VENTAS", "base_ventas", "numero"),
                ("VALOR COMISION #2", "liquidado_ventas", "numero"),
                ("VALOR COMISION X FC DV", "base_fc_dv", "numero"),
                ("VALOR COMISION #3", "liquidado_fc_dv", "numero"),
                ("VALOR A PAGAR", "valor_pagar", "numero"),
            ],
        },
    },
}

# Orden de carga: (clave en el resultado, tabla de Supabase)
TABLAS = [
    ("INDICADORES", "campo_indicadores"),
    ("RUTAS", "campo_liquidacion_rutas"),
    ("RUTAS_DINERO", "campo_comision_rutas"),
    ("LIDERES", "campo_liquidacion_lideres"),
    ("LIDERES_DINERO", "campo_comision_lideres"),
]


# Columnas que no todos los meses traen: si faltan no se avisa
COLUMNAS_OPCIONALES = {"LIDER ALQUERIA", "PROMEDIO LLAVES", "OBSERVACIONES", "COMENTARIO", "PROM. LLAVE"}


# ============================================================
# NOMBRES ALTERNATIVOS DE COLUMNAS
# ============================================================
# Si un mes el Excel trae el encabezado escrito distinto, agrégalo aquí.
# (Se comparan sin tildes, mayúsculas, espacios ni puntos.)

ALIAS_COLUMNAS = {
    "SUPERVISOR": ["NOMBRE SUPERVISOR", "SUPERVISOR ALQUERIA", "SUPERVISORA", "SUPERVISOR EFICACIA"],
    "LIDER": ["LIDER EFICACIA", "LIDER DE EJECUCION", "NOMBRE LIDER"],
    "RUTA": ["RUTAS", "NOMBRE RUTA", "COD RUTA"],
    "% CUMPL. CAPTURA": ["% CUMPLIMIENTO CAPTURA", "% CAPTURA"],
    "% CUMPLIMIENTO COBERTURA": ["% CUMPL. COBERTURA", "% COBERTURA"],
    "% CUMPLIMIENTO TP": ["% CUMPL. TP", "% TP", "% TIENDA PERFECTA", "% CUMPLIMIENTO TIENDA PERFECTA"],
    "% CUMPLIMIENTO VENTAS": ["% CUMPL. VENTAS", "% VENTAS", "% CUMPLIMIENTO VENTA"],
    "% CUMPLIMIENTO FC & DV": ["% CUMPL. FC & DV", "% FC & DV", "% CUMPLIMIENTO FC Y DV", "% CUMPLIMIENTO FC & DEV"],
    "% CUMPL. TP": ["% CUMPLIMIENTO TP", "% TP"],
    "% VENTAS": ["% CUMPLIMIENTO VENTAS", "% CUMPL. VENTAS"],
    "% FC & DV": ["% CUMPLIMIENTO FC & DV", "% CUMPL. FC & DV", "% FC Y DV"],
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
    numeros = pd.to_numeric(serie.apply(convertir_numero), errors="coerce")
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

def es_preliquidacion(archivo):
    nombre = re.sub(r"[^A-Z0-9]", "", normalizar(archivo.stem))
    return TEXTO_EN_NOMBRE in nombre


def buscar_archivos():
    """
    Devuelve (preliquidaciones, otros_excel).
    Las preliquidaciones van ordenadas de la más antigua a la más reciente
    (por mes del nombre y luego por fecha de modificación), así, si dos
    archivos traen el mismo mes, gana el más reciente.
    """
    excels = [
        a for a in CARPETA_BASE.iterdir()
        if a.is_file() and a.suffix.lower() in EXTENSIONES and not a.name.startswith("~$")
    ]
    preliq = [a for a in excels if es_preliquidacion(a)]
    otros = [a for a in excels if a not in preliq]

    def orden(a):
        mes = mes_desde_nombre_archivo(a)
        return ((mes[0] * 100 + mes[1]) if mes else 0, a.stat().st_mtime)

    return sorted(preliq, key=orden), otros


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
    datos = datos.dropna(how="all")
    # Número de fila real en Excel (1 = primera fila de la hoja)
    datos["__FILA_EXCEL"] = datos.index + 1
    return datos.reset_index(drop=True)


# ============================================================
# LIMPIEZA DE UNA HOJA
# ============================================================

def clave_columna(nombre):
    """Encabezado sin espacios ni puntuación: '% CUMPL.TP' == '% CUMPL. TP'."""
    base, _, repeticion = nombre.partition(" #")
    # "$ LIQUIDAR.1" (copias hechas con otras herramientas) = "$ LIQUIDAR #2"
    copia = re.search(r"\.(\d+)$", base)
    if copia and not repeticion:
        base, repeticion = base[:copia.start()], str(int(copia.group(1)) + 1)
    return re.sub(r"[^A-Z0-9%&$]", "", base) + (f"#{repeticion}" if repeticion else "")


def resolver_columna(esperada, tipo, disponibles, usadas):
    """
    Busca la columna del Excel que corresponde a la esperada:
    1) nombre exacto, 2) mismo nombre sin espacios/puntos, 3) nombres
    alternativos, 4) solo para texto: un único encabezado que lo contenga.
    """
    libres = [c for c in disponibles if c not in usadas and c != "__FILA_EXCEL"]
    if esperada in libres:
        return esperada

    por_clave = {clave_columna(c): c for c in libres}
    for candidato in [esperada] + ALIAS_COLUMNAS.get(esperada, []):
        encontrada = por_clave.get(clave_columna(candidato))
        if encontrada:
            return encontrada

    if tipo == "texto" and " #" not in esperada:
        objetivo = clave_columna(esperada)
        contienen = [c for c in libres if " #" not in c and objetivo in clave_columna(c)]
        if len(contienen) == 1:
            return contienen[0]
    return None


def preparar_hoja(datos, config, mes_respaldo):
    """Selecciona, convierte y renombra las columnas de una hoja
    (incluidas las de dinero, que luego se separan)."""
    columnas = config["columnas"] + (config["dinero"]["columnas"] if "dinero" in config else [])
    avisos = []

    # Relaciona cada columna esperada con la real del Excel
    mapa, usadas, faltantes, renombradas = {}, set(), [], []
    for esperada, _, tipo in columnas:
        real = resolver_columna(esperada, tipo, list(datos.columns), usadas)
        if real is None:
            faltantes.append(esperada)
        else:
            mapa[esperada] = real
            usadas.add(real)
            if real != esperada:
                renombradas.append(f'"{real}" como {esperada}')

    if renombradas:
        avisos.append("Encabezados leídos con otro nombre: " + "; ".join(renombradas))
    faltantes = [f for f in faltantes if f not in COLUMNAS_OPCIONALES]
    if faltantes:
        disponibles = [c for c in datos.columns if c not in usadas and c != "__FILA_EXCEL"
                       and not str(c).startswith("CATEGORIA EFEC")]
        avisos.append("Columnas faltantes (quedan vacías): " + ", ".join(faltantes)
                      + ("\n     Encabezados sin usar en la hoja: " + ", ".join(disponibles) if disponibles else ""))

    salida = pd.DataFrame(index=datos.index)
    salida["fila_excel"] = datos["__FILA_EXCEL"].astype(int)

    # Umbral de las llaves de ese mes, tomado del encabezado
    # "Categoria Efec 60%" / "Categoria Efec 70%" (Liquidación por Rutas)
    if config.get("detectar_umbral"):
        umbral = None
        for col in datos.columns:
            m = re.match(r"^CATEGORIA EFEC (\d+)%", str(col))
            if m:
                umbral = int(m.group(1))
                break
        salida["umbral_llave"] = umbral
        if umbral:
            avisos.append(f"ℹ️  Llaves de este mes: mínimo {umbral}% (encabezado \"Categoria Efec {umbral}%\").")

    for encabezado, destino, tipo in columnas:
        serie = datos[mapa[encabezado]] if encabezado in mapa else pd.Series([None] * len(datos), index=datos.index)

        if tipo == "pct":
            salida[destino] = columna_a_porcentaje(serie)
        elif tipo == "numero":
            salida[destino] = pd.to_numeric(serie.apply(convertir_numero), errors="coerce")
        elif tipo == "mes":
            partes = serie.apply(lambda v: interpretar_mes(v, mes_respaldo))
            salida["mes_orden"] = partes.apply(lambda p: p[0] * 100 + p[1] if p else None)
            salida["mes"] = partes.apply(lambda p: f"{NOMBRES_MES[p[1]]} {p[0]}" if p else None)
        else:
            salida[destino] = serie.apply(limpiar_texto)
            if destino == "regional":
                # "REGIONAL BOGOTA" y "BOGOTA" quedan iguales entre meses
                salida[destino] = salida[destino].apply(
                    lambda v: re.sub(r"^REGIONAL\s+", "", v, flags=re.I) if isinstance(v, str) else v)

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

def separar_dinero(df, config):
    """Divide una hoja en (tabla pública solo con %, tabla protegida con dinero)."""
    dinero = config["dinero"]
    cols_dinero = [d for _, d, _ in dinero["columnas"]]
    publica = df.drop(columns=cols_dinero)
    protegida = df[["fila_excel", "mes", "mes_orden"] + dinero["identidad"] + cols_dinero].copy()
    return publica, protegida


def cargar_preliquidacion(archivo=None, mostrar=True):
    """
    Devuelve (archivo, dict) con las claves de TABLAS:
    INDICADORES, RUTAS, RUTAS_DINERO, LIDERES, LIDERES_DINERO.
    Los DataFrames ya tienen los nombres de columna de Supabase.
    """
    def log(msg=""):
        if mostrar:
            print(msg)

    if not CARPETA_BASE.exists():
        raise FileNotFoundError(f"No existe la carpeta: {CARPETA_BASE}")

    if archivo is None:
        raise ValueError("Indica el archivo a leer (o usa cargar_todas).")
    archivo = Path(archivo)

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
            log(aviso if aviso.startswith("ℹ️") else f"⚠️  {aviso}")

        meses = sorted(df["mes"].dropna().unique().tolist()) if len(df) else []
        log(f"✅ Registros: {len(df):,}   Meses: {', '.join(meses) or '—'}")

        if "dinero" in config:
            # Porcentaje de la comisión que se paga según el Excel (no es dinero)
            base = pd.to_numeric(df["valor_comision"], errors="coerce")
            pagar = pd.to_numeric(df["valor_pagar"], errors="coerce")
            df["pct_comision"] = (pagar / base.where(base > 0) * 100).round(2)
            df = df.astype(object).where(pd.notna(df), None)
            resultado[clave], resultado[f"{clave}_DINERO"] = separar_dinero(df, config)
            log(f"🔒 Valores en dinero separados -> {config['dinero']['tabla']}")
        else:
            resultado[clave] = df

    return archivo, resultado


def cargar_todas(mostrar=True):
    """
    Lee todas las preliquidaciones de la carpeta y une sus datos.
    Devuelve (lista_de_archivos, dict con las claves de TABLAS).
    Si un mes aparece en varios archivos, se queda el del archivo más reciente.
    """
    def log(msg=""):
        if mostrar:
            print(msg)

    if not CARPETA_BASE.exists():
        raise FileNotFoundError(f"No existe la carpeta: {CARPETA_BASE}")

    archivos, otros = buscar_archivos()

    log(f"\n🔎 Archivos de preliquidación encontrados: {len(archivos)}")
    for a in archivos:
        log(f"   • {a.name}")
    if otros:
        log("   (Otros Excel en la carpeta que NO se procesan porque su nombre no dice 'PreLiquidacion':)")
        for a in otros:
            log(f"     - {a.name}")

    if not archivos:
        raise FileNotFoundError(
            f'No hay ningún Excel cuyo nombre contenga "PreLiquidacion" en {CARPETA_BASE}'
        )

    union = {}
    for archivo in archivos:
        _, bases = cargar_preliquidacion(archivo, mostrar)
        for clave, df in bases.items():
            if clave in union and len(df):
                meses_nuevos = set(df["mes_orden"].unique())
                repetidos = set(union[clave]["mes_orden"].unique()) & meses_nuevos
                if repetidos and clave in ("RUTAS", "INDICADORES", "LIDERES"):
                    log(f"⚠️  {archivo.name} repite {', '.join(sorted(str(m) for m in repetidos))} "
                        f"en {clave}: se usa este archivo por ser el más reciente.")
                union[clave] = union[clave][~union[clave]["mes_orden"].isin(meses_nuevos)]
                union[clave] = pd.concat([union[clave], df], ignore_index=True)
            else:
                union[clave] = df

    return archivos, union


# ============================================================
# EJECUCIÓN DIRECTA: genera CSV + control para revisar
# ============================================================

def main():
    print("\n" + "=" * 70)
    print(" EXTRACTOR · PRELIQUIDACIÓN DE CAMPO")
    print("=" * 70)
    print(f"\n📁 Carpeta: {CARPETA_BASE}")

    try:
        archivos, bases = cargar_todas()
    except Exception as error:
        print(f"\n❌ {error}")
        input("\nPresiona ENTER para salir...")
        sys.exit(1)

    CARPETA_CSV.mkdir(exist_ok=True)
    control = []

    print("\n" + "=" * 70)
    print(" GENERANDO CSV")
    print("=" * 70)

    for clave, tabla in TABLAS:
        if clave not in bases:
            control.append({"TABLA": tabla, "REGISTROS": 0, "ESTADO": "NO CARGADA"})
            print(f"❌ {tabla:<28} NO CARGADA")
            continue

        df = bases[clave]
        ruta_csv = CARPETA_CSV / f"{tabla}.csv"
        df.to_csv(ruta_csv, index=False, encoding="utf-8-sig")
        meses = ", ".join(sorted(df["mes"].dropna().unique().tolist()))
        control.append({"TABLA": tabla, "REGISTROS": len(df), "MESES": meses, "ESTADO": "OK"})
        print(f"✅ {tabla:<28} {len(df):>8,} registros   ({meses})")

    pd.DataFrame(control).to_excel(ARCHIVO_CONTROL, index=False, sheet_name="CONTROL")

    print("\n" + "=" * 70)
    print(" PROCESO TERMINADO")
    print("=" * 70)
    print(f"\n📁 CSV:     {CARPETA_CSV}")
    print(f"📋 Control: {ARCHIVO_CONTROL}\n")
    input("Presiona ENTER para cerrar...")


if __name__ == "__main__":
    main()