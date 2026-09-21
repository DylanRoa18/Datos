import os
import time
import pandas as pd
from dotenv import load_dotenv
from supabase import create_client

# ============================================================
# CONFIGURACIÓN
# ============================================================

load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise Exception(
        "No se encontraron SUPABASE_URL o SUPABASE_KEY en el archivo .env"
    )

supabase = create_client(
    SUPABASE_URL,
    SUPABASE_KEY
)

CARPETA_EXCEL = "BASE DE DATOS"

BATCH_SIZE = 500

# ============================================================
# CONFIGURACIÓN DE CARGA
# ============================================================

# True  = prueba de 500 registros de VENTA
# False = carga completa de las 7 bases
MODO_PRUEBA = False

REGISTROS_PRUEBA = 500

# True = elimina los datos actuales antes de cargar
# False = conserva los datos actuales
LIMPIAR_TABLAS = True


# ============================================================
# CONFIGURACIÓN DE LAS 7 BASES
# ============================================================

CONFIGURACION = {

    "VENTA": {
        "archivo": "VENTA.xlsx",
        "tabla": "venta",

        "columnas": {
            "REGIONAL": "regional",
            "RUTA": "ruta",
            "CADENA": "cadena",
            "PUNTO DE VENTA": "punto_de_venta",
            "SEGMENTO": "segmento",
            "MARCA": "marca",
            "PRODUCTO": "producto",
            "FECHA": "fecha",
            "VENTA NETA": "venta_neta",
            "VENTA KILOS": "venta_kilos",
            "CUOTA PESOS": "cuota_pesos",
            "CUOTA KILOS": "cuota_kilos",
            "% CUMPLIMIENTO CUOTA": "porcentaje_cumplimiento_cuota",
            "% CUMPLIMIENTO VENTA": "porcentaje_cumplimiento_venta"
        },

        "fechas": [
            "FECHA"
        ],

        "numericas": [
            "VENTA NETA",
            "VENTA KILOS",
            "CUOTA PESOS",
            "CUOTA KILOS",
            "% CUMPLIMIENTO CUOTA",
            "% CUMPLIMIENTO VENTA"
        ]
    },

    "PARTICIPACION": {
        "archivo": "PARTICIPACION.xlsx",
        "tabla": "participacion",

        "columnas": {
            "REGIONAL": "regional",
            "CANAL": "canal",
            "CADENA": "cadena",
            "FORMATO": "formato",
            "CÓDIGO CLIENTE": "codigo_cliente",
            "CIUDAD": "ciudad",
            "PUNTO DE VENTA": "punto_de_venta",
            "MERCADERISTA": "mercaderista",
            "OBJETIVO PARTICIPACION": "objetivo_participacion",
            "% PARTICIPACIÓN CATEGORIA": "porcentaje_participacion_categoria",
            "CARAS": "caras",
            "CATEGORÍA": "categoria",
            "MARCA": "marca",
            "FECHA VISITA": "fecha_visita",
            "TIPO EMPRESA": "tipo_empresa",
            "LIDER EJECUCION": "lider_ejecucion",
            "SUPERVISOR": "supervisor"
        },

        "fechas": [
            "FECHA VISITA"
        ],

        "numericas": [
            "OBJETIVO PARTICIPACION",
            "% PARTICIPACIÓN CATEGORIA",
            "CARAS"
        ]
    },

    "MERCADERISTA": {
        "archivo": "MERCADERISTA.xlsx",
        "tabla": "mercaderista",

        "columnas": {
            "REGIONAL": "regional",
            "RUTA": "ruta",
            "CANAL": "canal",
            "RAZÓN SOCIAL": "razon_social",
            "CÓDIGO CLIENTE": "codigo_cliente",
            "PDV": "pdv",
            "FORMATO": "formato",
            "FECHA": "fecha",
            "NOMBRE PLAN COMERCIAL": "nombre_plan_comercial",
            "CATEGORÍA OBJETIVO": "categoria_objetivo",
            "TIPO EXHIBICIÓN OBJETIVO": "tipo_exhibicion_objetivo",
            "TIPO EXHIBICIÓN IMPLEMENTADA": "tipo_exhibicion_implementada",
            "CANTIDAD IMPLEMENTACIÓN OBJETIVO": "cantidad_implementacion_objetivo",
            "CUMPLE IMPLEMENTACIÓN": "cumple_implementacion",
            "CANTIDAD IMPLEMENTADA": "cantidad_implementada",
            "CAUSAL": "causal",
            "% IMPLEMENTACIÓN": "porcentaje_implementacion",
            "SUPERVISOR": "supervisor",
            "LIDER DE EJECUCION": "lider_de_ejecucion",
            "MERCADERISTA": "mercaderista"
        },

        "fechas": [
            "FECHA"
        ],

        "numericas": [
            "CANTIDAD IMPLEMENTACIÓN OBJETIVO",
            "CANTIDAD IMPLEMENTADA",
            "% IMPLEMENTACIÓN"
        ]
    },

    "FECHAS CORTAS": {
        "archivo": "FECHAS CORTAS.xlsx",
        "tabla": "fechas_cortas",

        "columnas": {
            "REGIONAL": "regional",
            "RUTA": "ruta",
            "CANAL": "canal",
            "CADENA": "cadena",
            "FORMATO": "formato",
            "CÓDIGO CLIENTE": "codigo_cliente",
            "PUNTO DE VENTA": "punto_de_venta",
            "CATEGORIA": "categoria",
            "MARCA": "marca",
            "PRODUCTO": "producto",
            "FECHA DE VISITA": "fecha_de_visita",
            "FECHA DE VENCIMIENTO": "fecha_de_vencimiento",
            "CANTIDAD UNIDADES": "cantidad_unidades",
            "VALOR MONETARIO": "valor_monetario",
            "CANTIDAD DIAS VIDA UTIL": "cantidad_dias_vida_util",
            "POLÍTICA": "politica",
            "SUPERVISOR": "supervisor",
            "LIDER DE EJECUCION": "lider_de_ejecucion",
            "MERCADERISTA": "mercaderista"
        },

        "fechas": [
            "FECHA DE VISITA",
            "FECHA DE VENCIMIENTO"
        ],

        "numericas": [
            "CANTIDAD UNIDADES",
            "VALOR MONETARIO",
            "CANTIDAD DIAS VIDA UTIL"
        ]
    },

    "DEVOLUCIONES": {
        "archivo": "DEVOLUCIONES.xlsx",
        "tabla": "devoluciones",

        "columnas": {
            "REGIONAL": "regional",
            "CADENA": "cadena",
            "FORMATO": "formato",
            "CÓDIGO CLIENTE": "codigo_cliente",
            "PUNTO DE VENTA": "punto_de_venta",
            "FECHA": "fecha",
            "MARCA": "marca",
            "SEGMENTO": "segmento",
            "PRODUCTO": "producto",
            "VALOR": "valor",
            "UNIDADES": "unidades",
            "MOTIVOS DE DEVOLUCIÓN": "motivos_de_devolucion"
        },

        "fechas": [
            "FECHA"
        ],

        "numericas": [
            "VALOR",
            "UNIDADES"
        ]
    },

    "COBERTURA": {
        "archivo": "COBERTURA.xlsx",
        "tabla": "cobertura",

        "columnas": {
            "REGIONAL": "regional",
            "CANAL": "canal",
            "TIPOLOGIA": "tipologia",
            "CLIENTE": "cliente",
            "LIDER": "lider",
            "SUPERVISOR": "supervisor",
            "RUTA": "ruta",
            "PDV Programados": "pdv_programados",
            "PDV Visitados": "pdv_visitados",
            "% PUNTOS VISITADOS AUNQUE SEA UNA SOLA VEZ": "porcentaje_puntos_visitados_aunque_sea_una_sola_vez",
            "% VISITAS RUTERO AL CORTE": "porcentaje_visitas_rutero_al_corte",
            "% EFECTIVIDAD + EXTRA RUTA": "porcentaje_efectividad_extra_ruta"
        },

        "fechas": [],

        "numericas": [
            "PDV Programados",
            "PDV Visitados",
            "% PUNTOS VISITADOS AUNQUE SEA UNA SOLA VEZ",
            "% VISITAS RUTERO AL CORTE",
            "% EFECTIVIDAD + EXTRA RUTA"
        ]
    },

    "AGOTADOS": {
        "archivo": "AGOTADOS.xlsx",
        "tabla": "agotados",

        "columnas": {
            "REGIONAL": "regional",
            "RUTA": "ruta",
            "CANAL": "canal",
            "FORMATO": "formato",
            "CÓDIGO CLIENTE": "codigo_cliente",
            "PUNTO DE VENTA": "punto_de_venta",
            "CATEGORIA": "categoria",
            "MARCA": "marca",
            "PRODUCTO": "producto",
            "CAUSAL AGOTADO": "causal_agotado",
            "FECHA CAPTURA": "fecha_captura",
            "ESTADO AGOTADO": "estado_agotado",
            "SUPERVISOR": "supervisor",
            "LIDER DE EJECUCION": "lider_de_ejecucion",
            "MERCADERISTA": "mercaderista"
        },

        "fechas": [
            "FECHA CAPTURA"
        ],

        "numericas": []
    }
}


# ============================================================
# CONVERTIR NUMEROS
# ============================================================

def convertir_numero(valor):

    if pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    if texto == "":
        return None

    porcentaje = "%" in texto

    texto = (
        texto
        .replace("%", "")
        .replace("$", "")
        .replace(" ", "")
    )

    try:

        if "," in texto and "." in texto:

            if texto.rfind(",") > texto.rfind("."):

                texto = texto.replace(".", "")
                texto = texto.replace(",", ".")

            else:

                texto = texto.replace(",", "")

        elif "," in texto:

            texto = texto.replace(",", ".")

        numero = float(texto)

        if porcentaje:
            numero = numero / 100

        return numero

    except Exception:
        return None


# ============================================================
# LIMPIAR TEXTO
# ============================================================

def limpiar_texto(valor):

    if pd.isna(valor):
        return None

    texto = str(valor).strip()

    if texto == "":
        return None

    return texto


# ============================================================
# PREPARAR DATAFRAME
# ============================================================

def preparar_dataframe(df, config):

    columnas_origen = list(
        config["columnas"].keys()
    )

    faltantes = [
        columna
        for columna in columnas_origen
        if columna not in df.columns
    ]

    if faltantes:

        raise Exception(
            "Faltan columnas en el Excel: "
            + ", ".join(faltantes)
        )

    df = df[columnas_origen].copy()

    # --------------------------------------------------------
    # FECHAS
    # --------------------------------------------------------

    for columna in config["fechas"]:

        df[columna] = pd.to_datetime(
            df[columna],
            errors="coerce"
        )

        df[columna] = (
            df[columna]
            .dt.strftime("%Y-%m-%d")
        )

        df[columna] = df[columna].where(
            df[columna].notna(),
            None
        )

    # --------------------------------------------------------
    # NUMERICOS
    # --------------------------------------------------------

    for columna in config["numericas"]:

        df[columna] = df[columna].apply(
            convertir_numero
        )

    # --------------------------------------------------------
    # TEXTOS
    # --------------------------------------------------------

    columnas_numericas = set(
        config["numericas"]
    )

    columnas_fechas = set(
        config["fechas"]
    )

    for columna in columnas_origen:

        if (
            columna not in columnas_numericas
            and columna not in columnas_fechas
        ):

            df[columna] = df[columna].apply(
                limpiar_texto
            )

    # --------------------------------------------------------
    # RENOMBRAR
    # --------------------------------------------------------

    df = df.rename(
        columns=config["columnas"]
    )

    # --------------------------------------------------------
    # NaN → None
    # --------------------------------------------------------

    df = df.astype(object)

    df = df.where(
        pd.notna(df),
        None
    )

    return df


# ============================================================
# LIMPIAR TABLA SUPABASE
# ============================================================
def limpiar_tabla(tabla):

    print(f"\n🧹 Limpiando tabla: {tabla}")

    try:

        supabase.rpc(
            "limpiar_tabla",
            {
                "tabla_nombre": tabla
            }
        ).execute()

        print(f"✓ Tabla {tabla} limpiada")

    except Exception as e:

        print(f"❌ Error limpiando {tabla}: {e}")

        raise Exception(
            f"No se pudo limpiar la tabla {tabla}: {e}"
        )

# ============================================================
# SUBIR DATAFRAME POR LOTES
# ============================================================

def subir_dataframe(df, tabla):

    total = len(df)

    print()
    print(
        f"Tabla Supabase: {tabla}"
    )

    print(
        f"Registros a cargar: {total:,}"
    )

    print("-" * 60)

    registros_subidos = 0

    for inicio in range(
        0,
        total,
        BATCH_SIZE
    ):

        fin = min(
            inicio + BATCH_SIZE,
            total
        )

        lote = df.iloc[
            inicio:fin
        ]

        registros = lote.to_dict(
            orient="records"
        )

        intento = 0

        while True:

            try:

                supabase \
                    .table(tabla) \
                    .insert(registros) \
                    .execute()

                break

            except Exception as error:

                intento += 1

                print()
                print(
                    f"❌ Error en lote "
                    f"{inicio}-{fin}:"
                )

                print(error)

                if intento >= 3:

                    raise Exception(
                        f"No se pudo subir el lote "
                        f"{inicio}-{fin}"
                    )

                espera = intento * 2

                print(
                    f"Reintentando en "
                    f"{espera} segundos..."
                )

                time.sleep(espera)

        registros_subidos = fin

        porcentaje = (
            registros_subidos / total
        ) * 100

        print(
            f"Progreso: "
            f"{registros_subidos:,}/"
            f"{total:,} "
            f"({porcentaje:.2f}%)"
        )

    print()

    print(
        f"✓ Tabla {tabla} cargada correctamente."
    )


# ============================================================
# PROCESAR UNA BASE
# ============================================================

def procesar_base(nombre, config):

    archivo = os.path.join(
        CARPETA_EXCEL,
        config["archivo"]
    )

    if not os.path.exists(archivo):

        raise Exception(
            f"No se encontró el archivo: "
            f"{archivo}"
        )

    print()
    print("=" * 70)
    print(
        f"PROCESANDO: {nombre}"
    )
    print(
        f"Archivo: {config['archivo']}"
    )
    print("=" * 70)

    # --------------------------------------------------------
    # LEER EXCEL
    # --------------------------------------------------------

    df = pd.read_excel(
        archivo,
        dtype=object
    )

    print(
        f"Registros encontrados: "
        f"{len(df):,}"
    )

    # --------------------------------------------------------
    # PREPARAR DATOS
    # --------------------------------------------------------

    df = preparar_dataframe(
        df,
        config
    )

    # --------------------------------------------------------
    # MODO PRUEBA
    # --------------------------------------------------------

    if MODO_PRUEBA:

        df = df.head(
            REGISTROS_PRUEBA
        )

        print(
            f"MODO PRUEBA: "
            f"solo se cargarán "
            f"{len(df):,} registros."
        )

    # --------------------------------------------------------
    # LIMPIAR TABLA
    # --------------------------------------------------------

    if LIMPIAR_TABLAS:

        limpiar_tabla(
            config["tabla"]
        )

    # --------------------------------------------------------
    # SUBIR
    # --------------------------------------------------------

    subir_dataframe(
        df,
        config["tabla"]
    )


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

def main():

    inicio_total = time.time()

    print()
    print("=" * 70)
    print("SINCRONIZACIÓN EXCEL → SUPABASE")
    print("=" * 70)

    print()
    print(
        f"Supabase: {SUPABASE_URL}"
    )

    print()
    print(
        f"Modo prueba: {MODO_PRUEBA}"
    )

    print(
        f"Limpiar tablas: {LIMPIAR_TABLAS}"
    )

    print(
        f"Tamaño de lote: {BATCH_SIZE}"
    )

    print()

    # ========================================================
    # MODO PRUEBA
    # ========================================================

    if MODO_PRUEBA:

        print(
            "⚠️ MODO PRUEBA ACTIVADO"
        )

        print(
            f"Se cargarán máximo "
            f"{REGISTROS_PRUEBA} registros "
            f"de VENTA."
        )

        procesar_base(
            "VENTA",
            CONFIGURACION["VENTA"]
        )

    # ========================================================
    # CARGA COMPLETA
    # ========================================================

    else:

        print(
            "🚀 CARGA COMPLETA ACTIVADA"
        )

        print()
        print(
            "Se cargarán las 7 bases:"
        )

        for nombre, config in CONFIGURACION.items():

            print(
                f"  • {nombre}: "
                f"{config['archivo']}"
            )

        print()
        print(
            "Las tablas serán limpiadas "
            "antes de cargar los datos."
        )

        print()

        for nombre, config in CONFIGURACION.items():

            procesar_base(
                nombre,
                config
            )

    # ========================================================
    # FINAL
    # ========================================================

    tiempo_total = (
        time.time() - inicio_total
    )

    minutos = int(
        tiempo_total // 60
    )

    segundos = int(
        tiempo_total % 60
    )

    print()
    print("=" * 70)
    print("✅ PROCESO TERMINADO CORRECTAMENTE")
    print("=" * 70)

    print(
        f"Tiempo total: "
        f"{minutos} min "
        f"{segundos} seg"
    )

    print()
    print(
        "Los datos de Excel están "
        "sincronizados con Supabase."
    )

    print("=" * 70)


# ============================================================
# EJECUTAR
# ============================================================

if __name__ == "__main__":
    main()