import os
import pandas as pd
from pathlib import Path
from datetime import datetime

# ============================================================
# CONFIGURACIÓN
# ============================================================

# Si este archivo .py está dentro de la carpeta BASE DE DATOS,
# funcionará igualmente.
CARPETA_BASE = Path(__file__).resolve().parent / "BASE DE DATOS"

# Si BASE DE DATOS no está junto al .py, puedes poner aquí
# la ruta completa de tu carpeta.
#
# Ejemplo:
# CARPETA_BASE = Path(r"C:\Users\Dylan\Desktop\BASE DE DATOS")

ARCHIVO_SALIDA = CARPETA_BASE / "BASE_CONSOLIDADA.xlsx"

# ============================================================
# COLUMNAS ESPERADAS
# ============================================================

COLUMNAS_ESPERADAS = {

    "VENTA": [
        "REGIONAL",
        "RUTA",
        "CADENA",
        "PUNTO DE VENTA",
        "SEGMENTO",
        "MARCA",
        "PRODUCTO",
        "FECHA",
        "VENTA NETA",
        "VENTA KILOS",
        "CUOTA PESOS",
        "CUOTA KILOS",
        "% CUMPLIMIENTO CUOTA",
        "% CUMPLIMIENTO VENTA"
    ],

    "PARTICIPACION": [
        "REGIONAL",
        "CANAL",
        "CADENA",
        "FORMATO",
        "CÓDIGO CLIENTE",
        "CIUDAD",
        "PUNTO DE VENTA",
        "MERCADERISTA",
        "OBJETIVO PARTICIPACION",
        "% PARTICIPACIÓN CATEGORIA",
        "CARAS",
        "CATEGORÍA",
        "MARCA",
        "FECHA VISITA",
        "TIPO EMPRESA",
        "LIDER EJECUCION",
        "SUPERVISOR"
    ],

    "MERCADERISTA": [
        "REGIONAL",
        "RUTA",
        "CANAL",
        "RAZÓN SOCIAL",
        "CÓDIGO CLIENTE",
        "PDV",
        "FORMATO",
        "FECHA",
        "NOMBRE PLAN COMERCIAL",
        "CATEGORÍA OBJETIVO",
        "TIPO EXHIBICIÓN OBJETIVO",
        "TIPO EXHIBICIÓN IMPLEMENTADA",
        "CANTIDAD IMPLEMENTACIÓN OBJETIVO",
        "CUMPLE IMPLEMENTACIÓN",
        "CANTIDAD IMPLEMENTADA",
        "CAUSAL",
        "% IMPLEMENTACIÓN",
        "SUPERVISOR",
        "LIDER DE EJECUCION",
        "MERCADERISTA"
    ],

    "FECHAS CORTAS": [
        "REGIONAL",
        "RUTA",
        "CANAL",
        "CADENA",
        "FORMATO",
        "CÓDIGO CLIENTE",
        "PUNTO DE VENTA",
        "CATEGORIA",
        "MARCA",
        "PRODUCTO",
        "FECHA DE VISITA",
        "FECHA DE VENCIMIENTO",
        "CANTIDAD UNIDADES",
        "VALOR MONETARIO",
        "CANTIDAD DIAS VIDA UTIL",
        "POLÍTICA",
        "SUPERVISOR",
        "LIDER DE EJECUCION",
        "MERCADERISTA"
    ],

    "DEVOLUCIONES": [
        "REGIONAL",
        "CADENA",
        "FORMATO",
        "CÓDIGO CLIENTE",
        "PUNTO DE VENTA",
        "FECHA",
        "MARCA",
        "SEGMENTO",
        "PRODUCTO",
        "VALOR",
        "UNIDADES",
        "MOTIVOS DE DEVOLUCIÓN"
    ],

    "COBERTURA": [
        "REGIONAL",
        "CANAL",
        "TIPOLOGIA",
        "CLIENTE",
        "LIDER",
        "SUPERVISOR",
        "RUTA",
        "PDV Programados",
        "PDV Visitados",
        "% PUNTOS VISITADOS AUNQUE SEA UNA SOLA VEZ",
        "% VISITAS RUTERO AL CORTE",
        "% EFECTIVIDAD + EXTRA RUTA"
    ],

    "AGOTADOS": [
        "REGIONAL",
        "RUTA",
        "CANAL",
        "FORMATO",
        "CÓDIGO CLIENTE",
        "PUNTO DE VENTA",
        "CATEGORIA",
        "MARCA",
        "PRODUCTO",
        "CAUSAL AGOTADO",
        "FECHA CAPTURA",
        "ESTADO AGOTADO",
        "SUPERVISOR",
        "LIDER DE EJECUCION",
        "MERCADERISTA"
    ]
}

# ============================================================
# FUNCIONES
# ============================================================

def normalizar_nombre_archivo(nombre):
    """
    Convierte nombres de archivos a una forma comparable.
    """
    nombre = nombre.strip().upper()
    nombre = nombre.replace(".XLSX", "")
    nombre = nombre.replace(".XLS", "")
    return nombre


def normalizar_columna(nombre):
    """
    Limpia espacios innecesarios en los nombres de columnas.
    """
    nombre = str(nombre)
    nombre = nombre.replace("\n", " ")
    nombre = " ".join(nombre.split())
    return nombre.strip()


def buscar_archivo(nombre_base):
    """
    Busca el Excel correspondiente aunque esté en XLSX o XLS.
    """
    posibles = [
        CARPETA_BASE / f"{nombre_base}.xlsx",
        CARPETA_BASE / f"{nombre_base}.xls",
        CARPETA_BASE / f"{nombre_base.upper()}.xlsx",
        CARPETA_BASE / f"{nombre_base.upper()}.xls"
    ]

    for archivo in posibles:
        if archivo.exists():
            return archivo

    # Búsqueda flexible
    for archivo in CARPETA_BASE.iterdir():
        if archivo.is_file():
            nombre = normalizar_nombre_archivo(archivo.name)
            if nombre == normalizar_nombre_archivo(nombre_base):
                return archivo

    return None


def leer_excel(nombre_base, columnas_esperadas):
    """
    Lee un Excel completo y verifica sus columnas.
    """

    archivo = buscar_archivo(nombre_base)

    if archivo is None:
        print(f"\n❌ NO ENCONTRADO: {nombre_base}")
        return None

    print("\n" + "=" * 70)
    print(f"LEYENDO: {archivo.name}")
    print("=" * 70)

    try:
        # dtype=object permite conservar los valores originales
        df = pd.read_excel(
            archivo,
            sheet_name=0,
            dtype=object
        )
    except Exception as e:
        print(f"❌ Error leyendo {archivo.name}")
        print(f"   {e}")
        return None

    # Normalizar nombres de columnas
    df.columns = [
        normalizar_columna(col)
        for col in df.columns
    ]

    columnas_actuales = list(df.columns)

    # Verificar columnas faltantes
    faltantes = [
        columna
        for columna in columnas_esperadas
        if columna not in columnas_actuales
    ]

    # Verificar columnas adicionales
    adicionales = [
        columna
        for columna in columnas_actuales
        if columna not in columnas_esperadas
    ]

    if faltantes:
        print("\n⚠️ COLUMNAS FALTANTES:")
        for columna in faltantes:
            print(f"   - {columna}")

    if adicionales:
        print("\nℹ️ COLUMNAS ADICIONALES EN EL EXCEL:")
        for columna in adicionales:
            print(f"   + {columna}")

    # Mostrar información
    print(f"\n📊 Registros encontrados: {len(df):,}")
    print(f"📋 Columnas encontradas: {len(df.columns)}")

    # Eliminar únicamente filas completamente vacías
    registros_antes = len(df)

    df = df.dropna(
        how="all"
    ).reset_index(drop=True)

    registros_despues = len(df)

    if registros_antes != registros_despues:
        print(
            f"🧹 Filas completamente vacías eliminadas: "
            f"{registros_antes - registros_despues:,}"
        )

    print(f"✅ Registros finales: {len(df):,}")

    return df


# ============================================================
# INICIO
# ============================================================

print("\n")
print("=" * 70)
print(" EXTRACTOR DE BASES DE DATOS - POWER BI")
print("=" * 70)

print(f"\n📁 Carpeta de trabajo:")
print(CARPETA_BASE)

if not CARPETA_BASE.exists():
    print("\n❌ ERROR:")
    print("No se encontró la carpeta BASE DE DATOS.")
    print("\nRevisa que tengas una estructura como:")
    print("C:\\TuCarpeta\\")
    print("   ├── extraer_bases.py")
    print("   └── BASE DE DATOS")
    print("       ├── VENTA.xlsx")
    print("       ├── PARTICIPACION.xlsx")
    print("       ├── MERCADERISTA.xlsx")
    print("       ├── FECHAS CORTAS.xlsx")
    print("       ├── DEVOLUCIONES.xlsx")
    print("       ├── COBERTURA.xlsx")
    print("       └── AGOTADOS.xlsx")
    input("\nPresiona ENTER para salir...")
    exit()


# ============================================================
# LEER LAS 7 BASES
# ============================================================

bases = {}

for nombre, columnas in COLUMNAS_ESPERADAS.items():

    df = leer_excel(
        nombre,
        columnas
    )

    if df is not None:
        bases[nombre] = df


# ============================================================
# RESUMEN
# ============================================================

print("\n")
print("=" * 70)
print(" RESUMEN DE EXTRACCIÓN")
print("=" * 70)

total_registros = 0

for nombre in COLUMNAS_ESPERADAS.keys():

    if nombre in bases:

        cantidad = len(bases[nombre])

        total_registros += cantidad

        print(
            f"✅ {nombre:<20} "
            f"{cantidad:>12,} registros"
        )

    else:

        print(
            f"❌ {nombre:<20} "
            f"NO CARGADO"
        )

print("-" * 70)

print(
    f"TOTAL REGISTROS: "
    f"{total_registros:,}"
)


# ============================================================
# CREAR EXCEL CONSOLIDADO
# ============================================================

print("\n")
print("=" * 70)
print(" GENERANDO EXCEL CONSOLIDADO")
print("=" * 70)

try:

    with pd.ExcelWriter(
        ARCHIVO_SALIDA,
        engine="openpyxl"
    ) as writer:

        for nombre, df in bases.items():

            # Excel limita el nombre de las hojas
            nombre_hoja = nombre[:31]

            df.to_excel(
                writer,
                sheet_name=nombre_hoja,
                index=False
            )

    print("\n✅ EXCEL CONSOLIDADO GENERADO")
    print(f"📄 {ARCHIVO_SALIDA}")

except Exception as e:

    print("\n❌ ERROR GENERANDO EL EXCEL:")
    print(e)


# ============================================================
# GENERAR CSV INDIVIDUALES
# ============================================================

CARPETA_CSV = CARPETA_BASE / "CSV_EXTRAIDOS"

CARPETA_CSV.mkdir(
    exist_ok=True
)

print("\n")
print("=" * 70)
print(" GENERANDO CSV INDIVIDUALES")
print("=" * 70)

for nombre, df in bases.items():

    archivo_csv = CARPETA_CSV / f"{nombre}.csv"

    try:

        df.to_csv(
            archivo_csv,
            index=False,
            encoding="utf-8-sig"
        )

        print(
            f"✅ {archivo_csv.name} "
            f"({len(df):,} registros)"
        )

    except Exception as e:

        print(
            f"❌ Error creando {archivo_csv.name}: {e}"
        )


# ============================================================
# ARCHIVO DE CONTROL
# ============================================================

archivo_control = CARPETA_BASE / "CONTROL_EXTRACCION.xlsx"

control = []

for nombre in COLUMNAS_ESPERADAS.keys():

    if nombre in bases:

        df = bases[nombre]

        control.append({
            "BASE": nombre,
            "REGISTROS": len(df),
            "COLUMNAS": len(df.columns),
            "ESTADO": "OK"
        })

    else:

        control.append({
            "BASE": nombre,
            "REGISTROS": 0,
            "COLUMNAS": 0,
            "ESTADO": "NO CARGADO"
        })


control_df = pd.DataFrame(control)

try:

    with pd.ExcelWriter(
        archivo_control,
        engine="openpyxl"
    ) as writer:

        control_df.to_excel(
            writer,
            sheet_name="CONTROL",
            index=False
        )

    print("\n✅ ARCHIVO DE CONTROL:")
    print(archivo_control)

except Exception as e:

    print("\n❌ No se pudo crear el archivo de control:")
    print(e)


# ============================================================
# FINAL
# ============================================================

print("\n")
print("=" * 70)
print(" PROCESO TERMINADO")
print("=" * 70)

print("\nArchivos generados:")

print(f"\n📊 {ARCHIVO_SALIDA}")
print("   Contiene las 7 bases en diferentes hojas.")

print(f"\n📁 {CARPETA_CSV}")
print("   Contiene cada base individual en CSV.")

print(f"\n📋 {archivo_control}")
print("   Contiene el conteo de registros de cada base.")

print("\n")

input("Presiona ENTER para cerrar...")
