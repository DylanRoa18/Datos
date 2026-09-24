"""
============================================================
 SINCRONIZAR SUPABASE · SEGUIMIENTO DE CAMPO
============================================================
Sube a Supabase las 3 hojas del Excel de preliquidación usando
la limpieza de extraer_base.py (deben estar en la misma carpeta).

Modo de reemplazo:
  "MES"  -> borra en Supabase SOLO los meses que trae el Excel y los
            vuelve a cargar. Los meses anteriores se conservan, así el
            dashboard puede comparar contra el mes pasado.
  "TODO" -> vacía las tablas completas antes de cargar.

Archivo .env (junto a este script):
  SUPABASE_URL=https://bfbcplixcousebflrosx.supabase.co
  SUPABASE_KEY=<service_role key>   (nunca la pongas en el HTML)
============================================================
"""

import os
import sys
import time
from pathlib import Path

from dotenv import load_dotenv
from supabase import create_client

from extraer_base import HOJAS, cargar_preliquidacion


# ============================================================
# CONFIGURACIÓN
# ============================================================

load_dotenv(Path(__file__).resolve().parent / ".env")

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

MODO_REEMPLAZO = "MES"      # "MES" o "TODO"

MODO_PRUEBA = False         # True = sube solo unas filas por tabla
REGISTROS_PRUEBA = 20

BATCH_SIZE = 500
MAX_INTENTOS = 3


# ============================================================
# CONEXIÓN
# ============================================================

def conectar():
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError("Faltan SUPABASE_URL o SUPABASE_KEY en el archivo .env")
    return create_client(SUPABASE_URL, SUPABASE_KEY)


# ============================================================
# REINTENTOS
# ============================================================

def con_reintentos(accion, descripcion):
    """Ejecuta una operación contra Supabase con reintentos."""
    for intento in range(1, MAX_INTENTOS + 1):
        try:
            return accion()
        except Exception as error:
            print(f"   ❌ {descripcion} (intento {intento}/{MAX_INTENTOS}): {error}")
            if intento == MAX_INTENTOS:
                raise
            time.sleep(intento * 2)


# ============================================================
# LIMPIEZA EN SUPABASE
# ============================================================

def limpiar_tabla(supabase, tabla, meses):
    if MODO_REEMPLAZO == "TODO":
        print(f"🧹 Vaciando tabla completa: {tabla}")
        con_reintentos(
            lambda: supabase.table(tabla).delete().gte("id", 0).execute(),
            f"vaciar {tabla}",
        )
        return

    for mes in meses:
        print(f"🧹 Borrando mes {mes} de {tabla}")
        con_reintentos(
            lambda m=mes: supabase.table(tabla).delete().eq("mes_orden", m).execute(),
            f"borrar mes {mes} de {tabla}",
        )


# ============================================================
# CARGA POR LOTES
# ============================================================

def subir_registros(supabase, tabla, registros):
    total = len(registros)
    print(f"⬆️  Subiendo {total:,} registros a {tabla}")

    for inicio in range(0, total, BATCH_SIZE):
        fin = min(inicio + BATCH_SIZE, total)
        lote = registros[inicio:fin]
        con_reintentos(
            lambda l=lote: supabase.table(tabla).insert(l).execute(),
            f"lote {inicio}-{fin}",
        )
        print(f"   Progreso: {fin:,}/{total:,} ({fin / total * 100:.1f}%)")

    print(f"✅ {tabla} cargada.")


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

def main():
    inicio_total = time.time()

    print("\n" + "=" * 70)
    print(" SINCRONIZACIÓN EXCEL → SUPABASE · SEGUIMIENTO DE CAMPO")
    print("=" * 70)
    print(f"Modo de reemplazo: {MODO_REEMPLAZO}")
    if MODO_PRUEBA:
        print(f"MODO PRUEBA: solo {REGISTROS_PRUEBA} registros por tabla")

    supabase = conectar()
    print(f"Supabase: {SUPABASE_URL}")

    archivo, bases = cargar_preliquidacion()
    resumen = []

    for clave, config in HOJAS.items():
        tabla = config["tabla"]
        print("\n" + "-" * 70)

        if clave not in bases or bases[clave].empty:
            print(f"⚠️  {tabla}: sin datos, se omite (no se borra nada).")
            resumen.append((tabla, "OMITIDA", 0))
            continue

        df = bases[clave]
        if MODO_PRUEBA:
            df = df.head(REGISTROS_PRUEBA)

        meses = sorted(int(m) for m in df["mes_orden"].unique())
        registros = df.to_dict(orient="records")

        try:
            limpiar_tabla(supabase, tabla, meses)
            subir_registros(supabase, tabla, registros)
            resumen.append((tabla, "OK", len(registros)))
        except Exception as error:
            print(f"❌ {tabla}: {error}")
            resumen.append((tabla, "ERROR", 0))

    print("\n" + "=" * 70)
    print(" RESUMEN")
    print("=" * 70)
    print(f"Archivo: {archivo.name}")
    for tabla, estado, cantidad in resumen:
        print(f"{'✅' if estado == 'OK' else '❌'} {tabla:<28} {estado:<8} {cantidad:>8,}")
    print(f"\nTiempo total: {time.time() - inicio_total:.1f} s\n")

    if any(estado == "ERROR" for _, estado, _ in resumen):
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"\n❌ {error}")
        sys.exit(1)