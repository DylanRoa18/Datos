"""
============================================================
 CREAR USUARIOS DEL DASHBOARD (en bloque, con usuario y contraseña)
============================================================
Cada persona entra con un USUARIO (no con correo). El script lo genera
a partir del nombre, junto con una contraseña personalizada.

  Dylan Alejandro Rodriguez Roa (admin)  ->  usuario  DARODRIGUEZ
                                             clave    DyRo482*
  Ruta CTR_1                             ->  usuario  CTR1
                                             clave    Ctr1735*

Paso 1 · Ejecuta el script una vez: como aún no existe el Excel, crea
         "usuarios_dashboard.xlsx" con líderes, supervisores y rutas.
         (También: python crear_usuarios.py --plantilla para rehacerla.)

Paso 2 · Revisa el Excel:
         - NOMBRE en orden "Nombres Apellidos" (de ahí sale el usuario).
         - CREAR = NO para quien no deba tener acceso.
         - Agrega filas de administradores (ROL = admin, ALCANCE vacío).
         - ROL = consulta (ALCANCE vacío): ve Seguimiento y Comisiones
           de todos, pero NO Material POP.
         - USUARIO y CONTRASENA vacíos = se generan solos. Si escribes
           algo, se respeta.

Paso 3 · Ejecuta el script otra vez: ahora sí crea las cuentas.
         El mismo Excel queda con el USUARIO y la CONTRASENA de cada uno
         para que los entregues. Guárdalo en un lugar seguro.

Mes nuevo · Después de sincronizar un mes, ejecuta este script: agrega al
         Excel los líderes, supervisores y rutas que aún no tienen usuario
         (CREAR = REVISAR). Revisa su NOMBRE, pon CREAR = SI y ejecútalo
         otra vez para crearles la cuenta.

Roles: admin · consulta · lider · supervisor · ruta
Usa la misma .env de sincronizar_supabase.py (service_role key).
============================================================
"""

import os
import re
import secrets
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from supabase import create_client

CARPETA = Path(__file__).resolve().parent
ARCHIVO = CARPETA / "usuarios_dashboard.xlsx"

# Supabase necesita un correo por cuenta: se arma uno interno a partir del
# usuario (nunca se envían correos). Debe ser IGUAL al del index.html.
DOMINIO_USUARIOS = "tablero.local"

# Contraseña corta = 2 letras del nombre + 2 del apellido + 3 dígitos al azar + "*"
#   DyRo482*   (en rutas: código de la ruta, Ctr1735*)
# Los dígitos al azar evitan que alguien adivine la clave de otro.
DIGITOS_AL_AZAR = 3

# True = a TODAS las filas (con CREAR distinto de NO) se les genera y aplica
# una contraseña NUEVA con el formato corto, aunque ya tengan una en el Excel.
# Úsalo una vez para pasar a todos a claves cortas y vuelve a dejarlo en False.
# (También: python crear_usuarios.py --nuevas-claves)
NUEVAS_CLAVES_PARA_TODOS = False

ROLES = {"admin", "consulta", "lider", "supervisor", "ruta"}
SIN_ALCANCE = {"admin", "consulta"}   # ven a todos
NOMBRE_ROL = {"admin": "Administrador", "consulta": "Consulta", "lider": "Líder", "supervisor": "Supervisor", "ruta": "Ruta"}
PALABRA_ROL = {"admin": "Admin", "consulta": "Consulta", "lider": "Lider", "supervisor": "Super", "ruta": "Ruta"}

load_dotenv(CARPETA / ".env")
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")


def conectar():
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError("Faltan SUPABASE_URL o SUPABASE_KEY (service_role) en el archivo .env")
    return create_client(SUPABASE_URL, SUPABASE_KEY)


# ------------------------------------------------------------
# Usuario y contraseña a partir del nombre
# ------------------------------------------------------------
def sin_tildes(texto):
    texto = unicodedata.normalize("NFD", str(texto or ""))
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def partes_nombre(nombre):
    """('Nombres', 'Apellidos') suponiendo el orden 'Nombres Apellidos'."""
    palabras = [p for p in re.split(r"\s+", sin_tildes(nombre).strip()) if p]
    if len(palabras) <= 1:
        return palabras, []
    if len(palabras) == 2:
        return palabras[:1], palabras[1:]
    if len(palabras) == 3:
        return palabras[:1], palabras[1:]
    return palabras[:2], palabras[2:]


def generar_usuario(rol, nombre, alcance):
    if rol == "ruta":
        return re.sub(r"[^A-Z0-9]", "", sin_tildes(alcance).upper())
    nombres, apellidos = partes_nombre(nombre)
    iniciales = "".join(n[0] for n in nombres)
    apellido = apellidos[0] if apellidos else ""
    return re.sub(r"[^A-Z0-9]", "", (iniciales + apellido).upper())


def generar_contrasena(rol, nombre, alcance):
    if rol == "ruta":
        codigo = re.sub(r"[^A-Za-z0-9]", "", sin_tildes(alcance))
        base = codigo[:1].upper() + codigo[1:].lower()
    else:
        nombres, apellidos = partes_nombre(nombre)
        n = (nombres[0] if nombres else "Us")[:2]
        a = (apellidos[0] if apellidos else "Er")[:2]
        base = n[:1].upper() + n[1:].lower() + a[:1].upper() + a[1:].lower()
    digitos = "".join(str(secrets.randbelow(10)) for _ in range(DIGITOS_AL_AZAR))
    return f"{base}{digitos}*"


def correo_interno(usuario):
    return f"{usuario.lower()}@{DOMINIO_USUARIOS}"


# ------------------------------------------------------------
# Paso 1: plantilla
# ------------------------------------------------------------
def generar_plantilla(sb):
    filas, desde = [], 0
    while True:
        lote = sb.table("campo_liquidacion_rutas").select("mes_orden,ruta,lider,supervisor") \
            .range(desde, desde + 999).execute().data
        filas += lote
        if len(lote) < 1000:
            break
        desde += 1000
    if not filas:
        raise RuntimeError("No hay datos en campo_liquidacion_rutas. Sincroniza primero.")

    df = pd.DataFrame(filas)
    df = df[df["mes_orden"] == df["mes_orden"].max()]   # estructura del mes más reciente

    vacio = {"USUARIO": "", "CONTRASENA": "", "CREAR": "SI", "ESTADO": ""}
    registros = [{"ROL": "admin", "ALCANCE": "", "NOMBRE": "Dylan Alejandro Rodriguez Roa", **vacio}]
    for lider in sorted(df["lider"].dropna().unique()):
        registros.append({"ROL": "lider", "ALCANCE": lider, "NOMBRE": lider.title(), **vacio})
    for sup in sorted(df["supervisor"].dropna().unique()):
        registros.append({"ROL": "supervisor", "ALCANCE": sup, "NOMBRE": sup.title(), **vacio})
    for ruta in sorted(df["ruta"].dropna().unique()):
        registros.append({"ROL": "ruta", "ALCANCE": ruta, "NOMBRE": f"Ruta {ruta}", **vacio})

    pd.DataFrame(registros).to_excel(ARCHIVO, index=False)
    print(f"✅ Plantilla creada: {ARCHIVO}")
    print(f"   {len(registros)} filas. Revisa que NOMBRE esté como 'Nombres Apellidos'")
    print("   (algunos supervisores vienen al revés en el Excel) y ejecuta: python crear_usuarios.py")


# ------------------------------------------------------------
# Personas y rutas nuevas (aparecen en un mes nuevo)
# ------------------------------------------------------------
def clave_persona(texto):
    """Mismo nombre aunque cambie el orden, tildes o espacios (igual que en Supabase)."""
    return " ".join(sorted(re.sub(r"[^A-Z ]", " ", sin_tildes(str(texto)).upper()).split()))


def clave_ruta(texto):
    return re.sub(r"[^A-Z0-9]", "", sin_tildes(str(texto)).upper())


def leer_estructura(sb):
    filas, desde = [], 0
    while True:
        lote = sb.table("campo_liquidacion_rutas").select("mes_orden,ruta,lider,supervisor") \
            .range(desde, desde + 999).execute().data
        filas += lote
        if len(lote) < 1000:
            break
        desde += 1000
    df = pd.DataFrame(filas)
    return df[df["mes_orden"] == df["mes_orden"].max()] if len(df) else df


def agregar_nuevos(sb, df):
    """
    Agrega al Excel los líderes, supervisores y rutas del mes más reciente que
    aún no tienen fila. Quedan con CREAR = REVISAR para que revises el NOMBRE
    (de ahí sale el usuario) antes de crearlos.
    """
    estructura = leer_estructura(sb)
    if not len(estructura):
        return []
    ya = {(r.strip().lower(), clave_ruta(a) if r.strip().lower() == "ruta" else clave_persona(a))
          for r, a in zip(df["ROL"], df["ALCANCE"])}
    nuevos = []
    for rol, col in (("lider", "lider"), ("supervisor", "supervisor"), ("ruta", "ruta")):
        for alcance in sorted(estructura[col].dropna().unique()):
            clave = clave_ruta(alcance) if rol == "ruta" else clave_persona(alcance)
            if not clave or (rol, clave) in ya:
                continue
            ya.add((rol, clave))
            nuevos.append({"ROL": rol, "ALCANCE": alcance,
                           "NOMBRE": f"Ruta {alcance}" if rol == "ruta" else str(alcance).title(),
                           "USUARIO": "", "CONTRASENA": "", "CREAR": "REVISAR",
                           "ESTADO": "Nuevo: revisa NOMBRE (Nombres Apellidos) y pon CREAR = SI"})
    return nuevos


# ------------------------------------------------------------
# Paso 3: crear / actualizar
# ------------------------------------------------------------
def usuarios_existentes(sb):
    existentes, pagina = {}, 1
    while True:
        lote = sb.auth.admin.list_users(page=pagina, per_page=1000)
        for u in lote:
            if u.email:
                existentes[u.email.lower()] = u.id
        if len(lote) < 1000:
            break
        pagina += 1
    return existentes


def crear_usuarios(sb):
    if not ARCHIVO.exists():
        raise RuntimeError(f"No existe {ARCHIVO.name}. Ejecuta primero: python crear_usuarios.py --plantilla")

    # Si el Excel está abierto no se podrían guardar las contraseñas: se avisa ANTES de tocar nada
    try:
        with open(ARCHIVO, "r+b"):
            pass
    except PermissionError:
        raise RuntimeError(f"Cierra {ARCHIVO.name} en Excel y vuelve a ejecutar (no se cambió nada).")

    nuevas_claves = NUEVAS_CLAVES_PARA_TODOS or "--nuevas-claves" in sys.argv
    if nuevas_claves:
        print("🔑 Modo claves nuevas: se genera una contraseña corta para cada fila.\n")

    df = pd.read_excel(ARCHIVO, dtype=str).fillna("")
    df.columns = [c.strip().upper() for c in df.columns]
    for col in ("USUARIO", "CONTRASENA", "CREAR", "ESTADO"):
        if col not in df.columns:
            df[col] = ""

    # Quien aparece por primera vez en el mes más reciente se agrega al Excel
    nuevos = agregar_nuevos(sb, df)
    if nuevos:
        df = pd.concat([df, pd.DataFrame(nuevos)], ignore_index=True).fillna("")
        print(f"🆕 {len(nuevos)} nuevos agregados al Excel (CREAR = REVISAR):")
        for n in nuevos:
            print(f"     {NOMBRE_ROL[n['ROL']]:<11} {n['ALCANCE']}")
        print("   Revisa su NOMBRE (Nombres Apellidos), cambia CREAR a SI y vuelve a ejecutar.\n")

    existentes = usuarios_existentes(sb)
    usados = set()
    creados = actualizados = errores = 0

    for i, f in df.iterrows():
        if f["CREAR"].strip().upper() == "NO":
            df.at[i, "ESTADO"] = "Omitido (CREAR = NO)"
            continue
        if f["CREAR"].strip().upper() == "REVISAR":
            df.at[i, "ESTADO"] = "Nuevo: revisa NOMBRE (Nombres Apellidos) y pon CREAR = SI"
            continue

        rol = f["ROL"].strip().lower()
        alcance = f["ALCANCE"].strip()
        nombre = f["NOMBRE"].strip() or alcance
        if rol not in ROLES:
            df.at[i, "ESTADO"] = f"Error: rol '{rol}' no válido"
            errores += 1
            continue
        if rol not in SIN_ALCANCE and not alcance:
            df.at[i, "ESTADO"] = "Error: falta ALCANCE"
            errores += 1
            continue

        # Usuario: el escrito en el Excel o uno generado (único)
        usuario = re.sub(r"[^A-Z0-9_.-]", "", f["USUARIO"].strip().upper()) or generar_usuario(rol, nombre, alcance)
        base, n = usuario, 2
        while usuario in usados:
            usuario, n = f"{base}{n}", n + 1
        usados.add(usuario)
        email = correo_interno(usuario)

        clave = generar_contrasena(rol, nombre, alcance) if nuevas_claves else f["CONTRASENA"].strip()
        try:
            if email in existentes:
                if clave:
                    sb.auth.admin.update_user_by_id(existentes[email], {"password": clave})
                estado = "Actualizado"
                actualizados += 1
            else:
                clave = clave or generar_contrasena(rol, nombre, alcance)
                sb.auth.admin.create_user({"email": email, "password": clave, "email_confirm": True})
                estado = "Creado"
                creados += 1

            sb.table("usuarios_dashboard").upsert({
                "email": email, "usuario": usuario, "nombre": nombre, "rol": rol,
                "alcance": None if rol in SIN_ALCANCE else alcance, "activo": True,
            }).execute()

            df.at[i, "USUARIO"] = usuario
            df.at[i, "CONTRASENA"] = clave
            df.at[i, "ESTADO"] = estado
            print(f"✅ {usuario:<18} {NOMBRE_ROL[rol]:<14} {alcance}")
        except Exception as error:
            df.at[i, "ESTADO"] = f"Error: {error}"
            print(f"❌ {usuario}: {error}")
            errores += 1

    try:
        df.to_excel(ARCHIVO, index=False)
        destino = ARCHIVO
    except PermissionError:
        # Las cuentas ya se cambiaron: las contraseñas no se pueden perder
        destino = ARCHIVO.with_name(f"usuarios_dashboard_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        df.to_excel(destino, index=False)
        print(f"\n⚠️  {ARCHIVO.name} estaba abierto. Resultados guardados en {destino.name}")
    print(f"\nCreados: {creados} · Actualizados: {actualizados} · Con error: {errores}")
    print(f"📄 Usuarios y contraseñas guardados en: {destino}")
    print("   Este archivo tiene las contraseñas: guárdalo en un lugar seguro y no lo subas a GitHub.")


if __name__ == "__main__":
    try:
        cliente = conectar()
        # Si aún no existe el Excel (o se pide con --plantilla), primero se crea la plantilla.
        # Así funciona también con el botón ▶ de VS Code, que no pasa argumentos.
        if "--plantilla" in sys.argv or not ARCHIVO.exists():
            generar_plantilla(cliente)
            print("\n👉 Abre el Excel, revísalo y vuelve a ejecutar este mismo archivo para crear las cuentas.")
        else:
            crear_usuarios(cliente)
    except Exception as e:
        print(f"\n❌ {e}")
        sys.exit(1)