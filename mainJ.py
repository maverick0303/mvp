import os
import pandas as pd
from datetime import datetime, timedelta
from jinja2 import Environment, FileSystemLoader

# --- Limpiar pantalla al iniciar ---
os.system('cls' if os.name == 'nt' else 'clear')

# --- Configuración ---
RUTA_EXCEL = "jefatura.xlsx"       # nuevo archivo Excel con jefaturas
RUTA_TEMPLATES = "notificaciones/templates"  # revisa que sea correcta la ruta
UMBRAL_INACTIVO = 90
UMBRAL_DESACTIVADO = 120
OUTPUT_DIR = "correos_jefatura"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Leer Excel ---
print("📂 Leyendo Excel jefatura...")
df = pd.read_excel(RUTA_EXCEL)

# --- Limpieza de datos ---
print("🧹 Limpiando datos...")
df.columns = df.columns.str.strip().str.lower()
if "correo" in df.columns:
    df["correo"] = df["correo"].astype(str).str.strip().str.lower()
df = df.drop_duplicates(subset=["id_usuario"]).copy()

# --- Fechas y días inactivos ---
print("📅 Convirtiendo fechas...")
df["ultimo_login"] = pd.to_datetime(df["ultimo_login"], errors="coerce")

print("⏳ Calculando días inactivos...")
hoy = pd.Timestamp.today().normalize()
df["dias_inactivo"] = (hoy - df["ultimo_login"]).dt.days
df["dias_inactivo"] = df["dias_inactivo"].fillna(999).astype(int)

# --- Estado ---
def definir_estado(dias):
    if dias >= UMBRAL_DESACTIVADO:
        return "Desactivado"
    elif dias >= UMBRAL_INACTIVO:
        return "Inactivo"
    else:
        return "Activo"

df["estado"] = df["dias_inactivo"].apply(definir_estado)

# --- Filtrar candidatos ---
print("📋 Filtrando usuarios inactivos/desactivados...")
candidatos = df[df["estado"].isin(["Inactivo", "Desactivado"])].copy()

# --- Configurar Jinja2 ---
print("✉️ Cargando plantillas...")
env = Environment(loader=FileSystemLoader(RUTA_TEMPLATES))

tpl_usuario = env.get_template("correo_usuario.html")
tpl_jefatura = env.get_template("correo_jefatura.html")

# --- Generar correos individuales para usuarios ---
print("📨 Generando correos para usuarios...")
for _, row in candidatos.iterrows():
    html_usuario = tpl_usuario.render(
        NOMBRE_USUARIO=row["nombre"],
        DIAS_INACTIVO=row["dias_inactivo"],
        ESTADO=row["estado"],
        FECHA_LIMITE=(hoy + timedelta(days=14)).strftime("%d/%m/%Y"),
        NOMBRE_JEFATURA=row["id_jefatura"]  # aquí puedes mejorar con nombre real si tienes
    )
    archivo_usuario = os.path.join(OUTPUT_DIR, f"correo_usuario_{row['id_usuario']}.html")
    with open(archivo_usuario, "w", encoding="utf-8") as f:
        f.write(html_usuario)

# --- Generar correos resumen agrupados por jefatura ---
print("📨 Generando correos resumen para jefaturas...")
for id_jef, grupo in candidatos.groupby("id_jefatura"):
    filas_html = ""
    for _, row in grupo.iterrows():
        filas_html += f"<tr><td>{row['id_usuario']}</td><td>{row['nombre']}</td><td>{row['dias_inactivo']}</td><td>{row['estado']}</td></tr>"

    html_jefatura = tpl_jefatura.render(
        NOMBRE_JEFATURA=id_jef,  # o usa otro campo con nombre real de jefatura
        N_ENVIADOS=len(grupo),
        FILAS_TABLA=filas_html
    )
    archivo_jefatura = os.path.join(OUTPUT_DIR, f"correo_jefatura_{id_jef}.html")
    with open(archivo_jefatura, "w", encoding="utf-8") as f:
        f.write(html_jefatura)

print(f"✅ Proceso completado. Correos guardados en '{OUTPUT_DIR}'.")
