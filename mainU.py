import os
import pandas as pd
from datetime import datetime, timedelta
from jinja2 import Environment, FileSystemLoader

# --- Limpiar pantalla al iniciar ---
os.system('cls' if os.name == 'nt' else 'clear')

# --- Configuración ---
RUTA_EXCEL = "usuarios.xlsx"       # tu archivo Excel
RUTA_TEMPLATES = "notificaciones/templates"  # carpeta donde están los HTML
UMBRAL_INACTIVO = 90               # días para marcar como inactivo
UMBRAL_DESACTIVADO = 120           # días para marcar como desactivado

# --- Leer Excel ---
print("📂 Leyendo Excel...")
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

# --- Generar correos de usuario ---
print("📨 Generando correos para usuarios...")
for _, row in candidatos.iterrows():
    html_usuario = tpl_usuario.render(
        NOMBRE_USUARIO=row["nombre"],
        DIAS_INACTIVO=row["dias_inactivo"],
        ESTADO=row["estado"],
        FECHA_LIMITE=(hoy + timedelta(days=14)).strftime("%d/%m/%Y"),
        NOMBRE_JEFATURA="(Nombre de la Jefatura)"  # luego se puede enlazar
    )
    with open(f"correo_usuario_{row['id_usuario']}.html", "w", encoding="utf-8") as f:
        f.write(html_usuario)

# --- Generar correo de resumen para jefatura ---
print("📨 Generando correo de jefatura...")
filas_html = ""
for _, row in candidatos.iterrows():
    filas_html += f"<tr><td>{row['id_usuario']}</td><td>{row['nombre']}</td><td>{row['dias_inactivo']}</td><td>{row['estado']}</td></tr>"

html_jefatura = tpl_jefatura.render(
    NOMBRE_JEFATURA="Carlos Pérez",
    N_ENVIADOS=len(candidatos),
    FILAS_TABLA=filas_html
)

with open("correo_jefatura.html", "w", encoding="utf-8") as f:
    f.write(html_jefatura)

print("✅ Proceso completado. Archivos HTML creados.")
# === Resúmenes por jefatura usando jefatura.xlsx (nombre + correo) ===
# jefatura.xlsx tiene: id_jefatura, id_usuario, nombre (jefatura), correo (jefatura)
jef = pd.read_excel("jefatura.xlsx")
jef.columns = jef.columns.str.strip().str.lower()

# Nos quedamos con un único registro por jefatura (nombre y correo)
jef_unicas = (
    jef.dropna(subset=["id_jefatura"])
       .drop_duplicates(subset=["id_jefatura"])
       .rename(columns={"nombre": "nombre_jefatura", "correo": "correo_jefatura"})
       [["id_jefatura", "nombre_jefatura", "correo_jefatura"]]
)

# Unimos para que cada usuario candidato tenga el nombre/correo de su jefatura
cand_jef = candidatos.merge(jef_unicas, on="id_jefatura", how="left")

# Generamos UN HTML por cada jefatura (saludo con 'solo nombre')
for id_jef, g in cand_jef.groupby("id_jefatura"):
    nombre_jef = g["nombre_jefatura"].iloc[0] if pd.notna(g["nombre_jefatura"].iloc[0]) else f"Jefatura {id_jef}"

    filas_html = "".join(
        f"<tr><td>{r['id_usuario']}</td><td>{r['nombre']}</td>"
        f"<td>{r['dias_inactivo']}</td><td>{r['estado']}</td></tr>"
        for _, r in g.iterrows()
    )

    html_jef = tpl_jefatura.render(
        NOMBRE_JEFATURA=nombre_jef,  # << aquí va SOLO el nombre (ej: 'Pablo', 'Martin')
        N_ENVIADOS=len(g),
        FILAS_TABLA=filas_html
    )

    # Un archivo por jefatura, ej: correo_jefatura_JEF-1.html
    with open(f"correo_jefatura_{id_jef}.html", "w", encoding="utf-8") as f:
        f.write(html_jef)

print("✅ Resúmenes por jefatura generados (uno por cada id_jefatura).")
# === Fin del script ===
