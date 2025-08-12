import os
from pathlib import Path
import pandas as pd
from datetime import timedelta
from jinja2 import Environment, FileSystemLoader

# ── Limpia pantalla (comodidad)
os.system('cls' if os.name == 'nt' else 'clear')

# ── Config
UMBRAL_INACTIVO = 90
UMBRAL_DESACTIVADO = 120

BASE_DIR = Path(__file__).resolve().parent
RUTA_EXCEL_USUARIOS = BASE_DIR / "usuarios.xlsx"
RUTA_EXCEL_JEFATURA = BASE_DIR / "jefatura.xlsx"

# Intenta autodetectar la carpeta de plantillas entre las dos variantes que han usado
POSIBLES_TEMPLATES = [
    BASE_DIR / "notificaciones" / "templates",
    BASE_DIR / "notificacioes" / "templates",
    BASE_DIR / "templates",
]
RUTA_TEMPLATES = next((p for p in POSIBLES_TEMPLATES if p.exists()), None)
if RUTA_TEMPLATES is None:
    raise FileNotFoundError(
        "No encontré la carpeta de plantillas. Crea una en:\n"
        f" - {POSIBLES_TEMPLATES[0]}\n - {POSIBLES_TEMPLATES[1]}\n - {POSIBLES_TEMPLATES[2]}\n"
        "y coloca 'correo_usuario.html' y 'correo_jefatura.html' adentro."
    )

print(f"🔎 Usando carpeta de plantillas: {RUTA_TEMPLATES}")
print("📄 Plantillas disponibles:", [p.name for p in RUTA_TEMPLATES.iterdir()])

# ── Jinja
env = Environment(loader=FileSystemLoader(str(RUTA_TEMPLATES)))
tpl_usuario = env.get_template("correo_usuario.html")
tpl_jefatura = env.get_template("correo_jefatura.html")

# ── Leer usuarios
print("📂 Leyendo usuarios.xlsx ...")
u = pd.read_excel(RUTA_EXCEL_USUARIOS)
u.columns = u.columns.str.strip().str.lower()

# Normaliza campos clave
if "correo" in u.columns:
    u["correo"] = u["correo"].astype(str).str.strip().str.lower()
u["id_jefatura"] = u["id_jefatura"].astype(str).str.strip()

# Validación mínima
requeridas = {"id_usuario", "id_jefatura", "nombre", "correo", "ultimo_login"}
faltan = requeridas - set(u.columns)
if faltan:
    raise ValueError(f"Faltan columnas en usuarios.xlsx: {faltan}")

# Fechas y días/estado
print("📅 Convirtiendo fechas y calculando días...")
u["ultimo_login"] = pd.to_datetime(u["ultimo_login"], errors="coerce")
hoy = pd.Timestamp.today().normalize()
u["dias_inactivo"] = (hoy - u["ultimo_login"]).dt.days
u["dias_inactivo"] = u["dias_inactivo"].fillna(999).astype(int)

def estado_por_dias(d):
    if d >= UMBRAL_DESACTIVADO: return "Desactivado"
    if d >= UMBRAL_INACTIVO:    return "Inactivo"
    return "Activo"

u["estado"] = u["dias_inactivo"].apply(estado_por_dias)

# Filtrar candidatos
candidatos = u[u["estado"].isin(["Inactivo", "Desactivado"])].copy()
print(f"📋 Usuarios candidatos: {len(candidatos)}")

# ── Leer jefaturas
print("📂 Leyendo jefatura.xlsx ...")
j = pd.read_excel(RUTA_EXCEL_JEFATURA)
j.columns = j.columns.str.strip().str.lower()
# Normaliza id y renombra columnas a nombre/correo de jefatura
if "id_jefatura" not in j.columns:
    raise ValueError("En jefatura.xlsx debe existir la columna 'id_jefatura'.")
j["id_jefatura"] = j["id_jefatura"].astype(str).str.strip()
j = j.rename(columns={"nombre": "nombre_jefatura", "correo": "correo_jefatura"})

# Nos quedamos con un registro por jefatura
j_unicas = (
    j.dropna(subset=["id_jefatura"])
     .drop_duplicates(subset=["id_jefatura"])
     [["id_jefatura", "nombre_jefatura", "correo_jefatura"]]
)

# ── Merge: agregamos nombre_jefatura a cada usuario candidato
cand_jef = candidatos.merge(j_unicas, on="id_jefatura", how="left")

# Debug útil: ver si trajo el nombre de jefatura
print("🔎 Muestra post-merge (id, jefatura, nombre_jefatura):")
print(cand_jef[["id_usuario","id_jefatura","nombre_jefatura"]].head())

# ── Generar correos para usuarios (con nombre de jefatura)
print("📨 Generando correos para usuarios...")
gen_usuarios = 0
for _, row in cand_jef.iterrows():
    nombre_jef = row["nombre_jefatura"] if pd.notna(row["nombre_jefatura"]) else f"Jefatura {row['id_jefatura']}"
    html_usuario = tpl_usuario.render(
        NOMBRE_USUARIO=row["nombre"],
        DIAS_INACTIVO=row["dias_inactivo"],
        ESTADO=row["estado"],
        FECHA_LIMITE=(hoy + pd.Timedelta(days=14)).strftime("%d/%m/%Y"),
        NOMBRE_JEFATURA=nombre_jef
    )
    out_user = BASE_DIR / f"correo_usuario_{row['id_usuario']}.html"
    with open(out_user, "w", encoding="utf-8") as f:
        f.write(html_usuario)
    gen_usuarios += 1
print(f"   → Correos de usuario generados: {gen_usuarios}")

# ── Generar resúmenes por jefatura (usa tu plantilla actual con FILAS_TABLA)
print("📧 Generando resúmenes por jefatura...")
gen_jef = 0
for id_jef, g in cand_jef.groupby("id_jefatura"):
    nombre_jef = g["nombre_jefatura"].iloc[0] if pd.notna(g["nombre_jefatura"].iloc[0]) else f"Jefatura {id_jef}"

    filas_html = "".join(
        f"<tr><td>{r['id_usuario']}</td><td>{r['nombre']}</td>"
        f"<td>{r['dias_inactivo']}</td><td>{r['estado']}</td></tr>"
        for _, r in g.iterrows()
    )

    html_j = tpl_jefatura.render(
        NOMBRE_JEFATURA=nombre_jef,
        N_ENVIADOS=len(g),
        FILAS_TABLA=filas_html
    )
    out_j = BASE_DIR / f"correo_jefatura_{id_jef}.html"
    with open(out_j, "w", encoding="utf-8") as f:
        f.write(html_j)
    print(f"   → {out_j.name} (usuarios: {len(g)})")
    gen_jef += 1

print(f"✅ Listo. Usuarios: {gen_usuarios} | Resúmenes por jefatura: {gen_jef}")
