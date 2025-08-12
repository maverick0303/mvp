import pandas as pd
from datetime import datetime, timedelta
from jinja2 import Environment, FileSystemLoader

# --- Configuración ---
RUTA_EXCEL = "usuarios.xlsx"  # tu archivo Excel con id_usuario, id_jefatura, nombre, correo, ultimo_login
RUTA_TEMPLATES = "templates"  # carpeta donde están los HTML
UMBRAL_INACTIVO = 90          # días de inactividad para marcar como inactivo
UMBRAL_DESACTIVADO = 120      # días de inactividad para marcar como desactivado

# --- Leer Excel ---
df = pd.read_excel(RUTA_EXCEL)

# Asegurar formato fecha
df["ultimo_login"] = pd.to_datetime(df["ultimo_login"]).dt.date

# Calcular días inactivos y estado
hoy = datetime.now().date()
df["dias_inactivo"] = (hoy - df["ultimo_login"]).dt.days

def definir_estado(dias):
    if dias >= UMBRAL_DESACTIVADO:
        return "Desactivado"
    elif dias >= UMBRAL_INACTIVO:
        return "Inactivo"
    else:
        return "Activo"

df["estado"] = df["dias_inactivo"].apply(definir_estado)

# Filtrar solo usuarios candidatos a notificación
candidatos = df[df["estado"].isin(["Inactivo", "Desactivado"])].copy()

# --- Configurar Jinja2 ---
env = Environment(loader=FileSystemLoader(RUTA_TEMPLATES))

# Cargar plantillas
tpl_usuario = env.get_template("correo_usuario.html")
tpl_jefatura = env.get_template("correo_jefatura.html")

# --- Generar correos individuales para usuarios ---
for _, row in candidatos.iterrows():
    html_usuario = tpl_usuario.render(
        NOMBRE_USUARIO=row["nombre"],
        DIAS_INACTIVO=row["dias_inactivo"],
        ESTADO=row["estado"],
        FECHA_LIMITE=(hoy + timedelta(days=14)).strftime("%d/%m/%Y"),
        NOMBRE_JEFATURA="(Nombre de la Jefatura)"  # luego se puede leer de otra tabla o del Excel
    )
    with open(f"correo_usuario_{row['id_usuario']}.html", "w", encoding="utf-8") as f:
        f.write(html_usuario)

# --- Generar un correo de resumen para jefatura ---
# Para este MVP, generamos uno general, pero luego se puede separar por id_jefatura
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

print("✅ Correos generados. Revisa los archivos HTML creados.")
