import os
import smtplib
import pandas as pd
from datetime import datetime, timedelta
from jinja2 import Environment, FileSystemLoader
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

# --- Configuración ---
RUTA_EXCEL = "usuarios.xlsx"
RUTA_JEFATURAS = "jefatura.xlsx"
RUTA_TEMPLATES = "notificaciones/templates"
RUTA_ENVIADOS = "correos_enviados.csv"  # ahora guarda correos, no IDs
RUTA_JEFES_ENVIADOS = "jefes_enviados.csv"

UMBRAL_INACTIVO = 90
UMBRAL_DESACTIVADO = 120

REMITENTE = "pruebasunisimple@gmail.com"
CLAVE_APLICACION = "xvdn naan ivnn dhci"
PAUSA_SEGUNDOS = 1

# --- Leer archivos ---
df = pd.read_excel(RUTA_EXCEL)
jef = pd.read_excel(RUTA_JEFATURAS)
df.columns = df.columns.str.strip().str.lower()
jef.columns = jef.columns.str.strip().str.lower()
df["correo"] = df["correo"].astype(str).str.strip().str.lower()
jef["correo"] = jef["correo"].astype(str).str.strip().str.lower()
df = df.drop_duplicates(subset=["id_usuario"]).copy()

# --- Fechas ---
df["ultimo_login"] = pd.to_datetime(df["ultimo_login"], errors="coerce", dayfirst=True)
hoy = pd.Timestamp.today().normalize()
df["dias_inactivo"] = (hoy - df["ultimo_login"]).dt.days.fillna(999).astype(int)

def definir_estado(dias):
    if dias >= UMBRAL_DESACTIVADO:
        return "Desactivado"
    elif dias >= UMBRAL_INACTIVO:
        return "Inactivo"
    return "Activo"

df["estado"] = df["dias_inactivo"].apply(definir_estado)

# --- Candidatos por inactividad ---
candidatos = df[df["estado"].isin(["Inactivo", "Desactivado"])].copy()

# --- Preparar jefaturas ---
jef_unicas = (
    jef.dropna(subset=["id_jefatura"])
       .drop_duplicates(subset=["id_jefatura"])
       .rename(columns={"nombre": "nombre_jefatura", "correo": "correo_jefatura"})
       [["id_jefatura", "nombre_jefatura", "correo_jefatura"]]
)

cand_jef = candidatos.merge(jef_unicas, on="id_jefatura", how="left")

# --- Correos ya enviados (usuarios) ---
if os.path.exists(RUTA_ENVIADOS):
    enviados = pd.read_csv(RUTA_ENVIADOS)["correo"].tolist()
else:
    enviados = []

# --- Plantillas Jinja2 ---
env = Environment(loader=FileSystemLoader(RUTA_TEMPLATES))
tpl_usuario = env.get_template("correo_usuario.html")
tpl_jefatura = env.get_template("correo_jefatura.html")

# --- Conexión segura a Gmail ---
server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server.login(REMITENTE, CLAVE_APLICACION)

# --- Envío a usuarios ---
print("📨 Enviando correos a usuarios...")
correos_nuevos_enviados = []

for _, row in cand_jef.iterrows():
    correo_destino = row["correo"]

    if correo_destino in enviados:
        print(f"⏭ Ya fue enviado a: {correo_destino}, se omite.")
        continue

    html_usuario = tpl_usuario.render(
        NOMBRE_USUARIO=row["nombre"],
        DIAS_INACTIVO=row["dias_inactivo"],
        ESTADO=row["estado"],
        FECHA_LIMITE=(hoy + timedelta(days=14)).strftime("%d/%m/%Y"),
        NOMBRE_JEFATURA=row["nombre_jefatura"]
    )

    msg = MIMEMultipart('alternative')
    msg['From'] = REMITENTE
    msg['To'] = correo_destino
    msg['Subject'] = "Notificación de Inactividad"
    msg.attach(MIMEText(html_usuario, 'html'))

    try:
        server.sendmail(REMITENTE, correo_destino, msg.as_string())
        print(f"✅ Enviado a: {correo_destino}")
        correos_nuevos_enviados.append(correo_destino)
    except Exception as e:
        print(f"❌ Error enviando a {correo_destino}: {e}")
    time.sleep(PAUSA_SEGUNDOS)

# --- Guardar usuarios enviados ---
if correos_nuevos_enviados:
    df_nuevos = pd.DataFrame({"correo": correos_nuevos_enviados})
    if os.path.exists(RUTA_ENVIADOS):
        df_nuevos.to_csv(RUTA_ENVIADOS, mode="a", header=False, index=False)
    else:
        df_nuevos.to_csv(RUTA_ENVIADOS, index=False)

# --- Evitar reenviar resumen a jefaturas ---
if os.path.exists(RUTA_JEFES_ENVIADOS):
    jefes_ya_enviados = pd.read_csv(RUTA_JEFES_ENVIADOS)["id_jefatura"].tolist()
else:
    jefes_ya_enviados = []

print("📨 Enviando correos a jefaturas...")
jefes_enviados = []

for id_jef, g in cand_jef.groupby("id_jefatura"):
    if id_jef in jefes_ya_enviados:
        print(f"⏭ Ya se notificó a jefatura {id_jef}, se omite.")
        continue

    nombre_jef = g["nombre_jefatura"].iloc[0] or f"Jefatura {id_jef}"
    correo_destino = g["correo_jefatura"].iloc[0]

    filas_html = "".join(
        f"<tr><td>{r['id_usuario']}</td><td>{r['nombre']}</td><td>{r['dias_inactivo']}</td><td>{r['estado']}</td></tr>"
        for _, r in g.iterrows()
    )

    html_jef = tpl_jefatura.render(
        NOMBRE_JEFATURA=nombre_jef,
        N_ENVIADOS=len(g),
        FILAS_TABLA=filas_html
    )

    msg = MIMEMultipart('alternative')
    msg['From'] = REMITENTE
    msg['To'] = correo_destino
    msg['Subject'] = "Resumen de Cuentas Inactivas"
    msg.attach(MIMEText(html_jef, 'html'))

    try:
        server.sendmail(REMITENTE, correo_destino, msg.as_string())
        print(f"✅ Resumen enviado a jefatura: {correo_destino}")
        jefes_enviados.append(id_jef)
    except Exception as e:
        print(f"❌ Error enviando a jefatura {correo_destino}: {e}")
    time.sleep(PAUSA_SEGUNDOS)

server.quit()
print("📤 Todos los correos enviados.")

# --- Guardar jefaturas notificadas ---
if jefes_enviados:
    df_nuevos = pd.DataFrame({"id_jefatura": jefes_enviados})
    if os.path.exists(RUTA_JEFES_ENVIADOS):
        df_nuevos.to_csv(RUTA_JEFES_ENVIADOS, mode="a", header=False, index=False)
    else:
        df_nuevos.to_csv(RUTA_JEFES_ENVIADOS, index=False)
