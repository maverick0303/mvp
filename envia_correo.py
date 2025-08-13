import os
import pandas as pd
from datetime import datetime, timedelta
from jinja2 import Environment, FileSystemLoader
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

# --- Limpiar pantalla al iniciar ---
os.system('cls' if os.name == 'nt' else 'clear')

# --- Configuración ---
RUTA_EXCEL = "usuarios.xlsx"
RUTA_JEFATURAS = "jefatura.xlsx"
RUTA_TEMPLATES = "notificaciones/templates"
UMBRAL_INACTIVO = 90
UMBRAL_DESACTIVADO = 120

REMITENTE = "pruebasunisimple@gmail.com"
CLAVE_APLICACION = "xvdn naan ivnn dhci"
PAUSA_SEGUNDOS = 1  # pausa entre envíos para evitar bloqueos

# --- Leer Excel ---
print("📂 Leyendo Excel...")
df = pd.read_excel(RUTA_EXCEL)
jef = pd.read_excel(RUTA_JEFATURAS)

# --- Limpieza de datos ---
df.columns = df.columns.str.strip().str.lower()
jef.columns = jef.columns.str.strip().str.lower()
df["correo"] = df["correo"].astype(str).str.strip().str.lower()
jef["correo"] = jef["correo"].astype(str).str.strip().str.lower()
df = df.drop_duplicates(subset=["id_usuario"]).copy()

# --- Preparar jefaturas únicas ---
jef_unicas = (
    jef.dropna(subset=["id_jefatura"])
       .drop_duplicates(subset=["id_jefatura"])
       .rename(columns={"nombre": "nombre_jefatura", "correo": "correo_jefatura"})
       [["id_jefatura", "nombre_jefatura", "correo_jefatura"]]
)

# --- Fechas y días inactivos ---
df["ultimo_login"] = pd.to_datetime(df["ultimo_login"], errors="coerce", dayfirst=True)
hoy = pd.Timestamp.today().normalize()
df["dias_inactivo"] = (hoy - df["ultimo_login"]).dt.days.fillna(999).astype(int)

def definir_estado(dias):
    if dias >= UMBRAL_DESACTIVADO:
        return "Desactivado"
    elif dias >= UMBRAL_INACTIVO:
        return "Inactivo"
    else:
        return "Activo"

df["estado"] = df["dias_inactivo"].apply(definir_estado)

# --- Filtrar candidatos ---
candidatos = df[df["estado"].isin(["Inactivo", "Desactivado"])].copy()
# --- Merge con jefaturas para obtener nombre_jefatura ---
cand_jef = candidatos.merge(jef_unicas, on="id_jefatura", how="left")

# --- Configurar Jinja2 ---
env = Environment(loader=FileSystemLoader(RUTA_TEMPLATES))
tpl_usuario = env.get_template("correo_usuario.html")
tpl_jefatura = env.get_template("correo_jefatura.html")

# --- Conectar SMTP ---
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(REMITENTE, CLAVE_APLICACION)

# --- Enviar correos a usuarios ---
print("📨 Enviando correos a usuarios...")
for _, row in cand_jef.iterrows():  # <-- Cambiado a cand_jef
    html_usuario = tpl_usuario.render(
        NOMBRE_USUARIO=row["nombre"],
        DIAS_INACTIVO=row["dias_inactivo"],
        ESTADO=row["estado"],
        FECHA_LIMITE=(hoy + timedelta(days=14)).strftime("%d/%m/%Y"),
        NOMBRE_JEFATURA=row["nombre_jefatura"]  # <-- Ahora tomará correctamente
    )

    msg = MIMEMultipart('alternative')
    msg['From'] = REMITENTE
    msg['To'] = row["correo"]
    msg['Subject'] = "Notificación de Inactividad"
    msg.attach(MIMEText(html_usuario, 'html'))

    try:
        server.sendmail(REMITENTE, row["correo"], msg.as_string())
        print(f"✅ Correo enviado a {row['correo']} con jefatura {row['nombre_jefatura']}")
    except Exception as e:
        print(f"❌ Error enviando a {row['correo']}: {e}")
    time.sleep(PAUSA_SEGUNDOS)

# --- Enviar correos a jefaturas ---
print("📨 Enviando correos a jefaturas...")
for id_jef, g in cand_jef.groupby("id_jefatura"):
    nombre_jef = g["nombre_jefatura"].iloc[0] if pd.notna(g["nombre_jefatura"].iloc[0]) else f"Jefatura {id_jef}"

    filas_html = "".join(
        f"<tr><td>{r['id_usuario']}</td><td>{r['nombre']}</td>"
        f"<td>{r['dias_inactivo']}</td><td>{r['estado']}</td></tr>"
        for _, r in g.iterrows()
    )

    html_jef = tpl_jefatura.render(
        NOMBRE_JEFATURA=nombre_jef,
        N_ENVIADOS=len(g),
        FILAS_TABLA=filas_html
    )

    correo_destino = g["correo_jefatura"].iloc[0]
    msg = MIMEMultipart('alternative')
    msg['From'] = REMITENTE
    msg['To'] = correo_destino
    msg['Subject'] = "Resumen de Cuentas Inactivas"
    msg.attach(MIMEText(html_jef, 'html'))

    try:
        server.sendmail(REMITENTE, correo_destino, msg.as_string())
        print(f"✅ Correo enviado a {correo_destino}")
    except Exception as e:
        print(f"❌ Error enviando a {correo_destino}: {e}")
    time.sleep(PAUSA_SEGUNDOS)

server.quit()
print("📤 Todos los correos enviados.")
