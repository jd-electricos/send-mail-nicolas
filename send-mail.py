import json
import time
import win32com.client as win32
import os
import random
import pandas as pd
from datetime import datetime

# ------------------------------
# Rutas base
# ------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")
HTML_CLIENTES_DIR = os.path.join(BASE_DIR, "html", "clientes")
HTML_DISTRIB_DIR = os.path.join(BASE_DIR, "html", "distribuidor")
EXCEL_DIR = os.path.join(BASE_DIR, "providers")
PDF_DIR = os.path.join(BASE_DIR, "pdf")

EXCEL_PATH = os.path.join(EXCEL_DIR, "datos-clientes.xlsx")

# JSON generales (asesor)
data_path = os.path.join(DATA_DIR, "data.json")

# JSON subjects por rol
SUBJECT_CLIENTES_PATH = os.path.join(DATA_DIR, "subject-clientes.json")
SUBJECT_DISTRIB_PATH = os.path.join(DATA_DIR, "subject-distribuidores.json")

# PDFs por rol
PDF_CLIENTES = os.path.join(PDF_DIR, "JD-JAPS-CATALOGO.pdf")
PDF_DISTRIB = os.path.join(PDF_DIR, "brochure-jd-japs.pdf")

# ------------------------------
# Cargar data asesor
# ------------------------------
with open(data_path, "r", encoding="utf-8") as file:
    data = json.load(file)

asesor = data["asesor"]
correo_asesor = data["correoAsesor"]
numero_asesor = data["numeroAsesor"]

# ------------------------------
# Cargar subjects por rol
# ------------------------------
with open(SUBJECT_CLIENTES_PATH, "r", encoding="utf-8") as file:
    subjects_clientes = json.load(file)["subjects"]

with open(SUBJECT_DISTRIB_PATH, "r", encoding="utf-8") as file:
    subjects_distrib = json.load(file)["subjects"]

# ------------------------------
# Cargar plantillas por rol
# ------------------------------
def cargar_templates(path):
    templates = []
    for filename in sorted(os.listdir(path)):
        if filename.endswith(".html"):
            with open(os.path.join(path, filename), "r", encoding="utf-8") as f:
                templates.append(f.read())
    return templates

templates_clientes = cargar_templates(HTML_CLIENTES_DIR)
templates_distrib = cargar_templates(HTML_DISTRIB_DIR)

# ------------------------------
# Leer Excel
# ------------------------------
df = pd.read_excel(EXCEL_PATH)

df.columns = df.columns.astype(str).str.strip().str.lower()

registros = []
for _, row in df.iterrows():
    nombre = row.get("nombre")
    correo = row.get("correo")
    rol = row.get("rol")

    if pd.notna(nombre):
        registros.append({
            "nombre": str(nombre).strip(),
            "correo": str(correo).strip() if pd.notna(correo) else "",
            "rol": str(rol).strip().lower() if pd.notna(rol) else ""
        })

# ------------------------------
# Outlook
# ------------------------------
outlook = win32.Dispatch('Outlook.Application')

# ------------------------------
# Control
# ------------------------------
novedades = []
correos_enviados = 0

# ------------------------------
# Envío
# ------------------------------
for persona in registros:

    nombre = persona["nombre"]
    correo = persona["correo"]
    rol = persona["rol"]

    if not correo or "@" not in correo:
        novedades.append(f"{nombre}: correo inválido.")
        continue

    # ------------------------------
    # Configuración según rol
    # ------------------------------
    if rol == "cliente":
        templates = templates_clientes
        subjects = subjects_clientes
        pdf_path = PDF_CLIENTES

    elif rol == "distribuidor":
        templates = templates_distrib
        subjects = subjects_distrib
        pdf_path = PDF_DISTRIB

    else:
        novedades.append(f"{nombre}: rol no reconocido ({rol}).")
        continue

    # Validar PDF
    if not os.path.exists(pdf_path):
        novedades.append(f"No se encontró el PDF para rol {rol}.")
        break

    # ------------------------------
    # Crear correo
    # ------------------------------
    template = random.choice(templates)
    subject = random.choice(subjects)

    html_content = (
        template
        .replace("{{ cliente }}", nombre)
        .replace("{{ asesor }}", asesor)
        .replace("{{ correoAsesor }}", correo_asesor)
        .replace("{{ numeroAsesor }}", numero_asesor)
    )

    mail = outlook.CreateItem(0)
    mail.To = correo
    mail.Subject = subject
    mail.HTMLBody = html_content
    mail.Attachments.Add(pdf_path)
    mail.Send()

    correos_enviados += 1

    print(f"✔ Enviado a: {nombre} ({rol}) -> {correo}")

    # ------------------------------
    # Límite horario 5:30 PM
    # ------------------------------
    ahora = datetime.now()
    limite = ahora.replace(hour=17, minute=30, second=0, microsecond=0)

    if ahora > limite:
        novedades.append("Proceso detenido por límite horario.")
        break

    # # ------------------------------
    # # Espera aleatoria
    # # ------------------------------
    # wait_time = random.uniform(120, 240)
    # print(f"⏱ Esperando {wait_time/60:.2f} minutos...\n")
    # time.sleep(wait_time)
    
    # ------------------------------
    # Espera aleatoria (modo prueba)
    # ------------------------------
    wait_time = random.uniform(10, 20)
    print(f"⏱ Esperando {wait_time:.2f} segundos...\n")
    time.sleep(wait_time)

# ------------------------------
# REPORTE FINAL
# ------------------------------
reporte = []
reporte.append("REPORTE ENVÍO CORREOS\n")
reporte.append(f"Correos enviados: {correos_enviados}\n")

if novedades:
    reporte.append("NOVEDADES:")
    for n in novedades:
        reporte.append(f"- {n}")
else:
    reporte.append("Sin novedades.")

reporte_texto = "\n".join(reporte)

mail_reporte = outlook.CreateItem(0)
mail_reporte.To = "nicolasbernal@jdelectricos.com.co"
mail_reporte.Subject = "Reporte envío correos"
mail_reporte.Body = reporte_texto
mail_reporte.Send()

print("📧 Reporte enviado.")
print("🎉 Proceso finalizado.")