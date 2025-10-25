# -*- coding: utf-8 -*-
import io
import re
import unicodedata
import pandas as pd
import streamlit as st
from pdfminer.high_level import extract_text
from docx import Document
from google.oauth2.service_account import Credentials

# ─────────────────────────────
# CONFIGURACIÓN BÁSICA
# ─────────────────────────────
st.set_page_config(page_title="Extractor de ACTAS del Consejo", page_icon="📑", layout="wide")
st.title("📑 Extractor de ACTAS del Consejo de Investigación → CSV / Google Sheets")

# ─────────────────────────────
# FUNCIÓN: Inferir AÑO
# ─────────────────────────────
SPANISH_YEAR_WORDS = {
    "diecisiete": 2017, "dieciocho": 2018, "diecinueve": 2019,
    "veinte": 2020, "veintiuno": 2021,
    "veintidos": 2022, "veintidós": 2022,
    "veintitres": 2023, "veintitrés": 2023,
    "veinticuatro": 2024, "veinticinco": 2025,
    "veintiseis": 2026, "veintiséis": 2026,
    "veintisiete": 2027, "veintiocho": 2028,
    "veintinueve": 2029, "treinta": 2030
}

def infer_year_from_text(s: str):
    if not isinstance(s, str):
        return None
    m = re.search(r"\b(20\d{2})\b", s)
    if m:
        return int(m.group(1))
    t = unicodedata.normalize("NFKD", s).lower()
    t = t.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
    m = re.search(r"dos mil\s+([a-z]+)", t)
    if m:
        return SPANISH_YEAR_WORDS.get(m.group(1))
    return None

# ─────────────────────────────
# FUNCIÓN: Extraer texto
# ─────────────────────────────
def extract_text_from_any(file):
    if file.name.endswith(".pdf"):
        return extract_text(file)
    elif file.name.endswith(".docx"):
        doc = Document(file)
        return "\n".join([p.text for p in doc.paragraphs])
    else:
        return None

# ─────────────────────────────
# FUNCIÓN: Extraer estructura de datos desde texto
# ─────────────────────────────
def parse_acta_text(text, filename):
    lines = text.split("\n")
    entries = []
    acta_num = filename.replace(".pdf", "").replace(".docx", "")
    fecha = next((l for l in lines if "mil" in l.lower()), "")
    facultad = next((l for l in lines if "Facultad" in l or "Escuela" in l), "")
    current_type, title, director, estado = "", "", "", ""

    for line in lines:
        if "Proyecto" in line or "Categoría" in line or "Informe" in line:
            current_type = line.strip()
        if re.search(r"Director", line):
            director = line.strip()
        if re.search(r"Aprobado|Elevado|Rechazado|Observaciones", line, re.IGNORECASE):
            estado = line.strip()
            entries.append({
                "Acta": acta_num,
                "Fecha": fecha,
                "Facultad": facultad,
                "Tipo_tema": current_type,
                "Titulo_o_denominacion": title,
                "Director": director,
                "Estado": estado,
                "Fuente_archivo": filename
            })
            title, director, estado = "", "", ""
        if "Título" in line or "Titulo" in line:
            title = line.strip()

    return entries

# ─────────────────────────────
# INTERFAZ
# ─────────────────────────────
uploaded_files = st.file_uploader("📂 Subí tus actas (.pdf o .docx)", type=["pdf", "docx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        text = extract_text_from_any(f)
        if not text:
            st.warning(f"No se pudo leer {f.name}")
            continue
        data = parse_acta_text(text, f.name)
        all_data.extend(data)

    if not all_data:
        st.error("No se detectaron datos válidos en las actas.")
        st.stop()

    df = pd.DataFrame(all_data)
    df.insert(0, "Año", df["Fecha"].apply(infer_year_from_text))

    st.success("✅ Actas procesadas correctamente.")
    st.dataframe(df, use_container_width=True)

    # ───────────────
    # DESCARGAS
    # ───────────────
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Actas")
    excel_buf.seek(0)

    st.download_button("📘 Descargar Excel", data=excel_buf,
                       file_name="Actas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    csv_buf = io.StringIO()
    df.to_csv(csv_buf, index=False)
    st.download_button("📗 Descargar CSV", data=csv_buf.getvalue(),
                       file_name="Actas.csv", mime="text/csv")

    # ───────────────
    # SUBIR A GOOGLE DRIVE (opcional)
    # ───────────────
    st.subheader("Subir a Google Drive como hoja de cálculo (opcional)")

    try:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"])
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload

        if st.button("🚀 Crear hoja nativa en Drive"):
            drive_service = build("drive", "v3", credentials=creds)
            csv_bytes = df.to_csv(index=False).encode("utf-8")
            media = MediaIoBaseUpload(io.BytesIO(csv_bytes), mimetype="text/csv", resumable=False)
            file_metadata = {
                "name": "Actas Consejo",
                "mimeType": "application/vnd.google-apps.spreadsheet",
                "parents": [st.secrets.get("drive_folder_id", "1O7xo7cCGkSujhUXIl3fv51S-cniVLmnh")]
            }
            f = drive_service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink").execute()
            st.success(f"✅ Hoja creada correctamente: [Abrir en Drive]({f['webViewLink']})")
    except Exception as e:
        st.warning(f"⚠️ No se pudieron usar las credenciales de Google: {e}")

else:
    st.info("Subí tus actas en formato PDF o Word para comenzar.")
