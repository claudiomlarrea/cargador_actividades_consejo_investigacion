# -*- coding: utf-8 -*-
import io
import re
import csv
import unicodedata
import pandas as pd
import streamlit as st

from google.oauth2.service_account import Credentials

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN BÁSICA
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Cargador de Actas → Sheets", page_icon="📑", layout="wide")
st.title("📑 Cargador de Actas del Consejo → Google Sheets (con campo Año)")

# Carpeta destino por defecto (podés sobreescribir en st.secrets["drive_folder_id"])
DEFAULT_FOLDER_ID = "1O7xo7cCGkSujhUXIl3fv51S-cniVLmnh"

# Diagnóstico rápido (opcional, útil mientras configurás)
with st.expander("Diagnóstico (secrets y carpeta destino)", expanded=False):
    st.write("Secrets disponibles:", list(st.secrets.keys()))
    st.write("drive_folder_id:", st.secrets.get("drive_folder_id", DEFAULT_FOLDER_ID))
    st.write("¿gcp_service_account presente?:", "gcp_service_account" in st.secrets)

# ─────────────────────────────────────────────────────────────
# UTILIDADES: Inferir año desde 'Fecha'
# ─────────────────────────────────────────────────────────────
SPANISH_YEAR_WORDS = {
    "diecisiete": 2017, "dieciocho": 2018, "diecinueve": 2019,
    "veinte": 2020, "veintiuno": 2021,
    "veintidos": 2022, "veintidós": 2022,
    "veintitres": 2023, "veintitrés": 2023,
    "veinticuatro": 2024, "veinticinco": 2025,
    "veintiseis": 2026, "veintiséis": 2026,
    "veintisiete": 2027, "veintiocho": 2028,
    "veintinueve": 2029, "treinta": 2030,
}

def infer_year_from_text(s: str):
    if not isinstance(s, str):
        return None
    # Caso 1: hay un año numérico
    m = re.search(r"\b(20\d{2})\b", s)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            pass
    # Caso 2: “dos mil veinticuatro”, etc.
    t = unicodedata.normalize("NFKD", s).lower()
    t = t.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
    m = re.search(r"dos mil\s+([a-z]+)", t)
    if m:
        return SPANISH_YEAR_WORDS.get(m.group(1))
    return None

# ─────────────────────────────────────────────────────────────
# GOOGLE DRIVE HELPERS (imports perezosos para no romper si falta la lib)
# ─────────────────────────────────────────────────────────────
def get_creds(scopes):
    sa = st.secrets.get("gcp_service_account")
    if not sa:
        return None
    if isinstance(sa, dict):
        return Credentials.from_service_account_info(sa, scopes=scopes)
    # Si por alguna razón vino como string JSON (no usual en Streamlit), intentar parsear
    try:
        import json
        return Credentials.from_service_account_info(json.loads(sa), scopes=scopes)
    except Exception:
        return None

def delete_existing_by_name_in_folder(drive, name: str, folder_id: str):
    # Escapar comillas simples en nombre
    safe_name = name.replace("'", "\\'")
    q = (
        "name = '{}' and '{}' in parents and "
        "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    ).format(safe_name, folder_id)
    res = drive.files().list(q=q, fields="files(id)").execute()
    for f in res.get("files", []):
        drive.files().delete(fileId=f["id"]).execute()

def create_native_sheet_in_folder_from_df(df: pd.DataFrame, name: str, folder_id: str, creds):
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
    except ModuleNotFoundError:
        st.error(
            "Falta la librería `google-api-python-client`. "
            "Agregá `google-api-python-client` a tu requirements.txt y reiniciá la app."
        )
        return None

    drive = build("drive", "v3", credentials=creds)

    # (Opcional) eliminar duplicados por nombre en la carpeta
    delete_existing_by_name_in_folder(drive, name, folder_id)

    # Subir CSV en memoria y que se convierta en Hoja nativa
    csv_bytes = df.to_csv(index=False).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(csv_bytes), mimetype="text/csv", resumable=False)
    metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [folder_id],
    }
    file = drive.files().create(body=metadata, media_body=media, fields="id, webViewLink").execute()
    return file.get("webViewLink")

# ─────────────────────────────────────────────────────────────
# UI: Carga de archivo
# ─────────────────────────────────────────────────────────────
file = st.file_uploader("📂 Subí el archivo de Actas (.xlsx / .csv / .xls)", type=["xlsx", "csv", "xls"])

if not file:
    st.info("Subí un archivo para comenzar.")
    st.stop()

# Leer archivo
try:
    if file.name.lower().endswith(".csv"):
        df = pd.read_csv(file, encoding="utf-8", sep=",", on_bad_lines="skip")
    else:
        df = pd.read_excel(file)
except Exception as e:
    st.error(f"No pude leer el archivo: {e}")
    st.stop()

if df.empty:
    st.warning("El archivo se leyó pero no tiene filas.")
    st.stop()

# Insertar 'Año' como PRIMERA columna (sin romper el resto)
if "Fecha" in df.columns:
    df.insert(0, "Año", df["Fecha"].apply(infer_year_from_text))
else:
    df.insert(0, "Año", None)

st.success("✅ Archivo procesado. Se agregó la columna 'Año'.")
st.dataframe(df, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# Descargas locales
# ─────────────────────────────────────────────────────────────
st.subheader("Descargar")
# Excel
buf_xlsx = io.BytesIO()
try:
    with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Actas")
except Exception:
    buf_xlsx = io.BytesIO()
    with pd.ExcelWriter(buf_xlsx, engine="xlsxwriter") as w:
        df.to_excel(w, index=False, sheet_name="Actas")
buf_xlsx.seek(0)

st.download_button(
    "📘 Descargar Excel (Actas.xlsx)",
    data=buf_xlsx,
    file_name="Actas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# CSV
st.download_button(
    "📗 Descargar CSV (Actas.csv)",
    data=df.to_csv(index=False).encode("utf-8"),
    file_name="Actas.csv",
    mime="text/csv"
)

# ─────────────────────────────────────────────────────────────
# Creación automática en Google Drive (Hoja nativa)
# ─────────────────────────────────────────────────────────────
st.subheader("Crear Hoja de Cálculo de Google nativa en Drive")
folder_id = st.secrets.get("drive_folder_id", DEFAULT_FOLDER_ID)
default_name = "Actas - Consejo"
sheet_name = st.text_input("Nombre de la hoja en Drive", value=default_name)

scopes = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/drive.file"]
creds = get_creds(scopes)

if not creds:
    st.warning("No encuentro credenciales en `st.secrets['gcp_service_account']`. "
               "Cargalas en *Settings → Secrets*.")
else:
    if st.button("🚀 Crear hoja nativa en Drive"):
        link = create_native_sheet_in_folder_from_df(df, sheet_name, folder_id, creds)
        if link:
            st.success("✅ Hoja creada correctamente en tu carpeta de Drive.")
            st.write("Abrir:", link)
