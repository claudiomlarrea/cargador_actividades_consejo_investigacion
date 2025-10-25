import streamlit as st
import pandas as pd
import io
import re
import unicodedata
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import csv

# =============================
# CONFIGURACIÓN DE PÁGINA
# =============================
st.set_page_config(
    page_title="Cargador de Actividades Consejo de Investigación",
    layout="wide",
    page_icon="📑"
)
st.title("📑 Cargador de Actividades del Consejo de Investigación")

# =============================
# FUNCIÓN PARA EXTRAER AÑO
# =============================
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
    # buscar número directo “2024”, “2025”
    m = re.search(r"\b(20\d{2})\b", s)
    if m:
        try:
            return int(m.group(1))
        except:
            pass
    # buscar texto “dos mil veinticuatro”
    t = unicodedata.normalize("NFKD", s).lower()
    t = t.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    m = re.search(r"dos mil\s+([a-z]+)", t)
    if m:
        return SPANISH_YEAR_WORDS.get(m.group(1))
    return None


# =============================
# CARGA DE ARCHIVO
# =============================
uploaded_file = st.file_uploader("📂 Subí el archivo de actas (.xlsx, .csv o .xls)", type=["xlsx", "csv", "xls"])

if uploaded_file:
    try:
        # Leer Excel o CSV
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, encoding="utf-8", sep=",", on_bad_lines="skip")
        else:
            df = pd.read_excel(uploaded_file)

        # Agregar columna Año al inicio
        if "Fecha" in df.columns:
            df.insert(0, "Año", df["Fecha"].apply(infer_year_from_text))
        else:
            df.insert(0, "Año", None)

        st.success("✅ Archivo leído correctamente.")
        st.dataframe(df, use_container_width=True)

        # =============================
        # DESCARGAS LOCALES
        # =============================
        st.subheader("Descargar Excel / CSV")
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Actas")
        buffer_excel.seek(0)

        st.download_button(
            label="📘 Descargar Excel (Actas.xlsx)",
            data=buffer_excel,
            file_name="Actas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        buffer_csv = io.StringIO()
        df.to_csv(buffer_csv, index=False, quoting=csv.QUOTE_NONNUMERIC)
        st.download_button(
            label="📗 Descargar CSV (Actas.csv)",
            data=buffer_csv.getvalue(),
            file_name="Actas.csv",
            mime="text/csv"
        )

        # =============================
        # GOOGLE DRIVE / SHEETS
        # =============================
        st.subheader("Actualizar Google Sheets / Crear Hoja en Drive")

        scopes = [
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ]

        # Obtener credenciales desde secrets
        try:
            sa_info = st.secrets["gcp_service_account"]
            creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
        except Exception as e:
            st.error(f"⚠️ Error al cargar credenciales de Google: {e}")
            creds = None

        if creds:
            service = build("drive", "v3", credentials=creds)

            if st.checkbox("📄 Crear Hoja de Cálculo de Google nativa (conversión automática)"):
                file_metadata = {
                    "name": "Actas - Consejo",
                    "mimeType": "application/vnd.google-apps.spreadsheet"
                }

                # Subir CSV temporal convertido a Google Sheets
                csv_buffer = io.BytesIO()
                df.to_csv(csv_buffer, index=False)
                csv_buffer.seek(0)

                try:
                    uploaded = service.files().create(
                        body=file_metadata,
                        media_body=io.BytesIO(csv_buffer.read()),
                        fields="id"
                    ).execute()
                    sheet_id = uploaded.get("id")
                    st.success(f"✅ Hoja creada correctamente en Drive (ID: {sheet_id})")
                except Exception as e:
                    st.error(f"❌ Error al crear hoja en Drive: {e}")

        else:
            st.warning("No se encontraron credenciales válidas en `st.secrets['gcp_service_account']`.")

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {e}")

else:
    st.info("Subí un archivo Excel o CSV para comenzar.")
