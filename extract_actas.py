import os
import streamlit as st
import pandas as pd
from google.cloud import documentai_v1 as documentai
from google.oauth2 import service_account
from openpyxl import Workbook

# ==============================
# CONFIGURACIÓN DEL PROYECTO
# ==============================
PROJECT_ID = "extractor-de-texto-476314"
LOCATION = "us"  # región donde creaste el procesador
PROCESSOR_ID = "9d07fab065b8b880"  # ID de tu procesador Document AI

# Ruta del archivo JSON de credenciales (descargado de Google Cloud)
CREDENTIALS_PATH = "client_secret_1050909706701-ilv4mom0r2do2dppsunif1ip6o428hcn.apps.googleusercontent.com.json"

# ==============================
# FUNCIÓN PRINCIPAL DE EXTRACCIÓN
# ==============================
def extract_text_with_document_ai(file_path):
    """Extrae texto del PDF usando Google Document AI."""
    credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_PATH)
    client = documentai.DocumentProcessorServiceClient(credentials=credentials)

    name = f"projects/{PROJECT_ID}/locations/{LOCATION}/processors/{PROCESSOR_ID}"

    with open(file_path, "rb") as f:
        document = {"content": f.read(), "mime_type": "application/pdf"}

    request = {"name": name, "raw_document": document}
    result = client.process_document(request=request)

    text = result.document.text
    return text

# ==============================
# FLUJO STREAMLIT
# ==============================
def main():
    st.title("📄 Extractor de Actas del Consejo de Investigación – UCCuyo")
    st.write("Esta aplicación utiliza Google Document AI para extraer texto automáticamente de los archivos PDF de actas y generar una base institucional estructurada.")

    uploaded_files = st.file_uploader("Subí uno o más archivos PDF", type=["pdf"], accept_multiple_files=True)

    if uploaded_files:
        data = []
        for uploaded_file in uploaded_files:
            with open(uploaded_file.name, "wb") as f:
                f.write(uploaded_file.read())

            st.write(f"Procesando **{uploaded_file.name}**...")
            extracted_text = extract_text_with_document_ai(uploaded_file.name)
            data.append({"Archivo": uploaded_file.name, "Texto extraído": extracted_text[:5000]})

        df = pd.DataFrame(data)
        st.dataframe(df)

        # Exportar resultados a Excel
        output_path = "resultados_actas.xlsx"
        df.to_excel(output_path, index=False)
        st.success("✅ Proceso completado. Podés descargar el archivo Excel:")
        with open(output_path, "rb") as f:
            st.download_button("Descargar resultados", f, file_name=output_path)

if __name__ == "__main__":
    main()
