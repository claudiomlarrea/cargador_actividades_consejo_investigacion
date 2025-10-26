import os
import streamlit as st
import pandas as pd
from google.cloud import documentai_v1 as documentai
from google.oauth2 import service_account

# ==============================
# CONFIGURACIÓN DE GOOGLE CLOUD
# ==============================
PROJECT_ID = "extractor-de-texto-476314"
LOCATION = "us"  # región elegida en Document AI
PROCESSOR_ID = "9d0f7ab065b8b880"  # tu ID de procesador

# Ruta del archivo JSON con las credenciales del servicio
CREDENTIALS_PATH = "client_secret_1050909706701-ilv4mom0r2do2dppsunif1ip6o428hcn.apps.googleusercontent.com.json"

# ==============================
# FUNCIÓN DE EXTRACCIÓN DOCUMENT AI
# ==============================
def extract_text_with_document_ai(file_path):
    """Procesa un PDF con Document AI y devuelve el texto completo."""
    credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_PATH)
    client = documentai.DocumentProcessorServiceClient(credentials=credentials)

    name = f"projects/{PROJECT_ID}/locations/{LOCATION}/processors/{PROCESSOR_ID}"
    with open(file_path, "rb") as f:
        document = {"content": f.read(), "mime_type": "application/pdf"}

    request = {"name": name, "raw_document": document}
    result = client.process_document(request=request)
    return result.document.text

# ==============================
# FUNCIÓN STREAMLIT PRINCIPAL
# ==============================
def main():
    st.title("📄 Extractor de Actas del Consejo de Investigación – UCCuyo")
    st.caption("Usa inteligencia artificial (Document AI de Google Cloud) para leer y estructurar el contenido de las actas institucionales.")

    uploaded_files = st.file_uploader("Subí tus actas en PDF", type=["pdf"], accept_multiple_files=True)

    if uploaded_files:
        resultados = []
        for file in uploaded_files:
            # Guardar PDF temporal
            pdf_path = file.name
            with open(pdf_path, "wb") as f:
                f.write(file.read())

            st.info(f"Procesando **{pdf_path}** ...")
            texto_extraido = extract_text_with_document_ai(pdf_path)
            resultados.append({"Archivo": pdf_path, "Texto extraído": texto_extraido})

        df = pd.DataFrame(resultados)
        st.dataframe(df)

        # Exportar resultados a Excel
        output = "actas_extraidas.xlsx"
        df.to_excel(output, index=False)
        with open(output, "rb") as f:
            st.download_button("📥 Descargar resultados en Excel", f, file_name=output)
        st.success("✅ Extracción completada con Document AI")

if __name__ == "__main__":
    main()
