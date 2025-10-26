import os
import io
import time
import pandas as pd
import streamlit as st

# Document AI
from google.cloud import documentai_v1 as documentai
from google.oauth2 import service_account

# ===== (Opcional) Google Sheets: usa tu módulo existente si lo tienes =====
# Debe exponer una función: upload_dataframe_to_sheet(spreadsheet_id: str, sheet_name: str, df: pd.DataFrame)
try:
    from upload_to_sheets import upload_dataframe_to_sheet
    HAS_SHEETS = True
except Exception:
    HAS_SHEETS = False


# ==============================
#  CONFIGURACIÓN
# ==============================
# Recomendado: guarda el JSON de CUENTA DE SERVICIO en la raíz del proyecto (no el client_secret OAuth).
# Ejemplo: "service_account.json"
DEFAULT_SA_PATH = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")

# Tus datos de Document AI
# project_id numérico o string; usa el que ves en la consola (ID del proyecto)
PROJECT_ID = "extractor-de-texto-476314"     # ✏️ cámbialo si hace falta
LOCATION   = "us"                             # "us" o "southamerica-east1", según lo creaste
PROCESSOR_ID = "9d0f7ab065b8b880"             # ✏️ pega tu Processor ID

# Tipos de archivo: enfocamos en PDF
ALLOWED_TYPES = ["pdf"]


# ==============================
#  AUXILIARES
# ==============================
def build_docai_client(sa_path: str):
    """Crea cliente de Document AI desde un JSON de cuenta de servicio."""
    if not os.path.exists(sa_path):
        raise FileNotFoundError(
            f"No encuentro el archivo de credenciales: {sa_path}\n"
            "Descárgalo desde IAM > Cuentas de servicio > tu cuenta > Claves > Agregar clave (JSON)\n"
            "O define la variable de entorno GOOGLE_APPLICATION_CREDENTIALS."
        )
    credentials = service_account.Credentials.from_service_account_file(sa_path)
    return documentai.DocumentProcessorServiceClient(credentials=credentials)


def docai_process_file(client, project_id: str, location: str, processor_id: str, file_bytes: bytes, mime_type: str):
    name = f"projects/{project_id}/locations/{location}/processors/{processor_id}"
    raw_document = documentai.RawDocument(content=file_bytes, mime_type=mime_type)
    request = documentai.ProcessRequest(name=name, raw_document=raw_document)
    return client.process_document(request=request)


def parse_entities_to_rows(doc: documentai.Document) -> pd.DataFrame:
    """
    Convierte las entidades del Custom Extractor a filas.
    - Cada entidad trae: type (nombre de la etiqueta), mention_text (valor) y confidence.
    - Este parser genera una tabla 'Etiqueta | Valor | Confianza | Página'.
    """
    rows = []
    ents = doc.entities or []
    for e in ents:
        rows.append({
            "Etiqueta": e.type_,
            "Valor": e.mention_text,
            "Confianza": round(getattr(e, "confidence", 0.0), 3),
            "Página": (e.page_anchor.page_refs[0].page if (e.page_anchor and e.page_anchor.page_refs) else None)
        })
    return pd.DataFrame(rows)


def summarize_text(text: str, max_chars: int = 3000) -> str:
    """Muestra un recorte para vista previa en la UI."""
    return text if len(text) <= max_chars else text[:max_chars] + "\n...\n[Texto truncado]"


# ==============================
#  APP STREAMLIT
# ==============================
st.set_page_config(page_title="Extractor de Actas – Document AI", layout="wide")

st.title("📄 Extractor de Actas del Consejo (Google Document AI)")
st.caption("Subí tus PDF; la app usa tu **Custom Extractor** de Document AI para extraer etiquetas y texto. Luego descarga Excel/CSV o sube a Google Sheets.")

with st.expander("⚙️ Configuración (si necesitás cambiar algo)"):
    PROJECT_ID = st.text_input("Project ID", PROJECT_ID)
    LOCATION = st.selectbox("Location (región)", ["us", "southamerica-east1"], index=0 if LOCATION=="us" else 1)
    PROCESSOR_ID = st.text_input("Processor ID", PROCESSOR_ID)
    DEFAULT_SA_PATH = st.text_input("Ruta del JSON de cuenta de servicio", DEFAULT_SA_PATH)

st.divider()

uploaded_files = st.file_uploader("Subí uno o más PDFs", type=ALLOWED_TYPES, accept_multiple_files=True)

col_a, col_b = st.columns([1, 1])
with col_a:
    push_to_sheets = st.checkbox("Subir resultados a Google Sheets", value=False and HAS_SHEETS)
with col_b:
    if push_to_sheets and not HAS_SHEETS:
        st.warning("No pude importar `upload_to_sheets`. La subida a Sheets no estará disponible.")
    if push_to_sheets and HAS_SHEETS:
        spreadsheet_id = st.text_input("Spreadsheet ID (Google Sheets)", placeholder="1AbC...ID de tu hoja...")
        sheet_name = st.text_input("Nombre de pestaña (worksheet)", value="Actas")

process_btn = st.button("🚀 Procesar")

if process_btn:
    if not uploaded_files:
        st.error("Subí al menos un PDF.")
        st.stop()

    try:
        client = build_docai_client(DEFAULT_SA_PATH)
    except Exception as e:
        st.error(f"Error al crear cliente de Document AI: {e}")
        st.stop()

    all_rows = []
    previews = []

    progress = st.progress(0)
    for i, file in enumerate(uploaded_files, start=1):
        progress.progress(i / len(uploaded_files))
        pdf_bytes = file.read()

        try:
            result = docai_process_file(
                client=client,
                project_id=PROJECT_ID,
                location=LOCATION,
                processor_id=PROCESSOR_ID,
                file_bytes=pdf_bytes,
                mime_type="application/pdf",
            )
        except Exception as e:
            st.error(f"❌ Error procesando {file.name}: {e}")
            continue

        # Texto completo
        full_text = result.document.text or ""
        previews.append({"Archivo": file.name, "Vista previa del texto": summarize_text(full_text)})

        # Entidades (tus etiquetas del Custom Extractor)
        df_entities = parse_entities_to_rows(result.document)
        if df_entities.empty:
            # Si no hay entidades (no es un Custom Extractor o aún no entrenaste),
            # al menos devolvemos el texto completo como una “fila”
            all_rows.append(pd.DataFrame([{
                "Archivo": file.name,
                "Etiqueta": "TEXTO_COMPLETO",
                "Valor": full_text,
                "Confianza": "",
                "Página": ""
            }]))
        else:
            df_entities.insert(0, "Archivo", file.name)
            all_rows.append(df_entities)

    # Vista previa de texto
    if previews:
        st.subheader("📝 Vista previa (texto extraído)")
        st.dataframe(pd.DataFrame(previews), use_container_width=True)

    # Tabla de etiquetas
    if all_rows:
        result_df = pd.concat(all_rows, ignore_index=True)
        st.subheader("🏷️ Etiquetas extraídas (Custom Extractor)")
        st.dataframe(result_df, use_container_width=True)

        # Descargas
        c1, c2 = st.columns(2)
        with c1:
            csv_bytes = result_df.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Descargar CSV", csv_bytes, file_name="actas_document_ai.csv", mime="text/csv")
        with c2:
            xlsx_path = "actas_document_ai.xlsx"
            result_df.to_excel(xlsx_path, index=False)
            with open(xlsx_path, "rb") as f:
                st.download_button("⬇️ Descargar Excel", f, file_name=xlsx_path)

        # Google Sheets (opcional)
        if push_to_sheets and HAS_SHEETS and spreadsheet_id.strip():
            try:
                upload_dataframe_to_sheet(spreadsheet_id, sheet_name or "Actas", result_df)
                st.success("✅ Subido a Google Sheets correctamente.")
            except Exception as e:
                st.error(f"Error subiendo a Google Sheets: {e}")
