# app.py
# --------------------------------------------------------
# App Streamlit unificada:
# 1) Sube PDFs/DOCX
# 2) Procesa con Document AI (usa tu Custom Extractor)
# 3) Muestra etiquetas y texto
# 4) Descarga CSV/Excel
# 5) Sube a Google Sheets
# --------------------------------------------------------

import io
import streamlit as st
import pandas as pd

from extract_actas import process_with_document_ai, extract_text_local
from upload_to_sheets import upload_dataframe_to_sheet


st.set_page_config(page_title="Extractor de Actas – UCCuyo", layout="wide")
st.title("📑 Extractor de Actas del Consejo de Investigación – UCCuyo")

st.caption(
    "Cargá PDFs/DOCX. La app usa tu **Custom Extractor** de Google Document AI para leer "
    "las actas y devolver **etiquetas** (Proyecto, Director, Equipo, Fecha, etc.) y el texto completo."
)

uploaded_files = st.file_uploader(
    "📤 Subí uno o más archivos (PDF o DOCX)",
    type=["pdf", "docx", "txt"],
    accept_multiple_files=True,
)

process_btn = st.button("🚀 Procesar")

if process_btn:
    if not uploaded_files:
        st.error("Subí al menos un archivo.")
        st.stop()

    all_entities = []   # Lista de dataframes de entidades
    previews = []       # Vista previa de texto

    progress = st.progress(0)
    for idx, uf in enumerate(uploaded_files, start=1):
        progress.progress(idx / len(uploaded_files))

        file_bytes = uf.read()
        mime = "application/pdf" if uf.name.lower().endswith(".pdf") else "application/octet-stream"

        # Intentar Document AI
        try:
            full_text, df_ent = process_with_document_ai(file_bytes, mime_type=mime)
            origin = "Document AI"
        except Exception as e:
            st.warning(f"No se pudo procesar **{uf.name}** con Document AI. Intentando extracción local.\n\n> {e}")
            # Fallback local
            uf.seek(0)  # reset para leer desde extract_text_local
            full_text = extract_text_local(uf)
            df_ent = pd.DataFrame([{"Etiqueta": "TEXTO_COMPLETO", "Valor": full_text, "Confianza": "", "Página": ""}])
            origin = "Local"

        # Guardar vista previa de texto
        preview = full_text if len(full_text) <= 3000 else (full_text[:3000] + "\n...\n[texto truncado]")
        previews.append({"Archivo": uf.name, "Origen": origin, "Vista previa": preview})

        # Agregar columna Archivo a las entidades y acumular
        if not df_ent.empty:
            df_ent.insert(0, "Archivo", uf.name)
        else:
            df_ent = pd.DataFrame([{"Archivo": uf.name, "Etiqueta": "", "Valor": "", "Confianza": "", "Página": ""}])

        all_entities.append(df_ent)

    # Mostrar resultados
    st.subheader("📝 Vista previa del texto")
    st.dataframe(pd.DataFrame(previews), use_container_width=True)

    result_df = pd.concat(all_entities, ignore_index=True) if all_entities else pd.DataFrame()
    st.subheader("🏷️ Etiquetas extraídas")
    st.dataframe(result_df, use_container_width=True)

    # Descargas
    col1, col2 = st.columns(2)
    with col1:
        csv_bytes = result_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Descargar CSV", data=csv_bytes, file_name="actas_document_ai.csv", mime="text/csv")

    with col2:
        from io import BytesIO
        bio = BytesIO()
        result_df.to_excel(bio, index=False, engine="openpyxl")
        bio.seek(0)
        st.download_button("⬇️ Descargar Excel", data=bio, file_name="actas_document_ai.xlsx")

    st.divider()

    # Subir a Google Sheets
    st.subheader("📤 Subir resultados a Google Sheets")
    st.caption("Asegurate de haber compartido tu Sheet con el correo de la **service account** (Editor).")
    do_upload = st.checkbox("Subir ahora", value=False)
    if do_upload:
        ok = upload_dataframe_to_sheet(result_df)
        if ok:
            st.success("✅ Datos subidos correctamente a Google Sheets.")
        else:
            st.error("❌ No se pudo subir a Google Sheets. Revisá el ID/permiso de la service account.")
