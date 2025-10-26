# app.py
# --------------------------------------------------------
# Extractor de Actas UCCuyo – Document AI + Sheets
# --------------------------------------------------------

import io
import streamlit as st
import pandas as pd

from extract_actas import process_with_document_ai, extract_text_local
from upload_to_sheets import upload_dataframe_to_sheet


st.set_page_config(page_title="Extractor de Actas – UCCuyo", layout="wide")
st.title("📑 Extractor de Actas del Consejo de Investigación – UCCuyo")

st.caption(
    "Subí archivos PDF o DOCX. La app utiliza tu modelo personalizado de "
    "Google Document AI para reconocer automáticamente campos clave "
    "(proyecto, director, equipo, fecha, etc.) y subir los resultados a Google Sheets."
)

uploaded_files = st.file_uploader(
    "📤 Subí uno o más archivos (PDF o DOCX)",
    type=["pdf", "docx", "txt"],
    accept_multiple_files=True,
)

process_btn = st.button("🚀 Procesar")

if process_btn:
    if not uploaded_files:
        st.error("Por favor, subí al menos un archivo.")
        st.stop()

    all_entities = []
    previews = []

    progress = st.progress(0)
    for idx, uf in enumerate(uploaded_files, start=1):
        progress.progress(idx / len(uploaded_files))

        file_bytes = uf.read()
        mime = "application/pdf" if uf.name.lower().endswith(".pdf") else "application/octet-stream"

        try:
            full_text, df_ent = process_with_document_ai(file_bytes, mime_type=mime)
            origin = "Document AI"
        except Exception as e:
            st.warning(f"No se pudo procesar **{uf.name}** con Document AI. "
                       f"Se usará extracción local.\n\n> {e}")
            uf.seek(0)
            full_text = extract_text_local(uf)
            df_ent = pd.DataFrame([{
                "Etiqueta": "TEXTO_COMPLETO",
                "Valor": full_text,
                "Confianza": "",
                "Página": ""
            }])
            origin = "Local"

        # Vista previa de texto
        preview = full_text if len(full_text) <= 3000 else full_text[:3000] + "\n...\n[texto truncado]"
        previews.append({"Archivo": uf.name, "Origen": origin, "Vista previa": preview})

        # Añadir columna de archivo
        if not df_ent.empty:
            df_ent.insert(0, "Archivo", uf.name)
        all_entities.append(df_ent)

    st.subheader("📝 Vista previa del texto")
    st.dataframe(pd.DataFrame(previews), use_container_width=True)

    result_df = pd.concat(all_entities, ignore_index=True) if all_entities else pd.DataFrame()
    st.subheader("🏷️ Etiquetas extraídas")
    st.dataframe(result_df, use_container_width=True)

    # --- Descarga CSV ---
    csv_bytes = result_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Descargar CSV",
        data=csv_bytes,
        file_name="actas_document_ai.csv",
        mime="text/csv"
    )

    # --- Descarga Excel con limpieza de caracteres ilegales ---
    from io import BytesIO

    def clean_excel_text(x):
        if isinstance(x, str):
            return "".join(c for c in x if ord(c) in (9, 10, 13) or 32 <= ord(c) <= 126 or 160 <= ord(c))
        return x

    clean_df = result_df.applymap(clean_excel_text)

    bio = BytesIO()
    clean_df.to_excel(bio, index=False, engine="openpyxl")
    bio.seek(0)
    st.download_button(
        "⬇️ Descargar Excel",
        data=bio,
        file_name="actas_document_ai.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()

    # --- Subir a Google Sheets ---
    st.subheader("📤 Subir resultados a Google Sheets")
    st.caption("Asegurate de que el Google Sheet esté compartido con la cuenta de servicio.")
    do_upload = st.checkbox("Subir a Google Sheets", value=False)
    if do_upload:
        ok = upload_dataframe_to_sheet(clean_df)
        if ok:
            st.success("✅ Datos subidos correctamente a Google Sheets.")
        else:
            st.error("❌ No se pudo subir a Google Sheets. Revisa permisos e IDs.")
