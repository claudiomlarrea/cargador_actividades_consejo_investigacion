# -*- coding: utf-8 -*-
import io, re, unicodedata, datetime as dt
import pandas as pd
import streamlit as st

# Lectores de archivos
from pdfminer.high_level import extract_text as pdf_extract_text
from docx import Document as DocxDocument

# Librerías Google
try:
    import gspread
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
    HAS_GOOGLE = True
except Exception:
    HAS_GOOGLE = False

# ================= CONFIGURACIÓN GENERAL ================= #
st.set_page_config(page_title="Extractor de ACTAS → Google Sheets", page_icon="🗂️", layout="centered")
st.title("🗂️ Extractor de ACTAS del Consejo → Hoja automática en Google Drive")

# Carpeta por defecto (ya compartida)
DEFAULT_FOLDER_ID = "1O7xo7cCGkSujhUXIl3fv51S-cniVLmnh"
FOLDER_ID = st.secrets.get("drive_folder_id", DEFAULT_FOLDER_ID)

# ================= PATRONES ================= #
SECTION_DEFS = [
    {"name": "Proyectos de investigación", "patterns": [r"Presentación de Proyectos", r"Proyectos de Investigación", r"Proyectos de Convocatoria Abierta"]},
    {"name": "Proyectos de cátedra", "patterns": [r"Proyectos de Asignatura", r"Proyectos Cuadernos de Cátedra"]},
    {"name": "Informes finales", "patterns": [r"Presentación de Informes Finales", r"Informes Finales"]},
    {"name": "Informes de avance", "patterns": [r"Presentación de Informes de Avance", r"Informes de Avance"]},
    {"name": "Categorización", "patterns": [r"Categorización", r"Categorizaciones Extraordinarias", r"Categorización de investigadores", r"Categorización de docentes"]},
    {"name": "Jornadas de investigación", "patterns": [r"Jornadas de Investigación", r"Jornadas internas de investigación"]},
    {"name": "Trabajos Revista Cuadernos", "patterns": [r"Revista Cuadernos", r"Cuadernos de la Secretaría de Investigación", r"presentación de resúmenes", r"trabajos para la revista", r"resúmenes para Cuadernos"]},
    {"name": "Cursos", "patterns": [r"Cursos de capacitación", r"Cursos"]},
]
FACULTY_HDR = re.compile(r"^(Facultad|Instituto Superior|Vicerrectorado|Escuela)\b.*", re.IGNORECASE)
ITEM_SPLIT = re.compile(r"\n\s*(?:[\u2022•\-]|\d+\.)\s*")

# ================= FUNCIONES UTILITARIAS ================= #
def _norm(s: str) -> str:
    if not s: return ""
    s = unicodedata.normalize("NFKC", s).replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()

def read_pdf_bytes(file_bytes: bytes) -> str:
    with io.BytesIO(file_bytes) as bio:
        return _norm(pdf_extract_text(bio) or "")

def read_docx_bytes(file_bytes: bytes) -> str:
    with io.BytesIO(file_bytes) as bio:
        doc = DocxDocument(bio)
        return _norm("\n".join(p.text for p in doc.paragraphs))

def find_acta_number(text: str) -> str:
    m = re.search(r"ACTA\s+N[º°]?\s*([0-9]+)", text, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""

def find_date_text(text: str) -> str:
    m = re.search(r"a los\s+.+?d[ií]as.*?del mes de\s+.+?\s+de\s+dos mil.*", text, flags=re.IGNORECASE)
    return _norm(m.group(0)) if m else ""

def split_sections(text: str):
    hits = []
    for sec in SECTION_DEFS:
        for pat in sec["patterns"]:
            for m in re.finditer(pat, text, flags=re.IGNORECASE):
                hits.append((sec["name"], m.start()))
    if not hits: return [("General", 0, len(text))]
    hits.sort(key=lambda x: x[1])
    spans = []
    for i, (name, start) in enumerate(hits):
        end = hits[i+1][1] if i+1 < len(hits) else len(text)
        spans.append((name, start, end))
    return spans

def chunk_by_faculty(section_text: str):
    lines = [ln.strip() for ln in section_text.split("\n") if ln.strip()]
    blocks, current, buf = [], None, []
    for ln in lines:
        if FACULTY_HDR.match(ln):
            if buf:
                blocks.append((current, "\n".join(buf)))
                buf = []
            current = ln
        else:
            buf.append(ln)
    if buf: blocks.append((current, "\n".join(buf)))
    return blocks or [(None, section_text)]

def extract_candidate_items(text: str):
    parts = ITEM_SPLIT.split("\n" + text)
    cands = [p.strip() for p in parts if len(p.strip()) > 10]
    return cands or [text.strip()]

def infer_estado(text: str) -> str:
    t = text.lower()
    if "baja" in t: return "Baja"
    if "prórroga" in t or "prorroga" in t: return "Prórroga"
    if "aprob" in t and "elev" in t: return "Aprobado y elevado"
    if "aprob" in t: return "Aprobado"
    if "solicitud" in t: return "Solicitud"
    return ""

def infer_destino_publicacion(text: str) -> str:
    if re.search(r"\b(Cuadernos|revista\s+cuadernos)\b", text, re.IGNORECASE):
        return "Revista Cuadernos"
    return ""

def extract_title_director(text: str):
    title, director = "", ""
    m = re.search(r"Proyecto\s*:\s*(.+?)(?:\.\s*Director(?:a)?\s*:\s*([^.]+))?$", text, re.IGNORECASE)
    if m:
        title = m.group(1).strip()
        director = (m.group(2) or "").strip()
    return title, director

def build_dataframe(text: str, source_name: str) -> pd.DataFrame:
    acta = find_acta_number(text)
    fecha = find_date_text(text)
    rows = []
    for sec_name, s, e in split_sections(text):
        chunk = text[s:e].strip()
        for faculty, block in chunk_by_faculty(chunk):
            for item in extract_candidate_items(block):
                title, director = extract_title_director(item)
                rows.append({
                    "Acta": acta,
                    "Fecha": fecha,
                    "Facultad": faculty or "",
                    "Tipo_tema": sec_name,
                    "Titulo_o_denominacion": title or item[:200],
                    "Director": director,
                    "Estado": infer_estado(block + item),
                    "Destino_publicacion": infer_destino_publicacion(block + item),
                    "Fuente_archivo": source_name
                })
    return pd.DataFrame(rows)

# ================= GOOGLE DRIVE HELPERS ================= #
def get_creds(scopes):
    if not HAS_GOOGLE or "gcp_service_account" not in st.secrets:
        return None
    return Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)

def delete_existing_by_name_in_folder(drive, name: str, folder_id: str):
    safe_name = name.replace("'", "\\'")
    q = "name = '{}' and '{}' in parents and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false".format(safe_name, folder_id)
    res = drive.files().list(q=q, fields="files(id)").execute()
    for f in res.get("files", []):
        drive.files().delete(fileId=f["id"]).execute()

def create_native_sheet_in_folder_from_df(df: pd.DataFrame, name: str, folder_id: str):
    scopes = ["https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    creds = get_creds(scopes)
    if creds is None:
        raise RuntimeError("Faltan credenciales en st.secrets['gcp_service_account'].")

    drive = build("drive", "v3", credentials=creds)
    delete_existing_by_name_in_folder(drive, name, folder_id)

    csv_bytes = df.to_csv(index=False).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(csv_bytes), mimetype="text/csv", resumable=False)
    metadata = {"name": name, "mimeType": "application/vnd.google-apps.spreadsheet", "parents": [folder_id]}
    file = drive.files().create(body=metadata, media_body=media, fields="id, webViewLink").execute()
    return file.get("webViewLink")

# ================= INTERFAZ STREAMLIT ================= #
file = st.file_uploader("Subí el acta (PDF o DOCX)", type=["pdf", "docx"])

if file:
    raw = file.read()
    text = read_pdf_bytes(raw) if file.name.lower().endswith(".pdf") else read_docx_bytes(raw)
    if not text.strip():
        st.error("No se pudo leer el archivo.")
        st.stop()

    df = build_dataframe(text, file.name)
    if df.empty:
        st.warning("No se detectaron ítems en el documento.")
        st.stop()

    st.success("✅ Extracción completada")
    st.dataframe(df, use_container_width=True)

    # Descargas locales
    def df_to_excel_bytes(df):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False, sheet_name="Actas")
        return buf.getvalue()

    st.download_button("💾 Descargar Excel (Actas.xlsx)", df_to_excel_bytes(df), "Actas.xlsx")
    st.download_button("⬇️ Descargar CSV (Actas.csv)", df.to_csv(index=False).encode("utf-8"), "Actas.csv")

    # Creación automática en Drive
    if "gcp_service_account" not in st.secrets:
        st.info("Cargá tu Service Account en *Settings → Secrets* con la clave [gcp_service_account]. "
                "Y compartí la carpeta destino con permiso de Editor.")
    else:
        try:
            acta_num = df["Acta"].iloc[0] if (df["Acta"] != "").any() else ""
            name = f"Actas Consejo {acta_num or dt.datetime.now().strftime('%Y-%m-%d_%H%M')}"
            link = create_native_sheet_in_folder_from_df(df, name, FOLDER_ID)
            st.success(f"✅ Hoja creada en tu carpeta 'Actas de Consejo': **{name}**")
            st.write("🔗", link)
            st.caption("Podés vincularla directamente con tu tablero de Looker Studio.")
        except Exception as e:
            st.error(f"Ocurrió un error al crear la hoja en Drive: {e}")
else:
    st.caption("Subí un archivo PDF o DOCX para comenzar.")
