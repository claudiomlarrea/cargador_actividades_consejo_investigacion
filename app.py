# -*- coding: utf-8 -*-
import io, re, unicodedata
import pandas as pd
import streamlit as st

# Lectores de documentos
from pdfminer.high_level import extract_text as pdf_extract_text
from docx import Document as DocxDocument

# Google Sheets (solo si se cargan credenciales)
try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GS = True
except Exception:
    HAS_GS = False

# ============ CONFIGURACIÓN DE LA APP ============ #
st.set_page_config(page_title="Extractor de ACTAS → Excel/Sheets", page_icon="🗂️", layout="centered")
st.title("🗂️ Extractor de ACTAS del Consejo → Excel y Google Sheets")
st.caption("Subí un PDF o DOCX de acta. La app detecta Proyectos, Informes, Categorización, Jornadas, Cursos y trabajos para la Revista Cuadernos.")

# ---------- PATRONES DE SECCIONES ---------- #
SECTION_DEFS = [
    {"name": "Proyectos de investigación", "patterns": [r"Presentación de Proyectos", r"Proyectos de Investigación", r"Proyectos de Convocatoria Abierta"]},
    {"name": "Proyectos de cátedra",       "patterns": [r"Proyectos de Asignatura", r"Proyectos Cuadernos de Cátedra"]},
    {"name": "Informes finales",            "patterns": [r"Presentación de Informes Finales", r"Informes Finales"]},
    {"name": "Informes de avance",          "patterns": [r"Presentación de Informes de Avance", r"Informes de Avance"]},
    {"name": "Categorización",              "patterns": [r"Solicitud de Categorización", r"Categorizaciones Extraordinarias", r"Categorización de investigadores", r"Categorización de docentes"]},
    {"name": "Jornadas de investigación",   "patterns": [r"Jornadas de Investigación", r"Jornadas internas de investigación"]},
    {"name": "Trabajos Revista Cuadernos",  "patterns": [r"Revista Cuadernos", r"Cuadernos de la Secretaría de Investigación", r"presentación de resúmenes", r"trabajos para la revista", r"resúmenes para Cuadernos"]},
    {"name": "Cursos",                      "patterns": [r"Cursos de capacitación", r"Cursos"]},
]

FACULTY_HDR = re.compile(r"^(Facultad|Instituto Superior|Vicerrectorado|Escuela)\b.*", re.IGNORECASE)
ITEM_SPLIT = re.compile(r"\n\s*(?:[\u2022•\-]|[\u25CF\u25A0\u25E6]|\d+\.)\s*")

# ---------- UTILIDADES ---------- #
def _norm(s: str) -> str:
    if s is None: return ""
    s = s.replace("\x00", " ")
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()

def read_pdf_bytes(file_bytes: bytes) -> str:
    with io.BytesIO(file_bytes) as bio:
        return _norm(pdf_extract_text(bio))

def read_docx_bytes(file_bytes: bytes) -> str:
    with io.BytesIO(file_bytes) as bio:
        doc = DocxDocument(bio)
        return _norm("\n".join(p.text for p in doc.paragraphs))

def find_acta_number(text: str) -> str:
    m = re.search(r"ACTA\s+N[º°]?\s*([0-9]+)", text, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""

def find_date_text(text: str) -> str:
    m = re.search(r"En la ciudad.*?\n(.*)", text, flags=re.IGNORECASE)
    if m: return _norm(m.group(1))
    m2 = re.search(r"a los\s+.+?d[ií]as.*?del mes de\s+.+?\s+de\s+dos mil.*", text, flags=re.IGNORECASE)
    return _norm(m2.group(0)) if m2 else ""

def split_sections(text: str):
    named = []
    for i, sec in enumerate(SECTION_DEFS):
        pat = "|".join([p for p in sec["patterns"]])
        named.append(f"(?P<s{i}>\\b(?:{pat})\\b)")
    pattern = re.compile("|".join(named), flags=re.IGNORECASE)
    hits = []
    for m in pattern.finditer(text):
        for i, sec in enumerate(SECTION_DEFS):
            if m.group(f"s{i}"):
                hits.append((sec["name"], m.start()))
                break
    if not hits:
        return [("General", 0, len(text))]
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
    if buf:
        blocks.append((current, "\n".join(buf)))
    return blocks or [(None, section_text)]

def extract_candidate_items(text: str):
    parts = ITEM_SPLIT.split("\n" + text)
    cands = []
    for p in parts:
        p = p.strip(" ;\n\t")
        if len(p) < 6: continue
        if re.search(r"(Proyecto|Denominaci[oó]n|PROJOVI|Informe|Categorizaci[oó]n|Baja del proyecto|Revista|Cuadernos|Cursos?)", p, re.IGNORECASE):
            cands.append(p)
    return cands or [text.strip()]

def infer_estado(text: str) -> str:
    t = text.lower()
    if "baja del proyecto" in t or re.search(r"\bbaja\b", t): return "Baja"
    if "prórroga" in t or "prorroga" in t: return "Prórroga"
    if "aprob" in t and ("elev" in t or "enviado" in t): return "Aprobado y elevado"
    if "aprob" in t: return "Aprobado"
    if "solicitud de categorización" in t or "solicitud de categorizacion" in t: return "Solicitud"
    return ""

def infer_destino_publicacion(text: str) -> str:
    if re.search(r"\b(Cuadernos|revista\s+cuadernos|Cuadernos de la Secretar[ií]a de Investigaci[oó]n)\b", text, re.IGNORECASE):
        return "Revista Cuadernos"
    return ""

def extract_title_director(text: str):
    title, director = None, None
    m = re.search(r"Proyecto\s*:\s*(.+?)(?:\.\s*Director(?:a)?\s*:\s*([^.]+))?$", text, re.IGNORECASE)
    if m:
        title = m.group(1).strip(" .")
        if m.lastindex and m.lastindex >= 2 and m.group(2):
            director = m.group(2).strip(" .")
    if not director:
        m2 = re.search(r"Director(?:a)?\s*:\s*([^.]+)", text, re.IGNORECASE)
        if m2:
            director = m2.group(1).strip(" .")
    if not title:
        m3 = re.search(r"Denominaci[oó]n.*?:\s*(.+?)(?:\.|$)", text, re.IGNORECASE)
        if m3:
            title = m3.group(1).strip()
    if not title:
        m4 = re.search(r"PROJOVI\s*:\s*(.+?)(?:\.|$)", text, re.IGNORECASE)
        if m4:
            title = m4.group(1).strip()
    if not title:
        m5 = re.search(r"[«“\"']([^\"”»']+)[\"”»']", text)
        if m5:
            title = m5.group(1).strip()
    return title, director

def build_dataframe(text: str, source_name: str) -> pd.DataFrame:
    acta = find_acta_number(text)
    fecha = find_date_text(text)
    sections = split_sections(text)
    rows = []
    for sec_name, s, e in sections:
        chunk = text[s:e].strip()
        for faculty, block in chunk_by_faculty(chunk):
            for item in extract_candidate_items(block):
                title, director = extract_title_director(item)
                estado = infer_estado(block + " " + item)
                destino = infer_destino_publicacion(block + " " + item)
                rows.append({
                    "Acta": acta,
                    "Fecha": fecha,
                    "Facultad": faculty or "",
                    "Tipo_tema": sec_name,
                    "Titulo_o_denominacion": (title or item)[:400],
                    "Director": director or "",
                    "Estado": estado,
                    "Destino_publicacion": destino,
                    "Fuente_archivo": source_name,
                })
    df = pd.DataFrame(rows)
    if not df.empty:
        for c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df

# ---------- INTERFAZ STREAMLIT ---------- #
file = st.file_uploader("Subí el acta (PDF o DOCX)", type=["pdf","docx"])
col1, col2 = st.columns(2)
with col1:
    sheet_name = st.text_input("Nombre del Google Sheet", value="Base Consejo de Investigación")
with col2:
    ws_name = st.text_input("Nombre de la pestaña", value="Actas")

if file:
    suffix = file.name.split(".")[-1].lower()
    raw = file.read()
    text = read_pdf_bytes(raw) if suffix == "pdf" else read_docx_bytes(raw)
    if not text.strip():
        st.error("No se pudo leer el archivo.")
        st.stop()

    df = build_dataframe(text, file.name)
    if df.empty:
        st.warning("No se detectaron ítems.")
        st.stop()

    st.success("Extracción completada.")
    st.dataframe(df, use_container_width=True)

    # ---------- DESCARGAR EXCEL (con fallback) ----------
    st.subheader("Descargar Excel")
    def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
        buf = io.BytesIO()
        try:
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Actas")
            return buf.getvalue()
        except Exception:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Actas")
            return buf.getvalue()

    excel_bytes = df_to_excel_bytes(df)
    st.download_button(
        "💾 Descargar Excel (Actas.xlsx)",
        data=excel_bytes,
        file_name="Actas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ---------- SUBIR A GOOGLE SHEETS ----------
    st.subheader("Subir a Google Sheets (opcional)")
    st.caption("Configura en Streamlit Cloud → Settings → Secrets el JSON de Service Account (clave: gcp_service_account).")
    can_upload = HAS_GS and ("gcp_service_account" in st.secrets)
    do_upload = st.checkbox("Subir esta tabla a Google Sheets", value=False, disabled=(not can_upload))

    if do_upload and can_upload:
        try:
            scopes = ["https://www.googleapis.com/auth/spreadsheets"]
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
            client = gspread.authorize(creds)
            try:
                sh = client.open(sheet_name)
            except gspread.exceptions.SpreadsheetNotFound:
                sh = client.create(sheet_name)
            try:
                ws = sh.worksheet(ws_name)
            except gspread.exceptions.WorksheetNotFound:
                ws = sh.add_worksheet(title=ws_name, rows=1000, cols=20)
            ws.clear()
            ws.update([df.columns.tolist()] + df.astype(str).values.tolist())
            st.success(f"✅ Google Sheets actualizado: {sheet_name} / {ws_name} ({len(df)} filas)")
        except Exception as e:
            st.error(f"Error subiendo a Google Sheets: {e}")
