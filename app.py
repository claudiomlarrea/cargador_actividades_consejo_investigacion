# -*- coding: utf-8 -*-
import io, re, unicodedata
import pandas as pd
import streamlit as st
from pdfminer.high_level import extract_text as pdf_extract_text
from docx import Document
from google.oauth2.service_account import Credentials

# ──────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────
st.set_page_config(page_title="Extractor de ACTAS → Google Sheets", page_icon="📑", layout="wide")
st.title("📑 Extractor de ACTAS del Consejo → Base con 7 temas + Año")

DEFAULT_FOLDER_ID = "1O7xo7cCGkSujhUXIl3fv51S-cniVLmnh"  # fallback si no hay secret

# Diagnóstico opcional
with st.expander("Diagnóstico de configuración", expanded=False):
    st.write("Secrets:", list(st.secrets.keys()))
    st.write("drive_folder_id:", st.secrets.get("drive_folder_id", DEFAULT_FOLDER_ID))
    st.write("SA presente:", "gcp_service_account" in st.secrets)

# ──────────────────────────────────────────
# Utilidades
# ──────────────────────────────────────────
SPANISH_YEAR_WORDS = {
    "veinte": 2020, "veintiuno": 2021,
    "veintidos": 2022, "veintidós": 2022,
    "veintitres": 2023, "veintitrés": 2023,
    "veinticuatro": 2024, "veinticinco": 2025,
    "veintiseis": 2026, "veintiséis": 2026,
    "veintisiete": 2027, "veintiocho": 2028,
    "veintinueve": 2029, "treinta": 2030
}

def norm(s: str) -> str:
    if not isinstance(s, str): return ""
    s = unicodedata.normalize("NFKC", s).replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()

def extract_text_any(f) -> str:
    if f.name.lower().endswith(".pdf"):
        return norm(pdf_extract_text(f))
    if f.name.lower().endswith(".docx"):
        doc = Document(f)
        return norm("\n".join(p.text for p in doc.paragraphs))
    return ""

def get_acta_number(text: str, fname: str) -> str:
    m = re.search(r"ACTA\s+N[º°]?\s*([0-9]+)", text, re.IGNORECASE)
    if m: return m.group(1)
    m2 = re.search(r"([0-9]{2,4})", fname)
    return m2.group(1) if m2 else ""

def get_fecha(text: str) -> str:
    # línea larga con “dos mil …”
    m = re.search(r".*dos mil [[:alpha:]]+.*", text, re.IGNORECASE)
    if m: return norm(m.group(0))
    # primera línea larga del documento
    return text.split("\n")[0][:300]

def infer_year_from_text(s: str):
    if not isinstance(s, str): return None
    m = re.search(r"\b(20\d{2})\b", s)
    if m: 
        try: return int(m.group(1))
        except: pass
    t = unicodedata.normalize("NFKD", s).lower()
    t = t.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
    m2 = re.search(r"dos mil\s+([a-z]+)", t)
    if m2: return SPANISH_YEAR_WORDS.get(m2.group(1))
    return None

# ──────────────────────────────────────────
# Clasificación en los 7 temas
# ──────────────────────────────────────────
TOPIC_MAP = {
    "Proyectos de investigación": [
        r"\bproyectos? de (investigaci[oó]n|convocatoria abierta)\b",
        r"\bpresentaci[oó]n de proyectos?\b",
        r"\bprojovi\b", r"\bpid\b", r"\bppi\b"
    ],
    "Proyectos de investigación de cátedra": [
        r"\bproyectos? (de )?(asignatura|c[aá]tedra)\b",
        r"\bproyectos? cuadernos de c[aá]tedra\b"
    ],
    "Informes de avances": [
        r"\binformes? de avance\b", r"\bpresentaci[oó]n de informes? de avance\b"
    ],
    "Informes finales": [
        r"\binformes? finales?\b", r"\bpresentaci[oó]n de informes? finales?\b"
    ],
    "Categorización de investigadores o categorización de docentes": [
        r"\bcategorizaci[oó]n\b", r"\bsolicitud de categorizaci[oó]n\b",
        r"\bcategorizaciones? extraordinarias?\b"
    ],
    "Jornadas de investigación": [
        r"\bjornadas? de investigaci[oó]n\b", r"\bjornadas? internas\b"
    ],
    "Cursos de capacitación": [
        r"\bcursos? de capacitaci[oó]n\b", r"\bcursos?\b", r"\btaller(es)?\b", r"\bcapacitaci[oó]n\b"
    ],
}

TOPICS = list(TOPIC_MAP.keys())

def classify_topic(text: str) -> str:
    t = text.lower()
    for topic, pats in TOPIC_MAP.items():
        for pat in pats:
            if re.search(pat, t, re.IGNORECASE):
                return topic
    return "Proyectos de investigación"  # fallback conservador

# ──────────────────────────────────────────
# Heurísticas de extracción (título/director/estado/facultad)
# ──────────────────────────────────────────
FACULTY_PAT = re.compile(r"^(Facultad|Escuela|Instituto Superior|Vicerrectorado)\b.*", re.IGNORECASE)
STATE_WORDS = [
    ("Aprobado y elevado", r"\baprobado(?:s)?\b.*\belevad"),
    ("Aprobado", r"\baprobado(?:s)?\b"),
    ("Baja", r"\bbaja\b"),
    ("Prórroga", r"\bpr[oó]rrog"),
    ("Observaciones", r"\bobservaci[oó]n|observaciones"),
    ("Solicitud", r"\bsolicitud\b"),
]

def find_state(text: str) -> str:
    t = text.lower()
    for label, pat in STATE_WORDS:
        if re.search(pat, t): return label
    return ""

def find_director(text: str) -> str:
    m = re.search(r"Director(?:a)?\s*:\s*([^\.;\n]+)", text, re.IGNORECASE)
    return norm(m.group(1)) if m else ""

def find_title(text: str) -> str:
    for regex in [
        r"Denominaci[oó]n\s*:\s*(.+)",
        r"T[ií]tulo\s*:\s*(.+)",
        r"Proyecto\s*:\s*(.+)"
    ]:
        m = re.search(regex, text, re.IGNORECASE)
        if m:
            return norm(re.sub(r"\s*Director.*$", "", m.group(1)).strip(" .;"))
    # entre comillas
    m2 = re.search(r"[«“\"']([^\"”»']{6,})[\"”»']", text)
    if m2: return norm(m2.group(1))
    return ""

def block_by_faculty(text: str):
    """Divide por encabezados 'Facultad …' si los hay."""
    lines = [ln for ln in text.split("\n") if ln.strip()]
    blocks, current_fac, buf = [], "", []
    for ln in lines:
        if FACULTY_PAT.match(ln):
            if buf:
                blocks.append((current_fac, "\n".join(buf)))
                buf = []
            current_fac = ln.strip()
        else:
            buf.append(ln)
    if buf: blocks.append((current_fac, "\n".join(buf)))
    return blocks if blocks else [("", text)]

def split_items(txt: str):
    # bullets / numeraciones simples
    parts = re.split(r"\n\s*(?:[\u2022•\-•\*]|\d+\))\s*", "\n"+txt)
    parts = [p.strip() for p in parts if len(p.strip()) > 8]
    return parts if parts else [txt]

# ──────────────────────────────────────────
# Parse principal → DataFrame
# ──────────────────────────────────────────
def parse_acta_to_rows(text: str, fname: str):
    rows = []
    acta = get_acta_number(text, fname)
    fecha = get_fecha(text)
    for faculty, chunk in block_by_faculty(text):
        for item in split_items(chunk):
            if len(item) < 10: 
                continue
            topic = classify_topic(item + " " + chunk)
            title = find_title(item) or find_title(chunk) or ""
            director = find_director(item) or find_director(chunk) or ""
            state = find_state(item + " " + chunk)
            rows.append({
                "Año": infer_year_from_text(fecha),
                "Acta": acta,
                "Fecha": fecha,
                "Facultad": faculty,
                "Tipo_tema": topic,
                "Titulo_o_denominacion": title if title else item[:300],
                "Director": director,
                "Estado": state,
                "Fuente_archivo": fname
            })
    return rows

# ──────────────────────────────────────────
# Google Drive (opcional)
# ──────────────────────────────────────────
def get_creds(scopes):
    sa = st.secrets.get("gcp_service_account")
    if not sa: return None
    if isinstance(sa, dict):
        return Credentials.from_service_account_info(sa, scopes=scopes)
    try:
        import json
        return Credentials.from_service_account_info(json.loads(sa), scopes=scopes)
    except Exception:
        return None

def create_sheet_in_drive(df: pd.DataFrame, name: str, folder_id: str, creds):
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
    except ModuleNotFoundError:
        st.error("Falta `google-api-python-client` en requirements.txt.")
        return None
    drive = build("drive", "v3", credentials=creds)
    csv_bytes = df.to_csv(index=False).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(csv_bytes), mimetype="text/csv", resumable=False)
    metadata = {"name": name, "mimeType": "application/vnd.google-apps.spreadsheet", "parents": [folder_id]}
    f = drive.files().create(body=metadata, media_body=media, fields="id, webViewLink").execute()
    return f.get("webViewLink")

# ──────────────────────────────────────────
# UI
# ──────────────────────────────────────────
files = st.file_uploader("📂 Subí actas PDF o DOCX (una o varias)", type=["pdf", "docx"], accept_multiple_files=True)

if not files:
    st.info("Subí archivos para comenzar.")
    st.stop()

all_rows = []
for f in files:
    txt = extract_text_any(f)
    if not txt:
        st.warning(f"No pude leer: {f.name}")
        continue
    all_rows.extend(parse_acta_to_rows(txt, f.name))

if not all_rows:
    st.error("No se detectaron ítems en las actas.")
    st.stop()

df = pd.DataFrame(all_rows)

# Orden de columnas (Año primero)
ordered = ["Año","Acta","Fecha","Facultad","Tipo_tema","Titulo_o_denominacion","Director","Estado","Fuente_archivo"]
df = df[ordered]

st.success("✅ Actas procesadas.")
st.dataframe(df, use_container_width=True)

# ── Descargas
st.subheader("Descargar")
# Excel
# --- helper robusto para crear el Excel en memoria ---
def df_to_excel_bytes(df: pd.DataFrame) -> io.BytesIO:
    buf = io.BytesIO()
    try:
        # 1️⃣ Intentar con openpyxl (el más común en Streamlit Cloud)
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="Actas")
        buf.seek(0)
        return buf
    except Exception:
        # 2️⃣ Si openpyxl no está, probar con xlsxwriter
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False, sheet_name="Actas")
        buf.seek(0)
        return buf

# --- uso en el botón de descarga ---
st.subheader("Descargar")
buf_xlsx = df_to_excel_bytes(df)
st.download_button(
    "📘 Descargar Excel (Actas.xlsx)",
    data=buf_xlsx,
    file_name="Actas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# CSV (dejalo como está)
st.download_button("📗 CSV (Actas.csv)", data=df.to_csv(index=False).encode("utf-8"),
                   file_name="Actas.csv", mime="text/csv")

# CSV
st.download_button("📗 CSV (Actas.csv)", data=df.to_csv(index=False).encode("utf-8"),
                   file_name="Actas.csv", mime="text/csv")

# ── Enviar a Drive (opcional)
st.subheader("Crear Hoja de Google nativa en Drive (opcional)")
folder_id = st.secrets.get("drive_folder_id", DEFAULT_FOLDER_ID)
sheet_name = st.text_input("Nombre de la hoja en Drive", value="Actas Consejo")
creds = get_creds(["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/drive.file"])

if not creds:
    st.caption("Cargá tu Service Account en *Settings → Secrets*. Compartí la carpeta destino con permiso de Editor.")
else:
    if st.button("🚀 Crear hoja en Drive"):
        link = create_sheet_in_drive(df, sheet_name, folder_id, creds)
        if link:
            st.success("Hoja creada correctamente.")
            st.write("Abrir:", link)
