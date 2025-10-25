# -*- coding: utf-8 -*-
import io, re, unicodedata
import pandas as pd
import streamlit as st
from pdfminer.high_level import extract_text as pdf_extract_text
from docx import Document
from google.oauth2.service_account import Credentials

# ──────────────────────────────────────────
# CONFIGURACIÓN GENERAL
# ──────────────────────────────────────────
st.set_page_config(page_title="Extractor de ACTAS del Consejo", page_icon="📑", layout="wide")
st.title("📑 Extractor de ACTAS del Consejo → Base institucional")

DEFAULT_FOLDER_ID = "1O7xo7cCGkSujhUXIl3fv51S-cniVLmnh"  # Carpeta predeterminada

# ──────────────────────────────────────────
# UTILIDADES
# ──────────────────────────────────────────
SPANISH_YEAR_WORDS = {
    "veinte": 2020, "veintiuno": 2021, "veintidos": 2022, "veintidós": 2022,
    "veintitres": 2023, "veintitrés": 2023, "veinticuatro": 2024, "veinticinco": 2025,
    "veintiseis": 2026, "veintiséis": 2026, "veintisiete": 2027, "veintiocho": 2028
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
    m = re.search(r".*dos mil [[:alpha:]]+.*", text, re.IGNORECASE)
    if m: return norm(m.group(0))
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
# CLASIFICACIÓN POR TEMA
# ──────────────────────────────────────────
TOPIC_MAP = {
    "Proyectos de investigación": [
        r"\bproyectos? de (investigaci[oó]n|convocatoria abierta)\b",
        r"\bpresentaci[oó]n de proyectos?\b", r"\bprojovi\b", r"\bpid\b"
    ],
    "Proyectos de investigación de cátedra": [
        r"\bproyectos? (de )?(asignatura|c[aá]tedra)\b", r"\bproyectos? cuadernos\b"
    ],
    "Informes de avances": [
        r"\binformes? de avance\b", r"\bpresentaci[oó]n de informes? de avance\b"
    ],
    "Informes finales": [
        r"\binformes? finales?\b", r"\bpresentaci[oó]n de informes? finales?\b"
    ],
    "Categorización de investigadores o categorización de docentes": [
        r"\bcategorizaci[oó]n\b", r"\bsolicitud de categorizaci[oó]n\b"
    ],
    "Jornadas de investigación": [
        r"\bjornadas? de investigaci[oó]n\b", r"\bjornadas? internas\b"
    ],
    "Cursos de capacitación": [
        r"\bcursos? de capacitaci[oó]n\b", r"\btaller(es)?\b", r"\bcapacitaci[oó]n\b"
    ]
}

def classify_topic(text: str) -> str:
    t = text.lower()
    for topic, pats in TOPIC_MAP.items():
        for pat in pats:
            if re.search(pat, t, re.IGNORECASE):
                return topic
    return "Proyectos de investigación"

# ──────────────────────────────────────────
# PARSEO DE CAMPOS
# ──────────────────────────────────────────
def find_state(text: str) -> str:
    for label, pat in [
        ("Aprobado y elevado", r"\baprobado(?:s)?\b.*\belevad"),
        ("Aprobado", r"\baprobado(?:s)?\b"),
        ("Prórroga", r"\bpr[oó]rrog"),
        ("Baja", r"\bbaja\b"),
        ("Observaciones", r"\bobservaci[oó]n"),
        ("Solicitud", r"\bsolicitud\b")
    ]:
        if re.search(pat, text.lower()): return label
    return ""

def find_director(text: str) -> str:
    m = re.search(r"Director(?:a)?\s*:\s*([^\.;\n]+)", text, re.IGNORECASE)
    return norm(m.group(1)) if m else ""

def find_title(text: str) -> str:
    for regex in [
        r"Denominaci[oó]n\s*:\s*(.+)", r"T[ií]tulo\s*:\s*(.+)", r"Proyecto\s*:\s*(.+)"
    ]:
        m = re.search(regex, text, re.IGNORECASE)
        if m: return norm(re.sub(r"\s*Director.*$", "", m.group(1)).strip(" .;"))
    return ""

def block_by_faculty(text: str):
    lines = [ln for ln in text.split("\n") if ln.strip()]
    blocks, current_fac, buf = [], "", []
    for ln in lines:
        if re.match(r"^(Facultad|Escuela|Instituto|Vicerrectorado)", ln, re.IGNORECASE):
            if buf: blocks.append((current_fac, "\n".join(buf))); buf = []
            current_fac = ln.strip()
        else:
            buf.append(ln)
    if buf: blocks.append((current_fac, "\n".join(buf)))
    return blocks if blocks else [("", text)]

def split_items(txt: str):
    parts = re.split(r"\n\s*(?:[\u2022•\-•\*]|\d+\))\s*", "\n"+txt)
    return [p.strip() for p in parts if len(p.strip()) > 8]

def parse_acta_to_rows(text: str, fname: str):
    rows = []
    acta = get_acta_number(text, fname)
    fecha = get_fecha(text)
    for fac, chunk in block_by_faculty(text):
        for item in split_items(chunk):
            topic = classify_topic(item)
            title = find_title(item)
            director = find_director(item)
            estado = find_state(item)
            rows.append({
                "Año": infer_year_from_text(fecha),
                "Acta": acta,
                "Fecha": fecha,
                "Facultad": fac,
                "Tipo_tema": topic,
                "Titulo_o_denominacion": title or item[:250],
                "Director": director,
                "Estado": estado,
                "Fuente_archivo": fname
            })
    return rows

# ──────────────────────────────────────────
# EXPORTAR EXCEL (ROBUSTO)
# ──────────────────────────────────────────
def df_to_excel_bytes(df: pd.DataFrame) -> io.BytesIO:
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="Actas")
        buf.seek(0)
        return buf
    except Exception:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False, sheet_name="Actas")
        buf.seek(0)
        return buf

# ──────────────────────────────────────────
# GOOGLE DRIVE (opcional)
# ──────────────────────────────────────────
def get_creds(scopes):
    sa = st.secrets.get("gcp_service_account")
    if not sa: return None
    if isinstance(sa, dict):
        return Credentials.from_service_account_info(sa, scopes=scopes)
    import json
    return Credentials.from_service_account_info(json.loads(sa), scopes=scopes)

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
# INTERFAZ PRINCIPAL
# ──────────────────────────────────────────
files = st.file_uploader("📂 Subí actas (.pdf o .docx)", type=["pdf", "docx"], accept_multiple_files=True)
if not files:
    st.info("Subí archivos para comenzar.")
    st.stop()

all_rows = []
for f in files:
    text = extract_text_any(f)
    if not text:
        st.warning(f"No se pudo leer {f.name}")
        continue
    all_rows.extend(parse_acta_to_rows(text, f.name))

if not all_rows:
    st.error("No se detectaron ítems válidos en las actas.")
    st.stop()

df = pd.DataFrame(all_rows)
ordered = ["Año","Acta","Fecha","Facultad","Tipo_tema","Titulo_o_denominacion","Director","Estado","Fuente_archivo"]
df = df[ordered]

st.success("✅ Actas procesadas correctamente.")
st.dataframe(df, use_container_width=True)

# ──────────────────────────────────────────
# DESCARGAS
# ──────────────────────────────────────────
st.subheader("Descargar resultados")
buf_xlsx = df_to_excel_bytes(df)
st.download_button("📘 Descargar Excel (Actas.xlsx)", data=buf_xlsx,
                   file_name="Actas.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.download_button("📗 Descargar CSV (Actas.csv)",
                   data=df.to_csv(index=False).encode("utf-8"),
                   file_name="Actas.csv", mime="text/csv")

# ──────────────────────────────────────────
# SUBIR A DRIVE
# ──────────────────────────────────────────
st.subheader("Crear Hoja nativa en Google Drive (opcional)")
folder_id = st.secrets.get("drive_folder_id", DEFAULT_FOLDER_ID)
creds = get_creds(["https://www.googleapis.com/auth/drive.file"])
if creds and st.button("🚀 Crear hoja en Drive"):
    link = create_sheet_in_drive(df, "Actas Consejo", folder_id, creds)
    if link:
        st.success(f"✅ Hoja creada correctamente: [Abrir en Drive]({link})")
else:
    st.caption("Agregá tus credenciales en Settings → Secrets para activar esta función.")
