# -*- coding: utf-8 -*-
import io, re, unicodedata, json
from typing import List, Dict, Any, Tuple
import pandas as pd
import streamlit as st

# Lectura de PDF/DOCX
from pdfminer.high_level import extract_text as pdf_extract_text
from docx import Document

# Google Drive
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

# ──────────────────────────────────────────
# CONFIG STREAMLIT
# ──────────────────────────────────────────
st.set_page_config(page_title="Extractor de Órdenes del Día (Acumulativo)", page_icon="🗂️", layout="wide")
st.title("🗂️ Extractor de Órdenes del Día → Planilla estándar + Drive (Looker-ready, acumulativa)")

DEFAULT_FOLDER_ID = "REEMPLAZAR_CON_TU_CARPETA"  # fallback si no está en secrets
CSV_NAME  = "OrdenDelDia_Consejo.csv"
XLSX_NAME = "OrdenDelDia_Consejo.xlsx"
SHEET_NAME = "OrdenDelDia"

# ──────────────────────────────────────────
# UTILIDADES GENERALES
# ──────────────────────────────────────────
def norm(s: str) -> str:
    if not isinstance(s, str): return ""
    s = unicodedata.normalize("NFKC", s).replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()

def extract_text_any(uploaded) -> str:
    name = uploaded.name.lower()
    if name.endswith(".pdf"):
        return norm(pdf_extract_text(uploaded))
    if name.endswith(".docx"):
        doc = Document(uploaded)
        return norm("\n".join(p.text for p in doc.paragraphs))
    return ""

def find_date_header(text: str) -> Tuple[str, str]:
    head = text[:1500]
    m = re.search(r"\b(\d{1,2})[\/\-\._](\d{1,2})[\/\-\._](\d{2,4})\b", head)
    if m:
        d, mth, y = m.groups()
        y = "20"+y if len(y) == 2 else y
        y = int(y)
        d = f"{int(d):02d}"; mth = f"{int(mth):02d}"
        return str(y), f"{d}/{mth}/{y}"
    m2 = re.search(r"a\s+los\s+\d+\s+d[ií]as.*?mes\s+de\s+[a-záéíóú]+.*?dos\s+mil\s+([a-záéíóú]+)", head, re.I|re.S)
    if m2:
        mapa = {"veinte":2020,"veintiuno":2021,"veintidos":2022,"veintitres":2023,"veinticuatro":2024,"veinticinco":2025}
        y = unicodedata.normalize("NFKD", m2.group(1)).lower()
        y = y.replace("́","").replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
        if y in mapa:
            for ln in head.split("\n"):
                if len(ln.strip()) > 12:
                    return str(mapa[y]), ln.strip()
    m3 = re.search(r"\b(20\d{2})\b", head)
    if m3:
        return m3.group(1), ""
    return "", ""

# ──────────────────────────────────────────
# ESQUEMA FIJO DE COLUMNAS (Looker)
# ──────────────────────────────────────────
FIXED_COLUMNS = [
    "año","fecha",
    "proyectos de investigación","Nombre del proyecto de investigación","Director del Proyecto","Integrantes del equipo de investigación","Unidad académica de procedencia del proyecto",
    "Informe de avance","Nombre del proyecto de investigación del Informe de avance","Director del Proyecto del Informe de avance","Integrantes del equipo de investigación del Informe de avance","Unidad académica de procedencia del proyecto del Informe de avance",
    "Informe Final","Nombre del proyecto de investigación del Informe Final","Director del Proyecto del Informe Final","Integrantes del equipo de investigación del Informe Final","Unidad académica de procedencia del proyecto del Informe Final",
    "Proyectos de investigación de cátedra","Nombre del proyecto de investigación cátedra","Director del Proyecto del Informe de cátedra","Integrantes del equipo de investigación del proyecto de cátedra","Unidad académica de procedencia del proyecto de cátedra",
    "Publicación","Tipo de publicación (revista científica, libro, presentación a congreso, póster, revista Cuadernos, manual)","Docente o investigador incluida en la publicación","Unidad académica (Publicación)",
    "Categorización de docentes","Nombre del docente categorizado como investigador","Categoría alcanzada por el docente como docente investigador","Unidad académica (Categorización)",
    "Becario de beca cofinanciada doctoral","Nombre del becario doctoral","Becario de beca cofinanciada postdoctoral","Nombre del becario postdoctoral",
    "OTROS TEMAS"
]
def empty_row(base=None):
    row = {c:"" for c in FIXED_COLUMNS}
    if base: row.update({k:v for k,v in base.items() if k in row})
    return row

# ──────────────────────────────────────────
# PARSER ROBUSTO (exactitud de títulos)
# ──────────────────────────────────────────
SECTION_HEADERS = {
    "proyectos": re.compile(r"^(presentaci[oó]n\s+de\s+proyectos?|proyectos?\s+(de\s+)?investigaci[oó]n)\b", re.I),
    "final":     re.compile(r"^informes?\s+final(?:es)?\b", re.I),
    "avance":    re.compile(r"^informes?\s+de\s+avance\b", re.I),
    "catedra":   re.compile(r"(proyectos?\s+(de\s+)?c[aá]tedra|proyectos?\s+cuadernos)", re.I),
    "publica":   re.compile(r"^publicaci[oó]n(?:es)?\b", re.I),
    "categ":     re.compile(r"^categorizaci[oó]n", re.I),
    "beca":      re.compile(r"^becari[oa]s?\b", re.I),
}
DIRECTOR_LABELS = re.compile(r"\b(Director(?:a)?|Co[- ]?director(?:a)?)\b\s*:", re.I)
TEAM_LABELS     = re.compile(r"\b(Equipo(?:\s+de\s+(Trabajo|Investigaci[oó]n))?|Integrantes|Investigadores|Docentes|Estudiantes)\b\s*:", re.I)
UNIT_LABELS     = re.compile(r"\b(Facultad|Escuela|Instituto|Vicerrectorado)\b", re.I)
NO_TITLE_PREFIX = re.compile(r"^(L[ií]neas? de investigaci[oó]n|Enfoque[s]?:|Programa|Jornadas|PEI|Plan|Comisi[oó]n|Convocatoria)\b", re.I)

def split_lines(text: str) -> List[str]:
    text = re.sub(r"[\u2022\u2023\u25CF\u25CB\u25A0•►▪▫]+", "\n", text)
    lines = [norm(ln.strip(" \t-—–•")) for ln in text.split("\n")]
    return [ln for ln in lines if ln]

def is_section_header(ln: str) -> str:
    for key, rx in SECTION_HEADERS.items():
        if rx.search(ln): return key
    return ""

def is_faculty_line(ln: str) -> bool:
    return bool(re.match(r"^(Facultad|Escuela|Instituto|Vicerrectorado)\b", ln, re.I))

def looks_title_line(ln: str) -> bool:
    if not ln or len(ln) < 6: return False
    if NO_TITLE_PREFIX.search(ln): return False
    if DIRECTOR_LABELS.search(ln) or TEAM_LABELS.search(ln) or UNIT_LABELS.search(ln): return False
    if re.search(r"[«“\"'].*[»”\"']", ln): return True
    alpha = sum(ch.isalpha() for ch in ln)
    caps  = sum(ch.isupper() for ch in ln if ch.isalpha())
    return alpha >= 6 and (ln.istitle() or (alpha > 0 and caps/alpha >= 0.40))

def extract_after(label_rx: re.Pattern, chunk: List[str]) -> str:
    for ln in chunk:
        m = re.search(label_rx, ln)
        if m:
            return norm(re.sub(label_rx, "", ln)).strip(" .,:;–-\"'«»“”")
    return ""

def extract_unit(chunk: List[str]) -> str:
    for ln in chunk:
        if is_faculty_line(ln): return ln
    for ln in chunk:
        if UNIT_LABELS.search(ln): return ln
    return ""

def cut_before_director(chunk: List[str]) -> List[str]:
    out = []
    for ln in chunk:
        if DIRECTOR_LABELS.search(ln): break
        out.append(ln)
    return out

def first_title_from(chunk: List[str]) -> str:
    pre = cut_before_director(chunk)
    for ln in pre:
        if looks_title_line(ln):
            return ln.strip(" .,:;–-\"'«»“”")
    for ln in pre:
        if ln and not NO_TITLE_PREFIX.search(ln):
            return ln.strip(" .,:;–-\"'«»“”")
    return ""

def parse_items_by_section(lines: List[str], base_meta: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    section = ""
    buf: List[str] = []
    current_unit = ""

    def flush():
        nonlocal buf, section, current_unit
        if not buf: return
        chunk = [ln for ln in buf if ln]
        joined = " ".join(chunk)
        if NO_TITLE_PREFIX.match(chunk[0]) and not any(DIRECTOR_LABELS.search(x) for x in chunk):
            buf.clear(); return

        unit_here = extract_unit(chunk) or current_unit

        def make_row():
            r = empty_row(base_meta)
            r["Unidad académica de procedencia del proyecto"] = unit_here
            return r

        if section == "proyectos":
            r = make_row()
            r["proyectos de investigación"] = "Sí"
            r["Nombre del proyecto de investigación"] = first_title_from(chunk)
            r["Director del Proyecto"] = extract_after(DIRECTOR_LABELS, chunk)
            r["Integrantes del equipo de investigación"] = extract_after(TEAM_LABELS, chunk)
            r["Unidad académica de procedencia del proyecto"] = unit_here
            if r["Nombre del proyecto de investigación"]: rows.append(r)

        elif section == "final":
            r = make_row()
            r["Informe Final"] = "Sí"
            r["Nombre del proyecto de investigación del Informe Final"] = first_title_from(chunk)
            r["Director del Proyecto del Informe Final"] = extract_after(DIRECTOR_LABELS, chunk)
            r["Integrantes del equipo de investigación del Informe Final"] = extract_after(TEAM_LABELS, chunk)
            r["Unidad académica de procedencia del proyecto del Informe Final"] = unit_here
            if r["Nombre del proyecto de investigación del Informe Final"]: rows.append(r)

        elif section == "avance":
            r = make_row()
            r["Informe de avance"] = "Sí"
            r["Nombre del proyecto de investigación del Informe de avance"] = first_title_from(chunk)
            r["Director del Proyecto del Informe de avance"] = extract_after(DIRECTOR_LABELS, chunk)
            r["Integrantes del equipo de investigación del Informe de avance"] = extract_after(TEAM_LABELS, chunk)
            r["Unidad académica de procedencia del proyecto del Informe de avance"] = unit_here
            if r["Nombre del proyecto de investigación del Informe de avance"]: rows.append(r)

        elif section == "catedra":
            r = make_row()
            r["Proyectos de investigación de cátedra"] = "Sí"
            r["Nombre del proyecto de investigación cátedra"] = first_title_from(chunk)
            r["Director del Proyecto del Informe de cátedra"] = extract_after(DIRECTOR_LABELS, chunk)
            r["Integrantes del equipo de investigación del proyecto de cátedra"] = extract_after(TEAM_LABELS, chunk)
            r["Unidad académica de procedencia del proyecto de cátedra"] = unit_here
            if r["Nombre del proyecto de investigación cátedra"]: rows.append(r)

        elif section == "publica":
            r = make_row()
            r["Publicación"] = "Sí"
            tipo = ""
            if re.search(r"\brevista\b", joined, re.I): tipo = "revista científica"
            elif re.search(r"\blibro\b", joined, re.I): tipo = "libro"
            elif re.search(r"\b(congreso|ponencia|presentaci[oó]n)\b", joined, re.I): tipo = "presentación a congreso"
            elif re.search(r"p[oó]ster|poster", joined, re.I): tipo = "póster"
            elif re.search(r"\bcuadernos\b", joined, re.I): tipo = "revista Cuadernos"
            elif re.search(r"\bmanual\b", joined, re.I): tipo = "manual"
            r["Tipo de publicación (revista científica, libro, presentación a congreso, póster, revista Cuadernos, manual)"] = tipo
            r["Docente o investigador incluida en la publicación"] = extract_after(re.compile(r"(Autor(?:es)?|Docente|Investigador(?:es)?)\s*:", re.I), chunk)
            r["Unidad académica (Publicación)"] = unit_here
            rows.append(r)

        elif section == "categ":
            r = make_row()
            r["Categorización de docentes"] = "Sí"
            mcat = re.search(r"(Categor[ií]a\s*[:\-]?\s*[IVX]+|Investigador(?:\s+\w+){0,3})", joined, re.I)
            r["Categoría alcanzada por el docente como docente investigador"] = mcat.group(0) if mcat else ""
            cand = next((ln for ln in chunk if re.search(r"^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+", ln)), "")
            r["Nombre del docente categorizado como investigador"] = cand
            r["Unidad académica (Categorización)"] = unit_here
            rows.append(r)

        elif section == "beca":
            r = make_row()
            if re.search(r"postdoctoral", joined, re.I):
                r["Becario de beca cofinanciada postdoctoral"] = "Sí"
                r["Nombre del becario postdoctoral"] = first_title_from(chunk) or extract_after(re.compile(r"(Becari[oa]|Nombre)\s*:", re.I), chunk)
            else:
                r["Becario de beca cofinanciada doctoral"] = "Sí"
                r["Nombre del becario doctoral"] = first_title_from(chunk) or extract_after(re.compile(r"(Becari[oa]|Nombre)\s*:", re.I), chunk)
            rows.append(r)

        else:
            r = make_row()
            r["OTROS TEMAS"] = " ".join(chunk)
            rows.append(r)

        buf.clear()
        current_unit = unit_here

    for ln in lines:
        sec = is_section_header(ln)
        if sec:
            flush(); section = sec; continue
        if is_faculty_line(ln):
            flush(); current_unit = ln; continue
        buf.append(ln)
    flush()
    return rows

# ──────────────────────────────────────────
# GOOGLE DRIVE HELPERS
# ──────────────────────────────────────────
def get_creds(scopes):
    sa = st.secrets.get("gcp_service_account")
    if not sa: return None
    if isinstance(sa, dict):
        return Credentials.from_service_account_info(sa, scopes=scopes)
    try:
        return Credentials.from_service_account_info(json.loads(sa), scopes=scopes)
    except Exception:
        return None

def drive_client(creds):
    return build("drive", "v3", credentials=creds)

def drive_find_file(drive, name: str, folder_id: str) -> str:
    q = f"name = '{name}' and '{folder_id}' in parents and trashed = false"
    res = drive.files().list(q=q, fields="files(id,name)", pageSize=1).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else ""

def drive_upload_replace(drive, folder_id: str, name: str, data: bytes, mime: str, file_id_hint: str = ""):
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
    if file_id_hint:
        drive.files().update(fileId=file_id_hint, media_body=media).execute()
        return file_id_hint
    fid = drive_find_file(drive, name, folder_id)
    if fid:
        drive.files().update(fileId=fid, media_body=media).execute()
        return fid
    meta = {"name": name, "parents": [folder_id]}
    f = drive.files().create(body=meta, media_body=media, fields="id").execute()
    return f["id"]

def drive_read_csv_by_id(drive, file_id: str) -> pd.DataFrame:
    try:
        data = drive.files().get_media(fileId=file_id).execute()
        return pd.read_csv(io.BytesIO(data))
    except Exception:
        return pd.DataFrame(columns=FIXED_COLUMNS)

# ──────────────────────────────────────────
# UI
# ──────────────────────────────────────────
st.subheader("1) Subí el/los Órdenes del Día (PDF o DOCX)")
uploads = st.file_uploader("📂 Archivos", type=["pdf","docx"], accept_multiple_files=True)
if not uploads:
    st.info("Subí al menos un archivo para continuar.")
    st.stop()

# Parsear todo lo nuevo
all_rows = []
for up in uploads:
    raw = extract_text_any(up)
    if not raw:
        st.warning(f"No se pudo leer: {up.name}")
        continue
    year, date_str = find_date_header(raw)
    base = {"año": year, "fecha": date_str}
    lines = split_lines(raw)
    rows = parse_items_by_section(lines, base)
    if not rows:
        r = empty_row(base); r["OTROS TEMAS"] = raw[:1500] + ("…" if len(raw) > 1500 else ""); rows = [r]
    all_rows.extend(rows)

if not all_rows:
    st.error("No se detectaron ítems en los Órdenes del Día cargados.")
    st.stop()

# DataFrame nuevo (del/los archivos cargados)
df_new = pd.DataFrame(all_rows)
for col in FIXED_COLUMNS:
    if col not in df_new.columns:
        df_new[col] = ""
df_new = df_new[FIXED_COLUMNS]

# ──────────────────────────────────────────
# ACUMULAR con lo que ya existe en Drive (si existe)
# ──────────────────────────────────────────
st.subheader("2) Consolidación (acumular con el histórico)")

folder_id = st.secrets.get("drive_folder_id", DEFAULT_FOLDER_ID)
creds_full = get_creds(["https://www.googleapis.com/auth/drive"])
csv_id_secret  = st.secrets.get("drive_csv_file_id", "")
xlsx_id_secret = st.secrets.get("drive_xlsx_file_id", "")

existing_df = pd.DataFrame(columns=FIXED_COLUMNS)
if creds_full:
    drv = drive_client(creds_full)
    # Preferencia: leer por ID si lo diste en los secrets; si no, buscar por nombre.
    try:
        if csv_id_secret:
            existing_df = drive_read_csv_by_id(drv, csv_id_secret)
        else:
            maybe_id = drive_find_file(drv, CSV_NAME, folder_id)
            if maybe_id:
                existing_df = drive_read_csv_by_id(drv, maybe_id)
    except Exception:
        existing_df = pd.DataFrame(columns=FIXED_COLUMNS)

# Combinar y deduplicar
df = pd.concat([existing_df, df_new], ignore_index=True)

# Reglas de deduplicación: por Año + cualquiera de los campos de nombre de proyecto (proy/avances/finales)
df["__key__"] = df["año"].astype(str) + "||" + df["Nombre del proyecto de investigación"].fillna("") + "||" + \
                df["Nombre del proyecto de investigación del Informe de avance"].fillna("") + "||" + \
                df["Nombre del proyecto de investigación del Informe Final"].fillna("")
df.drop_duplicates(subset=["__key__"], inplace=True, ignore_index=True)
df.drop(columns=["__key__"], inplace=True)

# Mostrar consolidado
st.success("✅ Consolidado listo (histórico + nuevos).")
st.dataframe(df, use_container_width=True)

# ──────────────────────────────────────────
# Descargas locales
# ──────────────────────────────────────────
st.subheader("3) Descargar planillas")
csv_bytes = df.to_csv(index=False).encode("utf-8")
st.download_button("📗 CSV (OrdenDelDia_Consejo.csv)", data=csv_bytes, file_name=CSV_NAME, mime="text/csv")

def to_xlsx_bytes(df0: pd.DataFrame) -> io.BytesIO:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df0.to_excel(w, index=False, sheet_name=SHEET_NAME)
    buf.seek(0); return buf

xlsx_buf = to_xlsx_bytes(df)
st.download_button("📘 Excel (OrdenDelDia_Consejo.xlsx)", data=xlsx_buf, file_name=XLSX_NAME,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ──────────────────────────────────────────
# Subir/Reemplazar en Drive (mantener IDs → Looker auto)
# ──────────────────────────────────────────
st.subheader("4) Subir/Reemplazar en Google Drive (para Looker Studio)")
if not creds_full:
    st.caption("ℹ️ Configurá `gcp_service_account` en Secrets para habilitar Drive.")
else:
    if st.button("🚀 Subir/Reemplazar CSV y Excel en Drive"):
        try:
            drv = drive_client(creds_full)
            csv_id  = drive_upload_replace(drv, folder_id, CSV_NAME, csv_bytes, "text/csv",
                                           file_id_hint=csv_id_secret)
            xlsx_id = drive_upload_replace(drv, folder_id, XLSX_NAME, xlsx_buf.getvalue(),
                                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                           file_id_hint=xlsx_id_secret)
            st.success("✅ Archivos actualizados en Drive.")
            st.caption(f"CSV id: {csv_id} · XLSX id: {xlsx_id}")
            st.info("Looker Studio se actualiza solo al mantener los mismos IDs.")
        except Exception as e:
            st.error("❌ Error subiendo a Drive.")
            st.exception(e)
