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

# ──────────────────────────────────────────
# CONFIG STREAMLIT
# ──────────────────────────────────────────
st.set_page_config(page_title="Extractor Órdenes del Día (Acumulativo)", page_icon="🗂️", layout="wide")
st.title("🗂️ Órdenes del Día → Consolidado Looker (CSV/XLSX en Drive)")

DEFAULT_FOLDER_ID = "REEMPLAZAR_CON_TU_CARPETA"
CSV_NAME  = "OrdenDelDia_Consejo.csv"
XLSX_NAME = "OrdenDelDia_Consejo.xlsx"
SHEET_NAME = "OrdenDelDia"

# ──────────────────────────────────────────
# ESQUEMA FIJO
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
# UTILIDADES TEXTO
# ──────────────────────────────────────────
def norm(s: str) -> str:
    if not isinstance(s, str): return ""
    s = unicodedata.normalize("NFKC", s).replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()

def extract_text_any(up) -> str:
    name = up.name.lower()
    if name.endswith(".pdf"):
        return norm(pdf_extract_text(up))
    if name.endswith(".docx"):
        doc = Document(up)
        return norm("\n".join(p.text for p in doc.paragraphs))
    return ""

def find_date_header(text: str) -> Tuple[str, str]:
    head = text[:1500]
    m = re.search(r"\b(\d{1,2})[\/\-\._](\d{1,2})[\/\-\._](\d{2,4})\b", head)
    if m:
        d, mth, y = m.groups()
        y = "20"+y if len(y) == 2 else y
        y = int(y)
        return str(y), f"{int(d):02d}/{int(mth):02d}/{y}"
    m2 = re.search(r"a\s+los\s+\d+\s+d[ií]as.*?mes\s+de\s+[a-záéíóú]+.*?dos\s+mil\s+([a-záéíóú]+)", head, re.I|re.S)
    if m2:
        mapa = {"veinte":2020,"veintiuno":2021,"veintidos":2022,"veintitres":2023,"veinticuatro":2024,"veinticinco":2025}
        y = unicodedata.normalize("NFKD", m2.group(1)).lower()
        y = y.replace("́","").translate(str.maketrans("áéíóú","aeiou"))
        if y in mapa:
            for ln in head.split("\n"):
                if len(ln.strip()) > 12:
                    return str(mapa[y]), ln.strip()
    m3 = re.search(r"\b(20\d{2})\b", head)
    if m3: return m3.group(1), ""
    return "", ""

# ──────────────────────────────────────────
# PARSER ORDEN DEL DÍA (robusto)
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

def split_lines(text: str):
    text = re.sub(r"[\u2022\u2023\u25CF\u25CB\u25A0•►▪▫]+", "\n", text)
    lines = [norm(ln.strip(" \t-—–•")) for ln in text.split("\n")]
    return [ln for ln in lines if ln]

def is_section_header(ln: str) -> str:
    for k, rx in SECTION_HEADERS.items():
        if rx.search(ln): return k
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

def extract_after(label_rx: re.Pattern, chunk: list) -> str:
    for ln in chunk:
        m = re.search(label_rx, ln)
        if m: return norm(re.sub(label_rx, "", ln)).strip(" .,:;–-\"'«»“”")
    return ""

def extract_unit(chunk: list) -> str:
    for ln in chunk:
        if is_faculty_line(ln): return ln
    for ln in chunk:
        if UNIT_LABELS.search(ln): return ln
    return ""

def cut_before_director(chunk: list) -> list:
    out = []
    for ln in chunk:
        if DIRECTOR_LABELS.search(ln): break
        out.append(ln)
    return out

def first_title_from(chunk: list) -> str:
    pre = cut_before_director(chunk)
    for ln in pre:
        if looks_title_line(ln): return ln.strip(" .,:;–-\"'«»“”")
    for ln in pre:
        if ln and not NO_TITLE_PREFIX.search(ln): return ln.strip(" .,:;–-\"'«»“”")
    return ""

def parse_items_by_section(lines: list, base: dict) -> list[dict]:
    rows, section, buf, current_unit = [], "", [], ""
    def flush():
        nonlocal buf, section, current_unit
        if not buf: return
        chunk = [ln for ln in buf if ln]
        if NO_TITLE_PREFIX.match(chunk[0]) and not any(DIRECTOR_LABELS.search(x) for x in chunk):
            buf.clear(); return
        unit_here = extract_unit(chunk) or current_unit
        def make_row():
            r = empty_row(base); r["Unidad académica de procedencia del proyecto"] = unit_here; return r
        if section == "proyectos":
            r = make_row(); r["proyectos de investigación"]="Sí"
            r["Nombre del proyecto de investigación"]=first_title_from(chunk)
            r["Director del Proyecto"]=extract_after(DIRECTOR_LABELS, chunk)
            r["Integrantes del equipo de investigación"]=extract_after(TEAM_LABELS, chunk)
            r["Unidad académica de procedencia del proyecto"]=unit_here
            if r["Nombre del proyecto de investigación"]: rows.append(r)
        elif section == "final":
            r = make_row(); r["Informe Final"]="Sí"
            r["Nombre del proyecto de investigación del Informe Final"]=first_title_from(chunk)
            r["Director del Proyecto del Informe Final"]=extract_after(DIRECTOR_LABELS, chunk)
            r["Integrantes del equipo de investigación del Informe Final"]=extract_after(TEAM_LABELS, chunk)
            r["Unidad académica de procedencia del proyecto del Informe Final"]=unit_here
            if r["Nombre del proyecto de investigación del Informe Final"]: rows.append(r)
        elif section == "avance":
            r = make_row(); r["Informe de avance"]="Sí"
            r["Nombre del proyecto de investigación del Informe de avance"]=first_title_from(chunk)
            r["Director del Proyecto del Informe de avance"]=extract_after(DIRECTOR_LABELS, chunk)
            r["Integrantes del equipo de investigación del Informe de avance"]=extract_after(TEAM_LABELS, chunk)
            r["Unidad académica de procedencia del proyecto del Informe de avance"]=unit_here
            if r["Nombre del proyecto de investigación del Informe de avance"]: rows.append(r)
        elif section == "catedra":
            r = make_row(); r["Proyectos de investigación de cátedra"]="Sí"
            r["Nombre del proyecto de investigación cátedra"]=first_title_from(chunk)
            r["Director del Proyecto del Informe de cátedra"]=extract_after(DIRECTOR_LABELS, chunk)
            r["Integrantes del equipo de investigación del proyecto de cátedra"]=extract_after(TEAM_LABELS, chunk)
            r["Unidad académica de procedencia del proyecto de cátedra"]=unit_here
            if r["Nombre del proyecto de investigación cátedra"]: rows.append(r)
        elif section == "publica":
            r = make_row(); r["Publicación"]="Sí"
            joined=" ".join(chunk); tipo=""
            if re.search(r"\brevista\b", joined, re.I): tipo="revista científica"
            elif re.search(r"\blibro\b", joined, re.I): tipo="libro"
            elif re.search(r"\b(congreso|ponencia|presentaci[oó]n)\b", joined, re.I): tipo="presentación a congreso"
            elif re.search(r"p[oó]ster|poster", joined, re.I): tipo="póster"
            elif re.search(r"\bcuadernos\b", joined, re.I): tipo="revista Cuadernos"
            elif re.search(r"\bmanual\b", joined, re.I): tipo="manual"
            r["Tipo de publicación (revista científica, libro, presentación a congreso, póster, revista Cuadernos, manual)"]=tipo
            r["Docente o investigador incluida en la publicación"]=extract_after(re.compile(r"(Autor(?:es)?|Docente|Investigador(?:es)?)\s*:", re.I), chunk)
            r["Unidad académica (Publicación)"]=unit_here
            rows.append(r)
        elif section == "categ":
            r = make_row(); r["Categorización de docentes"]="Sí"
            joined=" ".join(chunk)
            mcat=re.search(r"(Categor[ií]a\s*[:\-]?\s*[IVX]+|Investigador(?:\s+\w+){0,3})", joined, re.I)
            r["Categoría alcanzada por el docente como docente investigador"]=mcat.group(0) if mcat else ""
            cand=next((ln for ln in chunk if re.search(r"^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+", ln)), "")
            r["Nombre del docente categorizado como investigador"]=cand
            r["Unidad académica (Categorización)"]=unit_here
            rows.append(r)
        elif section == "beca":
            r = make_row()
            joined=" ".join(chunk)
            if re.search(r"postdoctoral", joined, re.I):
                r["Becario de beca cofinanciada postdoctoral"]="Sí"
                r["Nombre del becario postdoctoral"]=first_title_from(chunk) or extract_after(re.compile(r"(Becari[oa]|Nombre)\s*:", re.I), chunk)
            else:
                r["Becario de beca cofinanciada doctoral"]="Sí"
                r["Nombre del becario doctoral"]=first_title_from(chunk) or extract_after(re.compile(r"(Becari[oa]|Nombre)\s*:", re.I), chunk)
            rows.append(r)
        else:
            r = make_row(); r["OTROS TEMAS"]=" ".join(chunk); rows.append(r)
        buf.clear(); current_unit = unit_here
    for ln in lines:
        sec = is_section_header(ln)
        if sec: flush(); section = sec; continue
        if is_faculty_line(ln): flush(); current_unit = ln; continue
        buf.append(ln)
    flush()
    return rows

# ──────────────────────────────────────────
# GOOGLE DRIVE
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

def drive_client(creds): return build("drive", "v3", credentials=creds)

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
# ESTADO: Cargar histórico al iniciar la sesión
# ──────────────────────────────────────────
if "consolidado" not in st.session_state:
    st.session_state.consolidado = pd.DataFrame(columns=FIXED_COLUMNS)
    creds_init = get_creds(["https://www.googleapis.com/auth/drive"])
    folder_id = st.secrets.get("drive_folder_id", DEFAULT_FOLDER_ID)
    csv_id_secret  = st.secrets.get("drive_csv_file_id", "")
    try:
        if creds_init:
            drv = drive_client(creds_init)
            if csv_id_secret:
                base = drive_read_csv_by_id(drv, csv_id_secret)
            else:
                fid = drive_find_file(drv, CSV_NAME, folder_id)
                base = drive_read_csv_by_id(drv, fid) if fid else pd.DataFrame(columns=FIXED_COLUMNS)
            # Asegurar esquema
            for c in FIXED_COLUMNS:
                if c not in base.columns: base[c] = ""
            st.session_state.consolidado = base[FIXED_COLUMNS]
    except Exception:
        pass  # si no existe aún, arrancamos vacío

# ──────────────────────────────────────────
# 1) Cargar archivos NUEVOS
# ──────────────────────────────────────────
st.subheader("1) Cargar nuevos Órdenes del Día (PDF/DOCX)")
uploads = st.file_uploader("📂 Archivos", type=["pdf","docx"], accept_multiple_files=True)

# Procesar lote actual
df_lote = pd.DataFrame(columns=FIXED_COLUMNS)
if uploads:
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
            r = empty_row(base); r["OTROS TEMAS"] = raw[:1500] + ("…" if len(raw) > 1500 else ""); rows=[r]
        all_rows.extend(rows)
    if all_rows:
        df_lote = pd.DataFrame(all_rows)
        for c in FIXED_COLUMNS:
            if c not in df_lote.columns: df_lote[c] = ""
        df_lote = df_lote[FIXED_COLUMNS]
        st.success(f"✅ Lote preparado: {len(df_lote)} fila(s).")
        st.dataframe(df_lote, use_container_width=True)
    else:
        st.info("No se detectaron ítems en los archivos cargados.")

# ──────────────────────────────────────────
# 2) Agregar al CONSOLIDADO de la sesión
# ──────────────────────────────────────────
st.subheader("2) Agregar este lote al consolidado")
colA, colB = st.columns(2)
with colA:
    if st.button("➕ Agregar lote al consolidado", disabled=df_lote.empty):
        df = pd.concat([st.session_state.consolidado, df_lote], ignore_index=True)
        # deduplicar por año + títulos de proyecto/avance/final
        key = df["año"].astype(str) + "||" + \
              df["Nombre del proyecto de investigación"].fillna("") + "||" + \
              df["Nombre del proyecto de investigación del Informe de avance"].fillna("") + "||" + \
              df["Nombre del proyecto de investigación del Informe Final"].fillna("")
        df = df.loc[~key.duplicated()].reset_index(drop=True)
        st.session_state.consolidado = df
        st.success(f"🧩 Lote agregado. Consolidado: {len(df)} fila(s).")
with colB:
    if st.button("🧹 Reiniciar consolidado (solo esta sesión)"):
        st.session_state.consolidado = pd.DataFrame(columns=FIXED_COLUMNS)
        st.info("Consolidado de sesión reiniciado. (No borra nada en Drive)")

# Mostrar consolidado actual
st.subheader("Consolidado en esta sesión (se sube a Drive)")
st.dataframe(st.session_state.consolidado, use_container_width=True)

# ──────────────────────────────────────────
# 3) Descargar consolidado
# ──────────────────────────────────────────
st.subheader("3) Descargas (consolidado)")
def to_xlsx_bytes(df0: pd.DataFrame) -> io.BytesIO:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df0.to_excel(w, index=False, sheet_name=SHEET_NAME)
    buf.seek(0); return buf

csv_bytes = st.session_state.consolidado.to_csv(index=False).encode("utf-8")
xlsx_buf  = to_xlsx_bytes(st.session_state.consolidado)

st.download_button("📗 CSV (consolidado)", data=csv_bytes, file_name=CSV_NAME, mime="text/csv")
st.download_button("📘 Excel (consolidado)", data=xlsx_buf, file_name=XLSX_NAME,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ──────────────────────────────────────────
# 4) Subir/Reemplazar en Drive (CONSOLIDADO)
# ──────────────────────────────────────────
st.subheader("4) Subir/Reemplazar en Google Drive (consolidado → Looker)")
folder_id = st.secrets.get("drive_folder_id", DEFAULT_FOLDER_ID)
creds = get_creds(["https://www.googleapis.com/auth/drive"])
csv_id_secret  = st.secrets.get("drive_csv_file_id", "")
xlsx_id_secret = st.secrets.get("drive_xlsx_file_id", "")

if not creds:
    st.caption("Configura `gcp_service_account` en Secrets para habilitar Drive.")
else:
    if st.button("🚀 Subir/Reemplazar CONSOLIDADO en Drive"):
        try:
            drv = drive_client(creds)
            csv_id  = drive_upload_replace(drv, folder_id, CSV_NAME, csv_bytes, "text/csv",
                                           file_id_hint=csv_id_secret)
            xlsx_id = drive_upload_replace(drv, folder_id, XLSX_NAME, xlsx_buf.getvalue(),
                                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                           file_id_hint=xlsx_id_secret)
            st.success("✅ Consolidado actualizado en Drive.")
            st.caption(f"CSV id: {csv_id} · XLSX id: {xlsx_id}")
            st.info("Looker Studio se actualiza solo al mantener los mismos IDs.")
        except Exception as e:
            st.error("❌ Error subiendo a Drive.")
            st.exception(e)
