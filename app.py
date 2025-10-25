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
st.set_page_config(page_title="Extractor de Órdenes del Día", page_icon="🗂️", layout="wide")
st.title("🗂️ Extractor de Órdenes del Día → Planilla estándar + Drive (Looker-ready)")

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
    """Devuelve texto plano de PDF o DOCX."""
    name = uploaded.name.lower()
    if name.endswith(".pdf"):
        return norm(pdf_extract_text(uploaded))
    if name.endswith(".docx"):
        doc = Document(uploaded)
        return norm("\n".join(p.text for p in doc.paragraphs))
    return ""

def find_date_header(text: str) -> Tuple[str, str]:
    """
    Intenta detectar fecha/año de la reunión en el encabezado o nombre.
    Formatos esperados: 20/02/2025, 21-08-25, 18_04_24, 23/10/2025, etc.
    """
    head = text[:1200]
    # dd[/-_]mm[/-_]yyyy | dd[/-_]mm[/-_]yy
    m = re.search(r"\b(\d{1,2})[\/\-\._](\d{1,2})[\/\-\._](\d{2,4})\b", head)
    if m:
        d, mth, y = m.groups()
        y = "20"+y if len(y) == 2 else y
        y = int(y)
        d = f"{int(d):02d}"; mth = f"{int(mth):02d}"
        return str(y), f"{d}/{mth}/{y}"
    # “a los … días del mes de … de dos mil …”
    m2 = re.search(r"a\s+los\s+\d+\s+d[ií]as.*?mes\s+de\s+[a-záéíóú]+.*?dos\s+mil\s+([a-záéíóú]+)", head, re.I|re.S)
    if m2:
        mapa = {
            "veinte":2020,"veintiuno":2021,"veintidos":2022,"veintitres":2023,"veinticuatro":2024,
            "veinticinco":2025,"veintiseis":2026,"veintisiete":2027,"veintiocho":2028,"veintinueve":2029,"treinta":2030
        }
        y = mapa.get(unicodedata.normalize("NFKD", m2.group(1)).replace("́","").lower())
        if y:
            # tomar primera línea larga como "fecha textual"
            for ln in head.split("\n"):
                if len(ln.strip()) > 12:
                    return str(y), ln.strip()
    # fallback: primer año 20xx
    m3 = re.search(r"\b(20\d{2})\b", head)
    if m3:
        return m3.group(1), ""
    return "", ""

# ──────────────────────────────────────────
# ESQUEMA FIJO DE COLUMNAS (para Looker Studio)
# ──────────────────────────────────────────
FIXED_COLUMNS = [
    "año",
    "fecha",

    "proyectos de investigación",
    "Nombre del proyecto de investigación",
    "Director del Proyecto",
    "Integrantes del equipo de investigación",
    "Unidad académica de procedencia del proyecto",

    "Informe de avance",
    "Nombre del proyecto de investigación del Informe de avance",
    "Director del Proyecto del Informe de avance",
    "Integrantes del equipo de investigación del Informe de avance",
    "Unidad académica de procedencia del proyecto del Informe de avance",

    "Informe Final",
    "Nombre del proyecto de investigación del Informe Final",
    "Director del Proyecto del Informe Final",
    "Integrantes del equipo de investigación del Informe Final",
    "Unidad académica de procedencia del proyecto del Informe Final",

    "Proyectos de investigación de cátedra",
    "Nombre del proyecto de investigación cátedra",
    "Director del Proyecto del Informe de cátedra",
    "Integrantes del equipo de investigación del proyecto de cátedra",
    "Unidad académica de procedencia del proyecto de cátedra",

    "Publicación",
    "Tipo de publicación (revista científica, libro, presentación a congreso, póster, revista Cuadernos, manual)",
    "Docente o investigador incluida en la publicación",
    "Unidad académica (Publicación)",  # ← desambiguado para mantener unicidad de columnas

    "Categorización de docentes",
    "Nombre del docente categorizado como investigador",
    "Categoría alcanzada por el docente como docente investigador",
    "Unidad académica (Categorización)",  # ← desambiguado

    # Becarios: unificamos como dos pares (doctoral/postdoctoral)
    "Becario de beca cofinanciada doctoral",
    "Nombre del becario doctoral",
    "Becario de beca cofinanciada postdoctoral",
    "Nombre del becario postdoctoral",

    "OTROS TEMAS"  # todo lo que no encaje arriba
]

def empty_row(base: Dict[str, Any]=None) -> Dict[str, Any]:
    row = {col: "" for col in FIXED_COLUMNS}
    if base:
        row.update({k:v for k,v in base.items() if k in row})
    return row

# ──────────────────────────────────────────
# PARSER DE SECCIONES (ÓRDENES DEL DÍA)
# ──────────────────────────────────────────
SECTION_MAP = {
    "proyectos": re.compile(r"^(proyectos? (de )?investigaci[oó]n|presentaci[oó]n de proyectos?)\b", re.I),
    "avance":    re.compile(r"^informes? de avance\b", re.I),
    "final":     re.compile(r"^informes? finales?\b", re.I),
    "catedra":   re.compile(r"(proyectos? (de )?c[aá]tedra|proyectos? cuadernos)", re.I),
    "publica":   re.compile(r"^publicaci[oó]n|^publicaciones\b", re.I),
    "categ":     re.compile(r"^categorizaci[oó]n", re.I),
    "beca":      re.compile(r"^becari[oa]s?", re.I),
}

def split_lines(text: str) -> List[str]:
    lines = [ln.strip(" -•\t") for ln in text.split("\n")]
    return [ln for ln in lines if ln]

def current_section_of(line: str) -> str:
    l = line.strip()
    for key, rx in SECTION_MAP.items():
        if rx.search(l):
            return key
    return ""

def extract_name_after(label: str, txt: str) -> str:
    m = re.search(label + r"\s*:\s*(.+)", txt, re.I)
    return norm(m.group(1)) if m else ""

def parse_people_list(s: str) -> str:
    s = re.sub(r"\s*[–—-]\s*", " – ", s)
    s = s.replace(" ,", ",")
    return norm(s)

def parse_unit(s: str) -> str:
    m = re.search(r"(Facultad|Escuela|Instituto|Vicerrectorado)[^\n]*", s, re.I)
    return norm(m.group(0)) if m else ""

def looks_title_line(s: str) -> bool:
    # heurística para líneas de TÍTULO
    if len(s) < 6: return False
    if re.search(r"(Director|Directora|Integrantes|Equipo|Codirector|Unidad)", s, re.I): return False
    cap = sum(1 for c in s if c.isupper())
    alpha = sum(1 for c in s if c.isalpha())
    return (alpha > 0 and (cap/alpha) > 0.4) or s.istitle() or "“" in s or '"' in s

def parse_items_by_section(lines: List[str], base_meta: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Crea filas WIDE conforme a FIXED_COLUMNS.
    Una fila por ítem. Campos no aplicables quedan vacíos.
    """
    rows: List[Dict[str, Any]] = []
    sec = ""
    buf: List[str] = []

    def flush_buffer(section: str, buffer: List[str]):
        if not buffer: return
        chunk = "\n".join(buffer)
        row_base = empty_row(base_meta)
        # Ruteo por sección
        if section == "proyectos":
            row_base["proyectos de investigación"] = "Sí"
            # Título
            t = extract_name_after(r"(Denominaci[oó]n|T[ií]tulo|Proyecto)", chunk) or \
                next((ln for ln in buffer if looks_title_line(ln)), "")
            row_base["Nombre del proyecto de investigación"] = t.strip("“”\"' ")
            # Director / Integrantes / Unidad
            row_base["Director del Proyecto"] = extract_name_after(r"Director(?:a)?", chunk)
            integ = extract_name_after(r"(Integrantes|Equipo)", chunk)
            row_base["Integrantes del equipo de investigación"] = parse_people_list(integ)
            row_base["Unidad académica de procedencia del proyecto"] = parse_unit(chunk)
            rows.append(row_base)

        elif section == "avance":
            row_base["Informe de avance"] = "Sí"
            t = extract_name_after(r"(Denominaci[oó]n|T[ií]tulo|Proyecto)", chunk) or \
                next((ln for ln in buffer if looks_title_line(ln)), "")
            row_base["Nombre del proyecto de investigación del Informe de avance"] = t.strip("“”\"' ")
            row_base["Director del Proyecto del Informe de avance"] = extract_name_after(r"Director(?:a)?", chunk)
            row_base["Integrantes del equipo de investigación del Informe de avance"] = parse_people_list(
                extract_name_after(r"(Integrantes|Equipo)", chunk)
            )
            row_base["Unidad académica de procedencia del proyecto del Informe de avance"] = parse_unit(chunk)
            rows.append(row_base)

        elif section == "final":
            row_base["Informe Final"] = "Sí"
            t = extract_name_after(r"(Denominaci[oó]n|T[ií]tulo|Proyecto)", chunk) or \
                next((ln for ln in buffer if looks_title_line(ln)), "")
            row_base["Nombre del proyecto de investigación del Informe Final"] = t.strip("“”\"' ")
            row_base["Director del Proyecto del Informe Final"] = extract_name_after(r"Director(?:a)?", chunk)
            row_base["Integrantes del equipo de investigación del Informe Final"] = parse_people_list(
                extract_name_after(r"(Integrantes|Equipo)", chunk)
            )
            row_base["Unidad académica de procedencia del proyecto del Informe Final"] = parse_unit(chunk)
            rows.append(row_base)

        elif section == "catedra":
            row_base["Proyectos de investigación de cátedra"] = "Sí"
            t = extract_name_after(r"(Denominaci[oó]n|T[ií]tulo|Proyecto|Asignatura)", chunk) or \
                next((ln for ln in buffer if looks_title_line(ln)), "")
            row_base["Nombre del proyecto de investigación cátedra"] = t.strip("“”\"' ")
            row_base["Director del Proyecto del Informe de cátedra"] = extract_name_after(r"Director(?:a)?", chunk)
            row_base["Integrantes del equipo de investigación del proyecto de cátedra"] = parse_people_list(
                extract_name_after(r"(Integrantes|Equipo|Docentes)", chunk)
            )
            row_base["Unidad académica de procedencia del proyecto de cátedra"] = parse_unit(chunk)
            rows.append(row_base)

        elif section == "publica":
            row_base["Publicación"] = "Sí"
            # Tipo
            tipo = ""
            for k in ["revista", "libro", "congreso", "póster", "poster", "cuadernos", "manual"]:
                if re.search(k, chunk, re.I):
                    mapa = {
                        "revista": "revista científica", "libro": "libro",
                        "congreso":"presentación a congreso", "póster":"póster", "poster":"póster",
                        "cuadernos":"revista Cuadernos", "manual":"manual"
                    }
                    tipo = mapa[k]; break
            row_base["Tipo de publicación (revista científica, libro, presentación a congreso, póster, revista Cuadernos, manual)"] = tipo
            # Autor y UA
            row_base["Docente o investigador incluida en la publicación"] = extract_name_after(r"(Autor(?:es)?|Docente|Investigador)", chunk)
            row_base["Unidad académica (Publicación)"] = parse_unit(chunk)
            rows.append(row_base)

        elif section == "categ":
            row_base["Categorización de docentes"] = "Sí"
            row_base["Nombre del docente categorizado como investigador"] = extract_name_after(r"(Docente|Nombre)", chunk) or \
                next((ln for ln in buffer if re.search(r"^[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑ ]+,", ln)), "")
            row_base["Categoría alcanzada por el docente como docente investigador"] = extract_name_after(r"(Categor[ií]a|Tipo)", chunk)
            row_base["Unidad académica (Categorización)"] = parse_unit(chunk)
            rows.append(row_base)

        elif section == "beca":
            # Detectar doctoral/postdoctoral
            if re.search(r"postdoctoral", chunk, re.I):
                row_base["Becario de beca cofinanciada postdoctoral"] = "Sí"
                row_base["Nombre del becario postdoctoral"] = extract_name_after(r"(Becari[oa]|Nombre)", chunk) or \
                    next((ln for ln in buffer if re.search(r"^[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑ ]+$", ln)), "")
            else:
                row_base["Becario de beca cofinanciada doctoral"] = "Sí"
                row_base["Nombre del becario doctoral"] = extract_name_after(r"(Becari[oa]|Nombre)", chunk) or \
                    next((ln for ln in buffer if re.search(r"^[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑ ]+$", ln)), "")
            rows.append(row_base)

        else:
            # OTROS TEMAS
            row_base["OTROS TEMAS"] = chunk
            rows.append(row_base)

    for ln in lines:
        sec_here = current_section_of(ln)
        if sec_here:
            # cambia de sección → flush
            flush_buffer(sec, buf)
            sec = sec_here
            buf = []
        else:
            buf.append(ln)
    flush_buffer(sec, buf)
    return rows

# ──────────────────────────────────────────
# GOOGLE DRIVE (reemplazo por nombre)
# ──────────────────────────────────────────
def get_creds(scopes):
    sa = st.secrets.get("gcp_service_account")
    if not sa:
        return None
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

def drive_upload_replace(drive, folder_id: str, name: str, data: bytes, mime: str):
    file_id = drive_find_file(drive, name, folder_id)
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
    if file_id:
        # update contenido (mantiene el mismo id → Looker sigue apuntando)
        drive.files().update(fileId=file_id, media_body=media).execute()
        return file_id
    else:
        meta = {"name": name, "parents": [folder_id]}
        f = drive.files().create(body=meta, media_body=media, fields="id").execute()
        return f["id"]

# ──────────────────────────────────────────
# UI
# ──────────────────────────────────────────
st.subheader("1) Subí el/los Órdenes del Día (PDF o DOCX)")
uploads = st.file_uploader("📂 Archivos", type=["pdf","docx"], accept_multiple_files=True)

if not uploads:
    st.info("Subí al menos un archivo para continuar.")
    st.stop()

all_rows = []
for up in uploads:
    raw = extract_text_any(up)
    if not raw:
        st.warning(f"No se pudo leer: {up.name}")
        continue

    # Año / Fecha (metadatos base para cada ítem)
    year, date_str = find_date_header(raw)
    base = {"año": year, "fecha": date_str}
    lines = split_lines(raw)
    rows = parse_items_by_section(lines, base)
    # si el documento no trae ninguna sección detectable, meterlo como "OTROS TEMAS"
    if not rows:
        r = empty_row(base)
        r["OTROS TEMAS"] = raw[:1500] + ("…" if len(raw) > 1500 else "")
        rows = [r]
    all_rows.extend(rows)

if not all_rows:
    st.error("No se detectaron ítems en los Órdenes del Día cargados.")
    st.stop()

# Construcción del DataFrame con columnas fijas (orden inmutable)
df = pd.DataFrame(all_rows)
# Asegurar todas las columnas (y el orden)
for col in FIXED_COLUMNS:
    if col not in df.columns:
        df[col] = ""
df = df[FIXED_COLUMNS]

st.success("✅ Órdenes del Día procesados.")
st.dataframe(df, use_container_width=True)

# ──────────────────────────────────────────
# Descargas locales
# ──────────────────────────────────────────
st.subheader("2) Descargar planillas")
# CSV
csv_bytes = df.to_csv(index=False).encode("utf-8")
st.download_button("📗 CSV (OrdenDelDia_Consejo.csv)", data=csv_bytes, file_name=CSV_NAME, mime="text/csv")

# XLSX
def to_xlsx_bytes(df0: pd.DataFrame) -> io.BytesIO:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df0.to_excel(w, index=False, sheet_name=SHEET_NAME)
    buf.seek(0); return buf

xlsx_buf = to_xlsx_bytes(df)
st.download_button("📘 Excel (OrdenDelDia_Consejo.xlsx)", data=xlsx_buf, file_name=XLSX_NAME,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ──────────────────────────────────────────
# Subida a Google Drive (reemplazo)
# ──────────────────────────────────────────
st.subheader("3) Subir/Reemplazar en Google Drive (para Looker Studio)")
folder_id = st.secrets.get("drive_folder_id", DEFAULT_FOLDER_ID)
creds = get_creds(["https://www.googleapis.com/auth/drive.file"])

if not creds:
    st.caption("ℹ️ Configurá `gcp_service_account` en Secrets para habilitar Drive.")
else:
    if st.button("🚀 Subir/Reemplazar CSV y Excel en Drive"):
        try:
            drv = drive_client(creds)
            csv_id  = drive_upload_replace(drv, folder_id, CSV_NAME,  csv_bytes, "text/csv")
            xlsx_id = drive_upload_replace(drv, folder_id, XLSX_NAME, xlsx_buf.getvalue(),
                                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("✅ Archivos actualizados en Drive con el mismo nombre (IDs preservados si ya existían).")
            st.caption(f"CSV id: {csv_id} · XLSX id: {xlsx_id}")
            st.info("Si tus fuentes de Looker Studio referencian estos archivos por ID o por nombre en esa carpeta, se verán actualizadas automáticamente.")
        except Exception as e:
            st.error(f"Error subiendo a Drive: {e}")
