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
st.set_page_config(page_title="Extractor de ACTAS del Consejo", page_icon="📑", layout="wide")
st.title("📑 Extractor de ACTAS → Base institucional (7 temas + Año)")

DEFAULT_FOLDER_ID = "1O7xo7cCGkSujhUXIl3fv51S-cniVLmnh"  # carpeta fallback

# ──────────────────────────────────────────
# UTILIDADES GENERALES
# ──────────────────────────────────────────
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
    head = text[:1500]
    m = re.search(
        r"a\s+los\s+\d+\s+d[ií]as.*?mes\s+de\s+[a-záéíóú]+.*?de\s+dos\s+mil\s+[a-záéíóú]+",
        head, flags=re.IGNORECASE | re.DOTALL
    )
    if m:
        return norm(m.group(0))
    for ln in head.split("\n"):
        ln = ln.strip()
        if len(ln) > 20:
            return ln
    return text.split("\n", 1)[0]

def infer_year_from_text(s: str, full_text: str = None):
    if not isinstance(s, str):
        s = ""
    t = unicodedata.normalize("NFKD", s).lower()
    t = t.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
    m = re.search(r"dos\s+mil\s+([a-z]+)", t)
    mapa = {
        "veinte": 2020, "veintiuno": 2021, "veintidos": 2022, "veintitres": 2023,
        "veinticuatro": 2024, "veinticinco": 2025, "veintiseis": 2026,
        "veintisiete": 2027, "veintiocho": 2028, "veintinueve": 2029, "treinta": 2030
    }
    if m and m.group(1) in mapa:
        return mapa[m.group(1)]
    scope = (full_text or s or "")[:1500]
    nums = [int(x) for x in re.findall(r"\b(20\d{2})\b", scope)]
    nums = [n for n in nums if 2000 <= n <= 2100]
    if nums:
        return max(nums)
    return None

# ──────────────────────────────────────────
# CLASIFICACIÓN (7 temas)
# ──────────────────────────────────────────
TOPIC_MAP = {
    "Proyectos de investigación": [
        r"\bproyectos? de (investigaci[oó]n|convocatoria abierta)\b",
        r"\bpresentaci[oó]n de proyectos?\b", r"\bprojovi\b", r"\bpid\b", r"\bppi\b"
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
        r"\bcategorizaci[oó]n\b", r"\bsolicitud de categorizaci[oó]n\b",
        r"\bcategorizaciones? extraordinarias?\b"
    ],
    "Jornadas de investigación": [
        r"\bjornadas? de investigaci[oó]n\b", r"\bjornadas? internas\b"
    ],
    "Cursos de capacitación": [
        r"\bcursos? de capacitaci[oó]n\b", r"\btaller(es)?\b", r"\bcapacitaci[oó]n\b"
    ],
}

def classify_topic(text: str) -> str:
    t = text.lower()
    for topic, pats in TOPIC_MAP.items():
        for pat in pats:
            if re.search(pat, t, re.IGNORECASE):
                return topic
    return "Proyectos de investigación"  # fallback

# ──────────────────────────────────────────
# EXTRACCIÓN / FILTROS DE ITEMS
# ──────────────────────────────────────────
NARRATIVE_STARTS = (
    r"^se\s", r"^los\s+informes", r"^las\s+categor[ií]as", r"^fueron\s+consultadas",
    r"^siendo\s+las\s+\d", r"^lectura\s+del\s+acta", r"^presentaci[oó]n de (informes|propuestas)",
    r"^nuevo?s?\s+requerimientos", r"^propuestas?\s+de\s+investigaci[oó]n\s+a\s+la\s+minera",
)

def is_narrative(item: str) -> bool:
    t = norm(item).lower()
    if len(t) < 20:
        return True
    for pat in NARRATIVE_STARTS:
        if re.search(pat, t):
            return True
    # si no contiene rastro de “proyecto/projovi/título/director”, es probable que sea narrativo
    if not re.search(r"(proyecto|projovi|pid|ppi|t[ií]tulo|denominaci[oó]n|director)", t):
        # pero podría ser un TÍTULO en MAYÚSCULAS (ej. SALUD MENTAL EN ADOLESCENTES…)
        first_line = t.split("\n", 1)[0]
        if not re.search(r"[A-ZÁÉÍÓÚÑ]{3,}", item):
            return True
    return False

def clean_person_titles(s: str) -> str:
    return re.sub(r"\b(Dr\.?|Dra\.?|Lic\.?|Prof\.?|Mg\.?|Ing\.?)\b\.?\s*", "", s, flags=re.IGNORECASE)

def extract_title_strict(text: str) -> str:
    """
    Devuelve SOLO el nombre del proyecto/actividad.
    Reglas:
      1) Prioriza campos rotulados: Denominación:/Título:/Proyecto:
      2) Si hay 'Director', toma lo que esté ANTES de 'Director'
      3) Si no, toma la primera línea que parezca un título (MAYÚSCULAS/Title Case)
      4) Limpia comillas y frases administrativas
    """
    t = norm(text)

    # 1) rotulados
    m = re.search(r"(Denominaci[oó]n|T[ií]tulo|Proyecto)\s*:\s*(.+)", t, re.IGNORECASE)
    if m:
        cand = m.group(2)
        cand = re.split(r"\bDirector(?:a)?\b\s*:", cand, flags=re.IGNORECASE)[0]
        cand = re.split(r"[;\n]", cand, maxsplit=1)[0]
        return norm(cand.strip(" .,:;–-\"'«»“”"))

    # 2) si aparece “Director: …”, tomar lo anterior a esa etiqueta
    if "Director" in t or "Directora" in t:
        pre = re.split(r"\bDirector(?:a)?\b\s*:", t, flags=re.IGNORECASE)[0]
        # suele venir “PROJOVI: TÍTULO …  Director: …”
        # quedarnos con la última línea no vacía de 'pre'
        pre_lines = [ln.strip() for ln in pre.split("\n") if ln.strip()]
        if pre_lines:
            cand = pre_l_
