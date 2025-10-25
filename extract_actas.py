# -*- coding: utf-8 -*-
"""
Extrae automáticamente temas de ACTAS (PDF/DOCX) del Consejo de Investigación
y genera un CSV normalizado listo para Google Sheets / Looker Studio.

Campos:
- Acta, Fecha, Facultad, Tipo_tema, Titulo_o_denominacion, Director, Estado,
  Destino_publicacion, Fuente_archivo
"""
import os, re, glob, unicodedata, yaml
import pandas as pd

# ---- Lectores de documentos
from pdfminer.high_level import extract_text as pdf_extract_text
from docx import Document as DocxDocument

# -------------------- Configuración --------------------
ACTAS_DIR = os.environ.get("ACTAS_DIR", "actas")
OUTPUT_CSV = os.environ.get("OUTPUT_CSV", "actas_extraccion.csv")
CONFIG_PATH = os.environ.get("CONFIG_PATH", "config_patterns.yaml")

# -------------------- Utilidades --------------------
def _norm(s: str) -> str:
    if s is None:
        return ""
    s = s.replace("\x00", " ")
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()

def read_pdf(path: str) -> str:
    try:
        return _norm(pdf_extract_text(path))
    except Exception as e:
        return ""

def read_docx(path: str) -> str:
    try:
        doc = DocxDocument(path)
        text = "\n".join(p.text for p in doc.paragraphs)
        return _norm(text)
    except Exception:
        return ""

def load_config(path: str):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def find_acta_number(text: str) -> str:
    m = re.search(r"ACTA\s+N[º°]?\s*([0-9]+)", text, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""

def find_date_text(text: str) -> str:
    # Toma la primera frase del encabezado con "En la ciudad" o líneas con día/mes
    m = re.search(r"En la ciudad.*?\n(.*)", text, flags=re.IGNORECASE)
    if m:
        return _norm(m.group(1))
    m2 = re.search(r"a los\s+.+?d[ií]as.*?del mes de\s+.+?\s+de\s+dos mil.*", text, flags=re.IGNORECASE)
    return _norm(m2.group(0)) if m2 else ""

def split_sections(text: str, config):
    """
    Devuelve lista de (nombre_seccion, start, end)
    según patrones del YAML.
    """
    entries = []
    # Construimos gran regex con nombres nombrados
    named_patterns = []
    for i, sec in enumerate(config["sections"]):
        pat = "|".join([re.escape(p) for p in sec["patterns"]])
        named_patterns.append(f"(?P<s{i}>\\b(?:{pat})\\b)")
    if not named_patterns:
        return [("General", 0, len(text))]
    pattern = re.compile("|".join(named_patterns), flags=re.IGNORECASE)
    hits = []
    for m in pattern.finditer(text):
        sec_name = None
        for i, sec in enumerate(config["sections"]):
            if m.group(f"s{i}"):
                sec_name = sec["name"]
                break
        if sec_name:
            hits.append((sec_name, m.start()))
    if not hits:
        return [("General", 0, len(text))]
    hits.sort(key=lambda x: x[1])
    spans = []
    for i, (name, start) in enumerate(hits):
        end = hits[i+1][1] if i+1 < len(hits) else len(text)
        spans.append((name, start, end))
    return spans

FACULTY_HDR = re.compile(r"^(Facultad|Instituto Superior|Vicerrectorado|Escuela)\b.*", re.IGNORECASE)

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
    if not blocks:
        return [(None, section_text)]
    return blocks

ITEM_SPLIT = re.compile(r"\n\s*(?:[\u2022•\-]|[\u25CF\u25A0\u25E6]|\d+\.)\s*")

def extract_candidate_items(text: str):
    # Divide por viñetas y también por frases con “Proyecto:”, “PROJOVI:”, “Denominación…”
    parts = ITEM_SPLIT.split("\n" + text)
    cands = []
    for p in parts:
        p = p.strip(" ;\n\t")
        if len(p) < 6:
            continue
        if re.search(r"(Proyecto|Denominaci[oó]n|PROJOVI|Informe|Categorizaci[oó]n|Baja del proyecto|Revista|Cuadernos)", p, re.IGNORECASE):
            cands.append(p)
    if not cands:
        cands = [text.strip()]
    return cands

def infer_estado(text: str) -> str:
    t = text.lower()
    if "baja del proyecto" in t or re.search(r"\bbaja\b", t):
        return "Baja"
    if "prórroga" in t or "prorroga" in t:
        return "Prórroga"
    if "aprob" in t and ("elev" in t or "enviado" in t):
        return "Aprobado y elevado"
    if "aprob" in t:
        return "Aprobado"
    if "solicitud de categorización" in t or "solicitud de categorizacion" in t:
        return "Solicitud"
    return ""

def infer_destino_publicacion(text: str) -> str:
    if re.search(r"\b(Cuadernos|revista\s+cuadernos|Cuadernos de la Secretar[ií]a de Investigaci[oó]n)\b", text, re.IGNORECASE):
        return "Revista Cuadernos"
    return ""

def extract_title_director(text: str):
    title, director = None, None
    # Proyecto: Título. Director: Nombre
    m = re.search(r"Proyecto\s*:\s*(.+?)(?:\.\s*Director(?:a)?\s*:\s*([^.]+))?$", text, re.IGNORECASE)
    if m:
        title = m.group(1).strip(" .")
        if m.lastindex and m.lastindex >= 2 and m.group(2):
            director = m.group(2).strip(" .")
    # Director explícito
    if not director:
        m2 = re.search(r"Director(?:a)?\s*:\s*([^.]+)", text, re.IGNORECASE)
        if m2:
            director = m2.group(1).strip(" .")
    # Denominación del Proyecto
    if not title:
        m3 = re.search(r"Denominaci[oó]n.*?:\s*(.+?)(?:\.|$)", text, re.IGNORECASE)
        if m3:
            title = m3.group(1).strip()
    # PROJOVI: título
    if not title:
        m4 = re.search(r"PROJOVI\s*:\s*(.+?)(?:\.|$)", text, re.IGNORECASE)
        if m4:
            title = m4.group(1).strip()
    # Título entre comillas
    if not title:
        m5 = re.search(r"[«“\"']([^\"”»']+)[\"”»']", text)
        if m5:
            title = m5.group(1).strip()
    return title, director

def main():
    config = load_config(CONFIG_PATH)
    rows = []

    files = sorted(glob.glob(os.path.join(ACTAS_DIR, "*.pdf"))) + \
            sorted(glob.glob(os.path.join(ACTAS_DIR, "*.PDF"))) + \
            sorted(glob.glob(os.path.join(ACTAS_DIR, "*.docx"))) + \
            sorted(glob.glob(os.path.join(ACTAS_DIR, "*.DOCX")))

    for fp in files:
        if fp.lower().endswith(".pdf"):
            text = read_pdf(fp)
        else:
            text = read_docx(fp)
        if not text:
            continue

        acta = find_acta_number(text)
        fecha = find_date_text(text)
        sections = split_sections(text, config)

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
                        "Fuente_archivo": os.path.basename(fp),
                    })

    df = pd.DataFrame(rows)
    # Limpieza básica
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip()
    # Orden recomendada de columnas
    cols = [
        "Acta","Fecha","Facultad","Tipo_tema","Titulo_o_denominacion",
        "Director","Estado","Destino_publicacion","Fuente_archivo"
    ]
    df = df[cols]
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8")
    print(f"✅ Extracción completada: {OUTPUT_CSV} ({len(df)} filas)")

if __name__ == "__main__":
    main()
