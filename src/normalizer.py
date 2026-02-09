# src/normalizer.py
import re
import unicodedata


def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join([c for c in s if not unicodedata.combining(c)])


def normalize_text_no_spaces(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = strip_accents(s)
    s = s.replace("ñ", "n")
    s = re.sub(r"[^a-z0-9\s]", "", s)
    s = re.sub(r"\s+", "", s)
    return s


def normalize_text_spaces(s: str) -> str:
    """minúscula/sin tildes/ñ->n, conservando 1 espacio."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = strip_accents(s)
    s = s.replace("ñ", "n")
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_money_to_float(val) -> float:
    if val is None:
        return 0.0
    s = str(val).strip().lower()
    s = s.replace("s/", "").replace("s\\", "").replace("soles", "")
    s = s.replace(",", "")
    m = re.search(r"(\d+(\.\d+)?)", s)
    if not m:
        return 0.0
    try:
        return float(m.group(1))
    except:
        return 0.0


def first_phone(telefonos: str) -> str:
    if telefonos is None:
        return ""
    s = str(telefonos).strip()
    parts = re.split(r"[;,/|]| y | - ", s)
    parts = [p.strip() for p in parts if p.strip()]
    return parts[0] if parts else s


# =========================
# Email
# =========================
_EMAIL_RE = re.compile(r"[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}", re.IGNORECASE)

def first_email(value) -> str:
    """
    Extrae el primer email válido desde un campo sucio:
    - si hay varios separados por coma/; espacio salto de línea -> toma el primero
    - devuelve en minúscula
    """
    if value is None:
        return ""

    s = str(value).strip().lower()
    if not s or s in ("nan", "none", "null"):
        return ""

    s = s.replace("mailto:", " ").strip()

    # normalizar separadores
    for sep in [",", ";", "|", "\n", "\t"]:
        s = s.replace(sep, " ")
    s = re.sub(r"\s+", " ", s).strip()

    m = _EMAIL_RE.search(s)
    if not m:
        return ""

    return m.group(0).strip(" ,.;:")


# =========================
# Empty handling
# =========================
def empty_to_blank(v) -> str:
    """Convierte valores vacíos a '' (None/NaN/'nan'/etc)."""
    if v is None:
        return ""
    try:
        if isinstance(v, float) and v != v:  # NaN
            return ""
    except:
        pass
    s = str(v).strip()
    if not s:
        return ""
    if s.lower() in ("nan", "none", "null", "n/a"):
        return ""
    return s


# =========================
# Document normalization
# =========================
def normalize_document_number(doc: str, *, is_ce: bool) -> str:
    """
    Normaliza el documento para ESCRIBIR en plantilla:
    - DNI: 8 dígitos (pad izquierda)
    - CE: 9 dígitos (pad izquierda)
    """
    if doc is None:
        return ""
    s = re.sub(r"\D", "", str(doc).strip())
    if not s:
        return ""
    if is_ce:
        return s.zfill(9)[:9]
    return s.zfill(8)[:8]


def doc_key8_for_match(doc: str) -> str:
    """
    Clave para MATCH: SIEMPRE 8 dígitos.
    - Extrae dígitos.
    - Si son 9 (CE), toma los ÚLTIMOS 8.
    """
    if doc is None:
        return ""
    s = re.sub(r"\D", "", str(doc).strip())
    if not s:
        return ""
    if len(s) >= 8:
        return s[-8:]
    return s.zfill(8)


def person_key_for_match(ap_paterno: str, ap_materno: str, nombres: str) -> str:
    """Clave fallback por nombres/apellidos (sin espacios/símbolos)."""
    return normalize_text_no_spaces(f"{ap_paterno} {ap_materno} {nombres}")