# src/core.py
from __future__ import annotations

import re
import pandas as pd
from datetime import datetime
from typing import Dict, Any, List, Set, Tuple, Optional

from .excel_utils import read_consolidado, load_template_catalogs, detect_columns
from .normalizer import (
    parse_money_to_float,
    normalize_text_spaces,
    first_phone,
    doc_key8_for_match,
    person_key_for_match,
)
from .correo import generar_correo_institucional
from .config import (
    CAMPUS_MAP, TIPO_ADMISION_MAP, OUTLOOK_DOMAIN,
    estado_estudiante_key,
    PROGRAMA_MASTER, es_virtual_por_pension
)

SHEET_TARGET = "SubidaEstudiantes"


# =========================
# Helpers locales
# =========================
def _only_digits(val: Any) -> str:
    s = "" if val is None else str(val)
    return "".join(ch for ch in s if ch.isdigit())

def _is_blank(val: Any) -> bool:
    if val is None:
        return True
    s = str(val).strip()
    if s == "":
        return True
    if s.lower() in {"nan", "none", "null"}:
        return True
    return False

def _as_dash(val: Any) -> str:
    """Convierte vacío/NaN/'nan' a '-'."""
    return "-" if _is_blank(val) else str(val).strip()

_EMAIL_RE = re.compile(r"[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}", re.IGNORECASE)

def _first_email(val: Any) -> str:
    """
    Extrae el primer email válido desde un campo sucio:
      ', kim@gmail.com' -> 'kim@gmail.com'
      'a@x.com, b@y.com' -> 'a@x.com'
      'a@x.com b@y.com' -> 'a@x.com'
    """
    if _is_blank(val):
        return ""
    s = str(val).strip().lower()

    # normalizar separadores
    for sep in [",", ";", "|", "\n", "\t"]:
        s = s.replace(sep, " ")
    s = re.sub(r"\s+", " ", s).strip()

    m = _EMAIL_RE.search(s)
    if not m:
        return ""
    return m.group(0).strip(" ,.;:")


# =========================
# Catalog map
# =========================
def _build_catalog_map_by_name(
    df: pd.DataFrame,
    key_col_hint: str,
    name_col_hint: str
) -> Tuple[Dict[str, Any], str, str]:
    cols_norm = {normalize_text_spaces(c): c for c in df.columns.astype(str)}

    def find_col(hints: List[str]) -> Optional[str]:
        for h in hints:
            hn = normalize_text_spaces(h)
            for cn, orig in cols_norm.items():
                if hn in cn:
                    return orig
        return None

    key_col = find_col([key_col_hint, "key", "codigo", "código", "id"])
    name_col = find_col([name_col_hint, "nombre", "descripcion", "descripción"])

    if not key_col or not name_col:
        raise ValueError(f"No se pudo detectar columnas de catálogo en hoja. Columnas: {list(df.columns)}")

    mapping = {}
    for _, r in df.iterrows():
        name = normalize_text_spaces(r.get(name_col, ""))
        if name:
            mapping[name] = r.get(key_col)
    return mapping, key_col, name_col


# =========================
# PROGRAMA + PLAN (TABLA MAESTRA)
# =========================
def _build_program_lookup() -> Dict[Tuple[str, bool], Tuple[str, str]]:
    lookup: Dict[Tuple[str, bool], Tuple[str, str]] = {}
    for nombre, is_virtual, codigo, plan in PROGRAMA_MASTER:
        key = (normalize_text_spaces(nombre).upper(), is_virtual)
        lookup[key] = (codigo, plan)
    return lookup

PROGRAM_LOOKUP = _build_program_lookup()

def _resolve_program(programa_academico: str, is_virtual: bool) -> Optional[Tuple[str, str]]:
    key = (normalize_text_spaces(programa_academico).upper(), is_virtual)
    if key in PROGRAM_LOOKUP:
        return PROGRAM_LOOKUP[key]

    prog_norm = key[0]
    for (name_norm, vflag), val in PROGRAM_LOOKUP.items():
        if vflag == is_virtual and (prog_norm in name_norm or name_norm in prog_norm):
            return val

    return None


# =========================
# PAGADO: detectar columna real por contenido (BF)
# =========================
def _guess_pagado_column(df: pd.DataFrame, detected_col: Optional[str]) -> str:
    candidates: List[str] = []

    if detected_col and detected_col in df.columns:
        candidates.append(detected_col)

    for c in df.columns.astype(str):
        cn = normalize_text_spaces(c)
        if "pagado" in cn and c not in candidates:
            candidates.append(c)

    if not candidates:
        for c in df.columns.astype(str):
            s = df[c].astype(str).head(400).str.upper()
            if (s.str.contains("S/").sum() >= 30) or (s.str.contains("175").sum() >= 30):
                candidates.append(c)

    if not candidates:
        raise ValueError("No se pudo localizar ninguna columna de PAGADO en el consolidado.")

    best = candidates[0]
    best_score = -1
    for c in candidates:
        vals = df[c].apply(parse_money_to_float)
        score = int((vals > 0).sum())
        if score > best_score:
            best_score = score
            best = c

    return best


# =========================
# OUTLOOK
# =========================
def _read_outlook(path: str) -> pd.DataFrame:
    if path.lower().endswith(".csv"):
        return pd.read_csv(path, sep=None, engine="python")
    return pd.read_excel(path, engine="openpyxl")

def _find_col_contains(df: pd.DataFrame, needle: str) -> Optional[str]:
    needle_n = normalize_text_spaces(needle)
    for c in df.columns.astype(str):
        if needle_n in normalize_text_spaces(c):
            return c
    return None


def read_outlook_indexes(outlook_path: str) -> Tuple[Set[str], Dict[str, str], Dict[str, str]]:
    """
    Retorna:
      - emails_existentes: set de correos (User principal name)
      - doc8_to_email: doc_key8 (desde Fax) -> correo
      - personkey_to_email: clave por apellidos+nombres -> correo (fallback)
    """
    df = _read_outlook(outlook_path)
    if df is None or df.empty:
        return set(), {}, {}

    email_col = (
        _find_col_contains(df, "user principal name")
        or _find_col_contains(df, "userprincipalname")
        or _find_col_contains(df, "mail")
        or _find_col_contains(df, "email")
    )
    fax_col = _find_col_contains(df, "fax")

    # Columnas de nombres (según export)
    display_col = _find_col_contains(df, "display name") or _find_col_contains(df, "displayname")
    given_col   = _find_col_contains(df, "given name") or _find_col_contains(df, "givenname") or _find_col_contains(df, "first name")
    surname_col = _find_col_contains(df, "surname") or _find_col_contains(df, "last name") or _find_col_contains(df, "lastname")

    if not email_col:
        raise ValueError("Outlook: No encontré la columna 'User principal name' (correo).")
    if not fax_col:
        raise ValueError("Outlook: No encontré la columna 'Fax' (documento).")

    emails_existentes: Set[str] = set()
    doc8_to_email: Dict[str, str] = {}
    personkey_to_email: Dict[str, str] = {}

    for _, row in df.iterrows():
        email = str(row.get(email_col, "")).strip().lower()
        if "@" not in email:
            continue

        emails_existentes.add(email)

        # 1) Index por documento (FAX -> doc_key8)
        fax_digits = _only_digits(row.get(fax_col))
        if fax_digits:
            k8 = doc_key8_for_match(fax_digits)  # SIEMPRE 8
            if k8 and k8 not in doc8_to_email:
                doc8_to_email[k8] = email

        # 2) Index por nombre/apellidos (fallback)
        display = str(row.get(display_col, "")).strip() if display_col else ""
        given = str(row.get(given_col, "")).strip() if given_col else ""
        surname = str(row.get(surname_col, "")).strip() if surname_col else ""

        ap_pat = ""
        ap_mat = ""
        nombres = ""

        if display:
            # Intento 1: "AP1 AP2, NOMBRES"
            if "," in display:
                left, right = display.split(",", 1)
                ap_parts = [p for p in left.strip().split() if p]
                nm_parts = [p for p in right.strip().split() if p]
                ap_pat = ap_parts[0] if len(ap_parts) >= 1 else ""
                ap_mat = ap_parts[1] if len(ap_parts) >= 2 else ""
                nombres = " ".join(nm_parts) if nm_parts else ""
            else:
                # Intento 2: "AP1 AP2 NOMBRES..."
                parts = [p for p in display.strip().split() if p]
                if len(parts) >= 3:
                    ap_pat = parts[0]
                    ap_mat = parts[1]
                    nombres = " ".join(parts[2:])

        # Fallback: given + surname
        if (not ap_pat and not ap_mat and not nombres) and (given or surname):
            ap_parts = [p for p in surname.strip().split() if p]
            ap_pat = ap_parts[0] if len(ap_parts) >= 1 else surname
            ap_mat = ap_parts[1] if len(ap_parts) >= 2 else ""
            nombres = given

        if ap_pat or ap_mat or nombres:
            pk = person_key_for_match(ap_pat, ap_mat, nombres)
            if pk and pk not in personkey_to_email:
                personkey_to_email[pk] = email

    return emails_existentes, doc8_to_email, personkey_to_email


# =========================
# HISTORIAL (opcional) para no repetir en otra ejecución
# =========================
def read_history_doc8(history_path: Optional[str]) -> Set[str]:
    """
    Lee un Excel/CSV histórico con una columna llamada:
    - doc8
    o
    - N°_Documento
    o
    - Documento
    y devuelve set(doc_key8)
    """
    if not history_path:
        return set()

    try:
        if history_path.lower().endswith(".csv"):
            hdf = pd.read_csv(history_path, sep=None, engine="python")
        else:
            hdf = pd.read_excel(history_path, engine="openpyxl")
    except:
        return set()

    if hdf is None or hdf.empty:
        return set()

    cols = {normalize_text_spaces(c): c for c in hdf.columns.astype(str)}
    pick = cols.get("doc8") or cols.get("n documento") or cols.get("documento") or cols.get("n° documento") or None
    if not pick:
        return set()

    out: Set[str] = set()
    for v in hdf[pick].dropna().astype(str):
        k8 = doc_key8_for_match(v)
        if k8:
            out.add(k8)
    return out


# =========================
# MAIN
# =========================
def generate_outputs(
    consolidado_path: str,
    outlook_path: str,
    template_path: str,
    history_path: Optional[str] = None,   # ✅ opcional
) -> Tuple[pd.DataFrame, pd.DataFrame]:

    df = read_consolidado(consolidado_path)
    col = detect_columns(df)

    missing = [k for k, v in col.items() if k in ["codigo", "ap_paterno", "ap_materno", "nombres"] and v is None]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias en consolidado: {missing}")

    # PAGADO
    pagado_col_real = _guess_pagado_column(df, col.get("pagado"))
    df["_pagado_num"] = df[pagado_col_real].apply(parse_money_to_float)
    df = df[df["_pagado_num"] > 0].copy()

    catalogs = load_template_catalogs(template_path)
    tipo_doc_df = catalogs.get("TipoDocumentos")
    sedes_df = catalogs.get("Sedes-Campus")

    if tipo_doc_df is None or sedes_df is None:
        raise ValueError("La plantilla no tiene una o más hojas requeridas: TipoDocumentos, Sedes-Campus.")

    tipo_doc_map, _, _ = _build_catalog_map_by_name(tipo_doc_df, key_col_hint="key", name_col_hint="nombre")
    sedes_map, _, sedes_name_col = _build_catalog_map_by_name(sedes_df, key_col_hint="nombre", name_col_hint="nombre")

    # Outlook index
    outlook_emails, outlook_doc8_to_email, outlook_personkey_to_email = read_outlook_indexes(outlook_path)
    existentes: Set[str] = set(outlook_emails)

    # Historial opcional (para no repetir en otra ejecución)
    history_doc8 = read_history_doc8(history_path)

    aprobados_rows: List[Dict[str, Any]] = []
    observados_rows: List[Dict[str, Any]] = []

    def add_observado(base_row: Dict[str, Any], motivo: str):
        rr = dict(base_row)
        rr["Motivo"] = motivo
        observados_rows.append(rr)

    for _, r in df.iterrows():
        base = {
            "Estudiante_Código": r[col["codigo"]],
            "Paterno": r[col["ap_paterno"]],
            "Materno": r[col["ap_materno"]],
            "Nombres": r[col["nombres"]],
        }

        # PROGRAMA + PLAN
        programa_raw = r[col["programa"]] if col.get("programa") else ""
        pension_raw = r[col["pension_escala"]] if col.get("pension_escala") else ""
        is_virtual = es_virtual_por_pension(pension_raw)

        resolved = _resolve_program(str(programa_raw), is_virtual)
        if not resolved:
            add_observado(base, f"No match PROGRAMA ACADÉMICO='{programa_raw}' (virtual={is_virtual}) según tabla maestra")
            continue

        program_code, plan_rel = resolved
        escuela_code = program_code

        # Campus
        campus_raw = str(r[col["sede_filial"]]).strip().upper() if col.get("sede_filial") else ""
        campus_mapped = CAMPUS_MAP.get(campus_raw, campus_raw)
        campus_ok = normalize_text_spaces(campus_mapped) in sedes_map or any(
            normalize_text_spaces(campus_mapped) == normalize_text_spaces(x)
            for x in sedes_df[sedes_name_col].dropna().astype(str).tolist()
        )
        if not campus_ok:
            add_observado(base, f"Campus no válido: origen='{campus_raw}' mapeado='{campus_mapped}'")
            continue

        # Tipo documento (Key)
        tipo_doc_raw = str(r[col["tipo_doc"]]).strip() if col.get("tipo_doc") else ""
        tipo_doc_norm = normalize_text_spaces(tipo_doc_raw)
        tipo_doc_key = None
        if tipo_doc_norm:
            for name, key in tipo_doc_map.items():
                if tipo_doc_norm == name or tipo_doc_norm in name:
                    tipo_doc_key = key
                    break
        if tipo_doc_key is None:
            add_observado(base, f"No match TipoDocumento para '{tipo_doc_raw}'")
            continue

        # Documento: DNI=8 / CE=9, MATCH=8 (desde doc_key8_for_match)
        doc_raw = str(r[col["documento"]]).strip() if col.get("documento") else ""
        if _is_blank(doc_raw):
            add_observado(base, "Documento vacío")
            continue

        doc_digits = _only_digits(doc_raw)
        if not doc_digits:
            add_observado(base, f"Documento inválido: '{doc_raw}'")
            continue

        is_ce = ("carnet" in tipo_doc_norm) or ("extran" in tipo_doc_norm) or (tipo_doc_norm == "ce")
        target_len = 9 if is_ce else 8

        # Si viene con más dígitos de lo permitido -> OBSERVADO
        if len(doc_digits) > target_len:
            add_observado(base, f"Documento con longitud inválida para {'CE' if is_ce else 'DNI'}: '{doc_digits}'")
            continue

        doc_out = doc_digits.zfill(target_len)
        doc_key8 = doc_key8_for_match(doc_out)  # MATCH SIEMPRE 8

        # Historial opcional: no repetir
        if doc_key8 and doc_key8 in history_doc8:
            add_observado(base, "Ya fue generado en un proceso anterior (historial).")
            continue

        # Correo personal, Teléfono, Dirección
        mail_personal = _first_email(r[col["mail_personal"]]) if col.get("mail_personal") else ""
        mail_personal = mail_personal if mail_personal else "-"

        tel = first_phone(r[col["telefonos"]]) if col.get("telefonos") else ""
        tel = "-" if _is_blank(tel) else str(tel).strip()

        direccion = _as_dash(r[col["direccion"]]) if col.get("direccion") else "-"

        # Fecha nacimiento
        fecha_txt = ""
        if col.get("fecha_nac"):
            v = r[col["fecha_nac"]]
            if pd.isna(v):
                fecha_txt = ""
            elif isinstance(v, (datetime, pd.Timestamp)):
                fecha_txt = v.strftime("%d/%m/%Y")
            else:
                parsed = pd.to_datetime(v, dayfirst=True, errors="coerce")
                fecha_txt = "" if pd.isna(parsed) else parsed.strftime("%d/%m/%Y")

        # Sexo
        sexo_raw = str(r[col["sexo"]]).strip().upper() if col.get("sexo") else ""
        sexo = "M" if "MASCULINO" in sexo_raw else ("F" if "FEMENINO" in sexo_raw else "")

        periodo_ing = str(r[col["inicio"]]).strip() if col.get("inicio") else ""

        # Tipo admisión
        modalidad_raw = str(r[col["modalidad"]]).strip() if col.get("modalidad") else ""
        modalidad_norm = modalidad_raw.strip().upper()
        adm_id = None
        if modalidad_raw:
            for k, v_id in TIPO_ADMISION_MAP.items():
                if k.strip().upper() == modalidad_norm:
                    adm_id = v_id
                    break
        if adm_id is None:
            add_observado(base, f"Modalidad sin mapeo: '{modalidad_raw}'")
            continue

        # Ciclo y estado
        try:
            ciclo_int = int(r[col["ciclo"]]) if col.get("ciclo") and pd.notna(r[col["ciclo"]]) else 0
        except:
            ciclo_int = 0
        estado_key = estado_estudiante_key(ciclo_int if ciclo_int else 0)

        # =========================
        # Correo institucional:
        # 1) Match por doc_key8 (Fax)
        # 2) Si no coincide, match por nombre+apellidos (fallback)
        # =========================
        correo_existente = outlook_doc8_to_email.get(doc_key8)

        if not correo_existente:
            pk = person_key_for_match(
                str(r[col["ap_paterno"]]),
                str(r[col["ap_materno"]]),
                str(r[col["nombres"]]),
            )
            correo_existente = outlook_personkey_to_email.get(pk)

        tiene_correo = "NO"
        estado_correo = "GENERAR"
        correo_inst: Optional[str] = None

        if correo_existente:
            correo_inst = correo_existente
            tiene_correo = "SI"
            estado_correo = "YA_TIENE"
        else:
            correo_inst, errs = generar_correo_institucional(
                nombres=str(r[col["nombres"]]),
                ap_paterno=str(r[col["ap_paterno"]]),
                ap_materno=str(r[col["ap_materno"]]),
                dominio=OUTLOOK_DOMAIN,
                existentes=existentes
            )
            if correo_inst is None:
                add_observado(base, "; ".join(errs) if errs else "No se pudo generar correo")
                continue

            existentes.add(correo_inst)

        out = {
            "Estudiante_Código": r[col["codigo"]],
            "Paterno": r[col["ap_paterno"]],
            "Materno": r[col["ap_materno"]],
            "Nombres": r[col["nombres"]],
            "EscuelaProfesional_Código": escuela_code,
            "Ciclo": ciclo_int,
            "PlanCurricular_RelationId": plan_rel,
            "ProgramaAcadémico_RelationId(-)": escuela_code,
            "Campus_Nombre": campus_mapped,
            "TipoDocument_Key": tipo_doc_key,
            "N°_Documento": doc_out,
            "CorreoInstitucional": correo_inst,
            "CorreoPersonal(-)": mail_personal,
            "Telefono": tel,
            "Departamento_Nombre(-)": "-",
            "Provincia_Nombre(-)": "-",
            "Distrito_Nombre(-)": "-",
            "Dirección": direccion,
            "F.Nacimiento (dd/mm/yyyy)": fecha_txt,
            "Sexo (M o F)": sexo,
            "PeriodoDeIngreso_RelationId": periodo_ing,
            "TipoDeAdmission_RelationId": adm_id,
            "EstadoEstudiante_Key": estado_key,

            "TieneCorreoInstitucional": tiene_correo,
            "EstadoCorreoInstitucional": estado_correo,
        }

        aprobados_rows.append(out)

    aprobados_df = pd.DataFrame(aprobados_rows)
    observados_df = pd.DataFrame(observados_rows)
    return aprobados_df, observados_df