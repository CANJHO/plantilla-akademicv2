# src/core.py
from __future__ import annotations

import pandas as pd
from datetime import datetime
from typing import Dict, Any, List, Set, Tuple, Optional

from .excel_utils import read_consolidado, load_template_catalogs, detect_columns
from .normalizer import (
    parse_money_to_float, normalize_text_spaces,
    first_phone, first_email, empty_to_blank
)

from .correo import generar_correo_institucional
from .config import (
    CAMPUS_MAP, TIPO_ADMISION_MAP, OUTLOOK_DOMAIN,
    estado_estudiante_key,
    PROGRAMA_MASTER, es_virtual_por_pension
)

SHEET_TARGET = "SubidaEstudiantes"


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
# ✅ PROGRAMA + PLAN (TABLA MAESTRA)
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

    # fallback por contiene (variaciones leves)
    prog_norm = key[0]
    for (name_norm, vflag), val in PROGRAM_LOOKUP.items():
        if vflag == is_virtual and (prog_norm in name_norm or name_norm in prog_norm):
            return val

    return None


# =========================
# ✅ PAGADO: detectar columna real por contenido (BF)
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
# ✅ OUTLOOK: leer por columnas EXACTAS:
# - Correo: "User principal name"
# - DNI: "Fax"
# =========================
def _read_outlook(path: str) -> pd.DataFrame:
    if path.lower().endswith(".csv"):
        # autodetect separador, evita CSV con ; o ,
        return pd.read_csv(path, sep=None, engine="python")
    return pd.read_excel(path, engine="openpyxl")

def _find_col_contains(df: pd.DataFrame, needle: str) -> Optional[str]:
    needle_n = normalize_text_spaces(needle)
    for c in df.columns.astype(str):
        if needle_n in normalize_text_spaces(c):
            return c
    return None

def _extract_digits(s: Any) -> str:
    txt = "" if s is None else str(s)
    digits = "".join(ch for ch in txt if ch.isdigit())
    return digits

def _extract_dni_from_fax(fax_val: Any) -> Optional[str]:
    d = _extract_digits(fax_val)
    # DNI Perú usualmente 8 dígitos; si viniera con más, tomamos últimos 8 si tiene sentido
    if len(d) == 8:
        return d
    if len(d) > 8:
        # a veces viene como 00000000 o con prefijos
        # preferimos los últimos 8 si parecen DNI
        return d[-8:]
    return None

def _norm_person_key(nombres: str, ap_pat: str, ap_mat: str) -> str:
    # clave muy tolerante: tokens ordenados
    full = normalize_text_spaces(f"{nombres} {ap_pat} {ap_mat}").upper()
    toks = [t for t in full.split() if t]
    toks.sort()
    return " ".join(toks)

def read_outlook_indexes(outlook_path: str) -> Tuple[Set[str], Dict[str, str], Dict[str, str]]:
    """
    Retorna:
      - emails_existentes: set de correos (User principal name)
      - dni_to_email: DNI -> correo
      - email_to_personkey: correo -> clave nombre/apellidos (si hay columnas de nombre)
    """
    df = _read_outlook(outlook_path)
    if df is None or df.empty:
        return set(), {}, {}

    email_col = _find_col_contains(df, "user principal name") or _find_col_contains(df, "userprincipalname") or _find_col_contains(df, "mail") or _find_col_contains(df, "email")
    fax_col = _find_col_contains(df, "fax")

    if not email_col:
        raise ValueError("Outlook: No encontré la columna 'User principal name' (correo).")
    if not fax_col:
        # tú dijiste que está como Fax, entonces lo exigimos
        raise ValueError("Outlook: No encontré la columna 'Fax' (DNI).")

    # columnas de nombre (opcional) para fallback por nombres
    given_col = _find_col_contains(df, "given name") or _find_col_contains(df, "givenname") or _find_col_contains(df, "first name") or _find_col_contains(df, "firstname")
    sur_col = _find_col_contains(df, "surname") or _find_col_contains(df, "last name") or _find_col_contains(df, "lastname")
    disp_col = _find_col_contains(df, "display name") or _find_col_contains(df, "displayname") or _find_col_contains(df, "nombre")

    emails_existentes: Set[str] = set()
    dni_to_email: Dict[str, str] = {}
    email_to_personkey: Dict[str, str] = {}

    for _, row in df.iterrows():
        email = str(row.get(email_col, "")).strip().lower()
        if "@" not in email:
            continue

        emails_existentes.add(email)

        dni = _extract_dni_from_fax(row.get(fax_col))
        if dni:
            # si el DNI aparece varias veces, conservamos el primero (o el último, es indistinto)
            dni_to_email[dni] = email

        # armar personkey si hay info
        if disp_col and str(row.get(disp_col, "")).strip():
            # displayname en una sola columna
            pk = normalize_text_spaces(str(row.get(disp_col))).upper()
            email_to_personkey[email] = pk
        else:
            # given + surname (si existen)
            gn = str(row.get(given_col, "")).strip() if given_col else ""
            sn = str(row.get(sur_col, "")).strip() if sur_col else ""
            if gn or sn:
                pk = normalize_text_spaces(f"{gn} {sn}").upper()
                email_to_personkey[email] = pk

    return emails_existentes, dni_to_email, email_to_personkey


def _match_por_nombres(
    nombres: str, ap_pat: str, ap_mat: str,
    email: str,
    email_to_personkey: Dict[str, str]
) -> bool:
    """
    Fallback si no hay DNI:
    - compara tokens (muy tolerante) contra displayname/given+surname si existe.
    """
    pk_out = email_to_personkey.get(email, "")
    if not pk_out:
        return False

    # tokens del alumno
    alumno_key = _norm_person_key(nombres, ap_pat, ap_mat)
    out_key = normalize_text_spaces(pk_out).upper()

    # condición mínima: al menos 1 apellido y 1 nombre aparezcan
    alumno_tokens = set(alumno_key.split())
    out_tokens = set(out_key.split())

    if not alumno_tokens or not out_tokens:
        return False

    # heurística: si coinciden 2+ tokens, lo tomamos como mismo alumno
    return len(alumno_tokens & out_tokens) >= 2


# =========================
# MAIN
# =========================
def generate_outputs(consolidado_path: str, outlook_path: str, template_path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = read_consolidado(consolidado_path)
    col = detect_columns(df)

    missing = [k for k, v in col.items() if k in ["codigo", "ap_paterno", "ap_materno", "nombres"] and v is None]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias en consolidado: {missing}")

    # ✅ PAGADO: usar la columna real por contenido
    pagado_col_real = _guess_pagado_column(df, col.get("pagado"))
    df["_pagado_num"] = df[pagado_col_real].apply(parse_money_to_float)
    df = df[df["_pagado_num"] > 0].copy()

    # plantilla fija interna
    catalogs = load_template_catalogs(template_path)

    tipo_doc_df = catalogs.get("TipoDocumentos")
    sedes_df = catalogs.get("Sedes-Campus")

    if tipo_doc_df is None or sedes_df is None:
        raise ValueError("La plantilla no tiene una o más hojas requeridas: TipoDocumentos, Sedes-Campus.")

    tipo_doc_map, _, _ = _build_catalog_map_by_name(tipo_doc_df, key_col_hint="key", name_col_hint="nombre")
    sedes_map, _, sedes_name_col = _build_catalog_map_by_name(sedes_df, key_col_hint="nombre", name_col_hint="nombre")

    # ✅ Outlook index por DNI(Fax) y correo(User principal name)
    outlook_emails, outlook_dni_to_email, outlook_email_to_personkey = read_outlook_indexes(outlook_path)

    # Set de “ocupados”: todo lo que ya existe en Outlook
    existentes: Set[str] = set(outlook_emails)

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

        # ✅ PROGRAMA + PLAN por tabla maestra + VIR
        programa_raw = r[col["programa"]] if col.get("programa") else ""
        pension_raw = r[col["pension_escala"]] if col.get("pension_escala") else ""
        is_virtual = es_virtual_por_pension(pension_raw)

        resolved = _resolve_program(str(programa_raw), is_virtual)
        if not resolved:
            add_observado(base, f"No match PROGRAMA ACADÉMICO='{programa_raw}' (virtual={is_virtual}) según tabla maestra")
            continue

        program_code, plan_rel = resolved
        escuela_code = program_code  # Pxx / PxxV

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

        # Tipo documento
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

        # Documento (DNI)
        doc = str(r[col["documento"]]).strip() if col.get("documento") else ""
        if not doc:
            add_observado(base, "Documento vacío")
            continue

        dni = _extract_digits(doc)
        if len(dni) != 8:
            # si tu documento no siempre es DNI 8 dígitos, quita esta validación
            # pero para tu caso, es DNI.
            dni = dni[-8:] if len(dni) > 8 else dni

        mail_personal = first_email(r[col["mail_personal"]]) if col.get("mail_personal") else ""
        mail_personal = mail_personal if mail_personal else "-"

        tel = first_phone(r[col["telefonos"]]) if col.get("telefonos") else ""
        tel = tel.strip() if isinstance(tel, str) else str(tel).strip()
        tel = tel if tel else "-"

        direccion_raw = empty_to_blank(r[col["direccion"]]) if col.get("direccion") else ""
        direccion = direccion_raw if direccion_raw else "-"

        # Fecha nacimiento dd/mm/yyyy texto
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

        # Tipo admisión (test válido)
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
        # ✅ VALIDACIÓN CORREO EXISTENTE POR DNI (Fax)
        # =========================
        correo_existente = outlook_dni_to_email.get(dni)
        tiene_correo = "NO"
        estado_correo = "GENERAR"
        correo_inst: Optional[str] = None

        if correo_existente:
            # mismo alumno por DNI -> usar correo existente, NO generar otro
            correo_inst = correo_existente
            tiene_correo = "SI"
            estado_correo = "YA_TIENE"
        else:
            # =========================
            # ✅ Fallback por nombres/apellidos si no hay DNI en Outlook
            # =========================
            # primero generamos un candidato y, si ese candidato ya existe,
            # verificamos si corresponde al mismo alumno por nombres.
            candidato, errs = generar_correo_institucional(
                nombres=str(r[col["nombres"]]),
                ap_paterno=str(r[col["ap_paterno"]]),
                ap_materno=str(r[col["ap_materno"]]),
                dominio=OUTLOOK_DOMAIN,
                existentes=existentes
            )
            if candidato is None:
                add_observado(base, "; ".join(errs) if errs else "No se pudo generar correo")
                continue

            # Si el candidato existe en Outlook, intentar decidir si es el mismo alumno por nombres
            if candidato in outlook_emails:
                if _match_por_nombres(
                    nombres=str(r[col["nombres"]]),
                    ap_pat=str(r[col["ap_paterno"]]),
                    ap_mat=str(r[col["ap_materno"]]),
                    email=candidato,
                    email_to_personkey=outlook_email_to_personkey
                ):
                    correo_inst = candidato
                    tiene_correo = "SI"
                    estado_correo = "YA_TIENE"
                else:
                    # no es el mismo alumno -> candidato fue generado evitando existentes,
                    # así que aquí solo asignamos
                    correo_inst = candidato
            else:
                correo_inst = candidato

            # Reservar para evitar duplicados dentro del mismo lote
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
            "N°_Documento": doc,
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

            # ✅ NUEVAS COLUMNAS (lo que pediste)
            "TieneCorreoInstitucional": tiene_correo,   # SI/NO
            "EstadoCorreoInstitucional": estado_correo, # YA_TIENE / GENERAR
        }

        aprobados_rows.append(out)

    aprobados_df = pd.DataFrame(aprobados_rows)
    observados_df = pd.DataFrame(observados_rows)
    return aprobados_df, observados_df