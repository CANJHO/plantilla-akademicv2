# src/correo.py
from dataclasses import dataclass
from typing import List, Set, Tuple, Optional

from .normalizer import normalize_text_no_spaces, normalize_text_spaces

@dataclass
class NombreParts:
    nombres: List[str]
    ap_paterno: str
    ap_materno: str

def split_names(nombres: str) -> List[str]:
    s = normalize_text_spaces(nombres)
    if not s:
        return []
    return s.split()

def normalize_surname_compuesto(apellido: str) -> str:
    # según tu regla: viene completo en AP. PATERNO, solo quitamos espacios y normalizamos
    # "De la Cruz" -> "delacruz"
    return normalize_text_no_spaces(apellido)

def build_email_local(nombres: List[str], ap_pat: str, ap_mat: str, mode: int) -> str:
    """
    Orden de intentos:
    1) primer_nombre.apellido_compuesto
    2) segundo_nombre.apellido_compuesto
    3) primer_nombre.segundo_apellido
    4) primer_nombre.apellido_compuesto+segundo_apellido
    5) dos_nombres.apellido_compuesto
    """
    ap_pat_n = normalize_surname_compuesto(ap_pat)
    ap_mat_n = normalize_text_no_spaces(ap_mat)

    n1 = normalize_text_no_spaces(nombres[0]) if len(nombres) >= 1 else ""
    n2 = normalize_text_no_spaces(nombres[1]) if len(nombres) >= 2 else ""
    n12 = (normalize_text_no_spaces(nombres[0] + nombres[1]) if len(nombres) >= 2 else "")

    if mode == 1:
        return f"{n1}.{ap_pat_n}"
    if mode == 2:
        return f"{n2}.{ap_pat_n}" if n2 else f"{n1}.{ap_pat_n}"
    if mode == 3:
        return f"{n1}.{ap_mat_n}" if ap_mat_n else f"{n1}.{ap_pat_n}"
    if mode == 4:
        combo = f"{ap_pat_n}{ap_mat_n}" if ap_mat_n else ap_pat_n
        return f"{n1}.{combo}"
    if mode == 5:
        return f"{n12}.{ap_pat_n}" if n12 else f"{n1}.{ap_pat_n}"
    return f"{n1}.{ap_pat_n}"

def generar_correo_institucional(nombres: str, ap_paterno: str, ap_materno: str, dominio: str, existentes: Set[str]) -> Tuple[Optional[str], List[str]]:
    intents = []
    nombres_list = split_names(nombres)
    if not nombres_list or not ap_paterno:
        return None, ["Faltan nombres/apellido paterno"]

    for mode in [1, 2, 3, 4, 5]:
        local = build_email_local(nombres_list, ap_paterno, ap_materno, mode)
        correo = f"{local}@{dominio}".lower()
        intents.append(correo)
        if correo not in existentes:
            return correo, []

    return None, [f"Correo duplicado (probados: {', '.join(intents)})"]