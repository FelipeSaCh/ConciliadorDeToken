"""Tabla de retenciones (TABLA_DE_RETENCIONES.xlsx): conceptos con tarifa numérica; la cuenta (col. F) es opcional."""
import re

import pandas as pd

from core.config import ARCHIVO_TABLA_RETENCIONES, ARCHIVO_TABLA_RETENCIONES_LOCAL, ASSETS_DIR
from core.errors import logger
from shared.cuentas import codigo_cuenta

COLUMNAS = ("concepto", "uvt", "pesos", "tarifa", "tipo", "cuenta")


def _numero(valor):
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _tipo(tipo, concepto):
    """Tipo de NIT de la fila; sin valor en la tabla se deduce del concepto (declarantes/PJ, no declarantes/PN)."""
    tipo = "" if pd.isna(tipo) else str(tipo).strip().upper()
    if tipo:
        return tipo
    if re.search(r"no declarantes", concepto, re.I):
        return "PN"
    return "PJ" if re.search(r"\(declarantes\)|personas jur[ií]dicas", concepto, re.I) else ""


def _desde_excel(ruta):
    df = pd.read_excel(ruta, header=0, usecols="A:F", dtype=object)
    df.columns = COLUMNAS[:len(df.columns)]
    reglas = []
    for f in df.to_dict("records"):
        concepto, tarifa = str(f["concepto"]).strip(), _numero(f["tarifa"])
        if pd.isna(f["concepto"]) or tarifa is None:
            continue
        cuenta = f.get("cuenta")
        reglas.append({
            "concepto": concepto,
            "uvt": _numero(f["uvt"]) or 0.0,
            "tarifa": tarifa,
            "tipo": _tipo(f.get("tipo"), concepto),
            "cuenta": "" if pd.isna(cuenta) else codigo_cuenta(re.sub(r"\.0+$", "", str(cuenta))),
        })
    return reglas


def cargar():
    """[{concepto, uvt, tarifa, tipo ('PJ'/'PN'/''), cuenta ('' si no tiene)}] con tarifa numérica."""
    for ruta in (ARCHIVO_TABLA_RETENCIONES, ARCHIVO_TABLA_RETENCIONES_LOCAL, ASSETS_DIR / ARCHIVO_TABLA_RETENCIONES.name):
        try:
            if ruta.exists():
                return _desde_excel(ruta)
        except Exception as e:
            logger.warning(f"No se pudo leer la tabla de retenciones de {ruta}: {e}")
    return []