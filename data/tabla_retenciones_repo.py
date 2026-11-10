"""Tabla de retenciones (TABLA_DE_RETENCIONES.xlsx): solo los conceptos con cuenta asignada se revisan."""
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


def _desde_excel(ruta):
    df = pd.read_excel(ruta, header=0, usecols="A:F", dtype=object)
    df.columns = COLUMNAS[:len(df.columns)]
    reglas = []
    for f in df.to_dict("records"):
        cuenta = codigo_cuenta(re.sub(r"\.0+$", "", str(f.get("cuenta") or "")))
        if not cuenta:
            continue
        reglas.append({
            "concepto": str(f["concepto"]).strip(),
            "uvt": _numero(f["uvt"]) or 0.0,
            "tarifa": _numero(f["tarifa"]),
            "tipo": str(f.get("tipo") or "").strip().upper(),
            "cuenta": cuenta,
        })
    return reglas


def cargar():
    """[{concepto, uvt, tarifa, tipo ('PJ'/'PN'/''), cuenta}] de los conceptos revisados."""
    for ruta in (ARCHIVO_TABLA_RETENCIONES, ARCHIVO_TABLA_RETENCIONES_LOCAL, ASSETS_DIR / ARCHIVO_TABLA_RETENCIONES.name):
        try:
            if ruta.exists():
                return _desde_excel(ruta)
        except Exception as e:
            logger.warning(f"No se pudo leer la tabla de retenciones de {ruta}: {e}")
    return []