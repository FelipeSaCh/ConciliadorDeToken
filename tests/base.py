"""Base común de las pruebas: ruta del proyecto y parámetros de fábrica (no lee la carpeta de red ni la copia local)."""
import sys
import unittest
from pathlib import Path
from unittest import mock

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from core.config import ASSETS_DIR  # noqa: E402
from data import parametros_repo, tabla_retenciones_repo  # noqa: E402


class PruebaBase(unittest.TestCase):
    def setUp(self):
        for nombre, valor in (("cargar", parametros_repo.por_defecto),
                              ("cargar_conceptos", parametros_repo.conceptos_por_defecto),
                              ("cargar_reglas", parametros_repo.reglas_por_defecto)):
            parche = mock.patch.object(parametros_repo, nombre, side_effect=valor)
            parche.start()
            self.addCleanup(parche.stop)
        parche = mock.patch.object(tabla_retenciones_repo, "cargar",
                                   side_effect=lambda: tabla_retenciones_repo._desde_excel(ASSETS_DIR / "TABLA_DE_RETENCIONES.xlsx"))
        parche.start()
        self.addCleanup(parche.stop)