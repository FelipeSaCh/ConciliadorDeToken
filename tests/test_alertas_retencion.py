import pandas as pd

from tests.base import PruebaBase

from core.config import GUIA_ALERTAS
from services.conciliacion.alertas import alertas_retencion, explicacion_alerta, tipo_alerta

GASTO = "143505 - MERCANCIAS"
RETE = "236540005 - Compras (2.5%)"
RETE_FUERA = "236525 - Servicios (4%)"
HONORARIOS = "511025 - ASESORIA JURIDICA"
LISTADO = {
    "860001942": "AUTORRETENEDORES",
    "830053812": "FIDEICOMISOS",
    "900927871": "solicitamos ... cincuenta por ciento (50%) ...",
    "800237815": "RGB",
}
PJ, PN = "999999999", "1017126000"


def evaluar(filas, fecha="2026-09-10", cuenta_ret=RETE, gasto=GASTO, tipo="FC1"):
    """filas: [(nit, base, retención)] -> lista de alertas por comprobante."""
    df = pd.DataFrame([{"Nit/C.C.": n, "Tipo": tipo, "Fecha": fecha, gasto: b, cuenta_ret: r} for n, b, r in filas])
    res = alertas_retencion(df, [gasto], [], [cuenta_ret], [], LISTADO, {})
    return [[t for t in a.split(" / ") if t] for a in res["ALERTA"]]


def tipos(alertas):
    return [tipo_alerta(t) for t in alertas]


class TestAlertasRetencion(PruebaBase):
    def test_sin_retencion_sobre_base(self):
        for nit in (PJ, PN):
            with self.subTest(nit=nit):
                [a] = evaluar([(nit, 5_000_000, 0)])
                self.assertEqual(tipos(a), ["Sin retención"])
                self.assertIn("Compras generales", a[0])

    def test_bajo_base_sin_alerta(self):
        self.assertEqual(evaluar([(PJ, 300_000, 0)]), [[]])

    def test_retencion_bajo_base(self):
        [a] = evaluar([(PJ, 300_000, 7_500)])
        self.assertEqual(tipos(a), ["Retención con base bajo el mínimo"])

    def test_tarifa_de_la_tabla(self):
        self.assertEqual(evaluar([(PJ, 5_000_000, 125_000)]), [[]])
        [a] = evaluar([(PJ, 5_000_000, 50_000)])
        self.assertEqual(tipos(a), ["Tarifa aplicada distinta a la tabla de retenciones"])

    def test_servicios_por_la_cuenta_de_retencion(self):
        self.assertEqual(evaluar([(PJ, 300_000, 12_000)], cuenta_ret=RETE_FUERA), [[]])
        [a] = evaluar([(PJ, 300_000, 50_000)], cuenta_ret=RETE_FUERA)
        self.assertEqual(tipos(a), ["Tarifa aplicada distinta a la tabla de retenciones"])
        [a] = evaluar([(PJ, 100_000, 4_000)], cuenta_ret=RETE_FUERA)
        self.assertEqual(tipos(a), ["Retención con base bajo el mínimo"])
        self.assertIn("Servicios generales", a[0])
        self.assertIn("$105.000", a[0])
        self.assertEqual(evaluar([(PJ, 105_000, 4_200)], cuenta_ret=RETE_FUERA), [[]])

    def test_servicio_especifico_por_el_nombre_de_la_cuenta(self):
        self.assertEqual(evaluar([(PJ, 2_000_000, 20_000)], cuenta_ret="236525010 - Transporte de carga (1%)"), [[]])

    def test_base_minima_de_compras_aproximada_a_miles(self):
        self.assertEqual(evaluar([(PJ, 523_800, 0)]), [[]])
        [a] = evaluar([(PJ, 524_000, 0)])
        self.assertIn("$524.000", a[0])
        [a] = evaluar([(PJ, 523_800, 13_095)])
        self.assertEqual(tipos(a), ["Retención con base bajo el mínimo"])

    def test_cuenta_responsable_de_iva_es_base(self):
        [a] = evaluar([(PJ, 300_000, 0)], gasto="513550001 - Transporte, fletes y acarreos responsable de iva",
                      cuenta_ret=RETE_FUERA)
        self.assertEqual(tipos(a), ["Sin retención"])
        self.assertEqual(evaluar([(PJ, 5_000_000, 125_000)], gasto="621001001 - Compra materia prima responsable de iva 19%"),
                         [[]])

    def test_honorarios_sin_base_minima(self):
        [a] = evaluar([(PJ, 5_000_000, 0)], gasto=HONORARIOS)
        self.assertEqual(tipos(a), ["Sin retención"])
        self.assertIn("Honorarios", a[0])

    def test_tarifa_y_detalle_de_la_retencion(self):
        df = pd.DataFrame([{"Nit/C.C.": PJ, "Tipo": "FC1", "Fecha": "2026-09-10", GASTO: b, RETE: r, RETE_FUERA: s}
                           for b, r, s in ((5_000_000, 125_000, 0), (5_000_000, 50_000, 100_000), (5_000_000, 0, 0))])
        res = alertas_retencion(df, [GASTO], [], [RETE, RETE_FUERA], [], LISTADO, {})
        self.assertEqual(res["% RETENCION"].tolist(), ["2,50 %", "4,00 % / 2,50 %", ""])
        self.assertEqual(res["DETALLE RETENCION"].tolist(), ["Compras (2.5%)", "Servicios (4%) / Compras (2.5%)", ""])

    def test_no_sujetos(self):
        self.assertEqual(evaluar([("860001942", 5_000_000, 0)]), [[]])
        self.assertEqual(evaluar([("830053812", 5_000_000, 0)]), [[]])
        [a] = evaluar([("860001942", 5_000_000, 125_000)])
        self.assertEqual(tipos(a), ["Retención practicada a no sujeto"])
        [a] = evaluar([(PJ, 5_000_000, 125_000)], tipo="FC REGIMEN SIMPLE")
        self.assertIn("RÉGIMEN SIMPLE", a[0])

    def test_sin_alertas_informativas(self):
        self.assertEqual(evaluar([("800237815", 5_000_000, 0)]), [[]])
        self.assertEqual(evaluar([("900927871", 5_000_000, 62_500)]), [[]])

    def test_nit_en_distintos_formatos(self):
        for nit in (860001942.0, "860001942-1", " 860001942 "):
            with self.subTest(nit=nit):
                self.assertEqual(evaluar([(nit, 5_000_000, 0)]), [[]])

    def test_toda_alerta_esta_documentada_en_la_guia(self):
        filas = [(PJ, 5_000_000, 0), (PJ, 300_000, 7_500), (PJ, 5_000_000, 50_000), ("860001942", 5_000_000, 125_000)]
        nombres_guia = {n for _, _, alertas in GUIA_ALERTAS for n, _, _ in alertas}
        for alertas in evaluar(filas):
            for texto in alertas:
                with self.subTest(texto=texto[:60]):
                    nombre = tipo_alerta(texto)
                    self.assertIn(nombre, nombres_guia)
                    self.assertTrue(explicacion_alerta(nombre)[2])