"""Conciliación completa sobre un libro real, ejecutada varias veces sobre el mismo archivo (re-procesos)."""
import shutil
import tempfile
from pathlib import Path
from unittest import mock

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from tests.base import PruebaBase
from tests.libro_prueba import HOJAS, SERIALES, crear_libro

from data.autorretenedores_repo import GestorAutorretenedores
from services.conciliacion import ConciliadorAuditoria
from services.conciliacion.preparacion import completar_columnas_token
from services.token_service import FormateadorToken
from shared.hojas import resolver_hojas

LISTADO = pd.DataFrame([{"NIT": "860001942", "COMENTARIO": "AUTORRETENEDORES"}])


class PruebaLibro(PruebaBase):
    def setUp(self):
        super().setUp()
        self.dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.dir, True)
        parche = mock.patch.object(GestorAutorretenedores, "como_dataframe", return_value=LISTADO)
        parche.start()
        self.addCleanup(parche.stop)
        self.mensajes = []

    def conciliar(self, ruta, hojas=None):
        return ConciliadorAuditoria(
            ruta, hojas or HOJAS, seriales_iva=[SERIALES["iva"]], seriales_base=[SERIALES["base"]],
            seriales_ret=[SERIALES["ret"]], filtro_grupo="Ambos", progress_callback=self.mensajes.append,
        ).ejecutar()

    @staticmethod
    def hoja(ruta, nombre):
        return pd.read_excel(ruta, sheet_name=nombre)

    @staticmethod
    def cambiar(ruta, hoja, celda, valor):
        wb = load_workbook(ruta)
        wb[hoja][celda] = valor
        wb.save(ruta)


class TestReproceso(PruebaLibro):
    def test_dos_ejecuciones_dan_el_mismo_resultado(self):
        ruta = crear_libro(self.dir / "libro.xlsx")
        self.conciliar(ruta)
        primera = {h: self.hoja(ruta, h) for h in ("DIAN VS CONT", "Token", "AUD-COMP")}
        self.conciliar(ruta)
        for nombre, df in primera.items():
            with self.subTest(hoja=nombre):
                pd.testing.assert_frame_equal(df, self.hoja(ruta, nombre))

    def test_columnas_inyectadas_no_se_duplican(self):
        ruta = crear_libro(self.dir / "libro.xlsx")
        for _ in range(3):
            self.conciliar(ruta)
        columnas = [str(c).split(".")[0] for c in self.hoja(ruta, "AUD-COMP").columns]  # pandas renombra repetidas: X.1
        for nombre in ("AUTORRETENCION", "BASE", "IVA", "ALERTA", "RETEFUENTE"):
            self.assertEqual(columnas.count(nombre), 1, nombre)

    def test_columna_inyectada_movida_por_el_usuario(self):
        ruta = crear_libro(self.dir / "libro.xlsx")
        self.conciliar(ruta)
        wb = load_workbook(ruta)
        ws = wb["AUD-COMP"]
        col = [c.value for c in ws[1]].index("ALERTA") + 1
        valores = [ws.cell(row=r, column=col).value for r in range(1, ws.max_row + 1)]
        ws.delete_cols(col)
        destino = ws.max_column + 1
        for r, v in enumerate(valores, start=1):
            ws.cell(row=r, column=destino, value=v)
        wb.save(ruta)
        self.conciliar(ruta)
        columnas = [str(c).split(".")[0] for c in self.hoja(ruta, "AUD-COMP").columns]
        self.assertEqual(columnas.count("ALERTA"), 1)

    def test_token_siempre_se_reescribe(self):
        ruta = crear_libro(self.dir / "libro.xlsx")
        self.conciliar(ruta)
        self.cambiar(ruta, "Token", "A2", "Factura electrónica")  # sin cambios de datos: igual se regenera
        self.mensajes.clear()
        self.conciliar(ruta)
        self.assertIn("Escribiendo hoja Token formateada...", self.mensajes)

    def test_cambio_de_concepto_entre_ejecuciones(self):
        ruta = crear_libro(self.dir / "libro.xlsx")
        self.conciliar(ruta)
        self.cambiar(ruta, "contabilidad", "D4", "CERTIFICADOS")
        self.conciliar(ruta)
        token = self.hoja(ruta, "Token")
        self.assertIn("GASTO CERTIFICADOS", token["CONCEPTO"].tolist())

    def test_num_ext_y_base_siempre_se_reconstruyen(self):
        df = pd.DataFrame({"Prefijo": ["FE", None], "Folio": [100.0, "0200"], "Num.Ext": ["VIEJO", ""],
                           "Total": [119000, 0], "IVA": [19000, 0], "BASE": [1, 2]})
        df = completar_columnas_token(df)
        self.assertEqual(df["Num.Ext"].tolist(), ["FE100", "0200"])
        self.assertEqual(list(df["BASE"]), [100000, 0])


class TestPersonales(PruebaLibro):
    def test_personal_desde_la_primera_ejecucion(self):
        ruta = crear_libro(self.dir / "libro.xlsx", tipos={"800333444": "PERSONAL"})
        resumen = self.conciliar(ruta)
        self.assertEqual(resumen["filas_personales"], 1)
        dvc = self.hoja(ruta, "DIAN VS CONT").fillna("")
        dian = dvc[dvc["Num.Ext"] == "FV200"]
        self.assertTrue(dian["Novedad"].str.startswith("DIAN - Gasto Personal").any())
        self.assertTrue(dian["Novedad"].str.startswith("CONT - Gasto Personal contabilizado").any())
        self.assertTrue(dian["Alertas"].str.contains("art. 107 E.T.").any())
        self.assertFalse(dvc["Alertas"].str.contains("Sin factura/documento soporte").any())

    def test_personal_marcado_en_reproceso(self):
        ruta = crear_libro(self.dir / "libro.xlsx")
        self.assertEqual(self.conciliar(ruta)["filas_personales"], 0)
        self.cambiar(ruta, "contabilidad", "C5", "PERSONAL")
        self.assertEqual(self.conciliar(ruta)["filas_personales"], 1)
        token = self.hoja(ruta, "Token")
        self.assertEqual(token.loc[token["Num.Ext"] == "FV200", "CONCEPTO"].iat[0], "PERSONAL RESTAURANTE")
        fila = token.index[token["Num.Ext"] == "FV200"][0] + 2
        relleno = load_workbook(ruta)["Token"].cell(row=fila, column=1).fill.fgColor.rgb
        self.assertTrue(str(relleno).upper().endswith("FFCCFF"))

    def test_personal_desmarcado_en_reproceso(self):
        ruta = crear_libro(self.dir / "libro.xlsx", tipos={"800333444": "PERSONAL"})
        self.conciliar(ruta)
        self.cambiar(ruta, "contabilidad", "C5", "GASTO")
        self.assertEqual(self.conciliar(ruta)["filas_personales"], 0)
        dvc = self.hoja(ruta, "DIAN VS CONT").fillna("")
        self.assertFalse(dvc["Novedad"].str.contains("Personal").any())

    def test_tipo_detalle_desactualizado_no_manda(self):
        ruta = crear_libro(self.dir / "libro.xlsx")
        self.conciliar(ruta)
        self.cambiar(ruta, "contabilidad", "C5", "PERSONAL")
        self.cambiar(ruta, "contabilidad", "E5", "GASTO RESTAURANTE")  # valor viejo escrito a mano
        self.mensajes.clear()
        self.assertEqual(self.conciliar(ruta)["filas_personales"], 1)
        self.assertTrue(any("TIPO-DETALLE distinto" in m for m in self.mensajes))


class TestHojas(PruebaLibro):
    def test_hojas_con_otro_nombre_se_encuentran(self):
        ruta = crear_libro(self.dir / "libro.xlsx", nombre_conta="Contabilidad 2026", nombre_terceros="Terceros ",
                           tipos={"800333444": "PERSONAL"})
        resumen = self.conciliar(ruta)
        self.assertEqual(resumen["filas_personales"], 1)
        self.assertTrue(any("se usa 'Contabilidad 2026'" in m for m in self.mensajes))
        token = self.hoja(ruta, "Token")
        self.assertEqual((token["TERCERO"] == "CREADO").sum(), 2)

    def test_contabilidad_por_contenido(self):
        ruta = crear_libro(self.dir / "libro.xlsx", nombre_conta="Hoja1")
        hojas, mensajes = resolver_hojas(ruta, {**HOJAS, "contabilidad": ""})
        self.assertEqual(hojas["contabilidad"], "Hoja1")
        self.assertTrue(any("NIT, TIPO y DETALLE" in m for m in mensajes))

    def test_token_por_contenido(self):
        ruta = crear_libro(self.dir / "libro.xlsx")
        wb = load_workbook(ruta)
        wb["Token"].title = "Facturas DIAN"
        wb.save(ruta)
        hojas, _ = resolver_hojas(ruta, HOJAS)
        self.assertEqual(hojas["principal"], "Facturas DIAN")

    def test_sin_contabilidad_lo_informa(self):
        ruta = crear_libro(self.dir / "libro.xlsx", nombre_conta="Otra")
        wb = load_workbook(ruta)
        for fila in wb["Otra"].iter_rows(min_row=3, max_row=3):
            fila[2].value = "CLASE"  # sin columna TIPO: no se reconoce por contenido
        wb.save(ruta)
        hojas, mensajes = resolver_hojas(ruta, HOJAS)
        self.assertEqual(hojas["contabilidad"], "")
        self.assertTrue(any("CONCEPTO NO CREADO" in m for m in mensajes))


class TestFormatoToken(PruebaLibro):
    def test_formato_legible(self):
        ruta = crear_libro(self.dir / "libro.xlsx", tipos={"800333444": "PERSONAL"})
        FormateadorToken(ruta, HOJAS).ejecutar()
        ws = load_workbook(ruta)["Token"]
        encabezados = [c.value for c in ws[1]]
        self.assertEqual(ws.freeze_panes, "A2")
        self.assertTrue(ws.auto_filter.ref)
        concepto = encabezados.index("CONCEPTO") + 1
        self.assertIsNotNone(ws.cell(row=1, column=concepto).comment)
        total = encabezados.index("Total") + 1
        self.assertEqual(ws.cell(row=2, column=total).number_format, "#,##0.00")
        letra = ws.cell(row=1, column=encabezados.index("Num.Ext") + 1).column_letter
        self.assertTrue(ws.column_dimensions[letra].hidden)
        self.assertEqual(ws.column_dimensions[letra].outlineLevel, 1)
        for c in range(1, len(encabezados) + 1):
            self.assertLessEqual(ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width, 47)
        no_creado = [r for r in range(2, ws.max_row) if ws.cell(row=r, column=concepto).value == "CONCEPTO NO CREADO"]
        self.assertTrue(no_creado)
        self.assertTrue(ws.cell(row=no_creado[0], column=concepto).font.bold)

class TestPersonalesEmitidos(PruebaLibro):
    def test_emitido_no_es_personal(self):
        ruta = crear_libro(self.dir / "libro.xlsx", tipos={"800333444": "PERSONAL"})
        wb = load_workbook(ruta)
        ws = wb["Token"]
        grupo = [c.value for c in ws[1]].index("Grupo") + 1
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=grupo, value="Emitido")
        wb.save(ruta)
        self.assertEqual(FormateadorToken(ruta, HOJAS).ejecutar()["filas_personales"], 0)


class TestFormatoUsuario(PruebaLibro):
    def test_conserva_colores_y_separadores(self):
        ruta = crear_libro(self.dir / "libro.xlsx")
        FormateadorToken(ruta, HOJAS).ejecutar()
        wb = load_workbook(ruta)
        ws = wb["Token"]
        orden = [ws.cell(row=r, column=3).value for r in range(2, 6)]
        ws.insert_rows(4)
        for c in range(1, ws.max_column + 1):
            ws.cell(row=2, column=c).fill = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
        ws.cell(row=3, column=3).font = Font(bold=True)
        wb.save(ruta)
        for _ in range(2):
            self.conciliar(ruta)
            ws = load_workbook(ruta)["Token"]
            self.assertEqual([ws.cell(row=r, column=3).value for r in (2, 3, 5, 6)], orden)
            self.assertTrue(all(ws.cell(row=4, column=c).value is None for c in range(1, ws.max_column + 1)))
            self.assertTrue(str(ws.cell(row=2, column=1).fill.fgColor.rgb).endswith("00B0F0"))
            self.assertTrue(ws.cell(row=3, column=3).font.b)
            self.assertIsNone(ws.cell(row=3, column=1).fill.fill_type)

    def test_sin_formato_del_usuario_se_ordena(self):
        ruta = crear_libro(self.dir / "libro.xlsx")
        FormateadorToken(ruta, HOJAS).ejecutar()
        self.mensajes.clear()
        self.conciliar(ruta)
        self.assertIn("Ordenando registros por Grupo, Tipo de Documento y Personales...", self.mensajes)