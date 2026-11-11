from copy import copy
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table

from core.config import COLUMNAS_OTROS_IMPUESTOS
from core.errors import ColumnaFaltanteError, ErrorSistema, ErrorUsuario, HojaNoEncontradaError, Reportador, logger
from shared.dataframes import a_numero, buscar_columna, es_personal, limpiar_encabezados
from shared.excel_io import escritor_excel, hojas_de, leer_hoja_por_palabra_clave
from shared.hojas import resolver_hojas
from shared.texto import limpiar_codigo, normalizar_ascii
from shared.totales import agregar_fila_totales_ws, quitar_fila_totales

COLUMNAS_CALCULADAS = ["CONCEPTO", "TERCERO", "TIPO", "TIPO-DETALLE", "BASE", "Num.Ext"]
COLUMNAS_MONTO_TOKEN = {c.upper() for c in ("IVA", "Total", "BASE", *COLUMNAS_OTROS_IMPUESTOS)}
COLUMNAS_OCULTAS = {'TIPO', 'TIPO-DETALLE', 'BASE', 'Num.Ext'}
COLUMNAS_SOBRESCRIBIR = {'CONCEPTO', 'TERCERO', 'TIPO', 'TIPO-DETALLE', 'NIT_CRUCE'}
DESCRIPCION_CALCULADAS = {
    "CONCEPTO": "TIPO + DETALLE de la hoja Contabilidad para el NIT del tercero (Emitido: NIT Receptor; Recibido: "
                "NIT Emisor). 'CONCEPTO NO CREADO' si el NIT no está en Contabilidad.",
    "TERCERO": "CREADO si el NIT del tercero está en la hoja Terceros; si no, NO CREADO.",
    "TIPO": "TIPO de la hoja Contabilidad. Si contiene PERSONAL la factura es gasto personal (fila rosada).",
    "TIPO-DETALLE": "TIPO + DETALLE de la hoja Contabilidad.",
    "BASE": "Total − IVA (0 si el total es 0).",
    "Num.Ext": "Prefijo + Folio: número con que se cruza la factura contra la contabilidad.",
}
COLUMNAS_CENTRADAS = {'PREFIJO', 'FOLIO', 'DIVISA', 'GRUPO', 'ESTADO', 'TERCERO', 'TIPO', 'NUM.EXT'}
ANCHO_MINIMO, ANCHO_MAXIMO = 9, 45
ANCHOS_FIJOS = {'CUFE/CUDE': 18}
FILAS_MUESTRA_ANCHO = 2000
COLOR_ENCABEZADO, COLOR_ENCABEZADO_CALCULADO = "1F4E78", "375623"
COLOR_PERSONAL, COLOR_NO_CREADO, COLOR_TOTALES = "FFCCFF", "FCE4D6", "DDEBF7"
COLORES_PROGRAMA = {COLOR_PERSONAL, COLOR_NO_CREADO, COLOR_TOTALES}  # rellenos que pone el formateo (no del usuario)
FUENTES_PROGRAMA = {"FF0000", "C00000"}  # nota crédito y 'NO CREADO'
FILA_ORIGEN = "_FILA_ORIGEN"  # fila de Excel de cada registro antes de formatear


def _color(color):
    rgb = getattr(color, "rgb", None) if color is not None else None
    return rgb[-6:].upper() if isinstance(rgb, str) else None


def _estilo_usuario(celda, fuente_defecto):
    """(relleno, fuente, borde) puestos por el usuario en la celda; None en lo que es del programa o por defecto."""
    relleno = celda.fill
    relleno = copy(relleno) if relleno.fill_type not in (None, "none", "gray125") and (
        _color(relleno.fgColor) not in COLORES_PROGRAMA) else None
    fuente = celda.font
    fuente = copy(fuente) if fuente != fuente_defecto and not (
        _color(fuente.color) in FUENTES_PROGRAMA and not fuente.i and not fuente.u and not fuente.strike) else None
    borde = celda.border
    borde = copy(borde) if any(getattr(borde, l).style for l in ("left", "right", "top", "bottom")) else None
    return (relleno, fuente, borde) if relleno or fuente or borde else None


class FormateadorToken(Reportador):
    COLUMNAS_REQUERIDAS_PRINCIPAL = ['NIT Emisor', 'NIT Receptor', 'Grupo', 'Prefijo', 'Folio', 'Total', 'IVA']
    COLUMNAS_REQUERIDAS_NIT = ['NIT']

    def __init__(self, file_path, sheet_names, progress_callback=None, hojas_resueltas=False):
        self.file_path = Path(file_path)
        self.sheet_names = dict(sheet_names)
        self.progress_callback = progress_callback
        self.hojas_resueltas = hojas_resueltas

    @staticmethod
    def _validar_columnas(df, columnas, nombre_hoja):
        for col in columnas:
            if col not in df.columns:
                raise ColumnaFaltanteError(nombre_hoja, col)

    def _cargar_hoja_nit(self, clave, hojas_excel):
        """Carga contabilidad/terceros: encabezado por palabra clave 'NIT'. Si falla, DataFrame vacío."""
        nombre = self.sheet_names.get(clave)
        if not nombre or nombre not in hojas_excel:
            return pd.DataFrame(columns=self.COLUMNAS_REQUERIDAS_NIT)
        try:
            df = limpiar_encabezados(leer_hoja_por_palabra_clave(self.file_path, nombre, "NIT"), mayus=True)
            self._validar_columnas(df, self.COLUMNAS_REQUERIDAS_NIT, nombre)
            df = df.dropna(subset=['NIT'])
            return df[df['NIT'].astype(str).str.strip() != '']
        except Exception as e:
            logger.warning(f"Error al procesar hoja de {clave} '{nombre}': {e}")
            return pd.DataFrame(columns=self.COLUMNAS_REQUERIDAS_NIT)

    def ejecutar(self):
        """Formatea el Token desde cero con la Contabilidad y los Terceros actuales. Se ejecuta completo en cada
        llamada (también en re-procesos): las columnas calculadas y la hoja se regeneran siempre."""
        if not self.file_path.exists():
            raise ErrorUsuario(f"El archivo no existe: {self.file_path}")
        try:
            hojas_excel = hojas_de(self.file_path)
        except Exception as e:
            raise ErrorSistema(f"No se pudo abrir el archivo Excel: {e}") from e

        if not self.hojas_resueltas:
            self._reportar("Ubicando las hojas Token, Contabilidad y Terceros en el libro...")
            self.sheet_names, mensajes = resolver_hojas(self.file_path, self.sheet_names)
            for mensaje in mensajes:
                self._reportar(mensaje)

        nombre_principal = self.sheet_names.get("principal")
        if not nombre_principal:
            raise ErrorUsuario("No se encontró la hoja del Token: selecciónala en Configurar auditoría.")
        if nombre_principal not in hojas_excel:
            raise HojaNoEncontradaError(nombre_principal, hojas_excel)

        self._reportar("Cargando hoja principal (Token)...")
        try:
            df_principal = quitar_fila_totales(limpiar_encabezados(pd.read_excel(self.file_path, sheet_name=nombre_principal)))
            df_principal = df_principal.dropna(how="all")
            self._validar_columnas(df_principal, self.COLUMNAS_REQUERIDAS_PRINCIPAL, nombre_principal)
        except ColumnaFaltanteError:
            raise
        except Exception as e:
            raise ErrorSistema(f"Error leyendo la hoja principal: {e}") from e
        columnas_originales = df_principal.columns.tolist()
        df_principal[FILA_ORIGEN] = df_principal.index + 2
        df_principal = df_principal.reset_index(drop=True)
        tabla_info, estilos, separadores = self._leer_hoja_existente(nombre_principal, df_principal[FILA_ORIGEN])

        df_conta = self._cargar_hoja_nit("contabilidad", hojas_excel)
        df_tercer = self._cargar_hoja_nit("terceros", hojas_excel)
        self._reportar(f"  Contabilidad: {len(df_conta)} NIT con concepto | Terceros: {len(df_tercer)} NIT creados")

        self._reportar("Cruzando datos entre hojas...")
        df_full, desactualizados = self._cruzar_datos(df_principal, df_conta, df_tercer)
        if desactualizados:
            self._reportar(f"  {desactualizados} fila(s) de Contabilidad tenían TIPO-DETALLE distinto de TIPO + DETALLE "
                           "(valor guardado de una versión anterior): se usó TIPO + DETALLE actual.")

        self._reportar("Calculando bases, Num.Ext y gastos personales...")
        es_personales = self._calcular_campos(df_full)
        no_creados = int((df_full['CONCEPTO'] == 'CONCEPTO NO CREADO').sum())
        self._reportar(f"  Personales: {int(es_personales.sum())} | Concepto no creado: {no_creados} | "
                       f"Tercero no creado: {int((df_full['TERCERO'] == 'NO CREADO').sum())}")

        df_export = df_full[self._orden_columnas(columnas_originales)].copy()
        df_export["_ES_PERSONAL"] = es_personales.values
        df_export[FILA_ORIGEN] = df_full[FILA_ORIGEN].values
        if estilos or separadores:
            self._reportar(f"Se conserva el orden, las filas de separación ({len(separadores)}) y los colores que ya "
                           "tenía la hoja Token.")
            df_export = df_export.sort_values(FILA_ORIGEN, kind="stable")
        else:
            self._reportar("Ordenando registros por Grupo, Tipo de Documento y Personales...")
            df_export = self._ordenar_jerarquico(df_export)
        n_registros = len(df_export)
        df_export = df_export.set_index(FILA_ORIGEN)
        if separadores:
            df_export = df_export.reindex(sorted(set(df_export.index) | set(separadores)))
            df_export["_ES_PERSONAL"] = df_export["_ES_PERSONAL"].fillna(False).astype(bool)
        filas_origen = df_export.index.tolist()
        df_export = df_export.reset_index(drop=True)
        es_personales_ordenado = df_export.pop("_ES_PERSONAL")
        estilos_filas = [estilos.get(f, {}) for f in filas_origen]

        self._reportar("Escribiendo hoja Token formateada...")
        self._escribir_excel(nombre_principal, df_export, es_personales_ordenado, tabla_info, estilos_filas)

        self._reportar("Formateo de Token completado con éxito.")
        return {
            "filas_procesadas": n_registros,
            "filas_personales": int(es_personales_ordenado.sum()),
            "escrito": True,
            "hojas": dict(self.sheet_names),
        }

    @staticmethod
    def _orden_columnas(columnas_originales):
        """Las calculadas que ya existían conservan su posición; las nuevas van al final."""
        calc_upper = [c.upper() for c in COLUMNAS_CALCULADAS]
        cols, insertadas = [], set()
        for col in columnas_originales:
            up = str(col).upper()
            if up in calc_upper:
                nombre = COLUMNAS_CALCULADAS[calc_upper.index(up)]
                cols.append(nombre)
                insertadas.add(nombre)
            else:
                cols.append(col)
        return cols + [c for c in COLUMNAS_CALCULADAS if c not in insertadas]

    @staticmethod
    def _ordenar_jerarquico(df):
        """Grupo (Emitido, Recibido, otros) → personales al final de cada grupo → tipo de documento."""
        if "Grupo" in df.columns:
            grupo = df["Grupo"].fillna("").astype(str).str.strip().str.upper()
            df["_ORDEN_GRUPO"] = np.select([grupo.str.contains("EMITIDO"), grupo.str.contains("RECIBIDO")], [0, 1], default=2)
        else:
            df["_ORDEN_GRUPO"] = 0

        col_tipo_doc = buscar_columna(df.columns, "TIPO DE DOCUMENTO", "TIPO DOCUMENTO")
        sort_cols = ["_ORDEN_GRUPO", "_ES_PERSONAL", "_ORDEN_DOC"]
        if col_tipo_doc:
            doc = df[col_tipo_doc].fillna("").astype(str).str.strip().str.upper()
            df["_ORDEN_DOC"] = np.select(
                [
                    doc.str.contains("FACTURA ELECTRÓNICA|FACTURA ELECTRONICA"),
                    doc.str.contains("DOCUMENTO EQUIVALENTE"),
                    doc.str.contains("APPLICATION RESPONSE"),
                ],
                [0, 1, 2], default=3,
            )
            sort_cols.append(col_tipo_doc)
        else:
            df["_ORDEN_DOC"] = 0

        df = df.sort_values(by=sort_cols, kind="stable").reset_index(drop=True)
        return df.drop(columns=["_ORDEN_GRUPO", "_ORDEN_DOC"])

    @staticmethod
    def _cruzar_datos(df_principal, df_conta, df_tercer):
        df_principal = df_principal.drop(
            columns=[c for c in df_principal.columns if str(c).upper() in COLUMNAS_SOBRESCRIBIR], errors='ignore'
        )

        # Emitido → cruza por NIT Receptor; Recibido u otro → NIT Emisor
        es_emitido = df_principal['Grupo'].fillna('').astype(str).str.strip().str.upper() == 'EMITIDO'
        df_principal['NIT_Cruce'] = np.where(
            es_emitido, limpiar_codigo(df_principal['NIT Receptor']), limpiar_codigo(df_principal['NIT Emisor'])
        )

        df_conta = df_conta.assign(NIT_clean=limpiar_codigo(df_conta['NIT'])).drop_duplicates(subset=['NIT_clean'])
        desactualizados = 0
        if {'TIPO', 'DETALLE'} <= set(df_conta.columns):
            # TIPO-DETALLE es la fórmula =CONCATENAR(TIPO;" ";DETALLE): se arma siempre con los datos actuales. El valor
            # guardado puede estar vacío (libro guardado con Python) o ser de una versión anterior del TIPO/DETALLE.
            compuesto = (limpiar_codigo(df_conta['TIPO']) + ' ' + limpiar_codigo(df_conta['DETALLE'])).str.strip()
            if 'TIPO-DETALLE' in df_conta.columns:
                actual = limpiar_codigo(df_conta['TIPO-DETALLE'])
                desactualizados = int(((actual != '') & (actual.str.upper() != compuesto.str.upper())).sum())
                df_conta['TIPO-DETALLE'] = compuesto.where(compuesto != '', actual)
            else:
                df_conta['TIPO-DETALLE'] = compuesto
        cols_a_traer = ['NIT_clean'] + [c for c in ('TIPO-DETALLE', 'TIPO') if c in df_conta.columns]

        df = df_principal.merge(df_conta[cols_a_traer], left_on='NIT_Cruce', right_on='NIT_clean', how='left')
        for col in ('TIPO-DETALLE', 'TIPO'):
            df[col] = df[col].fillna('') if col in df.columns else ''

        detalle = df['TIPO-DETALLE'].astype(str).str.strip()
        df['CONCEPTO'] = np.where(detalle == '', 'CONCEPTO NO CREADO', detalle)

        nits_terceros = set(limpiar_codigo(df_tercer['NIT']).unique())
        df['TERCERO'] = np.where(df['NIT_Cruce'].isin(nits_terceros) & (df['NIT_Cruce'] != ''), 'CREADO', 'NO CREADO')
        return df, desactualizados

    @staticmethod
    def _calcular_campos(df):
        df['Num.Ext'] = limpiar_codigo(df['Prefijo'], strip=False) + limpiar_codigo(df['Folio'], strip=False)
        total, iva = a_numero(df['Total']), a_numero(df['IVA'])
        df['BASE'] = np.where(total == 0, 0, total - iva)
        return es_personal(df)

    def _leer_hoja_existente(self, nombre_hoja, filas_datos):
        """(tabla de Excel, {fila: {columna: estilo del usuario}}, filas en blanco entre registros) de la hoja actual.
        Los estilos que pone el propio formateo no cuentan como del usuario."""
        tabla_info, estilos, separadores = None, {}, []
        try:
            wb = load_workbook(self.file_path)
            try:
                if nombre_hoja not in wb.sheetnames:
                    return tabla_info, estilos, separadores
                ws = wb[nombre_hoja]
                if ws.tables:
                    nombre, tbl = next(iter(ws.tables.items()))
                    tabla_info = nombre, tbl.headerRowCount, tbl.totalsRowCount, tbl.tableStyleInfo
                datos = set(filas_datos)
                if not datos:
                    return tabla_info, estilos, separadores
                encabezados = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
                fuente_defecto = wb._fonts[0]
                for fila in ws.iter_rows(min_row=2, max_row=max(datos), max_col=len(encabezados)):
                    r = fila[0].row
                    if r not in datos and any(c.value not in (None, "") for c in fila):
                        continue  # fila de totales u otra no leída como registro
                    if r not in datos:
                        separadores.append(r)
                    por_col = {n: e for n, c in zip(encabezados, fila) if n and (e := _estilo_usuario(c, fuente_defecto))}
                    if por_col:
                        estilos[r] = por_col
            finally:
                wb.close()
        except Exception as e:
            logger.warning(f"No se pudo leer el formato actual de la hoja '{nombre_hoja}': {e}")
        return tabla_info, estilos, separadores

    def _escribir_excel(self, nombre_principal, df_export, es_personales, tabla_info=None, estilos=None):
        try:
            with escritor_excel(self.file_path, hoja_activa=nombre_principal) as writer:
                df_export.to_excel(writer, index=False, sheet_name=nombre_principal)
                self._dar_formato(writer.sheets[nombre_principal], df_export, es_personales, tabla_info, estilos)
        except (ErrorUsuario, ErrorSistema):
            raise
        except Exception as e:
            raise ErrorSistema(f"Error escribiendo la hoja Token: {e}") from e

    @staticmethod
    def _ancho(df_export, col_name):
        nombre = str(col_name).strip()
        if nombre.upper() in ANCHOS_FIJOS:
            return ANCHOS_FIJOS[nombre.upper()]
        serie = df_export[col_name]
        muestra = serie.head(FILAS_MUESTRA_ANCHO)
        es_monto = nombre.upper() in COLUMNAS_MONTO_TOKEN
        datos = muestra.map(lambda v: 0 if pd.isna(v) else 10 if hasattr(v, "strftime")
                            else len(f"{float(v):,.2f}") if es_monto and isinstance(v, (int, float, np.number))
                            else len(str(v)))
        largo_datos = int(datos.quantile(0.98)) if len(datos) else 0
        if es_monto:  # la fila TOTALES suma toda la columna: debe caber sin mostrar ###
            largo_datos = max(largo_datos, len(f"{a_numero(serie).sum():,.2f}") + 1)
        largo_titulo = max((len(p) for p in nombre.split()), default=0)  # el encabezado se ajusta en varias líneas
        extra = 2 if nombre.upper() in ('CONCEPTO', 'TERCERO') else 0  # se resaltan en negrita
        return min(max(largo_datos, largo_titulo, ANCHO_MINIMO - 2) + 2 + extra, ANCHO_MAXIMO)

    def _dar_formato(self, sheet, df_export, es_personales, tabla_info, estilos=None):
        """Encabezado legible (calculadas en verde, con nota), anchos según el contenido, formatos de número y
        fecha, filtros, personales en rosado, notas crédito en rojo y 'NO CREADO' resaltado."""
        n_filas, n_cols = len(df_export), len(df_export.columns)
        ultima = get_column_letter(n_cols)
        columnas = [str(c).strip() for c in df_export.columns]
        upper = [c.upper() for c in columnas]
        calculadas = {c.upper() for c in COLUMNAS_CALCULADAS}

        borde = Border(bottom=Side(style="thin", color="FFFFFF"))
        for idx, nombre in enumerate(columnas, start=1):
            celda = sheet.cell(row=1, column=idx)
            es_calc = nombre.upper() in calculadas
            color = COLOR_ENCABEZADO_CALCULADO if es_calc else COLOR_ENCABEZADO
            celda.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            celda.font = Font(bold=True, color="FFFFFF")
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            celda.border = borde
            if es_calc:
                celda.comment = Comment(f"Columna calculada por el programa: {DESCRIPCION_CALCULADAS.get(nombre, '')}",
                                        "Conciliador")
        sheet.row_dimensions[1].height = 32
        sheet.freeze_panes = "A2"

        # Formatos por columna
        formatos = {}
        for idx, (nombre, col) in enumerate(zip(upper, df_export.columns), start=1):
            serie = df_export[col]
            if nombre in COLUMNAS_MONTO_TOKEN:
                formatos[idx] = ("#,##0.00", Alignment(horizontal="right", vertical="center"))
            elif pd.api.types.is_datetime64_any_dtype(serie):
                formatos[idx] = ("dd/mm/yyyy", Alignment(horizontal="center", vertical="center"))
            elif nombre.startswith("NIT") and pd.api.types.is_numeric_dtype(serie):
                formatos[idx] = ("0", Alignment(horizontal="center", vertical="center"))
            elif nombre in COLUMNAS_CENTRADAS:
                formatos[idx] = (None, Alignment(horizontal="center", vertical="center"))
        for idx, (formato, alineacion) in formatos.items():
            for r in range(2, n_filas + 2):
                celda = sheet.cell(row=r, column=idx)
                if formato:
                    celda.number_format = formato
                celda.alignment = alineacion

        # Colores, fuentes y bordes que ya tenía el usuario (se aplican antes de las marcas del programa)
        for r, por_col in enumerate(estilos or [], start=2):
            for idx, nombre in enumerate(columnas, start=1):
                relleno, fuente, borde = por_col.get(nombre) or (None, None, None)
                celda = sheet.cell(row=r, column=idx)
                if relleno:
                    celda.fill = relleno
                if fuente:
                    celda.font = fuente
                if borde:
                    celda.border = borde

        # Personales (fila rosada) y notas crédito (fuente roja)
        fill_personal = PatternFill(start_color=COLOR_PERSONAL, end_color=COLOR_PERSONAL, fill_type="solid")
        fuente_roja = Font(color="FF0000")
        col_tipo_doc = buscar_columna(df_export.columns, "TIPO DE DOCUMENTO", "TIPO DOCUMENTO")
        es_nota_credito = (
            df_export[col_tipo_doc].map(lambda v: "nota de credito" in normalizar_ascii(v) if pd.notna(v) and v else False)
            if col_tipo_doc else pd.Series(False, index=df_export.index)
        )
        for row_idx, (es_pers, es_nc) in enumerate(zip(es_personales, es_nota_credito), start=2):
            if not (es_pers or es_nc):
                continue
            for col_idx in range(1, n_cols + 1):
                cell = sheet.cell(row=row_idx, column=col_idx)
                if es_pers:
                    cell.fill = fill_personal
                if es_nc:
                    cell.font = fuente_roja

        # Concepto / tercero no creados: celda resaltada para que se vea qué falta crear
        fill_no_creado = PatternFill(start_color=COLOR_NO_CREADO, end_color=COLOR_NO_CREADO, fill_type="solid")
        fuente_no_creado = Font(bold=True, color="C00000")
        for nombre, valor in (("CONCEPTO", "CONCEPTO NO CREADO"), ("TERCERO", "NO CREADO")):
            if nombre not in upper:
                continue
            idx = upper.index(nombre) + 1
            for r, v in enumerate(df_export.iloc[:, idx - 1], start=2):
                if v == valor:
                    celda = sheet.cell(row=r, column=idx)
                    celda.fill = fill_no_creado
                    celda.font = fuente_no_creado

        if tabla_info:
            nombre_tabla, header_rc, totals_rc, style_info = tabla_info
            tabla = Table(displayName=nombre_tabla, ref=f"A1:{ultima}{n_filas + 1}")
            tabla.headerRowCount = header_rc
            tabla.totalsRowCount = totals_rc
            tabla.tableStyleInfo = style_info
            sheet.add_table(tabla)
        elif n_filas:
            sheet.auto_filter.ref = f"A1:{ultima}{n_filas + 1}"

        # Fila de totales de las columnas de valores (informativa: las lecturas del Token la descartan)
        montos = [i for i, c in enumerate(upper, start=1) if c in COLUMNAS_MONTO_TOKEN]
        fila_totales = agregar_fila_totales_ws(sheet, montos, 2, n_filas + 1)
        if fila_totales:
            fill_totales = PatternFill(start_color=COLOR_TOTALES, end_color=COLOR_TOTALES, fill_type="solid")
            for c in range(1, n_cols + 1):
                sheet.cell(row=fila_totales, column=c).fill = fill_totales

        # Anchos según el contenido; las calculadas técnicas quedan agrupadas (botón + para mostrarlas)
        for idx, col_name in enumerate(df_export.columns, start=1):
            dim = sheet.column_dimensions[get_column_letter(idx)]
            dim.width = self._ancho(df_export, col_name)
            if col_name in COLUMNAS_OCULTAS:
                dim.outlineLevel = 1
                dim.hidden = True