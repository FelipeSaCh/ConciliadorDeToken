"""Acta (memorando) en PDF con las inconsistencias de la conciliación DIAN vs Contabilidad."""
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from services.conciliacion.alertas import COL_ALERTAS, explicacion_alerta, tipo_alerta
from shared.formato import escapar_xml, nit_texto
from shared.texto import normalizar_num_ext

ORIGENES = {1: "DIAN", 2: "CONTAB.", 3: "DEVOL."}
SIN_SOPORTE = "Sin factura/documento soporte electrónico en DIAN (art. 771-2 E.T.)"
NOVEDADES_SIN_PAREJA = ['DIAN - Sin Pareja en Contabilidad', 'CONT - Sin Pareja en DIAN', 'DEVOL - Sin Pareja en DIAN', '']

ESTILO_BASE_TABLA = [
    ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 8),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
]


def _moneda(valor):
    return f"$ {pd.to_numeric(valor, errors='coerce'):,.2f}"


def _origen(prioridad, con_devol=False):
    if con_devol:
        return ORIGENES.get(prioridad, "DEVOL.")
    return "DIAN" if prioridad == 1 else "CONTAB."


HOJA_RESULTADO = 'DIAN VS CONT'
MARCA_TOTALES = 'TOTAL CONCILIADOS'


def _texto(serie):
    return serie.fillna('').astype(str).str.strip().replace({'nan': '', 'None': ''})


def _desde_dian_vs_cont(df):
    """Reconstruye el resultado de la conciliación desde la hoja DIAN VS CONT ya escrita en el libro."""
    from services.conciliacion.emparejamiento import _marcar_diferencias_valores

    df = df.reset_index(drop=True)
    es_totales = df.apply(lambda f: f.astype(str).str.upper().str.contains(MARCA_TOTALES, regex=False).any(), axis=1)
    fin_conciliados = int(es_totales.idxmax()) if es_totales.any() else 0
    vacia = df.isna().all(axis=1)
    en_bloque = pd.Series(np.arange(len(df)) < fin_conciliados, index=df.index)

    novedad = _texto(df['Novedad']) if 'Novedad' in df.columns else pd.Series('', index=df.index)
    es_dian = pd.Series(False, index=df.index)
    for col in ('Fecha Emisión', 'Total', 'Grupo', 'Estado'):
        if col in df.columns:
            es_dian |= _texto(df[col]) != ''
    es_dian |= novedad.str.startswith('DIAN')
    es_devol = ~es_dian & novedad.str.startswith('DEVOL')
    df['Prioridad_Fila'] = np.where(es_dian, 1, np.where(es_devol, 3, 2))

    # Parejas: en el bloque de conciliados cada fila DIAN abre un grupo; las filas vacías separan grupos
    corte = (es_dian | vacia).cumsum()
    num_ext = normalizar_num_ext(_texto(df['Num.Ext'])) if 'Num.Ext' in df.columns else pd.Series('', index=df.index)
    grupo = np.where(en_bloque, 'P' + corte.astype(str), 'S' + df.index.astype(str))
    df['Num.Ext_Clean'] = num_ext
    df['Num.Ext_Original'] = _texto(df['Num.Ext']) if 'Num.Ext' in df.columns else ''
    df['_Grupo'] = grupo
    df = df[~vacia & ~es_totales].copy()

    etiqueta = df.groupby('_Grupo')['Num.Ext_Clean'].transform(lambda s: next((v for v in s if v), ''))
    df['Num.Ext_Clean'] = np.where(df['_Grupo'].str.startswith('P'), etiqueta.where(etiqueta != '', df['_Grupo']), df['Num.Ext_Clean'])
    df['Conteo_Pareja'] = np.where(df['_Grupo'].str.startswith('P'), df.groupby('_Grupo')['_Grupo'].transform('count'), 1)
    df['Novedad'] = _texto(df['Novedad']) if 'Novedad' in df.columns else ''
    for col in ('BASE', 'IVA', 'NIT Emisor', 'Nombre Emisor', 'Tipo de documento', 'CUFE/CUDE', COL_ALERTAS):
        if col not in df.columns:
            df[col] = ''

    claves = df['Num.Ext_Clean'].copy()
    df['Num.Ext_Clean'] = df['_Grupo']  # clave única de pareja para calcular diferencias
    _marcar_diferencias_valores(df, df['Conteo_Pareja'] >= 2)
    df['Num.Ext_Clean'] = claves
    df['Diff_NIT'] = df['Novedad'].str.contains('No coincide tercero', regex=False)
    return df.drop(columns=['_Grupo'])


def _desde_hoja_auditoria(df, columna_alerta):
    """Solo alertas: AUD-COMP / MOVS_AUD con la columna ALERTA inyectada por la conciliación."""
    col = {str(c).strip().upper(): c for c in df.columns}
    tomar = lambda *nombres: df[next((col[n] for n in nombres if n in col), None)] if any(n in col for n in nombres) else ''
    salida = pd.DataFrame({
        'Prioridad_Fila': 2,
        'Num.Ext': tomar('NUM.EXT'),
        'Tipo de documento': tomar('TIPO'),
        'CUFE/CUDE': tomar('NÚMERO', 'NUMERO'),
        'NIT Emisor': tomar('NIT/C.C.', 'NIT'),
        'Nombre Emisor': tomar('TERCERO'),
        'BASE': tomar('BASE'),
        'IVA': tomar('IVA'),
        COL_ALERTAS: df[columna_alerta],
    }, index=df.index)
    salida = salida[_texto(salida['Tipo de documento']) != '']
    salida['Num.Ext_Original'] = _texto(salida['Num.Ext'])
    salida['Num.Ext_Clean'] = normalizar_num_ext(salida['Num.Ext_Original'])
    salida['Novedad'] = ''
    salida['Conteo_Pareja'] = 1
    for c in ('Diff_NIT', 'Diff_BASE', 'Diff_IVA'):
        salida[c] = False
    salida['Box_Color'] = ''
    return salida


def cargar_auditoria_guardada(ruta):
    """(DataFrame para el acta, descripción del origen) desde un libro ya conciliado.

    Usa la hoja DIAN VS CONT (acta completa); si no existe, una hoja con la columna ALERTA (solo alertas).
    """
    from shared.excel_io import hojas_de, leer_excel
    from core.errors import ErrorUsuario

    hojas = hojas_de(ruta)
    if HOJA_RESULTADO in hojas:
        return _desde_dian_vs_cont(leer_excel(ruta, rapido=True, sheet_name=HOJA_RESULTADO)), f"hoja '{HOJA_RESULTADO}'"
    for hoja in hojas:
        try:
            encabezado = leer_excel(ruta, rapido=True, sheet_name=hoja, nrows=0)
        except Exception:
            continue
        columna = next((c for c in encabezado.columns if str(c).strip().upper() == 'ALERTA'), None)
        if columna is not None:
            df = leer_excel(ruta, rapido=True, sheet_name=hoja)
            return _desde_hoja_auditoria(df, columna), f"hoja '{hoja}' (solo alertas tributarias)"
    raise ErrorUsuario(
        f"El archivo no tiene la hoja '{HOJA_RESULTADO}' ni una hoja de auditoría con la columna ALERTA. "
        "Ejecute primero la conciliación sobre ese archivo.")


class GeneradorActaErrores:
    def __init__(self, df_unificado, ruta_original):
        self.df = df_unificado.copy()
        self.ruta_original = Path(ruta_original)
        base = getSampleStyleSheet()
        self.estilo_titulo = ParagraphStyle('TituloActa', parent=base['Heading1'], fontName='Helvetica-Bold',
                                            fontSize=16, textColor=colors.black, alignment=TA_CENTER, spaceAfter=20)
        self.estilo_subtitulo = ParagraphStyle('Subtitulo', parent=base['Heading2'], fontName='Helvetica-Bold',
                                               fontSize=12, textColor=colors.darkgrey, spaceBefore=15, spaceAfter=10)
        self.estilo_normal = ParagraphStyle('TextoNormal', parent=base['Normal'], fontName='Helvetica',
                                            fontSize=10, alignment=TA_JUSTIFY, spaceAfter=10)
        self.estilo_celda = ParagraphStyle('Celda', parent=base['Normal'], fontName='Helvetica',
                                           fontSize=8, alignment=TA_LEFT)

    def nombre_sugerido(self):
        return f"Acta_Auditoria_{self.ruta_original.stem}_{datetime.now().strftime('%Y%m%d')}.pdf"

    def _p(self, texto):
        return Paragraph(str(texto), self.estilo_celda)

    def generar(self, ruta_pdf):
        elementos = []
        self._encabezado(elementos)
        hay_errores = False
        hay_errores |= self._seccion_parejas_discrepantes(elementos)
        hay_errores |= self._seccion_huerfanos(elementos)
        hay_errores |= self._seccion_otros_errores(elementos)
        hay_errores |= self._seccion_alertas(elementos)
        if not hay_errores:
            elementos.append(Paragraph("<b>RESULTADO EXCELENTE:</b> No se encontraron inconsistencias en la auditoría.", self.estilo_normal))

        SimpleDocTemplate(ruta_pdf, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40).build(elementos)
        return ruta_pdf

    def _encabezado(self, elementos):
        elementos.append(Paragraph("MEMORANDO DE AUDITORÍA CONTABLE", self.estilo_titulo))
        elementos.append(Paragraph(f"""
        <b>Fecha de emisión:</b> {datetime.now().strftime("%d/%m/%Y %H:%M")}<br/>
        <b>Archivo auditado:</b> {self.ruta_original.name}<br/><br/>
        <b>Asunto:</b> Acta de inconsistencias y correcciones requeridas.<br/><br/>
        El presente documento detalla las discrepancias halladas durante el cruce automatizado
        entre el reporte de la DIAN y los movimientos contables registrados. Se requiere la
        revisión y ajuste correspondiente en el software contable de los siguientes hallazgos:
        """, self.estilo_normal))
        elementos.append(Spacer(1, 10))

    def _seccion_parejas_discrepantes(self, elementos):
        df = self.df[self.df['Conteo_Pareja'] >= 2]
        df = df[(df['Diff_NIT'] == True) | (df['Diff_BASE'] == True) | (df['Diff_IVA'] == True) | (df['Box_Color'] == 'YELLOW')]  # noqa: E712
        if df.empty:
            return False

        elementos.append(Paragraph("1. DISCREPANCIAS EN REGISTROS EMPAREJADOS", self.estilo_subtitulo))
        elementos.append(Paragraph(
            "Las siguientes facturas fueron encontradas en ambos reportes, pero presentan diferencias en el "
            "<b>NIT del tercero</b>, los <b>valores contabilizados (Base/IVA)</b>, o <b>no generan efecto cero</b> "
            "tras sumar comprobantes y devoluciones:", self.estilo_normal))

        for num_ext, grupo in df.groupby('Num.Ext_Clean'):
            grupo = grupo.sort_values(by='Prioridad_Fila')
            primera = grupo.iloc[0]
            errores = [txt for cond, txt in (
                (primera['Box_Color'] == 'YELLOW', "No hay Efecto Cero"),
                (primera['Diff_BASE'] == True, "Diferencia en Base"),  # noqa: E712
                (primera['Diff_IVA'] == True, "Diferencia en IVA"),  # noqa: E712
                (primera['Diff_NIT'] == True, "NIT no coincide"),  # noqa: E712
            ) if cond]
            elementos.append(self._p(f"<b>Num.Ext: {num_ext}</b> - <i>Errores: {', '.join(errores)}</i>"))

            datos = [["Origen", "NIT Emisor", "Tercero", "Base", "IVA", "Tipo Doc"]]
            for _, f in grupo.iterrows():
                datos.append([
                    self._p(_origen(f['Prioridad_Fila'], con_devol=True)),
                    self._p(nit_texto(f['NIT Emisor'])),
                    self._p(str(f['Nombre Emisor'])[:25]),
                    _moneda(f['BASE']), _moneda(f['IVA']),
                    self._p(f['Tipo de documento']),
                ])
            tabla = Table(datos, colWidths=[50, 70, 160, 80, 80, 70])
            tabla.setStyle(TableStyle(ESTILO_BASE_TABLA + [
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (3, 1), (4, -1), 'RIGHT'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
            ]))
            elementos += [tabla, Spacer(1, 10)]
        return True

    def _seccion_huerfanos(self, elementos):
        df = self.df[(self.df['Conteo_Pareja'] == 1) & self.df['Novedad'].str.contains('Sin Pareja', na=False)]
        if df.empty:
            return False

        elementos.append(Paragraph("2. DOCUMENTOS FALTANTES (HUÉRFANOS)", self.estilo_subtitulo))
        elementos.append(Paragraph("Los siguientes registros existen en una plataforma pero no en la otra:", self.estilo_normal))

        datos = [["Origen", "Factura", "NIT", "Tercero", "Base", "Observación"]]
        for _, f in df.sort_values(by=['Prioridad_Fila', 'Num.Ext']).iterrows():
            datos.append([
                _origen(f['Prioridad_Fila']), str(f['Num.Ext']), nit_texto(f['NIT Emisor']),
                self._p(str(f['Nombre Emisor'])[:30]), _moneda(f['BASE']), self._p(f['Novedad']),
            ])
        tabla = Table(datos, colWidths=[50, 70, 70, 140, 70, 110], repeatRows=1)
        tabla.setStyle(TableStyle(ESTILO_BASE_TABLA + [('ALIGN', (4, 1), (4, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        elementos += [tabla, Spacer(1, 10)]
        return True

    def _seccion_otros_errores(self, elementos):
        nov = self.df['Novedad']
        df = self.df[
            ~nov.isin(NOVEDADES_SIN_PAREJA) & nov.notna()
            & ~nov.str.contains('Sin Pareja', na=False)
            & ~nov.str.startswith('DIAN - Gasto Personal', na=False)
            & ~nov.str.contains('Application Response', na=False, case=False)
            & ~nov.str.startswith('DIAN:')
        ]
        if df.empty:
            return False

        elementos.append(Paragraph("3. DUPLICADOS Y ERRORES DE ESTRUCTURA", self.estilo_subtitulo))
        elementos.append(Paragraph(
            "Se detectaron problemas en la digitación, documentos registrados más de una vez, o falta de número de factura:",
            self.estilo_normal))

        datos = [["Origen", "Factura Digitada", "Tercero", "Error Detectado"]]
        for _, f in df.sort_values(by=['Novedad', 'Prioridad_Fila']).iterrows():
            datos.append([
                _origen(f['Prioridad_Fila']), str(f['Num.Ext_Original']),
                self._p(str(f['Nombre Emisor'])[:35]), self._p(f['Novedad']),
            ])
        tabla = Table(datos, colWidths=[50, 100, 170, 190], repeatRows=1)
        tabla.setStyle(TableStyle(ESTILO_BASE_TABLA + [('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        elementos.append(tabla)
        return True

    def _alertas_detalle(self):
        """Una fila por (documento, alerta); la alerta de una pareja se reporta una sola vez."""
        if COL_ALERTAS not in self.df.columns:
            return pd.DataFrame()
        df = self.df[self.df[COL_ALERTAS].fillna('').astype(str).str.strip() != ''].copy()
        if df.empty:
            return df
        df['Texto'] = df[COL_ALERTAS].astype(str).str.split(' / ')
        df = df.explode('Texto')
        df['Texto'] = df['Texto'].str.strip()
        df = df[df['Texto'] != '']
        df['Tipo'] = df['Texto'].map(tipo_alerta).replace('', 'Otras')
        clave = df['Num.Ext_Clean'].fillna('').astype(str)
        df['Clave'] = clave.where(clave != '', 'fila' + df.index.astype(str))
        df = df.sort_values(by=['Clave', 'Prioridad_Fila'], kind='stable')
        return df.drop_duplicates(subset=['Clave', 'Tipo'], keep='first')

    def _documento(self, f):
        num_ext = str(f.get('Num.Ext_Original', '') or '').strip()
        if num_ext and num_ext.lower() != 'nan':
            return num_ext
        tipo, numero = str(f.get('Tipo de documento', '') or ''), str(f.get('CUFE/CUDE', '') or '')
        return f"{tipo} {numero}".replace('nan', '').strip() or 'Sin número'

    def _seccion_alertas(self, elementos):
        df = self._alertas_detalle()
        if df.empty:
            return False

        elementos.append(Paragraph("4. ALERTAS TRIBUTARIAS", self.estilo_subtitulo))
        elementos.append(Paragraph(
            "Errores tributarios detectados (retención en la fuente, IVA, soporte y periodo). "
            "Las diferencias menores a $50 o al 0,1 % del valor se ignoran.", self.estilo_normal))

        # Resumen por tipo de alerta, en el orden de la guía
        conteo = df['Tipo'].value_counts()
        datos = [["Alerta", "Casos", "Qué significa / qué revisar"]]
        orden = [n for n in conteo.index if n != 'Otras'] + (['Otras'] if 'Otras' in conteo.index else [])
        for nombre in sorted(orden, key=lambda n: (explicacion_alerta(n)[0] == '', -conteo[n])):
            significa = explicacion_alerta(nombre)[2] or 'Revisar el detalle de la alerta.'
            if nombre == SIN_SOPORTE:
                significa += ' Detalle en las secciones 2 y 3.'
            datos.append([self._p(f"<b>{escapar_xml(nombre)}</b>"), str(int(conteo[nombre])), self._p(escapar_xml(significa))])
        tabla = Table(datos, colWidths=[190, 40, 280], repeatRows=1)
        tabla.setStyle(TableStyle(ESTILO_BASE_TABLA + [('ALIGN', (1, 1), (1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        elementos += [tabla, Spacer(1, 12)]

        # Detalle por documento (los documentos sin soporte ya se listan en las secciones 2 y 3)
        df = df[df['Tipo'] != SIN_SOPORTE]
        if df.empty:
            return True
        elementos.append(Paragraph("<b>Detalle por documento</b>", self.estilo_normal))
        datos = [["Origen", "Documento", "NIT", "Tercero", "Alerta"]]
        for _, f in df.sort_values(by=['Tipo', 'Nombre Emisor', 'Clave'], kind='stable').iterrows():
            datos.append([
                _origen(f['Prioridad_Fila'], con_devol=True), self._p(escapar_xml(self._documento(f))), self._p(nit_texto(f['NIT Emisor'])),
                self._p(escapar_xml(str(f['Nombre Emisor'])[:30])), self._p(escapar_xml(f['Texto'])),
            ])
        tabla = Table(datos, colWidths=[45, 75, 65, 115, 210], repeatRows=1)
        tabla.setStyle(TableStyle(ESTILO_BASE_TABLA + [('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        elementos.append(tabla)
        return True