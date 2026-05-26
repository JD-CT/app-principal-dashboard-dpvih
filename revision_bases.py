#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# revision_bases.py - Modulo de Revision de Bases ITS/TAR/PrEP
# Proposito: Automatizar el proceso manual de 6 pasos descrito en
#   "Proceso de Revision de EESS, BMU, OBC"
# Fuente: Trama Unificada (73 columnas, 3 origenes: ITS, TAR, PrEP)
# Estilo: NVIH develop (lista plana de diccionarios, sin decoradores)
# Creado: 2026-05-25
# Dependencias: pandas, openpyxl

import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import logging
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
log = logging.getLogger(__name__)

# ──────────────────────────────────────────────
# CONFIGURACION: columnas esperadas
# ──────────────────────────────────────────────
COLUMNAS_ITS = [
    'sector_its', 'diresa_its', 'region_its', 'provincia_its', 'distrito_its',
    'eess_its', 'renipres_its', 'tipo_oferta_its', 'doc_brigadista_its',
    'nom_brigadista_its', 'ape_brigadista_its', 'brigada_its',
    'fecha_intervencion', 'tipo_lugar_its',
    'uid', 'edad', 'sexo', 'tipo_poblacion', 'nacionalidad',
    'fecha_consejeria', 'fecha_tamizaje', 'tipo_tamizaje', 'resultado',
    'vinculo_tipo', 'vinculo_estado', 'vinculo_eess_dest',
    'vinculo_fecha', 'vinculo_fecha_reg', 'fecha_reg_its', 'fecha_mod_its',
    'origen_registro'
]

COLUMNAS_TAR = [
    'sector_vih', 'diris_vih', 'depto_vih', 'provincia_vih', 'distrito_vih',
    'eess_vih', 'renipres_vih', 'doc_vih', 'nombre_vih', 'apellido_vih',
    'condicion_vih', 'fecha_inicio_tar', 'fecha_ult_atencion',
    'fecha_prox_cita', 'fecha_ult_abandono', 'fecha_recup_abandono',
    'fecha_derivacion', 'lugar_derivacion', 'fecha_creacion_vih', 'fecha_mod_vih'
]

COLUMNAS_PREP = [
    'sector_prep', 'diris_prep', 'depto_prep', 'provincia_prep', 'distrito_prep',
    'eess_prep', 'renipres_prep', 'doc_prep', 'nombre_prep', 'apellido_prep',
    'fecha_inicio_prep', 'eess_inicio_prep', 'modalidad_prep',
    'fecha_reingreso_prep', 'condicion_actual', 'fecha_ult_atencion_prep',
    'fecha_ult_tamizaje_vih', 'resultado_ult_tamizaje_vih',
    'fecha_creacion_prep', 'fecha_mod_prep'
]

COLUMNAS_ESPERADAS = COLUMNAS_ITS + COLUMNAS_TAR + COLUMNAS_PREP

# ──────────────────────────────────────────────
# VERIFICACIONES
# ──────────────────────────────────────────────
VERIFICACIONES = [
    {
        'n': 'filtro_vih_dvi',
        't': 'Filtro: solo registros tamizaje VIH y DVI',
        'd': 'Muestra los registros cuyo tipo de tamizaje es VIH o DVI. '
             'Se excluyen HEB (Hepatitis B), SIF (Sifilis), DSI y HEC (Hepatitis C). '
             'Estos registros filtrados son la base para todas las verificaciones.',
        'c': ['tipo_tamizaje'],
        'resaltar': 'tipo_tamizaje',
        'p': 'baja',
        'cat': 'filtro'
    },
    {
        'n': 'duplicados_tamizaje',
        't': 'Duplicados de paciente por UID repetido',
        'd': 'Detecta registros duplicados donde un mismo paciente (UID) '
             'aparece mas de una vez. Se debe conservar un solo registro '
             'por paciente, eliminando las copias.',
        'c': ['uid'],
        'resaltar': 'uid',
        'p': 'media',
        'cat': 'duplicados'
    },
    {
        'n': 'reactivos_sin_vinculacion',
        't': 'Reactivos VIH sin registro en TAR ni PrEP',
        'd': 'Pacientes con resultado REACTIVO o INCONCLUSO que no estan '
             'registrados en TAR (ESTABLECIMIENTO DE SALUD - VIH vacio) '
             'ni en PrEP (ESTABLECIMIENTO DE SALUD - PREP vacio). '
             'Requieren vinculacion inmediata a uno de los dos modulos.',
        'c': ['resultado', 'eess_vih', 'eess_prep'],
        'resaltar': 'resultado',
        'p': 'critica',
        'cat': 'brecha'
    },
    {
        'n': 'vinculados_sin_padron',
        't': 'Vinculacion efectiva sin respaldo en TAR ni PrEP',
        'd': 'Pacientes con estado EFECTIVA registrado en el modulo ITS, '
             'pero que no aparecen en TAR (ESTABLECIMIENTO DE SALUD - VIH vacio) '
             'ni en PrEP (ESTABLECIMIENTO DE SALUD - PREP vacio). '
             'La vinculacion deberia reflejarse en al menos un modulo.',
        'c': ['vinculo_estado', 'eess_vih', 'eess_prep'],
        'resaltar': 'vinculo_estado',
        'p': 'media',
        'cat': 'brecha'
    },
    {
        'n': 'en_tar_sin_tamizaje_its',
        't': 'En TAR sin tamizaje ITS registrado',
        'd': 'Paciente que aparece en TAR (con condicion_vih y fecha_inicio_tar) '
             'pero no tiene ningun registro en ITS (uid no aparece en datos ITS).',
        'c': ['uid', 'fecha_tamizaje', 'condicion_vih'],
        'resaltar': 'condicion_vih',
        'p': 'media',
        'cat': 'brecha'
    },
    {
        'n': 'vinculacion_fecha_inconsistente',
        't': 'Fecha de vinculacion posterior al inicio de TAR o PrEP',
        'd': 'La fecha de vinculacion registrada en ITS es posterior '
             'a la fecha de inicio en TAR o PrEP. Sugiere que el paciente '
             'ya estaba en tratamiento sin vinculacion formal registrada. '
             'La vinculacion debe ocurrir antes o el mismo dia del inicio.',
        'c': ['vinculo_fecha', 'fecha_inicio_tar', 'fecha_inicio_prep'],
        'resaltar': 'vinculo_fecha',
        'p': 'media',
        'cat': 'consistencia'
    },

    {
        'n': 'prep_sin_tamizaje_vih_reciente',
        't': 'PrEP sin tamizaje VIH en los ultimos 90 dias',
        'd': 'Paciente registrado en PrEP cuya fecha del ultimo tamizaje VIH '
             'supera los 90 dias o esta vacia. El protocolo exige tamizaje '
             'VIH trimestral para todos los usuarios de PrEP.',
        'c': ['condicion_actual', 'fecha_ult_tamizaje_vih'],
        'resaltar': 'condicion_actual',
        'p': 'media',
        'cat': 'consistencia'
    },
]

# ──────────────────────────────────────────────
# FUNCIONES DE VERIFICACION
# ──────────────────────────────────────────────
CRITERIOS = {v['n']: f"Columna(s): {', '.join(v['c'])}" for v in VERIFICACIONES}

RUTA_PLANTILLA = 'Trama_Unificada_encabezados.xlsx'


def _filtro_vih_dvi(df):
    """Filtra solo registros VIH y DVI (Dual VIH).
    Devuelve los que pasan el filtro para que se muestren en el reporte
    con la columna tipo_tamizaje resaltada en amarillo."""
    if 'tipo_tamizaje' not in df.columns:
        return pd.DataFrame()
    d = df[df['tipo_tamizaje'].str.upper().isin(['VIH', 'DVI'])]
    return d if not d.empty else pd.DataFrame()


def _duplicados_tamizaje(df):
    """Detectar duplicados por UID"""
    if 'uid' not in df.columns:
        return pd.DataFrame()
    dup_mask = df.duplicated(subset=['uid'], keep=False)
    d = df[dup_mask].copy()
    return d if not d.empty else pd.DataFrame()


def _reactivos_sin_vinculacion(df):
    """Reactivos VIH sin vinculacion a TAR ni PrEP"""
    if 'resultado' not in df.columns or 'eess_vih' not in df.columns or 'eess_prep' not in df.columns:
        return pd.DataFrame()
    # Reactivos que NO estan registrados en TAR (eess_vih nulo/vacio)
    # NI en PrEP (eess_prep nulo/vacio)
    reactivo = df['resultado'].str.upper().isin(['REACTIVO', 'INCONCLUSO'])
    no_tar = df['eess_vih'].isna() | (df['eess_vih'].astype(str).str.strip() == '')
    no_prep = df['eess_prep'].isna() | (df['eess_prep'].astype(str).str.strip() == '')
    d = df[reactivo & no_tar & no_prep]
    return d if not d.empty else pd.DataFrame()


def _vinculados_sin_padron(df):
    """Vinculacion efectiva en ITS sin registro en TAR ni PrEP"""
    if 'vinculo_estado' not in df.columns or 'eess_vih' not in df.columns or 'eess_prep' not in df.columns:
        return pd.DataFrame()
    efec = df['vinculo_estado'].str.lower() == 'efectiva'
    no_tar = df['eess_vih'].isna() | (df['eess_vih'].astype(str).str.strip() == '')
    no_prep = df['eess_prep'].isna() | (df['eess_prep'].astype(str).str.strip() == '')
    d = df[efec & no_tar & no_prep]
    return d if not d.empty else pd.DataFrame()


def _en_tar_sin_tamizaje_its(df):
    """En TAR pero sin tamizaje ITS"""
    if 'uid' not in df.columns or 'fecha_tamizaje' not in df.columns:
        return pd.DataFrame()
    d = df[
        (df['condicion_vih'].notna()) &
        (df['condicion_vih'] != '') &
        (df['fecha_tamizaje'].isna())
    ]
    return d if not d.empty else pd.DataFrame()


def _vinculacion_fecha_inconsistente(df):
    """Vinculacion posterior al inicio de TAR o PrEP"""
    if 'vinculo_fecha' not in df.columns:
        return pd.DataFrame()
    d = df[df['vinculo_fecha'].notna()].copy()
    if d.empty:
        return pd.DataFrame()
    d['_vf'] = pd.to_datetime(d['vinculo_fecha'], errors='coerce')
    condicion = pd.Series(False, index=d.index)
    if 'fecha_inicio_tar' in d.columns:
        d['_tar'] = pd.to_datetime(d['fecha_inicio_tar'], errors='coerce')
        condicion |= d['_vf'].notna() & d['_tar'].notna() & (d['_vf'] > d['_tar'])
    if 'fecha_inicio_prep' in d.columns:
        d['_prep'] = pd.to_datetime(d['fecha_inicio_prep'], errors='coerce')
        condicion |= d['_vf'].notna() & d['_prep'].notna() & (d['_vf'] > d['_prep'])
    d = d[condicion]
    drop_cols = [c for c in ['_vf', '_tar', '_prep'] if c in d.columns]
    d = d.drop(columns=drop_cols)
    return d if not d.empty else pd.DataFrame()
def _prep_sin_tamizaje_vih_reciente(df):
    """PrEP activo sin tamizaje VIH en ultimos 3 meses"""
    if 'condicion_actual' not in df.columns or 'fecha_ult_tamizaje_vih' not in df.columns:
        return pd.DataFrame()
    d = df[
        (df['condicion_actual'].notna()) &
        (df['condicion_actual'] != '')
    ].copy()
    d['_futv'] = pd.to_datetime(d['fecha_ult_tamizaje_vih'], errors='coerce')
    hoy = pd.Timestamp.now()
    d['_dias'] = (hoy - d['_futv']).dt.days if not d['_futv'].isna().all() else 999
    d = d[d['_futv'].isna() | (d['_dias'] > 90)]
    d = d.drop(columns=['_futv', '_dias'], errors='ignore')
    return d if not d.empty else pd.DataFrame()


# ──────────────────────────────────────────────
# FUNCION MAPA: nombre -> funcion
# ──────────────────────────────────────────────
FUNCIONES = {
    'filtro_vih_dvi': _filtro_vih_dvi,
    'duplicados_tamizaje': _duplicados_tamizaje,
    'reactivos_sin_vinculacion': _reactivos_sin_vinculacion,
    'vinculados_sin_padron': _vinculados_sin_padron,
        'en_tar_sin_tamizaje_its': _en_tar_sin_tamizaje_its,
    'vinculacion_fecha_inconsistente': _vinculacion_fecha_inconsistente,
        'prep_sin_tamizaje_vih_reciente': _prep_sin_tamizaje_vih_reciente,
}


# ──────────────────────────────────────────────
# ANALIZADOR PRINCIPAL
# ──────────────────────────────────────────────
class AnalizadorRevisionBases:
    """Analiza la Trama Unificada (ITS + TAR + PrEP) para Revision de Bases"""

    MAPA_COLUMNAS = {
        'codigo_paciente': 'uid',
        'paciente_sexo': 'sexo',
        'paciente_edad': 'edad',
        'paciente_tipo_poblacion': 'tipo_poblacion',
        'resultado_1': 'resultado',
        'tipo_de_vinculacion': 'vinculo_tipo',
        'estado_de_vinculacion': 'vinculo_estado',
        'fecha_de_vinculacion': 'vinculo_fecha',
        'fecha_de_registro': 'vinculo_fecha_reg',
        'eess_destino_de_vinculacion': 'vinculo_eess_dest',
        'condicion': 'condicion_vih',
        'fecha_de_inicio_de_tar': 'fecha_inicio_tar',
        'fecha_de_ultima_atencion': 'fecha_ult_atencion',
        'fecha_de_proxima_cita': 'fecha_prox_cita',
        'fecha_de_ultimo_abandono': 'fecha_ult_abandono',
        'fecha_de_ultima_recuperacion_de_abandono': 'fecha_recup_abandono',
        'fecha_de_derivacion': 'fecha_derivacion',
        'lugar_de_destino_de_derivacion': 'lugar_derivacion',
        'eess_de_inicio_prep': 'eess_inicio_prep',
        'modalidad_de_inicio_prep': 'modalidad_prep',
        'fecha_de_reingreso_a_la_prep': 'fecha_reingreso_prep',
        'fecha_de_ultima_atencion_prep': 'fecha_ult_atencion_prep',
        'fecha_de_resultado_ultimo_tamizaje_de_vih': 'fecha_ult_tamizaje_vih',
        'resultado_ultimo_tamizaje_de_vih': 'resultado_ult_tamizaje_vih',
        'condicion_actual': 'condicion_actual',
        'sector___its': 'sector_its',
        'diresa___its': 'diresa_its',
        'region___its': 'region_its',
        'provincia___its': 'provincia_its',
        'distrito___its': 'distrito_its',
        'eess___its': 'eess_its',
        'codigo_renipres___its': 'renipres_its',
        'tipo_oferta': 'tipo_oferta_its',
        'tipo_lugar_intervencion': 'tipo_lugar_its',
        'fecha_intervencion': 'fecha_intervencion',
        'fecha_de_tamizaje': 'fecha_tamizaje',
        'tipo_tamizaje': 'tipo_tamizaje',
        'fecha_de_registro___its': 'fecha_reg_its',
        'fecha_de_modificacion___its': 'fecha_mod_its',
        'origen_de_registro': 'origen_registro',
        'sector___vih': 'sector_vih',
        'diris_diresa_geresa___vih': 'diris_vih',
        'departamento_eess___vih': 'depto_vih',
        'provincia_eess___vih': 'provincia_vih',
        'distrito_eess___vih': 'distrito_vih',
        'establecimiento_de_salud___vih': 'eess_vih',
        'codigo_renipres___vih': 'renipres_vih',
        'sector___prep': 'sector_prep',
        'diris_diresa_geresa___prep': 'diris_prep',
        'departamento_eess___prep': 'depto_prep',
        'provincia_eess___prep': 'provincia_prep',
        'distrito_eess___prep': 'distrito_prep',
        'establecimiento_de_salud___prep': 'eess_prep',
        'codigo_renipres___prep': 'renipres_prep',
        'fecha_de_creacion___vih': 'fecha_creacion_vih',
        'fecha_de_modificacion___vih': 'fecha_mod_vih',
        'fecha_de_creacion___prep': 'fecha_creacion_prep',
        'fecha_de_modificacion___prep': 'fecha_mod_prep',
        'fecha_de_inicio_de_prep': 'fecha_inicio_prep',
    }

    ENCABEZADOS = {
        # ITS
        'sector_its': 'SECTOR - ITS',
        'diresa_its': 'DIRESA - ITS',
        'region_its': 'REGION - ITS',
        'provincia_its': 'PROVINCIA - ITS',
        'distrito_its': 'DISTRITO - ITS',
        'eess_its': 'EESS - ITS',
        'renipres_its': 'CODIGO RENIPRES - ITS',
        'uid': 'CODIGO UID',
        'edad': 'PACIENTE EDAD',
        'sexo': 'PACIENTE SEXO',
        'tipo_poblacion': 'PACIENTE TIPO POBLACION',
        'nacionalidad': 'NACIONALIDAD',
        'etnia': 'ETNIA',
        'brigada': 'BRIGADA',
        'fecha_intervencion': 'FECHA INTERVENCION',
        'tipo_lugar_its': 'TIPO LUGAR INTERVENCION',
        'tipo_oferta_its': 'TIPO OFERTA',
        'fecha_tamizaje': 'FECHA DE TAMIZAJE',
        'tipo_tamizaje': 'TIPO TAMIZAJE',
        'resultado': 'RESULTADO 1',
        'fecha_reg_its': 'FECHA DE REGISTRO - ITS',
        'fecha_mod_its': 'FECHA DE MODIFICACION - ITS',
        'vinculo_tipo': 'TIPO DE VINCULACION',
        'vinculo_estado': 'ESTADO DE VINCULACION',
        'vinculo_eess_dest': 'EESS DESTINO DE VINCULACION',
        'vinculo_fecha': 'FECHA DE VINCULACION',
        'vinculo_fecha_reg': 'FECHA DE REGISTRO',
        'origen_registro': 'ORIGEN DE REGISTRO',
        # TAR
        'sector_vih': 'SECTOR - VIH',
        'diris_vih': 'DIRIS/DIRESA/GERESA - VIH',
        'depto_vih': 'DEPARTAMENTO EESS - VIH',
        'provincia_vih': 'PROVINCIA EESS - VIH',
        'distrito_vih': 'DISTRITO EESS - VIH',
        'eess_vih': 'ESTABLECIMIENTO DE SALUD - VIH',
        'renipres_vih': 'CODIGO RENIPRES - VIH',
        'condicion_vih': 'CONDICION',
        'fecha_inicio_tar': 'FECHA DE INICIO DE TAR',
        'fecha_ult_atencion': 'FECHA DE ULTIMA ATENCION',
        'fecha_prox_cita': 'FECHA DE PROXIMA CITA',
        'fecha_ult_abandono': 'FECHA DE ULTIMO ABANDONO',
        'fecha_recup_abandono': 'FECHA DE ULTIMA RECUPERACION DE ABANDONO',
        'fecha_derivacion': 'FECHA DE DERIVACION',
        'lugar_derivacion': 'LUGAR DE DESTINO DE DERIVACION',
        'fecha_creacion_vih': 'FECHA DE CREACION - VIH',
        'fecha_mod_vih': 'FECHA DE MODIFICACION - VIH',
        # PrEP
        'sector_prep': 'SECTOR - PREP',
        'diris_prep': 'DIRIS/DIRESA/GERESA - PREP',
        'depto_prep': 'DEPARTAMENTO EESS - PREP',
        'provincia_prep': 'PROVINCIA EESS - PREP',
        'distrito_prep': 'DISTRITO EESS - PREP',
        'eess_prep': 'ESTABLECIMIENTO DE SALUD - PREP',
        'renipres_prep': 'CODIGO RENIPRES - PREP',
        'fecha_inicio_prep': 'FECHA DE INICIO DE PREP',
        'eess_inicio_prep': 'EESS DE INICIO PREP',
        'modalidad_prep': 'MODALIDAD DE INICIO PREP',
        'fecha_reingreso_prep': 'FECHA DE REINGRESO A LA PREP',
        'condicion_actual': 'CONDICION ACTUAL',
        'fecha_ult_atencion_prep': 'FECHA DE ULTIMA ATENCION PREP',
        'fecha_ult_tamizaje_vih': 'FECHA DE RESULTADO ULTIMO TAMIZAJE DE VIH',
        'resultado_ult_tamizaje_vih': 'RESULTADO ULTIMO TAMIZAJE DE VIH',
        'fecha_creacion_prep': 'FECHA DE CREACION - PREP',
        'fecha_mod_prep': 'FECHA DE MODIFICACION - PREP',
    }

    def __init__(self):
        self.df = None
        self.total_registros = 0
        self.resumen = []
        self.resultados = {}
        self._v_omitidas = []
        self.tiempos = {}

    def cargar_datos(self, ruta_excel, hoja=None):
        """Carga el Excel de la Trama Unificada"""
        if hoja is None:
            try:
                xl = pd.ExcelFile(ruta_excel)
                hoja = xl.sheet_names[0]
            except Exception:
                pass
        df = pd.read_excel(ruta_excel, sheet_name=hoja)
        # Normalizar nombres de columnas
        df.columns = [c.strip().lower().replace(' ', '_').replace('-', '_') for c in df.columns]
        # Mapear a nombres canonicos
        rename_map = {}
        for col in df.columns:
            if col in self.MAPA_COLUMNAS:
                rename_map[col] = self.MAPA_COLUMNAS[col]
        if rename_map:
            df = df.rename(columns=rename_map)
            log.info(f"Columnas mapeadas: {len(rename_map)}")
        self.df = df
        self.total_registros = len(df)
        log.info(f"Cargados {self.total_registros} registros")
        return self

    def _campos_existen(self, cols):
        """Verifica que las columnas necesarias existan"""
        faltan = [c for c in cols if c not in self.df.columns]
        return faltan if faltan else None

    
    def _formatear_hoja_detalle(self, ws, max_col):
        """Aplica colores y ancho a hoja ya escrita con pandas"""
        colores_f3 = {
            (1, 10): '4472C4', (11, 16): 'FFC000',
            (17, 25): '4472C4', (26, 33): 'FFC000',
            (34, 53): 'ED7D31', (54, 73): '70AD47',
            (74, 100): '70AD47',
        }
        font_f3 = Font(bold=True, color='FFFFFF', size=10)
        for inicio, fin in colores_f3:
            c = colores_f3[(inicio, fin)]
            fill_rango = PatternFill(start_color=c, end_color=c, fill_type='solid')
            for col in range(inicio, min(fin, max_col) + 1):
                cell = ws.cell(row=1, column=col)
                cell.fill = fill_rango
                cell.font = font_f3
                cell.alignment = Alignment(horizontal='center', vertical='center')
        for col_idx in range(1, max_col + 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = 13.0


    def analizar(self):
        """Ejecuta todas las verificaciones"""
        t_total = datetime.now()
        self.resumen = []
        self.resultados = {}

        log.info(f"Iniciando Revision de Bases sobre {self.total_registros} registros")

        # Paso 1: Ejecutar filtro VIH/DVI primero (siempre)
        log.info("  Aplicando filtro VIH/DVI antes del analisis...")
        filtro_fn = FUNCIONES.get('filtro_vih_dvi')
        df_filtrado = None
        if filtro_fn and 'tipo_tamizaje' in self.df.columns:
            df_filtrado = filtro_fn(self.df)
        if df_filtrado is not None and len(df_filtrado) > 0:
            log.info(f"  Filtro aplicado: {len(df_filtrado)} registros VIH/DVI de {self.total_registros} totales")
            self.df_original = self.df.copy()
            self.df = df_filtrado
            self.total_filtrados = len(df_filtrado)
        else:
            log.info("  No se pudo aplicar filtro VIH/DVI, se trabaja con todos los registros")
            self.df_original = None
            self.total_filtrados = self.total_registros

        for idx, item in enumerate(VERIFICACIONES, 1):
            t0 = pd.Timestamp.now()
            nombre = item['n']
            fn = FUNCIONES.get(nombre)

            faltan = self._campos_existen(item['c'])
            if faltan or fn is None:
                self._v_omitidas.append(nombre)
                detalle = f"Cols faltantes: {faltan}" if faltan else f"No implementada: {nombre}"
                self.resumen.append({
                    'ID': idx, 'Verificacion': item['t'],
                    'Categoria': item['cat'],                     'Registros': self.total_registros, 'Cantidad': 'N/A',
                    '%': 'N/A', 'Estado': 'OMITIDO',
                    'Criterio': CRITERIOS.get(nombre, ''),
                    'Descripcion': item['d'], 'Detalle': detalle,
                })
                log.warning(f"  OMITIDO: {nombre} -> {detalle}")
                continue

            try:
                # La verificacion #1 (filtro) usa df original, las demas usan df filtrado
                df_para_verif = self.df_original if nombre == 'filtro_vih_dvi' and hasattr(self, 'df_original') and self.df_original is not None else self.df
                registros = fn(df_para_verif)
                elapsed = (pd.Timestamp.now() - t0).total_seconds()
                self.tiempos[nombre] = elapsed

                total_base = len(df_para_verif)
                if registros is not None and len(registros) > 0:
                    cant = len(registros)
                    pct = (cant / total_base) * 100 if total_base > 0 else 0
                    self.resultados[nombre] = {
                        'id': idx, 'cantidad': cant, 'descripcion': item['d'],
                        'registros': registros, 'revisados': total_base,
                        'porcentaje': pct,
                    }
                    self.resumen.append({
                        'ID': idx, 'Verificacion': item['t'],
                        'Categoria': item['cat'],                         'Registros': total_base, 'Cantidad': cant,
                        '%': f"{pct:.2f}%", 'Estado': 'REVISAR',
                        'Criterio': CRITERIOS.get(nombre, ''),
                        'Descripcion': item['d'],
                    })
                    log.info(f"  -> {cant}/{total_base} registros ({pct:.2f}%) en {elapsed:.2f}s")
                else:
                    self.resumen.append({
                        'ID': idx, 'Verificacion': item['t'],
                        'Categoria': item['cat'],                         'Registros': total_base, 'Cantidad': 0,
                        '%': '0.00%', 'Estado': 'SIN PROBLEMAS',
                        'Criterio': CRITERIOS.get(nombre, ''),
                        'Descripcion': item['d'],
                    })
                    log.info(f"  -> 0 problemas en {elapsed:.2f}s")
            except Exception as e:
                elapsed = (pd.Timestamp.now() - t0).total_seconds()
                log.error(f"Error en {nombre}: {e}")
                self.resumen.append({
                    'ID': idx, 'Verificacion': item['t'],
                    'Categoria': item['cat'],                     'Registros': self.total_registros, 'Cantidad': 'ERROR',
                    '%': 'N/A', 'Estado': 'ERROR',
                    'Criterio': CRITERIOS.get(nombre, ''),
                    'Descripcion': item['d'], 'Detalle': str(e),
                })

        # Convertir columnas datetime a solo fecha (sin hora) en todos los resultados
        for nombre, res in self.resultados.items():
            registros = res.get('registros')
            if registros is not None and isinstance(registros, pd.DataFrame) and not registros.empty:
                for col in registros.select_dtypes(include=['datetime64']).columns:
                    registros[col] = registros[col].dt.date
        # Mismo para self.df
        if hasattr(self, 'df') and self.df is not None:
            for col in self.df.select_dtypes(include=['datetime64']).columns:
                self.df[col] = self.df[col].dt.date

        delta = (datetime.now() - t_total).total_seconds()
        log.info(f"Revision completa en {delta:.2f}s")
        return self

    
    def generar_reporte(self, nombre_base=None):
        """Genera Excel con resultados optimizado para 37k registros"""
        if nombre_base is None:
            nombre_base = f"Revision_Bases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        t_start = datetime.now()
        df_resumen = pd.DataFrame(self.resumen)
        amarillo = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

        # Paso 1: escribir hojas con pandas (rapido)
        with pd.ExcelWriter(nombre_base, engine='openpyxl') as writer:
            total_problemas = sum(
                r.get('Problemas', 0) for r in self.resumen
                if isinstance(r.get('Problemas', 0), (int, float))
            )
            criticos = sum(1 for r in self.resumen
                           if str(r.get('Prioridad', '')).lower() == 'critica'
                           and isinstance(r.get('Problemas', 0), (int, float))
                           and r.get('Problemas', 0) > 0)
            altos = sum(1 for r in self.resumen
                        if str(r.get('Prioridad', '')).lower() == 'alta'
                        and isinstance(r.get('Problemas', 0), (int, float))
                        and r.get('Problemas', 0) > 0)

            df_gen = pd.DataFrame({
                'Metrica': ['Total Registros', 'Verificaciones Registradas',
                    'Verificaciones con Problemas', 'Verificaciones Omitidas',
                    'Total Problemas Encontrados', 'Criticos', 'Alta Prioridad',
                    'Fecha Analisis', 'Script'],
                'Valor': [self.total_registros, len(VERIFICACIONES),
                    sum(1 for r in self.resumen
                        if isinstance(r.get('Problemas', 0), (int, float))
                        and r.get('Problemas', 0) > 0),
                    len(self._v_omitidas), total_problemas, criticos, altos,
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'Revision Bases v1.0']
            })
            df_gen.to_excel(writer, sheet_name='Resumen General', index=False)
            df_resumen.to_excel(writer, sheet_name='Resumen Verificaciones', index=False)

            for nombre, res in self.resultados.items():
                registros = res.get('registros')
                if registros is None or not isinstance(registros, pd.DataFrame) or registros.empty:
                    continue
                v = next((x for x in VERIFICACIONES if x['n'] == nombre), None)
                titulo = v['t'] if v else nombre
                safe_titulo = re.sub(r'[\/*?:\[\]]', '_', titulo)[:27]
                sn = f"{res['id']:02d}_{safe_titulo}"
                df_excel = registros.copy()
                rename_hdr = {k: v for k, v in self.ENCABEZADOS.items() if k in df_excel.columns}
                if rename_hdr:
                    df_excel = df_excel.rename(columns=rename_hdr)
                df_excel.to_excel(writer, sheet_name=sn, index=False)

        # Paso 2: aplicar formato con openpyxl (solo estilos)
        t1 = datetime.now()
        log.info(f"Escritura pandas: {(t1-t_start).total_seconds():.1f}s")
        wb = load_workbook(nombre_base)
        for nombre, res in self.resultados.items():
            registros = res.get('registros')
            if registros is None or not isinstance(registros, pd.DataFrame) or registros.empty:
                continue
            v = next((x for x in VERIFICACIONES if x['n'] == nombre), None)
            titulo = v['t'] if v else nombre
            safe_titulo = re.sub(r'[\/*?:\[\]]', '_', titulo)[:27]
            sn = f"{res['id']:02d}_{safe_titulo}"
            if sn not in wb.sheetnames:
                continue
            ws = wb[sn]

            self._formatear_hoja_detalle(ws, ws.max_column)

            # Resaltar todas las columnas del criterio
            cols_criterio = v['c'] if v and 'c' in v else []
            header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
            cols_resaltadas = []
            for col_crit in cols_criterio:
                col_excel = self.ENCABEZADOS.get(col_crit, col_crit)
                if col_excel in header:
                    col_idx = header.index(col_excel) + 1
                    for row in range(2, ws.max_row + 1):
                        ws.cell(row=row, column=col_idx).fill = amarillo
                    cols_resaltadas.append(col_excel)
            if cols_resaltadas:
                log.info(f"Columnas resaltadas en hoja '{sn}': {cols_resaltadas}")

        wb.save(nombre_base)
        wb.close()

        elapsed = (datetime.now() - t_start).total_seconds()
        log.info(f"Reporte final: {nombre_base} ({elapsed:.1f}s)")
        return nombre_base

