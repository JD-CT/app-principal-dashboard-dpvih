#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# revision_bases.py - Modulo de Revision de Bases ITS/TAR/PrEP
# Proposito: Automatizar el proceso manual de 6 pasos descrito en
#   "Proceso de Revision de EESS, BMU, OBC"
# Fuente: Trama Unificada (71 columnas, 3 origenes: ITS, TAR, PrEP)
# Estilo: NVIH develop (lista plana de diccionarios, sin decoradores)
# Creado: 2026-05-25
# Dependencias: pandas, openpyxl

import pandas as pd
import numpy as np
from datetime import datetime
import re
import logging

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
        'n': 'tamizajes_no_vih',
        't': 'Tamizajes que no son VIH/Dual',
        'd': 'Registros con tipo de tamizaje diferente a VIH o Dual. '
             'Según el proceso, deben filtrarse para quedarse solo con VIH/Dual.',
        'c': ['tipo_tamizaje'],
        'p': 'media',
        'cat': 'filtro'
    },
    {
        'n': 'duplicados_tamizaje',
        't': 'Duplicados por UID y fecha de tamizaje',
        'd': 'Mismo paciente (UID) tamizado en la misma fecha. '
             'Se elimina el último registro, quedándose con el primero.',
        'c': ['uid', 'fecha_tamizaje'],
        'p': 'alta',
        'cat': 'duplicados'
    },
    {
        'n': 'reactivos_sin_vinculacion',
        't': 'Reactivos VIH sin vinculación efectiva',
        'd': 'Pacientes con resultado REACTIVO en tamizaje VIH/Dual '
             'pero sin un registro de vinculación efectiva a TAR o PrEP.',
        'c': ['resultado', 'vinculo_estado'],
        'p': 'critica',
        'cat': 'brecha'
    },
    {
        'n': 'vinculados_sin_padron',
        't': 'Vinculados que no aparecen en TAR ni PrEP',
        'd': 'Paciente con vinculación efectiva registrada en ITS, '
             'pero sin datos correspondientes en TAR (condicion_vih) '
             'ni en PrEP (condicion_actual).',
        'c': ['vinculo_estado', 'condicion_vih', 'condicion_actual'],
        'p': 'alta',
        'cat': 'brecha'
    },
    {
        'n': 'reactivos_vih_sin_tar',
        't': 'Reactivos VIH sin inicio de TAR',
        'd': 'Paciente con resultado REACTIVO en tamizaje VIH '
             'pero sin fecha de inicio de TAR registrada.',
        'c': ['resultado', 'fecha_inicio_tar'],
        'p': 'critica',
        'cat': 'brecha'
    },
    {
        'n': 'en_tar_sin_tamizaje_its',
        't': 'En TAR sin tamizaje ITS registrado',
        'd': 'Paciente que aparece en TAR (con condicion_vih y fecha_inicio_tar) '
             'pero no tiene ningún registro en ITS (uid no aparece en datos ITS).',
        'c': ['uid', 'fecha_tamizaje', 'condicion_vih'],
        'p': 'alta',
        'cat': 'brecha'
    },
    {
        'n': 'vinculacion_fecha_inconsistente',
        't': 'Vinculación con fecha anterior al tamizaje',
        'd': 'La fecha de vinculación es anterior a la fecha de tamizaje. '
             'No es posible vincular antes de tamizar.',
        'c': ['vinculo_fecha', 'fecha_tamizaje'],
        'p': 'alta',
        'cat': 'consistencia'
    },
    {
        'n': 'poblacion_incoherente',
        't': 'Población clave incoherente con sexo/edad',
        'd': 'Registros donde la población clave no corresponde con '
             'el sexo biológico o la edad del paciente. '
             'Ej: HSH con sexo Femenino, Gestante con sexo Masculino, '
             'TS/TRA en menor de 15 años.',
        'c': ['sexo', 'tipo_poblacion', 'edad'],
        'p': 'alta',
        'cat': 'consistencia'
    },
    {
        'n': 'prep_sin_tamizaje_vih_reciente',
        't': 'PrEP sin tamizaje VIH en los últimos 3 meses',
        'd': 'Paciente en PrEP (con condicion_actual) cuya fecha del último '
             'tamizaje VIH es mayor a 90 días o está vacía.',
        'c': ['condicion_actual', 'fecha_ult_tamizaje_vih'],
        'p': 'alta',
        'cat': 'consistencia'
    },
]

# ──────────────────────────────────────────────
# FUNCIONES DE VERIFICACION
# ──────────────────────────────────────────────
CRITERIOS = {v['n']: f"Columna(s): {', '.join(v['c'])} | Prioridad: {v['p']}" for v in VERIFICACIONES}


def _tamizajes_no_vih(df):
    """Filtrar tamizajes que NO son VIH/Dual"""
    if 'tipo_tamizaje' not in df.columns:
        return pd.DataFrame()
    d = df[~df['tipo_tamizaje'].str.upper().str.contains('VIH|DUAL', na=False)]
    return d if not d.empty else pd.DataFrame()


def _duplicados_tamizaje(df):
    """Detectar duplicados por UID + fecha_tamizaje"""
    if 'uid' not in df.columns or 'fecha_tamizaje' not in df.columns:
        return pd.DataFrame()
    col_fecha = 'fecha_tamizaje'
    dup_mask = df.duplicated(subset=['uid', col_fecha], keep='first')
    d = df[dup_mask].copy()
    return d if not d.empty else pd.DataFrame()


def _reactivos_sin_vinculacion(df):
    """Reactivos VIH/Dual sin vinculación efectiva"""
    if 'resultado' not in df.columns or 'vinculo_estado' not in df.columns:
        return pd.DataFrame()
    d = df[
        (df['resultado'].str.upper() == 'REACTIVO') &
        (~df['vinculo_estado'].str.upper().str.contains('EFECTIVA|EFECTIVO', na=False))
    ]
    return d if not d.empty else pd.DataFrame()


def _vinculados_sin_padron(df):
    """Vinculacion efectiva pero sin dato en TAR ni PrEP"""
    req = ['vinculo_estado', 'condicion_vih', 'condicion_actual']
    if not all(c in df.columns for c in req):
        return pd.DataFrame()
    efec = df['vinculo_estado'].str.upper().str.contains('EFECTIVA|EFECTIVO', na=False)
    sin_tar = df['condicion_vih'].isna() | (df['condicion_vih'] == '')
    sin_prep = df['condicion_actual'].isna() | (df['condicion_actual'] == '')
    d = df[efec & sin_tar & sin_prep]
    return d if not d.empty else pd.DataFrame()


def _reactivos_vih_sin_tar(df):
    """Reactivo VIH sin inicio de TAR"""
    if 'resultado' not in df.columns or 'fecha_inicio_tar' not in df.columns:
        return pd.DataFrame()
    d = df[
        (df['resultado'].str.upper() == 'REACTIVO') &
        (df['fecha_inicio_tar'].isna() | (df['fecha_inicio_tar'] == ''))
    ]
    return d if not d.empty else pd.DataFrame()


def _en_tar_sin_tamizaje_its(df):
    """En TAR pero sin tamizaje ITS"""
    if 'uid' not in df.columns or 'fecha_tamizaje' not in df.columns:
        return pd.DataFrame()
    # Pacientes con datos TAR (condicion_vih no vacia)
    uid_cols_vih = [c for c in df.columns if 'condicion_vih' in c or 'fecha_inicio_tar' in c]
    if not uid_cols_vih:
        return pd.DataFrame()
    d = df[
        (df['condicion_vih'].notna()) &
        (df['condicion_vih'] != '') &
        (df['fecha_tamizaje'].isna())
    ]
    return d if not d.empty else pd.DataFrame()


def _vinculacion_fecha_inconsistente(df):
    """Vinculacion antes del tamizaje"""
    if 'vinculo_fecha' not in df.columns or 'fecha_tamizaje' not in df.columns:
        return pd.DataFrame()
    d = df[
        df['vinculo_fecha'].notna() &
        df['fecha_tamizaje'].notna()
    ].copy()
    d['_vf'] = pd.to_datetime(d['vinculo_fecha'], errors='coerce')
    d['_ft'] = pd.to_datetime(d['fecha_tamizaje'], errors='coerce')
    d = d[d['_vf'] < d['_ft']]
    d = d.drop(columns=['_vf', '_ft'], errors='ignore')
    return d if not d.empty else pd.DataFrame()


def _poblacion_incoherente(df):
    """Poblacion clave vs sexo/edad"""
    if 'sexo' not in df.columns or 'tipo_poblacion' not in df.columns:
        return pd.DataFrame()
    d = df.copy()
    if 'edad' in d.columns:
        d['edad'] = pd.to_numeric(d['edad'], errors='coerce')
        menor = (d['edad'] < 15) & (d['tipo_poblacion'].str.upper().str.contains('TS|TRA|TRANS', na=False))
    else:
        menor = pd.Series(False, index=d.index)
    hsh_mujer = (
        (d['tipo_poblacion'].str.upper().str.contains('HSH', na=False)) &
        (d['sexo'].str.upper().str.contains('FEMENINO|FEMENINA|MUJER', na=False))
    )
    gestante_hombre = (
        (d['tipo_poblacion'].str.upper().str.contains('GESTANTE|PG|EMBARAZADA', na=False)) &
        (d['sexo'].str.upper().str.contains('MASCULINO|HOMBRE|VARON', na=False))
    )
    mask = menor | hsh_mujer | gestante_hombre
    d = d[mask]
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
    'tamizajes_no_vih': _tamizajes_no_vih,
    'duplicados_tamizaje': _duplicados_tamizaje,
    'reactivos_sin_vinculacion': _reactivos_sin_vinculacion,
    'vinculados_sin_padron': _vinculados_sin_padron,
    'reactivos_vih_sin_tar': _reactivos_vih_sin_tar,
    'en_tar_sin_tamizaje_its': _en_tar_sin_tamizaje_its,
    'vinculacion_fecha_inconsistente': _vinculacion_fecha_inconsistente,
    'poblacion_incoherente': _poblacion_incoherente,
    'prep_sin_tamizaje_vih_reciente': _prep_sin_tamizaje_vih_reciente,
}


# ──────────────────────────────────────────────
# ANALIZADOR PRINCIPAL
# ──────────────────────────────────────────────
def _anotar_inconsistencia(d, cols, msg):
    """Marca las filas con problemas anotando descripción"""
    if d.empty:
        return d
    d = d.copy()
    d['_problema'] = msg
    return d


class AnalizadorRevisionBases:
    """Analiza la Trama Unificada (ITS + TAR + PrEP) para Revisión de Bases"""

    def __init__(self):
        self.df = None
        self.total_registros = 0
        self.resumen = []
        self.resultados = {}
        self._v_omitidas = []
        self.tiempos = {}

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

    def analizar(self):
        """Ejecuta todas las verificaciones"""
        t_total = datetime.now()
        self.resumen = []
        self.resultados = {}

        log.info(f"Iniciando Revision de Bases sobre {self.total_registros} registros")

        for idx, item in enumerate(VERIFICACIONES, 1):
            t0 = pd.Timestamp.now()
            nombre = item['n']
            fn = FUNCIONES.get(nombre)

            faltan = self._campos_existen(item['c'])
            if faltan or fn is None:
                self._v_omitidas.append(nombre)
                detalle = f"Cols faltantes: {faltan}" if faltan else f"No implementada: {nombre}"
                self.resumen.append({
                    'ID': idx, 'Verificación': item['t'],
                    'Categoría': item['cat'], 'Prioridad': item['p'],
                    'Registros': self.total_registros, 'Problemas': 'N/A',
                    '%': 'N/A', 'Estado': 'OMITIDO',
                    'Criterio': CRITERIOS.get(nombre, ''),
                    'Descripción': item['d'], 'Detalle': detalle,
                })
                log.warning(f"  OMITIDO: {nombre} -> {detalle}")
                continue

            try:
                registros = fn(self.df)
                elapsed = (pd.Timestamp.now() - t0).total_seconds()
                self.tiempos[nombre] = elapsed

                if registros is not None and len(registros) > 0:
                    cant = len(registros)
                    pct = (cant / self.total_registros) * 100
                    self.resultados[nombre] = {
                        'id': idx, 'cantidad': cant, 'descripcion': item['d'],
                        'registros': registros, 'revisados': self.total_registros,
                        'porcentaje': pct,
                    }
                    self.resumen.append({
                        'ID': idx, 'Verificación': item['t'],
                        'Categoría': item['cat'], 'Prioridad': item['p'],
                        'Registros': self.total_registros, 'Problemas': cant,
                        '%': f"{pct:.2f}%", 'Estado': 'REVISAR',
                        'Criterio': CRITERIOS.get(nombre, ''),
                        'Descripción': item['d'],
                    })
                    log.info(f"  → {cant} problemas ({pct:.2f}%) en {elapsed:.2f}s")
                else:
                    self.resumen.append({
                        'ID': idx, 'Verificación': item['t'],
                        'Categoría': item['cat'], 'Prioridad': item['p'],
                        'Registros': self.total_registros, 'Problemas': 0,
                        '%': '0.00%', 'Estado': 'SIN PROBLEMAS',
                        'Criterio': CRITERIOS.get(nombre, ''),
                        'Descripción': item['d'],
                    })
                    log.info(f"  → 0 problemas en {elapsed:.2f}s")
            except Exception as e:
                elapsed = (pd.Timestamp.now() - t0).total_seconds()
                log.error(f"Error en {nombre}: {e}")
                self.resumen.append({
                    'ID': idx, 'Verificación': item['t'],
                    'Categoría': item['cat'], 'Prioridad': item['p'],
                    'Registros': self.total_registros, 'Problemas': 'ERROR',
                    '%': 'N/A', 'Estado': 'ERROR',
                    'Criterio': CRITERIOS.get(nombre, ''),
                    'Descripción': item['d'], 'Detalle': str(e),
                })

        delta = (datetime.now() - t_total).total_seconds()
        log.info(f"Revision completa en {delta:.2f}s")
        return self

    def generar_reporte(self, nombre_base=None):
        """Genera Excel con resultados identico al formato CDD"""
        if nombre_base is None:
            nombre_base = f"Revision_Bases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        df_resumen = pd.DataFrame(self.resumen)

        with pd.ExcelWriter(nombre_base, engine='openpyxl') as writer:
            # Hoja Resumen General
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
                'Metrica': [
                    'Total Registros', 'Verificaciones Registradas',
                    'Verificaciones con Problemas', 'Verificaciones Omitidas',
                    'Total Problemas Encontrados', 'Criticos', 'Alta Prioridad',
                    'Fecha Analisis', 'Script'
                ],
                'Valor': [
                    self.total_registros, len(VERIFICACIONES),
                    sum(1 for r in self.resumen
                        if isinstance(r.get('Problemas', 0), (int, float))
                        and r.get('Problemas', 0) > 0),
                    len(self._v_omitidas),
                    total_problemas, criticos, altos,
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'Revision Bases v1.0'
                ]
            })
            df_gen.to_excel(writer, sheet_name='Resumen General', index=False)

            # Hoja Resumen Verificaciones
            df_resumen.to_excel(writer, sheet_name='Resumen Verificaciones', index=False)

            # Hojas detalladas
            for nombre, res in self.resultados.items():
                registros = res.get('registros')
                if registros is not None and isinstance(registros, pd.DataFrame) and not registros.empty:
                    # Buscar titulo legible
                    v = next((x for x in VERIFICACIONES if x['n'] == nombre), None)
                    titulo = v['t'] if v else nombre
                    safe_titulo = re.sub(r'[\\/*?:\[\]]', '_', titulo)[:27]
                    sheet_name = f"{res['id']:02d}_{safe_titulo}"
                    registros.to_excel(writer, sheet_name=sheet_name, index=False)

        log.info(f"Reporte generado: {nombre_base}")
        return nombre_base
