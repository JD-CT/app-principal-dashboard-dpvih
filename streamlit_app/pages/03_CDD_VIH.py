# -*- coding: utf-8 -*-
# Pagina Streamlit: Calidad del Dato VIH-TAR
# Creado: 2026-05-22
# Descripcion: Subir Excel TAR, ejecuta CDD VIH v3.1, muestra resumen y descarga resultado completo

import sys
import os
import tempfile
from io import BytesIO
from pathlib import Path

import streamlit as st
import pandas as pd

# Ruta al CDD
CDD_PATH = Path('/opt/whatsapp-helpdesk-bot/CDD_vih_v3.1_limpiado.py')
sys.path.insert(0, str(CDD_PATH.parent))

import importlib.util
spec = importlib.util.spec_from_file_location('cdd_vih', CDD_PATH)
cdd = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cdd)

st.set_page_config(
    page_title='Calidad del Dato VIH-TAR',
    page_icon='🩺',
    layout='wide',
)

# CSS SIHCE
st.markdown("""
<style>
    .main-header { background: linear-gradient(135deg, #1B3A5C, #2A5F8F); padding: 1.5rem; border-radius: 12px; margin-bottom: 2rem; }
    .main-header h1 { color: #fff; margin: 0; font-size: 1.8rem; font-weight: 600; }
    .main-header p { color: #B0D4F1; margin: 0.3rem 0 0 0; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>🩺 Calidad del Dato VIH-TAR</h1><p>Sube el reporte Excel de TAR para ejecutar el CDD v3.1</p></div>', unsafe_allow_html=True)

archivo = st.file_uploader(
    'Selecciona archivo Excel', type=['xlsx', 'xls'],
    help='Archivo con hoja "BD" que contenga los datos TAR a evaluar'
)

if archivo:
    with st.spinner('Ejecutando CDD VIH-TAR...'):
        # Guardar a temporal
        sufijo = '.xlsx' if archivo.name.endswith('.xlsx') else '.xls'
        with tempfile.NamedTemporaryFile(suffix=sufijo, delete=False) as tmp:
            tmp.write(archivo.getvalue())
            ruta_tmp = tmp.name

        try:
            analizador = cdd.AnalizadorCalidadDatos()
            analizador.cargar_datos(ruta_tmp)
            analizador.analizar()
            ruta_reporte = analizador.generar_reporte()

            # Cargar resumen como DF
            df_resumen = pd.DataFrame(analizador.resumen)

            # Tabla resumen
            st.subheader('📋 Resumen de verificaciones')

            col_config = {
                'Verificación': st.column_config.TextColumn('Verificación'),
                'Descripción': st.column_config.TextColumn('Descripción'),
                'Problemas': st.column_config.NumberColumn('Problemas'),
                'Prioridad': st.column_config.TextColumn('Prioridad'),
                'Categoría': st.column_config.TextColumn('Categoría'),
            }

            cols_show = [c for c in ['Verificación', 'Descripción', 'Problemas', 'Prioridad', 'Categoría'] if c in df_resumen.columns]
            st.dataframe(
                df_resumen[cols_show],
                column_config=col_config,
                use_container_width=True,
                hide_index=True,
            )

            # Metricas rapidas
            st.subheader('📊 Métricas rápidas')
            col1, col2, col3, col4 = st.columns(4)
            total = analizador.total_registros if hasattr(analizador, 'total_registros') else 0
            criticos = len([r for r in analizador.resumen if str(r.get('Prioridad', '')).lower() == 'critica' and r.get('Problemas', 0) > 0])
            altos = len([r for r in analizador.resumen if str(r.get('Prioridad', '')).lower() == 'alta' and r.get('Problemas', 0) > 0])

            col1.metric('Total registros', f'{total:,}')
            col2.metric('Verificaciones', len(analizador.resumen))
            col3.metric('Críticos', criticos)
            col4.metric('Alta prioridad', altos)

            # Boton descarga
            with open(ruta_reporte, 'rb') as f:
                buf = BytesIO(f.read())
            st.download_button(
                label='📥 Descargar reporte Excel completo',
                data=buf,
                file_name=os.path.basename(ruta_reporte),
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
            )

            st.success(f'Reporte generado: {os.path.basename(ruta_reporte)}')

        except Exception as e:
            st.error(f'Error al ejecutar CDD: {e}')
        finally:
            os.unlink(ruta_tmp)
else:
    st.info('Sube un archivo Excel para comenzar el análisis de calidad de datos.')
