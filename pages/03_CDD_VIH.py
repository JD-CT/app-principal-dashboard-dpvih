# -*- coding: utf-8 -*-
# Pagina Streamlit: Calidad del Dato VIH-TAR v3.2
# Creado: 2026-05-22 (ultima revision: 2026-05-25)
# Descripcion: Subir Excel TAR, ejecuta CDD VIH v3.2, muestra resumen y descarga Excel completo

import sys
import os
import tempfile
from io import BytesIO
from pathlib import Path

import streamlit as st
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cdd_vih_v3 import AnalizadorCalidadDatos

st.set_page_config(
    page_title='Calidad del Dato VIH-TAR',
    page_icon='🩺',
    layout='wide',
)

st.markdown("""
<style>
    .main-header { background: linear-gradient(135deg, #1B3A5C, #2A5F8F); padding: 1.5rem; border-radius: 12px; margin-bottom: 2rem; }
    .main-header h1 { color: #fff; margin: 0; font-size: 1.8rem; font-weight: 600; }
    .main-header p { color: #B0D4F1; margin: 0.3rem 0 0 0; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>🩺 Calidad del Dato VIH-TAR</h1><p>Sube el reporte Excel de TAR para ejecutar el CDD v3.2</p></div>', unsafe_allow_html=True)

archivo = st.file_uploader(
    'Selecciona archivo Excel', type=['xlsx', 'xls'],
    help='Archivo con hoja "BD" que contenga los datos TAR a evaluar'
)

if archivo:
    with st.spinner('Ejecutando CDD VIH-TAR...'):
        sufijo = '.xlsx' if archivo.name.endswith('.xlsx') else '.xls'
        with tempfile.NamedTemporaryFile(suffix=sufijo, delete=False) as tmp:
            tmp.write(archivo.getvalue())
            ruta_tmp = tmp.name

        try:
            analizador = AnalizadorCalidadDatos()
            analizador.cargar_datos(ruta_tmp)
            analizador.analizar()
            ruta_reporte = analizador.generar_reporte()

            df_resumen = pd.DataFrame(analizador.resumen)

            st.subheader('📋 Resumen de verificaciones')

            col_config = {
                'Verificación': st.column_config.TextColumn('Verificación'),
                'Detalle': st.column_config.TextColumn('Descripción'),
                'Problemas': st.column_config.NumberColumn('Problemas'),
                'Prioridad': st.column_config.TextColumn('Prioridad'),
                'Categoría': st.column_config.TextColumn('Categoría'),
            }

            cols_show = [c for c in ['Verificación', 'Detalle', 'Problemas', 'Prioridad', 'Categoría'] if c in df_resumen.columns]
            st.dataframe(
                df_resumen[cols_show],
                column_config=col_config,
                use_container_width=True,
                hide_index=True,
            )

            st.subheader('📊 Métricas rápidas')
            col1, col2, col3, col4 = st.columns(4)
            total = analizador.total_registros if hasattr(analizador, 'total_registros') else 0

            criticos = sum(
                1 for r in analizador.resumen
                if str(r.get('Prioridad', '')).lower() == 'critica'
                and r.get('Problemas', 0) > 0
            )
            altos = sum(
                1 for r in analizador.resumen
                if str(r.get('Prioridad', '')).lower() == 'alta'
                and r.get('Problemas', 0) > 0
            )

            col1.metric('Total registros', f'{total:,}')
            col2.metric('Verificaciones', len(analizador.resumen))
            col3.metric('Críticos', criticos)
            col4.metric('Alta prioridad', altos)

            # Generar Excel completo en memoria con todas las hojas
            import openpyxl
            from openpyxl.utils.dataframe import dataframe_to_rows

            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                # Hoja 1: Resumen General
                df_gen = pd.DataFrame({
                    'Métrica': ['Total Registros', 'Total Verificaciones', 'Registros con Problemas',
                                'Críticos', 'Alta Prioridad', 'Media Prioridad', 'Baja Prioridad'],
                    'Valor': [
                        total,
                        len(analizador.resumen),
                        sum(r.get('Problemas', 0) for r in analizador.resumen),
                        criticos, altos,
                        sum(1 for r in analizador.resumen
                            if str(r.get('Prioridad', '')).lower() == 'media'
                            and r.get('Problemas', 0) > 0),
                        sum(1 for r in analizador.resumen
                            if str(r.get('Prioridad', '')).lower() == 'baja'
                            and r.get('Problemas', 0) > 0),
                    ]
                })
                df_gen.to_excel(writer, sheet_name='Resumen General', index=False)

                # Hoja 2: Resumen Verificaciones (todo el resumen)
                df_resumen.to_excel(writer, sheet_name='Resumen Verificaciones', index=False)

                # Hojas detalladas: una por cada verificacion con problemas
                if hasattr(analizador, '_sheets') and analizador._sheets:
                    for nombre_meta in analizador._sheets:
                        # Buscar el registro en resumen
                        res = next((r for r in analizador.resumen if r['id'] == int(nombre_meta.split('_')[0])), None)
                        if res and not res.get('registros', pd.DataFrame()).empty:
                            df_det = res['registros']
                            sheet_name = nombre_meta[:31]  # Excel max 31 chars
                            df_det.to_excel(writer, sheet_name=sheet_name, index=False)

            buf.seek(0)

            st.download_button(
                label='📥 Descargar Excel completo (.xlsx)',
                data=buf,
                file_name=f'CDD_VIH_v3.2_completo.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
                use_container_width=True,
            )

            st.success('Reporte Excel completo listo para descargar. Incluye todas las hojas del CDD v3.2.')

        except Exception as e:
            st.error(f'Error al ejecutar CDD: {e}')
        finally:
            os.unlink(ruta_tmp)
else:
    st.info('Sube un archivo Excel para comenzar el análisis de calidad de datos.')
