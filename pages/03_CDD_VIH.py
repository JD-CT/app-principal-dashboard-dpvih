# -*- coding: utf-8 -*-
# Pagina Streamlit: Calidad del Dato VIH-TAR v3.2
# Usa generar_reporte() del CDD para producir Excel identico al script CLI
# Creado: 2026-05-22 (ultima revision: 2026-05-25)

import sys
import os
import tempfile
from io import BytesIO
from pathlib import Path

import streamlit as st
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cdd_vih_v3 import AnalizadorCalidadDatos



st.markdown("""
<style>
    .main-header { background: linear-gradient(135deg, #1B3A5C, #2A5F8F); padding: 1.5rem; border-radius: 12px; margin-bottom: 2rem; }
    .main-header h1 { color: #fff; margin: 0; font-size: 1.8rem; font-weight: 600; }
    .main-header p { color: #B0D4F1; margin: 0.3rem 0 0 0; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>🩺 Calidad del Dato VIH-TAR</h1><p>CDD v3.2 — Sube tu Excel TAR para análisis completo</p></div>', unsafe_allow_html=True)

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

        error_msg = None
        analizador = None
        ruta_reporte = None
        try:
            analizador = AnalizadorCalidadDatos()
            analizador.cargar_datos(ruta_tmp)
            analizador.analizar()
        except Exception as e:
            error_msg = str(e)
            st.error(f'Error en verificaciones: {error_msg}')

        if analizador is not None and hasattr(analizador, 'resumen') and analizador.resumen:
            df_resumen = pd.DataFrame(analizador.resumen)

            st.subheader('📋 Resumen de verificaciones')

            col_config = {
                'Verificación': st.column_config.TextColumn('Verificación'),
                'Descripción': st.column_config.TextColumn('Descripción'),
                'Problemas': st.column_config.NumberColumn('Problemas'),
                'Prioridad': st.column_config.TextColumn('Prioridad'),
                'Categoría': st.column_config.TextColumn('Categoría'),
                'Estado': st.column_config.TextColumn('Estado'),
            }

            cols_show = [c for c in ['Verificación', 'Descripción', 'Problemas', 'Prioridad', 'Categoría', 'Estado'] if c in df_resumen.columns]
            st.dataframe(
                df_resumen[cols_show],
                column_config=col_config,
                use_container_width=True,
                hide_index=True,
            )

            # --- METRICAS RAPIDAS ---
            st.subheader('📊 Métricas rápidas')

            total = analizador.total_registros if hasattr(analizador, 'total_registros') else 0

            def _get_num(v, default=0):
                if isinstance(v, (int, float)):
                    return v
                return default

            v_con_problemas = [
                r for r in analizador.resumen
                if _get_num(r.get('Problemas', 0)) > 0
            ]

            criticos = sum(1 for r in v_con_problemas
                           if str(r.get('Prioridad', '')).lower() == 'critica')
            altos = sum(1 for r in v_con_problemas
                        if str(r.get('Prioridad', '')).lower() == 'alta')
            omitidas = sum(1 for r in analizador.resumen
                           if str(r.get('Estado', '')).upper() == 'OMITIDO')
            con_error = sum(1 for r in analizador.resumen
                            if str(r.get('Estado', '')).upper() == 'ERROR')

            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric('Total registros', f'{total:,}')
            col2.metric('Con problemas', len(v_con_problemas))
            col3.metric('Críticos', criticos)
            col4.metric('Alta prioridad', altos)
            col5.metric('Omitidas / Error', f'{omitidas} / {con_error}')

            # --- DESCARGA EXCEL COMPLETO usando generar_reporte() ---
            try:
                ruta_reporte = analizador.generar_reporte()
                if ruta_reporte and os.path.exists(ruta_reporte):
                    with open(ruta_reporte, 'rb') as f:
                        excel_bytes = f.read()
                    st.download_button(
                        label='📥 Descargar Excel completo (.xlsx)',
                        data=excel_bytes,
                        file_name=os.path.basename(ruta_reporte),
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        type='primary',
                        use_container_width=True,
                    )
                    st.success(f'✅ Reporte generado: {os.path.basename(ruta_reporte)}')
                else:
                    st.warning('No se pudo obtener la ruta del reporte')
            except Exception as e:
                st.warning(f'No se pudo generar el reporte: {e}')
                # Fallback: construir manual
                try:
                    buf = BytesIO()
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        df_resumen.to_excel(writer, sheet_name='Resumen Verificaciones', index=False)
                        # Hojas detalladas
                        for idx, r in enumerate(analizador.resumen, 1):
                            registros = r.get('registros')
                            if registros is not None and isinstance(registros, pd.DataFrame) and not registros.empty:
                                sn = f"{idx:02d}_{str(r.get('Verificación', f'v{idx}'))[:27]}"
                                registros.to_excel(writer, sheet_name=sn, index=False)
                    buf.seek(0)
                    st.download_button(
                        label='📥 Descargar Excel (fallback)',
                        data=buf,
                        file_name='CDD_VIH_v3.2_completo.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        type='primary',
                        use_container_width=True,
                    )
                except:
                    csv_data = df_resumen.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label='📥 Descargar resumen CSV',
                        data=csv_data,
                        file_name='CDD_VIH_v3.2_resumen.csv',
                        mime='text/csv',
                    )
        else:
            if not error_msg:
                st.warning('El CDD se ejecutó pero no se obtuvieron resultados.')

        # Limpiar temporales
        try:
            os.unlink(ruta_tmp)
        except:
            pass
        if ruta_reporte:
            try:
                os.unlink(ruta_reporte)
            except:
                pass
else:
    st.info('Sube un archivo Excel para comenzar el análisis de calidad de datos.')
