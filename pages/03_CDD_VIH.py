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

        error_msg = None
        analizador = None
        try:
            analizador = AnalizadorCalidadDatos()
            analizador.cargar_datos(ruta_tmp)
            analizador.analizar()
        except Exception as e:
            error_msg = str(e)
            st.error(f'Error en verificaciones: {error_msg}')
            # Aun si hay error, continuamos para mostrar lo que se pueda

        if analizador is not None and hasattr(analizador, 'resumen') and analizador.resumen:
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

            # --- METRICAS RAPIDAS: solo con verificaciones que tienen problemas ---
            st.subheader('📊 Métricas rápidas')

            total = analizador.total_registros if hasattr(analizador, 'total_registros') else 0
            def _get_num(v, default=0):
                if isinstance(v, (int, float)):
                    return v
                return default

            verificaciones_con_problemas = [
                r for r in analizador.resumen
                if _get_num(r.get('Problemas', 0)) > 0
            ]
            total_problemas = sum(_get_num(r.get('Problemas', 0)) for r in analizador.resumen)

            criticos = sum(1 for r in verificaciones_con_problemas
                           if str(r.get('Prioridad', '')).lower() == 'critica')
            altos = sum(1 for r in verificaciones_con_problemas
                        if str(r.get('Prioridad', '')).lower() == 'alta')

            col1, col2, col3, col4 = st.columns(4)
            col1.metric('Total registros', f'{total:,}')
            col2.metric('Verif. con problemas', len(verificaciones_con_problemas))
            col3.metric('Críticos', criticos)
            col4.metric('Alta prioridad', altos)

            # --- DESCARGA .xlsx: construye Excel manual con datos del resumen ---
            try:
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    # Hoja 1: Resumen General
                    df_gen = pd.DataFrame({
                        'Metrica': ['Total Registros', 'Verificaciones con problemas',
                                    'Total problemas encontrados', 'Criticos', 'Alta prioridad'],
                        'Valor': [total, len(verificaciones_con_problemas),
                                  total_problemas, criticos, altos]
                    })
                    df_gen.to_excel(writer, sheet_name='Resumen General', index=False)

                    # Hoja 2: Resumen Verificaciones (todo el resultado)
                    df_resumen.to_excel(writer, sheet_name='Resumen Verificaciones', index=False)

                    # Hojas detalladas: una por cada verificacion con registros de problema
                    for idx, r in enumerate(analizador.resumen, 1):
                        registros = r.get('registros')
                        if registros is not None and isinstance(registros, pd.DataFrame) and not registros.empty:
                            sheet_name = f"{idx:02d}_{str(r.get('t', r.get('n', f'verif{idx}')))[:27]}"
                            registros.to_excel(writer, sheet_name=sheet_name, index=False)

                buf.seek(0)

                st.download_button(
                    label='📥 Descargar Excel completo (.xlsx)',
                    data=buf,
                    file_name='CDD_VIH_v3.2_completo.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    type='primary',
                    use_container_width=True,
                )

                st.success('Reporte Excel (.xlsx) listo para descargar.')
            except Exception as e:
                st.warning(f'No se pudo generar el Excel: {e}')
                # Como fallback, ofrecer CSV
                csv_data = df_resumen.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label='📥 Descargar resumen como CSV',
                    data=csv_data,
                    file_name='CDD_VIH_v3.2_resumen.csv',
                    mime='text/csv',
                )

        else:
            if not error_msg:
                st.warning('El CDD se ejecuto pero no se obtuvieron resultados.')

        # Limpiar archivo temporal
        try:
            os.unlink(ruta_tmp)
        except:
            pass
else:
    st.info('Sube un archivo Excel para comenzar el análisis de calidad de datos.')
