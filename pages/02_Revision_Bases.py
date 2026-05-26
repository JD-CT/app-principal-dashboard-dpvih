# -*- coding: utf-8 -*-
# Pagina Streamlit: Revision de Bases ITS/TAR/PrEP
# Fuente: Trama Unificada (71 columnas)
# Proceso automatizado basado en "Proceso de Revision de EESS, BMU, OBC"
# Creado: 2026-05-25

import sys
import os
import tempfile
from io import BytesIO
from pathlib import Path

import streamlit as st
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from revision_bases import AnalizadorRevisionBases

st.set_page_config(
    page_title='Revisión de Bases - ITS/TAR/PrEP',
    page_icon='🔍',
    layout='wide',
)

st.markdown("""
<style>
    .main-header { background: linear-gradient(135deg, #1B3A5C, #2A5F8F); padding: 1.5rem; border-radius: 12px; margin-bottom: 2rem; }
    .main-header h1 { color: #fff; margin: 0; font-size: 1.8rem; font-weight: 600; }
    .main-header p { color: #B0D4F1; margin: 0.3rem 0 0 0; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>🔍 Revisión de Bases — ITS / TAR / PrEP</h1><p>Sube la Trama Unificada (71 columnas) para analizar calidad, brechas y vinculaciones automáticamente</p></div>', unsafe_allow_html=True)

# ──────────────────────────
# INFO DEL PROCESO
# ──────────────────────────
with st.expander("📖 ¿Cómo funciona?", expanded=False):
    st.markdown("""
    **Proceso automatizado** (reemplaza los 6 pasos manuales en Excel):

    1. **Subes** el Excel generado por la consulta **Trama Unificada v1.6** (ITS + TAR + PrEP)
    2. El sistema ejecuta **9 verificaciones automáticas**:
       - Filtra solo tamizajes VIH/Dual (descarta sífilis/hepatitis)
       - Detecta duplicados por UID + fecha
       - Identifica reactivos sin vinculación efectiva
       - Cruza vinculaciones contra padrones TAR y PrEP
       - Verifica fechas y consistencia de datos
    3. Obtienes **KPIs + tablas + descarga .xlsx** del reporte completo
    """)

# ──────────────────────────
# CARGA DE ARCHIVO
# ──────────────────────────
archivo = st.file_uploader(
    'Selecciona el archivo Trama Unificada (.xlsx)',
    type=['xlsx', 'xls'],
    help='Archivo generado por la consulta SQL Trama Unificada v1.6 (71 columnas)'
)

if archivo:
    with st.spinner('Ejecutando Revisión de Bases...'):
        sufijo = '.xlsx' if archivo.name.endswith('.xlsx') else '.xls'
        with tempfile.NamedTemporaryFile(suffix=sufijo, delete=False) as tmp:
            tmp.write(archivo.getvalue())
            ruta_tmp = tmp.name

        error_msg = None
        analizador = None
        ruta_reporte = None
        try:
            analizador = AnalizadorRevisionBases()
            analizador.cargar_datos(ruta_tmp)
            analizador.analizar()
        except Exception as e:
            error_msg = str(e)
            st.error(f'Error en verificaciones: {error_msg}')
            import traceback
            st.code(traceback.format_exc())

        if analizador is not None and hasattr(analizador, 'resumen') and analizador.resumen:
            df_resumen = pd.DataFrame(analizador.resumen)

            st.subheader('📋 Resumen de verificaciones')

            col_config = {
                'ID': st.column_config.NumberColumn('#'),
                'Verificacion': st.column_config.TextColumn('Verificación'),
                'Descripcion': st.column_config.TextColumn('Descripción'),
                'Cantidad': st.column_config.NumberColumn('Cantidad'),
                'Categoria': st.column_config.TextColumn('Categoría'),
                'Estado': st.column_config.TextColumn('Estado'),
            }

            cols_show = [c for c in ['ID', 'Verificacion', 'Descripcion', 'Categoria', 'Cantidad', 'Estado'] if c in df_resumen.columns]
            st.dataframe(
                df_resumen[cols_show],
                column_config=col_config,
                use_container_width=True,
                hide_index=True,
            )

            # ────────── METRICAS RAPIDAS ──────────
            st.subheader('📊 Métricas rápidas')

            total = analizador.total_registros if hasattr(analizador, 'total_registros') else 0

            def _get_num(v, default=0):
                if isinstance(v, (int, float)):
                    return v
                return default

            v_con_problemas = [
                r for r in analizador.resumen
                if _get_num(r.get('Cantidad', 0)) > 0
            ]

            omitidas = sum(1 for r in analizador.resumen
                           if str(r.get('Estado', '')).upper() == 'OMITIDO')
            con_error = sum(1 for r in analizador.resumen
                            if str(r.get('Estado', '')).upper() == 'ERROR')
            total_problemas = sum(
                _get_num(r.get('Cantidad', 0)) for r in analizador.resumen
            )

            col1, col2, col3, col4 = st.columns(4)
            col1.metric('Total registros', f'{total:,}')
            col2.metric('Con problemas', len(v_con_problemas))
            col3.metric('Total problemas', f'{total_problemas:,}')
            col4.metric('Omitidas / Error', f'{omitidas} / {con_error}')

            # ────────── BOTON DESCARGA ──────────
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
                    st.warning('No se pudo generar el reporte.')
            except Exception as e:
                st.warning(f'No se pudo generar el reporte: {e}')
                try:
                    buf = BytesIO()
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        df_resumen.to_excel(writer, sheet_name='Resumen Verificaciones', index=False)
                        for nombre, res in analizador.resultados.items():
                            registros = res.get('registros')
                            if registros is not None and isinstance(registros, pd.DataFrame) and not registros.empty:
                                sn = f"{res['id']:02d}_{nombre}"[:31]
                                registros.to_excel(writer, sheet_name=sn, index=False)
                    buf.seek(0)
                    st.download_button(
                        label='📥 Descargar Excel (fallback)',
                        data=buf,
                        file_name='Revision_Bases.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    )
                except:
                    csv_data = df_resumen.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label='📥 Descargar resumen CSV',
                        data=csv_data,
                        file_name='Revision_Bases.csv',
                        mime='text/csv',
                    )
        else:
            if not error_msg:
                st.warning('La revisión se ejecutó pero no se obtuvieron resultados.')

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
    st.info('Sube un archivo Excel de la Trama Unificada para comenzar la revisión.')
