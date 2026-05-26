#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# PAGINA 01: DASHBOARD PrEP
# Migrado desde app_prep_dashboard.py para funcionar como page de Streamlit
# Compatible v2 y v3 del generador de indicadores

import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io, tempfile, os
from datetime import datetime

import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.colores import *
from components.css_sihce import CSS_GLOBAL
from components.kpi_card import render_kpi_row, render_kpi_doble

st.set_page_config(page_title="PrEP - SIHCE", page_icon="💊",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown(CSS_GLOBAL, unsafe_allow_html=True)

# --- HEADER ---
st.markdown(f"""
<div class="sihce-header">
  <div>
    <h1>💊 PrEP &mdash; Indicadores</h1>
    <p class="sub">Sistema de Informaci&oacute;n de Indicadores &middot; DPVIH-MINSA</p>
  </div>
  <div class="info">{datetime.now().strftime('%d/%m/%Y %H:%M')}<br>
    <a href="/" style="color:rgba(255,255,255,.6);font-size:10px">← Volver al menú</a>
  </div>
</div>
""", unsafe_allow_html=True)

# --- BOTON VOLVER ---
if st.button("← Volver al Menú Principal", type="secondary", use_container_width=False):
    st.switch_page("app_principal.py")

# --- CARGA ---
st.markdown("### 📂 Cargar indicadores PrEP")
col_archivo, col_info = st.columns([3, 1])

with col_archivo:
    archivo = st.file_uploader(
        "Selecciona el archivo Excel generado por generar_indicadores_prep.py",
        type=['xlsx'], label_visibility="collapsed"
    )

with col_info:
    st.markdown(f"""
    <div style="background:#E8F0FE;border-radius:8px;padding:8px 10px;margin-top:24px">
      <p style="font-size:10px;color:#003B6F;margin:0;line-height:1.4">
        📌 Compatible con v2 y v3<br>
        <small>Detección automática de columnas</small>
      </p>
    </div>
    """, unsafe_allow_html=True)

if archivo is None:
    st.info("⬆️ Sube un archivo Excel de indicadores PrEP para comenzar")
    st.stop()

# --- PROCESAR ---
with st.spinner("Cargando indicadores..."):

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(archivo.getbuffer())
        ruta_tmp = tmp.name

    xls = pd.ExcelFile(ruta_tmp)
    os.unlink(ruta_tmp)

    df_ind = pd.read_excel(xls, sheet_name='Indicadores')
    df_eess = pd.read_excel(xls, sheet_name='EESS') if 'EESS' in xls.sheet_names else pd.DataFrame()
    df_res = pd.read_excel(xls, sheet_name='Resumen_DIRIS') if 'Resumen_DIRIS' in xls.sheet_names else pd.DataFrame()
    df_meta = pd.read_excel(xls, sheet_name='Metadata') if 'Metadata' in xls.sheet_names else pd.DataFrame()

    # Detectar version
    is_v3 = 'I3_Activos_Mes' in df_ind.columns
    st_version = f"v{'3' if is_v3 else '2'}.x"

    col_i2 = 'I2_Inicios_Mes' if is_v3 else 'I2_Inicios'
    col_i3 = 'I3_Activos_Mes' if is_v3 else 'I3_Activos_Anio'
    col_i4 = 'I4_Seguimiento' if is_v3 else 'I4_Continuadores'
    col_i4p = 'I4_Seguimiento_Pct' if is_v3 else 'I4_Porcentaje'

    col_i5 = 'I5_Seroconversion' if is_v3 else 'I5_Seroconversiones'
    col_i5p = 'I5_Seroconversion_Pct' if is_v3 else 'I5_Porcentaje'
    col_i6 = 'I6_Descontinuacion' if is_v3 else 'I6_Descontinuaciones'
    col_i6p = 'I6_Descontinuacion_Pct' if is_v3 else 'I6_Porcentaje'
    col_i7 = 'I7_Reingreso' if is_v3 else 'I7_Reingresos'
    col_i7p = 'I7_Reingreso_Pct' if is_v3 else 'I7_Porcentaje'

    # Normalizar columnas v2/v3
    col_map = {}
    target_cols = {
        'Periodo': 'Periodo', 'Anio': 'Anio', 'DIRIS': 'DIRIS', 'Red': 'Red',
        'Sexo': 'Sexo', 'Grupo_Etareo': 'Grupo_Etareo',
        'Poblacion_Clave': 'Poblacion_Clave', 'Modalidad': 'Modalidad',
        'I2_Inicios': col_i2, 'I3_Activos': col_i3,
        'I4_Seguimiento': col_i4, 'I4_Porcentaje': col_i4p,
        'I5_Seroconversion': col_i5, 'I5_Porcentaje': col_i5p,
        'I6_Descontinuacion': col_i6, 'I6_Porcentaje': col_i6p,
        'I7_Reingreso': col_i7, 'I7_Porcentaje': col_i7p,
        'Total_Padron': 'Total_Padron' if 'Total_Padron' in df_ind.columns else None,
    }

    # Metadata
    total_padron = 0
    if not df_meta.empty:
        for _, row in df_meta.iterrows():
            k = str(row.iloc[0]).lower().strip() if len(row) > 0 else ''
            v = row.iloc[1] if len(row) > 1 else ''
            if 'total' in k and ('padron' in k or 'registro' in k):
                try:
                    total_padron = int(str(v).replace(',', '').replace('.', '').strip())
                except:
                    pass

    if total_padron == 0:
        tc = 'Total_Padron'
        if tc in df_ind.columns:
            total_padron = int(df_ind[tc].dropna().iloc[0]) if not df_ind[tc].dropna().empty else 0

    # --- CALCULOS ANUALES desde Resumen_DIRIS (personas unicas) ---
    anio = int(df_ind['Anio'].dropna().iloc[0]) if 'Anio' in df_ind.columns else 2026
    df_base = df_ind[
        (df_ind['Sexo'] == 'TODOS') &
        (df_ind['Grupo_Etareo'] == 'TODOS') &
        (df_ind['Poblacion_Clave'] == 'TODOS') &
        (df_ind['Modalidad'] == 'TODOS')
    ].copy()

    inicios_anio = int(df_base[col_i2].sum()) if not df_base.empty else 0
    activos_anio = int(df_base[col_i3].sum()) if not df_base.empty else 0

    # Resumen_DIRIS (anual con personas unicas)
    if not df_res.empty and 'DIRIS' in df_res.columns:
        df_anual = df_res
    else:
        df_anual = df_base

    # I.2 anual
    i2_anual = int(df_anual[col_i2].sum()) if col_i2 in df_anual.columns else inicios_anio
    # I.3 anual
    if is_v3 and 'I3_Activos_Anio' in df_anual.columns:
        i3_anual = int(df_anual['I3_Activos_Anio'].sum())
    elif not df_anual.empty:
        i3_anual = int(df_anual[col_i3].sum())
    else:
        i3_anual = activos_anio

    # I.4 - Nuevo + Continuador
    if is_v3 and 'I4_Nuevo' in df_anual.columns and 'I4_Continuador' in df_anual.columns:
        n4 = int(df_anual['I4_Nuevo'].sum()) + int(df_anual['I4_Continuador'].sum())
        denom_i4 = i3_anual
    else:
        n4 = int(df_anual[col_i4].sum()) if col_i4 in df_anual.columns else 0
        denom_i4 = i3_anual

    # I.5 - Seroconversion
    n5 = int(df_anual[col_i5].sum()) if col_i5 in df_anual.columns else 0
    denom_i5 = i3_anual if i3_anual > 0 else 1

    # I.6 - Descontinuacion
    n6 = int(df_anual[col_i6].sum()) if col_i6 in df_anual.columns else 0
    denom_i6 = i3_anual if i3_anual > 0 else 1

    # I.7 - Reingreso (n7 reingresos / n6 descontinuados)
    n7 = int(df_anual[col_i7].sum()) if col_i7 in df_anual.columns else 0

    # --- SIDEBAR FILTROS ---
    with st.sidebar:
        st.markdown(f"<h3 style='color:{AZUL_PRINCIPAL};font-size:15px;margin-bottom:8px'>📋 Filtros</h3>",
                    unsafe_allow_html=True)

        periodos = sorted(df_ind['Periodo'].dropna().unique().tolist())
        p_opts = ['TODOS'] + periodos
        p_sel = st.selectbox("Periodo (mes)", p_opts)

        diris_list = sorted(df_ind['DIRIS'].dropna().unique().tolist())
        d_opts = ['TODAS'] + diris_list
        d_sel = st.selectbox("DIRIS", d_opts)

        if 'Red' in df_ind.columns:
            d_filtrado = df_ind
            if d_sel != 'TODAS':
                d_filtrado = df_ind[df_ind['DIRIS'] == d_sel]
            redes = sorted(d_filtrado['Red'].dropna().unique().tolist())
            r_opts = ['TODAS'] + redes
            r_sel = st.selectbox("Red", r_opts)
        else:
            r_sel = 'TODAS'

        st.markdown("---")
        st.markdown(f"<p style='color:#888;font-size:11px;margin-bottom:4px'>Desgloses</p>",
                    unsafe_allow_html=True)

        sexos = sorted(df_ind['Sexo'].dropna().unique().tolist())
        sx_opts = ['TODOS'] + [s for s in sexos if s != 'TODOS']
        sx_sel = st.selectbox("Sexo", sx_opts)

        grupos = sorted(df_ind['Grupo_Etareo'].dropna().unique().tolist())
        ge_opts = ['TODOS'] + [g for g in grupos if g != 'TODOS']
        ge_sel = st.selectbox("Grupo Etareo", ge_opts)

        pks = sorted(df_ind['Poblacion_Clave'].dropna().unique().tolist())
        pk_opts = ['TODOS'] + [p for p in pks if p != 'TODOS']
        pk_sel = st.selectbox("Poblacion Clave", pk_opts)

        mods = sorted(df_ind['Modalidad'].dropna().unique().tolist())
        mod_opts = ['TODOS'] + [m for m in mods if m != 'TODOS']
        mod_sel = st.selectbox("Modalidad", mod_opts)

    # --- KPIs ---
    label_i4 = "I.4 Seguimiento" if is_v3 else "I.4 Continuadores"

    # EESS activas
    if not df_eess.empty:
        eess_activas = len(df_eess)
        eess_totales = eess_activas
    else:
        eess_activas = df_ind[df_ind['I2_Inicios_Mes' if is_v3 else 'I2_Inicios'] > 0]['EESS'].nunique() if 'EESS' in df_ind.columns else 0
        eess_totales = df_ind['EESS'].nunique() if 'EESS' in df_ind.columns else 0

    a_eess = eess_activas
    t_eess = max(eess_activas, eess_totales)

    kpis = [
        {'label': 'I.1 EESS activas SIHCE', 'value': f"{a_eess:,}",
         'sub': f"de {t_eess:,} EESS", 'color': 'azul'},
        {'label': 'I.2 Inician PrEP', 'value': f"{i2_anual:,}",
         'sub': f"(primer_prep_fecha en {anio})", 'color': 'azul'},
        {'label': 'I.3 Reciben PrEP', 'value': f"{i3_anual:,}",
         'sub': f"(5 condiciones activas en {anio})", 'color': 'azul'},
        {'label': label_i4, 'value': f"{round(n4/max(denom_i4,1)*100,1):.1f}%",
         'sub': f"{n4:,} / {denom_i4:,}<br><span style=\"font-size:9px;color:#999\">(nuevo y continuador)</span>",
         'color': 'verde'},
        {'label': 'I.5 Seroconversión', 'value': f"{round(n5/max(denom_i5,1)*100,2):.2f}%",
         'sub': f"{n5:,} / {denom_i5:,}", 'color': 'rojo'},
        {'label': 'I.6 Descontinuación', 'value': f"{round(n6/max(denom_i6,1)*100,1):.1f}%",
         'sub': f"{n6:,} / {denom_i6:,}", 'color': 'naranja'},
        {'label': 'I.7 Reingreso', 'value': f"{round(n7/max(n6,1)*100,1):.1f}%",
         'sub': f"{n7:,} reingresos / {n6:,} descont.", 'color': 'morado'},
    ]
    render_kpi_row(kpis)

    # --- GRAFICOS ---
    d_base = df_ind[
        (df_ind['Sexo'] == 'TODOS') &
        (df_ind['Grupo_Etareo'] == 'TODOS') &
        (df_ind['Poblacion_Clave'] == 'TODOS') &
        (df_ind['Modalidad'] == 'TODOS')
    ].copy()
    if p_sel != 'TODOS':
        d_base = d_base[d_base['Periodo'] == p_sel]
    if d_sel != 'TODAS':
        d_base = d_base[d_base['DIRIS'] == d_sel]
    if r_sel != 'TODAS':
        d_base = d_base[d_base['Red'] == r_sel]

    meses = sorted(d_base['Periodo'].unique().tolist()) if not d_base.empty else []

    if not d_base.empty:
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown(f'<div class="chart-box"><h4>I.2 Inicios de PrEP por mes ({anio})</h4>', unsafe_allow_html=True)
            g = d_base.groupby('Periodo', as_index=False)[col_i2].sum()
            c = alt.Chart(g).mark_bar(color='#003B6F').encode(
                x=alt.X('Periodo:N', sort=meses, title='', axis=alt.Axis(labelAngle=-45)),
                y=alt.Y(f'{col_i2}:Q', title='Personas'),
                tooltip=['Periodo', col_i2]
            ).properties(height=220)
            st.altair_chart(c, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col_g2:
            st.markdown(f'<div class="chart-box"><h4>I.2 Inicios por DIRIS (top 10)</h4>', unsafe_allow_html=True)
            g = d_base.groupby('DIRIS', as_index=False)[col_i2].sum().sort_values(col_i2, ascending=False).head(10)
            c = alt.Chart(g).mark_bar(color='#0059A3').encode(
                x=alt.X(f'sum({col_i2}):Q', title='Personas'),
                y=alt.Y('DIRIS:N', sort='-x', title=''),
                tooltip=['DIRIS', col_i2]
            ).properties(height=220)
            st.altair_chart(c, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # Tendencias
        col_g3, col_g4, col_g5, col_g6 = st.columns(4)
        for col_dest, col_tit, col_color, col_nums in [
            (col_g3, 'I.4 % Seguimiento', '#27ae60', [col_i3, col_i4]),
            (col_g4, 'I.5 % Seroconversión', '#e74c3c', [col_i3, col_i5]),
            (col_g5, 'I.6 % Descontinuación', '#f39c12', [col_i3, col_i6]),
            (col_g6, 'I.7 % Reingreso', '#8e44ad', [col_i6, col_i7]),
        ]:
            with col_dest:
                st.markdown(f'<div class="chart-box"><h4>{col_tit}</h4>', unsafe_allow_html=True)
                d_num, d_den = col_nums[0], col_nums[1]
                g = d_base.groupby('Periodo', as_index=False)[[d_num, d_den]].sum()
                g['Tasa'] = (g[d_den] / g[d_num].replace(0, np.nan) * 100).fillna(0)

                max_y = max(100, g['Tasa'].max() + 5) if not g.empty else 100
                c = alt.Chart(g).mark_line(point=True, color=col_color, strokeWidth=2).encode(
                    x=alt.X('Periodo:N', sort=meses, title='', axis=alt.Axis(labels=False)),
                    y=alt.Y('Tasa:Q', title='%', scale=alt.Scale(domain=[0, max_y])),
                    tooltip=['Periodo', alt.Tooltip('Tasa:Q', format='.1f')]
                ).properties(height=180)
                st.altair_chart(c, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

    # --- TABLAS ---
    with st.expander("📋 Resumen por DIRIS (anual)"):
        dr = df_res.copy() if not df_res.empty else pd.DataFrame()
        if not dr.empty and d_sel != 'TODAS':
            dr = dr[dr['DIRIS'] == d_sel]
        if not dr.empty:
            st.dataframe(dr, use_container_width=True, hide_index=True)
        else:
            st.info("Sin datos de resumen anual")

    with st.expander("📋 Detalle por filtros aplicados"):
        d_tabla = df_ind.copy()
        if p_sel != 'TODOS':
            d_tabla = d_tabla[d_tabla['Periodo'] == p_sel]
        if d_sel != 'TODAS':
            d_tabla = d_tabla[d_tabla['DIRIS'] == d_sel]
        if r_sel != 'TODAS':
            d_tabla = d_tabla[d_tabla['Red'] == r_sel]
        if ge_sel != 'TODOS':
            d_tabla = d_tabla[d_tabla['Grupo_Etareo'] == ge_sel]
        else:
            d_tabla = d_tabla[d_tabla['Grupo_Etareo'] == 'TODOS']
        if sx_sel != 'TODOS':
            d_tabla = d_tabla[d_tabla['Sexo'] == sx_sel]
        else:
            d_tabla = d_tabla[d_tabla['Sexo'] == 'TODOS']
        if pk_sel != 'TODOS':
            d_tabla = d_tabla[d_tabla['Poblacion_Clave'] == pk_sel]
        else:
            d_tabla = d_tabla[d_tabla['Poblacion_Clave'] == 'TODOS']
        d_tabla = d_tabla[d_tabla['Modalidad'] == 'TODOS']

        if not d_tabla.empty:
            cols_mostrar = [c for c in ['Periodo','DIRIS','Red','Sexo','Grupo_Etareo',
                'Poblacion_Clave','Modalidad',
                col_i2, col_i3, col_i4, col_i4p,
                col_i5, col_i5p, col_i6, col_i6p, col_i7, col_i7p] if c in d_tabla.columns]
            st.dataframe(d_tabla[cols_mostrar], use_container_width=True, hide_index=True)
        else:
            st.info("Sin datos para los filtros seleccionados")

    # Footer
    st.markdown(f"""
    <div class="sihce-footer">
      Dashboard PrEP &middot; SIHCE v{VERSION} &middot; {datetime.now().strftime('%d/%m/%Y %H:%M')} &middot;
      Archivo: {archivo.name} &middot; Padrón: {total_padron:,} &middot;
      I.2: {inicios_anio:,} &middot; I.3: {activos_anio:,}
    </div>
    """, unsafe_allow_html=True)
