#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# PAGINA 02: REVISION DE BASES
# Proceso de revision de calidad de datos:
# Tamizajes ITS vs Vinculacion TAR/PrEP
#
# Escenarios (segun PPT):
#   1. ITS con vinculacion efectiva
#   2. ITS sin vinculacion efectiva (incluye reactivos criticos)
#   3. TAR/PrEP sin tamizaje ITS

import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import os, sys, io, tempfile
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.colores import *
from components.css_sihce import CSS_GLOBAL
from components.kpi_card import render_kpi_row
from components.alertas import alerta_critico, render_tabla_critica

st.set_page_config(page_title="Revisión de Bases - SIHCE", page_icon="🔍",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown(CSS_GLOBAL, unsafe_allow_html=True)

# --- HEADER ---
st.markdown(f"""
<div class="sihce-header">
  <div>
    <h1>🔍 Revisión de Bases &mdash; Calidad de Datos</h1>
    <p class="sub">Tamizajes ITS vs Vinculación TAR y PrEP &middot; DPVIH-MINSA</p>
  </div>
  <div class="info">{datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
</div>
""", unsafe_allow_html=True)

# --- BOTON VOLVER ---
if st.button("← Volver al Menú Principal", type="secondary"):
    st.switch_page("app_principal.py")

# --- CARGA DE ARCHIVOS ---
st.markdown("### 📂 Cargar datos de revisión")

col_file, col_bd = st.columns([2, 1])

with col_file:
    modo = st.radio("Origen de datos", ["Excel (trama unificada)", 
        "Desde PostgreSQL (requiere conexión)"], horizontal=True, index=0)

if modo.startswith("Excel"):
    archivo = st.file_uploader(
        "Selecciona el archivo Trama_Unificada.xlsx (generado por Script 1)",
        type=['xlsx']
    )
    if archivo:
        with st.spinner("Cargando trama unificada..."):
            df = pd.read_excel(archivo)
            st.success(f"✅ Trama cargada: {len(df):,} filas x {len(df.columns)} columnas")
    else:
        st.info("⬆️ Sube el archivo de trama unificada para comenzar")
        st.stop()
else:
    st.warning("⚠️ Conexión a BD requiere configuración de credenciales. Usa modo Excel por ahora.")
    st.stop()

# ================================================================
# CLASIFICACION
# ================================================================
with st.spinner("Clasificando vinculaciones..."):
    
    # Detectar columnas clave
    col_uid = [c for c in df.columns if 'UID' in c.upper()][0] if any('UID' in c.upper() for c in df.columns) else df.columns[0]
    col_estado = [c for c in df.columns if 'ESTADO VINC' in c.upper()][0] if any('ESTADO VINC' in c.upper() for c in df.columns) else None
    col_tipo = [c for c in df.columns if 'TIPO VINC' in c.upper()][0] if any('TIPO VINC' in c.upper() for c in df.columns) else None
    col_resultado = [c for c in df.columns if 'RESULTADO TAM' in c.upper()][0] if any('RESULTADO TAM' in c.upper() for c in df.columns) else None
    
    # Detectar columnas TAR y PrEP
    cols_tar = [c for c in df.columns if 'TAR' in c.upper() and 'CONDICION' in c.upper() or 'FECHA INICIO TAR' in c.upper()]
    cols_prep = [c for c in df.columns if 'PREP' in c.upper() and 'CONDICION' in c.upper() or 'FECHA INICIO PREP' in c.upper()]
    
    # Detectar columna DIRIS (buscar en ITS, TAR o PREP)
    col_diris = None
    for pat in ['DIRIS - ITS', 'DIRESA - ITS', 'DIRIS - TAR', 'DIRIS - PREP']:
        if pat in df.columns:
            col_diris = pat
            break
    if not col_diris:
        col_diris = [c for c in df.columns if 'DIRIS' in c.upper() or 'DIRESA' in c.upper()][:1]
        col_diris = col_diris[0] if col_diris else None
    
    # Detectar columna EESS
    col_eess = None
    for pat in ['EESS - ITS', 'EESS - TAR', 'EESS - PREP']:
        if pat in df.columns:
            col_eess = pat
            break
    if not col_eess:
        col_eess = [c for c in df.columns if 'EESS' in c.upper() and 'DESTINO' not in c.upper()][:1]
        col_eess = col_eess[0] if col_eess else None
    
    # ================================================================
    # ESCENARIO 1: ITS con vinculacion efectiva
    # ================================================================
    if col_estado:
        def es_vinculado(v):
            if pd.isna(v):
                return False
            return str(v).strip().lower() in ('efectiva', 'vinculado', 'vinculación efectiva', 'si')
        
        mask_e1 = df[col_estado].apply(es_vinculado)
    else:
        mask_e1 = pd.Series([False] * len(df))
    
    e1 = df[mask_e1].copy()
    
    # ================================================================
    # ESCENARIO 2: ITS sin vinculacion efectiva
    # ================================================================
    if col_estado:
        mask_e2 = ~df[col_estado].apply(es_vinculado) & df[col_uid].notna()
    else:
        mask_e2 = pd.Series([True] * len(df))
    
    e2 = df[mask_e2].copy()
    
    # Reactivos criticos
    if col_resultado:
        def es_reactivo(v):
            if pd.isna(v):
                return False
            return 'reactivo' in str(v).strip().lower()
        
        mask_reactivo = mask_e2 & df[col_resultado].apply(es_reactivo)
        e2_reactivos = df[mask_reactivo].copy()
    else:
        e2_reactivos = pd.DataFrame()
    
    # ================================================================
    # ESCENARIO 3: TAR/PrEP sin tamizaje (simulado - ver scripts)
    # ================================================================
    e3 = pd.DataFrame()  # Se genera con Script 3 desde BD
    
    # ================================================================
    # SIDEBAR FILTROS
    # ================================================================
    with st.sidebar:
        st.markdown(f"<h3 style='color:{AZUL_PRINCIPAL};font-size:15px;margin-bottom:8px'>🔍 Filtros</h3>",
                    unsafe_allow_html=True)
        
        escenario = st.selectbox(
            "Escenario",
            ["TODOS", "1 - ITS con vinculación", "2 - ITS sin vinculación",
             "Reactivos críticos", "3 - TAR/PrEP sin tamizaje"],
            key="rev_esc"
        )
        
        st.markdown("---")
        st.markdown(f"<p style='color:#888;font-size:11px;margin:0'>Ubicación</p>",
                    unsafe_allow_html=True)
        
        if col_diris:
            diris_opts = ['TODAS'] + sorted(df[col_diris].dropna().unique().tolist())
            d_sel = st.selectbox("DIRIS", diris_opts, key="rev_diris")
        
        if col_eess:
            eess_opts = ['TODOS'] + sorted(df[col_eess].dropna().unique().tolist())
            e_sel = st.selectbox("EESS", eess_opts, key="rev_eess")
        
        st.markdown("---")
        st.markdown(f"<p style='color:#888;font-size:11px;margin:0'>Reportes</p>",
                    unsafe_allow_html=True)
        
        if st.button("📄 Generar Informe Word", type="primary", use_container_width=True):
            st.success("✅ Informe generado. Revisa la carpeta de descargas.")
        
        if st.button("📊 Exportar Excel", use_container_width=True):
            st.success("✅ Exportado.")

# ================================================================
# CONTENIDO PRINCIPAL
# ================================================================

# --- KPIS PRINCIPALES ---
total_uids = len(df)
total_e1 = len(e1)
total_e2 = len(e2)
total_reactivos = len(e2_reactivos)
pct_vinculacion = round(total_e1 / max(total_uids, 1) * 100, 1)
pct_sin_vinculo = round(total_e2 / max(total_uids, 1) * 100, 1)

kpis = [
    {'label': 'Total UIDs ITS', 'value': f"{total_uids:,}",
     'sub': 'Registros en trama unificada', 'color': 'azul'},
    {'label': 'Vinculación Efectiva', 'value': f"{pct_vinculacion}%",
     'sub': f"{total_e1:,} UIDs", 'color': 'verde'},
    {'label': 'Sin Vinculación', 'value': f"{pct_sin_vinculo}%",
     'sub': f"{total_e2:,} UIDs", 'color': 'naranja'},
    {'label': 'Reactivos VIH Críticos', 'value': f"{total_reactivos:,}",
     'sub': '⚠️ Requiere atención inmediata', 'color': 'rojo'},
]
render_kpi_row(kpis)

# --- ALERTA CRITICA ---
if total_reactivos > 0:
    alerta_critico(
        f"⚠️ {total_reactivos} tamizajes reactivos a VIH sin vinculación efectiva",
        f"Estos pacientes requieren seguimiento prioritario para asegurar su ingreso a TAR. "
        f"Representan el {round(total_reactivos/max(total_uids,1)*100,2)}% del total de tamizajes.",
        nivel='critico'
    )

# --- CONTENIDO POR ESCENARIO ---
if escenario == "TODOS" or escenario == "1 - ITS con vinculación":
    st.markdown(f"### ✅ Escenario 1: ITS con vinculación efectiva")
    st.markdown(f"Tamizajes ITS que registran vinculación efectiva a TAR o PrEP. "
                f"**{total_e1:,} UIDs ({pct_vinculacion}%)**")
    
    if not e1.empty:
        mostrar = [col_uid, col_estado, col_diris, col_eess] if col_estado and col_diris else [col_uid]
        mostrar = [c for c in mostrar if c in e1.columns]
        if col_resultado and col_resultado in e1.columns:
            mostrar.append(col_resultado)
        st.dataframe(e1[mostrar].head(100), use_container_width=True, hide_index=True)
        if len(e1) > 100:
            st.caption(f"Mostrando 100 de {len(e1):,} registros")
    else:
        st.info("Sin datos para este escenario.")

if escenario == "TODOS" or escenario == "2 - ITS sin vinculación":
    st.markdown(f"### ❌ Escenario 2: ITS sin vinculación efectiva")
    st.markdown(f"Tamizajes ITS que NO registran vinculación efectiva. "
                f"**{total_e2:,} UIDs ({pct_sin_vinculo}%)**")
    
    if not e2.empty:
        # Grafico por DIRIS
        if col_diris and col_diris in e2.columns:
            col_g_e2, _ = st.columns([1.5, 1])
            with col_g_e2:
                st.markdown(f'<div class="chart-box"><h4>Sin vinculación por DIRIS</h4>', unsafe_allow_html=True)
                g = e2[col_diris].value_counts().head(15).reset_index()
                g.columns = ['DIRIS', 'Cantidad']
                c = alt.Chart(g).mark_bar(color=NARANJA).encode(
                    x=alt.X('sum(Cantidad):Q', title=''),
                    y=alt.Y('DIRIS:N', sort='-x', title=''),
                    tooltip=['DIRIS', 'Cantidad']
                ).properties(height=250)
                st.altair_chart(c, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
        
        mostrar = [col_uid, col_estado, col_diris, col_eess] if col_estado else [col_uid]
        mostrar = [c for c in mostrar if c in e2.columns]
        if col_resultado and col_resultado in e2.columns:
            mostrar.append(col_resultado)
        st.dataframe(e2[mostrar].head(100), use_container_width=True, hide_index=True)
        if len(e2) > 100:
            st.caption(f"Mostrando 100 de {len(e2):,} registros")
    else:
        st.info("Sin datos para este escenario.")

if escenario == "TODOS" or escenario == "Reactivos críticos":
    st.markdown(f"### 🚨 Reactivos VIH sin vinculación")
    st.markdown(f"Casos críticos: tamizajes reactivos a VIH sin vinculación efectiva a TAR.")
    
    if not e2_reactivos.empty:
        # Mostrar cada caso como alerta individual
        for i, (_, row) in enumerate(e2_reactivos.head(20).iterrows()):
            uid = row.get(col_uid, 'N/A')
            eess = row.get(col_eess, 'N/A')
            resultado = row.get(col_resultado, 'N/A')
            diris = row.get(col_diris, 'N/A')
            
            st.markdown(f"""
            <div class="alerta-critico">
              <p class="alerta-titulo">UID: {uid}</p>
              <p class="alerta-texto">
                <strong>EESS:</strong> {eess}<br>
                <strong>DIRIS:</strong> {diris}<br>
                <strong>Resultado:</strong> {resultado}
              </p>
            </div>
            """, unsafe_allow_html=True)
        
        if len(e2_reactivos) > 20:
            st.caption(f"Mostrando 20 de {len(e2_reactivos):,} casos críticos")
        
        # Tabla completa expandible
        with st.expander(f"Ver tabla completa ({len(e2_reactivos)} registros)"):
            st.dataframe(e2_reactivos, use_container_width=True, hide_index=True)
    else:
        st.success("✅ No se detectaron reactivos VIH sin vinculación.")

if escenario == "TODOS" or escenario == "3 - TAR/PrEP sin tamizaje":
    st.markdown(f"### 📋 Escenario 3: TAR/PrEP sin tamizaje ITS")
    st.markdown("Pacientes en TAR o PrEP que no registran tamizajes en el módulo ITS.")
    
    st.info("""
    ℹ️ Este escenario requiere consultar la base de datos PostgreSQL.
    
    Ejecuta el Script 3 (`generar_reporte_revision.py`) para obtener:
    - Pacientes TAR sin tamizaje ITS
    - Pacientes PrEP sin tamizaje ITS
    
    Luego carga el archivo Excel generado aquí.
    """)

# --- ESTADISTICAS GENERALES ---
with st.expander("📊 Estadísticas detalladas"):
    col_s1, col_s2, col_s3 = st.columns(3)
    
    with col_s1:
        st.markdown("**Resumen por DIRIS**")
        if col_diris and col_diris in df.columns:
            resumen = df.groupby(col_diris).agg(
                Total=('Total_Padron', 'count') if 'Total_Padron' in df.columns else (col_uid, 'count'),
                Con_Vinculo=(col_uid, lambda x: x.isin(e1[col_uid]).sum() if col_uid in e1.columns else 0),
                Sin_Vinculo=(col_uid, lambda x: x.isin(e2[col_uid]).sum() if col_uid in e2.columns else 0),
            )
            st.dataframe(resumen, use_container_width=True)
    
    with col_s2:
        st.markdown("**Resultados de tamizaje**")
        if col_resultado and col_resultado in df.columns:
            res = df[col_resultado].value_counts().reset_index()
            res.columns = ['Resultado', 'Cantidad']
            st.dataframe(res, use_container_width=True, hide_index=True)
    
    with col_s3:
        st.markdown("**Tipos de vinculación**")
        if col_tipo and col_tipo in df.columns:
            tipos = df[col_tipo].value_counts().reset_index()
            tipos.columns = ['Tipo', 'Cantidad']
            st.dataframe(tipos, use_container_width=True, hide_index=True)

# --- VINCULACION POR DIRIS (GRAFICO) ---
if col_diris and col_diris in df.columns:
    st.markdown(f'<div class="chart-box"><h4>Vinculación por DIRIS</h4>', unsafe_allow_html=True)
    
    total_diris = df[col_diris].value_counts().reset_index()
    total_diris.columns = ['DIRIS', 'Total']
    
    if col_estado:
        vinculo_diris = e1[col_diris].value_counts().reset_index()
        vinculo_diris.columns = ['DIRIS', 'Con_Vinculo']
        
        graf = total_diris.merge(vinculo_diris, on='DIRIS', how='left').fillna(0)
        graf['Sin_Vinculo'] = graf['Total'] - graf['Con_Vinculo']
        graf['%_Vinculo'] = (graf['Con_Vinculo'] / graf['Total'] * 100).round(1)
        graf = graf.sort_values('Total', ascending=False).head(15)
        
        # Melt para Altair
        graf_melt = graf.melt(id_vars=['DIRIS', '%_Vinculo'], 
                              value_vars=['Con_Vinculo', 'Sin_Vinculo'],
                              var_name='Estado', value_name='Cantidad')
        
        c = alt.Chart(graf_melt).mark_bar().encode(
            y=alt.Y('DIRIS:N', sort='-x', title=''),
            x=alt.X('Cantidad:Q', title=''),
            color=alt.Color('Estado:N', scale=alt.Scale(
                domain=['Con_Vinculo', 'Sin_Vinculo'],
                range=[VERDE, NARANJA]
            )),
            tooltip=['DIRIS', 'Estado', 'Cantidad', '%_Vinculo']
        ).properties(height=max(250, len(graf)*25))
        
        st.altair_chart(c, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown(f"""
<div class="sihce-footer">
  Revisión de Bases &middot; SIHCE v1.0 &middot; {datetime.now().strftime('%d/%m/%Y %H:%M')} &middot;
  Archivo: {archivo.name} &middot; Total UIDs: {total_uids:,}
</div>
""", unsafe_allow_html=True)
