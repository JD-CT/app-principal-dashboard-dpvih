#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# MENU PRINCIPAL - SIHCE DPVIH MINSA
# Creado: 2026-05-18
# Descripcion: Punto de entrada. Muestra los modulos disponibles como tarjetas.

import streamlit as st
import sys, os
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from config.colores import AZUL_PRINCIPAL, AZUL_SECUNDARIO, VERSION
from components.css_sihce import CSS_GLOBAL

st.set_page_config(
    page_title="SIHCE - DPVIH MINSA",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown(CSS_GLOBAL, unsafe_allow_html=True)

MODULOS = [
    {
        'numero': '01',
        'icono': '💊',
        'titulo': 'Dashboard PrEP',
        'descripcion': 'Indicadores de Profilaxis Pre-Exposición (PrEP). '
                       'I.1 a I.7 con filtros jerárquicos, KPIs y gráficos de tendencia.',
        'archivo': 'pages/01_Prep_Dashboard.py',
        'color': AZUL_PRINCIPAL,
        'etiqueta': 'Activo',
        'badge': 'success'
    },
    {
        'numero': '02',
        'icono': '🔍',
        'titulo': 'Revisión de Bases',
        'descripcion': 'Proceso de revisión de calidad de datos: tamizajes ITS vs vinculación '
                       'a TAR y PrEP. 3 escenarios de análisis y generación de informes.',
        'archivo': 'pages/02_Revision_Bases.py',
        'color': AZUL_SECUNDARIO,
        'etiqueta': 'Activo',
        'badge': 'success'
    },
    {
        'numero': '03',
        'icono': '🩺',
        'titulo': 'Calidad del Dato VIH-TAR',
        'descripcion': 'CDD v3.1: Evalúa completitud, consistencia y duplicados '
                       'del reporte TAR. Sube Excel y descarga reporte completo.',
        'archivo': 'pages/03_CDD_VIH.py',
        'color': '#C0392B',
        'etiqueta': 'Nuevo',
        'badge': 'info'
    },
]

st.markdown(f"""
<div class="sihce-header">
  <div>
    <h1>🏥 SIHCE — DPVIH MINSA</h1>
    <p class="sub">Sistema de Información de Indicadores · Módulo de Monitoreo</p>
  </div>
  <div class="info">
    v{VERSION}<br>
    {datetime.now().strftime('%d/%m/%Y %H:%M')}
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<p style="font-size:13px;color:#666;margin-bottom:14px">
  Selecciona un módulo para acceder a sus indicadores y reportes:
</p>
""", unsafe_allow_html=True)

cols = st.columns(len(MODULOS))

for i, mod in enumerate(MODULOS):
    with cols[i]:
        badge_color = mod.get('badge', 'info')
        badge_html = f'<span class="badge badge-{badge_color}">{mod["etiqueta"]}</span>'
        
        st.markdown(f"""
        <div class="menu-card" id="modulo-{mod['numero']}">
          <div class="menu-icon">{mod['icono']}</div>
          <p class="menu-title">{mod['titulo']}</p>
          <p class="menu-desc">{mod['descripcion']}</p>
          <div style="margin-top:10px">{badge_html}</div>
        </div>
        """, unsafe_allow_html=True)
        
        if mod['archivo']:
            if st.button(f"Ingresar", key=f"btn_{mod['numero']}", type="primary",
                        use_container_width=True):
                st.switch_page(mod['archivo'])
        else:
            st.button("Próximamente", disabled=True, use_container_width=True)

# Footer
st.markdown("---")
col_info1, col_info2 = st.columns([1, 1])

with col_info1:
    st.markdown("""
    <div style="background:#f8f9fa;border-radius:10px;padding:14px 16px">
      <h4 style="color:#003B6F;font-size:13px;margin:0 0 8px 0">ℹ️ Acerca del Sistema</h4>
      <p style="font-size:11px;color:#666;margin:0;line-height:1.5">
        SIHCE (Sistema de Información de Indicadores) es una plataforma de monitoreo 
        desarrollada para la Dirección de Prevención y Control de VIH/SIDA (DPVIH) 
        del MINSA. Permite la visualización, análisis y revisión de indicadores 
        de los programas PrEP, TAR y tamizajes ITS.
      </p>
    </div>
    """, unsafe_allow_html=True)

with col_info2:
    st.markdown("""
    <div style="background:#f8f9fa;border-radius:10px;padding:14px 16px">
      <h4 style="color:#003B6F;font-size:13px;margin:0 0 8px 0">📁 Arquitectura</h4>
      <p style="font-size:11px;color:#666;margin:0;line-height:1.5">
        <strong>Módulos:</strong> PrEP · Revisión de Bases · CDD VIH-TAR<br>
        <strong>Datos:</strong> Excel (upload)<br>
        <strong>CDDs:</strong> VIH v3.1, ITS v2.0, PrEP v2.0
      </p>
    </div>
    """, unsafe_allow_html=True)

st.markdown(f"""
<div class="sihce-footer">
  SIHCE v{VERSION} &middot; DPVIH-MINSA &middot; {datetime.now().strftime('%d/%m/%Y %H:%M')}
</div>
""", unsafe_allow_html=True)
