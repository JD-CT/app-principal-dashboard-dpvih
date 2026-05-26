#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# MENU PRINCIPAL - EIE DPVIH MINSA
# Creado: 2026-05-18 (actualizado: 2026-05-25)
# Descripcion: Punto de entrada. Muestra los modulos disponibles como tarjetas.

import streamlit as st
import sys, os
from datetime import datetime, timezone, timedelta


def _hora_lima():
    """Retorna la fecha/hora actual en zona horaria de Lima (UTC-5)"""
    ahora = datetime.now(timezone(timedelta(hours=-5)))
    return ahora.strftime('%d/%m/%Y %H:%M')

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'streamlit_app'))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config.colores import AZUL_PRINCIPAL, AZUL_SECUNDARIO, VERSION
from components.css_sihce import CSS_GLOBAL

st.set_page_config(
    page_title="EIE - DPVIH MINSA",
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

st.markdown("""
<style>
    .main-header { background: linear-gradient(135deg, #1B3A5C, #2A5F8F); padding: 1.5rem; border-radius: 12px; margin-bottom: 2rem; }
    .main-header h1 { color: #fff; margin: 0; font-size: 1.8rem; font-weight: 600; }
    .main-header p { color: #B0D4F1; margin: 0.3rem 0 0 0; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

st.markdown(f'<div class="main-header"><h1>🏥 EIE — DPVIH MINSA</h1><p>Equipo de Información Estratégica · v{VERSION} · {_hora_lima()}</p></div>', unsafe_allow_html=True)

st.markdown("""
<p style="font-size:13px;color:#666;margin-bottom:14px">
  Menú principal
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

st.markdown(f"""
<div class="sihce-footer">
  EIE v{VERSION} &middot; DPVIH-MINSA &middot; {_hora_lima()}
</div>
""", unsafe_allow_html=True)
