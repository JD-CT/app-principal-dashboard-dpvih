#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Tarjetas KPI reutilizables estilo SIHCE

import streamlit as st

def render_kpi_row(kpis):
    """
    Renderiza una fila de tarjetas KPI.
    
    Args:
        kpis: list de dicts con keys:
            - label: str (texto arriba)
            - value: str (numero grande)
            - sub: str (texto debajo, opcional)
            - color: str (azul|verde|rojo|naranja|morado)
            - tooltip: str (opcional)
    """
    cols = st.columns(len(kpis))
    for i, kpi in enumerate(kpis):
        with cols[i]:
            color_class = kpi.get('color', 'azul')
            color_map = {
                'azul': '', 'verde': 'verde', 'rojo': 'rojo',
                'naranja': 'naranja', 'morado': 'morado', 'celeste': 'celeste'
            }
            extra = color_map.get(color_class, '')
            
            sub_html = ""
            if kpi.get('sub'):
                sub_html = f'<p class="kpi-sub">{kpi["sub"]}</p>'
            
            st.markdown(f"""
            <div class="kpi-card {extra}">
              <p class="kpi-label">{kpi.get('label', '')}</p>
              <p class="kpi-value">{kpi.get('value', '')}</p>
              {sub_html}
            </div>
            """, unsafe_allow_html=True)


def render_kpi_doble(label, valor_principal, valor_secundario,
                     color='azul', formato='{value}%', decimales=1):
    """
    KPI con valor principal grande y secundario pequeno debajo.
    Util para tasas (I.4, I.5, I.6, I.7)
    """
    color_class = color if color != 'azul' else ''
    st.markdown(f"""
    <div class="kpi-card {color_class}">
      <p class="kpi-label">{label}</p>
      <p class="kpi-value">{formato.format(value=round(valor_principal, decimales))}</p>
      <p class="kpi-sub">{valor_secundario}</p>
    </div>
    """, unsafe_allow_html=True)
