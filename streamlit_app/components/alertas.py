#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Componente para mostrar alertas criticas (reactivos sin vinculo, etc.)

import streamlit as st
from config.colores import ROJO, NARANJA

def alerta_critico(titulo, mensaje, nivel='critico'):
    """Alerta para casos criticos (reactivos sin vinculacion)."""
    color_borde = ROJO if nivel == 'critico' else NARANJA
    color_texto = ROJO if nivel == 'critico' else '#856404'
    bg = '#fff5f5' if nivel == 'critico' else '#fff3cd'
    
    st.markdown(f"""
    <div style="background:{bg};border:1px solid {color_borde};
                border-radius:8px;padding:10px 14px;margin:8px 0">
      <p style="font-weight:600;color:{color_texto};font-size:13px;margin:0">
        ⚠ {titulo}
      </p>
      <p style="color:#666;font-size:11px;margin:4px 0 0 0">{mensaje}</p>
    </div>
    """, unsafe_allow_html=True)

def render_tabla_critica(df, titulo, cols_mostrar=None):
    """Tabla con formato para datos criticos."""
    if df.empty:
        return
    
    st.markdown(f"**{titulo}** — {len(df)} registros")
    
    if cols_mostrar:
        cols_reales = [c for c in cols_mostrar if c in df.columns]
        st.dataframe(df[cols_reales], use_container_width=True, hide_index=True)
    else:
        st.dataframe(df, use_container_width=True, hide_index=True)
