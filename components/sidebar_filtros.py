#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Barra de filtros jerarquica estilo SIHCE
# Filtros anidados: Año > Periodo > DIRIS > Red > EESS

import streamlit as st
import pandas as pd
from config.colores import AZUL_PRINCIPAL, AZUL_SECUNDARIO

def sidebar_filtros_prep(df, key_prefix="prep"):
    """
    Barra lateral de filtros jerarquicos para dashboard PrEP.
    Retorna dict con: anio, periodo, diris, red, eess, sexo, grupo_etareo,
                      poblacion_clave, modalidad
    """
    with st.sidebar:
        st.markdown(f"<h3 style='color:{AZUL_PRINCIPAL};font-size:15px;margin-bottom:8px'>📋 Filtros</h3>",
                    unsafe_allow_html=True)
        
        # Año
        anios = sorted(df['Anio'].dropna().unique().tolist(), reverse=True)
        anio = st.selectbox("Año", anios, key=f"{key_prefix}_anio")
        
        # Periodo (mes)
        d_filtrado = df[df['Anio'] == anio]
        periodos = sorted(d_filtrado['Periodo'].dropna().unique().tolist())
        periodos_opts = ['TODOS'] + periodos
        periodo = st.selectbox("Periodo (mes)", periodos_opts, key=f"{key_prefix}_periodo")
        
        if periodo != 'TODOS':
            d_filtrado = d_filtrado[d_filtrado['Periodo'] == periodo]
        
        # DIRIS
        diris_list = sorted(d_filtrado['DIRIS'].dropna().unique().tolist())
        diris_opts = ['TODAS'] + diris_list
        diris = st.selectbox("DIRIS", diris_opts, key=f"{key_prefix}_diris")
        
        if diris != 'TODAS':
            d_filtrado = d_filtrado[d_filtrado['DIRIS'] == diris]
        
        # Red
        if 'Red' in d_filtrado.columns:
            redes = sorted(d_filtrado['Red'].dropna().unique().tolist())
            red_opts = ['TODAS'] + redes
            red = st.selectbox("Red", red_opts, key=f"{key_prefix}_red")
        else:
            red = 'TODAS'
        
        # EESS
        if 'EESS' in d_filtrado.columns:
            eess_list = sorted(d_filtrado['EESS'].dropna().unique().tolist())
            eess_opts = ['TODOS'] + eess_list
            eess = st.selectbox("EESS", eess_opts, key=f"{key_prefix}_eess")
        else:
            eess = 'TODOS'
        
        st.markdown("---")
        st.markdown(f"<p style='color:#888;font-size:11px;margin-bottom:4px'>Desgloses</p>",
                    unsafe_allow_html=True)
        
        # Sexo
        sexos = sorted(df['Sexo'].dropna().unique().tolist())
        sexos_opts = ['TODOS'] + [s for s in sexos if s != 'TODOS']
        sexo = st.selectbox("Sexo", sexos_opts, key=f"{key_prefix}_sexo")
        
        # Grupo Etnareo
        if 'Grupo_Etareo' in df.columns:
            ges = sorted(df['Grupo_Etareo'].dropna().unique().tolist())
            ge_opts = ['TODOS'] + [g for g in ges if g != 'TODOS']
            grupo_etareo = st.selectbox("Grupo Etareo", ge_opts, key=f"{key_prefix}_ge")
        else:
            grupo_etareo = 'TODOS'
        
        # Poblacion Clave
        if 'Poblacion_Clave' in df.columns:
            pks = sorted(df['Poblacion_Clave'].dropna().unique().tolist())
            pk_opts = ['TODOS'] + [p for p in pks if p != 'TODOS']
            poblacion_clave = st.selectbox("Poblacion Clave", pk_opts, key=f"{key_prefix}_pk")
        else:
            poblacion_clave = 'TODOS'
        
        # Modalidad
        if 'Modalidad' in df.columns:
            mods = sorted(df['Modalidad'].dropna().unique().tolist())
            mod_opts = ['TODOS'] + [m for m in mods if m != 'TODOS']
            modalidad = st.selectbox("Modalidad", mod_opts, key=f"{key_prefix}_mod")
        else:
            modalidad = 'TODOS'
        
        # Espacio
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Boton recargar (solo visual)
        st.markdown(f"""
        <div style='background:#f8f9fa;border-radius:8px;padding:8px 10px;margin-top:5px'>
          <p style='font-size:10px;color:#666;margin:0;text-align:center'>
            🔄 Filtros aplicados
          </p>
        </div>
        """, unsafe_allow_html=True)
    
    return {
        'anio': anio,
        'periodo': periodo,
        'diris': diris,
        'red': red,
        'eess': eess,
        'sexo': sexo,
        'grupo_etareo': grupo_etareo,
        'poblacion_clave': poblacion_clave,
        'modalidad': modalidad,
    }

def sidebar_filtros_revision():
    """
    Barra lateral para el modulo de Revision de Bases.
    Filtros: Anio, DIRIS, Red, EESS, Escenario
    """
    with st.sidebar:
        st.markdown(f"<h3 style='color:{AZUL_PRINCIPAL};font-size:15px;margin-bottom:8px'>🔍 Filtros Revisión</h3>",
                    unsafe_allow_html=True)
        
        anio = st.selectbox("Año", [2026, 2025, 2024], key="rev_anio")
        
        escenario = st.selectbox(
            "Escenario",
            ["TODOS", "1 - ITS con vinculación", "2 - ITS sin vinculación",
             "3 - TAR/PrEP sin tamizaje", "Reactivos críticos"],
            key="rev_escenario"
        )
        
        st.markdown("---")
        st.markdown(f"<p style='color:#888;font-size:11px;margin:0'>Ubicación</p>",
                    unsafe_allow_html=True)
        
        diris = st.selectbox("DIRIS", ["TODAS", "LIMA NORTE", "LIMA SUR", 
            "LIMA ESTE", "LIMA CENTRO", "CALLAO", "ICA"], key="rev_diris")
        
        red = st.selectbox("Red", ["TODAS", "RED 1", "RED 2", "RED 3", 
            "RED 4", "RED 5", "RED 6", "RED 7"], key="rev_red")
        
        eess = st.text_input("EESS (búsqueda)", placeholder="Nombre...", key="rev_eess")
        
        st.markdown("---")
        
        # Boton generar informe
        if st.button("📄 Generar Informe Word", type="primary", use_container_width=True):
            st.session_state['generar_informe'] = True
            st.success("✅ Informe generado (simulado)")
    
    return {
        'anio': anio,
        'escenario': escenario,
        'diris': diris,
        'red': red,
        'eess': eess,
    }
