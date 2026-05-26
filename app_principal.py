#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ROUTER: Navegacion principal con st.navigation

import streamlit as st


paginas = {
    "Menú principal": [
        st.Page("🏠_Menu_principal.py", title="Menú principal", default=True),
    ],
    "Módulos": [
        st.Page("pages/01_Prep_Dashboard.py", title="Dashboard PrEP"),
        st.Page("pages/02_Revision_Bases.py", title="Revisión de Bases"),
        st.Page("pages/03_CDD_VIH.py", title="Calidad del Dato VIH-TAR"),
    ],
}

pg = st.navigation(paginas, position="sidebar")
pg.run()
