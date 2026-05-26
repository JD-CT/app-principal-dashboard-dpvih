#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# CSS global estilo SIHCE - importado en todos los modulos

CSS_GLOBAL = """
<style>
/* Reset base */
.block-container{padding-top:0.8rem;padding-bottom:0.5rem}
.stApp{background:#f5f7fa}

/* Header SIHCE */
.sihce-header{
    background:linear-gradient(135deg,#003B6F,#0059A3);
    padding:14px 24px;border-radius:10px;
    margin-bottom:16px;display:flex;
    justify-content:space-between;align-items:center
}
.sihce-header h1{color:#fff;font-size:20px;margin:0;font-weight:600}
.sihce-header .sub{color:rgba(255,255,255,.7);font-size:11px;margin:1px 0 0 0}
.sihce-header .info{color:rgba(255,255,255,.4);font-size:10px;text-align:right;line-height:1.3}

/* Filtros */
.filtros{
    background:#f8f9fa;padding:10px 16px;border-radius:10px;
    margin-bottom:14px;display:flex;gap:8px;flex-wrap:wrap;align-items:end
}

/* KPIs */
.kpi-row{display:flex;gap:10px;margin-bottom:14px;flex-wrap:wrap}
.kpi-card{
    background:#fff;border-radius:10px;padding:12px 14px;
    flex:1;min-width:120px;box-shadow:0 1px 3px rgba(0,0,0,.06);
    border-left:4px solid #003B6F
}
.kpi-card .kpi-label{font-size:10px;color:#888;margin:0;text-transform:uppercase;letter-spacing:.3px}
.kpi-card .kpi-value{font-size:24px;font-weight:700;color:#003B6F;margin:2px 0;line-height:1.1}
.kpi-card .kpi-sub{font-size:10px;color:#aaa}
.kpi-card.verde{border-left-color:#27ae60}
.kpi-card.verde .kpi-value{color:#27ae60}
.kpi-card.rojo{border-left-color:#e74c3c}
.kpi-card.rojo .kpi-value{color:#e74c3c}
.kpi-card.naranja{border-left-color:#f39c12}
.kpi-card.naranja .kpi-value{color:#f39c12}
.kpi-card.morado{border-left-color:#8e44ad}
.kpi-card.morado .kpi-value{color:#8e44ad}
.kpi-card.celeste{border-left-color:#3498db}
.kpi-card.celeste .kpi-value{color:#3498db}

/* Graficos */
.chart-box{
    background:#fff;border-radius:10px;padding:12px;
    box-shadow:0 1px 3px rgba(0,0,0,.06);
    border:1px solid #eee;margin-bottom:12px
}
.chart-box h4{font-size:12px;color:#003B6F;margin:0 0 8px 0;
    border-left:3px solid #003B6F;padding-left:8px;font-weight:600}

/* Footer */
.sihce-footer{font-size:10px;color:#999;text-align:center;padding:12px 0 3px 0;border-top:1px solid #eee;margin-top:10px}

/* Menu principal */
.menu-card{
    background:#fff;border-radius:12px;padding:20px;margin-bottom:14px;
    box-shadow:0 2px 8px rgba(0,0,0,.06);border:1px solid #eee;
    cursor:pointer;transition:all .2s
}
.menu-card:hover{box-shadow:0 4px 16px rgba(0,0,0,.1);border-color:#003B6F}
.menu-card .menu-icon{font-size:36px;margin-bottom:8px}
.menu-card .menu-title{font-size:16px;font-weight:600;color:#003B6F;margin:0}
.menu-card .menu-desc{font-size:11px;color:#888;margin:4px 0 0 0}

/* Alertas */
.alerta-critico{
    background:#fff5f5;border:1px solid #e74c3c;border-radius:8px;
    padding:10px 14px;margin:8px 0
}
.alerta-critico .alerta-titulo{font-weight:600;color:#e74c3c;font-size:13px}
.alerta-critico .alerta-texto{color:#666;font-size:11px;margin:2px 0 0 0}

/* Badges */
.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:600}
.badge-success{background:#d4edda;color:#155724}
.badge-danger{background:#f8d7da;color:#721c24}
.badge-warning{background:#fff3cd;color:#856404}
.badge-info{background:#d1ecf1;color:#0c5460}

/* Sidebar - EIE */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1B3A5C 0%, #0F2744 100%);
}
section[data-testid="stSidebar"] .st-emotion-cache-16idsys {
    color: #B0D4F1 !important; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600;
}
section[data-testid="stSidebar"] p { color: #E8F0FE !important; }
section[data-testid="stSidebar"] a { color: #E8F0FE !important; text-decoration: none; }
section[data-testid="stSidebar"] a:hover { color: #fff !important; }
section[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.1); }

</style>
"""
