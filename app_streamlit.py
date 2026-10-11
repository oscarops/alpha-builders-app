"""Alpha Builders — aplicación integral en un único archivo.
Requiere ejecutar MIGRACION_ALPHA.sql y configurar secrets.toml antes del despliegue.
Los módulos de interfaz solo se ejecutan desde main(); las funciones puras son comprobables.
"""
from __future__ import annotations
import base64
import copy
import csv
import datetime as dt
import hashlib
import inspect
import hmac
import io
import json
import logging
import math
import os
import re
import secrets
import smtplib
import ssl
import time
import uuid
import zipfile
from collections import defaultdict
from email.message import EmailMessage
from html import escape as html_escape
from zoneinfo import ZoneInfo
import bcrypt
import openpyxl
import pandas as pd
import streamlit as st
import extra_streamlit_components as stx
from PIL import Image, ImageOps
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.drawing.image import Image as XLImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as PDFImage, KeepTogether
from supabase import create_client

LOG = logging.getLogger("alpha")
APP_VERSION = "2026.10.10-v8-interfaz"
COOKIE_NAME = "alpha_session_v5"
COOKIE_SECONDS = 30 * 86400
TABLES = {"checklists", "inspecciones", "incidencias", "rendimientos", "trabajadores"}
MAX_PHOTOS = 12
MAX_IMAGE_BYTES = 12 * 1024 * 1024

CARGOS_DISPONIBLES = ["Residente", "Asistente", "Maestro Mayor"]

EDIFICIOS_ALPHA = [
    "Tesla",
    "Lafuente",
    "Imagine",
    "Asimov",
    "Rubik",
    "Castle Rock",
    "Musk",
    "Wolf",
    "Dablanc",
    "Thomas Edison",
    "Westinghouse",
    "Smart",
    "Magnus",
    "Sparta",
]

NOMBRES_MESES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

UNIDADES_RUBRO = {"Enlucidos": "m2", "Fijos": "m2", "Fajas": "m", "Dinteles": "m"}

RENDIMIENTOS_TEORICOS = {"Enlucidos": 0.75, "Fijos": 0.50, "Fajas": 0.30, "Dinteles": 0.40}

ACTIVIDADES_MANANA_CLEAN = [
    "Verificación de asistencia del personal",
    "Verificación de los trabajos y la calidad",
    "Coordinación con otras especialidades",
    "Detección de incidencias",
]

ACTIVIDADES_TARDE_CLEAN = [
    "Verificación del avance físico de las actividades",
    "Corrección de incidencias",
    "Confirmación de materiales para el siguiente día",
    "Revisión del cumplimiento de la meta diaria",
]

OFICIOS_NOMINA_FORMATO = [
    "ALBAÑILES", "AYUDANTE", "SOLDADOR", "OPERADOR", "FIERRERO",
    "PINTORES", "MAESTRO SUPERVISOR", "CARPINTERO", "GUACHIMAN",
    "GYPSEROS", "ELECTRICOS", "PLOMEROS", "ALUMINIO Y VIDRIO"
]

RUROS_ROTATIVOS_FORMATO = [
    "BLINDOBARRAS", "TUBERÍAS AGUA POTABLE", "SOLSYSTECH (ELÉCTRICOS)",
    "TUBERÍAS HIDROSANITARIAS", "ESTRUSA (VIDRIOS)", "ASCENSORES INTERNACIONALES",
    "TUBERÍAS CONTRA INCENDIOS", "PISO FLOTANTE", "SOLDEINSA",
    "MÁRMOL Y PORCELÁNICO COVEÑA", "TERRAMODA (JARDINERAS)"
]

MAQUINARIAS_FORMATO = [
    "EXTENSION 110", "SIERRA CIRCULAR", "AMOLADORA", "SOLDADORA 220 + CARETA",
    "ARNES", "VIBRADOR", "TALADRO", "TIJERA"
]

APP_CSS = '\n    <style>\n    @import url(\'https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Montserrat:wght@500;600;700;800&display=swap\');\n\n    :root {\n        --bg-app: #f8fafc;\n        --text-app: #0f172a;\n        --card-bg: #ffffff;\n        --card-border: #cbd5e1;\n        --header-bg: #0f172a;\n        --header-text: #ffffff;\n        --sidebar-bg: #0b0f19;\n        --sidebar-card: #111827;\n        --sidebar-border: #1f2937;\n        --subtext: #64748b;\n        --hover-bg: #f1f5f9;\n        --table-row-even: #f8fafc;\n    }\n\n    @media (prefers-color-scheme: dark) {\n        :root {\n            --bg-app: #0f172a;\n            --text-app: #f8fafc;\n            --card-bg: #1e293b;\n            --card-border: #334155;\n            --header-bg: #020617;\n            --header-text: #ffffff;\n            --sidebar-bg: #030712;\n            --sidebar-card: #0f172a;\n            --sidebar-border: #1e293b;\n            --subtext: #94a3b8;\n            --hover-bg: #334155;\n            --table-row-even: #1e293b;\n        }\n    }\n\n    html, body, [class*="css"] { font-family: \'Plus Jakarta Sans\', sans-serif; }\n    h1, h2, h3, .brand-title { font-family: \'Montserrat\', sans-serif !important; letter-spacing: -0.03em !important; }\n\n    .block-container { \n        padding-top: 3.2rem !important; \n        padding-bottom: 1.2rem !important; \n        padding-left: 1rem !important; \n        padding-right: 1rem !important; \n        max-width: 100% !important; \n    }\n    .stApp { background-color: var(--bg-app) !important; color: var(--text-app) !important; }\n    .stApp p, .stApp label, .stApp span, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6 { color: var(--text-app); }\n    .stCaption, caption, small, [data-testid="stCaptionContainer"] { color: var(--subtext) !important; }\n\n    [data-testid="stInputInstructions"], div[data-testid="stInputInstructions"] { display: none !important; visibility: hidden !important; }\n    [data-testid="stHeader"] { background: transparent !important; z-index: 100 !important; }\n\n    /* BOTONES DE TOGGLE SIDEBAR */\n    [data-testid="stSidebarCollapseButton"] { display: block !important; visibility: visible !important; opacity: 1 !important; z-index: 999999 !important; }\n    [data-testid="collapsedControl"] { display: block !important; visibility: visible !important; opacity: 1 !important; position: fixed !important; top: 12px !important; left: 15px !important; z-index: 999999 !important; }\n\n    [data-testid="stSidebarCollapseButton"] button, [data-testid="collapsedControl"] button {\n        background-color: var(--header-bg) !important; \n        border: 1px solid var(--card-border) !important; \n        border-radius: 50% !important; \n        width: 34px !important; \n        height: 34px !important; \n        color: #ffffff !important; \n        box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important; \n        transition: all 0.2s ease !important;\n    }\n    [data-testid="stSidebarCollapseButton"] button:hover, [data-testid="collapsedControl"] button:hover {\n        background-color: #3b82f6 !important; \n        border-color: #3b82f6 !important; \n        transform: scale(1.08);\n    }\n    [data-testid="stSidebarCollapseButton"] svg, [data-testid="collapsedControl"] svg { fill: #ffffff !important; color: #ffffff !important; }\n\n    /* SIDEBAR */\n    [data-testid="stSidebar"] { \n        background-color: var(--sidebar-bg) !important; \n        border-right: 1px solid var(--sidebar-border) !important; \n        padding-top: 0px !important; \n        padding-left: 10px !important; \n        padding-right: 10px !important; \n        padding-bottom: 12px !important; \n    }\n    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0.4rem !important; padding-top: 0px !important; }\n    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] div { color: #ffffff !important; }\n\n    .sidebar-logo-card { \n        background-color: #ffffff; \n        border-radius: 10px; \n        padding: 6px 8px; \n        margin-top: 0px !important; \n        margin-bottom: 14px !important; \n        box-shadow: 0 4px 12px rgba(0,0,0,0.3); \n        width: 100% !important; \n        box-sizing: border-box; \n        text-align: center; \n        display: block; \n    }\n    [data-testid="stSidebar"] [data-testid="stImage"] { \n        width: 100% !important; \n        display: block !important; \n        margin-top: 4px !important; \n        margin-bottom: 8px !important; \n        clear: both !important; \n    }\n    [data-testid="stImage"] img { \n        border-radius: 10px !important; \n        width: 100% !important; \n        height: auto !important; \n        max-width: 100% !important; \n        object-fit: cover !important; \n        border: 1px solid var(--card-border) !important; \n        margin: 0 !important; \n        display: block !important; \n    }\n\n    .sidebar-profile-box { \n        background: var(--sidebar-card); \n        border: 1px solid var(--sidebar-border); \n        border-radius: 10px; \n        padding: 10px 8px !important; \n        text-align: center; \n        margin-top: 2px; \n        margin-bottom: 6px; \n        width: 100% !important; \n        box-shadow: 0 4px 10px rgba(0,0,0,0.3); \n        box-sizing: border-box; \n    }\n    .sidebar-user-nombres { font-size: 0.84rem; font-weight: 800; color: #ffffff !important; line-height: 1.2; }\n    .sidebar-user-apellidos { font-size: 0.80rem; font-weight: 700; color: #e2e8f0 !important; margin-bottom: 3px !important; line-height: 1.2; }\n    .sidebar-user-email { font-size: 0.65rem; color: #60a5fa !important; font-weight: 600; margin-bottom: 4px !important; word-break: break-all; }\n    .sidebar-user-cargo { \n        display: inline-block; \n        background: #1f2937 !important; \n        color: #ffffff !important; \n        border: 1px solid #374151 !important; \n        font-size: 0.58rem !important; \n        font-weight: 800 !important; \n        padding: 2px 7px !important; \n        border-radius: 12px !important; \n        text-transform: uppercase !important; \n    }\n\n    [data-testid="stSidebar"] hr { margin: 4px 0 !important; border-color: var(--sidebar-border) !important; }\n\n    /* SMART DASHBOARD GLASSMORPHISM */\n    .smart-dashboard-container {\n        background: radial-gradient(120% 120% at 50% 0%, #1e293b 0%, #0f172a 60%, #090d16 100%);\n        border: 1px solid rgba(255, 255, 255, 0.12);\n        border-radius: 14px;\n        padding: 12px 14px;\n        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);\n        color: #ffffff !important;\n        margin-top: 4px;\n        margin-bottom: 12px;\n        width: 100%;\n        box-sizing: border-box;\n    }\n    .smart-dashboard-container * { color: #ffffff !important; }\n\n    .smart-header-bar {\n        display: flex;\n        justify-content: space-between;\n        align-items: center;\n        flex-wrap: wrap;\n        gap: 8px;\n        border-bottom: 1px solid rgba(255, 255, 255, 0.08);\n        padding-bottom: 8px;\n        margin-bottom: 10px;\n    }\n    .smart-title { font-size: 1.1rem; font-weight: 800; letter-spacing: -0.02em; line-height: 1.1; }\n    .smart-user-sub {\n        font-size: 0.78rem;\n        color: #94a3b8 !important;\n        margin-top: 3px;\n        display: flex;\n        align-items: center;\n        flex-wrap: wrap;\n        gap: 6px;\n    }\n\n    .edificio-tag-badge {\n        display: inline-flex;\n        align-items: center;\n        background: rgba(59, 130, 246, 0.20);\n        border: 1px solid rgba(96, 165, 250, 0.45);\n        color: #93c5fd !important;\n        font-size: 0.64rem;\n        font-weight: 800;\n        padding: 2px 7px;\n        border-radius: 6px;\n        letter-spacing: 0.02em;\n    }\n\n    .smart-pill {\n        display: inline-flex;\n        align-items: center;\n        gap: 4px;\n        background: rgba(255, 255, 255, 0.06);\n        border: 1px solid rgba(255, 255, 255, 0.12);\n        backdrop-filter: blur(8px);\n        padding: 3px 8px;\n        border-radius: 16px;\n        font-size: 0.70rem;\n        font-weight: 700;\n    }\n\n    .widgets-grid {\n        display: grid;\n        grid-template-columns: repeat(auto-fit, minmax(135px, 1fr));\n        gap: 8px;\n    }\n\n    .widget-glass-card {\n        background: rgba(255, 255, 255, 0.04);\n        border: 1px solid rgba(255, 255, 255, 0.08);\n        border-radius: 10px;\n        padding: 8px 10px;\n        backdrop-filter: blur(10px);\n        display: flex;\n        flex-direction: column;\n        justify-content: space-between;\n        box-sizing: border-box;\n    }\n\n    .w-card-title {\n        font-size: 0.65rem;\n        font-weight: 800;\n        text-transform: uppercase;\n        color: #94a3b8 !important;\n        letter-spacing: 0.04em;\n        display: flex;\n        justify-content: space-between;\n        align-items: center;\n        margin-bottom: 4px;\n    }\n\n    .milestone-item {\n        display: flex;\n        justify-content: space-between;\n        align-items: center;\n        background: rgba(255, 255, 255, 0.03);\n        border: 1px solid rgba(255, 255, 255, 0.05);\n        border-radius: 6px;\n        padding: 4px 6px;\n        margin-bottom: 3px;\n    }\n    .milestone-name { font-size: 0.68rem; font-weight: 700; color: #e2e8f0 !important; }\n    .milestone-status-done {\n        background: #10b981;\n        color: #ffffff !important;\n        font-size: 0.58rem;\n        font-weight: 800;\n        padding: 1px 6px;\n        border-radius: 10px;\n    }\n    .milestone-status-pending {\n        background: rgba(239, 68, 68, 0.2);\n        border: 1px solid rgba(239, 68, 68, 0.4);\n        color: #fca5a5 !important;\n        font-size: 0.58rem;\n        font-weight: 800;\n        padding: 1px 6px;\n        border-radius: 10px;\n    }\n\n    /* DONA SVG */\n    .donut-container { display: flex; align-items: center; gap: 8px; }\n    .donut-chart-svg { width: 44px; height: 44px; transform: rotate(-90deg); flex-shrink: 0; }\n    .donut-bg { fill: none; stroke: #334155 !important; stroke-width: 4.5; }\n    .donut-progress { fill: none; stroke-width: 4.5; stroke-linecap: round; }\n    .donut-info-val { font-size: 0.98rem; font-weight: 900; line-height: 1; }\n    .donut-info-lbl { font-size: 0.60rem; color: #94a3b8 !important; font-weight: 600; margin-top: 1px; }\n    .stat-hero-number { font-size: 1.35rem; font-weight: 900; line-height: 1; margin-top: 2px; margin-bottom: 2px; }\n\n    /* PESTAÑAS */\n    .stTabs [data-baseweb="tab-list"] { \n        gap: 6px !important; \n        background-color: var(--card-border) !important; \n        padding: 4px !important; \n        border-radius: 10px !important; \n        border: 1px solid var(--card-border) !important; \n    }\n    .stTabs [data-baseweb="tab"] { \n        border-radius: 8px !important; \n        padding: 6px 14px !important; \n        background-color: var(--card-bg) !important; \n        border: 1px solid var(--card-border) !important; \n        box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important; \n        transition: all 0.2s ease !important; \n    }\n    .stTabs [data-baseweb="tab"]:hover { border-color: var(--subtext) !important; background-color: var(--hover-bg) !important; }\n    .stTabs [data-baseweb="tab"] p, .stTabs [data-baseweb="tab"] span { color: var(--text-app) !important; font-weight: 700 !important; font-size: 0.78rem !important; }\n    .stTabs [aria-selected="true"] { \n        background-color: var(--header-bg) !important; \n        border: 1px solid var(--header-bg) !important; \n        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.25) !important; \n    }\n    .stTabs [aria-selected="true"] p, .stTabs [aria-selected="true"] span, .stTabs [aria-selected="true"] div { \n        color: var(--header-text) !important; \n        font-weight: 800 !important; \n    }\n\n    .stButton > button { \n        background-color: var(--header-bg) !important; \n        color: #ffffff !important; \n        border-radius: 980px !important; \n        border: 1px solid var(--card-border) !important; \n        font-weight: 700 !important; \n        padding: 5px 12px !important; \n        font-size: 0.78rem !important; \n    }\n    .stButton > button p, .stButton > button span { color: #ffffff !important; }\n\n    .banner-item-header { \n        background-color: var(--header-bg) !important; \n        border: 1px solid var(--card-border) !important; \n        border-bottom: 2px solid #ffffff !important; \n        border-radius: 6px 6px 0 0 !important; \n        padding: 6px 10px !important; \n        margin-top: 4px !important; \n        margin-bottom: 0px !important; \n    }\n    .banner-item-header span { color: #ffffff !important; font-size: 0.80rem !important; font-weight: 800 !important; }\n\n    .card-item-body-compact { \n        border: 1px solid var(--card-border); \n        border-top: none; \n        border-radius: 0 0 6px 6px; \n        padding: 6px 8px 2px 8px; \n        background-color: var(--card-bg); \n        margin-bottom: 6px; \n    }\n\n    .worker-card-row { \n        background: var(--card-bg); \n        border: 1px solid var(--card-border); \n        border-radius: 8px; \n        padding: 8px 12px; \n        margin-bottom: 6px; \n        display: flex; \n        align-items: center; \n        justify-content: space-between; \n        gap: 8px; \n        box-shadow: 0 1px 3px rgba(0,0,0,0.03); \n    }\n    .worker-info-block { display: flex; flex-direction: column; flex: 1; min-width: 0; }\n    .worker-name-title { font-size: 0.82rem; font-weight: 800; color: var(--text-app); line-height: 1.2; word-break: break-word; }\n    .worker-meta-tags { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 4px; align-items: center; }\n    .tag-cargo-chip { \n        background: var(--card-border); \n        color: var(--text-app); \n        font-size: 0.68rem; \n        font-weight: 800; \n        padding: 2px 8px; \n        border-radius: 12px; \n        border: 1px solid var(--card-border); \n    }\n    .tag-edif-chip { \n        background: rgba(59, 130, 246, 0.15); \n        color: #3b82f6; \n        font-size: 0.68rem; \n        font-weight: 800; \n        padding: 2px 8px; \n        border-radius: 12px; \n        border: 1px solid rgba(59, 130, 246, 0.35); \n    }\n\n    @media (max-width: 768px) { \n        .block-container { padding-top: 3.6rem !important; padding-left: 0.4rem !important; padding-right: 0.4rem !important; }\n        .widgets-grid { grid-template-columns: 1fr 1fr; }\n        .smart-title { font-size: 0.95rem !important; }\n        .worker-card-row { padding: 6px 8px; }\n        .worker-name-title { font-size: 0.78rem; }\n    }\n\n    .incidencias-table { \n        width: 100%; \n        border-collapse: collapse !important; \n        margin-top: 4px; \n        margin-bottom: 8px; \n    }\n    .incidencias-table th { \n        background-color: var(--header-bg) !important; \n        color: #ffffff !important; \n        padding: 6px 8px !important; \n        font-size: 0.75rem !important; \n        font-weight: 700 !important; \n        border: 1px solid var(--card-border) !important; \n        text-align: left; \n    }\n    .incidencias-table th.center, .incidencias-table td.center { text-align: center !important; }\n    .incidencias-table td { \n        padding: 5px 7px !important; \n        font-size: 0.76rem !important; \n        border: 1px solid var(--card-border) !important; \n        vertical-align: middle !important; \n        background-color: var(--card-bg) !important; \n        color: var(--text-app) !important; \n    }\n    .incidencias-table tr:nth-child(even) td { \n        background-color: var(--table-row-even) !important; \n    }\n    #MainMenu {visibility: hidden;} footer {visibility: hidden;}\n    </style>\n    '



# Estilos complementarios 2026.10: diseño adaptativo y contraste para ambos modos.
# No modifica la lógica ni los componentes internos de Streamlit.
APP_CSS_V8 = """
<style>
:root {
  --ab-primary:#2459dc; --ab-primary-deep:#153fa5; --ab-accent:#7c55ef;
  --ab-surface:#ffffff; --ab-surface-soft:#f5f8ff;
  --ab-ink:#17233b; --ab-muted:#586781;
  --ab-outline:#dbe4f1; --ab-shadow:0 8px 28px rgba(31,54,101,.07);
  --ab-chip:#eaf1ff; --ab-chip-text:#234a94;
}
@media (prefers-color-scheme:dark){
 :root { --ab-primary:#80a7ff; --ab-primary-deep:#568af3; --ab-accent:#baa3ff;
 --ab-surface:#1e293b; --ab-surface-soft:#162237; --ab-ink:#f2f6ff;
 --ab-muted:#bbc7de; --ab-outline:#40516b; --ab-shadow:0 8px 25px rgba(0,0,0,.18);
 --ab-chip:#243b63; --ab-chip-text:#c6d8ff; }
}
/* Una app que usa tema oscuro manual respeta las variables de Streamlit si están presentes. */
.stApp { background:var(--bg-app)!important; }
.stApp [data-testid="stMainBlockContainer"] { max-width:1500px!important; margin:auto; }
.stApp [data-testid="stMainBlockContainer"] p,
.stApp [data-testid="stMainBlockContainer"] label,
.stApp [data-testid="stMainBlockContainer"] span { line-height:1.4; }
.stApp [data-testid="stMainBlockContainer"] h1,
.stApp [data-testid="stMainBlockContainer"] h2,
.stApp [data-testid="stMainBlockContainer"] h3 { color:var(--text-app)!important; }
.stApp [data-testid="stExpander"] {
 border:1px solid var(--ab-outline)!important; border-radius:13px!important;
 background:var(--ab-surface)!important; overflow:hidden!important;
 box-shadow:0 2px 7px rgba(0,0,0,.022);
}
.stApp [data-testid="stExpander"] summary { background:var(--ab-surface)!important; padding:0.65rem .8rem!important; }
.stApp [data-testid="stExpander"] summary:hover { background:var(--ab-surface-soft)!important; }
.stApp [data-testid="stExpander"] summary p,
.stApp [data-testid="stExpander"] summary span { color:var(--ab-ink)!important; font-weight:650!important; }
.stApp [data-testid="stVerticalBlockBorderWrapper"] > div[data-testid="stVerticalBlock"] { border-radius:14px; }
.stApp [data-testid="stMetric"] {
 border:1px solid var(--ab-outline); background:var(--ab-surface);
 border-radius:13px; padding:10px 13px; min-height:87px;
}
.stApp [data-testid="stMetric"] label { color:var(--ab-muted)!important; font-size:.76rem; }
.stApp [data-testid="stMetricValue"] { color:var(--ab-ink)!important; font-size:1.4rem; }
.stApp [data-testid="stDataFrame"] { border-radius:12px; overflow:hidden; }
.stApp [data-testid="stTabs"] [data-baseweb="tab-list"] {
 background:var(--ab-surface-soft)!important; border-color:var(--ab-outline)!important;
 flex-wrap:wrap!important; row-gap:4px!important;
}
.stApp [data-testid="stTabs"] [data-baseweb="tab"] { border-color:var(--ab-outline)!important; }
.stApp [data-testid="stTabs"] [data-baseweb="tab"] p { font-size:.79rem!important; }
.stApp .stButton > button[kind="primary"] { background:#2758c9!important; color:white!important; border:0!important; }
.stApp [data-testid="stTextInput"] input, .stApp [data-testid="stNumberInput"] input,
.stApp [data-testid="stTextArea"] textarea { border-radius:9px; }
.ab-heading {
 background:linear-gradient(120deg,var(--ab-surface),var(--ab-surface-soft));
 border:1px solid var(--ab-outline); border-left:4px solid var(--ab-primary);
 border-radius:14px; padding:14px 17px; margin:12px 0 16px;
}
.ab-heading .ab-eyebrow { color:var(--ab-primary)!important; font-size:.68rem; letter-spacing:.11em; font-weight:800; text-transform:uppercase; }
.ab-heading h3 { font-size:1.35rem!important; color:var(--ab-ink)!important; margin:3px 0 4px!important; }
.ab-heading p { font-size:.82rem!important; color:var(--ab-muted)!important; margin:0!important; }
.ab-day {display:flex;align-items:center;gap:9px;margin:15px 0 7px; border-bottom:1px solid var(--ab-outline);padding-bottom:7px;}
.ab-day-title {font-size:.86rem;font-weight:800;color:var(--ab-ink)!important;}
.ab-count {display:inline-flex;align-items:center;border-radius:100px; background:var(--ab-chip);color:var(--ab-chip-text)!important;font-weight:750;font-size:.67rem;padding:3px 9px;}
.ab-list-wrap {border:1px solid var(--ab-outline);border-radius:13px; overflow:hidden;background:var(--ab-surface);}
.ab-item {display:grid;grid-template-columns: minmax(0,1fr) auto;column-gap:16px;align-items:center;padding:12px 15px;border-bottom:1px solid var(--ab-outline);}
.ab-item:last-child{border-bottom:0;}
.ab-item-title {color:var(--ab-ink)!important;font-size:.83rem;font-weight:800;overflow-wrap:anywhere;}
.ab-item-sub {color:var(--ab-muted)!important;font-size:.72rem;margin-top:3px;}
.ab-item-tags {display:flex;flex-wrap:wrap;gap:5px;margin-top:5px;}
.ab-chip {display:inline-block;border:1px solid var(--ab-outline);border-radius:8px;padding:3px 8px;background:var(--ab-chip);color:var(--ab-chip-text)!important;font-size:.69rem;font-weight:700;}
.ab-chip-legacy {background:var(--ab-surface-soft);color:var(--ab-muted)!important;}
.ab-subtle {color:var(--ab-muted)!important;font-size:.77rem;}
.ab-inline-filter {display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:5px 0 10px;}
.ab-separator {height:1px;background:var(--ab-outline);margin:14px 0;}
@media(max-width:768px){
 .ab-heading{padding:12px;margin:10px 0;} .ab-heading h3{font-size:1.12rem!important;}
 .ab-item{padding:10px 12px;grid-template-columns:1fr;gap:4px;}
 .stApp [data-testid="stMetric"]{min-height:74px;padding:8px 10px;}
}
</style>
"""


def page_heading(eyebrow, title, subtitle):
    """Encabezado uniforme, evita mezclar títulos grandes y controles."""
    st.markdown(
        '<div class="ab-heading"><div class="ab-eyebrow">' + html_escape(eyebrow) + '</div>'
        '<h3>' + html_escape(title) + '</h3><p>' + html_escape(subtitle) + '</p></div>',
        unsafe_allow_html=True,
    )


def display_date(value, with_weekday=False):
    """Formato humano para mostrar, el ISO se mantiene almacenado."""
    day = date_valid(value)
    if day is None:
        return 'Fecha no especificada'
    weekdays=('lunes','martes','miércoles','jueves','viernes','sábado','domingo')
    prefix=weekdays[day.weekday()].capitalize()+', ' if with_weekday else ''
    return f'{prefix}{day.day} de {NOMBRES_MESES[day.month].lower()} de {day.year}'


def display_month(value):
    try:
        year,month=map(int,str(value)[:7].split('-'))
        return f'{NOMBRES_MESES[month]} de {year}'
    except (ValueError,KeyError):
        return 'Sin mes definido'


def record_date(table, record):
    return str(record.get('fecha') or record.get('fecha_compromiso') or '')[:10]


def record_title(table, record):
    """Nombre legible sin revelar claves internas ni IDs."""
    if table=='rendimientos':
        return str(record.get('rubro') or 'Actividad registrada') + ((' · ' + str(record.get('piso'))) if record.get('piso') else '')
    if table=='incidencias':
        return str(record.get('area') or 'Incidencia') + ' · ' + str(record.get('prioridad') or 'Media')
    if table=='checklists':
        return 'Jornada de inspección' + ((' · ' + str(record.get('hora_inicio'))) if record.get('hora_inicio') else '')
    if table=='trabajadores':
        return str(record.get('nombre') or 'Trabajador')
    if table=='inspecciones':
        return 'Libro de obra' + ((' · ' + str(record.get('residente'))) if record.get('residente') else '')
    return 'Registro de obra'


def record_note(table, record):
    if table=='rendimientos':
        value=rendimiento_values(record)
        return f"{value['avance']:.2f} {record.get('unidad') or ''} · {record.get('estado') or 'Sin evaluación'}"
    if table=='incidencias':
        return f"{record.get('estado') or 'Abierta'} · {record.get('responsable') or 'Sin responsable'}"
    if table=='trabajadores':
        return str(record.get('cargo') or 'Cargo sin registrar')
    return f"{record.get('usuario_email') or 'Responsable no indicado'}"


def day_heading(day, count):
    st.markdown(
        '<div class="ab-day"><span class="ab-day-title">📅 ' + html_escape(display_date(day,True)) +
        '</span><span class="ab-count">' + str(count) + (' registro' if count==1 else ' registros') + '</span></div>',
        unsafe_allow_html=True,
    )


def grouped_records(records,table):
    """Únicamente selecciona registros del periodo; no consulta la base de datos."""
    groups=defaultdict(list)
    for record in records:
        groups[record_date(table,record) or 'Sin fecha'].append(record)
    return sorted(groups.items(),key=lambda pair:pair[0],reverse=True)


def page_slice(records,key,size=12):
    """Paginación accesible, sin selector numérico de páginas ni IDs."""
    count=len(records)
    if count<=size:return records
    ranges=[(i,min(i+size,count)) for i in range(0,count,size)]
    index=st.selectbox('Ver bloque de registros', list(range(len(ranges))),
        format_func=lambda x:f'Registros {ranges[x][0]+1}–{ranges[x][1]} de {count}',key=key)
    start,end=ranges[index]
    return records[start:end]


def legacy_worker_backup(repo):
    """Lee sin alterar el respaldo histórico de la nómina anterior.

    Los trabajadores sin edificio válido se conservan visibles como respaldo,
    pero no pueden copiarse a la tabla sin asignación autorizada.
    """
    raw=get_config('user_trabajadores_'+repo.email)
    entries=obj(raw,[])
    if not isinstance(entries,list):return []
    cleaned=[];seen=set()
    for entry in entries:
        if not isinstance(entry,dict):continue
        owner=str(entry.get('usuario_email') or repo.email).strip().lower()
        if owner!=repo.email:continue
        name=str(entry.get('nombre') or '').strip().upper()
        role=str(entry.get('cargo') or '').strip().upper()
        building=str(entry.get('edificio') or '').strip()
        if not name:continue
        signature=(name,building.casefold())
        if signature in seen:continue
        seen.add(signature)
        cleaned.append({'nombre':name,'cargo':role or 'SIN ESPECIALIDAD','edificio':building or 'Sin edificio asignado','_legacy':True})
    return cleaned


def combined_workers(repo, building=None):
    """Trabajadores vigentes más respaldos legados, solo lectura y sin duplicados."""
    current = repo.own('trabajadores')
    seen={(str(w.get('nombre') or '').strip().upper(),str(w.get('edificio') or '').strip().casefold()) for w in current}
    for legacy in legacy_worker_backup(repo):
        ident=(legacy['nombre'],legacy['edificio'].casefold())
        if ident not in seen:
            current.append(legacy)
            seen.add(ident)
    if building is not None:
        return [w for w in current if str(w.get('edificio') or '').casefold()==str(building).casefold()]
    return current


def now_ec():
    return dt.datetime.now(ZoneInfo("America/Guayaquil"))


def json_text(value):
    return json.dumps(value, ensure_ascii=False, default=str, allow_nan=False)


def obj(value, fallback=None):
    if isinstance(value, (dict, list)):
        return copy.deepcopy(value)
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
            if isinstance(parsed, (dict, list)):
                return parsed
        except (ValueError, TypeError):
            pass
    return copy.deepcopy({} if fallback is None else fallback)


def number(value, default=0.0):
    try:
        n = float(value)
        return n if math.isfinite(n) else default
    except (ValueError, TypeError):
        return default


def norm_email(value):
    value = str(value or "").strip().lower()
    if not re.fullmatch(r"[a-z0-9.!#$&'*+/=?^_`{|}~-]+@[a-z0-9-]+(?:\.[a-z0-9-]+)+", value):
        raise ValueError("Ingrese un correo electrónico válido.")
    return value


def password_hash(password):
    salt = secrets.token_bytes(16)
    rounds = 600_000
    result = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, rounds)
    return f"pbkdf2_sha256${rounds}${salt.hex()}${result.hex()}"


def password_ok(password, stored):
    stored = str(stored or "")
    try:
        if stored.startswith("pbkdf2_sha256$"):
            _, rounds, salt, expected = stored.split("$")
            actual = hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt), int(rounds)).hex()
            return hmac.compare_digest(actual, expected)
        if stored.startswith("sha256_bcrypt$"):
            return bcrypt.checkpw(hashlib.sha256(password.encode()).hexdigest().encode(), stored[len("sha256_bcrypt$"):].encode())
    except (ValueError, TypeError):
        return False
    # No se aceptan contraseñas sin migrar ni comparaciones en texto claro.
    return False


def validate_password(password, repeated):
    if password != repeated:
        raise ValueError("Las contraseñas no coinciden.")
    if len(password) < 10 or len(password.encode()) > 512:
        raise ValueError("Use una contraseña de al menos 10 caracteres y hasta 512 bytes.")
    return password


def digest_token(raw):
    secret = str(st.secrets.get("SESSION_SECRET", ""))
    if len(secret) < 32:
        raise RuntimeError("Configure SESSION_SECRET con al menos 32 caracteres aleatorios.")
    return hmac.new(secret.encode(), str(raw).encode(), hashlib.sha256).hexdigest()


def image_b64(upload):
    if upload is None:
        return ""
    raw = upload.getvalue() if hasattr(upload, "getvalue") else upload.read()
    if len(raw) > MAX_IMAGE_BYTES:
        raise ValueError("Cada fotografía debe pesar como máximo 12 MB.")
    try:
        with Image.open(io.BytesIO(raw)) as im:
            if im.width * im.height > 40_000_000:
                raise ValueError("La fotografía supera 40 megapíxeles.")
            im = ImageOps.exif_transpose(im).convert("RGB")
            im.thumbnail((1400, 1400), Image.Resampling.LANCZOS)
            out = io.BytesIO()
            im.save(out, "JPEG", quality=78, optimize=True)
            return base64.b64encode(out.getvalue()).decode()
    except Exception as exc:
        raise ValueError("No se pudo procesar la fotografía. Use JPG, PNG o WEBP válido.") from exc


@st.cache_data(ttl=900, show_spinner=False, max_entries=64)
def image_stream(encoded):
    if not encoded:
        return None
    try:
        raw = base64.b64decode(encoded, validate=True)
        with Image.open(io.BytesIO(raw)) as im:
            im = ImageOps.exif_transpose(im).convert("RGB")
            im.thumbnail((1000, 1000))
            out = io.BytesIO(); im.save(out, "JPEG", quality=80); out.seek(0)
            return out
    except Exception:
        return None


def list_photos(item):
    photos = item.get("Fotos")
    if isinstance(photos, list) and photos:
        return [p for p in photos if isinstance(p, str) and p]
    return [item["Foto_B64"]] if item.get("Foto_B64") else []


def public_user(row):
    return {k: row.get(k) for k in ("correo", "nombres", "apellidos", "cargo", "edificios", "estado", "es_admin", "revision", "foto_b64")}


def active_user(row):
    return str(row.get("estado") or "Activo").casefold() in {"activo", "activa"}


@st.cache_resource
def database():
    url = str(st.secrets.get("SUPABASE_URL", ""))
    key = str(st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", ""))
    if not url or not key:
        raise RuntimeError("Configure SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY en los secretos del servidor.")
    digest_token("configuration-check")
    return create_client(url, key)


def all_rows(table, *, filters=None, in_filters=None, columns="*", order="id", client=None):
    client = client or database()
    rows = []; offset = 0
    while True:
        q = client.table(table).select(columns)
        for key, value in (filters or {}).items():
            q = q.eq(key, value)
        for key, values in (in_filters or {}).items():
            if not values:
                return []
            q = q.in_(key, list(values))
        result = q.order(order).range(offset, offset + 499).execute()
        batch = result.data or []
        rows.extend(batch)
        # Si llegaron menos de 500 filas, no ejecutar otra consulta vacía.
        if len(batch) < 500:
            return rows
        offset += len(batch)


@st.cache_data(ttl=30, show_spinner=False, max_entries=120)
def cached_rows(table, filter_items=(), in_items=(), columns="*", order="id"):
    return all_rows(table, filters=dict(filter_items), in_filters=dict(in_items), columns=columns, order=order)


def invalidate():
    cached_rows.clear()
    st.session_state.pop("_export_artifact", None)


@st.cache_data(ttl=180, show_spinner=False, max_entries=16)
def get_config(key):
    rows = database().table("app_config").select("value").eq("key", key).execute().data
    return rows[0]["value"] if rows else None


def set_config(key, value):
    result = database().table("app_config").upsert({"key": key, "value": str(value)}, on_conflict="key").execute()
    if not result.data:
        raise RuntimeError("No se confirmó la actualización de configuración.")
    get_config.clear()


def fetch_user(email, credentials=False):
    columns = "*" if credentials else "correo,nombres,apellidos,cargo,edificios,estado,es_admin,revision,foto_b64,must_change_password"
    rows = database().table("usuarios").select(columns).eq("correo", norm_email(email)).execute().data
    if len(rows or []) != 1:
        return None
    row = rows[0]
    row["edificios"] = obj(row.get("edificios"), [])
    if not isinstance(row["edificios"], list):
        row["edificios"] = []
    return row


def allowed_attempt(action, subject, limit=8):
    response = database().rpc("alpha_allow_attempt", {"p_action": action, "p_subject": hashlib.sha256(subject.encode()).hexdigest(), "p_limit": limit}).execute()
    return response.data is True


def send_reset(email):
    if not allowed_attempt("reset", email, 3):
        return
    user = fetch_user(email)
    if not user or not active_user(user):
        return
    raw = secrets.token_urlsafe(24)
    database().table("alpha_password_resets").upsert({"correo": email, "token_hash": digest_token(raw), "expires_at": (dt.datetime.now(dt.timezone.utc) + dt.timedelta(minutes=20)).isoformat()}, on_conflict="correo").execute()
    msg = EmailMessage()
    msg["Subject"] = "Alpha Builders — recuperación de contraseña"
    msg["From"] = str(st.secrets["SMTP_FROM"])
    msg["To"] = email
    msg.set_content(f"Tu código de recuperación es:\n\n{raw}\n\nVálido durante 20 minutos y para un solo uso. Si no lo solicitaste, ignora este mensaje.")
    host = str(st.secrets["SMTP_HOST"]); port = int(st.secrets.get("SMTP_PORT", 587))
    context = ssl.create_default_context()
    cls = smtplib.SMTP_SSL if port == 465 else smtplib.SMTP
    kwargs = {"context": context} if port == 465 else {}
    with cls(host, port, timeout=20, **kwargs) as server:
        if port != 465:
            server.starttls(context=context)
        server.login(str(st.secrets["SMTP_USER"]), str(st.secrets["SMTP_PASSWORD"]))
        server.send_message(msg)


class ConflictError(RuntimeError):
    pass


class Repository:
    """Autorización del servidor y escrituras con control de versión."""
    def __init__(self, user):
        self.user = user
        self.email = user["correo"]
        self.admin = bool(user.get("es_admin"))
        self.projects = set(user.get("edificios") or [])

    def own(self, table):
        if table not in TABLES:
            raise ValueError("Tabla no permitida.")
        return cached_rows(table, (("usuario_email", self.email),))

    def shared(self, table, email=None):
        if table not in TABLES:
            raise ValueError("Tabla no permitida.")
        field = "proyecto" if table in {"inspecciones", "incidencias"} else "edificio"
        filters = (("usuario_email", email),) if email else ()
        return cached_rows(table, filters, ((field, tuple(sorted(self.projects))),))

    def all(self, table):
        if not self.admin or table not in TABLES:
            raise PermissionError("Operación reservada al administrador.")
        return cached_rows(table)

    def _access(self, table, row, write=False):
        if self.admin or row.get("usuario_email") == self.email:
            return
        project = row.get("proyecto") if table in {"inspecciones", "incidencias"} else row.get("edificio")
        if project in self.projects and (not write or (table == "incidencias" and self.user.get("cargo") != "Maestro Mayor")):
            return
        raise PermissionError("No tiene permiso para este registro.")

    def _verify_write_session(self):
        """Leer permisos y sesión vigentes en cada escritura; no confiar en caché de UI."""
        raw = st.session_state.get("_session_raw")
        if not raw:
            raise PermissionError("La sesión ha expirado. Vuelva a iniciar sesión.")
        sessions = database().table("alpha_sessions").select("correo,expires_at").eq("token_hash", digest_token(raw)).execute().data or []
        if len(sessions) != 1 or sessions[0].get("correo") != self.email:
            raise PermissionError("Sesión revocada. Inicie sesión nuevamente.")
        expiry = dt.datetime.fromisoformat(str(sessions[0]["expires_at"]).replace("Z", "+00:00"))
        if expiry <= dt.datetime.now(dt.timezone.utc):
            raise PermissionError("La sesión ha expirado. Inicie sesión nuevamente.")
        current = fetch_user(self.email)
        if not current or not active_user(current) or current.get("must_change_password"):
            raise PermissionError("Su cuenta no tiene autorización para guardar en este momento.")
        self.user = current
        self.admin = bool(current.get("es_admin"))
        self.projects = set(current.get("edificios") or [])

    def save(self, table, payload, original=None):
        self._verify_write_session()
        if table not in TABLES:
            raise ValueError("Tabla no permitida.")
        payload = copy.deepcopy(payload)
        project_field = "proyecto" if table in {"inspecciones", "incidencias"} else "edificio"
        project = payload.get(project_field, (original or {}).get(project_field))
        previous_project = (original or {}).get(project_field)
        if project not in self.projects and not self.admin and not (original and project == previous_project and original.get("usuario_email") == self.email):
            raise PermissionError("Solicite al administrador la asignación a ese edificio.")
        if original:
            self._access(table, original, write=True)
            payload.pop("usuario_email", None)
            q = database().table(table).update(payload).eq("id", original["id"]).eq("revision", original["revision"])
            if not self.admin:
                q = q.eq("usuario_email", original["usuario_email"])
            result = q.execute()
            if not result.data:
                raise ConflictError("El registro cambió o fue eliminado en otra sesión. Su borrador sigue guardado; recargue el historial antes de volver a editar.")
        else:
            payload["usuario_email"] = self.email
            result = database().table(table).insert(payload).execute()
            if not result.data:
                raise RuntimeError("No se confirmó el guardado. Conserve el borrador y actualice el historial antes de reintentar.")
        invalidate()
        return result.data[0]

    def delete(self, table, row):
        self._verify_write_session()
        if table not in TABLES:
            raise ValueError("Tabla no permitida.")
        self._access(table, row, write=True)
        result = database().table(table).delete().eq("id", row["id"]).eq("revision", row["revision"]).execute()
        if not result.data:
            raise ConflictError("El registro cambió o ya no existe. Actualice el historial.")
        invalidate()


def notice_error(exc):
    if isinstance(exc, (ValueError, PermissionError, ConflictError)):
        st.error(str(exc))
    else:
        code = secrets.token_hex(4)
        LOG.error("Operación fallida %s (%s)", code, type(exc).__name__)
        st.error(f"La operación no pudo confirmarse. Sus campos permanecen disponibles. Revise conexión, migración y permisos. Referencia: {code}.")


def auth_user(cookie):
    raw = st.session_state.get("_session_raw") or cookie.get(COOKIE_NAME)
    if not raw:
        return None
    rows = database().table("alpha_sessions").select("correo,expires_at").eq("token_hash", digest_token(raw)).execute().data
    if not rows:
        return None
    exp = dt.datetime.fromisoformat(rows[0]["expires_at"].replace("Z", "+00:00"))
    if exp <= dt.datetime.now(dt.timezone.utc):
        return None
    user = fetch_user(rows[0]["correo"])
    if user and active_user(user):
        st.session_state["_session_raw"] = raw
        return user
    return None


def login_panel(cookie):
    st.title("Alpha Builders | Acceso")
    pending = st.session_state.get("_pending_cookie")
    if pending:
        if cookie.get(COOKIE_NAME) == pending:
            st.session_state["_session_raw"] = pending
            st.session_state.pop("_pending_cookie", None)
            st.rerun()
        cookie.set(COOKIE_NAME, pending, key="persist_login", max_age=COOKIE_SECONDS, secure=bool(st.secrets.get("COOKIE_SECURE", True)), same_site="lax")
        st.info("Conservando la sesión en este navegador…")
        if st.button("Continuar al portal"):
            st.session_state["_cookie_epoch"] = st.session_state.get("_cookie_epoch", 0) + 1
            st.rerun()
        st.stop()
    login, register, reset = st.tabs(["Iniciar sesión", "Registrarse", "Recuperar contraseña"])
    with login:
        with st.form("login"):
            email = st.text_input("Correo electrónico")
            password = st.text_input("Contraseña", type="password")
            pin = st.text_input("PIN de acceso", type="password", max_chars=4)
            remember = st.checkbox("Mantener la sesión iniciada durante 30 días")
            submit = st.form_submit_button("Entrar al portal", type="primary")
        if submit:
            try:
                email = norm_email(email)
                if not allowed_attempt("login", email):
                    raise ValueError("Demasiados intentos. Espere 15 minutos.")
                user = fetch_user(email, credentials=True)
                configured_pin = get_config("access_pin")
                if not configured_pin:
                    raise ValueError("El administrador debe configurar el PIN en Supabase.")
                valid_password = password_ok(password, user.get("password")) if user else password_ok(password, password_hash("invalid-account-padding"))
                if not user or not valid_password or not hmac.compare_digest(pin, str(configured_pin)):
                    raise ValueError("Correo, contraseña o PIN incorrectos.")
                if not active_user(user):
                    raise ValueError("Su cuenta está pendiente de aprobación o ha sido suspendida. Contacte al administrador.")
                if not str(user.get("password", "")).startswith("pbkdf2_sha256$"):
                    database().table("usuarios").update({"password": password_hash(password)}).eq("correo", email).execute()
                raw = secrets.token_urlsafe(40)
                age = COOKIE_SECONDS if remember else 12 * 3600
                database().table("alpha_sessions").insert({"token_hash": digest_token(raw), "correo": email, "expires_at": (dt.datetime.now(dt.timezone.utc) + dt.timedelta(seconds=age)).isoformat()}).execute()
                if remember:
                    st.session_state["_pending_cookie"] = raw
                else:
                    st.session_state["_session_raw"] = raw
                st.session_state["_cookie_epoch"] = st.session_state.get("_cookie_epoch", 0) + 1
                st.rerun()
            except Exception as exc:
                notice_error(exc)
    with register:
        with st.form("register"):
            names = st.text_input("Nombres")
            surname = st.text_input("Apellidos")
            email = st.text_input("Correo")
            password = st.text_input("Crear contraseña (mínimo 10 caracteres)", type="password")
            repeat = st.text_input("Repetir contraseña", type="password")
            role = st.selectbox("Cargo solicitado", CARGOS_DISPONIBLES)
            buildings = st.multiselect("Edificios solicitados", EDIFICIOS_ALPHA)
            pin = st.text_input("PIN de registro", type="password", max_chars=4)
            submit = st.form_submit_button("Solicitar registro")
        if submit:
            try:
                email = norm_email(email); validate_password(password, repeat)
                if not names.strip() or not surname.strip() or not buildings:
                    raise ValueError("Complete nombres, apellidos y al menos un edificio.")
                if not allowed_attempt("register", email, 3):
                    raise ValueError("Espere 15 minutos antes de volver a solicitar el registro.")
                current_pin = get_config("access_pin")
                if not current_pin or not hmac.compare_digest(pin, str(current_pin)):
                    raise ValueError("PIN incorrecto.")
                if fetch_user(email):
                    raise ValueError("El correo ya tiene una cuenta. Use recuperación de contraseña.")
                result = database().table("usuarios").insert({"correo": email, "nombres": names.strip(), "apellidos": surname.strip(), "password": password_hash(password), "cargo": role, "edificios": buildings, "estado": "Pendiente", "es_admin": False, "fecha_registro": now_ec().date().isoformat()}).execute()
                if not result.data:
                    raise RuntimeError("Registro no confirmado")
                invalidate(); st.success("Solicitud guardada. El administrador verificará su identidad, cargo y edificios antes de activar la cuenta.")
            except Exception as exc:
                notice_error(exc)
    with reset:
        smtp_ready = all(st.secrets.get(k) for k in ("SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD", "SMTP_FROM"))
        if not smtp_ready:
            st.info("La recuperación por correo requiere configurar SMTP. El administrador puede emitir una contraseña temporal desde el panel, tras verificar su identidad.")
        with st.form("request_reset"):
            email = st.text_input("Correo registrado")
            submit = st.form_submit_button("Enviar código de recuperación", disabled=not smtp_ready)
        if submit:
            try:
                send_reset(norm_email(email))
                st.success("Si la cuenta está activa, recibirá un código válido durante 20 minutos.")
            except Exception as exc:
                notice_error(exc)
        with st.form("confirm_reset"):
            email = st.text_input("Correo de la cuenta")
            token = st.text_input("Código recibido por correo")
            password = st.text_input("Nueva contraseña", type="password")
            repeat = st.text_input("Repetir nueva contraseña", type="password")
            submit = st.form_submit_button("Restablecer contraseña")
        if submit:
            try:
                email = norm_email(email); validate_password(password, repeat)
                if not allowed_attempt("reset_confirm", email):
                    raise ValueError("Demasiados intentos. Espere 15 minutos.")
                result = database().rpc("alpha_reset_password", {"p_email": email, "p_token_hash": digest_token(token.strip()), "p_password": password_hash(password)}).execute()
                if result.data is not True:
                    raise ValueError("Código inválido o vencido.")
                st.success("Contraseña actualizada. Las sesiones anteriores han sido revocadas.")
            except Exception as exc:
                notice_error(exc)


def parse_minutes(value, *, duration=False, empty_zero=True):
    text = str(value or "").strip()
    if not text and empty_zero:
        return 0
    if not re.fullmatch(r"\d{2}:\d{2}", text):
        raise ValueError("Escriba las horas como HH:MM; por ejemplo 07:00 o 00:30.")
    hh, mm = map(int, text.split(":"))
    if mm > 59 or hh > (99 if duration else 23):
        raise ValueError("Hora fuera del intervalo permitido.")
    return hh * 60 + mm


def time_calculation(start, end, dead="", lunch=False, meal=False, overnight=False):
    if not str(start or "").strip() and not str(end or "").strip():
        if str(dead or "").strip() not in {"", "00:00"} or lunch or meal:
            raise ValueError("Complete inicio y fin antes de descontar tiempos.")
        return 0.0
    if not str(start or "").strip() or not str(end or "").strip():
        raise ValueError("Complete inicio y fin, o deje ambos vacíos para registrar un avance sin tiempo.")
    a = parse_minutes(start, empty_zero=False); b = parse_minutes(end, empty_zero=False)
    if overnight:
        b += 1440
    elif b < a:
        raise ValueError("La hora final es anterior a la inicial. Confirme si termina al día siguiente.")
    minutes = b - a
    if minutes > 1440:
        raise ValueError("La jornada no puede superar 24 horas.")
    deductions = parse_minutes(dead, duration=True) + (15 if lunch else 0) + (60 if meal else 0)
    if deductions > minutes:
        raise ValueError("Los descuentos superan la duración de la jornada.")
    return (minutes - deductions) / 60


def rendimiento_payload(record):
    data = obj(record.get("datos"), {})
    if isinstance(data, dict) and (data.get("_rend_v4") or data.get("_rend_v5")):
        return data
    previous = obj(record.get("intervalo"), {})
    if isinstance(previous, dict) and (previous.get("_rend_v4") or previous.get("_rend_v5")):
        return previous
    fallback={"_legacy": True, "avance_manana": number(record.get("avance")), "avance_mediodia": 0.0, "avance_tarde": 0.0,
              "edificio": record.get("edificio") or "Sin Edificio Asignado", "piso": record.get("piso") or "", "trabajadores": [], "comentarios": ""}
    fallback.update(data if isinstance(data,dict) else {"_legacy_source":data})
    return fallback


def rendimiento_values(record):
    data = rendimiento_payload(record)
    building = data.get("edificio") or record.get("edificio") or "Sin Edificio Asignado"
    advance = sum(number(data.get(k)) for k in ("avance_manana", "avance_mediodia", "avance_tarde"))
    if data.get("_rend_v5"):
        elapsed = number(data.get("horas_jornada"))
        hh = number(record.get("horas_hh"))
        consumption = hh / advance if advance > 0 and hh > 0 else None
        productivity = advance / elapsed if elapsed > 0 else None
    else:
        # No se atribuye a HH una duración histórica de significado ambiguo.
        elapsed = number(record.get("horas_hh")) if data.get("_rend_v4") else None
        hh = consumption = None
        productivity = advance / elapsed if elapsed and elapsed > 0 else None
    return {"edificio": str(building), "piso": data.get("piso") or record.get("piso") or "", "avance": advance,
            "horas_jornada": elapsed, "hh": hh, "consumo": consumption, "productividad": productivity, "payload": data,
            "historico": not bool(data.get("_rend_v5"))}


def rendimiento_status(advance, hh, theoretical):
    if advance <= 0:
        return "SIN AVANCE" if hh > 0 else "SIN REGISTRO"
    if hh <= 0:
        return "SIN TIEMPO"
    if theoretical <= 0:
        return "SIN META"
    return "CUMPLE META" if hh / advance <= theoretical else "BAJO RENDIMIENTO"


def row_id():
    return uuid.uuid4().hex[:12]


def nested_get(data, path, default=None):
    value = data
    try:
        for part in path:
            value = value[part]
        return value
    except (KeyError, IndexError, TypeError):
        return default


def nested_set(data, path, value):
    target = data
    for part in path[:-1]:
        target = target[part]
    target[path[-1]] = value


def rerun_selected_module():
    """Redibuja el módulo activo, sin reconstruir cabecera, login ni dashboard."""
    if getattr(st, "fragment", None):
        st.rerun(scope="fragment")
    else:
        st.rerun()


class Editor:
    def __init__(self, module, user, draft):
        self.module = module; self.user = user; self.draft = draft
        self.data = draft["data"]
        self.slot = draft["slot"]
        self.prefix = f"edit:{user['correo']}:{module}:{self.slot}:"

    def persist(self):
        body = {"usuario_email": self.user["correo"], "modulo": self.module, "slot": self.slot, "datos": self.data}
        try:
            if self.draft.get("revision") is None:
                result = database().table("alpha_drafts").insert(body).execute()
            else:
                result = database().table("alpha_drafts").update({"datos": self.data}).eq("usuario_email", self.user["correo"]).eq("modulo", self.module).eq("slot", self.slot).eq("revision", self.draft["revision"]).execute()
            if not result.data:
                raise ConflictError("El borrador cambió en otro dispositivo. No se sobrescribió. Guarde una copia del borrador o recárguelo.")
            self.draft["revision"] = result.data[0]["revision"]
            # El listado de borradores solo se actualiza en operaciones concretas.
            self.draft["saved"] = now_ec().strftime("%H:%M:%S")
            self.draft.pop("error", None)
        except Exception as exc:
            self.draft["error"] = str(exc) if isinstance(exc, ConflictError) else "No se pudo guardar el borrador en el servidor. Mantenga esta página abierta y vuelva a intentarlo."

    def changed(self, key, path, kind):
        value = st.session_state[key]
        if kind == "date_input":
            value = value.isoformat()
        elif kind == "time_input":
            value = value.strftime("%H:%M")
        nested_set(self.data, path, value)
        self.persist()

    def field(self, kind, path, label, default=None, **kwargs):
        path = tuple(path) if isinstance(path, (tuple, list)) else (path,)
        value = nested_get(self.data, path, default)
        if value is None:
            value = default
        if nested_get(self.data, path, None) is None:
            nested_set(self.data, path, copy.deepcopy(value))
        # Las rutas incluyen identificadores estables de filas, no su posición visual.
        key = self.prefix + ":".join(str(p) for p in path)
        callback = dict(key=key, on_change=self.changed, args=(key, path, kind))
        if kind in {"text_input", "text_area"}:
            return getattr(st, kind)(label, value=str(value or ""), **callback, **kwargs)
        if kind == "number_input":
            if isinstance(default, int):
                value = int(number(value, default))
            else:
                value = number(value, default or 0)
            return st.number_input(label, value=value, **callback, **kwargs)
        if kind == "date_input":
            try:
                value = dt.date.fromisoformat(str(value)[:10])
            except ValueError:
                value = now_ec().date()
            return st.date_input(label, value=value, **callback, **kwargs)
        if kind == "time_input":
            try:
                time_value = dt.time.fromisoformat(str(value or default or "07:00")[:5])
            except ValueError:
                time_value = dt.time(7, 0)
            return st.time_input(label, value=time_value, step=300, **callback, **kwargs)
        if kind == "checkbox":
            return st.checkbox(label, value=bool(value), **callback, **kwargs)
        if kind == "segmented_control":
            options = list(kwargs.pop("options"))
            selected = value if value in options else None
            return st.segmented_control(label, options, default=selected, **callback, **kwargs)
        if kind == "selectbox":
            options = list(kwargs.pop("options"))
            if value not in options and value not in (None, ""):
                options.append(value)
            return st.selectbox(label, options, index=options.index(value) if value in options else 0, **callback, **kwargs)
        if kind == "multiselect":
            options = list(kwargs.pop("options")); value = value if isinstance(value, list) else []
            options = list(dict.fromkeys(options + value))
            return st.multiselect(label, options, default=value, **callback, **kwargs)
        raise ValueError("Control no soportado")

    def photos(self, path, label="Fotografías", limit=MAX_PHOTOS):
        path = tuple(path)
        photos = nested_get(self.data, path, [])
        if not isinstance(photos, list):
            photos = []
            nested_set(self.data, path, photos)
        upload_key = self.prefix + "upload:" + ":".join(map(str, path))
        def changed_photos():
            try:
                uploads = st.session_state.get(upload_key) or []
                additions = [image_b64(file) for file in uploads]
                existing = nested_get(self.data, path, [])
                combined = list(dict.fromkeys(existing + additions))
                if len(combined) > limit:
                    raise ValueError(f"Se permiten hasta {limit} fotografías en esta sección. Retire una antes de añadir otra.")
                nested_set(self.data, path, combined)
                self.draft.pop("photo_error", None)
                self.persist()
            except Exception as exc:
                self.draft["photo_error"] = str(exc)
        st.file_uploader(label, type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True, key=upload_key, on_change=changed_photos)
        for idx, photo in enumerate(photos):
            stream = image_stream(photo)
            if stream:
                st.image(stream, width=200, caption=f"{label} {idx + 1}")
            else:
                st.warning("Hay una fotografía histórica que no se puede interpretar; se conservará hasta que decida retirarla.")
            if st.button("Retirar fotografía", key=upload_key + ":remove:" + str(idx)):
                photos.pop(idx); nested_set(self.data, path, photos); self.persist(); rerun_selected_module()

    def rows(self, field, default):
        self.data.setdefault(field, {})
        if not self.data[field]:
            self.data[field][row_id()] = copy.deepcopy(default)
        return self.data[field]

    def add(self, field, default, label="Agregar fila"):
        if st.button("➕ " + label, key=self.prefix + "add:" + field):
            self.data[field][row_id()] = copy.deepcopy(default)
            self.persist(); rerun_selected_module()

    def remove(self, field, rid):
        if st.button("🗑️ Retirar fila", key=self.prefix + "del:" + field + rid):
            self.data[field].pop(rid, None); self.persist(); rerun_selected_module()

    def status(self):
        if self.draft.get("error"):
            st.warning(self.draft["error"])
            if st.button("Reintentar autoguardado", key=self.prefix + "retry"):
                self.persist(); rerun_selected_module()
            st.download_button("Descargar respaldo del borrador", json_text(self.data), file_name=f"borrador_{self.module}_{self.slot}.json", mime="application/json", key=self.prefix + "backup")
        else:
            st.caption("Borrador guardado en el servidor: " + self.draft.get("saved", "restaurado") + ". Los campos de texto se guardan al confirmar con Enter o salir del campo; las fotos al terminar de cargarse.")
        if self.draft.get("photo_error"):
            st.error(self.draft["photo_error"])

    def close(self, delete=False):
        if delete:
            result = database().table("alpha_drafts").delete().eq("usuario_email", self.user["correo"]).eq("modulo", self.module).eq("slot", self.slot).eq("revision", self.draft.get("revision", -1)).execute()
            if not result.data and self.draft.get("revision") is not None:
                st.warning("El registro se guardó, pero el borrador cambió en otra sesión y se conservó.")
        st.session_state.get("_editors", {}).pop(self.module, None)
        cached_rows.clear()
        st.session_state.setdefault("_cleanup_prefixes", []).append(self.prefix)


def begin_editor(module, user, data, original=None, slot=None):
    data = copy.deepcopy(data)
    data["_original"] = copy.deepcopy(original)
    data.setdefault("_request_id", str(uuid.uuid4()))
    slot = slot or ("record_" + str(original["id"]) if original else "new_" + uuid.uuid4().hex)
    existing = database().table("alpha_drafts").select("*").eq("usuario_email", user["correo"]).eq("modulo", module).eq("slot", slot).execute().data
    if existing:
        draft = {"slot": slot, "data": existing[0]["datos"], "revision": existing[0]["revision"]}
    else:
        draft = {"slot": slot, "data": data, "revision": None}
    st.session_state.setdefault("_editors", {})[module] = draft
    cached_rows.clear()
    editor = Editor(module, user, draft)
    if not existing:
        editor.persist()
    return editor


def editor_picker(module, user, default_factory, allow_new=True):
    draft = st.session_state.setdefault("_editors", {}).get(module)
    if draft:
        editor = Editor(module, user, draft)
        editor.status()
        a, b = st.columns(2)
        if a.button("Cerrar y conservar borrador", key=f"keep_{module}"):
            editor.persist(); editor.close(); rerun_selected_module()
        with b.popover("Descartar borrador"):
            st.caption("Esta acción elimina únicamente este borrador, no el registro original.")
            if st.button("Confirmar descarte", key=f"discard_{module}"):
                editor.close(delete=True); rerun_selected_module()
        return editor
    if not allow_new:
        st.caption("Su cargo actual permite consultar y editar sus históricos de este libro. La creación corresponde al cargo asignado para este tipo de libro.")
    if st.button("➕ Crear nuevo registro", key=f"new_{module}", type="primary", disabled=not allow_new):
        begin_editor(module, user, default_factory()); rerun_selected_module()
    drafts = cached_rows("alpha_drafts", (("usuario_email", user["correo"]), ("modulo", module)), columns="slot,revision,updated_at", order="updated_at")
    if drafts:
        choices = {f"Borrador {i+1} · {r['updated_at'][:16]}": r for i,r in enumerate(reversed(drafts))}
        choice = st.selectbox("Continuar un borrador", list(choices), key=f"draft_choice_{module}")
        if st.button("Recuperar borrador", key=f"restore_{module}"):
            row = choices[choice]
            recovered = all_rows("alpha_drafts", filters={"usuario_email": user["correo"], "modulo": module, "slot": row["slot"]}, columns="slot,datos,revision")
            if recovered:
                st.session_state["_editors"][module] = {"slot": recovered[0]["slot"], "data": recovered[0]["datos"], "revision": recovered[0]["revision"]}
                rerun_selected_module()
            else:
                st.warning("El borrador ya no está disponible. Actualice la página.")
    return None


def save_editor(editor, repo, table, payload):
    if editor.draft.get("photo_error"):
        raise ValueError("Corrija el problema de fotografías antes de guardar.")
    original = editor.data.get("_original")
    if not original:
        payload["client_request_id"] = editor.data["_request_id"]
        # Una respuesta perdida no debe insertar otra copia al reintentar.
        prior = database().table(table).select("*").eq("client_request_id", payload["client_request_id"]).eq("usuario_email", repo.email).execute().data
        if prior:
            editor.data["_original"] = prior[0]
            original = prior[0]
    result = repo.save(table, payload, original)
    editor.data["_original"] = result
    try:
        editor.close(delete=True)
    except Exception:
        editor.close()
        st.session_state["_flash"] = "Registro guardado. No se pudo retirar su borrador; se conserva como respaldo."
    else:
        st.session_state["_flash"] = "Registro guardado y confirmado por Supabase."
    st.rerun()


def project_options(user, existing=None):
    options = EDIFICIOS_ALPHA if user.get("es_admin") else list(user.get("edificios") or [])
    if existing and existing not in options:
        options = options + [existing]
    return options


def date_valid(value):
    try:
        return dt.date.fromisoformat(str(value)[:10])
    except (ValueError, TypeError):
        return None


def checklist_data(record=None):
    record = record or {}
    payload = obj(record.get("datos"), {})
    verifications = payload if isinstance(payload, list) else payload.get("Verificaciones", [])
    data = {"fecha": str(record.get("fecha") or now_ec().date()), "edificio": record.get("edificio") or "", "hora_inicio": str(record.get("hora_inicio") or "07:00")[:5], "hora_fin": str(record.get("hora_fin") or "16:00")[:5],
            "observacion_general": record.get("observacion_general") or "", "verificaciones": {}, "supervision": {}, "_extra": payload if isinstance(payload, dict) else {}}
    remaining = [v for v in verifications if isinstance(v, dict)]
    for jornada, activities in [("Mañana", ACTIVIDADES_MANANA_CLEAN), ("Tarde", ACTIVIDADES_TARDE_CLEAN)]:
        for idx, activity in enumerate(activities, 1):
            match = next((v for v in remaining if str(v.get("Jornada", "")).casefold() == jornada.casefold() and (v.get("Actividad") == activity or str(v.get("N°")) == str(idx))), None)
            value = copy.deepcopy(match or {"Jornada": jornada, "N°": idx, "Actividad": activity, "Estado": "Sin responder", "Observaciones": [], "Responsables": []})
            if match:
                remaining.remove(match)
            value["Estado"] = value.get("Estado") or "Sin responder"
            obs = value.get("Observaciones") or []
            value["Observaciones"] = "\n".join(map(str, obs)) if isinstance(obs, list) else str(obs)
            value["Fotos"] = list_photos(value)
            value["Responsables"] = {row_id(): v for v in value.get("Responsables", []) if isinstance(v, dict)}
            data["verificaciones"][row_id()] = value
    data["_otras_verificaciones"] = remaining
    for item in (payload.get("Supervision_Trabajos", []) if isinstance(payload, dict) else []):
        if isinstance(item, dict):
            entry = copy.deepcopy(item); entry["Fotos"] = list_photos(item)
            data["supervision"][row_id()] = entry
    return data


def render_checklist(editor, repo):
    """Reproduce los bloques mañana/tarde del diseño clásico usando borradores seguros."""
    d = editor.data
    st.markdown("#### ⚙️ Configuración de la Jornada")
    col_ed, col_resp, col_fecha, col_ini, col_fin = st.columns([2.1, 2.1, 1.55, 1.25, 1.25])
    with col_ed:
        editor.field("selectbox", "edificio", "Edificio / Proyecto:", "", options=[""] + project_options(repo.user, d.get("edificio")))
    with col_resp:
        st.text_input("Responsable:", value=full_name(repo.user), disabled=True, key=editor.prefix + "responsable_visible")
    with col_fecha:
        editor.field("date_input", "fecha", "Fecha:", str(now_ec().date()), disabled=bool(d.get("_original")))
    with col_ini:
        editor.field("time_input", "hora_inicio", "Hora Inicio:", "07:00")
    with col_fin:
        editor.field("time_input", "hora_fin", "Hora Fin:", "16:00")

    for turno, activities, title in [
        ("Mañana", ACTIVIDADES_MANANA_CLEAN, "🌅 Jornada de la Mañana"),
        ("Tarde", ACTIVIDADES_TARDE_CLEAN, "🌆 Jornada de la Tarde"),
    ]:
        st.markdown("#### " + title)
        selected = [(rid, item) for rid, item in d["verificaciones"].items() if item.get("Jornada") == turno]
        for rid, item in selected:
            st.markdown('<div class="banner-item-header"><span>N° ' +
                        html_escape(str(item.get("N°", ""))) + ' — ' +
                        html_escape(str(item.get("Actividad", ""))) + '</span></div>', unsafe_allow_html=True)
            with st.container(border=True):
                a, b, c = st.columns([1.25, 2.5, 1.9])
                with a:
                    editor.field("segmented_control", ("verificaciones", rid, "Estado"), "Estado general", None,
                                 options=["✓ Cumple", "✗ No Cumple"])
                with b:
                    editor.field("text_area", ("verificaciones", rid, "Observaciones"),
                                 "Observaciones del ítem (una por línea)", "", height=80)
                with c:
                    if turno == "Mañana" and item.get("N°") == 1 and not item.get("Fotos"):
                        st.caption("📷 Evidencia fotográfica opcional")
                    with st.expander("📷 Fotografías de evidencia", expanded=False):
                        editor.photos(("verificaciones", rid, "Fotos"), "Evidencia")
                if item.get("Actividad") == "Coordinación con otras especialidades" and (item.get("Estado") == "✗ No Cumple" or item.get("Responsables")):
                    st.markdown("**Responsables y especialidades involucradas**")
                    for pid, person in list(item["Responsables"].items()):
                        p1, p2, p3 = st.columns([2.5, 2.5, .65])
                        with p1:
                            editor.field("text_input", ("verificaciones", rid, "Responsables", pid, "nombre"), "Nombre", "")
                        with p2:
                            editor.field("text_input", ("verificaciones", rid, "Responsables", pid, "area"), "Área", "")
                        with p3:
                            if st.button("🗑️", key=editor.prefix + rid + pid + "remove", help="Quitar responsable"):
                                item["Responsables"].pop(pid, None)
                                editor.persist()
                                rerun_selected_module()
                    if st.button("➕ Agregar responsable", key=editor.prefix + rid + "add_responsable"):
                        item["Responsables"][row_id()] = {"nombre": "", "area": ""}
                        editor.persist()
                        rerun_selected_module()

    st.markdown("#### 🏗️ Supervisión de la ejecución de trabajos")
    default = {"Actividad": "", "Encargados": "", "Observaciones": "", "Fotos": []}
    for index, (rid, item) in enumerate(list(editor.rows("supervision", default).items()), 1):
        with st.container(border=True):
            st.markdown(f"**Actividad supervisada N° {index}**")
            c1, c2 = st.columns(2)
            with c1:
                editor.field("text_input", ("supervision", rid, "Actividad"), "Trabajo supervisado", "")
            with c2:
                editor.field("text_input", ("supervision", rid, "Encargados"), "Personal encargado", "")
            editor.field("text_area", ("supervision", rid, "Observaciones"), "Observaciones", "", height=72)
            with st.expander("📷 Fotos de la supervisión", expanded=False):
                editor.photos(("supervision", rid, "Fotos"))
            editor.remove("supervision", rid)
    editor.add("supervision", default, "Agregar actividad de supervisión")
    editor.field("text_area", "observacion_general", "Observación general de la jornada", "")
    if st.button("💾 Guardar Checklist de Obra", type="primary", key=editor.prefix + "save"):
        try:
            if not d.get("edificio"):
                raise ValueError("Seleccione el edificio de la jornada.")
            time_calculation(d["hora_inicio"], d["hora_fin"])
            verifications = []
            for item in d["verificaciones"].values():
                value = copy.deepcopy(item)
                value["Estado"] = value.get("Estado") or "Sin responder"
                value["Observaciones"] = [x.strip() for x in str(value.get("Observaciones") or "").splitlines() if x.strip()]
                value["Responsables"] = [v for v in value.get("Responsables", {}).values() if str(v.get("nombre", "")).strip()]
                verifications.append(value)
            supervision = [copy.deepcopy(v) for v in d["supervision"].values() if v.get("Actividad") or v.get("Fotos") or v.get("Observaciones")]
            if not any(v.get("Estado") != "Sin responder" or v.get("Observaciones") or v.get("Fotos") for v in verifications) and not supervision and not d.get("observacion_general"):
                raise ValueError("Registre al menos un control, actividad, observación o fotografía.")
            payload = copy.deepcopy(d.get("_extra", {}))
            payload.update({"Verificaciones": verifications + d.get("_otras_verificaciones", []), "Supervision_Trabajos": supervision})
            save_editor(editor, repo, "checklists", {
                "fecha": d["fecha"], "edificio": d["edificio"],
                "hora_inicio": d["hora_inicio"], "hora_fin": d["hora_fin"],
                "responsable": full_name(repo.user), "cargo": repo.user["cargo"],
                "observacion_general": d.get("observacion_general", ""), "datos": payload,
            })
        except Exception as exc:
            notice_error(exc)




def full_name(user):
    return f"{user.get('nombres') or ''} {user.get('apellidos') or ''}".strip()


def official_data(record=None):
    record=record or {}; payload=obj(record.get("datos"),{})
    if not isinstance(payload,dict): payload={"_legacy_items":payload}
    return {"fecha":str(record.get("fecha") or now_ec().date()),"proyecto":record.get("proyecto") or "", "residente":record.get("residente") or "", "hora_inicio":str(record.get("hora_inicio") or "07:00")[:5],"hora_fin":str(record.get("hora_fin") or "16:00")[:5],
            "datos":{**{"Superintendente":"","Fiscalizador":"","Ubicacion":record.get("frente") or "","Barrio":"","Hoja":"","Nomina_Conteo":{},"Rotativo_Conteo":{},"Clima_Condicion":record.get("clima") or "SOL","Clima_Obs":"","Maquinaria_Conteo":{},"Seguridad_Check":{},"Novedades":""},**payload},
            "actividades":{row_id():copy.deepcopy(v) for v in payload.get("Actividades_Ejecutadas",[]) if isinstance(v,dict)}}


OFICIO_ALIASES={"ALBAÑILES":["ALBAÑIL","ALBANIL"],"AYUDANTE":["AYUDANTE","PEON","PEÓN"],"SOLDADOR":["SOLDADOR"],"OPERADOR":["OPERADOR","MAQUINISTA"],"FIERRERO":["FIERRERO","HIERRERO"],"PINTORES":["PINTOR","PINTORES"],"MAESTRO SUPERVISOR":["MAESTRO","SUPERVISOR","CAPATAZ"],"CARPINTERO":["CARPINTERO"],"GUACHIMAN":["GUACHIMAN","GUARDIA"],"GYPSEROS":["GYPS","TABLAYESO"],"ELECTRICOS":["ELECTRIC","ELÉCTRIC"],"PLOMEROS":["PLOMER","HIDROSANITARI","FONTANER"],"ALUMINIO Y VIDRIO":["ALUMIN","VIDRI"]}


def workforce_counts(workers):
    counts={k:0 for k in OFICIOS_NOMINA_FORMATO}; unclassified=[]
    for worker in workers:
        role=str(worker.get("cargo") or "").upper()
        matches=[k for k,aliases in OFICIO_ALIASES.items() if any(v in role for v in aliases)]
        if len(matches)==1: counts[matches[0]]+=1
        else: unclassified.append(worker.get("nombre") or "Sin nombre")
    return counts,unclassified


def render_official(editor,repo):
    d=editor.data; original=d.get("_original")
    editor.field("date_input","fecha","Fecha",str(now_ec().date()),disabled=bool(original))
    editor.field("selectbox","proyecto","Proyecto","",options=[""]+project_options(repo.user,d.get("proyecto")))
    matching=[r for r in repo.own("checklists") if str(r.get("fecha"))[:10]==str(d["fecha"])[:10] and r.get("edificio")==d["proyecto"]]
    if matching:
        chosen=st.selectbox("Checklist del mismo edificio y fecha",matching,format_func=lambda r:f"{display_date(r['fecha'])} · {r.get('hora_inicio') or 'Sin hora'}",key=editor.prefix+"checklist")
        if st.button("Importar actividades y horario del checklist",key=editor.prefix+"importchk"):
            raw=obj(chosen.get("datos"),{})
            items=raw.get("Supervision_Trabajos",[]) if isinstance(raw,dict) else []
            for item in items:
                if item.get("Actividad"):
                    d["actividades"][row_id()]={"Descripcion":item["Actividad"],"Area":"","Unidad":"","Cantidad":0.0}
            d["hora_inicio"]=str(chosen.get("hora_inicio") or "07:00")[:5];d["hora_fin"]=str(chosen.get("hora_fin") or "16:00")[:5]
            d["datos"]["Checklist_Origen_ID"]=str(chosen["id"])
            for field in ("hora_inicio","hora_fin"):
                st.session_state.pop(editor.prefix+field,None)
            editor.persist();rerun_selected_module()
    else: st.caption("No hay checklist del mismo edificio y fecha para importar.")
    st.subheader("Datos generales")
    for key,label in [("Ubicacion","Ubicación"),("Barrio","Barrio"),("Superintendente","Superintendente"),("Fiscalizador","Fiscalizador"),("Hoja","Hoja N°")]:
        editor.field("text_input",("datos",key),label,"")
    if not d.get("residente"):d["residente"]=full_name(repo.user)
    editor.field("text_input","residente","Residente de obra",full_name(repo.user))
    a,b=st.columns(2)
    with a:editor.field("text_input","hora_inicio","Entrada","07:00")
    with b:editor.field("text_input","hora_fin","Salida","16:00")
    st.subheader("Personal nómina y subcontratos")
    workers=combined_workers(repo,d.get("proyecto"))
    if st.button("Actualizar conteo desde mi nómina del edificio",key=editor.prefix+"nomina"):
        counts,unclassified=workforce_counts(workers);d["datos"]["Nomina_Conteo"]=counts
        for key in list(st.session_state):
            if key.startswith(editor.prefix+"datos:Nomina_Conteo:"):del st.session_state[key]
        d["_unclassified"]=unclassified;editor.persist();rerun_selected_module()
    if d.get("_unclassified"):st.warning("Revise manualmente los oficios de: "+", ".join(d["_unclassified"]))
    a,b=st.columns(2)
    for col,field,labels in [(a,"Nomina_Conteo",OFICIOS_NOMINA_FORMATO),(b,"Rotativo_Conteo",RUROS_ROTATIVOS_FORMATO)]:
        with col:
            for label in labels:editor.field("number_input",("datos",field,label),label,0,min_value=0,step=1)
    st.subheader("Clima y maquinaria")
    editor.field("selectbox",("datos","Clima_Condicion"),"Clima observado","SOL",options=["SOL","NUBLADO","LLUVIA TENUE","LLUVIA INTENSA","GRANIZO"])
    editor.field("text_input",("datos","Clima_Obs"),"Observaciones del clima","")
    for label in MAQUINARIAS_FORMATO:editor.field("number_input",("datos","Maquinaria_Conteo",label),label,0,min_value=0,step=1)
    st.subheader("Seguridad, señalización y mitigación")
    labels=["Casco","Chalecos","Guantes","Gafas","Mascarilla","Auditivo","Conos","Cintas","Rótulos","Vallas","Extintor","Botiquin","Polvo","Ruido","Liquidos","Cerramiento","Limpieza"]
    cols=st.columns(3)
    for i,label in enumerate(labels):
        with cols[i%3]:editor.field("checkbox",("datos","Seguridad_Check",label),label,False)
    st.subheader("Actividades realizadas")
    default={"Descripcion":"","Area":"","Unidad":"","Cantidad":0.0}
    for rid,item in list(editor.rows("actividades",default).items()):
        with st.container(border=True):
            for key,label in [("Descripcion","Descripción"),("Area","Área de trabajo"),("Unidad","Unidad")]:editor.field("text_input",("actividades",rid,key),label,"")
            editor.field("number_input",("actividades",rid,"Cantidad"),"Cantidad",0.0,min_value=0.0,step=0.01)
            editor.remove("actividades",rid)
    editor.add("actividades",default,"Agregar actividad")
    editor.field("text_area",("datos","Novedades"),"Novedades y recomendaciones","")
    if st.button("Guardar libro oficial",type="primary",key=editor.prefix+"save"):
        try:
            time_calculation(d["hora_inicio"],d["hora_fin"])
            activities=[v for v in d["actividades"].values() if str(v.get("Descripcion","")).strip()]
            if not activities:raise ValueError("Ingrese al menos una actividad.")
            payload=copy.deepcopy(d["datos"]);payload["Actividades_Ejecutadas"]=activities;payload["Tipo_Registro"]="Libro_Obra_Oficial"
            day=dt.date.fromisoformat(d["fecha"])
            save_editor(editor,repo,"inspecciones",{"fecha":d["fecha"],"dia":["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"][day.weekday()],"proyecto":d["proyecto"],"residente":d["residente"],"frente":payload["Ubicacion"],"clima":payload["Clima_Condicion"],"hora_inicio":d["hora_inicio"],"hora_fin":d["hora_fin"],"datos":payload})
        except Exception as exc:notice_error(exc)


def maestro_data(record=None):
    record=record or {};payload=obj(record.get("datos"),{})
    acts=payload if isinstance(payload,list) else payload.get("Actividades_Maestro",[])
    normalized={}
    for item in acts:
        if not isinstance(item,dict):continue
        value=copy.deepcopy(item)
        for upper,lower in [("Actividad","actividad"),("Cantidad","cantidad"),("Personal_A_Cargo","personal_a_cargo"),("Piso","piso"),("Departamento","departamento")]:
            value[upper]=item.get(upper,item.get(lower,[] if upper=="Personal_A_Cargo" else ""))
        if not isinstance(value["Personal_A_Cargo"],list):value["Personal_A_Cargo"]=[str(value["Personal_A_Cargo"])]
        normalized[row_id()]=value
    return {"fecha":str(record.get("fecha") or now_ec().date()),"proyecto":record.get("proyecto") or "","actividades":normalized,"_extra":payload if isinstance(payload,dict) else {}}


def render_maestro(editor, repo):
    d = editor.data
    a, b, c = st.columns([1.4, 2.1, 1.6])
    with a:
        editor.field("date_input", "fecha", "Fecha de trabajo:*", str(now_ec().date()), disabled=bool(d.get("_original")))
    with b:
        editor.field("selectbox", "proyecto", "Proyecto / Edificio asignado:*", "", options=[""] + project_options(repo.user, d.get("proyecto")))
    with c:
        st.text_input("Maestro responsable:", value=full_name(repo.user), disabled=True, key=editor.prefix + "maestro_visible")
    workers = [f"{r['nombre']} ({r['cargo']})" for r in combined_workers(repo,d.get("proyecto"))]
    st.markdown(f"#### 🔨 Actividades realizadas, personal y metrajes ({len(d['actividades']) or 1} registro/s)")
    default = {"Actividad": "", "Cantidad": "", "Personal_A_Cargo": [], "Piso": "", "Departamento": "", "Observaciones": "", "Personal_Manual": ""}
    for index, (rid, item) in enumerate(list(editor.rows("actividades", default).items()), 1):
        st.markdown(f'<div class="banner-item-header"><span>Actividad N° {index}</span></div>', unsafe_allow_html=True)
        with st.container(border=True):
            c1, c2 = st.columns([2.6, 1.4])
            with c1:
                editor.field("text_input", ("actividades", rid, "Actividad"), "Descripción de la actividad", "")
            with c2:
                editor.field("text_input", ("actividades", rid, "Cantidad"), "Cantidad / avance y unidad", "")
            c3, c4, c5 = st.columns([2.0, 1.0, 1.0])
            with c3:
                editor.field("multiselect", ("actividades", rid, "Personal_A_Cargo"), "Personal a cargo", [], options=workers)
            with c4:
                editor.field("text_input", ("actividades", rid, "Piso"), "Piso", "")
            with c5:
                editor.field("text_input", ("actividades", rid, "Departamento"), "Departamento", "")
            d1, d2 = st.columns([3, 1])
            with d1:
                editor.field("text_input", ("actividades", rid, "Personal_Manual"), "Personal adicional (separado por comas)", "")
            with d2:
                editor.remove("actividades", rid)
            with st.expander("Observaciones de la actividad", expanded=False):
                editor.field("text_area", ("actividades", rid, "Observaciones"), "Observaciones", "")
    editor.add("actividades", default, "Agregar otra actividad")
    if st.button("💾 Guardar Reporte de Maestro Mayor", type="primary", key=editor.prefix + "save"):
        try:
            if not d.get("proyecto"):
                raise ValueError("Seleccione el edificio del reporte.")
            acts = []
            for item in d["actividades"].values():
                if not str(item.get("Actividad", "")).strip():
                    continue
                value = copy.deepcopy(item)
                value["Personal_A_Cargo"] = list(dict.fromkeys(value.get("Personal_A_Cargo", []) + [x.strip() for x in value.pop("Personal_Manual", "").split(",") if x.strip()]))
                acts.append(value)
            if not acts:
                raise ValueError("Ingrese al menos una actividad.")
            payload = copy.deepcopy(d.get("_extra", {}))
            payload.update({"Tipo_Registro": "Libro_Obra_Maestro", "Actividades_Maestro": acts, "Edificio": d["proyecto"], "Maestro": full_name(repo.user), "Fecha": d["fecha"]})
            day = dt.date.fromisoformat(d["fecha"])
            save_editor(editor, repo, "inspecciones", {
                "fecha": d["fecha"], "dia": ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"][day.weekday()],
                "proyecto": d["proyecto"], "residente": full_name(repo.user), "frente": "Reporte de Maestro Mayor",
                "clima": "N/A", "hora_inicio": "07:00", "hora_fin": "16:00", "datos": payload,
            })
        except Exception as exc:
            notice_error(exc)




def rend_data(record=None):
    record=record or {};payload=rendimiento_payload(record) if record else {}
    workers=payload.get("trabajadores") or ([{"nombre":record.get("trabajador", ""),"cargo":record.get("cargo_obrero", "")} ] if record else [])
    data={"fecha":str(record.get("fecha") or now_ec().date()),"edificio":payload.get("edificio") or record.get("edificio") or "","piso":payload.get("piso") or record.get("piso") or "","rubro":record.get("rubro") or "","unidad":record.get("unidad") or "",
          "hora_inicio":payload.get("hora_inicio", ""),"hora_fin":payload.get("hora_fin", ""),"hora_muerta":payload.get("hora_muerta", ""),"lunch":bool(payload.get("lunch")),"almuerzo":bool(payload.get("almuerzo")),"overnight":bool(payload.get("overnight")),
          "comentarios":payload.get("comentarios", ""),"teorico":number(record.get("rend_teorico")),"individuales":bool(payload.get("individuales")),"trabajadores":{row_id():{**v,"horas":number(v.get("horas"))} for v in workers if isinstance(v,dict)},"_extra":payload,
          "recalcular":not record or bool(payload.get("_rend_v5"))}
    for period in ("manana","mediodia","tarde"):
        data["avance_"+period]=number(payload.get("avance_"+period))
        data["fotos_"+period]=payload.get("fotos_"+period) or ([payload["foto_"+period]] if payload.get("foto_"+period) else [])
    return data


def render_rend(editor,repo):
    d=editor.data;original=d.get("_original");old=bool(original and not rendimiento_payload(original).get("_rend_v5"))
    if old:
        st.warning("Registro histórico: el significado de sus horas y rendimiento no está confirmado. Puede conservar los cálculos originales o recalcularlos después de revisar horarios y trabajadores.")
        editor.field("checkbox","recalcular","Recalcular con horas-hombre verificadas",False)
    st.subheader("Trabajadores asignados")
    editor.field("checkbox","individuales","Cada trabajador tiene una cantidad distinta de horas efectivas",False)
    default={"nombre":"","cargo":"","horas":0.0}
    for rid,item in list(editor.rows("trabajadores",default).items()):
        a,b,c=st.columns([2,2,1])
        with a:editor.field("text_input",("trabajadores",rid,"nombre"),"Nombre","")
        with b:editor.field("text_input",("trabajadores",rid,"cargo"),"Cargo","")
        with c:
            if d["individuales"]:editor.field("number_input",("trabajadores",rid,"horas"),"Horas efectivas",0.0,min_value=0.0,max_value=24.0,step=0.25)
        editor.remove("trabajadores",rid)
    editor.add("trabajadores",default,"Agregar trabajador")
    st.subheader("Ubicación y rubro")
    a,b,c=st.columns(3)
    with a:editor.field("selectbox","edificio","Edificio","",options=[""]+project_options(repo.user,d.get("edificio")))
    with b:editor.field("text_input","piso","Piso / nivel","")
    with c:editor.field("date_input","fecha","Fecha del trabajo",str(now_ec().date()))
    editor.field("text_input","rubro","Rubro","")
    editor.field("text_input","unidad","Unidad","")
    editor.field("number_input","teorico","Meta de consumo de mano de obra (HH/unidad; 0 = sin meta)",0.0,min_value=0.0,step=0.01)
    known=next((v for k,v in RENDIMIENTOS_TEORICOS.items() if k.casefold()==d.get("rubro","").strip().casefold()),None)
    if known:st.caption(f"Referencia del catálogo original: {known} HH/unidad. Verifique que la unidad y condiciones correspondan antes de usarla.")
    st.subheader("Jornada y descuentos")
    a,b,c=st.columns(3)
    with a:editor.field("text_input","hora_inicio","Inicio (HH:MM)","")
    with b:editor.field("text_input","hora_fin","Fin (HH:MM)","")
    with c:editor.field("text_input","hora_muerta","Tiempo muerto (vacío = 0)","")
    editor.field("checkbox","overnight","La jornada termina al día siguiente",False)
    a,b=st.columns(2)
    with a:editor.field("checkbox","lunch","Lunch: descontar 15 minutos",False)
    with b:editor.field("checkbox","almuerzo","Almuerzo: descontar 1 hora",False)
    st.caption("La duración calculada es el tiempo efectivo de la cuadrilla. Las HH suman las horas de todos los trabajadores. Los avances de cada periodo son cantidades adicionales, no acumulados.")
    st.subheader("Avances y fotografías")
    for period,label in [("manana","Mañana"),("mediodia","Mediodía"),("tarde","Tarde")]:
        with st.expander(label,expanded=True):
            editor.field("number_input","avance_"+period,"Avance de "+label.lower(),0.0,min_value=0.0,step=0.01,format="%.2f")
            editor.photos(("fotos_"+period,),"Fotografías de "+label.lower())
    editor.field("text_area","comentarios","Comentarios","")
    error=None;hours=hh=advance=0.0
    workers=[copy.deepcopy(v) for v in d["trabajadores"].values() if v.get("nombre", "").strip()]
    try:
        hours=time_calculation(d["hora_inicio"],d["hora_fin"],d["hora_muerta"],d["lunch"],d["almuerzo"],d["overnight"])
        workers=[copy.deepcopy(v) for v in d["trabajadores"].values() if v.get("nombre","").strip()]
        hh=sum(number(v.get("horas")) for v in workers) if d["individuales"] else len(workers)*hours
        advance=sum(number(d.get("avance_"+p)) for p in ("manana","mediodia","tarde"))
        a,b,c=st.columns(3);a.metric("Duración efectiva",f"{hours:.2f} h");b.metric("Horas-hombre",f"{hh:.2f} HH");c.metric("Avance",f"{advance:.2f} {d.get('unidad','')}")
        st.write(f"Productividad de cuadrilla: {advance/hours:.3f} {d.get('unidad','')}/h" if hours else "Productividad: sin tiempo de jornada")
        st.write(f"Consumo de mano de obra: {hh/advance:.3f} HH/{d.get('unidad','')}" if advance>0 and hh>0 else "Consumo de mano de obra: no calculable")
    except ValueError as exc:
        error=exc;st.warning(str(exc))
    if old and not d["recalcular"]:st.info("Al guardar se conservarán las cifras originales; la vista previa no sustituirá los cálculos históricos.")
    if st.button("Guardar rendimiento",type="primary",key=editor.prefix+"save"):
        try:
            preserve=old and not d["recalcular"]
            if not preserve and (not d.get("rubro","").strip() or not d.get("unidad","").strip() or not workers):raise ValueError("Complete rubro, unidad y al menos un trabajador.")
            if error and not preserve:raise error
            payload=copy.deepcopy(d.get("_extra",{}))
            payload.update({k:d[k] for k in ["edificio","piso","comentarios"]})
            for period in ("manana","mediodia","tarde"):
                payload["fotos_"+period]=d["fotos_"+period]
                payload["foto_"+period]=d["fotos_"+period][0] if d["fotos_"+period] else ""
            if preserve:
                baseline=rend_data(original)
                for key in ("hora_inicio","hora_fin","hora_muerta","lunch","almuerzo","overnight","individuales","teorico"):
                    if d.get(key)!=baseline.get(key):raise ValueError("Para cambiar horarios, descuentos o meta de un histórico, confirme el recálculo.")
                baseline_workers=[v for v in baseline["trabajadores"].values() if v.get("nombre", "").strip()]
                if workers!=baseline_workers:raise ValueError("Para cambiar trabajadores u horas de un histórico, confirme el recálculo.")
                for field in ("rubro","unidad"):
                    if d[field] != (original.get(field) or ""):raise ValueError("Para cambiar rubro o unidad de un histórico, confirme el recálculo.")
                for period in ("manana","mediodia","tarde"):
                    if number(d.get("avance_"+period))!=number(payload.get("avance_"+period)):raise ValueError("Para cambiar avances de un histórico, confirme el recálculo.")
                payload["trabajadores"]=workers
                values={k:original.get(k) for k in ["horas_hh","avance","rend_real","rend_teorico","esperado","estado","intervalo"]}
            else:
                payload.update({k:d[k] for k in ["hora_inicio","hora_fin","hora_muerta","lunch","almuerzo","overnight","individuales","avance_manana","avance_mediodia","avance_tarde"]})
                payload.update({"_rend_v5":True,"_rend_v4":False,"horas_jornada":hours,"trabajadores":[{**v,"horas":number(v.get("horas")) if d["individuales"] else hours} for v in workers],"unidad_consumo":"HH/unidad","unidad_productividad":"unidad/h"})
                if old:payload["_historico_original"]={k:original.get(k) for k in ["horas_hh","avance","rend_real","rend_teorico","estado"]}
                values={"horas_hh":hh,"avance":advance,"rend_real":hh/advance if advance>0 and hh>0 else None,"rend_teorico":d["teorico"],"esperado":hh/d["teorico"] if d["teorico"]>0 else None,"estado":rendimiento_status(advance,hh,d["teorico"]),"intervalo":"Jornada"}
            values.update({"fecha":d["fecha"],"edificio":d["edificio"],"piso":d["piso"],"rubro":d["rubro"].strip(),"unidad":d["unidad"].strip(),"trabajador":", ".join(v["nombre"] for v in workers),"cargo_obrero":", ".join(v.get("cargo","") for v in workers),"datos":payload})
            save_editor(editor,repo,"rendimientos",values)
        except Exception as exc:notice_error(exc)


def incident_data(record=None):
    record=record or {}
    return {"proyecto":record.get("proyecto") or "","area":record.get("area") or "","descripcion":record.get("descripcion") or "","responsable":record.get("responsable") or "","prioridad":record.get("prioridad") or "Media","fecha_compromiso":str(record.get("fecha_compromiso") or (now_ec().date()+dt.timedelta(days=3))),"estado":record.get("estado") or "Abierta"}


def render_incident(editor,repo):
    d=editor.data
    editor.field("selectbox","proyecto","Edificio","",options=[""]+project_options(repo.user,d.get("proyecto")))
    for key,label in [("area","Área / ubicación"),("responsable","Responsable")]:editor.field("text_input",key,label,"")
    editor.field("text_area","descripcion","Descripción de la incidencia","")
    editor.field("selectbox","prioridad","Prioridad","Media",options=["Alta","Media","Baja"])
    editor.field("date_input","fecha_compromiso","Fecha compromiso",str(now_ec().date()))
    editor.field("selectbox","estado","Estado","Abierta",options=["Abierta","Cerrada"])
    if st.button("Guardar incidencia",key=editor.prefix+"save",type="primary"):
        try:
            if any(not str(d.get(k,"")).strip() for k in ("area","descripcion","responsable")):raise ValueError("Complete área, descripción y responsable.")
            save_editor(editor,repo,"incidencias",{k:d[k] for k in ("proyecto","area","descripcion","responsable","prioridad","fecha_compromiso","estado")})
        except Exception as exc:notice_error(exc)


def record_building(table,record):
    if table=="rendimientos":return rendimiento_values(record)["edificio"]
    return str(record.get("proyecto") or record.get("edificio") or "Sin Edificio Asignado")


def text_value(value):
    if value is None:return "No determinado"
    if isinstance(value,bool):return "Sí" if value else "No"
    if isinstance(value,(list,dict)):return json_text(value)
    if isinstance(value,float):return f"{value:.4f}".rstrip("0").rstrip(".")
    return str(value)


def sections_for(table,record):
    """Modelo único compartido por PDF, Excel y vista de detalle."""
    data=obj(record.get("datos"),{})
    meta=[("Edificio",record_building(table,record)),("Fecha",record.get("fecha",record.get("fecha_compromiso",""))),("Registrado por",record.get("usuario_email", ""))]
    sections=[];photos=[]
    if table=="checklists":
        verifs=data if isinstance(data,list) else data.get("Verificaciones",[])
        supervision=data.get("Supervision_Trabajos",[]) if isinstance(data,dict) else []
        meta += [("Responsable",record.get("responsable", "")),("Cargo",record.get("cargo", "")),("Horario",f"{record.get('hora_inicio','')} – {record.get('hora_fin','')}"),("Observación general",record.get("observacion_general", ""))]
        rows=[]
        for item in verifs:
            if not isinstance(item,dict):continue
            observations=item.get("Observaciones",[])
            observations="\n".join(map(str,observations)) if isinstance(observations,list) else str(observations or "")
            people="; ".join(f"{r.get('nombre','')} ({r.get('area','')})" for r in item.get("Responsables",[]) if isinstance(r,dict))
            rows.append([item.get("Jornada",""),item.get("Actividad",""),item.get("Estado") or "Sin responder",observations,people])
            photos.extend([(f"{item.get('Jornada','')} · {item.get('Actividad','')} · Foto {i+1}",p) for i,p in enumerate(list_photos(item))])
        sections.append(("Verificaciones",["Jornada","Actividad","Estado","Observaciones","Responsables"],rows))
        rows=[]
        for item in supervision:
            if not isinstance(item,dict):continue
            rows.append([item.get("Actividad",""),item.get("Encargados",""),item.get("Observaciones","")])
            photos.extend([(f"Supervisión · {item.get('Actividad','')} · Foto {i+1}",p) for i,p in enumerate(list_photos(item))])
        sections.append(("Supervisión de trabajos",["Actividad","Encargados","Observaciones"],rows))
    elif table=="inspecciones":
        is_master=isinstance(data,list) or data.get("Tipo_Registro")=="Libro_Obra_Maestro" or "Actividades_Maestro" in data
        meta += [("Responsable",record.get("residente", "")),("Día",record.get("dia", "")),("Horario",f"{record.get('hora_inicio','')} – {record.get('hora_fin','')}")]
        if is_master:
            acts=data if isinstance(data,list) else data.get("Actividades_Maestro",[])
            rows=[]
            for item in acts:
                if not isinstance(item,dict):continue
                people=item.get("Personal_A_Cargo",item.get("personal_a_cargo",[]));people=", ".join(map(str,people)) if isinstance(people,list) else people
                rows.append([item.get("Actividad",item.get("actividad","")),people,item.get("Cantidad",item.get("cantidad","")),item.get("Piso",item.get("piso","")),item.get("Departamento",item.get("departamento","")),item.get("Observaciones",item.get("observaciones",""))])
            sections.append(("Actividades del maestro mayor",["Actividad","Personal","Cantidad","Piso","Departamento","Observaciones"],rows))
        else:
            meta += [("Ciudad",data.get("Ciudad","Quito")),("Hoja",data.get("Hoja","")),("Ubicación",data.get("Ubicacion",record.get("frente",""))),("Barrio",data.get("Barrio","")),("Superintendente",data.get("Superintendente","")),("Fiscalizador",data.get("Fiscalizador","")),("Clima",data.get("Clima_Condicion",record.get("clima",""))),("Observaciones del clima",data.get("Clima_Obs","")),("Novedades y recomendaciones",data.get("Novedades",""))]
            schedule=f"{record.get('hora_inicio','')} – {record.get('hora_fin','')}"
            for field,title in [("Nomina_Conteo","Nómina de personal"),("Rotativo_Conteo","Personal rotativo / subcontratos"),("Maquinaria_Conteo","Maquinaria y herramientas")]:
                mapping=data.get(field) or {}
                rows=[[k,v,schedule if field!="Maquinaria_Conteo" and number(v)>0 else ""] for k,v in mapping.items()] if isinstance(mapping,dict) else []
                sections.append((title,["Concepto","Cantidad","Horario"],rows))
            checks=data.get("Seguridad_Check") or {}
            sections.append(("Seguridad, señalización y mitigación",["Control","Registrado"],[[k,"Sí" if v else "No"] for k,v in checks.items()] if isinstance(checks,dict) else []))
            acts=data.get("Actividades_Ejecutadas") or []
            sections.append(("Actividades realizadas",["Descripción","Área","Unidad","Cantidad"],[[a.get("Descripcion",a.get("actividad","")),a.get("Area",""),a.get("Unidad",""),a.get("Cantidad",0)] for a in acts if isinstance(a,dict)]))
    elif table=="rendimientos":
        values=rendimiento_values(record);pay=values["payload"]
        meta += [("Piso",values["piso"]),("Rubro",record.get("rubro","")),("Unidad",record.get("unidad","")),("Avance total",values["avance"]),("Duración de jornada (h)",values["horas_jornada"]),("Horas-hombre (HH)",values["hh"]),("Productividad de cuadrilla (unidad/h)",values["productividad"]),("Consumo de mano de obra (HH/unidad)",values["consumo"]),("Meta (HH/unidad)",record.get("rend_teorico")),("Estado",record.get("estado","")),("Comentarios",pay.get("comentarios",""))]
        if values["historico"]:
            meta += [("Aviso histórico","HH y unidad del rendimiento original no verificadas; no se incluyen en totales de HH."),("Horas guardadas originalmente",record.get("horas_hh")),("Rendimiento guardado originalmente",record.get("rend_real"))]
        workers=pay.get("trabajadores") or [{"nombre":record.get("trabajador",""),"cargo":record.get("cargo_obrero","")}]
        sections.append(("Trabajadores",["Nombre","Cargo","Horas efectivas individuales"],[[w.get("nombre",""),w.get("cargo",""),w.get("horas")] for w in workers if isinstance(w,dict)]))
        sections.append(("Jornada y descuentos",["Inicio","Fin","Tiempo muerto","Lunch","Almuerzo"],[[pay.get("hora_inicio",""),pay.get("hora_fin",""),pay.get("hora_muerta", ""),"15 min" if pay.get("lunch") else "0 min","60 min" if pay.get("almuerzo") else "0 min"]]))
        sections.append(("Avances adicionales por periodo",["Mañana","Mediodía","Tarde"],[[pay.get("avance_manana",0),pay.get("avance_mediodia",0),pay.get("avance_tarde",0)]]))
        for period,label in [("manana","Mañana"),("mediodia","Mediodía"),("tarde","Tarde")]:
            values_p=pay.get("fotos_"+period) or ([pay["foto_"+period]] if pay.get("foto_"+period) else [])
            photos.extend([(f"{label} · Foto {i+1}",p) for i,p in enumerate(values_p)])
    elif table=="incidencias":
        meta += [("Área",record.get("area","")),("Descripción",record.get("descripcion","")),("Responsable",record.get("responsable","")),("Prioridad",record.get("prioridad","")),("Fecha compromiso",record.get("fecha_compromiso","")),("Estado",record.get("estado",""))]
    elif table=="trabajadores":
        meta += [("Nombre",record.get("nombre","")),("Cargo",record.get("cargo",""))]
    return meta,sections,photos


def excel_cell(cell,value):
    if isinstance(value,(dict,list)):value=json_text(value)
    if isinstance(value,str):
        value=re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]","",value)
        cell.value=value
        cell.data_type="s"
    else:cell.value=value
    cell.alignment=Alignment(vertical="top",wrap_text=True)


def csv_export(table,records):
    rows=[]
    for rec in records:
        meta,sections,_=sections_for(table,rec)
        row=dict(meta)
        for title,headers,values in sections:
            row[title]=json_text([dict(zip(headers,v)) for v in values])
        rows.append(row)
    df=pd.DataFrame(rows)
    for column in df.columns:
        df[column]=df[column].map(lambda v:"'"+v if isinstance(v,str) and v.lstrip().startswith(("=","+","-","@")) else v)
    return df.to_csv(index=False,sep=";").encode("utf-8-sig")


def summary_rend(records):
    groups=defaultdict(lambda:{"count":0,"avance":0.0,"hh":0.0,"legacy":0})
    for record in records:
        value=rendimiento_values(record);key=(value["edificio"],record.get("rubro", ""),record.get("unidad", ""))
        entry=groups[key];entry["count"]+=1;entry["avance"]+=value["avance"]
        if value["hh"] is None:entry["legacy"]+=1
        else:entry["hh"]+=value["hh"]
    return [[*key,g["count"],g["avance"],g["hh"],g["legacy"]] for key,g in sorted(groups.items())]


def export_excel(table,records):
    wb=openpyxl.Workbook();wb.remove(wb.active)
    dark="17365D";border=Border(left=Side(style="thin",color="CBD5E1"),right=Side(style="thin",color="CBD5E1"),top=Side(style="thin",color="CBD5E1"),bottom=Side(style="thin",color="CBD5E1"))
    def header(ws,row,values):
        for col,value in enumerate(values,1):
            cell=ws.cell(row,col);excel_cell(cell,value);cell.fill=PatternFill("solid",fgColor=dark);cell.font=Font(bold=True,color="FFFFFF");cell.border=border
        ws.row_dimensions[row].height=32
    if table=="rendimientos":
        ws=wb.create_sheet("Resumen")
        header(ws,1,["Edificio","Rubro","Unidad","Registros","Avance","HH verificadas","Registros históricos"])
        for row,values in enumerate(summary_rend(records),2):
            for col,value in enumerate(values,1):excel_cell(ws.cell(row,col),value)
        for col in "ABCDEFG":ws.column_dimensions[col].width=23
        ws.freeze_panes="A2"
    groups=defaultdict(list)
    for record in records:groups[record_building(table,record)].append(record)
    for building,items in sorted(groups.items()):
        name=re.sub(r"[\\/*?:\[\]]","_",building)[:31] or "Edificio";base=name;n=2
        while name.casefold() in {v.casefold() for v in wb.sheetnames}:
            suffix=f"_{n}";name=base[:31-len(suffix)]+suffix;n+=1
        ws=wb.create_sheet(name)
        for col in range(1,8):ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width=24
        header(ws,1,["ALPHA BUILDERS",table.upper(),building]);row=3
        for record in sorted(items,key=lambda r:str(r.get("fecha",r.get("fecha_compromiso",""))),reverse=True):
            meta,sections,photos=sections_for(table,record)
            header(ws,row,["Registro",display_date(record_date(table, record))]);row+=1
            for label,value in meta:
                excel_cell(ws.cell(row,1),label);ws.cell(row,1).font=Font(bold=True)
                ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=7);excel_cell(ws.cell(row,2),value)
                ws.row_dimensions[row].height=max(20,min(250,15*(1+len(text_value(value))//100+text_value(value).count("\n"))));row+=1
            for title,headers,values in sections:
                row+=1;header(ws,row,[title]);row+=1;header(ws,row,headers);row+=1
                for values_row in values:
                    longest=1
                    for col,value in enumerate(values_row,1):
                        cell=ws.cell(row,col);excel_cell(cell,value);cell.border=border
                        longest=max(longest,len(text_value(value))//25+text_value(value).count("\n")+1)
                    ws.row_dimensions[row].height=min(400,max(22,14*longest));row+=1
            for label,photo in photos:
                excel_cell(ws.cell(row,1),label);row+=1;stream=image_stream(photo)
                if stream:
                    image=XLImage(stream);ratio=min(420/image.width,220/image.height);image.width*=ratio;image.height*=ratio;ws.add_image(image,f"B{row}");ws.row_dimensions[row].height=175
                else:excel_cell(ws.cell(row,2),"Fotografía no interpretable; el original permanece almacenado.")
                row+=2
            row+=3
        ws.freeze_panes="A3";ws.print_title_rows="1:2";ws.page_setup.orientation="landscape";ws.page_setup.paperSize=ws.PAPERSIZE_A4;ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0;ws.sheet_properties.pageSetUpPr.fitToPage=True
    if not wb.sheetnames:wb.create_sheet("Sin registros")
    output=io.BytesIO();wb.save(output);return output.getvalue()


def export_pdf(table,records):
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    font="Helvetica";bold="Helvetica-Bold"
    for root in ("/usr/share/fonts/truetype/dejavu",os.path.join(os.path.dirname(__file__),"fonts")):
        regular=os.path.join(root,"DejaVuSans.ttf");heavy=os.path.join(root,"DejaVuSans-Bold.ttf")
        if os.path.exists(regular) and os.path.exists(heavy):
            if "AlphaSans" not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont("AlphaSans",regular));pdfmetrics.registerFont(TTFont("AlphaSansBold",heavy))
            font="AlphaSans";bold="AlphaSansBold";break
    out=io.BytesIO();doc=SimpleDocTemplate(out,pagesize=landscape(A4),leftMargin=28,rightMargin=28,topMargin=30,bottomMargin=32,title="Alpha Builders — "+table)
    styles=getSampleStyleSheet()
    cell=ParagraphStyle("AlphaCell",fontName=font,fontSize=8,leading=11,wordWrap="CJK")
    head=ParagraphStyle("AlphaHead",parent=cell,fontName=bold,textColor=colors.white)
    title=ParagraphStyle("AlphaTitle",fontName=bold,fontSize=15,leading=20,spaceAfter=10)
    section=ParagraphStyle("AlphaSection",fontName=bold,fontSize=10,leading=14,spaceBefore=8,spaceAfter=5)
    width=doc.width-12
    def para(value,style=cell):return Paragraph(html_escape(text_value(value)).replace("\n","<br/>"),style)
    def tab(headers,rows,widths=None):
        content=[[para(h,head) for h in headers]]+[[para(v) for v in row] for row in rows]
        value=Table(content,colWidths=widths or [width/len(headers)]*len(headers),repeatRows=1,hAlign="LEFT",splitByRow=1,splitInRow=1)
        value.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365D")),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#CBD5E1")),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]));return value
    story=[para("ALPHA BUILDERS | "+table.upper(),title),para("Generado: "+now_ec().strftime("%d/%m/%Y %H:%M")+f" · Registros: {len(records)}"),Spacer(1,8)]
    if table=="rendimientos":
        story += [para("Resumen por edificio, rubro y unidad",section),tab(["Edificio","Rubro","Unidad","Registros","Avance","HH verificadas","Históricos"],summary_rend(records)),para("Las HH históricas no verificadas se excluyen del total. Los avances se agrupan por rubro y unidad.")]
    for record in sorted(records,key=lambda r:(record_building(table,r),str(r.get("fecha",r.get("fecha_compromiso",""))))):
        meta,sections,photos=sections_for(table,record)
        story += [para(record_building(table,record)+" · "+display_date(record_date(table,record)),section),tab(["Campo","Valor"],meta,[width*.28,width*.72])]
        for label,headers,rows in sections:
            if rows:story += [para(label,section),tab(headers,rows)]
        for label,photo in photos:
            stream=image_stream(photo)
            if stream:
                im=PDFImage(stream);ratio=min(350/im.imageWidth,210/im.imageHeight);im.drawWidth=im.imageWidth*ratio;im.drawHeight=im.imageHeight*ratio
                story.append(KeepTogether([para(label,section),im,Spacer(1,8)]))
            else:story.append(para(label+": fotografía no interpretable; se conserva el original."))
        story.append(Spacer(1,15))
    def footer(canvas,doc):
        canvas.saveState();canvas.setFont(font,8);canvas.drawString(28,15,"Alpha Builders | "+APP_VERSION);canvas.drawRightString(landscape(A4)[0]-28,15,f"Página {doc.page}");canvas.restoreState()
    doc.build(story,onFirstPage=footer,onLaterPages=footer);return out.getvalue()


def export_controls(table,records,key):
    if not records:return
    with st.expander("Exportar Excel, PDF o CSV",expanded=False):
        st.caption("Se generará únicamente el formato solicitado, con los registros del filtro actual.")
        format_=st.selectbox("Formato",["Excel (.xlsx)","PDF (.pdf)","CSV (.csv)"],key="format_"+key)
        signature=hashlib.sha256(json_text([
            (str(r.get("id")), str(r.get("revision")), str(r.get("updated_at")))
            for r in records
        ]).encode()).hexdigest()
        artifact_key=(key,signature,format_)
        if st.button("Preparar descarga",key="prepare_"+key):
            try:
                with st.spinner("Preparando informe…"):
                    if format_.startswith("Excel"):content=export_excel(table,records);ext="xlsx";mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    elif format_.startswith("PDF"):content=export_pdf(table,records);ext="pdf";mime="application/pdf"
                    else:content=csv_export(table,records);ext="csv";mime="text/csv"
                st.session_state["_export_artifact"]={"key":artifact_key,"content":content,"ext":ext,"mime":mime}
            except Exception as exc:notice_error(exc)
        artifact=st.session_state.get("_export_artifact")
        if artifact and artifact["key"]==artifact_key:
            st.download_button("Descargar informe",artifact["content"],file_name=f"Alpha_{table}_{now_ec():%Y%m%d}.{artifact['ext']}",mime=artifact["mime"],key="download_"+key)


def show_record(table,record):
    meta,sections,photos=sections_for(table,record)
    st.dataframe(pd.DataFrame([[k,text_value(v)] for k,v in meta],columns=["Campo","Valor"]),hide_index=True,use_container_width=True)
    for title,headers,rows in sections:
        if rows:
            st.markdown("**"+title+"**")
            st.dataframe(pd.DataFrame([[text_value(v) for v in row] for row in rows],columns=headers),hide_index=True,use_container_width=True)
    for title,photo in photos:
        stream=image_stream(photo)
        if stream:st.image(stream,caption=title,width=320)


def history(table, records, repo, key, editable=True):
    """Biblioteca: edificio → mes → día → detalle. Los IDs son únicamente claves internas."""
    st.markdown('#### 📚 Archivo de registros')
    if not records:
        st.info('Todavía no hay registros de este tipo. Los nuevos aparecerán organizados por edificio y fecha.');return
    counts=defaultdict(int)
    for record in records:counts[record_building(table,record)]+=1
    buildings=sorted(counts,key=lambda b:(b=='Sin Edificio Asignado',b.casefold()))
    c1,c2=st.columns([1.4,1])
    with c1:building=st.selectbox('🏢 Edificio', ['Todos los edificios']+buildings,
        format_func=lambda b:b if b=='Todos los edificios' else f'{b} · {counts[b]}',key='hist_build_'+key)
    filtered=[r for r in records if building=='Todos los edificios' or record_building(table,r)==building]
    available_months=sorted({record_date(table,r)[:7] for r in filtered if len(record_date(table,r))>=7},reverse=True)
    with c2:month=st.selectbox('🗓️ Periodo', ['Todos los periodos']+available_months,
        format_func=lambda m:m if m=='Todos los periodos' else display_month(m),key='hist_month_'+key)
    if month!='Todos los periodos':filtered=[r for r in filtered if record_date(table,r).startswith(month)]
    if table=='incidencias':
        status=st.segmented_control('Estado', ['Todas','Abierta','Cerrada'],default='Todas',key='hist_status_'+key)
        if status not in ('Todas',None):filtered=[r for r in filtered if r.get('estado')==status]
    search=st.text_input('Buscar en los registros',placeholder='Rubro, área, responsable o edificio…',key='hist_search_'+key)
    if search.strip():
        term=search.strip().casefold()
        filtered=[r for r in filtered if term in ' '.join((record_building(table,r), record_title(table,r),record_note(table,r))).casefold()]
    a,b=st.columns([2,1])
    with a:st.caption(f'**{len(filtered)} registros** · Ordenados del más reciente al más antiguo')
    with b:st.caption('📤 Exportación del filtro actual')
    export_controls(table,filtered,'hist_export_'+key)
    if not filtered:
        st.info('No se encontraron registros con estos filtros.');return
    ordered=sorted(filtered,key=lambda r:(record_date(table,r),str(r.get('id') or '')),reverse=True)
    visible=page_slice(ordered,'hist_block_'+key,12)
    by_building=defaultdict(list)
    for record in visible:by_building[record_building(table,record)].append(record)
    # Único edificio visible evita el encabezado redundante.
    for building_name, items in sorted(by_building.items(),key=lambda p:p[0].casefold()):
        if building=='Todos los edificios' or len(by_building)>1:
            st.markdown('##### 🏢 '+html_escape(building_name))
        for day,day_records in grouped_records(items,table):
            day_heading(day,len(day_records))
            for record in day_records:
                identifier=key+':'+str(record['id'])
                title=record_title(table,record)
                note=record_note(table,record)
                with st.expander(f'📄 {title}  ·  {note}',expanded=False):
                    if st.checkbox('Ver información y fotografías',key='detail_'+identifier):
                        show_record(table,record)
                    if editable:
                        ca,cb=st.columns([1,1])
                        if ca.button('✏️ Editar registro',key='edit_'+identifier,use_container_width=True):
                            try:
                                repo._access(table,record,write=True)
                                if table=='checklists':module='Checklist';factory=checklist_data
                                elif table=='rendimientos':module='Rendimientos';factory=rend_data
                                elif table=='incidencias':module='Incidencias';factory=incident_data
                                else:
                                    master=is_master_book(record)
                                    module='Libro maestro' if master else 'Libro oficial'
                                    factory=maestro_data if master else official_data
                                begin_editor(module,repo.user,factory(record),record)
                                st.session_state['_next_route']=module;st.rerun()
                            except Exception as exc:notice_error(exc)
                        with cb.popover('🗑️ Eliminar registro',use_container_width=True):
                            st.warning('Esta acción no se puede deshacer.')
                            if st.button('Confirmar eliminación',key='delete_'+identifier):
                                try:repo.delete(table,record);st.session_state['_flash']='Registro eliminado.';st.rerun()
                                except Exception as exc:notice_error(exc)




def read_worker_file(upload,default_building,allowed_buildings):
    raw = upload.getvalue()
    if len(raw) > 10*1024*1024:
        raise ValueError("El archivo supera 10 MB.")
    return _read_worker_file_cached(raw, upload.name.lower(), default_building, tuple(allowed_buildings))


@st.cache_data(ttl=900, show_spinner=False, max_entries=32)
def _read_worker_file_cached(raw, file_name, default_building, allowed_buildings):
    import unicodedata
    if file_name.endswith(".csv"):
        try:text=raw.decode("utf-8-sig")
        except UnicodeDecodeError:text=raw.decode("latin-1")
        try:delimiter=csv.Sniffer().sniff(text[:8192],delimiters=";,\t").delimiter
        except csv.Error:delimiter=";" if ";" in text.splitlines()[0] else ","
        df=pd.read_csv(io.StringIO(text),sep=delimiter,dtype=str,keep_default_na=False)
    else:df=pd.read_excel(io.BytesIO(raw),dtype=str,keep_default_na=False)
    def normalized(value):
        value="".join(c for c in unicodedata.normalize("NFKD",str(value)) if not unicodedata.combining(c))
        return re.sub(r"[^a-z0-9]","",value.lower())
    mapping={}
    aliases={"nombre":{"nombre","nombrecompleto","trabajador"},"cargo":{"cargo","especialidad","cargoespecialidad"},"edificio":{"edificio","proyecto","edificioasignado","edificioproyecto"}}
    for column in df.columns:
        for target,names in aliases.items():
            if normalized(column) in names:mapping[target]=column
    if not {"nombre","cargo"}.issubset(mapping):raise ValueError("El archivo debe incluir las columnas Nombre y Cargo. También admite Nombre Completo y Cargo / Especialidad.")
    records=[];seen=set()
    for index,row in df.iterrows():
        name=str(row[mapping["nombre"]]).strip().upper();role=str(row[mapping["cargo"]]).strip().upper()
        if not name and not role:continue
        building=str(row[mapping["edificio"]]).strip() if "edificio" in mapping else default_building
        building=building or default_building
        canonical=next((b for b in allowed_buildings if b.casefold()==building.casefold()),None)
        if not name or not role:raise ValueError(f"Fila {index+2}: nombre o cargo vacío.")
        if not canonical:raise ValueError(f"Fila {index+2}: edificio no asignado a su cuenta: {building}.")
        identity=(name,canonical)
        if identity not in seen:records.append({"nombre":name,"cargo":role,"edificio":canonical});seen.add(identity)
    if len(records)>1000:raise ValueError("Importe como máximo 1.000 trabajadores por archivo.")
    return records


def import_workers(repo,records,operation_id):
    repo._verify_write_session()
    existing={(str(r.get("nombre") or "").strip().upper(),r.get("edificio")) for r in repo.own("trabajadores")}
    payload=[]
    for index,record in enumerate(records):
        if record["edificio"] not in repo.projects and not repo.admin:raise PermissionError("Edificio no autorizado.")
        identity=(record["nombre"].strip().upper(),record["edificio"])
        if identity in existing:continue
        existing.add(identity)
        payload.append({**record,"usuario_email":repo.email,"client_request_id":str(uuid.uuid5(uuid.NAMESPACE_URL,repo.email+":"+operation_id+":"+str(index)))})
    if not payload:return 0
    # Un único INSERT de un lote: Postgres lo confirma o lo rechaza completo.
    result=database().table("trabajadores").upsert(payload,on_conflict="client_request_id",ignore_duplicates=True).execute()
    if result.data is None:raise RuntimeError("No se confirmó la importación.")
    invalidate();return len(result.data)


def personal_module(repo):
    """Nómina segura: une solo para visualización SQL + respaldo histórico.

    Nunca borra trabajadores al filtrar, importar, actualizar o mostrar respaldos.
    """
    page_heading('Gestión de personal','Nómina por edificio',
        'Consulta tus cuadrillas, registra colaboradores y recupera trabajadores antiguos sin perder información.')
    projects=project_options(repo.user)
    current=sorted(repo.own('trabajadores'),key=lambda r:(str(r.get('edificio') or '').casefold(),str(r.get('nombre') or '').casefold()))
    historic=legacy_worker_backup(repo)
    signatures={(str(r.get('nombre') or '').strip().upper(),str(r.get('edificio') or '').strip().casefold()) for r in current}
    pending=[w for w in historic if (w['nombre'],w['edificio'].casefold()) not in signatures]
    shown=current+pending
    c1,c2,c3=st.columns(3)
    c1.metric('Personal registrado',len(current))
    c2.metric('Edificios en nómina',len({w.get('edificio') for w in shown if w.get('edificio')}))
    c3.metric('Pendientes de recuperar',len(pending))
    area=st.segmented_control('Sección',['👥 Nómina','➕ Registrar','📥 Importar','✏️ Gestionar'],
        default='👥 Nómina',key='v8_personal_section',label_visibility='collapsed') or '👥 Nómina'
    if area=='👥 Nómina':
        if pending:
            st.info(f'Encontramos {len(pending)} trabajador(es) en el respaldo de la versión anterior que aún no aparecen en la tabla actual. Se muestran junto a los demás, sin modificar tu base de datos.')
            canonical={p.casefold():p for p in projects}
            eligible=[{**w,'edificio':canonical[w['edificio'].casefold()]} for w in pending if w['edificio'].casefold() in canonical and w['nombre'] and w['cargo']]
            excluded=len(pending)-len(eligible)
            with st.expander('🛟 Recuperar mi nómina anterior',expanded=False):
                st.markdown('La copia histórica permanece intacta. **La recuperación solo añade trabajadores que no existan ya en la tabla actual.**')
                if excluded:
                    st.warning(f'{excluded} trabajador(es) conservan un edificio no asignado actualmente a tu cuenta. No se migrarán sin corregir su asignación.')
                if st.button(f'Incorporar {len(eligible)} trabajador(es) válidos a mi nómina',
                    key='v8_workers_recover',disabled=not eligible,type='primary'):
                    try:
                        count=import_workers(repo,[{k:w[k] for k in ('nombre','cargo','edificio')} for w in eligible],
                            'legacy-restore-'+hashlib.sha256(json_text(eligible).encode()).hexdigest())
                        st.session_state['_flash']=f'{count} trabajador(es) recuperados; no se eliminaron ni sobrescribieron los anteriores.'
                        st.rerun()
                    except Exception as exc:notice_error(exc)
        if not shown:
            st.info('Tu nómina está vacía. Empieza en Registrar o Importar.');return
        building_counts=defaultdict(int)
        for worker in shown:building_counts[worker.get('edificio') or 'Sin edificio asignado']+=1
        buildings=sorted(building_counts,key=lambda b:b.casefold())
        ca,cb=st.columns([1.3,2])
        with ca:building=st.selectbox('🏢 Edificio',['Todos los edificios']+buildings,
            format_func=lambda b:b if b=='Todos los edificios' else f'{b} · {building_counts[b]}',key='v8_personal_build')
        with cb:query=st.text_input('🔎 Buscar trabajador o especialidad',placeholder='Nombre o cargo…',key='v8_personal_query')
        visible=[w for w in shown if (building=='Todos los edificios' or (w.get('edificio') or 'Sin edificio asignado')==building) and
             (not query.strip() or query.strip().casefold() in (str(w.get('nombre') or '')+' '+str(w.get('cargo') or '')).casefold())]
        st.caption(f'{len(visible)} colaboradores encontrados · Los marcados como «Respaldo» aún no están sincronizados.')
        if not visible:st.info('No encontramos coincidencias.');return
        display=page_slice(visible,'v8_people_block',25)
        groups=defaultdict(list)
        for w in display:groups[w.get('edificio') or 'Sin edificio asignado'].append(w)
        for name,items in sorted(groups.items(),key=lambda x:x[0].casefold()):
            st.markdown('##### 🏢 '+html_escape(name))
            parts=[]
            for w in items:
                badge='<span class="ab-chip ab-chip-legacy">Respaldo</span>' if w.get('_legacy') else ''
                parts.append('<div class="ab-item"><div><div class="ab-item-title">👷 '+html_escape(str(w.get('nombre') or ''))+
                    '</div><div class="ab-item-tags"><span class="ab-chip">'+html_escape(str(w.get('cargo') or ''))+
                    '</span>'+badge+'</div></div><span class="ab-subtle">'+('Pendiente' if w.get('_legacy') else 'En nómina')+'</span></div>')
            st.markdown('<div class="ab-list-wrap">'+''.join(parts)+'</div>',unsafe_allow_html=True)
        with st.expander('📤 Descargar nómina actual',expanded=False):
            # Los respaldos se pueden exportar sin fingir que están confirmados en SQL.
            st.download_button('📄 Exportar lista visible (CSV)',
                pd.DataFrame([{'Nombre':w.get('nombre'),'Cargo':w.get('cargo'),'Edificio':w.get('edificio'),
                'Origen':'Respaldo histórico' if w.get('_legacy') else 'Nómina actual'} for w in visible]).to_csv(index=False,sep=';').encode('utf-8-sig'),
                file_name='Nomina_Alpha_Builders.csv',mime='text/csv',key='v8_people_csv')
            sql_visible=[w for w in visible if not w.get('_legacy')]
        if sql_visible:export_controls('trabajadores',sql_visible,'personal_v8')
        return
    if area=='➕ Registrar':
        if not projects:
            st.info('Solicita al administrador tu asignación a un edificio para registrar personal.');return
        st.markdown('#### Añadir un trabajador')
        with st.form('personal_new',clear_on_submit=True):
            name=st.text_input('Nombre completo',placeholder='Ej. Juan Carlos Pérez')
            role=st.text_input('Cargo / especialidad',placeholder='Ej. Albañil, ayudante, eléctrico…')
            building=st.selectbox('Edificio',projects)
            submit=st.form_submit_button('💾 Guardar trabajador',type='primary')
        if submit:
            try:
                if not name.strip() or not role.strip():raise ValueError('Completa el nombre y el cargo.')
                operation=st.session_state.setdefault('_worker_operation',uuid.uuid4().hex)
                count=import_workers(repo,[{'nombre':name.strip().upper(),'cargo':role.strip().upper(),'edificio':building}],operation)
                st.session_state.pop('_worker_operation',None)
                st.session_state['_flash']='Trabajador registrado.' if count else 'El trabajador ya existe en ese edificio.'
                st.rerun()
            except Exception as exc:notice_error(exc)
        return
    if area=='📥 Importar':
        if not projects:st.info('Necesitas un edificio asignado antes de importar trabajadores.');return
        st.markdown('#### Importar nómina Excel o CSV')
        with st.container(border=True):
            building=st.selectbox('Edificio predeterminado',projects,key='import_building')
            st.caption('Columnas: Nombre, Cargo y opcionalmente Edificio. No reemplazamos ni borramos trabajadores registrados.')
            upload=st.file_uploader('Selecciona el archivo de nómina',type=['csv','xlsx'],key='workers_upload')
            if upload:
                try:
                    records=read_worker_file(upload,building,projects)
                    st.dataframe(pd.DataFrame(records),hide_index=True,use_container_width=True)
                    st.caption(f'{len(records)} filas válidas. Los duplicados con el mismo nombre y edificio se omitirán.')
                    if st.button('Confirmar importación',type='primary',key='import_confirm'):
                        count=import_workers(repo,records,hashlib.sha256(upload.getvalue()+building.encode()).hexdigest())
                        st.session_state['_flash']=f'{count} trabajador(es) nuevos confirmados; se conservó la nómina existente.';st.rerun()
                except Exception as exc:notice_error(exc)
        with st.expander('🤝 Importar trabajadores de un compañero',expanded=False):
            if st.checkbox('Consultar nóminas compartidas',key='workers_load_colleagues'):
                users=colleagues(repo)
                if users:
                    colleague=st.selectbox('Compañero',users,format_func=lambda u:full_name(u)+' · '+u['correo'],key='workers_colleague')
                    records=[r for r in repo.shared('trabajadores',colleague['correo']) if r.get('edificio') in repo.projects]
                    st.caption(f'{len(records)} trabajadores en edificios compartidos.')
                    if st.button('Añadir a mi nómina',key='copy_workers'):
                        try:
                            count=import_workers(repo,[{k:r[k] for k in ('nombre','cargo','edificio')} for r in records],
                                'copy:'+colleague['correo']+':'+hashlib.sha256(json_text(records).encode()).hexdigest())
                            st.session_state['_flash']=f'{count} trabajador(es) importados, sin sobrescribir los anteriores.';st.rerun()
                        except Exception as exc:notice_error(exc)
                else:st.info('No hay compañeros con edificios compartidos.')
        return
    if area=='✏️ Gestionar':
        if not current:st.info('Aún no hay trabajadores sincronizados. Consulta Nómina para recuperar tu respaldo anterior.');return
        st.markdown('#### Editar un trabajador')
        st.caption('Selecciona a una persona. Los cambios solo se aplican al guardar.')
        selected=st.selectbox('Trabajador',current,
            format_func=lambda r:f"{r.get('nombre') or 'Sin nombre'} · {r.get('cargo') or 'Sin cargo'} · {r.get('edificio') or 'Sin edificio'}",
            key='v8_worker_edit')
        # Solo claves internas; no se presentan en pantalla.
        identity=f"{selected['id']}:{selected['revision']}"
        with st.form('edit_worker_'+identity):
            name=st.text_input('Nombre',value=selected.get('nombre') or '')
            role=st.text_input('Cargo / especialidad',value=selected.get('cargo') or '')
            options=project_options(repo.user,selected.get('edificio'))
            building=st.selectbox('Edificio asignado',options,index=options.index(selected['edificio']) if selected.get('edificio') in options else 0)
            submit=st.form_submit_button('💾 Actualizar trabajador',type='primary')
        if submit:
            try:
                if not name.strip() or not role.strip():raise ValueError('Nombre y cargo son obligatorios.')
                repo.save('trabajadores',{'nombre':name.strip().upper(),'cargo':role.strip().upper(),'edificio':building},selected)
                st.session_state['_flash']='Trabajador actualizado.';st.rerun()
            except Exception as exc:notice_error(exc)
        with st.popover('🗑️ Eliminar trabajador'):
            st.warning('Eliminar es permanente. Los reportes históricos conservarán sus nombres.')
            if st.button('Confirmar eliminación del trabajador',key='v8_worker_delete'):
                try:repo.delete('trabajadores',selected);st.session_state['_flash']='Trabajador eliminado.';st.rerun()
                except Exception as exc:notice_error(exc)




def colleagues(repo):
    users=cached_rows("usuarios",columns="correo,nombres,apellidos,cargo,edificios,estado",order="correo")
    return [u for u in users if u.get("correo")!=repo.email and active_user(u) and set(obj(u.get("edificios"),[])).intersection(repo.projects)]


def collaborative_module(repo):
    st.header("Espacio colaborativo")
    users=colleagues(repo)
    if not users:st.info("No hay compañeros activos en sus edificios asignados.");return
    user=st.selectbox("Compañero",users,format_func=lambda u:full_name(u)+" · "+u["correo"])
    st.caption("Solo se muestran registros de edificios compartidos: "+", ".join(sorted(set(obj(user.get("edificios"),[])).intersection(repo.projects))))
    table=st.selectbox("Registros",["checklists","inspecciones","incidencias","rendimientos","trabajadores"],format_func=lambda x:{"checklists":"Checklists","inspecciones":"Libros de obra","incidencias":"Incidencias","rendimientos":"Rendimientos","trabajadores":"Personal"}[x],key="colab_table")
    history(table,repo.shared(table,user["correo"]),repo,"colab_"+table+user["correo"],editable=False)


def logout(cookie):
    raw=st.session_state.get("_session_raw")
    if raw:database().table("alpha_sessions").delete().eq("token_hash",digest_token(raw)).execute()
    try:cookie.delete(COOKIE_NAME,key="logout_cookie")
    except KeyError:pass
    for key in list(st.session_state):del st.session_state[key]
    st.rerun()


def account_module(repo,cookie):
    st.header("Configuración de cuenta")
    user=repo.user
    with st.form("profile"):
        names=st.text_input("Nombres",value=user.get("nombres") or "");surname=st.text_input("Apellidos",value=user.get("apellidos") or "")
        photo=st.file_uploader("Foto de perfil",type=["png","jpg","jpeg","webp"])
        submit=st.form_submit_button("Guardar perfil")
    if submit:
        try:
            if not names.strip() or not surname.strip():raise ValueError("Complete nombres y apellidos.")
            payload={"nombres":names.strip(),"apellidos":surname.strip()}
            if photo:payload["foto_b64"]=image_b64(photo)
            result=database().table("usuarios").update(payload).eq("correo",repo.email).eq("revision",user["revision"]).execute()
            if not result.data:raise ConflictError("El perfil cambió en otra sesión. Recargue la página.")
            invalidate();st.session_state["_flash"]="Perfil actualizado.";st.rerun()
        except Exception as exc:notice_error(exc)
    with st.form("account_access_request"):
        role=st.selectbox("Solicitar cargo",CARGOS_DISPONIBLES,index=CARGOS_DISPONIBLES.index(user["cargo"]) if user.get("cargo") in CARGOS_DISPONIBLES else 0)
        buildings=st.multiselect("Solicitar edificios",EDIFICIOS_ALPHA,default=[x for x in user.get("edificios",[]) if x in EDIFICIOS_ALPHA])
        submit=st.form_submit_button("Enviar solicitud al administrador")
    if submit:
        try:
            result=database().table("alpha_access_requests").upsert({"correo":repo.email,"cargo":role,"edificios":buildings,"estado":"Pendiente"},on_conflict="correo").execute()
            if not result.data:raise RuntimeError("Solicitud no confirmada")
            st.success("Solicitud registrada en el panel administrador.")
        except Exception as exc:notice_error(exc)
    password_change(repo,cookie)


def password_change(repo,cookie):
    with st.form("change_password"):
        old=st.text_input("Contraseña actual",type="password")
        new=st.text_input("Nueva contraseña (mínimo 10 caracteres)",type="password")
        repeat=st.text_input("Repetir nueva contraseña",type="password")
        submit=st.form_submit_button("Cambiar contraseña y cerrar sesiones")
    if submit:
        try:
            validate_password(new,repeat)
            if not allowed_attempt("password_change",repo.email):raise ValueError("Demasiados intentos. Espere 15 minutos.")
            row=fetch_user(repo.email,credentials=True)
            if not password_ok(old,row.get("password")):raise ValueError("La contraseña actual no es correcta.")
            response=database().rpc("alpha_set_password",{"p_email":repo.email,"p_password":password_hash(new),"p_must_change":False}).execute()
            if response.data is not True:raise RuntimeError("Cambio de contraseña no confirmado.")
            logout(cookie)
        except Exception as exc:notice_error(exc)


def admin_module(repo):
    if not repo.admin:raise PermissionError("Acceso reservado al administrador.")
    st.header("Panel administrador")
    section=st.selectbox("Administrar",["Registros","Usuarios y permisos","Solicitudes","PIN de acceso","Revisión de migración"],key="admin_section")
    if section=="Registros":
        table=st.selectbox("Tipo",["checklists","inspecciones","incidencias","rendimientos","trabajadores"],key="admin_table")
        records=repo.all(table)
        users=sorted({r.get("usuario_email") or "Sin propietario" for r in records})
        selected=st.selectbox("Usuario",["Todos"]+users,key="admin_owner")
        records=[r for r in records if selected=="Todos" or (r.get("usuario_email") or "Sin propietario")==selected]
        # El administrador supervisa y exporta; solo edita sus propios formularios desde sus módulos.
        history(table,records,repo,"admin_"+table,editable=False)
    elif section=="Usuarios y permisos":
        users=all_rows("usuarios",columns="correo,nombres,apellidos,cargo,edificios,estado,es_admin,revision",order="correo")
        if not users:return
        target=st.selectbox("Cuenta",users,format_func=lambda u:full_name(u)+" · "+u["correo"]+" · "+str(u.get("estado")))
        with st.form("admin_user_"+target["correo"]+str(target["revision"])):
            role=st.selectbox("Cargo autorizado",CARGOS_DISPONIBLES,index=CARGOS_DISPONIBLES.index(target["cargo"]) if target.get("cargo") in CARGOS_DISPONIBLES else 0)
            buildings=st.multiselect("Edificios autorizados",EDIFICIOS_ALPHA,default=[x for x in obj(target.get("edificios"),[]) if x in EDIFICIOS_ALPHA])
            status=st.selectbox("Estado",["Activo","Pendiente","Suspendido"],index=["Activo","Pendiente","Suspendido"].index(target["estado"]) if target.get("estado") in ["Activo","Pendiente","Suspendido"] else 2)
            administrator=st.checkbox("Permisos de administrador",value=bool(target.get("es_admin")))
            submit=st.form_submit_button("Guardar permisos")
        if submit:
            try:
                if target.get("es_admin") and (not administrator or status!="Activo") and sum(bool(u.get("es_admin")) and active_user(u) for u in users)<=1:raise ValueError("Debe conservar al menos un administrador activo.")
                response=database().rpc("alpha_update_permissions",{"p_email":target["correo"],"p_revision":target["revision"],"p_cargo":role,"p_edificios":buildings,"p_estado":status,"p_admin":administrator}).execute()
                if response.data is not True:raise ConflictError("La cuenta cambió en otra sesión. Recargue.")
                invalidate();st.session_state["_flash"]="Permisos actualizados; las sesiones de esa cuenta se revocaron.";st.rerun()
            except Exception as exc:notice_error(exc)
        with st.popover("Restablecimiento administrado de contraseña"):
            verified=st.checkbox("He verificado la identidad del titular por un canal de confianza",key="verified_"+target["correo"])
            if st.button("Emitir contraseña temporal",disabled=not verified,key="reset_admin_"+target["correo"]):
                try:
                    temporary=secrets.token_urlsafe(15)
                    response=database().rpc("alpha_set_password",{"p_email":target["correo"],"p_password":password_hash(temporary),"p_must_change":True}).execute()
                    if response.data is not True:raise RuntimeError("Cambio no confirmado")
                    st.session_state["_temporary_password"]={"correo":target["correo"],"password":temporary}
                except Exception as exc:notice_error(exc)
            temporary=st.session_state.get("_temporary_password")
            if temporary and temporary["correo"]==target["correo"]:
                st.code(temporary["password"]);st.caption("Entréguela únicamente al titular. Tendrá que cambiarla al entrar.")
                if st.button("Ocultar contraseña temporal"):st.session_state.pop("_temporary_password",None);st.rerun()
    elif section=="Solicitudes":
        requests=all_rows("alpha_access_requests",filters={"estado":"Pendiente"},order="correo")
        if not requests:st.info("No hay solicitudes pendientes.")
        for request in requests:
            with st.container(border=True):
                st.write(request["correo"],request["cargo"],", ".join(request.get("edificios") or []))
                if st.button("Aprobar solicitud",key="approve_"+request["correo"]):
                    try:
                        response=database().rpc("alpha_approve_access",{"p_email":request["correo"],"p_revision":request["revision"]}).execute()
                        if response.data is not True:raise ConflictError("La solicitud cambió. Recargue.")
                        invalidate();st.rerun()
                    except Exception as exc:notice_error(exc)
                if st.button("Rechazar solicitud",key="reject_"+request["correo"]):
                    result=database().table("alpha_access_requests").update({"estado":"Rechazada"}).eq("correo",request["correo"]).eq("revision",request["revision"]).execute()
                    if not result.data:st.warning("La solicitud cambió. Actualice.")
                    else:st.rerun()
    elif section=="PIN de acceso":
        with st.form("admin_pin"):
            pin=st.text_input("Nuevo PIN de 4 dígitos",type="password",max_chars=4)
            submit=st.form_submit_button("Actualizar PIN")
        if submit:
            try:
                if not re.fullmatch(r"\d{4}",pin):raise ValueError("El PIN debe tener cuatro dígitos.")
                set_config("access_pin",pin);st.success("PIN actualizado. Compártalo solo con las personas autorizadas.")
            except Exception as exc:notice_error(exc)
    else:
        issues=all_rows("alpha_migration_issues",order="id")
        if issues:st.dataframe(pd.DataFrame(issues),hide_index=True,use_container_width=True);st.info("Los casos ambiguos conservan los datos originales. Revise las coincidencias antes de editar o eliminar personal.")
        else:st.success("No se registraron conflictos de migración.")


def dashboard(repo):
    user = repo.user
    today = now_ec().date()
    date_label = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")[today.weekday()]
    date_label += f", {today.day} de {NOMBRES_MESES[today.month]} de {today.year}"
    email = user["correo"]
    # Vista ligera: nunca descargar fotos/JSON completos para mostrar 4 indicadores.
    own_filter = (("usuario_email", email),)
    checks = cached_rows("checklists", own_filter, columns="fecha") if user.get("cargo") != "Maestro Mayor" else []
    books = cached_rows("inspecciones", own_filter, columns="fecha,tipo:datos->>Tipo_Registro")
    yield_rows = cached_rows("rendimientos", own_filter, columns="estado")
    sql_workers = cached_rows("trabajadores", own_filter, columns="nombre,edificio")
    signatures = {(str(w.get("nombre") or "").strip().upper(), str(w.get("edificio") or "").strip().casefold()) for w in sql_workers}
    workers_count = len(sql_workers) + sum(
        (w["nombre"], w["edificio"].casefold()) not in signatures
        for w in legacy_worker_backup(repo)
    )
    master_role = user.get("cargo") == "Maestro Mayor"
    # Los libros antiguos sin marca de tipo se inspeccionan solo en la fecha de hoy.
    today_text = today.isoformat()
    role_books_today = [r for r in books if str(r.get("fecha"))[:10] == today_text and
                        (str(r.get("tipo") or "") == "Libro_Obra_Maestro") == master_role and bool(r.get("tipo"))]
    if not role_books_today and any(str(r.get("fecha"))[:10] == today_text and not r.get("tipo") for r in books):
        legacy_today = cached_rows("inspecciones", (("usuario_email",email),("fecha",today_text)), columns="fecha,datos")
        role_books_today = [r for r in legacy_today if is_master_book(r) == master_role]
    books_for_role = role_books_today
    valid_yields = [r for r in yield_rows if r.get("estado") in {"CUMPLE META", "BAJO RENDIMIENTO"}]
    efficiency = round(100 * sum(r.get("estado") == "CUMPLE META" for r in valid_yields) / len(valid_yields)) if valid_yields else None
    efficiency_label = f"{efficiency}%" if efficiency is not None else "—"
    circle_value = efficiency if efficiency is not None else 0
    efficiency_detail = "Registros con meta verificada" if efficiency is not None else "Sin meta verificable"
    color = "#10b981" if efficiency is not None and efficiency >= 75 else "#f59e0b" if efficiency is not None and efficiency >= 50 else "#ef4444" if efficiency is not None else "#94a3b8"
    incidence_projects = tuple(sorted(repo.projects))
    open_incidents = sum(r.get("estado") == "Abierta" for r in
        (cached_rows("incidencias", (), (("proyecto", incidence_projects),), columns="estado") if incidence_projects else []))
    tags = "".join('<span class="edificio-tag-badge">' + html_escape(str(x)) + "</span>" for x in user.get("edificios") or [])
    def milestone(label, checked):
        badge = '<span class="milestone-status-done">✓ Registrado</span>' if checked else '<span class="milestone-status-pending">⏳ Pendiente</span>'
        return '<div class="milestone-item"><span class="milestone-name">' + html_escape(label) + '</span>' + badge + '</div>'
    milestones = milestone("Libro de Obra Maestro" if user.get("cargo") == "Maestro Mayor" else "Checklist", bool(books_for_role and any(str(r.get("fecha"))[:10] == today_text for r in books_for_role)) if user.get("cargo") == "Maestro Mayor" else any(str(r.get("fecha"))[:10] == today_text for r in checks))
    if user.get("cargo") != "Maestro Mayor":
        milestones += milestone("Libro de Obra", any(str(r.get("fecha"))[:10] == today_text for r in books_for_role))
    output = (
        '<div class="smart-dashboard-container"><div class="smart-header-bar"><div>'
        '<div class="smart-title">Alpha Builders | Portal de Obra</div>'
        '<div class="smart-user-sub">' + html_escape(full_name(user)) + ' &bull; <b>' + html_escape(str(user.get("cargo") or "")) + '</b> ' + tags + '</div>'
        '</div><div class="smart-pill">📅 ' + html_escape(date_label) + '</div></div>'
        '<div class="widgets-grid">'
        '<div class="widget-glass-card"><div class="w-card-title"><span>🎯 Hitos Diarios</span><span>Hoy</span></div>' + milestones + '</div>'
        '<div class="widget-glass-card"><div class="w-card-title"><span>⚡ Rendimiento</span><span>Meta</span></div>'
        '<div class="donut-container"><svg class="donut-chart-svg" viewBox="0 0 36 36">'
        '<circle class="donut-bg" cx="18" cy="18" r="15.9155" />'
        '<circle class="donut-progress" cx="18" cy="18" r="15.9155" stroke="' + color + '" stroke-dasharray="' + str(circle_value) + ', 100" />'
        '</svg><div><div class="donut-info-val" style="color:' + color + '!important">' + efficiency_label + '</div>'
        '<div class="donut-info-lbl">' + efficiency_detail + '</div></div></div></div>'
        '<div class="widget-glass-card"><div class="w-card-title"><span>🚨 Incidencias</span><span>Abiertas</span></div>'
        '<div class="stat-hero-number" style="color:#f87171!important">' + str(open_incidents) + '</div>'
        '<div style="font-size:.65rem;color:#94a3b8!important">En proyectos asignados</div></div>'
        '<div class="widget-glass-card"><div class="w-card-title"><span>👷 Personal</span><span>Registrados</span></div>'
        '<div class="stat-hero-number" style="color:#60a5fa!important">' + str(workers_count) + '</div>'
        '<div style="font-size:.65rem;color:#94a3b8!important">Personal a tu cargo</div></div>'
        '</div></div>'
    )
    st.markdown(output, unsafe_allow_html=True)
    st.markdown('<div style="height:4px"></div>', unsafe_allow_html=True)




@st.cache_data(show_spinner=False)
def _logo_data(uri, stamp_ns):
    with open(uri, "rb") as file:
        return base64.b64encode(file.read()).decode("ascii")


def sidebar_original(repo, cookie):
    """Perfil, logo y configuración visual del diseño de 6.000 líneas."""
    user = repo.user
    email = user["correo"]
    with st.sidebar:
        logo_file = next((name for name in ("alpha.473f0c2dc3c48a682723-2.webp", "images.png") if os.path.isfile(name)), None)
        if logo_file:
            try:
                encoded = _logo_data(logo_file, os.stat(logo_file).st_mtime_ns)
                mime = "image/webp" if logo_file.endswith(".webp") else "image/png"
                st.markdown('<div class="sidebar-logo-card"><img src="data:' + mime + ';base64,' + encoded + '" style="width:100%;height:auto;display:block;border-radius:8px;" alt="Alpha Builders"></div>', unsafe_allow_html=True)
            except OSError:
                st.markdown("### 🏗️ Alpha Builders")
        else:
            st.markdown("### 🏗️ Alpha Builders")
        photo = image_stream(user.get("foto_b64"))
        if photo:
            st.image(photo, use_container_width=True)
        else:
            profile_file = next((name for name in ("perfil.jpg", "perfil.png", "perfil.jpeg", "avatar.png") if os.path.isfile(name)), None)
            if profile_file:
                st.image(profile_file, use_container_width=True)
        badges = "".join('<span class="edificio-tag-badge" style="margin:2px;">' + html_escape(str(b)) + '</span>' for b in (user.get("edificios") or []))
        if not badges:
            badges = '<span style="font-size: .7rem; color: #94a3b8;">Sin proyectos asignados</span>'
        st.markdown(f'''<div class="sidebar-profile-box">
            <div class="sidebar-user-nombres">{html_escape(str(user.get('nombres') or ''))}</div>
            <div class="sidebar-user-apellidos">{html_escape(str(user.get('apellidos') or ''))}</div>
            <div class="sidebar-user-email">{html_escape(email)}</div>
            <div style="margin:5px 0"><span class="sidebar-user-cargo">{html_escape(str(user.get('cargo') or ''))}</span></div>
            <div style="display:flex;flex-wrap:wrap;justify-content:center;gap:3px">{badges}</div>
            </div>''', unsafe_allow_html=True)
        if repo.admin:
            st.markdown('<div style="text-align:center;background:#111827;border:1px solid #374151;border-radius:6px;padding:4px;font-size:.7rem;font-weight:800;color:white">ADMINISTRADOR GENERAL</div>', unsafe_allow_html=True)
        with st.expander("⚙️ Configuración de Cuenta", expanded=False):
            account_module(repo, cookie)
        if st.button("🔄 Actualizar información", key="sidebar_refresh", use_container_width=True):
            invalidate()
            st.rerun()
        if st.button("Cerrar Sesión", key="sidebar_logout", use_container_width=True):
            logout(cookie)
        st.caption(APP_VERSION)


def is_master_book(record):
    raw = obj(record.get("datos"), {})
    return isinstance(raw, list) or raw.get("Tipo_Registro") == "Libro_Obra_Maestro" or "Actividades_Maestro" in raw


def open_lazy_tabs(labels, key):
    """Siempre monta una sola pestaña. 1.55+ respeta el diseño nativo."""
    try:
        if "on_change" in inspect.signature(st.tabs).parameters:
            return st.tabs(labels, on_change="rerun", key=key), True
    except (TypeError, ValueError):
        pass
    # Versiones anteriores: la API nativa ejecuta TODOS los paneles y bloquea
    # la pantalla. La alternativa es un selector horizontal de carga diferida.
    import contextlib
    try:
        selected = st.radio("Navegación", labels, key=key+"_legacy", horizontal=True, label_visibility="collapsed")
    except TypeError:
        selected = st.radio("Navegación", labels, key=key+"_legacy", horizontal=True)
    class _LazyPanel:
        def __init__(self, open_): self.open = open_
        def __enter__(self):
            if self.open:
                self._ctx = st.container()
                self._ctx.__enter__()
            return self
        def __exit__(self, *args):
            if self.open:
                return self._ctx.__exit__(*args)
            return False
    return [_LazyPanel(label == selected) for label in labels], True


def history_rend_by_building(records, repo):
    """Resumen general, edificios y bloques cronológicos compactos."""
    st.markdown('#### 📊 Rendimientos registrados')
    if not records:
        st.info('Todavía no hay rendimientos. Los registros se clasificarán automáticamente por edificio, mes y jornada.');return
    by_build=defaultdict(list)
    for record in records:by_build[record_building('rendimientos',record)].append(record)
    buildings=sorted(by_build,key=lambda x:(x=='Sin Edificio Asignado',x.casefold()))
    p1,p2,p3=st.columns(3)
    p1.metric('Edificios',len(buildings));p2.metric('Registros',len(records))
    p3.metric('Con cálculo HH',sum(rendimiento_values(r)['hh'] is not None for r in records))
    a,b=st.columns([1.5,1])
    with a:building=st.selectbox('🏢 Ver rendimientos', ['Todos los edificios']+buildings,
        format_func=lambda x:x if x=='Todos los edificios' else f'{x} · {len(by_build[x])}',key='v8_rend_building')
    filtered=records if building=='Todos los edificios' else by_build[building]
    months=sorted({record_date('rendimientos',r)[:7] for r in filtered if len(record_date('rendimientos',r))>=7},reverse=True)
    with b:month=st.selectbox('🗓️ Mes', ['Todos los meses']+months,
        format_func=lambda x:display_month(x) if x!='Todos los meses' else x,key='v8_rend_month')
    filtered=[r for r in filtered if month=='Todos los meses' or record_date('rendimientos',r).startswith(month)]
    search=st.text_input('Buscar rubro, piso o trabajador',placeholder='Ej. enlucidos, piso 2…',key='v8_rend_search')
    if search.strip():
        text=search.strip().casefold()
        filtered=[r for r in filtered if text in ' '.join((record_title('rendimientos',r),record_building('rendimientos',r),str(r.get('trabajador') or ''))).casefold()]
    st.caption(f'**{len(filtered)} rendimientos** · Selecciona un registro para ver fotografías, editarlo o eliminarlo.')
    export_controls('rendimientos',filtered,'rend_general_v8')
    if not filtered:
        st.info('No hay rendimientos con estos filtros.');return
    ordered=sorted(filtered,key=lambda r:(record_date('rendimientos',r),str(r.get('id') or '')),reverse=True)
    visible=page_slice(ordered,'rend_v8_block',12)
    groups=defaultdict(list)
    for item in visible:groups[record_building('rendimientos',item)].append(item)
    for building_name,items in sorted(groups.items(),key=lambda x:x[0].casefold()):
        st.markdown('##### 🏢 '+html_escape(building_name))
        if building!='Todos los edificios':
            export_controls('rendimientos',by_build[building_name],'rend_build_v8_'+hashlib.md5(building_name.encode()).hexdigest()[:9])
        for day,day_records in grouped_records(items,'rendimientos'):
            day_heading(day,len(day_records))
            for record in day_records:
                result=rendimiento_values(record)
                with st.container(border=True):
                    st.markdown('**'+html_escape(record_title('rendimientos',record))+'**')
                    c1,c2,c3=st.columns(3)
                    unit=str(record.get('unidad') or '')
                    c1.metric('Avance',f"{result['avance']:.2f} {unit}")
                    c2.metric('Horas-hombre',f"{result['hh']:.2f} HH" if result['hh'] is not None else 'No verificadas')
                    c3.metric('Productividad',f"{result['productividad']:.3f} {unit}/h" if result['productividad'] is not None else '—')
                    record_actions('rendimientos',record,repo,'rend_view_'+str(record['id']))




def record_actions(table, record, repo, key, editable=True):
    """Detalles bajo demanda; conserva editar/eliminar con protección de versión."""
    with st.expander("Ver detalle, fotografías y acciones", expanded=False):
        if st.checkbox("Mostrar contenido y evidencias", key=key + "_detail"):
            show_record(table, record)
        if editable:
            a, b = st.columns(2)
            if a.button("✏️ Editar", key=key + "_edit"):
                try:
                    repo._access(table, record, write=True)
                    factories = {"checklists": ("Checklist", checklist_data), "rendimientos": ("Rendimientos", rend_data), "incidencias": ("Incidencias", incident_data)}
                    module, factory = factories[table]
                    begin_editor(module, repo.user, factory(record), record)
                    st.session_state["_next_route"] = module
                    st.rerun()
                except Exception as exc:
                    notice_error(exc)
            with b.popover("🗑️ Eliminar"):
                st.caption("La eliminación es permanente.")
                if st.button("Confirmar eliminación", key=key + "_delete"):
                    try:
                        repo.delete(table, record)
                        st.session_state["_flash"] = "Registro eliminado."
                        st.rerun()
                    except Exception as exc:
                        notice_error(exc)


def incidencias_original_matrix(records, repo):
    """Matriz de incidencias del primer diseño, sin HTML sin escapar."""
    st.markdown("### Matriz de Incidencias en tus Proyectos")
    a, b = st.columns([2, 2])
    with a:
        buildings = sorted({str(r.get("proyecto") or "") for r in records if r.get("proyecto")})
        selected_building = st.selectbox("🏢 Seleccionar Proyecto / Edificio:", ["-- Todos Mis Proyectos --"] + buildings, key="inc_original_building")
    with b:
        selected_state = st.segmented_control("Filtrar por Estado:", ["Todos", "Abierta", "Cerrada"], default="Todos", key="inc_original_state")
    visible = [r for r in records if (selected_building == "-- Todos Mis Proyectos --" or r.get("proyecto") == selected_building) and (selected_state in ("Todos", None) or r.get("estado") == selected_state)]
    st.markdown(f"**{len(visible)} incidencia(s) según el filtro**")
    if not visible:
        st.info("No hay incidencias registradas para esta selección.")
        return
    content = []
    for i, item in enumerate(sorted(visible, key=lambda r: (str(r.get("fecha_compromiso") or ""),str(r.get("id")))), 1):
        safe = lambda k: html_escape(str(item.get(k) or ""))
        priority = str(item.get("prioridad") or "Media")
        status = str(item.get("estado") or "Abierta")
        badges = " ".join(("☒" if priority == x else "☐") + " " + x for x in ("Alta", "Media", "Baja"))
        color = "#dc2626" if priority == "Alta" else "#d97706" if priority == "Media" else "#16a34a"
        status_color = "#dc2626" if status == "Abierta" else "#16a34a"
        content.append(
            '<tr><td class="center">' + str(i) + '</td>'
            '<td><b>' + safe("area") + '</b><div style="font-size:.7rem">' + safe("proyecto") + '</div></td>'
            '<td>' + safe("descripcion") + '</td>'
            '<td>' + safe("responsable") + '</td>'
            '<td style="color:' + color + ';font-weight:700">' + html_escape(badges) + '</td>'
            '<td class="center">' + safe("fecha_compromiso") + '</td>'
            '<td style="color:' + status_color + ';font-weight:800">' + html_escape(status) + '</td></tr>'
        )
    st.markdown(
        '<div style="overflow-x:auto"><table class="incidencias-table"><thead><tr>'
        '<th class="center">N°</th><th>Área / Edificio</th><th>Descripción</th><th>Responsable</th>'
        '<th>Prioridad</th><th class="center">Fecha Comp.</th><th>Estado</th>'
        '</tr></thead><tbody>' + "".join(content) + '</tbody></table></div>', unsafe_allow_html=True
    )
    # Exportación con el mismo filtro de la matriz: ambas librerías generan fotos si existen.
    export_controls("incidencias", visible, "inc_matrix")
    with st.expander("⚙️ Cambiar estado de una incidencia", expanded=False):
        editable = [r for r in visible if r.get("usuario_email") == repo.email or repo.admin or r.get("proyecto") in repo.projects]
        if editable:
            record = st.selectbox("Incidencia", editable, format_func=lambda r: str(r.get("area") or "") + " · " + str(r.get("proyecto") or "") + " · " + display_date(r.get("fecha_compromiso")), key="inc_matrix_edit")
            new_status = "Cerrada" if record.get("estado") == "Abierta" else "Abierta"
            if st.button("Cambiar estado a: " + new_status, key="inc_matrix_toggle"):
                try:
                    repo.save("incidencias", {"estado": new_status}, record)
                    st.session_state["_flash"] = "Estado de incidencia actualizado."
                    st.rerun()
                except Exception as exc:
                    notice_error(exc)
        else:
            st.caption("No hay incidencias con permiso de edición.")


def module_original(label, repo):
    """Restaura la distribución de los formularios originales, sin perder repositorio seguro."""
    spec = {
        "Checklist": ("checklists", "Checklist", checklist_data, render_checklist, "### Checklist – Control de Obra", "Supervisión técnica y control de ejecución diaria en obra.", "➕ Crear Nueva Jornada de Inspección"),
        "Libro de Obra": ("inspecciones", "Libro oficial", official_data, render_official, "### Libro de Obra – Formato Oficial", "Estructura técnica de control diario y sincronización con Checklist.", "➕ Llenar Libro de Obra"),
        "Libro de Obra Maestro": ("inspecciones", "Libro maestro", maestro_data, render_maestro, "### Libro de Obra – Maestro Mayor", "Registro diario de actividades, cuadrillas y metrajes ejecutados en obra.", "➕ Llenar Libro de Obra Maestro"),
        "Levantamiento de Incidencias": ("incidencias", "Incidencias", incident_data, render_incident, "### Levantamiento de Incidencias", "Registro y seguimiento técnico de novedades en obra.", "➕ Registrar Incidencia"),
        "Control de Rendimiento": ("rendimientos", "Rendimientos", rend_data, render_rend, "### Control de Rendimiento por Personal", "Registro de trabajadores, edificios, rubros y horas, con cálculo automático y fotografías.", "➕ Registrar Rendimiento"),
    }
    table, module, factory, renderer, header, caption, button = spec[label]
    page_heading("Control de obra",re.sub(r"^#+\s*", "", header),caption)
    editor = st.session_state.get("_editors", {}).get(module)
    if editor:
        selected = editor_picker(module, repo.user, factory)
    else:
        st.caption("Presione el botón para abrir el formulario. Al guardar se cerrará automáticamente, y el borrador se conserva si sale sin terminar.")
        # Conservar botones originales, incluyendo recuperación explícita de borradores.
        if st.button(button, type="primary", key="original_new_" + module):
            begin_editor(module, repo.user, factory())
            rerun_selected_module()
        drafts = cached_rows("alpha_drafts", (("usuario_email", repo.email), ("modulo", module)), columns="slot,revision,updated_at", order="updated_at")
        if drafts:
            with st.expander(f"📁 Recuperar mis borradores ({len(drafts)})", expanded=False):
                choices = {f"Borrador {i+1} · {r['updated_at'][:16]}": r for i,r in enumerate(reversed(drafts))}
                selected_name = st.selectbox("Seleccione un borrador", list(choices), key="restore_choice_" + module)
                if st.button("Continuar borrador", key="restore_button_" + module):
                    draft = choices[selected_name]
                    recovered = all_rows("alpha_drafts", filters={"usuario_email": repo.email, "modulo": module, "slot": draft["slot"]}, columns="slot,datos,revision")
                    if recovered:
                        row = recovered[0]
                        st.session_state.setdefault("_editors", {})[module] = {"slot": row["slot"], "data": row["datos"], "revision": row["revision"]}
                        rerun_selected_module()
                    else:
                        st.warning("El borrador ya no existe. Actualice la página.")
        selected = None
    if selected:
        st.divider()
        st.markdown("**✏️ Edición de registro existente**" if selected.data.get("_original") else "**📝 Nuevo registro**")
        renderer(selected, repo)
        st.divider()
    if table == "incidencias":
        records = repo.shared(table)
    else:
        records = repo.own(table)
    if table == "inspecciones":
        records = [r for r in records if is_master_book(r) == (module == "Libro maestro")]
    if table == "rendimientos":
        history_rend_by_building(records, repo)
    elif table == "incidencias":
        incidencias_original_matrix(records, repo)
        history(table, records, repo, "original_" + module)
    else:
        history(table, records, repo, "original_" + module)


# Los widgets dentro de este fragmento no reconstruyen el dashboard al cambiar.
# Conserva un solo archivo y funciona sin dependencia adicional en versiones modernas.
_FRAGMENT = getattr(st, "fragment", None)


@(_FRAGMENT if _FRAGMENT is not None else (lambda fn: fn))
def render_active_module(name, repo):
    role = repo.user.get("cargo")
    if name == "Libro de Obra Maestro" and role != "Maestro Mayor":
        raise PermissionError("El libro de obra maestro es exclusivo del Maestro Mayor.")
    if role == "Maestro Mayor" and name in {"Checklist", "Libro de Obra", "Levantamiento de Incidencias"}:
        raise PermissionError("El módulo no está habilitado para el cargo Maestro Mayor.")
    if name == "Panel Admin" and not repo.admin:
        raise PermissionError("La administración requiere permisos activos.")
    if name == "Personal a Cargo":
        personal_module(repo)
    elif name == "Colaborativo":
        collaborative_module(repo)
    elif name == "Panel Admin":
        admin_module(repo)
    else:
        module_original(name, repo)


def main():
    for prefix in st.session_state.pop("_cleanup_prefixes", []):
        for key in list(st.session_state):
            if key.startswith(prefix):
                del st.session_state[key]
    st.set_page_config(page_title="Alpha Builders | Portal Ejecutivo", page_icon="🏗️", layout="wide", initial_sidebar_state="expanded")
    st.markdown(APP_CSS, unsafe_allow_html=True)
    st.markdown(APP_CSS_V8, unsafe_allow_html=True)
    # Streamlit puede usar tema oscuro manual aunque Windows esté en modo claro.
    # La preferencia del usuario prevalece sobre prefers-color-scheme.
    try:
        actual_theme = st.context.theme.type
    except (AttributeError, RuntimeError, KeyError):
        actual_theme = None
    if actual_theme in ("dark", "light"):
        tokens = (
            "--bg-app:#0f172a;--text-app:#f8fafc;--card-bg:#1e293b;"
            "--card-border:#334155;--header-bg:#020617;--subtext:#94a3b8;"
            "--hover-bg:#334155;--table-row-even:#1e293b;"
            "--ab-surface:#1e293b;--ab-surface-soft:#162237;--ab-ink:#f2f6ff;"
            "--ab-muted:#bbc7de;--ab-outline:#40516b;--ab-chip:#243b63;"
            "--ab-chip-text:#c6d8ff;--ab-primary:#80a7ff"
        ) if actual_theme == "dark" else (
            "--bg-app:#f8fafc;--text-app:#0f172a;--card-bg:#ffffff;"
            "--card-border:#cbd5e1;--header-bg:#0f172a;--subtext:#64748b;"
            "--hover-bg:#f1f5f9;--table-row-even:#f8fafc;"
            "--ab-surface:#ffffff;--ab-surface-soft:#f5f8ff;--ab-ink:#17233b;"
            "--ab-muted:#586781;--ab-outline:#dbe4f1;--ab-chip:#eaf1ff;"
            "--ab-chip-text:#234a94;--ab-primary:#2459dc"
        )
        st.markdown('<style>:root{' + tokens + '}</style>', unsafe_allow_html=True)
    try:
        if get_config("alpha_schema_version") != "5":
            st.error("El esquema depurado v5 no está activado. No se intentará usar el sistema anterior.")
            st.caption("Revise la migración y SUPABASE_SERVICE_ROLE_KEY / SESSION_SECRET en los secretos de Streamlit.")
            st.stop()
        cookie = stx.CookieManager(key="alpha_cookie_" + str(st.session_state.get("_cookie_epoch", 0)))
        user = auth_user(cookie)
        if user is None:
            login_panel(cookie)
            st.stop()
        if st.session_state.get("_user_identity") != user["correo"]:
            for key in list(st.session_state):
                if key.startswith("edit:") or key.startswith("alpha_tabs_") or key.startswith("alpha_rend_edificios") or key in {"_editors", "_export_artifact", "_temporary_password", "_route"}:
                    del st.session_state[key]
            st.session_state["_user_identity"] = user["correo"]
        st.session_state.pop("_pending_cookie", None)
        repo = Repository(user)
        if user.get("must_change_password"):
            st.warning("Debe cambiar su contraseña temporal antes de acceder al portal.")
            password_change(repo, cookie)
            st.stop()
        sidebar_original(repo, cookie)
        if st.session_state.get("_flash"):
            st.success(st.session_state.pop("_flash"))
        dashboard(repo)
        master = user.get("cargo") == "Maestro Mayor"
        # El libro de obra maestro no aparece ni puede editarse desde cargo diferente.
        visible = (["Libro de Obra Maestro", "Personal a Cargo", "Control de Rendimiento", "Colaborativo"]
                   if master else ["Checklist", "Libro de Obra", "Personal a Cargo", "Levantamiento de Incidencias", "Control de Rendimiento", "Colaborativo"])
        if repo.admin:
            visible.append("Panel Admin")
        labels = ["📋 " + x if x == "Checklist" else "🏗️ " + x if "Libro de Obra" in x else "👷 " + x if x == "Personal a Cargo" else "🚨 " + x if x == "Levantamiento de Incidencias" else "📊 " + x if x == "Control de Rendimiento" else "🤝 " + x if x == "Colaborativo" else "⚙️ " + x for x in visible]
        role_key = "alpha_tabs_maestro" if master else "alpha_tabs_obra"
        next_module = st.session_state.pop("_next_route", None)
        next_map = {"Libro maestro": "Libro de Obra Maestro", "Libro oficial": "Libro de Obra", "Rendimientos": "Control de Rendimiento", "Incidencias": "Levantamiento de Incidencias", "Checklist": "Checklist"}
        next_label = next_map.get(next_module, next_module)
        if next_label in visible:
            st.session_state[role_key] = labels[visible.index(next_label)]
        tabs, lazy = open_lazy_tabs(labels, role_key)
        for name, tab in zip(visible, tabs):
            with tab:
                if lazy and not tab.open:
                    continue
                render_active_module(name, repo)
    except Exception as exc:
        notice_error(exc)
        st.caption("No se interpretan errores de lectura como registros vacíos. Verifique conexión, configuración y permisos.")
        if st.button("Reintentar conexión"):
            invalidate()
            st.rerun()




if __name__=="__main__":
    main()
