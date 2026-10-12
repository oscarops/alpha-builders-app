"""Alpha Builders — aplicación integral en un único archivo.
Requiere ejecutar MIGRACION_ALPHA.sql y configurar secrets.toml antes del despliegue.
Los módulos de interfaz solo se ejecutan desde main(); las funciones puras son comprobables.
"""
from __future__ import annotations
import base64
import ast
import calendar
import urllib.parse
import urllib.request
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
import zlib
import unicodedata
from xml.sax.saxutils import escape as xml_escape
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
from reportlab.lib.pagesizes import A4, letter, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as PDFImage, KeepTogether, PageBreak
from supabase import create_client

LOG = logging.getLogger("alpha")
APP_VERSION = "2026.10.12-v20.3-generales-nomina-clima"
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

APP_CSS = '\n    <style>\n    @import url(\'https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Montserrat:wght@500;600;700;800&display=swap\');\n\n    :root {\n        --bg-app: #f8fafc;\n        --text-app: #0f172a;\n        --card-bg: #ffffff;\n        --card-border: #cbd5e1;\n        --header-bg: #0f172a;\n        --header-text: #ffffff;\n        --sidebar-bg: #0b0f19;\n        --sidebar-card: #111827;\n        --sidebar-border: #1f2937;\n        --subtext: #64748b;\n        --hover-bg: #f1f5f9;\n        --table-row-even: #f8fafc;\n    }\n\n    @media (prefers-color-scheme: dark) {\n        :root {\n            --bg-app: #0f172a;\n            --text-app: #f8fafc;\n            --card-bg: #1e293b;\n            --card-border: #334155;\n            --header-bg: #020617;\n            --header-text: #ffffff;\n            --sidebar-bg: #030712;\n            --sidebar-card: #0f172a;\n            --sidebar-border: #1e293b;\n            --subtext: #94a3b8;\n            --hover-bg: #334155;\n            --table-row-even: #1e293b;\n        }\n    }\n\n    html, body, [class*="css"] { font-family: \'Plus Jakarta Sans\', sans-serif; }\n    h1, h2, h3, .brand-title { font-family: \'Montserrat\', sans-serif !important; letter-spacing: -0.03em !important; }\n\n    .block-container { \n        padding-top: 3.2rem !important; \n        padding-bottom: 1.2rem !important; \n        padding-left: 1rem !important; \n        padding-right: 1rem !important; \n        max-width: 100% !important; \n    }\n    /* Los elementos nativos usan la paleta real del tema activo. */\n    /* Los elementos nativos usan la paleta real del tema activo. */\n    .stCaption, caption, small, [data-testid="stCaptionContainer"] { color: var(--subtext) !important; }\n\n    [data-testid="stInputInstructions"], div[data-testid="stInputInstructions"] { display: none !important; visibility: hidden !important; }\n    [data-testid="stHeader"] { background: transparent !important; z-index: 100 !important; }\n\n    /* BOTONES DE TOGGLE SIDEBAR */\n    [data-testid="stSidebarCollapseButton"] { display: block !important; visibility: visible !important; opacity: 1 !important; z-index: 999999 !important; }\n    [data-testid="collapsedControl"] { display: block !important; visibility: visible !important; opacity: 1 !important; position: fixed !important; top: 12px !important; left: 15px !important; z-index: 999999 !important; }\n\n    [data-testid="stSidebarCollapseButton"] button, [data-testid="collapsedControl"] button {\n        background-color: var(--header-bg) !important; \n        border: 1px solid var(--card-border) !important; \n        border-radius: 50% !important; \n        width: 34px !important; \n        height: 34px !important; \n        color: #ffffff !important; \n        box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important; \n        transition: all 0.2s ease !important;\n    }\n    [data-testid="stSidebarCollapseButton"] button:hover, [data-testid="collapsedControl"] button:hover {\n        background-color: #3b82f6 !important; \n        border-color: #3b82f6 !important; \n        transform: scale(1.08);\n    }\n    [data-testid="stSidebarCollapseButton"] svg, [data-testid="collapsedControl"] svg { fill: #ffffff !important; color: #ffffff !important; }\n\n    /* SIDEBAR */\n    [data-testid="stSidebar"] { \n        background-color: var(--sidebar-bg) !important; \n        border-right: 1px solid var(--sidebar-border) !important; \n        padding-top: 0px !important; \n        padding-left: 10px !important; \n        padding-right: 10px !important; \n        padding-bottom: 12px !important; \n    }\n    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0.4rem !important; padding-top: 0px !important; }\n    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] div { color: #ffffff !important; }\n\n    .sidebar-logo-card { \n        background-color: #ffffff; \n        border-radius: 10px; \n        padding: 6px 8px; \n        margin-top: 0px !important; \n        margin-bottom: 14px !important; \n        box-shadow: 0 4px 12px rgba(0,0,0,0.3); \n        width: 100% !important; \n        box-sizing: border-box; \n        text-align: center; \n        display: block; \n    }\n    [data-testid="stSidebar"] [data-testid="stImage"] { \n        width: 100% !important; \n        display: block !important; \n        margin-top: 4px !important; \n        margin-bottom: 8px !important; \n        clear: both !important; \n    }\n    [data-testid="stImage"] img { \n        border-radius: 10px !important; \n        width: 100% !important; \n        height: auto !important; \n        max-width: 100% !important; \n        object-fit: cover !important; \n        border: 1px solid var(--card-border) !important; \n        margin: 0 !important; \n        display: block !important; \n    }\n\n    .sidebar-profile-box { \n        background: var(--sidebar-card); \n        border: 1px solid var(--sidebar-border); \n        border-radius: 10px; \n        padding: 10px 8px !important; \n        text-align: center; \n        margin-top: 2px; \n        margin-bottom: 6px; \n        width: 100% !important; \n        box-shadow: 0 4px 10px rgba(0,0,0,0.3); \n        box-sizing: border-box; \n    }\n    .sidebar-user-nombres { font-size: 0.84rem; font-weight: 800; color: #ffffff !important; line-height: 1.2; }\n    .sidebar-user-apellidos { font-size: 0.80rem; font-weight: 700; color: #e2e8f0 !important; margin-bottom: 3px !important; line-height: 1.2; }\n    .sidebar-user-email { font-size: 0.65rem; color: #60a5fa !important; font-weight: 600; margin-bottom: 4px !important; word-break: break-all; }\n    .sidebar-user-cargo { \n        display: inline-block; \n        background: #1f2937 !important; \n        color: #ffffff !important; \n        border: 1px solid #374151 !important; \n        font-size: 0.58rem !important; \n        font-weight: 800 !important; \n        padding: 2px 7px !important; \n        border-radius: 12px !important; \n        text-transform: uppercase !important; \n    }\n\n    [data-testid="stSidebar"] hr { margin: 4px 0 !important; border-color: var(--sidebar-border) !important; }\n\n    /* SMART DASHBOARD GLASSMORPHISM */\n    .smart-dashboard-container {\n        background: radial-gradient(120% 120% at 50% 0%, #1e293b 0%, #0f172a 60%, #090d16 100%);\n        border: 1px solid rgba(255, 255, 255, 0.12);\n        border-radius: 14px;\n        padding: 12px 14px;\n        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);\n        color: #ffffff !important;\n        margin-top: 4px;\n        margin-bottom: 12px;\n        width: 100%;\n        box-sizing: border-box;\n    }\n    .smart-dashboard-container * { color: #ffffff !important; }\n\n    .smart-header-bar {\n        display: flex;\n        justify-content: space-between;\n        align-items: center;\n        flex-wrap: wrap;\n        gap: 8px;\n        border-bottom: 1px solid rgba(255, 255, 255, 0.08);\n        padding-bottom: 8px;\n        margin-bottom: 10px;\n    }\n    .smart-title { font-size: 1.1rem; font-weight: 800; letter-spacing: -0.02em; line-height: 1.1; }\n    .smart-user-sub {\n        font-size: 0.78rem;\n        color: #94a3b8 !important;\n        margin-top: 3px;\n        display: flex;\n        align-items: center;\n        flex-wrap: wrap;\n        gap: 6px;\n    }\n\n    .edificio-tag-badge {\n        display: inline-flex;\n        align-items: center;\n        background: rgba(59, 130, 246, 0.20);\n        border: 1px solid rgba(96, 165, 250, 0.45);\n        color: #93c5fd !important;\n        font-size: 0.64rem;\n        font-weight: 800;\n        padding: 2px 7px;\n        border-radius: 6px;\n        letter-spacing: 0.02em;\n    }\n\n    .smart-pill {\n        display: inline-flex;\n        align-items: center;\n        gap: 4px;\n        background: rgba(255, 255, 255, 0.06);\n        border: 1px solid rgba(255, 255, 255, 0.12);\n        backdrop-filter: blur(8px);\n        padding: 3px 8px;\n        border-radius: 16px;\n        font-size: 0.70rem;\n        font-weight: 700;\n    }\n\n    .widgets-grid {\n        display: grid;\n        grid-template-columns: repeat(auto-fit, minmax(135px, 1fr));\n        gap: 8px;\n    }\n\n    .widget-glass-card {\n        background: rgba(255, 255, 255, 0.04);\n        border: 1px solid rgba(255, 255, 255, 0.08);\n        border-radius: 10px;\n        padding: 8px 10px;\n        backdrop-filter: blur(10px);\n        display: flex;\n        flex-direction: column;\n        justify-content: space-between;\n        box-sizing: border-box;\n    }\n\n    .w-card-title {\n        font-size: 0.65rem;\n        font-weight: 800;\n        text-transform: uppercase;\n        color: #94a3b8 !important;\n        letter-spacing: 0.04em;\n        display: flex;\n        justify-content: space-between;\n        align-items: center;\n        margin-bottom: 4px;\n    }\n\n    .milestone-item {\n        display: flex;\n        justify-content: space-between;\n        align-items: center;\n        background: rgba(255, 255, 255, 0.03);\n        border: 1px solid rgba(255, 255, 255, 0.05);\n        border-radius: 6px;\n        padding: 4px 6px;\n        margin-bottom: 3px;\n    }\n    .milestone-name { font-size: 0.68rem; font-weight: 700; color: #e2e8f0 !important; }\n    .milestone-status-done {\n        background: #10b981;\n        color: #ffffff !important;\n        font-size: 0.58rem;\n        font-weight: 800;\n        padding: 1px 6px;\n        border-radius: 10px;\n    }\n    .milestone-status-pending {\n        background: rgba(239, 68, 68, 0.2);\n        border: 1px solid rgba(239, 68, 68, 0.4);\n        color: #fca5a5 !important;\n        font-size: 0.58rem;\n        font-weight: 800;\n        padding: 1px 6px;\n        border-radius: 10px;\n    }\n\n    /* DONA SVG */\n    .donut-container { display: flex; align-items: center; gap: 8px; }\n    .donut-chart-svg { width: 44px; height: 44px; transform: rotate(-90deg); flex-shrink: 0; }\n    .donut-bg { fill: none; stroke: #334155 !important; stroke-width: 4.5; }\n    .donut-progress { fill: none; stroke-width: 4.5; stroke-linecap: round; }\n    .donut-info-val { font-size: 0.98rem; font-weight: 900; line-height: 1; }\n    .donut-info-lbl { font-size: 0.60rem; color: #94a3b8 !important; font-weight: 600; margin-top: 1px; }\n    .stat-hero-number { font-size: 1.35rem; font-weight: 900; line-height: 1; margin-top: 2px; margin-bottom: 2px; }\n\n    /* PESTAÑAS */\n    .stTabs [data-baseweb="tab-list"] { \n        gap: 6px !important; \n        background-color: var(--card-border) !important; \n        padding: 4px !important; \n        border-radius: 10px !important; \n        border: 1px solid var(--card-border) !important; \n    }\n    .stTabs [data-baseweb="tab"] { \n        border-radius: 8px !important; \n        padding: 6px 14px !important; \n        background-color: var(--card-bg) !important; \n        border: 1px solid var(--card-border) !important; \n        box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important; \n        transition: all 0.2s ease !important; \n    }\n    .stTabs [data-baseweb="tab"]:hover { border-color: var(--subtext) !important; background-color: var(--hover-bg) !important; }\n    .stTabs [data-baseweb="tab"] p, .stTabs [data-baseweb="tab"] span { color: var(--text-app) !important; font-weight: 700 !important; font-size: 0.78rem !important; }\n    .stTabs [aria-selected="true"] { \n        background-color: var(--header-bg) !important; \n        border: 1px solid var(--header-bg) !important; \n        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.25) !important; \n    }\n    .stTabs [aria-selected="true"] p, .stTabs [aria-selected="true"] span, .stTabs [aria-selected="true"] div { \n        color: var(--header-text) !important; \n        font-weight: 800 !important; \n    }\n\n    .stButton > button { \n        background-color: var(--header-bg) !important; \n        color: #ffffff !important; \n        border-radius: 980px !important; \n        border: 1px solid var(--card-border) !important; \n        font-weight: 700 !important; \n        padding: 5px 12px !important; \n        font-size: 0.78rem !important; \n    }\n    .stButton > button p, .stButton > button span { color: #ffffff !important; }\n\n    .banner-item-header { \n        background-color: var(--header-bg) !important; \n        border: 1px solid var(--card-border) !important; \n        border-bottom: 2px solid #ffffff !important; \n        border-radius: 6px 6px 0 0 !important; \n        padding: 6px 10px !important; \n        margin-top: 4px !important; \n        margin-bottom: 0px !important; \n    }\n    .banner-item-header span { color: #ffffff !important; font-size: 0.80rem !important; font-weight: 800 !important; }\n\n    .card-item-body-compact { \n        border: 1px solid var(--card-border); \n        border-top: none; \n        border-radius: 0 0 6px 6px; \n        padding: 6px 8px 2px 8px; \n        background-color: var(--card-bg); \n        margin-bottom: 6px; \n    }\n\n    .worker-card-row { \n        background: var(--card-bg); \n        border: 1px solid var(--card-border); \n        border-radius: 8px; \n        padding: 8px 12px; \n        margin-bottom: 6px; \n        display: flex; \n        align-items: center; \n        justify-content: space-between; \n        gap: 8px; \n        box-shadow: 0 1px 3px rgba(0,0,0,0.03); \n    }\n    .worker-info-block { display: flex; flex-direction: column; flex: 1; min-width: 0; }\n    .worker-name-title { font-size: 0.82rem; font-weight: 800; color: var(--text-app); line-height: 1.2; word-break: break-word; }\n    .worker-meta-tags { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 4px; align-items: center; }\n    .tag-cargo-chip { \n        background: var(--card-border); \n        color: var(--text-app); \n        font-size: 0.68rem; \n        font-weight: 800; \n        padding: 2px 8px; \n        border-radius: 12px; \n        border: 1px solid var(--card-border); \n    }\n    .tag-edif-chip { \n        background: rgba(59, 130, 246, 0.15); \n        color: #3b82f6; \n        font-size: 0.68rem; \n        font-weight: 800; \n        padding: 2px 8px; \n        border-radius: 12px; \n        border: 1px solid rgba(59, 130, 246, 0.35); \n    }\n\n    @media (max-width: 768px) { \n        .block-container { padding-top: 3.6rem !important; padding-left: 0.4rem !important; padding-right: 0.4rem !important; }\n        .widgets-grid { grid-template-columns: 1fr 1fr; }\n        .smart-title { font-size: 0.95rem !important; }\n        .worker-card-row { padding: 6px 8px; }\n        .worker-name-title { font-size: 0.78rem; }\n    }\n\n    .incidencias-table { \n        width: 100%; \n        border-collapse: collapse !important; \n        margin-top: 4px; \n        margin-bottom: 8px; \n    }\n    .incidencias-table th { \n        background-color: var(--header-bg) !important; \n        color: #ffffff !important; \n        padding: 6px 8px !important; \n        font-size: 0.75rem !important; \n        font-weight: 700 !important; \n        border: 1px solid var(--card-border) !important; \n        text-align: left; \n    }\n    .incidencias-table th.center, .incidencias-table td.center { text-align: center !important; }\n    .incidencias-table td { \n        padding: 5px 7px !important; \n        font-size: 0.76rem !important; \n        border: 1px solid var(--card-border) !important; \n        vertical-align: middle !important; \n        background-color: var(--card-bg) !important; \n        color: var(--text-app) !important; \n    }\n    .incidencias-table tr:nth-child(even) td { \n        background-color: var(--table-row-even) !important; \n    }\n    #MainMenu {visibility: hidden;} footer {visibility: hidden;}\n    </style>\n    '



# Estilos complementarios 2026.10: diseño adaptativo y contraste para ambos modos.
# No modifica la lógica ni los componentes internos de Streamlit.
APP_CSS_V10 = """
<style>
:root {
  --ab-primary:#2459dc; --ab-primary-deep:#153fa5; --ab-accent:#7c55ef;
  --ab-surface:#ffffff; --ab-surface-soft:#f5f8ff; --ab-surface-soft-2:#eef4ff;
  --ab-ink:#17233b; --ab-muted:#586781; --ab-strong:#0f172a;
  --ab-outline:#dbe4f1; --ab-outline-strong:#c6d3e8; --ab-shadow:0 8px 28px rgba(31,54,101,.07);
  --ab-chip:#eaf1ff; --ab-chip-text:#234a94; --ab-success:#0f9b6f; --ab-danger:#e84d6e;
}
@media (prefers-color-scheme:dark){
 :root { --ab-primary:#8fb3ff; --ab-primary-deep:#6c99ff; --ab-accent:#c0b2ff;
 --ab-surface:#182334; --ab-surface-soft:#142032; --ab-surface-soft-2:#1d2b40;
 --ab-ink:#f5f8ff; --ab-muted:#c0cbe0; --ab-strong:#ffffff;
 --ab-outline:#3f526d; --ab-outline-strong:#546a87; --ab-shadow:0 8px 25px rgba(0,0,0,.22);
 --ab-chip:#243b63; --ab-chip-text:#d7e4ff; --ab-success:#39c48f; --ab-danger:#ff7b96; }
}
/* El fondo principal lo gestiona el tema de Streamlit. */
.stApp [data-testid="stMainBlockContainer"] { max-width:1520px!important; margin:auto; }
.stApp [data-testid="stMainBlockContainer"] p,
.stApp [data-testid="stMainBlockContainer"] label,
.stApp [data-testid="stMainBlockContainer"] span { line-height:1.45; }
/* No forzar el mismo color a todos los controles; tema nativo de Streamlit. */
/* Los títulos fuera de tarjetas usan el color nativo del tema. */

/* Superficies */
.stApp [data-testid="stExpander"] {
 border:1px solid var(--ab-outline)!important; border-radius:15px!important;
 background:linear-gradient(180deg,var(--ab-surface),var(--ab-surface-soft))!important; overflow:hidden!important;
 box-shadow:0 3px 10px rgba(0,0,0,.025);
}
.stApp [data-testid="stExpander"] summary { background:transparent!important; padding:0.72rem .9rem!important; }
.stApp [data-testid="stExpander"] summary:hover { background:var(--ab-surface-soft)!important; }
.stApp [data-testid="stExpander"] summary p,
.stApp [data-testid="stExpander"] summary span { color:var(--ab-strong)!important; font-weight:760!important; }
.stApp [data-testid="stVerticalBlockBorderWrapper"] > div[data-testid="stVerticalBlock"] { border-radius:14px; }
.stApp [data-testid="stMetric"] {
 border:1px solid var(--ab-outline); background:linear-gradient(180deg,var(--ab-surface),var(--ab-surface-soft));
 border-radius:15px; padding:12px 14px; min-height:92px; box-shadow:var(--ab-shadow);
}
.stApp [data-testid="stMetric"] label { color:var(--ab-muted)!important; font-size:.77rem; font-weight:700!important; }
.stApp [data-testid="stMetricValue"] { color:var(--ab-ink)!important; font-size:1.45rem; }
.stApp [data-testid="stDataFrame"] { border-radius:14px; overflow:hidden; }
.stApp [data-testid="stTabs"] [data-baseweb="tab-list"] {
 background:var(--ab-surface-soft)!important; border-color:var(--ab-outline)!important;
 flex-wrap:wrap!important; row-gap:4px!important;
}
.stApp [data-testid="stTabs"] [data-baseweb="tab"] { border-color:var(--ab-outline)!important; }
.stApp [data-testid="stTabs"] [data-baseweb="tab"] p { font-size:.79rem!important; }

/* Botones y formularios */
.stApp .stButton > button,
.stApp [data-testid="stFormSubmitButton"] > button {
 background:linear-gradient(135deg,var(--ab-primary),var(--ab-primary-deep))!important;
 color:#fff!important; border:0!important; border-radius:12px!important; font-weight:800!important;
 min-height:2.65rem!important; box-shadow:0 8px 18px rgba(36,89,220,.18)!important;
}
.stApp .stButton > button:hover,
.stApp [data-testid="stFormSubmitButton"] > button:hover { transform:translateY(-1px); filter:brightness(1.03); }
.stApp .stButton > button p,.stApp .stButton > button span,
.stApp [data-testid="stFormSubmitButton"] > button p,
.stApp [data-testid="stFormSubmitButton"] > button span { color:#fff!important; font-weight:800!important; }
.stApp .stButton > button:disabled,
.stApp [data-testid="stFormSubmitButton"] > button:disabled {
 background:linear-gradient(135deg,#94a3b8,#64748b)!important; color:#f8fafc!important; opacity:.85!important;
}
.stApp [data-testid="stTextInput"] input,
.stApp [data-testid="stNumberInput"] input,
.stApp [data-testid="stTextArea"] textarea,
.stApp [data-testid="stSelectbox"] [data-baseweb="select"],
.stApp [data-testid="stMultiSelect"] [data-baseweb="select"],
.stApp [data-testid="stDateInput"] input,
.stApp [data-testid="stTimeInput"] input {
 border-radius:11px!important; border:1px solid var(--ab-outline-strong)!important;
 background:var(--ab-surface)!important; color:var(--ab-ink)!important;
}
.stApp [data-testid="stCheckbox"] label,
.stApp [data-testid="stRadio"] label,
.stApp [data-testid="stSelectbox"] label,
.stApp [data-testid="stTextInput"] label,
.stApp [data-testid="stNumberInput"] label,
.stApp [data-testid="stTextArea"] label,
.stApp [data-testid="stDateInput"] label,
.stApp [data-testid="stTimeInput"] label,
.stApp [data-testid="stFileUploader"] label,
.stApp [data-testid="stMultiSelect"] label {
 color:var(--ab-strong)!important; font-weight:780!important; font-size:.84rem!important;
}
.stApp [data-testid="stFileUploaderDropzone"] {
 border:1px dashed var(--ab-outline-strong)!important; border-radius:14px!important; background:var(--ab-surface-soft)!important;
}
.stApp [data-testid="stAlertContainer"] [role="alert"] { border-radius:14px!important; }
.stApp [data-testid="stPopover"] button { border-radius:10px!important; }

.ab-heading {
 background:linear-gradient(120deg,var(--ab-surface),var(--ab-surface-soft));
 border:1px solid var(--ab-outline); border-left:4px solid var(--ab-primary);
 border-radius:16px; padding:16px 18px; margin:12px 0 16px; box-shadow:var(--ab-shadow);
}
.ab-heading .ab-eyebrow { color:var(--ab-primary)!important; font-size:.68rem; letter-spacing:.11em; font-weight:800; text-transform:uppercase; }
.ab-heading h3 { font-size:1.38rem!important; color:var(--ab-ink)!important; margin:3px 0 4px!important; }
.ab-heading p { font-size:.82rem!important; color:var(--ab-muted)!important; margin:0!important; }
.ab-day {display:flex;align-items:center;gap:9px;margin:15px 0 7px; border-bottom:1px solid var(--ab-outline);padding-bottom:7px;}
.ab-day-title {font-size:.88rem;font-weight:800;color:var(--ab-ink)!important;}
.ab-count {display:inline-flex;align-items:center;border-radius:100px; background:var(--ab-chip);color:var(--ab-chip-text)!important;font-weight:750;font-size:.67rem;padding:3px 9px;}
.ab-list-wrap {border:1px solid var(--ab-outline);border-radius:15px; overflow:hidden;background:var(--ab-surface);box-shadow:var(--ab-shadow);}
.ab-item {display:grid;grid-template-columns: minmax(0,1fr) auto;column-gap:16px;align-items:center;padding:13px 16px;border-bottom:1px solid var(--ab-outline);}
.ab-item:last-child{border-bottom:0;}
.ab-item:hover{background:var(--ab-surface-soft);}
.ab-item-title {color:var(--ab-ink)!important;font-size:.85rem;font-weight:800;overflow-wrap:anywhere;}
.ab-item-sub {color:var(--ab-muted)!important;font-size:.73rem;margin-top:4px;}
.ab-item-tags {display:flex;flex-wrap:wrap;gap:5px;margin-top:6px;}
.ab-chip {display:inline-block;border:1px solid var(--ab-outline);border-radius:999px;padding:4px 9px;background:var(--ab-chip);color:var(--ab-chip-text)!important;font-size:.69rem;font-weight:800;}
.ab-chip-legacy {background:var(--ab-surface-soft-2);color:var(--ab-muted)!important;}
.ab-subtle {color:var(--ab-muted)!important;font-size:.77rem;font-weight:700;}
.ab-inline-filter {display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:5px 0 10px;}
.ab-separator {height:1px;background:var(--ab-outline);margin:14px 0;}
.ab-section-card{border:1px solid var(--ab-outline);border-radius:16px;background:linear-gradient(180deg,var(--ab-surface),var(--ab-surface-soft));padding:14px 16px;box-shadow:var(--ab-shadow);}
.ab-note{color:var(--ab-muted)!important;font-size:.76rem;}

@media(max-width:768px){
 .ab-heading{padding:12px;margin:10px 0;} .ab-heading h3{font-size:1.12rem!important;}
 .ab-item{padding:10px 12px;grid-template-columns:1fr;gap:4px;}
 .stApp [data-testid="stMetric"]{min-height:74px;padding:8px 10px;}
 .stApp .stButton > button,.stApp [data-testid="stFormSubmitButton"] > button{min-height:2.55rem!important; font-size:.84rem!important;}
}

/* ALPHA BUILDERS V10 | Accesibilidad, responsive y tema dual */
:root {
  --ab-highlight:#3a6ce6; --ab-glow:rgba(59,104,216,.13);
  --ab-strong:#0b1830; --ab-outline-strong:#bbcbe3;
  --ab-surface-soft-2:#edf3fd;
  --ab-success:#087d61; --ab-danger:#b42341;
}
@media(prefers-color-scheme:dark){
 :root { --ab-highlight:#9dbaff; --ab-glow:rgba(88,144,255,.17);
 --ab-strong:#f8fbff; --ab-outline-strong:#5b6f8c;
 --ab-surface-soft-2:#1b2d44; --ab-success:#5be0ab; --ab-danger:#ff99ae; }
}
html { scroll-behavior:smooth; }
.stApp [data-testid="stMainBlockContainer"] { max-width:1480px!important; }
.stApp [data-testid="stMainBlockContainer"] > div { min-width:0; }
/* Main area inputs retain contrast even when the dark theme is manually selected. */
.stApp [data-testid="stMainBlockContainer"] input,
.stApp [data-testid="stMainBlockContainer"] textarea,
.stApp [data-testid="stMainBlockContainer"] [data-baseweb="select"] input {
 color:var(--ab-ink)!important; caret-color:var(--ab-primary)!important;
}
.stApp [data-testid="stMainBlockContainer"] input::placeholder,
.stApp [data-testid="stMainBlockContainer"] textarea::placeholder { color:var(--ab-muted)!important; opacity:.88; }
.stApp [data-testid="stMainBlockContainer"] [data-testid="stMetric"] p,
.stApp [data-testid="stMainBlockContainer"] [data-testid="stMetric"] span { color:var(--ab-ink)!important; }
.stApp [data-testid="stMainBlockContainer"] [data-testid="stExpander"] summary,
.stApp [data-testid="stMainBlockContainer"] [data-testid="stExpander"] summary p,
.stApp [data-testid="stMainBlockContainer"] [data-testid="stExpander"] summary span {color:var(--ab-ink)!important;}
/* Secondary actions should never look like primary destructive actions. */
.stApp .stButton button[kind="secondary"],
.stApp [data-testid="stPopover"] button,
.stApp [data-testid="stDownloadButton"] button {
 background:var(--ab-surface)!important; color:var(--ab-ink)!important;
 border:1px solid var(--ab-outline-strong)!important;
 box-shadow:none!important; border-radius:11px!important;
}
.stApp .stButton button[kind="secondary"] p,
.stApp .stButton button[kind="secondary"] span,
.stApp [data-testid="stPopover"] button p,
.stApp [data-testid="stDownloadButton"] button p { color:var(--ab-ink)!important; }
.stApp .stButton button[kind="secondary"]:hover,
.stApp [data-testid="stDownloadButton"] button:hover { background:var(--ab-surface-soft-2)!important; border-color:var(--ab-primary)!important; }
.stApp .stButton button[kind="primary"],
.stApp [data-testid="stFormSubmitButton"] button[kind="primary"] {
 background:linear-gradient(105deg,#2553bd,#426de6)!important;
 color:#fff!important; border:1px solid transparent!important;
 box-shadow:0 8px 22px rgba(32,79,184,.15)!important;
}
.stApp .stButton button[kind="primary"] p,
.stApp .stButton button[kind="primary"] span,
.stApp [data-testid="stFormSubmitButton"] button[kind="primary"] p {color:#fff!important;}
.stApp [data-testid="stSegmentedControl"] {max-width:100%;overflow-x:auto;}
.stApp [data-testid="stSegmentedControl"] label {white-space:normal!important;}
.stApp [data-testid="stFileUploaderDropzone"] {min-height:100px;}
.stApp [data-testid="stSidebar"] {background:linear-gradient(180deg,#0a1328 0%,#101e3b 100%)!important;}
.stApp [data-testid="stSidebar"] [data-testid="stImage"] img { object-fit:contain!important;max-height:155px; }
.stApp [data-testid="stSidebar"] .sidebar-profile-box {border:1px solid #30415f!important;border-radius:14px;}
.stApp [data-testid="stSidebar"] .sidebar-user-email {color:#aecaff!important;}
.stApp [data-testid="stSidebar"] .stButton button {min-height:2.55rem!important;}
.stApp [data-testid="stSidebar"] .stButton button p {font-weight:780!important;}
.stApp [data-testid="stExpander"] {margin-top:3px;}
/* Dashboard premium: dark cards stay legible in either app theme. */
.ab-hero {
 position:relative;overflow:hidden;margin:7px 0 18px;padding:21px 23px;
 background:linear-gradient(120deg,#0b1530,#1a2e54 65%,#263a60);
 border:1px solid rgba(181,204,255,.18);border-radius:20px;
 color:#fff!important;box-shadow:0 12px 32px rgba(10,25,57,.19);
}
.ab-hero::after{content:"";position:absolute;right:-70px;top:-115px;width:330px;height:330px;
 background:radial-gradient(circle,rgba(119,162,255,.19),transparent 70%);pointer-events:none;}
.ab-hero-top{position:relative;z-index:1;display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:14px;}
.ab-hero-title{font-family:'Montserrat',sans-serif;font-size:1.25rem;font-weight:800;letter-spacing:-.03em;line-height:1.2;color:#fff!important;}
.ab-hero-kicker{font-size:.65rem;color:#9fc5ff!important;font-weight:800;letter-spacing:.17em;text-transform:uppercase;margin-bottom:4px;}
.ab-hero-desc{font-size:.80rem;color:#e6edff!important;margin-top:7px;line-height:1.45;}
.ab-hero-date{border-radius:100px;border:1px solid rgba(255,255,255,.17);background:rgba(255,255,255,.08);
 padding:7px 12px;color:#f8fbff!important;font-size:.75rem;font-weight:750;}
.ab-hero-tags{display:flex;flex-wrap:wrap;gap:5px;margin-top:12px;}
.ab-hero-tags span{display:inline-block;border:1px solid rgba(176,206,255,.28);background:rgba(89,141,235,.16);
 border-radius:100px;font-size:.67rem;font-weight:800;padding:4px 9px;color:#e4efff!important;}
.ab-overview-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:11px;margin:14px 0 2px;}
.ab-overview-card{background:rgba(255,255,255,.075);border:1px solid rgba(221,234,255,.14);border-radius:15px;
 padding:12px 13px;min-width:0;backdrop-filter:blur(6px);}
.ab-overview-label{display:flex;justify-content:space-between;gap:7px;align-items:center;
 color:#c4d6f7!important;font-size:.69rem;font-weight:800;margin-bottom:9px;}
.ab-overview-value{font-size:1.47rem;line-height:1.1;font-weight:850;color:#fff!important;letter-spacing:-.035em;}
.ab-overview-sub{font-size:.66rem;color:#c8d5ea!important;margin-top:7px;line-height:1.4;}
.ab-status-line{display:flex;align-items:center;justify-content:space-between;gap:5px;
 padding:6px 0;border-top:1px solid rgba(255,255,255,.11);font-size:.68rem;color:#e8f0ff!important;}
.ab-status-line:first-of-type{border-top:0;}
.ab-status-pill{padding:3px 7px;border-radius:100px;font-size:.63rem;font-weight:850;white-space:nowrap;}
.ab-status-ok{color:#9ff1cb!important;background:rgba(24,179,113,.18);}
.ab-status-wait{color:#ffe5a8!important;background:rgba(245,166,35,.14);}
.ab-mini-bar{height:7px;border-radius:100px;background:rgba(255,255,255,.13);overflow:hidden;margin-top:9px;}
.ab-mini-bar>span{height:100%;display:block;border-radius:100px;background:linear-gradient(90deg,#4edeb7,#87bbff);}
.stApp [data-testid="stMainBlockContainer"] .ab-hero :is(p,span,div,b,strong){color:inherit;}
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-hero-title,
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-overview-value{color:#fff!important;}
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-hero-kicker{color:#9fc5ff!important;}
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-hero-desc,
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-status-line{color:#e6edff!important;}
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-overview-sub,
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-overview-label{color:#cbdaf7!important;}
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-hero-tags span{color:#e4efff!important;}
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-status-pill.ab-status-ok{color:#a6ffd7!important;}
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-status-pill.ab-status-wait{color:#ffe5a8!important;}
/* Improved history: unified readable expandable cards. */
.ab-record-eyebrow {font-size:.69rem;color:var(--ab-muted)!important;font-weight:760;margin:9px 0;}
.ab-record-summary {background:var(--ab-surface-soft)!important;border:1px solid var(--ab-outline)!important;border-radius:12px;
 padding:9px 13px; margin:5px 0 10px;}
.ab-record-summary-title{color:var(--ab-strong)!important;font-size:.88rem;font-weight:850;}
.ab-record-summary-sub{color:var(--ab-muted)!important;font-size:.72rem;margin-top:3px;}
.ab-field-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:8px;}
.ab-field{border:1px solid var(--ab-outline);border-radius:12px;background:var(--ab-surface);padding:9px 12px;}
.ab-field-label{font-size:.64rem;font-weight:800;color:var(--ab-muted)!important;letter-spacing:.04em;text-transform:uppercase;}
.ab-field-value{margin-top:3px;font-size:.89rem;font-weight:800;color:var(--ab-strong)!important;overflow-wrap:anywhere;}
.ab-layout-label {font-weight:800;font-size:.80rem;color:var(--ab-strong)!important;margin:15px 0 8px;}
.ab-tag-inline{border:1px solid var(--ab-outline);border-radius:100px;color:var(--ab-chip-text)!important;
 background:var(--ab-chip);font-weight:800;padding:3px 9px;font-size:.71rem;}
.ab-empty-hint{border:1px dashed var(--ab-outline-strong);background:var(--ab-surface-soft);border-radius:13px;
 color:var(--ab-muted)!important;padding:15px;font-size:.8rem;}
/* Contrast and scrollbars */
.stApp [data-testid="stDataFrame"]{border:1px solid var(--ab-outline)!important;}
.stApp [data-testid="stToast"]{border-radius:12px!important;}
@media(max-width:1100px){
 .ab-overview-grid{grid-template-columns:repeat(2,minmax(0,1fr));}
}
@media(max-width:640px){
 .ab-hero{padding:15px 13px;border-radius:15px;margin-top:3px;}
 .ab-hero-title{font-size:1.08rem;}
 .ab-overview-grid{gap:7px;}
 .ab-overview-card{padding:10px;}
 .ab-overview-value{font-size:1.24rem;}
 .ab-hero-date{font-size:.67rem;}
 .ab-field-row{grid-template-columns:repeat(2,minmax(0,1fr));}
 .ab-field{padding:9px 8px;}
}
@media(max-width:360px){
 .ab-overview-grid,.ab-field-row{grid-template-columns:1fr;}
}
@media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;transition-duration:0.01ms!important;animation-duration:0.01ms!important;}}


/* Single-click record headers: full width clickable, no additional checkbox. */
.stApp [class*="st-key-v10_open_record_"] button,
.stApp [class*="st-key-v10_open_yield_"] button {
 display:flex!important;justify-content:flex-start!important;text-align:left!important;
 width:100%!important;height:auto!important;min-height:3.15rem!important;
 white-space:normal!important;overflow-wrap:anywhere!important;
 padding:11px 15px!important;margin:4px 0 5px!important;
 background:var(--ab-surface)!important; color:var(--ab-strong)!important;
 border:1px solid var(--ab-outline)!important; border-left:4px solid var(--ab-primary)!important;
 border-radius:13px!important;box-shadow:var(--ab-shadow)!important;
}
.stApp [class*="st-key-v10_open_record_"] button:hover,
.stApp [class*="st-key-v10_open_yield_"] button:hover {
 background:var(--ab-surface-soft-2)!important;border-color:var(--ab-primary)!important;
}
.stApp [class*="st-key-v10_open_record_"] button p,
.stApp [class*="st-key-v10_open_record_"] button span,
.stApp [class*="st-key-v10_open_yield_"] button p,
.stApp [class*="st-key-v10_open_yield_"] button span {color:var(--ab-strong)!important;font-size:.83rem!important;font-weight:790!important;}


/* V11: Especificidad alta: no heredar color oscuro dentro del hero */
.stApp [data-testid="stMainBlockContainer"] section.ab-hero,
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-hero-title,
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-hero-title span,
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-overview-value,
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-overview-card .ab-status-line span:first-child {
 color:#ffffff!important;
}
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-overview-label,
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-overview-label span,
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-overview-sub,
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-hero-desc {
 color:#e0ebff!important;
}
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-hero-kicker {color:#9fc5ff!important;}
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-status-pill.ab-status-ok {color:#a6ffd7!important;}
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-status-pill.ab-status-wait {color:#ffe5a8!important;}
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-hero-date,
.stApp [data-testid="stMainBlockContainer"] section.ab-hero .ab-hero-tags span {color:#f2f6ff!important;}
/* Menos margen superior y hero realmente compacto */
.stApp .block-container {padding-top:1.35rem!important;padding-bottom:1rem!important;}
.stApp [data-testid="stMainBlockContainer"] {padding-top:1.1rem!important;}
.ab-hero {padding:15px 19px!important; margin:0 0 11px!important;border-radius:18px!important;}
.ab-hero-title {font-size:1.15rem!important;}
.ab-hero-tags {margin-top:7px!important;}
.ab-overview-grid {gap:9px!important;margin:10px 0 0!important;}
.ab-overview-card {padding:10px 12px!important;border-radius:13px!important;}
.ab-overview-label {margin-bottom:5px!important;}
.ab-overview-value {font-size:1.36rem!important;}
.ab-hero-top {gap:7px!important;}
/* Sidebar compacta, logo reducido, fotografía destacada */
.stApp [data-testid="stSidebar"] .sidebar-logo-card {max-width:188px!important; margin:3px auto 8px!important;padding:5px!important;}
.stApp [data-testid="stSidebar"] [data-testid="stImage"] img {
 width:100%!important;height:196px!important;max-height:196px!important;
 object-fit:cover!important;object-position:center top!important;
 border:3px solid #94b7ff!important;border-radius:18px!important;
}
.stApp [data-testid="stSidebar"] .sidebar-profile-box {padding:9px 8px!important;margin:7px 0 8px!important;}
.stApp [data-testid="stSidebar"] .stButton button[kind="secondary"] {
 background:#172949!important;color:#fff!important;border:1px solid #486188!important;
 box-shadow:none!important;
}
.stApp [data-testid="stSidebar"] .stButton button[kind="secondary"] p {color:white!important;}
/* Navegación de app por módulos, sin radio buttons ni contenido duplicado */
.stApp [data-testid="stSegmentedControl"] {margin:0 0 4px!important;}
.stApp [data-testid="stSegmentedControl"] > div {max-width:100%!important;}
.stApp [data-testid="stSegmentedControl"] label {
 border-radius:11px!important;border-color:var(--ab-outline)!important;
 font-size:.81rem!important;font-weight:760!important;cursor:pointer!important;
}
/* Color de selección definido de forma conjunta con el texto en APP_CSS_V19. */
.ab-module-tabs{margin:6px 0 11px;padding:7px 11px;border:1px solid var(--ab-outline);
 border-radius:15px;background:var(--ab-surface);box-shadow:var(--ab-shadow);}
.ab-module-tabs-title{font-size:.67rem;color:var(--ab-muted)!important;text-transform:uppercase;font-weight:850;letter-spacing:.1em;}
.ab-welcome{border:1px solid var(--ab-outline);background:linear-gradient(110deg,var(--ab-surface),var(--ab-surface-soft));
 border-radius:16px;padding:15px 18px;margin:9px 0 15px;box-shadow:var(--ab-shadow);}
.ab-welcome h3{margin:0 0 4px!important;color:var(--ab-strong)!important;font-size:1.18rem!important;}
.ab-welcome p{color:var(--ab-muted)!important;margin:0;font-size:.8rem;}
.ab-section-mini{margin:7px 0 12px;font-size:.75rem;font-weight:800;color:var(--ab-muted)!important;}
.ab-record-table-wrap{overflow-x:auto;max-width:100%;border:1px solid var(--ab-outline);border-radius:14px;
 background:var(--ab-surface);box-shadow:var(--ab-shadow);margin:8px 0 13px;}
.ab-record-table{border-collapse:collapse;width:100%;font-size:.82rem;color:var(--ab-ink);table-layout:auto;}
.ab-record-table th{background:var(--ab-surface-soft-2);color:var(--ab-strong)!important;
 text-align:left;font-size:.71rem;letter-spacing:.025em;font-weight:850;padding:11px 12px;
 border-bottom:2px solid var(--ab-outline);white-space:nowrap;}
.ab-record-table td{padding:10px 12px;border-top:1px solid var(--ab-outline);
 vertical-align:top;color:var(--ab-ink)!important;overflow-wrap:anywhere;min-width:80px;max-width:480px;white-space:pre-wrap;}
.ab-record-table tr:nth-child(even) td{background:var(--ab-surface-soft)!important;}
.ab-record-table tbody tr:hover td{background:var(--ab-surface-soft-2)!important;}
.ab-record-table td.ab-table-primary{font-weight:770;color:var(--ab-strong)!important;}
.ab-detail-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(165px,1fr));gap:9px;margin:8px 0 13px;}
.ab-detail-cell{background:var(--ab-surface);border:1px solid var(--ab-outline);border-radius:12px;padding:10px 12px;min-width:0;}
.ab-detail-label{font-size:.66rem!important;color:var(--ab-muted)!important;font-weight:800;text-transform:uppercase;letter-spacing:.04em;}
.ab-detail-value{font-size:.87rem!important;color:var(--ab-strong)!important;font-weight:750;margin-top:5px;overflow-wrap:anywhere;}
.ab-screen-tile{margin:3px 0 9px!important;}
.ab-screen-tile .stButton button{min-height:3.5rem!important;}
.stApp [data-testid="stDataFrame"] {box-shadow:var(--ab-shadow)!important;}
@media(max-width:650px){
 .stApp .block-container{padding-top:.9rem!important;padding-left:.58rem!important;padding-right:.58rem!important;}
 .ab-hero{padding:12px!important;border-radius:14px!important;}
 .ab-hero-title{font-size:1.06rem!important;}
 .ab-overview-grid{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:7px!important;}
 .ab-overview-card{padding:9px!important;}
 .ab-detail-grid{grid-template-columns:repeat(2,minmax(0,1fr));gap:7px;}
 .ab-record-table th,.ab-record-table td{padding:8px 9px;}
}


/* V12 — Escritorio modular. Respeta el tema claro y oscuro actual. */
.ab-home-intro {margin:1px 0 6px;}
.ab-home-kicker {font-size:.7rem;letter-spacing:.12em;font-weight:900;color:var(--ab-primary)!important;}
.ab-home-intro h2 {font-size:1.52rem!important;letter-spacing:-.04em;margin:.17rem 0 .22rem!important;color:var(--ab-ink)!important;}
.ab-home-intro p {font-size:.86rem!important;line-height:1.55;color:var(--ab-muted)!important;margin:0;}
.ab-home-widget-head {display:flex;align-items:center;gap:10px;border-bottom:1px solid var(--ab-outline);padding:2px 0 13px;margin-bottom:14px;}
.ab-home-widget-icon {height:42px;width:42px;border-radius:13px;background:var(--ab-chip);display:grid;place-items:center;font-size:1.28rem;flex:0 0 auto;}
.ab-home-widget-head h3 {margin:0!important;color:var(--ab-ink)!important;font-size:1.06rem!important;letter-spacing:-.01em;}
.ab-home-widget-head p {margin:2px 0 0!important;font-size:.72rem!important;color:var(--ab-muted)!important;}
.ab-home-cal {width:100%;border-collapse:separate;border-spacing:4px;text-align:center;margin:3px 0 12px;table-layout:fixed;}
.ab-home-cal th {font-size:.73rem;padding:5px 1px;color:var(--ab-muted)!important;font-weight:800;}
.ab-home-cal td {position:relative;height:39px;border:1px solid var(--ab-outline);border-radius:9px;color:var(--ab-ink)!important;background:var(--ab-surface-soft);font-weight:720;font-size:.79rem;}
.ab-home-cal .ab-cal-empty {background:transparent;border:0;}
.ab-home-cal .ab-cal-today {background:var(--ab-chip);border:2px solid var(--ab-primary);color:var(--ab-chip-text)!important;}
.ab-home-cal small {position:absolute;right:1px;top:-3px;font-weight:850;font-size:.59rem;color:white!important;background:#2961d6;border-radius:100px;min-width:15px;height:15px;line-height:15px;}
.ab-home-answer {background:var(--ab-chip);color:var(--ab-chip-text)!important;border:1px solid var(--ab-outline);padding:16px;border-radius:13px;text-align:center;font-size:1.45rem;font-weight:850;letter-spacing:-.03em;overflow-wrap:anywhere;margin-top:9px;}
.ab-home-weather {background:linear-gradient(110deg,#152648,#334f80);color:white!important;border-radius:16px;padding:18px 20px;margin-bottom:12px;}
.stApp [data-testid="stMainBlockContainer"] .ab-home-weather,
.stApp [data-testid="stMainBlockContainer"] .ab-home-weather div,
.stApp [data-testid="stMainBlockContainer"] .ab-home-weather span{color:white!important;}
.ab-weather-big{font-size:2.25rem;font-weight:850;letter-spacing:-.055em;line-height:1.1;margin-bottom:4px;}
@media (max-width:720px){
 .ab-home-intro h2{font-size:1.27rem!important;}
 .ab-home-cal {border-spacing:2px;}
 .ab-home-cal td {height:34px;font-size:.69rem;}
 .ab-home-widget-head{gap:8px;}
}
</style>
"""



APP_CSS_V14 = """
<style>
/* Checklist y formularios: máxima legibilidad en ambos temas. */
.stApp [data-testid="stMainBlockContainer"] .banner-item-header,
.stApp [data-testid="stMainBlockContainer"] .banner-item-header * {
 color:#ffffff!important; opacity:1!important; text-shadow:none!important;
}
.stApp [data-testid="stMainBlockContainer"] .banner-item-header {
 background:linear-gradient(110deg,#132344,#284c87)!important;
 border:1px solid #345786!important; border-radius:12px 12px 0 0!important;
 padding:11px 16px!important; margin-top:14px!important;
}
.stApp [data-testid="stMainBlockContainer"] .banner-item-header span {
 font-size:.85rem!important; letter-spacing:.01em; font-weight:800!important;
}
.stApp [data-testid="stMainBlockContainer"] div[data-testid="stRadio"] label,
.stApp [data-testid="stMainBlockContainer"] div[data-testid="stCheckbox"] label,
.stApp [data-testid="stMainBlockContainer"] div[data-testid="stTextArea"] label,
.stApp [data-testid="stMainBlockContainer"] div[data-testid="stTextInput"] label {
 color:var(--ab-strong,var(--text-app))!important; opacity:1!important;
}
/* Tarjetas de colaboración y navegación con contraste real. */
.ab-link-card{border:1px solid var(--ab-outline);border-radius:14px;
 background:linear-gradient(130deg,var(--ab-surface),var(--ab-surface-soft));
 padding:15px 18px;margin:12px 0;box-shadow:var(--ab-shadow);}
.ab-link-title{font-size:1rem;font-weight:850;color:var(--ab-ink)!important;}
.ab-link-meta{font-size:.79rem;font-weight:600;color:var(--ab-muted)!important;}
.ab-request-priority{display:inline-flex;border-radius:20px;padding:4px 10px;
 font-size:.7rem;font-weight:850;background:var(--ab-chip);color:var(--ab-chip-text)!important;}
.stApp [data-testid="stMainBlockContainer"] .ab-link-card p,
.stApp [data-testid="stMainBlockContainer"] .ab-link-card span,
.stApp [data-testid="stMainBlockContainer"] .ab-link-card strong{
 color:var(--ab-ink)!important;}
@media(max-width:768px){.ab-link-card{padding:11px 13px;}}
</style>
"""


APP_CSS_V17 = """
<style>
/* Tarjetas y navegación accesibles en claro y oscuro */
.stApp [data-testid="stMainBlockContainer"] [data-testid="stSegmentedControl"] {
  border-radius:12px!important; overflow-x:auto; max-width:100%;
}
/* Etiquetas segmentadas: contraste definido por fondo y texto juntos en V19. */
.stApp [data-testid="stMainBlockContainer"] [data-testid="stButton"] button {
  white-space:normal!important; overflow-wrap:anywhere!important; min-height:2.9rem!important;
  text-align:left!important; transition:background .15s ease, border-color .15s ease;
}
.stApp [data-testid="stMainBlockContainer"] button[kind="secondary"]:hover {
  background:var(--ab-surface-soft-2)!important; border-color:var(--ab-primary)!important;
}
.ab-v17-page{text-align:center;padding:.7rem;color:var(--ab-muted)!important;font-size:.82rem;font-weight:800}
.ab-v17-help {background:var(--ab-surface-soft)!important;border:1px solid var(--ab-outline);
 border-radius:12px;padding:11px 15px;margin:5px 0 12px;color:var(--ab-ink)!important}
@media (max-width:760px){.stApp [data-testid="stMainBlockContainer"] [data-testid="stButton"] button {font-size:.77rem!important}}
</style>
"""


APP_CSS_V18 = """
<style>
/* Contraste reforzado para cualquier bloque azul u oscuro */
.stApp [data-testid="stMainBlockContainer"] .banner-item-header,
.stApp [data-testid="stMainBlockContainer"] .banner-item-header *,
.stApp [data-testid="stMainBlockContainer"] .smart-dashboard-container,
.stApp [data-testid="stMainBlockContainer"] .smart-dashboard-container *,
.stApp [data-testid="stMainBlockContainer"] .ab-hero,
.stApp [data-testid="stMainBlockContainer"] .ab-hero *,
.stApp [data-testid="stMainBlockContainer"] .ab-home-weather,
.stApp [data-testid="stMainBlockContainer"] .ab-home-weather *,
.stApp [data-testid="stMainBlockContainer"] .widget-glass-card,
.stApp [data-testid="stMainBlockContainer"] .widget-glass-card *,
.stApp [data-testid="stMainBlockContainer"] .milestone-item,
.stApp [data-testid="stMainBlockContainer"] .milestone-item *,
.stApp [data-testid="stSidebar"] .sidebar-profile-box,
.stApp [data-testid="stSidebar"] .sidebar-profile-box *,
.stApp [data-testid="stSidebar"] .smart-dashboard-container,
.stApp [data-testid="stSidebar"] .smart-dashboard-container * {
  color:#ffffff !important;
  -webkit-text-fill-color:#ffffff !important;
  opacity:1 !important;
  text-shadow:none !important;
}

/* Subtextos claros sobre superficies azules */
.stApp [data-testid="stMainBlockContainer"] .smart-user-sub,
.stApp [data-testid="stMainBlockContainer"] .w-card-title,
.stApp [data-testid="stMainBlockContainer"] .ab-hero-desc,
.stApp [data-testid="stMainBlockContainer"] .ab-overview-label,
.stApp [data-testid="stMainBlockContainer"] .ab-overview-sub,
.stApp [data-testid="stMainBlockContainer"] .ab-hero-date,
.stApp [data-testid="stMainBlockContainer"] .ab-hero-tags span,
.stApp [data-testid="stMainBlockContainer"] .milestone-name,
.stApp [data-testid="stMainBlockContainer"] .edificio-tag-badge,
.stApp [data-testid="stMainBlockContainer"] .smart-pill {
  color:#e8f1ff !important;
  -webkit-text-fill-color:#e8f1ff !important;
}

/* El título de actividad sobre azul debe verse siempre */
.stApp [data-testid="stMainBlockContainer"] .banner-item-header span,
.stApp [data-testid="stMainBlockContainer"] .banner-item-header strong,
.stApp [data-testid="stMainBlockContainer"] .banner-item-header p {
  color:#ffffff !important;
  -webkit-text-fill-color:#ffffff !important;
  font-weight:800 !important;
}

/* Botones oscuros/azules: texto siempre visible */
.stApp [data-testid="stMainBlockContainer"] .stButton button,
.stApp [data-testid="stMainBlockContainer"] [data-testid="stFormSubmitButton"] button {
  color:inherit;
}
.stApp [data-testid="stMainBlockContainer"] .stButton button p,
.stApp [data-testid="stMainBlockContainer"] .stButton button span,
.stApp [data-testid="stMainBlockContainer"] [data-testid="stFormSubmitButton"] button p,
.stApp [data-testid="stMainBlockContainer"] [data-testid="stFormSubmitButton"] button span {
  opacity:1 !important;
}

/* Expanders con header oscuro */
.stApp [data-testid="stExpander"] summary,
.stApp [data-testid="stExpander"] summary * {
  opacity:1 !important;
}
</style>
"""


APP_CSS_V19 = """
<style>
/* Alpha Builders v18.7 | Contraste independiente del modo del sistema.
   Evitar reglas globales sobre .stApp span/label/p y no inferir el tema
   exclusivamente de prefers-color-scheme, ya que Streamlit permite un
   tema manual distinto al del dispositivo. */

/* Menús y navegación: siempre fondo y texto del MISMO diseño, en ambos temas. */
.stApp [data-testid="stSegmentedControl"] label,
.stApp [data-testid="stMainBlockContainer"] [data-testid="stSegmentedControl"] label {
  background:#162849!important;
  border:1px solid #3b557b!important;
  color:#f8fbff!important;
  -webkit-text-fill-color:#f8fbff!important;
  font-weight:740!important;
}
.stApp [data-testid="stSegmentedControl"] label :is(p,span,div,strong,small),
.stApp [data-testid="stMainBlockContainer"] [data-testid="stSegmentedControl"] label :is(p,span,div,strong,small) {
  color:#f8fbff!important;
  -webkit-text-fill-color:#f8fbff!important;
  opacity:1!important;
}
.stApp [data-testid="stSegmentedControl"] label:is(:hover,:focus-within) {
  background:#254575!important;
  border-color:#7998d7!important;
}
.stApp [data-testid="stSegmentedControl"] label:has(input:checked),
.stApp [data-testid="stSegmentedControl"] label[aria-checked="true"] {
  background:#2555ca!important;
  border-color:#7ba6ff!important;
  color:#fff!important;
  -webkit-text-fill-color:#fff!important;
}
.stApp [data-testid="stSegmentedControl"] label:has(input:checked) *,
.stApp [data-testid="stSegmentedControl"] label[aria-checked="true"] * {
  color:#fff!important;
  -webkit-text-fill-color:#fff!important;
}

/* Inputs de base: utilizar el color nativo (el navegador lo cambia al cambiar tema).
   Se mantiene el borde azul y se evita imponer un fondo blanco en modo oscuro. */
.stApp [data-testid="stMainBlockContainer"] [data-testid="stTextInput"] input:not(:disabled),
.stApp [data-testid="stMainBlockContainer"] [data-testid="stNumberInput"] input:not(:disabled),
.stApp [data-testid="stMainBlockContainer"] [data-testid="stTextArea"] textarea:not(:disabled),
.stApp [data-testid="stMainBlockContainer"] [data-testid="stDateInput"] input:not(:disabled),
.stApp [data-testid="stMainBlockContainer"] [data-testid="stTimeInput"] input:not(:disabled) {
  background:transparent!important;
  color:inherit!important;
  -webkit-text-fill-color:currentColor!important;
  opacity:1!important;
}
/* Campos no editables: contraste fijo entre su propio fondo y texto. */
.stApp [data-testid="stMainBlockContainer"] input:disabled,
.stApp [data-testid="stMainBlockContainer"] textarea:disabled {
  background:#203351!important;
  color:#f8fbff!important;
  -webkit-text-fill-color:#f8fbff!important;
  opacity:1!important;
  border:1px solid #526c91!important;
}
.stApp [data-testid="stMainBlockContainer"] input:disabled::placeholder {
  color:#dbe8ff!important;
  -webkit-text-fill-color:#dbe8ff!important;
}
/* Combos: el texto del placeholder SIEMPRE se puede leer. */
.stApp [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"],
.stApp [data-testid="stSelectbox"] [data-baseweb="select"] [data-testid="stMarkdownContainer"] {
  opacity:1!important;
}
.stApp [data-testid="stSelectbox"] [data-baseweb="select"] [aria-disabled="true"] {
  opacity:1!important;
}

/* Tarjetas con su propio fondo oscuro: nunca usar texto oscuro. */
.stApp [data-testid="stMainBlockContainer"] .banner-item-header,
.stApp [data-testid="stMainBlockContainer"] .banner-item-header *,
.stApp [data-testid="stMainBlockContainer"] .ab-hero,
.stApp [data-testid="stMainBlockContainer"] .ab-hero *,
.stApp [data-testid="stMainBlockContainer"] .ab-home-weather,
.stApp [data-testid="stMainBlockContainer"] .ab-home-weather * {
  color:#fff!important;
  -webkit-text-fill-color:#fff!important;
}
/* Badges del dashboard sobre azul: conservan colores legibles. */
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-hero-kicker,
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-status-pill.ab-status-ok {
  color:#b9f6db!important;
  -webkit-text-fill-color:#b9f6db!important;
}
.stApp [data-testid="stMainBlockContainer"] .ab-hero .ab-status-pill.ab-status-wait {
  color:#ffe3a3!important;
  -webkit-text-fill-color:#ffe3a3!important;
}
/* Botones primarios tienen siempre su fondo azul y su texto blanco. */
.stApp [data-testid="stMainBlockContainer"] button[kind="primary"],
.stApp [data-testid="stMainBlockContainer"] button[kind="primary"] :is(p,span,strong) {
  color:#fff!important;
  -webkit-text-fill-color:#fff!important;
}
</style>
"""


APP_CSS_V20 = """
<style>
/* 18.9 - Fichas de historial compactas, sin afectar botones de formularios. */
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-v17_hist_card_"] button {
  min-height:2.25rem!important;
  height:auto!important;
  padding:.36rem .7rem!important;
  font-size:.77rem!important;
  line-height:1.28!important;
  border-radius:10px!important;
  text-align:left!important;
  justify-content:flex-start!important;
  box-shadow:none!important;
  white-space:normal!important;
}
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-v17_hist_card_"] button[kind="secondary"] {
  background:var(--ab-surface)!important;
  color:var(--ab-ink)!important;
  -webkit-text-fill-color:var(--ab-ink)!important;
  border:1px solid var(--ab-outline)!important;
}
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-v17_hist_card_"] button[kind="secondary"] :is(p,span) {
  color:var(--ab-ink)!important;
  -webkit-text-fill-color:var(--ab-ink)!important;
  font-weight:700!important;
}
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-v17_hist_card_"] button[kind="primary"] {
  background:#2459cc!important;
  color:#ffffff!important;
  -webkit-text-fill-color:#ffffff!important;
  border:1px solid #2459cc!important;
}
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-v17_hist_card_"] button[kind="primary"] :is(p,span) {
  color:#ffffff!important;
  -webkit-text-fill-color:#ffffff!important;
  font-weight:800!important;
}
/* Planilla en pantalla: misma organización de 6 columnas del Word; evidencia mini opcional. */
.ab-planilla-frame {border:1px solid var(--ab-outline);border-radius:13px;overflow:hidden;
  box-shadow:var(--ab-shadow);margin:7px 0 13px;background:var(--ab-surface)}
.ab-planilla-scroll {overflow-x:auto;max-width:100%;}
.ab-planilla-table {border-collapse:collapse;width:100%;min-width:750px;table-layout:fixed;
  font-size:.75rem;background:var(--ab-surface);color:var(--ab-ink)}
.ab-planilla-table th,.ab-planilla-table td {border:1px solid var(--ab-outline-strong);padding:7px 9px;
  vertical-align:middle;overflow-wrap:anywhere;color:var(--ab-ink)!important;}
.ab-planilla-table thead tr.ab-planilla-meta th {background:#fff1c8!important;color:#132038!important;
  -webkit-text-fill-color:#132038!important;text-align:center;font-size:.78rem;font-weight:850;}
.ab-planilla-table thead tr.ab-planilla-head th {background:#f3dfaa!important;color:#132038!important;
  -webkit-text-fill-color:#132038!important;text-align:left;font-weight:830;}
.ab-planilla-table tbody tr:nth-child(even) {background:var(--ab-surface-soft);}
.ab-planilla-table tbody tr:hover {background:var(--ab-surface-soft-2);}
.ab-planilla-table td:nth-child(3) {font-weight:700;}
.ab-photo-thumb {width:88px;height:62px;object-fit:cover;border:1px solid var(--ab-outline);
  display:block;margin:auto;border-radius:7px;background:var(--ab-surface-soft);}
.ab-planilla-empty{font-size:.75rem;color:var(--ab-muted)!important;font-weight:600;}
.ab-compact-record-head {margin-top:.32rem;margin-bottom:.45rem;}
@media(max-width:650px){
 .stApp [data-testid="stMainBlockContainer"] [class*="st-key-v17_hist_card_"] button{min-height:2.2rem!important;
  padding:.32rem .52rem!important;font-size:.71rem!important;}
 .ab-planilla-table {min-width:700px;font-size:.71rem;}
 .ab-planilla-table th,.ab-planilla-table td{padding:6px 7px;}
 .ab-photo-thumb{width:70px;height:52px;}
}
</style>
"""


APP_CSS_V19_1 = """
<style>
/* Borradores: diferenciar continuar (amarillo) de descartar (rojo) sin herencia de tema. */
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-ab_draft_keep_action"] button,
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-ab_draft_keep_action"] button[kind="secondary"] {
  background:#facc15!important; color:#422006!important;
  -webkit-text-fill-color:#422006!important; border:1px solid #ca8a04!important;
  border-radius:12px!important; box-shadow:0 2px 7px rgba(161,98,7,.13)!important;
  font-weight:800!important;
}
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-ab_draft_keep_action"] button :is(p,span,strong) {
  color:#422006!important; -webkit-text-fill-color:#422006!important; font-weight:800!important;
}
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-ab_draft_keep_action"] button:hover {
  background:#fde047!important; border-color:#a16207!important;
}
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-ab_draft_discard_action"] button {
  background:#dc2626!important; color:#ffffff!important;
  -webkit-text-fill-color:#ffffff!important; border:1px solid #b91c1c!important;
  border-radius:12px!important; box-shadow:0 2px 7px rgba(185,28,28,.13)!important;
  font-weight:800!important;
}
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-ab_draft_discard_action"] button :is(p,span,strong) {
  color:#ffffff!important; -webkit-text-fill-color:#ffffff!important; font-weight:800!important;
}
.stApp [data-testid="stMainBlockContainer"] [class*="st-key-ab_draft_discard_action"] button:hover {
  background:#b91c1c!important; border-color:#991b1b!important;
}
</style>
"""


APP_CSS_V19_3 = """
<style>
.ab-rend-table-wrap {overflow-x:auto;max-width:100%;border:1px solid var(--ab-outline);
 background:var(--ab-surface);border-radius:12px;margin:10px 0 14px;box-shadow:var(--ab-shadow)}
.ab-rend-table {width:100%;min-width:2110px;table-layout:auto;border-collapse:collapse;
 color:var(--ab-ink);font-size:.75rem}
.ab-rend-table th {padding:10px 9px;text-align:left;white-space:nowrap;
 background:#18345f!important;color:#ffffff!important;-webkit-text-fill-color:#ffffff!important;
 border-right:1px solid #355482;position:sticky;top:0}
.ab-rend-table td {padding:8px 9px;vertical-align:middle;overflow-wrap:anywhere;
 border-right:1px solid var(--ab-outline);border-bottom:1px solid var(--ab-outline);
 color:var(--ab-ink)!important;-webkit-text-fill-color:var(--ab-ink)!important}
.ab-rend-table tbody tr:nth-child(even){background:var(--ab-surface-soft)}
.ab-rend-table tbody tr:hover{background:var(--ab-surface-soft-2)}
.ab-rend-table td:nth-child(1){white-space:nowrap;font-weight:750}
.ab-rend-table td:nth-child(4){font-weight:800}
.ab-rend-table td:nth-child(5){min-width:210px}
.ab-rend-table td:nth-child(6){min-width:180px;white-space:pre-line;line-height:1.55} 
.ab-rend-table td:nth-child(7){min-width:130px;white-space:pre-line;line-height:1.55}
.ab-rend-table td:nth-child(14){font-weight:800}
.ab-rend-table .ab-photo-thumb{width:76px;height:55px;object-fit:cover;border-radius:7px}
</style>
"""

APP_CSS_V19_7 = """
<style>
/* Compactar sin invadir el menú flotante de Streamlit (Share/Deploy). */
.stApp .block-container { padding-top:1.95rem!important; padding-bottom:1rem!important; }
.stApp [data-testid="stMainBlockContainer"] { padding-top:0!important; }
.stApp [data-testid="stHeader"]{min-height:3rem!important;}
@media (max-width:768px){
 .stApp .block-container {padding-top:3.3rem!important;}
}
</style>
"""

# V19.8: CSS cargado sin generar elementos verticales vacíos en el cuerpo.
# st.html envía los contenidos formados exclusivamente por <style> al contenedor
# de eventos de Streamlit, sin consumir espacio en el layout.
APP_CSS_V19_8 = """
<style>
/* Mantener despejada la barra Share/Deploy y empezar el dashboard justo debajo. */
.stApp .block-container,
.stApp [data-testid="stMain"] .block-container {
  padding-top: 4.15rem !important;
  padding-bottom: 1rem !important;
}
.stApp [data-testid="stMainBlockContainer"] {
  padding-top: 0 !important;
}
@media (max-width: 768px) {
  .stApp .block-container,
  .stApp [data-testid="stMain"] .block-container {
    padding-top: 3.8rem !important;
  }
}
</style>
"""


def render_app_styles():
    """Inyecta todos los estilos sin añadir 11 bloques vacíos arriba del dashboard.

    Conserva el mismo orden de prioridades CSS y el comportamiento de ambos temas.
    En versiones antiguas sin st.html usa un solo bloque de respaldo.
    """
    css = "".join((
        APP_CSS, APP_CSS_V10, APP_CSS_V14, APP_CSS_V17, APP_CSS_V18,
        APP_CSS_V19, APP_CSS_V20, APP_CSS_V19_1, APP_CSS_V19_3,
        APP_CSS_V19_7, APP_CSS_V19_8,
    ))
    if hasattr(st, "html"):
        st.html(css)
    else:
        st.markdown(css, unsafe_allow_html=True)


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
    return str(record.get("residente") or record.get("responsable") or "Abrir ficha para consultar")


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
    index=st.selectbox('Bloque de registros', list(range(len(ranges))),
        format_func=lambda x:f'Registros {ranges[x][0]+1}–{ranges[x][1]} de {count}',key=key)
    start,end=ranges[index]
    return records[start:end]


def replaced_worker_buildings(repo):
    """Edificios cuyo respaldo histórico NO debe mezclarse tras una sustitución.

    No elimina el archivo histórico; conserva una marca por propietario y edificio.
    """
    raw = obj(get_config('alpha_nomina_reemplazos_' + repo.email), [])
    if not isinstance(raw, list):
        return set()
    return {str(building).strip().casefold() for building in raw if str(building).strip()}


def legacy_worker_backup(repo):
    """Lee sin alterar el respaldo histórico de la nómina anterior.

    Los trabajadores sin edificio válido se conservan visibles como respaldo,
    pero no pueden copiarse a la tabla sin asignación autorizada.
    """
    raw=get_config('user_trabajadores_'+repo.email)
    entries=obj(raw,[])
    if not isinstance(entries,list):return []
    cleaned=[];seen=set()
    excluded_buildings = replaced_worker_buildings(repo)
    for entry in entries:
        if not isinstance(entry,dict):continue
        owner=str(entry.get('usuario_email') or repo.email).strip().lower()
        if owner!=repo.email:continue
        name=str(entry.get('nombre') or '').strip().upper()
        role=str(entry.get('cargo') or '').strip().upper()
        building=str(entry.get('edificio') or '').strip()
        if building.casefold() in excluded_buildings:continue
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
    try:
        ultimo_libro_oficial_usuario.clear()
    except NameError:
        pass
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
    """Admite entradas intuitivas, pero convierte todo a minutos comprobables.

    Inicio y fin: 7, 7:00, 07:00, 16:30, 4 pm o 7:30 am.
    Tiempo muerto: 0:30, 30 min, 0.5, 1 o 1h 30min. Los números
    sin unidad del tiempo muerto representan HORAS, no minutos.
    """
    raw = str(value if value is not None else '').strip().lower().replace(',', '.')
    text = re.sub(r'\s+', ' ', raw)
    if not text:
        if empty_zero:
            return 0
        raise ValueError('Escribe la hora de inicio y de fin.')
    if duration:
        # Minutos expresos, por ejemplo 30 min o 45m.
        match = re.fullmatch(r'(\d+(?:\.\d+)?)\s*(?:min|minutos?|m)', text)
        if match:
            minutes = float(match.group(1))
            if not minutes.is_integer():
                raise ValueError('Los minutos deben ser un número entero.')
            result = int(minutes)
        else:
            # Horas y minutos expresos, por ejemplo 1h 30min.
            match = re.fullmatch(r'(\d+)\s*(?:h|horas?)\s*(\d{1,2})\s*(?:min|minutos?|m)?', text)
            if match:
                hours, minute_part = map(int, match.groups())
                if minute_part > 59:
                    raise ValueError('Los minutos deben estar entre 0 y 59.')
                result = 60 * hours + minute_part
            else:
                match = re.fullmatch(r'(\d{1,3}):(\d{1,2})', text)
                if match:
                    hours, minute_part = map(int, match.groups())
                    if minute_part > 59:
                        raise ValueError('Los minutos deben estar entre 0 y 59.')
                    result = 60 * hours + minute_part
                else:
                    # 1, 1.5, 1 h, 0.25 horas (fracciones son decimales).
                    match = re.fullmatch(r'(\d+(?:\.\d+)?)\s*(?:h|horas?)?', text)
                    if not match:
                        raise ValueError('Tiempo muerto: usa 0:30, 30 min, 0.5 o 1 hora.')
                    minutes = float(match.group(1)) * 60
                    if not math.isfinite(minutes):
                        raise ValueError('Duración no válida.')
                    result = round(minutes)
        if result > 99 * 60 + 59:
            raise ValueError('El tiempo muerto supera el límite permitido.')
        return result

    match = re.fullmatch(r'(\d{1,2})(?::(\d{1,2}))?\s*(am|pm|a\.m\.|p\.m\.)?', text)
    if not match:
        raise ValueError('Hora no reconocida. Ejemplos: 7, 7:00, 16:30, 4 pm.')
    hours = int(match.group(1))
    minute_part = int(match.group(2) or 0)
    period = match.group(3)
    if minute_part > 59:
        raise ValueError('Los minutos deben estar entre 0 y 59.')
    if period:
        if not 1 <= hours <= 12:
            raise ValueError('Con AM/PM usa horas de 1 a 12.')
        hours = hours % 12 + (12 if period.startswith('p') else 0)
    elif hours > 23:
        raise ValueError('Las horas deben estar entre 0 y 23.')
    return 60 * hours + minute_part


def normalized_clock(value, *, duration=False):
    """Devuelve el formato estable HH:MM después de validar la entrada."""
    minutes = parse_minutes(value, duration=duration, empty_zero=duration)
    return f'{minutes // 60:02d}:{minutes % 60:02d}'


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
    if isinstance(data, dict) and (data.get("_rend_v4") or data.get("_rend_v5") or data.get("_rend_v6")):
        return data
    previous = obj(record.get("intervalo"), {})
    if isinstance(previous, dict) and (previous.get("_rend_v4") or previous.get("_rend_v5") or previous.get("_rend_v6")):
        return previous
    fallback={"_legacy": True, "avance_manana": number(record.get("avance")), "avance_mediodia": 0.0, "avance_tarde": 0.0,
              "edificio": record.get("edificio") or "Sin Edificio Asignado", "piso": record.get("piso") or "", "trabajadores": [], "comentarios": ""}
    fallback.update(data if isinstance(data,dict) else {"_legacy_source":data})
    return fallback


def rendimiento_values(record):
    data = rendimiento_payload(record)
    building = data.get("edificio") or record.get("edificio") or "Sin Edificio Asignado"
    advance = sum(number(data.get(k)) for k in ("avance_manana", "avance_mediodia", "avance_tarde"))
    if data.get("_rend_v6"):
        elapsed=number(data.get('horas_jornada'))
        hh=consumption=None
        productivity=advance/elapsed if elapsed>0 else None
    elif data.get("_rend_v5"):
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
            "historico": not bool(data.get("_rend_v5") or data.get("_rend_v6"))}


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
        # En formularios dinámicos un control puede desaparecer cuando se cambia
        # de capítulo, se recupera un borrador o se recarga la nómina. Streamlit
        # aún puede emitir su callback anterior: no es un error de guardado.
        if key not in st.session_state:
            LOG.debug('Widget retirado antes de procesar el cambio: %s', key)
            return
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
            # Permitir conteos sin rellenar hasta que el usuario los registre.
            # 0 sigue siendo un valor válido distinto de «sin registrar».
            allow_empty=bool(kwargs.pop('allow_empty',False))
            if allow_empty and value in (None, ''):
                return st.number_input(label, value=None,
                    placeholder='Sin registrar', **callback, **kwargs)
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
            # Solo para campos expresamente opcionales antes de seleccionar.
            # Mantiene el valor existente al editar y nunca guarda el placeholder.
            empty_placeholder = kwargs.pop("empty_placeholder", None)
            if empty_placeholder is not None:
                options = [candidate for candidate in options if candidate not in (None, "")]
                if value not in options and value not in (None, ""):
                    options.append(value)
                position = options.index(value) if value in options else None
                return st.selectbox(label, options, index=position, placeholder=empty_placeholder,
                                    **callback, **kwargs)
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

    def activity_photo(self, path):
        """Exactamente una foto por actividad; edición y borradores independientes."""
        path = tuple(path)
        photos = nested_get(self.data, path, [])
        if isinstance(photos, str):
            photos = [photos] if photos else []
        if not isinstance(photos, list):
            photos = []
        nested_set(self.data, path, photos)
        key = self.prefix + 'activity-photo:' + ':'.join(map(str, path))

        def on_photo():
            upload = st.session_state.get(key)
            if upload is None:
                return
            try:
                encoded = image_b64(upload)
                nested_set(self.data, path, [encoded])
                self.draft.pop('photo_error', None)
                self.persist()
            except Exception as exc:
                self.draft['photo_error'] = str(exc)

        st.file_uploader('Fotografía de esta actividad', type=['jpg', 'jpeg', 'png', 'webp'],
                         accept_multiple_files=False, key=key, on_change=on_photo,
                         help='Adjunta una fotografía como evidencia. Puedes reemplazarla seleccionando otra.')
        if photos:
            photo = image_stream(photos[0])
            if photo:
                st.image(photo, width=260, caption='Evidencia de esta actividad')
            else:
                st.warning('La foto guardada no se puede visualizar; no se eliminará automáticamente.')
            if st.button('🗑️ Quitar foto', key=key + ':remove'):
                nested_set(self.data, path, [])
                self.persist()
                rerun_selected_module()
        else:
            st.caption('📷 Pendiente: adjunta una fotografía antes de guardar el libro.')

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
        with a:
            with st.container(key='ab_draft_keep_action'):
                if st.button('🟡 Cerrar y conservar borrador', key=f'keep_{module}',
                             use_container_width=True):
                    editor.persist()
                    if editor.draft.get('error'):
                        st.error('No se confirmó el guardado: se conservará el borrador abierto.')
                    else:
                        editor.close(); rerun_selected_module()
        with b:
            with st.container(key='ab_draft_discard_action'):
                with st.popover('🔴 Descartar borrador', use_container_width=True):
                    st.warning('Se eliminará este borrador. Los libros ya guardados no se borrarán.')
                    if st.button('Confirmar descarte',key=f'discard_{module}',type='primary'):
                        editor.close(delete=True); rerun_selected_module()
        return editor
    if not allow_new:
        st.caption("Su cargo actual permite consultar y editar sus históricos de este libro. La creación corresponde al cargo asignado para este tipo de libro.")
    if st.button("➕ Crear nuevo registro", key=f"new_{module}", type="primary", disabled=not allow_new):
        begin_editor(module, user, default_factory()); rerun_selected_module()
    drafts = cached_rows("alpha_drafts", (("usuario_email", user["correo"]), ("modulo", module)), columns="slot,revision,updated_at", order="updated_at")
    if drafts:
        choices = {f"Borrador {i+1} · {hora_borrador_ecuador(r['updated_at'])}": r for i,r in enumerate(reversed(drafts))}
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


# Ubicaciones sugeridas a partir del listado entregado por el usuario.
# Modificables en cada libro; no se convierten en ubicaciones GPS.
DIRECCIONES_EDIFICIOS = {
    'TESLA': 'CALLE LUXEMBURGO Y CALLE HOLANDA',
    'LAFUENTE': 'AV. REPÚBLICA DE EL SALVADOR Y CALLE MOSCÚ',
    'IMAGINE': 'AV. GONZÁLEZ SUÁREZ Y CALLE MUROS',
    'ASIMOV': 'CALLE CHECOSLOVAQUIA Y AV. 6 DE DICIEMBRE',
    'RUBIK': 'CALLE AUSTRIA Y CALLE CHECOSLOVAQUIA',
    'CASTLE ROCK': 'CALLE FINLANDIA Y CALLE SUECIA',
    'MUSK': 'AV. GONZÁLEZ SUÁREZ Y CALLE MUROS',
    'WOLF': 'AV. DE LOS SHYRIS Y AV. REPÚBLICA DE EL SALVADOR',
    'DABLANC': 'CALLE SUIZA Y CALLE CHECOSLOVAQUIA',
    'THOMAS EDISON': 'AV. DE LOS SHYRIS Y CALLE SUECIA',
    'WESTINGHOUSE': 'CALLE LUXEMBURGO Y CALLE HOLANDA',
    'SMART': 'CALLE MOSCÚ Y AV. REPÚBLICA DE EL SALVADOR',
    'MAGNUS': 'AV. DE LOS SHYRIS Y AV. NACIONES UNIDAS',
    'SPARTA': 'AV. DE LOS SHYRIS Y CALLE BÉLGICA',
}

VALORES_PREDETERMINADOS_LIBRO = {
    'Barrio': 'BENALCAZAR',
    'Ciudad': 'QUITO',
    'Superintendente': 'ING. PABLO ESPINOSA',
    'Fiscalizador': 'ING. DIEGO CHARVET',
}


def direccion_sugerida_edificio(edificio):
    return DIRECCIONES_EDIFICIOS.get(str(edificio or '').strip().upper(), '')


def autocompletar_datos_obra(datos, edificio, proyecto_anterior=None):
    """Sincroniza calle con edificio y conserva ediciones hechas EN ese edificio.

    Al seleccionar un nuevo edificio, la calle anterior debe desaparecer aunque
    hubiera sido escrita manualmente; de lo contrario se guardaría el libro de
    un proyecto con la ubicación de otro. Sin cambiar de edificio, las entradas
    manuales no se sobreescriben con los valores sugeridos.

    Devuelve las rutas simples alteradas para eliminar widgets obsoletos de
    Streamlit antes de dibujar Datos generales.
    """
    changed=[]
    # La dirección sigue derivándose únicamente del edificio. Los valores
    # administrativos predefinidos son editables y no dependen del proyecto.

    edificio_actual=str(edificio or '').strip().casefold()
    anterior=str(proyecto_anterior or '').strip().casefold()
    cambio_edificio=bool(edificio_actual) and edificio_actual != anterior
    target=direccion_sugerida_edificio(edificio)
    current=str(datos.get('Ubicacion') or '').strip()
    current_auto=str(datos.get('_Ubicacion_Auto') or '').strip()

    if target and (cambio_edificio or not current or current.upper()=='POR CONFIRMAR' or
                   (current_auto and current.casefold()==current_auto.casefold())):
        if current!=target:
            datos['Ubicacion']=target
            changed.append('Ubicacion')
    # No modificar una dirección personalizada si el proyecto no ha cambiado.
    datos['_Ubicacion_Auto']=target
    return changed

DISCIPLINAS_LIBRO_OFICIAL = [
    'ALBAÑILERÍA - GYPSUM - PINTURA Y PORCELANATO',
    'ESTRUCTURA', 'CARPINTERÍA', 'ALUMINIO Y VIDRIO',
    'HIDROSANITARIO - CONTRAINCENDIOS', 'ELÉCTRICO - ELECTRÓNICO',
    'OTROS PROYECTOS',
]

def hora_borrador_ecuador(value):
    """Mostrar hora de Supabase en Guayaquil; timestamp sin zona se interpreta UTC."""
    try:
        raw = str(value or '').strip().replace('Z', '+00:00')
        stamp = dt.datetime.fromisoformat(raw)
        if stamp.tzinfo is None:
            stamp = stamp.replace(tzinfo=dt.timezone.utc)
        return stamp.astimezone(ZoneInfo('America/Guayaquil')).strftime('%d/%m/%Y · %H:%M')
    except (ValueError, TypeError, OverflowError):
        return str(value or 'Sin fecha')


def anexos_fotos_validas(datos):
    raw = datos.get('Anexos_Fotos', []) if isinstance(datos,dict) else []
    return [photo for photo in raw if isinstance(photo,str) and photo.strip()] if isinstance(raw,list) else []


def obtener_fotos_maestro(repo, proyecto, fecha):
    """Fotos consultables solo de un Maestro vinculado en la misma jornada y edificio."""
    activities = v17_linked_master_activities(repo, proyecto, fecha)
    found = []
    for index, item in enumerate(activities, 1):
        photo = item.get('Foto_B64')
        if isinstance(photo,str) and photo:
            label = f"{index}. {item.get('Actividad') or 'Actividad'} · {item.get('Maestro') or 'Maestro Mayor'}"
            found.append({'label':label,'foto':photo,'actividad':str(item.get('Actividad') or '')})
    return found


def agregar_fotos_anexo(datos, nuevas, limite=MAX_PHOTOS):
    """Deduplicar anexos; no modificar nunca el registro del Maestro."""
    current = anexos_fotos_validas(datos)
    merged = list(dict.fromkeys(current + [x for x in nuevas if isinstance(x,str) and x]))
    if len(merged)>limite:
        raise ValueError(f'El libro admite máximo {limite} fotografías en anexos.')
    datos['Anexos_Fotos'] = merged
    return len(merged)-len(current)


def render_anexos_libro(editor, repo):
    d=editor.data
    datos=d['datos']
    st.markdown('#### 📷 Anexos y fotografías')
    st.caption('Arrastra fotografías desde tu dispositivo. También puedes incorporar las evidencias del Maestro Mayor vinculado para este mismo día y edificio.')
    editor.photos(('datos','Anexos_Fotos'), 'Fotografías del Libro de Obra', limit=MAX_PHOTOS)
    st.divider()
    st.markdown('#### 👷 Evidencias del Maestro Mayor')
    if not d.get('proyecto') or not d.get('fecha'):
        st.info('Selecciona primero el edificio y la fecha para consultar al Maestro Mayor.')
    else:
        try:
            disponible=obtener_fotos_maestro(repo, d['proyecto'], str(d['fecha'])[:10])
        except Exception as exc:
            notice_error(exc)
            disponible=[]
        if not disponible:
            st.caption('No hay fotografías del Maestro Mayor vinculado para esta jornada.')
        else:
            labels={f'{ix} · {item["label"]}':item for ix,item in enumerate(disponible,1)}
            choices=st.multiselect('Selecciona las evidencias para añadirlas al libro',list(labels),
                key=editor.prefix+'anexos_fotos_maestro',placeholder='Selecciona fotografías…')
            if choices:
                gallery=st.columns(min(3,len(choices)))
                for i,label in enumerate(choices):
                    with gallery[i%len(gallery)]:
                        image=image_stream(labels[label]['foto'])
                        if image:st.image(image,caption=labels[label]['actividad'],use_container_width=True)
                if st.button('📥 Incorporar fotografías seleccionadas',
                             key=editor.prefix+'copiar_fotos_maestro',use_container_width=True):
                    try:
                        total=agregar_fotos_anexo(datos,[labels[key]['foto'] for key in choices])
                        editor.persist()
                        st.success(f'{total} fotografía(s) incorporada(s). El libro del Maestro no se modifica.')
                        rerun_selected_module()
                    except Exception as exc:
                        notice_error(exc)
    st.divider()
    editor.field('text_area',('datos','Novedades'),'Novedades y recomendaciones','',height=130)

def official_data(record=None):
    """Formulario idéntico en conceptos a Ejemplo.xlsx; edición sin datos ficticios.

    Fecha, residente y actividades se ingresan expresamente. Los datos
    permanentes se recuperan únicamente de un libro previo del MISMO edificio.
    """
    nuevo=not bool(record)
    record=record or {}
    payload=obj(record.get('datos'),{})
    if not isinstance(payload,dict):payload={'_legacy_items':payload}
    # Los datos generales cuentan con una sugerencia inicial editable.
    # Los demás pasos NO reciben cantidades ni clima ficticios.
    sugerencia_general=VALORES_PREDETERMINADOS_LIBRO if nuevo else {}
    suggested={
        'Superintendente':sugerencia_general.get('Superintendente',''),
        'Fiscalizador':sugerencia_general.get('Fiscalizador',''),
        'Ubicacion':record.get('frente') or direccion_sugerida_edificio(record.get('proyecto')),
        'Barrio':sugerencia_general.get('Barrio',''),
        'Ciudad':sugerencia_general.get('Ciudad',''), 'Hoja':'00053',
        'Disciplina_Principal':'', 'Anexos_Fotos':[],
        'Nomina_Conteo':{label:None for label in OFICIOS_NOMINA_FORMATO},
        'Rotativo_Conteo':{label:None for label in RUROS_ROTATIVOS_FORMATO},
        'Clima_Condicion':record.get('clima') or '',
        'Clima_Obs':'', 'Maquinaria_Conteo':{label:None for label in MAQUINARIAS_FORMATO},
        'Seguridad_Check':{label:False for label in (
            'Casco','Chalecos','Guantes','Gafas','Mascarilla','Auditivo',
            'Conos','Cintas','Rótulos','Vallas','Extintor','Botiquin',
            'Polvo','Ruido','Liquidos','Cerramiento','Limpieza')},
        'Novedades':'',
    }
    values={**suggested,**payload}
    values['Hoja']='00053'
    # Si ya existe un libro guardado, respetar sus valores aunque sean vacíos.
    # En libros nuevos, el último libro propio del edificio puede reemplazar
    # los valores predeterminados si el usuario los había cambiado.
    if str(values.get('Ubicacion') or '').strip().upper() in ('', 'POR CONFIRMAR'):
        values['Ubicacion']=direccion_sugerida_edificio(record.get('proyecto'))
    for block in ('Nomina_Conteo','Rotativo_Conteo','Maquinaria_Conteo','Seguridad_Check'):
        if not isinstance(values.get(block),dict):values[block]={}
        values[block]={**suggested[block],**values[block]}
    return {
        'fecha':str(record.get('fecha') or ''),
        'proyecto':record.get('proyecto') or '',
        'residente':record.get('residente') or '',
        'hora_inicio':str(record.get('hora_inicio') or ('07:00' if nuevo else ''))[:5],
        'hora_fin':str(record.get('hora_fin') or ('16:00' if nuevo else ''))[:5],
        'datos':values,
        'actividades':{row_id():copy.deepcopy(v) for v in payload.get('Actividades_Ejecutadas',[]) if isinstance(v,dict)},
    }




# ==============================================================================
# V18: Catálogo de capítulos y rubros del Excel de estructura de costos de Dablanc
# Se conservan EXCLUSIVAMENTE código, descripción y unidad; ni precios ni cantidades.
# Incluye códigos únicos de las cuatro hojas de planilla, priorizando la hoja 03.
# ==============================================================================
CATALOGO_RUBROS_B85 = (
    'c-qvxPjlQzmM8dCU|BXh)TBrS{v*_uW*rJr601mwMN+D)?&<>q0g_N60#QH&i`1HpSzq?PHkZEUu?}Un)w!p(IW<4Pex>#I+yS@)&+#Xy'
    '>Yg#Hl9Z&spF7;2KY#!H`TzM{6mF93F3pPXUjO4iem9?898S#dUVe9Rd2)7oe)?f{d2;pLKfM&cvDj};ew@!f&*rBeKN!a+#`WpR<>l<k'
    '_`Pxd>GGr0)MicRr{^ahuCZ<X3kNW+t}j2$uRmSR>Yuyp=d-gnv#0<2^z7vFRQinv<6#3iK7oF&Pv-AGUKz8q)6b`~53^&rQ=hfGJo(Jv'
    '+1Odv^6vut-Tc#=%l7Z4Fxc7U#p#D@=o&uke0F>y4bPn5@IJ9AIsM2+@%i-l^7Lc<YnZb+OvSI2v70bs?^)Z~r?b<!)ZP|9|9Ca~aC$vc'
    'uwaJHPv-0wGKGtCe|`D^r=I<F^5OVY`Un%I>G*_AkFEjC&dJ%y{Q6P>hKZSfJpaVH!=;iI+{B4Fy_|nKz5Zz6_Z@%4WoI~_Pp{;KwI+!3'
    ';rjC988n&A=O<TE_ZCdi@yGLz*Q}rVr!YaYPsh^tTQD~#zx{N2G2;tmoLpU;%um^hAG1Y~5MiQbm)9q6XLGikMm3dE6XdEZHWgPNWlpwW'
    'vVJ|gJU(TgYkWSrgh8Dlf4%>BdHN6RGvGh#_I`6rIAgQ3x&=(*`HVGTdxo*l?+t8@vjvmIruON7u-W>jm*3^PMXu~OzU40CO&YGZab`Rk'
    'ACmhx4Ws1gmo@v%>fish=2eq#!`WuH-iB$oO5*i4Gd5u!8p}A(vq=1iptG@;ZIWlKsd2H3^KDqLpM6><%W$bQg{eK3;fz(d-i0Zm{-QK>'
    'TTP>EZLFSt{^KrLheooF(%oasdMv^;`$CNi7UMG8?(h?gDB&;#{Ks9Euaetx8q|(qP3ElQY7X-3J}HuH9i}*hMVkF_7Zd#j^*qd?C@!)*'
    ';k~^HpZ=QB69PM77G*0&w!~o;aekjfj1W4Io1&-?@Ok_tER0ARWc-j6TlOzvTwI7qZQ&!fj1P#WzykAKxXhC@4UOwOTo;=x-^Rwl)ouYJ'
    'd}JT<C1n3N54UliglSx`pPdx!6M*|(7>9;yY~s9N(^a%57X_3&mJ4Zeui2Q4?~UWQSY?a}zfj)-4D@anMt2EgoTtBpMx4g?Y&<WGk2h@E'
    'Sf?<=Y#HBDp9eg}dfdhl?g5zZw+xHX7kwQAGo8iTJZ97Xalu$BUK=SJHqIt&6E^pS!PfEV=QN4<@--lUv1aQf`}7QEjLlsRgFKH5rCGoS'
    '3DY9Rn5?s9#%7m&+F5uXvM*=<W^9iPbI;}xzxuYEmLw#3jZ57Q#$JW6=(k~d&oI8SNojurO5^wC=3%QBui~%Kk#Y0y|8Mkz<xGs9ZeK-='
    '=Eyi%!;XWoRa2cm{Q@dMn=33+ks#rcZLp`mtZSx*$qpkvUCj|$gT9KfNwj5Cn59el!|WcIZMa}C+dSE%@dH5*jCq!<Hc9*yb_o+pY(GRG'
    'pNzQN9kkF=r$T1YY#kR|Cfx3FSWZ3?J_h)OX{}boSUioCH4YeXwhHqt+uqp=g3@nMvR+o<YLgYX-G%Ubm2we5NL_wl@oj`+S2L%G3oG-;'
    'x6yr^FNu9;O>so!#_T|B)7rqS{5YyMW~dATkGF+JTG+sqvwX$ak+0$w_IS>OPQeLvV#eNwCD}Y|(x;z`EfXHbI=;<`xwY-~WH7Oe9|$94'
    'qgy9i*eq-lZ^%R5c3KeZTjC{WwA;r`u_HLjc3V_zI`UnFhig5z2k%(F+4=<+5}_j3Y&?$RJ2v2`mI({CGw`5jxA(P2Bb&OXKeM?j<*Y&$'
    '=*jMqC7W+1+_qs(oUrz^oAl}D9f<ZOnKSwJho&<74j$kPq0PIic$u)H0mho{z8bdsU`>p>9~{?+R^%bzSVOQvc9s~<goD|x4$H=^exKB$'
    ';#NAEi4PIXhq6zVjErY9cB-<SmB%RqTr;LDjD!1xvD*u>cNZ-ewqF*4BUfT_`PZm8b~U{uZ@Qi=y<tx%)TP)?Yx!P}-06V^Ha=Y!pj30N'
    '$HPI|N#a<<Zg&gDVfdx3$#YRJ*olcgO?NOnj(DL+L1L<=GAWZlY?G@}X}epXWHT;4&krw*lUMV<Gmev|pV?MmXEI}h70m+Pq+wyKYYPKK'
    'UCgxgO=fJfO?JyR&<)=|@6v3+#wr7u9YB(epiGGe%=|W7<dsALT}hg1oKiGLY9Up~QooRAAtv^oXAkcJd5{@P;kX{0^0Zj)O46za)Bu_Q'
    '#4@-tlOfSbDI!Tz%;L7;EM?2X)KJ`u4-wNXm%zX5-`kv$%9bz^*?wTmTQHq$o80r=yV}8KFfC(6cV~Cg*+;I9(&?o1bYyNP*;GdclHg<='
    'Wj8sT`{#UgMijPnPm7z1vOFv0dp)1Et^JEa87SQV=U`73_EZpEO_2bk9dCafrn_ZKv_w1cr0(2+V`rVU$AEUB+)%Nv`H=vFA}Fm6E5LRq'
    '#&0AmK!*Vg!Tk|S*zc>bfa8;YTxFgcOdc>xBqKf(*y1A2Z<4fkMqHFF?BEwI&#NJ+?a(8br3pL#(|ApF>WkA8H~VE=+GzE74~25)O^k=?'
    'k7#9_yf)$QBa-RR<a-xq%bfA&1^eVMC;CPqo_)nH$g8fwM^j|%m^P1s3G3cEve-^C4#Gva6$kSRvQsqWWQJYEU$RPNJxD{Q;KeW63x~$q'
    'EQ-0Ub{XF=>Bk0j^Z0||Sd(h)VXKw%{L`%hFqnH{upz<l0JHwUI2Fh6>584Wt2kUFDKk^qc|ncR9z4o<*+qM#PLH)afX}(BP?EE&5H}bY'
    'Qya#UOe;J%sYEZ^2=^Il9B)I-s(G@=fgNj2{tH*7nBx6KcG9WC@Ee<*efT0=M^w{5`RXy78m%=pO&~m~<MBR3#aNuXVJVVh2}Q<0tL<I>'
    '^lw52u}=jZx;~zaEZ@PoI-gw|Hj1LAI$?3!?J|y6i!5ckW07s)U?x5mZ8b?{R}MlWT!QvrX~>LCP$y*e2OP?e%zn$rRkAxzyVe&wjkN)t'
    'LTq$t8@ELQnq_a3bj!K?HfIRLCwg4E6v=w~^k=j$<2B>5d&*sCal7kU(g%%smcxuMfnoC-HZRLi)((A_J0TR1pb4@CHSmq`kHEXPyZE0}'
    '1cQK>HZ?EyiQt71dpa&p&-N#vMOL4#sBlB`g?!jH$?Q@kN|@2H4kc$X_j$m-91!Rh{DEgmbKz|f%979Z`{b6N7P2>t=-X}`0bKlb5;;AR'
    'RQATr!=zwlN|Qfqt|?u{DI9smJGK>eDeIi+;*4XeiND&#DL5T>i$i9wv#&|vPn$pkZkyv?dSv!<b)<jUp_`I#>D{8>c<~Q*)L402#M{Jx'
    '=Gkh)HV~kIv_m>$9fIQ_C3+7=%m^DmNUn?uIUNDnUTtc|QAJZvv1+0t=`?CAH!5tzNEn~w+;Re!oN@DCqaO@2s0K##Z;dz{mAR46N1N+#'
    '=ae(+|NFlJ<L^zEDL!n`m@TnM3Sz^VWB9iLq<)Y)?1SbqM&nyBbjO*LeQ_Rh_IN~DfxaNYwB)`=#c;S^oG&p057R~X^w%&wBEBgY1KBk>'
    'Dg#GJls^5kVMg|W$5e+}b$w*w?Y|hCBwps>3*HT|k?7_nTjFq$<&_iTo*BK#lIX89Mxb9+<ACXRN3#3ud*cmw2VW3fm_xd-u!l{Wpdf!_'
    'r`HmkxTO^?Av!|#584vPNyw!8?J8{Wki(@vvq>O4GrPJzH)b|d#-IMhFke}J_>N#4hs%9I0>?e741(C4+Qz%fqc=W6{NcMp<Ln1#$_(C>'
    '4$M)(xB_ChuL5>rxjG_OA&PM7udY8GUzNj{1Z<eqXW%GwSiuOx@D>yaVkyk7#Wk8JT;e|BZp9B^Cd?)zP_1FQJf1SqT~chqBW0<ZcmqGx'
    '2zT3zJF}R&C3rsFt=8nHPlwK5$H+$RKiyb{?J?;nB_M1CVpBmxh6FHl#?kaSslyKERoG$XPFi~r4Ia(e<nXhpc_29=ctsO=jA%j<O`AI3'
    'yPC6=miDCCvYI!w9SWoMt?&wmD3&lLWVn+D_CJTo0L7ewX3IGG33Od~1{G9Gi|R1Q$%9>?f)D`nYrTxK7L23Drytp=nr3WOkD-6QD2yk5'
    'm>hYa)*EN^C%QvXe6IsV_10CNPR%{KZlDjz$V$T~l$?T(DuT*THs_5x6v_^XEQc_;lof&#o3gpf;{s#XQ9=!SN?Hl+hG@tzi==>NY?wDW'
    '#2vU}eZAb}aTJS)O|%Q)ymE{~5B62Lt?pvBLD*@(joSu~!-GNf{shY-X6*j-SAG<50WELU_NK%b(45%<+cH=U#vbP<3q4=g$qMbRmmq7('
    '*?(D@WTk~xhE~X91>;6d!|ZDNm}g8;;BhMG>|6ANY{}W;yp9;W{rwM?Yu6HsAB>#+R4#G6AyX=;Z^V?2(IsiY>jZ<f8L_b%Uojm_U&%TZ'
    '0VVPPgsUpG;gm5n6YdF`bqz2P9Ab)`*e}#E(B^N~8Z4>;y^wM3wm+<`f9sA>*XgGgmPRcER0Nmy$Mv&;%R}ypAzvj7tda-z=Ruh?aA<r@'
    '))D&)qEOp7x@+f$I6Rm+lr=X`zvSC06-pggEBv80Ry!Tej29A$%0k|l#HBA`zD(9}POSz8DwUxtYp;ylYd9CPD)C2(A0K?y#kd_FeWWe^'
    '9!WjinsP)0GQ+^D97kt@67&G}m5uoMG33A~2<`}bOwSgcH)N5l4vzOQ>sZQ(!E83RKJl7Z9!!X1M@kC4)p>6XmB%d|gEukzEhGfY*i>y`'
    'ix}4sxBVIxEecDe_zX^zHPH?h_)W&x4Lmp3jOp00k9W~s*lLN36fWa6e%E{UNgy@v*m6C*HY}!@{vE^sAnUaLQ47T7N)Y4M7#tm@@Mp4P'
    '`oIALZ^GEbeEJuNRlUNWAH8T{cq)cfV|<S&=0sos1c~6w2qb`@2Q8hThebYShd~U9G;aL~7JPyQpYX2d8(`%x6Q;@KA$Md%jNi+F01k}d'
    'UDqIZI&KrfR3NI6=s_?DCp=8=k7l6H-({~2$7Nz4zA1iPS}H1?Nz>Wxq0>ocRb0gLpamO2G>CCwT-qdz9=;_sOp!^{fC08;r)A7HS?xng'
    ');COcWO*C#uI5O27Xp}5u|sIsFoFCNkQHz>PoDlIV2cUn*O7s8Ep9va@jr<fb}`P(W(O60DcMbE?2{v0_sNmV|FT^wGXRi0X`IeCAE)0u'
    'SL##9uncgc7u4hhBpbRdP7XK^5*<$^R;v002BMeCtUD|WCajaMkjex*aUW@7niF03F5SJUX>UyqEe(mKA+a@INGW{gFoh4uoraKgm1*I#'
    'a#BqTJu7aPfD*D!6?Ga3gMRHd5Uf!)s>o1+<V)lI4=@VU7(g9Xc(ge+7#s{Yp@BJsbq2&d&Y59a>f#w2*d;SgH(MqO>K`ik5JnJbHX6Ft'
    '^qGC!XAa?Xh%I)f_5s?ok=xgx-#@v}<xa(OIdg>ybU-`ca=bPWF)KFZnW(F<;CBs(Ij~Ia@nX887=S0p*QYITqpjga9fVySxYBA~LqykZ'
    'H+IS_ng?(oCA8c^_uK{hsM-$(;(5u8x~ynp)+c^_#exsu58s{r;XCw=h>IL;;;AOu#{G`jeMpQCFN>|yUBt#ONja(|DgxnphjP2vL5!T4'
    'TWqx;E(gx&_JyMcf>~u1rQv-D@sJqu58|SZlu~339TUREa?&~5G7w(pgH|XV`QzdrEZmR}TRgdAsw-C}_*`=1>6()e1fE1t7;o|0%I1&}'
    '0ouwBv~J-KAq6%jXak-Pkzod-^f^kSY1OEPZ8j3gQNq~=a^elw7OE_+#uwv$B?(*&fu$j^bUoM_0$<bKX|`u<cqZ;`7^Y)TXCJ?Q!nY-O'
    'OLEv#nW3bXw&sd^8Sb-kGfPfVqK9XlAU+5{zwZP=aZw8wDaQDS0-9iP6cUOguVckTbKLZ@H<^7*yiJtE+pHi5$9KC#x(e5~J8`#3X=us2'
    '4ulntOn~p}NbmM?dG^tHyi6cT3c|}>oNg<^0Jd9YtR9wkIS-9D*%s~uw5i=zSNff2d(Ey2N$k|tYqEZA_@F{pKY9B3;WcVXEl6Kvfns19'
    'UC`01^D>7ObWjs=*O^u$px%cCKLqJqOqaX)yAUy1X$;`B2hbfkLJSIY0a~LG%|67*SVT*6=Hf0ay(UIjI#P%&=GameP&GJ)fesiR^^z9&'
    'U7qc>+7UK%bRJNE3x&*hzB(dQe9(zW3F){^2R~j+G$d-GKkr$NK$ph^Dwjun%T$Exc5Ig>YvWH$J~VfT6Le9oR8g%E>eMFQB?t;k8pcaQ'
    '?RmlgK<U{k6qgm-$_BIaVQmKZtXz{Yr??J0(BFh%0L5SAk8>Bfb`!uH>2Sg*fQsObL?A6fA3tWGw~2=XXq#eUw4GipdSB~nV1J-J(oG+9'
    'yACGOvP_0d$+R0!UZf#--lPX_<dTU{K~rrmk385DYPy56fiQSb-CAhm9W~q9P3=WQ1i%5;LFVZ55O3gkCP-T15A);RAI*-cyQGkT__|Xb'
    'qN9KK9Rd};5!+#Cgg<$lh3oOs4d7>JUU+U;mW|n%5(Wp5+-Gbd82i-MeVbkSlR-vb`zY|uQh`?PfhgXcB3W(H*a#*M!Nkxg$-Xrh$D2Hu'
    'mLACN;UaYs)EV?0A}u3JlU>f1a*cdn-A!Ebi}SZn|1)mQCa~qZ4ALB0YNzjYn!DUlTU!eqmbyRq6z0AOvo(}p9DvmVm=3#$4KqGtUz{zn'
    '^|JjT!9+1FcKLmWfq|E^#^g~ke6uyojF;z7Z=H%t1`zsYb}8XF3xw-7-ZuP~A-|W;-I&6=$?uvLTHkSCSE=b`x>T?0DQI0YQ}{z^iaZFx'
    'bLC2(GVD@H5QFm?wMiPK9;9Vv|24zwzey`n8Th^4-pLWVVeKP4xF_%z?fY%-^gs?a)cz$!;@XN^^6$C(qa|>IY414ABQ*A%YMhh7<D5>a'
    '$_+1-5X{XE9g_9pz%dEJC1X+s+z;^&9{Y&J4<4b2PEGA00XsTh6XI=Z$?;u9bFQ+$@2W_PLTNYL*6g2bAAIHcYL%NYE_!6v)X~p$8N`q;'
    'M~grro9#oc?}RBZ-D`t#xJ}(=(5edhiQKw64Vf64*Ss1JCs88#cyVK*$;j{YH*0xdZJb*s#51Ok-X#&-TxpHbAFgXi-m;dKF%7yywJRxP'
    'rqE6ML$xh3kLMh`!+IRXJb_Cumti}!;PFg@Ry)u(+vQX|k#3GkS{!9DauC$l!K*dEkgL!*i*KoZoRX5W+&i*Vh{sb5&O*2+Cz4_L*siXb'
    '(tPB@G}TeqWUX^Lds2Nq;~EN{nXip^zrVQpbZ(qnT|fQvhvV5LBA%Z!X;WAJGi=Cr|HF6ehi9`7SaXjT)nh-<&=Y<8YTs7X_vL$ir`=u7'
    'svZrp%RueYer}h7jP7izcILb`%1VgykC*HpSk<mYhOt_S%&nK?$Lq`4@uy3+OzyWGDo>F(X4+ty#ue|>#cGgs^Kw_i6+v-k@&u@OwFwKn'
    'iw4<)Opv{xy5N$%#(-?vZ6zj!*kP!*aQ(2kQ+ZNFFL{*ZY~ab)uuX4|qUTpBnQ+aRxs$cg94(sojyzdM$%dWVkKCpz+90O7(B3y%=>Bbh'
    '5ml*Y|20Y5?3nds!SxiX@2gK2Czo$N&MuD)&p^Uf6B>!`m>9FIjyd?xA9E!$h|xOgXs3OG=09tkZbSZ0$bnv6pPZZ-Z!c%(Cve3H(j3n|'
    '{`S)e0oXiH>Jz0G&A2$Zf*PgfWKIpWaa@?}yck4>*@}2*zYSCBbjM&N##an}YNJGGh2vBrix=gIc8gpp%~YP<kh85hyDdcyR;`)}w#S2)'
    'jfQmyj3%)u{Q;%-GAUwSAhS81O6NddLV^|bR=nQjYi~t;f!jf0`XmYZ<}sUdXeU0c;s~}JuHsmd+X?Ja$V*~I$-75}wlS{F!zf&Z#mjpU'
    'k9&#t?-QsYm$R0}f*G2iXY-nI2UpiGgl>uHJaD;m3!1#!RxUBx`OP}Md_TBjo-83cj@R)`Sas?iz7H44R&aG^mn~FMS4Q4?%PgL-J@a<s'
    'm@#dQcDY)qBFyC4C4tyD1WczXu9{HvVM!o6Wy!W$>5&VzX*L|M*~hn*L<tTx!R!UI|89YW4_Wuuv$L}k?y+ig?{(wU#1=v94<)t<;)Y6W'
    'u_+y<(`!~u-<;^Yp(0!E>puPU8}#Z-p0zPU;T9St3^4l3_DL?4{1l!6$@^U{-Q{~?)LK`vq@%AwaIn<bTh8Q1XA;rpi__dp!&NG3(glHg'
    'G##R%*m?wws^T<MDs%duIouAN%4z6S)?|S4_a8jOhL&d4w)fl9%M;OO?>=ixde;EO=teAe(>q}QM-Q=U{+dG5>pO=mey@MPN#rzTqQ3{9'
    '_jTe<8uqKCDVDZ9)8zt6e7a3TUOI1fD=eL7mg)Z9#Gr+bYxRN%A9@#vO@D=C(HF8ufk1EuAjnPy0>W)TR2`Qt04zjXI{^rY=60+_TkX-K'
    'c$HWjUmguLsSc#KYYIVnB7LBwuj3&iExZjVf(4@FdO>)8cm{!fNbzfnaLsxXnyOU!Wlo(O;f)saNXICiIqs|4AK}UTiuRp!>OTEkU-&PT'
    'gJmq;B%lXnm&I^&9yg(PpeyPhS`*Y{-EuaZ$pcUSFHz~<-EZ#kBRZ+-{ZuKU8J0oZ+(%3xI!soH{PNG7nHA#33E;3p>h0;9%a8Mub9~05'
    '!!b=PU@%|-ALZKf3Uxy$o9@6JIzSft4TrBic|BO??Zq}bZ+VC*=)T_J@hk>8OjgqEf33sg84bh<&^UrzZt1Uf@d5-XXxO5J4n(?}>*PE8'
    'fU?OEMd-vEA}Xdk7*iB5OhG%x_;Y(p_})0Q2o0359v*Fo`o6a6Hy>0U@~^BJU!Mx;@f-@7C(YsfzL(bHH#V5v-0XX9uAsTO?BSXk(XO|<'
    'hELyi`sS_Go!qXsyM~9wb_aeP0oAYuZEs&GM9dt~UBjBQnU?!@Hh(`mo{dv@f`+wa2cs8w)QmS&HG4X6fe9$TcBn<@P|LxgrpsROL}o~j'
    '(!LfF148jkj#z9s!bL%NB``x|kmty04+j%KBu^kOOGC6XE8p+Q!<O=XO&e_NxZlFGA7b;3w40}mP2zaZGwuHtLv*x5<aRUC{imAt)5ex{'
    '{Qu}7cFmtxI|pxg(rFpXdC9ivddu_vZqoQmTszvD8*AFsUE$b~0vIoS1LcI8aHbV<)8HF<EJs8Urqcl)ab=u7&!;&-+W;*)jVSX)3EyJ7'
    '!-UW(SzCmzi1xPJUUw(b7YQ`nB*7Ovw1FX?C(*n~!}T9`vE-G-6h$K}qIg}9U3gL#C;Bh+rDmtM>*Qv48}cH*n%kznc-(H6hI}MXKj$&J'
    'A67f^Sqoo0et8$(l-2pU`AW3Jtdg>&cAEJsc6Z?l)1*66qI{l1&<+RbC&a9lFf)o*<!-nT25PghypC5PXu!*m9i{jx8_}FcBsXztT!)+F'
    'E73w+4%(J0JzH;+yKH?+!ugHHSJH)J*!`#LyJW#P0l&qVv+YQgkM_k&bK*<p;SP?$GD55Wo~g`Z#ft6HxZMgv&{w?rb1pqHMYJ@rB^1F`'
    'Zy0Urt+Ev_#~24pmS8dY7X*;i24b}9-71!f(KBear{K0#^z@fJ;+FOL>EBi&KhQRo$!)Ub-lC@Q;qxeVE42Qu<pQdszAvjs-(n#nh@ta+'
    '3x#)~G#N8Tw@Ld}x_@ZN8+XRUqj6o%3@^i`za!$8mtw`jZJvI4ATE#kJTxpb82hTvSY!1fDN4ImCE@R^3FMfwgfQ(y>VDD8-ERPxI?>p9'
    '@5L8kTtNtLS?Ku|<FkYL49a3sOop-8Hj^bwHdV=uT&cdUaelBVGdDn*!!ralZ(+c8KWsc0G^og1;Tmro*vFI}?2Bu(Vj}L@d7K(y#1<~)'
    'j|DL69^A}Wisp_f?(O7yp9gE}^R?Un(xxD2rs}$PoIg%S$@C3pOcZ}9=pO{R*?0oxCMkCK(h=d|^?5+Ed5~7Jk-wk%0KX)9>bMQVbI#)j'
    '{lSls#Oe;gBbZ$qG2%Mqd*s~$pNIV}NvNv!=};(d{{bdi-3TLUSM_57Sq=;@UeCn~X61-@?bBkjTfvhcn{ZulN!ez6Q3Eqvi0gBDOqzTu'
    '4T&kE0*D-(^kHovcha~PLS5KtPl*U+iK+@W$-26=(3$Y>^YE310k{0(we$!jHe}3fT)lA#tPW<D%2N%7b2=)qJpGK9k=Pd|Ovy|MYWG=}'
    'sC^#GDz)ks&+54o16LnVN2ovljd<M+gCSf}(47XPDqrYD9e=H_))?8>74aIq+eLCqthgtizvbD<qZemvh5i&4cLb5C^~&N&7*?L=Jiccv'
    'TfLs$HNzN+q7Q-EjYN#a?hFMq2hbT3@Okt}4hb~r_YV#DJT^7<kgo38^y&Z@mFibw#f_N#HQAu5X$hYl@|Ql=E!8aG*6^G9_!kwp1Q}jq'
    'S~oJ{15wqoCc-<Nc?2c!!G2t1Dek{J{w5mgc<^KtQ#o;kT5!Cg6bFDe!{PJi8KsN{;yDAUx2JzbQ}rr~_`|EMt{U`SgC2W326?;9%2#6>'
    'BtAtlhqQZM4@XqeDCE3>6@kNxmqH<L;+NBw<*|wjV2DLo23Gnm@r;L2@GXc{%JVqlELvwlSVs<zOE)^>COK2xe_p{gE1pm*CURcd{1C9`'
    'SM#PC5HteAbOU|HoL#EH8lmbz5Q4NhYJg2uRI{!<SsJZG(Y2tummyEtc@kCWy%dGb%YUz<yU=*2e+WGg*DT(K=@OQpypa)w{N3R0*>Cc<'
    'DPCaGuWi%a(lJ>ou%`p&G0I-1X(RgLvqic05j=PIy5-6XL;=eRu=SyPVVceTi2U2$km483Y@X&)<8Dw76lq4Eb)W>IpyeQbHZ98eu##B1'
    '!`cc&UCVxYs^V23F6K0a94JCz8sXZBLaHEIs9Cfd^oub1sep{`J5?l)5?x2FwddYB_6bIrp-+8`TeW)6^wxsQFc^CPw~VD>60~03L$+8v'
    'vXOY@clXlz6|In_QXK+$a=-=17E!!!S#l^&?GXVRG+|{ST^~6;gPTep79lZ%iF0voD!Zf`$(jfx<1_dYgcIQ_>;b*YIN+{8c@{Xu9q2UR'
    '+&U;iM*31*)qX7P{ucyNlUecRZ}UP7?avCzw?iYSOC-#^a#OVRD7>RG$KXq2(NI<5^Z97-A|d75Fioz-Yb(E&UdU@quhQ02ZDm0QU7$<j'
    'W$E0@jJIk0kStQD{mw_on!ini;`X^JxuBq9`NaanV_aouRQf9yFQp2rnUFxff^*C4wFDSi1TCOI-fO~v6mZAmxbJEX9vzXY-nnA4xN4tt'
    'cdvu?xq)s-z1qq?9)yL4*6O3>!h$BY`-$5dl7@2npjr}w#&J7w+R}?gp_J){3a*?RgmANc@-)ALrO-(`l*JU2*`=g3t}f<%rT}+ON<F`+'
    'kF3VMHJOap=Oa%_$Yl^2p&A>&F*U^aC_o>v)u%@`=ZiRNdrSkz?ola4AJKFAaGt5Q_9zUE-dtG&J^rT9UL^}xSrLkZkl<Q#DvF#;U2;Nv'
    '#U7JXC(0<}91_q|TdN8xXVd&qr7v4FTH7d<p6}SD;Lw9truos)L~rA{lX`Sj{g8RBn?_H6E;j9;Y{0#x=VFRm$`h|H2lB-@kScEzii2xa'
    '55TF#l*%Ykn1b(hjjnHayEwu81tO{!aOY{$h#n#X&`sJT9J2q>LrSuVx?!MdARf0(dUTK`z!jpJxx^v|zpfm<hrGJnVI3BFVH_GJK2p$>'
    'hUh%q8&i7tKmiRJ`un4}34uq^WVk}^P8-Vn2wHE_qc2O|ksc&-<finxHl;^dYOgAdY|-REQgV4VdFt{c6qOMfy~sQ#V4U$V76Idqa(-bv'
    ';_2snm8EQ!yso3h^Tr;TW5E3uxU?tZ>Nx8=A*uBV<m(IQ5SFfG$hSQ(4shMvL0(h0sRC>vVa1>f7uB=+gpjXOe<1&1nWze17w`DyK0pNb'
    '_&V1G-wFkKhM~&c)6Y=+Nwsu|x%X`fuVJjRvDI?q;A#^xX@>SI`<olEhkOqZr;y*=HDy|fIx^$nS^!-EFa_jlfZB3k$`ScMMS*4t(R&Wx'
    'bUBFpy?rA!H&TlgP*czdVD?EvU#a0;MPXXIRwR^Qm9_#5-!M-?+Dyjm|2NEns9z50bG`;<pX&Yd5Y^k8oE!hmu-Ph{8waQ77sM(D)^p{p'
    'Ef`;e-`qhPu_M2P(eCTlr0<9eo^2oun+BIq#A-M4N{Z{)T^{7X`vz{ma#IrU>RT!}mk`fZ3+rLs!Mea=eI1m)KP#L;1fvFn)Cp<5hz%4>'
    '_1@oxZhC~zCAF&!20W{uT=nYG6|!RKb6R69k@ll4fyn#cJCz{d*8-ZQ3mc9WT!-5jBn&kKdwBd^jW%=*%NZMZExq@xxo(~uWn^GT_HR(8'
    'sy95YT=b%z!mT^|7_M6q<bGG=uq}KB#0^EBx;!GRv%7)RIAfLMc{}wZ;F;#lX{gW8zKNI~ifHBTKF$mNHp;T~)W@_HWycf<vSHg(cpWbA'
    '{mPRlV}IYaGXAGLeqKG7%)a0Ycdc5@EjH!R^YbAW8mBx`e%`#NsA=m<ZA({J(P|Z_t(H~Po#`a=6k=r56yq8-FT_))?o?DT6m7XzZJTMD'
    'nrhPv`ilHiM8-tJ3vs2ts80u)$d-xf>-cnIhYZ~Dt$`vvkt>23arp!d!@f`0Z-`mL5E#unjfR+>o^9enDLp_B)NjujTg8R;!}y!(<rd#S'
    'mK%10nldpKUj7;$Z?3k7+5$sQ@L3`YMu5MX&B+1qkkiwj^V=*98ymC%%+t&B`zrOXWYn`XypMVL*fq@GyEI#bsX|*XHrps$UUF56Oq3bl'
    'M$VrR3!zH6+rFDP6^}Y8d+sL9wxBHtNhl`2L~89ZOWxsI*e%kQZFJ@l!<Mch6VvI_6!k`?$pU%@g+jfu-qU2$ZZd}^<8L?NV+k1I!i9MH'
    '%NksV=mcudge@L@P9P17F=rZ<GQT;!c8g6af3cJX++duQj&b-)P8p<1hB>d?P}WiBSEbzQu`q><7uC-pdBw=HCHuHm6St|rY|*WmCr^LH'
    ')c98wR}eEu6>g5;de;^%XBZ)@{Q@5#n80Q6PAtD0i`kcX`1BV*Gu?hP;G`=IIb)``IO&uOKuLB&*>`>B*_20ss6rOPy45IJN5<U`?V8S0'
    '5gDqK?kz7k5mg4o7BteE3=5a%Z7y``6N#P8Os;2m^{CotXoi(i#U$m2V~v3bNC;k-97q`tm(Re&MNQ{S9@ziN0VuTmAY+PTC=0HN5}stO'
    'ZUwh*{!MwT2DIkK-5ReJl_hlGec>tG8|1VeHT&N4o;jcLFiZ!pW++j_5C-tR!*ib&c_kg%W?L!Hr4sGJh~xRk(t8S&{>fi)YG8t65G<Ze'
    'ye>*fV)qXhqcdC$Groo!m_Opwa3wC?yyCJ>aY(}5nyh;B!C)rOI$~n*wL@{6D8?jj-|FQdJnv9_&&%gGk3&c~i4p?H9zNF%<!azmW7wuD'
    '3mT$SFH7QQrJbV?xaIrv;<i+hL-qA#s~y+C;HYM*e|4Yip5LTDsPu>I1q@Mm+m_WHXpGW-mgzecv6p5|>MYu&@b=EpB;e6U+c>|b>G6VD'
    'oWne^mE$*ud7gK<ip%i6Ttu1Dn^dX`zL~RL@^@h{`w~@|_UC%$Gg3mPJeDH51a-?eEz(eL=J0ozc=J5I$&=d{J}9U669z}5!>U#dTDWdQ'
    's}-L>e6LrQVyx~$W2yu_01PXK`3CJ~@x+eI`IQ@@$@*<o1p^YFVVi<j?#P#(^4P;5*GE*2EeSh$(w8{n4J^||yk<&Iu}hndfGH0v$SQQB'
    'vC`I3?HT??Om$qgKh3vE0=aa2gTRrbrE-NGulaz7qN6vrm$;*6kJ`s&S$V^JPmrJ$&lb+ONxq1Ex0Da#TKX8&3G6rQd+s~LkyD-&g><Tt'
    'mXWg|w84aQ5a@}euQn&*rbP8prJ6g}XljWUhPZg)ZVsXuW(J7Yo_;4DZEQ4gsze0IKEx5f<po&@8W*=9FSJ1X+LL|_YxT%1lW;+QGa7PX'
    '5#nuHd~pb0P&r*^4{Uj1pFL6;gg_cJNqIZtI$mwWZMrwvw&ipLQW0jUjPZNTJWFmIHpjA^$)sKOrMi}NpKXmp;{+b^Wm<0KW+4Evx<JaN'
    'cac4otg)qDCWXcgsUW`3G~TnW+wuu0Y?7j!5t_!4mkV@+c2+nvbHy{yp*Y^;r=c26B?e%VQLTSDocYtwt1u@R3*9;;VQ=F=9#C&<+IffT'
    'Cltg;M~l9mYK8_;n7S=6jNDS}w^Q--5WMNQ<aa{Jm9%VOc;q4#xf;P^wOV~%B!H0^LVZo7F-o_!cgs$sk4!@{ZwctVCOrfRKt+DbG&D%T'
    'B4{MX&-HEr?=WHM3EdC_V*hH+IcPqu-t*!LT!C}Pl#Ng*F6EAel94&N2)syv{8eHiwxw9o&xl>-=LjP~NE;mhx8wQadIXr&97xz$o+Ph9'
    '33|I?1Fo~-SxiMRp|PW7c=!MbMF-%jd#u`^y7peZ`@2Wxca3aS$`2j2>@nPOhJ%+g@x4kt)s&>-4r(2dqtPCE`f^N~tfw<e(hM*@#UE$X'
    'HMg71bqzcDAR2<;YT5R&DO0o<Xq@a8jMMMJ4d@fhNXqiBKyvolKAKRki36Rgjl}D66Kd{k14$s&kgY9b##m9jrL-e-l)&rZ%piyuUogE?'
    '<ga4E8R#v3Z8)n{+heE#iTc%>tVk-yLUVrLg;;6tJy!fqz?skTS^<KE(lf^^=yhyJz@e`kvQ+wU@Rs(Ptpex_ZSP5qMIK+Go;6#!VvAQ|'
    'Lgj`{AGNDzVAxUOc>xXMxmwj%OTiz~RL?za<6yWt+EPZfpwo^sQb*)XhMLyx-K#aQ0oQB_8j}%P#sOnPY8WjsX0zt5ws}j8IohVUl8x=e'
    '<B}qZ$F^x;v|Ncm>er<0GaFvI0MEO)@l=`u_&Os2XL7T!(?n9u8<Sgy(vh%*)VTOIgSSp6JloTB&y53ba(?>Zh*}6!n2wd9LJX^ZX})^G'
    '7jhiAx3BDdE{Dn0L8@~8eG=#FqdtJow$;>=avP|RZ;VV_Vqol#rFU9kc%G`d74L{d;TK-tgqQ@y1u32_rd!G<eyNtsMXESSQC@xLqvVol'
    '4SGyfGykP{2$_wJ-wsH#oIAB?KvM&V?voE$Y1VheNCC#`g;~aF=?Kz<XaUDoa88*BPfFcxk*wu71NVE=D2{?czf;&T2v;R<QVyCBPg>ft'
    '8jJ333AAtICU6pdQ3kf&WI}9mlV=Tm!SK!_;~+R>>TLzLQ^K~-2?(c(sd&+OYm&7`TlS=YNE((h%P}*xiHi2CbTgIMbOf}e5;e9oL10V!'
    '`dqFfs!ono`KWlZhxB;$J4}0PM=^oQl^(hlRRrohoMvOM88Lme5>7FtKarHV2#a=6E`afeNnX?H5WTX|+BKHzjwmj`OhvY)Dv@g`XhA4h'
    '><Mxp_C&k+s&ME51e17-d$rhY83g@Nc1t?d@}r=fqXdLSZmeo%$&F>(n#r82LI87B335kAUI)`r-z9j;c89HF2M^oh*z%gG+-q_R*}ewo'
    '7?oOZC_-k-dIx3)WDVybhbEvj?AgFLT|dTJi0QD;IuHMNRNd%q!?e>J1dRCERg^KToLCx1<;K-6WvAN?6Mgw%hZ;f(uCl{7*_PijL`>PY'
    '+h$S6Djz}TOvSkRr#NTU=j|@V?P6arRt*@#GC|D#oSk3LIChCs-wsqBb*}?3wVj0q>r#_w2%y#bej*AQ{Y1BO$ye6ukzfk8M?dNU^j;ag'
    'fviz(3e*od#8j%Wjj>#xDZsS!aQxAbRv#oUg0x6V_EBEs<g!j;MP3qF__7jVjS`GCRp#zCDYiV43OP8pRiU4(S}K%;Y#n!bRfh{NuQ=N2'
    'Ro|_u-?`&MCY%=?Y4-qEniqMTs_7pUH(}wi^=-UP$`|Vi7*%a#3>m_EG9j(}x^{PhlJ0zUoo+fQT2BzS<gp~h3$#*UwF)`qu7TinAk?=s'
    '>iZ=Ggx>)XcgbBe{u`&UnvUBHo)hz~6TDE13S8=SJ>K>Jyh-<2co|0tvuD2$CYfhyNZ-J*%Wq$YX<Noy-1@y1<$ZZTmPu7^jQrud^FMs|'
    'cgAt@^z#btNE<=p1tp-jM$mf%y*+~7A?TeE^fp27j-Z#`*-$XNF&G5BKZ4%Wj3F3_LCxiK3`UzCP42;KL5+Zyn)}S;I1_E68rxkQ{nVbo'
    '=H9bJK{<sMGZWb{u8jMY9>5v`P!5~|um_DmOYRIJuNIGjaEC(ZXTci+p(pnxa(*3h==2AjlwI>a?e0@YH%9i6yN;KX>*%>_txU6*Ry$Xw'
    ')k&+TvsJo92t^WOaHQrdrsWKraMKwM)Abd+2rF10SRFOFk5f%cV9Hriq-E*d?bDyJrz3W-J^ghTr<epT3QA)ZcbddZk&JonW)!E((70ox'
    'd-^jZovwJ6G@kZhdsY4BHDQLGW~d&k!bm?u!B7ZW3o#uEp`BpmaK5;>JiPc|SguSuEuG>j-DGIj+P%eZ<WLYR_TOK_L9@scHHLv`=aJvq'
    '83v-CNoP2SeokBhL@foLt&I>pPQ&KM+YeAbSN?v0`gsfX1JutUCJ>g<Q?`$Ha1oZ-Ur#^pgnP3Yb+ov%yn7?~F0~QLY?G|0-oQW^m)-2p'
    'YL^gRTa(p_9<yzZ5SO|SNB7)kW;~WCdCbhmI_I4kx$y7TgRt!JSY3jgr>Bg<(_`^Oxb@R7ydfk)O^i*EXqtPyOKZ_f>-Q3xCMg~5prxc{'
    'E8)03FC9n2)^yrLT4d=kspj;K-m|H}JCnXHwbI5hdzEwjaJh3?T-%DWb+KDzvPE6*uZz#~!wci&)m&Lau@V$Ccjm35%=mnHq1&F4@kU`d'
    '!!We;?jZWkO@w~E0T^0(e;9qo2$a_!#cN2j+a1&~-D=`eK!$03YxGTE9G4d<eV?RT&8BsGnb#$E9{T3Wo{$4KF?A^e+nUW$-G%G5PQm2K'
    'HKH}o!OzzGTAPNzCO3ltnWXu4m^SX#Tz7dmn_ubqPJ6Vtu3i@985S}h%z&;p8cQoHTz?!^$Ebo)NO~OLQl6bFT&G3^6$Hz2Kj7PhD?WTI'
    '4;|J||B2&}-0Mx~4VtY7PpfBYCxbfiC<JTBAhdGavxkuD$FC>~_yKPk$4{r6#tT#`{k-KM>ICpw0Q3xYg7|~fo#A-FutDipWjcC5TC9Gx'
    '$hIi%#=BeHOY3R-G)IrWbDQ`zlk)`T-4;v^5!Dl4MSm29W{*be*nmO5zDi0D#W}0!NMQ$=HR&azv}4QcBO{r>viitK$!OVqWTZM^IelcL'
    '3T(N3WTY)<d3|K0FU9iv$Y?9W3VP{0l5PN>P;|D!DD#tJs94F!jf=~ZtJ9AkjEj?Zvo~ia#y_5gMU?%MJP^t?YMXoGk)Z+G+JAfmn!P`o'
    '=6jvJ5h*m6GMBch*r26GTjYo=moZ>^X(%#X#(+8ckp-P#16}e2?({jhdeu*4#Oym>9je(qrQ_^j+S2Yu(=WGUk4Mw5cf-u?kM$A|?qCR='
    'iqkN(hedn|7=Jj74k8!~(Pg=z0X`f?H?{6$IE)U$9KJgGaahC0p`8=gZrIlQFk8z&--q2?cLYcm!@1rVARUN50;G%S++ZY-c9N%KfHWYI'
    'Z9QP6v}5tiF=NrrjW^!@&<k|W{)Wbk=Zpr^i*wH%4d&{>ywPB#dQX7)qrrTw$>s&4!KV7poOXgeU-$7R-=qWm9`Vg@^94G1>)XJ0NSq#n'
    'emgo9i|E4nCS7RZ-ERZm^#$Iy;UgWq|4rbvOB@Vylxp{>KOF?1m$boT5P)8~2E%;4dWjlXg8=kWG_VH&=y@n`hE1Y&rv~n@5$MRhVdNdM'
    'BJc-8XlG#{FY3s9gT(r-r@hf${oGH@-hOqIRxf48j8E+$)Y>n0`lz*(Zf|(Iqg$^(Hm?6-zrSZ4Y0yvFHP_QY<n&kDjqGvoc5%PI@sDtK'
    '{dO1slj;<2J09&u26->Wgv7ANU{H=zyOBxWi(!%&?idWh&1^T4$@?Ja8TCnsnemUGrOP!#J|`FrM1RwCG!FS;=46jJf!)aS9T+FDo7t1Y'
    ';skax*KAmvz-}h=42%=l;tAx^LAQ*HD9mqV#%V9JaR^t9nU<B2o7(jDLA=_nlBFnXDJq;1(EbQ$S0JB-_x#Qk4H@hK`6<T|C7MA__dtA`'
    '$G3Pb+|U)?Wysd#S;%>|(cqeUPAgr~6cDXF5${91V^4$Hdw>q&E1sP@5E*wZL&ldFX*hQexM2fT_Lh9tv3_St5|Bgme4B?D%i;#2_kol_'
    'H>J)eV)HO#2W2I!+N2n1R7zPV9nuKnh)V0UQy$o@7kzeJn;ipzO7m6xHA;CV5HDY$WfKhfDYQ4$UX@D=6O?WzrFNs+-r&4^Pe1>IZ#Q%;'
    ';^6xYdLFU|UvSWKurnEvm#?1_XT)uVj#+U=+*i=gi!<WFLdVQFduHBu%+j9s7rJI_&l?OKv$y9JhK`vGG*=F?Qf5=P4s0HlYg`^6&yd$_'
    'f;4^ejMp@R>%&;UZb{3rcD8G04z7KuB)oQ|(%L+fHY}6Y<{`7eiL|yT2q&d}fw`_E!h(bS6)z6ATjLK7bqTRHkCDyuT@=F`2noCcS{9Yv'
    'KF_wpz}e|)lP4>|R{$EXb}1iaSlqC$fXijZ{OankRd*X8@KuxQCj3kGnNVdr!)K3In=smn%BTw{O##o}VVWwurXijH_(g%6iYnD{Sz|qh'
    '=RoRl9GtGBBcm)M_o6)@jG7G@5CIO)-i9f!Vh6?5qE%UPCc2Bkk-Ki+H@A6!t@)0H{)ojuT7Mlpbu)HUz;n3E;<zjg(JUKLBy1t2Yled1'
    'uKo{Zh^bXue!5jqt$k6IIxY=Rdk@g-<G!If`=WLY)$Ia>jY5Ba8;IAPT{Ct2CXFBP^`XX?Ix7M=;)yyHp)r6ygy5U_PaN*|;03s_R3ldT'
    ';?Jgpo0LEF3D509!L+PM?z85~HO~}>yFkW4D$7P7;)V!U8;-dDjv%0|y`lRCZSR9j><t3i*&DiV(5_@Np{$e&k^na`nOX2sBj)4}-x)t@'
    '5~dyQq8W%}btC!1k(_R%U^tT3jWivO6m%k4lfg(<_l#TS*tH9ENY>c3(<0ep*G`M%j9oh|k~?<ov`F6AwbLTCcs!{#>ZomJ1%rXQcbzpI'
    '4Ai~p?8#uD?mcIlgMqrYoNWyS>e+F&J#y7_%Wpd)S4|7#j$Ab@kT-JGv_SsIRnr0mBUeoeL|xBc#nIExP-M3t&qZgrl<sC?rwa5W`-iI>'
    '?!}SD!yiajoV5UMx(}vHSCMnL4`&S1nU?OO%98&HdL?<WxW~A@9Tg_<_G`FXvI(DhlZi2ZdydbqzrQdJENk-lmHn34lf^4<^8PDr6glpQ'
    'K1|b^1lS2vdPnr(wfc~jx!sAsS0}{Fjw#dGifzp6!YHd)@g<EFjF=REDkgr^v`@!O$+YakAJK(78S)X03W?m##uCOw*?TY%y^mTCS26d+'
    'OQSHDs;_DdS})bbX#zFf^O%+ziUX_P<K>KTV>#jx$jZ*z-zp-Y-Myg|1A{;wwZ$x}T++X{SfB5$k4&VESzCXM&<LK{Sum21Gn$RfiQ5>*'
    '%aO9uosRhIpxN8V&b-Mep*eBQ<|L2Vt)BU<Cwb)8T4ScxQv0Yrd1z<$xT%HgrRQ|#e5<*gJoCHVQQvLZirgRFo22R%E#;1=&bvg%_mUSa'
    '1*PagqZzsF+NH8LY<6{hZp@g_M}Zn;#dep6oyADqb}`dYfDQp@KL^kz0G;OmlJ!g7b`8{Vpah`z96(buO8>c_)B*+10wn<J>mA73yLeIZ'
    'eV#>0h_#`PdTS!M+|fG*$k73rBY|ukkTnuWf4jit4%|^-I-qt%kyZ+;c8=YCTJ5xY{j}Pd^!sT$CodR4ucMu|Xjfs*{whMF*lptcJ}DX{'
    'b#&8bO-4iMX45i9L+K{hvPMI7&9!BZhtokdtK>DSwssA}?WWbvhSyD}T@b&UPRkPtYx0K~Fqf(<<DwP{(!xwf!RWrio{WOgO`>g%g6UdE'
    'wTfduln%xo1!H!N$Qcc#oo;s=lnzF{pTCDzM_~H<V`@J)*dJ5NYSaBOyZFwTJOfvUiLbn%o$&M*Y&v8m_|vF5D&1W37gBXpy4ghEX<cCY'
    'Cf$7*Hb%FbPCH7kn@&4UzniXeWrH4C9UXAmyKs@L<1PBbw&<jeAOk2XjQQE@>clubCBz7Lt6gJd90U`H1i&8_Jg4%4><2xN`}Si`hWDc+'
    '!Iw%TKe2b4Jh|Q8wXSmgo$(vPJ@O!zY1?N6(4AWPPB-ienVJ1J=+qwdrFicQIyKEB>lwW|qdxUp^$I)j<YAkzev!ZDAFaPL=HVucpecAI'
    '(gjWFH7GBBdo_>}EzNAj6g;?l!xTLp&a*1E?Anj$sl3dx8+_~PCQC7TGzx)=p~dhem5Zd>KZI-C<zB<#u5p(;{)Qxjy4*)M%!}@FAKN$B'
    '9&?Ddy-o6%?nN!VRkj`AQ}EnDl!W8J^ZF?XD!-RX<D>C{eoBIB+Dk<^nfysV6*VV*JJ@qIKOGm>ae4g*DTq7d)}6~sI`k)sYIada#;D6n'
    'ImjG8Tjb$a{wU*Vk==;*WBHS`Y``>#+w)aKJw=z-V7N%bVw>ECD<0Fn5XCT!*<0g^Nf&(g?&1SHfO>ur8rQ7d(_gdK6#HS#g+1Xke2H}y'
    '_8lR$wY-K~_vHM#b5nT@uWh5fP4-}S_0axIkNaIAsdNXc{^rNNA-VeI^oYC;d~!he2s6&C!pPQxX&JNMy=%1(wCx?T{Z!lDquTT5mdW-3'
    '*<R7sib{RGgYt4qCT;LhSxwDxkNYB<<mH<%L+F)*DP|%|M5Dv?j-9y!02FN~-F*d+JX>T%^z_SyzbCzB`a>R*SHL{^-tOw@m*p<yM<hHq'
    'Ii!1!qF!a|WE*FBJZK7h3PO%N55Nd0jCHn6ZrBv?7?uoT9DJC~U(irf#f%xq2k3&l8om7JZaf*k>gZcM8^74-L_8nA%=~5MeOCv4hUd>*'
    '7k*<kWE!5ju~m9=8W{M%Ye9qfLqKe^8zdM4;!e6j_??;cO>(;{?}th0LCcuGNculoT6P->or`a)H$XQT(+CX9F+MLNgULvm#?;dUIvO5#'
    'L36-h(jD6R#_W!6Y_}Ub+Q!ba3=NMv{LUK!XrTO2A8w%zVdLO9t{#Z32J)hLl-T?z{1UTq!!xzpvb?|hz-&ATgJy6$D4*@cN3)B}$f9j_'
    'i*J&!!G!oEuc>)1*5}&FEQN$Vd?e4X4(tc#)#SzZ7W`#gUw?wPvyGoVtK;^$-l9;l@qTZXe;C*A-+wUPmRDF}c&=IK2Nx$89B=CdK38p2'
    'P=~>J?Ee9cC;@69ofag*=^=Ohm~IO*cs8cjz_bt8jKWHo57sM~RI3-7g7urQ4-D6mam`ngy^w<zvdPD}_gJSbMC*Bo=A?}%@SJ@Yiof#N'
    'lsoEMyVbMXWX`ZJUi)lbB*aH*_)XukkQ_7HiWtfUk-L|_`PmN>qZv6lbd5BNG|&I~;sMa1Xb|H!cEt{^Eb8!_5qRc=DV2SIZa3fp-wW2z'
    'dfl{JcunzZs3G<{h?5kqpmM!r!H&e+urMCMp#I}7#-a$i5d}@q5^>Y<jJz)%3{?oD>fQEj9<Ga<4DW~MypBVj<26?r%Op1l3f;XnoH-$8'
    'DvUCA7O}p$^pVb&@K3iq4vn`R+h}GVn)&c6>qR9Z`rxeo!Ev}>S@=gUpxp!Lp!^#0rA`mhun%<mKG1pil{=Yiw>u64C@;Quw(vtk`Dq*5'
    '>%_KxWXIo+aQgq)bNaWD{T^g?U?X}k5;y3c4GFVNJME!m>+z#S(^{=PRk49!2AZsA>)<qa8rtrSvgg8BR?ouhhidnY73S8oMxr{Z`&Kwk'
    'C1EK1_Dk@Wq=w@%t@d6c?=80tb{nr**LWyrYo@I4vK|btje!sO7UkIog(C1CjRwFczDS%9zx{ZI(ZnQ=cFblej74%27D)s?tGCQ@n!K`{'
    '$>-Ue{nMQp(JNlBQ`1k--H+*gqkgcUND_YyjoMRkp4}(HC13^<;~mVf=Df7q+v=INfp{QQKP21h8Is6OT*hBSc2$}2TpM5ZWb(kD{%iDY'
    '_0uxC2d-~6Xq$5}&(_h?pVy#rGPZm`Cnqf_&}!zNe7XjuXSaa}<9YJ1%`_j`w079^LOjB_&(bYK72xbRpPw6(BhNTAOwwHBFP-XCC^qh!'
    'r=N?2=Tow+U#o_eJDHtdXh!DmhiP_UQisMOkHeqXq05tk(A#riJix6Y;}Gwm=;5Z_;Q-^HwiRBqCe0yP{#kY=e?ql0;LtNBv)}X4NpE!q'
    'BYKGKp3p=OxziJ<=pk1EBNr!B)<-vh*8$K=?CXh#e0$K*_kf{KyF^YID4}JJ+hXHqyI{m)suHOA&{P@PG)-&A8DQm@ROcA9lg`LDm8@as'
    ';VX77eg;yHiF8QV$cmRtVlw;kOUz}FI}<-Z$9BAfHF&IlvQ5Akr3H#>x4+xz`);RAGwGvoTQor*iSUVLwf`f7U_lzMrSfGUQZ3f~Pq(k4'
    ')xZC%aSZjt&d)Vp+jL(=++ojjXymr*PZ+$2?m}j4qvFvR5-|O7vRc7y#f*05`r>s}(X2>T8(q7WkM8+?KmM*NzjY~6eX<l!DXjVxZZv#='
    'NiSz?1Ebj=?ZPG8J7T=`J575_ZlA}?Y~9)1##j-iIxt0b$-WtGIcaCm;}ONQ2rG}yHd!rlU2{w*zy?eVt&L-y66`gn+&ui1IeDr4)t*wH'
    'AhKR6f3=C4{!EMePC#MLZVdbISIlxl;_`0-qbVdU)|F^o{_V7={2s|_=_RW_X?N)GZt%~bO<}ev0^U|f(XwX;eUwd5^(~z5w2AnF@jtyb'
    'h1;g+nE<a%(K7+KK%4M3jP4R9&^5+@&l9RahF{~;zS(9I=YNWM3f;0U)%Ys-!*?%@vmdM@XR^{@VIsBAI?W5P^=44h2_lAgx}2<|yWNd`'
    'KL%YqWj}PV6u$f+b!}T?yEY{NYaakh1#m`xye$DRP2HSG2ypY0J4{zux{R67)POh*5O*ICR|S#hVlal+z}Rabz6#>(1rn$rdtppf4A0&J'
    'WZKQ$j-I=x6SaHa-J^F?#R&HK0;`F!4~Wd?7iG{urG((oWTv5as<bEmO3^YdFq-u%i+ilw)93O81g7O+ZQ6%S<gJ9|3rE+?kYJk2a);VX'
    'SpQoy1g{V9$B}@3<=-gP5#zc;6Tv4)czdCb%tNLltu=MvfdX@VvQ1gKbisW%WMga6&8d1v4UhR%Ia-P2tGPlAl{-PHE=?WbF_xj#(pfsX'
    '>pbb~(5|MCecoy(Y%AkF(Hxc5#A)T)#%at=&v6;I0Ok%$o~>+>u7p{@^Li%!<osGg>37e(o-pW{dOfA`+`~3og;l_>k59~Q*3{ecmO0UP'
    '+RbBX;-pI1!)0pHB>AP<SyB@wsR8Rro!$}aX`An0mwRK`JxQL1Y}+$Kz<%%Im2n<_$#|&ZY8R&&6NC`+FhIhAW5FmwFf81&HyWIvW%_$<'
    '(_h_$ts}Y0u(&ECiiY{qXDEKnox>`*)$PyLtE-mP_5)dCI@SuyK9kq78fvOHQcbliTLlP412j}rYY&CYvYHmu^);KkUooszv@&epxJyz7'
    '*UEwhQja9Hhmbb(QQ;Qj)w=t~9JHzZ;nY~aGK_vQ=E*j@$<_&kf1=%biS@VnyBSbYJX!2`2;lViEe+KZhtVsdmR+ZuBe$Z;WlqJEd6UWK'
    'Hi_9FG5blZ28q*4(lo)DKmi#7>3#aIaXj7jztMOLEyvewTs=v!_hxmb-(ah_+NtRfQ0}BN!Bh3QQ=OWf8J<r+*D0Pu(75M%6PD%f$@t^b'
    'pVa_l>v9huHF>{E9l^<kfS~!CHhZ;v(KVUfMdm(7Ubx<K<b~@uS3sWk2XmrHuC?6y+DiSExY2(+yEuVksr{W(ip1<9X>Psqc*&&LXAzvm'
    '97;j#8Wj12>e1=!(QLSg#vWXrrH@4sK61;mJHN5|Ld&H_i&pL(lgRMaLuCcfBlCw&rbSL>&_nKYlTUld%Sby(jVH>R?2kLF+Io?OaJSEx'
    '$2pr>-S^s!l{CFqo~<k<CV2EMsBH~{xlH2JI6nPAgEZ8keaGSqgXtbiL)`2K(sbNSUNt?DM*rskN1mW@oGmkB7HyMz%~}R6qxwY2LncUg'
    'mYZa5!pk+Q@OI5k08n92>YZ1vqz5=PNhD`763N#hnWK<wJfAdB)>tV0_t_(%Ed8jQkx=?Uxg(K)hiN2lB$9sld?j#rc5=*<1L!y-YII6#'
    '3{7hq^U4C^&Ax_FOH!cjbuM!08O_N$&Tlbog>MYJ{LJiQcIqW*wV9dJIP@knQ09#?$j^VrG&rW}UFa@o`Miur{UnhQ*B1jsD$xZ~yV!6a'
    'bpVQ&BdT1H5JWz>VoKRkXm0hdUgDYiQgV`l9{Os&V-s(2#WcoSg8a0_i(6^J)X_37k8x>QUsQ{A@HYXux-5++E;exlcOhC16490q7>Qr<'
    'pS&_`Yt|fc{o;-F5E$R{EQS8zWA2M)_>RVtQk1T-#-vs)Uh;phmFB2);^<AcWAq>X{of3JFZmf0uegq4IXytEf;9GK3z0xjwe8}NXAjiV'
    '4pa=If1vol5iGzq`<o6`anHWbaz?QqFC*=LFe@vz?`e7GL`anrX8zvBYc_J4c&a77$g}$dl8`mGO+3Zdn~G<58;h#W6!&hjPYZKi6ddQQ'
    'Bz@Vjhr~L=U5Do`wr4iSEgttR{x=xAatNF;vq)T0sz=?k;A-g_xZr<sel34>BJ;o+QsmQNTW>36e)fGvtIB{n^LvkB&skMODl>`l4052l'
    '`Jo)v5H6`9RX*YAeEj<RGW=WejUtOTW%#qo#cG{*?IwRqWy15-vKLWHl=Mhn1E$oSuqnmSmk4^{3&sJLzRvTcuNTk-KZu=G|5D-aL1&!3'
    '=SG|AN2fCPw;J4j4B&Uf4I8t#*Dh@6XM60s5tHtBgfc5WMh!v!v_Ac!QYw(mhbt9jaVF@igO7b{=%7dFMEj0;zpdzZ#B#rZB!4~+?YTFc'
    'zv*Y8ga6&z`RRY&2P((!b)G<>R!l(s-^QHe1<hJ_r>^uF<GxQ)9BWh2{eUFknFch)RL#B-CTqu_f(ezcqaPC{bQ1YJ7g>>TM<Y#U)tOaP'
    'ib%zrG_^gDnkU-=YtuN628Zht7vVNE-i2wJtZ7XeGdT5E1ZkR@PpXXjew{t2d8MWD(*!lIDE9|GqE$_kOr?t}tEXS?lawbX%h7H%0SmMS'
    'D1*1$X`F1BQYl}Ck#63}w_!}{+h2#lRLlW=v$g*AeOtr&wsn2`&*<A8*0-aXbL%<Z?KJvc{KqY~JE-fw$!dF2*R`jZ_sRqb``i=JCr$5n'
    ')3ZOi>4R>1yN`a_P49G$KNwlpG8kRgG8kRgG8k9aG8kLeG8kLeGN@l~*Iy8Akp!JrLR-zJJ(ob61m>jY+GmTv)Mq=`lLuQ<ODy1ZFEoVK'
    '<iXMt+8y~7`e#Z5em`25T0>J{{(86$(|8+7-LTpC<1S9|2HoUAbf`{?*H7uvl>QLPvPPRq8uXG{6AkHf2q`sQYoei)a)dA>rI{(q97<2m'
    'mt_s4*L}KW4+kIz;0*r&dH_mcZcsUQx9z8dj5X7<)?dOFb?wF5#%X*bf|Lh5M@Rlq6mAWBUyCf@$u3waz6K5HmWump8A2rxF`nIagX8e_'
    'HLC)i_5zk(4X(X}8}Kw2wY6gus3p@Umt8w*qAYv>%ktoSuYqEk3r?5-qFdDtiYVDW{c@A7arKy(diu*6cg;(2kGY&YH{O5tFyCcfc0@O6'
    'CN6UKX<axp;(4%)TWNYx$UEYVnY9#;q4t$>wLB18#9#RC7=aVCi=SQ%r^ei>C@u;%x5;|T%<C{^yK#Z3x(73660Sqzcc*Vp>1pQAd_&Uk'
    'qY?jc2Ty1j2i3zMO#9nizBI~mF2KGqc&WG19s6?~r!S4wF5M=;w=az>zZLK0=5Q%K`I@nHVNI_7lmLO{To~_?TXtG!@%PTv+Y8NzFxLQ{'
    '@WRw<rpZA@Y-&8j+1fa<Awk`_Mdmqjjdyonjf3EP!B#K2i(eR*;mQc!Ets(JJo5TNAfCF>&NrstvX7@}ZK6hEl&yAaQNfqI%oFh37&X~7'
    'C(4;Vm|A7RX`b*QExT>a;o*2yI%Fl<T5a$|A0SIA;j5Bkmf#Lj`)0d`Bu&Dc{uDgIV|a2VRXJSL2FRI63E^1)#$rVV&#cxQSsq)N@C7nL'
    'wgeU^a9kBLzBhi8Y<bzl97<aiG@D@#%~=f21#Tmi-NG0c@5Qvt!Auz3FWg;MkNr?iwx=`e1Ca-V4=eIm7Np!%5}zgko)cS$?MEy{MK+y|'
    'G-jrwjhX3aV`e(en3;|>W~O6}ndx|AW;)K8nT|DPrelqnX=%*B2KaP+A<o)XBdgWOYqhbfHn&yQf6%@oJ1vb8EK^e7p)=l|o&4Bp>efxY'
    'o~E@}GM%2LN|-fb^An0F>%2$lz&U?eX)4t{gBRo`G6Sj~40cf+m7z5eSsU3GibBx^%5rvagvP=9W8a!=UYf6Lrcck$KVH2AN5&>9jMGcV'
    'CM?*0Pgd+dUK$^F+hxWMO7@G-O!OSyKq*|~&5kKE+n2`aoHdyoO-$o#_weeic?1O~EaUU>oE^)o*DYe4UU=S1<97+9F2Nf^%<jBp%JT~U'
    'y~|@Ghq80V0n--#{r~yz9{c0POT#(}{<~?dR>oPp23?iSg-L%Ux4b&O(y6}Ea@4Q1+F$t}^J#;=%WQtvGFxT)eix&S|CjhGtFDGGYSlx4'
    '33mPbJxyV^DabV1DYrr2fe#+jGwyRJ9l;K^^K8k4&uioO>d?1s>;1d;r-xVXjq@;%!q>+8<R_lYMPs^6Oh<-%|C>WVwI<H{-!jy<<{$p>'
    'VIZ)F<&z3fHg@nZp>h4jxSCySKFwEKJKe48J2dZ)&rk79y{5#orfM6DXd~UX!!{<>##Gz5#)=%1bgPbI9cf#l9?ahrU@J&InKl(K8HiSv'
    'wv{WjGN-y$c+6+JBBp=sOK%>;4_kP~EH`$=!U&E03o6e9Q9u#hEmuh=74;C@e183z(EO*|zN(n>!{<vV`@;2jT8BgPM<HN4q?5q<U(14<'
    'if2Ka@}6$$P$HCW>W=e@PD`6dXLZR}-DtCGDomo0cjT<eI87YdkJH3~Vw@%p6eBcoXg*342ZB+WI1r4~#DQdlCJxO<Y2u7voF)zgy_z^Q'
    'Qqq&s+UVTm*+a6zmoG;uvPrjY37y*go%U+AMq8~N(6f^Bv&go>KgPY8e}gY|JNw&1bsx~Dr@9ZMNku0{9HnFm2$1?ZQo4Dd?P-PgD8n=g'
    '*AbNEq-Bv=r%}>ez=Y^+Ka9oiVXBi$w={{COt7#2>(<z2yXcOlLhS<r)sT`Q3g^^!wL-YoPd*L{v~~}zX>|}g+#$#)>X*^OD2fXZL%P!6'
    'K<MfTg`<Y@-2<hkr*yh0{g$rH>}GZz7Ml>VS-1~v!POHQAsG4?6*TUCnTrpZ9;rPceWm!#Ds4k+0$QH`5^<FvT;9b+axc=+;g3b06q*S_'
    'CH6D;?+h$umD8KhwtM`c&}Lj57AV~3X&e9{3zqx%E{U+qPw(FvcOl#i^fYvMm}c8D48oZREBuF;D8e+#Z9?Y40W}tsg-5wWvehaq<QQQ!'
    '5Q5q~Y(M+M#m#ZtNDo{+{essn;Rq<|3lnn1i?U4_OL87$$IyO3K^UT?9pap=5!^gehu642S8{E{8*3BE_ltMI0Buzy)H6Xhr(TIGNkU>%'
    'nGJ=mR~a%J3SF-<WO9X$XYP?}N`Fb6?k6T!`Hm)mbtn~rt4v<wAX07MK2Fi|ighSejN^E2vjg{;9oT9IR`Uz2=YB!0fiVV7LpUu2K1g1S'
    '(V(Mu&;4X^X1XY`;ja4cjKKJKx1w9uvHwtnh9G>mTD3-HmlfCdA6XUtpMCm~_d>U;{k)g+S*am9n*Mnyc}<TPgI#S>e8^MWJK~??ADK)Q'
    'U}kg(!T+_9yvoqQK;rOc-Fk6c17l}!#S48EW~pK6flV#2%YaP;*5cBTJdl+=@)P60w1SCvI%1PNFa@l5Y2XF;U4c5L3(rZUJ0_^$H+fu2'
    'VbgnUPUvYx14fF>)&t(gA>KlV7s3k7_^@2Y*{8$H+3)EFE+Ig%3qp8Wa-d=_UR6b<Rc22O)-9|uX=;Dn!se3m-G~Q&fW|5xc!ihaAw8Y)'
    'LqS8EXH%@VViLCDg6;IL(ql-syHMn@zPmW5K^z5yvE#OST9*S9#w}C+_%-k@AXjA`ctbB1`e)enhAhx}ZCd`kQ52?BfmWnpd*DT1l>|gG'
    'BG_k7f4SM+f?}1!L6^di(jDyL^yx3_nBQvK@h4LC9Zq_CWPJ6<U2NPX%RF1hyVcXrsNsOXPPawbYo_F@)Y{Ibd^G-1sOwnZg}33QZS`Zr'
    '*{Dn4vQL!yU>px`FQpCUU=yKv)K8#VFR<KyIYUcGeaLS^eivFy?_C&y=ORvFQzkcAy41qiFSv$-t(X;5;^8+!^ro7h`t!A5&c2I*$@G8g'
    '^>ANE$x`a9-=L?5Esz)={5)NSj7>w`Z4d$Ky+AvSrs75(J|(%%9vC^^S(=D1eXPga97w&Cy<5&7xaE3*0+q4yc#Cxunv4k?BJ)~_yVvYZ'
    'U+@aQ6$>}J8jI<ZG7n7{E;`e6M^`Tf>Oi9AG)(F-i!f0IFqiD9W{*z$LGmQN&7|-G3517W56~q7y@%bj;i;oEl-Y!8OP?8h*t|>npu$t4'
    '6g#-{gqQHdEvULoGoR@K6WBCUt!Oq~SnK4TF~Ln9b4fvip_f8=fv<YQj8J84uK9@4TX;pJ&@>Y#TI<y?yjF6PfZ}1ob`Dn?<6w~{>m|o~'
    'L66fd53JVHW_h?`VhpsPOrV>dz~w#<+gLG?(d=W@f;Wf4PxgYhDEKd!b~*6-Z3-b0O5Kv_qW9HXg?yyK=QD1W9kLfDZbhy6@+~=dq+kTB'
    'I}p|!H(X24Wa$3e#_?E4dYX6neP~Q4#sQ=bGCSwRSXieUmft5F59<zuEtd=64@d4iZ{UoKV+o%^&Eq8~t}++zGvZ(j_PAkB2a1Ujb-0wk'
    '+jPqFtS23nIe|jDKUWysHcK@Y7FvKO5aO3Ktk3KpiTDg!j<=rtEm|r<<7mgskM|ekAllFQ_QUDH%+`Dje9jh#<%_CCaeHf+rKwiz7UW!5'
    '^8Hdgq6<)<OP^GXl!i&uI?e7Z>2-CYy*pGD3h~qFZ^SCNP{cqGiDyKFRwt>4z2=kEp&jyZ#B1YAm@gB|nx-+xS48U#K@&|oo-^fzHte_G'
    'fB63TB|GQd{{y5#{pNRjJwMMn!>CgiM?tMw9Y@*bH&B6XnJsJY0r{a5-NvwO6+}GpYD8&Rpb<q=IV2T#mY5JIH+ID?tcruH<Dns46;g0K'
    '^*L6@=g7B=%X|3J67j%_s9+FG9z5-ZOaSA>jG-WKoJ%vjsY|}av`TBsmyI{o!$L!+V%jmg6R6^0oI&!fMtVDw4szGPTIc2J5C9JM-o4G)'
    'PFM;EQ8E!cVT>ie>gX&zor!cbnT?5P$qfdI&%*m~0pT*XdfO132)f>Y2ch*UTr)O1S+TtW?nZW`KK&K0z-3SbCDstTa-TQVnE8Nd5)4ws'
    'U!%0-q~l$57aBYEfj7Ig&VlA|YZ(R#S%2j9Drm$7UdEu|mGB{-Fjc-kZhrp~UJKo3+EwP>-yAE)(hFmn#F%<vF-tHOr2?T~Yf@zJOc)s*'
    'akz=vYVS_8?VU#ZYau8)NL&p`$EZANRF|^*T;E-}-tdsCQO)0lEH#Xei=1Wz)V(-of{*IypYxcFU=CNmw5%XBFVWE;64b70?kV8;vJF%O'
    't^uyPxX&`nw`zG*8#d1y&LJ~&#hq5JI8}FbxH^T-6o~zQiD`UWvu38Lop_P+T~1w9d9rKYS#^0D$v|eE-m^tgYFE<aJX?6pJ_W9aL2fU2'
    '^SFNn#lsdV=HnGs2+OTX`9Y&$W;c>umt=IqfR;r|HvK4z=%PYEDX|LXOr*)JMw4^6j&oL)EFHvGc2(<$F;%!uwosQ^%w(d}R_}h844_1o'
    'xL6yUPP$KgYZsQxD0gQB%(V>F6%?`6hbjqyfO0A*%uNxowlcETByjmzIV&P8ZA7;aZ#8V=b*bI0z#aJ|e#JCJ@lPIV<FPN${W<~d7T{)='
    'CcO6AFc6QA``1hfhs#5s%h*uE#ku<Te^qEKisnt0@)W6NYZ&I|_%<%10M&KTMmwk3&ZXPAjdr{=5ZTmgG{uKG67(JENMP#;4rh{j!;7PT'
    '$`|39vbsERV8cb2Rosj}{#rI}njKAt&{fkL6A*ZEv#~BRdKY^9<Zi9hjQxF+MEu@;6VqzXu0=0|*<9G~tEIr4z^9+9ku@oKE>Tonf9?JV'
    'SA2w+l6jWi0@oMnf#UJQS~M&qf4>gbmz&?q6HPS6V8yP00I`tkmFY?oOq;%wcBn|CRSF}Igu(NfF(R`GIawWf2Ul<hgIV0Ryx~GNefqiB'
    ';+u^=WWmTEK!31hm5&4zOimT-L$)~kE@WM}8e7TZ=k>dojrU=jr0t1n4pLN=TSC6q(=Y5;H11wA<K)orNS;$Gf?Rx_A6^(IujY^_7K&$Y'
    '+~y-oK!#q8f^9Y^JXcKy{RNmP@$_fbdTFfTb{n-%_zN{HL~i@@s27bsmH7xa7xkTZ;>yi@dJp0;E7YTE6rr@n5Zt+a`k!(4V)G+bd^6Ke'
    'w*0jUqCcD7Gx~|Lr4?Sr%;*lc$$e-XfS>L51w9&f&}e*qc%86jUxc+YL^tvI`8D5#OWLR7^CN4vWIPXr8F~0S%F-+nTbF4Of82pc#nryc'
    'Ql|8=;p!&A=V#lSu-KK=2y2G9=iZU1erv#TCPqq+%O(7=?~>IzEJ)K0C`3$c9)gO6+D|_L2F)%qOy`CBg~hOb%$lqU^J$u52!_A$o-yrj'
    'L#JMf#&+E}UK?Na4KZBun|Qm3(^TKe<#{4$bjL*Fwri9a!?=jkY|Cwr$Lv-=O#VrXUm~oS!Sn`<q2hdp>#m8j8<gqOUpB)ioqEPu%ZTZy'
    'kgwNE=87zolHh&(5Z(%dR3g3z^CqRoDH#i{L*pFo@!c?tuN{-f6DF9syfohZ{#@HE;A^~QJ7Tt2X27gsG~G-$t?52;M#_4KxRMa<zJ5*U'
    '2AJ2*#HOh36Q&|TW4Go8%b^h%!bGM40s&B{2+<wn!o(U5!i!_s0>c7H2Fg>6%2B5vo0d6X{Art*Ns#PosBeoK8^7%YEYyTF3Pjrq7$5Fj'
    '_cgcpuKI!pL-{6Rgf6iqozp1@c}mnShPVwE;$iF4S1l?>q7vT4A!IJ!YAJoOKr})of1mvnKkD1~JP~QZBsys{i)P5y7bn-on+|~p4X0(p'
    'cbE{XZD`4>QZ6D>mP(&Z9U#)sPs)xi;o&2FJF{$ey<?W7z8wzeq=<6v3>@Gz5`?xYhT+C=6!tTk<y1U8hdaj|pWupvkJ-XuUC+0l^(|z8'
    '<Al5E7=>_X5z0mLoCj)TjbHB|z%~2A^d6?uo@I-Lo|!DpCFp!ADcBB$B1;`#H?3v=F+Zfc<^HtHKE$aL)sQL<>f;-GqMjMV@c^yN7iCQX'
    '{mR=`eXJV4*B+_RlAnybEmTaKGs&WF>F{jB8%&ZruJlYjfm<IORr8LSLzWOaqeRg0mA<@F&$A`^qk7rfMqa?X{h9{THhc9li&GsEB6)KM'
    '<3Q6p-^E?>#+i@;$Kub%b$q{LzNn7DD$Br_4LsEx<dIn(g-iXSJ53S7KCP3V9Gr(YS+>2?w{fdBplz~~EMDpxdnFg~YmiJ^j+v;?xA$u{'
    '%?I(R6L9&u&36$P7n1$Y!<sSscuXS(uL!^J$aFza2PQAQ$piec(xXpDqNDG6D7stk-lEci;8g<f#db<WcT#ILy<5hWR3P~J%WjT#t$NLR'
    'j{?`0j{=0}4=0*0p=v7oqvP`(KDl^LHaBjvJTDh$8>eE1#WmWF5xDhm%c-cu31A)f@SN+lzjqL`a_eVzb2E6^{OWuGAp=KD)<eNhzAR-e'
    '!4L`4ZO6(@B`Maqa_AVnPS#x7b%&UsHHJ5sOuQ8tv2i!Jsl2s`d6ck^b&+2rF%BlkrV8B#f{9GN;9FN6l$J`#uiHWbE~q6Pbhb*zlN)oW'
    '+>)}xG!*W+Vv}uK+A;@lT0<_U#t1iQB2s}+MjDG`ndC9%%7tmipj={58~(_$hRSMxeY6X6QGEftDxvXMMfj3X17oWe7UH~l^IKt|AlMxQ'
    'wQuz!vCCz=s4hFNAYmYEtAf0-r8eGXsjeli)n&{KSbl@N)yTmrrB8niQ&D1|X(Adaf9(R9{yteJx>Z0QDw9UiGlbt)kAP`tyn!AuQKV8z'
    'cHhL|R<K0toBYbL13vIg_9&JVtU&I*WU6Ec)%2m>u8!3M*|{c6PoAR?qfuB+k%oNARn|?suHJC_5-zk%fNu40r!hQClsyf&O~`huahKdC'
    'vF6eM5MsoXzo3iPFYuBHKQEYl%NQtSyG&9^)cwluM-CvzGrdi24<zSSePgdQu};1rm@{-*m@sWqyNc{sz;!|`Bk^R~Y!laEx2gg}R}vtQ'
    '`g>Z8CI*&<T8pMR9cL->OC%^YViX$dr+-@sAM{d6t<n%rMwyr<hMc-=L=&S*9rmGV*<*GAzxB^gYIWf)XUul!z;SCo6DEcfO#cvk&|FHM'
    '1j5_qOp+HulB<n;l=Jt@dMcy#7q8So)c=6drU|pN)+Fy>5Q@f<boRk$)BceG1f~{1!AE>|i~zEwpz~E6kxeF}N8aD9Ic)r8Ch%%`V?Zr-'
    '8q69!z+t#HfvfUz-#TGLZNIZvuBy_4byfWc3uUjV8(rAA)P%0%G-hJ2Du!P=yLH|gPRiYw85YbbP0f99>=w?ZDiJUShX+zbb%~;*A!@G|'
    'Y9Bq=+>U}9=4om5bx?S{YF-X;Zjr67J<Duf^NT-!1`Nga&N{=#Yx#9jcsl^ieyz@j$;<ND!3cMYB&zcl#F0zUPTRDboU#47!%DKJB~!g)'
    'rxN?nEL&;VZ(8zDhcBPELJoPm+p?pd?TL<tPVHBN1B0Kg9qnBy6jqcKz@RLD(Lv;iJL*NIJFu~tuV8~h3ghYmMnn2Dk0rCQ@O)`?3v^*P'
    'mF6>0;lbqcld7YZ5lk$aGpVd6thZ1IUsa?H!g3N>R*3QS6LHg5pR!rD<DSt@+Xvg>YGdd<1EaB{C12vMSRZf4*SLD>x}#;#h-1X~NMXlD'
    'LaHgJT4x)bx(3Zete-6i=8eYs%=ePOd?A@z1#OOZd`%hpSLnk5ZedQ%*ELymFlHGY)N!Xsx#%D-14m)WZQwi75tz-K4B?*H8pS5m?1ai4'
    '^cfSZUcnv}17-i%X*scS2E7Sk#k2y=sk}cJ<>MEbiV&*c6de>eZP!o#28vb`u9LNnaWH`f6pN(%UN%Qjwk2<N0|BM=T(q^G^W(RA)yWa}'
    '&a)y>Qi-@Piy&~z)by%xX555~&o$v_ZvX9aorTUd_h`4srR!N6p4tQzd3Ll}5HxaswppsW8JSc&n5`B}FT)V(3%Yl#`INE`5ni{MQ1i&v'
    '7O!RyW0-+8?ntooMfuD(zb!YQacGo9swv+XW2T%hrSN7m$s94xT8A1tlBW_f^UJu+(i^%77NgT?MQF_7(Oej(>zfcRpvE+-p}jUpInI{Z'
    '_UY#cvJ~+-z6Y<JVCRd>!2x(F_;R~|{GkJTRry$LyqFe{aKS*he8*<3dt_4#c;!o`O2NIwqI&0g26YR7&U70O{4lZ(6E2p+4B>5>t*`);'
    'arwU8JHIG^wq50Pk@jl?A>?)u{L?WW49kdCM+Bxd%o41D2@HF*0Z5LMKz4?ifSse`m0)=o7w`6^nBZsJy*7{Ft%vaS<j9}27tO0PCZ$<M'
    'aWp7*QanP203p9pNDEW`%9v1ts#`mU2J19^`sG&G1g&yCAV-DlVZXkCgFcG&E-GFdpSxTb^Zyj*gak%i^-UuiNtqUYh!@N>#waKh!YOS$'
    'vRga=XB&9aL9AevzDG}+c*l!pyw_o0O=Cy0-GtG`UkM{P0jtzE8cXr3W^0dV&5Q6gmHsBBY*0!WcFkr%ul-u`>Z>rylT_yqccvO3P=NWX'
    '10)?jv!&2U6`>hv$J^mR&81P!-!{`J`mWXNgS#*in9rhOVEj-fgLR*X4{g=X>0>a)@%v8r`)a7SUF{ad4wN1m!cidbbwGHBJzt7ifWi+)'
    '6T1U^q|L43dP<LH;6>$EPKApUZ%9UvDWlV~Twj4mcVCU$oo+#bW(1O}jc?$F-o+c5X4;N0!k4i#ztAM>c@uR=<!keixNE}0*y0Je4i()q'
    '2jPeW92&wJ0zs3i40X%XZukU<0#A!w{5UMR)KggtpSp$u=kZpTtnIm6iQ6JuG#SV9=#`wMTMZnq9|01A9U<&c{jx)^L9QH~N(fdVN6ry$'
    'ROzrQ;0XFtmMwYCy|Imd0`rAw{RaF|7#`dVEENj{ajHX9bP<`vpU7#drQ1hsV%AOOY(nDo*LE9A1cxD;>U6z?iy;XeQ1WdhU~$5PdtBcT'
    'B-%T5gc?dBuoL~O>@jUNm*OG%Q^5Bs&_b>#I63m*`q1RanHY59jPE}~9dzdB#8~Qb&mW~daQnC^b}M56b~G)X{VA{P5&dMGmxcY=K3X3x'
    'GA2|dsn=YsU{dvX1-TU!4cSdl-IO)bde{H|g<Y|3gD?>N6_3dhk%T5VQkPa$TPfY5cIrR~qO^fjKq@ixWBLJ|vUTOh>btXvAqR|+cz3qJ'
    '7oUCand&pF3kCXZ%zA7<<TEuM3>4S7Fu8#=jF?=cVFVi8j`MQSH@Ef#Ht`g#{^KB+=dAfUfA7q(-usEBC-t0mdW-Y(CJKVQJV^%E@i5{A'
    'Hx+)#;H?O*E%8{7^Ih*bP~y)cCEMy-6?3wWIE`44bXd@)rz%%k<fmZoXEpnGi;EW+nQ$fCB;mq+YmUg-0Oz3uK8s;&lXlhCe`jXCq_hcH'
    'A!>f6m{XqbP$l#Fyaqf}r$Q~2_zt?ZsSqfLsX5K4et_WJX-}+I(~bZm8=~~g0!J>aA~yPKF(Qs;H0Z{QJfy-nrg(BTaVXAtj(FOLhigd2'
    '_(>sax=E~3ldpJpq)TeG=RsExQguWf^5k7srW6O?6a{0E3J;X_$q0TNb%vin%S(}CRUU)(G&9~(&$-X4Kc70?cy88u?rtCA_C4f;S|+0n'
    'nd<{cm5+aF_PnvqG{kK=fGsJMLd6oO(am8S_yyK`R=<AJ(bWajzTXwDb%2hfMGHFGk9$HWNkqj3N!J*9Brkyy0i|f_1@yoq*X%b4D)<v&'
    'tu~P+yfsP@^b(I}^PoBbQ%{T@(_kV{^a^L*nDsuL4q8H<P1YQ1ELl;BOuW_oON?gv2(t?K(#9-s-aANYsZs!zZ>vknxE7+o3c0BB9pm=0'
    'f~a>KphSe8xP!RGg{#nf>DbUL>Hsc#`e$vL_~}l)-UDxgu3b3`zQWW=;50>rpM7vFLEH+affzl<%C!AvU%&qWM)h_g'
)


@st.cache_data(show_spinner=False)
def catalogo_rubros_v18():
    """Catálogo estático integrado en el archivo: sin consultas SQL para navegarlo."""
    return json.loads(zlib.decompress(base64.b85decode(CATALOGO_RUBROS_B85)).decode('utf-8'))


def v18_chapter_code(value):
    text=str(value or '').strip().upper()
    hit=re.search(r'(?:CAP[-\s]*)?(\d{1,2})(?!\d)',text)
    return ('CAP-'+hit.group(1).zfill(2)) if hit else ''


def v18_chapter_label(code):
    target=v18_chapter_code(code)
    for ch,title in catalogo_rubros_v18()['capitulos']:
        if ch==target:
            return f'{ch[4:]} {title}'
    return ''


def codigo_rubro_orden_numerico(codigo):
    """Clave de orden natural: capítulo, rubro y subcódigo (todos numéricos).

    Ej.: CAP-05-130 > CAP-05-129 > CAP-05-13 > CAP-05-09.
    También soporta sufijos presentes en el Excel: CAP-23-06 (11) y CAP-27-113E.
    No cambia los códigos originales ni inventa códigos que no estén en el catálogo.
    """
    codigo=str(codigo or '').strip().upper()
    patron=re.fullmatch(r'CAP-(\d+)-(\d+)(?:\s*\((\d+)\)|([A-Z]+))?',codigo)
    if patron:
        capitulo=int(patron.group(1))
        rubro=int(patron.group(2))
        subnumero=int(patron.group(3)) if patron.group(3) else 0
        sufijo=patron.group(4) or ''
        subletra=0
        for letra in sufijo:
            subletra=subletra*26+ord(letra)-ord('A')+1
        return (capitulo,rubro,subnumero,subletra,codigo)
    # Compatibilidad con posibles códigos adicionales de futuros Excel.
    numeros=tuple(int(parte) for parte in re.findall(r'\d+',codigo))
    return (*((numeros+(0,0,0,0))[:4]),codigo)


def v18_catalog_matches(chapter='', query='', limit=100):
    """Capítulos del 01 al 28 y rubros numéricamente ascendentes EN CADA capítulo.

    El usuario pidió que la navegación empiece por 1, 2, 3… en todos los capítulos.
    Por eso el catálogo general queda en orden natural: primero el capítulo 01, luego
    el 02, etc.; y dentro de cada capítulo los rubros avanzan de menor a mayor según
    el valor numérico real del código. También se respetan subcódigos como
    CAP-23-06 (11) y CAP-27-113E.
    """
    records=catalogo_rubros_v18()['rubros']
    q=normalize_template_key(query)
    code=v18_chapter_code(chapter)
    matches=[]
    for ident,description,unit in records:
        if code and not ident.startswith(code+'-'):
            continue
        if q and q not in normalize_template_key(ident+' '+description+' '+unit):
            continue
        matches.append((ident,description,unit))
    matches.sort(key=lambda item:codigo_rubro_orden_numerico(item[0]))
    return matches[:limit] if limit is not None else matches


def v18_catalog_record(code):
    for ident,description,unit in catalogo_rubros_v18()['rubros']:
        if ident==code:
            chapter=v18_chapter_label(ident)
            return {'Capitulo':chapter,'Codigo_Rubro':ident,'Descripcion':description,
                    'Area':'','Unidad':unit,'Cantidad':0.0}
    return None


def v18_activity_chapter(activity):
    """Preserva capítulo elegido. Libros previos: solo asignar si hay coincidencia inequívoca."""
    code=v18_chapter_code(activity.get('Capitulo') or activity.get('Codigo_Rubro'))
    if code:
        return v18_chapter_label(code) or str(activity.get('Capitulo') or '')
    description=normalize_template_key(activity.get('Descripcion') or '')
    if not description:
        return ''
    found={ident[:6] for ident,title,unit in catalogo_rubros_v18()['rubros']
           if normalize_template_key(title)==description}
    return v18_chapter_label(next(iter(found))) if len(found)==1 else ''


def v18_catalog_page(repo):
    """Pantalla consultiva de catálogo. No cambia datos de otros usuarios ni el Excel base."""
    page_heading('Catálogo técnico','Capítulos y rubros de Dablanc',
        'Busca por especialidad, código o descripción. Las unidades proceden del Excel que compartiste.')
    book=catalogo_rubros_v18()
    chapters=sorted(book['capitulos'],key=lambda pair:int(pair[0][4:]));items=book['rubros']
    m1,m2=st.columns(2)
    m1.metric('Capítulos',len(chapters))
    m2.metric('Rubros disponibles',len(items))
    st.caption('Fuente: Estructura de Costos Dablanc · Hojas 03, 02, 00 y 01. '
               'El catálogo no incluye cantidades ni valores económicos. '
               'Para usarlo en un Libro de Obra entra en Actividades → Añadir desde catálogo.')
    c1,c2=st.columns([1,1.6])
    with c1:
        chapter=st.selectbox('📂 Filtrar capítulo',['Todos los capítulos']+[c for c,_ in chapters],
            format_func=lambda c:c if c=='Todos los capítulos' else v18_chapter_label(c),
            key='v18_catalog_chapter')
    with c2:
        query=st.text_input('🔎 Buscar rubro o código',placeholder='Ej.: enlucido, CAP-05-23…',key='v18_catalog_search')
    prefix='' if chapter=='Todos los capítulos' else chapter
    full=v18_catalog_matches(prefix,query,limit=5000)
    if not full:
        st.info('No se encontraron rubros con esos filtros.');return
    pages=max(1,(len(full)+24)//25)
    page_key='v18_catalog_page_idx'
    # Un cambio de capítulo o de búsqueda empieza en la primera página.
    filter_signature=(prefix,normalize_template_key(query))
    filter_state_key='v18_catalog_last_filter'
    if st.session_state.get(filter_state_key)!=filter_signature:
        st.session_state[page_key]=0
        st.session_state[filter_state_key]=filter_signature
    if page_key not in st.session_state:st.session_state[page_key]=0
    page=max(0,min(int(st.session_state[page_key]),pages-1))
    st.session_state[page_key]=page
    st.caption(f'{len(full)} rubros encontrados · Capítulos 01 → 28; rubros dentro de cada capítulo de menor a mayor · Mostrando {page*25+1}–{min((page+1)*25,len(full))}')
    if pages>1:
        ca,cb,cc=st.columns([1,1.2,1])
        with ca:st.button('← Anteriores',key='v18_cat_prev',disabled=page==0,
            on_click=set_module_view,args=(page_key,max(0,page-1)),use_container_width=True)
        with cb:st.markdown(f'<div class="ab-v17-page">Página {page+1} de {pages}</div>',unsafe_allow_html=True)
        with cc:st.button('Siguientes →',key='v18_cat_next',disabled=page==pages-1,
            on_click=set_module_view,args=(page_key,min(pages-1,page+1)),use_container_width=True)
    display=full[page*25:(page+1)*25]
    styled_detail_table(['Capítulo','Código','Rubro / descripción','Unidad'],
        [[v18_chapter_label(code),code,desc,unit] for code,desc,unit in display])


def v18_catalog_activity_picker(editor):
    """Selecciona capítulo→rubro o busca por texto→capítulo sin miles de widgets."""
    with st.container(border=True):
        st.markdown('#### 📚 Añadir rubro del catálogo')
        st.caption('Capítulos del 01 al 28 y rubros de menor a mayor dentro de cada capítulo. Elige primero un capítulo o busca directamente un rubro en todo el catálogo. '
                   'Se completarán el capítulo, código, descripción y unidad; tú indicarás el área y cantidad.')
        route=st.radio('Modo de búsqueda',['Por capítulo','Por nombre o código'],horizontal=True,
            key=editor.prefix+'v18_search_mode')
        chapters=sorted(catalogo_rubros_v18()['capitulos'],key=lambda pair:int(pair[0][4:]))
        chapter=''
        if route=='Por capítulo':
            chapter=st.selectbox('📂 Capítulo', [c for c,_ in chapters],
                format_func=v18_chapter_label,key=editor.prefix+'v18_search_chapter')
        query=st.text_input('🔎 Buscar rubro',placeholder='Descripción o código del rubro…',
            key=editor.prefix+'v18_search_query')
        if route=='Por nombre o código' and len(query.strip())<2:
            st.info('Escribe al menos dos caracteres para buscar entre los rubros.');return
        found=v18_catalog_matches(chapter,query,limit=125)
        if not found:
            st.info('No se encontraron coincidencias. También puedes añadir un rubro manual debajo.');return
        if len(found)>=125:
            st.caption('Mostrando las primeras 125 coincidencias. Escribe más palabras para afinar la búsqueda.')
        choices=[item[0] for item in found]
        display={code:(desc,unit) for code,desc,unit in found}
        chosen=st.selectbox('Seleccionar rubro',choices,index=None,
            placeholder='Elige un rubro…',
            format_func=lambda code:f'{code} · {display[code][0]}  ({display[code][1]})',
            key=editor.prefix+'v18_search_result')
        if st.button('➕ Agregar rubro seleccionado',type='primary',
                     disabled=not chosen,use_container_width=True,
                     key=editor.prefix+'v18_add_catalog'):
            item=v18_catalog_record(chosen)
            if item:
                editor.data.setdefault('actividades',{})[row_id()]=item
                editor.persist()
                rerun_selected_module()



def v19_aplicar_rubro_seleccionado(editor, rid, codigo, session_key=None):
    """Autorrelleno no destructivo de un rubro: conserva área y cantidad manuales."""
    catalog=v18_catalog_record(codigo)
    if not catalog:
        return
    current=editor.data.setdefault('actividades',{}).get(rid)
    if not isinstance(current,dict):
        return
    for key in ('Capitulo','Codigo_Rubro','Descripcion','Unidad'):
        current[key]=catalog[key]
        # Los campos normales pueden haber estado montados en una ejecución previa.
        st.session_state.pop(editor.prefix+'actividades:'+str(rid)+':'+key,None)
    editor.persist()


def _v19_9_on_choose_rubro(editor, rid, key, other_key):
    """El código y la descripción son dos formas de elegir el MISMO rubro."""
    code=st.session_state.get(key)
    if code:
        v19_aplicar_rubro_seleccionado(editor,rid,code)
        st.session_state.pop(other_key,None)


def _v19_9_manual_mode(editor, rid, key_code, key_desc):
    item=editor.data.setdefault('actividades',{}).get(rid,{})
    if st.session_state.get(editor.prefix+'v19_9_manual:'+str(rid)):
        item['Codigo_Rubro']=''
        st.session_state.pop(key_code,None)
        st.session_state.pop(key_desc,None)
        editor.persist()


def _v20_on_choose_rubro(editor, rid, selection_key):
    """Un único selector sincroniza código, descripción y unidad del rubro."""
    code=st.session_state.get(selection_key)
    if code:
        # No alterar los valores manuales de área ni cantidad ejecutada.
        v19_aplicar_rubro_seleccionado(editor,rid,code)
    else:
        row=editor.data.setdefault('actividades',{}).get(rid,{})
        if isinstance(row,dict):
            for field in ('Codigo_Rubro','Descripcion','Unidad'):
                row[field]=''
            editor.persist()


def formato_rubro_unificado(codigo, catalogo):
    """Etiqueta buscable: código, descripción y unidad del rubro."""
    descripcion,unidad=catalogo[codigo]
    return f'{codigo} · {descripcion} ({unidad})'


def v19_catalogo_en_rubro(editor, rid, item):
    """Selección unificada por código o descripción en el Libro de Obra.

    Cada entrada contiene ambos textos y su unidad. La búsqueda del selectbox
    de Streamlit puede coincidir con cualquier fragmento del elemento visible.
    Conserva el código como identidad interna para no modificar los borradores.
    """
    chapters=['']+[v18_chapter_label(ch) for ch,_ in sorted(
        catalogo_rubros_v18()['capitulos'],key=lambda pair:int(pair[0][4:]))]
    if item.get('Codigo_Rubro') and not item.get('Capitulo'):
        item['Capitulo']=v18_chapter_label(item['Codigo_Rubro'])
    chapter=editor.field('selectbox',('actividades',rid,'Capitulo'),
        'Capítulo', '', options=chapters)

    manual_key=editor.prefix+'v19_9_manual:'+str(rid)
    manual_default=bool(item.get('Descripcion') and not item.get('Codigo_Rubro'))
    manual=st.checkbox('Actividad manual fuera del catálogo', value=manual_default,
                       key=manual_key,
                       on_change=_v19_9_manual_mode,
                       args=(editor,rid,
                             editor.prefix+'v19_9_code:'+str(rid),
                             editor.prefix+'v19_9_desc:'+str(rid)))
    if manual:
        a,b=st.columns([2,1])
        with a:
            editor.field('text_input',('actividades',rid,'Descripcion'),
                'Descripción de la actividad','')
        with b:
            editor.field('text_input',('actividades',rid,'Unidad'),'Unidad','')
        return

    current_code=str(item.get('Codigo_Rubro') or '').strip()
    if current_code and v18_chapter_code(chapter)!=v18_chapter_code(current_code):
        for field in ('Codigo_Rubro','Descripcion','Unidad'):
            item[field]=''
        current_code=''
        editor.persist()
    if not chapter:
        st.caption('Selecciona primero un capítulo.')
        return

    options=v18_catalog_matches(chapter,'',limit=None)
    codes=[code for code,_,_ in options]
    lookup={code:(description,unit) for code,description,unit in options}
    # Una clave por capítulo evita selecciones persistentes de otro capítulo.
    selection_key=(editor.prefix+'v20_1_rubro_unificado:'+str(rid)+':'+
                   (v18_chapter_code(chapter) or 'ninguno'))
    selected_code=current_code if current_code in lookup else ''
    st.selectbox('Rubro (buscar por código o descripción)',codes,
        index=codes.index(selected_code) if selected_code else None,
        placeholder='Escribe el código o parte de la descripción…',
        format_func=lambda code:formato_rubro_unificado(code,lookup),
        key=selection_key,on_change=_v20_on_choose_rubro,
        args=(editor,rid,selection_key))

    chosen=lookup.get(selected_code)
    a,b,c=st.columns([1,2,0.7])
    with a:
        st.text_input('Rubro',value=selected_code or 'Sin seleccionar',
            disabled=True,key=editor.prefix+'v20_read_code:'+str(rid)+':'+selected_code)
    with b:
        st.text_input('Descripción',value=chosen[0] if chosen else '',
            disabled=True,key=editor.prefix+'v20_read_desc:'+str(rid)+':'+selected_code)
    with c:
        st.text_input('Unidad',value=chosen[1] if chosen else '',
            disabled=True,key=editor.prefix+'v20_read_unit:'+str(rid)+':'+selected_code)


@st.cache_data(ttl=60, show_spinner=False, max_entries=128)
def ultimo_libro_oficial_usuario(correo, edificio):
    """Último libro oficial guardado por ESTA cuenta y edificio.

    El resumen es de solo lectura; las actividades de otras fechas jamás se copian.
    El TTL reduce consultas de red mientras se editan campos del formulario.
    """
    if not correo or not edificio:
        return None
    rows=(database().table('inspecciones')
        .select('id,fecha,datos,hora_inicio,hora_fin')
        .eq('usuario_email', correo).eq('proyecto', edificio)
        .order('fecha', desc=True).limit(100).execute().data or [])
    for row in rows:
        if is_master_book(row):
            continue
        payload=obj(row.get('datos'), {})
        if isinstance(payload, dict):
            return {'datos':payload,
                    'hora_inicio':row.get('hora_inicio') or '',
                    'hora_fin':row.get('hora_fin') or '',
                    'fecha':row.get('fecha') or ''}
    return None


LIBRO_CAMPOS_REPETIBLES = (
    'Ubicacion','Barrio','Ciudad','Superintendente','Fiscalizador',
    'Nomina_Conteo','Rotativo_Conteo','Maquinaria_Conteo','Seguridad_Check',
    'Clima_Condicion','Clima_Obs','Disciplina_Principal',
)
LIBRO_BLOQUES_REPETIBLES = (
    'Nomina_Conteo','Rotativo_Conteo','Maquinaria_Conteo','Seguridad_Check',
)


def precargar_ultimo_libro_oficial(datos, anterior, proyecto, auto_anterior=None):
    """Aplicar recomendaciones una vez por proyecto, respetando ediciones manuales.

    Actualiza únicamente valores que aún sean predeterminados, vacíos o que
    sigan iguales a la recomendación aplicada anteriormente. Los valores
    personales editados se conservan. Nunca copia fecha, residente, actividades,
    novedades ni fotografías del día anterior.
    """
    prev=anterior or {}
    if not prev:
        return [], {}
    current=datos.setdefault('datos', {})
    values=prev.get('datos') or {}
    if not isinstance(values, dict):
        return [], {}
    baseline=official_data().get('datos', {})
    old_auto=auto_anterior if isinstance(auto_anterior, dict) else {}
    changes=[]
    applied={}
    for field in LIBRO_CAMPOS_REPETIBLES:
        if field not in values or values[field] is None:
            continue
        # La ubicación procede de la tabla de edificios. Un histórico puede
        # tener una calle anterior o mal registrada: nunca reemplazar con ella
        # la dirección sugerida del edificio recién seleccionado.
        if field=='Ubicacion' and direccion_sugerida_edificio(proyecto):
            continue
        source=values[field]
        if field in LIBRO_BLOQUES_REPETIBLES:
            if not isinstance(source, dict):
                continue
            target=current.setdefault(field,{})
            previous=old_auto.get(field,{}) if isinstance(old_auto.get(field),dict) else {}
            new_auto={}
            for label,value in source.items():
                existing=target.get(label)
                default=(baseline.get(field) or {}).get(label)
                if existing in (None,'') or existing==default or existing==previous.get(label,object()):
                    if existing!=value:
                        target[label]=copy.deepcopy(value)
                        changes.append(('datos',field,label))
                    new_auto[label]=copy.deepcopy(value)
            if new_auto:
                applied[field]=new_auto
            continue
        existing=current.get(field)
        default=baseline.get(field)
        prior_applied=old_auto.get(field,object())
        if existing in (None,'','POR CONFIRMAR') or existing==default or existing==prior_applied:
            # Ubicación personalizada del usuario no se sustituye.
            if field=='Ubicacion':
                current_address=str(existing or '').strip().casefold()
                known={x.casefold() for x in DIRECCIONES_EDIFICIOS.values()}
                if existing and current_address not in known and existing!=prior_applied:
                    continue
            if existing!=source:
                current[field]=copy.deepcopy(source)
                changes.append(('datos',field))
            applied[field]=copy.deepcopy(source)
    for field in ('hora_inicio','hora_fin'):
        source=str(prev.get(field) or '').strip()[:5]
        if not source:
            continue
        existing=str(datos.get(field) or '').strip()
        default='07:00' if field=='hora_inicio' else '16:00'
        if not existing or existing==default or existing==old_auto.get(field,object()):
            if existing!=source:
                datos[field]=source
                changes.append((field,))
            applied[field]=source
    # La hoja cambia para la siguiente jornada, pero siempre es editable.
    value=str(values.get('Hoja') or '').strip()
    if value.isdigit():
        suggestion='00053'
        existing=str(current.get('Hoja') or '').strip()
        if existing in ('','1') or existing==str(old_auto.get('Hoja','')):
            if existing!=suggestion:
                current['Hoja']=suggestion
                changes.append(('datos','Hoja'))
            applied['Hoja']=suggestion
    return changes, applied


def v17_previous_book_values(repo, project):
    """Retrocompatibilidad con consultas anteriores del mismo usuario."""
    record=ultimo_libro_oficial_usuario(repo.email, project)
    if not record:
        return {}
    payload=record['datos']
    suggested={k:copy.deepcopy(payload[k]) for k in LIBRO_CAMPOS_REPETIBLES if k in payload}
    if str(payload.get('Hoja') or '').strip().isdigit():
        suggested['Hoja']='00053'
    return suggested


def v17_linked_master_activities(repo, project, date_iso):
    """Actividades de Maestro Mayor solo con vínculo activo, edificio y día exactos."""
    if not project or not date_iso:return []
    if project not in repo.projects:return []
    links=[x for x in validated_maestro_links(repo) if x['edificio']==project]
    result=[]
    for link in links:
        rows=cached_rows('inspecciones',(
            ('usuario_email',link['maestro']),('proyecto',project),('fecha',date_iso)),
            columns='id,fecha,proyecto,residente,datos',order='id')
        for row in rows:
            if not is_master_book(row):continue
            raw=obj(row.get('datos'),{})
            activities=raw if isinstance(raw,list) else raw.get('Actividades_Maestro',[])
            if not isinstance(activities,list):continue
            for item in activities:
                if not isinstance(item,dict):continue
                title=str(item.get('Actividad') or item.get('actividad') or '').strip()
                if not title:continue
                people=item.get('Personal_A_Cargo') or item.get('personal_a_cargo') or []
                if not isinstance(people,list):people=[str(people)]
                result.append({
                    'Maestro':link['nombre'] or 'Maestro Mayor',
                    'Actividad':title,
                    'Avance':str(item.get('Cantidad') or item.get('cantidad') or ''),
                    'Piso':str(item.get('Piso') or item.get('piso') or ''),
                    'Departamento':str(item.get('Departamento') or item.get('departamento') or ''),
                    'Personal':', '.join(map(str,people)),
                    'Observaciones':str(item.get('Observaciones') or ''),
                    'Unidad':str(item.get('Unidad') or ''),
                    'Capitulo':str(item.get('Capitulo') or ''),
                    'Codigo_Rubro':str(item.get('Codigo_Rubro') or ''),
                    'Area':str(item.get('Area') or ''),
                    'Foto_B64':next(iter(list_photos(item)),''),
                })
    return result


def v17_import_master_activity(editor, activity):
    """Copia asistida: el residente mantiene el control de la cantidad final."""
    d=editor.data
    quantity=str(activity.get('Avance') or '')
    m=re.search(r'(?<!\w)(\d+(?:[.,]\d+)?)\s*(m2|m²|m3|m³|ml|m|u|unidades?)\b',quantity,re.I)
    numeric=float(m.group(1).replace(',','.')) if m else 0.0
    unit=m.group(2) if m else ''
    place=' · '.join(x for x in [activity.get('Piso'),activity.get('Departamento')] if x)
    d.setdefault('actividades',{})[row_id()]={
        'Capitulo':activity.get('Capitulo',''),'Codigo_Rubro':activity.get('Codigo_Rubro',''),
        'Descripcion':activity.get('Actividad',''),'Area':activity.get('Area') or place,
        'Unidad':activity.get('Unidad') or unit,'Cantidad':numeric,
    }
    editor.persist()
    rerun_selected_module()


OFICIO_ALIASES={"ALBAÑILES":["ALBAÑIL","ALBANIL"],"AYUDANTE":["AYUDANTE","PEON","PEÓN"],"SOLDADOR":["SOLDADOR"],"OPERADOR":["OPERADOR","MAQUINISTA"],"FIERRERO":["FIERRERO","HIERRERO"],"PINTORES":["PINTOR","PINTORES"],"MAESTRO SUPERVISOR":["MAESTRO","SUPERVISOR","CAPATAZ"],"CARPINTERO":["CARPINTERO"],"GUACHIMAN":["GUACHIMAN","GUARDIA"],"GYPSEROS":["GYPS","TABLAYESO"],"ELECTRICOS":["ELECTRIC","ELÉCTRIC"],"PLOMEROS":["PLOMER","HIDROSANITARI","FONTANER"],"ALUMINIO Y VIDRIO":["ALUMIN","VIDRI"]}


def normalizar_cargo_nomina(value):
    import unicodedata
    raw=unicodedata.normalize('NFKD', str(value or '').upper())
    return ''.join(c for c in raw if not unicodedata.combining(c))


def workforce_counts(workers):
    """Recalcula TODAS las categorías, incluidas las que ahora valen cero.

    En cargos compuestos se toma la primera especialidad reconocible y no se
    duplica al trabajador en dos oficios diferentes. No se inventan categorías.
    """
    counts={k:0 for k in OFICIOS_NOMINA_FORMATO}
    unclassified=[]
    for worker in workers:
        role=normalizar_cargo_nomina(worker.get('cargo'))
        main_role=re.split(r'\s*[/|;,]\s*',role,maxsplit=1)[0]
        matches=[name for name, aliases in OFICIO_ALIASES.items()
                 if any(normalizar_cargo_nomina(alias) in main_role for alias in aliases)]
        if not matches:
            matches=[name for name, aliases in OFICIO_ALIASES.items()
                     if any(normalizar_cargo_nomina(alias) in role for alias in aliases)]
        if matches:
            counts[matches[0]]+=1
        else:
            unclassified.append(worker.get('nombre') or 'Sin nombre')
    return counts,unclassified


def actualizar_conteo_nomina_editor(editor, repo):
    """Se llama como on_click ANTES de montar los number_input del formulario.

    Reemplaza el bloque completo y sincroniza las claves antiguas de widgets,
    evitando que st.session_state vuelva a mostrar datos anteriores.
    """
    edificio=editor.data.get('proyecto')
    if not edificio:
        editor.draft['nomina_notice']='Selecciona un edificio antes de actualizar su nómina.'
        return
    try:
        cached_rows.clear()
        workers=combined_workers(repo,edificio)
        counts,unclassified=workforce_counts(workers)
        editor.data['datos']['Nomina_Conteo']={key:counts.get(key,0) for key in OFICIOS_NOMINA_FORMATO}
        editor.data['_unclassified']=unclassified
        editor.data['_nomina_sugerida_proyecto']=edificio
        editor.data['_nomina_sugerida_conteos']=copy.deepcopy(editor.data['datos']['Nomina_Conteo'])
        for label in OFICIOS_NOMINA_FORMATO:
            key=editor.prefix+'datos:Nomina_Conteo:'+label
            # Callback previo al montaje: actualiza además el widget existente.
            st.session_state[key]=int(counts.get(label,0))
        editor.draft['nomina_notice']=(
            f'Conteo actualizado: {len(workers)} trabajadores visibles en {edificio}. '
            'Los oficios ausentes quedaron en cero.')
        editor.persist()
    except Exception as exc:
        LOG.exception('Error recalculando la nómina para el Libro de Obra')
        editor.draft['nomina_notice']='No se pudo actualizar. Conservamos los conteos anteriores. Comprueba la conexión.'


def render_official(editor,repo):
    """Formulario por pasos y referencia lateral por día del Maestro vinculado."""
    d=editor.data
    original=d.get('_original')
    d['datos']['Hoja']='00053'
    # Recuperación no destructiva de borradores creados en versiones anteriores:
    # sugerir los datos generales una única vez cuando aún estén vacíos.
    # No modificar libros guardados ni tocar valores que ya escribió el usuario.
    if not original and not d.get('_v203_general_initialized'):
        modified=False
        for key, default in VALORES_PREDETERMINADOS_LIBRO.items():
            if str(d['datos'].get(key) or '').strip() in ('', 'POR CONFIRMAR'):
                d['datos'][key]=default
                st.session_state.pop(editor.prefix+'datos:'+key,None)
                modified=True
        for key, default in (('hora_inicio','07:00'),('hora_fin','16:00')):
            if not str(d.get(key) or '').strip():
                d[key]=default
                st.session_state.pop(editor.prefix+key,None)
                modified=True
        d['_v203_general_initialized']=True
        if modified:
            editor.persist()
    st.markdown('#### 🗓️ Libro de obra — datos de la jornada')
    top1,top2=st.columns([1.35,1])
    with top1:
        editor.field('selectbox','proyecto','Edificio de obra','',
                     options=project_options(repo.user,d.get('proyecto')),
                     empty_placeholder='Seleccionar edificio')
    with top2:
        # A diferencia de la versión anterior, la fecha NUEVA no se rellena
        # automáticamente. El usuario la selecciona y se guarda como borrador.
        date_key=editor.prefix+'v17_date'
        def _date_changed():
            current=st.session_state.get(date_key)
            d['fecha']=current.isoformat() if current else ''
            editor.persist()
        initial=None
        if d.get('fecha'):
            try:initial=dt.date.fromisoformat(str(d['fecha'])[:10])
            except ValueError:initial=None
        st.date_input('Fecha del libro (selecciona una)',value=initial,
            key=date_key,on_change=_date_changed,disabled=bool(original))
    # Ubicación obligatoria: derivada únicamente del edificio seleccionado.
    previous=d.get('_v19_last_project')
    if original and str(d['datos'].get('Ubicacion') or '').strip() not in ('','POR CONFIRMAR') \
            and previous is None:
        # Primera apertura de un libro guardado: respetar la dirección histórica.
        d['_v19_last_project']=d.get('proyecto')
        previous=d.get('proyecto')
    changed=autocompletar_datos_obra(d['datos'],d.get('proyecto'),previous)
    for field in changed:
        # Se borra antes de dibujar el campo de ubicación; de lo contrario
        # Streamlit seguiría mostrando el valor anterior de session_state.
        st.session_state.pop(editor.prefix+'datos:'+field,None)
    project_changed=(previous!=d.get('proyecto'))
    if project_changed:
        d['_v19_last_project']=d.get('proyecto')
    # La ubicación NO se toma de datos obsoletos, formularios anteriores ni texto libre.
    # Cuando cambia el edificio, el texto mostrado también cambia, sin edición posible.
    expected_address=direccion_sugerida_edificio(d.get('proyecto'))
    if d['datos'].get('Ubicacion')!=expected_address:
        d['datos']['Ubicacion']=expected_address
        changed.append('Ubicacion')
    if changed or project_changed:
        editor.persist()
    if not original and d.get('proyecto') and d.get('_v20_prefill_project')!=d['proyecto']:
        try:
            previous_book=ultimo_libro_oficial_usuario(repo.email,d['proyecto'])
            updates,recommended=precargar_ultimo_libro_oficial(
                d,previous_book,d['proyecto'],d.get('_v20_prefill_auto'))
            for path in updates:
                # Los widgets con estados antiguos no deben tapar las sugerencias nuevas.
                st.session_state.pop(editor.prefix+':'.join(map(str,path)),None)
            if previous_book and isinstance(previous_book.get('datos',{}).get('Nomina_Conteo'),dict):
                # La nómina repetitiva del último libro no se sobreescribe al abrir el paso 2.
                d['_nomina_sugerida_proyecto']=d['proyecto']
                d['_nomina_sugerida_conteos']=copy.deepcopy(d['datos']['Nomina_Conteo'])
            d['_v20_prefill_auto']=recommended
            d['_v20_prefill_project']=d['proyecto']
            editor.persist()
        except Exception as exc:
            LOG.warning('No se pudieron recuperar los datos del último libro: %s',type(exc).__name__)
    # El libro anterior NO puede sobrescribir una dirección de la tabla oficial.
    expected_address=direccion_sugerida_edificio(d.get('proyecto'))
    if d['datos'].get('Ubicacion')!=expected_address:
        d['datos']['Ubicacion']=expected_address
        editor.persist()
    phases=['📋 Datos generales','👷 Nómina y subcontratos','🌤️ Clima y seguridad',
            '🏗️ Actividades y Maestro','📷 Anexos','✅ Revisar y guardar']
    phase=st.segmented_control('Paso de llenado',phases,default=phases[0],
        key=editor.prefix+'v17_phase',label_visibility='collapsed') or phases[0]
    st.caption('Los campos se guardan como borrador. Solo aparece el paso que selecciones.')

    if phase==phases[0]:
        st.markdown('#### 📋 Datos generales del proyecto')
        fields=[('Ubicacion','Ubicación de la obra'),('Barrio','Barrio'),('Ciudad','Ciudad'),
                ('Superintendente','Superintendente'),('Fiscalizador','Fiscalizador'),('Hoja','Número de hoja')]
        a,b=st.columns(2)
        for i,(k,label) in enumerate(fields):
            with (a if i%2==0 else b):
                if k=='Ubicacion':
                    st.text_input(label,value=expected_address or 'SELECCIONA UN EDIFICIO',
                        disabled=True,help='Dirección fija según el edificio, no se puede editar.')
                elif k=='Hoja':
                    st.text_input(label,value='00053',disabled=True)
                else:
                    editor.field('text_input',('datos',k),label,'')
        editor.field('text_input','residente','Residente de obra (obligatorio y manual)','')
        a,b=st.columns(2)
        with a:editor.field('text_input','hora_inicio','Entrada','',placeholder='Ej.: 07:00')
        with b:editor.field('text_input','hora_fin','Salida','',placeholder='Ej.: 16:00')
        if d.get('fecha') and d.get('proyecto'):
            matching=[r for r in repo.own('checklists')
                if str(r.get('fecha'))[:10]==str(d['fecha'])[:10] and r.get('edificio')==d['proyecto']]
            if matching:
                with st.expander('📥 Incorporar horarios del checklist',expanded=False):
                    chosen=st.selectbox('Checklist de la misma fecha',matching,
                        format_func=lambda r:f"{display_date(r['fecha'])} · {r.get('hora_inicio') or 'Sin hora'}",
                        key=editor.prefix+'checklist')
                    if st.button('Usar horarios y actividades de supervisión',key=editor.prefix+'importchk'):
                        raw=obj(chosen.get('datos'),{})
                        for item in raw.get('Supervision_Trabajos',[]) if isinstance(raw,dict) else []:
                            if item.get('Actividad'):
                                d.setdefault('actividades',{})[row_id()]={
                                    'Descripcion':item['Actividad'],'Area':'','Unidad':'','Cantidad':0.0}
                        d['hora_inicio']=str(chosen.get('hora_inicio') or '07:00')[:5]
                        d['hora_fin']=str(chosen.get('hora_fin') or '16:00')[:5]
                        d['datos']['Checklist_Origen_ID']=str(chosen['id'])
                        for field in ('hora_inicio','hora_fin'):
                            st.session_state.pop(editor.prefix+field,None)
                        editor.persist();rerun_selected_module()
        return

    if phase==phases[1]:
        st.markdown('#### 👷 Nómina fija y contratistas')
        workers=combined_workers(repo,d.get('proyecto')) if d.get('proyecto') else []
        # Los conteos arrancan en cero cuando aún no hay libro previo.
        # Una nómina existente solo se importa si el usuario lo solicita mediante
        # el botón «Actualizar conteo desde mi nómina»; de otro modo podría
        # sobrescribir una planilla que está editando manualmente.
        st.button('🔄 Actualizar conteo desde mi nómina',key=editor.prefix+'nomina',
                  on_click=actualizar_conteo_nomina_editor,args=(editor,repo))
        notice=editor.draft.get('nomina_notice')
        if notice:
            if notice.startswith('Conteo actualizado:'):st.success(notice)
            else:st.warning(notice)
        if d.get('_unclassified'):
            st.warning('Revisa manualmente los oficios de: '+', '.join(d['_unclassified']))
        a,b=st.columns(2)
        for col,field,labels in [(a,'Nomina_Conteo',OFICIOS_NOMINA_FORMATO),
                                  (b,'Rotativo_Conteo',RUROS_ROTATIVOS_FORMATO)]:
            with col:
                st.markdown('**Personal de nómina**' if field=='Nomina_Conteo' else '**Subcontratos y personal rotativo**')
                for label in labels:
                    editor.field('number_input',('datos',field,label),label,None,
                                 min_value=0,step=1,allow_empty=True)
        return

    if phase==phases[2]:
        st.markdown('#### 🌤️ Clima y maquinaria (formato oficial)')
        editor.field('selectbox',('datos','Clima_Condicion'),'Clima observado','',
            options=['SOL','NUBLADO','LLUVIA TENUE','LLUVIA INTENSA','GRANIZO'],
            empty_placeholder='Seleccionar clima')
        editor.field('text_input',('datos','Clima_Obs'),'Observaciones del clima','')
        with st.expander('🏗️ Equipos y maquinaria',expanded=False):
            a,b=st.columns(2)
            for i,label in enumerate(MAQUINARIAS_FORMATO):
                with (a if i%2==0 else b):
                    editor.field('number_input',('datos','Maquinaria_Conteo',label),label,None,
                                 min_value=0,step=1,allow_empty=True)
        with st.expander('🦺 Seguridad, señalización y mitigación',expanded=False):
            labels=['Casco','Chalecos','Guantes','Gafas','Mascarilla','Auditivo','Conos','Cintas',
                'Rótulos','Vallas','Extintor','Botiquin','Polvo','Ruido','Liquidos','Cerramiento','Limpieza']
            cols=st.columns(3)
            for i,label in enumerate(labels):
                with cols[i%3]:editor.field('checkbox',('datos','Seguridad_Check',label),label,False)
        return

    if phase==phases[3]:
        st.markdown('#### 🏷️ Disciplina del encabezado principal')
        editor.field('selectbox',('datos','Disciplina_Principal'),
            'Especialidad que aparecerá en la cabecera principal del Excel',
            '', options=DISCIPLINAS_LIBRO_OFICIAL,
            empty_placeholder='Seleccionar disciplina')
        left=st.container()
        with left:
            st.markdown('#### 🏗️ Actividades y rubros')
            default={'Capitulo':'','Codigo_Rubro':'','Descripcion':'','Area':'','Unidad':'','Cantidad':0.0}
            for idx,(rid,item) in enumerate(list(editor.rows('actividades',default).items()),1):
                with st.container(border=True):
                    st.markdown(f'**Rubro {idx}**')
                    v19_catalogo_en_rubro(editor,rid,item)
                    a,b=st.columns([1.7,1])
                    with a:editor.field('text_input',('actividades',rid,'Area'),'Área o ubicación','')
                    with b:editor.field('number_input',('actividades',rid,'Cantidad'),'Cantidad ejecutada',
                                 0.0,min_value=0.0,step=0.01)
                    editor.remove('actividades',rid)
            editor.add('actividades',default,'Agregar otro rubro / actividad')
        with st.container(border=True):
            st.markdown('#### 👷 Planilla de avance del Maestro Mayor')
            st.caption('Referencia del día y edificio seleccionados. Se muestran las seis columnas de la planilla, más una miniatura cuando hay evidencia. No se copia nada automáticamente.')
            if not d.get('proyecto') or not d.get('fecha'):
                st.info('Selecciona primero el edificio y la fecha de la jornada.')
            else:
                try:
                    master_acts=v17_linked_master_activities(repo,d['proyecto'],d['fecha'])
                    if not master_acts:
                        st.info('No existe un Libro del Maestro Mayor vinculado para '+display_date(d['fecha'])+'.')
                    else:
                        st.caption(f'{len(master_acts)} actividades de esta jornada.')
                        date_raw=str(d['fecha'])[:10]
                        try:day_view=dt.date.fromisoformat(date_raw).strftime('%d/%m/%Y')
                        except ValueError:day_view=date_raw
                        master_planilla_preview(linked_master_planilla_rows(master_acts),d['proyecto'],day_view,show_thumbs=True)
                        a,b=st.columns([2.4,1],vertical_alignment='bottom')
                        with a:
                            choice=st.selectbox('Actividad para copiar al libro oficial',[None]+list(range(len(master_acts))),
                                format_func=lambda i:'Selecciona una actividad…' if i is None else
                                    f"{master_acts[i]['Actividad']} · {master_acts[i]['Avance']}",
                                key=editor.prefix+'v17_master_activity')
                        with b:
                            if st.button('➕ Copiar rubro',key=editor.prefix+'v17_copy_master',
                                disabled=choice is None,use_container_width=True):
                                v17_import_master_activity(editor,master_acts[choice])
                except Exception as exc:notice_error(exc)
        return

    if phase==phases[4]:
        render_anexos_libro(editor,repo)
        return

    # Último paso: verificación y guardado. Las novedades se escriben en Anexos.
    st.markdown('#### ✅ Revisar y guardar libro de obra')
    summary=[('Edificio',d.get('proyecto') or 'Pendiente'),
             ('Fecha',display_date(d['fecha']) if d.get('fecha') else 'Pendiente'),
             ('Residente',d.get('residente') or 'Pendiente'),
             ('Número de hoja',d['datos'].get('Hoja') or '—'),
             ('Rubros registrados',sum(bool(str(x.get('Descripcion') or '').strip()) for x in d.get('actividades',{}).values()))]
    st.dataframe(pd.DataFrame(summary,columns=['Campo','Valor']),hide_index=True,use_container_width=True)
    st.caption('Las novedades, recomendaciones y fotografías se gestionan en la pestaña 📷 Anexos.')
    if st.button('💾 Guardar libro oficial',type='primary',key=editor.prefix+'save',use_container_width=True):
        try:
            if not d.get('proyecto'):raise ValueError('Selecciona el edificio antes de guardar.')
            if not d.get('fecha'):raise ValueError('Selecciona la fecha del libro antes de guardar.')
            day=dt.date.fromisoformat(str(d['fecha'])[:10])
            time_calculation(d['hora_inicio'],d['hora_fin'])
            if not str(d.get('residente') or '').strip():
                raise ValueError('Escribe el residente de obra antes de guardar.')
            activities=[x for x in d.get('actividades',{}).values() if str(x.get('Descripcion') or '').strip()]
            if not activities:raise ValueError('Registra al menos un rubro en Actividades y Maestro.')
            payload=copy.deepcopy(d['datos'])
            payload['Hoja']='00053'
            chosen_discipline=str(payload.get('Disciplina_Principal') or '').upper().strip()
            if chosen_discipline not in DISCIPLINAS_LIBRO_OFICIAL:
                raise ValueError('Selecciona una disciplina en Actividades y Maestro antes de guardar.')
            payload['Disciplina_Principal']=chosen_discipline
            for field in ('Ubicacion','Barrio','Ciudad','Superintendente','Fiscalizador'):
                payload[field]=str(payload.get(field) or '').strip().upper()
            # Esta marca es solamente una sugerencia de la interfaz, no un campo del Excel.
            payload.pop('_Ubicacion_Auto',None)
            payload['Actividades_Ejecutadas']=copy.deepcopy(activities)
            payload['Tipo_Registro']='Libro_Obra_Oficial'
            save_editor(editor,repo,'inspecciones',{
                'fecha':d['fecha'],
                'dia':['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'][day.weekday()],
                'proyecto':d['proyecto'],'residente':d['residente'],'frente':payload.get('Ubicacion') or '',
                'clima':payload.get('Clima_Condicion') or '',
                'hora_inicio':d['hora_inicio'],'hora_fin':d['hora_fin'],'datos':payload,
            })
        except Exception as exc:notice_error(exc)


def v18_8_catalog_picker_maestro(editor):
    """Selector de rubros del mismo catálogo que utiliza el Libro de Obra oficial."""
    with st.container(border=True):
        st.markdown('#### 📚 Seleccionar rubro del catálogo')
        st.caption('Selecciona por capítulo o busca el nombre/código. Se rellenan rubro, unidad y capítulo; el avance y los responsables los indicas tú.')
        mode=st.radio('Buscar rubro',['Por capítulo','Por nombre o código'],horizontal=True,
            key=editor.prefix+'v188_catalog_mode')
        chapter=''
        if mode=='Por capítulo':
            chapters=sorted(catalogo_rubros_v18()['capitulos'],key=lambda ch:int(ch[0][4:]))
            chapter=st.selectbox('Capítulo',[ch for ch,_ in chapters],
                format_func=v18_chapter_label,key=editor.prefix+'v188_catalog_chapter')
        search=st.text_input('Buscar código o descripción',placeholder='Ej.: enlucido, CAP-05-01',
            key=editor.prefix+'v188_catalog_search')
        if mode=='Por nombre o código' and len(search.strip())<2:
            st.info('Escribe dos o más caracteres para buscar en todo el catálogo.')
            return
        found=v18_catalog_matches(chapter,search,limit=125)
        if not found:
            st.info('Sin coincidencias. Puedes escribir la actividad manualmente debajo.')
            return
        if len(found)>=125:
            st.caption('Se muestran hasta 125 coincidencias; afina tu búsqueda para localizar el rubro.')
        descriptions={code:(desc,unit) for code,desc,unit in found}
        code=st.selectbox('Rubro del catálogo',[row[0] for row in found],index=None,
            placeholder='Elige un rubro',key=editor.prefix+'v188_catalog_result',
            format_func=lambda c:f'{c} · {descriptions[c][0]} ({descriptions[c][1]})')
        if st.button('➕ Añadir este rubro al avance',disabled=not code,
            key=editor.prefix+'v188_catalog_add',use_container_width=True):
            rubric=v18_catalog_record(code)
            if rubric:
                item={
                    'Actividad':rubric['Descripcion'],'Capitulo':rubric['Capitulo'],
                    'Codigo_Rubro':rubric['Codigo_Rubro'],'Unidad':rubric['Unidad'],
                    'Cantidad':'','Area':'','Piso':'','Departamento':'',
                    'Personal_A_Cargo':[],'Personal_Manual':'','Fotos':[],
                    'Observaciones':'',
                }
                editor.data.setdefault('actividades',{})[row_id()]=item
                editor.persist()
                rerun_selected_module()


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
        existing_photos = value.get('Fotos')
        if not isinstance(existing_photos, list):
            existing_photos = [value['Foto_B64']] if value.get('Foto_B64') else []
        value['Fotos'] = [p for p in existing_photos if isinstance(p,str) and p]
        normalized[row_id()]=value
    extras=payload if isinstance(payload,dict) else {}
    return {"fecha":str(record.get("fecha") or now_ec().date()),"proyecto":record.get("proyecto") or "",
            "actividades":normalized,"_extra":extras}


def render_maestro(editor, repo):
    d = editor.data
    a, b, c = st.columns([1.4, 2.1, 1.6])
    with a:
        editor.field("date_input", "fecha", "Fecha de trabajo:*", str(now_ec().date()), disabled=bool(d.get("_original")))
    with b:
        editor.field("selectbox", "proyecto", "Proyecto / Edificio asignado:*", "",
                     options=project_options(repo.user, d.get("proyecto")),
                     empty_placeholder="Seleccionar edificio")
    with c:
        st.text_input("Maestro responsable:", value=full_name(repo.user), disabled=True, key=editor.prefix + "maestro_visible")
    workers = [f"{r['nombre']} ({r['cargo']})" for r in combined_workers(repo,d.get("proyecto"))]
    st.markdown(f"#### 🔨 Actividades realizadas, personal y metrajes ({len(d['actividades']) or 1} registro/s)")
    st.caption("Cada actividad nueva necesita una fotografía. La planilla de avance tendrá: departamento, área, actividad, unidad, cantidad y responsable.")
    v18_8_catalog_picker_maestro(editor)
    default = {"Actividad": "", "Cantidad": "", "Unidad":"", "Area":"", "Capitulo":"", "Codigo_Rubro":"", "Personal_A_Cargo": [], "Piso": "", "Departamento": "", "Observaciones": "", "Personal_Manual": "", "Fotos": []}
    for index, (rid, item) in enumerate(list(editor.rows("actividades", default).items()), 1):
        st.markdown(f'<div class="banner-item-header"><span>Actividad N° {index}</span></div>', unsafe_allow_html=True)
        with st.container(border=True):
            if item.get('Codigo_Rubro'):
                st.caption('📚 '+str(item.get('Capitulo') or '')+' · Código: '+str(item['Codigo_Rubro']))
            c1, c2 = st.columns([2.7, 1.3])
            with c1:
                editor.field("text_input", ("actividades", rid, "Actividad"), "Actividad / rubro", "")
            with c2:
                editor.field("text_input", ("actividades", rid, "Unidad"), "Unidad", "")
            c3, c4, c5 = st.columns([1.25, 1.4, 1.25])
            with c3:
                editor.field("text_input", ("actividades", rid, "Cantidad"), "Cantidad ejecutada", "")
            with c4:
                editor.field("text_input", ("actividades", rid, "Area"), "Área de trabajo", "")
            with c5:
                editor.field("text_input", ("actividades", rid, "Departamento"), "Departamento", "")
            c6,c7=st.columns([2.0,1.0])
            with c6:
                editor.field("multiselect", ("actividades", rid, "Personal_A_Cargo"), "Personal responsable", [], options=workers)
            with c7:
                editor.field("text_input", ("actividades", rid, "Piso"), "Piso (opcional)", "")
            d1, d2 = st.columns([3, 1])
            with d1:
                editor.field("text_input", ("actividades", rid, "Personal_Manual"), "Personal adicional (separado por comas)", "")
            with d2:
                editor.remove("actividades", rid)
            with st.expander("📷 Fotografía de esta actividad", expanded=not bool(item.get('Fotos'))):
                editor.activity_photo(("actividades", rid, "Fotos"))
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
            # En libros nuevos cada actividad debe contar con evidencia
            # independiente. Los libros históricos sin fotos pueden editarse
            # sin perder datos ni imponerles una obligación retroactiva.
            if not d.get('_original'):
                missing_photos = [i for i,act in enumerate(acts,1) if not act.get('Fotos')]
                if missing_photos:
                    raise ValueError('Adjunta una fotografía a cada actividad antes de guardar. Pendiente en actividad(es): '
                                     + ', '.join(map(str,missing_photos)) + '.')
            payload = copy.deepcopy(d.get("_extra", {}))
            # Los requerimientos NO se crean ni actualizan desde los libros de obra.
            # Los registros antiguos permanecen intactos dentro de _extra en ediciones.
            payload.update({"Tipo_Registro": "Libro_Obra_Maestro", "Actividades_Maestro": acts,
                            "Edificio":d["proyecto"],
                            "Maestro":full_name(repo.user), "Fecha":d["fecha"]})
            day = dt.date.fromisoformat(d["fecha"])
            save_editor(editor, repo, "inspecciones", {
                "fecha": d["fecha"], "dia": ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"][day.weekday()],
                "proyecto": d["proyecto"], "residente": full_name(repo.user), "frente": "Reporte de Maestro Mayor",
                "clima": "N/A", "hora_inicio": "07:00", "hora_fin": "16:00", "datos": payload,
            })
        except Exception as exc:
            notice_error(exc)




def rend_data(record=None):
    """Editor compacto v19.3, sin inferir horarios ni unidades de registros antiguos."""
    record=record or {}
    payload=rendimiento_payload(record) if record else {}
    workers=payload.get('trabajadores') or ([{'nombre':record.get('trabajador',''),'cargo':record.get('cargo_obrero','')}] if record else [])
    stored=str(payload.get('codigo_rubro') or record.get('rubro') or '')
    code=stored if v18_catalog_record(stored) else ''
    # Un nombre histórico solo se vincula a un código si hay coincidencia única.
    old_description=str(payload.get('descripcion_rubro') or record.get('rubro') or '')
    if not code and old_description:
        candidates=[r[0] for r in catalogo_rubros_v18()['rubros'] if r[1].strip().casefold()==old_description.strip().casefold()]
        if len(candidates)==1:code=candidates[0]
    found=v18_catalog_record(code) if code else None
    data={'fecha':str(record.get('fecha') or now_ec().date()),
        'edificio':payload.get('edificio') or record.get('edificio') or '',
        'piso':payload.get('piso') or record.get('piso') or '',
        'area':payload.get('area') or payload.get('piso') or record.get('piso') or '',
        'codigo_rubro':code,
        'descripcion_rubro':payload.get('descripcion_rubro') or (found['Descripcion'] if found else old_description),
        'unidad':record.get('unidad') or (found['Unidad'] if found else ''),
        'hora_inicio':payload.get('hora_inicio',''),'hora_fin':payload.get('hora_fin',''),
        'hora_muerta':payload.get('hora_muerta',''),
        'lunch':bool(payload.get('lunch')),'almuerzo':bool(payload.get('almuerzo')),
        'overnight':bool(payload.get('overnight')),'comentarios':payload.get('comentarios',''),
        'trabajadores':{row_id():{'nombre':str(v.get('nombre') or ''),'cargo':str(v.get('cargo') or '')}
                        for v in workers if isinstance(v,dict) and str(v.get('nombre') or '').strip()},
        '_extra':payload,'_original':record if record else None}
    for period in ('manana','mediodia','tarde'):
        data['avance_'+period]=number(payload.get('avance_'+period))
        data['fotos_'+period]=payload.get('fotos_'+period) or ([payload['foto_'+period]] if payload.get('foto_'+period) else [])
    return data


def _rend_select_catalog(editor):
    """Un solo buscador por código o descripción, con filtro opcional por capítulo."""
    d=editor.data
    chapter_options=['Todos los capítulos']+[code for code,_ in sorted(
        catalogo_rubros_v18()['capitulos'],key=lambda p:int(p[0][4:]))]
    chapter=st.selectbox('Capítulo (filtro)',chapter_options,
        format_func=lambda ch:ch if ch=='Todos los capítulos' else v18_chapter_label(ch),
        key=editor.prefix+'rend_chapter')
    chosen_code=str(d.get('codigo_rubro') or '')
    candidates=v18_catalog_matches('' if chapter=='Todos los capítulos' else chapter,
                                    '',limit=None)
    choices=[code for code,_,_ in candidates]
    lookup={code:(description,unit) for code,description,unit in candidates}
    # Conservar el rubro persistido al cambiar de filtro: no sustituirlo hasta
    # que la persona seleccione expresamente otro elemento.
    selected_key=editor.prefix+'rend_unified:'+chapter
    def _apply():
        code=st.session_state.get(selected_key)
        if not code:return
        row=v18_catalog_record(code)
        if row:
            d['codigo_rubro']=code
            d['descripcion_rubro']=row['Descripcion']
            d['unidad']=row['Unidad']
            editor.persist()
    initial=choices.index(chosen_code) if chosen_code in choices else None
    st.selectbox('Rubro (buscar por código o descripción)',choices,index=initial,
        placeholder='Escribe el código o parte de la descripción…',
        format_func=lambda code:formato_rubro_unificado(code,lookup),
        key=selected_key,on_change=_apply)
    row=v18_catalog_record(chosen_code) if chosen_code else None
    a,b,c=st.columns([1,2,0.8])
    with a:st.text_input('Rubro',value=chosen_code or 'Sin seleccionar',disabled=True)
    with b:st.text_input('Descripción',value=(row['Descripcion'] if row else d.get('descripcion_rubro') or ''),disabled=True)
    with c:st.text_input('Unidad',value=(row['Unidad'] if row else d.get('unidad') or ''),disabled=True)
    if not chosen_code and d.get('_original'):
        st.caption('Registro antiguo sin código confirmado. Selecciona un rubro del catálogo antes de modificarlo.')


def render_rend(editor,repo):
    """Formulario de rendimiento diario. Producción por hora efectiva; sin HH ni metas."""
    d=editor.data
    page_heading('Control de obra','Registro de rendimiento',
        'Registra tiempos, descuentos, avance y evidencias de mañana, mediodía y tarde.')
    a,b,c=st.columns([1.1,1.3,1])
    with a:editor.field('selectbox','edificio','Edificio','',
        options=['']+project_options(repo.user,d.get('edificio')),
        empty_placeholder='Seleccionar edificio')
    with b:editor.field('text_input','area','Área / ubicación','')
    with c:editor.field('date_input','fecha','Fecha de trabajo',str(now_ec().date()))
    st.markdown('#### 📚 Rubro del catálogo')
    _rend_select_catalog(editor)
    st.markdown('#### 👷 Personal responsable')
    st.caption('Registra al menos un encargado y su cargo. Ambos son obligatorios.')
    default={'nombre':'','cargo':''}
    for rid,worker in list(editor.rows('trabajadores',default).items()):
        a,b=st.columns(2)
        with a:editor.field('text_input',('trabajadores',rid,'nombre'),'Encargado *','')
        with b:editor.field('text_input',('trabajadores',rid,'cargo'),'Cargo *','')
        editor.remove('trabajadores',rid)
    editor.add('trabajadores',default,'Añadir encargado')
    st.markdown('#### 🕒 Jornada')
    a,b,c=st.columns(3)
    with a:editor.field('text_input','hora_inicio','Inicio','',placeholder='Ej.: 7, 7:00 o 07:00')
    with b:editor.field('text_input','hora_fin','Fin','',placeholder='Ej.: 16, 16:00 o 4 pm')
    with c:editor.field('text_input','hora_muerta','Tiempo muerto','',placeholder='Ej.: 0:30, 30 min o 0.5')
    st.caption('Horarios flexibles: también puedes escribir 7:30 o 4 pm. En tiempo muerto, 1 significa una hora y 0.5 media hora.')
    a,b,c=st.columns(3)
    with a:editor.field('checkbox','lunch','Lunch (15 minutos)',False)
    with b:editor.field('checkbox','almuerzo','Almuerzo (60 minutos)',False)
    with c:editor.field('checkbox','overnight','Termina al día siguiente',False)
    st.markdown('#### 📸 Avances y evidencias')
    period=st.segmented_control('Periodo',['🌅 Mañana','☀️ Mediodía','🌇 Tarde'],
        default='🌅 Mañana',key=editor.prefix+'rend_period') or '🌅 Mañana'
    period_id,period_name={'🌅 Mañana':('manana','mañana'),'☀️ Mediodía':('mediodia','mediodía'),
                           '🌇 Tarde':('tarde','tarde')}[period]
    with st.container(border=True):
        editor.field('number_input','avance_'+period_id,'Cantidad ejecutada en '+period_name,
                     0.0,min_value=0.0,step=.01,format='%.2f')
        editor.photos(('fotos_'+period_id,),'Fotografía de '+period_name)
    editor.field('text_area','comentarios','Observaciones (opcional)','')
    error=None
    try:
        hours=time_calculation(d['hora_inicio'],d['hora_fin'],d['hora_muerta'],
                               d['lunch'],d['almuerzo'],d['overnight'])
        progress=sum(number(d.get('avance_'+p)) for p in ('manana','mediodia','tarde'))
        unit=d.get('unidad') or ''
        a,b,c=st.columns(3)
        a.metric('Tiempo efectivo',f'{hours:.2f} h')
        b.metric('Avance total',f'{progress:.2f} {unit}')
        c.metric('Rendimiento',f'{progress/hours:.3f} {unit}/h' if hours>0 else 'Sin calcular')
    except ValueError as exc:
        error=exc
        st.warning(str(exc))
        hours=progress=0.0
    if st.button('💾 Guardar rendimiento',type='primary',key=editor.prefix+'rend_save'):
        try:
            if not d.get('edificio'):raise ValueError('Selecciona un edificio.')
            if not d.get('codigo_rubro') or not v18_catalog_record(d['codigo_rubro']):
                raise ValueError('Selecciona un rubro válido del catálogo.')
            if not d['hora_inicio'] or not d['hora_fin']:
                raise ValueError('Debes ingresar hora de inicio y fin.')
            responsible=list(d.get('trabajadores',{}).values())
            if not any(str(w.get('nombre') or '').strip() and str(w.get('cargo') or '').strip() for w in responsible):
                raise ValueError('Agrega al menos un encargado con su cargo.')
            if any(bool(str(w.get('nombre') or '').strip()) != bool(str(w.get('cargo') or '').strip()) for w in responsible):
                raise ValueError('Cada encargado registrado debe tener nombre y cargo.')
            if error:raise error
            if hours<=0:raise ValueError('El tiempo efectivo debe ser mayor que cero.')
            # El formulario acepta formatos libres, pero SQL guarda HH:MM consistente.
            start_saved=normalized_clock(d['hora_inicio'])
            end_saved=normalized_clock(d['hora_fin'])
            dead_saved=normalized_clock(d['hora_muerta'],duration=True)
            original=d.get('_original')
            payload=copy.deepcopy(d.get('_extra') or {})
            payload.update({'_rend_v6':True,'_rend_v5':False,'_rend_v4':False,
                'codigo_rubro':d['codigo_rubro'],'descripcion_rubro':d['descripcion_rubro'],
                'area':d['area'].strip(),'edificio':d['edificio'],'piso':d['area'].strip(),
                'hora_inicio':start_saved,'hora_fin':end_saved,'hora_muerta':dead_saved,
                'lunch':d['lunch'],'almuerzo':d['almuerzo'],'overnight':d['overnight'],
                'comentarios':d.get('comentarios',''),'horas_jornada':hours,
                'trabajadores':[{'nombre':w['nombre'].strip(),'cargo':w.get('cargo','').strip()} for w in d['trabajadores'].values() if w.get('nombre','').strip()],
                'avance_manana':d['avance_manana'],'avance_mediodia':d['avance_mediodia'],
                'avance_tarde':d['avance_tarde'],'unidad_productividad':d['unidad']+'/h'})
            for period in ('manana','mediodia','tarde'):
                images=d['fotos_'+period]
                payload['fotos_'+period]=images
                payload['foto_'+period]=images[0] if images else ''
            if original:
                payload['_historico_original']={k:original.get(k) for k in ('horas_hh','rend_real','rend_teorico','esperado','estado')}
            values={'fecha':d['fecha'],'edificio':d['edificio'],'piso':d['area'].strip(),
                'rubro':d['codigo_rubro'],'unidad':d['unidad'],'trabajador':', '.join(w['nombre'] for w in payload['trabajadores']),
                'cargo_obrero':', '.join(w['cargo'] for w in payload['trabajadores']),
                'datos':payload,'avance':progress,'horas_hh':None,'rend_real':None,
                'rend_teorico':None,'esperado':None,
                'estado':'CON AVANCE' if progress>0 else 'SIN AVANCE','intervalo':'Jornada'}
            save_editor(editor,repo,'rendimientos',values)
        except Exception as exc:
            notice_error(exc)


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
                photos.extend([(f"Actividad: {item.get('Actividad',item.get('actividad',''))} · Evidencia {j + 1}",photo)
                               for j,photo in enumerate(list_photos(item))])
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
            photos.extend((f'Anexo fotográfico {i}',photo) for i,photo in enumerate(anexos_fotos_validas(data),1))
    elif table=="rendimientos":
        values=rendimiento_values(record)
        pay=values['payload']
        code=pay.get('codigo_rubro') or (record.get('rubro') if v18_catalog_record(record.get('rubro')) else '')
        description=pay.get('descripcion_rubro') or (v18_catalog_record(code) or {}).get('Descripcion') or record.get('rubro','')
        unit=record.get('unidad') or ''
        meta += [('Área',pay.get('area') or values['piso']),('Rubro',code),
                 ('Descripción',description),('Unidad',unit),
                 ('Inicio',pay.get('hora_inicio') or 'Sin registro'),
                 ('Fin',pay.get('hora_fin') or 'Sin registro'),
                 ('Tiempo muerto',pay.get('hora_muerta') or 'Sin registro'),
                 ('Lunch','15 min' if pay.get('lunch') else 'No'),
                 ('Almuerzo','60 min' if pay.get('almuerzo') else 'No'),
                 ('Rendimiento',f"{values['productividad']:.3f} {unit}/h" if values['productividad'] is not None else 'Sin verificar'),
                 ('Comentarios',pay.get('comentarios',''))]
        workers=pay.get('trabajadores') or []
        if workers:
            sections.append(('Personal asignado',['Nombre','Cargo'],
                [[w.get('nombre',''),w.get('cargo','')] for w in workers if isinstance(w,dict)]))
        sections.append(('Avances por periodo',['Mañana','Mediodía','Tarde'],[[
            f"{number(pay.get('avance_manana')):.2f} {unit}",
            f"{number(pay.get('avance_mediodia')):.2f} {unit}",
            f"{number(pay.get('avance_tarde')):.2f} {unit}"]]))
        for period,label in [('manana','Mañana'),('mediodia','Mediodía'),('tarde','Tarde')]:
            images=pay.get('fotos_'+period) or ([pay['foto_'+period]] if pay.get('foto_'+period) else [])
            photos.extend([(f'{label} · Fotografía {i+1}',photo) for i,photo in enumerate(images)])
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


def rendimiento_excel_v193(records):
    """Libro Excel con las mismas columnas de la planilla y miniaturas de evidencias."""
    from openpyxl.drawing.image import Image as WorksheetImage
    wb=openpyxl.Workbook();ws=wb.active;ws.title='Rendimientos'
    ws.append(REND_PLANILLA_HEADERS)
    fill=PatternFill('solid',fgColor='18345F')
    for cell in ws[1]:cell.fill=fill;cell.font=Font(color='FFFFFF',bold=True)
    widths=[23,19,21,17,48,35,23,10,12,12,16,14,16,20,18,18,18]
    for col,width in enumerate(widths,1):ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width=width
    ws.freeze_panes='E2';ws.auto_filter.ref=f'A1:Q{len(records)+1}'
    for n,rec in enumerate(records,2):
        row=_rend_planilla_data(rec)
        for col,key in enumerate(('Fecha','Edificio','Área','Rubro','Descripción','Encargado','Cargo','Unidad','Inicio','Fin',
                                  'Tiempo muerto','Lunch','Almuerzo','Rendimiento'),1):
            excel_cell(ws.cell(n,col),row[key])
        ws.row_dimensions[n].height=max(63,min(210,15*(1+max(row['Encargado'].count('\n'),row['Cargo'].count('\n')))))
        for col,period in ((15,'Mañana'),(16,'Mediodía'),(17,'Tarde')):
            photos=row[period]
            if photos:
                try:
                    stream=image_stream(photos[0]);
                    if stream:
                        photo=WorksheetImage(stream);photo.width=76;photo.height=54
                        ws.add_image(photo,f'{openpyxl.utils.get_column_letter(col)}{n}')
                    else:ws.cell(n,col,'No disponible')
                except Exception:ws.cell(n,col,'No disponible')
            else:ws.cell(n,col,'Sin foto')
        if n%2==0:
            for cell in ws[n]:
                if cell.fill.patternType!='solid':cell.fill=PatternFill('solid',fgColor='F1F5FB')
    buffer=io.BytesIO();wb.save(buffer);return buffer.getvalue()


def rendimiento_pdf_v193(records):
    """Planilla horizontal A3 y anexo fotográfico por registro."""
    from reportlab.lib.pagesizes import A3
    buffer=io.BytesIO()
    doc=SimpleDocTemplate(buffer,pagesize=landscape(A3),leftMargin=25,rightMargin=25,topMargin=28,bottomMargin=28)
    styles=getSampleStyleSheet()
    head=ParagraphStyle('rh',parent=styles['Normal'],fontSize=7,leading=9,textColor=colors.white,fontName='Helvetica-Bold')
    body=ParagraphStyle('rb',parent=styles['Normal'],fontSize=6.5,leading=8.5)
    title=ParagraphStyle('rt',parent=styles['Heading1'],fontSize=15,leading=19)
    story=[Paragraph('ALPHA BUILDERS · PLANILLA DE RENDIMIENTOS',title),Spacer(1,10)]
    heads=REND_PLANILLA_HEADERS[:14]
    data=[[Paragraph(html_escape(h),head) for h in heads]]
    for rec in records:
        r=_rend_planilla_data(rec)
        data.append([Paragraph(html_escape(str(r[k])).replace('\n','<br/>'),body) for k in
              ('Fecha','Edificio','Área','Rubro','Descripción','Encargado','Cargo','Unidad','Inicio','Fin','Tiempo muerto','Lunch','Almuerzo','Rendimiento')])
    available=doc.width
    factors=[.085,.075,.085,.075,.17,.135,.10,.055,.06,.06,.07,.055,.06,.10]
    scales=sum(factors)
    tbl=Table(data,colWidths=[available*x/scales for x in factors],repeatRows=1,hAlign='LEFT')
    tbl.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#18345F')),
        ('VALIGN',(0,0),(-1,-1),'TOP'),('GRID',(0,0),(-1,-1),.4,colors.HexColor('#C5D4E8')),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F1F5FB')]),
        ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5)]))
    story.append(tbl)
    hasphotos=any(any(_rend_planilla_data(rec)[p] for p in ('Mañana','Mediodía','Tarde')) for rec in records)
    if hasphotos:
        story.append(PageBreak())
        story.append(Paragraph('ANEXO · EVIDENCIAS FOTOGRÁFICAS',title))
        for rec in records:
            r=_rend_planilla_data(rec)
            for period in ('Mañana','Mediodía','Tarde'):
                for photo in r[period]:
                    raw=image_stream(photo)
                    if not raw:continue
                    try:
                        with Image.open(raw) as im:
                            im=ImageOps.exif_transpose(im);w,h=im.size
                        image=PDFImage(raw)
                        k=min(320/max(1,w),210/max(1,h))
                        image.drawWidth=w*k;image.drawHeight=h*k
                        story.extend([Paragraph(html_escape(
                            f"{r['Edificio']} · {r['Rubro']} · {period}"),styles['Heading3']),image,Spacer(1,10)])
                    except Exception:continue
    doc.build(story)
    return buffer.getvalue()


def export_excel(table,records):
    if table=='rendimientos':return rendimiento_excel_v193(records)
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
    if table=='rendimientos':return rendimiento_pdf_v193(records)
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


# ==============================================================================
# EXPORTADOR DEL LIBRO OFICIAL EN LA PLANTILLA ORIGINAL APORTADA POR EL USUARIO
# ==============================================================================
# Los bytes del XLSX se incrustan en el propio .py para preservar logo, estilos,
# celdas combinadas, dibujos, impresión y fórmulas sin añadir archivos al repo.
BOOK_TEMPLATE_B85 = (
    'P)h>@6aWSQ2mk;8AppLl<BOF6001rq000vJ2m)I}Z*FvDZggK%d2nTOT`qWSY#^cp0w|yW0ssI20000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '00000000000IXEYPQx$|yd&`sk*}z6;eZfo+5?XxAORm>>}_Jzu`TZc?ccLDr9!AQs@h8)ws&TBvYE+j*}Lik658M!abAx_1zrbZ'
    'ecOoFm;2F`sF0L5$^{RNSOJRJ^z>x%vWkEeS9@$kM-;E6MBPEJunrMC@61E$74d6oWmI~hT96mx@uk$k6L=a?#t_rVEzH%sqv~$S'
    '%il`gtQXbI_qj~f2o<BVS`iDCAKv^S8HM@WYA`{+_Y7TQOrQ*Q0MxrWCd-WJ8Hl(s6mlzP^z9!`+h@&reS1X$hx^G~`%rUbX$3nQ'
    '@r=#-havC$@58$8kw0Lv2C65O=t1@DmRvgd8Peh{ghjm@WM4CjsdcZc|IxjJzZ^zcZs$i9lKU&b13HMO5@fj^E}_5-5F)J{;IVBl'
    'phL`^N?@Le=e<3;fZx~yI!scZS%O?Y9PaxXXn<EMr*AqQ%_Ez?!QI8r;q#P2M9#Ma`wRGqhFm?03<QZRY{FsNPqty?a6GtU$mwgq'
    'Y~hwA?WbP=009600{~D<0|XQR1^@^E001EXwN)@9^Z)<=OacG^3jhcLUvgz^b1yD(Wo&aGq67jcpa22@00000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '00000000000000000022l1)#;Fbsz8Nc@M$_to8y5JF3LK}fqHRDOU-+z+*lqa+*Mzb70JR2qW=M}NJ3{L0(iJdNteA@<@ltcwy9'
    'J0rGtt%jHSadihOCUTZWajfCQ5q8z})%KZ3N_i-Dy^kuZxTv8^$=@SJ)A2;H5YH}G8VM7n9EKJ>nSolyO<CSz_^tufLR0OnhOoCc'
    'pz4#)n}64qrs)koh&fE`(wP;=b7E)Nu6)RBA@%thRa0m=X|Q6RaylY;Uu0~+^F*7!6a9QGPMpY6A~X{CGS5fHC7|_xTC_Cr%bj_|'
    'HwnY5$neR6iyx@o0RRC1{{sL}O9KQH00saE0000X0FAEv!+8S$0Ko_V01p5F0C;RKcW-iQVsCG2E_iKhtW|Aqqc{-$e$xI2#CM9#'
    'TL>y$@fK~QcDKrA+dHY!$_1xb0n8cGq^kD6@0f(N-JDLj?nQ|mdxpo)%sd{>KW^(vd!(Xcya{!DhIEZKIWJiA5bEESGsD-lPLie|'
    '6>n&$KhaMAc=6$n^9>j4d(PJy05n~wmr}NYX?A%@Yto(ZmNsBnaZwWqOYvZKt)QgnN=jv2nHEBxSrgXiJA^=-g27iSmeZ8yeN7v='
    '0}-?$66kj&Yr7Gx&QF1>iCFiok@LC*h<jEs`80s)TAc^;hlUGsUqRSg>}p#9Uk`p5!C;Lif$8X|n&pCbd?n8S+T4*oiXS5r<2NbT'
    '9_QdRB*zr=k*OLUSPAc)th@u3_X38|Z=+$5)}WODlJ`zD_rT54FV0u2qBlF~YNTyHlUh}((zS|oGAo#*MW|0;$v5=PhUnY4XBAkj'
    '$;9$>^Wq@qYoQf%MfyrELC7OKI3mkNNO^)VqDoTH5J?l>NYLRvuHV(%fSr^a6s|4k-#rr)E*9tuVug|90lDw4i7d5VRH2>(ci%b)'
    'e;Dsl+N~vT?_SB09N&TeAWV|0XiW%rM``!}*U*8P2uAT;OQFI4d3p&2_(~q398mOPe_Q8}Z~yzVA9*ph{M4|gu4On;k{Z4@^9^r`'
    'oy><cV;uj0D1;Z}ob<BXmsnwjx&!q*GF_4FXu&AxS@ClIGunR)HUBb<wjU}|)y!@f-E=QvP)lul%bJ33Lfya!n&H#i<z_J6vO<<1'
    '65Dd&((Lv=QT9**V{AEU6GX3=h5ECVq)Fu1wh_A?GMpqqhVQ4ym_|`*yY2*8Q)gggzM?d=O`tZ+S~E!favm?fXlbT>i5C&HY1LU('
    'lCiFdKpiFK1s;kxKAn>)2ZgI?aGYW|OnK0)yzJyKgSZ&5#14ukXzCa!vt7gSr<UP6wq+zvYGv*uOS9NjWm4Tg_&<8!pz&aQttgfy'
    '5ppR=zJ{04g05oHfuME)0pHhTW7m(74HTUjo*53FA|sAH$8gh`?M`r-WbW*RtO|ei&JN_8!_$<=9&V?)n8P|y<7~fs*j(*4_T_qW'
    '_ku;L^1gri-yFX}^c6kPd3JN6WAgcGd3B=iWw!i&J3G}mx{A~2bpL3vi2hn;e~w3(|Cms7D7YF2nVO?aE`9<40RR6308mQ<1QY-U'
    '00;m803iU92bAmt0RR9$0{{RT00;qiY%gDOWo&aVcW-iQVsCG2E_iKhE^=jTb0DGw0VtpV0RR910000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000IZVBYQ!)M'
    'MfZjNgVDZmGJQ~nac0p{m~ESWfN+#JOkxLHedO<JQaZ`NkZ!UuNW%9XiM!i>cdhs&9~x&BWJwBwtafE%YX$Gc%grN*z-UWsoh1bW'
    '1=#1;SG!kgk>jUenm&lAVuJ$1+;61}+E9lfaXr~MshsbSW9e(zqdsCya+9VH($6&@KhqS4Qh`5|w;+n4kIw&VyQ*q5Jv;r;k>v{('
    'GVsvG1w?_qCRXrUCy^24W&6#V_8iYbllxf8@y(VPvNajLoIeIb#1qG<734Tsgt=c6Mq{g=3^sOxIklE9a<^(I40<Zx*vHJzvoUY~'
    '!N{}N=RW`d0RR6308mQ<1QY-U00;m803iU;0kB`QaR315gaZH=0001ZY%h0ja%*#FWo2}8FLP*RWpptvcx`N)RBdnCFcAJeY5yVe'
    'caqq7M?kf<LaQ`QXtbNQZzdS9NMe?lQa0_s?=Fc0QYXkN0Y032?&Z1XGP-%mGwWHDl~l!;P}?PzC{neQ#d=Jh9;bttSe53*l4q(A'
    'V{#M~xtY9wH#(?tTWy5U7DOnjG1+Ln8#+#vZbZ&2TkS*vPgbhTxrV>x+NpM>;L946XAX1S(8;+hNE2aL_KBfZE18OWmF{y<=q5-h'
    'GOnTiY9n`*iI%5*!sWc&?stP!<vR$mkeSrSIw-O7boj6?RLK_^^!-A8o?0&@d>DK^lcv6Zl_{5LsVcS7Hbi!snqRk{IElm4F4F7%'
    'dxfFCQ;KJa13ZOdZ!3ugT`YDA=Djgk*abtn%Hduv$K-bshtu1@jR(FvO$L4(CWFLvy@4A}6F+9bT^vpSkjbdl$*i1=c6=@7LO<<h'
    'rL~g!QO)22v`U=G$mzV65(I%HSfyBv$p<?8PF*iRr`qSM6o=}}wlrVNMJ7@$pejnNXK)x3Om+)0-Oi9yaUj+&rSf^oQ3MfeNB8$L'
    'Dmw$Hdwx5|mH5PuYOgV4gTY-{D7D3#2gv4Pdv$KCm#11ji#w4)upbQf?PuKr+FgqXHSM0aGOhOum|5^j+^OuFT<Q%LMXaS*@qMPx'
    'E<HPnBG31tpnLj79X<-V-e|zUZMv_t%DWpF8|r?zJl=~cg@uFKZ4awTRR(Orua!%DalpRtm-;+36|m4mGbOdK1_1J^0icW)aM|B)'
    '{oyqJXU#Na|1%LZRVRY81^RXzhY^jDg$qEarv}2=0t4IksqcpDGB9d^aQ<IL7=4`)&R(1(tv{Tte`HfPsmFVvG1NvgEHFAUgTN23'
    'Q9`>hgV7>mn_Z=nnrWlakr{^9xul?vrq+{CdSx_vp(nwdtN#)`MlrY#ZLu`5J&(B_y)t@gM1zLGUKTbrazUfxw3kAu5eyn8n{W4;'
    'q|pvNi3fFYffr_EgNDh*`tibyaL{P6Jt<{IJ81M;3uZKfhAA|#`^b$M*`U!v8STX=Dx+2^8YXM6eZ)6qgFF)ryRgYl{qp?_00960'
    '00030|E%3xbK}UBuHnDpsH?qi?2<}#>gaaFRFMQgg5UuFMO`}DJ!4;tJu&U^p85Clfw@x1m2c|M+PvzHCo(CBdIgdT6zj)d|MJT('
    '|5*I#AAkDekN^4i|M$E9{OR4zyWjo#AAkDWulm0qZa=>J-CzFk)4SWd-+$Ks{`_CR{^RfeUjOpff16MX|NG|dPe1?g=l@y!^6SsP'
    '{7v@$_aEN<@yDO1+xeI2hAyk*ze@l4$M-+};h%r}@ee=Sif>`X?Ylo1e9vK#!!n0e4pk0y4owd095y+$Ic#(2a@ggt&*6~6F^4{f'
    'A%`)CQx5kzoO5^x@L~Ef{6P^smT-Hm5~dsa>D}GC-#wOala&xwe4I<jVUfc!hgA+$4s{Mq4(l8?IkY)!bLeu|<*?7;ki#*DK8GQP'
    'F^5xtPmlNMeSpu&kKxw;7|vnEm&Yv+0lr>-dp~}7T<tRXP{`!w?)Um}KYMzZw(0cnJ~^EE3tY}1-)zavoCRb7S(?nJm6eJ9^BCZr'
    '|IRC0a`*V9YDfcFLpG2WvW0Yz9b^wVK#q_eGC)Si$z=Y0yf>L^I-7i+V==kp;|fs<cu09gA=5S$^7Zj6d<$~>ehvb%fGi;^NCl}O'
    '4P*`3Kw8Ka(m{5RJ>&p6LVCym86hXgJ>(2|FuDA?x}U$EI=toQKW)=<51oS^e{DZzj~lk+^4N7bXM8u&u{>`2<TovB$z0JAvVv45'
    '^Y27$@+tXJpZ%9=Y{~pfvW9FREo2MnAUl)!?7cUce}N8#d*_Qk<RiZG9x~v*7=wIB&*oqK4WDo`_mDH>!ETsQ>27~ZRMWPf-hG%K'
    '9ln_)4F$PP8@kD!yDn@^5)ottsZ5fHOwx#MvS&vPZf6bIKw8Ka(m{4O!aeSvgUS40c{G{7uO2ekEzkG!Xlv4FZnCE&Cmi8D<P3Q*'
    'VKsX8^OkC+ZK`H|`1@uu|7L#=;;YG?(JX9DQVnDUsZ5e;O!RHePpPw~G!1TN4cS0i$QIH;b~wVl-9PhB`(QGkpN=N;`KgBtCiDF~'
    'nz(AR$AuG)@E&r8JeaVWFK?@6+UB3Y1WBL3Ko*cCWCf`pHKc*8Asa{w*+M$V&SY+OALK()M36i8lgu8hkGSO?GC)Si333lPLmo_6'
    '#MieKF>T9U^xS068NP)zX(NyYWC>Y8Do71!AZy45(n7Y74ze>z8VQm%668)Ziup{JKV`u!_mBZHLQar-$Qklr!bYaowr_c`x$M)t'
    'Bz@w0lRcaIW{c8HNDI=Ev?5icn$$#^e}e0H@Om-$_yfDmo;+>1uZ{cCCtbJM6QM2lb;W(zlZYMn?TymZiie&(f**M=^rV3_@)%C9'
    'xv%cI?;Q8(x$%4-XZO`ZT$DaXxSj79_X6!^8Uf}*-&edszwA?~df=Q7Bzqq6%@(;*Z3a0%4_a{5lC&aKq?*)3N=pp~aW_99*>F`G'
    'zePO~%->@6bYshXU2$Lb;I`wwy^-tE4pb+a`4F;4jU&J3o-~j~9>Zxq_QKTuo*T|aAJa$j+w3g;5ErGV^4s|y@j*_5$F~F-29`cZ'
    'oI1lI7bIwL16q<+q>5CNnn-DocJOJCb`bNb);22J=Yqt(uDB0DV&C5OxghO8(;z=&=lCPP=bki>M&2{0`Php=;-JsA;e{Z>zK^d8'
    'avE;FCCD%obwPf}uYTNwMJ`Ctf~%IK6{#ZCq$W}tq#b-3q#aip<m0O6g2cYgzMu~oL1N$D_PHSK5S|YbcWKYNVIYmXXHN667lXt>'
    'pKZenL56*=3v!yaz9q;orS(B(&zHX4V7U98v>+`>D^f+ONlm0QNIUp6NITRt$j4RB1&Mu~eZd!k#J;`lb3xi6JRc<P<(_xLKpJ_^'
    'oaSRM28n|{+lCi{4EtUe<TR;$OORp0>x0a%1m0jU{GGHQElDd<MXE_nq%=r7_%uii$6MR@!b3Lpb;W%+Lt@|F_PHSKKwXg8Gxa0C'
    '=bki>M&2{0`Php=;-JsA;e{Z>zSjjg4T0YhWEdX%AoB}PH<&qpCoM=z(u!1(YEly^4bl!i4bl!ZJwraOdhQH~eVu*57tWB_x3_&R'
    'NIL`<WIp`k_uP{P(#U(}G#`60NF4OpHoOpI*!Q|1r}^?*f(&zMA7nl?yump8J83~$l2)XORFj%WX^?jCX^^;D(Avfq&XCyG75Cu`'
    'iG6$9=Yq5YbwTEnMt;vdX&{ZfXHN667lXt>pKZenL56*=3vwE7za>Z;+^1uo4|tW0Y5MmY(1Nritw<HACN+`LAnoAOAnmx)Gvwo{'
    '=gyGW*Vz|*;S7m=d)w!Nv_nXPyu+((+@(G5hJiHlo;l6OUJMckeYOoR1R3_dF34$?{+1xa4Bnj~@9-)cqx<iAZ_(u~Cg~g!wB)K4'
    'sUp>+CQ=%t9ef(39Yh-B<ErO^#J<kH;0r-w-`@7QAng!bkl9PeBfsaKG>}H#GpG63i$UU`&$i)(Aj7^_1*wtDn}WRT%Le0j`HLD`'
    'l#YE!3(}IbB2}cC)I>^xj0c|t$yM#V>X{(9uPg4$AF{b`Z<GcZZ^{Lkzo_B&+>-{<$YVIo$6gAO8_q`02N@T=W)6A##w#_KeR+`i'
    'iyB*`<9U2XK50Q(l2)XORFj%WX^^~9!&Pnk7M~rAb6;27hal~W)qGTY+vkETT&cOsU)1n>?nwh_<guUTV=o4YgFf4a7lI7?UU!Ck'
    '<CU7}ZJpB}1*Fd*@A4Nlw#Ws^D>YoTB&|plsU|g%(ja-IhO656Ej}M4uhbOxAxK`Sv3)Mc!j+o4{6!7F=bki>Mt;wy`Php=@=A?u'
    'cp=E&aHZz9c%|m!8~%Xla!Yxl`O6$z<Pznz8?IWCR-}qllbT3rqP%v)Rc-thpHGz6Zi@SmD6ievK9{In^w5uZew%mxGAHhvzrmg~'
    'kjD5u&+j@r8GW4oQVoBTcFzsxxWT<OCBM2I7rFOjJaKh<dXFey=DaCU8_=h3^7x29uHx04%Ms|k^~nn#(vq|yRiv8KL`oC2ga51_'
    '#J~QrAGiN!eRK1}=Jvz?{qYa~TKL&)>^45mPfXiL^Lt3QT-9;aj;r=YF3xa>pVBur<oEB<_I;ZFPN^piq>;yPiZtIRdN;-os`82Z'
    '<es;2j$dhhkNZQU!XNn+FWbDSMZ3h{ev$Jli_>3}PJh&qKE(trxN1pSkt$M6Y9ggA+QI8&^z_@-xUV(3bLWQqZEM`u75C+DTjRdH'
    'k!v&@`lsaW__ivco-~j~(kaq>XWf&|k>+>EKSa8`b=&`8Y4OU;o0=<JtGUhJfOgAkHKYY;Nm`LAQcY?irOokL4Og{B&yU8ruPg4u'
    'Lo~0|*gjWSSmg=>^|motHIPQqDN<UP-TX}FxGF8oRuu{>UTS$$VTFq?$zQ^TKUBTt#TU|ov?Q%a6{#jQk<!9=@rA2eqZbO}#h2nf'
    '6vm4$w$Bw7R=L7Jy=`=bVb#c0r$}jGwlOWtC@svWP+0NW$(sr*Tuu3ia~Q9t-12G)X-Qg<DpE~qBBh1#Y6@4iMlTe`t0~2OD2!KA'
    'Y@aJEta61tQu4|pn~C*yQ?5DO)W}t*NNIDnF>TH$ZO*9BT=DYAo0=<JF!_LSATOBQ@`4FzNm`LAQcY?irOokz30Jj7FEqysCdGYd'
    'ju%X9pKC6xa?OEy+vp0zs*$Tsk<!9!V_KL|T9^?FoBn|JI<HJGx0DS}KIHQ~Ta>;UAT3Bs(u!1(YEly^DJ&j*as`B|+IiJ8S3tP0'
    'EAGn+<G#I7T9{E<7^x=>q>*%rloUofM@kBdR4A-?0pl&_un#tDfA(Dnw#XI6KR)5AC22*fNHwX6lorNI8eG-JZ}Hjj9QSp_eJG5V'
    'G;E(MEUa>cfqL7Rp2J8Z=@cm~%x*q8hs9NCVYaGJSn-O)n+hvjqxhKJjo<W90{KLPv>+`>D^f+ONlm1*FkYkJs@CZFM@jDMiu+I)'
    'uTj`OS6Eo(3Ip}F(G`YOBUhawrG?qXv@oNzFrz|YSG&ZZ`+--%@G(fZp77oFrSn_TlC&aKq?*)3N(<v92Ci!3xA^=yjF%XS`%oA!'
    'G1xv=SXku>1NFAi6^2zKSDhlIh1tfmFr&0Eqe5X<yF&26{u04mXU(7K+amWUiHimcu3C~-q>5CNnn-D3cJS%*F#IV)8^6Wp3&Xyy'
    'xDSP4-`@7Q!t6RhGJ~~M=?s=MkVeueQd*dO#k4S^v@oMWVON{`f3VT{vsdzMkt>XIf38}RR-}qllbT3rVRrCoVffpGHhzoG7lwUZ'
    'aUTl9zP;^pg>lMnt6X7NHE`8PIz>tgv)f7wGfE3HDin6LDZf5<hmVr^g?L-!3d7)k!BtDrid2zmQWGgHj8lHDYU8)~d|{Z@cg1}u'
    'j8lHw=L+ML-&VQ8uxj9{k#vfb7RD)mT$L8~<dpxz)u#NHTl(qUM}4XJO$hmQbX(*K!{C3xRZG%}RFP^@6Dch$9K^@uI{NLWA2uI8'
    'J#`&Dez(bW^hn8dbgt^SYR6T3Bi9z^`?kupg;fJrjighgv@Oo}<Epf+C+GVgrYS7mivFfokzt-YeICc$7A9df#7?e@Uyi`PM}xE^'
    'tw<HACN-orX+vsBTT(~bMVfzbdm~qmO~lia<H!76Y_>0b7Y1n{jieLlo^&QXM7sPM`A^Bc+U)*AnBIPNJ>k*NX5{Jn89rt2s`_rX'
    '<X%N$u)VZZckZULbK{CPRgvmQ^N+hJ-g@>%h&69&LuyG|Qb*eHu=Yl-Lc8D7w`qUMKOQLVOFq+%G?2zf^9S9Xj6SDdF#jl}d+s~O'
    'ed!x<@{dw_h>P4CUY_`Ppy?CZu5&KD+|p0)KBsS&`IKMljEm+<z8j^bMw(YGxpBpf6{+TJHAWg_$M?doxp6~kNn27!+VLRv@gS1V'
    '%$tvEel+=%y`Dev)_c-G8hH#SqvT_f?z!r0^f`S3b(<a5LtK=;Qhe$YfTmA;yHayuQtX?{r?{6dr!dpzc%8sj>77k6iX}I$;;K2F'
    'id6Hq8Y9;n4r0UGYDrsCN80fq_VFOn=E6Zt7j@<jRn9+vinrdA2GYo5I2ozA(pz-7;cWCmeip*3^1Iqagz01a^o~cLU-snXTU?ap'
    'hwBOpTjlbzTc0<s;;J-1+*ZxoYK&Zd*tp@%x1=qpBkgz)`*;v(e&HZ|e%TYABX7MY4WyCBa58!-KU~!~#~<JGXFvN<y(+(}U16A>'
    'w&nSK&i^46Ta=vRzZ<znTnyWn@z&G)aJ8c1ZPk&IXM#rV$NVvS2wC%{Hl&ucC3U174{IL}D-AMyi|$B2XU~X_y!D<mkVYQE$>^yd'
    '@pC!HpUd+>S{knk@@f|(rngEX$oy^Mw#Ws^3lg@<oey!vU=>%T=R@39&D&~>-1!h2H@x|lv?X<<9S>q34<gMk9E8s=e<8wK?@0q`'
    '<T0G)pT_iXoc*lsx#4W|!o#tJ@T&Z-cD3Rz%&R}AvxLuh1%gWx-))~eA7a&#t5&3nRFfLgnzSLcq%El<?II=97^C!caHRCv;Agx7'
    '!F|*73l5}_$8a)w>g<QhCuiI6LVgy)tMa?rrHs2U%>Jxd!aNMjFUQ&<U6XnIJ8aT|v?Q%a6{#jQq%~<nYDrsCN7@<bFX`g9y^pKX'
    'vtPIe{n;;jId_aVHJ|f)(m)y`&3FCD=&Ag0ZRKnmUdYcvcvXH^yUsEFxpnr>pBQ~f|52r|Dxn2eElDd<MXE^+X-(RYTGE!(k#_N('
    'NMo}*CymX>o$K(*H!MmX8hX3+ga*<WDUB`M)>E<BSDH7R?KWPBE$sXJs@Se};bVGJnPU4g|6Sxaqx2FQX+c_&R-}qllN!>Rv>~;m'
    'EvX~zNPE(ObR_kpfi#j%q<hkt^bm>v!ol6uE^1tENv)(a>(7{ygbBPq^=E^C?{-V>i5f3z*sAGUdwd-JlD#Zm#haRc-72HxH9^f)'
    '&3x<g@EFIinQtpOhU{fz%Uj>R=GZ&#+ZDfF_H2I7eTVq7NiP-T_th~jN+<hYvKJ-&Yi?;M9%l9oedJ-DUURqHbKkjm<QQFtAKw=}'
    'X)N{P|GB%XT`rkE6=OPsroR<m^3k;|a=#T{vd5_JwkZ9rh?ESojh_Fl;J&JO2-$CM&3(;lzTTSqHpQ=(od{a)+s<Dv{$6%mwJY8F'
    '*X-P|=e|SyY0b|cj*-&eim%zPct7uZ`dl;=4=VfR9eGfv*WAVT+;=V>dG`DK5Eoq%^1lODyL>Z!3Q_($@HKzIZj0RSz}M_|_`5A~'
    'zXSLih^4L4??7y16*nf21H~_x{bttO*UY#6^v`<DeVf;Oy_Wm7{BAm~+7)j-JB94I?+|}BY4P^sa>Zw-f}Z<^;vr<e(jyPy^qM>L'
    'p8L-65ZrO1Z~JBneP8YR(8q9DApK`<jnZeyqy=e7T9GPJO=?JM(uUNMwxo`<Bkf5C(vj4Y2GU46k?u)n(nF-eKYx3*i#httqRY83'
    '9qZq`&#&g#s&srrT9B5c6{#ZCq=vL6ZAdL?OX^5F(w=l69Z5ZDAdRFG>7H~ZJs5rA`=~#bF1>ts|Ar4VyWEnVPNpw<Zr<l_bG22;'
    '8|l827NjL<MXE?OsUfXN8&XT!k~-3kv?m=%M^aB3NF(WF^d%h+-@MPSxY(-XbFj`vU){f@`5upF;gZYO^c`h4@8>5|{wZAfn#;T2'
    '{rq3Q{^RfeUjOpffBTDG3+cB0`{wRXKmYLO|5^O<>(9UZP4@ox(}$~FJ{#p<K888_*Yq{a&3pVq-u7n%^Q`E*ZA>~`kd~wssUp>+'
    'hO{PaNG)kg>PS1%o^&7`Nj+&GjieLlo^&QX7=7iZicN2%FT8u}H@6dp<X`7MtNM1i<)7YtO|Fd?xlrtH2IicWT(u%qq?*)_)}#%o'
    'C2dI^X-C?V4x}TgCk>?0C^^ZTjK0oqpugv;b6hpQll39e<u#Q*$&@c`{N_-?fZK&)l!g+>g#udY?<l4xS;?vusUp>+hO{PaNG)kg'
    '>PS1%o^&7`Nj+&GjYcjMyYJIb0=dU5Ta|`lR0ySfS>iW`60S+OP>j-0><<p+)5k&wo>j3Vtw<HACN-orX+vsBTT(~bk@lnm=}78H'
    '18FqUf8u22WaQ2$;f_vLo#U!B6#JnSLMdO9_|2h&fv*e2C=JCvNhb{jwEPHKkt$M6YDjC+hSZX_q>i*B?MVmHk<^n0(rDyDvHLzf'
    'qXcrH*sA%jFCOf+bm>#kxrd(VLtU@%9OZH|dP0&sM~Re#Vw67gkd~wssUp>+hO{PaNG)kg>PS1%o^&7`Nj+&GjYerGM(G)atIkHc'
    'Y93ET50MI?lxNw$^^BsQ{qzXpLNQ813FJZnE$z0_=P0C#RFfLgnzSLcq%El<?MQplfpjGGq=7UVxlruBPtGV@bvDwFEh**k?+9>8'
    'DNlueb17jY>`JjI@%!Wj@~3F?2R$zewk+7HU{%5Df;9zO7i?3owqV<Wbp_iMY+tZL!Hxy%3pNyNe6-|4<jve^Cj2<k$0J+iPBV{x'
    '1%P<UW8&W&PnZ$Ac#P81Od#bQPk@V0pe1QVsz^1dA+1RpQcK#BI?|4`Cml#fQcoI4qmet!*nOX#W&*iT?5+6e<I&^0^ASq<z4*U5'
    '6dMMnPh5=BPy)G7Kui4<fBu@8RFP^@Lt2wIq?WWLb)+3>Pdbo}q@Fa8Mk5!B-S=rIfm|r|ruj6KK$_$7DeiSfxSwo<o0ejfmJ&#_'
    '-Q=7?TiUIq>sAG;3RV}aDcHJTn}W3k+ZL=V*sfsvf*lHWELdN#p<v^q%?+H)+{010>s>tdmiIKC$9KWwEOVU^?kA3LBc<(NUc0bW'
    '$%$u4T9GPJO=?JM(uUNMwxo`<Bkf5C(vj4Y2GVHc&NBAnNY65XTqySD^fZ*m_p2k6r$)Gc`Rkwm{L9}iSBQSwC8JNA=0-}J;WXD)'
    'xn?-c<*JHQlN!>Rv>~;mEvX~zNPE(ObR_kpfixPqW;o5|s<V-Mt{XO%W-m`obKl+!2fLp**fny`W;od8suigs)ue{BCT&P9X-n!z'
    'JJOzXARS3PX&{Y8t{J=U)AI}mySC9a6IPX)d1|ox_GUQa{lposk!yxCUane^DpE~qNNduD)RMNOj<h4~Ne9xA)RP9%XylsVjF+p<'
    'My{E#vCz!)AxXs<@0**s+@hc2MM@r<jM8REOVWx|k!n&yT9Y=Umb4{xq#bEbI*^W}o-~j~qqG^L^ocK5osH6FB5~Y1ecsZQo8jd5'
    '@<sgH7>ryqocwatid2zmQbSskHl&ucC3U17X-_(kj-;M6kVYfd3@5)_bvAO%JURIN{M6w0?dKWJfG=MpYbLC6&48A+F<G@DRiv8K'
    'kk+IPsU>Ym9cf3}lMbXKsV5Dj(a1GpKaS-23|F0vTr*G3dp|!l?|pkS9Q|ItNY)HTzqTrU-9%cEDpE~qNNduD)RMNOj<h4~Ne9xA'
    ')RP9%XylsV=$EU`My{DBN59jjNL~4{iBsUq7s;C86xdd|W;g}rs)|&T8q%7yA+@9}sUz)3d(we)B=w|$G#a^PI0fdavyp4&=_#;2'
    ')b85Ra3Fm7B3U$ePw&E3xoEIz#Z?ulCN-orX+vsBTT(~bk@lnm=}78H18FpJ(b$h8eQ@IR*EYIn;&`5y4yF$jxbm5Xli<r2$)e#T'
    '*jBk{uxiCs6{#jQq%~<nYDrsCN7|A0qyy<l>PZ7>G;+~!63kU+BmL*`jWR~1>rK-KOk6n{j)E^=B#VZlU|Z#);V77^DpE~qNNduD'
    ')RMNOj<h4~Ne9xA)RP9%Xyl^dD446xMlPBsN5Ru4aj}`Z-|O)#c_RGmpO{>3BAdx@_G_!$YbDNpxoSnKNHwV;tw|eFOWKk;(vGwz'
    '9Y{w~PZ~(0k&A@0U#>bExk#Rz{Z5~k^>QS??JCpdHnVCt{k2u@s&>Ja1zQ!YDp*~xreN!WZ3@;FY+JCdV7r3t3w9{jv0#0{hJuZc'
    'mR!{~bHlyx6S{dnzZd*$%UnrMPK5QLR#$$Ay4<4gC{i*JHcHPvq$O!Zsz^1dA+1RpQcK#BI?|4`Cml#fQcoI4qfvVHF-oK1s<Tnr'
    'Oe9W(^$AwjZpKEC{*NN_Pj<O{UH6wr-$@J7lC&aKq?*)_)}#%oC2dI^X-C?V4x}TgCk>>LbRyl8&ZLJ({3k%BkMX(gv*XJxW~Uw-'
    'RHmmM4vB4J`s|prB2}cC)R5Ms4XGt<NgZiN+LI2XBdI41q|wNodN?HJs<V+h^~85_^Us;STzyJ>xy39RzCqYlxyLKML71ynq>5CN'
    '8q%7yA+@9}sUz)3d(we)B=w|$G#a^RI3?z)vyqGD$tm%dt51n9x0p4<_X^u8*9_k)%vCE=MXE^+X-(RYTGE!(k#?j#=|DP?deT4|'
    'ja)OF5_8qr$TeecA1*w{x%!ZJ`mDU{@0}v04^A8s+bZ|q#33<PRiv8Kkk+IPsU>Ym9cf3}lMbXKsV5Dj(a1H!Au(5-ja)NNzD4-U'
    ')n~+)Tg=WgHm6LR;f&Zex@I^d=BkQRlN!>Rv>~;mEvX~zNPE(ObR_kpfixPqW;i3}s<V-6=E-*p>jRFi{Un$p;>#EDBVr@>*o1lT'
    'lB-svid2&t(wej(wWKYnBkf3g(t&g&^`wC`8o6dTBIc^Ik!$A35%KhaRabuD#CHo{zDU*##={F+<<2u$wc@IZRFfLgnzSLcq%El<'
    '?MQplfpjGGq=7UVxn?*q=Bl%iYbI>Gd0%+7ef75sPoK(_KcV5k*jBj{4c{)zRVz|Os!0uLP1=xJ(w5YbcBDP&Ksu6o(m)!GTr_;U'
    'Fjt+8Tr^J(ho_HEyY7kRa*O^6O{C-rjZxYRX-Qg<DpE~qNNduD)RMNOj<h4~Ne9xA)RP9%Xp}Z%ln#fv>THxY6NxXJrVrD*ax;9p'
    '@a2p62PY%f4BsxyRVz|Os!0uLP1=xJ(w5YbcBDP&Ksu6o(m)!GTr+&TFjt+8Tr*F;U3mI1#49($w+mmsNY)G|!L}-WLPJ`SDpE~q'
    'NNduD)RMNOj<h4~Ne9xA)RP9%XylsVB$%ttMy{DB-!80AS^WQfpTy-RvIiqher=UIv2gOsRVz|Os!0uLP1=xJ(w5YbcBDP&Ksu6o'
    '(m)!GTqK<Qa@E<$Me^k2SD*TE?Gp>%AAI>D*@=bk54KhA#KQLnb5%vENeyXD+K^h(mei4Uq&?|CI+A+QKpKr)Go1W#)!E23^W@}L'
    'AEkQjW;pr1e37gfPJV5bYlf3wuBu2isUfXN8&XT!k~-3kv?m=%M^aB3NTZQ!hLc~eIvcrWo}B#Z3G}s_;k$z`UnFaWlV4lqn&IS^'
    't141WYDjC+hSZX_q>i*B?MVmHk<^n0(rDzG;pCUA&PJ}8CnvvOuRi&`++y~8#-@+yi)+3)*fzRm_~u})sz^1dA+1RpQcK#BI?|4`'
    'Cml#fQcoI4qmgTdlV7en8@Xnlocw;h`sDX=i&-;#uduCh&w=?~VXj({DpE~qNNduD)RMNOj<h4~Ne9xA)RP9%XylsV<d>_?My{DB'
    '-z)s}>Vx0QEoRMd@N27FGaUSK)rwS+YEna5lQyK5v?X<<9cfQGkdCCDG>}Fk*9-^0Ty-{b%{)2yoj%|5%1>y*P5NpcEqQo~mUdIH'
    'Wx-Yjs|r>ZtSQ*KV4H%q1=|*^E7-1J`+^+`b}U$5u%TdMw6vsXdL9%1yy?EM?i?+R>Cre6*1xH9{hBxw4qwJUL5}8Y;#|0}PBmej'
    'YJ#aIm}-KlCYWl1sV10ef~h8$YJ#aIm}-KlCYWl1sV10e3g&AH=4-+_)fCLvWLB!_>eFKVlR~(Z`DpIk#DQ^Kmp*wem}<ga)dW*b'
    'Fx3Q8O)%92Q%x|{1XE2g)dW*bFx3Q8O)%92Q%x|{6wKEY%-4i<swtSS$*fe<)!$Hj^PZnO@o26le)2?I=WD_`)dW*bFx3Q8O)%92'
    'Q%x|{1XE2g)dW*bFx3Q8O)%92Q%x|{1XE4Hd`-c8O<1Rzg87=vN;O@5;5_|ZasD=oM{_lC<{a1gnmBbXtW!-`r<!1@38tE0stKl='
    'V5$kGnqaC4rkY@?38tE0stKl=V5%vYuPK<X3F}l-Fkh2dsix8>`u6?rZ{=_PRGjHgHyM<f3yK5jxYq~8k#u35g2Fll1yfKk1qD-3'
    'Fa-ruP%s4rQ&2Dk1yfKk1qD-3Fa-ruP{Djq!F*6yr=WuQpv+1^73b8y>5t2A-g7ejXf7YlrsF!F52w?Gb;<|pln+e#z?2V6`M{J9'
    'O!>f+4@~*Mln+e#z?2V6`M{J9O!)-!`2_R%V4d;_=JPQt<#Y9s_4J3!JQ;Cl9oPBSaeTvaVV#~UVV!D%sV10ef~h8$YJ#aIm}-Kl'
    'CYWl1sV10ef~h8$YJ#aIm}&~<YYOIT!aCIy%-3XAs_E*(?3?#|AM>NRa}(dm9M}1C6W_~RSf`q>PBp<)6HGP1R1-`!!Bi7WHNjL9'
    'Of|t&6HGP1R1-`!!BkT)UsEt&6V|DwV7?}^QcbV=NQ`GcH)VR;G*3pHY{zx}WW?EaVVzD!Sf_ko$_J)=V9E!kd|=84rhH(^2c~>r'
    '$_J)=V9E!kd|=84rhJ0=e1iFWuuk~|^ZA$+^0~SG!25EO*}ywmdNML|HPMzgv{k{Xg4G3U3broTreJNswgu}7wkz1aV26So3)UBG'
    'DA*Xy)f9go>A>48S$B?mT}`%5pYD|Z6A(8y)AaH`{`5EfdBn{Rmse}kzYMPD<!^u0cDYI2$n;fpW54Pj6#wTR^^p7HAO5L-$nj_U'
    '=f%H;FYP}hqhO1IsW1$7RTxZ#!BiMbg~3!9OohQz7)*u1R2WQ!!BiORP%ssSbt(*|!eA;4b}HC?!BiObK0F%#n&;+bnrL3RFdMKY'
    '|Ey8?GT9}S%a?YAsU{3@RTE4#!Bi7WHNjL9Of|t&6HGP1R1-`!!Bi7WHNjL9Of|t&6HGP1R8ug2z6$0Z{kTpwg?0XXWmY<0O%u&4'
    '*JJ}$SCg50^uxebHNjL9Of|t&6HGP1R1-`!!Bi7WHNjL9Of|t&6HGP1R1-`!!Bi7WHNjL<Fke$JUlZ1;reMA%vr<jdMDxlu*?`s6'
    'WaetZz*aTER1-`!!Bi7WHNjL9Of|t&6HGP1R1-`!!Bi7WHNjL9Of|t&6HGP1R8uftQ!rl>)~Tjoz9zF$P188@$~D<^)zxI?YQnTu'
    'HNjL9Of|t&6HGP1R1-`!!Bi7WHNjL9Of|t&6HGP1R1-`!!Bi7WHNjL<Fke$JUlZ1;reMA%vr<jdbo0tJ*|62sWaetZ@K!a!R1-`!'
    '!Bi7WHNjL9Of|t&6HGP1R1-`!!Bi7WHNjL9Of|t&6HGP1R8uftQ!rl>)~Tjoz9zF$P1AJq$~D=r)zxI?YQpeVHNjL9Of|t&6HGP1'
    'R1-`!!Bi7WHNjL9Of|t&6HGP1R1-`!!Bi7WHNjL<Fke$JUlZ1;reMA%vr<jdbo0tJ*|62sWaetZ@K!a!R1-`!!Bi7WHNjL9Of|t&'
    '6HGP1R1-`!!Bi7WHNjL9Of|t&6HGP1R8uftQ!rl>)~Tjoz9zF$P1CUR$~D>S)zxI?YQij6HNjL9Of|t&6HGP1R1-`!!Bi7WHNjL9'
    'Of|t&6HGP1R1-`!!Bi7WHNjL<Fke$JUlZ1;reMA%vqDX`*PoAHZZeyXM@wrmbMtZ9@`ko5SXHpPU`@f+1=|#?E!ehTUBPw*+ZXIm'
    'uw%jcf(->5qq&;m&m*0Wo4NV8nX4&UsitYzdEIl<ZJ3|xPkrNC*T3CFb2VYkx`^xCrC-~dm}<ga)dW*bFx3Q8O)%92Q%x|{1XE2g'
    ')dW*bFx3Q8O)%92Q%x|{6wKEY%$=LK?i}~}n!?`F)t1|7uzBU0Y|iRxGIKRyZmXJLIyb>o6HGP1R1-`!!Bi7WHNjL9Of|t&6HGP1'
    'R1-`!!Bi7WHNjL9Of?1bH3joEVV!CU=4&!5otvh?=9O!*IjgJ5%+-Xst!jd)CYWl1sV10ef~h8$YJ#aIm}-KlCYWl1sV10ef~h8$'
    'YJ#aIm}-KlreMCNV7?}-Q%%8qO=hK<rorZwYqB}3tI5pOgt@J1f~h8$YJ#aIm}-KlCYWl1sV10ef~h8$YJ#aIm}-KlCYWl1sV10e'
    'f~lrpzNTQlCahCU!F)|-rJAO}=9O!*IjgJ5%+-Xst!jd)CYWl1sV10ef~h8$YJ#aIm}-KlCYWl1sV10ef~h8$YJ#aIm}-KlreMCN'
    'V7?}-Q%%99_YFM$^K>)b@_j0|(`fU`J=vty^<?IH!sJ#x!PFB>J;BryOg+KW6HGn9)DuiS!PFB>J;BryOg+KW6HGn9)DuiS!PHYQ'
    '-%~K(6V|DxU|CNQOFd1)%`5k0vsTxWnd=F&TlEA}PcZcaQ%^AU1XE8i^#oH-F!cmePcZcaQ%^AU1XE8i^#oH-F!cmePr-ao!F*3x'
    'r=EgkJw+__G>tc}+>=dPT~B7NCroeE6HGn9)DuiS!PFB>J;BryOg+KW6HGn9)DuiS!PFB>J;BryOg+KW6HGk?^F0OgJz<@C3YPU0'
    'vDDKv;Jk8AHg9!3nYo@Yzg15#^#oH-F!cmePcZcaQ%^AU1XE8i^#oH-F!cmePcZcaQ%^AU1XE8i^%Tta6wLR8b?PZt)>FhnPam#7'
    '6~EkMHWiPS_GIR!;<V)rZB?+UV0FQof~^a-DOg*uZNa*N?FzOp*r8y@g7pO(3N}V_J;k3#Iu$o_Q*krj(<7F8nueU$JwJW0nQMA}'
    'ispL440my3>)c$Nrk=1@J;BryOg+KW6HGn9)DuiS!PFB>J;BryOg+KW6HGn9)DuiS1@k=xbLS_nJIB3QPhs=)mS6lG!iQ<hdF7sL'
    '>gswjb3I{-tDaywKf%-!Og+KW6HGn9)DuiS!PFB>J;BryOg+KW6HGn9)DuiS!PFB>Jq7bU1@k>&oq7tE^%Swx(=_P3a!)pQbv>E6'
    'o-oH%PcZcaQ%^AU1XE8i^#oH-F!cmePcZcaQ%^AU1XE8i^#oH-F!cmePcZcq%=Z+`_k?xoDOlE1#8OYwsPoD_+2qyrWafIpBv(Db'
    ')DuiS!PFB>J;BryOg+KW6HGn9)DuiS!PFB>J;BryOg+KW6HGn9)Kf6uQ!w8X)~TmpSx*s5Jxv49EB9oxSJ#u7>j|@5^#oH-F!cme'
    'PcZcaQ%^AU1XE8i^#oH-F!cmePcZcaQ%^AU1XE8i^#oH-!F*4_d{0=Xo`PjPMJ)9+%{s5#lg(aTPiC$s%yQKeOg+KW6HGn9)DuiS'
    '!PFB>J;BryOg+KW6HGn9)DuiS!PFB>J;BryOg#njJq7bUVV!ykmh}{|)YCNUymC)Adv!gTxt=h~RZlSW1XE8i^#oH-F!cmePcZca'
    'Q%^AU1XE8i^#oH-F!cmePcZcaQ%^AU6wLP&%=d(K>M2;(Q^ZnF)2#E#J=yHl^<?IH!Yo%k!PFB>J;BryOg+KW6HGn9)DuiS!PFB>'
    'J;BryOg+KW6HGn9)DuiS!PHYQ-%~K(6V|DxU|CNQOFd1q&MWt1vsc%Xnd=F&T=fJ~PcZcaQ%^AU1XE8i^#oH-F!cmePcZcaQ%^AU'
    '1XE8i^#oH-F!cmePr-ao!F*3xr=EgkJw+_^ba(yv_~j<E`FOOnCo?x6r!8-2tAbSps|(f?Y+bNT!P<gt3)U5ESFnA-4h1_FtS{J5'
    'urZqJDgHdt`M8;zkDE<@Up~KI<q=CgO|#DHo}cb)_L`obqPd<h%U#^qIyWDusVD4JPcZcaQ%^AU1XE8i^#oH-F!cmePcZcaQ%^AU'
    '1XE8i^#oH-!F*4_roSAYKbM}H&vCOqLD|CTFNHB6zng}gR}RXCuP!Jv7ZirM3JRtZ6ih+E6ckKB!4wosLBSLhOhLgE6ih+E6ckKB'
    '!4wosLBSLhOhLgER4^Y@F!zARz3M5f%X*4f>S-EwUb!b5zPg^wTu&J0swbFwf~hB%dV;AZn0kV#CzyJIsVA6vf~hB%dV;AZn0kV#'
    'CzyJIsVA6v3g&wX=6k|k^%N}YDPpOoY1nz?o^1H)dNOl8VVJ9)VCo5`o?z+;rk-Hx38tQ4>ItTvVCo5`o?z+;rk-Hx38tQ4>ItTv'
    'VCpHD?<tt?3G38Tu&k$urJkl?=aqZ1;j8P(%=Ls}u6ly0CzyJIsVA6vf~hB%dV;AZn0kV#CzyJIsVA6vf~hB%dV;AZn0kV#r(nLP'
    'V7@1;Q%}LNo+6fdnueWM?#YI)t|v3s6Nb6!38tQ4>ItTvVCo5`o?z+;rk-Hx38tQ4>ItTvVCo5`o?z+;rk-Hx38tQc`JRGhPsp)Z'
    'LB)mXUnD$YDX3}MdF7zu05$#O%)77o2WQ@4nEU;PrpB;Njlt9yOpU?R7)*`9)EG>S!PFQ`jlt9yOpU?R7)*`9)EMl4X}hx>NwV|G'
    '!mmQ`0je`1+$NzZK&+hSu`26QC>2DaArT;{mh|?$EY7hE|9u?^UE0nYT`ZYrq?8{#v;G|e#?&yThB37t2JOc|gP30X)3-MG#+j%6'
    '{CQ<y{Ty4~v3e`-SUT^bNexYEXi`Iy8k*G5q=qInG^wFU4NYojQbUs(n$*ywh9)&Msi8>?O|8{;tk&w^fO-y1t?l0k>g*(S#L*kp'
    'e*V0&r+$tto_Z^ubl%028lKefq=qLoJgMPH4Nq!#Qp1xPp49N9h9@;Vso_ZtPilBl!;>1GTC1K~tDf{6o?6?#71WuhBMv-G*FTxx'
    '|B@Fc^P`=6>aCp2-JaaIJsq@XgZ6yTUJTmHL3=f5uLteTpuHWmcZ2qR&^`>>$3govXrBk|%h8Ib<Im&#WZv7^b3Yxg9e%d&w4X_@'
    '`&M-SOJ34x*VwJl(IxlsO@D2<j=MoeuZ4~pbkv}u1|2o%s6j^!I%?2SgN_<>)S#mV9X05vK}QWbtyMa$mA4{4_w{&feJkn@jy>&%'
    '&MSNBhu8DJ+NR!$CmnL}q{ds38lKefq=qLoJgMPH4Nq!#Qp1xPp49N9h9@;Vso_ZtPilBl!&7V3Q)|_ep2JgX7oLtdSnsC&(0OG~'
    '{qR~m^;SIT7>g%0JgMPH4Nq!#Qp1xPp49N9h9@;Vso_ZtPilBl!;>1G)bON+CpA2^Rz0;=J?S|-wRYj@h$Bz?q4Ua~`r);B>aBRv'
    'As0_-cv8cY8lKefq=qLoJgMPH4Nq!#Qp1xPp49N9h9@;Vso_ZtPilB-t$J#$deU=vYVE?)5l5c(L+6z}^}}oN)LZeSLoS}w@T7(('
    'H9V=|NexeGcv8cY8lKefq=qLoJgMPH4Nq!#Qp1xPp49NvTJ_Xg^`z(U)Y^roBaS@n$ImN!YWLLgT@I(-iYMJ-@uY?)H9V=|NexeG'
    'cv8cY8lKefq=qLoJgMPH4Nq!#Qp1xPp49N9hNsr5r`D<`J%^{(E<7D^<Y~WjUfEMSo>ou26;C>{;z<opYIst^lNz4X@T7((H9V=|'
    'NexeGcv8cY8lKefq=qLoJgMPH4Nt9APpwr?dJa#mU3fa;$kTq-yt1d`Qo5df8&>|W*?+9!*&*Qj8#h4dIY6laN)1qIfKmgL8lcnw'
    'r3NT9K&b&r4Nz);QUjD4pws}R1}HT^sr@i$KOT+8k258p+G{UB9dQI|ziVC@)Nv{;pmd}Klp3Ja0Hp>fH9)BWN)1qIfKmgL8lcnw'
    'r3NT9K&b&r4Nz);QUjD4pws}R1}HT^wN^p3b^+>$15oqzkKy;fq<@a$dN~L8=gOOw+v6LzCxiBM(4Gz2^FeztXfFrt)u6o|v^RtH'
    'cF^7p+WSHKFlZkK?bD!r9<(om_QRn4IA~u7?Wb>TFbDJJd+oQZ`(N_nW`4Bue^d2V<`y?J=?;u0H8iQANexYEXi`Iy8k*G5q=qIn'
    'G^wFU4NYojQbUs(n$*zL+Sxl1np!LWrt))WYR{GL&w8}M_h+5=>*sa<0=)kvFE#ZGZPBFbF7HihXi`Iy8k*G5q=qInG^wFU4NYoj'
    'QbUs(n$*ywh9)&Msi8>?O=@Urt!iqmYSMFPYOQMOZKP?xeO_5pzr~*4qvqa<Cf#<?q=qInG^wFU4NYojQbUs(n$*ywh9)&Msi8>?'
    'O=@UTLz5bs)X=1crq-&a)~Y5who;u5rrt)H_RHs$HT6qu(bQYfq{}Xv)X=1cCN(svp-Bx*YG_hJlNy@T(4>YYH8iQANexYEXi`Iy'
    '8k*G5)LPZlTGgcI(9~Mh)Z0kYe)p87(|5e$>}Q<+$Ia;<fBw6F`p?I|{N?9g{*@nF{^+|qesKBEzxng;e)|tMb3eQmKD`w_I^@Df'
    '4L)k{QG<^feAM8h1|K!}sKG}KK5Fn$gO3_~)Zn8AA2s->!AA`~tyMm)RX%zSKCM+gy^Z+nXU;3X8TGSk(bQYfq%$s>)X=1cCN(sv'
    'p-Bx*YG_hJlNy@T(4>YYH8iQANexYEXi`Iy8k*G5)LPZlTGgcI(9~Mh)Z0kYe&oEerharSntCgmbi_rI8k*G5q=qInG^wFU4NYoj'
    'QbUs(n$*ywh9)&Msi8>?O=@UTLz5bsTC18`tD5v2np&%xdK+olFPvA_)Q+ZQW|?~{nslv2lNy@T(4>YYH8iQANexYEXi`Iy8k*G5'
    'q=qInG^wFU4NYojQbUs(np&%xTC1A$9GY6IntB^)+RvL;*3>ShRa0+8lg_eeQbUs(n$*ywh9)&Msi8>?O=@UTLz5bs)X=1cCN(sv'
    'p-Bx*YG_hJQ)^XIYgLn;LsM&2Q*Q%J%k>Z9_rK)Df$3=HntJ2<bu2`WZ`_^?+S5UMHfYZW?Zu$I9JE)1_Il9X4BFd4dpBtB2kpb4'
    'eH^q;gZ6pQz8tM+I{rM)KRdj)GO+YkG#zcEX}@k>_q}Q9ht>0cHyy2L(xDbjYG_hJlNy@T(4>YYH8iQANexYEXi`Iy8k*G5q=qIn'
    'G^wFU4NYojYOQK&t-Lq+xv$4-tETqa(U0L=_DklKHTA=4(bQYfq(d#5)X=1cCN(svp-Bx*YG_hJlNy@T(4>YYH8iQANexYEXi`Iy'
    '8k*G5)LPZlTGgcI(9~Mh)Z6I2X}@G%SyMl(7EQesO*+(~NexYEXi`Iy8k*G5q=qInG^wFU4NYojQbUs(n$*ywh9)&Msi8>?O|4Z;'
    'tyN8W4o$69O}&jY?U&3eYwCy9qN%r{NrzfAsi8>?O=@UTLz5bs)X=1cCN(svp-Bx*YG_hJlNy@T(4>YYH8iQAskN%9wW>+ap{cd1'
    'skf1){je!b^LOO4<M+@1_Lsl@>0h4z^2dMv1D9|9-u&;c;Y+`?7EiquPrBUVNexeGcv8cY8lKefq=qLoJgMPH4Nq!#Qp1xPp49N9'
    'h9@;Vso_ZtPpwr?tyNEY4o|J!+}wZ0$^G2%)3<p0$lr$!d4B^@{n}bY^;SgbdW$GEM5!T44N+={QbUv)qSO$jhA1^esUb=YQEG@%'
    'LzEh#)DWeHC^ba2Rz<Z|Md>+2wRRz@$DyeGz<KresD5!RqIxT$biqZG8luz?rG_XqM5!T44N+={QbUv)qSO$jhA1^esUb=YQEG@%'
    'LzEh#TC1X3tD^KAqFTEU)#Fgqe&oEmsD5=VqIxT$bj3xK8luz?rG_XqM5!T44N+={QbUv)qSO$jhA1^esUb=YQEG@%LzEh#TC1X3'
    'tD^KAqFTEU)#Fgqe(1dNzsmdFwRq~Sc+wpgPilBl!;>1G)bON+CpA2&;YkfoYIst^lNz4X@T7((H9V=|NexeGcxtVBYOQ+Gb9id)'
    '!c&hUPwVxs<M+Sh#dZ8>=bm~i*KxNeH*QY{?b)C`AG8;P_Hxi(4chBLdoyTn2kqUUy&tp>gZ6RIJ`LLELHlyF;_3MFIKPhfR<7f{'
    '-Q1pkGrA*=ethD(A3OE_beH{i-u*9nDXL#&izr=lAK&!Xo+(3=o<o!xqSO$jhA1^esUb=YQEG@%LzEh#)DWeHC^bZ>AxaHVYKUsB'
    'ifXOAL;1O{$7?S{wHFV+rOLV=Jg+XQUtrJw&b$9WC@Z3L(M6ORqSO$jhA1^esUb=YQEG@%LzEh#)DWeHC^bZ>AxaHVYKT%plp3O1'
    'tD;(~qVyc1TDuU{<M2IdKYCtWRKLO&QN0yWy6PfI4N+={QbUv)qSO$jhA1^esUb=YQEG@%LzEh#)DWeHC^bZ>AxaHVtyNL2RZ)5l'
    'QLSBw>TxJ)KYU(YRKLU)QN0yWy6hrK4N+={QbUv)qSO$jhA1^esUb=YQEG@%LzEh#)DWeHC^bZ>AxaHVtyNL2RZ)5lQLSBw>TxJ)'
    'KZjmjRKLX*QN0yWy6qxL4N+={QbUv)qSO$jhA1^esUb=YQEG@%LzEh#)DWeHC^bZ>AxaHVtyNL2RZ)5lQLSBw>TxJ)zkFU@RKLX*'
    'QN0yWy6qxL4N+={QbUv)qSO$jhA1^esUb=YQEG@%LzEh#)DWeHC^bZ>AxaHVtyNL2RZ)5lQLSBw>TxJ)zkFU@RKLX*QN0yWy6qxL'
    '4N+={QbUv)qSO$jhA1^esUb=YQEG@%LzEh#)DWeHC^bZ>AxaHVtyNL2RZ)5lQLSBw>TxJ)zkFU@RKLX*QN0yWy6qxL4N+={QbUv)'
    'qSO$jhA1^esUb=YQEG@%LzEh#)DWeHC^bZ>AxaHVtyNL2RZ)5lQLSBw>TxJ)zkOa^RKLa+QN0yWy6z%M4N+={QbUv)qSO$jhA1^e'
    'sUb=YQEG@%LzEh#)DWeHC^bZ>AxaHVtyNL2RZ)5lQLSBw>Tx7$8(z%A3+MeWdU4@A+PSFS%FW#E$&K67L3=i6&j;<rpuHTlSA+I?'
    '(B2H%+d+FbXzvH@!=QZ}v`>TfdC<Nbt%y4QJkD?Cy_K7JZ&lP0hobiD=e6IXwtkO2e~&s^5vBX?<D33mxtY5mO0Pwf8luz?rG_Xq'
    'M5!T44N+={QbUv)qSO$jhA1^esUb=YQLR-`t(EsEKlk-`?S-iJ;^EDF+pnKj7uD~vMO1G^l<vETQsX^J4N+={QbUv)qSO$jhA1^e'
    'sUb=YQEG@%LzEh#)DWeHC^bZ>A*!`1s<kRg&mpR{3sF4|-=p^H=ha2^du$QaTM?!EE~3;BrG_XqM5!T44N+={QbUv)qSO$jhA1^e'
    'sUb=YQEG@%LzEh#)DYEL71df5rRNaU+J&edhobiD=ha2^du$QaTM?!EE~3;BrG_XqM5!T44N+={QbUv)qSO$jhA1^esUb=YQEG@%'
    'LzEh#)DYEL71df5rRNaU+J&edhobiD=ha2^du$QaTM?!EE~3;BrG_XqM5!T44N+={QbUv)qSO$jhA1^esUb=YQEG@%LzEh#)DYEL'
    '71df5rRNaU+J&edhobiD=ha2^du$QaTM?!EE~3;BrG_XqM5!T44N+={QbUv)qSO$jhA1^esUb=YQEG@%LzEh#)DYEL71bL5tK%!T'
    'e;hoAsP^22s2+!+_Uq@>MfH1Z5!G7}rTZ?T)DWeHC^bZ>AxaHVYKT%plp3Pc5T%AFHAJZ)N)1tJh*Cq88luz?)mjzRS{0?|5Y^g+'
    's2+!+_Uq@>MfH1Z5!G7}rTZ?T)DWeHC^bZ>AxaHVYKT%plp3Pc5T%AFHAJZ)N)1tJh*Cq88luz?)mjzRS{0?|5Y^g+s2+!+_Uq@>'
    'MfH1Z5v3dL_cxpFyNFUllp3Pc5T%AFHAJZ)N)1tJh*Cq88luz?rG_XqM5!T44N+={QbSa0Ra9$Ll%7LWYZsz=9EiGo^f$Pfw_mj0'
    '%)hmVzqh|X+WF0V(4Gw1(?NSSXwL`j#h|?$v{!@ndeGhs+S@^UH)!t%?Zcpb9JEh^_Ic309PM1x(auE;p8I;VBI;WlirTNASALJW'
    'ebn!<XYWzpS`nrD?&IUR^PBmgAxf`Blp3Pc5T%AFHAJZ)N)1tJh*Cq88luz?rG_XqM5!T44N<LCQLR-`dJa*oU5M&&C~CicN)-Qb'
    '_&@&o^xH@M8e2T|cJ{BR<GK0{rHe1`P-=Km!;>1G)bON+CpA2&;YkfoYIst^lNz4X@T7((H9V=|Nexe}RZp!|PkIhdt=-(5eI?98'
    '^bg{=eYD>`uk7i#vGRN2`y>2??+g5Y|NiFDcJYHa+K1-*C*0nB`_EVY8a!V75P8>fgqb7G9D(MDG)Jg8V$Bh3j%ag)n<L&F0q2M~'
    'N60y1&JlEusB?s!;}4_w<0vA}FaPNw-oLl}uRs6k-~ax9^FQ|Q5`Ft<zmZ;9UccFvce38fJDG047*m6o8p6~7riL#yc&VXF4P0v2'
    'QiGNnveba3hATB#si8^@RBD(~gH&tvovgL`PNwJZ)EeI&;ka$VskNI&_dhf0g42E=y)vhMuq~W=E1YxyhLakc)ZnBBCp9>!!AT8H'
    'YH(76lNy}V;G_m8H8`okNexbFa8iSl8k}0IoLZ}#^c<X8D_?T_5J!J8?&r}fd+O)f;;FadN#|cYso_ZtPilBl!;>1G)bON+CpA2&'
    ';YkfoYIst^lNz4X@T7((H9V=|skQ2<wdzUF;i<I?Pe&Yi+K;1G_SBEH#Zzy^la9Z5Qp1xPp49N9h9@;Vso_ZtPilBl!;>1G)bON+'
    'CpA2&;YkfoYIst^Q)|^zYt@sU!&7S)o{l*3w4X+=?5Uq>i>Ka-C!K!rq=qLoJgMPH4Nq!#Qp1xPp49N9h9@;Vso_ZtPilBl!;>1G'
    ')bON+r`D>c)~Y8xho{yqJRNc5X+Ml!*;7B%7EiquPdfbKNexeGcv8cY8lKefq=qLoJgMPH4Nq!#Qp1xPp49N9h9@;Vso_ZtPpwr?'
    'tyNEY4o|IJcsk<1)6MXhe*4}29rXU^x2pSJ_2Qa-v~yCum0OkDlUui^gZ6CDo)6lKL3=r9uLkY)puHKiw}bX>(B2Q)he7)|XrBh{'
    '^Pqh>S|N4(d7OP#!r-~DN2{dzi$|X?e9KYmH&=d(y8k6FJ@wmd@ua)&<M!JrKm8g{dJa!&cv8cY8lKefq=qLoJgMPH4Nq!#Qp1xP'
    'p49N9h9@;Vso|-$>Z!Hz7Uk#g)SkQWbi~nH)P5}0UwOakN1ferqu+l7I-do2oGE_)trb8z^a4l?KxzO|1CSbk)BvOgAT<D~0Z0u%'
    'Y5-CLkQ#v00Hg*WH2|prNDV-(RY0v(Kza^9tz7^*;t0@w^Yp)&(!b-{CAE_3t&nQ1{vD?VDK$u`K}ro$YLHTclp3VeAf*N=HAtyJ'
    'N)1wKkWzz`8l==9r3R_iDyh~gDLn_N)-FgLaYSlAabEYI;Wz!_THd01E1q=0<t<7LPilBl!;>1G)bON+CpA2&;YkfoYIst^lNz4X'
    '@T7((H9V=|Nexe}RZp!|PkIhdtzCFJ;>go}_q?*FesV3IdMloE!o`ysp49N9h9@;Vso_ZtPilBl!;>1G)bON+CpA2&;YkfoYIst^'
    'lNz2{tDah`p7b1^TD$Oc#F3}{#Cc^;{p4Ca^;SITgo`INJgMPH4Nq!#Qp1xPp49N9h9@;Vso_ZtPilBl!;>1G)bON+CpA2^Rz0;='
    'J?S|-wRYj@h$Bz?iSx>y`pLC;>aBRv2^UXlcv8cY8lKefq=qLoJgMPH4Nq!#Qp1xPp49N9h9@;Vso_ZtPilB-t$J#$deU=vYVE?)'
    '5l5c(6X%sZ^^<Gy)LZeS6E2?A@T7((H9V=|NexeGcv8cY8lKefq=qLoJgMPH4Nq!#Qp1xPp49NvTJ_Xg^`z(U)Y^roBaS@nC(bK-'
    '>L=IYskh=uCtN(K;YkfoYIst^lNz4X@T7((H9V=|NexeGcv8cY8lKefq=qLoJgMQSwd$$0>PgSxskIAFM;v(KyPIFz)BP`bae_M9'
    'xu@RBy}<3st=rQ<dp2m#2kph6y&SYxgZ6sR-VEB?L3=l7?+5L}pnV*)PlNV((7qh4csl+(&Trzqojv#Ucy0C69~|Dq`IhJL4fFoj'
    'xCGRXt_6?|wvTW7Ys)8;y8%eA1&|tm)BvOgAT<D~0Z0u%Y5-CLkQ#v00Hg*WH2|prNDV-(RY0wkHzYp?p!VF&?ZY43Pv0U3POjAT'
    '^X6{%`{s53g1i4EFFo}eZ1JS~EuPePLsG+&8lKefq=qLoJgMPH4Nq!#Qp1xPp49N9h9@;Vso_ZtPilB-t$J#$deU=vYVE>Pk3&!U'
    'ar4Ta`my!=A547G`2FruJn49gCpA2&;YkfoYIst^lNz4X@T7((H9V=|NexeGcv8cY8lKefq=qLoJhfImwN^dpIXty?;i<==r~SBj'
    'Wl#OsT0HeuJn49gCpA2&;YkfoYIst^lNz4X@T7((H9V=|NexeGcv8cY8lKefq=qLoJhfImwN^dpIXty?;i<==r~SBjWl#OsT0Heu'
    'Jn49gCpA2&;YkfoYIst^lNz4X@T7((H9V=|NexeGcv8cY8lKefq=qLoJhfImwN^dpIXty?;i<==r~SBjWl#OsT0HeuJn49gCpA2&'
    ';YkfoYIst^lNz4X@T7((H9V=|NexeGcv8cY8lKefq=qLoJhfImwN^dpIXty?;i<==r~SBjWl#OsT0HeuJn49gCpA2&;YkfoYIst^'
    'lNz4X@T7((H9V=|NexeGcv8cY8lKefq=qLoJhfImwN^dpIXty?;i<==r~SBjWl#OsT0HeuJn49gCpA2&;YkfoYIst^lNz4X@T7(('
    'H9V=|NexeGcv8cY8lKefq=qLoJhfImwN^dpIXty?;i<==r~SBjWl#OsT0HeuJn49gCpA2&;YkfoYIst^lNz4X@T7((H9V=|NexeG'
    'cv8cY8lKefq=qLoJhfImwN^dpIXty?;i<=wr@Qe%Jp3HG`(O3q1a-7?QoWU%xZ9Ikx2J>lY|x$$+KWMZIcTp2?e(C&8ML>9_HNML'
    '588)8`#5Ny2JQ2peK}eob^Ljp-^6<>H}T#osUr?a?KjZtzD3>j^XvIr)X|D3opT@G_UFop+zn59EuPfyq=qLoJgMPH4Nq!#Qp1xP'
    'p49N9h9@;Vso_ZtPilB-t$J#$yhZuBug7aIJhc}OH<r8o+<9eB#|4#lzlZOUar5Z%dt}^y=<+oN-`~1nOlM!-v(zxAhA}masbNeF'
    'V`>;v!<ZVz)G(%oF*S^-VN4BUY8X?)m>R~^ei*bL2MuF-?N8s@<`L#@KY(6YSU<>?cdXvZJC+W-Xi`Iy8k*G5q=qInG^wFU4NYoj'
    'QbUs(n$*ywh9)&Msi8>?O=@UTLsM(@9jmn)KCAcJH*g)i7*5B7=ik=nTfG0{tgGJLegPGzIr;(DxA(v3<v*(WowktDjTcgCkWzz`'
    '8l==9r3NWANU1?e4N_{5QiGHlq|_j#1}Qa2sX<B&QfiQDt&(c3{w1jA5Y^hnUxGaj|53HyK(8*UUuDaiRBuI;uDpm+LzEh#)DWeH'
    'C^bZ>AxaHVYKT%plp3Pc5T%AFHAJZ)N)1tJh*CpTYgJThRg|7XRBIQadK`+{ub@{K)$g)JRBuI;?!1UnLzEh#)DWeHC^bZ>AxaHV'
    'YKT%plp3Pc5T%AFHAJZ)N)1tJh*CpTYgJThRg|7XRBIQadK`+{@1R%yXH>t-7EiquPrCEsNexeGcv8cY8lKefq=qLoJgMPH4Nq!#'
    'Qp1xPp49N9h9@;Vso_ZtPpwr?tyNEY4o|IJc<OQJX}^PB*;BvE7EiquPrCEsNexeGcv8cY8lKefq=qLoJgMPH4Nq!#Qp1xPp49N9'
    'h9@;Vso_ZtPpwr?tyNEY4o|IJc<OQJX}^PB*;BvE7EiquPrCEsNexeGcv8cY8lKefq=qLoJgMPH4Nq!#Qp1xPp49N9h9@;Vso_Zt'
    'Ppwr?tyNEY4o|IJc<OQF=`_5Z|GIBraQ|yu{72o<&H?pSF6C}dZrz>^+Ot7>K4>on?d71o8noAg_GZxD4%)jxdp~F&2JPdZeHygS'
    'gZAZU1<>*5aegW9tz61`tALI;1hij0ult5{>Q~tFH>9H#PrB+pzU|MIOSv1K^jbWr;YkfoYIst^lNz4X@T7((H9V=|NexeGcv8cY'
    '8lKef)LQk_T6shAb6=0wUU+IR9$v~%``Po#p4ug~d^e2y55s#_@uZ6@p451MQp1xPp49N9h9@;Vso_ZtPilBl!;>1G)bON+CpA2&'
    ';YkfoYItg`dTOnD(sOuf?ZQ)!!}q8C?kP{d@-Mh{JFS3vD}Z#g1&|tm)BvOgAT<D~0Z0u%Y5-CLkQ#v00Hg*WH2|prNDV-008#^x'
    '8h~1>fLg17^c;X%y8zVV5YT?nOab+iYWW+jw*p8fS^%j5NDV-008#^x8i3RQqy`{00I2~;4M1uDQUj10fYbn_1|T&6sR5|93aGUT'
    'NY4SNwF^K!4gu{4%oI?+pcX*A6+pVk0!R%&Y5-CLkQ#v00Hg*WH2|prNDV-008#^x8i3RQqy`{00I2~;4M43`K&@3kdJaIXT>$EF'
    '2xz}orhxjvv;gX@0MY>#KxzO|1CSbk)BvOgAT<D~0Z0u%Y5-CLkQ#v00Hg*WH2|prNDV-00BWrQYOMm&a{y}X0#J`bK>L9*1=KI3'
    '1yFAVkS?wOQUj10fYbn_1|T&6sR2k0KxzO|1CSbk)BvOgAT<D~0Z0u%Y5-CLP-_)XYZZ{515j%hfO;GP+Aoq<{*SVLAuXPIE1q<5'
    '#giJI)bON+CpA2&;YkfoYIst^lNz4X@T7((H9V=|NexeGcv8cY8lGCKo?5G(^c<dAyYO_x`;$}s+>+CNksNvY`QQHX*FXKs^I!h>'
    '&wt=^+Bg&Z$Ia;<fBw6F`p?I|{N?9g{+0IVyE}dw)X)3R|M`d8`(J1O`A^tA`y9BN%g=%P_T2qXeYt<}_YW_Ah|oF?s1Ffaj^J`c'
    'mm|C!@#P3GM}#>-%n@UbAag{SBg`Ce<_I)Lq&Y&(5o?Y=jN*@@2sgj{r-wNB+_>q#`i?2TlH2{S12;Z%_~Bdiev1#ESoQsFk9>aB'
    '@!;d5@w-3TlcVwNIga*p(4Gz2^FeztXfFrt)u6o|v^RtHcF^7p+WSHKFlZkK?bD!r9<(om_QRn4IA~u7?Wb?;{?D?$XHEOj((hUK'
    'UuE!X9ku%>pW%7NsI~P$0*d-`MNw;2QEOF^8j93Vq=q6j6se&|4Ml1wQbUm%ique~h9Wf-si8;>MQSKgLy;Pa)KH{`qTWV|_EY7R'
    '6t$~oRn%Hl)SoMgTC0j$tBTZ6q=q6j6se&|4Ml1wQbUm%ique~h9Wf-si8;>MQSKgLy;Pa)KH{`A~h8CHd3_TE3c%eoj$9g)~cfZ'
    'Tv61|re`naue25`BQ+SQ!AK28YA{lRks6HDV59~kH5jSENDW46Fj9k&8jRFnqy{537^%TX4Mx3<80}ZfD=}&Z(8{Q_%BVkA81>_6'
    'Wz=G2qy{537^%TX4Mu7(QiG8ijMQMH1|u~XsliAMMrtrpgOM7H)L^6rBQ+SQ!Kk+pqy2VyB}VNGS{b!g8TIE1qkclIj9RRW)L^6r'
    'BQ+SQ!AK28YA{lRks6HDV59~kH5jSENDW46Fj9k&8jRFnqy{5381*({v|lj)e;NJt4}bjm|NQ<>|Mihzdbkoy{WN;^Ut_<&{T3e%'
    'Q2c}Yc=93g-))ZkE$7I8xj7=s5n7Jeas-znx*Xx<h%ZNgIU>vvVvZPd1eqhs9AV~&Ge@8~BFzzMj#xY5;H#Ua{o46A!Pd{H`)8h?'
    'ouqCrPg3{4f&G@Soc-wXLB!wjD?U7lKS%sI;?EI(j`(xLpCkSp@#ly?NBlYB&k=u)_;bXcBmNxm=ZHT?{5j&!5r2;OJL2GHO><|x'
    '+TWS&&|3e_Y_0yCsfH~zXsIDf4OnWpQiGKms?<QGhAA~jsUb=YP-=KmgOeJX)WD>MB{e9iAxRBLYB*AZQENARZm+)e>a?Fiuk=r^'
    'c2%v6TC0rIV59~kH5jSENDW46Fj9k&8jRFnqy{537^%TX4Mu7(QiG8ijMQMH1|u~XsliAMMrtr>ZOCZ;Z@z=due@#E|5CsEZDG_~'
    'VdVDs*6qolJsq@XgZ6yTUJTmHL3=f5uLteTpuHWmcZ2qR&^`>>$3govXrBk|%b@)*Xg?0x*FpQ~Xk$kE-Sax{MtlPI{>^zSqt+@T'
    'H5jSENDW46Fj9k&8jRFnqy{537^%TX4Mu7(QiG8ijMQMH1|u~XsliAMMrtrpgOM7HS{pOk51&_J)NZJiQEQcv8jRFnqy{537^%TX'
    '4Mu7(QiG8ijMQMH1|u~XsliAMMrtrpgOM7H)L^6rBQ+SQ!AK28t&JJ&htDf9YB$u%sI|&S4Mu7(QiG8ijMQMH1|u~XsliAMMrtrp'
    'gOM7H)L^6rBQ+SQ!AK28YA{lRks6HDV5A14*2awX!{?P4wHs<>)LLbv1|u~XsliAMMrtrpgOM7H)L^6rBQ+SQ!AK28YA{lRks6HD'
    'V59~kH5jSENDW46Fj9k2Yhy<H;qyw2+6}cbYOOL-gOM7H)L^6rBQ+SQ!AK28YA{lRks6HDV59~kH5jSENDW46Fj9k&8jRFnqy{53'
    '7^%UiwK1dprunzHTh8rRT2ZxDQK^AS4OD8NQUjG5sMJ8E1}ZgBsewujRBE771C<)6)Ig;MDm754fl3WjYM@dBl^UqjK-Jn9)qbeF'
    '(tA_CdG22_e0Ki6yF7o-$8+bOH8+Uh@{{OtgqI_}j`;A&RscI9fA63DE>HD$2|E9ZsB?s!Bkml5=ZHK<==ryfz432-s=swS`d0*`'
    'BO)Ck>4-`HhCn&~4bAm81XTZuNOgqj5${*H`dN7Mehj@5)pl2{?|8buet&!aA%FMVvv)yI>N!lRK~fEo{hrI$0Ue+IhsYy7e1Sgv'
    '^lJ2}flm#4YS2?do*MA{I|QD8hbYslL8gW=HHfJp%pVzH{>X6BtI?zervBp5^?cq>qF2JyuB#Q3uCQQI1Ctt<)WD<$rhe6}n2uO5'
    '>D6FT1Ctt<)WD<$CN(hmcL*l`4#A{XgGmicYG6_WlRq+;{E@+=SA$6nO#Q_pOv|siq5qmsU%3B;S&I3porg~s{ua-_qG1pp-wxuF'
    'QG7az&mQ8#r>#H#7Wp*y<G1u;^zxUZ_-YhikK&tAd^?KoM!(PZqu=L;(aS%M;-^vkJc?gNAN_~XNB?p3@~@-#)5FULpDVNM=hNt&'
    'i$AwN|LNcNFQd6#_9NQWelh+@b^o<H)35l<`19Wk{|z0zp8t4?UeCXiqu29~m+1BUeHFc)x2WXx{4*zeJu;(@560;Ap}#fsw}$@K'
    '(BB&RTSI?q=x+=CZK1y{^tXln_$D)vf4-wp_WID@7W&&le_Q77?osCN?osCN?osCN?osCN?osCN?osCN?osCN?osCN?osCN?osIP'
    'CiHg``s2(0llPCio6sLW$RPXpp}(8Z-%aT6CiHg``r|7lqW}NzKRrJ9_uJk6)54?IhyM1T1{{6-(BJ;kKBJEx`rCh+U-a=qfBR46'
    'N<RMn<C>z^hyL~-sS<ts(BJ+u6rztG`rE%ldi3!_fBSbpPCouVTt%-B{q5i2GWz(Tzx~_WMIS%(w|{f1=;Md}_U~mAef-eh{+&aT'
    'kH3HKf#~(2zx~UNqmLi@+rPjt`uL&0{fmC0k01KmKg~b+`1|MJN3Re4?Vr&Zef-eh{yALH#}EDOpQjOh{LtUNmqs5y^taDd$;aQ9'
    'k?8fIza{jyg#MP$-xB&;LVru>ZwdW{`^ViH`ddSPYv?cBKkmZ)<1XAk?!x`!F5Exv!u{he+&}KZ{o^j&KkmZ)<1XAk?!x`!j&F7u'
    '{rwT{A9vyYaTo3%r*Qu`h5N@T+&@m?{&5QTk5jmRoWlL%6z(6VaQ`@k`^PEVKThHPaSHd3Q@DSe!u{iPll}Ye6z(6VaQ`@k`^PEV'
    'KThHPaSHd3Q@DSeZnJ;?ox=U&bQ}5$_m9*5t)PSd{yT;H$LTiv?~hZsf1JYo<HTn$NB{o|_m9*5Q^k`%zi|IJ?ccX4`uL&0{U?G('
    'A3yXL?jNW9hrlJj|8V~}?LXfs`uL&0{inM`A3yZB|7?fo<A?tCZ_b{4{BZv`h5N^8|E{~y??3dnf1l0h<A?tCZ<-f<{LtV2-L9gK'
    'ANmXTkJJ8*O_JY#xPP4XZ)FgD{LtV2#mmvh5B=?5JQ#ia(BJ+gLdnMu_m5M!f1LI&28n+Ep}%ndIEDMiDcnC!;r?+7_m5M!f1JYo'
    ';}q^6r*Qu`h5N@T+&@m?{&5QTk5jmRoWlL%6z(6VaQ`@k`^PEVKThHPaSHd3Q@DSe!u{hE?jNUc|2T#F#}w`#Q@DRj;r=m&`^OaS'
    'A5*x0OyT}9h5N@8?jKXQe@x;2F@^ib6z(5WxPMIH{xOC7#}w`#Q@DRj;r=m&`^OaSA5*x0OyT}9h5N@8?jKXQe@x;2F@^ib6z(5W'
    'xPMIH{xOC7#}w`#Q@DRj;r=m&`^OaSA5-}CA5*x0OySpmOyT}9h5N@8?jKXQe@x-me@x;2F@<0MF@^ib6n_236z(5WxPMIH{xOC7'
    '#}t13#}w`#Q~32CQ@DRj;n#mm;r=m&U;i<M`^OaSA5*x0OyT}9g<t<Mh5N@8e*MQ3?jKY5^&eBXe@x;2F@^ib6z(5W`1K!CxPMIH'
    '{xOC7#}w`#Q@DRj;r=m&`^OaSA5*x0OyT}9h5N@8?jKXQe@x;2F@^ib6z(5WxPMIH{xOC7#}w`#Q@DRj;r=m&`^OaSA5*x0OyT}9'
    'hx^AI?jLiwf6U?jF^Bud9PS@;xPQ#y{xOI9#~kh-bGU!Z;r=m)`^OyaA9J{W%;Ekqhx^AI?jLiwf6U?jF^Bud9PS@;xPQ#y{xOI9'
    '#~kh-bGU!Z;r=m)`^OyaA9J{W%;Ekqhx^AI?jLiwf6U?jF^Bud9PS@;xPQ#y{xOI9#~kh-bGU!Z;r=m)`^OyaA9J{W%;Ekqhx^AI'
    '?jLiwf6U?jF^Bud9PS@;xPQ#y{xOI9#~kh-bGU!Z;r=m)`^OyaA9J{W%;Ekqhx^AI?jLiwf6U?jF^Bud9PS@;xPQ#y{xOI9#~kh-'
    'bGU!Z;r=m)`^OyaA9J{W%;Ekqhx^AI?jLiwf6U?jF^Bud9PS@;xPQ#y{xOI9#~kh-bGU!Z;r=m)`^OyaA9J{W%;Ekqhx^AI?jLiw'
    'f6U?jv4s1_67C;MxPL6+{;`Dn#}e)zOSpe5;r_9N`^OUQA4|A@EaCpKg!{)5?jK9Ie=Ontv4s1_67C;MxPL6+{;`Dn#}e)zOSpe5'
    ';r_9N`^OUQA4|A@EaCpKg!{)5?jK9Ie=Ontv4s1_67C;MxPL6+{;`Dn#}e)zOSpe5;r_9N`^OUQA4|A@EaCpKg!{)5?jK9Ie=Ont'
    'v4s1_67C;MxPL6+{;`Dn#}e)zOSpe5;r_9N`^OUQA4|A@EaCpKg!{)5?jK9Ie=Ontv4s1_67C;MxPL6+{;`Dn#}e)zOSpe5;r_9N'
    '`^OUQA4|A@EaCpKg!{)5?jK9Ie=Ontv4s1_67C;MxPL6+{;`Dn#}e)zOSpe5;r_9N`^OUQA4|A@EaCpKg!{)5?jK9Ie=Ontv4s1_'
    '67C;MxPL6+{;`Jp#~SV*Yq)=`;r_9P`^OsYA8WXOtl|E#hWp1F?jLKof2`sDv4;D{8txx!xPPqS{;`Jp#~SV*Yq)=`;r_9P`^OsY'
    'A8WXOtl|E#hWp1F?jLKof2`sDv4;D{8txx!xPPqS{;`Jp#~SV*Yq)=`;r_9P`^OsYA8WXOtl|E#hWp1F?jLKof2`sDv4;D{8txx!'
    'xPPqS{;`Jp#~SV*Yq)=`;r_9P`^OsYA8WXOtl|E#hWp1F?jLKof2`sDv4;D{8txx!xPPqS{;`Jp#~SV*Yq)=`;r_9P`^OsYA8WXO'
    'tl|E#hWp1F?jLKof2`sDv4;D{8txx!xPPqS{;`Jp#~SV*Yq)=`;r_9P`^OsYA8WXOtl|E#hWp1F?jLKof2`sDv4;D{8txx!xPPqS'
    '{;`Jp#~SV*Yq)=`;r_9O`^OgUA6vM8Y~lW~h5N@A?jKvYe{A9Yv4#7`7VaNgxPNTn{;`Go#}@7%TeyF0;r_9O`^OgUA6vM8Y~lW~'
    'h5N@A?jKvYe{A9Yv4#7`7VaNgxPNTn{;`Go#}@7%TeyF0;r_9O`^OgUA6vM8Y~lW~h5N@A?jKvYe{A9Yv4#7`7VaNgxPNTn{;`Go'
    '#}@7%TeyF0;r_9O`^OgUA6vM8Y~lW~h5N@A?jKvYe{A9Yv4#7`7VaNgxPNTn{;`Go#}@7%TeyF0;r_9O`^OgUA6vM8Y~lW~h5N@A'
    '?jKvYe{A9Yv4#7`7VaNgxPNTn{;`Go#}@7%TeyF0;r_9O`^OgUA6vM8Y~lW~h5N@A?jKvYe{A9Yv4#7`7VaNgxPNTn{;`Go#}@7%'
    'TeyF0;r_9O`^OgUA6vM8Y~lWK^C%oZ_yul$`<p-f^>@F$nZEmfq?t_$!Y~wsuTpvd`(lfKNLO8WjSwH|LJLibdV8lmg5NeI!@Mxe'
    'TYSqhN^>sb^Wu-&dRSMV(8YAReh1IVU%F_Uwr=|KG`_1z&kQ0E3AEHfM+G!E`-2@7*i(T66*xixh(dxv79xQy9f2kh;t-8MlL&!`'
    '9MB{}Bq9eii4ckegH%KUG>MRl$N}v{NJiv<b|PdWazHx~(ve^gk4S)SX9!5-fOaB8By#X)CpP)+UWQm={{sL3|Nj600RR8IoZE8S'
    '))j_d1*2=R#s*0c(%7B>K*$7$gGliZ^-hgtTT{hyX=T#G^db68eX;&OlQ@a{u~(hkAlCl0h{M`zL-`iKm*3qV-rtw^kB^_DZ=b&Z'
    '`26ZD9fz~mFTadOhx>Q0&hliMUfybR_JY{VY{o1TOF1r`UA{?XF~ltrv4PpO*_PSP?3P)cOfJMG7kr%YWtmK;z~;#`_W99qiUqU4'
    'afw;NY-~1WHZhwpo0`#a7Re;lb(zq$-kI&0DbCk(LitnN%yGmNmpLwj<5XPcIAV$$JC4p>abw33Q{2RHbX^oTaU3zlO&v$qOmS1k'
    'ow%N{<0$WniycQy`wNbvIia}VIAV%R97pp-af#!I70KB78*~1|_RJLL{Eaz(Vv6JZjV~NWb+ps4<A`a0&fl2xC#E>g-`M#ZbN<8>'
    '$N3w(9>!b`Vu~ZC<8<D~oOdzBao)$yyOz)YUL$o~oOfh-GWI$#K9O73%9v|q>{>xqB#Gxl!gE4w&rETi6A8}=F~#wmNIWMJo)cn<'
    '<2jLdP9!`h#1zMKBJrF^cut5Zj^{+;Ig#+35K|n_iNv{0IJaVo<C;!f(+StKnBq9^iSwRt-o+HhYh-Z#0_RU`&#Xv-^A|XOVv0MF'
    'Ki5jawUW42kZFI;U*P<SX@8u*;QR&7pP1q}f5G(-xE{n5$M?A4`~}XRnBq8p!TAfEKQYCf$e-&Wa6JUq1F|BCoxhm#C$?v%JUD-W'
    '^CzbL@md#L4}t4JOmSQfvGW&m{={^goWI!li#dN{isSsnu7{ZGK}>PHp2g04%y}169I-r!{dZl=xs9FMm~$(p&*Zf(cCEx*D`JYn'
    'S_v1L6WY-I$ZXF{bIEfe<~bpz{qdZLJttzG6Jm<PIT0>2C$!<?)SO^egbU3HZ8%PIf|=qpC$!->%?W0TQ%!5bajI!%ic{XT;W*`;'
    'nc|dpZ8%POXQnviPaBR?{+KCF`O}8uls{&QQ>|#jajF$&MaZ1LjPob9XQnvkFXQ})DGqBoWUhyd>p@I$c#jL2^Ote{#1w~h8#3oF'
    '<NS#!j_WpaJ!D)DVv55xGGxwQ#`zOd9M@mw{AHX!F~#9ME@ZBUjO#&6akvhKne#W}{E6+E6=CN5%{YHziaU`%*F(njkhva^X@Ab&'
    'jPoa^{c-+g&fkplC#E>g-^}$e<9ZNN9Ny!?%=w#f{=^i=`I|X^GtQrw;!fnx^)Taln7JO16(M#0QqG^)o|*FC{Glh>nQ4D`KMFI~'
    '!;I@eOmSQfsq=@<YNufzC+E*zDx{o0F~xEIQrAPu^&qA=Ue8kJFXjA+DUS1(I)5qWPfT&Vo~5pbl<PrEam4a4^*K!G9K@XWDLT23'
    '%DZU;_Z7`0=Qicsis@KR<kq#4a;>DU6=X%2`rOgG?aXuzycWotDvs9zF~#v*ntCpwQ`P=Fmyqc=eSYX173cFqrZ~?f^pT45TtcQe'
    '&n0w~it}7TRs^}&tv2{KgV>&#_UE+#U99%!wE&soJYUe$D$er-nU2$WXLo3(<K(>C#Rhh<8g?AzJrDA&x7y&Gw{H#XTQ%&@<aNSc'
    'F|b$Eu;cjtU{4#^(`wjpyuR4e2KKZXb{wxS_KJbMqJ|a6UNOkUYJYrxx2I*M+<H#1r!`X?uOaq|fxV)JeVkm=_OyXLt%e=PdAFwx'
    '>}fUZIL^ELpyK%cV0Xx@2==tvU~SsdGE<yug*~mA;<#4qTLb%64LgqS0rstdeXE8YcOrkT2lTBw4LgqO!A>=>Q`NBIkQKpBRU5pX'
    'cB;%2=XzkLYNj~;d)H1iuv68r<GB9pR0BIz4LgqWXQvw2scP79Tn~1tft{*`9mn~zQw{7?HSD+(`Exz6Q`NBJcs;XI4eV4ktT<#v'
    'uv66r^KPfgOnG-buv0bDadJJ_sRnkc8utD;e|D;YovMZ%$N95U4eV4k>^NG_@*qEWs}0`c>{|o-Rt@_zIk)z$fqkon9d{zPt`+vJ'
    '8g?AlioIfBuc%?gu~!WCwA$dDu%~6Fx#T&)p4Lpq$#cT)Ft9t+u=mGv!k#v;r`53Icuv^U2KKZXb{ya1>{|o-Rt-Cj^Df`2IKIb;'
    'DUS1Q7aQ2cYS?j@cPmvksurq9ieR6r4d>jsW1kxAQ?=nZuaE3gty9^iitLdTR;g@KMfOMvYgD$VB6}o-6)GE4kv)>a`jqXd$R0^y'
    'b;{;cWRIk<Hf3unvPV)_nX)kz*&`{eOWBr+?2!~!rEE$qN|6-Sq-;qoN!gKFk+LDRAZ0&lJ<4{}a+KYu)hL@$i&6HX)}m}hEk)Uh'
    'O7Csi{8{`V$)!qf*_2w8vM04BWlL&F%8t~ElntpxAbUV-K(>IEs_ay)RFM=Gs_av(Q`x3U@NL=sS^cs3vq)u+YK_Vk)e@B*sue05'
    'R0}`$eb#+!`z%k{om!o;Ikh-tZ)$DI*3^=Z9iJ5+8$JtD_NCUPND9kRcBNLOY)UQm*y~yAAt@|L*^ye2vLUq~Wj|^?%63#zZp*&T'
    'x{hs~#VC7GYf-kMmZI!Ltwh;~N?vZ!UGA(kkra}eTXdQ`kv-c^%T9KkR-H(4Ny%Gwl2(#zBrPP_M_Na+jkJtp7b)$#MeDn>5Js}G'
    'rejNINrxnt*1SclnM+*WqPNVgE!kRHS|S;w7q8JSZY(9)Nm@y=k<|Ij@q8+gD<p437D#H52FX&SMbe4%NCuH>_MtkTIi61?a)soL'
    '$O1_%(jZxiv`9LU9!V|*ddmh>LV1p^xe~cT@<wEVq!wwAEJa!*ok)*l5V=OOu_9+fZb8n^F-u7gXrxyb+3c~cv618w)dRZdmGw5;'
    'ZOd&WD?et;PZ|3mV|L!s%yYENmB<1~EixckS!lD*w$5goEnz)JUtNhTkkldrl9kmsn{kVA_To~}b2QPF$O1_%G9X!5db9Jk@@C_$'
    'Cczwg1eM4FNi8xUS-Bm--3T`$xECS)KS%pti7b%RA_J0@8x7oNaGQa<3~D6IvByw}ERfV91Co_n7TmFL!vaaJ7D0_2f<|PCq!sCq'
    '^dbY2TrGtGI|?h2TO_&kevNj&5m_Q>MLHzC$bckQd!@$iN+YsF(u#CQdXWK1uJ%Zcos&jniKG?jkn|!0l3eYS8apkG$P!5_(jn<Z'
    '1|+%KAvJbG8j&TER-{AHiwyiMyV@r;c263SC6ZR8L(+>3NOHA5YV3YAB1<H#NQa~s8Ia^^Kh@ZMYDAVuT9FP(FESv>)jq1Rd(?<5'
    'k+dQml3qkFPk}eg)o!Y>6V-?;k+dQml3rv$qW8tLA2Q30$P!5_(jn<Z1|+%KOEq?v8j&TER-{AHiwsEg?wIyNMz|4KB56fBB)!Oh'
    'WOaNUf}ar9-v=0@_eXSWWVsuWC6ZR8L(+>3NLI&J6410CUq~R5s~uNkH?9#`B56fBB)!OhWcA}~7G%nQe8~c9L2Dv*7#lSomq=QX'
    '4oNREAXyz>9Xau&ydRXuW7E2WUA{&Qz$KDaq(jn+3`kaPHJ*5T<M^OA9h+L0HFh-{ktLE=q(jn+3`kbT_UVbI*VX*sXU4R?VK=r>'
    'qjrg;73q-lA_J1uu_b)s31l4`KkKG79y_9qnx#u5tw@KY7a5SOj&0c!Pm$`_xF@T16uZ2Q8o)~=tw@KY7a5SOj;&^X8t(XDCC)e9'
    'OR#g@sHwd~(u#CQdXWLi>exQ#ryGwC_~EhXzJML#M$O|Tl2)Wc(u)j8Rw8;-iJz8JPKo4dd)L_YZbX(yT9FP(FESum9b4}wo-)$0'
    ';eA1^{2Dv?jmQ#7E7Bq9MFu3R<KqPUwCC|b5zMLX3;1+GqX!q3NLrB&NiQ-WS@{(zeuK&{P$9|H>aVfW--s-cv?3jnUSvSBIzDc2'
    'qGRpGGhDY)Z}ZdIvd}_B<6AxHh-;A+U0Od^`kzl+S9JrSd%;R?`{O-gdRfqZ%jxACs{Vx*3>v?lO5D*0t)`b73fS7n(}hbf-_Vk{'
    '&?-vftMqb%7gS%-zz>z`aU1+2<8nu(u}?0)!J-{sRy2M!zTDyO>uU^-3)72VzGdR|%eQyWce}gC2mSZq>EoBLotWxNCw_~f=f8cr'
    'e|7fw@cH>Hx_f+l`k{Ed`}pTqXW=aR{Nd?`>Tq~EboZY>-@VtbdGIekeOJXKjr2$H_);7HOOJEt@6gFv^w-1X_YZGho&6)^=`_B`'
    'r=wXqDMyoeG9P6ZlVViH*?gYrb@Hou_D>zqyQjl<-#^}k*JmM(=GCSu%X;Xex-YAFJ=~m!SyZ)Exmnlca2~R#9J*~^m$`qd626Ga'
    'dR=bo%@CE<dY#V)eM+=Z%&xxC@0=&GV*1T`XbtkRtZwwNNr>h{H*EAt=SdRfTkY&T(Y{vO`a19Gs`m$aUwM1>^2L{<eEH(HSMdLK'
    'Mb+KbQ(JMAt3i~jY5vQVyvWOVUX+v3<SHJICh0sLT_ja8N<&)Y^DD7x`r9iRhv<4;wRKnb`C1n#9_u1?I;G=PJC7$(b+cJ-%guJ3'
    'pU2ZEZ;Slj|Ek++UF+vGx>nS@x`}Re9m~aVqhF|Zb>8dZ>ceKyy4vXj^zpt*8p$+LZj1Gokq~}NMsxhcWD@1;O?8!*`t9f@@9Ryz'
    'c5V~x`lhV<o54q!q*0Tv=e5!r?W(m?(;g3EF|6x9Duw>>JjtRiSCVvQb*>{kChVAVr~G(=ze~RSjS%`@R-J+0lTZaDze?FvdZ8~R'
    '8P7*o)1q+Ricyy2=_oH}b3IX%W>@L>tI3~4i+aA!x9xGBkB>hMU#``8m_%Jw>e9Iircw3HMiugNq4dW}NyfjOz17WiRn|(~JXbQ$'
    'Pi5hExmtd9UJu@1{OTHIdR#E1^L%tMPs3<Zrse2DZ;6ejq005n%6Jm8{HtrE>S(u`)b*gm?{vB9;rBUu{fVQaPk(jv(`hI(4UMMh'
    'WHy>)#dM^XsOF=28OK?fUsba_d{vZMsq*!8-EUkJUW+>AWu5=;Gx%z``TY9jC#}WZ-QoR%7WT*ccltV^_zw~{(c$6!hc7=oKYjAJ'
    '9{Y&?^z{7v^xe-teYn4Sdw<YRXVJT-r{^F4sV~#;r{3H@fBz}^boc50aP#nYE%-{;=Wp*G^)DIe=-tEf=4o;N({9e9r^Cbj$LHfM'
    '=GED!r^EB%?m@Sa!==~%L;d!6KYV+*`{Ci^`(H+ncf%JyJRSc0`NRGF^XvZr009600{~D<0|XQR1^@^E001EXlL<#a!3O{U6CwZr'
    '6951Jcx*3pXk~3>FLY>SZDlbocx`O#TFs9ew-vt!$p1jhUDnL#dto`yr#JTQdf}B_<VZ=BhF;EaKvKIa6afn7)JuT^Y0*o8-g+pC'
    '255i=x#W*AU?WZch&*yen&FCyz5YOwVs|ey)bBn1=Do-F-jH+mldG7a3nC=tN!JVR_#Pt3jL&JZ?0Qco<4W5@QsHEd8Ba*pTNCMh'
    'a_{4h?&1R#k(eNmPvimadXZA;LA5GpFpT9Lo{|K9TW}F$1-D{ZoeO*kdSg}%eZN_aF-<&_;26|SKnBeTdbC*38S(CI)Q=d<ktnGL'
    'W=x!E<s_FWVdv+8M#!}sFo7;G>v}NLoL^4JmGTh7Qo#>h&o_Tw_1@hoKFEcb+9O+%8=JpeE*F@ehh|Q(oNg!eM`61jZ1o!eQ@glP'
    'TmNtM833Nm;7KzsWt~Bz>9@liTLNX54&9wrP_ywR{k0v=L8sXd>y~~4%JkQFpK;t74I38TfHHhz2jBC<ey3*P4JgAmckuO5uN96g'
    'ya7dwCg;28W~<$9=IAYC!P%h`-f1@d)-Z<`D2Rq_bm=J;JW=~%EymBd7z0298CEnwYMqh=o<SV<QpF`2QkgQm_E3rwE@9XYg8)LW'
    '?uXldV{v>ya7kM80w!{2Nb^B*CTOa<-UHC;l^`#_`R2v*uU|a>&x<d<c=7yS=!h;Om8r7i4so(9%l+{8KmPBpzeE51_uqc_@=tEO'
    'EaPAO^Y^d5{qCpK!%{0u@$%1Kz54pAmw)-g_y73Pp}i;Yv<#onm`HR&F41Ei!vh*q*!;;<+)QE;VQNX>2y{DCk5ps<Pu7^ZaD8H%'
    '@Tq_-;{x7aJ+ru;MPj9>W9ElZYyposXMHZ526(9FDp`24N|x@)#j1>ZjPZqQ?g36L2ai@Mq*v-vIf#fw^OWI4;U!6kLi#s;PKe{_'
    'KBLsK>oJ`PF8M;C&nW6+>NLrOPOU&Hgbry8ep$N+f~zgt96v>U&RkDDBo{UkmIY=G@)N==Gv3E5g=3e>365FW;73?RE|O<!F)QOn'
    'Qh}$JgmE;Q6DeIXj|4n#$>|R<BsQnpkJ&o5kwVdPm(C;1c^Nt6=Yt5xsgpBJqB8bVc@80p(J5Ch_%XLvh`t9u;AF46pHgDE`}K+7'
    '2_(2Or6NG|FDv2j<$c2K^|)pWOk7!}Ct@qR_5^iCM}M`nLi>mihA;7)5cK3z2ZKIOEt?mlA4HH<4v8zU9$-6G^?gDlL8h<SO_4`b'
    'S|NKzmV6JN<Mp+~v&Kn`#Xij^;B_l>rvetE<2fI(*?AdwLZJbKcySu#kpvy3;2!PMdKzJ?cIkWRj@PxY{P!kn4A$kdo32ZcAvczR'
    '#Qw&%nP6s3Pz1~bQ*`7?w;*D<zYx)N)rhT}atnLW7H)vX$m(}un!K+0Il)5k%b%zB`Cwg-Tk*T}JNxp=;k7PjAMRJ!+YIq4IVDg|'
    'b`#faYjfM$c>njU&AzpK+tA!LG`9_n>lN&SG&DsI1MN+*)nhh(;=TJm3(DA;S~GGa&DM{E*P8hl2F;$7*_hn!1EvvN>gRz8EZkCH'
    'BZCB2zo9BRi*O42LxI`xT*_SaQlgYg*fui5W=EA=(``1NR`D^PXWN}-f632GD6uNW{RV72W#h0_r80E0l@Dru8|sGb(rmSE2x!@x'
    'kSI-Nlc+gJv^ImUA)$B9-Vck8gGF1b-?(bW6hLk36?pGP@V4Id8ug43?8rfnKG%NAT)pwGcFueC^gd&<as<HE;;yGVn(N-j(@&R$'
    '+8cP<B4M18g_TXh_{ubRGJ?0^JUYw!?mrc9H}!R=aH&O4n_{yX3L35Uhw!ydBG+;QOUjvoCFl~?cnFKmLo=Lqy#?$W!)2U8@JRjc'
    'h1n90UuH^VEB<X#ODd!qVi{!?H(4k1PfQgNh|##~>F3{u14~R2$;fqJ>Hbp5g?hpL(#V0=Z7-9>g3MIe@1+rK>FgF#e3mSnzZ&tk'
    '2Wv^b0(YN9^Gh^kEAbdZh&5V)_U@cYI6eq6SI;S&hix-i5r5YzQ=a{;L#`qpW0poZFDd10p21B)+0rq0TN{+_@@EEv?wWAA)D`xf'
    '%kJxIwYFSQA3H_Sw9<+$o9=A-;i`OzRZ&GPR$1=8o^v~!WCwxjEEDfu`%Bb{$t-G`@opa0Njo34slCUJpj3UDd+fF?5>89JZRrQO'
    '>rtyK#HO({fp3n^%V*ns`V5lC5Ka+SEIaLH3EY|1GCdW>w^Ke}=a)=oi6-;8Zc4LM{+KKfonLjm(A7EShsSw0z!E^4o)=cwR?yk)'
    '*@C**<)Ybw-3n)I?6rQ?7T!v1j#;<DW^2(!TU~8V)OD|JX4VeNp4y~gV=1`WwIUq<3k<Duc9@&xUdS_7D(|%63atd~dcXA>z4{<*'
    '3@U!RF{;#Sb-&VX^lFt}qfrY+jlduF!_R@9isGP=efuzm(;&9azko2~J6}P>n**;qGagsD`MRNMyl=jO2tw;Ci0q37G|?Y9c;HB&'
    '6^?>>*b4`h!7yl6>fx|iX}4;<${=hG!yXjv=D7E{hb{~_=nseEaU-lW2Y{>ly+)<iuMH~A_NX6@gHe6x12``~SMV02b8c=6xHP=('
    '{RjX6|NjF3P)h>@6aWSQ2mk;8ApmJiNR4m{006^~000dD004MwFLQKxY-MvUcx`OtU2SjUxDozdp#MS87X471BgvMPI2UKTN!Gvt'
    '*F}MBfdbuPA=`2+oL+pAyz5+X|NVxN98s2NQZy;kr~|BRNhTi-XNEJw8SdTtT^MNJyd?GGXsTORrmlI>B3}B@%T)i@=h@&!*V4?5'
    'mTnM7-c*0}QvLn?pZ|21X0L(wy!O0IgC?SMs;{$bb88ssV(o=)dKGWH2=1)nBy=<Qn!Fh4Ch^>5ihdLZ#?Uk!BXs>pFIu<_7c7m0'
    'Zt}}^Gg!po#?Ab>ANbj8-j=R~i`$=GqBwEq0YKeZw!6@FmXi#%U9#^azuM?0^cP8-#;fcK+BM?U%3pX*xD#VyxQnVKXqm4iYiyXq'
    '+SqpqXG^w`c;9?9efs@fv<+urmTHT58)Z}7e!HX<cYa#J#gVQRlkylZflObuziID&c=yh{GR?2rkLc@{!NG;U{<@8`9|y(T`}g@{'
    'fB&lK#=c9eeao@y9hmux@a(w$?0^0;nBAD>f0yw3^1~M`gx43X`1}elLMzTekHE$N%_zxnf47RGiYk@`Gl+QO_LnIB9?j4_V11w!'
    'dhGr#{Z0Gk25`kfjV$6IPP7bo8fcYYiQLdDp7^lI;#B+OCQ1AqJ-c#4KX@(f4bg>M7RrZ1A2>6IRdiGQJ4a9LJ9MfJlb88apUp7Z'
    'RlO#@8yxG^z&q&J_=CHS0CnQ8GVR$5{f9VMwhpbmlbr4a$%KudMk1_(a!&Fq*qn5t4ykSN<nPKlRQk5F<f9wTfdYq9miJz3a*H#b'
    'jA^TNAZGNZ3uH-x{4E86^@HF|TTYM$fD89`u&`ua63yU?R(}5cx&ir*U=1upIDd}v^h@HtTEj8>57IdBmk7nnW3F{dF8t`AKh(^e'
    'bi=?PgW@ZIQQ#f+_9jnS=A+ePU|w1FWHPz2uWj?%9uJ*iJ^@leltXs%SIdyA8|N6ZA1%F|2a71IacF@NNp);`am^;8kCLd@GxXO&'
    ')ar)7-o_L3*V0>cqqVo3c|kVkak2!vbZ^Tz*B}_hwfnokTV+76#D7_%_blEZ>cv?G#@qef(sy6t$PLg8?SFy)5)4|fU8j2JE&Xi>'
    '%@zyR;pC!DYJD*N4$yP}Qy+%ZE*KlnshSg23L0MOcbVZmQTKgiYw$Am^IS^9=y7yC%%#DoM$E%Tl*Wl{)b_i|-ZS{vsPl-A{eT+{'
    'BMOW(+`K82il=G@3M*h!%1EI?r_ohfyB0Y+9HHEZ39LLqcIVuK^+)Q&zNsfkz3}Tx!GL;DAbru2Vr4W}+p|e+f7*}1NaJ3@8Cj>>'
    '^2)0gx`UA{*$3R%h;v}J(Alx~NAwQ-Sg3ocOxD#ek)_G_l`ylteStj2s3jHD>+Ft7XFEU<rQ5^YbciodToVC~;W4sf&RbV6g0WdP'
    '7H@YVB_XG4=U(72j;_W3rSObGB<CNa5R151w(6``P6<|B4tW^K!pP>W0l5n{5EKbIULNnZ2YNazyu5*Z_E@6ig;$~q;I&_l?2d@B'
    'x(+*_#rO`NXEkz^T$?2_H<NdGg0a{k{wcsDWtI$C!|uN3$*@T=6)Ap;pu1^x0vw?(TQA6EaKQ1s4QkM??`<OCfO>G;30*WJpj{Ud'
    'L1_d{AT2APWKN(q-v=aaUN|fIQ?)`D@Z8~QNu4l@7B#ixs1WcQNzG<ilY34DbhwSx*|Jw>%DTifm^q>`4+46jCUP}wL&~i1Ew8eB'
    ';gC96kOWLOqyFR#W%K(DC-6p@eF({n3oi(sQP$zVS8o}OsN!I^!YU`A_6B9{qaqFT5mJoHk45rg@dc*Pz*{TY#+w}vxf<4XtG6CV'
    'u%L3H6&9+fD1WWFn@#YFiYidKMDgt*Pvxw>e+c}SDD;X)?(f|FS8W|9{%`P8RDQ95`yQ0-Kygy$FR&|!0=pHNb|jF)O8^PQIqFan'
    'DAqdNP~f9JND!f8Am#EPiE>cxqz>g!Ay}zo0B@yapqlC!NK_VaF4-nQzE%u!Sp-)oa||?mtmq7-m32CQ+J<v4ndKFe6@}}%^qka3'
    'g?%N%z{hGJ6LlGe1lgjf2NlVj7cL%IX3vQk)p52NES6HME)k_t%PtY6QtK{B71L&1#pT1IQ1+adQEhW6S-=>T(zPhI1PL>k&m{96'
    'n^I7Z%=(?nyhn~dZrnp;F)yh*X4zw;sNth2^9taR!4n&&m_tmB6NN=J!v*s}i3DlbEhKM`jn!(lrzPK(IOM$2(WJSRL~<}YzemVH'
    'RQVKg2oXI-0uV7$OerNW%7{>KG7?G4TFi1)L{&IFSQbS~gs2cZlH_s7%#Iw30kZ|h(Kd5z;IM7W$xl%!oJ~whYl;<W_b<%bBu+51'
    'OTjS-LuOBa$CQxCKJ1<r4{{M?IjdG-QG3K-fx-a}lDuCa<Q{B4?o9;}nboydi3@?kzAh0nr>}dGp~PqwpVMhR5h9Iw``x$02nkC+'
    '7U)pRB}8WWew&sw_rr#av4r+}AU6~EhOj9D9*a$oc%^CO3g0~u?Df$QtHKnGDA*-&bR){99!Ha+D>6LpJPnT{vjW3d)6hdo<U*0~'
    'c`ar<WF{;}z??>qQ-qoprzA0w6imFFMDp!N+`$``m2Jco15rc+j~ZbQrIlDsL3G+|XNqfzR$<~1c5^ie?0cMw-Sb5xWIeGUA2q^9'
    'mzwHwiA&u`MA(t!nvWEA&51!9N(5PvM4DRsexz_#9+5W`-rXj-?lgY>zln)(hf0Z*RVquIh*GJ9HEc%$Wa%O@E{VH56lKt_V37+f'
    'Sn42D+H^T(KVOQn6CEH^HS4h+PLP38G$kpfTE7xV+-&X5XyDEG8i`1NBic+CxN&<<r%3yrxSP-3E-yFh5F4u-*A&x4Nr&K!P7E*6'
    'YbtXjBeYGlT021|g(q$9@MQ77ZWSq7px|0-C7J{g8sV!7DUr$0i5|p#P8lhSgjG*nz+!l=R~mP7?oc%&r7+oM46O^_`M|KTJD#fS'
    'T{=k7l$5<V^w@ShN{MKz{*toiBtk#H($lVjMI;ZC8t30~WMT3P*d8LXc#jnXVwRB}^WfD5EN0y!MPizJM&j3w_Kc#sm3qP3t~3i<'
    'rdz|ywzVD4=7xQL#qeSVpzux;DTmI!6vwjH<wUrVUuSfMld^l=&k^kW^`T8VA5cURydKoGOy8QOti^_x#i-{qik*-F&|C^d0+%&0'
    'M4_;Aoi8LN3MjLT&t~4FVarQZ#-8dBo3CTOz+@d;;ba|RY@2z1o~LeF=9bx|C~*?yiO~<X10!nEqHaAR$&9r7h^2Ji@gU9BSVVx3'
    '*tm$JM6i8z5V|U52#y6+_E<k6K`5p|Wm58-w;*#FZ$_^;4#`318nuS0#}YF+c(WqLlr`_Pxe~EBk;y6X<P~mmil|Q_$(18hi#Z~+'
    '7!kS5!#^CE>c)|&ZqNxZYzND7C`z<T$=rB;I5iwcjh%U2JMsGM>}oA?LvN~oLZ_VuSO@d1A7p;ia@sYtx!hIfXPfA}xXhggUVgIn'
    'TQ9)dOK;_FgY5I$&8c2}{)e~pw;}LR`P4uCZ*i74G1aTjKl?B1%yLj4kX-b0n&t1>Hu0zW|2{rkPo6%`hJzdP;bvfuyzyW%es~&;'
    '?Z=0wr`g0DnvcK3n8F}RZ+DhG)z?|JxiyS*vGzhYy$bzB5~uMhyIRDd5wBMM!ZXrM;<?Lo?Ri-k7(>&X7!xCO{Ro}sZP~YJ08b_5'
    'U`z1N)wQXPeJSAQC+Y!w9Q$PGm><TLIhc)1YhXL>&EUovjR&)_HGFdHhmYgg7{fm1!)}>|WmVwEx0xS$fgkPXYCl8xRhT6BdSnpB'
    'ekP0*ou&SaD0Tk_00030{{sL}O9KQH00saE0000X00Og+YUcz1080-502BZK0C;RKb7*05Wn@!ya%pa7b1ryoY@}CBPa{VVy+_J_'
    'X!%4Gfel(^SIEq&rl)P1^h^)kGZyA%fwLejHet*vIqors99H>B{$qJP2D5Yxk)?1l{ZU;XuimQ}e*EL^v-mQ9Xl|E_liuOMpeN>w'
    'tL63W;?qg5u15PGdZJm)7T2@S%f<Yp_vgIneN2A*Vc0Y)fxC<5q<6Dg-5>Y+&DG8PZq^(u@8=8b{I-0!o2~Hg!>4|8|1h6jH#hV7'
    '>h5#@XfXJte>b~b^u*Qjaj`n-9lbm3iO0q5?~il2Svfp>-%EzgZ8BUXMyJk-ObMGh8unMoupc(s&EUgvb3eO+Z+N7cKYW?@67hfj'
    'T%GD)qi2#+dq&6cV(IKu$;!s-D$^q^wH2MFU|y2c>wk)!Vv%Hw67dtJzEF8uyRj8hacT_}S)7_=q~s}Ebjs^Y6_q*`g53FmkmQW@'
    '!fP{Ohf{J+v+z<?kdbsrZl?f7+9L*iNtt!WLKou$k!c0T;hmePip3^%olzD?d|qpwX<wJhX+%>z#nhTFDK%CoU+Tho7Sus(`OLb4'
    'G71Q*9Hn$-1+03DC`Dm&T`&{>b;H?;Dm`IY+nOY=$9T`~3!js#VsDBia}ecna4<gUkyXCR-ifb~NooB}8|CAj_s4?)<?-q}td?vw'
    'HM+=b00=QVA3E*O6lF!JV*6E{D!2Zdyr6MSqD1~u!%lf?016vJF0qKWYkj1ZbIP$Yz>c5#3augblv)2JE2pLPKe65hsw|r2ZEsI?'
    '<^ZokSL8HLLQcw8;H?&-6YaU{VZZhy_9n2TJA6QE0>V{C6y&8UU{ZNB6b=+1lz><6FH$N`bxtwaBrj=~HsA|}U^9asameFo>EQ!g'
    'FA<a&8PtncwVRif0bREFW=$XJt8~HIwcAi>Y#v-<X7cq{ms<l;EUl9$@Ye$671B=B`rB^eL8eB}gT1F>QyBIL67DFspp3mUa+xjw'
    'L5`pbi7^aT&&2Y=0v0Io;#o_0mX&3-lfi*Aj*B<cFm?jhDc9$92e5#JEJ#-YWJv?uR1hoXQ%+#xBEY_X+U7!oqWG~)dFgZ+%*8I`'
    'pdjMJidM>4>)RCYo3se;=to_)v%!l?j3zCsg1UfWc(^4F4+o5Z>(qcz7;OSpfoK>F8&gj-30SN-YXue}A+?jWVS1T7U2O%73l%w3'
    'fi<H8<6L&pT1$^4j*bT6=b-AUVy?c*B=pW~YAjUIHCrW>2F92XpqjvZv11cVX<bC@)wp)hsf>+J(0}j&=amZ<a~&9lr2@c|%#AUP'
    'y6}n-#A#z#FjQx{ORasNi!ijbgBo|;VzOXug*x)LF%Q%juvoe<*73stW~)2#>C{%}SIyd}<qp<ch-mzV<E1rxG;y`&csJp}49pbd'
    'fsG_B1F@k1gi&gB$+Ikx<A4FMic_0m5PBbGHYnutG-W77d$DHSvdFFzGoX)le46Cq<kUs*gOv^dJ9u#|+U};+?HmOf6#F8mWS!$j'
    'kn9>w|13eExymvPX~kcLgF8L89R={#P8VKHCl+pPRDLV+yII)^HVkZj+g49^av>5nmVyGJyurj2sh3c6e8g<DhmmigE&#>X{52IL'
    'TiLOrQOEcdeHp|Cx1T}r#`j+<oZ>u&JutmQ627ZjA0NZXzW@LL|NjF3P)h>@6aWSQ2mk;8Apllq_;te^000)E0RR{P004MwFJy9I'
    'cWG{Cb1!6aVRvb6XE82#ZEWnFTW?&+m8IVYnEycVJxX!E2v#-X+A!d87wj3#GmA=<2=gM4w5r@j|M&SKDV0sR8#CQcCk9avN+~9p'
    'JehgU%DvZKd#&I6(=T_o7oSg$Pd5+u?=IDM^77(zfBo?O=KiO5m;dX34iPUeo}RDn-(TH6+@IcE{^j&^`A^^f_y6{rU*12y{rRaq'
    'Uhn|-PjC6>yUUNy&!67DdGmDr@pO0f^xea!(>?coczC?Kdgh;xKfQVXc=a<6eRumtOX=S{eR@1yy?^?6Iz6|SdoI5`-qnBM(eAEp'
    '?k~ST9~sA5Pq(-E{`%v?<I9&mJU-mL+`4|a{a*g&&3p#)=J^M>@h=}fd>>uV?$^8KSI+l7KK%UkC+5wsU;E}Kxa<5A=STkAuRHzn'
    'd~y8?FRrmhYIk{Y{g-!_mfsZI<(rqs`gC*sa^wE<4>#98d^zIl-+%t$@#5zFyGwm}aesBk%YSos_0!3G_v!v8p6JzE9{P7r&tGm`'
    '{Q2?b-Q{1$(8cv(#o(5(vRF~NO=4&hyAGpj<Ns~`wDNDSdHJ7t=YRYD>&veEKfLgFH`k93PY)lSzq@|8d-L$&!_D>S>l?}2sO+2b'
    '3x7V)zmC#E8)b+tixRfDEK&<YSZ!qSxfC7Bzs;xp=6v*D|2{v9mv>}7=`ZgIH|MKy!CCU*^!SI<EeHGg^oP^a&A;+m$n(R@ck$J$'
    '_T%l%r*U)3`+D_u-v08;{+~|ImnYXAuK#>@x_^E-M<36>_Tm2N<ISh1%Ztair@J3dyeE&peXn>gxj3FVm`{&4_s{cyuHHU9Kc1ei'
    'KhB#UIF3J@uAg6?@2_{9pXlF?b3TQqPxDz{z5V6GBWLN=TaM&{bLM<LUS4p*yxcnfg8$12#Vh;uQ9eCBJ^%CR;f`1E>6t_3P|oMg'
    ')n{JtmxKEHE%TS&Kg<_ye*JHUbiVoP`FVYN!wb`{p06%G{O;lUPf!2*=K13v=k@#d@5jG$_5FW#L7HVQNNJ0YGJBzVsa%G=itKgf'
    'mQvDs)0GadF39tB`2|aq4MPX7#ZV>+n_U$tcs>+*uDW;`bK`1$!4fXg_pSfo#Cd!EL;zp@%TO{O3n(kPni16y$|fglGD)Nuvk1M#'
    'R;+iiyA3AuTHRo>@Ta?7228q3rkOEB8@^Lp<Sfe2P6U~BX;!Kr%~8Q*n4+>QMU++yl}T434oP_etVP*sj<Rns*<b>iuG}inUo6rp'
    'lp12xrU?2X>XD|p$Tso6pkp0vgrdh$!6Xj~5S@+a#d7RU2N9ch!uOVYY&1nSZ!p<lvN)Ldbt)4dYOcu$soHE6x)D(o_zxArNVOMv'
    'Tre56)Ls>pNf}gC!6}iF?4tH$RB$%OW;d8@Fj*W-!a86wLNZbYF{GEm_T)u`=oHzQVWO+Gx_H#eq*@hIO+i>&Az*8u!_nuC4aP?Y'
    'X@lQw&L&_dny%dDFqT@R&5=162L~$*10Z2NEG+b?7FjCVA1~*)=BQv|TJY6KC!+U&iFe!@G(6r~>3Z`vjZMIAFj*Z;(lTHYrK1vI'
    'q$|6n&G8T?u`r*hx<Waqu2V_JC191-u?<mthci*lh$LebGOL!IH&xcxe9y~>c#nqnXm};9hC)KKNNJ*7RXY11a_MEFNrA;QN;s8t'
    'fX+0bzZPp8Xg18qhCiW**26?XCfrGCBZ^MLDCKKS8%=g;vU+Kv&={@K=p=7O=}z<!K@-!oC@N~vyy`VF-%<Ur!{HEW4RvtrljpIG'
    '>2^+*q0TT2hHY<6ZrR#KljYH59U-Y}K57?}C{ZIHwX;f5wTt8od}b_EEcK|JPS>&fS`vh@YfOiG<Jff*P1WeE=Fs;*=S5?>b!!Xj'
    '*3>#GSnqw7Rz?v<<BBXXG8OSgFf4|6P_4)EdMa|HPDdqlbrK}Zg=5c=(RP@aIHZn7w$zf`<KaCXUVNgX){(&$rJd=y3o|71VUGf~'
    'G!!R2r#wRL4vqX#Ym?T!3?2NdE+OeR!6UevEUZgD^*n40TWqpBo#i{7b#$;g2IaGvTQw9Ywn9K8Xb2ISOIi)j^?F#|#H5szl}(H}'
    '(zd2_lrrE-a<0w;x2?S1F4{Al^YZ9k_giZJR;?q1Ek1=*{G5iE*)1UbD#1um>j+k91Yz>^=lg*=odjKyEc#tn0#X?j;e5C$LN&Q&'
    '6|%{jh26Eu^0mo2O4zErjtqv`@T-X{646PXB6TLmJt)TxallOA@7DPb#hY}by15A-EDL`~*+i)=h>0;BFhsqkHgB?dlht{Xb)>K{'
    '6jNiRHpzUdTfjr40&k*a@C<Ysl8(db>TQqNIu@Qa11UNkj=hAeM6=pzwsoq_8%=g?vV3i_kQO!!Uos=%T(BguFeS{2XS>7Bi*r~K'
    '*mPKD6Wg81rUsPs#j|0Ax;`a6EfusL*}Iy}-qP7<vO1dRb+oXx>83Lxjp8PnSdULq7)F@hW$lfN(;rrEQd~f`8xY4}K7igNvZ{y|'
    'i$K!W7*x5K$F5B_nk<hd>u6zPb+WVdhvaLZCk@SFM}TGnCr0C22!PFDm;Od6p<qRTivnX}3?!RKH4g!w%*G<mf<1xFn{3`>ao$9)'
    'qlHzNs%92)1U(JAE_Fnd^{I%Q@d@Dq)xE!V8#)jR+mm!MIWN4|Fvw*{y0nNE)@!_<s`DXir=(q*EMJ=}<Ou778u*l9l|;M4*6<Ly'
    ';ha>oR~2<Sc0~H9GaZHdK-Mh$MavSj^$6FZi$iFijIq@QsbB9uZ@tOZn=I(}zK#~wu^^Y)m_U;07BG7kX`o(fIpbu8OGy4^%6g#M'
    'q*a~45RGuR((cGWr=tUI3e8|#k8d1XJG;}_ozCi=&N^CH>`kT-GtM9pOI1dy7Aa|Eea?H_hq_WK9d=Kr1(&_7ph@t+VQ3O0jEtv-'
    'k-g!m(9>%MdDkYpHd(MXSw{;yM)cJhF2M=`OKlRFi`XhmYMo^|HH<#2vR)NsM@G1tC{m_+js1X(6Ps<VM$@|W-?%AkG}&mfAeyYB'
    'g+-E~=xGAo!nA3Qkzg&~x=@3qV$WzK<-_VtMzOtHyw?;?jvTv&Jhmk14CM-|N3ngo|Gd#;qsfA3vW^zE)vkS;aTltvL(KKZOlJ^='
    'K}ud~rlLT3t=TvbZxVv^47e-l7YZ$`S9nZQY6;gvD@o%6<=f^>HgB>zZ?cXSRubdKPD`J`T<|6xBgEX*4Cd0w*5XFDhgH^N@-%W_'
    'Y1q}Ng$-`54}&YlJJ0SBQ9I@Jj?u17c5Sj=ZDQ8Z!jgW$$N6+>4sJ&GS_s%i5=L_}=oXAh^{{B-4Soaz+cn({Xj0i-ES2C;V8rXH'
    'okAVEYm;4@EMA+Kb+oWOhS4XQ;4RHuBr+z>9%X5?(r8x*6gG#|n+!!r;iy$i16W4f#F?XjO+wXF6fr<jx~*Q@>b2F?Yi1oSY^>Ng'
    'W{47L9akg;i9mBw1$!>h+giIXuepZ<@g~?j1?`FV>oUwZbyfp#p|mt@!QjziL)xD8+ZXoS+^w9=W!BNccEYWCMx+7LfJex@B4cB7'
    '4ue%DIvqRva9A|)3VSu$n(3B=vYy^%W&vRn>y-?nGWs5`r`wx1*}Tbuyvahou<YU_e`ZS*pYJ%qqE44qXcKe6-LcAVoSF|*n;=~y'
    '0E5lpF$llF(u9P<HW?H-@HQDZD%j1!?oMa<PG=o0EZetUcE&mMG&3)cHJ6NrO-M<UPQrXBjpNoP#gx%9yMD?skMdY5B#FCiCIqNp'
    'I^9?g*evX3VOM8i*U`cn9jiLqs=CQ!%9=%H8p58Acfl6r3Bfq*)*r^O0nwUBSU$nRw#0<E6UrXE50L;2jI%eIY&2ONP1e!E&J-(l'
    'TV_Rq?#pVd%aTJ2n-v+Hu7(oTalO|@BP<!e{UR%zPhs^hI-=Q|Rt-IyCpdH`VC+t3`A%mYEo@@KI=CjMzy$80q9-j4e-rGQjVjtn'
    '9}k;quCiyp&vnr)VqTk^1q3?<N{$NmwUVXZjszP`mPeCyw6J+N#d2JpXaWmc$kQnrgMxwKv|$lYW_4hfX>v~iL(ve_gM{$0X)?$m'
    'DdRJK`(f`8?+Nkp6JooL7PjPptRy-y$mn81j5k5?sA30sG<Y|=;g7pR+*H$YAb#}`TNZ$?z`{bKD7J}9$;xH9_jGn`vU+V|7t+Ea'
    '^BSq(Z!()4aD_b`jlrtxJ(A5O;Z+XH!g?myrcjKGK^QS3VZ2#NSE6F!*i|B;>zl2BtvA_vllAl_b{#D&QVos03F>rM%;|5E5Y^K%'
    '1jVPk)v*(|bJUrR3^U}EH9waE4%y6u;bYId3FDlWrghxbo9x<T`PyV1Ev#w^cN58VsJIx<yDphLaM5~HSehvljyuy)Ng`W|Nd!lp'
    '7&(Q=VA&H8WrS9ZQzn+Z^|#Svc{Eu^3)>6Hcug@Jrd{BF61j$)@_NQ3Gz`v@5Pn#_iED(2&A?M5Ird&0Xu`&h=@K@Xk&f$oLcGyr'
    'c{Eu^3mXgm?))Ky#-Lu)#9)wbPP`7nFz8v`syu80hU*%%!)4za1hKHN5)g!n)!CtGYoi)3w|Z@>*A`c=*>$wA))zl>{bs@yoz5t1'
    '_1lIP7DA<1XcJ{#f4(1RPbWn)v}0qjJ9YsWY2(cjhbB~~hPm}ee(0X*?Am1c+GHIqYzZb8a{dV3oS}In4~uCo4V~$`>dx9U4x9Bi'
    'Jt%X=$u?usLc5~@l4e;c&ZSo3#dIuNS-;U_c{Eu_3)>}KfegZA#Y~4B8YaZhnYgY+)3KnI{;<k=9aA@&6rG_gU7Crw*ITlH5Y6D)'
    '0&5c5*qzSqbXM<l*3rV&UOIjt*dt{U{?4;v_I8XI2GbK=Vs`DYXp*$n5x;sNv2uhrp*Eq@DT3oKtUwoS^JZapZL)f8;ug}vDqVa<'
    'q(QVf+=E8^6$_q@%3}eaM__z>*qKhKtSWN&g!}OXEGgy;he<VuJF&HuF@oA?ve9IDG+9Rrn^M;2kM^3dF0;l2JnL~ol&!g#y(WC~'
    'uxOG>>waQLZga<|&K!?6Fkl;`Ul6K;$;no)Z8TXPP1e!E+A`S@iV<ID!j-OguMxN@W}qmQij=w>c0%0P`13s@gH@fJLft`=+zQfI'
    '!{%bgt3d7@;*BQDqscm2SORENABbPc0O#}FI@D`=IuZ<7gXGQwyCtDF8H}8q&~V1oRrEYK5_D|j0x^ZM=<p2^?meBor?cLkj$20y'
    '8w-m%rAJW%mPzdCFcO)mf1TJX){Cb#TREy0);P~Hj;_>Ia~_3Ly}%TkWHD1hGunF{x0B<pO_r}s*3rU-&=`oY>*C>E`FzLxJxTWa'
    'TQ(s#QvY&2Y@UvG^feZ@B#a4=-_^p3jtWq*HW7vzT9Rt>CL2wbN0W85u-!pEO#d=f3UiT+JqWZcMj@VpMNRV|dsyB?;g4+-6^RSe'
    'Z-1sKlUPW+PH>s9b23HsMw5*u%cIFUT3GLDl1@{MOckj$r2b_~7^kWXMRa5VMjh6(zGJaZXigKMPxTsJYiul9Rs-;}ZM~BKSa+Vz'
    'Mw8{yWF0N63|KVV>_knT3GpN#3W``9hn#82kc(jt>kCWHsV~7-fy)UqosI!z8V2t*2L&5S;H$5v+na^mEbM|T>_Wb<5}AfiiuB>>'
    'Z!$1(Omse?IBZsJY|}64s4T4Z0aq3WgA5rvwXjwc<D|g%gl^5WsP>-D_Jv*E7uK(%g;hC_j5iRNXn@3420MsIqv+kq9J*7n92ZTD'
    '9%^QGKqI5$WknIaiAFj@WS21#5!*M3?7KGEwaI$5iC;$x8%xWLj7^~{je;Gd{uQ=wMM}wVeLDVr?O%2vM}lr6hWoL?8xzLP4ggyL'
    '9lNKhgUY?ul-u4R-n_}`yoq1P7dC791pG~sqQ9~5moou_t$vG|WA5rie^`H$XoE`<v$alUGAwKX1bgJz!;pGLclfe-lZ__Jqscm2'
    'Sec5on5*LZJ`=8RQWX)eSmJf^AO(<(|8LSak5ro|(<D}tG-v_!8YBvX?aH1`b`>v28AI3;;*BQDqscm2*w8(DeMn7MxFwTDEBIJD'
    '^CXf|Rv8rt4?QXi+cRVAxjvi>gB^_GFnNlwu)5~dk@)tK_73qzljYH59W5-b?=!o%iC!6wu<OD{`!sH50hfs1QnONrRj=8=261L%'
    'v0(>`)01zg*H|g84j?4CL+P}Y^&3r=N0W85u(-h{tcAKTHdB$XcuP(f;8Ic=x+mxLH+gvn>U69hK^sEFLT__pk*$E}dc&ZFuB~CF'
    '(9HIQ-Dt8rnk=M+&9z0zicTeghMA6u7$JHj>eYlE9hUXxxSfu)(m938J}Nf!CXmM|PDcXvbIKH>bhTO7jV8;Z$vRrttT0G&mCQ{7'
    'L}5$(^sJu~Vr2D^EaDKuVN=bA9|(3sXC6fuF&WB{E{!Z8pW%?FGbKxm+ZT3TOh4Y<d>S{mx8HBQ$%1;5b+oXqpshEo`AG=@OMnhW'
    'h?9WO3Sb$;Vn}&d?=>_12U_)rW9E(#8eV+kEG8$_tC1QwXm-}$u1!|2O~N`_Sc1!ga)`fFPwcvYEEpjo19MO&21>=`epr`jGR-+u'
    'WUvb4n4Hgaa^%>3i37KN4V|)C*o`L3qsc<PuqqQ53fVCFoDfSwb)sp+TxA_c_E8OYSS@U7L1VxY7zyLJX5oL3nk0wcK4W9$WF2uk'
    'Ic_vr9!=KK!e-nK*=aFC#WcOwxGba8S=8u!<pZ9e-{Yop8E12LMuZvjXW8^)rZX}*#2{qAdM)eS>1;Gv9!=KK!U84<$D}UU;Gd2|'
    'un<up17l$)wSpJCsfXoFd_XEQNau<rQ=5=cC`y@v*l9=r4cpZ>^QgV`w^`WrSlF<R7B*9W#B`R@H=G;;bCIIZu@P=XkCRHbp&wNX'
    'OJQP28<rs9^o5n63F%ywp%gYs<em=VPxnk`&vX{g`U~r5VF#83Zf5;uGaZK<)16pJ)<K!>G*+<<gVm#=Nh<`2!ITK;B3)Wxx=p6C'
    'M~Ihl8A_VI6EL>F$@2atVI3`O>_(Z6HM0wT`kTx=od6Y!V-mLS)~jR>yF=_IoY)dXuojnTo{C6~4CUUn(+%7D@Xd64cRIV%S+LVt'
    'M+?gbxRSkQfy4Wp=}5*o78A43b4#NIH`-xY*n!<7N_xE4Vnr+CE1iyqONRU*VT@TZF~z&n*`3bnoz6O1ST^I}8_=w2KmAQ;5=64i'
    'PH)v}ny#|P!?LhiCsn6MQg4#2zwQDWduSAEGU5}WPdWIFCL2vwN0W85utDd6xoQ%D7_R9^z`|c4IW*g3)fffO*u%20#baEn0FvG%'
    'GaU<><l6DHmBq^fn53Qcw|SH0d6T$~7S<Fi?SOm0BZ<#<i(3dz&FD=O`mZMYKWxq&Xooo9j_uj(6QQz5Y_*3s@vK8z%oJk8X1eyR'
    'UfX(;)%7NEAzxV1L9rp=Q(Y-~gh^*I7H*)4^d^wx3)CESrW0}=S<;)SI5{7ewJ4s{Q_Y}T2g}WGvVb-VyII)vSlGCZ7FIF^X%699'
    'N`VRS3^Yd)Cb6*Zs(Erlaf&@EnyBhriUjDi#3;&EM_rZhtY<{kB=m7k_H~!G(PX2^f@rdi7PhwBjZ{XAM*1*2#1>_JW<oL~OFHc3'
    'CVSZ0geV|b*w}^F4fUGG#F314NZHtMK0&*t`|EksMw5*u3!=$7T3FQQ20M@<>KLR*Na)O1jTJ5z*HaHuQF>S{toB1sZx#>*-o%hd'
    'Tzq9h>|OH-cUv`HC$eue*=VvLnyjOR?N!>Q7=(EC`-N3=*F|B87>Fe9E~!X1%~8>WK=W<_OT(p%=}4f+DXdkq1n68KTUik}wkO1U'
    'LcI8dIIg3Gb*bwbh7nEU-)If$^wjZmoc<t~6nLflVN=ar#WO+AJsr<_5%7mmQZ|??pgK_wC(2ZJr?Y1|%g=Py(ZZ(8R*nQiOlWjE'
    '01GJ+P!XA}m%z}qOC^t5n>bs2Z%FMND~d444M&2g8S&I6YY8&Bve9It$?|BjkS{C?bCGE<1J<e3&?M`W1{K=~<HU(b%g_(2UdzQ4'
    '{H3d?ZE9s|!PC(iSXfh7-^tS8O0YYf-RZ2}>8zuLWtyr{(31=v5>Cx78MgZIHwjV^L#zfy(}%51d~%Gm>2%P(0*8^*zt}0RNDd9G'
    '9kDYYUmyEsVK)oA9t)e+(ZZ4r%DSJcLWT?rTZUzGu4`eM$>jVdC%^f8KadtSG2QlqPG_2#z{2L4ECz#;ZS?5s9l@d7vwn9vt9LqS'
    '9W5;Rn$<&>2#G5*PL8bTL=uRJHrJV5+|X+=N99e*!1=W@7tWO>Xlt;^F5XN+8Hh|-CGy&Lr?Weq#XFs}juzJ0F0)}mY!lNEJ3Sq<'
    '^DOzp!p`|H`QXQOOG0mwQPD&91dtHf3^ScMnW{bZ;+g1_y+gcfljUoZb+oX^zfh7oL>h?Er+Q7(H^d@5z|v}oeCi+8k-*u}s!bS7'
    'r<wIvD$0|NasABtBbitS=eCpMMw8{yWF0Lm8nsgORDi{~hwB&<N5XuC@y!<2{=k^S)+R>AtQ{4Ji;SXdXA?T9)P?h|I&BJ3JKM8<'
    'qsj7UvW^zEPV>f!iDU8jGb7-@d#&PRrgb-%BsxaM$IT(GLCH#kMqnW`oix}AC~!Um04fWItaaEq#2ZbPN0W85u%Y!HJ)(M@&U%wd'
    '*5B0F8(oya5*`iPVLKg6VN<b$U-TH6N3l>$PN7A9f{l{Igfd~Zuvys6!miH3uA_y;2{EwdQt1|m(y@jQc|29CNaTa-Xc+gH*N#I+'
    'a!aZZB+fA*W&lN}Loipww4MYEhjS9MV?Xwu&aO@Vu{P22#gcAnPtFP)Zv4xK58tCu($W8V-~7_~?#G9pzxV0o>+|;4FV1g3U!mNu'
    'cb$La{E~loe7Jl0#?vp)7uUbMyCf8XRPW1+>%aVZ`_0#w`HQ=&>&J(8mzNjMeEaRq{hzq~a_9Z$KYse*@yl`k{__uy7dP+UU8Y|J'
    'viZ-?4}bi4_33n>`HA!8`Mam*FSjoK{CM;3@~>?4vc!w}W4!G47N(yaW1yi;>^h994Gbecqx{?J$J5pG&BOg~-(UQ4cYFWzmM^}$'
    '{P_I*>Ft|0PuCw$cUMo}-Q94wPY)lSzq@|8d-L$&!_D>S&HKlzpKtDedZW>*y}7%(xxYLg>R(;YDR%0V2C0HB4818fXHKY+B9}0p'
    '1@<++c_56DBb3g|GwBnycQnJ}<W&?J_uEjk#{ud;qiK72E3e)@efr(Q^`D+D?jP#MtNWi$`RUW?`uW}E6a#$y(j%wZ$J5pOr}<^h'
    'l=EqI{rfMc!}(u-e{MW+mR$V!^Z$5we|mR$#Yu5K8-Drl$Z2@>mbd7Fb5I#}q2t{A;v}5^Ago=^2Y-3V|KrmXsovLz`t<np{LiO{'
    'yNh|_-Q^<(e*RTgpE>X5abCUs`YrQu?;pm^?d{iJ@eikFl8KpFoNhRCfF@_o>|m4I4@HlO<6WI})jT!r+mAQz-=FU1!~J?{+W+*a'
    'DPK;_d1y}$w>R(S!=3-}_|uQ|_VMEL)$O}Wt`q({zv{2P?ak}MI6qk~*sp(|pWXSLneWDzx9sWDmy7rI=5+tuuAZ;HzL^jAr+T`*'
    '&G*+IAHII~&rj{~`~L#~0RR6308mQ<1PTBE0000003iTW;6j1UIRF68IRF3?0001ZY%gtPWNBe9X>DO=Wic*rZfA*5PDc$28VUda'
    '01Zh<L{b0%0IC200A2_J0000y7Z}<A0004VQb$4nuFf3k00004XF*Lt006O%3;baP00009a7bBm000ie000ie0hKEb8vp<~V@X6o'
    'RCr$P-3hqXRkc6x)4i3m*yMzY24ZMxIgf~<;ed(*4&fXc1S+E=;sB*)PB`HRA_9sSs33@nnWmB>4&VT3nx);Y4Q85k`+wf|d))r+'
    'VNiQ+ar{55$9C(te`~M3_L}$mzDFDW@sEG}PaPly2YLM4-~I-F|NZwLee}^Ik36zr;qV{-@gI#?L&Mr@uYJ=^H-*CzXwFbGKkF8I'
    '7bNi4zy9@w7hbsj`s=%O>&B>!=KUyGbImo!j2ZLyIy5p`0?iU?<_~a!ciwp?_<QfY$MCPe{`$m;6T5fs-nMPq<(6CSv!DI!XFl_p'
    'Mhnje|M{Q)>8%`KY}KmOxN+mcY6&!FXpH_m0P*_v+ix@c`0?XUI_acsw%O*(U;c7Q>^vaIqG9#bS06QM6bDi}8u>1PW(f`NzWZ)4'
    '5zku288c?|?b~;gO*UC=wbj_0?Tgi`1-=Q0es43bwboi=$Bqr#CD5GV!{jgXKmGL6=bwLmhYlUS@P#k1C_|QP&#oUc^Z~yS{nezh'
    '1ezx_h<Nbw%P-%2^Ucb<vK<qC?sK0j$Wlu!wTP-8Na5!t@RluG-f+VWOZvX%4-Jg;_~VZc95`_O_19l%rIlj!1Ug&BoWVa%^v6q9'
    'U3JxwBS)(JP3{p(pt(WAym|9_^ynd8nJxA%aUXlLXfZsz3Yp26wjiT!k@<07k^q`HG`OMZ)MB~VGQbZ$|M|~1KGdNJ!TeC(md+nE'
    '?z04%B{Y=J@|IqD=^%;V0z$uP_29_|Wz3|7$s8Jp8*jXEN#<|<P%{;SY3_`mSej9*ez43k%joBURH0IQ^{ZbM;;mb^4sLRtb=HBm'
    'YuD~eU;2`Z!{wJ>&itCn#(kDRvxJ&ynLJjncKgPd-z%0r_~3)DxZ(=8g9#HR7$;7gsA}AH+ik&3PMtdS)?07A{r1}j4<5`LuD<$e'
    '+~0fey==1tnkUqZt<1;FO{}i=>)|_h?mTSRu&17S>g~7R{?niS^yfeS+4z^g{N;~-{Nuatz8l=6w~#;l;SXcRjFAAYyz)wY-;&m^'
    'nZt*nU%(>|VEYaoI-Gajd9H<QTt35*(o@u9+KgrLt+(D9I&`QDq~yQkW1r>^HDfoLy`h20)5j^~0|pG3K7IP{fB*Z~HPFPOd6khL'
    'K<3V!+oekvErLo>Cs_i`8XC<n+?4?M8f&Z}+e_fDz4n@9UQC-N&WtgWaRu`p18Tw8v16BCe)*MGUU~TN;Ryb3lYe$;eE2|&e<w5%'
    'V1ZCgic!(WS+%&J7V!!F7A;zAzy0<XTyVid4?XnGJMYBgdbfz~2JbM{P98x?25`2YJb7~a_U*A!`>(m?8uOp@*o&t3HfjyNXyU~M'
    'kt$=<)jZ`4)C|AkiYsom*=BwE^qDkilHzTamZN9~Q22}0A;10YZ(n-pr6-?!a`x=mPe1+ii!Z)7Yu2ns9(hE#@44rmxJe-{ep<ic'
    '91DadO7R$Id|nWEORA6s8ejY?p>e1Aq;13&0wKs4O*T)ceLZvJkw@Ni(@ihG{IUi<HWTZjIjKJI{PWM>c;k(X+OJ=~{rBI0*Ijo#'
    '=%9nTcI~Rt!#j5DxbC{^>KpQzvDVMUZ_yw5eNu%6kbO+b|66FX>7r$*%UB~Y*5jxdoBxtaE)m4}oeyz&fDk9K8IQc-x#ym{;f5QI'
    'KKkelH{7sw>(-isrIuQ1x#gDo@|VBN%<R4N(n~M9?6Ua;Fr~!&s8ORZENSCk1krMH1KJ~4WNV%h7HDvN_kZk^k7!~rLMP8poH+5@'
    '-~M)+Hf_H6#V?jF0aWh``Pe}7n2!!2i1V?}Cq4Yfp|8^X8<hJ`Y>_2Z*Z)_C!9NX+Klb9VL^LL2OZ%?-XP<qxo?wL)R$%yi76cd`'
    'UhWH0ec~R(4<rWWzwENhK2FlBDwhBs`4thKq(zy9{|58tM^KFq^&LNQ(-2fkXZaCjbAy`sBPh+P_mnH>Z@cZb0}eP~z4g|MUCXC|'
    'v2?~5I{>e~;xX4<cil&7%G;`ou)WgL_^=RQ3HAV4xM=?WCqSgCQ`XP)(dH-K+i$-eH--P<4}WNU*UX@1{v402{6Hr^bm&m`H$lwm'
    '3Hs_6LWz50$p-}~WFNEhFM$OQG5<5tS5!=}AjVJ{3tQ5f|NZZOpH+cSh9*lu6Z>UE6RQNrr>v@4%Uku4NWf!3!rWT`i44IJKtc&J'
    'iT-!L`yC|0)Rxuf7cVpvA!zf#2Oo3`SbOcY6U6KqV*`i~FP2^ygrG%>7MEUnY1l3TifW2Q0o-I!oH1rDG!$JmC$)x2h%*U7ZzhPh'
    '#T!Cje5*PW4wf)E`mzKBuhPWxM0wtt_gI!xS>?mie3}CIeO85LF+xLBNxo~Zy>|EAcV{{Ft6I8R_ye3HZc)e}DG-?#{lkY3r-(oO'
    '=}*~aeuyGyhb}@!cfbDiukX3%p1be9`+)}@c<ZgV5MYw#l=`R)y~)QOdo15BKJ?H-b!PtyEWs0E32<c{H*OsINs}h&WYsmd-09P&'
    '|MaIn_3HK4Uq>m!V+%3zwx9j%XOkyS9x-CXz4zYx=9_P3p7H1^cVoz`S+j1r<rXv}M~>7dp$VWw5I;aXWy%zZ@T#k>y5o*J68+v9'
    'RoH?rMyMr!+IO{cuVcAbF;F$UK49Bzx7}*1tu%v>Ew<QV(@i(sW}9uc+;U6E)?06#`R%vgew}sJS$_HDMSmP?VZboPC_V8_W!GPS'
    'y<5gtzxvfqojUP=I^l=k8mF56*MI%jA%`5o!-&+gXHPRAd;mVUC3`a%o{Y5e$}3B08q%+P<ty&1ZmO2AvdSvo{N^{ib?Y{M{CH--'
    '#99{m@BjYq!w)|^SJ_>5+2xsMp1}(L90!@v7C-vYk1$zr#T7T$V1xVayU*J~gHL3B_St85+G!`(V4CdEp~LLivojwii`LJ-2O6AQ'
    '9I&6jEw*QM0i958*|Oyxd+ec;pE`Bwop;_TI?)RYF?#>~_dB^r`bix!f6_@OY5gv|@WNuY1wzbUbV>}W@7%ewp2Mcpde&KI;ZW64'
    'mPk|`J9ey#*}#4#twmWP$xQSM#NU`PV^F3q;g^hBjs{s~nPpnHZcW-`drO4-o_p@;6-uuB#1l`LOib5V)_@}p9y~ZVVWRTd;(tL9'
    '<2G&DwAJc=D}jt|r3&*Uose^}dqBgle)X%Pk3KruiBd|$b6bk0nwc|avJDlJSUg4DA^_q55F?yYB$A6Rx(GKT?ufI^{1AmE?g2#S'
    'AwaV=?F$!yjwYKT$U@P%`R1FW;V6=X8Hy%AF?ql$Bcn>r;Tdb;Y%E=^BZG}M-k2W~Ny2#PrI)ySD9u0h)KmC4Z<4=VaOPD*h73W&'
    'w0zy#`ko&qa|oIr{_uz414uh?;J^rJ46@jvK^<LVjWyz-APJvF)?a`9e*O9h=-40<JAc#)=0^Yz;7l^-1s7b<IJIy{pu|P!_w2LJ'
    'j;07h0M4hn+Z2BSAP&yMSe=kS(z4bhsB!!t^h2CS$v;_*Q?9=H>ig`o59=Oy;DHkRMjLGupTLSs={C@l#HBa%%FxqKKaEeE71TGW'
    '*^jgc)Dh+%J$kf-5fTE>$Wq}Ga4uW?;upUNfgFI?USYVnpkdUgQ876Yql;FpS{-xDF?t`DK<A7e$(jWq$D1GF916DLDj!~c`Q;WC'
    'f(1gfnJ<~dO1?3QB+O5(uAp9F5?0|R`n9N$R)z>cUV#`OIC+GuwkAAz)p_Tg_u6Z(Nq9h%&zUnvdh^zH3VCiT^4zYw?i!t_>rX!U'
    'WHbQ6*(0;5{gHOnRae#4nU5`_C|f`L@WYasP1TnqSA{_DS;Q-B7b7%u@7_J<C!gRD6D;?_3oo#Zd>auR39wKc+)Q!Xn9z|NSfF(6'
    '+7&A?QCh!p+JHwqvnr4+G}3L`wiO~$(GEN85X}lx)+&5#jhIEh^{sEAM5EcP^H?xgBDHAs*#JF@o_5-4nGXSmt^@c01TMX1w2em{'
    'b(GzKc9{85Muwan(9FdbUz~-q6!Y7tymdNT=v}<UM2?ZJ;M5{DQe6?Cs)xk_wU(Y_?t4<xph1Ha;BwK*=rMS*^lyIi8}Y#^t}fu!'
    'S6_8^5=rhh`|i81kj2e-_0?C$TIHw#kJWjjdV#ZG;x{a$RTlI@iaq3BcQ9mB76Sm$FX!7uNbS7y&Y4^wV9r-=yzxfFKuJ31oO4V9'
    'Z!zT1u8yt)_7$W1?YEy-&MOZkUcdhJuWS8GR_#Q7Y|l933=(pzFMa7t<=<uDTmZs{q!6}8aO|<iy8Fci`RlxS^AOafsEJo!x7eVj'
    '7nh>my?cv;xpU{LtmUY&cbxi-H{PHr7wEoy`wkc|06j!=@Vz=5al{c{``Xu%LZXym!-kb;XaYHT)D(apeDJ~6>W|bNS;Pg>BOAyf'
    'k%y+N3L-_NcjK~k1`7s`Wz;~_4q0#ePd)WiQiTtI=<MWVZ;370nB*2Jf=65E{9k$H6-)RZur*<XJ{oOZS_)M|0rkp!IHPm~9_%gW'
    '&_4-}HP>7-#PUBHJe@ikUX0MdN#!v&RH4u~S3dyx5^#LN-HFNF%GlVc=*y{Grd^#?dWcgw8+*_iL@=fHFU$i14o{mDH)F;OIp10#'
    'D3XMZ3q2pghc}=pzCt7pD(JZyXy3kl=I4hjL7$5sU@f<pOd6R=p!4bEKdyGjA%~bZDGg+aqg`>u6+iy*k2MhPz07R9=9+89jvcGa'
    'At0mzuTQdF7~(*C33mj-oxu8=z4zYRBbmiaZ0}sf5f>LU$o6`BaztP8$B0E}p`%>sfY`gHk2*L*BUT4i2N1@~jO2gWWtU;H2=hBU'
    'H9^quKO7jlWy_ZNw5QmN0dlP(UL<5Z(Vw>}g_(r`0r<4q&#^FPU{^2mb_8((MG7A$A^`J?U{ZDI(j_xqDGiQ8jNthtqZr`&H3{l9'
    '#TL<9!iGb~*fwQsuYj9^89X=JY_ojS*euLXT>n>0rLq!>xRRc`yDFQ=rn^`OPBj$jF@8`qmABD#E`BlC0)h5xM)>>RH4{9{LgAu&'
    'Fi!vtSF)vj^Jz&|k&{zyx7~JLiRK-5+%fa>gYLv=;e-hj5QmlI@9OpP%P(^>_Zk&SJC+>2)>>;xf#lKIl;}?~3*?xI=|CR!%ekOt'
    'UmW!OaTlJU^T_;NciqKxFf>k?G9{C7!Nmx*zwh^dDP5Uz{%O;u2^3i|wk)R&0_XxXKGY!&0Wngj=*OyBPypi<pr%aDo;}<B&2pB|'
    '0465LAOvx<OSf*_+@Xtpx#n)W?Z%ege)F5(%;drVH*Go25m?l!7n^RnsU5f6a!Zfw2MDYPie2^qcEmheSj&oV5K4#{sji$?fhZ12'
    '&LKmF*b8x+Hf?_X^Pfiy<zL*JN8hJUpJ-<Bo8SKkp*}%~u9H`{6j%P6<K={r8xF~2piE}oXokw<b1qfi;zKBgFSfTe^Q3CDZrxh^'
    'g}~^JUfZCdLBo*`-s`JtH%u45yjpbf$A{jvs}Zs!CnMVg$;`Uzu50Uf7$=arvxHz=LI0t>J^uLPHN2V<ojfEN$VvH-+Mn%=`b|rs'
    'jlK5TOA?Z<xUXB%fd?LF($+lhm%sd_vlORUTu?KAvD5wc-|x06=<~rl$T0wQL?dOUK$$5BMQe;Pbi#06Amsd1nN`dz&)<3Ho!Tqu'
    'lqBiNsrZshE<w-mQL)&iMBrXt=hV1S-o|@^fV=O$Tg}ENNgkAZ7AAsHPlqymAO{?90H1Iiy@)cDt10I&#$a%jF-%aZ{U-CuRYG>j'
    'ydK3VB*HV#JX0?$@k9pw^7QG`F|i%yc7!-x@rH;|a1cBWT74S`hsK8wv{4Xab*A8;W`0aH+qzles$>ZZdZ`_*j!;BKkHrsilBlV8'
    '2F)5l{>rC`lvUt!4y*}cb!!2JZQHhW*QazB?$``}$yS@4RYX^$L|*oSn2A!NyNvctcJ-EN@j>qhty;C>IyzaFkm1}LPd@o%T@K2C'
    'g;bHcV{2DpjY8wQ%5m&6bm-8ykvofffk}HIa8or|Jj|gO8<G|Z(3PZe`st^q^R1dZNFa}8vicTYeFW8+CV0jodHv;wYa0^!I_+<K'
    ';~SB$L`fA~Z(M?@Pl~1#jZ*%u+attGyLRm~YwYKa;DqO%d#;0(eqw%z%|v-TCJt0!#T8c!l6gF5&6<^P!M1e%<=&qoRSyPGnh|F!'
    't+Z0;`}glJ6`Ob2(Q>+{322loq5MT!nmY2xBO{Uu_{0-WAj{hrEKsg?yhMM;l}+-wQ%*U>q_xgPX(VsaqD8I`9#iHl)Lp*u#vAs^'
    'p@QH5AzrLyCI~<Sf(M$+l*j61mYadh{}$9be`QY(6`=tHh&dxI5PkREcV|xN!8PgTn{O6RZs*rufBpFJ<2k-@^ytx$Ns}hIBCC$p'
    'e*gR5$K6S?&HUgMmG$b?E8Fool~R>89(2$_M;&z(8|<*d4pA{NsM{*iTh*ykr(_)j)l*f+*aM!<KR)gD<-EXKBN<};Ko*vJzf<Vl'
    '+!XS*Kvt=41Yk~KO#|jF<W<@cvRN)Lv<vhmPq_Z|?c4Xl3oneGfqSDeksSeHf|!XPfOtl5P-p_>ZRX*D`pZUWvSfaO2DOK91Y-E;'
    'x`3HYHrd46pZw$}T-C8%3uOx<H`kmC2*H$5cY}ZmlsaoE+G9D%{NR;tXsfNZLTq1E{PD*h*C@@IGe<VoC&O>O^;X3wiFA)W_OPk*'
    'SeVaKO{VteklD+oP8i%g5dG|g7%)bO;H(1ifFdfSd{d{bd-TyqEis8cx~uE(f4F04z%94jl7$l|P892rO{Os^$K;kzKQ<CGNo<b%'
    'gCG3BJiyIbBit^Ikoibx3{q$^^57taettj@ngUH88=iaaIeto-_-xUX2Jo=M4in9B|Lj46`l=5l%Oj?6p20%g#rbtc<?R9iD~5?_'
    'Z7S}=oUR=#CCS#MBxw;M2{L7yzyg|(HqupzMfN!7oOARXt_<!S#s?pK5dD4k-Df$G?2?atvQg3luw%&qWH!I=zWc_6${DAO%8KH0'
    'E_YPIr?ih8IT8m1G9NWs>lNRP-(*#mi+D9Or`~6ueR81%Hl^Y0juixb(K&#4mH7yOd3YA#tAHkCS=$7%5Hw7gGKDC`>JaX^<BmIC'
    'd+oJpdhLpZ1sSul2xPR4$!sQeHPZOxV;>?K7YNR#?LsNXDzTJ>XA?)}1Sg(&A{xy<d#A$!l79rIM|Ir=K@!%8c+`lwwlE1?8KxSY'
    'KVo*u3{lH2aGOT*09-Ol=Wj<m^Nk-$XQ?W5`0(Koz$i$(LXfX2Z09DKFAGp&2q`K$_~3&RX@rb}1`rt0gv_IbK{DeJ9`G24gRqgY'
    '3M7m_!I4GaDr6yO(1H@hQH{hBZ8Sb}=1ke%Yb-2|f0ah5EUpPs$1{pf1$+7Bm*@VjYE5j}1mt#&TY^kQl<hO0^T#UK-gD1A5ey}F'
    'G3j8u(<ChiLSQVX$fP`kUXICIFlf-AvP%$Kz?GpgDS>_a+usg73(cNAJ7lyUJ%+6Ol80mIhs<bW^#R#BUsaea0Z_CdZKSUP(p8ZJ'
    '0&J08?5N+2Vgm#-n#p9m0phAsm;jSh<ZXh}svvn4q|n!rYDCfIIQF=4<9_|?U)LpRq>!MDj#0SaXiGDZ&*LwD`OC?ZCm(v~p&DK1'
    'r<C*0Ki@IRepN6DC`+W;moE0-e}Am9DL|92A*hTO7GrBf=FXk#t@bdx@QJ891XbLtK7IONa@SpVv10s8Lz*qBw$aKMGvkbweCIpg'
    '5!OIi%y0byk<Exk!k2iGd!|jBX2uoYM!cUA_*^l3L4dXVNz0<41}R1byyb{x7h`oTuuE|cNZytiV^*1QoN^)~ysQmP=xfT7dnr`5'
    'chuOmyhiHLGL4=9m7<Z!dtk(f5elulfINTdsi$h<9KOb`4r%>xm>+Nkk3II7X1GU>9*l%7;s^&woV5|uj*_;Vblh>r>7>l?QC9=4'
    'lX*%%vkpUs4AEm}WM;4AFvKqIg9i^5_h+1O#={RkT%6dsK=PUGisFpqA-^!y#`0`4tdfNpk?EPPy6P%tM<eZ-`T5U(uD92SNCv<9'
    ')vuh<WHPfV=FxU(oO5rYJ-H1*=8>T>g6v{GQUMZ1)ZnQLjD->)i{xmmf?yRL2H*m$95XY<26rbKy!YOFOAg64*#J}GDxYdZ_wL<0'
    'bm)*zW`OdR+rR!u(klrzaV9sCBTSkA$$-lO%~Gqz*uGWCa(It8pj33ofRc^?eP#64E2qjR37OAmU#Gy6{Kz2D5w;lS<i=<Uf_aFM'
    ')upTDPp?3!tWitm5!eyE(Tp8QXhZ|T7UW#Piz*~GtMW0h+y=4Mq*Ew}v?2T26mtw2W!}7b6qUE-PJ{+$hYe5iptqfckrPfZIEMPo'
    '7hZV5_1<2#wh(;;Ymla}D{Hc>#<C6?ACc6tW5<>)TWSe)Sn6xZu>h|4CD|H5=9B!xwh7d<Td%lAeYHz-{s2SyfD0YD6f$<~*q{CE'
    'XJoU4tPvMrE35^ORc@r=lUG$RdB+`hj2=Dunrp5>M*GgiF!$eozsM8WH{N(7)kH$OOqei1N<)0-op*+Vjf^!}>-_m^X7C|u;$CKw'
    '1Kd6^95reb!Jc{M8AfHfXdhw;1XwO-NMn=(aFCFUGLvK@MV)rqY080A$|zPj8daQd!U@M8f4o%ZCJE61No+v!aN(4AY{cq_Es5vB'
    'fwZT>x&f)_$}4d1mDtD-pGMQbgZ!U|sF;*+AkDIbHr35rdCtdGiEJR&i=_(!3T~$;YknvSMIfKAz4qFgfUkf3>uz!DufM)Zq{hQJ'
    'oN=d}cA_A`fj)joT6xIznmg{Y%PyFNTN7Y|z4zYRl|cz_-@bhjCO<2(+`i?9T+FX&kMMRBZRn}eURlVBMYBew2V3;%D)I#&c?i7G'
    'MjP$9=bn5k%n_@bM&QvM!9q(!5!oO*0HYqlyruyY^9mxVY<ldm$A0mPU$jxaEJlV#yyaHXu3bBnf`n$!+algZX1l9{smPWSFS}^<'
    '4msoyUYlF=hkxUjgyVQXP+E+8krhx)6lC;hED0ZYL}?6g=LKedU`G&ba>>w#KRt$<M~4J-vC1l|<VRW#8ELoMZaV~=(v`unGYRU9'
    'Yq*ipSsg1vy3WLkIJj+_A?m8Dt||^J2eW3)(&ZthCYOGJp6x8j!hp#fiU#XqHvYmTS|Y83c*Ni3(?+_1gp*Y{uW&?GU)rd#;5k0X'
    'r9f0<cQ>if7gbvlJZt&!e*5kB$Rm#+CaT=-?A-Ue-~FyYs;~GEBA&h0S!W#;UuCI%<&%}kPI?w19(LGaiaJsUt1qkpyk@8m;u{1v'
    '7U2RUQ~N)T`SaCkc42!xghwh^VTBcP?5M-M{v*rDN&DGx>#esgJ^=~>m}<){w?qTB-FDlQNd#THcD0rYWVQT`6s4M~2a5CQZ>X;('
    ';s+f-LW-BfmgsXeh=e@ALXS+2v&q=f$OGK$@re&vIA+WkCyMas#WY|iAudeW(kQ-_fNaf*_|&yt%QQd)lRQhrMmCa-azV2fX)nI`'
    ';`|uN7ASfDrkifss#PnQCNf*nty?z@MB}f0j6~EOh+5e%P8~X9*2LWcq%r1g0jEOamHCe}KZHW_0gJUxrLdh;2Ze3GfB{DzeRO{C'
    'lYP}Z3@x#+Q>RW2W^%L5He-?;0fcQpW94lU>=s*WL7P!1&(4k_K%LIyqmMqC=m)#+zWb}Mz8Y3B8z`J1wku4GY<$EKN4V%8e)!=K'
    'il)811q`VM7xQa7;?Sh}*0;VTx69sQ-SU3@`YH43+V{Tqy?l`1yghsNEXwcPxpS&MGO<er-jKHn*#i$ekX>GU@kQ0sq_YbfYUFd8'
    '30D$VEs2w7cy;Nemy)NJ{4Jwa%@>al^zkXcqoev58|1NylvhE_1Gs`+bhii)iw=OErS(nxF*gFfOlz)l5r#HJFsoRCt(;+olQCiG'
    'egnl7-hA^-OJW9TKxvB!^+A~w3Zg$kivIi*$2@cDd*Ww-EG(P>Yb9e*<<zNDwUD|8;*lB8glr5WphZ;D2GC%Z0_V=mN2OcKl`4pA'
    'oF<>)ZD(JsN@~=TYG~AKZwaiK@7=q%vxGoz7*+5C`NIqu4n6eHIdkTiloGJDoz9oH=}v5CzOnXmw{5oBhTAe8j!{qnkFv2=K(a+f'
    'o0KJ_co`LMzWL@RB~Nh!@-2@0f-S>*fk5xxmgr}Gx-;+CAo_a#`R4@_tB^|WlAL0dIE{t&TA2B9aA8S)LPlk_^H$-o7DKI*8&Q_v'
    '%<oMimRj7z!X#z@c?6h{L7<}qtF5+L^^3cJyIGGOJz{2AO2DyMPNelSPb$C_l6$f|yEKWJ_?d*EJ9aq_J;;*sOBURSZ40tZnKDH|'
    'aYmEwBgOMTw%|HgIU2lH`{lr0cinY<lN)V}i?&=A8aHm748hrL0O5=dgOSI3?zzY1$}LedHFD%g?W@4e28#x<h(7~62J@#Oa0Y@>'
    '1kp9>v9tm<ckbL+QFoTu;*Pn)6ey;`YUiDI4jJ>u7r{XX9h9TVoI)Wx3I#M#wAMX;iP1#YRVNUXR2(g7yP{&OED14x+IZuQ5dg9H'
    '&_fS}8@(m@1NzKKjn%!(gq26;4>A4Z=j2|+%tqNF!3JEK#J45LO8%DZB@FXcVG9t1KLoehA9%Zc=EG`WjnS!sD=O-l`z6)WAYF?#'
    'pu@C1zIb`ev11!|9u{?l#n~PU18)Ii)VKYN4n^iSANS8%e(9o;<fcs8$fPaY7sQZNa>K&RFVW?KV68+&;Jgx-B2u7>r+l4bFA*%a'
    'p#;LH@GRT_!D3pFpDAEMZ?k95HV^2uD!f{r1_fMn(M2(*fH9dPE5_0c!C27<$gJ+5#38Jl+VP#jJ!Qs>8D<J3GfB6Mjho(ujmV(-'
    'm|;QUwQ1Al$}6uFnRJUk&Vg`%H0TthZW|;fG6KP&#10=m909{4O*9iLk^$R0wpV7NDY)JZlhaQ>J=R(@<YHrY-R#p(Kh33F@oC>)'
    '5~ixDC5$}RU3Xo)+eXT>13TV^8*b=SVHF@x5UO;QkNK6Sc&hmW4?HkQUOfHur$3Dn<-`1fh*wF3wl-OM>B1Ahz0&z7w^OwImte}<'
    '*dC2Qmmw1Cg2B;f>9*T$Q%18;r^N#<yX>+{F1duBFTC(V4U>*c7Nkaeu$5*=>#+L#`|i7squ^Ye<Kc%NPRB~&UAlB(CiC2pGOda!'
    '6*WEPph1I1jT$9UX^Ak=U|Eh&!=y=*Qa|irw2Viqn!|<-(+^W3QHYM3G!|t0oX`lO9g`4!2VJu$%QcD1CV-f%RAAHWOQ({Ro?pfz'
    'uJ%Mqh^b_IjaAt(C4&at``I-D1PV@C7ER_5{e-dHTe~D&M1N{O$)y3^O!TF}E@dbXp4mOPcA8*-_USY_@d`v9Gi(#v(}s{p@F<i?'
    '3D&Zx{zLO9()19TL}ugADrp=YxDBO#sUXv)O_MDM#^Z65)xC9=xVU7){FZ1!PdxF&vb7NbpW_Eq$`Nn6=_Wabkcex`;>Zq<mokqj'
    '!`q4BLsU!m_0Q6464etxRGGKIA^c8b6*tm|IzyavVc;!#f-(|y2y{&YrI3d%Ah8GImD(B4OZ0;jyzCw*IC@?`zxmylibB=?uoeBh'
    ')^_Nv(h9_)*)9M#8Qk40R&3d_<z08(71Nfjb81;3igh&pnqY|eBfNQUZ5kxFQ>VpaW_)z%IRg7Jst*0uTW^(D#*ZJbyW%}%m6{u;'
    'At(3c#AGX5*e<7|VIHO>OyR5jB+UyB0+H?A5#2p<R1)M!MGH|Dv3Ns91THNywk+a~>um8y{X|uF3xpUW&PxVk3`1z*LV|N!_XzEn'
    '-D7~=cH1pS(xype_<+JQ2$9-<@WX16jVx9<>Zqegh%N0r_Sj=)$f?P|D#UU|PFH7BV;G(p3E-8Q71L6h3X)Wj9EkuW8m%#1ho=%Q'
    'c)X$|D#|2;$IG?(Fvlp{Ndt;^jtMLzEn<LjgAF#2s@;<8o7){A33<{v#SsA}Lxv2|8tdM;j`K#t$qG7&v(7q;p=;}f8ckhW|9KCs'
    'kv?1qQA|jEYO^CvpFUkD$V<4}4L96iRKO%Rw?Rzf)J93BT%Q&Q(J#3PWC<~L+;K-qo@sa5X(t_x77W5+$xTy|wxmZkQX$I+Q^Hz>'
    'E~YBwhS<theiMdFdrbcLt~(o^B&0{xG?wFA9XfQt6TKU(3uo1L?FfQrufc-{Ge%8!4Gn_TOqgdCAvusb$ovQxyMO=wWHYb!7oX5%'
    '2<|?8jyvU)lTJVVG~>x9pE!8%peLVvod5X8&c_|sd+DWX(I2S)=d~JBg(Kj@_BveIF@nVfZA234BHQ`rpD)G)fMS<@;{n`)DQjQ&'
    'YSUm*P&r1TigsiYZ`E>56MN_r`C!|&ZP#0Gy>-`J*C;yag6_gSaLzgB@b{_`=Lh^H|2QUNkaf<SIb0Ihlu6ucm8)1qo`3Mc2V-EX'
    'Y~kconFx**1EIl0Awz>X1*PvbpO1ZPD*Bb@nAUw2O(f@b@)_9>8;jwC81dnYD8;49{rKaLJ7zR)p)r)2eB#Ag0+B<Qy#4mulQ!r|'
    'GmBD(r@5-%rQ|2P+0Sw*<Zo}k{m1_O`!2WKGE(7EOMT8iU-&``c|zRh+6h=uImpo*oF0M)NJXR<UU)&@aPGP19&o?`GC@LuU>9F}'
    'al;*V-0}MBuUFRqu7nYAXXxF#w;a-{RV!PH8-bXOiyJ1Z6fn;*)&7!*76{RAqTedjE7$rW2?CKVwRjcI`4GcAU(6&`>Y3A&kTzz3'
    'V<NJ3QfLTM{SZRemULRJ{V%=rlIV|NAo*l1^dX>-1XiGD&z^~X%xy$~>>D5kdU-n`f#!!l{9$O6ZIUwprHvRdLbc`qu{YZ%;yFZ+'
    'ci(+y;K2UNE?bj!`Q?}G(q%6hLWS+qr;qSUu!(8ShqJSN5<n0#6oDN9fq3Sr`|Y=%-P2F7QSD#rNUV&Or9@ey8@lG2YvfOAa*Wv6'
    'N{4}fs`=ruy2p&tqy4h|0wJ<W#pxB)GsPvZ(o%##m5I1c>w@HL#FIi5EE|D?4?fsrw4Wu8YCPeD6H3+8OEYKAb;J=zaJcL^bLLEw'
    'Rz3Udv#Mzns`V2tC;_!#4heriKj`_8Qv)7`iJ3wGJj(nVZ@jTxFhoOu7(?wZRb%w%(e;#2O9AnKAn(2R7yobq4$-Su&t;ZT92-_y'
    '<tw+}KB@lufp_%n(hr@`xh;5uI-Aq6x*LI7K$*nIB%~%*9TaX&Mg-7N)&9a4QPiC?sG7vp+)AD{aL7`gW7;YLs*@oDiT<im76^wP'
    'dT4Yc`R9*eK~_15(JiI)I9lHrL&A6EiwiTGj2Z{e5+JM<3zB2ru3bCEFc|}|Tq=T2#e954xa)o*{9b$Q#q|+awU)CWwvk7TZLaT9'
    'aY-T62w@)Hqrps8!xKZ5DCtD=l<(re)mL8~Z?;!BfSkQN_UhG(6H~11+!`?P3eDunlX+D%pp#cOy-IJSjnx}<?b_8|4kCkkWf#KR'
    'sA_%%y7r-OQ~<noc~cS;!h}~e80zG)Ocpakf}78bxA37uho;&s3<ydrv6z~|{0e{UN1m9OQ%kdQ54J`XG<O^BZ=o`j^Xt{i&Cn5?'
    'UxQLymLlXKs>!A$`lU^)vLo~7TQa+J?AQ?p^UUuCh@rJ_z4aEU(ifZ0^#MJxng7i<-=t+$fAPf^XUv$P#3rTD10k=x^2(^fLJfiB'
    'PcQ`Mop+wHt7CUeob5#n9E6|kY(xgBo_-LMmO0~*ULok-y?ZPoot%2=srk1>0$!!(Ft23m>cobeuet$I?aoZ>Xr}g8sus|-vk^jr'
    'pi5NdBOeMt#ViVmfxXJ&Dj<j@3Y>DvDUK9@76uF^i6!M4m3E?Eb=6g};<$0+-~$E>SZAGeircxX5H%{8I(2Gpy?{C0J2V8k8>51i'
    '2E|BsG%g}7N#5?!qeuLS8JH?TN%d|Vb!N4nKA(K@$+DO2F1+x<@XSw1O^Rbo%4bIot`DVsk)+%RJ8s1A(S|g`eR)p;j5%qT$7Cb<'
    '#zwNNV`>yY-@biyT9Uc0&`IgXtAwAB)i;C>0YtN_-8a7RjpP{4K_U2SnN{~RF~Bv_g_b*V`flC2*)d6rN+qH3<Hy%N^T$y<xgod('
    'q@~Q3#u&be7u!30$Usbp0GNk=<Trl6TgHfWE#Y=7>M9{EOX=hzYw;=pEhrP>4lmjmhP082rSakQr6-7-S`e*o%pYNN8^lIHWy-eW'
    '<Ei$udfXyFaX}C#KUqRBml1r9KmPdW5m==JKuS_YX(UPK&Yj7e-E9Qiq0EPdgyC!^sw?{B6Aow<hj27Drj2~Qgh|dK5gv2QF*t-('
    'zkdCaHdHs7EcODi4&@j#X;U{Kxu<4+sYxDCEHTRKV&iBE;<OB-;<`a*<mO_|xcTOryML=duGdC0n9zFyi*kAE{C;2tq!QzAmvZ=c'
    'bRd$@yFOGN&RQ~C%dXTEe+9*jKu)1>y3yGky)kP!9RAKPXNnl&VJib~zy0=6Q+C8E+>DtI$feGON%k$oWt0_65?EWKF4#IQs{n~z'
    '30&48UWlulbNtNXAM;BIHuB1-_Lt;`EpTyVwJdb~=3DmC^tos-ip0rv2vb7H!8xn+)1^xnu1};|^4AV9XDCQ!0SF#x7`}i1{u3rl'
    'kh{1)H&jI+?mR{*6l7BUOL>zIC4aI4!D^QPvw6&n!zf@SjIvh=L8l_O{@m4HImovZ$$pkQJVd*2+G(e8yVyG;O|84`y7IP~7XeBh'
    'q8-tn){nl@OqrK6d9dv$2sl15$OAVC1gII>qyhBYPOO`EQrsK?fqhMKBZjA|2n=%A2;d^E6-CTH02h>ITGm0xnt?cSKENlf9@8Dq'
    'jF6*VOdJOvBMq>hFmuIV={9ZJs8HEdF1L}8R+t1NZER2ntRy6wVSCC)S^G~psW<t6qSX5mNshPw{`)IMS`wPFz0pb$BWIA$2scTZ'
    'NUSEw;mpr}ojk`5MEaWF){TdF<j9eBS4#Q2RHjk_$~8%Jr4hH%HrvROn(x%j@6jj+Ah_a+E36tcXpmy5Li55LFJRvQT=?os5Uc3O'
    'uri6mwbx#2xe~7A)2X+}j3ViR3obC}-r?q{OyK;~Q%|XG&ZuAUN*Nk9Y?xL~iqbn12~Q`Dvx3M>_<8f@X>P32Uk@KX+$xQsRVIN`'
    'Tz>iG#s?pKFeT%OC!QEJY82T7z}&fWF>&N@%W*VxE=Ooaj2NMoGknb96#-s*?KS@DtzIf0)IRdaBh{9?LHAp}7YDIASb61@J9X;h'
    'NbY0Bk+BX^AVOpWgtik$Rvne1Ocp>!?sz~%)Nj682kDm$EOBimG^gY}^jHp44xIBtXwwML0H?yu=xubDy+XqhCi937#9JKL62r`6'
    'Qm1W<*J8Xtc1H=yc|kl9l1=lfx;5M+K#6AvAcX*+Ax4nwn@14-9swkiK~P3!ud3WYRz;+|GE;w3gV39)-9Cjldna=O8Be>&_a*ii'
    'o=zh)j*#q~5f?K?96bUtA*YGL1VI)S@EE{3B9_Ub{Ro%I%xIpo1$x3{uc(Q0cvYwRUl6AOcF96y@ED-FB+(E8V}U}$MS>TeAll55'
    '%oJ&n1uy0?qBqGN5vER9Tapj^*!R8n+|KnhHT*As;l#)qCrRg@ys|3C2k~79F(M)UVP0qo5=oFn@<2oRxm9ehJh&l96e1G$tZ@un'
    'uqJLzG4%~H=A30xxpf;+JCtOra%S@jgeI?wSaqe;4r9|?^vf&rWmP$t8GC^|LZa@xLKfKqWjR(}dHl!=DyYha-WK5*4n{kg<fibb'
    '*jV=UHn>TECi9?J7xMo5*fu`G{qyd-%vqa<eDJo>Uu~DO<kQuNfjK178iSb6TYCgd0+Y4suH^FoLBI9ZTPr+TN5{y@kY==Nj*(rf'
    'udp_{J*(FGAlwEH97wFhLFk)+$tc2DW`a~9VwAA6onu=9DUgMb!YcBBGJ_8$d4xMiNoipSF`5C8ur-Di+^Vt!SY<2>>vJ^-0m{6B'
    'P$mS8_z7ip4A>ziw{1S_75W}Clec*lGRuXfT}q(mjsP4gh*#y2{1Fw{&Lce#0|!uxczvJdkiqgFZwPf*M|8`KcFh<+5SI`m8bw!+'
    '<`!2z(-8gn9LgTXK(-7Xo**;XE>P7haSK6^!T}Pzu&UA&56Davs$1u6#^Pkf<i?`YfVJkS4^82d6qVg$V6Qv|Jmy4&KZAuWdPZeh'
    '2F4&`zWRgkQg6$5fFG2JzpI-xA6D5eBPjdA*`-KmYdeOF1#=2`E!(@O=q+_#SxMHccfu@q<&{@v&6=gR)Sc>A1L;vUt}YUcy~7SW'
    '=oVvpmk1pIDYM6?0&?Ug5cv~K)DpLVRMj1%%y5O^Nj8N55|1exiJ9Oi;~$`SRUylgNWz4X0c02R#ghQaYUU2H$Cy96m@Et%Tn5}O'
    ';!TLc0v?dhZ@u-_?|%2YUAlB}am=^-0e4N%x^-)PzWZ)IREq6&5~_eiQe<uZ5Oq`$0H?BCDB#I|i7L7Z6pgvIAai7s9xe1J1p@am'
    'OOmCT!8OO*;#Ed(-3gUT++ME(&AXTCz+CO%*g~Fv{&`XQ+;h*7!KT6zb9}sw*mQ2D>M<H{oQV_;4?OUIi<QU7+<=i67I=u<fD0qN'
    '`R1G5k2l?P)0OJK!i#lFlUN|15d}1Txk0!%q$T-uL9t~3PS@_`lP6Dp^wCFu@rz%0bQ7j=F>}s2=luA`KaR~1Uw{4eoPOranH~}7'
    'f=@sFbS1T4zkX+)d8P|@6k@q~_kLc55-@=vr0RR`z1OM}PdxEE-}z3jUcF@H!|Q)IMgVO4_wSFJ8+D&PeI#8zWR=uAd-m*vJIA`?'
    'jyq<|m=Q0m9pE4TurE2$+;-b-_U+lTr&K+0;zY}xuJLdDO8el^Auql3(qoP}hT-$kCX>hN`9!m3=ZYVg0eLKw&fs|A`6+_fBREwE'
    '&6{qzY5e%{%sgSj1j@bWqKl?Zohm#OK8?<C#~mk9L!7_;7&&sJ5WM~N+dFpbC}wC(b4@MM=UsQ*Wg*M0y6UQ{ufF<}DN~Y2n0&kK'
    'w)_0&KkpK5Bbikc>BRF(3;2;o9?8sm@4Yueo_F4P#~yp^jyvvn{`u$cyYIdbDIrFboX$S`Y`(&pIR&>M$$tCo$5C`<9O$ykF5_FS'
    'd)$|gHvW-Uq93ENW5+s%hG4n+->a4n0|Fk4`3n*#h6nJIPC7{ziddgg5I__(ZrnH?K<}b<=+L3uTHJVh!wokmd35^n%P;4ca5hsb'
    'pLyn)?YG~a;f;d^4Pvb=w%CFp_u6YORY04SFNZ`n6|vbRmt2yTnQf*`o3_?kYpuQZ+DvO<_wL=1jUGLk)0}(mxg0BnoSApqZMROH'
    'I_<pk&QCu1Bmp7=8vJ=x?ErPNGC$AjTmMh55X|JW&ps;?$nIFxqTe~?C9Z8V5zh_tp>MG$#9a;q&!oqC*w-L40M11$XZ|<p|8B5P'
    '3zVQQ5c3QDTW-09;l+lyR{|J};zo@c)v{$vsaQaZHYPdhsH4OyQ!%HA-+c4U6<YIBhzu;gy%q5StWBFXbd@jh=FFLsKdGiqA%6GW'
    'ce{n{UjHj&M!}i4k%YHNJj|cEu6~PY#>#J#$P@ER3D;eBUA5_!PC9q)%r(zC>nw>X-h)X9YG3)N8IvYW+GUqrI4oe`YS0s_n7j(H'
    'x>v>gRWf--!%`RE(*g6FpF%0?o_p?L6)9JWe(bTwRK^raQO?GKPh0Tt!w<Xn?z7K6{9a`};D7^Kv}kd`1s5bw;6vSfg83DAg>|*n'
    'R#O*%$aX~%>ze3gmtB_e?!5C(=9f0Omo3DkNEJw9#*F#VkA9TT(b>yI?6AWQ<6-UFx3B(4?gu~kf!mp+;xXl)hjs1R)x6lYb-tF%'
    'D^N4Pw8n{K`^0xKXPh8LH3ov0Di9n3%0h@HNTQoxr_B}?D5^=Fw!TTn=TTBkqb@5j&zftl>8>l2OV`SIz5rAv+=0b+{>2W6Q&@Fm'
    '$mp3BA$%o2ptZb1hYsp7R{Ts;#uZjrVFkXTKhSi`G5kOVI^l#9<Q@r}JTf`g*<+7AbQqF(LSM6mxjd1si;zYPx#NBJ-N(1AI^>W;'
    '^eY;dWP;i^enW-~A$`?;@q<qaagM-%0R!TLpHg7CK;!d#AcwP}q|JUTr+BlOfR<zBdF`99ivf~pr%#{mk}So_;cKk1#u;av;bJAw'
    'wF_o|9AMP^Ge17A+5AI#7vB<$-i@UXI_MxlFPUp2By0g)ewvnmR;2WvCGd>$2YQuD=Fsl(hHl-uNgIk3@oN80Ul}aXoH&-jV!}@f'
    'Rr5>!|9_ibKPZRmsuTtFTSjFJ9w3(JSMvY<_rI?~Q?bxgZYjj1PEZpkZRm|8aIr}F+CA+$pRj~aeC=yr!&ab6Jj#bAfCEW3VqFP$'
    '>*HZ+gc>9tXaKc-kyA9wJ%a}iPL1FY=3Pf+np_U?F|xRFV66mD`x8&n#AdPIKNtF-YJPd<)7AQYQlQqfew3IfcVrho^E|^Vfh)qe'
    '1t-w)2{;u&aJx!Sxa-yLe)l_%f>kdmw@bh}4BaL3D<9IehEcbnf6Qkv!-fshdBm9_gMME0qe<v{rA1&`-F~kAWb4(d*T8`T&EO-`'
    '^8kKh_Z~fZ5UjxgQnzBZn6OlVPXtx-lS~SVl%EQyhk6YfpFavlj2NMHS4WfN^aMgv!4{_T=FL-X^->#cw2@4}9`d<YAy%6?haeWC'
    '9@&XWT={TQx#ju%>}QirHgR2)`T322M&MG1o`lTIv>HIU1P#~Gy*>2ML;PUEgbAuc@7}%h)2z!cznlQ-j$&_jvc>-nJn+D)ufE#N'
    'qV})y<+HJO$^ZWxRL$>r>3csd5Zi;s=Z^wP+<4=Sd-m+<g2VPqa>^;EDEY44g3Mh&GHBPXos_O0Y}>Z25L9}Z%}s2pt+ryhS3dZU'
    '$S}W>EaFAK&Rmk$6Rfx1dX{gv;f7r2*kg~CB;*x6T>>50(H(EwZMQw<m}5jUg|oWeo{bd|kL>N0g;vd+In&jHNUk0tQ;Q&@U{(8V'
    'UCfrdevrBaKQUx^#$?3^`l*0M<9}@P%bdC@^_%U186ha`=TjgS<rd;kZ54U<-+#XuG$;i<V@vG3L?t$9q$}KBeDOu5vLt`V%~2ry'
    'Ok`^>7%XWYD{>lDbh!kLKGfD?o%cw(DESHw8VEK<X$$i@1|DN0xXJh5fB*jh%B<*p^_6b300000NkvXXu0mjfP)h>@6aWSQ2mk;8'
    'ApkiswUOHZ007Vd001KZ004MwFL!TpYjbF2Wpr~dUvgz^b1!pfWo2|RE_iKhE^=jTbF7igYQr!LfbWC8gVDWq(!&^|PA_B7>)Jhl'
    'XDf+WA{$Fv*SDW{C@iJ0%WnEXpFW}O?rWaZCl-#0EVQ}-h2&F+G+1~&KCd4@u{a8DN`w~9$gt~HciR`7oYWy6V`fz;#1=-8{AdjO'
    '5ogCbWu)5aQ<<IAv<xOY|Lz7fo2I!p<u?Pm1yk*Vg|ZJDP{(twoBw8~zK<TClK+^I<i-jUiaXWID$<n!#e&wC8ef7|S7k7_ZQB2|'
    '$t99tIiN_D;)T^!&s^oLev9-YX%>Izz5xIL|NjF3P)h>@6aWSQ2mk;8ApkEd^T@sc001Nb001KZ004MwFJy9IcWG{Cb1z?VWo&aV'
    'WO8A5X>MmRE_iKhE^=jTbA*pUY6CG0MfZi?!DwG=ri)O@c(N&hY)f(gVQW05W7*hFNN!(4Hw_fJ)c^eV(6qnBO<lRvki@{K2T&|F'
    'xgg%a`E+=B163P|qb&&=xO0X5{QNY1@<y$X>KbxYog&r1+FE``tajx{Rj16-n<bS<t)InXCi_L6@pTw>Sbj5LeqgGjGf<9h1a-RS'
    'ZvLO0mL*vJNcIw0+8+uGvAa8visxoPYmDpy{p6$0;sNIgAAXy^0RRC1{{sL}O9KQH00saE0000X06>_eI<o-)05lZ<04D$d0C;RK'
    'aB^vGbY*f=Wps3DZfA2ZaB^vGbY*f=Wps3DZfA2bE@EkJ?3Bw&!$1^<|1@5zE>sXia3L<GMN8vlX=$of8dI89x)O?w5YVKgDWc*+'
    'T>Apz!}th3fGZcS#0PMr|HO(YDn;A~XTq6tnKNf5AH%WWVHY6+v?+UtFoQCE4}NS%jH-x<HGDO+K7HU8^$C4mgN7koju^=%tx?ny'
    '3F>-6=bVG_<?i1FweB{h6BMPNIo6BLhhxm*R&5qHNa?BZyPXT?r_=i5#i_W~A6`{T{UZWz@2Zz8-!RFss+%fz4A(fp0#^9S*Qhg`'
    'b4a6rC1e>(Q%o%LOEV|Sc!AaeeS?06I!`Z)#n?M##Y(|3OC{8NFLc|=DmPkg=y=C2at7?sbN#3hd4Z3zU8z=b#VRUpJJ|2UF`9Nu'
    'SsCPm{UEf0ru%Ar|7WRjVB0XQHwIV3t=Va&KTvealV&H4`$Oyead2budhtyCf2q5M<a=F}tNgjHR{knwTlr%khk*U2g9Es6hpDme'
    'xa>D&G-$PP$n_9SCV$Sfx#stnn`ez9wAjHM?oj8cjutT{LkI`~As_^VfDjM@LO=)z0U;m+gn$qb0=*D;0RRC1{{sL}O9KQH00saE'
    '0000X0Q7J`OrZb(0L}ma01yBG0C;RKV_|G#Lug@XZZ3FjY)p>93c@fDMDK$C5cY0UtEfoR9t00PtM`zuZ6Mi{Y!Ll>BUJG+FmK>7'
    'b^G*AJTS5iq+nf?AdpPxEe#5;m-FfXf|I6RyFjSmg$!-8T+~K8a~icJkqyKO#+0U`lx#-ynu{<Y<>elt*D2RHNS-3<o=3#wovg}o'
    'Cw&$`QyUQ#bela0n@@mnpOiDdRoj_=N%v3W?@99p009600{~D<0|XQR1^@^E001EXr5X;fSOEY4j{*Py5da7QWN%|oa&K^RFJo_V'
    'WiEJaY#^cp0VtpV0RR9100000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '000000000000000000000000000000000000F08~PQxG=#qUjghlcx>(k_c>+AdMEE9Xo$qwx;m+lCg}0Pc@(vuD|hN!zW~%tSBY'
    'JLktaKswsW!drtjqEs0oL5Mt1tmHy1GxRtg`&Y;_mMTt*QbC4xV9+SPyhvG<uu{W=E-TPhfbkqjnS@mtT3K5q7#p?%NloAs$_*E#'
    'mejhXUgC<fce(@|h2aoOu#{6vaSimVj);y(&N^iEp^FB<8HNHRC~GhY2>#Yu(9-;BXeK?S6n0m+PRF=?HfQayQ*VvvgqzJK*u+f?'
    'm!IJ0$=$4}>I+rh6$6x~oFz;HwWZFdhEeU!;FTIH4twR@Q@uXy;nSPG0jDYMY3q9}oNnTtU?G5C@BX%*TioeQ*){@j&&5yLvZu*Y'
    'd~-V=qdW?up-)1eka-v<QJf6Impai8e_f~DkjGsACk~R3^oS=wp5n7Flz#yL0RR6308mQ<1QY-U00;m803iU~cH%*t0RR9W0{{RJ'
    '00;qOZ(~q$Z*X%jVQ_FRcx`MTq67gbpa1~?00000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
    '000000000000000000000000000000000000000000000000000001-l0k0UFc3xe1#$<1`<QZ41cgB{NbD3YkitgnWY1C~n+Zh<'
    '<XEtKo1Uc?>qtTzImxQKnfas7r%%f3edSmKG1=gAJ|1P9fj6PBev|XZ_4o5{oF&x0&@On$`2iBYR%c)2GKLx=T1bp6e9HM2v7QK#'
    '%oZx0MkISu$`C7!)W(enWoZq}g4tEzQCwu%m4H2hFQ7QDkA}Q=oHY2qrbRF<UwT>}Y64dBrmmee8tE_9PitaGp+t7KH{hfgijwG7'
    'V0O{sL1j`5Te&jY!IX?usa*mo-Uf0Gx(#iqZA?lwIBCFOh%DLP6x#)7zjT5YC+CfhR(m9@cCBx^+SLi8nulkdSOLuZ{o~Be1u1B+'
    '*Sd1}$*{Ix)wo-r;&ib!=y^yuC0$!|kQU`qNBjhQJOu9;dh8zO{>SaZg3a#O;`ZU@17;Th9r*5uX{c)L4<sC`d+T4)W4#VDjqn=o'
    '=}@k=IzmDBaSZ=BkaNm4y4Ga6)qVrT>&{1^eF{(gdspMDQFfWpEe~c=yuEDo4*&rF{{sL}O928c02BZQ00;m803iUrq~nW~0RR9l'
    '1^@sP0000000000000000000009!+EZggdCbYE0?aAk8{E_iKhP)h*<EdUe%1^@^E001EXwN)@9^Z)<=OacG^3jhEB0000000000'
    '007Pd003WdWo&aVE^=jTb5KhG0xbX(00saE0000X0FAEv!+8S$0Ko_V01p5F00000000000002<1^@tfY%h0ja%*C5Z)+}iZER3W'
    '0Rk-m6aWSQ2mk;8ApnvGl<Win002J&000^Q000000000000000nF;^^cx*3Ua%F6DFL!TpYhrJ2Yc6<gY%X$TY;#ac0Rk-m6aWSQ'
    '2mk;8App?<uwS!r008KO0{|ER000000000000000-3$N#cx*3sZ*psMXk}$|b1!pfWo2|RE_iKhP)h*<EdUe%1^@^E001EXlL<#a'
    '!3O{U6CwZr6951J00000000000072*004MwFLY>SZDlWXXk~3>F)ny*Y*0%90xbX(00saE0000X0BKA}jc^M90K<>~01W^D00000'
    '000000002Eh5!I~Y%g<kd2D5KE_iKhP)h*<EdUe%1^@^E001EX0<(~6=L7%%OAi156aWAK0000000000002{x004MwFLP*Na%E&w'
    'baH8KXLBxiZER3W0Rk-m6aWSQ2mk;8Apllq_;te^000)E0RR{P000000000000000Y?lB4cx*3Za$$FAZfA2ZWO8A5X>MmRE_iKh'
    'P)h*<EdUAt00000001EXRNz8^&N%=8&N%=86951J0000000000003pQ004MwFKuOHX<;vEZDD6+F)na!XHZK40xbX(00saE0000X'
    '068(Wk=p<O0MG#d03!eZ00000000000001F?f?LIY%h0ja%*#FWo2}8FJE$HY;!MjXk}$|F)ny*Y%X$TY;#ac0Rk-m6aWSQ2mk;8'
    'ApkEd^T@sc001Nb001KZ000000000000000fbReRcx*3Za$$FAZfA2ZUvgz^b1!6aVRvb6XE82#ZEP-bWo&a$O928c02BZQ00;m8'
    '03iTCn4~(h0RR9r6#xJy0000000000000000DtfR0C;RKaB^vGbY*f=Wps3DZfA2ZaB^vGbY*f=Wps3DZfA2bE@EkJP)h*<EdUe%'
    '1^@^E001EX^l(5-p#T5?&Hw-a5C8xG0000000000004LL004MwFJobBV?$_RX>KlfZER3W0Rk-m6aWSQ2mk;8ApoTs4zO4O0055y'
    '000pH000000000000000M)LpwWN%|oa&K^RFJo_VWiEJaY*0%90xbX(00saE0000X0Nr-tL7V{q03QPY01yBG00000000000002h'
    '^#A~5Z(~q$Z*X%jVQ_FRcx`M@O9ci10000G01yB?1ONc1_y7O^00'
)


@st.cache_data(ttl=3600, show_spinner=False)
def official_template_bytes():
    return base64.b85decode(BOOK_TEMPLATE_B85)


def normalize_template_key(value):
    """Normalización para etiquetas de plantilla con o sin tildes."""
    deaccent=''.join(x for x in unicodedata.normalize('NFKD',str(value or ''))
                     if not unicodedata.combining(x))
    return re.sub(r'[^A-Z0-9]+',' ',deaccent.upper()).strip()


def official_template_activity_zone(activity):
    # V18: ubicar por capítulo del presupuesto antes que por palabras del rubro.
    code=v18_chapter_code(activity.get('Capitulo') or activity.get('Codigo_Rubro'))
    if code in {'CAP-03','CAP-04'}:return 'estructura'
    if code=='CAP-09':return 'carpinteria'
    if code=='CAP-10':return 'aluminio'
    if code in {'CAP-12','CAP-13','CAP-14','CAP-15','CAP-23'}:return 'hidrosanitario'
    if code in {'CAP-16','CAP-17','CAP-18','CAP-19','CAP-20','CAP-21','CAP-22'}:return 'electrico'
    if code=='CAP-27':return 'otros'
    if code:return 'albanileria'
    s=normalize_template_key(str(activity.get('Descripcion',''))+' '+str(activity.get('Area','')))
    if any(x in s for x in ('ELECTRIC','ELECTRONIC','CABLE','ILUMINACION','BLINDOBARRA')):
        return 'electrico'
    if any(x in s for x in ('HIDROSAN','SANITARI','AGUA POTABLE','TUBERIA','CONTRAINCENDIO','INCENDIO','DESAGUE','PLUVIAL')):
        return 'hidrosanitario'
    if any(x in s for x in ('ALUMINIO','VIDRIO','VENTANA','MAMPARA')):
        return 'aluminio'
    if any(x in s for x in ('CARPINTERIA','PUERTA DE MADERA','MUEBLE','CLOSET','GABINETE')):
        return 'carpinteria'
    if any(x in s for x in ('HORMIGON','ESTRUCTURA','ENCOFRADO','LOSA','VIGA','COLUMNA','ARMADO DE ACERO','CIMENTACION')):
        return 'estructura'
    if any(x in s for x in ('LIMPIEZA GENERAL','JARDINERIA','SUBCONTRATO','OTRO PROYECTO')):
        return 'otros'
    return 'albanileria'


def v197_append_excel_annexes(original_xlsx, record):
    """Agrega una hoja Anexos y fotos sin resalvar la plantilla ni perder dibujos.

    Opera a nivel de OOXML: las partes originales se copian byte por byte.
    """
    data=obj(record.get('datos'),{})
    photos=anexos_fotos_validas(data)
    if not photos:
        return original_xlsx
    ready=[]
    for photo in photos:
        stream=image_stream(photo)
        if stream:
            ready.append(stream.getvalue())
    if not ready:
        return original_xlsx
    ns_main='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ns_rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg_rel='http://schemas.openxmlformats.org/package/2006/relationships'
    pkg_ct='http://schemas.openxmlformats.org/package/2006/content-types'
    ns_drawing='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    ns_a='http://schemas.openxmlformats.org/drawingml/2006/main'
    with zipfile.ZipFile(io.BytesIO(original_xlsx),'r') as source:
        workbook=source.read('xl/workbook.xml').decode('utf-8')
        relationships=source.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        types=source.read('[Content_Types].xml').decode('utf-8')
        sheet_ids=[int(x) for x in re.findall(r'sheetId="(\d+)"',workbook)]
        sheet_id=max(sheet_ids or [0])+1
        relationship_ids=[int(x) for x in re.findall(r'Id="rId(\d+)"',relationships)]
        rel_id='rId'+str(max(relationship_ids or [0])+1)
        if '<sheet name="Anexos"' in workbook:
            raise RuntimeError('El Excel ya contiene una hoja de Anexos; evite duplicarla.')
        sheet_tag=(f'<sheet name="Anexos" sheetId="{sheet_id}" r:id="{rel_id}"/>')
        if '</sheets>' not in workbook or '</Relationships>' not in relationships:
            raise RuntimeError('La estructura OOXML de la plantilla ha cambiado.')
        workbook=workbook.replace('</sheets>',sheet_tag+'</sheets>',1)
        relationships=relationships.replace('</Relationships>',
            '<Relationship Id="'+rel_id+'" Type="'+ns_rel+'/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',1)
        for part,contenttype in (
            ('/xl/worksheets/sheet2.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
            ('/xl/drawings/drawing2.xml','application/vnd.openxmlformats-officedocument.drawing+xml')):
            if 'PartName="'+part+'"' not in types:
                types=types.replace('</Types>',f'<Override PartName="{part}" ContentType="{contenttype}"/></Types>',1)
        if 'Extension="jpeg"' not in types:
            types=types.replace('</Types>','<Default Extension="jpeg" ContentType="image/jpeg"/></Types>',1)

        def cell(where,val):
            escaped=xml_escape(str(val or ''))
            return f'<c r="{where}" t="inlineStr"><is><t xml:space="preserve">{escaped}</t></is></c>'
        building=str(record.get('proyecto') or '').upper()
        day=str(record.get('fecha') or '')[:10]
        rows=['<row r="1">'+cell('A1','ANEXOS FOTOGRÁFICOS · '+building)+'</row>',
              '<row r="2">'+cell('A2','FECHA: '+day+'   HOJA: 00053')+'</row>']
        anchors=[]
        drawing_relationships=[]
        for i,photo in enumerate(ready,1):
            caption_row=4+(i-1)*13
            image_row=caption_row+1
            rows.append(f'<row r="{caption_row}">'+cell('A'+str(caption_row),'EVIDENCIA '+str(i))+'</row>')
            rid=f'rId{i}'
            drawing_relationships.append(f'<Relationship Id="{rid}" Type="{ns_rel}/image" Target="../media/annex_{i}.jpeg"/>')
            anchors.append(f"""<xdr:twoCellAnchor editAs="oneCell">
                <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{image_row-1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                <xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{image_row+10}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
                <xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{i}" name="Anexo {i}"/><xdr:cNvPicPr/></xdr:nvPicPr>
                <xdr:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
                <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>
                </xdr:twoCellAnchor>""")
        sheet_xml=(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{ns_main}" xmlns:r="{ns_rel}">'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
            '<cols><col min="1" max="8" width="14" customWidth="1"/></cols><sheetData>'+
            ''.join(rows)+'</sheetData><drawing r:id="rId1"/></worksheet>')
        drawings_xml=(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<xdr:wsDr xmlns:xdr="{ns_drawing}" xmlns:a="{ns_a}" xmlns:r="{ns_rel}">'+
            ''.join(anchors)+'</xdr:wsDr>')
        worksheet_relationships=(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{pkg_rel}"><Relationship Id="rId1" Type="{ns_rel}/drawing" Target="../drawings/drawing2.xml"/></Relationships>')
        drawing_rels_xml=(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{pkg_rel}">'+''.join(drawing_relationships)+'</Relationships>')
        changed={'xl/workbook.xml':workbook.encode(),'xl/_rels/workbook.xml.rels':relationships.encode(),
                 '[Content_Types].xml':types.encode()}
        result=io.BytesIO()
        with zipfile.ZipFile(result,'w') as target:
            for info in source.infolist():
                target.writestr(info,changed.get(info.filename,source.read(info.filename)))
            target.writestr('xl/worksheets/sheet2.xml',sheet_xml.encode())
            target.writestr('xl/worksheets/_rels/sheet2.xml.rels',worksheet_relationships.encode())
            target.writestr('xl/drawings/drawing2.xml',drawings_xml.encode())
            target.writestr('xl/drawings/_rels/drawing2.xml.rels',drawing_rels_xml.encode())
            for i,photo in enumerate(ready,1):
                target.writestr(f'xl/media/annex_{i}.jpeg',photo)
        return result.getvalue()


def official_book_excel_from_template(record):
    """Completa Ejemplo.xlsx SIN reconstruir sus diseños, fórmulas ni dibujos.

    A diferencia de cargar/guardar el libro con openpyxl, solo modifica las celdas
    del XML de la hoja. Las demás piezas del paquete OOXML permanecen intactas.
    """
    if is_master_book(record):
        raise ValueError('Esta plantilla corresponde al Libro de Obra Oficial, no al libro del Maestro Mayor.')
    data=obj(record.get('datos'),{})
    if not isinstance(data,dict):
        data={}
    building=str(record.get('proyecto') or data.get('Edificio') or '').strip()
    if not building:
        raise ValueError('El libro no tiene un edificio asignado.')
    start=str(record.get('hora_inicio') or '07:00')[:5]
    end=str(record.get('hora_fin') or '16:00')[:5]
    def as_clock(value):
        try:
            return dt.datetime.strptime(value,'%H:%M').strftime('%I:%M %p').lstrip('0')
        except ValueError:
            return str(value)
    replacements={
        'A2':building.upper(),
        'F3':'EDIFICIO '+building.upper(),
        'A10':'PERSONAL ASIGNADO - '+building.upper(),
        'Y2':'00053',
        'A55':str(data.get('Disciplina_Principal') or DISCIPLINAS_LIBRO_OFICIAL[0]).strip().upper(),
        'H5':str(data.get('Ubicacion') or record.get('frente') or ''),
        'H6':str(data.get('Barrio') or ''),
        'H7':str(data.get('Ciudad') or 'QUITO').upper(),
        'Q6':display_date(record.get('fecha') or ''),
        'Q8':'',
        'U5':'RESIDENTE: '+str(record.get('residente') or ''),
        'U6':'FISCALIZADOR: '+str(data.get('Fiscalizador') or ''),
        'U7':'SUPERINTENDENTE: '+str(data.get('Superintendente') or ''),
        'A46':str(data.get('Clima_Obs') or ''),
        'D250':str(data.get('Novedades') or ''),
    }
    # Nómina según la obra, no los valores de ejemplo del archivo original.
    counts=data.get('Nomina_Conteo') or {}
    if not isinstance(counts,dict):counts={}
    normalized_counts={normalize_template_key(k):v for k,v in counts.items()}
    for index,role in enumerate(OFICIOS_NOMINA_FORMATO):
        r=11+index
        raw_qty=normalized_counts.get(normalize_template_key(role))
        qty=max(0,int(number(raw_qty))) if raw_qty is not None else None
        replacements[f'A{r}']=role
        replacements[f'J{r}']=qty if qty is not None else ''
        replacements[f'L{r}']=as_clock(start) if qty else ''
        replacements[f'N{r}']=as_clock(end) if qty else ''
    # El formulario no tiene apartado 'otros proyectos': se deja en blanco.
    for r in range(26,37):
        replacements[f'J{r}']=''
        replacements[f'L{r}']=''
        replacements[f'N{r}']=''
    # Personal rotativo: alinear las etiquetas con los datos realmente ingresados.
    rotations=data.get('Rotativo_Conteo') or {}
    if not isinstance(rotations,dict):rotations={}
    for i in range(13):
        r=11+i
        name=RUROS_ROTATIVOS_FORMATO[i] if i<len(RUROS_ROTATIVOS_FORMATO) else ''
        replacements[f'Q{r}']=name
        raw_qty=rotations.get(name) if name else None
        qty=max(0,int(number(raw_qty))) if raw_qty is not None else None
        replacements[f'V{r}']=qty if qty is not None else ''
        replacements[f'X{r}']=as_clock(start) if qty else ''
        replacements[f'Y{r}']=as_clock(end) if qty else ''
    machines=data.get('Maquinaria_Conteo') or {}
    if not isinstance(machines,dict):machines={}
    for i in range(8):
        r=40+i
        name=MAQUINARIAS_FORMATO[i]
        raw_qty=machines.get(name)
        qty=max(0,int(number(raw_qty))) if raw_qty is not None else None
        replacements[f'Q{r}']=name
        replacements[f'V{r}']=qty if qty is not None else ''
        replacements[f'X{r}']=as_clock(start) if qty else ''
        replacements[f'Y{r}']=as_clock(end) if qty else ''
    # Clima: mañana/tarde. Se marca solamente la condición elegida.
    weather=normalize_template_key(data.get('Clima_Condicion') or record.get('clima') or '')
    for i,label in enumerate(('SOL','NUBLADO','LLUVIA TENUE','LLUVIA INTENSA','GRANIZO')):
        r=40+i
        replacements[f'B{r}']='X' if weather==label else ''
        replacements[f'E{r}']='X' if weather==label else ''
    # Limpia todos los vistos del ejemplo; escribe únicamente los marcados.
    safety=data.get('Seguridad_Check') or {}
    if not isinstance(safety,dict):safety={}
    lookup={normalize_template_key(k):bool(v) for k,v in safety.items()}
    boxes={
        'E49':'Chalecos','M49':'Guantes','Q49':'Casco','S49':'Gafas','V49':'Mascarilla','Y49':'Auditivo',
        'E50':'Conos','I50':'Cintas','M50':'Rótulos','Q50':'Vallas','S50':'Extintor','V50':'Botiquin',
        'E51':'Polvo','I51':'Ruido','Q51':'Liquidos','S51':'Cerramiento','V51':'Limpieza',
    }
    for cell,label in boxes.items():
        replacements[cell]='X' if lookup.get(normalize_template_key(label),False) else ''
    # Filas de la plantilla por especialidad, conservando los rótulos originales.
    zones={
       'albanileria':list(range(56,81)),
       'estructura':list(range(82,111)),
       'carpinteria':list(range(112,147)),
       'aluminio':[r for r in range(148,159) if r!=150],  # fila 150 sin celdas X/Y en la plantilla
       'hidrosanitario':list(range(161,203)),
       'electrico':list(range(204,232)),
       'otros':list(range(233,249)),
    }
    activities=data.get('Actividades_Ejecutadas') or []
    if not isinstance(activities,list):activities=[]
    # Limpiar también columnas de capítulo/código en filas disponibles.
    # La plantilla conserva rótulos de cabecera pero no datos de ejemplo.
    for zone_rows in zones.values():
        for row in zone_rows:
            for col in ('A','B','V','X','Y'):
                replacements[f'{col}{row}']=''
    used={k:0 for k in zones}
    visible_activity_rows=set()
    for activity in activities:
        if not isinstance(activity,dict) or not str(activity.get('Descripcion') or '').strip():
            continue
        zone=official_template_activity_zone(activity)
        if used[zone]>=len(zones[zone]):
            raise ValueError('La plantilla se quedó sin filas para '+zone+'. Divida las actividades antes de exportar.')
        row=zones[zone][used[zone]]
        used[zone]+=1
        visible_activity_rows.add(row)
        replacements[f'A{row}']=v18_activity_chapter(activity)
        replacements[f'B{row}']=str(activity.get('Descripcion') or '')
        replacements[f'V{row}']=str(activity.get('Area') or '')
        replacements[f'X{row}']=str(activity.get('Unidad') or '')
        replacements[f'Y{row}']=number(activity.get('Cantidad',0))
    # No se trasladan valores de ejemplo como el antiguo proyecto Rubik.
    # Parchar elementos de celda existentes mantiene XML completo y el logo.
    # Conserva fórmulas SUM y actualiza también sus resultados en caché para
    # lectores que muestran la vista previa sin recalcular la hoja.
    cached_totals={
        'J24':sum(replacements.get(f'J{r}',0) or 0 for r in range(11,24)),
        'V24':sum(replacements.get(f'V{r}',0) or 0 for r in range(11,24)),
        'J37':0,
    }
    template=official_template_bytes()
    output=io.BytesIO()
    pattern=re.compile(r'<c\b[^<>]*?r="(?P<ref1>[A-Z]+[0-9]+)"[^<>]*?/>|<c\b[^<>]*?r="(?P<ref2>[A-Z]+[0-9]+)"[^<>]*?>.*?</c>',re.DOTALL)
    with zipfile.ZipFile(io.BytesIO(template),'r') as source, zipfile.ZipFile(output,'w') as dest:
        for info in source.infolist():
            raw=source.read(info.filename)
            if info.filename=='xl/worksheets/sheet1.xml':
                xml=raw.decode('utf-8')
                seen=set()
                def replace(match):
                    ref=match.group('ref1') or match.group('ref2')
                    if ref not in replacements:
                        return match.group(0)
                    seen.add(ref)
                    value=replacements[ref]
                    original=match.group(0)
                    start_tag=original[:original.index('>')+1] if not original.endswith('/>') else original[:-2]+'>'
                    if isinstance(value,(int,float)) and not isinstance(value,bool):
                        start_tag=re.sub(r'\s+t="[^"]+"','',start_tag)
                        payload='<v>'+str(value)+'</v>'
                    else:
                        start_tag=re.sub(r'\s+t="[^"]+"','',start_tag)[:-1]+' t="inlineStr">'
                        safe=xml_escape(str(value),{'"':'&quot;','\'':'&apos;'})
                        payload='<is><t xml:space="preserve">'+safe+'</t></is>'
                    return start_tag+payload+'</c>'
                xml=pattern.sub(replace,xml)
                # Alinear los TRES títulos con el eje del edificio F3:V3.
                # Los rangos A1:Y1 y A2:V2 tienen centros diferentes. Reanclar
                # sus textos en F1 y F2 y unir F:V permite centrar exactamente
                # encima de EDIFICIO <NOMBRE>, reservando A:E para el logotipo.
                for row_num, merged_original, target_merge in (
                    (1,'A1:Y1','F1:V1'), (2,'A2:V2','F2:V2')):
                    source_cell=f'A{row_num}'
                    target_cell=f'F{row_num}'
                    source_match=re.search(r'<c\b[^>]*\br="'+source_cell+r'"[^>]*>.*?</c>',xml)
                    target_match=re.search(r'<c\b[^>]*\br="'+target_cell+r'"[^>]*/>',xml)
                    if not source_match or not target_match or f'ref="{merged_original}"' not in xml:
                        raise RuntimeError('La estructura del encabezado oficial ha cambiado.')
                    source_markup=source_match.group(0)
                    # Mantener el estilo original del título, incluida su alineación.
                    heading_markup=source_markup.replace('r="'+source_cell+'"','r="'+target_cell+'"',1)
                    source_style=re.search(r'\bs="(\d+)"',source_markup)
                    source_replacement=f'<c r="{source_cell}" s="2"/>'
                    xml=xml.replace(source_markup,source_replacement,1)
                    xml=xml.replace(target_match.group(0),heading_markup,1)
                    xml=xml.replace(f'ref="{merged_original}"',f'ref="{target_merge}"',1)
                # Algunas filas de actividades del Excel original están ocultas.
                # Revelar SOLO las que contienen rubros para que se vean al abrir.
                if visible_activity_rows:
                    def show_used_row(match):
                        markup=match.group(0)
                        return re.sub(r'\s+hidden="1"','',markup) if int(match.group(1)) in visible_activity_rows else markup
                    xml=re.sub(r'<row\b[^>]*\br="([0-9]+)"[^>]*>',show_used_row,xml)
                for target,total in cached_totals.items():
                    cache_pattern=r'(<c\s+r="'+target+r'"[^>]*>\s*<f>[^<]*</f>\s*<v>)[^<]*(</v>)'
                    xml,changed=re.subn(cache_pattern,lambda m:m.group(1)+str(total)+m.group(2),xml,count=1)
                    if changed!=1:
                        raise RuntimeError('La fórmula totalizadora '+target+' no existe en la plantilla.')
                missing=set(replacements)-seen
                if missing:
                    raise RuntimeError('La plantilla cambió y faltan celdas: '+', '.join(sorted(missing)[:15]))
                raw=xml.encode('utf-8')
            dest.writestr(info,raw)
    return v197_append_excel_annexes(output.getvalue(),record)


@st.cache_data(ttl=900,show_spinner=False,max_entries=45)
def v18_cached_official_excel(serialized_record):
    # La consulta de la ficha autoriza el dato. Aquí únicamente se genera el archivo.
    return official_book_excel_from_template(json.loads(serialized_record))



@st.cache_data(ttl=900, show_spinner=False, max_entries=24)
def v196_official_excel_preview(serialized_record):
    """Vista HTML del XLSX oficial rellenado (incluye celdas combinadas y estilos).

    Reutiliza los mismos bytes que la descarga; las filas en blanco de actividades
    se compactan para que la consulta no se vuelva interminable. No modifica el XLSX.
    """
    workbook_bytes=v18_cached_official_excel(serialized_record)
    book=openpyxl.load_workbook(io.BytesIO(workbook_bytes),data_only=True,read_only=False)
    sheet=book.active
    # Mantener cabecera, nómina, clima y anexos. En ACTIVIDADES no mostrar
    # los encabezados de las seis disciplinas que NO eligió el usuario.
    # El archivo descargable permanece intacto y conserva el formato completo.
    discipline=str(obj(json.loads(serialized_record).get('datos'),{}).get('Disciplina_Principal') or '').strip().upper()
    activity_sections={
        DISCIPLINAS_LIBRO_OFICIAL[0]: (55,80),
        DISCIPLINAS_LIBRO_OFICIAL[1]: (81,110),
        DISCIPLINAS_LIBRO_OFICIAL[2]: (111,146),
        DISCIPLINAS_LIBRO_OFICIAL[3]: (147,158),
        DISCIPLINAS_LIBRO_OFICIAL[4]: (159,202),
        DISCIPLINAS_LIBRO_OFICIAL[5]: (203,231),
        DISCIPLINAS_LIBRO_OFICIAL[6]: (232,248),
    }
    visible=set(range(1,55)) | set(range(249,271))
    section=activity_sections.get(discipline)
    if section:
        # Incluye el rótulo elegido y solo las filas que contienen un rubro.
        visible.add(section[0])
        for rid in range(section[0]+1,section[1]+1):
            row=sheet[rid]
            if any(str(sheet[f'{column}{rid}'].value or '').strip()
                   for column in ('A','B','V','X','Y')):
                visible.add(rid)
    visible_rows=sorted(visible)
    merged_owner={}
    merges={}
    for rng in sheet.merged_cells.ranges:
        first=(rng.min_row,rng.min_col)
        included=sum(1 for r in visible_rows if rng.min_row<=r<=rng.max_row)
        if not included:continue
        merges[first]=(rng.max_col-rng.min_col+1,included)
        # Solo asignar posiciones existentes en la vista, sin miles de claves extras.
        for r in visible_rows:
            if rng.min_row<=r<=rng.max_row:
                for c in range(rng.min_col,rng.max_col+1):
                    merged_owner[(r,c)]=first
    def color_hex(color):
        if color and color.type=='rgb' and color.rgb and isinstance(color.rgb,str):
            return '#'+color.rgb[-6:]
        return None
    body=[]
    for rid in visible_rows:
        row_cells=[]
        for col in range(1,26):
            owner=merged_owner.get((rid,col))
            if owner and owner!=(rid,col):continue
            cell=sheet.cell(rid,col)
            span,rows=merges.get((rid,col),(1,1))
            attrs=[]
            if span>1:attrs.append(f'colspan="{span}"')
            if rows>1:attrs.append(f'rowspan="{rows}"')
            style=[]
            if cell.fill and cell.fill.patternType=='solid':
                fill=color_hex(cell.fill.fgColor)
                if fill:style.append('background:'+fill)
            fc=color_hex(cell.font.color) if cell.font else None
            if fc:style.append('color:'+fc)
            if cell.font and cell.font.bold:style.append('font-weight:750')
            if cell.font and cell.font.sz:
                style.append('font-size:'+str(min(17,max(9,float(cell.font.sz)*0.92)))+'px')
            if cell.alignment:
                if cell.alignment.horizontal in ('center','right','left'):
                    style.append('text-align:'+cell.alignment.horizontal)
                if cell.alignment.vertical in ('center','top','bottom'):
                    style.append('vertical-align:'+('middle' if cell.alignment.vertical=='center' else cell.alignment.vertical))
            value=cell.value
            if value is None:label=''
            elif isinstance(value,(dt.datetime,dt.date)):label=value.strftime('%d/%m/%Y')
            elif isinstance(value,float) and value.is_integer():label=str(int(value))
            else:label=str(value)
            attrs.append('style="'+';'.join(style)+'"')
            row_cells.append('<td '+ ' '.join(attrs)+'>'+html_escape(label).replace('\n','<br>')+'</td>')
        height=sheet.row_dimensions[rid].height or 15
        body.append('<tr style="height:'+str(max(20,min(45,int(height*1.28))))+'px">'+''.join(row_cells)+'</tr>')
    logo=''
    if sheet._images:
        try:
            im_bytes=sheet._images[0]._data()
            mime='image/png' if im_bytes.startswith(b'\x89PNG') else 'image/jpeg'
            logo='<img class="book-logo" alt="Alpha Builders" src="data:'+mime+';base64,'+base64.b64encode(im_bytes).decode('ascii')+'">'
        except Exception:
            pass
    book.close()
    css="""
    body{margin:0;padding:0;background:#f5f7fb;font-family:Arial,Helvetica,sans-serif;color:#111827}
    .book-hint{padding:12px 14px;color:#334155;background:#eaf0fb;font-size:12px;border-bottom:1px solid #ccd5e5}
    .book-scroller{height:565px;overflow:auto;background:#fff;border:1px solid #cbd5e1;position:relative}
    .book-sheet{position:relative;min-width:1730px;background:white}
    .book-logo{position:absolute;top:8px;left:12px;width:140px;max-height:75px;object-fit:contain;z-index:4;pointer-events:none}
    table{table-layout:fixed;border-collapse:collapse;width:100%;color:#111827;font-family:Arial,sans-serif;font-size:11px}
    td{border:1px solid #848484;padding:2px 4px;line-height:1.14;vertical-align:middle;
       word-break:break-word;white-space:pre-wrap;min-height:18px;background-clip:padding-box}
    tr{page-break-inside:avoid}
    """
    # Las columnas pequeñas mantienen la división original en secciones.
    cols=''.join('<col style="width:'+('135' if i==1 else '47' if i<=15 else '70')+'px">' for i in range(1,26))
    return ('<!doctype html><html lang="es"><head><meta charset="utf-8"><style>'+css+
        '</style></head><body><div class="book-hint">Vista del Excel oficial: desplázate horizontalmente y hacia abajo. '
        'En actividades se muestra solo la disciplina seleccionada; el Excel descargado conserva la plantilla completa.</div>'
        '<div class="book-scroller"><div class="book-sheet">'+logo+'<table><colgroup>'+cols+
        '</colgroup><tbody>'+''.join(body)+'</tbody></table></div></div></body></html>')


def v196_show_official_preview(record):
    import streamlit.components.v1 as components
    st.markdown('#### 📄 Vista previa del Libro de Obra oficial')
    st.caption('Se muestra la disciplina seleccionada, conservando el formato oficial.')
    html=v196_official_excel_preview(json_text(record))
    components.html(html,height=635,scrolling=False)
    images=anexos_fotos_validas(obj(record.get('datos'),{}))
    if images:
        with st.expander(f'📷 Anexos fotográficos ({len(images)})',expanded=False):
            for i,photo in enumerate(images,1):
                file=image_stream(photo)
                if file:st.image(file,caption=f'Anexo {i}',width=360)


def official_books_package(records):
    books=[r for r in records if not is_master_book(r)]
    if not books:
        raise ValueError('No hay libros oficiales compatibles con la plantilla.')
    if len(books)==1:
        return official_book_excel_from_template(books[0]), 'xlsx'
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as bundle:
        used=set()
        for index,record in enumerate(books,1):
            name=re.sub(r'[^A-Za-z0-9_-]+','_',str(record.get('proyecto') or 'Edificio'))[:38]
            day=str(record.get('fecha') or '')[:10]
            basename=f'Libro_Obra_{name}_{day}'
            filename=basename+'.xlsx'
            suffix=2
            while filename.lower() in used:
                filename=f'{basename}_{suffix}.xlsx';suffix+=1
            used.add(filename.lower())
            bundle.writestr(filename, official_book_excel_from_template(record))
    return out.getvalue(),'zip'



def export_controls(table,records,key):
    if not records:return
    with st.expander("Exportar Excel, PDF o CSV",expanded=False):
        st.caption("Se generará únicamente el formato solicitado, con los registros del filtro actual.")
        choices=["Excel (.xlsx)","PDF (.pdf)","CSV (.csv)"]
        if table=="inspecciones" and any(not is_master_book(r) for r in records):
            choices.insert(0,"Libro de Obra — Excel oficial")
        format_=st.selectbox("Formato",choices,key="format_"+key)
        signature=hashlib.sha256(json_text([
            (str(r.get("id")), str(r.get("revision")), str(r.get("updated_at")))
            for r in records
        ]).encode()).hexdigest()
        artifact_key=(key,signature,format_)
        if st.button("Preparar descarga",key="prepare_"+key):
            try:
                with st.spinner("Preparando informe…"):
                    if format_.startswith("Libro oficial"):
                        content,ext=official_books_package(records)
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if ext=="xlsx" else "application/zip"
                    elif format_.startswith("Excel"):
                        content=export_excel(table,records);ext="xlsx";mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    elif format_.startswith("PDF"):content=export_pdf(table,records);ext="pdf";mime="application/pdf"
                    else:content=csv_export(table,records);ext="csv";mime="text/csv"
                st.session_state["_export_artifact"]={"key":artifact_key,"content":content,"ext":ext,"mime":mime}
            except Exception as exc:notice_error(exc)
        artifact=st.session_state.get("_export_artifact")
        if artifact and artifact["key"]==artifact_key:
            st.download_button("Descargar informe",artifact["content"],file_name=f"Alpha_{table}_{now_ec():%Y%m%d}.{artifact['ext']}",mime=artifact["mime"],key="download_"+key)


def styled_detail_table(headers, rows):
    """Tabla responsiva y escapada para lectura, con cabeceras y filas alternadas."""
    if not rows:
        st.caption('No hay filas disponibles en esta sección.')
        return
    if len(rows)>180:
        st.dataframe(pd.DataFrame([[text_value(v) for v in r] for r in rows],columns=headers),
                     hide_index=True,use_container_width=True)
        return
    head=''.join('<th scope="col">'+html_escape(str(h))+'</th>' for h in headers)
    body=[]
    for row in rows:
        cells=[]
        for index,value in enumerate(row):
            value_text=text_value(value)
            cells.append('<td'+(' class="ab-table-primary"' if index==0 else '')+'>'+html_escape(value_text)+'</td>')
        body.append('<tr>'+''.join(cells)+'</tr>')
    st.markdown('<div class="ab-record-table-wrap"><table class="ab-record-table">'
                '<thead><tr>'+head+'</tr></thead><tbody>'+''.join(body)+'</tbody>'
                '</table></div>',unsafe_allow_html=True)


def master_planilla_activities(record):
    """Proyecta libros nuevos y antiguos al formato de seis columnas del Word."""
    raw=obj(record.get('datos'),{})
    if isinstance(raw,dict):
        items=raw.get('Actividades_Maestro') or []
        maestro=str(raw.get('Maestro') or record.get('residente') or '').strip()
    elif isinstance(raw,list):
        items=raw;maestro=str(record.get('residente') or '').strip()
    else:
        items=[];maestro=str(record.get('residente') or '').strip()
    result=[]
    for item in items:
        if not isinstance(item,dict):continue
        activity=str(item.get('Actividad') or item.get('actividad') or '').strip()
        if not activity:continue
        people=item.get('Personal_A_Cargo') or item.get('personal_a_cargo') or []
        if isinstance(people,str):people=[people]
        if not isinstance(people,list):people=[]
        people=[str(p).strip() for p in people if str(p).strip()]
        extra=str(item.get('Personal_Manual') or '').strip()
        if extra:people.extend([part.strip() for part in extra.split(',') if part.strip()])
        responsible=', '.join(dict.fromkeys(people)) or maestro
        area=str(item.get('Area') or item.get('area') or item.get('Piso') or item.get('piso') or '').strip()
        amount=str(item.get('Cantidad') or item.get('cantidad') or '').strip()
        unit=str(item.get('Unidad') or item.get('unidad') or '').strip()
        if not unit and amount:
            matched=re.fullmatch(r'\s*([0-9]+(?:[.,][0-9]+)?)\s*(m2|m²|m3|m³|ml|m|u|und|unidades|kg|glb)\s*',amount,re.I)
            if matched:amount=matched.group(1);unit=matched.group(2)
        result.append({
            'Departamento':str(item.get('Departamento') or item.get('departamento') or ''),
            'Area':area,'Actividad':activity,'Unidad':unit,'Cantidad':amount,'Responsable':responsible,
            'Capitulo':str(item.get('Capitulo') or ''),'Codigo_Rubro':str(item.get('Codigo_Rubro') or ''),
            'Fotos':list_photos(item),'Observaciones':str(item.get('Observaciones') or ''),
        })
    return result


def master_planilla_pdf(record):
    """PDF en orientación horizontal que reproduce la tabla del Word proporcionado.

    Las fotografías se agregan como anexo para no alterar las seis columnas del original.
    """
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import PageBreak
    regular='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
    heavy='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
    font='Helvetica';bold='Helvetica-Bold'
    if os.path.exists(regular) and os.path.exists(heavy):
        if 'AlphaPlanilla' not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont('AlphaPlanilla',regular))
            pdfmetrics.registerFont(TTFont('AlphaPlanillaBold',heavy))
        font='AlphaPlanilla';bold='AlphaPlanillaBold'
    payload=obj(record.get('datos'),{})
    project=str(record.get('proyecto') or (payload.get('Edificio') if isinstance(payload,dict) else '') or 'SIN EDIFICIO')
    date_raw=str(record.get('fecha') or (payload.get('Fecha') if isinstance(payload,dict) else '') or '')[:10]
    try:date_str=dt.date.fromisoformat(date_raw).strftime('%d/%m/%Y')
    except ValueError:date_str=date_raw or '—'
    activities=master_planilla_activities(record)
    out=io.BytesIO()
    doc=SimpleDocTemplate(out,pagesize=landscape(letter),leftMargin=45,rightMargin=45,topMargin=34,bottomMargin=36,
        title='Planilla de avance - '+project,author='Alpha Builders')
    width=doc.width
    # Misma proporción que la tabla de Word, con columna ancha para ACTIVIDAD.
    proportions=[.082,.082,.555,.071,.088,.122]
    widths=[width*p for p in proportions]
    body=ParagraphStyle('plan_body',fontName=font,fontSize=7,leading=9.5,wordWrap='CJK')
    head=ParagraphStyle('plan_head',parent=body,fontName=bold,fontSize=8,leading=10.5)
    title=ParagraphStyle('plan_title',parent=head,fontSize=10,leading=13,alignment=1)
    caption=ParagraphStyle('plan_caption',parent=head,fontSize=10,leading=13,spaceAfter=10)
    def p(v,style=body):
        return Paragraph(html_escape(str(v if v is not None else '')).replace('\n','<br/>'),style)
    header=[p('PLANILLA DE AVANCE - PROYECTO “'+project+'”',title),'','',p('FECHA: '+date_str,head),'','']
    names=['DEPTO.','ÁREA','ACTIVIDAD','UNIDAD','CANTIDAD','RESPONSABLE']
    rows=[header,[p(v,head) for v in names]]
    for a in activities:
        rows.append([p(a[k]) for k in ('Departamento','Area','Actividad','Unidad','Cantidad','Responsable')])
    # La hoja original tiene 28 filas vacías bajo el encabezado; respetarlas en jornadas cortas.
    rows.extend([['']*6 for _ in range(max(0,28-len(activities)))])
    table=Table(rows,colWidths=widths,repeatRows=2,hAlign='CENTER',splitByRow=1,
        rowHeights=[23,23]+[None]*max(len(rows)-2,0))
    table.setStyle(TableStyle([
        ('SPAN',(0,0),(2,0)),('SPAN',(3,0),(5,0)),
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#fff1c8')),
        ('GRID',(0,0),(-1,-1),.6,colors.HexColor('#111827')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('LEFTPADDING',(0,0),(-1,-1),3.3),('RIGHTPADDING',(0,0),(-1,-1),3.3),
        ('TOPPADDING',(0,0),(-1,-1),1.8),('BOTTOMPADDING',(0,0),(-1,-1),1.8),
    ]))
    story=[table]
    images=[(i,a,photo) for i,a in enumerate(activities,1) for photo in a['Fotos']]
    if images:
        story.append(PageBreak())
        story.append(p('ANEXO FOTOGRÁFICO - '+project,caption))
        story.append(p('Fecha: '+date_str))
        story.append(Spacer(1,12))
        for index,a,photo in images:
            raw=image_stream(photo)
            if raw is None:continue
            try:
                img=PDFImage(raw)
                ratio=min(300/img.imageWidth,180/img.imageHeight,1)
                img.drawWidth=img.imageWidth*ratio;img.drawHeight=img.imageHeight*ratio
                info=p('Actividad '+str(index)+': '+a['Actividad']+' - '+a['Area'],head)
                story.append(KeepTogether([info,Spacer(1,5),img,Spacer(1,12)]))
            except Exception:
                continue
    def foot(canvas,document):
        canvas.saveState();canvas.setFont(font,7);canvas.setFillColor(colors.HexColor('#64748b'))
        canvas.drawString(doc.leftMargin,18,'Alpha Builders | Planilla de avance diario')
        canvas.drawRightString(landscape(letter)[0]-doc.rightMargin,18,'Página '+str(document.page))
        canvas.restoreState()
    doc.build(story,onFirstPage=foot,onLaterPages=foot)
    return out.getvalue()


@st.cache_data(ttl=900,show_spinner=False,max_entries=30)
def cached_master_planilla_pdf(serialized_record):
    return master_planilla_pdf(json.loads(serialized_record))



@st.cache_data(ttl=900,show_spinner=False,max_entries=180)
def master_photo_thumb_b64(photo_b64):
    """Miniatura segura, de carga diferida: no inserta imágenes grandes en las tablas."""
    if not isinstance(photo_b64,str) or not photo_b64:
        return ''
    try:
        raw=image_stream(photo_b64)
        if raw is None:return ''
        with Image.open(raw) as im:
            im=ImageOps.exif_transpose(im).convert('RGB')
            im.thumbnail((180,124),Image.Resampling.LANCZOS)
            stream=io.BytesIO()
            im.save(stream,format='JPEG',quality=68,optimize=True)
        return base64.b64encode(stream.getvalue()).decode('ascii')
    except Exception:
        return ''


def master_planilla_preview(rows,project='',day='',show_thumbs=True):
    """Planilla consultiva con las 6 columnas del Word y miniatura opcional.

    Nunca mezcla fotos con los metrajes; la séptima columna es solo vista previa.
    El PDF oficial conserva exactamente las seis columnas originales.
    """
    if not rows:
        st.info('No hay actividades registradas en esta jornada.')
        return
    any_photo=bool(show_thumbs and any(r.get('Fotos') or r.get('Foto_B64') for r in rows))
    ncols=7 if any_photo else 6
    header=['DEPTO.','ÁREA','ACTIVIDAD','UNIDAD','CANTIDAD','RESPONSABLE']
    if any_photo:header.append('FOTO')
    safe_project=html_escape(str(project or 'Sin edificio'))
    safe_day=html_escape(str(day or 'Sin fecha'))
    cols=[9,10,36,9,10,26]
    if any_photo:cols=[8,8,34,7,9,21,13]
    colgroup='<colgroup>'+''.join(f'<col style="width:{v}%">' for v in cols)+'</colgroup>'
    meta=(f'<tr class="ab-planilla-meta"><th colspan="{ncols-1}">'
          f'PLANILLA DE AVANCE - PROYECTO “{safe_project}”</th>'
          f'<th>FECHA: {safe_day}</th></tr>')
    cols_html=''.join('<th scope="col">'+html_escape(label)+'</th>' for label in header)
    trs=[]
    for row in rows:
        values=[row.get(k,'') for k in ('Departamento','Area','Actividad','Unidad','Cantidad','Responsable')]
        tds=''.join('<td>'+html_escape(str(x if x is not None else ''))+'</td>' for x in values)
        if any_photo:
            photos=row.get('Fotos') or ([row.get('Foto_B64')] if row.get('Foto_B64') else [])
            thumb=master_photo_thumb_b64(photos[0]) if photos else ''
            thumb_html=(f'<img class="ab-photo-thumb" loading="lazy" alt="Fotografía de actividad" '
                f'src="data:image/jpeg;base64,{thumb}">' if thumb else '<span class="ab-planilla-empty">Sin foto</span>')
            tds+='<td>'+thumb_html+'</td>'
        trs.append('<tr>'+tds+'</tr>')
    html=('<div class="ab-planilla-frame"><div class="ab-planilla-scroll"><table class="ab-planilla-table">'
          +colgroup+'<thead>'+meta+'<tr class="ab-planilla-head">'+cols_html+'</tr></thead><tbody>'
          +''.join(trs)+'</tbody></table></div></div>')
    st.markdown(html,unsafe_allow_html=True)


def linked_master_planilla_rows(activities):
    """Las actividades compartidas se presentan con los mismos campos de la planilla."""
    return [{
        'Departamento':a.get('Departamento') or '',
        'Area':a.get('Area') or a.get('Piso') or '',
        'Actividad':a.get('Actividad') or '',
        'Unidad':a.get('Unidad') or '',
        'Cantidad':a.get('Avance') or '',
        'Responsable':a.get('Personal') or a.get('Maestro') or '',
        'Foto_B64':a.get('Foto_B64') or '',
    } for a in activities]

def show_master_planilla(record):
    """Fichas de Libro Maestro: planilla visible al seleccionar la jornada."""
    rows=master_planilla_activities(record)
    st.markdown('**Planilla de avance diario**')
    if not rows:
        st.info('Este libro no contiene actividades registradas.')
        return
    project=str(record.get('proyecto') or 'Sin edificio')
    date_raw=str(record.get('fecha') or '')[:10]
    try:date_display=dt.date.fromisoformat(date_raw).strftime('%d/%m/%Y')
    except ValueError:date_display=date_raw
    master_planilla_preview(rows,project,date_display,show_thumbs=True)
    if any(r.get('Fotos') for r in rows):
        with st.expander('📷 Ver fotografías a tamaño completo',expanded=False):
            for r in rows:
                for photo in r.get('Fotos',[]):
                    stream=image_stream(photo)
                    if stream:st.image(stream,caption=r.get('Actividad') or 'Evidencia',width=340)


def show_record(table,record):
    """Un solo bloque visible por ficha, evitando tablas/fotos interminables."""
    if table=='inspecciones' and is_master_book(record):
        show_master_planilla(record)
        return
    meta,sections,photos=sections_for(table,record)
    sections=[(title,heads,rows) for title,heads,rows in sections if rows]
    choices=['🧾 Información general']+[title for title,_,_ in sections]
    if photos:
        choices.append('📷 Fotografías')
    unique='v11_record_section_'+table+'_'+str(record.get('id') or 'sin-id')
    selected=st.selectbox('Contenido del registro',choices,key=unique)
    if selected=='🧾 Información general':
        cells=[]
        for label,value in meta:
            cells.append('<div class="ab-detail-cell"><div class="ab-detail-label">'+html_escape(str(label))+
                         '</div><div class="ab-detail-value">'+html_escape(text_value(value))+'</div></div>')
        st.markdown('<div class="ab-detail-grid">'+''.join(cells)+'</div>',unsafe_allow_html=True)
    elif selected=='📷 Fotografías':
        st.caption(str(len(photos))+' fotografías adjuntas')
        for i,(title,photo) in enumerate(photos):
            stream=image_stream(photo)
            if stream:st.image(stream,caption=title,width=380)
    else:
        section=next(((title,headers,rows) for title,headers,rows in sections if title==selected),None)
        if section:
            _,headers,rows=section
            styled_detail_table(headers,rows)


def history(table, records, repo, key, editable=True):
    """Historial intuitivo: filtros optativos, tarjetas con un clic y una sola ficha."""
    st.markdown('#### 📁 Mis registros')
    if not records:
        st.info('Aún no hay registros. Abre Registrar para crear el primero.');return
    building_choices=sorted({record_building(table,r) for r in records},key=str.casefold)
    building='Todos los edificios';month='Todos los meses';search=''
    with st.expander('🔎 Filtrar por edificio, fecha o palabra',expanded=False):
        a,b=st.columns(2)
        with a:building=st.selectbox('Edificio',['Todos los edificios']+building_choices,key='v17_hist_build_'+key)
        filtered_month_source=[r for r in records if building=='Todos los edificios' or record_building(table,r)==building]
        months=sorted({record_date(table,r)[:7] for r in filtered_month_source if len(record_date(table,r))>=7},reverse=True)
        with b:month=st.selectbox('Mes',['Todos los meses']+months,
            format_func=lambda x:display_month(x) if x!='Todos los meses' else x,key='v17_hist_month_'+key)
        search=st.text_input('Buscar en actividades o responsables',placeholder='Escribe para buscar…',key='v17_hist_search_'+key)
    filtered=[]
    for r in records:
        if building!='Todos los edificios' and record_building(table,r)!=building:continue
        if month!='Todos los meses' and not record_date(table,r).startswith(month):continue
        if search.strip() and search.strip().casefold() not in ' '.join((record_building(table,r),record_title(table,r),record_note(table,r))).casefold():continue
        filtered.append(r)
    if table=='incidencias':
        status=st.segmented_control('Estado',['Todas','Abierta','Cerrada'],default='Todas',key='v17_hist_status_'+key)
        if status and status!='Todas':filtered=[r for r in filtered if r.get('estado')==status]
    st.caption(f'{len(filtered)} registros · Elige una tarjeta para consultar sus datos.')
    with st.expander('📤 Exportar registros',expanded=False):
        export_controls(table,filtered,'hist_export_'+key)
    if not filtered:st.info('No encontramos registros con esos filtros.');return
    ordered=sorted(filtered,key=lambda r:(record_date(table,r),str(r.get('id') or '')),reverse=True)
    page_size=10
    total_pages=(len(ordered)+page_size-1)//page_size
    page_key='v17_hist_page_'+key
    page=min(max(int(st.session_state.get(page_key,0)),0),total_pages-1)
    st.session_state[page_key]=page
    if total_pages>1:
        c1,c2,c3=st.columns([1,1.3,1])
        with c1:st.button('← Anteriores',key='v17_hist_prev_'+key,disabled=page==0,
                          on_click=set_module_view,args=(page_key,max(0,page-1)),use_container_width=True)
        with c2:st.markdown(f'<div class="ab-v17-page">Página {page+1} de {total_pages}</div>',unsafe_allow_html=True)
        with c3:st.button('Siguientes →',key='v17_hist_next_'+key,disabled=page>=total_pages-1,
                          on_click=set_module_view,args=(page_key,min(total_pages-1,page+1)),use_container_width=True)
    visible=ordered[page*page_size:(page+1)*page_size]
    selected_key='v17_hist_selected_'+key
    selected_id=str(st.session_state.get(selected_key) or '')
    st.caption('Pulsa una fecha para desplegar su planilla justo debajo. Pulsa otra vez para cerrar.')
    for r in visible:
        identity=str(r.get('id'))
        label=f"{display_date(record_date(table,r))} · {record_building(table,r)} · {record_title(table,r)}"
        selected=selected_id==identity
        target='' if selected else identity
        st.button(('▾ ' if selected else '▸ ')+label,key='v17_hist_card_'+key+'_'+identity,
            on_click=set_module_view,args=(selected_key,target),use_container_width=True,
            type='primary' if selected else 'secondary')
        if selected:
            selected_record=r
            with st.container(border=True):
                st.markdown('#### '+html_escape(record_title(table,selected_record)))
                st.caption(display_date(record_date(table,selected_record),True)+' · '+record_building(table,selected_record))
                if table=='inspecciones' and not is_master_book(selected_record):
                    view=st.segmented_control('Vista del libro',
                        ['📄 Planilla Excel oficial','🧾 Datos y secciones'],
                        default='📄 Planilla Excel oficial',
                        key='v196_official_view_'+key+'_'+str(selected_record.get('id')),
                        label_visibility='collapsed') or '📄 Planilla Excel oficial'
                    if view=='📄 Planilla Excel oficial':
                        try:v196_show_official_preview(selected_record)
                        except Exception as exc:
                            LOG.exception('Error generando vista previa del Libro de Obra')
                            st.warning('No fue posible mostrar la vista previa. Puedes descargar el Excel oficial.')
                    else:show_record(table,selected_record)
                else:
                    show_record(table,selected_record)
                if table=='inspecciones' and is_master_book(selected_record):
                    try:
                        document=cached_master_planilla_pdf(json_text(selected_record))
                        safe_project=re.sub(r'[^A-Za-z0-9_-]+','_',str(selected_record.get('proyecto') or 'Edificio'))[:42]
                        st.download_button('⬇️ Descargar avance del Maestro — PDF (formato planilla)',
                            data=document,
                            file_name=f'Planilla_Avance_{safe_project}_{str(selected_record.get("fecha") or "")[:10]}.pdf',
                            mime='application/pdf',
                            key='v188_master_plan_pdf_'+key+'_'+str(selected_record.get('id')),
                            on_click='ignore',use_container_width=True)
                    except Exception as exc:notice_error(exc)
                if table=='inspecciones' and not is_master_book(selected_record):
                    friendly=re.sub(r'[^A-Za-z0-9_-]+','_',str(selected_record.get('proyecto') or 'Edificio'))
                    try:
                        file_data=v18_cached_official_excel(json_text(selected_record))
                        st.download_button('⬇️ Descargar Libro de Obra — Excel oficial',
                            data=file_data,
                            file_name=f"Libro_Obra_{friendly}_{str(selected_record.get('fecha'))[:10]}.xlsx",
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            key='v18_official_direct_'+key+'_'+str(selected_record.get('id')),
                            on_click='ignore',use_container_width=True)
                    except Exception as exc:notice_error(exc)
                if editable:
                    ca,cb=st.columns(2)
                    identity=key+':'+str(selected_record['id'])
                    if ca.button('✏️ Editar registro',key='edit_'+identity,use_container_width=True):
                        try:
                            repo._access(table,selected_record,write=True)
                            if table=='checklists':module='Checklist';factory=checklist_data
                            elif table=='rendimientos':module='Rendimientos';factory=rend_data
                            elif table=='incidencias':module='Incidencias';factory=incident_data
                            else:
                                master=is_master_book(selected_record)
                                module='Libro maestro' if master else 'Libro oficial'
                                factory=maestro_data if master else official_data
                            begin_editor(module,repo.user,factory(selected_record),selected_record)
                            st.session_state['_next_route']=module;st.rerun()
                        except Exception as exc:notice_error(exc)
                    with cb.popover('🗑️ Eliminar registro',use_container_width=True):
                        st.warning('Esta acción no se puede deshacer.')
                        if st.button('Confirmar eliminación',key='delete_'+identity):
                            try:repo.delete(table,selected_record);st.session_state['_flash']='Registro eliminado.';st.rerun()
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
    existing={(str(r.get("nombre") or "").strip().upper(),str(r.get("edificio") or "").strip()) for r in repo.own("trabajadores")}
    payload=[]
    for index,record in enumerate(records):
        if record["edificio"] not in repo.projects and not repo.admin:
            raise PermissionError("Edificio no autorizado.")
        identity=(record["nombre"].strip().upper(),record["edificio"])
        if identity in existing:
            continue
        existing.add(identity)
        payload.append({**record,"usuario_email":repo.email,
            "client_request_id":str(uuid.uuid5(uuid.NAMESPACE_URL,repo.email+":"+operation_id+":"+str(index)))})
    if not payload:
        return 0
    try:
        result=database().table("trabajadores").upsert(payload,on_conflict="client_request_id",ignore_duplicates=True).execute()
        if result.data is None:
            raise RuntimeError("No se confirmó la importación.")
        invalidate();return len(result.data)
    except Exception:
        inserted=0
        for item in payload:
            single={k:v for k,v in item.items() if k!="client_request_id"}
            try:
                result=database().table("trabajadores").insert(single).execute()
                if result.data:
                    inserted+=1
            except Exception:
                continue
        if inserted<=0:
            raise
        invalidate();return inserted



def replacement_source_workers(repo, colleague_email, building):
    """Obtener la nómina VISIBLE real del compañero, SQL + legado sin duplicados.

    Conserva permisos por edificio y NO lee datos de otros proyectos.
    """
    colleague_email=norm_email(colleague_email)
    if colleague_email==repo.email:
        raise ValueError('Elige una cuenta distinta de la tuya.')
    user=fetch_user(colleague_email)
    if not user or not active_user(user):
        raise PermissionError('La cuenta del compañero no está activa.')
    if building not in repo.projects or building not in (user.get('edificios') or []):
        raise PermissionError('El edificio debe estar autorizado para ambos usuarios.')
    source=all_rows('trabajadores',filters={'usuario_email':colleague_email,'edificio':building},
                    columns='nombre,cargo,edificio',order='id')
    # El legado se guardó antes de migrar trabajadores a SQL. La pantalla de su
    # propietario utiliza ambas fuentes: la sustitución debe hacer lo mismo.
    owner=type('_NominaOwner',(),{'email':colleague_email})()
    source += [w for w in legacy_worker_backup(owner)
               if str(w.get('edificio') or '').casefold()==building.casefold()]
    cleaned=[];seen=set()
    for row in source:
        name=str(row.get('nombre') or '').strip().upper()
        cargo=str(row.get('cargo') or '').strip().upper()
        if not name or not cargo:
            raise ValueError('Hay nombres o cargos incompletos en la nómina de origen. Corrígelos antes de reemplazar.')
        ident=(name,building.casefold())
        if ident in seen:continue
        seen.add(ident)
        cleaned.append({'nombre':name,'cargo':cargo,'edificio':building})
    if len(cleaned)>1000:raise ValueError('No se permite copiar más de 1.000 trabajadores de una vez.')
    return cleaned


def worker_snapshot(records):
    """Huella estable para impedir copiar un listado distinto del previsualizado."""
    return hashlib.sha256(json_text(records).encode('utf-8')).hexdigest()


def replace_collaborative_workers(repo, colleague_email, building, expected_snapshot):
    """Sustituye TU nómina del edificio sin insertar duplicados preexistentes.

    Importante: PostgREST no ofrece una transacción para varias solicitudes.
    Se archiva ANTES de mutar; primero se insertan los nombres ausentes,
    luego se actualizan las especialidades coincidentes y solo al final
    se eliminan los trabajadores que no aparecen en la nómina de origen.
    Un fallo deja una copia recuperable y el detalle de la fase del error.
    Reejecutar es seguro porque se vuelve a comparar con el estado SQL real.
    """
    repo._verify_write_session()
    if building not in repo.projects:
        raise PermissionError('El edificio ya no está asignado a tu cuenta.')
    source = replacement_source_workers(repo, colleague_email, building)
    if not source:
        raise ValueError('El compañero no tiene trabajadores en este edificio. No se alteró tu nómina.')
    if worker_snapshot(source) != expected_snapshot:
        raise ConflictError('La nómina de origen cambió desde la vista previa. Revísala y confirma nuevamente.')

    old_rows = all_rows('trabajadores', filters={'usuario_email': repo.email, 'edificio': building},
                        columns='*', order='id')
    previous = combined_workers(repo, building)
    previous_count = len(previous)
    old_legacy = [r for r in legacy_worker_backup(repo)
                  if str(r.get('edificio') or '').casefold() == building.casefold()]

    # La operación es por persona y edificio; no por ID ni por especialidad.
    def norm_name(row):
        return ' '.join(str(row.get('nombre') or '').upper().split())

    target = {norm_name(w): w for w in source}
    if len(target) != len(source):
        raise ValueError('Hay nombres repetidos en la nómina de origen. Revísalos antes de sustituirla.')
    old_by_name = {}
    obsolete = []
    for row in old_rows:
        name = norm_name(row)
        if name in old_by_name:
            obsolete.append(row)
        else:
            old_by_name[name] = row
    # Reutilizar la fila original de los nombres coincidentes impide colisiones
    # con restricciones UNIQUE(usuario_email, edificio, nombre) de Postgres.
    missing = [w for name, w in target.items() if name not in old_by_name]
    changed = [(old_by_name[name], w) for name, w in target.items()
               if name in old_by_name and str(old_by_name[name].get('cargo') or '').strip().upper() != w['cargo']]
    obsolete += [row for name, row in old_by_name.items() if name not in target]

    stamp = now_ec().strftime('%Y%m%d_%H%M%S')
    backup_key = ('alpha_nomina_archivo_' + hashlib.sha256(repo.email.encode('utf-8')).hexdigest()[:16]
                  + '_' + stamp + '_' + uuid.uuid4().hex[:8])
    # Respaldar antes de modificar ningún trabajador, incluso si el origen y
    # destino ya comparten nombres. La copia histórica original NO se borra.
    try:
        set_config(backup_key, json_text({
            'version': 18, 'accion': 'sustitucion_colaborativa', 'propietario': repo.email,
            'origen': norm_email(colleague_email), 'edificio': building,
            'fecha': now_ec().isoformat(), 'trabajadores_sql': old_rows,
            'trabajadores_respaldo': old_legacy,
        }))
    except Exception as exc:
        LOG.exception('No se pudo respaldar la nómina antes de la sustitución')
        raise ConflictError('No se pudo guardar el respaldo de seguridad en Supabase. '
                            'No se modificó ningún trabajador. Comprueba los permisos de app_config '
                            'y los registros del servidor antes de intentar de nuevo.') from exc

    # v18.5: toda la sustitución ocurre en UNA transacción PostgreSQL.
    # Requiere ejecutar Alpha_Builders_v18_5_CORRECCION_NOMINA.sql una sola vez.
    # La función SQL conserva IDs coincidentes, copia solo los nombres faltantes,
    # retira los sobrantes y oculta el respaldo del edificio al confirmar el commit.
    try:
        response = database().rpc('alpha_replace_workers_atomic', {
            'p_email': repo.email,
            'p_building': building,
            'p_workers': source,
        }).execute()
        result = response.data
        if isinstance(result, list) and len(result) == 1 and isinstance(result[0], dict):
            result = result[0]
        if not isinstance(result, dict) or int(result.get('final', -1)) != len(source):
            raise RuntimeError('La operación no devolvió la confirmación final esperada.')
        invalidate()
        get_config.clear()
        return previous_count, len(source)
    except Exception as exc:
        invalidate()
        LOG.exception('Fallo en la sustitución atómica de nómina para el edificio %s', building)
        pg_code = str(getattr(exc, 'code', '') or '')
        message = str(getattr(exc, 'message', '') or '')
        if pg_code in {'PGRST202', '42883'} or 'alpha_replace_workers_atomic' in message.lower():
            raise ConflictError(
                'Falta instalar la actualización SQL v18.5 en Supabase. '
                'La nómina actual no se modificó. Ejecuta primero el archivo '
                'Alpha_Builders_v18_5_CORRECCION_NOMINA.sql en SQL Editor.'
            ) from exc
        if pg_code == '23505':
            violated = re.search(r'(?:constraint|index) [\"\']?([a-zA-Z0-9_]+)', message, re.I)
            name = violated.group(1) if violated else 'no identificada'
            raise ConflictError(
                f'La base de datos todavía tiene una restricción de duplicados ({name}). '
                'La sustitución fue revertida por completo; tu nómina sigue intacta. '
                'Comparte ese nombre de restricción para adaptar el SQL a tu esquema.'
            ) from exc
        if pg_code == '23503':
            raise ConflictError(
                'Algunos trabajadores están vinculados mediante relaciones de base de datos '
                '(restricción de clave foránea). PostgreSQL canceló el reemplazo y conserva '
                'todos los registros. Se necesita revisar esa relación antes de continuar.'
            ) from exc
        if isinstance(exc, (ValueError, PermissionError, ConflictError)):
            raise
        safe_code = pg_code if re.fullmatch(r'[0-9A-Z]{5}', pg_code) else 'sin código'
        raise ConflictError(
            f'No se completó la sustitución atómica (código {safe_code}). '
            'PostgreSQL revirtió cualquier modificación de esta operación. '
            'Revisa los registros de Streamlit y comparte el código si el error continúa.'
        ) from exc


def update_legacy_worker(repo, worker, *, new_name=None, new_role=None, new_building=None, delete=False):
    """Editar o retirar personal histórico visible, sin tocar la nómina de otros usuarios.

    Se archiva la versión previa antes de modificar el respaldo que alimenta la
    pantalla. No se elimina el respaldo de seguridad.
    """
    repo._verify_write_session()
    email=repo.email
    old_name=str(worker.get('nombre') or '').strip().upper()
    old_building=str(worker.get('edificio') or '').strip()
    if old_building not in repo.projects:
        raise PermissionError('No tienes permiso para modificar este edificio.')
    key='user_trabajadores_'+email
    entries=obj(get_config(key),[])
    if not isinstance(entries,list):raise ValueError('El archivo histórico de nómina tiene un formato no compatible.')
    match=[]
    for i,item in enumerate(entries):
        if not isinstance(item,dict):continue
        if str(item.get('usuario_email') or email).strip().lower()!=email:continue
        if str(item.get('nombre') or '').strip().upper()==old_name and str(item.get('edificio') or '').strip().casefold()==old_building.casefold():
            match.append(i)
    if len(match)!=1:
        raise ConflictError('El trabajador histórico cambió o está duplicado; actualiza y revisa antes de editar.')
    if not delete:
        new_name=str(new_name or '').strip().upper()
        new_role=str(new_role or '').strip().upper()
        new_building=str(new_building or '').strip()
        if not new_name or not new_role:raise ValueError('El nombre y el cargo son obligatorios.')
        if new_building not in repo.projects:raise PermissionError('El edificio de destino no está autorizado.')
        for w in combined_workers(repo):
            if str(w.get('nombre') or '').strip().upper()==new_name and str(w.get('edificio') or '').strip().casefold()==new_building.casefold():
                if (new_name,new_building.casefold())!=(old_name,old_building.casefold()):
                    raise ValueError('Ya existe un trabajador con ese nombre en el edificio seleccionado.')
    # Respaldo de seguridad exclusivo para el propietario, ANTES del cambio.
    safe_email=hashlib.sha256(email.encode('utf-8')).hexdigest()[:16]
    archive='alpha_nomina_archivo_'+safe_email+'_'+now_ec().strftime('%Y%m%d_%H%M%S')+'_'+uuid.uuid4().hex[:8]
    set_config(archive,json_text({'fecha':now_ec().isoformat(),'edificio':old_building,
                                   'accion':'eliminar' if delete else 'editar','trabajadores':entries}))
    updated=copy.deepcopy(entries)
    if delete:
        updated.pop(match[0])
    else:
        updated[match[0]].update({'nombre':new_name,'cargo':new_role,'edificio':new_building})
    set_config(key,json_text(updated))
    invalidate()
    return True


def personal_module(repo):
    """Gestión de nómina por edificio; copia explícita reemplaza solo el edificio aprobado."""
    page_heading('Gestión de personal','Nómina por edificio',
        'Consulta tus cuadrillas y gestiona el personal por edificio sin perder registros existentes.')
    projects=project_options(repo.user)
    current=sorted(repo.own('trabajadores'),key=lambda r:(str(r.get('edificio') or '').casefold(),str(r.get('nombre') or '').casefold()))
    historic=legacy_worker_backup(repo)
    signatures={(str(r.get('nombre') or '').strip().upper(),str(r.get('edificio') or '').strip().casefold()) for r in current}
    pending=[w for w in historic if (w['nombre'],w['edificio'].casefold()) not in signatures]
    shown=current+pending
    area=st.segmented_control('Sección de personal',
        ['👥 Ver nómina','➕ Registrar','📥 Importar','✏️ Gestionar'],
        default='👥 Ver nómina',key='v16_personal_section',label_visibility='collapsed') or '👥 Ver nómina'
    if area=='👥 Ver nómina':
        st.caption(f"{len(shown)} trabajador(es) · {len({w.get('edificio') for w in shown if w.get('edificio')})} edificio(s) · "
                   f"{len({str(w.get('cargo') or '').upper() for w in shown if w.get('cargo')})} especialidad(es)")
        # Respaldo histórico visible junto a los trabajadores actuales, sin botones de recuperación ni mutaciones automáticas.
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
        st.caption(f'{len(visible)} colaboradores encontrados en tu nómina.')
        if not visible:st.info('No encontramos coincidencias.');return
        display=page_slice(visible,'v8_people_block',25)
        groups=defaultdict(list)
        for w in display:groups[w.get('edificio') or 'Sin edificio asignado'].append(w)
        for name,items in sorted(groups.items(),key=lambda x:x[0].casefold()):
            st.markdown('##### 🏢 '+html_escape(name))
            parts=[]
            for w in items:
                badge=''  # Presentación unificada de la nómina, conservando internamente el origen.
                parts.append('<div class="ab-item"><div><div class="ab-item-title">👷 '+html_escape(str(w.get('nombre') or ''))+
                    '</div><div class="ab-item-tags"><span class="ab-chip">'+html_escape(str(w.get('cargo') or ''))+
                    '</span>'+badge+'</div></div><span class="ab-subtle">'+'En nómina'+'</span></div>')
            st.markdown('<div class="ab-list-wrap">'+''.join(parts)+'</div>',unsafe_allow_html=True)
        with st.expander('📤 Descargar nómina actual',expanded=False):
            # Los respaldos se pueden exportar sin fingir que están confirmados en SQL.
            st.download_button('📄 Exportar lista visible (CSV)',
                pd.DataFrame([{'Nombre':w.get('nombre'),'Cargo':w.get('cargo'),'Edificio':w.get('edificio'),
                'Origen':'Copia histórica (sin sincronizar)' if w.get('_legacy') else 'Nómina actual'} for w in visible]).to_csv(index=False,sep=';').encode('utf-8-sig'),
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
        with st.expander('🤝 Reemplazar nómina con la de un compañero',expanded=False):
            st.info('Se sustituirá únicamente TU nómina del edificio seleccionado. Los demás edificios y la nómina del compañero no se modifican.')
            users=colleagues(repo)
            if not users:
                st.info('No hay compañeros activos con edificios compartidos.')
            else:
                colleague=st.selectbox('Selecciona a un compañero',[None]+users,
                    format_func=lambda u:'Elige a un compañero…' if u is None else full_name(u)+' · '+str(u.get('cargo') or ''),
                    key='v16_collab_worker_source')
                if colleague is not None:
                    common=sorted(set(projects).intersection(obj(colleague.get('edificios'),[])))
                    if not common:
                        st.warning('No tienes edificios compartidos con esta cuenta.')
                    else:
                        chosen_building=st.selectbox('Edificio que deseas actualizar',common,key='v16_collab_worker_building')
                        try:
                            source=replacement_source_workers(repo,colleague['correo'],chosen_building)
                            previous=[w for w in combined_workers(repo,chosen_building)]
                            a,b=st.columns(2)
                            a.metric('Tu nómina visible',len(previous))
                            b.metric('Nómina disponible del compañero',len(source))
                            st.caption('Se cuentan los trabajadores de la tabla actual y del respaldo histórico, sin duplicar nombres por edificio.')
                            if source:
                                st.dataframe(pd.DataFrame(source).rename(columns={'nombre':'Nombre','cargo':'Cargo','edificio':'Edificio'}),
                                    hide_index=True,use_container_width=True)
                                if previous:
                                    backup_csv=pd.DataFrame([{'Nombre':w.get('nombre'),'Cargo':w.get('cargo'),
                                        'Edificio':w.get('edificio')} for w in previous]).to_csv(index=False,sep=';').encode('utf-8-sig')
                                    st.download_button('📥 Descargar copia de mi nómina actual',backup_csv,
                                        file_name='Respaldo_nomina_'+re.sub(r'[^a-zA-Z0-9_-]+','_',chosen_building)+'.csv',
                                        mime='text/csv',key='v16_backup_workers')
                                st.warning(f'Esta operación reemplazará los {len(previous)} trabajadores que aparecen en tu nómina de {chosen_building} por {len(source)} del compañero. Confirma únicamente si deseas realizar el cambio.')
                                confirm=st.checkbox('Confirmo que deseo sustituir mi nómina de este edificio',
                                    value=False,key='v16_worker_replace_approval')
                                if st.button('🔄 Reemplazar nómina de '+chosen_building,
                                    key='v16_collab_replace_confirm',type='primary',disabled=not confirm):
                                    try:
                                        removed,added=replace_collaborative_workers(repo,colleague['correo'],chosen_building,worker_snapshot(source))
                                        st.session_state['_flash']=(f'Nómina actualizada en {chosen_building}: '
                                            f'{removed} trabajador(es) anteriores sustituidos por {added} del compañero.')
                                        st.rerun()
                                    except Exception as exc:
                                        notice_error(exc)
                            else:
                                st.info('Este compañero no tiene trabajadores guardados en este edificio. No se permite reemplazar una nómina por una lista vacía.')
                        except Exception as exc:
                            notice_error(exc)
        return
    if area=='✏️ Gestionar':
        if not shown:
            st.info('No hay trabajadores para gestionar. Registra o importa personal primero.');return
        st.markdown('#### ✏️ Administrar trabajadores')
        st.caption('Puedes editar nombre, especialidad y edificio, o eliminar un trabajador después de confirmarlo.')
        choices=sorted(shown,key=lambda x:(str(x.get('edificio') or ''),str(x.get('nombre') or '')))
        selected=st.selectbox('Buscar trabajador para editar',[None]+choices,
            format_func=lambda r:'Selecciona un trabajador…' if r is None else
                f"{r.get('nombre')} · {r.get('cargo')} · {r.get('edificio')}",key='v17_worker_edit')
        if selected is None:return
        is_legacy=bool(selected.get('_legacy'))
        st.caption('Registro de la nómina histórica' if is_legacy else 'Registro sincronizado en Supabase')
        selected_signature=hashlib.sha256((repo.email+'|'+str(selected.get('nombre'))+'|'+str(selected.get('edificio'))+'|'+str(selected.get('id'))).encode()).hexdigest()[:12]
        with st.form('v17_edit_worker_'+selected_signature):
            name=st.text_input('Nombre completo',value=selected.get('nombre') or '')
            role=st.text_input('Cargo / especialidad',value=selected.get('cargo') or '')
            options=project_options(repo.user,selected.get('edificio'))
            building=st.selectbox('Edificio asignado',options,
                index=options.index(selected['edificio']) if selected.get('edificio') in options else 0)
            submit=st.form_submit_button('💾 Guardar cambios',type='primary',use_container_width=True)
        if submit:
            try:
                if not name.strip() or not role.strip():raise ValueError('Nombre y cargo son obligatorios.')
                if is_legacy:
                    update_legacy_worker(repo,selected,new_name=name,new_role=role,new_building=building)
                else:
                    repo.save('trabajadores',{'nombre':name.strip().upper(),
                        'cargo':role.strip().upper(),'edificio':building},selected)
                st.session_state['_flash']='Cambios guardados correctamente.';st.rerun()
            except Exception as exc:notice_error(exc)
        with st.expander('🗑️ Eliminar este trabajador',expanded=False):
            st.warning('Se quitará al trabajador de TU nómina. Los libros y reportes históricos no se modifican.')
            confirm=st.checkbox('Confirmo que deseo eliminar a '+str(selected.get('nombre') or ''),key='v17_delete_ack_'+selected_signature)
            if st.button('Eliminar trabajador de mi nómina',type='primary',
                         key='v17_delete_worker_'+selected_signature,disabled=not confirm):
                try:
                    if is_legacy:update_legacy_worker(repo,selected,delete=True)
                    else:repo.delete('trabajadores',selected)
                    st.session_state.pop('v17_worker_edit',None)
                    st.session_state['_flash']='Trabajador eliminado de tu nómina.';st.rerun()
                except Exception as exc:notice_error(exc)



MAESTRO_LINKS_CONFIG = "alpha_maestro_links_v14"


def maestro_links():
    """Asociaciones privadas, sin alterar las tablas de registros existentes."""
    raw=obj(get_config(MAESTRO_LINKS_CONFIG),[])
    if not isinstance(raw,list):return []
    links=[];seen=set()
    for item in raw:
        if not isinstance(item,dict):continue
        try:
            master=norm_email(item.get('maestro'))
            staff=norm_email(item.get('supervisor'))
        except ValueError:
            continue
        project=str(item.get('edificio') or '').strip()
        identity=(master,staff,project.casefold())
        if not project or master==staff or identity in seen:continue
        links.append({'maestro':master,'supervisor':staff,'edificio':project})
        seen.add(identity)
    return links


def user_for_link(email):
    """Verificar cargo, estado y edificios ACTUALES; un vínculo antiguo no concede acceso."""
    return fetch_user(email)


def validated_maestro_links(repo):
    if repo.user.get('cargo') not in {'Residente','Asistente'}:
        return []
    results=[]
    for link in maestro_links():
        if link['supervisor']!=repo.email or link['edificio'] not in repo.projects:
            continue
        master=user_for_link(link['maestro'])
        if not master or not active_user(master) or master.get('cargo')!='Maestro Mayor':continue
        if link['edificio'] not in master.get('edificios',[]):continue
        results.append({**link,'nombre':full_name(master)})
    return results


def save_maestro_link(repo,link,unlink=False):
    repo._verify_write_session()
    if not repo.admin:raise PermissionError('Solo el administrador puede modificar vínculos de obra.')
    master_email=norm_email(link['maestro'])
    staff_email=norm_email(link['supervisor'])
    project=str(link['edificio'])
    if unlink:
        current=maestro_links()
        updated=[x for x in current if (x['maestro'],x['supervisor'],x['edificio'])!=(master_email,staff_email,project)]
        set_config(MAESTRO_LINKS_CONFIG,json_text(updated))
        get_config.clear()
        return False
    master=user_for_link(master_email)
    staff=user_for_link(staff_email)
    if not master or not staff or not active_user(master) or not active_user(staff):
        raise ValueError('Las dos cuentas deben estar activas.')
    if master.get('cargo')!='Maestro Mayor' or staff.get('cargo') not in ('Residente','Asistente'):
        raise ValueError('Vincule un Maestro Mayor con un Residente o Asistente.')
    if project not in set(master.get('edificios',[])).intersection(staff.get('edificios',[])):
        raise ValueError('Ambos usuarios deben tener asignado el mismo edificio.')
    current=maestro_links()
    target=(master['correo'],staff['correo'],project.casefold())
    filtered=[x for x in current if (x['maestro'],x['supervisor'],x['edificio'].casefold())!=target]
    if not unlink:filtered.append({'maestro':master['correo'],'supervisor':staff['correo'],'edificio':project})
    set_config(MAESTRO_LINKS_CONFIG,json_text(filtered))
    get_config.clear()
    return not unlink


def admin_maestro_links(repo):
    st.markdown('### 👷 Vincular Maestro Mayor a Residente o Asistente')
    st.caption('El vínculo permite consultar libros y gestionar solicitudes independientes de materiales o planos del edificio en común. '
               'No concede edición ni abre otros proyectos.')
    users=all_rows('usuarios',columns='correo,nombres,apellidos,cargo,edificios,estado',order='correo')
    masters=[u for u in users if active_user(u) and u.get('cargo')=='Maestro Mayor']
    supervisors=[u for u in users if active_user(u) and u.get('cargo') in {'Residente','Asistente'}]
    if not masters or not supervisors:
        st.info('Para crear vínculos se necesita al menos un Maestro Mayor y un Residente o Asistente activos.')
    if masters and supervisors:
        with st.container(border=True):
            a,b=st.columns(2)
            with a:
                master=st.selectbox('Maestro Mayor',masters,index=None,placeholder='Seleccionar Maestro Mayor',format_func=lambda u:full_name(u)+' · '+u['correo'],key='v193_maestro_link_master')
            with b:
                staff=st.selectbox('Residente o Asistente',supervisors,index=None,placeholder='Seleccionar Residente o Asistente',format_func=lambda u:full_name(u)+' · '+u['cargo'],key='v193_maestro_link_staff')
            shared=(sorted(set(obj(master.get('edificios'),[])).intersection(obj(staff.get('edificios'),[])))
                    if master is not None and staff is not None else [])
            if master is None or staff is None:
                st.caption('Selecciona primero las dos cuentas para ver los edificios en común.')
            elif not shared:
                st.warning('Las cuentas seleccionadas no comparten ningún edificio.')
            else:
                building=st.selectbox('Edificio del vínculo',shared,key='v193_maestro_link_build')
                if st.button('🔗 Vincular en este edificio',type='primary',key='maestro_link_save'):
                    try:
                        save_maestro_link(repo,{'maestro':master['correo'],'supervisor':staff['correo'],'edificio':building})
                        st.session_state['_flash']='Vínculo creado. El supervisor ya puede consultar al Maestro Mayor.'
                        st.rerun()
                    except Exception as exc:notice_error(exc)
    links=maestro_links()
    st.markdown(f'#### Vínculos actuales ({len(links)})')
    if not links:st.info('Aún no has vinculado ningún Maestro Mayor.');return
    names={u['correo']:full_name(u) for u in users}
    for index,link in enumerate(links):
        with st.container(border=True):
            c1,c2=st.columns([5,1],vertical_alignment='center')
            with c1:
                st.markdown('**🏗️ '+html_escape(link['edificio'])+'**')
                st.caption('Maestro: '+names.get(link['maestro'],link['maestro'])+'  ·  Supervisión: '+names.get(link['supervisor'],link['supervisor']))
            with c2:
                if st.button('Desvincular',key='maestro_unlink_'+str(index)):
                    try:
                        save_maestro_link(repo,link,unlink=True)
                        st.session_state['_flash']='Vínculo retirado.';st.rerun()
                    except Exception as exc:notice_error(exc)


def colleagues(repo):
    users=cached_rows('usuarios',columns='correo,nombres,apellidos,cargo,edificios,estado',order='correo')
    return [u for u in users if u.get('correo')!=repo.email and active_user(u)
            and set(obj(u.get('edificios'),[])).intersection(repo.projects)
            # El acceso al Maestro Mayor se gestiona exclusivamente por vínculo.
            and not (repo.user.get('cargo') in {'Residente','Asistente'}
                     and u.get('cargo')=='Maestro Mayor')]


# ==============================================================================
# V15 — REQUERIMIENTOS DEL MAESTRO, INDEPENDIENTES DEL LIBRO DE OBRA
# Cada petición es un registro distinto en app_config, evitando sobrescribir las
# solicitudes de otros usuarios. No se necesita modificar tablas ni migraciones.
# ==============================================================================
WORK_REQ_VERSION = "alpha_req_v15_"
WORK_REQ_KINDS = ("Material", "Plano")
WORK_REQ_PRIORITIES = ("Normal", "Alta", "Urgente")
WORK_REQ_STATUSES = ("Pendiente", "En gestión", "Atendido", "No disponible", "Cancelado")


def work_req_prefix(master_email, building):
    master = norm_email(master_email)
    building = str(building or '').strip()
    if not building:
        raise ValueError('Debe seleccionar un edificio válido.')
    # El hash evita publicar correos en las claves de configuración.
    scope = hashlib.sha256((master+'|'+building.casefold()).encode('utf-8')).hexdigest()[:24]
    return WORK_REQ_VERSION + scope + '_'


@st.cache_data(ttl=25, show_spinner=False, max_entries=160)
def cached_work_requests(prefix):
    records=[]
    offset=0
    while True:
        # Solo consultar el prefijo autorizado, no descargar todas las cuentas.
        resp=(database().table('app_config').select('key,value')
              .like('key',prefix+'%').order('key').range(offset,offset+199).execute())
        rows=resp.data or []
        for row in rows:
            item=obj(row.get('value'),{})
            if not isinstance(item,dict) or not row.get('key','').startswith(prefix):
                continue
            if item.get('tipo') in WORK_REQ_KINDS:
                records.append({**item,'_config_key':row['key']})
        if len(rows)<200:break
        offset+=len(rows)
    return sorted(records,key=lambda x:(str(x.get('creado_en') or ''),str(x.get('titulo') or '')),reverse=True)


def work_requests_for(repo,master_email,building):
    """La autorización se comprueba antes de leer cualquier requerimiento."""
    master_email=norm_email(master_email)
    if repo.user.get('cargo')=='Maestro Mayor' and repo.email==master_email:
        if building not in repo.projects:
            raise PermissionError('No tiene asignado ese edificio.')
    elif repo.user.get('cargo') in ('Residente','Asistente'):
        if not any(x['maestro']==master_email and x['edificio']==building
                   for x in validated_maestro_links(repo)):
            raise PermissionError('Este Maestro Mayor no está vinculado a su cuenta para ese edificio.')
    else:
        raise PermissionError('No está autorizado para consultar estos requerimientos.')
    return cached_work_requests(work_req_prefix(master_email,building))


def create_work_request(repo,building,kind,title,quantity,unit,plan_ref,priority,due,notes):
    repo._verify_write_session()
    if repo.user.get('cargo')!='Maestro Mayor':
        raise PermissionError('Solo el Maestro Mayor puede crear requerimientos.')
    if building not in repo.projects:
        raise PermissionError('El edificio no está asignado al Maestro Mayor.')
    if kind not in WORK_REQ_KINDS or priority not in WORK_REQ_PRIORITIES:
        raise ValueError('Tipo o prioridad no válido.')
    title=str(title or '').strip()
    notes=str(notes or '').strip()
    quantity=str(quantity or '').strip()
    unit=str(unit or '').strip()
    plan_ref=str(plan_ref or '').strip()
    if not 3<=len(title)<=180:
        raise ValueError('Describa el material o plano (de 3 a 180 caracteres).')
    if len(notes)>3000 or len(plan_ref)>180 or len(unit)>35 or len(quantity)>60:
        raise ValueError('Uno de los campos supera el límite de longitud.')
    if kind=='Material' and not quantity:
        raise ValueError('Indique la cantidad solicitada de material.')
    if due and not date_valid(str(due)):
        raise ValueError('Seleccione una fecha válida.')
    stamp=now_ec().isoformat(timespec='seconds')
    rid=uuid.uuid4().hex
    payload={'id':rid,'maestro':repo.email,'edificio':building,'tipo':kind,
       'titulo':title,'cantidad':quantity if kind=='Material' else '',
       'unidad':unit if kind=='Material' else '',
       'referencia_plano':plan_ref if kind=='Plano' else '',
       'prioridad':priority,'necesario_para':str(due or ''),'observaciones':notes,
       'estado':'Pendiente','respuesta':'','actualizado_por':'',
       'creado_en':stamp,'actualizado_en':stamp}
    key=work_req_prefix(repo.email,building)+rid
    result=database().table('app_config').insert({'key':key,'value':json_text(payload)}).execute()
    if not result.data:
        raise RuntimeError('No se pudo confirmar el requerimiento. Puede volver a intentarlo.')
    cached_work_requests.clear()
    return payload


def respond_work_request(repo,request,status,response):
    repo._verify_write_session()
    master_email=norm_email(request.get('maestro'))
    building=str(request.get('edificio') or '')
    key=str(request.get('_config_key') or '')
    prefix=work_req_prefix(master_email,building)
    if not key.startswith(prefix):
        raise PermissionError('Identificador de requerimiento inválido.')
    if repo.user.get('cargo') not in ('Residente','Asistente') or not any(
        link['maestro']==master_email and link['edificio']==building
        for link in validated_maestro_links(repo)):
        raise PermissionError('No tiene autorización para responder a este requerimiento.')
    if status not in ('Pendiente','En gestión','Atendido','No disponible'):
        raise ValueError('Estado de atención inválido.')
    response=str(response or '').strip()
    if len(response)>2500:
        raise ValueError('La respuesta no debe superar 2.500 caracteres.')
    current=(database().table('app_config').select('key,value').eq('key',key).execute().data or [])
    if len(current)!=1:
        raise ConflictError('El requerimiento ya no existe. Actualice la consulta.')
    original=current[0]['value']
    entry=obj(original,{})
    if (entry.get('maestro')!=master_email or entry.get('edificio')!=building
            or str(entry.get('id'))!=key[len(prefix):]):
        raise PermissionError('El contenido de esta solicitud no coincide con el edificio vinculado.')
    if entry.get('estado')=='Cancelado':
        raise ValueError('El Maestro Mayor canceló este requerimiento.')
    entry.update({'estado':status,'respuesta':response,'actualizado_por':repo.email,
                  'actualizado_en':now_ec().isoformat(timespec='seconds')})
    result=(database().table('app_config').update({'value':json_text(entry)})
            .eq('key',key).eq('value',original).execute())
    if not result.data:
        raise ConflictError('El requerimiento fue actualizado por otra persona. Consulte de nuevo.')
    cached_work_requests.clear()
    return entry


def cancel_work_request(repo,request):
    repo._verify_write_session()
    if repo.user.get('cargo')!='Maestro Mayor' or repo.email!=request.get('maestro'):
        raise PermissionError('Solo el Maestro Mayor autor puede cancelar su solicitud.')
    building=str(request.get('edificio') or '')
    if building not in repo.projects:
        raise PermissionError('El edificio dejó de estar asignado.')
    key=str(request.get('_config_key') or '')
    prefix=work_req_prefix(repo.email,building)
    if not key.startswith(prefix):raise PermissionError('Solicitud no válida.')
    rows=database().table('app_config').select('value').eq('key',key).execute().data or []
    if len(rows)!=1:raise ConflictError('Esta solicitud ya no existe.')
    raw=rows[0]['value']; item=obj(raw,{})
    if item.get('maestro')!=repo.email or item.get('edificio')!=building:
        raise PermissionError('Solicitud no autorizada.')
    if item.get('estado')=='Atendido':
        raise ValueError('Una solicitud atendida no se puede cancelar.')
    item.update({'estado':'Cancelado','actualizado_por':repo.email,
                 'actualizado_en':now_ec().isoformat(timespec='seconds')})
    result=(database().table('app_config').update({'value':json_text(item)})
             .eq('key',key).eq('value',raw).execute())
    if not result.data:raise ConflictError('La solicitud cambió; vuelva a consultar.')
    cached_work_requests.clear()


def display_work_requests(repo, requests, mode):
    """Fichas compactas, cada solicitud se abre a voluntad."""
    if not requests:
        st.info('Todavía no hay requerimientos de materiales o planos registrados para este edificio.')
        return
    kinds=['Todos','Material','Plano']
    statuses=['Todos']+list(WORK_REQ_STATUSES)
    c1,c2=st.columns(2)
    with c1:kind_filter=st.selectbox('Tipo de solicitud',kinds,key='work_kind_'+mode)
    with c2:status_filter=st.selectbox('Estado',statuses,key='work_status_'+mode)
    filtered=[r for r in requests
              if (kind_filter=='Todos' or r.get('tipo')==kind_filter)
              and (status_filter=='Todos' or r.get('estado')==status_filter)]
    pending=sum(r.get('estado') in ('Pendiente','En gestión') for r in requests)
    st.caption(f'📦 {len(filtered)} solicitudes visibles · {pending} pendientes o en gestión')
    for idx,req in enumerate(filtered):
        key=req.get('_config_key','')
        icon='🧱' if req.get('tipo')=='Material' else '📐'
        status=req.get('estado') or 'Pendiente'
        label=f"{icon} {req.get('titulo') or 'Solicitud'} · {status}"
        with st.expander(label,expanded=False):
            c1,c2,c3=st.columns(3)
            c1.markdown('**Tipo**');c1.write(req.get('tipo',''))
            c2.markdown('**Prioridad**');c2.write(req.get('prioridad','Normal'))
            c3.markdown('**Fecha necesaria**');c3.write(display_date(req.get('necesario_para')) if req.get('necesario_para') else 'Sin fecha')
            if req.get('tipo')=='Material':
                st.markdown('**Cantidad solicitada**: '+str(req.get('cantidad') or '—')+' '+str(req.get('unidad') or ''))
            if req.get('tipo')=='Plano' and req.get('referencia_plano'):
                st.markdown('**Plano o referencia**: '+str(req.get('referencia_plano')))
            if req.get('observaciones'):
                st.markdown('**Observaciones del Maestro Mayor**')
                st.write(req['observaciones'])
            st.caption('Creado: '+str(req.get('creado_en',''))[:16].replace('T',' '))
            if req.get('respuesta'):
                st.markdown('**Respuesta del residente/asistente**')
                st.write(req['respuesta'])
                st.caption('Actualizado: '+str(req.get('actualizado_en') or '')[:16].replace('T',' '))
            if mode=='supervisor' and status!='Cancelado':
                with st.form('work_req_response_'+key):
                    options=['Pendiente','En gestión','Atendido','No disponible']
                    initial=status if status in options else 'Pendiente'
                    new_status=st.selectbox('Estado de atención',options,index=options.index(initial))
                    answer=st.text_area('Respuesta o gestión realizada',value=req.get('respuesta') or '',max_chars=2500)
                    submit=st.form_submit_button('💾 Guardar respuesta',type='primary')
                if submit:
                    try:
                        respond_work_request(repo,req,new_status,answer)
                        st.session_state['_flash']='Requerimiento actualizado.'
                        st.rerun()
                    except Exception as exc:notice_error(exc)
            if mode=='master' and status not in ('Cancelado','Atendido'):
                with st.popover('Cancelar requerimiento'):
                    st.warning('La solicitud dejará de aparecer como pendiente. Los demás datos se conservarán.')
                    if st.button('Confirmar cancelación',key='req_cancel_'+key):
                        try:
                            cancel_work_request(repo,req)
                            st.session_state['_flash']='Solicitud cancelada.'
                            st.rerun()
                        except Exception as exc:notice_error(exc)


def maestro_requests_module(repo):
    """Módulo propio del Maestro Mayor, sin exigir libros diarios."""
    if repo.user.get('cargo')!='Maestro Mayor':
        raise PermissionError('Este módulo pertenece al Maestro Mayor.')
    page_heading('Coordinación de obra','Requerimientos de materiales y planos',
        'Solicita directamente materiales o planos que necesitas para tus edificios. '
        'No forman parte del Libro de Obra.')
    view=st.segmented_control('Requerimientos',['➕ Nueva solicitud','📋 Mis solicitudes'],
                              default='➕ Nueva solicitud',key='req_master_view') or '➕ Nueva solicitud'
    buildings=sorted(repo.projects)
    if not buildings:
        st.warning('Necesitas un edificio asignado para registrar solicitudes.')
        return
    building=st.selectbox('🏢 Edificio',buildings,key='req_master_building')
    if view=='📋 Mis solicitudes':
        display_work_requests(repo,work_requests_for(repo,repo.email,building),'master')
        return
    kind=st.segmented_control('¿Qué necesitas?',['🧱 Material','📐 Plano'],
                              default='🧱 Material',key='req_master_kind') or '🧱 Material'
    selected_kind='Material' if kind=='🧱 Material' else 'Plano'
    with st.form('req_master_new_'+selected_kind,clear_on_submit=True):
        title=st.text_input('Material solicitado' if selected_kind=='Material' else 'Nombre o descripción del plano',max_chars=180)
        qty='';unit='';ref=''
        if selected_kind=='Material':
            c1,c2=st.columns([2,1])
            with c1:qty=st.text_input('Cantidad',placeholder='Ej. 30',max_chars=60)
            with c2:unit=st.selectbox('Unidad',['un','sacos','kg','m','m²','m³','litros','cajas','rollos','otro'])
        else:
            ref=st.text_input('Especialidad, código o referencia (opcional)',
                              placeholder='Ej. Plano eléctrico E-04, última revisión',max_chars=180)
        col_a,col_b=st.columns(2)
        with col_a:priority=st.selectbox('Prioridad',list(WORK_REQ_PRIORITIES))
        with col_b:
            due=st.date_input('Fecha en que se necesita',value=None)
        notes=st.text_area('Observaciones o justificación',max_chars=3000,
                           placeholder='Ubicación, características, uso previsto, detalles de la solicitud…')
        submit=st.form_submit_button('📨 Enviar requerimiento',type='primary',use_container_width=True)
    if submit:
        try:
            create_work_request(repo,building,selected_kind,title,qty,unit,ref,priority,
                                due.isoformat() if due else '',notes)
            st.session_state['_flash']='Requerimiento registrado. Los residentes o asistentes vinculados pueden consultarlo.'
            st.rerun()
        except Exception as exc:notice_error(exc)


def maestro_collaboration(repo):
    links=validated_maestro_links(repo)
    if not links:
        st.info('No tienes un Maestro Mayor vinculado. Pide al administrador que cree un vínculo en Admin → Vínculos Maestro Mayor.')
        return
    options=[None]+sorted({x['maestro'] for x in links})
    names={x['maestro']:x['nombre'] for x in links}
    chosen=st.selectbox('Selecciona un Maestro Mayor',options,
        format_func=lambda v:'Selecciona a quién consultar' if v is None else names.get(v,v),key='maestro_collab_select')
    if chosen is None:
        st.info('Elige un Maestro Mayor para consultar sus solicitudes y sus libros de obra.');return
    authorized=sorted({x['edificio'] for x in links if x['maestro']==chosen})
    building=st.selectbox('Edificio autorizado',authorized,key='maestro_collab_build')
    if not any(x['maestro']==chosen and x['edificio']==building for x in validated_maestro_links(repo)):
        raise PermissionError('Ya no tienes permiso para consultar este Maestro Mayor.')
    tab=st.segmented_control('Información',['📦 Requerimientos','📚 Libros de obra'],
        default='📦 Requerimientos',key='maestro_collab_view') or '📦 Requerimientos'
    if tab=='📦 Requerimientos':
        st.caption('Solicitudes independientes del Libro de Obra: materiales o planos que necesita el Maestro Mayor.')
        display_work_requests(repo,work_requests_for(repo,chosen,building),'supervisor')
        return
    # Los libros solo se consultan al seleccionar específicamente su pestaña.
    records=[r for r in cached_rows('inspecciones',(('usuario_email',chosen),('proyecto',building)))
             if is_master_book(r)]
    history('inspecciones',records,repo,'maestro_libros_'+chosen+'_'+building,editable=False)


def collaborative_module(repo):
    page_heading('Trabajo en equipo','Espacio colaborativo',
        'Elige expresamente a un compañero para consultar sus registros en proyectos compartidos.')
    sections=['👥 Compañeros']
    if repo.user.get('cargo') in ('Residente','Asistente'):
        sections.append('👷 Maestro Mayor')
    view=st.segmented_control('Sección de colaboración',sections,
        default='👥 Compañeros',key='collab_section_v14') or sections[0]
    if view=='👷 Maestro Mayor':
        maestro_collaboration(repo);return
    users=colleagues(repo)
    if not users:
        st.info('No hay compañeros activos que compartan edificios contigo.');return
    user=st.selectbox('¿Con quién quieres colaborar?', [None]+users,
        format_func=lambda u:'Selecciona un compañero' if u is None else full_name(u)+' · '+u.get('cargo',''),
        key='collab_person_v14')
    if user is None:
        st.info('Selecciona un compañero para ver sus registros. No se consulta ningún historial antes de elegirlo.')
        return
    shared=sorted(set(obj(user.get('edificios'),[])).intersection(repo.projects))
    st.caption('Edificios compartidos: '+', '.join(shared))
    types=['checklists','inspecciones','incidencias','rendimientos','trabajadores']
    titles={'checklists':'Checklist','inspecciones':'Libros de obra','incidencias':'Incidencias',
            'rendimientos':'Rendimientos','trabajadores':'Personal'}
    table=st.selectbox('Qué deseas consultar',[None]+types,
        format_func=lambda t:'Selecciona el tipo de registros' if t is None else titles[t],key='collab_records_v14')
    if table is None:return
    records=[r for r in repo.shared(table,user['correo'])
             if record_building(table,r) in shared]
    history(table,records,repo,'colab_'+table+user['correo'],editable=False)


def logout(cookie):
    raw=st.session_state.get("_session_raw")
    if raw:database().table("alpha_sessions").delete().eq("token_hash",digest_token(raw)).execute()
    try:cookie.delete(COOKIE_NAME,key="logout_cookie")
    except KeyError:pass
    for key in list(st.session_state):del st.session_state[key]
    st.rerun()


def account_module(repo,cookie):
    page_heading('Cuenta','Configuración de cuenta',
        'Gestiona tu perfil, solicitudes de acceso y seguridad de inicio de sesión desde un panel más claro y ordenado.')
    user=repo.user
    section=st.segmented_control('Sección de cuenta',['👤 Perfil','🪪 Accesos','🔐 Seguridad'],
        default='👤 Perfil',key='v9_account_section',label_visibility='collapsed') or '👤 Perfil'

    if section=='👤 Perfil':
        c1,c2=st.columns([1.15,1.85])
        with c1:
            with st.container(border=True):
                st.markdown('#### Vista actual')
                photo=image_stream(user.get('foto_b64'))
                if photo:
                    st.image(photo,width=180)
                else:
                    st.info('Aún no has cargado una foto de perfil.')
                st.markdown(f"**{html_escape(full_name(user))}**")
                st.caption(user.get('correo') or '')
                st.caption('Cargo: '+str(user.get('cargo') or 'Sin cargo'))
                st.caption('Edificios: '+(', '.join(user.get('edificios') or []) or 'Sin edificios asignados'))
        with c2:
            with st.container(border=True):
                st.markdown('#### Actualizar datos personales')
                with st.form('profile'):
                    names=st.text_input('Nombres',value=user.get('nombres') or '')
                    surname=st.text_input('Apellidos',value=user.get('apellidos') or '')
                    photo=st.file_uploader('Foto de perfil',type=['png','jpg','jpeg','webp'])
                    submit=st.form_submit_button('Guardar perfil',type='primary',use_container_width=True)
                if submit:
                    try:
                        if not names.strip() or not surname.strip():
                            raise ValueError('Complete nombres y apellidos.')
                        payload={'nombres':names.strip(),'apellidos':surname.strip()}
                        if photo:
                            payload['foto_b64']=image_b64(photo)
                        result=database().table('usuarios').update(payload).eq('correo',repo.email).eq('revision',user['revision']).execute()
                        if not result.data:
                            raise ConflictError('El perfil cambió en otra sesión. Recargue la página.')
                        invalidate();st.session_state['_flash']='Perfil actualizado.';st.rerun()
                    except Exception as exc:
                        notice_error(exc)
        return

    if section=='🪪 Accesos':
        with st.container(border=True):
            st.markdown('#### Solicitar cambios de acceso')
            st.caption('Envía una solicitud al administrador para ajustar tu cargo o los edificios asignados.')
            with st.form('account_access_request'):
                role=st.selectbox('Solicitar cargo',CARGOS_DISPONIBLES,index=CARGOS_DISPONIBLES.index(user['cargo']) if user.get('cargo') in CARGOS_DISPONIBLES else 0)
                buildings=st.multiselect('Solicitar edificios',EDIFICIOS_ALPHA,default=[x for x in user.get('edificios',[]) if x in EDIFICIOS_ALPHA])
                submit=st.form_submit_button('Enviar solicitud al administrador',use_container_width=True)
            if submit:
                try:
                    result=database().table('alpha_access_requests').upsert({'correo':repo.email,'cargo':role,'edificios':buildings,'estado':'Pendiente'},on_conflict='correo').execute()
                    if not result.data:
                        raise RuntimeError('Solicitud no confirmada')
                    st.success('Solicitud registrada en el panel administrador.')
                except Exception as exc:
                    notice_error(exc)
        return

    if section=='🔐 Seguridad':
        with st.container(border=True):
            st.markdown('#### Cambiar contraseña')
            st.caption('Por seguridad, al cambiarla se cerrarán las demás sesiones activas.')
            password_change(repo,cookie)



def password_change(repo,cookie):
    with st.form("change_password"):
        old=st.text_input("Contraseña actual",type="password")
        new=st.text_input("Nueva contraseña (mínimo 10 caracteres)",type="password")
        repeat=st.text_input("Repetir nueva contraseña",type="password")
        submit=st.form_submit_button("Cambiar contraseña y cerrar sesiones",use_container_width=True)
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
    page_heading('Administración','Panel administrador','Selecciona una herramienta. Solo cargaremos la sección que abras.')
    tools=[('🔗 Vincular Maestro Mayor','Conecta equipos por edificio','Vínculos Maestro Mayor'),
           ('📚 Revisar registros','Consulta documentos por módulo','Registros'),
           ('👥 Usuarios y permisos','Administra cargos y edificios','Usuarios y permisos'),
           ('📨 Solicitudes','Aprueba o rechaza solicitudes','Solicitudes'),
           ('🔒 PIN de acceso','Actualiza el código de registro','PIN de acceso'),
           ('🧩 Migración','Revisa incompatibilidades pendientes','Revisión de migración')]
    section=st.session_state.get('v17_admin_section')
    st.markdown('**Herramientas administrativas**')
    for start in (0,3):
        cols=st.columns(3)
        for col,(label,description,destination) in zip(cols,tools[start:start+3]):
            with col:
                st.button(label,key='v17_admin_button_'+destination,
                    on_click=set_module_view,args=('v17_admin_section',destination),
                    use_container_width=True,type='primary' if section==destination else 'secondary')
                st.caption(description)
    if not section:
        st.info('Elige una de las seis herramientas superiores para empezar.');return
    st.markdown('---')
    if section=="Vínculos Maestro Mayor":
        admin_maestro_links(repo)
        return
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
    """Panel ejecutivo: datos reales, consultas ligeras e interfaz responsiva."""
    user = repo.user
    today = now_ec().date()
    date_label = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")[today.weekday()]
    date_label += f", {today.day} de {NOMBRES_MESES[today.month]} de {today.year}"
    email = user["correo"]
    own_filter = (("usuario_email", email),)
    master_role = user.get("cargo") == "Maestro Mayor"
    checks = cached_rows("checklists", own_filter, columns="fecha") if not master_role else []
    books = cached_rows("inspecciones", own_filter, columns="fecha,tipo:datos->>Tipo_Registro")
    yield_rows = cached_rows("rendimientos", own_filter, columns="estado")
    sql_workers = cached_rows("trabajadores", own_filter, columns="nombre,edificio")
    signatures = {(str(w.get("nombre") or "").strip().upper(), str(w.get("edificio") or "").strip().casefold()) for w in sql_workers}
    worker_count = len(sql_workers) + sum(
        (w["nombre"], w["edificio"].casefold()) not in signatures for w in legacy_worker_backup(repo)
    )
    today_text = today.isoformat()
    books_today = [r for r in books if str(r.get("fecha"))[:10] == today_text and
                   bool(r.get("tipo")) and (str(r.get("tipo")) == "Libro_Obra_Maestro") == master_role]
    if not books_today and any(str(r.get("fecha"))[:10] == today_text and not r.get("tipo") for r in books):
        legacy = cached_rows("inspecciones", (("usuario_email", email), ("fecha", today_text)), columns="fecha,datos")
        books_today = [r for r in legacy if is_master_book(r) == master_role]
    check_ok = any(str(r.get("fecha"))[:10] == today_text for r in checks)
    book_ok = bool(books_today)
    goals = [(("Libro maestro" if master_role else "Checklist"), book_ok if master_role else check_ok)]
    if not master_role:
        goals.append(("Libro de obra", book_ok))
    valid = [r for r in yield_rows if r.get("estado") in {"CUMPLE META", "BAJO RENDIMIENTO"}]
    efficiency = (round(100 * sum(r.get("estado") == "CUMPLE META" for r in valid) / len(valid)) if valid else None)
    incidence_projects = tuple(sorted(repo.projects))
    open_incidents = sum(r.get("estado") == "Abierta" for r in
        (cached_rows("incidencias", (), (("proyecto", incidence_projects),), columns="estado") if incidence_projects else []))
    tags = ''.join('<span>'+html_escape(str(x))+'</span>' for x in (user.get('edificios') or []))
    if not tags:
        tags = '<span>Sin edificio asignado</span>'
    def metric_card(icon, title, value, sub, extra=''):
        return ('<div class="ab-overview-card">'
                '<div class="ab-overview-label"><span>'+html_escape(title)+'</span><span>'+icon+'</span></div>'
                '<div class="ab-overview-value">'+html_escape(str(value))+'</div>'
                '<div class="ab-overview-sub">'+html_escape(sub)+'</div>'+extra+'</div>')
    goals_content = ''.join(
        '<div class="ab-status-line"><span>'+html_escape(title)+'</span><span class="ab-status-pill '
        +('ab-status-ok' if done else 'ab-status-wait')+'">'+('✓ Listo' if done else 'Pendiente')+'</span></div>'
        for title,done in goals
    )
    card_goals = ('<div class="ab-overview-card"><div class="ab-overview-label"><span>HITOS DE HOY</span><span>🎯</span></div>'
                  +goals_content+'</div>')
    label = (f'{efficiency}%' if efficiency is not None else '—')
    bar = '<div class="ab-mini-bar"><span style="width:'+str(efficiency or 0)+'%"></span></div>'
    card_performance = metric_card('⚡','META CUMPLIDA',label,
        'Rendimientos evaluados' if efficiency is not None else 'Aún no hay metas verificables',bar)
    card_incidents = metric_card('🚨','INCIDENCIAS ABIERTAS',open_incidents,'En tus edificios asignados')
    card_workers = metric_card('👷','PERSONAL EN NÓMINA',worker_count,'Incluye el respaldo histórico sin duplicar')
    markup = ('<section class="ab-hero" aria-label="Resumen de tu obra">'
              '<div class="ab-hero-top"><div>'
              '<div class="ab-hero-kicker">PANEL EJECUTIVO DE OBRA</div>'
              '<div class="ab-hero-title">Alpha Builders <span style="opacity:.8">/ Portal de obra</span></div>'
              '<div class="ab-hero-desc">'+html_escape(full_name(user))+' · '+html_escape(str(user.get('cargo') or ''))+'</div>'
              '<div class="ab-hero-tags">'+tags+'</div></div>'
              '<div class="ab-hero-date">📅 '+html_escape(date_label)+'</div></div>'
              '<div class="ab-overview-grid">'+card_goals+card_performance+card_incidents+card_workers+'</div>'
              '</section>')
    st.markdown(markup,unsafe_allow_html=True)


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
            s_left,s_photo,s_right=st.columns([1,5,1])
            with s_photo:st.image(photo, use_container_width=True)
        else:
            profile_file = next((name for name in ("perfil.jpg", "perfil.png", "perfil.jpeg", "avatar.png") if os.path.isfile(name)), None)
            if profile_file:
                s_left,s_photo,s_right=st.columns([1,5,1])
                with s_photo:st.image(profile_file, use_container_width=True)
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


def rendimiento_personal_historico(record):
    """Personal real de la fuente original, tanto nueva como histórica.

    En v4/v5/v6 el listado se conserva en datos.trabajadores. Versiones anteriores
    guardaban solamente las columnas trabajador / cargo_obrero de la tabla SQL.
    No modifica ni rellena datos históricos en Supabase.
    """
    pay=rendimiento_payload(record)
    candidates=pay.get('trabajadores')
    if isinstance(candidates, dict):
        candidates=list(candidates.values())
    if not isinstance(candidates, list) or not candidates:
        candidates=[]
        for alt in ('personal_asignado','personal','cuadrilla'):
            value=pay.get(alt)
            if isinstance(value,dict):value=list(value.values())
            if isinstance(value,list) and value:
                candidates=value
                break
    people=[]
    for person in candidates:
        if isinstance(person,dict):
            name=str(person.get('nombre') or person.get('Nombre') or
                     person.get('trabajador') or person.get('Trabajador') or '').strip()
            role=str(person.get('cargo') or person.get('Cargo') or
                     person.get('cargo_obrero') or person.get('especialidad') or '').strip()
        elif isinstance(person,str):
            name=person.strip()
            role=''
        else:
            continue
        if name or role:
            people.append({'nombre':name,'cargo':role})
    # En datos v4-v6 puede no estar el personal o tener una lista parcial.
    # Las columnas SQL originales se mantienen como respaldo; si el listado
    # tiene un trabajador sin cargo, se puede completar cuando solo hay uno.
    stored_name=str(record.get('trabajador') or '').strip()
    stored_role=str(record.get('cargo_obrero') or '').strip()
    if not people:
        if stored_name or stored_role:
            # Se conserva el texto original íntegro: no separar por comas,
            # porque algunos nombres/cargos históricos pueden contenerlas.
            people.append({'nombre':stored_name,'cargo':stored_role})
    elif len(people)==1:
        if not people[0]['nombre'] and stored_name:
            people[0]['nombre']=stored_name
        if not people[0]['cargo'] and stored_role:
            people[0]['cargo']=stored_role
    return people


def _rend_planilla_data(record):
    """Planilla de rendimiento: datos actuales e históricos, sin inventar personal."""
    pay=rendimiento_payload(record)
    value=rendimiento_values(record)
    code=pay.get('codigo_rubro') or (record.get('rubro') if v18_catalog_record(record.get('rubro')) else '')
    description=pay.get('descripcion_rubro') or (v18_catalog_record(code) or {}).get('Descripcion') or record.get('rubro') or ''
    unit=record.get('unidad') or ''
    photos={p:(pay.get('fotos_'+p) or ([pay['foto_'+p]] if pay.get('foto_'+p) else [])) for p in ('manana','mediodia','tarde')}
    people=rendimiento_personal_historico(record)
    # Una línea por encargado mantiene el orden que tenía cada cargo en SQL.
    encargados='\n'.join(person['nombre'] or 'Sin nombre registrado' for person in people) if people else 'Sin registro'
    cargos='\n'.join(person['cargo'] or 'Sin cargo registrado' for person in people) if people else 'Sin registro'
    return {'Fecha':display_date(record_date('rendimientos',record)),
            'Edificio':value['edificio'],'Área':pay.get('area') or value['piso'],
            'Rubro':code or '—','Descripción':description,
            'Encargado':encargados,'Cargo':cargos,'Unidad':unit,
            'Inicio':pay.get('hora_inicio') or '—','Fin':pay.get('hora_fin') or '—',
            'Tiempo muerto':pay.get('hora_muerta') or ('00:00' if pay.get('_rend_v6') else '—'),
            'Lunch':'Sí (15 min)' if pay.get('lunch') else 'No',
            'Almuerzo':'Sí (60 min)' if pay.get('almuerzo') else 'No',
            'Rendimiento':f"{value['productividad']:.3f} {unit}/h" if value['productividad'] is not None else 'Sin verificar',
            'Mañana':photos['manana'],'Mediodía':photos['mediodia'],'Tarde':photos['tarde']}


REND_PLANILLA_HEADERS=['Fecha','Edificio','Área','Rubro','Descripción','Encargado','Cargo','Unidad',
     'Inicio','Fin','Tiempo muerto','Lunch','Almuerzo','Rendimiento',
     'Foto mañana','Foto mediodía','Foto tarde']


def _rend_photo_cell(images):
    if not images:return '<span class="ab-planilla-empty">Sin foto</span>'
    thumb=master_photo_thumb_b64(images[0])
    if not thumb:return '<span class="ab-planilla-empty">No disponible</span>'
    return (f'<img class="ab-photo-thumb" loading="lazy" alt="Evidencia fotográfica" '
            f'src="data:image/jpeg;base64,{thumb}">')


def rendimiento_planilla_preview(records):
    """Planilla horizontal con 14 columnas; miniaturas ligeras, scroll horizontal en móvil."""
    headers=''.join('<th scope="col">'+html_escape(h)+'</th>' for h in REND_PLANILLA_HEADERS)
    rows=[]
    for rec in records:
        row=_rend_planilla_data(rec)
        cells=''.join('<td>'+html_escape(str(row[k] or '—'))+'</td>' for k in (
            'Fecha','Edificio','Área','Rubro','Descripción','Encargado','Cargo','Unidad','Inicio','Fin',
            'Tiempo muerto','Lunch','Almuerzo','Rendimiento'))
        cells+=''.join('<td>'+_rend_photo_cell(row[k])+'</td>' for k in ('Mañana','Mediodía','Tarde'))
        rows.append('<tr>'+cells+'</tr>')
    st.markdown('<div class="ab-rend-table-wrap"><table class="ab-rend-table">'
        '<thead><tr>'+headers+'</tr></thead><tbody>'+''.join(rows)+'</tbody></table></div>',
        unsafe_allow_html=True)


def history_rend_by_building(records,repo):
    """Consulta directa: buscar, seleccionar ficha y revisar planilla."""
    page_heading('Consulta de obra','Planilla de rendimientos',
        'Filtra por edificio, mes o rubro; abre una ficha antes de recorrer la tabla.')
    if not records:
        st.info('No hay rendimientos registrados todavía.');return
    buildings=sorted({record_building('rendimientos',r) for r in records},key=str.casefold)
    a,b,c=st.columns([1.3,1.05,1.15])
    with a:building=st.selectbox('Edificio',['Todos los edificios']+buildings,key='v193_rend_building')
    months=sorted({record_date('rendimientos',r)[:7] for r in records
                    if (building=='Todos los edificios' or record_building('rendimientos',r)==building)
                    and len(record_date('rendimientos',r))>=7},reverse=True)
    with b:month=st.selectbox('Mes',['Todos los meses']+months,
        format_func=lambda m:display_month(m) if m!='Todos los meses' else m,key='v193_rend_month')
    available_days=sorted({record_date('rendimientos',r) for r in records
        if len(record_date('rendimientos',r))==10
        and (building=='Todos los edificios' or record_building('rendimientos',r)==building)
        and (month=='Todos los meses' or record_date('rendimientos',r).startswith(month))},reverse=True)
    with c:day=st.selectbox('Día',['Todos los días']+available_days,
        format_func=lambda value:display_date(value) if value!='Todos los días' else value,
        key='v196_rend_day')
    search=st.text_input('Buscar por código, descripción, área o trabajador',
        placeholder='Ej.: CAP-05-09, enlucido, subsuelo…',key='v193_rend_query')
    filtered=[]
    for rec in records:
        if building!='Todos los edificios' and record_building('rendimientos',rec)!=building:continue
        if month!='Todos los meses' and not record_date('rendimientos',rec).startswith(month):continue
        if day!='Todos los días' and record_date('rendimientos',rec)!=day:continue
        data=_rend_planilla_data(rec)
        search_parts=[data['Rubro'],data['Descripción'],data['Área'],data['Edificio'],data['Encargado'],data['Cargo']]
        if search.strip() and search.strip().casefold() not in ' '.join(str(p) for p in search_parts).casefold():continue
        filtered.append(rec)
    if not filtered:
        st.info('No hay registros para estos filtros.');return
    ordered=sorted(filtered,key=lambda r:(record_date('rendimientos',r),str(r.get('id') or '')),reverse=True)
    # Selección inmediata debajo del buscador, sin tener que bajar al final de la tabla.
    selected=st.selectbox('Abrir registro para editar o consultar fotografías',
        [None]+ordered,index=0,
        format_func=lambda r:'Seleccionar registro' if r is None else
          f"{display_date(record_date('rendimientos',r))} · {_rend_planilla_data(r)['Rubro']} · {_rend_planilla_data(r)['Área']}",
        key='rend_v194_selected')
    if selected is not None:
        data=_rend_planilla_data(selected)
        with st.container(border=True):
            st.markdown('#### '+html_escape(str(data['Rubro']))+' · '+html_escape(str(data['Descripción'])))
            st.caption(display_date(record_date('rendimientos',selected))+' · '+str(data['Edificio'])+' · '+str(data['Área']))
            a,b,c=st.columns(3)
            a.metric('Inicio',data['Inicio']);b.metric('Fin',data['Fin']);c.metric('Rendimiento',data['Rendimiento'])
            st.markdown('**Personal responsable:**')
            person_left,person_right=st.columns(2)
            with person_left:st.caption('Encargado');st.text(data['Encargado'])
            with person_right:st.caption('Cargo');st.text(data['Cargo'])
            for period,label in [('Mañana','🌅 Mañana'),('Mediodía','☀️ Mediodía'),('Tarde','🌇 Tarde')]:
                photos=data[period]
                if photos:
                    with st.expander(label+' · Ver fotografías',expanded=False):
                        for i,photo in enumerate(photos):
                            image_bytes=image_stream(photo)
                            if image_bytes:st.image(image_bytes,caption=f'{label} · Foto {i+1}',width=380)
            a,b=st.columns(2)
            if a.button('✏️ Editar',key='rend_v193_edit_'+str(selected['id']),use_container_width=True):
                try:
                    repo._access('rendimientos',selected,write=True)
                    begin_editor('Rendimientos',repo.user,rend_data(selected),selected)
                    st.session_state['_next_route']='Rendimientos';st.rerun()
                except Exception as exc:notice_error(exc)
            with b.popover('🗑️ Eliminar registro',use_container_width=True):
                st.warning('La eliminación es permanente.')
                if st.button('Confirmar eliminación',key='rend_v193_delete_'+str(selected['id'])):
                    try:repo.delete('rendimientos',selected);st.session_state['_flash']='Rendimiento eliminado.';st.rerun()
                    except Exception as exc:notice_error(exc)

    st.caption(f'{len(filtered)} registros · La tabla permite desplazamiento horizontal en celular.')
    page_size=10
    count=(len(ordered)+page_size-1)//page_size
    page_key='rend_v193_page'
    page_filter=(building,month,day,search.strip().casefold())
    if st.session_state.get('v196_rend_page_filter')!=page_filter:
        st.session_state[page_key]=0
        st.session_state['v196_rend_page_filter']=page_filter
    page=max(0,min(int(st.session_state.get(page_key,0)),count-1))
    a,b,c=st.columns([1,1,1])
    with a:st.button('← Anterior',key='rend_v193_prev',disabled=page==0,
        on_click=set_module_view,args=(page_key,max(0,page-1)),use_container_width=True)
    with b:st.caption(f'Página {page+1} de {count}')
    with c:st.button('Siguiente →',key='rend_v193_next',disabled=page>=count-1,
        on_click=set_module_view,args=(page_key,min(count-1,page+1)),use_container_width=True)
    visible=ordered[page*page_size:(page+1)*page_size]
    rendimiento_planilla_preview(visible)
    with st.expander('⬇️ Exportar registros filtrados',expanded=False):
        export_controls('rendimientos',filtered,'rend_v193_export')


def record_actions(table, record, repo, key, editable=True):
    """Mostrar u ocultar la ficha en un clic; evita leer fotos de otros registros."""
    open_key='_v10_record_detail_'+table+':'+repo.email
    active=st.session_state.get(open_key)==key
    if st.button('▴  Ocultar detalle' if active else '▸  Ver detalle, fotografías y acciones',
                 key='v10_open_yield_'+key,use_container_width=True):
        active=not active
        st.session_state[open_key]=key if active else None
    if active:
        with st.container(border=True):
            show_record(table,record)
            if editable:
                a,b=st.columns(2)
                if a.button('✏️ Editar registro',key=key+'_edit',use_container_width=True):
                    try:
                        repo._access(table,record,write=True)
                        factories={'checklists':('Checklist',checklist_data),
                                   'rendimientos':('Rendimientos',rend_data),
                                   'incidencias':('Incidencias',incident_data)}
                        module,factory=factories[table]
                        begin_editor(module,repo.user,factory(record),record)
                        st.session_state['_next_route']=module
                        st.rerun()
                    except Exception as exc:notice_error(exc)
                with b.popover('🗑️ Eliminar registro',use_container_width=True):
                    st.warning('La eliminación es permanente.')
                    if st.button('Confirmar eliminación',key=key+'_delete'):
                        try:
                            repo.delete(table,record)
                            st.session_state['_flash']='Registro eliminado.'
                            st.rerun()
                        except Exception as exc:notice_error(exc)


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
        '<div class="ab-record-table-wrap"><table class="ab-record-table"><thead><tr>'
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




def set_module_view(key, value):
    """Callback seguro: cambia pestaña antes de crear el segmented_control.

    Streamlit rechaza actualizar desde el cuerpo principal una clave ya asociada
    a un widget renderizado durante esa misma ejecución.
    """
    st.session_state[key] = value

def module_original(label, repo):
    """Un módulo = una vista. Registrar y consultar sin pantalla de resumen redundante."""
    spec = {
        'Checklist':('checklists','Checklist',checklist_data,render_checklist,'Checklist de obra','Verificaciones y supervisión diaria','Nueva jornada'),
        'Libro de Obra':('inspecciones','Libro oficial',official_data,render_official,'Libro de obra oficial','Registro técnico de la jornada','Nuevo libro de obra'),
        'Libro de Obra Maestro':('inspecciones','Libro maestro',maestro_data,render_maestro,'Libro de obra maestro','Cuadrillas, actividades y avances ejecutados','Nuevo libro maestro'),
        'Levantamiento de Incidencias':('incidencias','Incidencias',incident_data,render_incident,'Incidencias de obra','Registro y seguimiento de novedades','Nueva incidencia'),
        'Control de Rendimiento':('rendimientos','Rendimientos',rend_data,render_rend,'Control de rendimientos','Avances y productividad del personal','Nuevo rendimiento'),
    }
    table,module,factory,renderer,title,subtitle,button=spec[label]
    page_heading('Control de obra',title,subtitle)
    selection_key='v14_mode_'+module.replace(' ','_')
    editor_active=bool(st.session_state.get('_editors',{}).get(module))
    options=['➕ Registrar','📚 Consultar']
    if editor_active and st.session_state.get(selection_key)!='➕ Registrar':
        st.session_state[selection_key]='➕ Registrar'
    choice=st.segmented_control('Vista de este módulo',options,default='➕ Registrar',key=selection_key,
                                 label_visibility='collapsed') or '➕ Registrar'
    if choice=='➕ Registrar':
        editor=st.session_state.get('_editors',{}).get(module)
        if editor:
            selected=editor_picker(module,repo.user,factory)
        else:
            st.caption('Abre el formulario para comenzar. Se conserva un borrador si sales antes de guardar.')
            if st.button('➕ '+button,type='primary',key='original_new_'+module):
                begin_editor(module,repo.user,factory())
                rerun_selected_module()
            drafts=cached_rows('alpha_drafts',(('usuario_email',repo.email),('modulo',module)),
                columns='slot,revision,updated_at',order='updated_at')
            if drafts:
                with st.expander(f'📁 Recuperar borradores ({len(drafts)})',expanded=False):
                    choices={f"Borrador {i+1} · {hora_borrador_ecuador(r['updated_at'])}":r for i,r in enumerate(reversed(drafts))}
                    selected_name=st.selectbox('Selecciona borrador',list(choices),key='restore_choice_'+module)
                    if st.button('Continuar borrador',key='restore_button_'+module):
                        draft=choices[selected_name]
                        recovered=all_rows('alpha_drafts',filters={'usuario_email':repo.email,'modulo':module,'slot':draft['slot']},
                                           columns='slot,datos,revision')
                        if recovered:
                            row=recovered[0]
                            st.session_state.setdefault('_editors',{})[module]={'slot':row['slot'],'data':row['datos'],'revision':row['revision']}
                            rerun_selected_module()
                        else:st.warning('El borrador ya no existe.')
            selected=None
        if selected:
            st.markdown('#### ✏️ Editando un registro' if selected.data.get('_original') else '#### 📝 Nuevo registro')
            renderer(selected,repo)
        return
    # Consulta diferida: las fotos, las secciones y la tabla se procesan solo bajo demanda.
    records=repo.shared(table) if table=='incidencias' else repo.own(table)
    if table=='inspecciones':
        records=[r for r in records if is_master_book(r)==(module=='Libro maestro')]
    if table=='rendimientos':
        history_rend_by_building(records,repo)
    elif table=='incidencias':
        section=st.segmented_control('Presentación de incidencias',['📁 Fichas','📊 Matriz'],
            default='📁 Fichas',key='v11_incidents_view') or '📁 Fichas'
        if section=='📊 Matriz':incidencias_original_matrix(records,repo)
        else:history(table,records,repo,'original_'+module)
    else:
        history(table,records,repo,'original_'+module)


# Los widgets dentro de este fragmento no reconstruyen el dashboard al cambiar.
# Conserva un solo archivo y funciona sin dependencia adicional en versiones modernas.
_FRAGMENT = getattr(st, "fragment", None)


@(_FRAGMENT if _FRAGMENT is not None else (lambda fn: fn))
def render_active_module(name, repo, cookie):
    role = repo.user.get("cargo")
    if name in {"Libro de Obra Maestro","Requerimientos"} and role != "Maestro Mayor":
        raise PermissionError("Este módulo es exclusivo del Maestro Mayor.")
    if role == "Maestro Mayor" and name in {"Checklist", "Libro de Obra", "Levantamiento de Incidencias"}:
        raise PermissionError("El módulo no está habilitado para el cargo Maestro Mayor.")
    if name == "Panel Admin" and not repo.admin:
        raise PermissionError("La administración requiere permisos activos.")
    if name == "Requerimientos":
        maestro_requests_module(repo)
    elif name == "Personal a Cargo":
        personal_module(repo)
    elif name == "Catálogo de Rubros":
        v18_catalog_page(repo)
    elif name == "Colaborativo":
        collaborative_module(repo)
    elif name == "Mi cuenta":
        account_module(repo,cookie)
    elif name == "Panel Admin":
        admin_module(repo)
    else:
        module_original(name, repo)



# ==============================================================================
# V12 — ESCRITORIO PERSONAL (datos separados por cuenta, sin nuevas tablas)
# ==============================================================================
HOME_WIDGETS = {
    "calendar": ("📅", "Calendario", "Agenda y eventos de obra"),
    "tasks": ("✅", "Tareas", "Pendientes y prioridades"),
    "notes": ("📝", "Mis notas", "Ideas y apuntes privados"),
    "weather": ("🌤️", "Clima", "Tiempo real por ciudad"),
    "calculator": ("🧮", "Calculadora", "Operaciones rápidas"),
    "construction": ("📐", "Cálculos de obra", "Metrajes y rendimientos"),
    "converter": ("🔁", "Conversor", "Unidades habituales"),
    "bookmarks": ("🔗", "Enlaces", "Páginas frecuentes"),
}
HOME_DEFAULT = ["calendar", "tasks", "notes", "weather"]
HOME_CITIES = {
    "Quito": (-0.1807, -78.4678), "Guayaquil": (-2.1894, -79.8891),
    "Cuenca": (-2.9001, -79.0059), "Ambato": (-1.2491, -78.6167),
    "Ibarra": (0.3517, -78.1223), "Otavalo": (0.2346, -78.2625),
    "Manta": (-0.9676, -80.7089), "Riobamba": (-1.6635, -78.6547),
}


def home_storage_key(repo, section):
    """No incluye correos en las claves visibles de configuración."""
    if section not in {"layout", "tasks", "notes", "events", "bookmarks", "city"}:
        raise ValueError("Sección personal desconocida.")
    fingerprint = hashlib.sha256(norm_email(repo.email).encode("utf-8")).hexdigest()[:32]
    return f"alpha_home_v12_{fingerprint}_{section}"


def home_load(repo, section, fallback):
    value = get_config(home_storage_key(repo, section))
    parsed = obj(value, fallback) if isinstance(fallback, (dict, list)) else value
    if isinstance(fallback, (dict, list)) and not isinstance(parsed, type(fallback)):
        return copy.deepcopy(fallback)
    return copy.deepcopy(fallback) if parsed is None else parsed


def home_save(repo, section, data):
    """Confirma la identidad activa antes de escribir; sin sobrescribir otros módulos."""
    repo._verify_write_session()
    if section == "city":
        if data not in HOME_CITIES:
            raise ValueError("La ciudad no está disponible.")
        serialized = str(data)
    else:
        serialized = json_text(data)
        if len(serialized.encode("utf-8")) > 100_000:
            raise ValueError("El espacio personal alcanzó su capacidad. Reduce el contenido.")
    set_config(home_storage_key(repo, section), serialized)


def home_active_widgets(repo):
    layout = home_load(repo, "layout", HOME_DEFAULT)
    if not isinstance(layout, list):
        return HOME_DEFAULT.copy()
    selected = []
    for widget in layout:
        if isinstance(widget, str) and widget in HOME_WIDGETS and widget not in selected:
            selected.append(widget)
    return selected or HOME_DEFAULT.copy()


def rerun_home():
    if getattr(st, "fragment", None) is not None:
        try:
            st.rerun(scope="fragment")
            return
        except (ValueError, RuntimeError):
            pass
    st.rerun()


def home_widget_title(widget):
    icon, name, description = HOME_WIDGETS[widget]
    st.markdown('<div class="ab-home-widget-head"><div class="ab-home-widget-icon">'
        + icon + '</div><div><h3>' + html_escape(name) + '</h3><p>'
        + html_escape(description) + '</p></div></div>', unsafe_allow_html=True)


def valid_home_records(items, max_items=150):
    if not isinstance(items, list):
        return []
    return [r for r in items[:max_items] if isinstance(r, dict) and isinstance(r.get("id"), str)]


def home_widget_tasks(repo):
    items = valid_home_records(home_load(repo, "tasks", []))
    open_items = [x for x in items if not x.get("done")]
    st.caption(f"{len(open_items)} pendiente(s) · {len(items)-len(open_items)} completada(s)")
    with st.expander('➕ Nueva tarea',expanded=not items):
        with st.form("home_add_task", clear_on_submit=True):
            title = st.text_input("Nueva tarea", placeholder="Ej. Revisar el hormigonado", max_chars=180)
            c1, c2 = st.columns(2)
            with c1: due = st.date_input("Fecha límite", now_ec().date(), key="home_task_due")
            with c2: priority = st.selectbox("Prioridad", ["Normal", "Alta", "Baja"])
            add = st.form_submit_button("+ Añadir tarea", use_container_width=True)
    if add:
        try:
            if not title.strip():
                raise ValueError("Escribe el nombre de la tarea.")
            if len(items) >= 150:
                raise ValueError("Máximo 150 tareas. Elimina algunas antes de agregar más.")
            items.append({"id":uuid.uuid4().hex, "title":title.strip(), "date":due.isoformat(),
                          "priority":priority, "done":False})
            home_save(repo, "tasks", items);rerun_home()
        except Exception as e:notice_error(e)
    if not items:
        st.info("Tu lista está vacía. Agrega una tarea para empezar.")
        return
    show_completed = st.checkbox("Mostrar completadas", value=False, key="home_completed_tasks")
    ordered = sorted(items, key=lambda x:(bool(x.get("done")),
        0 if x.get("priority")=="Alta" else 1, str(x.get("date") or "")))
    displayed = [x for x in ordered if show_completed or not x.get("done")][:12]
    for item in displayed:
        title_text = str(item.get("title") or "Tarea")[:180]
        row_a, row_b = st.columns([8, 1], vertical_alignment="center")
        with row_a:
            check = st.checkbox(title_text, value=bool(item.get("done")), key="h_task_"+item["id"])
        with row_b:
            delete = st.button("✕", key="h_del_task_"+item["id"], help="Eliminar tarea", use_container_width=True)
        due_text = str(item.get("date") or "")
        st.caption(("Prioridad alta · " if item.get("priority")=="Alta" else "") + due_text)
        if delete:
            try:
                home_save(repo, "tasks", [x for x in items if x["id"]!=item["id"]]);rerun_home()
            except Exception as e:notice_error(e)
        if check != bool(item.get("done")):
            try:
                item["done"] = check;home_save(repo, "tasks", items);rerun_home()
            except Exception as e:notice_error(e)
    if len(open_items)>12 and not show_completed:
        st.caption("Mostrando las primeras 12 tareas. Prioriza y completa para ver las siguientes.")


def home_widget_notes(repo):
    notes = valid_home_records(home_load(repo, "notes", []), max_items=80)
    with st.expander('➕ Nueva nota',expanded=not notes):
        with st.form("home_add_note", clear_on_submit=True):
            title = st.text_input("Título", max_chars=120, placeholder="Ej. Observaciones de obra")
            body = st.text_area("Nota", height=95, max_chars=6000,
                                placeholder="Escribe tus apuntes aquí…")
            add = st.form_submit_button("+ Guardar nota", use_container_width=True)
    if add:
        try:
            if not title.strip() and not body.strip():
                raise ValueError("La nota necesita título o contenido.")
            if len(notes)>=80:
                raise ValueError("Máximo 80 notas guardadas.")
            notes.insert(0,{"id":uuid.uuid4().hex,"title":title.strip() or "Sin título",
                            "body":body.strip(),"updated":now_ec().strftime('%Y-%m-%d %H:%M')})
            home_save(repo,"notes",notes);rerun_home()
        except Exception as e:notice_error(e)
    if not notes:
        st.info("Tus notas estarán aquí, disponibles al volver a iniciar sesión.")
        return
    selected = st.selectbox("Abrir nota", notes,
        format_func=lambda n:str(n.get("title") or "Sin título")[:100], key="home_note_select")
    with st.expander('✏️ Editar nota seleccionada',expanded=False):
        with st.form("home_edit_note_"+selected["id"]):
            new_title=st.text_input("Editar título",value=str(selected.get("title") or "")[:120],max_chars=120)
            new_body=st.text_area("Editar nota",value=str(selected.get("body") or "")[:6000],height=140,max_chars=6000)
            saved=st.form_submit_button("Guardar cambios",use_container_width=True)
    if saved:
        try:
            selected["title"] = new_title.strip() or "Sin título"
            selected["body"] = new_body.strip()
            selected["updated"] = now_ec().strftime('%Y-%m-%d %H:%M')
            home_save(repo,"notes",notes);rerun_home()
        except Exception as e:notice_error(e)
    with st.popover("Eliminar esta nota",use_container_width=True):
        st.warning("La nota se eliminará permanentemente de tu espacio personal.")
        if st.button("Confirmar eliminación",key="h_confirm_note_del"):
            try:
                home_save(repo,"notes",[n for n in notes if n["id"]!=selected["id"]]);rerun_home()
            except Exception as e:notice_error(e)


def home_calendar_html(year, month, events):
    months=calendar.monthcalendar(year,month)
    days=("L","M","M","J","V","S","D")
    today=now_ec().date()
    counts=defaultdict(int)
    for e in events:
        when=str(e.get("date") or "")
        if when.startswith(f"{year:04d}-{month:02d}-"):
            try:counts[int(when[8:10])]+=1
            except ValueError:pass
    pieces=['<table class="ab-home-cal"><thead><tr>']
    pieces.extend('<th>'+d+'</th>' for d in days)
    pieces.append('</tr></thead><tbody>')
    for week in months:
        pieces.append('<tr>')
        for day in week:
            if not day:
                pieces.append('<td class="ab-cal-empty"></td>');continue
            active=(year,month,day)==(today.year,today.month,today.day)
            event_count=counts.get(day,0)
            pieces.append('<td'+(' class="ab-cal-today"' if active else '')+'><span>'+str(day)+'</span>'+(
                '<small title="Eventos">'+str(event_count)+'</small>' if event_count else '')+'</td>')
        pieces.append('</tr>')
    pieces.append('</tbody></table>')
    return ''.join(pieces)


def home_widget_calendar(repo):
    events=valid_home_records(home_load(repo,"events",[]),max_items=150)
    today=now_ec().date()
    a,b=st.columns(2)
    with a: month=st.selectbox("Mes",list(range(1,13)),index=today.month-1,
         format_func=lambda i:NOMBRES_MESES[i],key="home_calendar_month")
    with b: year=st.selectbox("Año",list(range(today.year-1,today.year+4)),
         index=1,key="home_calendar_year")
    st.markdown(home_calendar_html(year,month,events),unsafe_allow_html=True)
    with st.expander("+ Agendar un evento",expanded=False):
        with st.form("home_new_event",clear_on_submit=True):
            title=st.text_input("Evento",max_chars=150,placeholder="Ej. Reunión de coordinación")
            c1,c2=st.columns(2)
            with c1: day=st.date_input("Fecha",value=today,key="home_event_date")
            with c2: when=st.time_input("Hora",value=dt.time(8,0),key="home_event_time")
            add=st.form_submit_button("Guardar en calendario",use_container_width=True)
        if add:
            try:
                if not title.strip():raise ValueError("Escribe el nombre del evento.")
                if len(events)>=150:raise ValueError("Máximo 150 eventos.")
                events.append({"id":uuid.uuid4().hex,"title":title.strip(),
                               "date":day.isoformat(),"time":when.strftime('%H:%M')})
                home_save(repo,"events",events);rerun_home()
            except Exception as e:notice_error(e)
    monthly=sorted([x for x in events if str(x.get("date",'')).startswith(f"{year:04d}-{month:02d}-")],
                   key=lambda x:(str(x.get("date")),str(x.get("time"))))
    if not monthly:
        st.caption("Sin eventos agendados para este mes.")
    else:
        st.caption(f"{len(monthly)} evento(s) este mes")
        for event in monthly[:8]:
            c1,c2=st.columns([8,1],vertical_alignment="center")
            c1.markdown("**"+html_escape(str(event.get("title") or "Evento"))+"**")
            c1.caption(f"{event.get('date','')} · {event.get('time','')}")
            if c2.button("✕",key="h_event_del_"+event["id"],help="Eliminar evento"):
                try:
                    home_save(repo,"events",[e for e in events if e["id"]!=event["id"]]);rerun_home()
                except Exception as e:notice_error(e)
        if len(monthly)>8:st.caption(f"Mostrando 8 de {len(monthly)} eventos del mes.")


@st.cache_data(ttl=900,show_spinner=False,max_entries=20)
def home_weather_data(city):
    if city not in HOME_CITIES:
        raise ValueError("Ciudad no disponible.")
    latitude,longitude=HOME_CITIES[city]
    query=urllib.parse.urlencode({
        'latitude':latitude,'longitude':longitude,
        'current':'temperature_2m,relative_humidity_2m,apparent_temperature,weather_code,wind_speed_10m',
        'daily':'temperature_2m_max,temperature_2m_min,precipitation_probability_max',
        'timezone':'America/Guayaquil','forecast_days':1})
    request=urllib.request.Request('https://api.open-meteo.com/v1/forecast?'+query,
                                   headers={'User-Agent':'AlphaBuilders/1.0'})
    with urllib.request.urlopen(request,timeout=5) as response:
        data=json.load(response)
    if not isinstance(data,dict) or not isinstance(data.get('current'),dict):
        raise ValueError("La respuesta del servicio meteorológico es inválida.")
    return data


def home_weather_description(code):
    try:code=int(code)
    except (ValueError,TypeError):return "Condición sin datos"
    if code==0:return "Despejado ☀️"
    if code in (1,2,3):return "Parcialmente nublado ⛅"
    if code in (45,48):return "Neblina 🌫️"
    if code in (51,53,55,56,57):return "Llovizna 🌦️"
    if code in (61,63,65,66,67,80,81,82):return "Lluvia 🌧️"
    if code in (71,73,75,77,85,86):return "Nieve ❄️"
    if code in (95,96,99):return "Tormenta ⛈️"
    return "Condición variable"


def home_widget_weather(repo):
    """Clima automático al abrir Inicio; cacheado 15 minutos por ciudad.

    Cambiar de ciudad muestra inmediatamente el nuevo pronóstico. El guardado
    de la ciudad es individual; ninguna consulta de clima escribe en Supabase.
    """
    city = home_load(repo, 'city', 'Quito')
    if city not in HOME_CITIES:
        city = 'Quito'

    chosen = st.selectbox(
        'Ciudad', list(HOME_CITIES),
        index=list(HOME_CITIES).index(city),
        key='home_weather_city',
    )
    if chosen != city:
        # Una única escritura cuando se elige otra ciudad; no hace falta botón.
        try:
            home_save(repo, 'city', chosen)
        except Exception as exc:
            notice_error(exc)

    # Se ejecuta directamente al renderizar el widget. La caché evita
    # solicitudes repetidas en cada rerun del formulario o cambio de pestaña.
    try:
        with st.spinner('Consultando el clima…', show_time=False):
            data = home_weather_data(chosen)
        current = data['current']
        daily = data.get('daily') or {}
        temperature = current.get('temperature_2m')
        if temperature is None:
            raise ValueError('Sin temperatura disponible.')

        st.markdown(
            '<div class="ab-home-weather"><div class="ab-weather-big">'
            + html_escape(f'{float(temperature):.1f} °C') + '</div><div>'
            + html_escape(home_weather_description(current.get('weather_code')))
            + '</div></div>',
            unsafe_allow_html=True,
        )
        a, b = st.columns(2)
        a.metric('Sensación', str(current.get('apparent_temperature', '—')) + ' °C')
        b.metric('Humedad', str(current.get('relative_humidity_2m', '—')) + ' %')
        maxima = daily.get('temperature_2m_max')
        minima = daily.get('temperature_2m_min')
        if maxima and minima:
            st.caption(
                'Máx. ' + str(maxima[0]) + ' °C · Mín. '
                + str(minima[0]) + ' °C · Viento '
                + str(current.get('wind_speed_10m', '—')) + ' km/h'
            )
        st.caption('Fuente: Open-Meteo · datos almacenados en caché hasta 15 minutos')
    except Exception:
        st.warning('No se pudo cargar el clima en este momento. Inténtalo más tarde.')
        st.caption('El resto de herramientas continúa funcionando con normalidad.')


def home_safe_calculate(expression):
    """Evalúa aritmética básica con AST, sin eval(), funciones ni acceso a objetos."""
    expression=str(expression).strip().replace('×','*').replace('÷','/').replace('^','**')
    if len(expression)>120:raise ValueError('La operación es demasiado larga.')
    tree=ast.parse(expression,mode='eval')
    limit=[0]
    def walk(node,depth=0):
        limit[0]+=1
        if limit[0]>60 or depth>20:
            raise ValueError('La operación es demasiado compleja.')
        if isinstance(node,ast.Expression):return walk(node.body,depth+1)
        if isinstance(node,ast.Constant) and type(node.value) in (int,float):
            value=float(node.value)
        elif isinstance(node,ast.UnaryOp) and isinstance(node.op,(ast.UAdd,ast.USub)):
            x=walk(node.operand,depth+1)
            value=x if isinstance(node.op,ast.UAdd) else -x
        elif isinstance(node,ast.BinOp) and isinstance(node.op,(ast.Add,ast.Sub,ast.Mult,ast.Div,ast.Mod,ast.Pow)):
            a=walk(node.left,depth+1);b=walk(node.right,depth+1)
            if isinstance(node.op,ast.Pow):
                if abs(b)>12 or abs(a)>1e8:raise ValueError('Potencia fuera de límites.')
                value=a**b
            elif isinstance(node.op,ast.Add):value=a+b
            elif isinstance(node.op,ast.Sub):value=a-b
            elif isinstance(node.op,ast.Mult):value=a*b
            elif isinstance(node.op,ast.Div):value=a/b
            else:value=a%b
        else:
            raise ValueError('Usa solo números, +, −, ×, ÷, %, ^ y paréntesis.')
        if isinstance(value,complex) or not math.isfinite(value) or abs(value)>1e14:
            raise ValueError('El resultado está fuera de límites.')
        return value
    try:return walk(tree)
    except ZeroDivisionError as exc:raise ValueError('No es posible dividir entre cero.') from exc
    except SyntaxError as exc:raise ValueError('Revisa los signos de la operación.') from exc
    except OverflowError as exc:raise ValueError('La operación supera el límite numérico.') from exc


def home_widget_calculator(repo):
    with st.form('home_calc_form'):
        expression=st.text_input('Operación',placeholder='Ej. (12.5 × 8) / 2',max_chars=120,
                                 key='home_calc_expr')
        solve=st.form_submit_button('Calcular =',use_container_width=True)
    if solve:
        try:
            result=home_safe_calculate(expression)
            st.session_state['home_calc_result'] = (f'{result:,.8f}'.rstrip('0').rstrip('.') or '0') if abs(result)<1e14 else str(result)
        except Exception as e:st.session_state['home_calc_result']=str(e)
    if st.session_state.get('home_calc_result'):
        st.markdown('<div class="ab-home-answer">'+html_escape(st.session_state['home_calc_result'])+'</div>',
                    unsafe_allow_html=True)
    st.caption('Permite operaciones con paréntesis, potencias y porcentajes mediante división.')


def home_widget_construction(repo):
    mode=st.selectbox('Tipo de cálculo',['Volumen de hormigón','Área de trabajo','Productividad por hora'],
                      key='home_construction_type')
    if mode=='Volumen de hormigón':
        a,b=st.columns(2)
        with a:length=st.number_input('Largo (m)',min_value=0.,value=5.,step=.5,key='hc_len')
        with b:width=st.number_input('Ancho (m)',min_value=0.,value=4.,step=.5,key='hc_wid')
        depth=st.number_input('Espesor (m)',min_value=0.,value=.15,step=.01,format='%.2f',key='hc_depth')
        volume=length*width*depth
        st.metric('Volumen geométrico',f'{volume:,.3f} m³')
        st.caption(f'Con 5 % adicional como referencia: {volume*1.05:,.3f} m³. Verifica pérdidas reales en obra.')
    elif mode=='Área de trabajo':
        a,b=st.columns(2)
        with a:length=st.number_input('Largo (m)',min_value=0.,value=8.,key='hc_area_l')
        with b:height=st.number_input('Alto/ancho (m)',min_value=0.,value=3.,key='hc_area_h')
        st.metric('Área estimada',f'{length*height:,.2f} m²')
    else:
        a,b=st.columns(2)
        with a:amount=st.number_input('Cantidad ejecutada',min_value=0.,value=30.,key='hc_prod_amt')
        with b:hours=st.number_input('Horas-hombre',min_value=.1,value=16.,key='hc_prod_hh')
        st.metric('Productividad',f'{amount/hours:,.3f} unidades/HH')
        st.caption('Indica la producción por hora-hombre. Verifica la unidad del rubro.')


def home_widget_converter(repo):
    types={
        'Longitud':{'mm':0.001,'cm':0.01,'m':1.,'km':1000.},
        'Área':{'cm²':.0001,'m²':1.,'ha':10000.},
        'Volumen':{'L':.001,'m³':1.,'cm³':.000001},
        'Masa':{'g':.001,'kg':1.,'t':1000.},
    }
    kind=st.selectbox('Magnitud',list(types),key='home_conv_type')
    options=list(types[kind])
    value=st.number_input('Cantidad',value=1.,key='home_conv_value')
    c1,c2=st.columns(2)
    with c1:from_unit=st.selectbox('De',options,key='home_conv_from')
    with c2:to_unit=st.selectbox('A',options,index=min(1,len(options)-1),key='home_conv_to')
    output=value*types[kind][from_unit]/types[kind][to_unit]
    st.markdown('<div class="ab-home-answer">'+html_escape(f'{output:,.6g} {to_unit}')+'</div>',unsafe_allow_html=True)


def home_widget_bookmarks(repo):
    items=valid_home_records(home_load(repo,'bookmarks',[]),max_items=40)
    with st.expander('➕ Añadir enlace',expanded=not items):
        with st.form('home_new_bookmark',clear_on_submit=True):
            name=st.text_input('Nombre del enlace',max_chars=90,placeholder='Ej. Plataforma de planos')
            url=st.text_input('URL',max_chars=600,placeholder='https://…')
            add=st.form_submit_button('+ Guardar enlace',use_container_width=True)
    if add:
        try:
            uri=urllib.parse.urlsplit(url.strip())
            if uri.scheme not in ('http','https') or not uri.netloc or uri.username or uri.password:
                raise ValueError('Ingresa una URL válida que comience por https:// o http://')
            if not name.strip():raise ValueError('Agrega un nombre al enlace.')
            if len(items)>=40:raise ValueError('Máximo 40 enlaces personales.')
            items.append({'id':uuid.uuid4().hex,'name':name.strip(),'url':url.strip()})
            home_save(repo,'bookmarks',items);rerun_home()
        except Exception as e:notice_error(e)
    if not items:st.caption('Agrega enlaces frecuentes para abrirlos desde Inicio.')
    for item in items:
        c1,c2=st.columns([8,1],vertical_alignment='center')
        link=str(item.get('url') or '')
        if urllib.parse.urlsplit(link).scheme in ('http','https'):
            c1.link_button('↗ '+str(item.get('name') or 'Abrir enlace')[:90],link,use_container_width=True)
        if c2.button('✕',key='h_delete_link_'+item['id'],help='Eliminar enlace'):
            try:home_save(repo,'bookmarks',[x for x in items if x['id']!=item['id']]);rerun_home()
            except Exception as e:notice_error(e)


HOME_RENDERERS = {
    'calendar':home_widget_calendar,'tasks':home_widget_tasks,'notes':home_widget_notes,
    'weather':home_widget_weather,'calculator':home_widget_calculator,
    'construction':home_widget_construction,'converter':home_widget_converter,
    'bookmarks':home_widget_bookmarks,
}


def home_customize(repo, selected):
    with st.popover('⚙️ Personalizar widgets',use_container_width=False):
        st.markdown('**Tu inicio, tus herramientas**')
        st.caption('Selecciona cuáles aparecen. Puedes cambiar el orden con las flechas.')
        with st.form('home_customize_form'):
            chosen=st.multiselect('Widgets visibles',list(HOME_WIDGETS),default=selected,
                      format_func=lambda key:HOME_WIDGETS[key][0]+' '+HOME_WIDGETS[key][1],
                      key='home_widget_choices_'+hashlib.sha1(str(selected).encode()).hexdigest()[:8])
            apply=st.form_submit_button('Guardar selección',use_container_width=True)
        if apply:
            try:
                if not chosen:raise ValueError('Deja al menos una herramienta activa.')
                updated=[name for name in selected if name in chosen]+[name for name in chosen if name not in selected]
                home_save(repo,'layout',updated);rerun_home()
            except Exception as e:notice_error(e)
        if len(selected)>1:
            moving=st.selectbox('Ordenar herramientas',selected,
                          format_func=lambda name:HOME_WIDGETS[name][1],key='home_reorder_select')
            idx=selected.index(moving)
            up,down=st.columns(2)
            if up.button('↑ Subir',disabled=idx==0,key='home_move_up',use_container_width=True):
                try:
                    reordered=selected.copy();reordered[idx-1],reordered[idx]=reordered[idx],reordered[idx-1]
                    home_save(repo,'layout',reordered);rerun_home()
                except Exception as e:notice_error(e)
            if down.button('↓ Bajar',disabled=idx==len(selected)-1,key='home_move_down',use_container_width=True):
                try:
                    reordered=selected.copy();reordered[idx+1],reordered[idx]=reordered[idx],reordered[idx+1]
                    home_save(repo,'layout',reordered);rerun_home()
                except Exception as e:notice_error(e)


@(_FRAGMENT if _FRAGMENT is not None else (lambda fn:fn))
def render_home_widgets(repo):
    selected=home_active_widgets(repo)
    a,b=st.columns([5,1],vertical_alignment='center')
    with a:
        st.markdown('<div class="ab-home-intro"><div class="ab-home-kicker">TU ESPACIO PERSONAL</div>'
                    '<h2>Mi escritorio</h2><p>Organiza tu día, tus notas y tus herramientas. '
                    'Elige qué mostrar y cambia el orden cuando quieras.</p></div>',unsafe_allow_html=True)
    with b:home_customize(repo,selected)
    st.caption(f'{len(selected)} widget(s) activos · Preferencias privadas de tu cuenta · Guardado en Supabase')
    for offset in range(0,len(selected),2):
        columns=st.columns([1,1],gap='medium')
        for col,widget in zip(columns,selected[offset:offset+2]):
            with col:
                with st.container(border=True):
                    home_widget_title(widget)
                    HOME_RENDERERS[widget](repo)


def main():
    for prefix in st.session_state.pop('_cleanup_prefixes',[]):
        for key in list(st.session_state):
            if key.startswith(prefix):del st.session_state[key]
    st.set_page_config(page_title='Alpha Builders | Portal Ejecutivo',page_icon='🏗️',layout='wide',initial_sidebar_state='expanded')
    render_app_styles()
    try:
        actual_theme=st.context.theme.type
    except (AttributeError,RuntimeError,KeyError):
        actual_theme=None
    if actual_theme in ('dark','light'):
        tokens=(
            '--bg-app:#0f172a;--text-app:#f8fafc;--card-bg:#1e293b;'
            '--card-border:#334155;--header-bg:#020617;--subtext:#94a3b8;'
            '--hover-bg:#334155;--table-row-even:#1e293b;'
            '--ab-surface:#1e293b;--ab-surface-soft:#162237;--ab-ink:#f2f6ff;'
            '--ab-muted:#bbc7de;--ab-outline:#40516b;--ab-chip:#243b63;'
            '--ab-chip-text:#c6d8ff;--ab-primary:#80a7ff;--ab-strong:#f8fbff;'
            '--ab-outline-strong:#5b6f8c;--ab-surface-soft-2:#1b2d44;--ab-highlight:#9dbaff'
        ) if actual_theme=='dark' else (
            '--bg-app:#f8fafc;--text-app:#0f172a;--card-bg:#ffffff;'
            '--card-border:#cbd5e1;--header-bg:#0f172a;--subtext:#64748b;'
            '--hover-bg:#e2e8f0;--table-row-even:#f8fafc;'
            '--ab-surface:#ffffff;--ab-surface-soft:#f5f8ff;--ab-ink:#17233b;'
            '--ab-muted:#586781;--ab-outline:#dbe4f1;--ab-chip:#eaf1ff;'
            '--ab-chip-text:#234a94;--ab-primary:#2459dc;--ab-strong:#0b1830;'
            '--ab-outline-strong:#bbcbe3;--ab-surface-soft-2:#edf3fd;--ab-highlight:#3a6ce6'
        )
        theme_css = '<style>:root{'+tokens+'}</style>'
        if hasattr(st, "html"):
            st.html(theme_css)
        else:
            st.markdown(theme_css, unsafe_allow_html=True)
    try:
        if get_config('alpha_schema_version')!='5':
            st.error('El esquema depurado v5 no está activado. No se intentará usar el sistema anterior.')
            st.caption('Revise la migración y los secretos de Streamlit.')
            st.stop()
        cookie=stx.CookieManager(key='alpha_cookie_'+str(st.session_state.get('_cookie_epoch',0)))
        user=auth_user(cookie)
        if user is None:
            login_panel(cookie);st.stop()
        if st.session_state.get('_user_identity')!=user['correo']:
            for key in list(st.session_state):
                if key.startswith(('edit:','alpha_tabs_','alpha_rend_edificios','_v10_record_detail_','_v10_hist_active_','v11_')) or key in {'_editors','_export_artifact','_temporary_password','_route'}:
                    del st.session_state[key]
            st.session_state['_user_identity']=user['correo']
        st.session_state.pop('_pending_cookie',None)
        repo=Repository(user)
        if user.get('must_change_password'):
            st.warning('Debe cambiar su contraseña temporal antes de acceder al portal.')
            password_change(repo,cookie);st.stop()
        sidebar_original(repo,cookie)
        if st.session_state.get('_flash'):
            st.success(st.session_state.pop('_flash'))
        dashboard(repo)
        master=user.get('cargo')=='Maestro Mayor'
        visible=(['Libro de Obra Maestro','Requerimientos','Personal a Cargo','Control de Rendimiento','Colaborativo']
                 if master else ['Checklist','Libro de Obra','Personal a Cargo','Levantamiento de Incidencias','Control de Rendimiento','Colaborativo'])
        visible=['Inicio']+visible+['Catálogo de Rubros','Mi cuenta']
        if repo.admin:visible.append('Panel Admin')
        short={'Inicio':'🏠 Inicio','Checklist':'📋 Checklist','Libro de Obra':'🏗️ Libro de obra',
               'Libro de Obra Maestro':'🏗️ Libro maestro','Requerimientos':'📦 Requerimientos','Personal a Cargo':'👷 Personal',
               'Levantamiento de Incidencias':'🚨 Incidencias','Control de Rendimiento':'📊 Rendimientos',
               'Colaborativo':'🤝 Colaborativo','Catálogo de Rubros':'📚 Rubros',
               'Mi cuenta':'👤 Mi cuenta','Panel Admin':'⚙️ Admin'}
        labels=[short[v] for v in visible]
        route_key='v11_navigation_'+('master' if master else 'staff')+('_admin' if repo.admin else '')
        next_module=st.session_state.pop('_next_route',None)
        next_map={'Libro maestro':'Libro de Obra Maestro','Libro oficial':'Libro de Obra',
                  'Rendimientos':'Control de Rendimiento','Incidencias':'Levantamiento de Incidencias','Checklist':'Checklist'}
        next_label=next_map.get(next_module,next_module)
        if next_label in visible:
            st.session_state[route_key]=short[next_label]
        st.markdown('<div class="ab-module-tabs-title">SELECCIONA UN MÓDULO</div>',unsafe_allow_html=True)
        chosen=st.segmented_control('Navegación',labels,default=short['Inicio'],key=route_key,
                                     label_visibility='collapsed') or short['Inicio']
        active=visible[labels.index(chosen)]
        if active=='Inicio':
            render_home_widgets(repo)
            return
        render_active_module(active,repo,cookie)
    except Exception as exc:
        notice_error(exc)
        st.caption('No se interpretan errores de lectura como registros vacíos. Verifica conexión y permisos.')
        if st.button('Reintentar conexión'):
            invalidate();st.rerun()


if __name__=="__main__":
    main()
