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
APP_VERSION = "2026.10.10-v12.1-clima-auto"
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
.stApp { background:var(--bg-app)!important; }
.stApp [data-testid="stMainBlockContainer"] { max-width:1520px!important; margin:auto; }
.stApp [data-testid="stMainBlockContainer"] p,
.stApp [data-testid="stMainBlockContainer"] label,
.stApp [data-testid="stMainBlockContainer"] span { line-height:1.45; }
.stApp [data-testid="stMainBlockContainer"] label,
.stApp [data-testid="stMainBlockContainer"] .stMarkdown,
.stApp [data-testid="stMainBlockContainer"] p,
.stApp [data-testid="stMainBlockContainer"] span:not([data-baseweb="tag"]) { color:var(--ab-ink)!important; }
.stApp [data-testid="stMainBlockContainer"] h1,
.stApp [data-testid="stMainBlockContainer"] h2,
.stApp [data-testid="stMainBlockContainer"] h3,
.stApp [data-testid="stMainBlockContainer"] h4 { color:var(--text-app)!important; }

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
.stApp [data-testid="stSegmentedControl"] label:has(input:checked) {
 background:var(--ab-chip)!important;border-color:var(--ab-primary)!important;
}
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
    selected_period=st.segmented_control('Periodo del avance', ['🌤️ Mañana','☀️ Mediodía','🌇 Tarde'],
        default='🌤️ Mañana',key=editor.prefix+'period_view') or '🌤️ Mañana'
    period,label={'🌤️ Mañana':('manana','Mañana'),'☀️ Mediodía':('mediodia','Mediodía'),
                  '🌇 Tarde':('tarde','Tarde')}[selected_period]
    with st.container(border=True):
        editor.field('number_input','avance_'+period,'Avance de '+label.lower(),0.0,
                     min_value=0.0,step=0.01,format='%.2f')
        editor.photos(('fotos_'+period,),'Fotografías de '+label.lower())
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


def show_record(table,record):
    """Un solo bloque visible por ficha, evitando tablas/fotos interminables."""
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
    """Navegación en tres pasos: filtros → seleccionar ficha → única sección visible."""
    st.markdown('#### 📚 Consultar registros')
    if not records:
        st.info('Todavía no hay registros guardados en este módulo.');return
    counts=defaultdict(int)
    for rec in records:
        counts[record_building(table,rec)]+=1
    buildings=sorted(counts,key=lambda x:x.casefold())
    a,b=st.columns([1.25,1])
    with a:
        building=st.selectbox('🏢 Edificio',['Todos los edificios']+buildings,
           format_func=lambda x:x if x=='Todos los edificios' else f'{x} · {counts[x]}',key='hist_build_'+key)
    filtered=[r for r in records if building=='Todos los edificios' or record_building(table,r)==building]
    months=sorted({record_date(table,r)[:7] for r in filtered if len(record_date(table,r))>=7},reverse=True)
    with b:
        month=st.selectbox('🗓️ Mes',['Todos los meses']+months,
            format_func=lambda x:display_month(x) if x!='Todos los meses' else x,key='hist_month_'+key)
    if month!='Todos los meses':filtered=[r for r in filtered if record_date(table,r).startswith(month)]
    if table=='incidencias':
        status=st.segmented_control('Estado',['Todas','Abierta','Cerrada'],default='Todas',key='hist_status_'+key)
        if status and status!='Todas':filtered=[r for r in filtered if r.get('estado')==status]
    search=st.text_input('🔎 Buscar',placeholder='Actividad, responsable, rubro…',key='hist_search_'+key)
    if search.strip():
        needle=search.strip().casefold()
        filtered=[r for r in filtered if needle in ' '.join((record_building(table,r),record_title(table,r),record_note(table,r))).casefold()]
    st.caption(f'{len(filtered)} registros encontrados. Selecciona uno para ver su ficha.')
    with st.expander('📤 Exportar resultados filtrados',expanded=False):
        export_controls(table,filtered,'hist_export_'+key)
    if not filtered:
        st.info('No hay resultados con estos filtros.');return
    ordered=sorted(filtered,key=lambda r:(record_date(table,r),str(r.get('id') or '')),reverse=True)
    records_per_page=15
    visible=page_slice(ordered,'hist_block_'+key,records_per_page)
    options=[None]+visible
    selected=st.selectbox('Selecciona un registro',options,
        format_func=lambda r:'Elige una ficha para ver su contenido' if r is None else
          f"{display_date(record_date(table,r))} · {record_title(table,r)} · {record_building(table,r)}",
        key='v11_selected_record_'+key)
    if selected is None:
        st.info('Elige un registro del selector para visualizar sus detalles y acciones.');return
    with st.container(border=True):
        st.markdown('#### '+html_escape(record_title(table,selected)))
        st.caption(display_date(record_date(table,selected),True)+' · '+record_building(table,selected))
        show_record(table,selected)
        if editable:
            ca,cb=st.columns(2)
            identity=key+':'+str(selected['id'])
            if ca.button('✏️ Editar esta ficha',key='edit_'+identity,use_container_width=True):
                try:
                    repo._access(table,selected,write=True)
                    if table=='checklists':module='Checklist';factory=checklist_data
                    elif table=='rendimientos':module='Rendimientos';factory=rend_data
                    elif table=='incidencias':module='Incidencias';factory=incident_data
                    else:
                        master=is_master_book(selected)
                        module='Libro maestro' if master else 'Libro oficial'
                        factory=maestro_data if master else official_data
                    begin_editor(module,repo.user,factory(selected),selected)
                    st.session_state['_next_route']=module
                    st.rerun()
                except Exception as exc:notice_error(exc)
            with cb.popover('🗑️ Eliminar ficha',use_container_width=True):
                st.warning('Esta acción es irreversible.')
                if st.button('Confirmar eliminación',key='delete_'+identity):
                    try:repo.delete(table,selected);st.session_state['_flash']='Registro eliminado.';st.rerun()
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
    area=st.segmented_control('Sección',['📌 Resumen','👥 Ver nómina','➕ Registrar','📥 Importar','✏️ Gestionar'],
        default='📌 Resumen',key='v8_personal_section',label_visibility='collapsed') or '👥 Ver nómina'
    if area=='📌 Resumen':
        st.markdown('<div class="ab-welcome"><h3>Gestión de cuadrillas</h3>'
                    '<p>Consulta el personal existente o elige una acción sin cargar toda la lista.</p></div>',
                    unsafe_allow_html=True)
        a,b,c=st.columns(3)
        for col,caption,target in [(a,'👥 Consultar nómina','👥 Ver nómina'),
                                   (b,'➕ Registrar trabajador','➕ Registrar'),
                                   (c,'📥 Importar nómina','📥 Importar')]:
            if col.button(caption,key='v11_personal_'+target,use_container_width=True):
                st.session_state['v8_personal_section']=target;st.rerun()
        return
    if area=='👥 Ver nómina':
        if pending:
            st.info(f'Encontramos {len(pending)} trabajador(es) en el respaldo de la versión anterior que aún no aparecen en la tabla actual. Se muestran junto a los demás, sin modificar tu base de datos.')
            canonical={p.casefold():p for p in projects}
            eligible=[{**w,'edificio':canonical[w['edificio'].casefold()]} for w in pending if w['edificio'].casefold() in canonical and w['nombre'] and w['cargo']]
            excluded=len(pending)-len(eligible)
            with st.expander('🛟 Recuperar mi nómina anterior',expanded=False):
                st.markdown('La copia histórica permanece intacta. **La recuperación solo añade trabajadores que no existan ya en la tabla actual.**')
                st.download_button('⬇️ Descargar respaldo de nómina (CSV)',
                    data=pd.DataFrame([{'Nombre':w['nombre'],'Cargo':w['cargo'],'Edificio':w['edificio']} for w in pending]).to_csv(index=False,sep=';').encode('utf-8-sig'),
                    file_name='Respaldo_Nomina_Alpha_Builders.csv',mime='text/csv',key='v10_worker_backup_csv')
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
    st.header("Panel administrador")
    section=st.selectbox("Administrar",["Resumen","Registros","Usuarios y permisos","Solicitudes","PIN de acceso","Revisión de migración"],key="admin_section")
    if section=="Resumen":
        st.info("Selecciona un apartado del administrador para consultar únicamente esa información.")
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


def history_rend_by_building(records, repo):
    """Resumen con una tabla ligera; abrir una única ficha bajo demanda."""
    st.markdown('#### 📊 Rendimientos por edificio')
    if not records:
        st.info('No hay rendimientos registrados todavía.');return
    by_build=defaultdict(list)
    for rec in records:by_build[record_building('rendimientos',rec)].append(rec)
    buildings=sorted(by_build,key=lambda x:x.casefold())
    a,b=st.columns([1.5,1])
    with a:
        building=st.selectbox('🏢 Edificio',['Todos los edificios']+buildings,
            format_func=lambda x:x if x=='Todos los edificios' else f'{x} · {len(by_build[x])}',key='v8_rend_building')
    filtered=records if building=='Todos los edificios' else by_build[building]
    months=sorted({record_date('rendimientos',r)[:7] for r in filtered if len(record_date('rendimientos',r))>=7},reverse=True)
    with b:
        month=st.selectbox('🗓️ Mes',['Todos los meses']+months,
            format_func=lambda x:display_month(x) if x!='Todos los meses' else x,key='v8_rend_month')
    filtered=[r for r in filtered if month=='Todos los meses' or record_date('rendimientos',r).startswith(month)]
    search=st.text_input('🔎 Buscar rubro o trabajador',placeholder='Rubro, piso, trabajador…',key='v8_rend_search')
    if search.strip():
        needle=search.strip().casefold()
        filtered=[r for r in filtered if needle in ' '.join((record_title('rendimientos',r),record_building('rendimientos',r),str(r.get('trabajador') or ''))).casefold()]
    st.caption(f'{len(filtered)} registros · Puedes consultar uno sin desplegar los demás')
    with st.expander('📤 Exportar rendimientos',expanded=False):
        export_controls('rendimientos',filtered,'rend_general_v8')
        if building!='Todos los edificios':
            export_controls('rendimientos',by_build[building],'rend_build_v11_'+hashlib.md5(building.encode()).hexdigest()[:9])
    if not filtered:
        st.info('No encontramos registros para esos filtros.');return
    ordered=sorted(filtered,key=lambda r:(record_date('rendimientos',r),str(r.get('id') or '')),reverse=True)
    visible=page_slice(ordered,'rend_v8_block',12)
    rows=[]
    for record in visible:
        value=rendimiento_values(record)
        rows.append([display_date(record_date('rendimientos',record)),record_building('rendimientos',record),
            record.get('rubro') or 'Actividad',f"{value['avance']:.2f} {record.get('unidad') or ''}",
            f"{value['hh']:.2f} HH" if value['hh'] is not None else 'No verificado',str(record.get('estado') or 'Sin meta')])
    styled_detail_table(['Fecha','Edificio','Rubro','Avance','Horas-hombre','Estado'],rows)
    selected=st.selectbox('Abrir ficha de rendimiento',[None]+visible,
        format_func=lambda r:'Selecciona el registro que deseas consultar' if r is None else
            f"{display_date(record_date('rendimientos',r))} · {record_title('rendimientos',r)} · {record_building('rendimientos',r)}",
        key='v11_selected_yield')
    if selected is None:
        return
    data=rendimiento_values(selected)
    with st.container(border=True):
        st.markdown('#### '+html_escape(record_title('rendimientos',selected)))
        c1,c2,c3=st.columns(3)
        c1.metric('Avance',f"{data['avance']:.2f} {selected.get('unidad') or ''}")
        c2.metric('Horas-hombre',f"{data['hh']:.2f} HH" if data['hh'] is not None else 'Sin verificar')
        c3.metric('Productividad',f"{data['productividad']:.3f} {selected.get('unidad') or ''}/h" if data['productividad'] is not None else '—')
        show_record('rendimientos',selected)
        a,b=st.columns(2)
        if a.button('✏️ Editar rendimiento',key='v11_edit_yield_'+str(selected['id']),use_container_width=True):
            try:
                repo._access('rendimientos',selected,write=True)
                begin_editor('Rendimientos',repo.user,rend_data(selected),selected)
                st.session_state['_next_route']='Rendimientos';st.rerun()
            except Exception as exc:notice_error(exc)
        with b.popover('🗑️ Eliminar rendimiento',use_container_width=True):
            st.warning('Esta acción es irreversible.')
            if st.button('Confirmar eliminación',key='v11_delete_yield_'+str(selected['id'])):
                try:repo.delete('rendimientos',selected);st.session_state['_flash']='Registro eliminado.';st.rerun()
                except Exception as exc:notice_error(exc)


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


def module_original(label, repo):
    """Un módulo = una vista. Resumen, registro e historial nunca se dibujan juntos."""
    spec = {
        'Checklist':('checklists','Checklist',checklist_data,render_checklist,'Checklist de obra','Verificaciones y supervisión diaria','Nueva jornada'),
        'Libro de Obra':('inspecciones','Libro oficial',official_data,render_official,'Libro de obra oficial','Registro técnico de la jornada','Nuevo libro de obra'),
        'Libro de Obra Maestro':('inspecciones','Libro maestro',maestro_data,render_maestro,'Libro de obra maestro','Cuadrillas, actividades y avances ejecutados','Nuevo libro maestro'),
        'Levantamiento de Incidencias':('incidencias','Incidencias',incident_data,render_incident,'Incidencias de obra','Registro y seguimiento de novedades','Nueva incidencia'),
        'Control de Rendimiento':('rendimientos','Rendimientos',rend_data,render_rend,'Control de rendimientos','Avances y productividad del personal','Nuevo rendimiento'),
    }
    table,module,factory,renderer,title,subtitle,button=spec[label]
    page_heading('Control de obra',title,subtitle)
    selection_key='v11_mode_'+module.replace(' ','_')
    editor_active=bool(st.session_state.get('_editors',{}).get(module))
    options=['📌 Resumen','➕ Registrar','📚 Consultar']
    if editor_active:
        # La edición activa tiene prioridad sobre las consultas.
        if st.session_state.get(selection_key)!='➕ Registrar':
            st.session_state[selection_key]='➕ Registrar'
    choice=st.segmented_control('Vista de este módulo',options,default='📌 Resumen',key=selection_key,
                                 label_visibility='collapsed') or '📌 Resumen'
    if choice=='📌 Resumen':
        st.markdown('<div class="ab-welcome"><h3>¿Qué deseas hacer?</h3>'
                    '<p>Elige una acción. Los formularios y los registros solo se cargarán cuando los abras.</p></div>',
                    unsafe_allow_html=True)
        a,b=st.columns(2)
        if a.button('➕ '+button,key='v11_quick_new_'+module,use_container_width=True,type='primary'):
            st.session_state[selection_key]='➕ Registrar';st.rerun()
        if b.button('📚 Consultar historial',key='v11_quick_history_'+module,use_container_width=True):
            st.session_state[selection_key]='📚 Consultar';st.rerun()
        st.caption('Los registros existentes y los borradores no se modifican al cambiar de vista.')
        return
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
                    choices={f"Borrador {i+1} · {r['updated_at'][:16]}":r for i,r in enumerate(reversed(drafts))}
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
    st.markdown(APP_CSS,unsafe_allow_html=True)
    st.markdown(APP_CSS_V10,unsafe_allow_html=True)
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
        st.markdown('<style>:root{'+tokens+'}</style>',unsafe_allow_html=True)
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
        visible=(['Libro de Obra Maestro','Personal a Cargo','Control de Rendimiento','Colaborativo']
                 if master else ['Checklist','Libro de Obra','Personal a Cargo','Levantamiento de Incidencias','Control de Rendimiento','Colaborativo'])
        visible=['Inicio']+visible+['Mi cuenta']
        if repo.admin:visible.append('Panel Admin')
        short={'Inicio':'🏠 Inicio','Checklist':'📋 Checklist','Libro de Obra':'🏗️ Libro de obra',
               'Libro de Obra Maestro':'🏗️ Libro maestro','Personal a Cargo':'👷 Personal',
               'Levantamiento de Incidencias':'🚨 Incidencias','Control de Rendimiento':'📊 Rendimientos',
               'Colaborativo':'🤝 Colaborativo','Mi cuenta':'👤 Mi cuenta','Panel Admin':'⚙️ Admin'}
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
