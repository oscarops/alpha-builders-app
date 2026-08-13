import base64
import datetime
import io
import json
import os
import pandas as pd
from PIL import Image, ImageOps
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from supabase import create_client, Client
from streamlit_local_storage import LocalStorage

# Importaciones de ReportLab para generación de PDFs
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ==========================================
st.set_page_config(
    page_title="Alpha Builders | Portal Ejecutivo",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    h1, h2, h3, .brand-title { font-family: 'Montserrat', sans-serif !important; letter-spacing: -0.03em !important; }

    .block-container { padding-top: 1rem !important; padding-bottom: 1.5rem !important; padding-left: 2.5rem !important; padding-right: 2.5rem !important; max-width: 100% !important; }
    .stApp { background-color: #ffffff !important; color: #121318 !important; }
    .stApp p, .stApp label, .stApp span, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6 { color: #121318; }
    .stCaption, caption, small, [data-testid="stCaptionContainer"] { color: #5a5f6e !important; }

    [data-testid="stInputInstructions"], div[data-testid="stInputInstructions"] { display: none !important; visibility: hidden !important; }

    [data-testid="stSidebarCollapseButton"] { display: block !important; visibility: visible !important; opacity: 1 !important; z-index: 999999 !important; }
    [data-testid="collapsedControl"] { display: block !important; visibility: visible !important; opacity: 1 !important; position: fixed !important; top: 15px !important; left: 15px !important; z-index: 999999 !important; }

    [data-testid="stSidebarCollapseButton"] button, [data-testid="collapsedControl"] button {
        background-color: #1c1e26 !important; border: 1px solid #323646 !important; border-radius: 50% !important; width: 36px !important; height: 36px !important; color: #ffffff !important; box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important; transition: all 0.2s ease !important;
    }
    [data-testid="stSidebarCollapseButton"] button:hover, [data-testid="collapsedControl"] button:hover {
        background-color: #ff8c00 !important; border-color: #ff8c00 !important; transform: scale(1.08);
    }
    [data-testid="stSidebarCollapseButton"] svg, [data-testid="collapsedControl"] svg { fill: #ffffff !important; color: #ffffff !important; }

    [data-testid="stSidebar"] { background-color: #121318 !important; border-right: 2px solid #282a36 !important; padding-top: 0px !important; padding-left: 12px !important; padding-right: 12px !important; padding-bottom: 15px !important; }
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0.5rem !important; padding-top: 0px !important; }
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] div { color: #ffffff !important; }

    .sidebar-logo-card { background-color: #ffffff; border-radius: 12px; padding: 8px 10px; margin-top: 0px !important; margin-bottom: 20px !important; box-shadow: 0 4px 12px rgba(0,0,0,0.3); width: 100% !important; box-sizing: border-box; text-align: center; display: block; }
    [data-testid="stSidebar"] [data-testid="stImage"] { width: 100% !important; display: block !important; margin-top: 6px !important; margin-bottom: 10px !important; clear: both !important; }
    [data-testid="stSidebar"] [data-testid="stImage"] img { border-radius: 12px !important; width: 100% !important; height: auto !important; max-width: 100% !important; object-fit: cover !important; border: 1px solid #323646 !important; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4); margin: 0 !important; display: block !important; }

    .sidebar-profile-box { background: #1c1e26; border: 1px solid #323646; border-radius: 12px; padding: 10px 8px !important; text-align: center; margin-top: 4px; margin-bottom: 8px; width: 100% !important; box-shadow: 0 4px 10px rgba(0,0,0,0.3); box-sizing: border-box; }
    .sidebar-user-nombres { font-size: 0.88rem; font-weight: 800; color: #ffffff !important; line-height: 1.2; }
    .sidebar-user-apellidos { font-size: 0.85rem; font-weight: 700; color: #e0e4ed !important; margin-bottom: 4px !important; line-height: 1.2; }
    .sidebar-user-email { font-size: 0.68rem; color: #72b2ff !important; font-weight: 600; margin-bottom: 6px !important; word-break: break-all; }
    .sidebar-user-cargo { display: inline-block; background: #323646 !important; color: #ffffff !important; border: 1px solid #484e5e !important; font-size: 0.60rem !important; font-weight: 800 !important; padding: 2px 8px !important; border-radius: 14px !important; text-transform: uppercase !important; }

    [data-testid="stSidebar"] hr { margin: 6px 0 !important; border-color: #282a36 !important; }
    [data-testid="stSidebar"] [data-testid="stExpander"] { background-color: #1c1e26 !important; border: 1px solid #323646 !important; border-radius: 10px !important; margin-top: 2px !important; margin-bottom: 6px !important; }
    [data-testid="stSidebar"] [data-testid="stExpander"] summary { background-color: #282c36 !important; padding: 6px 8px !important; }
    [data-testid="stSidebar"] [data-testid="stExpander"] summary * { color: #ffffff !important; font-weight: 700 !important; font-size: 0.78rem !important; }

    .executive-card-studio { background: linear-gradient(145deg, #f3f6fc 0%, #e8edf7 100%); border: 1px solid #b8c4d8; border-left: 7px solid #121318; border-radius: 22px; padding: 22px 28px; box-shadow: 0 12px 35px rgba(0,0,0,0.06); margin-bottom: 20px; width: 100%; box-sizing: border-box; }
    .brand-title { font-family: 'Montserrat', sans-serif !important; font-weight: 700 !important; font-size: 2.4rem !important; background: linear-gradient(90deg, #121318 0%, #3a4256 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; text-shadow: 0 2px 12px rgba(0,0,0,0.08); letter-spacing: -0.03em !important; }

    .kpi-card-studio { background: linear-gradient(145deg, #eceff6 0%, #dbe2ef 100%); border: 1px solid #aebacf; border-radius: 20px; padding: 18px; text-align: center; box-shadow: 0 8px 25px rgba(0,0,0,0.06); transition: all 0.3s ease; }
    .kpi-card-studio:hover { transform: translateY(-3px); box-shadow: 0 14px 35px rgba(0,0,0,0.12); filter: brightness(1.02); }
    .kpi-val-studio { font-size: 2.5rem; font-weight: 900; color: #121318 !important; }
    .kpi-lbl-studio { font-size: 0.72rem; color: #4a5060 !important; text-transform: uppercase; font-weight: 800; }

    .stTabs [data-baseweb="tab-list"] { gap: 10px; background-color: #e2e5ec !important; padding: 6px; border-radius: 16px; border: 1px solid #c2c7d2; }
    .stTabs [data-baseweb="tab"] { border-radius: 12px !important; padding: 10px 24px !important; background-color: transparent !important; }
    .stTabs [data-baseweb="tab"] p, .stTabs [data-baseweb="tab"] span { color: #121318 !important; font-weight: 700 !important; }
    .stTabs [aria-selected="true"] { background-color: #121318 !important; border-radius: 12px !important; }
    .stTabs [aria-selected="true"] p, .stTabs [aria-selected="true"] span, .stTabs [aria-selected="true"] div { color: #ffffff !important; font-weight: 900 !important; }

    .stButton > button { background-color: #121318 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; font-weight: 800 !important; padding: 10px 22px !important; }
    .stButton > button p, .stButton > button span { color: #ffffff !important; }

    .streamlit-expanderHeader { background-color: #e8eaee !important; border-radius: 12px !important; border: 1px solid #c2c7d2 !important; font-weight: 700 !important; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. CONEXIÓN Y CARGA SUPABASE
# ==========================================
@st.cache_resource
def init_supabase():
    url = st.secrets.get("SUPABASE_URL", "")
    key = st.secrets.get("SUPABASE_KEY", "")
    if not url or not key:
        st.error("⚠️ Credenciales SUPABASE_URL / SUPABASE_KEY no configuradas.")
        st.stop()
    return create_client(url, key)

supabase = init_supabase()
local_storage = LocalStorage()

DEFAULT_TRABAJADORES = []

def load_db_from_supabase():
    try:
        res_pin = supabase.table("app_config").select("*").eq("key", "access_pin").execute()
        access_pin = res_pin.data[0]["value"] if res_pin.data else "1254"
    except Exception:
        access_pin = "1254"

    try:
        res_usr = supabase.table("usuarios").select("*").execute()
        db_usuarios = []
        db_fotos = {}
        admin_emails = []
        for row in res_usr.data:
            c = row["correo"].lower().strip()
            db_usuarios.append({
                "Nombres": row["nombres"],
                "Apellidos": row["apellidos"],
                "Correo": c,
                "Password": row["password"],
                "Cargo": row["cargo"],
                "Fecha_Registro": str(row["fecha_registro"]),
                "Estado": row.get("estado", "Activo")
            })
            if row.get("foto_b64"):
                db_fotos[c] = row["foto_b64"]
            if row.get("es_admin"):
                admin_emails.append(c)
    except Exception:
        db_usuarios = []
        db_fotos = {}
        admin_emails = ["oscarsebitas2013@gmail.com"]

    if "oscarsebitas2013@gmail.com" not in admin_emails:
        admin_emails.append("oscarsebitas2013@gmail.com")

    try:
        res_trab = supabase.table("trabajadores").select("*").execute()
        db_trabajadores = [{"nombre": r["nombre"], "cargo": r["cargo"]} for r in res_trab.data] if res_trab.data else DEFAULT_TRABAJADORES
    except Exception:
        db_trabajadores = DEFAULT_TRABAJADORES

    # Carga Checklists
    db_checklists = {}
    try:
        res_chk = supabase.table("checklists").select("*").execute()
        for r in res_chk.data:
            c = r["usuario_email"].lower().strip()
            if c not in db_checklists:
                db_checklists[c] = []
            
            datos_parsed = r["datos"] if isinstance(r["datos"], list) else json.loads(r["datos"])
            db_checklists[c].append({
                "db_id": r["id"],
                "Fecha": str(r["fecha"]),
                "Hora_Inicio": r.get("hora_inicio", "07:00"),
                "Hora_Fin": r.get("hora_fin", "17:00"),
                "Edificio": r["edificio"],
                "Responsable": r.get("responsable", ""),
                "Cargo": r.get("cargo", ""),
                "Observacion_General": r.get("observacion_general", ""),
                "Datos": datos_parsed
            })
    except Exception:
        pass

    # Carga Formato de Inspección
    db_inspecciones = {}
    try:
        res_insp = supabase.table("inspecciones").select("*").execute()
        for r in res_insp.data:
            c = r["usuario_email"].lower().strip()
            if c not in db_inspecciones:
                db_inspecciones[c] = []
            
            datos_parsed = r["datos"] if isinstance(r["datos"], dict) else json.loads(r["datos"])
            db_inspecciones[c].append({
                "db_id": r["id"],
                "Fecha": str(r["fecha"]),
                "Dia": r.get("dia", ""),
                "Proyecto": r["proyecto"],
                "Residente": r.get("residente", ""),
                "Frente": r.get("frente", ""),
                "Clima": r.get("clima", ""),
                "Hora_Inicio": r.get("hora_inicio", "07:00"),
                "Hora_Fin": r.get("hora_fin", "17:00"),
                "Datos": datos_parsed
            })
    except Exception:
        pass

    # Carga Rendimientos
    db_rendimientos = {}
    try:
        res_rnd = supabase.table("rendimientos").select("*").execute()
        for r in res_rnd.data:
            c = r["usuario_email"].lower().strip()
            if c not in db_rendimientos:
                db_rendimientos[c] = []
            db_rendimientos[c].append({
                "db_id": r["id"],
                "Usuario_Registro": c,
                "Cargo_Registrador": r.get("cargo_obrero", ""),
                "Fecha": str(r["fecha"]),
                "Trabajador": r["trabajador"],
                "Cargo_Obrero": r.get("cargo_obrero", ""),
                "Rubro": r["rubro"],
                "Horas Trabajadas (HH)": float(r["horas_hh"]),
                "Avance": float(r["avance"]),
                "Unidad": r["unidad"],
                "Rend. Real (HH/Unid)": float(r["rend_real"]),
                "Rend. Teórico": float(r["rend_teorico"]),
                "Estado": r["estado"]
            })
    except Exception:
        pass

    return {
        "access_pin": access_pin,
        "admin_emails": admin_emails,
        "db_fotos_perfil_b64": db_fotos,
        "db_usuarios": db_usuarios,
        "db_checklists": db_checklists,
        "db_inspecciones": db_inspecciones,
        "db_rendimientos": db_rendimientos,
        "db_trabajadores": db_trabajadores,
    }

if "db_loaded" not in st.session_state or not st.session_state.db_loaded:
    p_data = load_db_from_supabase()
    st.session_state.access_pin = p_data["access_pin"]
    st.session_state.admin_emails = p_data["admin_emails"]
    st.session_state.db_fotos_perfil_b64 = p_data["db_fotos_perfil_b64"]
    st.session_state.db_usuarios = p_data["db_usuarios"]
    st.session_state.db_checklists = p_data["db_checklists"]
    st.session_state.db_inspecciones = p_data["db_inspecciones"]
    st.session_state.db_rendimientos = p_data["db_rendimientos"]
    st.session_state.db_trabajadores = p_data["db_trabajadores"]
    st.session_state.db_loaded = True

# PERSISTENCIA DE SESIÓN
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_email = ""
    st.session_state.usuario_nombres = ""
    st.session_state.usuario_apellidos = ""
    st.session_state.usuario_cargo = ""

if not st.session_state.autenticado:
    saved_token = local_storage.getItem("user_session_email")
    if saved_token:
        mail_clean = saved_token.strip().lower()
        u_match = next((u for u in st.session_state.db_usuarios if u["Correo"] == mail_clean), None)
        if u_match:
            st.session_state.autenticado = True
            st.session_state.usuario_email = mail_clean
            st.session_state.usuario_nombres = u_match["Nombres"]
            st.session_state.usuario_apellidos = u_match["Apellidos"]
            st.session_state.usuario_cargo = u_match["Cargo"]

def render_estado_badge(estado_str):
    if not estado_str:
        return '<span style="color: #64748b; font-weight: 600;">Sin Responder</span>'
    if "Cumple" in estado_str or estado_str in ["Sí", "Operativo", "Completado"]:
        return f'<span style="background-color: #dcfce7; color: #16a34a; font-weight: 800; padding: 3px 10px; border-radius: 8px; border: 1px solid #bbf7d0; font-size: 0.82rem;">{estado_str}</span>'
    elif "No" in estado_str or estado_str in ["Fuera de servicio", "Retrasado"]:
        return f'<span style="background-color: #fee2e2; color: #dc2626; font-weight: 800; padding: 3px 10px; border-radius: 8px; border: 1px solid #fca5a5; font-size: 0.82rem;">{estado_str}</span>'
    else:
        return f'<span style="background-color: #f1f5f9; color: #121318; font-weight: 800; padding: 3px 10px; border-radius: 8px; border: 1px solid #cbd5e1; font-size: 0.82rem;">{estado_str}</span>'

def image_to_base64(image_file):
    if image_file is not None:
        try:
            img = Image.open(image_file)
            img = ImageOps.exif_transpose(img)
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            return base64.b64encode(buffered.getvalue()).decode("utf-8")
        except Exception:
            return None
    return None

def base64_to_image(b64_str):
    if b64_str:
        try:
            img_data = base64.b64decode(b64_str)
            img = Image.open(io.BytesIO(img_data))
            img = ImageOps.exif_transpose(img)
            return img
        except Exception:
            return None
    return None

def get_repo_image_b64(filenames):
    for filename in filenames:
        if os.path.exists(filename):
            try:
                with open(filename, "rb") as f:
                    return base64.b64encode(f.read()).decode("utf-8")
            except Exception:
                pass
    return None

# ==========================================
# GENERADORES DE REPORTES EXCEL, PDF Y CSV
# ==========================================
def export_dataframe_to_excel_csv(df):
    """Exportación en CSV delimitado por punto y coma (para compatibilidad con Excel)."""
    df_clean = df.drop(columns=["Foto_B64", "db_id"], errors="ignore")
    return df_clean.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")

def export_checklist_to_excel_file(jornada_dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Checklist Obra"

    ws.merge_cells("A1:F1")
    ws["A1"] = f"INSPECCIÓN DE OBRA - {jornada_dict.get('Edificio', '')} ({jornada_dict.get('Fecha', '')})"
    ws["A1"].font = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    ws["A1"].fill = PatternFill(start_color="121318", end_color="121318", fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["A2"] = f"Hora Inicio: {jornada_dict.get('Hora_Inicio', 'N/A')}"
    ws["B2"] = f"Hora Fin: {jornada_dict.get('Hora_Fin', 'N/A')}"
    ws["C2"] = f"Responsable: {jornada_dict.get('Responsable', '')}"

    headers = ["Jornada", "N°", "Actividad", "Estado", "Observaciones", "Evidencia Fotográfica"]
    ws.append([])
    ws.append(headers)

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=4, column=col)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill(start_color="121318", end_color="121318", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    ws.row_dimensions[4].height = 25
    datos = jornada_dict.get("Datos", [])
    start_row = 5

    for row_idx, item in enumerate(datos, start=start_row):
        obs_val = item.get("Observaciones", "")
        if item.get("Actividades_Especificas"):
            sub_acts = " | Actividades: " + ", ".join([f"• {a['Actividad']}" for a in item["Actividades_Especificas"] if a.get("Actividad")])
            obs_val += sub_acts

        ws.append([item.get("Jornada", ""), item.get("N°", ""), item.get("Actividad", ""), item.get("Estado", ""), obs_val])
        ws.row_dimensions[row_idx].height = 140

        for c_i in range(1, 7):
            cell_txt = ws.cell(row=row_idx, column=c_i)
            cell_txt.border = thin_border
            cell_txt.font = Font(name="Arial", size=9)
            if c_i < 6:
                cell_txt.alignment = Alignment(vertical="center", wrap_text=True)

        foto_b64 = item.get("Foto_B64")
        if foto_b64:
            try:
                img_data = base64.b64decode(foto_b64)
                img_pil = Image.open(io.BytesIO(img_data))
                img_pil = ImageOps.exif_transpose(img_pil)
                img_pil = img_pil.resize((600, 450), Image.Resampling.LANCZOS)
                img_stream = io.BytesIO()
                img_pil.save(img_stream, format="PNG", quality=100)
                img_stream.seek(0)

                img_xlsx = OpenpyxlImage(img_stream)
                img_xlsx.width = 280
                img_xlsx.height = 140
                ws.add_image(img_xlsx, f"F{row_idx}")
            except Exception:
                pass

    if jornada_dict.get("Observacion_General"):
        last_r = len(datos) + start_row
        ws.cell(row=last_r, column=1, value="OBSERVACIÓN GENERAL:").font = Font(name="Arial", bold=True, size=10)
        ws.cell(row=last_r, column=3, value=jornada_dict.get("Observacion_General"))

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 35
    ws.column_dimensions['F'].width = 45

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def export_checklist_to_pdf_file(jornada_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []

    title_style = ParagraphStyle('TitleStyle', fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#121318'), alignment=1, spaceAfter=8)
    sub_style = ParagraphStyle('SubStyle', fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#333333'), spaceAfter=10)
    header_style = ParagraphStyle('HeaderStyle', fontName='Helvetica-Bold', fontSize=8.5, textColor=colors.white, alignment=1)
    cell_style = ParagraphStyle('CellStyle', fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#121318'))

    story.append(Paragraph(f"CHECKLIST DIARIO DE OBRA — {jornada_dict.get('Edificio', '').upper()}", title_style))
    story.append(Paragraph(f"<b>Fecha:</b> {jornada_dict.get('Fecha', '')} | <b>Horario:</b> {jornada_dict.get('Hora_Inicio', '')} - {jornada_dict.get('Hora_Fin', '')} | <b>Responsable:</b> {jornada_dict.get('Responsable', '')}", sub_style))

    data = [[
        Paragraph("<b>Jornada</b>", header_style),
        Paragraph("<b>N°</b>", header_style),
        Paragraph("<b>Actividad</b>", header_style),
        Paragraph("<b>Estado</b>", header_style),
        Paragraph("<b>Observaciones</b>", header_style)
    ]]
    
    for item in jornada_dict.get("Datos", []):
        obs = item.get("Observaciones", "")
        if item.get("Actividades_Especificas"):
            sub_acts = ", ".join([a['Actividad'] for a in item["Actividades_Especificas"] if a.get("Actividad")])
            if sub_acts:
                obs += f" (Específicas: {sub_acts})"

        data.append([
            Paragraph(str(item.get("Jornada", "")), cell_style),
            Paragraph(str(item.get("N°", "")), cell_style),
            Paragraph(str(item.get("Actividad", "")), cell_style),
            Paragraph(str(item.get("Estado", "") or "N/A"), cell_style),
            Paragraph(str(obs), cell_style)
        ])

    table = Table(data, colWidths=[55, 25, 190, 65, 215])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#121318')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(table)

    if jornada_dict.get("Observacion_General"):
        story.append(Spacer(1, 10))
        story.append(Paragraph(f"<b>Observación General:</b> {jornada_dict.get('Observacion_General')}", sub_style))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def export_inspeccion_to_excel_file(insp_dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inspección Diaria"

    # Estilos visuales
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), 
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), 
        bottom=Side(style='thin', color='CBD5E1')
    )
    fill_header = PatternFill(start_color="121318", end_color="121318", fill_type="solid")
    fill_section = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    
    font_main_title = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    font_section_title = Font(name="Arial", bold=True, color="1E293B", size=10)
    font_tbl_header = Font(name="Arial", bold=True, color="FFFFFF", size=9.5)
    font_bold = Font(name="Arial", bold=True, color="121318", size=9.5)
    font_regular = Font(name="Arial", size=9)

    # 1. Título Principal
    ws.merge_cells("A1:C1")
    ws["A1"] = f"FORMATO DE INSPECCIÓN DIARIA - {insp_dict.get('Proyecto', '').upper()}"
    ws["A1"].font = font_main_title
    ws["A1"].fill = fill_header
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Meta Información
    meta_info = [
        ["Fecha:", f"{insp_dict.get('Fecha', '')} ({insp_dict.get('Dia', '')})", ""],
        ["Residente de Obra:", insp_dict.get("Residente", ""), ""],
        ["Frente Inspeccionado:", insp_dict.get("Frente", ""), ""],
        ["Clima:", insp_dict.get("Clima", ""), f"Horario: {insp_dict.get('Hora_Inicio', '')} - {insp_dict.get('Hora_Fin', '')}"]
    ]
    for r in meta_info:
        ws.append(r)
        r_idx = ws.max_row
        ws.cell(row=r_idx, column=1).font = font_bold
        ws.cell(row=r_idx, column=2).font = font_regular
        ws.cell(row=r_idx, column=3).font = font_regular

    ws.append([])
    datos = insp_dict.get("Datos", {})

    # 2. Avance General
    r_av_head = ws.max_row + 1
    ws.merge_cells(f"A{r_av_head}:C{r_av_head}")
    ws[f"A{r_av_head}"] = "1. AVANCE GENERAL"
    ws[f"A{r_av_head}"].font = font_section_title
    ws[f"A{r_av_head}"].fill = fill_section
    ws[f"A{r_av_head}"].alignment = Alignment(vertical="center", indent=1)

    ws.append(["Actividad", "% Prog / % Ejec", "Estado"])
    r_hdr1 = ws.max_row
    for col_i in range(1, 4):
        c = ws.cell(row=r_hdr1, column=col_i)
        c.font = font_tbl_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")

    for av in datos.get("Avance", []):
        prog_ejec_str = f"Prog: {av.get('% Prog', 0)}% | Ejec: {av.get('% Ejec', 0)}%"
        ws.append([av.get("Actividad", ""), prog_ejec_str, av.get("Estado", "") or "N/A"])
        r_i = ws.max_row
        for col_i in range(1, 4):
            cell = ws.cell(row=r_i, column=col_i)
            cell.font = font_regular
            cell.border = thin_border
            if col_i in [2, 3]:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.append([])

    # 3. Check List General Agrupado por Secciones (Formato estructurado por módulos)
    r_chk_main = ws.max_row + 1
    ws.merge_cells(f"A{r_chk_main}:C{r_chk_main}")
    ws[f"A{r_chk_main}"] = "2. CHECK LIST Y CONTROL GENERAL DE OBRA"
    ws[f"A{r_chk_main}"].font = font_section_title
    ws[f"A{r_chk_main}"].fill = fill_section
    ws[f"A{r_chk_main}"].alignment = Alignment(vertical="center", indent=1)

    checklist_groups = datos.get("Checklist", {})

    for sec_name, items in checklist_groups.items():
        if items:
            ws.append([])
            r_sec = ws.max_row
            ws.merge_cells(f"A{r_sec}:C{r_sec}")
            ws[f"A{r_sec}"] = f"■ {sec_name.upper()}"
            ws[f"A{r_sec}"].font = font_bold
            
            # Cabecera de la sub-tabla
            ws.append(["Ítem / Aspecto Inspeccionado", "Cumple / Estado", "Observación"])
            r_hdr_sub = ws.max_row
            ws.row_dimensions[r_hdr_sub].height = 22
            
            for col_i in range(1, 4):
                c = ws.cell(row=r_hdr_sub, column=col_i)
                c.font = font_tbl_header
                c.fill = fill_header
                c.alignment = Alignment(horizontal="center", vertical="center")

            # Filas de la sección
            for it in items:
                it_nombre = it.get("Item") or it.get("Aspecto") or it.get("Revisar") or it.get("Equipo", "")
                it_estado = it.get("Cumple") or it.get("Estado", "") or "N/A"
                it_obs = it.get("Observación", "")

                ws.append([it_nombre, it_estado, it_obs])
                r_i = ws.max_row
                
                c1 = ws.cell(row=r_i, column=1)
                c2 = ws.cell(row=r_i, column=2)
                c3 = ws.cell(row=r_i, column=3)

                c1.font = font_regular
                c2.font = font_regular
                c3.font = font_regular

                c1.border = thin_border
                c2.border = thin_border
                c3.border = thin_border

                c1.alignment = Alignment(vertical="center", wrap_text=True)
                c2.alignment = Alignment(horizontal="center", vertical="center")
                c3.alignment = Alignment(vertical="center", wrap_text=True)

    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 45

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def export_inspeccion_to_pdf_file(insp_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    story = []

    title_style = ParagraphStyle('TitleStyle', fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#121318'), alignment=1, spaceAfter=6)
    sub_title = ParagraphStyle('SubTitle', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#121318'), spaceBefore=8, spaceAfter=4)
    sec_header_style = ParagraphStyle('SecHeaderStyle', fontName='Helvetica-Bold', fontSize=8.5, textColor=colors.HexColor('#1e293b'), spaceBefore=6, spaceAfter=2)
    header_style = ParagraphStyle('HeaderStyle', fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)
    cell_style = ParagraphStyle('CellStyle', fontName='Helvetica', fontSize=7.5, textColor=colors.HexColor('#121318'))

    story.append(Paragraph(f"FORMATO DE INSPECCIÓN DIARIA DE OBRA", title_style))
    
    meta_text = f"<b>Proyecto:</b> {insp_dict.get('Proyecto', '')} | <b>Fecha:</b> {insp_dict.get('Fecha', '')} ({insp_dict.get('Dia', '')})<br/>" \
                f"<b>Residente:</b> {insp_dict.get('Residente', '')} | <b>Frente:</b> {insp_dict.get('Frente', '')}<br/>" \
                f"<b>Clima:</b> {insp_dict.get('Clima', '')} | <b>Horario:</b> {insp_dict.get('Hora_Inicio', '')} - {insp_dict.get('Hora_Fin', '')}"
    
    story.append(Paragraph(meta_text, cell_style))
    story.append(Spacer(1, 6))

    datos = insp_dict.get("Datos", {})

    story.append(Paragraph("1. AVANCE GENERAL", sub_title))
    av_data = [[
        Paragraph("<b>Actividad</b>", header_style),
        Paragraph("<b>% Prog</b>", header_style),
        Paragraph("<b>% Ejec</b>", header_style),
        Paragraph("<b>Estado</b>", header_style)
    ]]
    for av in datos.get("Avance", []):
        av_data.append([
            Paragraph(av.get("Actividad", ""), cell_style),
            Paragraph(str(av.get("% Prog", 0)), cell_style),
            Paragraph(str(av.get("% Ejec", 0)), cell_style),
            Paragraph(str(av.get("Estado", "") or "N/A"), cell_style)
        ])
    t_av = Table(av_data, colWidths=[190, 75, 75, 210])
    t_av.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#121318')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_av)

    story.append(Spacer(1, 6))
    story.append(Paragraph("2. CHECK LIST Y CONTROL GENERAL DE OBRA", sub_title))

    checklist_groups = datos.get("Checklist", {})

    for sec_name, items in checklist_groups.items():
        if items:
            story.append(Paragraph(f"📌 <b>{sec_name.upper()}</b>", sec_header_style))
            chk_data = [[
                Paragraph("<b>Ítem / Aspecto Inspeccionado</b>", header_style),
                Paragraph("<b>Cumple / Estado</b>", header_style),
                Paragraph("<b>Observación</b>", header_style)
            ]]

            for it in items:
                it_nombre = it.get("Item") or it.get("Aspecto") or it.get("Revisar") or it.get("Equipo", "")
                it_estado = it.get("Cumple") or it.get("Estado", "") or "N/A"
                it_obs = it.get("Observación", "")

                chk_data.append([
                    Paragraph(str(it_nombre), cell_style),
                    Paragraph(str(it_estado), cell_style),
                    Paragraph(str(it_obs), cell_style)
                ])

            t_sec = Table(chk_data, colWidths=[240, 100, 210])
            t_sec.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#121318')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ]))
            story.append(t_sec)
            story.append(Spacer(1, 4))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# 3. CONSTANTES
# ==========================================
EDIFICIOS_ALPHA = [
    "Tesla", "Lafuente", "Imagine", "Asimov", "Rubik", "Castle Rock",
    "Musk", "Wolf", "Dablanc", "Thomas Edison", "Westinghouse", "Smart",
]

UNIDADES_RUBRO = {"Enlucidos": "m2", "Fijos": "m2", "Fajas": "m", "Dinteles": "m"}
RENDIMIENTOS_TEORICOS = {"Enlucidos": 0.75, "Fijos": 0.50, "Fajas": 0.30, "Dinteles": 0.40}

ACTIVIDADES_MANANA = [
    "Verificación de asistencia del personal", "Distribución de cuadrillas por frente de trabajo",
    "Recorrido inicial de obra", "Supervisión de la ejecución de los trabajos",
    "Verificación de los trabajos y la calidad", "Coordinación con otras especialidades",
    "Corrección de observaciones detectadas",
]

ACTIVIDADES_TARDE = [
    "Recorrido de seguimiento de los frentes de trabajo", "Verificación del avance físico de las actividades",
    "Control del rendimiento de las cuadrillas", "Supervisión de la ejecución de los trabajos",
    "Verificación de los trabajos y la calidad", "Revisión de observaciones pendientes",
    "Verificación de trabajos corregidos", "Verificación del orden y limpieza de los frentes de trabajo",
    "Confirmación de materiales para el siguiente día", "Revisión del cumplimiento de la meta diaria",
    "Cierre de actividades en campo",
]
import base64
import datetime
import io
import json
import os
import pandas as pd
from PIL import Image, ImageOps
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from supabase import create_client, Client
from streamlit_local_storage import LocalStorage

# Importaciones de ReportLab para generación de PDFs
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ==========================================
st.set_page_config(
    page_title="Alpha Builders | Portal Ejecutivo",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    h1, h2, h3, .brand-title { font-family: 'Montserrat', sans-serif !important; letter-spacing: -0.03em !important; }

    .block-container { padding-top: 1rem !important; padding-bottom: 1.5rem !important; padding-left: 2.5rem !important; padding-right: 2.5rem !important; max-width: 100% !important; }
    .stApp { background-color: #ffffff !important; color: #121318 !important; }
    .stApp p, .stApp label, .stApp span, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6 { color: #121318; }
    .stCaption, caption, small, [data-testid="stCaptionContainer"] { color: #5a5f6e !important; }

    [data-testid="stInputInstructions"], div[data-testid="stInputInstructions"] { display: none !important; visibility: hidden !important; }

    [data-testid="stSidebarCollapseButton"] { display: block !important; visibility: visible !important; opacity: 1 !important; z-index: 999999 !important; }
    [data-testid="collapsedControl"] { display: block !important; visibility: visible !important; opacity: 1 !important; position: fixed !important; top: 15px !important; left: 15px !important; z-index: 999999 !important; }

    [data-testid="stSidebarCollapseButton"] button, [data-testid="collapsedControl"] button {
        background-color: #1c1e26 !important; border: 1px solid #323646 !important; border-radius: 50% !important; width: 36px !important; height: 36px !important; color: #ffffff !important; box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important; transition: all 0.2s ease !important;
    }
    [data-testid="stSidebarCollapseButton"] button:hover, [data-testid="collapsedControl"] button:hover {
        background-color: #ff8c00 !important; border-color: #ff8c00 !important; transform: scale(1.08);
    }
    [data-testid="stSidebarCollapseButton"] svg, [data-testid="collapsedControl"] svg { fill: #ffffff !important; color: #ffffff !important; }

    [data-testid="stSidebar"] { background-color: #121318 !important; border-right: 2px solid #282a36 !important; padding-top: 0px !important; padding-left: 12px !important; padding-right: 12px !important; padding-bottom: 15px !important; }
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0.5rem !important; padding-top: 0px !important; }
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] div { color: #ffffff !important; }

    .sidebar-logo-card { background-color: #ffffff; border-radius: 12px; padding: 8px 10px; margin-top: 0px !important; margin-bottom: 20px !important; box-shadow: 0 4px 12px rgba(0,0,0,0.3); width: 100% !important; box-sizing: border-box; text-align: center; display: block; }
    [data-testid="stSidebar"] [data-testid="stImage"] { width: 100% !important; display: block !important; margin-top: 6px !important; margin-bottom: 10px !important; clear: both !important; }
    [data-testid="stSidebar"] [data-testid="stImage"] img { border-radius: 12px !important; width: 100% !important; height: auto !important; max-width: 100% !important; object-fit: cover !important; border: 1px solid #323646 !important; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4); margin: 0 !important; display: block !important; }

    .sidebar-profile-box { background: #1c1e26; border: 1px solid #323646; border-radius: 12px; padding: 10px 8px !important; text-align: center; margin-top: 4px; margin-bottom: 8px; width: 100% !important; box-shadow: 0 4px 10px rgba(0,0,0,0.3); box-sizing: border-box; }
    .sidebar-user-nombres { font-size: 0.88rem; font-weight: 800; color: #ffffff !important; line-height: 1.2; }
    .sidebar-user-apellidos { font-size: 0.85rem; font-weight: 700; color: #e0e4ed !important; margin-bottom: 4px !important; line-height: 1.2; }
    .sidebar-user-email { font-size: 0.68rem; color: #72b2ff !important; font-weight: 600; margin-bottom: 6px !important; word-break: break-all; }
    .sidebar-user-cargo { display: inline-block; background: #323646 !important; color: #ffffff !important; border: 1px solid #484e5e !important; font-size: 0.60rem !important; font-weight: 800 !important; padding: 2px 8px !important; border-radius: 14px !important; text-transform: uppercase !important; }

    [data-testid="stSidebar"] hr { margin: 6px 0 !important; border-color: #282a36 !important; }
    [data-testid="stSidebar"] [data-testid="stExpander"] { background-color: #1c1e26 !important; border: 1px solid #323646 !important; border-radius: 10px !important; margin-top: 2px !important; margin-bottom: 6px !important; }
    [data-testid="stSidebar"] [data-testid="stExpander"] summary { background-color: #282c36 !important; padding: 6px 8px !important; }
    [data-testid="stSidebar"] [data-testid="stExpander"] summary * { color: #ffffff !important; font-weight: 700 !important; font-size: 0.78rem !important; }

    .executive-card-studio { background: linear-gradient(145deg, #f3f6fc 0%, #e8edf7 100%); border: 1px solid #b8c4d8; border-left: 7px solid #121318; border-radius: 22px; padding: 22px 28px; box-shadow: 0 12px 35px rgba(0,0,0,0.06); margin-bottom: 20px; width: 100%; box-sizing: border-box; }
    .brand-title { font-family: 'Montserrat', sans-serif !important; font-weight: 700 !important; font-size: 2.4rem !important; background: linear-gradient(90deg, #121318 0%, #3a4256 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; text-shadow: 0 2px 12px rgba(0,0,0,0.08); letter-spacing: -0.03em !important; }

    .kpi-card-studio { background: linear-gradient(145deg, #eceff6 0%, #dbe2ef 100%); border: 1px solid #aebacf; border-radius: 20px; padding: 18px; text-align: center; box-shadow: 0 8px 25px rgba(0,0,0,0.06); transition: all 0.3s ease; }
    .kpi-card-studio:hover { transform: translateY(-3px); box-shadow: 0 14px 35px rgba(0,0,0,0.12); filter: brightness(1.02); }
    .kpi-val-studio { font-size: 2.5rem; font-weight: 900; color: #121318 !important; }
    .kpi-lbl-studio { font-size: 0.72rem; color: #4a5060 !important; text-transform: uppercase; font-weight: 800; }

    .stTabs [data-baseweb="tab-list"] { gap: 10px; background-color: #e2e5ec !important; padding: 6px; border-radius: 16px; border: 1px solid #c2c7d2; }
    .stTabs [data-baseweb="tab"] { border-radius: 12px !important; padding: 10px 24px !important; background-color: transparent !important; }
    .stTabs [data-baseweb="tab"] p, .stTabs [data-baseweb="tab"] span { color: #121318 !important; font-weight: 700 !important; }
    .stTabs [aria-selected="true"] { background-color: #121318 !important; border-radius: 12px !important; }
    .stTabs [aria-selected="true"] p, .stTabs [aria-selected="true"] span, .stTabs [aria-selected="true"] div { color: #ffffff !important; font-weight: 900 !important; }

    .stButton > button { background-color: #121318 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; font-weight: 800 !important; padding: 10px 22px !important; }
    .stButton > button p, .stButton > button span { color: #ffffff !important; }

    .streamlit-expanderHeader { background-color: #e8eaee !important; border-radius: 12px !important; border: 1px solid #c2c7d2 !important; font-weight: 700 !important; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. CONEXIÓN Y CARGA SUPABASE
# ==========================================
@st.cache_resource
def init_supabase():
    url = st.secrets.get("SUPABASE_URL", "")
    key = st.secrets.get("SUPABASE_KEY", "")
    if not url or not key:
        st.error("⚠️ Credenciales SUPABASE_URL / SUPABASE_KEY no configuradas.")
        st.stop()
    return create_client(url, key)

supabase = init_supabase()
local_storage = LocalStorage()

DEFAULT_TRABAJADORES = []

def load_db_from_supabase():
    try:
        res_pin = supabase.table("app_config").select("*").eq("key", "access_pin").execute()
        access_pin = res_pin.data[0]["value"] if res_pin.data else "1254"
    except Exception:
        access_pin = "1254"

    try:
        res_usr = supabase.table("usuarios").select("*").execute()
        db_usuarios = []
        db_fotos = {}
        admin_emails = []
        for row in res_usr.data:
            c = row["correo"].lower().strip()
            db_usuarios.append({
                "Nombres": row["nombres"],
                "Apellidos": row["apellidos"],
                "Correo": c,
                "Password": row["password"],
                "Cargo": row["cargo"],
                "Fecha_Registro": str(row["fecha_registro"]),
                "Estado": row.get("estado", "Activo")
            })
            if row.get("foto_b64"):
                db_fotos[c] = row["foto_b64"]
            if row.get("es_admin"):
                admin_emails.append(c)
    except Exception:
        db_usuarios = []
        db_fotos = {}
        admin_emails = ["oscarsebitas2013@gmail.com"]

    if "oscarsebitas2013@gmail.com" not in admin_emails:
        admin_emails.append("oscarsebitas2013@gmail.com")

    try:
        res_trab = supabase.table("trabajadores").select("*").execute()
        db_trabajadores = [{"nombre": r["nombre"], "cargo": r["cargo"]} for r in res_trab.data] if res_trab.data else DEFAULT_TRABAJADORES
    except Exception:
        db_trabajadores = DEFAULT_TRABAJADORES

    # Carga Checklists
    db_checklists = {}
    try:
        res_chk = supabase.table("checklists").select("*").execute()
        for r in res_chk.data:
            c = r["usuario_email"].lower().strip()
            if c not in db_checklists:
                db_checklists[c] = []
            
            datos_parsed = r["datos"] if isinstance(r["datos"], list) else json.loads(r["datos"])
            db_checklists[c].append({
                "db_id": r["id"],
                "Fecha": str(r["fecha"]),
                "Hora_Inicio": r.get("hora_inicio", "07:00"),
                "Hora_Fin": r.get("hora_fin", "17:00"),
                "Edificio": r["edificio"],
                "Responsable": r.get("responsable", ""),
                "Cargo": r.get("cargo", ""),
                "Observacion_General": r.get("observacion_general", ""),
                "Datos": datos_parsed
            })
    except Exception:
        pass

    # Carga Formato de Inspección
    db_inspecciones = {}
    try:
        res_insp = supabase.table("inspecciones").select("*").execute()
        for r in res_insp.data:
            c = r["usuario_email"].lower().strip()
            if c not in db_inspecciones:
                db_inspecciones[c] = []
            
            datos_parsed = r["datos"] if isinstance(r["datos"], dict) else json.loads(r["datos"])
            db_inspecciones[c].append({
                "db_id": r["id"],
                "Fecha": str(r["fecha"]),
                "Dia": r.get("dia", ""),
                "Proyecto": r["proyecto"],
                "Residente": r.get("residente", ""),
                "Frente": r.get("frente", ""),
                "Clima": r.get("clima", ""),
                "Hora_Inicio": r.get("hora_inicio", "07:00"),
                "Hora_Fin": r.get("hora_fin", "17:00"),
                "Datos": datos_parsed
            })
    except Exception:
        pass

    # Carga Rendimientos
    db_rendimientos = {}
    try:
        res_rnd = supabase.table("rendimientos").select("*").execute()
        for r in res_rnd.data:
            c = r["usuario_email"].lower().strip()
            if c not in db_rendimientos:
                db_rendimientos[c] = []
            db_rendimientos[c].append({
                "db_id": r["id"],
                "Usuario_Registro": c,
                "Cargo_Registrador": r.get("cargo_obrero", ""),
                "Fecha": str(r["fecha"]),
                "Trabajador": r["trabajador"],
                "Cargo_Obrero": r.get("cargo_obrero", ""),
                "Rubro": r["rubro"],
                "Horas Trabajadas (HH)": float(r["horas_hh"]),
                "Avance": float(r["avance"]),
                "Unidad": r["unidad"],
                "Rend. Real (HH/Unid)": float(r["rend_real"]),
                "Rend. Teórico": float(r["rend_teorico"]),
                "Estado": r["estado"]
            })
    except Exception:
        pass

    return {
        "access_pin": access_pin,
        "admin_emails": admin_emails,
        "db_fotos_perfil_b64": db_fotos,
        "db_usuarios": db_usuarios,
        "db_checklists": db_checklists,
        "db_inspecciones": db_inspecciones,
        "db_rendimientos": db_rendimientos,
        "db_trabajadores": db_trabajadores,
    }

if "db_loaded" not in st.session_state or not st.session_state.db_loaded:
    p_data = load_db_from_supabase()
    st.session_state.access_pin = p_data["access_pin"]
    st.session_state.admin_emails = p_data["admin_emails"]
    st.session_state.db_fotos_perfil_b64 = p_data["db_fotos_perfil_b64"]
    st.session_state.db_usuarios = p_data["db_usuarios"]
    st.session_state.db_checklists = p_data["db_checklists"]
    st.session_state.db_inspecciones = p_data["db_inspecciones"]
    st.session_state.db_rendimientos = p_data["db_rendimientos"]
    st.session_state.db_trabajadores = p_data["db_trabajadores"]
    st.session_state.db_loaded = True

# PERSISTENCIA DE SESIÓN
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_email = ""
    st.session_state.usuario_nombres = ""
    st.session_state.usuario_apellidos = ""
    st.session_state.usuario_cargo = ""

if not st.session_state.autenticado:
    saved_token = local_storage.getItem("user_session_email")
    if saved_token:
        mail_clean = saved_token.strip().lower()
        u_match = next((u for u in st.session_state.db_usuarios if u["Correo"] == mail_clean), None)
        if u_match:
            st.session_state.autenticado = True
            st.session_state.usuario_email = mail_clean
            st.session_state.usuario_nombres = u_match["Nombres"]
            st.session_state.usuario_apellidos = u_match["Apellidos"]
            st.session_state.usuario_cargo = u_match["Cargo"]

def render_estado_badge(estado_str):
    if not estado_str:
        return '<span style="color: #64748b; font-weight: 600;">Sin Responder</span>'
    if "Cumple" in estado_str or estado_str in ["Sí", "Operativo", "Completado"]:
        return f'<span style="background-color: #dcfce7; color: #16a34a; font-weight: 800; padding: 3px 10px; border-radius: 8px; border: 1px solid #bbf7d0; font-size: 0.82rem;">{estado_str}</span>'
    elif "No" in estado_str or estado_str in ["Fuera de servicio", "Retrasado"]:
        return f'<span style="background-color: #fee2e2; color: #dc2626; font-weight: 800; padding: 3px 10px; border-radius: 8px; border: 1px solid #fca5a5; font-size: 0.82rem;">{estado_str}</span>'
    else:
        return f'<span style="background-color: #f1f5f9; color: #121318; font-weight: 800; padding: 3px 10px; border-radius: 8px; border: 1px solid #cbd5e1; font-size: 0.82rem;">{estado_str}</span>'

def image_to_base64(image_file):
    if image_file is not None:
        try:
            img = Image.open(image_file)
            img = ImageOps.exif_transpose(img)
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            return base64.b64encode(buffered.getvalue()).decode("utf-8")
        except Exception:
            return None
    return None

def base64_to_image(b64_str):
    if b64_str:
        try:
            img_data = base64.b64decode(b64_str)
            img = Image.open(io.BytesIO(img_data))
            img = ImageOps.exif_transpose(img)
            return img
        except Exception:
            return None
    return None

def get_repo_image_b64(filenames):
    for filename in filenames:
        if os.path.exists(filename):
            try:
                with open(filename, "rb") as f:
                    return base64.b64encode(f.read()).decode("utf-8")
            except Exception:
                pass
    return None

# ==========================================
# GENERADORES DE REPORTES EXCEL, PDF Y CSV
# ==========================================
def export_dataframe_to_excel_csv(df):
    """Exportación en CSV delimitado por punto y coma (para compatibilidad con Excel)."""
    df_clean = df.drop(columns=["Foto_B64", "db_id"], errors="ignore")
    return df_clean.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")

def export_checklist_to_excel_file(jornada_dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Checklist Obra"

    ws.merge_cells("A1:F1")
    ws["A1"] = f"INSPECCIÓN DE OBRA - {jornada_dict.get('Edificio', '')} ({jornada_dict.get('Fecha', '')})"
    ws["A1"].font = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    ws["A1"].fill = PatternFill(start_color="121318", end_color="121318", fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["A2"] = f"Hora Inicio: {jornada_dict.get('Hora_Inicio', 'N/A')}"
    ws["B2"] = f"Hora Fin: {jornada_dict.get('Hora_Fin', 'N/A')}"
    ws["C2"] = f"Responsable: {jornada_dict.get('Responsable', '')}"

    headers = ["Jornada", "N°", "Actividad", "Estado", "Observaciones", "Evidencia Fotográfica"]
    ws.append([])
    ws.append(headers)

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=4, column=col)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill(start_color="121318", end_color="121318", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    ws.row_dimensions[4].height = 25
    datos = jornada_dict.get("Datos", [])
    start_row = 5

    for row_idx, item in enumerate(datos, start=start_row):
        obs_val = item.get("Observaciones", "")
        if item.get("Actividades_Especificas"):
            sub_acts = " | Actividades: " + ", ".join([f"• {a['Actividad']}" for a in item["Actividades_Especificas"] if a.get("Actividad")])
            obs_val += sub_acts

        ws.append([item.get("Jornada", ""), item.get("N°", ""), item.get("Actividad", ""), item.get("Estado", ""), obs_val])
        ws.row_dimensions[row_idx].height = 140

        for c_i in range(1, 7):
            cell_txt = ws.cell(row=row_idx, column=c_i)
            cell_txt.border = thin_border
            cell_txt.font = Font(name="Arial", size=9)
            if c_i < 6:
                cell_txt.alignment = Alignment(vertical="center", wrap_text=True)

        foto_b64 = item.get("Foto_B64")
        if foto_b64:
            try:
                img_data = base64.b64decode(foto_b64)
                img_pil = Image.open(io.BytesIO(img_data))
                img_pil = ImageOps.exif_transpose(img_pil)
                img_pil = img_pil.resize((600, 450), Image.Resampling.LANCZOS)
                img_stream = io.BytesIO()
                img_pil.save(img_stream, format="PNG", quality=100)
                img_stream.seek(0)

                img_xlsx = OpenpyxlImage(img_stream)
                img_xlsx.width = 280
                img_xlsx.height = 140
                ws.add_image(img_xlsx, f"F{row_idx}")
            except Exception:
                pass

    if jornada_dict.get("Observacion_General"):
        last_r = len(datos) + start_row
        ws.cell(row=last_r, column=1, value="OBSERVACIÓN GENERAL:").font = Font(name="Arial", bold=True, size=10)
        ws.cell(row=last_r, column=3, value=jornada_dict.get("Observacion_General"))

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 35
    ws.column_dimensions['F'].width = 45

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def export_checklist_to_pdf_file(jornada_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []

    title_style = ParagraphStyle('TitleStyle', fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#121318'), alignment=1, spaceAfter=8)
    sub_style = ParagraphStyle('SubStyle', fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#333333'), spaceAfter=10)
    header_style = ParagraphStyle('HeaderStyle', fontName='Helvetica-Bold', fontSize=8.5, textColor=colors.white, alignment=1)
    cell_style = ParagraphStyle('CellStyle', fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#121318'))

    story.append(Paragraph(f"CHECKLIST DIARIO DE OBRA — {jornada_dict.get('Edificio', '').upper()}", title_style))
    story.append(Paragraph(f"<b>Fecha:</b> {jornada_dict.get('Fecha', '')} | <b>Horario:</b> {jornada_dict.get('Hora_Inicio', '')} - {jornada_dict.get('Hora_Fin', '')} | <b>Responsable:</b> {jornada_dict.get('Responsable', '')}", sub_style))

    data = [[
        Paragraph("<b>Jornada</b>", header_style),
        Paragraph("<b>N°</b>", header_style),
        Paragraph("<b>Actividad</b>", header_style),
        Paragraph("<b>Estado</b>", header_style),
        Paragraph("<b>Observaciones</b>", header_style)
    ]]
    
    for item in jornada_dict.get("Datos", []):
        obs = item.get("Observaciones", "")
        if item.get("Actividades_Especificas"):
            sub_acts = ", ".join([a['Actividad'] for a in item["Actividades_Especificas"] if a.get("Actividad")])
            if sub_acts:
                obs += f" (Específicas: {sub_acts})"

        data.append([
            Paragraph(str(item.get("Jornada", "")), cell_style),
            Paragraph(str(item.get("N°", "")), cell_style),
            Paragraph(str(item.get("Actividad", "")), cell_style),
            Paragraph(str(item.get("Estado", "") or "N/A"), cell_style),
            Paragraph(str(obs), cell_style)
        ])

    table = Table(data, colWidths=[55, 25, 190, 65, 215])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#121318')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(table)

    if jornada_dict.get("Observacion_General"):
        story.append(Spacer(1, 10))
        story.append(Paragraph(f"<b>Observación General:</b> {jornada_dict.get('Observacion_General')}", sub_style))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def export_inspeccion_to_excel_file(insp_dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inspección Diaria"

    # Estilos visuales
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), 
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), 
        bottom=Side(style='thin', color='CBD5E1')
    )
    fill_header = PatternFill(start_color="121318", end_color="121318", fill_type="solid")
    fill_section = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    
    font_main_title = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    font_section_title = Font(name="Arial", bold=True, color="1E293B", size=10)
    font_tbl_header = Font(name="Arial", bold=True, color="FFFFFF", size=9.5)
    font_bold = Font(name="Arial", bold=True, color="121318", size=9.5)
    font_regular = Font(name="Arial", size=9)

    # 1. Título Principal
    ws.merge_cells("A1:C1")
    ws["A1"] = f"FORMATO DE INSPECCIÓN DIARIA - {insp_dict.get('Proyecto', '').upper()}"
    ws["A1"].font = font_main_title
    ws["A1"].fill = fill_header
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Meta Información
    meta_info = [
        ["Fecha:", f"{insp_dict.get('Fecha', '')} ({insp_dict.get('Dia', '')})", ""],
        ["Residente de Obra:", insp_dict.get("Residente", ""), ""],
        ["Frente Inspeccionado:", insp_dict.get("Frente", ""), ""],
        ["Clima:", insp_dict.get("Clima", ""), f"Horario: {insp_dict.get('Hora_Inicio', '')} - {insp_dict.get('Hora_Fin', '')}"]
    ]
    for r in meta_info:
        ws.append(r)
        r_idx = ws.max_row
        ws.cell(row=r_idx, column=1).font = font_bold
        ws.cell(row=r_idx, column=2).font = font_regular
        ws.cell(row=r_idx, column=3).font = font_regular

    ws.append([])
    datos = insp_dict.get("Datos", {})

    # 2. Avance General
    r_av_head = ws.max_row + 1
    ws.merge_cells(f"A{r_av_head}:C{r_av_head}")
    ws[f"A{r_av_head}"] = "1. AVANCE GENERAL"
    ws[f"A{r_av_head}"].font = font_section_title
    ws[f"A{r_av_head}"].fill = fill_section
    ws[f"A{r_av_head}"].alignment = Alignment(vertical="center", indent=1)

    ws.append(["Actividad", "% Prog / % Ejec", "Estado"])
    r_hdr1 = ws.max_row
    for col_i in range(1, 4):
        c = ws.cell(row=r_hdr1, column=col_i)
        c.font = font_tbl_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")

    for av in datos.get("Avance", []):
        prog_ejec_str = f"Prog: {av.get('% Prog', 0)}% | Ejec: {av.get('% Ejec', 0)}%"
        ws.append([av.get("Actividad", ""), prog_ejec_str, av.get("Estado", "") or "N/A"])
        r_i = ws.max_row
        for col_i in range(1, 4):
            cell = ws.cell(row=r_i, column=col_i)
            cell.font = font_regular
            cell.border = thin_border
            if col_i in [2, 3]:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.append([])

    # 3. Check List General Agrupado por Secciones (Formato estructurado por módulos)
    r_chk_main = ws.max_row + 1
    ws.merge_cells(f"A{r_chk_main}:C{r_chk_main}")
    ws[f"A{r_chk_main}"] = "2. CHECK LIST Y CONTROL GENERAL DE OBRA"
    ws[f"A{r_chk_main}"].font = font_section_title
    ws[f"A{r_chk_main}"].fill = fill_section
    ws[f"A{r_chk_main}"].alignment = Alignment(vertical="center", indent=1)

    checklist_groups = datos.get("Checklist", {})

    for sec_name, items in checklist_groups.items():
        if items:
            ws.append([])
            r_sec = ws.max_row
            ws.merge_cells(f"A{r_sec}:C{r_sec}")
            ws[f"A{r_sec}"] = f"■ {sec_name.upper()}"
            ws[f"A{r_sec}"].font = font_bold
            
            # Cabecera de la sub-tabla
            ws.append(["Ítem / Aspecto Inspeccionado", "Cumple / Estado", "Observación"])
            r_hdr_sub = ws.max_row
            ws.row_dimensions[r_hdr_sub].height = 22
            
            for col_i in range(1, 4):
                c = ws.cell(row=r_hdr_sub, column=col_i)
                c.font = font_tbl_header
                c.fill = fill_header
                c.alignment = Alignment(horizontal="center", vertical="center")

            # Filas de la sección
            for it in items:
                it_nombre = it.get("Item") or it.get("Aspecto") or it.get("Revisar") or it.get("Equipo", "")
                it_estado = it.get("Cumple") or it.get("Estado", "") or "N/A"
                it_obs = it.get("Observación", "")

                ws.append([it_nombre, it_estado, it_obs])
                r_i = ws.max_row
                
                c1 = ws.cell(row=r_i, column=1)
                c2 = ws.cell(row=r_i, column=2)
                c3 = ws.cell(row=r_i, column=3)

                c1.font = font_regular
                c2.font = font_regular
                c3.font = font_regular

                c1.border = thin_border
                c2.border = thin_border
                c3.border = thin_border

                c1.alignment = Alignment(vertical="center", wrap_text=True)
                c2.alignment = Alignment(horizontal="center", vertical="center")
                c3.alignment = Alignment(vertical="center", wrap_text=True)

    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 45

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def export_inspeccion_to_pdf_file(insp_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    story = []

    title_style = ParagraphStyle('TitleStyle', fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#121318'), alignment=1, spaceAfter=6)
    sub_title = ParagraphStyle('SubTitle', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#121318'), spaceBefore=8, spaceAfter=4)
    sec_header_style = ParagraphStyle('SecHeaderStyle', fontName='Helvetica-Bold', fontSize=8.5, textColor=colors.HexColor('#1e293b'), spaceBefore=6, spaceAfter=2)
    header_style = ParagraphStyle('HeaderStyle', fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)
    cell_style = ParagraphStyle('CellStyle', fontName='Helvetica', fontSize=7.5, textColor=colors.HexColor('#121318'))

    story.append(Paragraph(f"FORMATO DE INSPECCIÓN DIARIA DE OBRA", title_style))
    
    meta_text = f"<b>Proyecto:</b> {insp_dict.get('Proyecto', '')} | <b>Fecha:</b> {insp_dict.get('Fecha', '')} ({insp_dict.get('Dia', '')})<br/>" \
                f"<b>Residente:</b> {insp_dict.get('Residente', '')} | <b>Frente:</b> {insp_dict.get('Frente', '')}<br/>" \
                f"<b>Clima:</b> {insp_dict.get('Clima', '')} | <b>Horario:</b> {insp_dict.get('Hora_Inicio', '')} - {insp_dict.get('Hora_Fin', '')}"
    
    story.append(Paragraph(meta_text, cell_style))
    story.append(Spacer(1, 6))

    datos = insp_dict.get("Datos", {})

    story.append(Paragraph("1. AVANCE GENERAL", sub_title))
    av_data = [[
        Paragraph("<b>Actividad</b>", header_style),
        Paragraph("<b>% Prog</b>", header_style),
        Paragraph("<b>% Ejec</b>", header_style),
        Paragraph("<b>Estado</b>", header_style)
    ]]
    for av in datos.get("Avance", []):
        av_data.append([
            Paragraph(av.get("Actividad", ""), cell_style),
            Paragraph(str(av.get("% Prog", 0)), cell_style),
            Paragraph(str(av.get("% Ejec", 0)), cell_style),
            Paragraph(str(av.get("Estado", "") or "N/A"), cell_style)
        ])
    t_av = Table(av_data, colWidths=[190, 75, 75, 210])
    t_av.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#121318')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_av)

    story.append(Spacer(1, 6))
    story.append(Paragraph("2. CHECK LIST Y CONTROL GENERAL DE OBRA", sub_title))

    checklist_groups = datos.get("Checklist", {})

    for sec_name, items in checklist_groups.items():
        if items:
            story.append(Paragraph(f"📌 <b>{sec_name.upper()}</b>", sec_header_style))
            chk_data = [[
                Paragraph("<b>Ítem / Aspecto Inspeccionado</b>", header_style),
                Paragraph("<b>Cumple / Estado</b>", header_style),
                Paragraph("<b>Observación</b>", header_style)
            ]]

            for it in items:
                it_nombre = it.get("Item") or it.get("Aspecto") or it.get("Revisar") or it.get("Equipo", "")
                it_estado = it.get("Cumple") or it.get("Estado", "") or "N/A"
                it_obs = it.get("Observación", "")

                chk_data.append([
                    Paragraph(str(it_nombre), cell_style),
                    Paragraph(str(it_estado), cell_style),
                    Paragraph(str(it_obs), cell_style)
                ])

            t_sec = Table(chk_data, colWidths=[240, 100, 210])
            t_sec.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#121318')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ]))
            story.append(t_sec)
            story.append(Spacer(1, 4))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# 3. CONSTANTES
# ==========================================
EDIFICIOS_ALPHA = [
    "Tesla", "Lafuente", "Imagine", "Asimov", "Rubik", "Castle Rock",
    "Musk", "Wolf", "Dablanc", "Thomas Edison", "Westinghouse", "Smart",
]

UNIDADES_RUBRO = {"Enlucidos": "m2", "Fijos": "m2", "Fajas": "m", "Dinteles": "m"}
RENDIMIENTOS_TEORICOS = {"Enlucidos": 0.75, "Fijos": 0.50, "Fajas": 0.30, "Dinteles": 0.40}

ACTIVIDADES_MANANA = [
    "Verificación de asistencia del personal", "Distribución de cuadrillas por frente de trabajo",
    "Recorrido inicial de obra", "Supervisión de la ejecución de los trabajos",
    "Verificación de los trabajos y la calidad", "Coordinación con otras especialidades",
    "Corrección de observaciones detectadas",
]

ACTIVIDADES_TARDE = [
    "Recorrido de seguimiento de los frentes de trabajo", "Verificación del avance físico de las actividades",
    "Control del rendimiento de las cuadrillas", "Supervisión de la ejecución de los trabajos",
    "Verificación de los trabajos y la calidad", "Revisión de observaciones pendientes",
    "Verificación de trabajos corregidos", "Verificación del orden y limpieza de los frentes de trabajo",
    "Confirmación de materiales para el siguiente día", "Revisión del cumplimiento de la meta diaria",
    "Cierre de actividades en campo",
]
# ==========================================
# 7. ASIGNACIÓN DE PESTAÑAS Y MÓDULOS
# ==========================================
tab_chk = tabs_app[0]
tab_didactico = tabs_app[1]
tab_rend = tabs_app[2]

# ------------------------------------------
# MÓDULO 1: CHECKLIST DIARIO
# ------------------------------------------
with tab_chk:
    if "creando_jornada" not in st.session_state:
        st.session_state.creando_jornada = False

    st.markdown("### Check List Diario – Control de Obra")
    st.caption("Supervisión diaria de frentes de trabajo con verificación manual y registro de horas.")

    if not st.session_state.creando_jornada:
        if st.button("➕ Crear Nueva Jornada de Inspección", type="primary"):
            st.session_state.creando_jornada = True

    if st.session_state.creando_jornada:
        st.markdown("---")
        with st.container():
            st.markdown("#### Configuración de la Nueva Jornada")
            
            cfg_c1, cfg_c2, cfg_c3, cfg_c4, cfg_c5 = st.columns([2, 2, 2, 1.5, 1.5])
            with cfg_c1:
                edificio_val = st.selectbox("Edificio / Proyecto:", ["-- Seleccione --"] + EDIFICIOS_ALPHA, index=0, key="sel_edificio")
            with cfg_c2:
                st.text_input("Responsable:", value=user_nombre_completo, disabled=True)
            with cfg_c3:
                fecha_val = st.date_input("Fecha:", datetime.date.today(), key="sel_fecha")
            with cfg_c4:
                hora_inicio_val = st.time_input("Hora Inicio:", datetime.time(7, 0), key="sel_hora_inicio")
            with cfg_c5:
                hora_fin_val = st.time_input("Hora Fin:", datetime.time(17, 0), key="sel_hora_fin")

            st.markdown("---")

            with st.form("form_checklist_jornada"):
                st.markdown("#### 🌅 Jornada de la Mañana")
                resp_manana = []

                for idx, act in enumerate(ACTIVIDADES_MANANA, 1):
                    st.markdown(f"**N° {idx}. {act}**")
                    
                    sub_actividades_m = []
                    if act == "Supervisión de la ejecución de los trabajos":
                        st.caption("📌 Indique las actividades específicas a realizar durante la jornada de la mañana:")
                        df_sub_m = pd.DataFrame([
                            {"Actividad": ""},
                            {"Actividad": ""}
                        ])
                        sub_edited_m = st.data_editor(
                            df_sub_m,
                            num_rows="dynamic",
                            column_config={
                                "Actividad": st.column_config.TextColumn("Descripción de la Actividad a Realizar")
                            },
                            key=f"sub_m_{idx}"
                        )
                        sub_actividades_m = sub_edited_m.to_dict(orient="records")

                    c_sel, c_obs, c_foto = st.columns([2, 3, 3])

                    with c_sel:
                        est = st.segmented_control(
                            "Estado General",
                            ["✓ Cumple", "✗ No Cumple", "N/A"],
                            key=f"m_st_{idx}"
                        )
                    with c_obs:
                        ob = st.text_input("Observación", key=f"m_ob_{idx}", placeholder="Observaciones...", label_visibility="collapsed")
                    with c_foto:
                        ft = st.file_uploader("Foto Evidencia (Opcional)", type=["jpg", "jpeg", "png"], key=f"m_ft_{idx}")

                    st.markdown("<hr style='margin: 8px 0; border-color: #c2c7d2;'>", unsafe_allow_html=True)
                    
                    ft_b64 = image_to_base64(ft) if ft is not None else None
                    resp_manana.append({
                        "Jornada": "Mañana",
                        "N°": idx,
                        "Actividad": act,
                        "Estado": est,
                        "Observaciones": ob,
                        "Actividades_Especificas": sub_actividades_m,
                        "Foto_B64": ft_b64,
                        "Foto_Adjunta": "Sí" if ft is not None else "No"
                    })

                st.markdown("#### 🌆 Jornada de la Tarde")
                resp_tarde = []

                for idx, act in enumerate(ACTIVIDADES_TARDE, 1):
                    st.markdown(f"**N° {idx}. {act}**")
                    
                    sub_actividades_t = []
                    if act == "Supervisión de la ejecución de los trabajos":
                        st.caption("📌 Indique las actividades específicas a realizar durante la jornada de la tarde:")
                        df_sub_t = pd.DataFrame([
                            {"Actividad": ""},
                            {"Actividad": ""}
                        ])
                        sub_edited_t = st.data_editor(
                            df_sub_t,
                            num_rows="dynamic",
                            column_config={
                                "Actividad": st.column_config.TextColumn("Descripción de la Actividad a Realizar")
                            },
                            key=f"sub_t_{idx}"
                        )
                        sub_actividades_t = sub_edited_t.to_dict(orient="records")

                    c_sel, c_obs, c_foto = st.columns([2, 3, 3])

                    with c_sel:
                        est = st.segmented_control(
                            "Estado General",
                            ["✓ Cumple", "✗ No Cumple", "N/A"],
                            key=f"t_st_{idx}"
                        )
                    with c_obs:
                        ob = st.text_input("Observación", key=f"t_ob_{idx}", placeholder="Observaciones...", label_visibility="collapsed")
                    with c_foto:
                        ft = st.file_uploader("Foto Evidencia (Opcional)", type=["jpg", "jpeg", "png"], key=f"t_ft_{idx}")

                    st.markdown("<hr style='margin: 8px 0; border-color: #c2c7d2;'>", unsafe_allow_html=True)
                    
                    ft_b64 = image_to_base64(ft) if ft is not None else None
                    resp_tarde.append({
                        "Jornada": "Tarde",
                        "N°": idx,
                        "Actividad": act,
                        "Estado": est,
                        "Observaciones": ob,
                        "Actividades_Especificas": sub_actividades_t,
                        "Foto_B64": ft_b64,
                        "Foto_Adjunta": "Sí" if ft is not None else "No"
                    })

                st.markdown("#### 📝 Cierre de la Inspección")
                obs_general_val = st.text_area("Observación General de la Inspección:", placeholder="Ingrese observaciones generales, recomendaciones o novedades globales del frente de obra...", key="obs_general_inp")

                btn_guardar_chk = st.form_submit_button("Guardar Jornada de Inspección", type="primary")

            if btn_guardar_chk:
                if edificio_val == "-- Seleccione --" or not edificio_val:
                    st.error("⚠️ Por favor seleccione un Edificio o Proyecto válido.")
                else:
                    manana_respondida = [item for item in resp_manana if item["Estado"] is not None]
                    tarde_respondida = [item for item in resp_tarde if item["Estado"] is not None]

                    all_chk_data = manana_respondida + tarde_respondida

                    if len(all_chk_data) == 0:
                        st.error("⚠️ Por favor responda al menos a una actividad (Mañana o Tarde) para guardar.")
                    else:
                        try:
                            supabase.table("checklists").insert({
                                "usuario_email": user_email,
                                "edificio": edificio_val,
                                "fecha": fecha_val.strftime("%Y-%m-%d"),
                                "hora_inicio": hora_inicio_val.strftime("%H:%M"),
                                "hora_fin": hora_fin_val.strftime("%H:%M"),
                                "responsable": user_nombre_completo,
                                "cargo": user_cargo,
                                "observacion_general": obs_general_val.strip(),
                                "datos": all_chk_data
                            }).execute()

                            st.session_state.db_loaded = False
                            st.success(f"¡Jornada guardada permanentemente en Supabase para **{edificio_val}**!")
                            st.session_state.creando_jornada = False
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al guardar checklist en Supabase: {e}")

    st.markdown("---")

    # HISTORIAL DE CHECKLISTS POR EDIFICIO
    st.markdown("### Historial General de Checklists Creados")

    mis_jornadas = st.session_state.db_checklists.get(user_email, [])

    if len(mis_jornadas) > 0:
        col_edif_sel, _ = st.columns([2, 2])
        with col_edif_sel:
            edificio_filtro = st.selectbox(
                "🏢 Seleccionar Edificio / Proyecto:",
                ["-- Todos los Edificios --"] + EDIFICIOS_ALPHA,
                key="filtro_edificio_historial"
            )

        jornadas_filtradas = [j for j in mis_jornadas if j.get("Edificio") == edificio_filtro] if edificio_filtro != "-- Todos los Edificios --" else mis_jornadas.copy()

        st.caption(f"Mostrando **{len(jornadas_filtradas)}** de **{len(mis_jornadas)}** checklist(s) registrados.")

        if len(jornadas_filtradas) > 0:
            meses_nombres = {
                "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
                "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
                "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
            }

            jornadas_con_index = [(orig_idx, j_item) for orig_idx, j_item in enumerate(mis_jornadas) if j_item in jornadas_filtradas]
            jornadas_con_index.sort(key=lambda x: x[1]['Fecha'], reverse=True)

            grupos_meses = {}
            for orig_idx, j in jornadas_con_index:
                f_str = j['Fecha']
                try:
                    partes = f_str.split("-")
                    nombre_mes = f"{meses_nombres.get(partes[1], 'Mes')} {partes[0]}"
                except Exception:
                    nombre_mes = "Otros"
                
                if nombre_mes not in grupos_meses:
                    grupos_meses[nombre_mes] = []
                grupos_meses[nombre_mes].append((orig_idx, j))

            for mes_anio, lista_j in grupos_meses.items():
                with st.expander(f"📅 {mes_anio} ({len(lista_j)} checklists)", expanded=True):
                    for idx_rel, (orig_idx, j) in enumerate(lista_j):
                        col_j_info, col_j_del = st.columns([8, 1])
                        
                        with col_j_del:
                            if st.button("🗑️", key=f"quick_del_{orig_idx}", help="Borrar checklist permanentemente"):
                                try:
                                    supabase.table("checklists").delete().eq("id", j["db_id"]).execute()
                                    st.session_state.db_loaded = False
                                    st.success("¡Checklist eliminado!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error al eliminar: {e}")

                        with col_j_info:
                            with st.expander(f"📌 {j['Edificio']} — {j['Fecha']} (Horario: {j.get('Hora_Inicio', 'N/A')} - {j.get('Hora_Fin', 'N/A')})"):
                                st.markdown("#### Actividades Registradas y Evidencias:")
                                df_data = pd.DataFrame(j["Datos"])
                                for _, row in df_data.iterrows():
                                    estado_badge = render_estado_badge(row['Estado'])
                                    st.markdown(f"- **[{row['Jornada']}] N° {row['N°']}. {row['Actividad']}**: {estado_badge}", unsafe_allow_html=True)
                                    
                                    sub_tasks = row.get("Actividades_Especificas", [])
                                    if sub_tasks:
                                        for st_item in sub_tasks:
                                            if st_item.get("Actividad"):
                                                st.markdown(f"  * ▫️ *Actividad:* {st_item['Actividad']}")

                                    if row['Observaciones']:
                                        st.caption(f"Obs: {row['Observaciones']}")
                                    
                                    if row.get("Foto_B64") is not None:
                                        img_evidencia = base64_to_image(row["Foto_B64"])
                                        if img_evidencia:
                                            with st.popover("📷 Vista previa"):
                                                st.image(img_evidencia, caption=row['Actividad'], use_container_width=True)

                                    st.markdown("<hr style='margin: 4px 0; border-color: #cbd5e1;'>", unsafe_allow_html=True)

                                if j.get("Observacion_General"):
                                    st.info(f"**Observación General:** {j.get('Observacion_General')}")

                                c_dl1, c_dl2 = st.columns(2)
                                with c_dl1:
                                    excel_bytes = export_checklist_to_excel_file(j)
                                    st.download_button(
                                        label="📊 Descargar Excel (.xlsx)",
                                        data=excel_bytes,
                                        file_name=f"Checklist_{j['Edificio'].replace(' ', '_')}_{j['Fecha']}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key=f"dl_xlsx_{orig_idx}"
                                    )
                                with c_dl2:
                                    pdf_bytes = export_checklist_to_pdf_file(j)
                                    st.download_button(
                                        label="📄 Descargar PDF (.pdf)",
                                        data=pdf_bytes,
                                        file_name=f"Checklist_{j['Edificio'].replace(' ', '_')}_{j['Fecha']}.pdf",
                                        mime="application/pdf",
                                        key=f"dl_pdf_{orig_idx}"
                                    )
        else:
            st.warning("No hay checklists registrados para el edificio seleccionado.")
    else:
        st.info("Aún no hay checklists guardados.")

# ------------------------------------------
# MÓDULO 2: INSPECCIÓN DIARIA
# ------------------------------------------
with tab_didactico:
    st.markdown("### Formato de Inspección Diaria de Obra")
    st.caption("Supervisión técnica paso a paso con tabuladores y registro permanente.")

    dias_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    
    col_f_out, _ = st.columns([1, 2])
    with col_f_out:
        did_fecha_fuera = st.date_input("Fecha de Inspección:", datetime.date.today(), key="did_fecha_live")
    
    dia_auto_es = dias_es[did_fecha_fuera.weekday()]

    with st.form("form_didactico_1_6"):
        st.markdown("#### 1. Información General")
        c1, c2 = st.columns(2)
        with c1:
            did_proyecto = st.selectbox("Proyecto:", ["-- Seleccione --"] + EDIFICIOS_ALPHA, index=0, key="did_proy")
            did_dia = st.text_input("Día:", value=dia_auto_es, disabled=True)
            did_residente = st.text_input("Residente de Obra:", value=user_nombre_completo, key="did_res")

        with c2:
            did_frente = st.text_input("Frente Inspeccionado:", placeholder="Ej. Bloque A - Piso 3", key="did_fre")
            did_clima = st.selectbox("Clima:", ["Soleado", "Nublado", "Lluvia"], index=None, placeholder="Seleccionar clima...", key="did_cli_single")
            c_h1, c_h2 = st.columns(2)
            with c_h1:
                did_h_ini = st.time_input("Hora inicio:", datetime.time(7, 0), key="did_hini")
            with c_h2:
                did_h_fin = st.time_input("Hora fin:", datetime.time(17, 0), key="did_hfin")

        st.markdown("---")

        # 2. AVANCE GENERAL
        st.markdown("#### 2. Avance General")
        sub_actividades_pdf = ["Movimiento de tierras", "Estructura", "Mampostería", "Enlucidos", "Instalaciones", "Acabados"]
        avance_datos = []
        for act in sub_actividades_pdf:
            col_a1, col_a2, col_a3, col_a4 = st.columns([3, 2, 2, 2])
            with col_a1:
                st.write(f"**{act}**")
            with col_a2:
                p_prog = st.number_input(f"% Prog. ({act})", min_value=0.0, max_value=100.0, step=1.0, key=f"prog_{act}")
            with col_a3:
                p_ejec = st.number_input(f"% Ejec. ({act})", min_value=0.0, max_value=100.0, step=1.0, key=f"ejec_{act}")
            with col_a4:
                est_act = st.selectbox(
                    f"Estado ({act})", 
                    ["En Proceso", "Completado", "Retrasado", "N/A"], 
                    index=None, 
                    placeholder="Seleccionar...", 
                    key=f"est_{act}"
                )
            avance_datos.append({"Actividad": act, "% Prog": p_prog, "% Ejec": p_ejec, "Estado": est_act})

        st.markdown("---")

        # 3. CHECK LIST GENERAL
        st.markdown("#### 3. Check List General")
        
        tab_sec1, tab_sec2, tab_sec3, tab_sec4, tab_sec5 = st.tabs([
            "🛡️ Seguridad Industrial", "🧱 Mampostería", "🏗️ Hormigón", "🔌 Instalaciones", "🎨 Acabados"
        ])

        with tab_sec1:
            sec_seguridad = ["Personal con casco", "Uso correcto de EPP", "Arnés", "Andamios", "Señalización", "Orden y limpieza", "Extintores", "Botiquín"]
            seg_resp = []
            for item in sec_seguridad:
                cs1, cs2, cs3 = st.columns([3, 2, 3])
                with cs1:
                    st.write(f"• {item}")
                with cs2:
                    st_val = st.segmented_control(f"Seg_{item}", ["Sí", "No", "N/A"], key=f"seg_{item}", label_visibility="collapsed")
                with cs3:
                    obs_val = st.text_input(f"Obs_{item}", placeholder="Observaciones...", key=f"seg_obs_{item}", label_visibility="collapsed")
                seg_resp.append({"Item": item, "Cumple": st_val, "Observación": obs_val})

        with tab_sec2:
            sec_mamp = ["Muros aplomados", "Muros nivelados", "Chicotes", "Tensores", "Juntas", "Limpieza"]
            mamp_resp = []
            for item in sec_mamp:
                cm1, cm2, cm3 = st.columns([3, 2, 3])
                with cm1:
                    st.write(f"• {item}")
                with cm2:
                    st_val = st.segmented_control(f"Mamp_{item}", ["Sí", "No"], key=f"mamp_{item}", label_visibility="collapsed")
                with cm3:
                    obs_val = st.text_input(f"Obs_mamp_{item}", placeholder="Observaciones...", key=f"mamp_obs_{item}", label_visibility="collapsed")
                mamp_resp.append({"Item": item, "Cumple": st_val, "Observación": obs_val})

        with tab_sec3:
            sec_horm = ["Acero conforme planos", "Recubrimiento", "Vibrado", "Sin cangrejeras", "Sin juntas frías", "Curado"]
            horm_resp = []
            for item in sec_horm:
                ch1, ch2, ch3 = st.columns([3, 2, 3])
                with ch1:
                    st.write(f"• {item}")
                with ch2:
                    st_val = st.segmented_control(f"Horm_{item}", ["Sí", "No"], key=f"horm_{item}", label_visibility="collapsed")
                with ch3:
                    obs_val = st.text_input(f"Obs_horm_{item}", placeholder="Observaciones...", key=f"horm_obs_{item}", label_visibility="collapsed")
                horm_resp.append({"Item": item, "Cumple": st_val, "Observación": obs_val})

        with tab_sec4:
            sec_inst = ["Eléctrica", "Sanitaria", "Hidráulica", "Ductería", "Cajas niveladas"]
            inst_resp = []
            for item in sec_inst:
                ci1, ci2, ci3 = st.columns([3, 2, 3])
                with ci1:
                    st.write(f"• {item}")
                with ci2:
                    st_val = st.segmented_control(f"Inst_{item}", ["Sí", "No"], key=f"inst_{item}", label_visibility="collapsed")
                with ci3:
                    obs_val = st.text_input(f"Obs_inst_{item}", placeholder="Observaciones...", key=f"inst_obs_{item}", label_visibility="collapsed")
                inst_resp.append({"Item": item, "Cumple": st_val, "Observación": obs_val})

        with tab_sec5:
            sec_acab = ["Enlucidos", "Cerámica", "Pintura", "Puertas", "Ventanas"]
            acab_resp = []
            for item in sec_acab:
                ca1, ca2, ca3 = st.columns([3, 2, 3])
                with ca1:
                    st.write(f"• {item}")
                with ca2:
                    st_val = st.segmented_control(f"Acab_{item}", ["Sí", "No"], key=f"acab_{item}", label_visibility="collapsed")
                with ca3:
                    obs_val = st.text_input(f"Obs_acab_{item}", placeholder="Observaciones...", key=f"acab_obs_{item}", label_visibility="collapsed")
                acab_resp.append({"Item": item, "Cumple": st_val, "Observación": obs_val})

        st.markdown("---")

        # 4. CONTROL DE PERSONAL
        st.markdown("#### 4. Control de Personal")
        sec_pers = ["Personal completo", "Contratistas completos", "Rendimiento adecuado", "Retrasos"]
        pers_resp = []
        for item in sec_pers:
            cp1, cp2, cp3 = st.columns([3, 2, 3])
            with cp1:
                st.write(f"• {item}")
            with cp2:
                st_val = st.segmented_control(f"Pers_{item}", ["Sí", "No"], key=f"pers_{item}", label_visibility="collapsed")
            with cp3:
                obs_val = st.text_input(f"Obs_pers_{item}", placeholder="Observación...", key=f"pers_obs_{item}", label_visibility="collapsed")
            pers_resp.append({"Aspecto": item, "Cumple": st_val, "Observación": obs_val})

        st.markdown("---")

        # 5. MATERIALES
        st.markdown("#### 5. Materiales")
        sec_mat = ["Material suficiente", "Material conforme", "Material almacenado correctamente", "Material deteriorado"]
        mat_resp = []
        for item in sec_mat:
            cmat1, cmat2, cmat3 = st.columns([3, 2, 3])
            with cmat1:
                st.write(f"• {item}")
            with cmat2:
                st_val = st.segmented_control(f"Mat_{item}", ["Sí", "No"], key=f"mat_{item}", label_visibility="collapsed")
            with cmat3:
                obs_val = st.text_input(f"Obs_mat_{item}", placeholder="Observación...", key=f"mat_obs_{item}", label_visibility="collapsed")
            mat_resp.append({"Revisar": item, "Cumple": st_val, "Observación": obs_val})

        st.markdown("---")

        # 6. EQUIPOS
        st.markdown("#### 6. Equipos")
        sec_eq = ["Mezcladora", "Vibrador", "Cortadora", "Compresor", "Herramienta eléctrica"]
        eq_resp = []
        for item in sec_eq:
            ceq1, ceq2, ceq3 = st.columns([3, 2, 3])
            with ceq1:
                st.write(f"• {item}")
            with ceq2:
                st_val = st.segmented_control(f"Eq_{item}", ["Operativo", "Fuera de servicio"], key=f"eq_{item}", label_visibility="collapsed")
            with ceq3:
                obs_val = st.text_input(f"Obs_eq_{item}", placeholder="Observación...", key=f"eq_obs_{item}", label_visibility="collapsed")
            eq_resp.append({"Equipo": item, "Estado": st_val, "Observación": obs_val})

        btn_guardar_did = st.form_submit_button("💾 Guardar Formato de Inspección", type="primary")

        if btn_guardar_did:
            if did_proyecto == "-- Seleccione --":
                st.error("⚠️ Por favor seleccione un Proyecto.")
            elif not did_clima:
                st.error("⚠️ Por favor seleccione las condiciones del clima.")
            else:
                payload_insp = {
                    "Avance": avance_datos,
                    "Checklist": {
                        "Seguridad Industrial": seg_resp,
                        "Mampostería": mamp_resp,
                        "Hormigón": horm_resp,
                        "Instalaciones": inst_resp,
                        "Acabados": acab_resp,
                        "Control de Personal": pers_resp,
                        "Materiales": mat_resp,
                        "Equipos": eq_resp
                    }
                }
                try:
                    supabase.table("inspecciones").insert({
                        "usuario_email": user_email,
                        "proyecto": did_proyecto,
                        "fecha": did_fecha_fuera.strftime("%Y-%m-%d"),
                        "dia": dia_auto_es,
                        "residente": did_residente,
                        "frente": did_frente,
                        "clima": did_clima,
                        "hora_inicio": did_h_ini.strftime("%H:%M"),
                        "hora_fin": did_h_fin.strftime("%H:%M"),
                        "datos": payload_insp
                    }).execute()

                    st.session_state.db_loaded = False
                    st.success(f"¡Formato de inspección registrado y guardado permanentemente para **{did_proyecto}**!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al guardar inspección en Supabase: {e}")

    st.markdown("---")

    # HISTORIAL DE FORMATOS DE INSPECCIÓN DIARIA (CON EDICIÓN Y ELIMINACIÓN)
    st.markdown("### Historial General de Formatos de Inspección Creados")

    mis_inspecciones = st.session_state.db_inspecciones.get(user_email, [])

    if len(mis_inspecciones) > 0:
        col_edif_insp, _ = st.columns([2, 2])
        with col_edif_insp:
            edif_insp_filtro = st.selectbox("🏢 Seleccionar Proyecto para consultar:", ["-- Todos los Proyectos --"] + EDIFICIOS_ALPHA, key="filtro_edif_inspecciones")

        insps_filtradas = [i for i in mis_inspecciones if i.get("Proyecto") == edif_insp_filtro] if edif_insp_filtro != "-- Todos los Proyectos --" else mis_inspecciones.copy()

        st.caption(f"Mostrando **{len(insps_filtradas)}** de **{len(mis_inspecciones)}** inspección(es) registradas.")

        if len(insps_filtradas) > 0:
            insps_filtradas.sort(key=lambda x: x['Fecha'], reverse=True)

            for idx_insp, insp in enumerate(insps_filtradas, 1):
                col_i_info, col_i_del = st.columns([8, 1])

                with col_i_del:
                    if st.button("🗑️", key=f"del_insp_{idx_insp}", help="Borrar inspección permanentemente"):
                        try:
                            supabase.table("inspecciones").delete().eq("id", insp["db_id"]).execute()
                            st.session_state.db_loaded = False
                            st.success("¡Inspección eliminada!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al eliminar: {e}")

                with col_i_info:
                    with st.expander(f"📌 {insp['Proyecto']} — {insp['Fecha']} ({insp['Dia']}) | Frente: {insp.get('Frente', 'N/A')}"):
                        st.markdown(f"**Residente:** {insp.get('Residente', '')} | **Clima:** {insp.get('Clima', '')}")
                        st.markdown(f"**Horario:** {insp.get('Hora_Inicio', '')} - {insp.get('Hora_Fin', '')}")
                        
                        st.markdown("##### Avance General de Actividades:")
                        df_av = pd.DataFrame(insp.get("Datos", {}).get("Avance", []))
                        if not df_av.empty:
                            st.dataframe(df_av, use_container_width=True)

                        # BOTÓN Y FORMULARIO DE EDICIÓN
                        with st.popover("✏️ Editar Inspección"):
                            st.markdown(f"#### Modificar Inspección — {insp['Proyecto']} ({insp['Fecha']})")
                            with st.form(f"form_edit_insp_{insp['db_id']}"):
                                e_frente = st.text_input("Frente Inspeccionado:", value=insp.get("Frente", ""))
                                e_clima = st.selectbox("Clima:", ["Soleado", "Nublado", "Lluvia"], index=["Soleado", "Nublado", "Lluvia"].index(insp.get("Clima", "Soleado")) if insp.get("Clima") in ["Soleado", "Nublado", "Lluvia"] else 0)
                                e_residente = st.text_input("Residente:", value=insp.get("Residente", ""))

                                st.markdown("##### Actualizar Avance General")
                                e_avance_datos = []
                                current_avance = insp.get("Datos", {}).get("Avance", [])
                                for item_av in current_avance:
                                    c_e1, c_e2, c_e3, c_e4 = st.columns([3, 2, 2, 2])
                                    act_n = item_av.get("Actividad", "")
                                    with c_e1:
                                        st.write(f"**{act_n}**")
                                    with c_e2:
                                        p_p = st.number_input(f"% Prog ({act_n})", min_value=0.0, max_value=100.0, value=float(item_av.get("% Prog", 0)), key=f"e_prog_{insp['db_id']}_{act_n}")
                                    with c_e3:
                                        p_e = st.number_input(f"% Ejec ({act_n})", min_value=0.0, max_value=100.0, value=float(item_av.get("% Ejec", 0)), key=f"e_ejec_{insp['db_id']}_{act_n}")
                                    with c_e4:
                                        e_st = st.selectbox(f"Estado ({act_n})", ["En Proceso", "Completado", "Retrasado", "N/A"], index=["En Proceso", "Completado", "Retrasado", "N/A"].index(item_av.get("Estado", "En Proceso")) if item_av.get("Estado") in ["En Proceso", "Completado", "Retrasado", "N/A"] else 0, key=f"e_est_{insp['db_id']}_{act_n}")
                                    e_avance_datos.append({"Actividad": act_n, "% Prog": p_p, "% Ejec": p_e, "Estado": e_st})

                                btn_save_edit = st.form_submit_button("💾 Guardar Cambios en Inspección", type="primary")

                                if btn_save_edit:
                                    updated_payload = insp.get("Datos", {})
                                    updated_payload["Avance"] = e_avance_datos

                                    try:
                                        supabase.table("inspecciones").update({
                                            "frente": e_frente,
                                            "clima": e_clima,
                                            "residente": e_residente,
                                            "datos": updated_payload
                                        }).eq("id", insp["db_id"]).execute()

                                        st.session_state.db_loaded = False
                                        st.success("¡Inspección actualizada correctamente!")
                                        st.rerun()
                                    except Exception as err:
                                        st.error(f"Error al actualizar inspección: {err}")

                        st.markdown("<br>", unsafe_allow_html=True)

                        c_dl_i1, c_dl_i2 = st.columns(2)
                        with c_dl_i1:
                            excel_insp_bytes = export_inspeccion_to_excel_file(insp)
                            st.download_button(
                                label="📊 Descargar Excel (.xlsx)",
                                data=excel_insp_bytes,
                                file_name=f"Inspeccion_{insp['Proyecto'].replace(' ', '_')}_{insp['Fecha']}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_insp_xlsx_{idx_insp}"
                            )
                        with c_dl_i2:
                            pdf_insp_bytes = export_inspeccion_to_pdf_file(insp)
                            st.download_button(
                                label="📄 Descargar PDF (.pdf)",
                                data=pdf_insp_bytes,
                                file_name=f"Inspeccion_{insp['Proyecto'].replace(' ', '_')}_{insp['Fecha']}.pdf",
                                mime="application/pdf",
                                key=f"dl_insp_pdf_{idx_insp}"
                            )
        else:
            st.warning("No hay formatos de inspección para el proyecto seleccionado.")
    else:
        st.info("Aún no hay formatos de inspección guardados.")

# ==========================================
# 8. MÓDULO 3: CONTROL DE RENDIMIENTO
# ==========================================
with tab_rend:
    st.markdown("### Control de Rendimiento por Trabajador")
    st.caption("Asignación de rubros, cálculo de Horas-Hombre (HH) y diagnóstico de productividad.")

    col1, col2 = st.columns(2)
    with col1:
        nombres_obreros = [t["nombre"] for t in st.session_state.db_trabajadores]
        if len(nombres_obreros) > 0:
            trabajador_sel = st.selectbox(f"Seleccionar Trabajador ({len(nombres_obreros)} Activos):", nombres_obreros)
            cargo_actual = next((t["cargo"] for t in st.session_state.db_trabajadores if t["nombre"] == trabajador_sel), "OBRERO")
            st.info(f"**Cargo en obra:** {cargo_actual}")
        else:
            trabajador_sel = None
            cargo_actual = "OBRERO"
            st.warning("No hay obreros registrados. Agregue obreros en la plantilla superior.")

    with col2:
        rubros_opciones = ["Enlucidos", "Fijos", "Fajas", "Dinteles"]
        rubro_sel = st.selectbox("Seleccionar Rubro:", rubros_opciones)
        unidad_medida = UNIDADES_RUBRO[rubro_sel]
        st.caption(f"Unidad de medida: **{unidad_medida}**")

    st.markdown("---")
    st.markdown("#### Horarios Trabajados (Intervalos)")

    col_h1, col_h2, col_h3 = st.columns(3)
    with col_h1:
        h1 = st.checkbox("07:00 - 10:00 (3 Horas HH)")
    with col_h2:
        h2 = st.checkbox("10:00 - 13:00 (3 Horas HH)")
    with col_h3:
        h3 = st.checkbox("14:00 - 16:00 (2 Horas HH)")

    horas_acumuladas = (3.0 if h1 else 0.0) + (3.0 if h2 else 0.0) + (2.0 if h3 else 0.0)
    st.markdown(f"**Total Horas-Hombre:** `{horas_acumuladas} HH`")

    st.markdown("#### Avance Ejecutado")
    avance_cant = st.number_input(f"Cantidad ejecutada ({unidad_medida}):", min_value=0.0, step=0.1, format="%.2f")

    if st.button("Registrar Rendimiento", type="primary"):
        if not trabajador_sel:
            st.error("Debe seleccionar un trabajador.")
        elif horas_acumuladas == 0:
            st.warning("Seleccione al menos un intervalo de horario.")
        elif avance_cant <= 0:
            st.warning("Ingrese un avance mayor a 0.")
        else:
            rend_real = round(horas_acumuladas / avance_cant, 3)
            rend_teorico = RENDIMIENTOS_TEORICOS.get(rubro_sel, 1.0)
            estado_diag = "EFICIENTE" if rend_real <= rend_teorico else "EXCESO DE HH"

            try:
                supabase.table("rendimientos").insert({
                    "usuario_email": user_email,
                    "cargo_obrero": cargo_actual,
                    "fecha": datetime.date.today().strftime("%Y-%m-%d"),
                    "trabajador": trabajador_sel,
                    "rubro": rubro_sel,
                    "horas_hh": horas_acumuladas,
                    "avance": avance_cant,
                    "unidad": unidad_medida,
                    "rend_real": rend_real,
                    "rend_teorico": rend_teorico,
                    "estado": estado_diag
                }).execute()

                st.session_state.db_loaded = False
                st.success(f"Rendimiento registrado correctamente en Supabase para {trabajador_sel}.")
                st.rerun()
            except Exception as e:
                st.error(f"Error registrando rendimiento: {e}")

    st.markdown("---")
    st.markdown("### Registros de Rendimiento Guardados")

    mis_rendimientos = st.session_state.db_rendimientos.get(user_email, [])
    if len(mis_rendimientos) > 0:
        df_mis_r = pd.DataFrame(mis_rendimientos)
        df_display = df_mis_r.drop(columns=["db_id", "Usuario_Registro", "Cargo_Registrador"], errors="ignore")
        if not df_display.empty:
            df_display.index = range(1, len(df_display) + 1)
        st.dataframe(df_display, use_container_width=True)

        csv_bytes_r = export_dataframe_to_excel_csv(df_display)
        st.download_button(label="📥 Descargar Rendimientos en CSV (Excel)", data=csv_bytes_r, file_name=f"Rendimientos_{user_email}.csv", mime="text/csv")
    else:
        st.info("Aún no existen registros en su historial.")

# ==========================================
# 9. MÓDULO ADMINISTRADOR
# ==========================================
if es_admin:
    tab_admin = tabs_app[3]
    with tab_admin:
        st.markdown("### Panel de Control Administrador")
        st.caption("Módulo exclusivo para supervisar inspecciones, rendimientos, usuarios y configuración.")

        st.markdown("#### 🔐 Código de Seguridad de Acceso y Registro (PIN)")
        col_pin1, col_pin2 = st.columns([2, 1])

        with col_pin1:
            pin_actual = st.session_state.get("access_pin", "1254")
            with st.form("form_pin_clean"):
                nuevo_pin_input = st.text_input("Nuevo Código PIN (4 dígitos):", value=pin_actual, max_chars=4, type="password", help="Código requerido para iniciar sesión y registrar cuentas nuevas.")
                btn_pin_save = st.form_submit_button("Guardar Nuevo Código PIN", type="primary")

            if btn_pin_save:
                if len(nuevo_pin_input.strip()) == 4 and nuevo_pin_input.strip().isdigit():
                    try:
                        supabase.table("app_config").update({"value": nuevo_pin_input.strip()}).eq("key", "access_pin").execute()
                        st.session_state.access_pin = nuevo_pin_input.strip()
                        st.success(f"¡Código PIN de acceso actualizado con éxito en Supabase a: **{nuevo_pin_input.strip()}**!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al actualizar PIN: {e}")
                else:
                    st.error("El PIN debe constar exactamente de 4 números.")

        with col_pin2:
            st.info(f"**PIN Actual Configurado:** `{st.session_state.get('access_pin', '1254')}`")

        st.markdown("---")

        st.markdown("#### 👁️ Checklists Subidos por Todos los Participantes")
        todas_las_jornadas_admin = []
        for u_mail, j_lista in st.session_state.db_checklists.items():
            for j_item in j_lista:
                j_copy = j_item.copy()
                j_copy["Usuario_Correo"] = u_mail
                todas_las_jornadas_admin.append(j_copy)

        if len(todas_las_jornadas_admin) > 0:
            col_adm_f1, col_adm_f2 = st.columns(2)
            with col_adm_f1:
                filtro_edif_admin = st.selectbox("Filtrar por edificio / proyecto:", ["-- Todos los Edificios --"] + EDIFICIOS_ALPHA, key="admin_filter_edif_chk")
            with col_adm_f2:
                usuarios_lista_chk = ["-- Todos los Usuarios --"] + sorted(list(set([j["Usuario_Correo"] for j in todas_las_jornadas_admin])))
                filtro_usr_chk = st.selectbox("Filtrar por participante:", usuarios_lista_chk, key="admin_filter_usr_chk")

            jornadas_admin_filtradas = todas_las_jornadas_admin.copy()
            if filtro_edif_admin != "-- Todos los Edificios --":
                jornadas_admin_filtradas = [j for j in jornadas_admin_filtradas if j.get("Edificio") == filtro_edif_admin]
            if filtro_usr_chk != "-- Todos los Usuarios --":
                jornadas_admin_filtradas = [j for j in jornadas_admin_filtradas if j["Usuario_Correo"] == filtro_usr_chk]

            jornadas_admin_filtradas.sort(key=lambda x: x['Fecha'], reverse=True)

            for idx_adm, j_adm in enumerate(jornadas_admin_filtradas, 1):
                resp_str = j_adm.get("Responsable", "") or j_adm["Usuario_Correo"]
                with st.expander(f"📌 [{j_adm.get('Edificio', 'N/A')}] {j_adm['Fecha']} — {resp_str} ({j_adm['Usuario_Correo']})"):
                    st.markdown(f"**Usuario:** `{j_adm['Usuario_Correo']}` | **Responsable:** {resp_str}")
                    
                    df_data_adm = pd.DataFrame(j_adm["Datos"])
                    for _, r_adm in df_data_adm.iterrows():
                        badge_adm = render_estado_badge(r_adm['Estado'])
                        st.markdown(f"- **[{r_adm['Jornada']}] N° {r_adm['N°']}. {r_adm['Actividad']}**: {badge_adm}", unsafe_allow_html=True)

                    c_ad_dl1, c_ad_dl2 = st.columns(2)
                    with c_ad_dl1:
                        excel_bytes_adm = export_checklist_to_excel_file(j_adm)
                        st.download_button(
                            label=f"📊 Descargar Excel",
                            data=excel_bytes_adm,
                            file_name=f"Checklist_{j_adm['Usuario_Correo']}_{j_adm['Edificio']}_{j_adm['Fecha']}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_xlsx_adm_{idx_adm}"
                        )
                    with c_ad_dl2:
                        pdf_bytes_adm = export_checklist_to_pdf_file(j_adm)
                        st.download_button(
                            label=f"📄 Descargar PDF",
                            data=pdf_bytes_adm,
                            file_name=f"Checklist_{j_adm['Usuario_Correo']}_{j_adm['Edificio']}_{j_adm['Fecha']}.pdf",
                            mime="application/pdf",
                            key=f"dl_pdf_adm_{idx_adm}"
                        )
        else:
            st.info("Ningún participante ha registrado checklists aún.")

        st.markdown("---")

        st.markdown("#### 📑 Formatos de Inspección Subidos por Todos los Participantes")
        todas_las_inspecciones_admin = []
        for u_mail, i_lista in st.session_state.db_inspecciones.items():
            for i_item in i_lista:
                i_copy = i_item.copy()
                i_copy["Usuario_Correo"] = u_mail
                todas_las_inspecciones_admin.append(i_copy)

        if len(todas_las_inspecciones_admin) > 0:
            col_adm_i1, col_adm_i2 = st.columns(2)
            with col_adm_i1:
                filtro_edif_insp_adm = st.selectbox("Filtrar inspecciones por edificio:", ["-- Todos los Edificios --"] + EDIFICIOS_ALPHA, key="admin_filter_edif_insp")
            with col_adm_i2:
                usuarios_lista_insp = ["-- Todos los Usuarios --"] + sorted(list(set([i["Usuario_Correo"] for i in todas_las_inspecciones_admin])))
                filtro_usr_insp_adm = st.selectbox("Filtrar inspecciones por usuario:", usuarios_lista_insp, key="admin_filter_usr_insp")

            insps_admin_filtradas = todas_las_inspecciones_admin.copy()
            if filtro_edif_insp_adm != "-- Todos los Edificios --":
                insps_admin_filtradas = [i for i in insps_admin_filtradas if i.get("Proyecto") == filtro_edif_insp_adm]
            if filtro_usr_insp_adm != "-- Todos los Usuarios --":
                insps_admin_filtradas = [i for i in insps_admin_filtradas if i["Usuario_Correo"] == filtro_usr_insp_adm]

            insps_admin_filtradas.sort(key=lambda x: x['Fecha'], reverse=True)

            for idx_i_adm, i_adm in enumerate(insps_admin_filtradas, 1):
                with st.expander(f"📌 [{i_adm.get('Proyecto', 'N/A')}] {i_adm['Fecha']} ({i_adm['Dia']}) — Usuario: {i_adm['Usuario_Correo']}"):
                    st.markdown(f"**Residente:** {i_adm.get('Residente', '')} | **Clima:** {i_adm.get('Clima', '')}")
                    
                    c_ad_idl1, c_ad_idl2 = st.columns(2)
                    with c_ad_idl1:
                        excel_insp_bytes_adm = export_inspeccion_to_excel_file(i_adm)
                        st.download_button(
                            label="📊 Descargar Excel (.xlsx)",
                            data=excel_insp_bytes_adm,
                            file_name=f"Inspeccion_{i_adm['Proyecto'].replace(' ', '_')}_{i_adm['Fecha']}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_insp_xlsx_adm_{idx_i_adm}"
                        )
                    with c_ad_idl2:
                        pdf_insp_bytes_adm = export_inspeccion_to_pdf_file(i_adm)
                        st.download_button(
                            label="📄 Descargar PDF (.pdf)",
                            data=pdf_insp_bytes_adm,
                            file_name=f"Inspeccion_{i_adm['Proyecto'].replace(' ', '_')}_{i_adm['Fecha']}.pdf",
                            mime="application/pdf",
                            key=f"dl_insp_pdf_adm_{idx_i_adm}"
                        )
        else:
            st.info("Ningún participante ha registrado formatos de inspección aún.")

        st.markdown("---")

        st.markdown("#### 📊 Rendimientos Subidos por Todos los Participantes")
        todos_los_rendimientos = []
        for u_mail, r_lista in st.session_state.db_rendimientos.items():
            for r_item in r_lista:
                r_copy = r_item.copy()
                r_copy["Usuario_Correo"] = u_mail
                todos_los_rendimientos.append(r_copy)

        if len(todos_los_rendimientos) > 0:
            df_rend_admin = pd.DataFrame(todos_los_rendimientos)
            cols_first = ["Usuario_Correo", "Fecha", "Trabajador", "Rubro", "Horas Trabajadas (HH)", "Avance", "Unidad", "Rend. Real (HH/Unid)", "Rend. Teórico", "Estado"]
            df_rend_admin = df_rend_admin.reindex(columns=[c for c in cols_first if c in df_rend_admin.columns])
            if not df_rend_admin.empty:
                df_rend_admin.index = range(1, len(df_rend_admin) + 1)

            st.dataframe(df_rend_admin, use_container_width=True)

            csv_rend_admin_bytes = export_dataframe_to_excel_csv(df_rend_admin)
            st.download_button(
                label="📥 Descargar Todos los Rendimientos (Excel CSV)",
                data=csv_rend_admin_bytes,
                file_name=f"Rendimientos_Globales_{datetime.date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        else:
            st.info("Ningún participante ha registrado rendimientos aún.")

        st.markdown("---")

        st.markdown("#### Gestión de Administradores de la Plataforma")
        col_adm1, col_adm2 = st.columns([2, 1])

        with col_adm1:
            with st.form("form_admin_add_clean"):
                nuevo_admin_mail = st.text_input("Ingrese correo para conceder permisos de Administrador:", placeholder="usuario@correo.com")
                btn_admin_add = st.form_submit_button("Otorgar Acceso Administrador")

            if btn_admin_add:
                if nuevo_admin_mail:
                    mail_clean = nuevo_admin_mail.strip().lower()
                    try:
                        supabase.table("usuarios").update({"es_admin": True}).eq("correo", mail_clean).execute()
                        st.session_state.db_loaded = False
                        st.success(f"Se otorgaron permisos de administrador en Supabase a: {mail_clean}")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error actualizando administrador: {e}")

        with col_adm2:
            st.markdown("**Administradores Actuales:**")
            for adm in st.session_state.admin_emails:
                st.write(f"- `{adm}`")

        st.markdown("---")

        st.markdown("#### Usuarios Activos y Gestión de Cuentas")
        
        lista_correos = [u["Correo"] for u in st.session_state.db_usuarios]
        
        col_del_usr1, col_del_usr2 = st.columns([2, 1])
        with col_del_usr1:
            usuario_a_eliminar = st.selectbox("Seleccionar cuenta de usuario a eliminar:", lista_correos, key="sel_user_del")
        with col_del_usr2:
            st.write("") 
            st.write("")
            if st.button("🗑️ Eliminar Cuenta Seleccionada", type="secondary"):
                if usuario_a_eliminar == user_email:
                    st.error("No puedes eliminar la cuenta activa con la que estás con sesión iniciada.")
                else:
                    try:
                        supabase.table("usuarios").delete().eq("correo", usuario_a_eliminar).execute()
                        st.session_state.db_loaded = False
                        st.success(f"Cuenta de usuario **{usuario_a_eliminar}** eliminada correctamente de Supabase.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al eliminar usuario: {e}")

        st.markdown("<br>", unsafe_allow_html=True)

        db_usuarios_privados = []
        for u in st.session_state.db_usuarios:
            u_copy = u.copy()
            u_copy["Password"] = "••••••••"
            db_usuarios_privados.append(u_copy)

        df_users = pd.DataFrame(db_usuarios_privados)
        if not df_users.empty:
            df_users.index = range(1, len(df_users) + 1)
        st.dataframe(df_users, use_container_width=True)

        st.markdown("#### Resumen Global de Actividad por Usuario")
        resumen_actividad = []
        for u in st.session_state.db_usuarios:
            e = u["Correo"]
            num_c = len(st.session_state.db_checklists.get(e, []))
            num_i = len(st.session_state.db_inspecciones.get(e, []))
            num_r = len(st.session_state.db_rendimientos.get(e, []))
            resumen_actividad.append({
                "Usuario": f"{u['Nombres']} {u['Apellidos']}".strip() or e,
                "Correo": e,
                "Cargo": u["Cargo"],
                "Checklists": num_c,
                "Inspecciones": num_i,
                "Rendimientos": num_r,
                "Estado": u["Estado"]
            })

        df_act = pd.DataFrame(resumen_actividad)
        if not df_act.empty:
            df_act.index = range(1, len(df_act) + 1)
        st.dataframe(df_act, use_container_width=True)

        csv_admin_bytes = export_dataframe_to_excel_csv(df_act)
        st.download_button(
            label="📥 Descargar Reporte de Usuarios (Excel CSV)",
            data=csv_admin_bytes,
            file_name=f"Reporte_Usuarios_AlphaBuilders_{datetime.date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )