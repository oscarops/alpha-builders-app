import base64
import datetime
import io
import json
import os
import pandas as pd
from PIL import Image
import streamlit as st

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS TRÍCROMAS
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
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@500;700;800;900&family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    h1, h2, h3, .brand-title {
        font-family: 'Montserrat', sans-serif !important;
        letter-spacing: -0.03em !important;
    }

    /* AJUSTE SUPERIOR DE PÁGINA QUE SE ADAPTA AL 100% DE PANTALLA COMPLETA */
    .block-container {
        padding-top: 0.8rem !important;
        padding-bottom: 1.2rem !important;
        max-width: 100% !important;
        width: 100% !important;
    }

    /* 1. FONDO PRINCIPAL: BLANCO PURO */
    .stApp {
        background-color: #ffffff !important;
        color: #121318 !important;
        transition: all 0.3s ease;
    }

    label, p, span, div, h1, h2, h3, h4, h5, h6, .stMarkdown {
        color: #121318 !important;
        font-weight: 500;
    }

    .stCaption, caption, small, [data-testid="stCaptionContainer"] {
        color: #5a5f6e !important;
    }

    /* BOTÓN PARA COLAPSAR SIDEBAR SIEMPRE VISIBLE */
    [data-testid="stSidebarCollapseButton"], 
    [data-testid="collapsedControl"] {
        display: block !important;
        visibility: visible !important;
        opacity: 1 !important;
        z-index: 999999 !important;
    }

    [data-testid="stSidebarCollapseButton"] button, 
    [data-testid="collapsedControl"] button {
        background-color: #1c1e26 !important;
        border: 1px solid #323646 !important;
        border-radius: 50% !important;
        width: 32px !important;
        height: 32px !important;
        color: #ffffff !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
        transition: all 0.2s ease !important;
    }

    [data-testid="stSidebarCollapseButton"] button:hover, 
    [data-testid="collapsedControl"] button:hover {
        background-color: #ff8c00 !important;
        border-color: #ff8c00 !important;
        transform: scale(1.08);
    }

    [data-testid="stSidebarCollapseButton"] svg, 
    [data-testid="collapsedControl"] svg {
        fill: #ffffff !important;
        color: #ffffff !important;
    }

    /* 2. BARRA LATERAL (SIDEBAR): FIJA, SIN SCROLL Y CON ESPACIO CLARO */
    [data-testid="stSidebar"] {
        background-color: #121318 !important;
        border-right: 2px solid #282a36 !important;
        padding-top: 0px !important;
        padding-left: 12px !important;
        padding-right: 12px !important;
        padding-bottom: 15px !important;
        width: 250px !important;
        min-width: 250px !important;
        max-width: 250px !important;
        resize: none !important;
    }

    [data-testid="stSidebarResizer"] {
        display: none !important;
        pointer-events: none !important;
    }

    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
        gap: 0.5rem !important;
        padding-top: 0px !important;
    }

    [data-testid="stSidebar"] label, 
    [data-testid="stSidebar"] p, 
    [data-testid="stSidebar"] span, 
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3, 
    [data-testid="stSidebar"] div {
        color: #ffffff !important;
    }

    /* TARJETA DEL LOGO ARRIBA CON MARGEN INFERIOR GRANDE */
    .sidebar-logo-card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 8px 10px;
        margin-top: 0px !important;
        margin-bottom: 24px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        width: 100% !important;
        box-sizing: border-box;
        text-align: center;
        display: block;
    }

    /* FOTO DE PERFIL DEL MISMO ANCHO EXACTO Y DESPLAZADA ABAJO */
    [data-testid="stSidebar"] [data-testid="stImage"] {
        width: 100% !important;
        display: block !important;
        margin-top: 10px !important;
        margin-bottom: 12px !important;
        clear: both !important;
    }

    [data-testid="stSidebar"] [data-testid="stImage"] > div {
        width: 100% !important;
    }

    [data-testid="stSidebar"] [data-testid="stImage"] img {
        border-radius: 12px !important;
        width: 100% !important;
        height: auto !important;
        max-width: 100% !important;
        object-fit: cover !important;
        border: 1px solid #323646 !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4);
        margin: 0 !important;
        display: block !important;
    }

    /* TARJETA DE INFORMACIÓN DE PERFIL */
    .sidebar-profile-box {
        background: #1c1e26;
        border: 1px solid #323646;
        border-radius: 12px;
        padding: 10px 8px !important;
        text-align: center;
        margin-top: 4px;
        margin-bottom: 8px;
        width: 100% !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.3);
        box-sizing: border-box;
    }

    .sidebar-user-nombres {
        font-size: 0.88rem;
        font-weight: 800;
        color: #ffffff !important;
        margin-top: 1px;
        margin-bottom: 1px !important;
        line-height: 1.2;
    }

    .sidebar-user-apellidos {
        font-size: 0.85rem;
        font-weight: 700;
        color: #e0e4ed !important;
        margin-bottom: 4px !important;
        line-height: 1.2;
    }

    .sidebar-user-email {
        font-size: 0.68rem;
        color: #72b2ff !important;
        font-weight: 600;
        margin-bottom: 6px !important;
        word-break: break-all;
    }

    .sidebar-user-cargo {
        display: inline-block;
        background: #323646 !important;
        color: #ffffff !important;
        border: 1px solid #484e5e !important;
        font-size: 0.60rem !important;
        font-weight: 800 !important;
        padding: 2px 8px !important;
        border-radius: 14px !important;
        text-transform: uppercase !important;
    }

    [data-testid="stSidebar"] hr {
        margin: 6px 0 !important;
        border-color: #282a36 !important;
    }

    /* EXPANDER DE CONFIGURACIÓN */
    [data-testid="stSidebar"] [data-testid="stExpander"] {
        background-color: #1c1e26 !important;
        border: 1px solid #323646 !important;
        border-radius: 10px !important;
        margin-top: 2px !important;
        margin-bottom: 6px !important;
    }

    [data-testid="stSidebar"] [data-testid="stExpander"] summary {
        background-color: #282c36 !important;
        padding: 6px 8px !important;
    }

    [data-testid="stSidebar"] [data-testid="stExpander"] summary * {
        color: #ffffff !important;
        font-weight: 700 !important;
        font-size: 0.78rem !important;
    }

    [data-testid="stSidebar"] .stButton > button {
        padding: 6px 12px !important;
        font-size: 0.78rem !important;
        margin-top: 2px !important;
    }

    /* 3. TARJETA PRINCIPAL CON BORDE NEGRO Y LETRA DISTINTIVA */
    .executive-card-studio {
        background: linear-gradient(145deg, #f3f6fc 0%, #e8edf7 100%);
        border: 1px solid #b8c4d8;
        border-left: 7px solid #121318;
        border-radius: 22px;
        padding: 22px 28px;
        box-shadow: 0 12px 35px rgba(0,0,0,0.06);
        margin-bottom: 20px;
        transition: all 0.3s ease;
    }
    .executive-card-studio:hover {
        transform: translateY(-2px);
        box-shadow: 0 18px 45px rgba(0,0,0,0.12);
        border-color: #121318;
        border-left-color: #121318;
    }

    .brand-title {
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 900 !important;
        font-size: 2.4rem !important;
        background: linear-gradient(90deg, #121318 0%, #3a4256 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-shadow: 0 2px 12px rgba(0,0,0,0.08);
        letter-spacing: -0.04em !important;
    }

    .kpi-card-studio {
        background: linear-gradient(145deg, #eceff6 0%, #dbe2ef 100%);
        border: 1px solid #aebacf;
        border-radius: 20px;
        padding: 18px;
        text-align: center;
        box-shadow: 0 8px 25px rgba(0,0,0,0.06);
        transition: all 0.35s cubic-bezier(0.16, 1, 0.3, 1);
    }
    .kpi-card-studio:hover {
        background: linear-gradient(145deg, #e3eaf7, #d0dbee);
        border-color: #121318;
        transform: translateY(-5px) scale(1.02);
        box-shadow: 0 16px 35px rgba(0,0,0,0.12);
    }
    .kpi-val-studio {
        font-size: 2.5rem;
        font-weight: 900;
        color: #121318 !important;
        letter-spacing: -0.03em;
    }
    .kpi-lbl-studio {
        font-size: 0.72rem;
        color: #4a5060 !important;
        text-transform: uppercase;
        font-weight: 800;
        letter-spacing: 0.08em;
        margin-top: 2px;
    }

    /* PESTAÑAS (SEGMENT CONTROL) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: #e2e5ec !important;
        padding: 6px;
        border-radius: 16px;
        border: 1px solid #c2c7d2;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 12px !important;
        padding: 10px 24px !important;
        background-color: transparent !important;
        border: none !important;
        transition: all 0.2s ease;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: rgba(0,0,0,0.04);
    }
    .stTabs [data-baseweb="tab"] p, 
    .stTabs [data-baseweb="tab"] span {
        color: #4a5060 !important;
        font-weight: 700 !important;
    }

    .stTabs [aria-selected="true"] {
        background-color: #121318 !important;
        border-radius: 12px !important;
        box-shadow: 0 6px 18px rgba(0,0,0,0.3) !important;
    }
    .stTabs [aria-selected="true"] p, 
    .stTabs [aria-selected="true"] span,
    .stTabs [aria-selected="true"] div {
        color: #ffffff !important;
        font-weight: 900 !important;
    }

    /* BOTONES PRIMARIOS NEGROS */
    .stButton > button {
        background-color: #121318 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        font-weight: 800 !important;
        padding: 10px 22px !important;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.25) !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    .stButton > button p, .stButton > button span {
        color: #ffffff !important;
    }
    .stButton > button:hover {
        background-color: #2c303d !important;
        transform: translateY(-2px);
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.35) !important;
    }

    .streamlit-expanderHeader {
        background-color: #e8eaee !important;
        border-radius: 12px !important;
        border: 1px solid #c2c7d2 !important;
        font-weight: 700 !important;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. PERSISTENCIA EN DISCO (LOCAL_DB.JSON Y REPO)
# ==========================================
DB_FILE = "local_db.json"

def get_repo_image_b64(filenames):
    for filename in filenames:
        if os.path.exists(filename):
            try:
                with open(filename, "rb") as f:
                    return base64.b64encode(f.read()).decode("utf-8")
            except Exception:
                pass
    return None

def load_persistent_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "admin_emails": ["oscarsebitas2013@gmail.com"],
        "db_fotos_perfil_b64": {},
        "db_usuarios": [
            {
                "Nombres": "Oscar Sebastián",
                "Apellidos": "Narváez Ojeda",
                "Correo": "oscarsebitas2013@gmail.com",
                "Password": "Al678554",
                "Cargo": "Residente",
                "Fecha_Registro": "2026-07-26",
                "Estado": "Activo",
            }
        ],
        "db_checklists": {},
        "db_rendimientos": {},
    }

def save_persistent_db():
    data_to_save = {
        "admin_emails": st.session_state.get("admin_emails", ["oscarsebitas2013@gmail.com"]),
        "db_fotos_perfil_b64": st.session_state.get("db_fotos_perfil_b64", {}),
        "db_usuarios": st.session_state.get("db_usuarios", []),
        "db_checklists": st.session_state.get("db_checklists", {}),
        "db_rendimientos": st.session_state.get("db_rendimientos", {}),
    }
    try:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(data_to_save, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

if "db_loaded" not in st.session_state:
    p_data = load_persistent_db()
    st.session_state.admin_emails = p_data.get("admin_emails", ["oscarsebitas2013@gmail.com"])
    st.session_state.db_fotos_perfil_b64 = p_data.get("db_fotos_perfil_b64", {})
    st.session_state.db_usuarios = p_data.get("db_usuarios", [])
    st.session_state.db_checklists = p_data.get("db_checklists", {})
    st.session_state.db_rendimientos = p_data.get("db_rendimientos", {})
    st.session_state.db_loaded = True

def image_to_base64(image_file):
    if image_file is not None:
        try:
            img = Image.open(image_file)
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
            return Image.open(io.BytesIO(img_data))
        except Exception:
            return None
    return None

def export_dataframe_to_excel_csv(df):
    df_clean = df.drop(columns=["Foto_B64"], errors="ignore")
    return df_clean.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")

# ==========================================
# 3. BASE DE DATOS Y ESTADOS DE SESIÓN
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_email = ""
    st.session_state.usuario_nombres = ""
    st.session_state.usuario_apellidos = ""
    st.session_state.usuario_cargo = ""

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
]

TRABAJADORES_NO_MINA = [
    {"nombre": "ACHINA AGUAGUIÑA BYRON ALEXANDER", "cargo": "BODEGA"},
    {"nombre": "AGUALONGO PILAMUNGA LUIS LENIN", "cargo": "GYPSERO/ALBAÑIL"},
    {"nombre": "ALTAMIRANO GUALAN WILLIAM PATRICIO", "cargo": "GYPSERO"},
    {"nombre": "BUNSHI CAYANCELA SANTIAGO EFRAIN", "cargo": "ALBAÑIL"},
    {"nombre": "CAYAMBE SANDOVAL LUIS ANTONIO", "cargo": "ALBAÑIL"},
    {"nombre": "CUASCOTA INLAGO JOSE LIZARDO", "cargo": "ALBAÑIL"},
    {"nombre": "CUERO BAMONTES DEIBINZON ESTALIN", "cargo": "AYUDANTE"},
    {"nombre": "GUANOLUISA VACA LUIS FERNANDO", "cargo": "ALBAÑIL"},
    {"nombre": "LLUGLLUNA FARINANGO SEGUNDO MANUEL", "cargo": "ALBAÑIL"},
    {"nombre": "MORALES OTUNA VERONICA JAQUELINE", "cargo": "AYUDANTE"},
    {"nombre": "OCHOA MORAN MIGUEL BERNARDO", "cargo": "GYPSERO"},
    {"nombre": "PAGUAY RAMOS DILAN ANDRES", "cargo": "GYPSERO"},
    {"nombre": "ROMERO ANDRANGO LUIS ENRIQUE", "cargo": "GYPSERO"},
    {"nombre": "SANGUCHO FONSECA EDGAR XAVIER", "cargo": "ALBAÑIL"},
    {"nombre": "TARAPUES MONARCO CARLOS ANDRES", "cargo": "GYPSERO"},
    {"nombre": "TONATO TACO LUIS EUCLIDES", "cargo": "ALBAÑIL"},
    {"nombre": "TOSCANO ALTAMIRANO JEREMMY WENDLEY", "cargo": "AYUDANTE"},
    {"nombre": "TRONCOSO COBEÑA CRISTOPHER GEOVANNY", "cargo": "AYUDANTE"},
    {"nombre": "TUTASI CASILLAS JORGE GEOVANI", "cargo": "FIERRERO"},
    {"nombre": "CHAVEZ GUITARRA JOSE GREGORIO", "cargo": "GYPSERO"},
    {"nombre": "CORDOVA FLORES ERICK DARIO", "cargo": "GYPSERO / AYUDANTE"},
    {"nombre": "CABRERA CAMPO ANNDY JEREMIAS", "cargo": "GYPSERO / OPERADOR"},
    {"nombre": "CHELA OCHOA RAUL", "cargo": "GYPSERO/ALBAÑIL"},
    {"nombre": "SEMBLANTES TIPANLUISA JAVIER PATRICIO", "cargo": "GYPSERO/ALBAÑIL"},
    {"nombre": "FUEREZ COYAGO JOSE SANTOS", "cargo": "HERRAMIENTAS"},
    {"nombre": "ALTAMIRANO CORDOVA HECTOR LUIS", "cargo": "PINTOR"},
    {"nombre": "ACOSTA AGUILAR JORGE PATRICIO", "cargo": "SOLDADOR"},
    {"nombre": "TARAPUES CASTRO JOAO ALEXANDER", "cargo": "SOLDADOR"},
]

UNIDADES_RUBRO = {"Enlucidos": "m2", "Fijos": "m2", "Fajas": "m", "Dinteles": "m"}
RENDIMIENTOS_TEORICOS = {"Enlucidos": 0.75, "Fijos": 0.50, "Fajas": 0.30, "Dinteles": 0.40}

ACTIVIDADES_MANANA = [
    "Verificación de asistencia del personal",
    "Distribución de cuadrillas por frente de trabajo",
    "Recorrido inicial de obra",
    "Verificación de los frentes de trabajo",
    "Revisión del cumplimiento de planos y especificaciones",
    "Supervisión de la ejecución de los trabajos",
    "Verificación de calidad de los trabajos ejecutados",
    "Verificación de materiales disponibles en cada frente",
    "Coordinación con otras especialidades",
    "Corrección de observaciones detectadas",
]

ACTIVIDADES_TARDE = [
    "Recorrido de seguimiento de los frentes de trabajo",
    "Verificación del avance físico de las actividades",
    "Control del rendimiento de las cuadrillas",
    "Inspección de calidad de los trabajos ejecutados",
    "Revisión de observaciones pendientes",
    "Verificación de trabajos corregidos",
    "Verificación del orden y limpieza de los frentes de trabajo",
    "Confirmación de materiales para el siguiente día",
    "Revisión del cumplimiento de la meta diaria",
    "Cierre de actividades en campo",
]

# ==========================================
# 4. MÓDULO DE LOGIN & REGISTRO (CON images.png)
# ==========================================
if not st.session_state.autenticado:
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])

    with col_l2:
        if os.path.exists("images.png"):
            with open("images.png", "rb") as image_file:
                encoded_logo = base64.b64encode(image_file.read()).decode("utf-8")
            st.markdown(
                f"""
                <div style="text-align: center; margin-top: 10px; margin-bottom: 10px;">
                    <img src="data:image/png;base64,{encoded_logo}" style="width: 320px; max-width: 100%; pointer-events: none;">
                </div>
            """,
                unsafe_allow_html=True,
            )

        tab_login, tab_register, tab_reset = st.tabs(["Iniciar Sesión", "Registrarse", "¿Olvidaste tu Contraseña?"])

        with tab_login:
            st.markdown("### Iniciar Sesión")
            st.caption("Ingrese sus credenciales registradas.")

            login_email = st.text_input("Correo electrónico:", placeholder="nombre@correo.com", key="log_email")
            login_pass = st.text_input("Contraseña:", type="password", key="log_pass")

            if st.button("Entrar al Portal", type="primary", use_container_width=True):
                if login_email and login_pass:
                    mail_clean = login_email.strip().lower()
                    u_match = next((u for u in st.session_state.db_usuarios if u["Correo"] == mail_clean), None)

                    if u_match:
                        if u_match["Password"] == login_pass:
                            st.session_state.autenticado = True
                            st.session_state.usuario_email = mail_clean
                            st.session_state.usuario_nombres = u_match["Nombres"]
                            st.session_state.usuario_apellidos = u_match["Apellidos"]
                            st.session_state.usuario_cargo = u_match["Cargo"]

                            if mail_clean not in st.session_state.db_checklists:
                                st.session_state.db_checklists[mail_clean] = []
                            if mail_clean not in st.session_state.db_rendimientos:
                                st.session_state.db_rendimientos[mail_clean] = []
                            st.rerun()
                        else:
                            st.error("Contraseña incorrecta.")
                    else:
                        st.error("El usuario no existe. Complete el registro.")
                else:
                    st.error("Ingrese su correo y contraseña.")

        with tab_register:
            st.markdown("### Crear una Cuenta Nueva")
            st.caption("Complete la información para habilitar su acceso.")

            col_n, col_a = st.columns(2)
            with col_n:
                reg_nombres = st.text_input("Nombres:", placeholder="Ej. Juan Carlos")
            with col_a:
                reg_apellidos = st.text_input("Apellidos:", placeholder="Ej. Pérez Gómez")

            reg_email = st.text_input("Correo electrónico:", placeholder="ejemplo@correo.com", key="reg_email")
            
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                reg_pass = st.text_input("Crear contraseña:", type="password", key="reg_pass")
            with col_p2:
                reg_pass_repeat = st.text_input("Repetir contraseña:", type="password", key="reg_pass_rep")

            reg_cargo = st.selectbox("Cargo / Rol en Obra:", ["Residente", "Asistente", "Ayudante"])

            if st.button("Completar Registro", type="primary", use_container_width=True):
                if reg_nombres and reg_apellidos and reg_email and reg_pass and reg_pass_repeat:
                    if reg_pass != reg_pass_repeat:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        mail_clean = reg_email.strip().lower()
                        exists = any(u["Correo"] == mail_clean for u in st.session_state.db_usuarios)
                        if exists:
                            st.warning("Este correo ya se encuentra registrado.")
                        else:
                            st.session_state.autenticado = True
                            st.session_state.usuario_email = mail_clean
                            st.session_state.usuario_nombres = reg_nombres.strip()
                            st.session_state.usuario_apellidos = reg_apellidos.strip()
                            st.session_state.usuario_cargo = reg_cargo

                            st.session_state.db_usuarios.append({
                                "Nombres": reg_nombres.strip(),
                                "Apellidos": reg_apellidos.strip(),
                                "Correo": mail_clean,
                                "Password": reg_pass,
                                "Cargo": reg_cargo,
                                "Fecha_Registro": datetime.date.today().strftime("%Y-%m-%d"),
                                "Estado": "Activo"
                            })

                            if mail_clean not in st.session_state.db_checklists:
                                st.session_state.db_checklists[mail_clean] = []
                            if mail_clean not in st.session_state.db_rendimientos:
                                st.session_state.db_rendimientos[mail_clean] = []

                            save_persistent_db()
                            st.success("¡Registro completado exitosamente!")
                            st.rerun()
                else:
                    st.error("Por favor complete todos los campos requeridos.")

        with tab_reset:
            st.markdown("### Recuperación de Contraseña")
            st.caption("Restablezca su acceso de forma segura.")

            reset_email = st.text_input("Ingrese su correo registrado:", placeholder="ejemplo@correo.com", key="rst_email")
            
            col_rp1, col_rp2 = st.columns(2)
            with col_rp1:
                new_pass = st.text_input("Nueva contraseña:", type="password", key="rst_pass")
            with col_rp2:
                new_pass_rep = st.text_input("Repetir nueva contraseña:", type="password", key="rst_pass_rep")

            if st.button("Restablecer Contraseña", type="primary", use_container_width=True):
                if reset_email and new_pass and new_pass_rep:
                    if new_pass != new_pass_rep:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        mail_clean = reset_email.strip().lower()
                        u_match = next((u for u in st.session_state.db_usuarios if u["Correo"] == mail_clean), None)

                        if u_match:
                            u_match["Password"] = new_pass
                            save_persistent_db()
                            st.success(f"Contraseña actualizada con éxito para {mail_clean}.")
                        else:
                            st.error("El correo ingresado no está registrado.")
                else:
                    st.error("Complete todos los campos.")

    st.stop()

# ==========================================
# 5. BARRA LATERAL (LOGO MÁS ARRIBA CON SEPARACIÓN)
# ==========================================
user_email = st.session_state.usuario_email
user_nombres = st.session_state.usuario_nombres
user_apellidos = st.session_state.usuario_apellidos
user_cargo = st.session_state.usuario_cargo
es_admin = user_email in st.session_state.admin_emails

with st.sidebar:
    # LOGO EN RECUADRO CON FONDO BLANCO BIEN ARRIBA Y SEPARADO
    logo_filename = "alpha.473f0c2dc3c48a682723-2.webp"
    if not os.path.exists(logo_filename):
        logo_filename = "images.png"

    if os.path.exists(logo_filename):
        ext = "webp" if logo_filename.endswith(".webp") else "png"
        with open(logo_filename, "rb") as image_file:
            encoded_sidebar_logo = base64.b64encode(image_file.read()).decode("utf-8")
        st.markdown(
            f"""
            <div class="sidebar-logo-card">
                <img src="data:image/{ext};base64,{encoded_sidebar_logo}" style="width: 100%; max-width: 100%; pointer-events: none; display: block; margin: 0 auto;">
            </div>
        """,
            unsafe_allow_html=True,
        )

    b64_foto = st.session_state.db_fotos_perfil_b64.get(user_email, None)
    if not b64_foto:
        b64_foto = get_repo_image_b64(["perfil.jpg", "perfil.png", "perfil.jpeg", "avatar.png"])

    img_obj = base64_to_image(b64_foto)

    if img_obj is not None:
        st.image(img_obj, use_container_width=True)

    st.markdown(
        f"""
        <div class="sidebar-profile-box">
            <div class="sidebar-user-nombres">{user_nombres}</div>
            <div class="sidebar-user-apellidos">{user_apellidos}</div>
            <div class="sidebar-user-email">{user_email}</div>
            <div class="sidebar-user-cargo">{user_cargo}</div>
        </div>
    """,
        unsafe_allow_html=True,
    )

    if es_admin:
        st.markdown("<div style='text-align: center; margin-bottom: 4px; font-size: 0.65rem; color: #ffffff; font-weight: 800; background: #1c1e26; padding: 3px; border-radius: 6px; border: 1px solid #323646;'>ADMINISTRADOR GENERAL</div>", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    with st.expander("⚙️ Configuración de Cuenta", expanded=False):
        st.caption("Ajustes personales y fotografía:")
        
        edit_nombres = st.text_input("Nombres:", value=st.session_state.usuario_nombres, key="sb_nom")
        edit_apellidos = st.text_input("Apellidos:", value=st.session_state.usuario_apellidos, key="sb_ape")
        
        cargos_lista = ["Residente", "Asistente", "Ayudante"]
        idx_c = cargos_lista.index(user_cargo) if user_cargo in cargos_lista else 0
        edit_cargo = st.selectbox("Cargo:", cargos_lista, index=idx_c, key="sb_car")

        edit_pass = st.text_input("Nueva Contraseña:", type="password", key="sb_pass")
        edit_pass_rep = st.text_input("Repetir Contraseña:", type="password", key="sb_pass_rep")

        nueva_foto_file = st.file_uploader("Actualizar Foto de Perfil", type=["jpg", "jpeg", "png"], key="sb_foto_file")
        if nueva_foto_file is not None:
            b64_str = image_to_base64(nueva_foto_file)
            if b64_str:
                st.session_state.db_fotos_perfil_b64[user_email] = b64_str
                save_persistent_db()

        if st.button("Guardar Ajustes", type="primary", use_container_width=True):
            if edit_pass.strip() or edit_pass_rep.strip():
                if edit_pass != edit_pass_rep:
                    st.error("Las nuevas contraseñas no coinciden.")
                    st.stop()

            st.session_state.usuario_nombres = edit_nombres.strip()
            st.session_state.usuario_apellidos = edit_apellidos.strip()
            st.session_state.usuario_cargo = edit_cargo

            for u in st.session_state.db_usuarios:
                if u["Correo"] == user_email:
                    u["Nombres"] = edit_nombres.strip()
                    u["Apellidos"] = edit_apellidos.strip()
                    u["Cargo"] = edit_cargo
                    if edit_pass.strip():
                        u["Password"] = edit_pass.strip()

            save_persistent_db()
            st.success("Configuración actualizada correctamente.")
            st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

# ==========================================
# 6. DASHBOARD PRINCIPAL
# ==========================================
user_nombre_completo = f"{user_nombres} {user_apellidos}".strip()

st.markdown(
    f"""
    <div class="executive-card-studio">
        <h1 class="brand-title" style="font-size: 2.5rem; font-weight: 900; margin: 0;">Portal de Control e Inspección</h1>
        <p style="color: #5a5f6e; margin-top: 6px; font-size: 1.05rem;">{user_nombre_completo} — <b>{user_cargo}</b></p>
    </div>
""",
    unsafe_allow_html=True,
)

usr_chks = len(st.session_state.db_checklists.get(user_email, []))
usr_rnds = len(st.session_state.db_rendimientos.get(user_email, []))

k1, k2, k3 = st.columns(3)
with k1:
    st.markdown(
        '<div class="kpi-card-studio"><div class="kpi-val-studio">28</div><div class="kpi-lbl-studio">Obreros Activos</div></div>',
        unsafe_allow_html=True,
    )
with k2:
    st.markdown(
        f'<div class="kpi-card-studio"><div class="kpi-val-studio">{usr_chks}</div><div class="kpi-lbl-studio">Checklists Guardados</div></div>',
        unsafe_allow_html=True,
    )
with k3:
    st.markdown(
        f'<div class="kpi-card-studio"><div class="kpi-val-studio">{usr_rnds}</div><div class="kpi-lbl-studio">Reportes de Rendimiento</div></div>',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

pestanas = ["Checklist Diario", "Control de Rendimiento"]
if es_admin:
    pestanas.append("Panel Admin")

tabs_app = st.tabs(pestanas)
tab_chk = tabs_app[0]
tab_rend = tabs_app[1]

# ==========================================
# 7. MÓDULO 1: CHECKLIST DIARIO
# ==========================================
with tab_chk:
    st.markdown("### Check List Diario – Control de Obra")
    st.caption("Supervisión diaria de frentes de trabajo con verificación manual obligatoria.")

    if "creando_jornada" not in st.session_state:
        st.session_state.creando_jornada = False

    col_btn_j, _ = st.columns([2, 2])
    with col_btn_j:
        if st.button("➕ Crear Nueva Jornada de Inspección", type="primary", use_container_width=True):
            st.session_state.creando_jornada = True

    if st.session_state.creando_jornada:
        st.markdown("---")
        with st.container():
            st.markdown("#### Configuración de la Nueva Jornada")
            col_m1, col_m2, col_m3 = st.columns(3)

            with col_m1:
                edificio_val = st.selectbox("Edificio / Proyecto:", EDIFICIOS_ALPHA, key="sel_edificio")
            with col_m2:
                st.text_input("Responsable:", value=user_nombre_completo, disabled=True, help="Cargado de su inicio de sesión.")
            with col_m3:
                fecha_val = st.date_input("Fecha de Inspección:", datetime.date.today(), key="sel_fecha")

            st.markdown("---")

            with st.form("form_checklist_jornada"):
                st.markdown("#### 🌅 Jornada de la Mañana")
                resp_manana = []

                for idx, act in enumerate(ACTIVIDADES_MANANA, 1):
                    st.markdown(f"**N° {idx}. {act}**")
                    c_sel, c_obs, c_foto = st.columns([2, 3, 3])

                    with c_sel:
                        est = st.radio(
                            "Estado",
                            ["✓ Cumple", "✗ No Cumple", "N/A"],
                            index=None,
                            key=f"m_st_{idx}",
                            horizontal=True,
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
                        "Foto_B64": ft_b64,
                        "Foto_Adjunta": "Sí" if ft is not None else "No"
                    })

                st.markdown("#### 🌆 Jornada de la Tarde")
                resp_tarde = []

                for idx, act in enumerate(ACTIVIDADES_TARDE, 1):
                    st.markdown(f"**N° {idx}. {act}**")
                    c_sel, c_obs, c_foto = st.columns([2, 3, 3])

                    with c_sel:
                        est = st.radio(
                            "Estado",
                            ["✓ Cumple", "✗ No Cumple", "N/A"],
                            index=None,
                            key=f"t_st_{idx}",
                            horizontal=True,
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
                        "Foto_B64": ft_b64,
                        "Foto_Adjunta": "Sí" if ft is not None else "No"
                    })

                btn_guardar_chk = st.form_submit_button("Guardar Jornada de Inspección", type="primary")

            if btn_guardar_chk:
                all_chk_data = resp_manana + resp_tarde
                
                sin_responder = [item["Actividad"] for item in all_chk_data if item["Estado"] is None]
                if sin_responder:
                    st.error(f"⚠️ Seleccione el estado de todas las actividades ({len(sin_responder)} pendientes).")
                else:
                    df_chk_save = pd.DataFrame(all_chk_data)

                    if user_email not in st.session_state.db_checklists:
                        st.session_state.db_checklists[user_email] = []

                    st.session_state.db_checklists[user_email].append({
                        "Fecha": fecha_val.strftime("%Y-%m-%d"),
                        "Edificio": edificio_val,
                        "Responsable": user_nombre_completo,
                        "Cargo": user_cargo,
                        "Datos": df_chk_save.to_dict(orient="records")
                    })

                    save_persistent_db()
                    st.success(f"Jornada guardada para **{edificio_val}**.")
                    st.session_state.creando_jornada = False
                    st.rerun()

    # HISTORIAL DE JORNADAS
    st.markdown("---")
    st.markdown("### Historial de Jornadas e Inspecciones Creadas")

    mis_jornadas = st.session_state.db_checklists.get(user_email, [])

    if len(mis_jornadas) > 0:
        for idx_j, j in enumerate(reversed(mis_jornadas), 1):
            with st.expander(f"📌 Jornada #{len(mis_jornadas) - idx_j + 1} | Edificio: {j['Edificio']} | Fecha: {j['Fecha']} | Responsable: {j['Responsable']}"):
                df_data = pd.DataFrame(j["Datos"])

                st.markdown("#### Detalle de Actividades e Inspección")
                for r_idx, row in df_data.iterrows():
                    col_det1, col_det2, col_det3 = st.columns([4, 2, 2])
                    with col_det1:
                        st.markdown(f"**[{row['Jornada']}] N° {row['N°']}. {row['Actividad']}**")
                        st.caption(f"Observaciones: {row['Observaciones'] if row['Observaciones'] else 'Sin observaciones'}")
                    with col_det2:
                        st.markdown(f"**Estado:** `{row['Estado']}`")
                    with col_det3:
                        if row.get("Foto_B64") is not None:
                            img_evidencia = base64_to_image(row["Foto_B64"])
                            if img_evidencia:
                                with st.popover("📷 Ver Foto"):
                                    st.image(img_evidencia, caption=f"Evidencia: {row['Actividad']}", use_container_width=True)
                        else:
                            st.caption("Sin foto")

                    st.markdown("<hr style='margin: 4px 0; border-color: #c2c7d2;'>", unsafe_allow_html=True)

                csv_bytes = export_dataframe_to_excel_csv(df_data)
                st.download_button(
                    label=f"📥 Descargar Reporte CSV (Excel) - {j['Edificio']}",
                    data=csv_bytes,
                    file_name=f"Checklist_{j['Edificio'].replace(' ', '_')}_{j['Fecha']}.csv",
                    mime="text/csv",
                    key=f"dl_{idx_j}"
                )
    else:
        st.info("Aún no ha creado jornadas de inspección. Presione 'Crear Nueva Jornada' para comenzar.")

# ==========================================
# 8. MÓDULO 2: CONTROL DE RENDIMIENTO
# ==========================================
with tab_rend:
    st.markdown("### Control de Rendimiento por Trabajador")
    st.caption("Asignación de rubros, cálculo de Horas-Hombre (HH) y diagnóstico de productividad.")

    col1, col2 = st.columns(2)
    with col1:
        nombres_obreros = [t["nombre"] for t in TRABAJADORES_NO_MINA]
        trabajador_sel = st.selectbox("Seleccionar Trabajador (28 Activos):", nombres_obreros)
        cargo_actual = next(t["cargo"] for t in TRABAJADORES_NO_MINA if t["nombre"] == trabajador_sel)
        st.info(f"**Cargo en obra:** {cargo_actual}")

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
        if horas_acumuladas == 0:
            st.warning("Seleccione al menos un intervalo de horario.")
        elif avance_cant <= 0:
            st.warning("Ingrese un avance mayor a 0.")
        else:
            rend_real = round(horas_acumuladas / avance_cant, 3)
            rend_teorico = RENDIMIENTOS_TEORICOS.get(rubro_sel, 1.0)
            estado_diag = "EFICIENTE" if rend_real <= rend_teorico else "EXCESO DE HH"

            nuevo_registro = {
                "Usuario_Registro": user_email,
                "Cargo_Registrador": user_cargo,
                "Fecha": datetime.date.today().strftime("%Y-%m-%d"),
                "Trabajador": trabajador_sel,
                "Cargo_Obrero": cargo_actual,
                "Rubro": rubro_sel,
                "Horas Trabajadas (HH)": horas_acumuladas,
                "Avance": avance_cant,
                "Unidad": unidad_medida,
                "Rend. Real (HH/Unid)": rend_real,
                "Rend. Teórico": rend_teorico,
                "Estado": estado_diag,
            }

            if user_email not in st.session_state.db_rendimientos:
                st.session_state.db_rendimientos[user_email] = []

            st.session_state.db_rendimientos[user_email].append(nuevo_registro)
            save_persistent_db()
            st.success(f"Rendimiento registrado correctamente para {trabajador_sel}.")

    st.markdown("---")
    st.markdown("### Registros de Rendimiento Guardados")

    mis_rendimientos = st.session_state.db_rendimientos.get(user_email, [])
    if len(mis_rendimientos) > 0:
        df_mis_r = pd.DataFrame(mis_rendimientos)
        st.dataframe(df_mis_r, use_container_width=True)

        csv_bytes_r = export_dataframe_to_excel_csv(df_mis_r)
        st.download_button(label="📥 Descargar Rendimientos en CSV (Excel)", data=csv_bytes_r, file_name=f"Rendimientos_{user_email}.csv", mime="text/csv")
    else:
        st.info("Aún no existen registros en su historial.")

# ==========================================
# 9. MÓDULO ADMINISTRADOR
# ==========================================
if es_admin:
    tab_admin = tabs_app[2]
    with tab_admin:
        st.markdown("### Panel de Control Administrador")
        st.caption("Módulo exclusivo para monitoreo de usuarios y asignación de permisos.")

        st.markdown("#### Gestión de Administradores de la Plataforma")
        col_adm1, col_adm2 = st.columns([2, 1])

        with col_adm1:
            nuevo_admin_mail = st.text_input("Ingrese correo para conceder permisos de Administrador:", placeholder="usuario@correo.com")
            if st.button("Otorgar Acceso Administrador"):
                if nuevo_admin_mail:
                    mail_clean = nuevo_admin_mail.strip().lower()
                    if mail_clean not in st.session_state.admin_emails:
                        st.session_state.admin_emails.append(mail_clean)
                        save_persistent_db()
                        st.success(f"Se otorgaron permisos de administrador a: {mail_clean}")
                        st.rerun()
                    else:
                        st.warning("El correo ingresado ya es administrador.")

        with col_adm2:
            st.markdown("**Administradores Actuales:**")
            for adm in st.session_state.admin_emails:
                st.write(f"- `{adm}`")

        st.markdown("---")

        st.markdown("#### Usuarios Activos en la Plataforma")
        df_users = pd.DataFrame(st.session_state.db_usuarios)
        st.dataframe(df_users, use_container_width=True)

        st.markdown("#### Resumen Global de Actividad por Usuario")
        resumen_actividad = []
        for u in st.session_state.db_usuarios:
            e = u["Correo"]
            num_c = len(st.session_state.db_checklists.get(e, []))
            num_r = len(st.session_state.db_rendimientos.get(e, []))
            resumen_actividad.append({
                "Usuario": f"{u['Nombres']} {u['Apellidos']}".strip() or e,
                "Correo": e,
                "Cargo": u["Cargo"],
                "Checklists Guardados": num_c,
                "Rendimientos Registrados": num_r,
                "Estado": u["Estado"]
            })

        df_act = pd.DataFrame(resumen_actividad)
        st.dataframe(df_act, use_container_width=True)

        csv_admin_bytes = export_dataframe_to_excel_csv(df_act)
        st.download_button(
            label="📥 Descargar Reporte de Usuarios (Excel CSV)",
            data=csv_admin_bytes,
            file_name=f"Reporte_Usuarios_AlphaBuilders_{datetime.date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )