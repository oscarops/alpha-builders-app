import base64
import datetime
import io
import pandas as pd
from PIL import Image
import streamlit as st

# ==========================================
# 1. CONFIGURACIÓN Y ESTILOS STUDIO APPLE SLATE
# ==========================================
st.set_page_config(
    page_title="Alpha Builders | Cazadores de Inversiones",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# LOGO OFICIAL ALPHA BUILDERS EN BASE64
LOGO_ALPHA_B64 = "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
st.image(base64.b64decode(LOGO_ALPHA_B64), use_column_width=False, width=320)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

    * {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif !important;
        transition: all 0.25s ease;
    }

    /* Fondo Lienzo Principal Studio Light Slate */
    .stApp {
        background: #f4f5f8 !important;
        color: #1a1c23 !important;
    }

    /* Texto Universal de Alta Legibilidad */
    label, p, span, div, h1, h2, h3, h4, h5, h6, .stMarkdown {
        color: #1a1c23 !important;
    }

    .stCaption, caption, small, [data-testid="stCaptionContainer"] {
        color: #6c707a !important;
    }

    /* SIDEBAR CORREGIDO (BLANCO PURO) */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #e1e3e8 !important;
        padding-top: 15px !important;
    }

    /* Tarjeta de Perfil en Sidebar */
    .sidebar-profile-box {
        background: #f8f9fa;
        border: 1px solid #e1e3e8;
        border-radius: 18px;
        padding: 18px;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.03);
    }

    .sidebar-user-name {
        font-size: 1.15rem;
        font-weight: 800;
        color: #1a1c23 !important;
        margin-top: 10px;
        margin-bottom: 2px;
    }

    .sidebar-user-email {
        font-size: 0.85rem;
        color: #0071e3 !important;
        font-weight: 600;
        margin-bottom: 6px;
    }

    .sidebar-user-cargo {
        display: inline-block;
        background: #1a1c23;
        color: #ffffff !important;
        font-size: 0.72rem;
        font-weight: 700;
        padding: 4px 12px;
        border-radius: 20px;
        text-transform: uppercase;
    }

    /* TARJETAS PRINCIPALES Y BANNER */
    .executive-card-light {
        background: #ffffff;
        border: 1px solid #e1e3e8;
        border-radius: 20px;
        padding: 28px;
        box-shadow: 0 8px 25px rgba(0,0,0,0.04);
        margin-bottom: 25px;
    }

    /* TARJETAS KPIS ELEVADAS CON SOMBRA Y HOVER */
    .kpi-card-studio {
        background: #ffffff;
        border: 1px solid #e1e3e8;
        border-radius: 18px;
        padding: 22px;
        text-align: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
    }
    .kpi-card-studio:hover {
        border-color: #1a1c23;
        transform: translateY(-3px);
        box-shadow: 0 10px 25px rgba(0,0,0,0.08);
    }
    .kpi-val-studio {
        font-size: 2.7rem;
        font-weight: 900;
        color: #1a1c23 !important;
        letter-spacing: -0.03em;
    }
    .kpi-lbl-studio {
        font-size: 0.75rem;
        color: #6c707a !important;
        text-transform: uppercase;
        font-weight: 700;
        letter-spacing: 0.08em;
        margin-top: 4px;
    }

    /* INPUTS Y SELECCIONES CLARAS */
    .stTextInput input, .stSelectbox > div > div, .stNumberInput input, .stDateInput input {
        background-color: #ffffff !important;
        color: #1a1c23 !important;
        border: 1px solid #d1d5db !important;
        border-radius: 12px !important;
        padding: 12px 16px !important;
        font-size: 0.95rem !important;
        font-weight: 500 !important;
    }

    .stTextInput input:focus, .stSelectbox > div > div:focus, .stNumberInput input:focus {
        border-color: #0071e3 !important;
        box-shadow: 0 0 0 3px rgba(0, 113, 227, 0.15) !important;
    }

    .stTextInput input:disabled {
        background-color: #f0f2f5 !important;
        color: #6c707a !important;
        border-color: #e1e3e8 !important;
    }

    /* PESTAÑAS (SEGMENT CONTROL APPLE STUDIO) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: #e5e7eb !important;
        padding: 6px;
        border-radius: 16px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 12px;
        padding: 10px 24px;
        background-color: transparent !important;
        color: #4b5563 !important;
        font-weight: 600;
        border: none !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #1a1c23 !important;
        font-weight: 800;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08) !important;
    }

    /* BOTONES MONOCROMÁTICOS ELEGANTES */
    .stButton > button {
        background-color: #1a1c23 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        font-weight: 700 !important;
        padding: 11px 26px !important;
        box-shadow: 0 4px 12px rgba(26, 28, 35, 0.2) !important;
    }
    .stButton > button * {
        color: #ffffff !important;
    }
    .stButton > button:hover {
        background-color: #0071e3 !important;
        box-shadow: 0 6px 18px rgba(0, 113, 227, 0.3) !important;
        transform: translateY(-2px);
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. FUNCIONES DE APOYO Y ALMACENAMIENTO PERSISTENTE (BASE64)
# ==========================================
def image_to_base64(image_file):
    """Guarda la imagen de perfil en formato Base64 para que no se borre al actualizar controles"""
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
    """Reconstruye la imagen almacenada desde Base64"""
    if b64_str:
        try:
            img_data = base64.b64decode(b64_str)
            return Image.open(io.BytesIO(img_data))
        except Exception:
            return None
    return None

# ==========================================
# 3. BASE DE DATOS Y ESTADOS DE SESIÓN
# ==========================================
if "admin_emails" not in st.session_state:
    st.session_state.admin_emails = ["oscarsebitas2013@gmail.com"]

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_email = ""
    st.session_state.usuario_nombres = ""
    st.session_state.usuario_apellidos = ""
    st.session_state.usuario_cargo = ""

if "db_fotos_perfil_b64" not in st.session_state:
    st.session_state.db_fotos_perfil_b64 = {}

if "db_usuarios" not in st.session_state:
    st.session_state.db_usuarios = [
        {
            "Nombres": "Oscar Sebastián",
            "Apellidos": "Narváez Ojeda",
            "Correo": "oscarsebitas2013@gmail.com",
            "Password": "Al678554",
            "Cargo": "Residente",
            "Fecha_Registro": "2026-07-26",
            "Estado": "Activo",
        }
    ]

if "db_checklists" not in st.session_state:
    st.session_state.db_checklists = {}

if "db_rendimientos" not in st.session_state:
    st.session_state.db_rendimientos = {}

# NÓMINA OFICIAL DE EDIFICIOS
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

# NÓMINA REAL DE 28 TRABAJADORES OPERATIVOS
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
# 4. MÓDULO DE LOGIN & REGISTRO
# ==========================================
if not st.session_state.autenticado:
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])

    with col_l2:
        st.markdown(
            f"""
            <div class="executive-card-light" style="text-align: center; margin-top: 40px;">
                <img src="data:image/png;base64,{LOGO_ALPHA_B64}" style="max-width: 280px; width: 100%; margin-bottom: 15px;">
                <p style="color: #6c707a; font-size: 1.05rem; font-weight: 500;">Portal Corporativo de Control de Obra y Calidad</p>
            </div>
        """,
            unsafe_allow_html=True,
        )

        tab_login, tab_register, tab_reset = st.tabs(["Iniciar Sesión", "Registrarse", "¿Olvidaste tu Contraseña?"])

        # --- INICIAR SESIÓN ---
        with tab_login:
            st.markdown("### Iniciar Sesión")
            st.caption("Ingrese sus credenciales corporativas registradas.")

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
                            st.error("Contraseña incorrecta. Intente nuevamente.")
                    else:
                        st.error("El usuario no existe. Complete el registro.")
                else:
                    st.error("Ingrese su correo y contraseña.")

        # --- REGISTRARSE ---
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

                            st.success("¡Registro completado exitosamente!")
                            st.rerun()
                else:
                    st.error("Por favor complete todos los campos requeridos.")

        # --- RECUPERACIÓN DE CONTRASEÑA ---
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
                            st.success(f"Contraseña actualizada con éxito para {mail_clean}.")
                        else:
                            st.error("El correo ingresado no está registrado.")
                else:
                    st.error("Complete todos los campos.")

    st.stop()

# ==========================================
# 5. BARRA LATERAL CON PERSISTENCIA DE IMAGEN Y CONFIGURACIÓN
# ==========================================
user_email = st.session_state.usuario_email
user_nombre_completo = f"{st.session_state.usuario_nombres} {st.session_state.usuario_apellidos}".strip()
user_cargo = st.session_state.usuario_cargo
es_admin = user_email in st.session_state.admin_emails

with st.sidebar:
    # Logo en Sidebar
    st.markdown(f'<img src="data:image/png;base64,{LOGO_ALPHA_B64}" style="width:100%; max-width:210px; margin-bottom:15px; display:block; margin-left:auto; margin-right:auto;">', unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; font-weight: 800;'>Perfil de Usuario</h4>", unsafe_allow_html=True)

    # Carga persistente de imagen en Base64
    b64_foto = st.session_state.db_fotos_perfil_b64.get(user_email, None)
    img_obj = base64_to_image(b64_foto)

    if img_obj is not None:
        st.image(img_obj, use_column_width=True)

    st.markdown(
        f"""
        <div class="sidebar-profile-box">
            <div class="sidebar-user-name">{user_nombre_completo}</div>
            <div class="sidebar-user-email">{user_email}</div>
            <div class="sidebar-user-cargo">{user_cargo}</div>
        </div>
    """,
        unsafe_allow_html=True,
    )

    if es_admin:
        st.markdown("<div style='text-align: center; margin-bottom: 15px; font-size: 0.78rem; color: #1a1c23; font-weight: 800; background: #e5e7eb; padding: 8px; border-radius: 12px; border: 1px solid #d1d5db;'>ADMINISTRADOR GENERAL</div>", unsafe_allow_html=True)

    st.markdown("---")

    # Módulo de Configuración de Cuenta Limpio en Sidebar
    with st.expander("⚙️ Configuración de Cuenta", expanded=False):
        st.caption("Ajustes personales y de seguridad:")
        
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

            st.success("Configuración actualizada correctamente.")
            st.rerun()

    st.markdown("---")
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

    st.markdown("---")
    st.caption("Alpha Builders Portal v13.0\nCazadores de Inversiones Edition")

# ==========================================
# 6. DASHBOARD PRINCIPAL
# ==========================================
st.markdown(
    f"""
    <div class="executive-card-light">
        <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap;">
            <div>
                <h1 style="font-size: 2.3rem; letter-spacing: -0.03em; font-weight: 900; margin: 0;">Alpha Builders</h1>
                <p style="color: #6c707a; margin-top: 4px; font-size: 1.05rem;">Portal de Control e Inspección | Usuario Activo: <b>{user_nombre_completo}</b> ({user_cargo})</p>
            </div>
            <img src="data:image/png;base64,{LOGO_ALPHA_B64}" style="max-width: 240px; width: 100%; margin-top: 10px;">
        </div>
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

# Pestañas principales
pestanas = ["Checklist Diario", "Control de Rendimiento"]
if es_admin:
    pestanas.append("Panel Admin")

tabs_app = st.tabs(pestanas)
tab_chk = tabs_app[0]
tab_rend = tabs_app[1]

# ==========================================
# 7. MÓDULO 1: CHECKLIST DIARIO (CON CASILLAS OBLIGATORIAS SINE SELECCIONAR Y SELECCIÓN DE 12 EDIFICIOS)
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

    # FORMULARIO DE CREACIÓN DE JORNADA
    if st.session_state.creando_jornada:
        st.markdown("---")
        with st.container():
            st.markdown("#### Configuración de la Nueva Jornada")
            col_m1, col_m2, col_m3 = st.columns(3)

            with col_m1:
                # SELECCIÓN ENTRE LOS 12 EDIFICIOS
                edificio_val = st.selectbox("Edificio / Proyecto:", EDIFICIOS_ALPHA, key="sel_edificio")
            with col_m2:
                # RESPONSABLE AUTOMÁTICO BLOQUEADO
                st.text_input("Responsable:", value=user_nombre_completo, disabled=True, help="Obtenido automáticamente de la sesión activa.")
            with col_m3:
                fecha_val = st.date_input("Fecha de Inspección:", datetime.date.today(), key="sel_fecha")

            st.markdown("---")

            with st.form("form_checklist_jornada"):
                st.markdown("#### Jornada de la Mañana")
                resp_manana = []

                for idx, act in enumerate(ACTIVIDADES_MANANA, 1):
                    st.markdown(f"**N° {idx}. {act}**")
                    c_sel, c_obs, c_foto = st.columns([2, 3, 3])

                    with c_sel:
                        # INDEX=NONE: OBLIGA A SELECCIONAR MANUALMENTE
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
                        ft = st.file_uploader("Foto (Opcional)", type=["jpg", "jpeg", "png"], key=f"m_ft_{idx}")

                    st.markdown("<hr style='margin: 8px 0; border-color: #e1e3e8;'>", unsafe_allow_html=True)
                    resp_manana.append({"Jornada": "Mañana", "N°": idx, "Actividad": act, "Estado": est, "Observaciones": ob, "Foto_Objeto": ft, "Foto_Adjunta": "Sí" if ft is not None else "No"})

                st.markdown("#### Jornada de la Tarde")
                resp_tarde = []

                for idx, act in enumerate(ACTIVIDADES_TARDE, 1):
                    st.markdown(f"**N° {idx}. {act}**")
                    c_sel, c_obs, c_foto = st.columns([2, 3, 3])

                    with c_sel:
                        # INDEX=NONE: OBLIGA A SELECCIONAR MANUALMENTE
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
                        ft = st.file_uploader("Foto (Opcional)", type=["jpg", "jpeg", "png"], key=f"t_ft_{idx}")

                    st.markdown("<hr style='margin: 8px 0; border-color: #e1e3e8;'>", unsafe_allow_html=True)
                    resp_tarde.append({"Jornada": "Tarde", "N°": idx, "Actividad": act, "Estado": est, "Observaciones": ob, "Foto_Objeto": ft, "Foto_Adjunta": "Sí" if ft is not None else "No"})

                btn_guardar_chk = st.form_submit_button("Guardar Jornada de Inspección", type="primary")

            if btn_guardar_chk:
                all_chk_data = resp_manana + resp_tarde
                
                # Validar que no queden campos vacíos
                sin_responder = [item["Actividad"] for item in all_chk_data if item["Estado"] is None]
                if sin_responder:
                    st.error(f"⚠️ Por favor seleccione el estado de todas las actividades ({len(sin_responder)} pendientes).")
                else:
                    df_chk_save = pd.DataFrame(all_chk_data)

                    st.session_state.db_checklists[user_email].append({
                        "Fecha": fecha_val.strftime("%Y-%m-%d"),
                        "Edificio": edificio_val,
                        "Responsable": user_nombre_completo,
                        "Cargo": user_cargo,
                        "Datos": df_chk_save
                    })

                    st.success(f"Jornada guardada exitosamente para el edificio **{edificio_val}**.")
                    st.session_state.creando_jornada = False
                    st.rerun()

    # HISTORIAL DE JORNADAS CREADAS
    st.markdown("---")
    st.markdown("### Historial de Jornadas e Inspecciones Creadas")

    mis_jornadas = st.session_state.db_checklists.get(user_email, [])

    if len(mis_jornadas) > 0:
        for idx_j, j in enumerate(reversed(mis_jornadas), 1):
            with st.expander(f"📌 Jornada #{len(mis_jornadas) - idx_j + 1} | Edificio: {j['Edificio']} | Fecha: {j['Fecha']} | Responsable: {j['Responsable']}"):
                df_display = j["Datos"]
                st.dataframe(df_display.drop(columns=["Foto_Objeto"]), use_container_width=True)

                csv_item = df_display.drop(columns=["Foto_Objeto"]).to_csv(index=False).encode("utf-8")
                st.download_button(
                    label=f"Descargar CSV - {j['Edificio']}",
                    data=csv_item,
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

            st.session_state.db_rendimientos[user_email].append(nuevo_registro)
            st.success(f"Rendimiento registrado correctamente para {trabajador_sel}.")

    st.markdown("---")
    st.markdown("### Registros de Rendimiento Guardados")

    mis_rendimientos = st.session_state.db_rendimientos.get(user_email, [])
    if len(mis_rendimientos) > 0:
        df_mis_r = pd.DataFrame(mis_rendimientos)
        st.dataframe(df_mis_r, use_container_width=True)

        csv_r = df_mis_r.to_csv(index=False).encode("utf-8")
        st.download_button(label="Descargar Rendimientos (CSV)", data=csv_r, file_name=f"Rendimientos_{user_email}.csv", mime="text/csv")
    else:
        st.info("Aún no existen registros en su historial.")

# ==========================================
# 9. MÓDULO ADMINISTRADOR
# ==========================================
if es_admin:
    tab_admin = tabs_app[2]
    with tab_admin:
        st.markdown("### Panel de Control Administrador")
        st.caption("Módulo exclusivo para monitoreo de usuarios, cargos y asignación de permisos.")

        # --- GESTIÓN DE ADMINS ---
        st.markdown("#### Gestión de Administradores de la Plataforma")
        col_adm1, col_adm2 = st.columns([2, 1])

        with col_adm1:
            nuevo_admin_mail = st.text_input("Ingrese correo para conceder permisos de Administrador:", placeholder="usuario@correo.com")
            if st.button("Otorgar Acceso Administrador"):
                if nuevo_admin_mail:
                    mail_clean = nuevo_admin_mail.strip().lower()
                    if mail_clean not in st.session_state.admin_emails:
                        st.session_state.admin_emails.append(mail_clean)
                        st.success(f"Se otorgaron permisos de administrador a: {mail_clean}")
                        st.rerun()
                    else:
                        st.warning("El correo ingresado ya es administrador.")

        with col_adm2:
            st.markdown("**Administradores Actuales:**")
            for adm in st.session_state.admin_emails:
                st.write(f"- `{adm}`")

        st.markdown("---")

        # --- MONITOR DE USUARIOS ---
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

        csv_admin = df_act.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Descargar Reporte de Usuarios Activos (CSV)",
            data=csv_admin,
            file_name=f"Reporte_Usuarios_AlphaBuilders_{datetime.date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )