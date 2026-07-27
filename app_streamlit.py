import datetime
import pandas as pd
import streamlit as st

# ==========================================
# 1. CONFIGURACIÓN Y ESTILOS APPLE LIGHT
# ==========================================
st.set_page_config(
    page_title="Alpha Builders | Portal Oficial de Obra",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    /* Fondo General Apple Light Canvas */
    .stApp {
        background-color: #f5f5f7 !important;
        color: #1d1d1f !important;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Helvetica Neue", Helvetica, Arial, sans-serif;
    }

    /* Legibilidad de textos y etiquetas */
    label, p, span, div, h1, h2, h3, h4, h5, h6, .stMarkdown {
        color: #1d1d1f !important;
    }
    
    .stCaption, caption, small, [data-testid="stCaptionContainer"] {
        color: #6e6e73 !important;
    }

    /* Sidebar Estilo Apple Light */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #d2d2d7 !important;
        padding-top: 20px !important;
    }
    [data-testid="stSidebar"] * {
        color: #1d1d1f !important;
    }

    /* Tarjeta de Avatar en Sidebar */
    .sidebar-profile-card {
        text-align: center;
        padding: 15px 10px;
        background-color: #f5f5f7;
        border-radius: 16px;
        border: 1px solid #e5e5e7;
        margin-bottom: 20px;
    }

    /* Entradas de texto e inputs legibles */
    .stTextInput input, .stSelectbox > div > div, .stNumberInput input, .stDateInput input {
        background-color: #ffffff !important;
        color: #1d1d1f !important;
        border: 1px solid #d2d2d7 !important;
        border-radius: 12px !important;
        padding: 12px 16px !important;
        font-size: 0.95rem !important;
        font-weight: 500 !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04) !important;
    }

    .stTextInput input:focus, .stSelectbox > div > div:focus, .stNumberInput input:focus {
        border-color: #0071e3 !important;
        box-shadow: 0 0 0 4px rgba(0, 113, 227, 0.15) !important;
    }

    ::placeholder {
        color: #86868b !important;
        opacity: 1 !important;
    }

    /* Tarjetas Contenedoras */
    .apple-card-light {
        background-color: #ffffff;
        border: 1px solid #e5e5e7;
        border-radius: 20px;
        padding: 30px;
        box-shadow: 0 4px 18px rgba(0, 0, 0, 0.04);
        margin-bottom: 25px;
    }

    /* Tarjetas KPIs */
    .kpi-card-light {
        background-color: #ffffff;
        border: 1px solid #e5e5e7;
        border-radius: 18px;
        padding: 22px;
        text-align: center;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.03);
        transition: all 0.2s ease;
    }
    .kpi-card-light:hover {
        border-color: #d2d2d7;
        transform: translateY(-2px);
    }
    .kpi-value-light {
        font-size: 2.4rem;
        font-weight: 700;
        color: #1d1d1f !important;
        letter-spacing: -0.03em;
    }
    .kpi-label-light {
        font-size: 0.8rem;
        color: #86868b !important;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.06em;
        margin-top: 6px;
    }

    /* Pestañas (Segment Control) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #e8e8ed !important;
        padding: 5px;
        border-radius: 14px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px;
        padding: 10px 22px;
        background-color: transparent !important;
        color: #6e6e73 !important;
        font-weight: 500;
        border: none !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #1d1d1f !important;
        font-weight: 600;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08) !important;
    }

    /* Botones Accento Azul Apple */
    .stButton > button {
        background-color: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 10px 22px !important;
        box-shadow: 0 2px 6px rgba(0, 113, 227, 0.25) !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button * {
        color: #ffffff !important;
    }
    .stButton > button:hover {
        background-color: #0077ed !important;
        box-shadow: 0 4px 12px rgba(0, 113, 227, 0.35) !important;
        transform: translateY(-1px);
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. ESTADOS DE SESIÓN Y GESTIÓN DE ROLES
# ==========================================
if "admin_emails" not in st.session_state:
    st.session_state.admin_emails = ["oscarsebitas2013@gmail.com"]

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_email = ""
    st.session_state.usuario_nombres = ""
    st.session_state.usuario_apellidos = ""
    st.session_state.usuario_cargo = ""

if "db_fotos_perfil" not in st.session_state:
    st.session_state.db_fotos_perfil = {}

if "db_usuarios" not in st.session_state:
    st.session_state.db_usuarios = [
        {
            "Nombres": "Oscar Sebastián",
            "Apellidos": "Narváez Ojeda",
            "Correo": "oscarsebitas2013@gmail.com",
            "Cargo": "Residente",
            "Fecha_Registro": "2026-07-26",
            "Estado": "Activo",
        }
    ]

if "db_checklists" not in st.session_state:
    st.session_state.db_checklists = {}

if "db_rendimientos" not in st.session_state:
    st.session_state.db_rendimientos = {}

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
# 3. MÓDULO DE LOGIN & REGISTRO
# ==========================================
if not st.session_state.autenticado:
    st.markdown(
        """
        <div class="apple-card-light" style="text-align: center; max-width: 620px; margin: 40px auto 20px auto;">
            <h1 style="font-size: 2.6rem; letter-spacing: -0.03em; font-weight: 700;">Alpha Builders</h1>
            <p style="color: #6e6e73; font-size: 1.05rem;">Portal Oficial de Control de Obra y Calidad</p>
        </div>
    """,
        unsafe_allow_html=True,
    )

    col_c1, col_c2, col_c3 = st.columns([1, 2, 1])

    with col_c2:
        tab_login, tab_register = st.tabs(["Iniciar Sesión", "Registrarse"])

        # --- INICIAR SESIÓN ---
        with tab_login:
            st.markdown("### Iniciar Sesión")
            st.caption("Ingrese sus credenciales de acceso registradas.")

            login_email = st.text_input("Correo electrónico:", placeholder="nombre@correo.com", key="log_email")
            login_pass = st.text_input("Contraseña:", type="password", key="log_pass")

            if st.button("Entrar", type="primary", use_container_width=True):
                if login_email and login_pass:
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = login_email.strip().lower()

                    u_match = next((u for u in st.session_state.db_usuarios if u["Correo"] == st.session_state.usuario_email), None)
                    if u_match:
                        st.session_state.usuario_nombres = u_match["Nombres"]
                        st.session_state.usuario_apellidos = u_match["Apellidos"]
                        st.session_state.usuario_cargo = u_match["Cargo"]
                    else:
                        st.session_state.usuario_nombres = login_email.split("@")[0].title()
                        st.session_state.usuario_apellidos = ""
                        st.session_state.usuario_cargo = "Residente"
                        st.session_state.db_usuarios.append({
                            "Nombres": st.session_state.usuario_nombres,
                            "Apellidos": "",
                            "Correo": st.session_state.usuario_email,
                            "Cargo": st.session_state.usuario_cargo,
                            "Fecha_Registro": datetime.date.today().strftime("%Y-%m-%d"),
                            "Estado": "Activo"
                        })

                    if st.session_state.usuario_email not in st.session_state.db_checklists:
                        st.session_state.db_checklists[st.session_state.usuario_email] = []
                    if st.session_state.usuario_email not in st.session_state.db_rendimientos:
                        st.session_state.db_rendimientos[st.session_state.usuario_email] = []
                    st.rerun()
                else:
                    st.error("Por favor ingrese correo y contraseña.")

            st.markdown("---")
            st.caption("Acceso corporativo directo:")
            col_g, col_o = st.columns(2)
            with col_g:
                if st.button("Google Workspace", use_container_width=True):
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = "oscarsebitas2013@gmail.com"
                    st.session_state.usuario_nombres = "Oscar Sebastián"
                    st.session_state.usuario_apellidos = "Narváez Ojeda"
                    st.session_state.usuario_cargo = "Residente"
                    st.rerun()
            with col_o:
                if st.button("Microsoft Outlook", use_container_width=True):
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = "usuario.outlook@outlook.com"
                    st.session_state.usuario_nombres = "Usuario"
                    st.session_state.usuario_apellidos = "Outlook"
                    st.session_state.usuario_cargo = "Asistente"
                    st.rerun()

        # --- REGISTRARSE ---
        with tab_register:
            st.markdown("### Crear una Cuenta Nueva")
            st.caption("Complete el formulario para habilitar su perfil de acceso.")

            col_n, col_a = st.columns(2)
            with col_n:
                reg_nombres = st.text_input("Nombres:", placeholder="Ej. Juan Carlos")
            with col_a:
                reg_apellidos = st.text_input("Apellidos:", placeholder="Ej. Pérez Gómez")

            reg_email = st.text_input("Correo electrónico:", placeholder="ejemplo@correo.com", key="reg_email")
            reg_pass = st.text_input("Nueva contraseña:", type="password", key="reg_pass")
            reg_cargo = st.selectbox("Cargo / Rol en Obra:", ["Residente", "Asistente", "Ayudante"])

            if st.button("Registrarte", type="primary", use_container_width=True):
                if reg_nombres and reg_apellidos and reg_email and reg_pass:
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = reg_email.strip().lower()
                    st.session_state.usuario_nombres = reg_nombres.strip()
                    st.session_state.usuario_apellidos = reg_apellidos.strip()
                    st.session_state.usuario_cargo = reg_cargo

                    st.session_state.db_usuarios.append({
                        "Nombres": reg_nombres.strip(),
                        "Apellidos": reg_apellidos.strip(),
                        "Correo": st.session_state.usuario_email,
                        "Cargo": reg_cargo,
                        "Fecha_Registro": datetime.date.today().strftime("%Y-%m-%d"),
                        "Estado": "Activo"
                    })

                    if st.session_state.usuario_email not in st.session_state.db_checklists:
                        st.session_state.db_checklists[st.session_state.usuario_email] = []
                    if st.session_state.usuario_email not in st.session_state.db_rendimientos:
                        st.session_state.db_rendimientos[st.session_state.usuario_email] = []

                    st.success("¡Registro completado!")
                    st.rerun()
                else:
                    st.error("Por favor complete todos los campos requeridos.")

    st.stop()

# ==========================================
# 4. BARRA LATERAL (LIMPIA Y ORDENADA)
# ==========================================
user_email = st.session_state.usuario_email
user_nombre_completo = f"{st.session_state.usuario_nombres} {st.session_state.usuario_apellidos}".strip()
user_cargo = st.session_state.usuario_cargo
es_admin = user_email in st.session_state.admin_emails

with st.sidebar:
    st.markdown("<h3 style='margin-bottom: 15px;'>Perfil de Usuario</h3>", unsafe_allow_html=True)

    foto_perfil = st.session_state.db_fotos_perfil.get(user_email, None)
    if foto_perfil is not None:
        st.image(foto_perfil, use_column_width=True)

    st.markdown(f"**Usuario:** {user_nombre_completo}")
    st.markdown(f"**Correo:** `{user_email}`")
    st.markdown(f"**Cargo:** `{user_cargo}`")
    
    if es_admin:
        st.info("Rol: Administrador General")

    st.markdown("---")
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

    st.markdown("---")
    st.caption("Alpha Builders Portal v6.0\nEstilo Apple Light Canvas")

# ==========================================
# 5. DASHBOARD PRINCIPAL
# ==========================================
st.markdown(
    f"""
    <div class="apple-card-light">
        <h1 style="font-size: 2.2rem; letter-spacing: -0.03em;">Alpha Builders</h1>
        <p style="color: #6e6e73;">Panel de Control | Usuario: <b>{user_nombre_completo}</b> ({user_cargo})</p>
    </div>
""",
    unsafe_allow_html=True,
)

usr_chks = len(st.session_state.db_checklists.get(user_email, []))
usr_rnds = len(st.session_state.db_rendimientos.get(user_email, []))

k1, k2, k3 = st.columns(3)
with k1:
    st.markdown(
        '<div class="kpi-card-light"><div class="kpi-value-light">28</div><div class="kpi-label-light">Obreros Activos</div></div>',
        unsafe_allow_html=True,
    )
with k2:
    st.markdown(
        f'<div class="kpi-card-light"><div class="kpi-value-light">{usr_chks}</div><div class="kpi-label-light">Checklists Guardados</div></div>',
        unsafe_allow_html=True,
    )
with k3:
    st.markdown(
        f'<div class="kpi-card-light"><div class="kpi-value-light">{usr_rnds}</div><div class="kpi-label-light">Reportes de Rendimiento</div></div>',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# Pestañas principales incluyendo Configuración de Cuenta
pestanas = ["Checklist Diario", "Control de Rendimiento", "Configuración de Cuenta"]
if es_admin:
    pestanas.append("Panel Admin")

tabs_app = st.tabs(pestanas)
tab_chk = tabs_app[0]
tab_rend = tabs_app[1]
tab_config = tabs_app[2]

# ==========================================
# 6. MÓDULO 1: CHECKLIST DIARIO
# ==========================================
with tab_chk:
    st.markdown("### Check List Diario – Control de Obra")
    st.caption("Supervisión diaria de calidad y avance en el frente de trabajo.")

    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        proyecto_val = st.text_input("Proyecto:", value="Alpha Builders - Obra Central")
    with col_m2:
        residente_val = st.text_input("Responsable:", value=user_nombre_completo)
    with col_m3:
        fecha_val = st.date_input("Fecha:", datetime.date.today())

    st.markdown("---")

    with st.form("form_checklist_official"):
        st.markdown("#### Jornada de la Mañana")
        resp_manana = []

        for idx, act in enumerate(ACTIVIDADES_MANANA, 1):
            st.markdown(f"**N° {idx}. {act}**")
            c_sel, c_obs, c_foto = st.columns([2, 3, 3])

            with c_sel:
                est = st.radio("Estado", ["✓ Cumple", "✗ No Cumple", "N/A"], key=f"m_st_{idx}", horizontal=True)
            with c_obs:
                ob = st.text_input("Observación", key=f"m_ob_{idx}", placeholder="Observaciones...", label_visibility="collapsed")
            with c_foto:
                ft = st.file_uploader("Foto (Opcional)", type=["jpg", "jpeg", "png"], key=f"m_ft_{idx}")

            st.markdown("<hr style='margin: 8px 0; border-color: #e5e5e7;'>", unsafe_allow_html=True)
            resp_manana.append({"Jornada": "Mañana", "N°": idx, "Actividad": act, "Estado": est, "Observaciones": ob, "Foto_Objeto": ft, "Foto_Adjunta": "Sí" if ft is not None else "No"})

        st.markdown("#### Jornada de la Tarde")
        resp_tarde = []

        for idx, act in enumerate(ACTIVIDADES_TARDE, 1):
            st.markdown(f"**N° {idx}. {act}**")
            c_sel, c_obs, c_foto = st.columns([2, 3, 3])

            with c_sel:
                est = st.radio("Estado", ["✓ Cumple", "✗ No Cumple", "N/A"], key=f"t_st_{idx}", horizontal=True)
            with c_obs:
                ob = st.text_input("Observación", key=f"t_ob_{idx}", placeholder="Observaciones...", label_visibility="collapsed")
            with c_foto:
                ft = st.file_uploader("Foto (Opcional)", type=["jpg", "jpeg", "png"], key=f"t_ft_{idx}")

            st.markdown("<hr style='margin: 8px 0; border-color: #e5e5e7;'>", unsafe_allow_html=True)
            resp_tarde.append({"Jornada": "Tarde", "N°": idx, "Actividad": act, "Estado": est, "Observaciones": ob, "Foto_Objeto": ft, "Foto_Adjunta": "Sí" if ft is not None else "No"})

        btn_guardar_chk = st.form_submit_button("Guardar Checklist Diario", type="primary")

    if btn_guardar_chk:
        all_chk_data = resp_manana + resp_tarde
        df_chk_save = pd.DataFrame(all_chk_data)

        st.session_state.db_checklists[user_email].append({
            "Fecha": fecha_val.strftime("%Y-%m-%d"),
            "Proyecto": proyecto_val,
            "Responsable": residente_val,
            "Cargo": user_cargo,
            "Datos": df_chk_save
        })

        st.success(f"Checklist guardado correctamente para la cuenta {user_email}.")
        st.dataframe(df_chk_save.drop(columns=["Foto_Objeto"]), use_container_width=True)

        st.markdown("#### Evidencia Fotográfica Capturada")
        has_images = False
        cols_gal = st.columns(3)
        g_idx = 0

        for row_item in all_chk_data:
            if row_item["Foto_Objeto"] is not None:
                has_images = True
                with cols_gal[g_idx % 3]:
                    st.image(row_item["Foto_Objeto"], caption=f"[{row_item['Jornada']}] N° {row_item['N°']}: {row_item['Actividad']}", use_column_width=True)
                g_idx += 1

        if not has_images:
            st.info("No se adjuntaron fotografías en esta inspección.")

        csv_chk = df_chk_save.drop(columns=["Foto_Objeto"]).to_csv(index=False).encode("utf-8")
        st.download_button(label="Descargar Checklist Diario (CSV)", data=csv_chk, file_name=f"Checklist_{fecha_val.strftime('%Y%m%d')}_{user_email}.csv", mime="text/csv")

# ==========================================
# 7. MÓDULO 2: CONTROL DE RENDIMIENTO
# ==========================================
with tab_rend:
    st.markdown("### Control de Rendimiento por Trabajador")
    st.caption("Seleccione al operario, indique el rubro y registre sus horas ejecutadas.")

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
# 8. MÓDULO: CONFIGURACIÓN DE CUENTA (ESTILO FACEBOOK SETTINGS)
# ==========================================
with tab_config:
    st.markdown("### Configuración de Cuenta y Perfil")
    st.caption("Actualice sus datos personales, foto de perfil y configuraciones de acceso.")

    col_cfg1, col_cfg2 = st.columns([1, 2])

    with col_cfg1:
        st.markdown("#### Foto de Perfil")
        foto_actual = st.session_state.db_fotos_perfil.get(user_email, None)
        if foto_actual is not None:
            st.image(foto_actual, caption="Vista previa actual", width=180)
        else:
            st.info("No se ha configurado foto de perfil.")

        nueva_foto = st.file_uploader("Subir nueva fotografía (JPG, PNG)", type=["jpg", "jpeg", "png"], key="upload_config")
        if nueva_foto is not None:
            st.session_state.db_fotos_perfil[user_email] = nueva_foto
            st.success("Fotografía cargada correctamente.")

    with col_cfg2:
        st.markdown("#### Datos Personales")
        edit_nombres = st.text_input("Nombres:", value=st.session_state.usuario_nombres)
        edit_apellidos = st.text_input("Apellidos:", value=st.session_state.usuario_apellidos)
        
        cargos_lista = ["Residente", "Asistente", "Ayudante"]
        idx_c = cargos_lista.index(user_cargo) if user_cargo in cargos_lista else 0
        edit_cargo = st.selectbox("Cargo / Rol en Obra:", cargos_lista, index=idx_c)

        st.text_input("Correo electrónico:", value=user_email, disabled=True, help="El correo electrónico no se puede modificar.")

        if st.button("Guardar Cambios de Perfil", type="primary"):
            st.session_state.usuario_nombres = edit_nombres.strip()
            st.session_state.usuario_apellidos = edit_apellidos.strip()
            st.session_state.usuario_cargo = edit_cargo

            # Actualizar en la base de datos global de usuarios
            for u in st.session_state.db_usuarios:
                if u["Correo"] == user_email:
                    u["Nombres"] = edit_nombres.strip()
                    u["Apellidos"] = edit_apellidos.strip()
                    u["Cargo"] = edit_cargo

            st.success("Información de perfil actualizada exitosamente.")
            st.rerun()

# ==========================================
# 9. MÓDULO ADMINISTRADOR Y GESTIÓN DE ROLES
# ==========================================
if es_admin:
    tab_admin = tabs_app[3]
    with tab_admin:
        st.markdown("### Panel de Control Administrador")
        st.caption("Módulo exclusivo para monitoreo de usuarios y gestión de permisos.")

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