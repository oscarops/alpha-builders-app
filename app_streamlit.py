import datetime
import pandas as pd
import streamlit as st

# ==========================================
# 1. CONFIGURACIÓN Y CSS OFICIAL LIQUID GLASS (APPLE STORE)
# ==========================================
st.set_page_config(
    page_title="Alpha Builders | Portal Oficial de Obra",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilos CSS de grado producción estilo Apple App Store / Liquid Glass
st.markdown(
    """
    <style>
    /* Fondo Canvas Apple Dark Space */
    .stApp {
        background: radial-gradient(circle at 50% -20%, #1c2333 0%, #0c0e14 60%, #040507 100%);
        color: #ffffff !important;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Segoe UI", Roboto, sans-serif;
    }

    /* LEGIBILIDAD DE LABELS Y TEXTOS GENERALES */
    label, p, span, h1, h2, h3, h4, h5, h6, .stMarkdown {
        color: #ffffff !important;
        font-weight: 500;
    }
    
    .stCaption, caption, small, [data-testid="stCaptionContainer"] {
        color: #98989d !important;
    }

    /* RECUADROS DE ENTRADA CORREGIDOS (DARK LIQUID GLASS) */
    .stTextInput input, .stSelectbox > div > div, .stNumberInput input, .stDateInput input {
        background-color: rgba(20, 24, 35, 0.85) !important;
        color: #ffffff !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        border-radius: 12px !important;
        padding: 12px 16px !important;
        font-size: 0.95rem !important;
        font-weight: 500 !important;
        backdrop-filter: blur(15px) !important;
        -webkit-backdrop-filter: blur(15px) !important;
    }

    /* FOCO EN INPUTS - BORDE AZUL ELECTRIC APPLE */
    .stTextInput input:focus, .stSelectbox > div > div:focus, .stNumberInput input:focus {
        border-color: #2997ff !important;
        box-shadow: 0 0 12px rgba(41, 151, 255, 0.4) !important;
        background-color: rgba(25, 30, 45, 0.95) !important;
    }

    /* PLACEHOLDERS EN GRIS CLARO LEGBLE */
    ::placeholder {
        color: #8e8e93 !important;
        opacity: 1 !important;
    }

    /* TARJETAS CONTENEDORAS GLASSMORPHISM */
    .glass-card-official {
        background: rgba(255, 255, 255, 0.04);
        backdrop-filter: blur(25px);
        -webkit-backdrop-filter: blur(25px);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 24px;
        padding: 32px;
        box-shadow: 0 20px 50px rgba(0, 0, 0, 0.5);
        margin-bottom: 25px;
    }

    /* TARJETAS KPIS */
    .kpi-card-official {
        background: rgba(255, 255, 255, 0.03);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 20px;
        padding: 22px;
        text-align: center;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
        transition: all 0.3s ease;
    }
    .kpi-card-official:hover {
        background: rgba(255, 255, 255, 0.07);
        border-color: rgba(255, 255, 255, 0.25);
        transform: translateY(-2px);
    }
    .kpi-value-official {
        font-size: 2.3rem;
        font-weight: 700;
        color: #ffffff !important;
        letter-spacing: -0.02em;
    }
    .kpi-label-official {
        font-size: 0.8rem;
        color: #98989d !important;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.08em;
        margin-top: 6px;
    }

    /* PESTAÑA TIPO SEGMENT CONTROL APPLE */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: rgba(255, 255, 255, 0.06) !important;
        backdrop-filter: blur(15px);
        padding: 6px;
        border-radius: 16px;
        border: 1px solid rgba(255, 255, 255, 0.12);
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 12px;
        padding: 10px 24px;
        background-color: transparent !important;
        color: #8e8e93 !important;
        font-weight: 500;
        border: none !important;
    }
    .stTabs [aria-selected="true"] {
        background: rgba(255, 255, 255, 0.18) !important;
        backdrop-filter: blur(10px);
        color: #ffffff !important;
        font-weight: 600;
        border: 1px solid rgba(255, 255, 255, 0.25) !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
    }

    /* BOTONES ACCENTO AZUL APPLE */
    .stButton > button {
        background: linear-gradient(180deg, #0071e3 0%, #005bb5 100%) !important;
        color: #ffffff !important;
        border-radius: 12px !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 12px 24px !important;
        box-shadow: 0 4px 15px rgba(0, 113, 227, 0.35) !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        background: linear-gradient(180deg, #147ce5 0%, #0066cc 100%) !important;
        box-shadow: 0 6px 20px rgba(0, 113, 227, 0.5) !important;
        transform: translateY(-1px);
    }

    /* OCULTAR ELEMENTOS PREDETERMINADOS */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. ESTADOS DE SESIÓN Y REGISTRO
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_email = ""
    st.session_state.usuario_nombre = ""
    st.session_state.usuario_cargo = ""

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
# 3. PÁGINA DE LOGIN / REGISTRO
# ==========================================
if not st.session_state.autenticado:
    st.markdown(
        """
        <div class="glass-card-official" style="text-align: center; max-width: 650px; margin: 40px auto 20px auto;">
            <h1 style="font-size: 2.5rem; letter-spacing: -0.03em;">ALPHA BUILDERS</h1>
            <p style="color: #98989d;">Portal Oficial de Control de Obra y Calidad</p>
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
            st.caption("Ingrese sus credenciales registradas.")

            login_email = st.text_input("Correo electrónico:", placeholder="nombre@correo.com", key="log_email")
            login_pass = st.text_input("Contraseña:", type="password", key="log_pass")

            if st.button("Entrar", type="primary", use_container_width=True):
                if login_email and login_pass:
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = login_email.strip().lower()
                    st.session_state.usuario_nombre = login_email.split("@")[0].title()
                    st.session_state.usuario_cargo = "Residente"

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
                if st.button("🌐 Con Google", use_container_width=True):
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = "usuario.google@gmail.com"
                    st.session_state.usuario_nombre = "Usuario Google"
                    st.session_state.usuario_cargo = "Residente"
                    st.rerun()
            with col_o:
                if st.button("🏢 Con Outlook", use_container_width=True):
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = "usuario.outlook@outlook.com"
                    st.session_state.usuario_nombre = "Usuario Outlook"
                    st.session_state.usuario_cargo = "Residente"
                    st.rerun()

        # --- REGISTRARSE ---
        with tab_register:
            st.markdown("### Crear una Cuenta Nueva")
            st.caption("Complete todos los campos para registrar su acceso.")

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
                    st.session_state.usuario_nombre = f"{reg_nombres.strip()} {reg_apellidos.strip()}"
                    st.session_state.usuario_cargo = reg_cargo

                    if st.session_state.usuario_email not in st.session_state.db_checklists:
                        st.session_state.db_checklists[st.session_state.usuario_email] = []
                    if st.session_state.usuario_email not in st.session_state.db_rendimientos:
                        st.session_state.db_rendimientos[st.session_state.usuario_email] = []

                    st.success("¡Registro completado!")
                    st.rerun()
                else:
                    st.error("Por favor complete todos los campos de registro.")

    st.stop()

# ==========================================
# 4. BARRA LATERAL CON PERFIL
# ==========================================
user_email = st.session_state.usuario_email
user_nombre = st.session_state.usuario_nombre
user_cargo = st.session_state.usuario_cargo

with st.sidebar:
    st.markdown("### 👤 Perfil de Sesión")
    st.markdown(f"**Usuario:** `{user_nombre}`")
    st.markdown(f"**Correo:** `{user_email}`")
    st.markdown(f"**Cargo:** `{user_cargo}`")

    st.markdown("---")
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

    st.markdown("---")
    st.caption("Alpha Builders Official Portal\nDiseño Liquid Glass")

# ==========================================
# 5. PANEL PRINCIPAL / DASHBOARD
# ==========================================
st.markdown(
    f"""
    <div class="glass-card-official">
        <h1>ALPHA BUILDERS</h1>
        <p>Panel de Control | Usuario: <b>{user_nombre}</b> ({user_cargo})</p>
    </div>
""",
    unsafe_allow_html=True,
)

usr_chks = len(st.session_state.db_checklists.get(user_email, []))
usr_rnds = len(st.session_state.db_rendimientos.get(user_email, []))

k1, k2, k3 = st.columns(3)
with k1:
    st.markdown(
        '<div class="kpi-card-official"><div class="kpi-value-official">28</div><div class="kpi-label-official">Obreros Activos</div></div>',
        unsafe_allow_html=True,
    )
with k2:
    st.markdown(
        f'<div class="kpi-card-official"><div class="kpi-value-official">{usr_chks}</div><div class="kpi-label-official">Checklists Guardados</div></div>',
        unsafe_allow_html=True,
    )
with k3:
    st.markdown(
        f'<div class="kpi-card-official"><div class="kpi-value-official">{usr_rnds}</div><div class="kpi-label-official">Reportes de Rendimiento</div></div>',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

tab_chk, tab_rend = st.tabs(
    ["📋 Checklist Diario", "📊 Control de Rendimiento por Trabajador"]
)

# ==========================================
# 6. MÓDULO 1: CHECKLIST DIARIO
# ==========================================
with tab_chk:
    st.markdown("### 📋 Check List Diario – Control de Obra")
    st.caption("Supervisión diaria de calidad y avance en el frente de trabajo.")

    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        proyecto_val = st.text_input("🏢 Proyecto:", value="Alpha Builders - Obra Central")
    with col_m2:
        residente_val = st.text_input("👷 Responsable:", value=user_nombre)
    with col_m3:
        fecha_val = st.date_input("📅 Fecha:", datetime.date.today())

    st.markdown("---")

    with st.form("form_checklist_official"):
        st.markdown("#### 🌅 Jornada de la Mañana")
        resp_manana = []

        for idx, act in enumerate(ACTIVIDADES_MANANA, 1):
            st.markdown(f"**N° {idx}. {act}**")
            c_sel, c_obs, c_foto = st.columns([2, 3, 3])

            with c_sel:
                est = st.radio("Estado", ["✓ Cumple", "✗ No Cumple", "N/A"], key=f"m_st_{idx}", horizontal=True)
            with c_obs:
                ob = st.text_input("Observación", key=f"m_ob_{idx}", placeholder="Observaciones...", label_visibility="collapsed")
            with c_foto:
                ft = st.file_uploader("📷 Foto (Opcional)", type=["jpg", "jpeg", "png"], key=f"m_ft_{idx}")

            st.markdown("<hr style='margin: 8px 0; border-color: rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
            resp_manana.append({"Jornada": "Mañana", "N°": idx, "Actividad": act, "Estado": est, "Observaciones": ob, "Foto_Objeto": ft, "Foto_Adjunta": "Sí 📷" if ft is not None else "No"})

        st.markdown("#### 🌆 Jornada de la Tarde")
        resp_tarde = []

        for idx, act in enumerate(ACTIVIDADES_TARDE, 1):
            st.markdown(f"**N° {idx}. {act}**")
            c_sel, c_obs, c_foto = st.columns([2, 3, 3])

            with c_sel:
                est = st.radio("Estado", ["✓ Cumple", "✗ No Cumple", "N/A"], key=f"t_st_{idx}", horizontal=True)
            with c_obs:
                ob = st.text_input("Observación", key=f"t_ob_{idx}", placeholder="Observaciones...", label_visibility="collapsed")
            with c_foto:
                ft = st.file_uploader("📷 Foto (Opcional)", type=["jpg", "jpeg", "png"], key=f"t_ft_{idx}")

            st.markdown("<hr style='margin: 8px 0; border-color: rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
            resp_tarde.append({"Jornada": "Tarde", "N°": idx, "Actividad": act, "Estado": est, "Observaciones": ob, "Foto_Objeto": ft, "Foto_Adjunta": "Sí 📷" if ft is not None else "No"})

        btn_guardar_chk = st.form_submit_button("💾 Guardar Checklist Diario", type="primary")

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

        st.success(f"✅ Checklist guardado correctamente para el usuario **{user_email}**.")
        st.dataframe(df_chk_save.drop(columns=["Foto_Objeto"]), use_container_width=True)

        st.markdown("#### 🖼️ Evidencia Fotográfica Capturada")
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
        st.download_button(label="📥 Descargar Checklist Diario (CSV)", data=csv_chk, file_name=f"Checklist_{fecha_val.strftime('%Y%m%d')}_{user_email}.csv", mime="text/csv")

# ==========================================
# 7. MÓDULO 2: CONTROL DE RENDIMIENTO
# ==========================================
with tab_rend:
    st.markdown("### 📊 Control de Rendimiento por Trabajador")
    st.caption("Seleccione al operario, indique el rubro y registre sus horas ejecutadas.")

    col1, col2 = st.columns(2)
    with col1:
        nombres_obreros = [t["nombre"] for t in TRABAJADORES_NO_MINA]
        trabajador_sel = st.selectbox("👷 Seleccionar Trabajador (28 Activos):", nombres_obreros)
        cargo_actual = next(t["cargo"] for t in TRABAJADORES_NO_MINA if t["nombre"] == trabajador_sel)
        st.info(f"**Cargo en obra:** {cargo_actual}")

    with col2:
        rubros_opciones = ["Enlucidos", "Fijos", "Fajas", "Dinteles"]
        rubro_sel = st.selectbox("🧱 Seleccionar Rubro:", rubros_opciones)
        unidad_medida = UNIDADES_RUBRO[rubro_sel]
        st.caption(f"Unidad de medida: **{unidad_medida}**")

    st.markdown("---")
    st.markdown("#### 🕒 Horarios Trabajados (Intervalos)")

    col_h1, col_h2, col_h3 = st.columns(3)
    with col_h1:
        h1 = st.checkbox("07:00 - 10:00 (3 Horas HH)")
    with col_h2:
        h2 = st.checkbox("10:00 - 13:00 (3 Horas HH)")
    with col_h3:
        h3 = st.checkbox("14:00 - 16:00 (2 Horas HH)")

    horas_acumuladas = (3.0 if h1 else 0.0) + (3.0 if h2 else 0.0) + (2.0 if h3 else 0.0)
    st.markdown(f"⏱️ **Total Horas-Hombre:** `{horas_acumuladas} HH`")

    st.markdown("#### 📐 Avance Ejecutado")
    avance_cant = st.number_input(f"Cantidad ejecutada ({unidad_medida}):", min_value=0.0, step=0.1, format="%.2f")

    if st.button("💾 Registrar Rendimiento", type="primary"):
        if horas_acumuladas == 0:
            st.warning("⚠️ Seleccione al menos un intervalo de horario.")
        elif avance_cant <= 0:
            st.warning("⚠️ Ingrese un avance mayor a 0.")
        else:
            rend_real = round(horas_acumuladas / avance_cant, 3)
            rend_teorico = RENDIMIENTOS_TEORICOS.get(rubro_sel, 1.0)
            estado_diag = "EFICIENTE ✅" if rend_real <= rend_teorico else "EXCESO DE HH ⚠️"

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
            st.success(f"¡Rendimiento registrado correctamente para **{trabajador_sel}**!")

    st.markdown("---")
    st.markdown("### 📜 Registros de Rendimiento Guardados")

    mis_rendimientos = st.session_state.db_rendimientos.get(user_email, [])
    if len(mis_rendimientos) > 0:
        df_mis_r = pd.DataFrame(mis_rendimientos)
        st.dataframe(df_mis_r, use_container_width=True)

        csv_r = df_mis_r.to_csv(index=False).encode("utf-8")
        st.download_button(label="📥 Descargar Rendimientos (CSV)", data=csv_r, file_name=f"Rendimientos_{user_email}.csv", mime="text/csv")
    else:
        st.info("Aún no existen registros en su historial.")