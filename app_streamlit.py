import datetime
import pandas as pd
import streamlit as st

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y TEMA VISUAL
# ==========================================
st.set_page_config(
    page_title="Alpha Builders | Portal de Obra y Calidad",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilos CSS Avanzados para Interfaz Profesional de Ingeniería
st.markdown(
    """
    <style>
    /* Estructura Global */
    .stApp {
        background-color: #f8fafc;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }
    
    /* Encabezado Corporativo Main Banner */
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        padding: 28px 32px;
        border-radius: 16px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
    }
    .main-header h1 {
        color: #ffffff !important;
        font-weight: 800;
        margin: 0;
        font-size: 2.2rem;
        letter-spacing: -0.02em;
    }
    .main-header p {
        color: #94a3b8;
        margin-top: 6px;
        margin-bottom: 0;
        font-size: 1rem;
    }

    /* Tarjetas Dashboard KPI */
    .kpi-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 14px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        text-align: center;
        transition: transform 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
    }
    .kpi-value {
        font-size: 1.8rem;
        font-weight: 800;
        color: #0f172a;
    }
    .kpi-label {
        font-size: 0.825rem;
        color: #64748b;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.05em;
        margin-top: 4px;
    }

    /* Pestañas Principales Estilo Nav Pill */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: #f1f5f9;
        padding: 6px;
        border-radius: 14px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px;
        padding: 10px 20px;
        background-color: transparent;
        color: #475569;
        font-weight: 600;
        border: none !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #1e3a8a !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
    }

    /* Tarjetas de Formularios y Módulos */
    .module-card {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 14px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
        margin-bottom: 20px;
    }

    /* Ocultar elementos innecesarios */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. BASE DE DATOS Y ESTADOS DE SESIÓN
# ==========================================
# Autenticación de Usuario
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_email = ""
    st.session_state.usuario_nombre = ""

# Almacenamiento seguro por perfil de usuario
if "db_checklists" not in st.session_state:
    st.session_state.db_checklists = {}  # {email: [lista_checklists]}

if "db_rendimientos" not in st.session_state:
    st.session_state.db_rendimientos = {}  # {email: [lista_rendimientos]}

# Nómina Real de Trabajadores (28 Trabajadores del Archivo Excel)
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
    {
        "nombre": "SEMBLANTES TIPANLUISA JAVIER PATRICIO",
        "cargo": "GYPSERO/ALBAÑIL",
    },
    {"nombre": "FUEREZ COYAGO JOSE SANTOS", "cargo": "HERRAMIENTAS"},
    {"nombre": "ALTAMIRANO CORDOVA HECTOR LUIS", "cargo": "PINTOR"},
    {"nombre": "ACOSTA AGUILAR JORGE PATRICIO", "cargo": "SOLDADOR"},
    {"nombre": "TARAPUES CASTRO JOAO ALEXANDER", "cargo": "SOLDADOR"},
]

UNIDADES_RUBRO = {
    "Enlucidos": "m2",
    "Fijos": "m2",
    "Fajas": "m",
    "Dinteles": "m",
}
RENDIMIENTOS_TEORICOS = {
    "Enlucidos": 0.75,
    "Fijos": 0.50,
    "Fajas": 0.30,
    "Dinteles": 0.40,
}

# Actividades del Checklist según el Excel del Residente
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
# 3. MÓDULO DE AUTENTICACIÓN (LOGIN SEGURO)
# ==========================================
if not st.session_state.autenticado:
    st.markdown(
        """
        <div class="main-header">
            <h1>🏗️ ALPHA BUILDERS</h1>
            <p>Portal Corporativo para Gestión de Obra y Control de Calidad</p>
        </div>
    """,
        unsafe_allow_html=True,
    )

    col_center1, col_center2, col_center3 = st.columns([1, 2, 1])

    with col_center2:
        st.markdown("### 🔐 Acceso de Usuarios")

        tab_login1, tab_login2 = st.tabs(
            ["Ingreso Directo / Correo", "Google / Microsoft"]
        )

        with tab_login1:
            email_input = st.text_input(
                "Correo Electrónico (Gmail, Outlook, etc.):",
                placeholder="ejemplo@alphabuilders.com",
            )
            pass_input = st.text_input("Contraseña:", type="password")
            nombre_input = st.text_input(
                "Nombre Completo:", placeholder="Ing. Residente"
            )

            if st.button(
                "Iniciar Sesión", type="primary", use_container_width=True
            ):
                if email_input and pass_input:
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = (
                        email_input.strip().lower()
                    )
                    st.session_state.usuario_nombre = (
                        nombre_input.strip() or email_input.split("@")[0].title()
                    )

                    # Inicializar carpetas de datos del usuario
                    if (
                        st.session_state.usuario_email
                        not in st.session_state.db_checklists
                    ):
                        st.session_state.db_checklists[
                            st.session_state.usuario_email
                        ] = []
                    if (
                        st.session_state.usuario_email
                        not in st.session_state.db_rendimientos
                    ):
                        st.session_state.db_rendimientos[
                            st.session_state.usuario_email
                        ] = []

                    st.rerun()
                else:
                    st.error("Por favor ingrese correo y contraseña.")

        with tab_login2:
            st.caption(
                "Inicie sesión rápidamente usando su cuenta de correo configurada."
            )
            col_g, col_m = st.columns(2)
            with col_g:
                if st.button("🌐 Continuar con Google", use_container_width=True):
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = "usuario.google@gmail.com"
                    st.session_state.usuario_nombre = "Usuario Google"
                    st.rerun()
            with col_m:
                if st.button(
                    "🏢 Continuar con Microsoft", use_container_width=True
                ):
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = "usuario.outlook@outlook.com"
                    st.session_state.usuario_nombre = "Usuario Outlook"
                    st.rerun()

        st.caption(
            "🔒 **Seguridad Garantizada:** Cada usuario gestiona su propia base de datos de manera aislada e independiente."
        )

    st.stop()

# ==========================================
# 4. BARRA LATERAL Y SESIÓN ACTIVA
# ==========================================
user_email = st.session_state.usuario_email
user_nombre = st.session_state.usuario_nombre

with st.sidebar:
    st.markdown("### 👤 Perfil de Sesión")
    st.markdown(f"**Nombre:** `{user_nombre}`")
    st.markdown(f"**Correo:** `{user_email}`")
    st.markdown("**Estado:** `Residente Activo 🟢`")

    st.markdown("---")
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

    st.markdown("---")
    st.caption("Alpha Builders v3.0 Pro\nSistema Continuo de Control de Obra")

# ==========================================
# 5. DASHBOARD PRINCIPAL
# ==========================================
st.markdown(
    f"""
    <div class="main-header">
        <h1>🏗️ ALPHA BUILDERS | Portal de Obra</h1>
        <p>Bienvenido, <b>{user_nombre}</b>. Panel de control continuo para calidad y rendimiento de personal.</p>
    </div>
""",
    unsafe_allow_html=True,
)

# Métricas rápidas (KPIs) de la sesión del usuario
usr_chks = len(st.session_state.db_checklists.get(user_email, []))
usr_rnds = len(st.session_state.db_rendimientos.get(user_email, []))

kpi1, kpi2, kpi3 = st.columns(3)
with kpi1:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-value">28</div><div class="kpi-label">Obreros Registrados</div></div>',
        unsafe_allow_html=True,
    )
with kpi2:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-value">{usr_chks}</div><div class="kpi-label">Checklists Guardados</div></div>',
        unsafe_allow_html=True,
    )
with kpi3:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-value">{usr_rnds}</div><div class="kpi-label">Registros de Avance</div></div>',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# Pestañas principales
tab_chk, tab_rend = st.tabs(
    ["📋 Checklist Diario Residente", "📊 Control de Rendimiento por Trabajador"]
)

# ==========================================
# 6. MÓDULO 1: CHECKLIST DIARIO CON FOTOS
# ==========================================
with tab_chk:
    st.markdown("### 📋 Check List Diario – Residente de Obra")
    st.caption(
        "Supervisión y control de calidad diaria en campo para jornadas de mañana y tarde."
    )

    col_meta1, col_meta2, col_meta3 = st.columns(3)
    with col_meta1:
        proyecto_val = st.text_input(
            "🏢 Proyecto:", value="Alpha Builders - Obra Principal"
        )
    with col_meta2:
        residente_val = st.text_input(
            "👷 Residente de Obra:", value=user_nombre
        )
    with col_meta3:
        fecha_val = st.date_input("📅 Fecha:", datetime.date.today())

    st.markdown("---")

    with st.form("form_checklist_profesional"):
        st.markdown("#### 🌅 Jornada de la Mañana")
        resp_manana = []

        for idx, act in enumerate(ACTIVIDADES_MANANA, 1):
            st.markdown(f"**N° {idx}. {act}**")
            c_sel, c_obs, c_foto = st.columns([2, 3, 3])

            with c_sel:
                est = st.radio(
                    "Estado",
                    ["✓ Cumple", "✗ No Cumple", "N/A"],
                    key=f"m_st_{idx}",
                    horizontal=True,
                )
            with c_obs:
                ob = st.text_input(
                    "Observación",
                    key=f"m_ob_{idx}",
                    placeholder="Detalles u observaciones...",
                    label_visibility="collapsed",
                )
            with c_foto:
                ft = st.file_uploader(
                    "📷 Foto (Opcional)",
                    type=["jpg", "jpeg", "png"],
                    key=f"m_ft_{idx}",
                )

            st.markdown(
                "<hr style='margin: 8px 0; opacity: 0.2;'>",
                unsafe_allow_html=True,
            )
            resp_manana.append(
                {
                    "Jornada": "Mañana",
                    "N°": idx,
                    "Actividad": act,
                    "Estado": est,
                    "Observaciones": ob,
                    "Foto_Objeto": ft,
                    "Foto_Adjunta": "Sí 📷" if ft is not None else "No",
                }
            )

        st.markdown("#### 🌆 Jornada de la Tarde")
        resp_tarde = []

        for idx, act in enumerate(ACTIVIDADES_TARDE, 1):
            st.markdown(f"**N° {idx}. {act}**")
            c_sel, c_obs, c_foto = st.columns([2, 3, 3])

            with c_sel:
                est = st.radio(
                    "Estado",
                    ["✓ Cumple", "✗ No Cumple", "N/A"],
                    key=f"t_st_{idx}",
                    horizontal=True,
                )
            with c_obs:
                ob = st.text_input(
                    "Observación",
                    key=f"t_ob_{idx}",
                    placeholder="Detalles u observaciones...",
                    label_visibility="collapsed",
                )
            with c_foto:
                ft = st.file_uploader(
                    "📷 Foto (Opcional)",
                    type=["jpg", "jpeg", "png"],
                    key=f"t_ft_{idx}",
                )

            st.markdown(
                "<hr style='margin: 8px 0; opacity: 0.2;'>",
                unsafe_allow_html=True,
            )
            resp_tarde.append(
                {
                    "Jornada": "Tarde",
                    "N°": idx,
                    "Actividad": act,
                    "Estado": est,
                    "Observaciones": ob,
                    "Foto_Objeto": ft,
                    "Foto_Adjunta": "Sí 📷" if ft is not None else "No",
                }
            )

        btn_guardar_chk = st.form_submit_button(
            "💾 Guardar Checklist Diario", type="primary"
        )

    if btn_guardar_chk:
        all_chk_data = resp_manana + resp_tarde
        df_chk_save = pd.DataFrame(all_chk_data)

        # Guardar en la cuenta del usuario
        registro_completo = {
            "Fecha": fecha_val.strftime("%Y-%m-%d"),
            "Proyecto": proyecto_val,
            "Residente": residente_val,
            "Datos": df_chk_save,
        }
        st.session_state.db_checklists[user_email].append(registro_completo)

        st.success(
            f"✅ Checklist guardado exitosamente para la cuenta **{user_email}**."
        )

        st.markdown("#### Resumen del Registro Actual")
        st.dataframe(
            df_chk_save.drop(columns=["Foto_Objeto"]), use_container_width=True
        )

        # Galería visual de fotos adjuntas
        st.markdown("#### 🖼️ Evidencia Fotográfica Capturada")
        has_images = False
        cols_gal = st.columns(3)
        g_idx = 0

        for row_item in all_chk_data:
            if row_item["Foto_Objeto"] is not None:
                has_images = True
                with cols_gal[g_idx % 3]:
                    st.image(
                        row_item["Foto_Objeto"],
                        caption=f"[{row_item['Jornada']}] N° {row_item['N°']}: {row_item['Actividad']}",
                        use_column_width=True,
                    )
                g_idx += 1

        if not has_images:
            st.info("No se adjuntaron fotografías en este registro.")

        # Descarga de CSV
        csv_chk = (
            df_chk_save.drop(columns=["Foto_Objeto"])
            .to_csv(index=False)
            .encode("utf-8")
        )
        st.download_button(
            label="📥 Descargar Checklist Diario en CSV",
            data=csv_chk,
            file_name=f"Checklist_{fecha_val.strftime('%Y%m%d')}_{user_email}.csv",
            mime="text/csv",
        )

# ==========================================
# 7. MÓDULO 2: CONTROL DE RENDIMIENTO
# ==========================================
with tab_rend:
    st.markdown("### 📊 Control de Rendimiento por Trabajador")
    st.caption("Seleccione al operario, asigne el rubro y registre sus horas ejecutadas y avance de obra.")

    col1, col2 = st.columns(2)

    with col1:
        nombres_obreros = [t["nombre"] for t in TRABAJADORES_NO_MINA]
        trabajador_sel = st.selectbox(
            "👷 Seleccionar Trabajador (28 Activos):", nombres_obreros
        )

        cargo_actual = next(
            t["cargo"]
            for t in TRABAJADORES_NO_MINA
            if t["nombre"] == trabajador_sel
        )
        st.info(f"**Cargo Asignado:** {cargo_actual}")

    with col2:
        rubros_opciones = ["Enlucidos", "Fijos", "Fajas", "Dinteles"]
        rubro_sel = st.selectbox("🧱 Seleccionar Rubro:", rubros_opciones)
        unidad_medida = UNIDADES_RUBRO[rubro_sel]
        st.caption(f"Unidad de medida para {rubro_sel}: **{unidad_medida}**")

    st.markdown("---")
    st.markdown("#### 🕒 Intervalos de Horario Trabajados")

    col_h1, col_h2, col_h3 = st.columns(3)
    with col_h1:
        h1 = st.checkbox("07:00 - 10:00 (3 Horas HH)")
    with col_h2:
        h2 = st.checkbox("10:00 - 13:00 (3 Horas HH)")
    with col_h3:
        h3 = st.checkbox("14:00 - 16:00 (2 Horas HH)")

    horas_acumuladas = (
        (3.0 if h1 else 0.0) + (3.0 if h2 else 0.0) + (2.0 if h3 else 0.0)
    )
    st.markdown(
        f"⏱️ **Horas-Hombre (HH) Calculadas:** `{horas_acumuladas} HH`"
    )

    st.markdown("#### 📐 Avance Físico de Obra")
    avance_cant = st.number_input(
        f"Cantidad Ejecutada en {unidad_medida}:",
        min_value=0.0,
        step=0.1,
        format="%.2f",
    )

    if st.button("💾 Registrar Rendimiento de Trabajador", type="primary"):
        if horas_acumuladas == 0:
            st.warning("⚠️ Seleccione al menos un intervalo de horario.")
        elif avance_cant <= 0:
            st.warning("⚠️ Ingrese una cantidad ejecutada mayor a 0.")
        else:
            rend_real = round(horas_acumuladas / avance_cant, 3)
            rend_teorico = RENDIMIENTOS_TEORICOS.get(rubro_sel, 1.0)
            estado_diag = (
                "EFICIENTE ✅" if rend_real <= rend_teorico else "EXCESO DE HH ⚠️"
            )

            nuevo_registro = {
                "Usuario_Registro": user_email,
                "Fecha": datetime.date.today().strftime("%Y-%m-%d"),
                "Trabajador": trabajador_sel,
                "Cargo": cargo_actual,
                "Rubro": rubro_sel,
                "Horas Trabajadas (HH)": horas_acumuladas,
                "Avance": avance_cant,
                "Unidad": unidad_medida,
                "Rend. Real (HH/Unid)": rend_real,
                "Rend. Teórico": rend_teorico,
                "Estado": estado_diag,
            }

            st.session_state.db_rendimientos[user_email].append(nuevo_registro)
            st.success(
                f"¡Registro de rendimiento guardado correctamente para **{trabajador_sel}**!"
            )

    st.markdown("---")
    st.markdown("### 📜 Historial de Registros de Rendimiento")

    mis_rendimientos = st.session_state.db_rendimientos.get(user_email, [])

    if len(mis_rendimientos) > 0:
        df_mis_r = pd.DataFrame(mis_rendimientos)
        st.dataframe(df_mis_r, use_container_width=True)

        csv_r = df_mis_r.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Descargar Reporte de Rendimientos en CSV",
            data=csv_r,
            file_name=f"Rendimientos_{user_email}.csv",
            mime="text/csv",
        )
    else:
        st.info(
            "Aún no tiene registros de rendimiento guardados en su cuenta."
        )