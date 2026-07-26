import datetime
import pandas as pd
import streamlit as st

# Configuración inicial de la aplicación
st.set_page_config(
    page_title="Alpha Builders - Checklist Diario Residente",
    page_icon="📋",
    layout="wide",
)

st.title("🏗️ ALPHA BUILDERS")
st.subheader("📋 Check List Diario – Residente de Obra")

# Pestañas principales para navegar entre la app y el checklist
tab_checklist, tab_rendimiento = st.tabs(
    ["📋 Checklist Residente", "📊 Control de Rendimiento por Trabajador"]
)

# ==========================================
# MÓDULO 1: CHECKLIST DIARIO RESIDENTE
# ==========================================
with tab_checklist:
    st.markdown("### Datos Generales del Control")

    col_meta1, col_meta2, col_meta3 = st.columns(3)
    with col_meta1:
        proyecto = st.text_input("🏢 Proyecto:", value="Alpha Builders - Obra Central")
    with col_meta2:
        residente = st.text_input("👷 Residente de Obra:", value="")
    with col_meta3:
        fecha = st.date_input("📅 Fecha:", datetime.date.today())

    st.markdown("---")

    # Definición de actividades según el Excel
    actividades_manana = [
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

    actividades_tarde = [
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

    # Formulario interactivo
    with st.form("form_checklist"):
        st.markdown("### 🌅 Jornada de la Mañana")
        respuestas_manana = []

        for idx, act in enumerate(actividades_manana, 1):
            c1, c2, c3 = st.columns([1, 4, 5])
            with c1:
                st.write(f"**N° {idx}**")
            with c2:
                st.write(act)
            with c3:
                col_sel, col_obs = st.columns([2, 3])
                with col_sel:
                    estado = st.radio(
                        "Estado",
                        ["✓ Cumple", "✗ No Cumple", "N/A"],
                        key=f"m_{idx}",
                        label_visibility="collapsed",
                        horizontal=True,
                    )
                with col_obs:
                    obs = st.text_input(
                        "Observación",
                        key=f"obs_m_{idx}",
                        placeholder="Observaciones...",
                        label_visibility="collapsed",
                    )
            respuestas_manana.append(
                {
                    "Jornada": "Mañana",
                    "N°": idx,
                    "Actividad": act,
                    "Estado": estado,
                    "Observaciones": obs,
                }
            )

        st.markdown("---")
        st.markdown("### 🌆 Jornada de la Tarde")
        respuestas_tarde = []

        for idx, act in enumerate(actividades_tarde, 1):
            c1, c2, c3 = st.columns([1, 4, 5])
            with c1:
                st.write(f"**N° {idx}**")
            with c2:
                st.write(act)
            with c3:
                col_sel, col_obs = st.columns([2, 3])
                with col_sel:
                    estado = st.radio(
                        "Estado",
                        ["✓ Cumple", "✗ No Cumple", "N/A"],
                        key=f"t_{idx}",
                        label_visibility="collapsed",
                        horizontal=True,
                    )
                with col_obs:
                    obs = st.text_input(
                        "Observación",
                        key=f"obs_t_{idx}",
                        placeholder="Observaciones...",
                        label_visibility="collapsed",
                    )
            respuestas_tarde.append(
                {
                    "Jornada": "Tarde",
                    "N°": idx,
                    "Actividad": act,
                    "Estado": estado,
                    "Observaciones": obs,
                }
            )

        guardar_chk = st.form_submit_button(
            "💾 Guardar y Generar Checklist", type="primary"
        )

    if guardar_chk:
        if not residente.strip():
            st.warning("⚠️ Por favor ingrese el nombre del Residente de Obra.")
        else:
            todos_datos = respuestas_manana + respuestas_tarde
            df_chk = pd.DataFrame(todos_datos)

            st.success(
                f"✅ Checklist guardado exitosamente por **{residente}** el {fecha.strftime('%d/%m/%Y')}."
            )
            st.markdown("#### Resumen del Checklist Diario")
            st.dataframe(df_chk, use_container_width=True)

            # Preparar CSV para descarga
            csv_chk = df_chk.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Descargar Checklist en CSV",
                data=csv_chk,
                file_name=f"Checklist_{fecha.strftime('%Y%m%d')}_{residente.replace(' ', '_')}.csv",
                mime="text/csv",
            )

# ==========================================
# MÓDULO 2: CONTROL DE RENDIMIENTO POR TRABAJADOR
# ==========================================
with tab_rendimiento:
    st.markdown("### Registro Diario por Trabajador")

    if "registros" not in st.session_state:
        st.session_state.registros = []

    TRABAJADORES = [
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

    col1, col2 = st.columns(2)
    with col1:
        nombres = [t["nombre"] for t in TRABAJADORES]
        trabajador_sel = st.selectbox(
            "👷 Seleccionar Trabajador (28 Activos):", nombres
        )
        cargo_actual = next(
            t["cargo"] for t in TRABAJADORES if t["nombre"] == trabajador_sel
        )
        st.info(f"**Cargo:** {cargo_actual}")

    with col2:
        rubro_sel = st.selectbox(
            "🧱 Seleccionar Rubro:", ["Enlucidos", "Fijos", "Fajas", "Dinteles"]
        )
        unidad = UNIDADES_RUBRO[rubro_sel]
        st.caption(f"Unidad de medida: **{unidad}**")

    st.markdown("#### 🕒 Horarios de Trabajo Ejecutados")
    c_int1, c_int2, c_int3 = st.columns(3)
    with c_int1:
        h1 = st.checkbox("07:00 - 10:00 (3 Horas)")
    with c_int2:
        h2 = st.checkbox("10:00 - 13:00 (3 Horas)")
    with c_int3:
        h3 = st.checkbox("14:00 - 16:00 (2 Horas)")

    horas_totales = (3.0 if h1 else 0.0) + (3.0 if h2 else 0.0) + (2.0 if h3 else 0.0)
    st.write(f"⏱️ **Horas Trabajadas Acumuladas:** `{horas_totales} HH`")

    avance = st.number_input(
        f"Ingrese la cantidad ejecutada ({unidad}):",
        min_value=0.0,
        step=0.1,
        format="%.2f",
    )

    if st.button("💾 Registrar Actividad", type="primary"):
        if horas_totales == 0:
            st.warning("⚠️ Debe seleccionar al menos un intervalo de horario.")
        elif avance <= 0:
            st.warning("⚠️ Por favor ingrese un avance mayor a 0.")
        else:
            rend_real = round(horas_totales / avance, 3)
            rend_teorico = RENDIMIENTOS_TEORICOS.get(rubro_sel, 1.0)
            estado = (
                "EFICIENTE ✅" if rend_real <= rend_teorico else "EXCESO DE HH ⚠️"
            )

            st.session_state.registros.append(
                {
                    "Trabajador": trabajador_sel,
                    "Cargo": cargo_actual,
                    "Rubro": rubro_sel,
                    "Horas Trabajadas (HH)": horas_totales,
                    "Avance": avance,
                    "Unidad": unidad,
                    "Rend. Real (HH/Unid)": rend_real,
                    "Rend. Teórico": rend_teorico,
                    "Estado": estado,
                }
            )
            st.success(
                f"¡Registro guardado exitosamente para {trabajador_sel} en {rubro_sel}!"
            )

    if len(st.session_state.registros) > 0:
        df_r = pd.DataFrame(st.session_state.registros)
        st.dataframe(df_r, use_container_width=True)
        csv_r = df_r.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Descargar Reporte en CSV",
            data=csv_r,
            file_name="Reporte_Productividad_AlphaBuilders.csv",
            mime="text/csv",
        )