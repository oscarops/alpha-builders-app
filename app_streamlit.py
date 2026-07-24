import pandas as pd
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Alpha Builders - Control de Rendimiento por Trabajador",
    page_icon="🏗️",
    layout="wide",
)

st.title("🏗️ ALPHA BUILDERS")
st.subheader("Registro Diario y Control de Rendimiento por Trabajador")

# Inicializar almacenamiento de registros en la sesión de Streamlit
if "registros" not in st.session_state:
    st.session_state.registros = []

# --- BASE DE DATOS REAL DE TRABAJADORES ACTIVOS (28 TRABAJADORES) ---
TRABAJADORES = [
    {
        "nombre": "ACHINA AGUAGUIÑA BYRON ALEXANDER",
        "cargo": "BODEGA",
    },
    {
        "nombre": "AGUALONGO PILAMUNGA LUIS LENIN",
        "cargo": "GYPSERO/ALBAÑIL",
    },
    {
        "nombre": "ALTAMIRANO GUALAN WILLIAM PATRICIO",
        "cargo": "GYPSERO",
    },
    {
        "nombre": "BUNSHI CAYANCELA SANTIAGO EFRAIN",
        "cargo": "ALBAÑIL",
    },
    {
        "nombre": "CAYAMBE SANDOVAL LUIS ANTONIO",
        "cargo": "ALBAÑIL",
    },
    {
        "nombre": "CUASCOTA INLAGO JOSE LIZARDO",
        "cargo": "ALBAÑIL",
    },
    {
        "nombre": "CUERO BAMONTES DEIBINZON ESTALIN",
        "cargo": "AYUDANTE",
    },
    {
        "nombre": "GUANOLUISA VACA LUIS FERNANDO",
        "cargo": "ALBAÑIL",
    },
    {
        "nombre": "LLUGLLUNA FARINANGO SEGUNDO MANUEL",
        "cargo": "ALBAÑIL",
    },
    {
        "nombre": "MORALES OTUNA VERONICA JAQUELINE",
        "cargo": "AYUDANTE",
    },
    {
        "nombre": "OCHOA MORAN MIGUEL BERNARDO",
        "cargo": "GYPSERO",
    },
    {
        "nombre": "PAGUAY RAMOS DILAN ANDRES",
        "cargo": "GYPSERO",
    },
    {
        "nombre": "ROMERO ANDRANGO LUIS ENRIQUE",
        "cargo": "GYPSERO",
    },
    {
        "nombre": "SANGUCHO FONSECA EDGAR XAVIER",
        "cargo": "ALBAÑIL",
    },
    {
        "nombre": "TARAPUES MONARCO CARLOS ANDRES",
        "cargo": "GYPSERO",
    },
    {
        "nombre": "TONATO TACO LUIS EUCLIDES",
        "cargo": "ALBAÑIL",
    },
    {
        "nombre": "TOSCANO ALTAMIRANO JEREMMY WENDLEY",
        "cargo": "AYUDANTE",
    },
    {
        "nombre": "TRONCOSO COBEÑA CRISTOPHER GEOVANNY",
        "cargo": "AYUDANTE",
    },
    {
        "nombre": "TUTASI CASILLAS JORGE GEOVANI",
        "cargo": "FIERRERO",
    },
    {
        "nombre": "CHAVEZ GUITARRA JOSE GREGORIO",
        "cargo": "GYPSERO",
    },
    {
        "nombre": "CORDOVA FLORES ERICK DARIO",
        "cargo": "GYPSERO / AYUDANTE",
    },
    {
        "nombre": "CABRERA CAMPO ANNDY JEREMIAS",
        "cargo": "GYPSERO / OPERADOR",
    },
    {"nombre": "CHELA OCHOA RAUL", "cargo": "GYPSERO/ALBAÑIL"},
    {
        "nombre": "SEMBLANTES TIPANLUISA JAVIER PATRICIO",
        "cargo": "GYPSERO/ALBAÑIL",
    },
    {
        "nombre": "FUEREZ COYAGO JOSE SANTOS",
        "cargo": "HERRAMIENTAS",
    },
    {
        "nombre": "ALTAMIRANO CORDOVA HECTOR LUIS",
        "cargo": "PINTOR",
    },
    {
        "nombre": "ACOSTA AGUILAR JORGE PATRICIO",
        "cargo": "SOLDADOR",
    },
    {
        "nombre": "TARAPUES CASTRO JOAO ALEXANDER",
        "cargo": "SOLDADOR",
    },
]

# Unidades por defecto para cada rubro
UNIDADES_RUBRO = {
    "Enlucidos": "m2",
    "Fijos": "m2",
    "Fajas": "m",
    "Dinteles": "m",
}

# Rendimientos teóricos de referencia (HH/Unid)
RENDIMIENTOS_TEORICOS = {
    "Enlucidos": 0.75,
    "Fijos": 0.50,
    "Fajas": 0.30,
    "Dinteles": 0.40,
}

# --- FORMULARIO DE REGISTRO ---
st.markdown("---")
st.markdown("### 📋 Formulario de Registro de Actividad")

col1, col2 = st.columns(2)

with col1:
    # 1. Selección de Trabajador
    nombres = [t["nombre"] for t in TRABAJADORES]
    trabajador_sel = st.selectbox(
        "👷 Seleccionar Trabajador (28 Activos):", nombres
    )

    # Mostrar cargo correspondiente al lado
    cargo_actual = next(
        t["cargo"] for t in TRABAJADORES if t["nombre"] == trabajador_sel
    )
    st.info(f"**Cargo:** {cargo_actual}")

with col2:
    # 2. Selección de Rubro
    rubros_disponibles = ["Enlucidos", "Fijos", "Fajas", "Dinteles"]
    rubro_sel = st.selectbox("🧱 Seleccionar Rubro:", rubros_disponibles)
    unidad = UNIDADES_RUBRO[rubro_sel]
    st.caption(f"Unidad de medida: **{unidad}**")

# 3. Selección de Horarios e Intervalos de Trabajo
st.markdown("#### 🕒 Horarios de Trabajo Ejecutados")
col_int1, col_int2, col_int3 = st.columns(3)

with col_int1:
    h1 = st.checkbox("07:00 - 10:00 (3 Horas)")
with col_int2:
    h2 = st.checkbox("10:00 - 13:00 (3 Horas)")
with col_int3:
    h3 = st.checkbox("14:00 - 16:00 (2 Horas)")

# Calcular total de horas ejecutadas según los intervalos marcados
horas_totales = 0.0
if h1:
    horas_totales += 3.0
if h2:
    horas_totales += 3.0
if h3:
    horas_totales += 2.0

st.write(f"⏱️ **Horas Trabajadas Acumuladas:** `{horas_totales} HH`")

# 4. Ingrese el Avance de Obra
st.markdown("#### 📐 Avance Realizado")
avance = st.number_input(
    f"Ingrese la cantidad ejecutada ({unidad}):",
    min_value=0.0,
    step=0.1,
    format="%.2f",
)

# Botón para Registrar
if st.button("💾 Registrar Actividad", type="primary"):
    if horas_totales == 0:
        st.warning("⚠️ Debe seleccionar al menos un intervalo de horario.")
    elif avance <= 0:
        st.warning("⚠️ Por favor ingrese un avance mayor a 0.")
    else:
        # Cálculos de rendimiento
        rend_real = round(horas_totales / avance, 3)
        rend_teorico = RENDIMIENTOS_TEORICOS.get(rubro_sel, 1.0)
        estado = (
            "EFICIENTE ✅" if rend_real <= rend_teorico else "EXCESO DE HH ⚠️"
        )

        # Agregar el registro
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

# --- TABLA Y RESUMEN DE REGISTROS ---
st.markdown("---")
st.markdown("### 📊 Historial de Registros de Obra")

if len(st.session_state.registros) > 0:
    df_registros = pd.DataFrame(st.session_state.registros)
    st.dataframe(df_registros, use_container_width=True)

    # Botón para descargar reporte
    csv = df_registros.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Descargar Reporte en CSV",
        data=csv,
        file_name="Reporte_Productividad_AlphaBuilders.csv",
        mime="text/csv",
    )
else:
    st.info(
        "Aún no hay registros guardados. Utiliza el formulario superior para añadir avances."
    )