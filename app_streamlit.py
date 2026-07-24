import streamlit as st
import pandas as pd
from datetime import datetime

# Configuración de página adaptable
st.set_page_config(
    page_title="Alpha Builders - Control de Rendimiento",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilo personalizado para el encabezado y botones
st.markdown("""
    <style>
    .stApp { background-color: #f3f4f6; }
    .header-title { font-size: 24px; font-weight: bold; color: #111827; }
    .header-sub { font-size: 14px; color: #4b5563; margin-bottom: 20px; }
    div.stButton > button {
        background-color: #000000 !important;
        color: #ffffff !important;
        font-weight: bold;
        border-radius: 8px;
        padding: 12px;
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# Encabezado Corporativo
st.markdown("<div class='header-title'>ALPHA BUILDERS</div>", unsafe_allow_html=True)
st.markdown("<div class='header-sub'>Cazadores de Inversiones • Control de Rendimiento de Obra</div>", unsafe_allow_html=True)
st.divider()

# Formulario de entrada
with st.container():
    col1, col2 = st.columns(2)
    with col1:
        actividad = st.text_input("Rubro / Actividad", placeholder="Ej. Mampostería").strip().upper()
        unidad = st.text_input("Unidad", placeholder="m2, m3, kg").strip().lower()
    with col2:
        horario = st.text_input("Intervalo de Tiempo", placeholder="08:00 - 16:00")
        apu = st.number_input("Rendimiento Diario APU Planificado", min_value=0.0, step=1.0)

    st.subheader("👥 Cuadrilla en Frente")
    c_m, c_o, c_p = st.columns(3)
    with c_m:
        maestros = st.number_input("Maestros", min_value=0, step=1)
    with c_o:
        obreros = st.number_input("Obreros / Oficiales", min_value=0, step=1)
    with c_p:
        peones = st.number_input("Peones / Ayudantes", min_value=0, step=1)

    st.subheader("⏱️ Distribución de Tiempos (Horas)")
    c_tp, c_tc, c_tnp = st.columns(3)
    with c_tp:
        tp = st.number_input("Trabajo Productivo (TP)", min_value=0.0, step=0.5)
    with c_tc:
        tc = st.number_input("Trabajo Contributivo (TC)", min_value=0.0, step=0.5)
    with c_tnp:
        tnp = st.number_input("Trabajo No Productivo (TNP)", min_value=0.0, step=0.5)

    st.subheader("📊 Avance Físico")
    avance = st.number_input("¿Cuánto avanzó en el frente?", min_value=0.0, step=1.0)

st.divider()

# Botón de Procesamiento
if st.button("⚡ Ejecutar Auditoría Analítica"):
    if not actividad or not unidad:
        st.warning("⚠️ Debes ingresar el nombre del 'Rubro' y la 'Unidad'.")
    else:
        personal_total = maestros + obreros + peones
        tiempo_total = tp + tc + tnp

        if personal_total == 0 or tiempo_total == 0:
            st.error("❌ Asigna al menos 1 trabajador y horas de trabajo válidas.")
        else:
            # Cálculos de Ingeniería
            hh = personal_total * tiempo_total
            rend_hora = avance / tiempo_total
            idr = (avance / apu) * 100 if apu > 0 else 0

            # Resultados en pantalla
            st.success(f"✅ Diagnóstico procesado para: **{actividad}**")
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Esfuerzo del Frente", f"{hh:.2f} H-H")
            m2.metric("Rendimiento por Hora", f"{rend_hora:.2f} {unidad}/h")
            m3.metric("Cumplimiento (IDR)", f"{idr:.0f}%" if apu > 0 else "Sin APU")

            if apu > 0:
                if idr >= 95:
                    st.info(f"🎯 **EN META:** Lograste {avance:.2f} {unidad}, cumpliendo el objetivo contractual del APU.")
                else:
                    st.error(f"💸 **RIESGO DE RETRASO:** Completaste {avance:.2f} {unidad} frente a la meta diaria de {apu:.2f} {unidad}.")