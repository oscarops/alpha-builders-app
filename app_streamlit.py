import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Alpha Builders - Control de Rendimiento",
    page_icon="🏗️",
    layout="wide",
)

st.title("🏗️ ALPHA BUILDERS")
st.subheader("Control de Rendimiento de Mano de Obra")
st.write(
    "Ingresa o modifica las actividades, cantidades ejecutadas y horas-hombre para calcular los rendimientos."
)

# Datos iniciales
datos_iniciales = [
    {
        "Actividad / Rubro": "Excavación manual",
        "Unidad": "m3",
        "Cant. Ejecutada": 12.5,
        "Horas-Hombre (HH)": 10.0,
        "Rend. Teórico (HH/Unid)": 0.85,
    },
    {
        "Actividad / Rubro": "Hormigón en zapatas",
        "Unidad": "m3",
        "Cant. Ejecutada": 8.0,
        "Horas-Hombre (HH)": 16.0,
        "Rend. Teórico (HH/Unid)": 1.80,
    },
    {
        "Actividad / Rubro": "Mampostería de ladrillo",
        "Unidad": "m2",
        "Cant. Ejecutada": 25.0,
        "Horas-Hombre (HH)": 30.0,
        "Rend. Teórico (HH/Unid)": 1.10,
    },
    {
        "Actividad / Rubro": "Enlucido interior",
        "Unidad": "m2",
        "Cant. Ejecutada": 40.0,
        "Horas-Hombre (HH)": 28.0,
        "Rend. Teórico (HH/Unid)": 0.75,
    },
]

# Crear DataFrame inicial
df_inicial = pd.DataFrame(datos_iniciales)

# Tabla editable e interactiva en Streamlit
st.subheader("Tabla de Datos Interactiva")
df_editado = st.data_editor(
    df_inicial,
    num_rows="dynamic",  # Permite agregar/eliminar filas fácilmente
    use_container_width=True,
)

# Botón para realizar el cálculo
if st.button("⚡ Calcular Rendimientos", type="primary"):
    # Realizar cálculos
    def calcular_rend_real(row):
        cant = row["Cant. Ejecutada"]
        hh = row["Horas-Hombre (HH)"]
        return round(hh / cant, 3) if cant > 0 else 0.0

    def evaluar_estado(row):
        real = row["Rend. Real (HH/Unid)"]
        teorico = row["Rend. Teórico (HH/Unid)"]
        if real <= 0:
            return "INCOMPLETO"
        return "EFICIENTE ✅" if real <= teorico else "EXCESO DE HH ⚠️"

    df_editado["Rend. Real (HH/Unid)"] = df_editado.apply(
        calcular_rend_real, axis=1
    )
    df_editado["Estado / Diagnóstico"] = df_editado.apply(
        evaluar_estado, axis=1
    )

    st.success("¡Cálculos actualizados con éxito!")
    st.dataframe(df_editado, use_container_width=True)

    # Opción para descargar los resultados en CSV
    csv = df_editado.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Descargar Reporte en CSV",
        data=csv,
        file_name="Reporte_Rendimientos_AlphaBuilders.csv",
        mime="text/csv",
    )