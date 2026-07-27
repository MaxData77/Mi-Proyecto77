import streamlit as st
import pandas as pd
import numpy as np

st.title("Mi primera App con Gráficos y Tablas")
st.write("Esta aplicación incluye componentes interactivos avanzados.")

nombre = st.text_input("¿Cómo te llamas?")
if nombre:
    st.success(f"¡Hola, {nombre}! Bienvenido a tu app de Streamlit.")

st.header("📊 Tabla de Datos de Ejemplo")
datos = pd.DataFrame(
    np.random.randn(20, 3),
    columns=['Ventas', 'Costos', 'Ganancias']
)
st.dataframe(datos) 

st.header("📈 Gráfico de Rendimiento")
st.line_chart(datos)

st.header("⚙️ Acciones")
if st.button("Calcular Resumen Estadístico"):
    st.write("Aquí tienes un resumen de tus datos:")
    st.write(datos.describe())
