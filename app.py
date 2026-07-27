import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px  # Importamos Plotly Express

# Configuración inicial de la página para aprovechar todo el ancho
st.set_page_config(page_title="Revisión de Órdenes y Detenciones", layout="wide")

# 1. TÍTULO PRINCIPAL
st.title("🛠️ Sistema de Control y Revisión de Detenciones")
st.markdown("Plataforma interactiva con analítica avanzada en Plotly.")
st.write("---")

# 2. GENERACIÓN DE DATOS SIMULADOS
@st.cache_data
def cargar_datos():
    np.random.seed(42)
    categorias = ['Mecánica', 'Eléctrica', 'Operacional', 'Falta de Suministro']
    equipos = ['Línea A', 'Línea B', 'Chancador', 'Molino 1']
    
    df = pd.DataFrame({
        'Fecha': pd.date_range(start='2026-07-01', periods=30, freq='D'),
        'Equipo': np.random.choice(equipos, 30),
        'Tipo Falla': np.random.choice(categorias, 30),
        'Horas Detención': np.random.uniform(0.5, 6.0, 30).round(1),
        'Estado': np.random.choice(['Resuelto', 'En Proceso', 'Pendiente'], 30)
    })
    return df

df_original = cargar_datos()

# 3. FILTROS EN LA BARRA LATERAL (Sidebar)
st.sidebar.header("🎯 Filtros de Búsqueda")

tipos_seleccionados = st.sidebar.multiselect(
    "Selecciona Tipo de Falla:",
    options=df_original['Tipo Falla'].unique(),
    default=df_original['Tipo Falla'].unique()
)

estados_seleccionados = st.sidebar.multiselect(
    "Selecciona Estado de la Orden:",
    options=df_original['Estado'].unique(),
    default=df_original['Estado'].unique()
)

df_filtrado = df_original[
    (df_original['Tipo Falla'].isin(tipos_seleccionados)) & 
    (df_original['Estado'].isin(estados_seleccionados))
]

# 4. INDICADORES CLAVE DE RENDIMIENTO (KPIs)
st.subheader("📊 Resumen Operativo")
kpi1, kpi2, kpi3 = st.columns(3)

with kpi1:
    st.metric(label="Total Incidentes", value=len(df_filtrado))
with kpi2:
    horas_totales = df_filtrado['Horas Detención'].sum().round(1)
    st.metric(label="Horas de Parada Totales", value=f"{horas_totales} hrs")
with kpi3:
    if not df_filtrado.empty:
        peor_equipo = df_filtrado.groupby('Equipo')['Horas Detención'].sum().idxmax()
        st.metric(label="Equipo más Crítico", value=peor_equipo)
    else:
        st.metric(label="Equipo más Crítico", value="N/A")

st.write("---")

# 5. DISTRIBUCIÓN VISUAL: TABLA Y GRÁFICOS (3 Columnas)
col_tabla, col_grafico1, col_grafico2 = st.columns([1.2, 0.9, 0.9])

with col_tabla:
    st.subheader("📋 Registro de Detenciones")
    st.dataframe(df_filtrado, use_container_width=True, height=380)

if not df_filtrado.empty:
    with col_grafico1:
        st.subheader("📈 Tiempos por Equipo")
        datos_barras = df_filtrado.groupby('Equipo', as_index=False)['Horas Detención'].sum()
        
        fig_barras = px.bar(
            datos_barras, 
            x='Horas Detención', 
            y='Equipo', 
            orientation='h',
            labels={'Horas Detención': 'Horas', 'Equipo': 'Equipo'},
            color='Horas Detención',
            color_continuous_scale=px.colors.sequential.Blugrn
        )
        fig_barras.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=350, coloraxis_showscale=False)
        st.plotly_chart(fig_barras, use_container_width=True)

    with col_grafico2:
        st.subheader("🍕 Distribución de Fallas")
        datos_circular = df_filtrado.groupby('Tipo Falla', as_index=False).size()
        
        fig_circular = px.pie(
            datos_circular, 
            values='size', 
            names='Tipo Falla',
            hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_circular.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=350)
        fig_circular.update_traces(textinfo='percent+label')
        fig_circular.update_layout(showlegend=False)
        
        st.plotly_chart(fig_circular, use_container_width=True)
else:
    with col_grafico1:
        st.warning("Sin datos para graficar.")
