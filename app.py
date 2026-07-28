import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import openpyxl
import io

# 1. Configuración de página en modo ancho (Wide)
st.set_page_config(
    page_title="Gestión y Control de Ordenes OT",
    layout="wide"
)

COLOR_VERDE = "#00FF00"
COLOR_ROJO = "#FF0000"

# 2. Inyección de CSS para rediseñar la interfaz
st.markdown("""
    <style>
    .stApp {
        background-color: #FFC000 !important;
    }
    .metric-card {
        background-color: #000000;
        color: white;
        padding: 10px 15px;
        border-radius: 12px;
        text-align: center;
        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.3);
    }
    .metric-card h3 {
        color: #FFFFFF !important;
        font-size: 19px !important;
        font-weight: 600 !important;
        margin-bottom: 0px !important;
    }
    .metric-card h1 {
        color: #FFFFFF !important;
        font-size: 40px !important;
        font-weight: bold !important;
        margin: 0 !important;
        line-height: 1.1;
    }
    .upload-container {
        background-color: #000000;
        padding: 15px;
        border-radius: 15px;
        text-align: center;
        color: white;
        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.3);
    }
    div[data-testid="stFileUploader"] {
        background-color: #1A1A1A !important;
        border-radius: 10px;
        padding: 10px;
        border: 1px dashed #FFC000 !important;
    }
    div[data-testid="stFileUploader"] label {
        color: #FFC000 !important;
    }
    </style>
""", unsafe_allow_html=True)

# 3. Función de Procesamiento con Todas las Reglas de Negocio Oficiales
def procesar_archivo_ot(file_bytes):
    resultado = {
        "equipo": "Sin equipo",
        "orden": "OT Sin Orden",
        "turno": "N/A",
        "faltantes": 0,
        "detalle": "Ninguno",
        "estado": "Cumple",
        "campos_validados": {}
    }
    try:
        file_bytes.seek(0)
        file_stream = io.BytesIO(file_bytes.read())
        wb = openpyxl.load_workbook(file_stream, data_only=True)
        sheet = wb.active
        
        # --- EXTRACCIÓN DE DATOS DE CABECERA ---
        val_equipo = sheet['G7'].value
        val_orden = sheet['J42'].value or sheet['G25'].value
        val_turno = sheet['G19'].value
        
        if val_equipo and str(val_equipo).strip(): resultado["equipo"] = str(val_equipo).strip()
        if val_orden and str(val_orden).strip(): resultado["orden"] = str(val_orden).strip()
        if val_turno and str(val_turno).strip(): resultado["turno"] = str(val_turno).strip()

        # --- DICCIONARIO GENERAL DE CELDAS A EVALUAR ---
        campos_criticos = {
            # PÁGINA 1: Antecedentes y Trabajo
            'HORÓMETRO': sheet['G13'].value,
            'ORDEN DE PEDIDO': sheet['G25'].value,
            'MOTIVO DETENCIÓN': sheet['AB25'].value,
            'FECHA INICIO': sheet['X9'].value,
            'HORA INICIO': sheet['AB9'].value,
            'FECHA FINAL': sheet['X11'].value,
            'HORA FINAL': sheet['AB11'].value,
            'HORA TRABAJO INICIO': sheet['B42'].value,
            'HORA TRABAJO TERMINO': sheet['F42'].value,
            'CÓDIGO SMCS': sheet['Q42'].value,
            'CÓDIGO MODIFICADOR': sheet['T42'].value,
            'CÓDIGO TRABAJO': sheet['W42'].value,
            'DESCRIPCIÓN SÍNTOMA': sheet['Z42'].value,
            'CÓDIGO SÍNTOMA': sheet['AO42'].value,
            'DESCRIPCION CAUSA': sheet['AR42'].value,
            'CÓDIGO CAUSA': sheet['BH42'].value,
            'TIPO TAREA': sheet['BO42'].value,
            'TAREA PRINCIPAL': sheet['BV42'].value,
            
            # PÁGINA 2: Informe SIMS
            'N° PIEZA QUE FALLÓ': sheet['B189'].value,
            'DESCRIPCIÓN DE LA PIEZA': sheet['E189'].value,
            'CANTIDAD': sheet['X189'].value,
            'CÓDIGO SERVICIO': sheet['AA189'].value,
            'N° GRUPO': sheet['AE189'].value,
            'DESCRIPCIÓN DEL GRUPO': sheet['AJ186'].value,
            'FIN VIDA ÚTIL?': sheet['AR189'].value,
            'COMENTARIOS SIMS': sheet['AU189'].value,
            
            # PÁGINA 2: Firmas Responsables
            'JEFE TURNO (NOMBRE)': sheet['C238'].value,
            'JEFE TURNO (RUT)': sheet['C244'].value,
            'TECNICO (NOMBRE)': sheet['BD239'].value,
            'TECNICO (RUT)': sheet['BD243'].value,
        }
        
        campos_con_hallazgo = []

        # --- VALIDACIÓN DINÁMICA DE REGLAS DE NEGOCIO ---
        for campo, valor in campos_criticos.items():
            val_str = str(valor).strip().upper() if valor is not None else ""
            hubo_alerta = False
            
            if val_str == "":
                hubo_alerta = True
            elif campo == 'DESCRIPCIÓN SÍNTOMA' and val_str == "SIN INFORMACION":
                hubo_alerta = True
            elif campo == 'CÓDIGO SÍNTOMA' and val_str == "156":
                hubo_alerta = True
            elif campo == 'DESCRIPCION CAUSA' and val_str == "OTROS":
                hubo_alerta = True
            elif campo == 'CÓDIGO CAUSA' and val_str in ["6.6", "6,6", "7.1", "7,1"]:
                hubo_alerta = True

            if hubo_alerta:
                campos_con_hallazgo.append(campo)
                resultado["campos_validados"][campo] = "No cumple"
            else:
                resultado["campos_validados"][campo] = "Cumple"

        # --- REGLA AVANZADA: RESUMEN ANÁLISIS DE FALLA ---
        analisis_texto = " ".join([
            str(sheet['E205'].value or ""), 
            str(sheet['E211'].value or ""), 
            str(sheet['E216'].value or "")
        ]).strip().upper()
        
        resultado["campos_validados"]['RESUMEN ANÁLISIS DE FALLA'] = "Cumple"
        if analisis_texto == "":
            campos_con_hallazgo.append('RESUMEN ANÁLISIS DE FALLA')
            resultado["campos_validados"]['RESUMEN ANÁLISIS DE FALLA'] = "No cumple"
        elif "CAMBIO" in analisis_texto:
            val_b189 = sheet['B189'].value
            if val_b189 is None or str(val_b189).strip() == "":
                if 'N° PIEZA QUE FALLÓ' not in campos_con_hallazgo:
                    campos_con_hallazgo.append('N° PIEZA QUE FALLÓ')
                resultado["campos_validados"]['N° PIEZA QUE FALLÓ'] = "No cumple"

        # --- CONSOLIDACIÓN DEL ESTADO FINAL ---
        cant_faltantes = len(campos_con_hallazgo)
        if cant_faltantes > 0:
            resultado["faltantes"] = cant_faltantes
            resultado["detalle"] = ", ".join(campos_con_hallazgo)
            resultado["estado"] = "No cumple"
        else:
            resultado["detalle"] = "Completo"
            resultado["estado"] = "Cumple"
            
        return resultado
    except Exception as e:
        resultado["estado"] = "No cumple"
        resultado["detalle"] = f"Error: {str(e)}"
        return resultado

# --- ENCABEZADO DE INTERFAZ ---
st.markdown("<h2 style='text-align: center; color: black; font-weight: bold; font-size: 32px;'>GESTION Y CONTROL EN LOS PROCESOS OPERACIONALES</h2>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: black; font-size: 22px;'>Revisión de Ordenes de Trabajo OT</h3>", unsafe_allow_html=True)

col_left, col_right = st.columns()

# ================= COLUMNA IZQUIERDA (Cargador) =================
with col_left:
    st.markdown("""
        <div class="upload-container">
            <p style="color: #FFC000; font-weight: bold; font-size: 14px; margin-bottom: 8px;">
                Carga los archivos a revisar<br>(Excel .XLSX)
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader("", accept_multiple_files=True, type=['xlsx'])
    ejecutar = st.button("Ejecutar revisión", use_container_width=True)

# ================= COLUMNA DERECHA (Dashboard Analítico) =================
with col_right:
    if not ejecutar or not uploaded_files:
        # Estado Inicial vacío (Dashboard en 0)
        m1, m2, m3, m4 = st.columns(4)
        for m, txt in zip([m1, m2, m3, m4], ["OT Revisadas", "OT con observación", "Hallazgos detectados", "OT completa"]):
            with m: st.markdown(f'<div class="metric-card"><h3>{txt}</h3><h1>0</h1></div>', unsafe_allow_html=True)

        g1, g2 = st.columns(2)
        with g1:
            st.write("**Campos Revisados**")
            # CORREGIDO DE RAÍZ: Estructura de diccionario válida y completa con datos iniciales reales
            df_empty_bar = pd.DataFrame({'Campo': ['Esperando archivos...'], 'Porcentaje': [0]})
            fig_bar = px.bar(df_empty_bar, x='Porcentaje', y='Campo', orientation='h', color_discrete_sequence=['#CCCCCC'])
            fig_bar.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0))
            st.plotly_chart(fig_bar, use_container_width=True)
        with g2:
            st.write("**Total OT Revisadas**")
            # CORREGIDO DE RAÍZ: Estructura de diccionario válida y completa con datos iniciales reales
            df_empty_pie = pd.DataFrame({'Estado': ['Sin datos'], 'Cantidad': [1]})
            fig_pie = px.pie(df_empty_pie, values='Cantidad', names='Estado', hole=0.6, color_discrete_sequence=['#CCCCCC'])
            fig_pie.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0))
            st.plotly_chart(fig_pie, use_container_width=True)

        st.write("**Resumen por OT**")
        df_empty_table = pd.DataFrame(columns=['Archivo', 'Equipo', 'Orden', 'Turno', 'Cant. Faltantes', 'Detalle Campos Faltantes', 'Estado'])
        st.dataframe(df_empty_table, use_container_width=True)

    else:
        # --- PROCESAMIENTO ACTIVO ---
        lista_resumen = []
        conteo_campos = {}

        for f in uploaded_files:
            datos_ot = procesar_archivo_ot(f)
            lista_resumen.append({
                'Archivo': f.name,
                'Equipo': datos_ot['equipo'],
                'Orden': datos_ot['orden'],
                'Turno': datos_ot['turno'],
