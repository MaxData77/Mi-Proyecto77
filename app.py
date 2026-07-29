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
ARCHIVOS_POR_PAGINA = 3

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
    /* Oculta la lista nativa de archivos cargados; usamos nuestra propia lista paginada */
    div[data-testid="stFileUploaderFileList"] {
        display: none !important;
    }
    .file-pill {
        background-color: #262626;
        color: #FFC000;
        border-radius: 8px;
        padding: 6px 10px;
        margin-bottom: 6px;
        font-size: 13px;
        text-align: left;
    }
    .file-page-info {
        color: #FFC000;
        font-size: 12px;
        text-align: center;
        margin-top: 6px;
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

# ================= BARRA LATERAL (Cargador) =================
# Al vivir en st.sidebar, Streamlit agrega automáticamente la flecha
# de colapsar/expandir en la esquina superior izquierda.
with st.sidebar:
    st.markdown("""
        <div class="upload-container">
            <p style="color: #FFC000; font-weight: bold; font-size: 14px; margin-bottom: 8px;">
                Carga los archivos a revisar<br>(Excel .XLSX)
            </p>
        </div>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader("", accept_multiple_files=True, type=['xlsx'])

    # --- LISTA PROPIA PAGINADA (3 archivos por página) ---
    if uploaded_files:
        total_archivos = len(uploaded_files)
        total_paginas = max(1, (total_archivos - 1) // ARCHIVOS_POR_PAGINA + 1)

        if "pagina_archivos" not in st.session_state:
            st.session_state.pagina_archivos = 1
        # Si se suben/quitan archivos y la página queda fuera de rango, la ajustamos
        if st.session_state.pagina_archivos > total_paginas:
            st.session_state.pagina_archivos = total_paginas

        pagina_actual = st.session_state.pagina_archivos
        inicio = (pagina_actual - 1) * ARCHIVOS_POR_PAGINA
        fin = inicio + ARCHIVOS_POR_PAGINA

        for f in uploaded_files[inicio:fin]:
            tamano_kb = f.size / 1024
            st.markdown(
                f'<div class="file-pill">📄 {f.name}<br><small>{tamano_kb:.1f} KB</small></div>',
                unsafe_allow_html=True
            )

        col_prev, col_info, col_next = st.columns([1, 2, 1])
        with col_prev:
            if st.button("◀", disabled=(pagina_actual <= 1), use_container_width=True):
                st.session_state.pagina_archivos -= 1
                st.rerun()
        with col_info:
            st.markdown(
                f'<div class="file-page-info">Página {pagina_actual} de {total_paginas}</div>',
                unsafe_allow_html=True
            )
        with col_next:
            if st.button("▶", disabled=(pagina_actual >= total_paginas), use_container_width=True):
                st.session_state.pagina_archivos += 1
                st.rerun()

    ejecutar = st.button("Ejecutar revisión", use_container_width=True)

# ================= CONTENIDO PRINCIPAL (Dashboard Analítico) =================
if not ejecutar or not uploaded_files:
    # Estado Inicial vacío (Dashboard en 0)
    m1, m2, m3, m4 = st.columns(4)
    for m, txt in zip([m1, m2, m3, m4], ["OT Revisadas", "OT con observación", "Hallazgos detectados", "OT completa"]):
        with m: st.markdown(f'<div class="metric-card"><h3>{txt}</h3><h1>0</h1></div>', unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        st.write("**Campos Revisados**")
        df_empty_bar = pd.DataFrame({'Campo': ['Esperando archivos...'], 'Porcentaje': [0]})
        fig_bar = px.bar(df_empty_bar, x='Porcentaje', y='Campo', orientation='h', color_discrete_sequence=['#CCCCCC'])
        fig_bar.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig_bar, use_container_width=True)
    with g2:
        st.write("**Total OT Revisadas**")
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
            'Cant. Faltantes': datos_ot['faltantes'],
            'Detalle Campos Faltantes': datos_ot['detalle'],
            'Estado': datos_ot['estado']
        })

        for campo, estado_campo in datos_ot['campos_validados'].items():
            if campo not in conteo_campos:
                conteo_campos[campo] = {'Cumple': 0, 'No cumple': 0}
            conteo_campos[campo][estado_campo] += 1

    df_resumen = pd.DataFrame(lista_resumen)

    total_ot = len(df_resumen)
    ot_con_observacion = int((df_resumen['Estado'] == 'No cumple').sum())
    hallazgos_totales = int(df_resumen['Cant. Faltantes'].sum())
    ot_completas = int((df_resumen['Estado'] == 'Cumple').sum())

    # --- MÉTRICAS SUPERIORES ---
    m1, m2, m3, m4 = st.columns(4)
    valores = [total_ot, ot_con_observacion, hallazgos_totales, ot_completas]
    titulos = ["OT Revisadas", "OT con observación", "Hallazgos detectados", "OT completa"]
    for m, txt, val in zip([m1, m2, m3, m4], titulos, valores):
        with m:
            st.markdown(f'<div class="metric-card"><h3>{txt}</h3><h1>{val}</h1></div>', unsafe_allow_html=True)

    # --- GRÁFICAS ---
    g1, g2 = st.columns(2)
    with g1:
        st.write("**Campos Revisados**")
        if len(conteo_campos) > 0:
            df_conteo = pd.DataFrame([
                {
                    'Campo': campo,
                    'Porcentaje': (v['No cumple'] / total_ot) * 100 if total_ot > 0 else 0
                }
                for campo, v in conteo_campos.items()
            ]).sort_values('Porcentaje', ascending=True)

            fig_bar = px.bar(
                df_conteo, x='Porcentaje', y='Campo', orientation='h',
                color_discrete_sequence=[COLOR_ROJO]
            )
            fig_bar.update_layout(height=550, margin=dict(l=0, r=0, t=10, b=0))
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("No hay datos de campos para mostrar.")

    with g2:
        st.write("**Total OT Revisadas**")
        df_pie = pd.DataFrame({
            'Estado': ['Cumple', 'No cumple'],
            'Cantidad': [ot_completas, ot_con_observacion]
        })
        fig_pie = px.pie(
            df_pie, values='Cantidad', names='Estado', hole=0.6,
            color='Estado',
            color_discrete_map={'Cumple': COLOR_VERDE, 'No cumple': COLOR_ROJO}
        )
        fig_pie.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig_pie, use_container_width=True)

    # --- TABLA RESUMEN ---
    st.write("**Resumen por OT**")
    st.dataframe(df_resumen, use_container_width=True)
