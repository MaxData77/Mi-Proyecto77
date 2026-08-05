import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import openpyxl
import io
import unicodedata

# 1. Configuración de página en modo ancho (Wide)
st.set_page_config(
    page_title="Gestión y Control de Ordenes OT",
    layout="wide"
)

COLOR_VERDE = "#4C68A2"
COLOR_ROJO = "#FF0000"
ARCHIVOS_POR_PAGINA = 3

# Carácter "Hangul Filler" (U+3164) que Excel deja en checkboxes no marcados;
# hay que tratarlo como vacío igual que un espacio en blanco.
CARACTER_FANTASMA = '\u3164'

# Orden fijo de los 14 ítems que se muestran en el gráfico de barras
BAR_ITEMS_ORDEN = [
    'HORÓMETRO',
    'MOTIVO DETENCIÓN DEL EQUIPO',
    'CÓDIGO COMPONENTE SMCS',
    'CÓDIGO MODIFICADOR',
    'CÓDIGO TRABAJO',
    'DESCRIPCIÓN DEL SÍNTOMA',
    'CÓDIGO SÍNTOMA',
    'DESCRIPCIÓN DE LA CAUSA',
    'CÓDIGO CAUSA',
    'DESCRIPCIÓN DE ACTIVIDADES',
    'INFORME SIMS',
    'RESUMEN ANÁLISIS DE FALLA',
    'JEFE DE TURNO NOMBRE Y RUT',
    'TÉCNICO RESPONSABLE NOMBRE Y RUT',
]

# 2. Inyección de CSS para rediseñar la interfaz
st.markdown("""
    <style>
    .stApp {
        background-color: #C6DAF8 !important;
    }
    .metric-card {
        background-color: #4C68A2;
        color: white;
        padding: 6px 15px;
        border-radius: 10px;
        border: 1px solid #FFC000;
        text-align: center;
        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.3);
    }
    .metric-card h3 {
        color: #FFFFFF !important;
        font-size: 17px !important;
        font-weight: 600 !important;
        margin-bottom: 0px !important;
    }
    .metric-card h1 {
        color: #FFFFFF !important;
        font-size: 34px !important;
        font-weight: bold !important;
        margin: 0 !important;
        line-height: 1.2;
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
    /* Línea indicadora de la pestaña activa (Resumen por OT / Encargado de OT) en verde */
    div[data-baseweb="tab-highlight"] {
        background-color: #00CC00 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] p {
        color: #00CC00 !important;
        font-weight: 700 !important;
    }
    </style>
""", unsafe_allow_html=True)


def _limpiar(valor):
    """Convierte el valor de una celda a texto limpio: quita espacios y el
    carácter fantasma que Excel deja en checkboxes no marcados."""
    if valor is None:
        return ""
    return str(valor).replace(CARACTER_FANTASMA, "").strip()


def _quitar_tildes(texto):
    """Normaliza texto para comparar sin tildes (CAMBIÓ -> CAMBIO)."""
    nfkd = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd if not unicodedata.combining(c)])


def _esta_marcado(valor):
    """True si la celda de un checkbox contiene una 'X' (mayúscula o minúscula)."""
    return _limpiar(valor).upper() == "X"


# 3. Función de Procesamiento con Todas las Reglas de Negocio Oficiales
def procesar_archivo_ot(file_bytes):
    resultado = {
        "equipo": "OT sin equipo",
        "orden": "OT sin orden",
        "turno": "OT sin turno",
        "faltantes": 0,
        "detalle": "Ninguno",
        "estado": "Cumple",
        "categoria": "-",
        "seccion": "-",
        "jefe_turno_nombre": "Sin dato",
        "tecnico_nombre": "Sin dato",
        "sims_estado": "Sin datos",
        "campos_bar": {item: "Cumple" for item in BAR_ITEMS_ORDEN}
    }
    try:
        file_bytes.seek(0)
        file_stream = io.BytesIO(file_bytes.read())
        wb = openpyxl.load_workbook(file_stream, data_only=True)
        sheet = wb.active

        # --- EXTRACCIÓN DE DATOS DE CABECERA (para columnas Equipo/Orden/Turno) ---
        val_equipo = _limpiar(sheet['G7'].value)
        resultado["equipo"] = val_equipo if val_equipo else "OT sin equipo"

        val_turno = _limpiar(sheet['G19'].value)
        resultado["turno"] = val_turno if val_turno else "OT sin turno"

        # Orden: usa J42 si es válido, si no G25; si ambas están vacías o dicen "no" -> "OT sin orden"
        val_j42 = _limpiar(sheet['J42'].value)
        val_g25 = _limpiar(sheet['G25'].value)

        def _valor_valido_orden(v):
            return v != "" and v.upper() != "NO"

        if _valor_valido_orden(val_j42):
            resultado["orden"] = val_j42
        elif _valor_valido_orden(val_g25):
            resultado["orden"] = val_g25
        else:
            resultado["orden"] = "OT sin orden"

        # --- NOMBRES PARA "ENCARGADO DE OT" ---
        nombre_jefe = _limpiar(sheet['C238'].value)
        nombre_tecnico = _limpiar(sheet['BD239'].value)
        resultado["jefe_turno_nombre"] = nombre_jefe if nombre_jefe else "Sin dato"
        resultado["tecnico_nombre"] = nombre_tecnico if nombre_tecnico else "Sin dato"

        # --- CAMPOS MAESTROS: (etiqueta, celda, categoría, sección) ---
        CAMPOS_MAESTRO = [
            ("EQUIPO", 'G7', "No Crítico", "Antecedentes de la Detención"),
            ("HORÓMETRO", 'G13', "Crítico", "Antecedentes de la Detención"),
            ("TURNO", 'G19', "No Crítico", "Antecedentes de la Detención"),
            ("ORDEN DE PEDIDO / SALIDA DE BODEGA", 'G25', "No Crítico", "Antecedentes de la Detención"),
            ("FECHA INICIO DETENCIÓN", 'X9', "No Crítico", "Antecedentes de la Detención"),
            ("HORA INICIO DETENCIÓN", 'AB9', "No Crítico", "Antecedentes de la Detención"),
            ("FECHA FINAL DETENCIÓN", 'X11', "No Crítico", "Antecedentes de la Detención"),
            ("HORA FINAL DETENCIÓN", 'AB11', "No Crítico", "Antecedentes de la Detención"),
            ("MOTIVO DETENCIÓN DEL EQUIPO", 'AB25', "Crítico", "Antecedentes de la Detención"),

            ("HORA INICIO TRABAJO", 'B42', "No Crítico", "Información del Trabajo"),
            ("HORA TERMINO TRABAJO", 'F42', "No Crítico", "Información del Trabajo"),
            ("N° ORDEN SERVICIO", 'J42', "No Crítico", "Información del Trabajo"),
            ("CÓDIGO COMPONENTE SMCS", 'Q42', "Crítico", "Información del Trabajo"),
            ("CÓDIGO MODIFICADOR", 'T42', "Crítico", "Información del Trabajo"),
            ("CÓDIGO TRABAJO", 'W42', "Crítico", "Información del Trabajo"),
            ("DESCRIPCIÓN DEL SÍNTOMA", 'Z42', "Crítico", "Información del Trabajo"),
            ("CÓDIGO SÍNTOMA", 'AO42', "Crítico", "Información del Trabajo"),
            ("DESCRIPCIÓN DE LA CAUSA", 'AR42', "Crítico", "Información del Trabajo"),
            ("CÓDIGO CAUSA", 'BH42', "Crítico", "Información del Trabajo"),
            ("TIPO TAREA", 'BO42', "No Crítico", "Información del Trabajo"),
            ("TAREA PRINCIPAL", 'BV42', "No Crítico", "Información del Trabajo"),

            ("JEFE TURNO (NOMBRE)", 'C238', "Crítico", "Firma Responsables"),
            ("JEFE TURNO (RUT)", 'C244', "Crítico", "Firma Responsables"),
            ("TÉCNICO (NOMBRE)", 'BD239', "Crítico", "Firma Responsables"),
            ("TÉCNICO (RUT)", 'BD243', "Crítico", "Firma Responsables"),
        ]

        campos_con_hallazgo = []  # lista de (etiqueta, categoria, seccion)
        campos_estado = {}        # etiqueta -> "Cumple" / "No cumple"

        for label, celda, categoria, seccion in CAMPOS_MAESTRO:
            val_str = _limpiar(sheet[celda].value).upper()
            no_cumple = (val_str == "")

            # --- Reglas especiales adicionales ---
            if label == 'DESCRIPCIÓN DEL SÍNTOMA' and val_str == "SIN INFORMACION":
                no_cumple = True
            elif label == 'CÓDIGO SÍNTOMA' and val_str == "156":
                no_cumple = True
            elif label == 'DESCRIPCIÓN DE LA CAUSA' and val_str == "OTROS":
                no_cumple = True
            elif label == 'CÓDIGO CAUSA' and val_str in ["6.6", "6,6", "7.1", "7,1"]:
                no_cumple = True
            elif label == 'ORDEN DE PEDIDO / SALIDA DE BODEGA' and val_str == "NO":
                no_cumple = True

            campos_estado[label] = "No cumple" if no_cumple else "Cumple"
            if no_cumple:
                campos_con_hallazgo.append((label, categoria, seccion))

        # --- UBICACIÓN DEL EQUIPO: TALLER (R21) o TERRENO (Y21) ---
        if not (_esta_marcado(sheet['R21'].value) or _esta_marcado(sheet['Y21'].value)):
            campos_con_hallazgo.append(("UBICACIÓN DEL EQUIPO (Taller/Terreno)", "No Crítico", "Antecedentes de la Detención"))

        # --- TIPO Y RESPONSABILIDAD DE LA DETENCIÓN ---
        # Grilla 3x2: filas Planeado/Imprevisto/Accidente, columnas Dealer(AQ)/Customer(AW).
        # Se espera exactamente una celda marcada con "X".
        celdas_grilla = ['AQ13', 'AQ15', 'AQ17', 'AW13', 'AW15', 'AW17']
        if not any(_esta_marcado(sheet[c].value) for c in celdas_grilla):
            campos_con_hallazgo.append(("TIPO Y RESPONSABILIDAD DE LA DETENCIÓN", "No Crítico", "Antecedentes de la Detención"))

        # --- EQUIPO ENTREGADO: SI (BL10) o NO (BO10) ---
        if not (_esta_marcado(sheet['BL10'].value) or _esta_marcado(sheet['BO10'].value)):
            campos_con_hallazgo.append(("EQUIPO ENTREGADO (Sí/No)", "No Crítico", "Antecedentes de la Detención"))

        # --- DESCRIPCIÓN DE ACTIVIDADES (combina Tipo Tarea + Tarea Principal) ---
        bar_actividades_cumple = (
            campos_estado['TIPO TAREA'] == "Cumple" and campos_estado['TAREA PRINCIPAL'] == "Cumple"
        )

        # --- JEFE DE TURNO y TÉCNICO combinados (Nombre + RUT) ---
        bar_jefe_cumple = (
            campos_estado['JEFE TURNO (NOMBRE)'] == "Cumple" and campos_estado['JEFE TURNO (RUT)'] == "Cumple"
        )
        bar_tecnico_cumple = (
            campos_estado['TÉCNICO (NOMBRE)'] == "Cumple" and campos_estado['TÉCNICO (RUT)'] == "Cumple"
        )

        # --- RESUMEN ANÁLISIS DE FALLA (E205 + E211 + E216) ---
        analisis_texto = " ".join([
            _limpiar(sheet['E205'].value),
            _limpiar(sheet['E211'].value),
            _limpiar(sheet['E216'].value)
        ]).strip()
        resumen_vacio = (analisis_texto == "")
        if resumen_vacio:
            campos_con_hallazgo.append(("RESUMEN ANÁLISIS DE FALLA", "Crítico", "Resumen Análisis de Falla"))
        bar_resumen_cumple = not resumen_vacio

        # --- LÓGICA SIMS ---
        # AU189 ya NO se revisa como campo del bloque SIMS. En cambio, si en AU189
        # está escrito "N/A SIMS", la OT se considera directamente "No aplica SIMS".
        au189_normalizado = " ".join(_limpiar(sheet['AU189'].value).upper().split())
        no_aplica_sims = (au189_normalizado == "N/A SIMS")

        campos_sims = {
            'N° PIEZA QUE FALLÓ': sheet['B189'].value,
            'DESCRIPCIÓN DE LA PIEZA': sheet['E189'].value,
            'CANTIDAD': sheet['X189'].value,
            'CÓDIGO SERVICIO': sheet['AA189'].value,
            'N° GRUPO': sheet['AE189'].value,
            'DESCRIPCIÓN DEL GRUPO': sheet['AJ189'].value,
            '¿LLEGÓ AL FIN DE SU VIDA ÚTIL?': sheet['AR189'].value,
        }
        sims_faltantes = [k for k, v in campos_sims.items() if _limpiar(v) == ""]

        # --- Estado para la pestaña "Registro Sims" (siempre se calcula, sin importar el gatillo) ---
        if no_aplica_sims:
            resultado["sims_estado"] = "No aplica SIMS"
        elif len(sims_faltantes) == 0:
            resultado["sims_estado"] = "Sims completo"
        elif len(sims_faltantes) == len(campos_sims):
            resultado["sims_estado"] = "Falta SIMS"
        else:
            resultado["sims_estado"] = "Sims incompleto"

        bar_sims_cumple = True
        if not no_aplica_sims:
            # Si en E205, AB25 o B98 aparece "cambio"/"cambia"/"reemplaza", se exige el bloque SIMS.
            texto_disparador = _quitar_tildes(" ".join([
                _limpiar(sheet['E205'].value),
                _limpiar(sheet['AB25'].value),
                _limpiar(sheet['B98'].value)
            ]).upper())
            palabras_gatillo = ["CAMBIO", "CAMBIA", "REEMPLAZA"]
            requiere_sims = any(p in texto_disparador for p in palabras_gatillo)

            if requiere_sims:
                if len(sims_faltantes) == len(campos_sims):
                    campos_con_hallazgo.append(("Falta SIMS", "Crítico", "Registro Informe SIMS"))
                    bar_sims_cumple = False
                elif len(sims_faltantes) > 0:
                    campos_con_hallazgo.append(
                        (f"SIMS incompleto: falta {', '.join(sims_faltantes)}", "Crítico", "Registro Informe SIMS")
                    )
                    bar_sims_cumple = False
        # Si no_aplica_sims es True, bar_sims_cumple se mantiene en True y no se agrega ningún hallazgo.

        # --- Consolidación de los 14 ítems del gráfico de barras ---
        resultado["campos_bar"] = {
            'HORÓMETRO': campos_estado['HORÓMETRO'],
            'MOTIVO DETENCIÓN DEL EQUIPO': campos_estado['MOTIVO DETENCIÓN DEL EQUIPO'],
            'CÓDIGO COMPONENTE SMCS': campos_estado['CÓDIGO COMPONENTE SMCS'],
            'CÓDIGO MODIFICADOR': campos_estado['CÓDIGO MODIFICADOR'],
            'CÓDIGO TRABAJO': campos_estado['CÓDIGO TRABAJO'],
            'DESCRIPCIÓN DEL SÍNTOMA': campos_estado['DESCRIPCIÓN DEL SÍNTOMA'],
            'CÓDIGO SÍNTOMA': campos_estado['CÓDIGO SÍNTOMA'],
            'DESCRIPCIÓN DE LA CAUSA': campos_estado['DESCRIPCIÓN DE LA CAUSA'],
            'CÓDIGO CAUSA': campos_estado['CÓDIGO CAUSA'],
            'DESCRIPCIÓN DE ACTIVIDADES': "Cumple" if bar_actividades_cumple else "No cumple",
            'INFORME SIMS': "Cumple" if bar_sims_cumple else "No cumple",
            'RESUMEN ANÁLISIS DE FALLA': "Cumple" if bar_resumen_cumple else "No cumple",
            'JEFE DE TURNO NOMBRE Y RUT': "Cumple" if bar_jefe_cumple else "No cumple",
            'TÉCNICO RESPONSABLE NOMBRE Y RUT': "Cumple" if bar_tecnico_cumple else "No cumple",
        }

        # --- CONSOLIDACIÓN DEL ESTADO FINAL ---
        cant_faltantes = len(campos_con_hallazgo)
        resultado["faltantes"] = cant_faltantes
        if cant_faltantes > 0:
            # Ya NO se agrega la etiqueta (Crítico)/(No Crítico) al texto del detalle.
            resultado["detalle"] = ", ".join([label for label, _, _ in campos_con_hallazgo])
            resultado["estado"] = "No cumple"
            secciones_unicas = sorted(set(sec for _, _, sec in campos_con_hallazgo))
            resultado["seccion"] = ", ".join(secciones_unicas)
            resultado["categoria"] = "Crítico" if any(cat == "Crítico" for _, cat, _ in campos_con_hallazgo) else "No Crítico"
        else:
            resultado["detalle"] = "Completo"
            resultado["estado"] = "Cumple"
            resultado["seccion"] = "-"
            resultado["categoria"] = "-"

        return resultado
    except Exception as e:
        resultado["estado"] = "No cumple"
        resultado["detalle"] = f"Error: {str(e)}"
        return resultado


# --- ENCABEZADO DE INTERFAZ ---
st.markdown("<h2 style='text-align: center; color: black; font-weight: bold; font-size: 32px;'>GESTION Y CONTROL EN LOS PROCESOS OPERACIONALES</h2>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: black; font-size: 22px;'>Revisión de Ordenes de Trabajo OT</h3>", unsafe_allow_html=True)

# ================= BARRA LATERAL (Cargador) =================
with st.sidebar:
    st.markdown("""
        <div class="upload-container">
            <p style="color: #FFC000; font-weight: bold; font-size: 14px; margin-bottom: 8px;">
                Carga los archivos a revisar<br>(Excel .XLSX)
            </p>
        </div>
    """, unsafe_allow_html=True)

    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

    uploaded_files = st.file_uploader(
        "", accept_multiple_files=True, type=['xlsx'],
        key=f"uploader_{st.session_state.uploader_key}"
    )

    ejecutar = st.button("Ejecutar revisión", use_container_width=True)

    if st.button("🗑️ Borrar archivos", use_container_width=True):
        st.session_state.uploader_key += 1
        st.rerun()

# ================= CONTENIDO PRINCIPAL (Dashboard Analítico) =================
if not ejecutar or not uploaded_files:
    m1, m2, m3, m4 = st.columns(4)
    for m, txt in zip([m1, m2, m3, m4], ["OT Revisadas", "OT con observación", "Hallazgos detectados", "OT completa"]):
        with m: st.markdown(f'<div class="metric-card"><h3>{txt}</h3><h1>0</h1></div>', unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        st.write("**Cumplimiento de Campos por OT**")
        df_empty_bar = pd.DataFrame({'Campo': ['Esperando archivos...'], 'Porcentaje': [0]})
        fig_bar = px.bar(df_empty_bar, x='Porcentaje', y='Campo', orientation='h', color_discrete_sequence=['#CCCCCC'])
        fig_bar.update_layout(height=550, margin=dict(l=0, r=0, t=10, b=0))
        fig_bar.update_xaxes(title_text='Porcentaje de campos por OT')
        st.plotly_chart(fig_bar, use_container_width=True)
    with g2:
        st.write("**Cumplimiento general por OT**")
        df_empty_pie = pd.DataFrame({'Estado': ['Sin datos'], 'Cantidad': [1]})
        fig_pie = px.pie(df_empty_pie, values='Cantidad', names='Estado', hole=0.6, color_discrete_sequence=['#CCCCCC'])
        fig_pie.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig_pie, use_container_width=True)

    st.write("**Resumen por OT**")
    df_empty_table = pd.DataFrame(columns=['Documento OT', 'Equipo', 'Orden', 'Turno', 'Categoría', 'Sección', 'Cant. Faltantes', 'Detalle Campos Faltantes', 'Estado'])
    st.dataframe(df_empty_table, use_container_width=True)

else:
    # --- PROCESAMIENTO ACTIVO ---
    lista_resumen = []
    conteo_campos = {item: {'Cumple': 0, 'No cumple': 0} for item in BAR_ITEMS_ORDEN}

    for f in uploaded_files:
        datos_ot = procesar_archivo_ot(f)
        lista_resumen.append({
            'Documento OT': f.name,
            'Equipo': datos_ot['equipo'],
            'Orden': datos_ot['orden'],
            'Turno': datos_ot['turno'],
            'Categoría': datos_ot['categoria'],
            'Sección': datos_ot['seccion'],
            'Cant. Faltantes': datos_ot['faltantes'],
            'Detalle Campos Faltantes': datos_ot['detalle'],
            'Estado': datos_ot['estado'],
            'Jefe de Turno': datos_ot['jefe_turno_nombre'],
            'Técnico Responsable': datos_ot['tecnico_nombre'],
            'Estado SIMS': datos_ot['sims_estado'],
        })

        for campo, estado_campo in datos_ot['campos_bar'].items():
            conteo_campos[campo][estado_campo] += 1

    df_resumen = pd.DataFrame(lista_resumen)
    df_resumen.index = range(1, len(df_resumen) + 1)
    df_resumen.index.name = "Qty"

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
        st.write("**Cumplimiento de Campos por OT**")
        filas_barras = []
        for campo in BAR_ITEMS_ORDEN:
            v = conteo_campos[campo]
            total = v['Cumple'] + v['No cumple']
            pct_cumple = (v['Cumple'] / total * 100) if total > 0 else 100
            pct_no_cumple = (v['No cumple'] / total * 100) if total > 0 else 0
            filas_barras.append({'Campo': campo, 'Estado': 'Cumple', 'Porcentaje': pct_cumple})
            filas_barras.append({'Campo': campo, 'Estado': 'No cumple', 'Porcentaje': pct_no_cumple})

        df_barras = pd.DataFrame(filas_barras)

        fig_bar = px.bar(
            df_barras, x='Porcentaje', y='Campo', color='Estado', orientation='h',
            category_orders={'Campo': BAR_ITEMS_ORDEN[::-1]},
            color_discrete_map={'Cumple': COLOR_VERDE, 'No cumple': COLOR_ROJO},
            text=df_barras['Porcentaje'].apply(lambda p: f'{p:.1f}%' if p > 0 else '')
        )
        fig_bar.update_traces(textposition='inside', insidetextanchor='middle')
        # Texto blanco dentro de ambas barras (azul #4C68A2 y roja) para buen contraste
        for trace in fig_bar.data:
            if trace.name == 'Cumple':
                trace.textfont = dict(color='white', size=12)
            else:
                trace.textfont = dict(color='white', size=12)
        fig_bar.update_layout(
            barmode='stack', height=550, margin=dict(l=0, r=0, t=10, b=0), legend_title_text=''
        )
        fig_bar.update_xaxes(title_text='Porcentaje de campos por OT')
        st.plotly_chart(fig_bar, use_container_width=True)

    with g2:
        st.write("**Cumplimiento general por OT**")
        etiqueta_cumple = f'Cumple {ot_completas} OT'
        etiqueta_no_cumple = f'No cumple {ot_con_observacion} OT'
        df_pie = pd.DataFrame({
            'Estado': [etiqueta_cumple, etiqueta_no_cumple],
            'Cantidad': [ot_completas, ot_con_observacion]
        })
        fig_pie = px.pie(
            df_pie, values='Cantidad', names='Estado', hole=0.6,
            color='Estado',
            color_discrete_map={etiqueta_cumple: COLOR_VERDE, etiqueta_no_cumple: COLOR_ROJO}
        )
        # Solo el porcentaje dentro de la dona: texto blanco en ambas franjas (azul y roja).
        # El total de OT (cumple/no cumple) queda en la leyenda, dentro del recuadro del gráfico.
        fig_pie.update_traces(
            texttemplate='%{percent}',
            textposition='inside',
            textfont=dict(color=['white', 'white'], size=13)
        )
        fig_pie.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0), legend_title_text='')
        st.plotly_chart(fig_pie, use_container_width=True)

    # --- TABLA RESUMEN, ENCARGADO DE OT Y REGISTRO SIMS (en pestañas) ---
    columnas_resumen = ['Documento OT', 'Equipo', 'Orden', 'Turno', 'Categoría', 'Sección',
                         'Cant. Faltantes', 'Detalle Campos Faltantes', 'Estado']

    tab_resumen, tab_encargado, tab_sims = st.tabs(["Resumen por OT", "Encargado de OT", "Registro Sims"])

    with tab_resumen:
        st.dataframe(df_resumen[columnas_resumen], use_container_width=True)

    with tab_encargado:
        df_no_cumple = df_resumen[df_resumen['Estado'] == 'No cumple'].copy()
        if df_no_cumple.empty:
            st.info("No hay OT con observaciones: todas están completas.")
        else:
            df_encargado = df_no_cumple.rename(columns={
                'Detalle Campos Faltantes': 'Detalle Campo Faltante',
            })[['Documento OT', 'Jefe de Turno', 'Técnico Responsable', 'Sección', 'Detalle Campo Faltante']]
            st.dataframe(df_encargado, use_container_width=True, hide_index=True)

    with tab_sims:
        df_sims = df_resumen[['Documento OT', 'Jefe de Turno', 'Técnico Responsable', 'Sección', 'Estado SIMS']].rename(
            columns={'Estado SIMS': 'Estado'}
        )
        st.dataframe(df_sims, use_container_width=True, hide_index=True)
