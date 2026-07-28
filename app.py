import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px

# Configuración de página ancha
st.set_page_config(page_title="Auditor Masivo de OTs", layout="wide")

st.title("📋 Auditor Masivo de Órdenes de Trabajo")
st.markdown("Carga múltiples archivos Excel en lote. Clasificación automática por nivel de criticidad de celdas vacías.")
st.write("---")

def auditar_archivos_masivos(lista_archivos):
    reporte_errores = []
    total_archivos = len(lista_archivos)
    archivos_con_errores = 0
    
    for archivo in lista_archivos:
        try:
            wb = openpyxl.load_workbook(archivo, data_only=True)
            hojas_disponibles = wb.sheetnames
            
            if "OT FORMATO IMPRIMIR" in hojas_disponibles:
                hoja1 = wb["OT FORMATO IMPRIMIR"]
            else:
                hoja1 = wb[hojas_disponibles[0]]
                
            if len(hojas_disponibles) > 1:
                hoja2 = wb[hojas_disponibles[1]]
            else:
                hoja2 = hoja1
            
            nombre_archivo = archivo.name
            errores_en_este_archivo = 0

            # -------------------------------------------------------------------------
            # CONFIGURACIÓN DE CAMPOS Y SU NIVEL DE CRITICIDAD
            # Estructura: (Página, Nombre, Hoja, [Celdas], Criticidad)
            # -------------------------------------------------------------------------
            definicion_campos = [
                # --- PÁGINA 1 ---
                ("Página 1", "EQUIPO", hoja1, ["G7"], "🚨 CRÍTICO"),
                ("Página 1", "HOROMETRO", hoja1, ["G13"], "🚨 CRÍTICO"),
                ("Página 1", "TURNO", hoja1, ["G19"], "🚨 CRÍTICO"),
                ("Página 1", "ORDEN DE PEDIDO / SALIDA DE BODEGA", hoja1, ["G25"], "⚠️ NO CRÍTICO"),
                ("Página 1", "INICIO (Fecha y Hora)", hoja1, ["X9", "AB9"], "🚨 CRÍTICO"),
                ("Página 1", "FINAL (Fecha y Hora)", hoja1, ["X11", "AB11"], "🚨 CRÍTICO"),
                ("Página 1", "UBICACIÓN (Taller o Terreno)", hoja1, ["R21", "Y21"], "🚨 CRÍTICO"),
                ("Página 1", "MOTIVO DETENCIÓN DEL EQUIPO", hoja1, ["AB25"], "🚨 CRÍTICO"),
                ("Página 1", "TIPO DETENCIÓN: PLANEADO", hoja1, ["AQ13"], "⚠️ NO CRÍTICO"),
                ("Página 1", "TIPO DETENCIÓN: IMPREVISTO", hoja1, ["AQ15"], "⚠️ NO CRÍTICO"),
                ("Página 1", "TIPO DETENCIÓN: ACCIDENTE", hoja1, ["AQ17"], "⚠️ NO CRÍTICO"),
                ("Página 1", "RESPONSABILIDAD: DEALER (FINNING)", hoja1, ["AQ13", "AQ15", "AQ17"], "🚨 CRÍTICO"),
                ("Página 1", "RESPONSABILIDAD: CUSTOMER (CLIENTE)", hoja1, ["AW13", "AW15", "AW17"], "🚨 CRÍTICO"),
                ("Página 1", "EQUIPO ENTREGADO (SI o NO)", hoja1, ["BL10", "BO10"], "🚨 CRÍTICO"),
                
                # --- PÁGINA 2 ---
                ("Página 2", "N° PIEZA QUE FALLÓ", hoja2, ["B189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "DESCRIPCIÓN DE LA PIEZA", hoja2, ["E189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "CANTIDAD", hoja2, ["X189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "CÓDIGO SERVICIO", hoja2, ["AA189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "N° GRUPO", hoja2, ["AE189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "DESCRIPCIÓN DEL GRUPO", hoja2, ["AJ186", "AJ189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "¿Llegó al fin de su vida útil?", hoja2, ["AR189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "COMENTARIOS", hoja2, ["AU189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "RESUMEN ANÁLISIS DE FALLA", hoja2, ["E205", "E211", "E216"], "🚨 CRÍTICO"),
                ("Página 2", "VALIDACIÓN DE OT POR JEFE TURNO", hoja2, ["C238", "C244"], "🚨 CRÍTICO"),
                ("Página 2", "TECNICO RESPONSABLE", hoja2, ["BD239", "BD243"], "🚨 CRÍTICO"),
            ]

            for num_pagina, nombre, hoja_obj, lista_celdas, criticidad in definicion_campos:
                campo_completado = False
                
                for celda_id in lista_celdas:
                    try:
                        valor = hoja_obj[celda_id].value
                        texto_limpio = str(valor).strip().lower() if valor is not None else ""
                        if valor is not None and texto_limpio not in ["", "no", "none"]:
                            campo_completado = True
                            break
                    except:
                        pass
                
                if not campo_completado:
                    errores_en_este_archivo += 1
                    reporte_errores.append({
                        "Archivo Excel": nombre_archivo,
                        "Página": num_pagina,
                        "Campo Incompleto": nombre,
                        "Celdas Mapeadas": ", ".join(lista_celdas),
                        "Criticidad": criticidad
                    })
            
            if errores_en_este_archivo > 0:
                archivos_con_errores += 1

        except Exception as e:
            reporte_errores.append({
                "Archivo Excel": archivo.name,
                "Página": "Error Técnico",
                "Campo Incompleto": "Error general de formato",
                "Celdas Mapeadas": "N/A",
                "Criticidad": "🚨 CRÍTICO"
            })
            archivos_con_errores += 1

    return pd.DataFrame(reporte_errores), total_archivos, archivos_con_errores

# Cargador masivo de archivos
archivos_cargados = st.file_uploader(
    "📂 Sube una o varias Órdenes de Trabajo al mismo tiempo (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if archivos_cargados:
    with st.spinner("Auditando criticidad en lote..."):
        df_errores, cant_total, cant_con_errores = auditar_archivos_masivos(archivos_cargados)
        cant_perfectos = cant_total - cant_con_errores
        
        st.subheader("📊 Resumen de la Auditoría Masiva")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Archivos Subidos", cant_total)
        m2.metric("Archivos 100% Completos ✅", cant_perfectos)
        m3.metric("Archivos con Errores ❌", cant_con_errores)
        
        st.write("---")
        
        if df_errores.empty:
            st.success("🎉 ¡Excelente! No se encontraron celdas vacías en ninguna jerarquía de criticidad.")
            fig = px.pie(values=[1], names=["100% Correctos"], color_discrete_sequence=["#2ecc71"], hole=0.4)
            st.plotly_chart(fig)
        else:
            col_izq, col_der = st.columns([1.3, 0.7])
            
            with col_izq:
                st.subheader("⚠️ Registro de Omisiones por Gravedad")
                
                # FILTRO DIRECTO EN INTERFAZ PARA TU CONTROL
                selector_crit = st.radio("Filtrar por nivel de riesgo:", ["Mostrar Todos los Errores", "Solo Errores 🚨 CRÍTICO", "Solo Advertencias ⚠️ NO CRÍTICO"], horizontal=True)
                
                if selector_crit == "Solo Errores 🚨 CRÍTICO":
                    df_mostrar = df_errores[df_errores['Criticidad'] == "🚨 CRÍTICO"]
                elif selector_crit == "Solo Advertencias ⚠️ NO CRÍTICO":
                    df_mostrar = df_errores[df_errores['Criticidad'] == "⚠️ NO CRÍTICO"]
                else:
                    df_mostrar = df_errores
                
                st.dataframe(df_mostrar[["Archivo Excel", "Página", "Campo Incompleto", "Celdas Mapeadas", "Criticidad"]], use_container_width=True, height=450)
            
            with col_der:
                st.subheader("📈 Volumen de Errores por Criticidad")
                # Gráfico circular dinámico basado únicamente en el volumen de celdas vacías detectadas
                conteo_crit = df_errores['Criticidad'].value_counts().reset_index()
                conteo_crit.columns = ['Nivel', 'Cantidad']
                
                fig = px.pie(
                    conteo_crit, 
                    values="Cantidad", 
                    names="Nivel", 
                    hole=0.4,
                    color="Nivel",
                    color_discrete_map={"🚨 CRÍTICO": "#e74c3c", "⚠️ NO CRÍTICO": "#f1c40f"}
                )
                fig.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=350)
                st.plotly_chart(fig, use_container_width=True)
