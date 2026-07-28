import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
import io

# 1. Configuración de página ancha
st.set_page_config(page_title="Auditor Masivo de OTs", layout="wide", page_icon="📋")

st.title("📋 Auditor Masivo de Órdenes de Trabajo")
st.markdown("Carga múltiples archivos Excel en lote. Clasificación automática con motor de alertas y reglas de negocio avanzadas.")
st.write("---")

# 2. Función de auditoría técnica
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
            # 1. LISTA DE CAMPOS OBLIGATORIOS BASE
            # -------------------------------------------------------------------------
            definicion_campos = [
                # --- ÍTEMS PRIORITARIOS PÁGINA 1 ---
                ("Página 1", "HORÓMETRO", hoja1, ["G13"], "🚨 CRÍTICO"),
                ("Página 1", "MOTIVO DETENCIÓN DEL EQUIPO", hoja1, ["AB25"], "🚨 CRÍTICO"),
                ("Página 1", "CÓDIGO COMPONENTE SMCS", hoja1, ["Q42"], "🚨 CRÍTICO"),
                ("Página 1", "CÓDIGO MODIFICADOR", hoja1, ["T42"], "🚨 CRÍTICO"),
                ("Página 1", "CÓDIGO TRABAJO", hoja1, ["W42"], "🚨 CRÍTICO"),
                ("Página 1", "TIPO TAREA", hoja1, ["BO42"], "🚨 CRÍTICO"),
                ("Página 1", "TAREA PRINCIPAL", hoja1, ["BV42"], "🚨 CRÍTICO"),
                ("Página 1", "DESCRIPCIÓN DE ACTIVIDADES", hoja1, ["Z42"], "🚨 CRÍTICO"),
                
                # --- RESPONSABILIDAD MATRIZ UNIFICADA ---
                # Revisa las 6 celdas juntas (Dealer: AQ13, AQ15, AQ17 | Customer: AW13, AW15, AW17)
                ("Página 1", "TIPO DETENCIÓN / RESPONSABILIDAD", hoja1, ["AQ13", "AQ15", "AQ17", "AW13", "AW15", "AW17"], "🚨 CRÍTICO"),

                # --- FIRMAS PÁGINA 2 ---
                ("Página 2", "FIRMA JEFE TURNO (NOMBRE + RUT)", hoja2, ["C238", "C244"], "🚨 CRÍTICO"),
                ("Página 2", "FIRMA TÉCNICO RESPONSABLE (NOMBRE + RUT)", hoja2, ["BD239", "BD243"], "🚨 CRÍTICO"),

                # --- OTROS CAMPOS DE APOYO Y SECUNDARIOS ---
                ("Página 1", "EQUIPO", hoja1, ["G7"], "🚨 CRÍTICO"),
                ("Página 1", "ORDEN DE PEDIDO / SALIDA DE BODEGA", hoja1, ["G25"], "⚠️ NO CRÍTICO"),
                ("Página 1", "INICIO (Fecha y Hora)", hoja1, ["X9", "AB9"], "🚨 CRÍTICO"),
                ("Página 1", "FINAL (Fecha y Hora)", hoja1, ["X11", "AB11"], "🚨 CRÍTICO"),
                ("Página 1", "UBICACIÓN (Taller o Terreno)", hoja1, ["R21", "Y21"], "🚨 CRÍTICO"),
                ("Página 1", "EQUIPO ENTREGADO (SI o NO)", hoja1, ["BL10", "BO10"], "🚨 CRÍTICO"),
                ("Página 1", "HORA INICIO ACTIVIDAD", hoja1, ["B42"], "🚨 CRÍTICO"),
                ("Página 1", "HORA TERMINO ACTIVIDAD", hoja1, ["F42"], "🚨 CRÍTICO"),
                ("Página 1", "Nº ORDEN SERVICIO", hoja1, ["J42"], "🚨 CRÍTICO"),

                # --- PÁGINA 2 CAMPOS OPCIONALES ---
                ("Página 2", "DESCRIPCIÓN DE LA PIEZA", hoja2, ["E189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "CANTIDAD", hoja2, ["X189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "CÓDIGO SERVICIO", hoja2, ["AA189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "N° GRUPO", hoja2, ["AE189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "DESCRIPCIÓN DEL GRUPO", hoja2, ["AJ186", "AJ189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "¿Llegó al fin de su vida útil?", hoja2, ["AR189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "COMENTARIOS", hoja2, ["AU189"], "⚠️ NO CRÍTICO")
            ]

            # Evaluar celdas vacías generales
            for num_pagina, nombre, hoja_obj, lista_celdas, criticidad in definicion_campos:
                campo_completado = False
                for celda_id in lista_celdas:
                    valor = hoja_obj[celda_id].value
                    texto_limpio = str(valor).strip().lower() if valor is not None else ""
                    if valor is not None and texto_limpio not in ["", "no", "none", "ㅤ"]:
                        campo_completado = True
                        break
                
                if not campo_completado:
                    errores_en_este_archivo += 1
                    reporte_errores.append({
                        "Archivo Excel": nombre_archivo,
                        "Página": num_pagina,
                        "Campo / Alerta": nombre,
                        "Celdas Mapeadas": ", ".join(lista_celdas),
                        "Detalle del Error": "Celda vacía u omitida",
                        "Criticidad": criticidad
                    })

            # -------------------------------------------------------------------------
            # 2. MOTOR DE ALERTAS CRÍTICAS (SÍNTOMAS Y CAUSAS)
            # -------------------------------------------------------------------------
            val_z42 = hoja1["Z42"].value
            val_ao42 = hoja1["AO42"].value
            val_ar42 = hoja1["AR42"].value
            val_bh42 = hoja1["BH42"].value

            alertas_especificas = [
                ("Página 1", "DESCRIPCIÓN DE SÍNTOMA", "Z42", val_z42, "SIN INFORMACION", "Contiene texto prohibido 'SIN INFORMACION'", "🚨 CRÍTICO", "igual_texto"),
                ("Página 1", "CÓDIGO SÍNTOMA", "AO42", val_ao42, "156", "Alerta: Código restringido '156'", "🚨 CRÍTICO", "igual_texto"),
                ("Página 1", "DESCRIPCIÓN DE LA CAUSA", "AR42", val_ar42, "OTROS", "Texto no permitido 'OTROS'", "🚨 CRÍTICO", "igual_texto"),
                ("Página 1", "CÓDIGO CAUSA", "BH42", val_bh42, ["6.6", "6,6", "7.1", "7,1"], "Código de falla crítico detectado", "🚨 CRÍTICO", "en_lista")
            ]

            for num_pag, nom_alerta, celda_id, val_real, val_prohibido, msg_error, crit, tipo_verif in alertas_especificas:
                if val_real is not None:
                    txt_real_clean = str(val_real).strip().upper()
                    disparar_alerta = False
                    
                    if tipo_verif == "igual_texto" and txt_real_clean == str(val_prohibido).upper():
                        disparar_alerta = True
                    elif tipo_verif == "en_lista" and str(val_real).strip() in val_prohibido:
                        disparar_alerta = True
                        msg_error = f"Contiene código de falla crítico ({str(val_real).strip()})"
                        
                    if disparar_alerta:
                        errores_en_este_archivo += 1
                        reporte_errores.append({
                            "Archivo Excel": nombre_archivo,
                            "Página": num_pag,
                            "Campo / Alerta": nom_alerta,
                            "Celdas Mapeadas": celda_id,
                            "Detalle del Error": msg_error,
                            "Criticidad": crit
                        })

            # -------------------------------------------------------------------------
            # 3. REGLA CONDICIONAL "REGISTRO INFORME SIMS"
            # -------------------------------------------------------------------------
            contiene_cambio = False
            celda_origen_cambio = ""
            for c_id in ["E205", "E211", "E216"]:
                val_c = hoja2[c_id].value
                if val_c and "cambio" in str(val_c).lower():
                    contiene_cambio = True
                    celda_origen_cambio = c_id
                    break
            
            val_b189 = hoja2["B189"].value
            txt_b189 = str(val_b189).strip().lower() if val_b189 is not None else ""
            es_b189_vacia = (val_b189 is None or txt_b189 in ["", "no", "none", "ㅤ"])

            if contiene_cambio and es_b189_vacia:
                errores_en_este_archivo += 1
                reporte_errores.append({
                    "Archivo Excel": nombre_archivo,
                    "Página": "Página 2",
                    "Campo / Alerta": "REGISTRO INFORME SIMS",
                    "Celdas Mapeadas": f"E205/B189 ({celda_origen_cambio})",
                    "Detalle del Error": f"Se detectó la palabra 'cambio' en {celda_origen_cambio}, por lo que el N° Pieza (B189) en Informe SIMS es obligatorio.",
                    "Criticidad": "🚨 CRÍTICO"
                })

            if errores_en_este_archivo > 0:
                archivos_con_errores += 1

        except Exception as e:
            reporte_errores.append({
                "Archivo Excel": archivo.name,
                "Página": "Error Técnico",
                "Campo / Alerta": "Error de estructura",
                "Celdas Mapeadas": "N/A",
                "Detalle del Error": f"No se pudo procesar: {str(e)}",
                "Criticidad": "🚨 CRÍTICO"
            })
            archivos_con_errores += 1

    return reporte_errores, total_archivos, archivos_con_errores


# 3. Interfaz de Usuario
archivos_subidos = st.file_uploader(
    "Selecciona las Órdenes de Trabajo en formato Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if archivos_subidos:
    if st.button("🚀 Iniciar Auditoría Masiva", type="primary"):
        with st.spinner("Procesando y auditando archivos..."):
            errores, total_arch, arch_con_error = auditar_archivos_masivos(archivos_subidos)
            df_errores = pd.DataFrame(errores)

        st.success("¡Auditoría finalizada con éxito!")
        st.write("---")

        # 4. Indicadores Clave (KPIs)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Archivos", total_arch)
        col2.metric("Con Errores / Alertas", arch_con_error)
        col3.metric("Archivos Correctos", total_arch - arch_con_error)
        col4.metric("Total Hallazgos", len(df_errores))

        st.write("---")

        # 5. Visualizaciones de Errores
        if not df_errores.empty:
            col_graph1, col_graph2 = st.columns(2)
            
            with col_graph1:
                st.subheader("Distribución por Criticidad")
                fig_crit = px.pie(
                    df_errores, 
                    names="Criticidad", 
                    color="Criticidad",
                    color_discrete_map={"🚨 CRÍTICO": "#FF4B4B", "⚠️ NO CRÍTICO": "#FFAA00"}
                )
                st.plotly_chart(fig_crit, use_container_width=True)

            with col_graph2:
                st.subheader("Top 15 Campos / Alertas más Frecuentes")
                top_campos = df_errores["Campo / Alerta"].value_counts().reset_index()
                top_campos.columns = ["Campo / Alerta", "Cantidad"]
                
                fig_bar = px.bar(
                    top_campos.head(15), 
                    x="Cantidad", 
                    y="Campo / Alerta", 
                    orientation="h", 
                    color="Cantidad",
                    color_continuous_scale="Reds"
                )
                fig_bar.update_layout(
                    yaxis={"autorange": "reversed"},
                    height=500
                )
                st.plotly_chart(fig_bar, use_container_width=True)

            # 6. Tabla Detallada con Filtros
            st.subheader("📊 Detalle de Hallazgos")
            
            filtro_criticidad = st.multiselect(
                "Filtrar por Criticidad:", 
                options=df_errores["Criticidad"].unique(), 
                default=df_errores["Criticidad"].unique()
            )
            
            df_filtrado = df_errores[df_errores["Criticidad"].isin(filtro_criticidad)]
            st.dataframe(df_filtrado, use_container_width=True)

            # 7. Descarga del Informe
            st.subheader("📥 Descargar Reporte")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_errores.to_excel(writer, index=False, sheet_name="Hallazgos")
            processed_data = output.getvalue()

            col_dl1, col_dl2 = st.columns(2)
            col_dl1.download_button(
                label="📄 Descargar Informe en Excel (.xlsx)",
                data=processed_data,
                file_name="Reporte_Auditoria_OTs.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            col_dl2.download_button(
                label="📝 Descargar Informe en CSV (.csv)",
                data=df_errores.to_csv(index=False).encode("utf-8"),
                file_name="Reporte_Auditoria_OTs.csv",
                mime="text/csv"
            )
        else:
            st.balloons()
            st.success("🎉 ¡Increíble! Todos los archivos procesados cumplen al 100% con los estándares y no registran errores.")
else:
    st.info("👋 Por favor, carga uno o más archivos de Órdenes de Trabajo en formato `.xlsx` arriba para empezar.")
