import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
import io

# 1. Configuración de página ancha
st.set_page_config(page_title="Auditor Masivo de OTs", layout="wide", page_icon="📋")

st.title("📋 Auditor Masivo de Órdenes de Trabajo")
st.markdown("Auditoría enfocada **exclusivamente en los 15 ítems prioritarios y críticos** con métricas de cumplimiento porcentual.")
st.write("---")

# Lista exacta de los 15 ítems prioritarios
ITEMS_CRITICOS_NOMBRES = [
    "HORÓMETRO",
    "MOTIVO DETENCIÓN DEL EQUIPO",
    "CÓDIGO COMPONENTE SMCS",
    "CÓDIGO MODIFICADOR",
    "CÓDIGO TRABAJO",
    "DESCRIPCIÓN DE SÍNTOMA",
    "CÓDIGO SÍNTOMA",
    "DESCRIPCIÓN DE LA CAUSA",
    "CÓDIGO CAUSA",
    "TIPO TAREA",
    "TAREA PRINCIPAL",
    "DESCRIPCIÓN DE ACTIVIDADES",
    "REGISTRO INFORME SIMS",
    "FIRMA JEFE TURNO (NOMBRE + RUT)",
    "FIRMA TÉCNICO RESPONSABLE (NOMBRE + RUT)"
]

# 2. Función de auditoría técnica
def auditar_archivos_masivos(lista_archivos):
    reporte_errores = []
    detalle_ots = []
    
    total_archivos = len(lista_archivos)
    archivos_con_errores = 0
    
    # Contadores de cumplimiento por cada uno de los 15 ítems críticos
    conteo_exito_item = {item: 0 for item in ITEMS_CRITICOS_NOMBRES}

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

            # --- Detección de datos SIMS (Pieza en B189 y Descripción en E189) ---
            val_b189 = hoja2["B189"].value
            val_e189 = hoja2["E189"].value
            txt_b189 = str(val_b189).strip() if val_b189 is not None else ""
            txt_e189 = str(val_e189).strip() if val_e189 is not None else ""
            
            tiene_sims_valido = bool(txt_b189 and txt_b189.lower() not in ["none", "no", "", "ㅤ"])

            # -------------------------------------------------------------------------
            # 1. EVALUACIÓN DE CELDAS VACÍAS
            # -------------------------------------------------------------------------
            definicion_campos_criticos = [
                ("Página 1", "HORÓMETRO", hoja1, ["G13"]),
                ("Página 1", "MOTIVO DETENCIÓN DEL EQUIPO", hoja1, ["AB25"]),
                ("Página 1", "CÓDIGO COMPONENTE SMCS", hoja1, ["Q42"]),
                ("Página 1", "CÓDIGO MODIFICADOR", hoja1, ["T42"]),
                ("Página 1", "CÓDIGO TRABAJO", hoja1, ["W42"]),
                ("Página 1", "DESCRIPCIÓN DE SÍNTOMA", hoja1, ["Z42"]),
                ("Página 1", "CÓDIGO SÍNTOMA", hoja1, ["AO42"]),
                ("Página 1", "DESCRIPCIÓN DE LA CAUSA", hoja1, ["AR42"]),
                ("Página 1", "CÓDIGO CAUSA", hoja1, ["BH42"]),
                ("Página 1", "TIPO TAREA", hoja1, ["BO42"]),
                ("Página 1", "TAREA PRINCIPAL", hoja1, ["BV42"]),
                ("Página 1", "DESCRIPCIÓN DE ACTIVIDADES", hoja1, ["Z42"]),
                ("Página 2", "FIRMA JEFE TURNO (NOMBRE + RUT)", hoja2, ["C238", "C244"]),
                ("Página 2", "FIRMA TÉCNICO RESPONSABLE (NOMBRE + RUT)", hoja2, ["BD239", "BD243"])
            ]

            # Evaluar los 14 ítems directos
            for num_pagina, nombre_item, hoja_obj, lista_celdas in definicion_campos_criticos:
                campo_completado = False
                for celda_id in lista_celdas:
                    valor = hoja_obj[celda_id].value
                    texto_limpio = str(valor).strip().lower() if valor is not None else ""
                    if valor is not None and texto_limpio not in ["", "no", "none", "ㅤ"]:
                        campo_completado = True
                        break
                
                if campo_completado:
                    conteo_exito_item[nombre_item] += 1
                else:
                    errores_en_este_archivo += 1
                    reporte_errores.append({
                        "Archivo Excel": nombre_archivo,
                        "Página": num_pagina,
                        "Campo / Alerta": nombre_item,
                        "Celdas Mapeadas": ", ".join(lista_celdas),
                        "Detalle del Error": f"El campo obligatorio '{nombre_item}' está vacío",
                        "Criticidad": "🚨 CRÍTICO"
                    })

            # -------------------------------------------------------------------------
            # 2. MOTOR DE ALERTAS DE CONTENIDO INVÁLIDO
            # -------------------------------------------------------------------------
            val_z42 = hoja1["Z42"].value
            val_ao42 = hoja1["AO42"].value
            val_ar42 = hoja1["AR42"].value
            val_bh42 = hoja1["BH42"].value

            alertas_especificas = [
                ("Página 1", "DESCRIPCIÓN DE SÍNTOMA", "Z42", val_z42, "SIN INFORMACION", "Contiene texto no permitido 'SIN INFORMACION'", "igual_texto"),
                ("Página 1", "CÓDIGO SÍNTOMA", "AO42", val_ao42, "156", "Alerta: Código restringido '156'", "igual_texto"),
                ("Página 1", "DESCRIPCIÓN DE LA CAUSA", "AR42", val_ar42, "OTROS", "Texto no permitido 'OTROS'", "igual_texto"),
                ("Página 1", "CÓDIGO CAUSA", "BH42", val_bh42, ["6.6", "6,6", "7.1", "7,1"], "Código de falla crítico detectado", "en_lista")
            ]

            for num_pag, nom_alerta, celda_id, val_real, val_prohibido, msg_error, tipo_verif in alertas_especificas:
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
                            "Criticidad": "🚨 CRÍTICO"
                        })

            # -------------------------------------------------------------------------
            # 3. EVALUACIÓN Y REGLA CONDICIONAL "REGISTRO INFORME SIMS"
            # -------------------------------------------------------------------------
            requiere_sims = False
            celda_origen_cambio = ""
            for c_id in ["E205", "E211", "E216"]:
                val_c = hoja2[c_id].value
                if val_c:
                    txt_c = str(val_c).lower()
                    if "cambi" in txt_c or "reemplaz" in txt_c:
                        requiere_sims = True
                        celda_origen_cambio = c_id
                        break

            # Evaluación de cumplimiento del ítem SIMS
            if requiere_sims:
                if tiene_sims_valido:
                    conteo_exito_item["REGISTRO INFORME SIMS"] += 1
                else:
                    errores_en_este_archivo += 1
                    reporte_errores.append({
                        "Archivo Excel": nombre_archivo,
                        "Página": "Página 2",
                        "Campo / Alerta": "REGISTRO INFORME SIMS",
                        "Celdas Mapeadas": f"{celda_origen_cambio} / B189",
                        "Detalle del Error": f"Se detectó acción de cambio en {celda_origen_cambio}, pero el N° Pieza (B189) en Informe SIMS está vacío.",
                        "Criticidad": "🚨 CRÍTICO"
                    })
            else:
                # Si no requería SIMS (no hubo cambio), cuenta como no defectuoso
                conteo_exito_item["REGISTRO INFORME SIMS"] += 1

            if errores_en_este_archivo > 0:
                archivos_con_errores += 1

            # Guardar resumen de la OT procesada
            detalle_ots.append({
                "Archivo Excel": nombre_archivo,
                "Estado OT": "🚨 Con Errores" if errores_en_este_archivo > 0 else "✅ Sin Errores",
                "Requiere SIMS": "Sí" if requiere_sims else "No",
                "Tiene Dato SIMS (B189)": "Sí" if tiene_sims_valido else "No",
                "N° Pieza (B189)": txt_b189 if tiene_sims_valido else "N/A",
                "Descripción Pieza (E189)": txt_e189 if tiene_sims_valido else "N/A"
            })

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

    return reporte_errores, detalle_ots, total_archivos, archivos_con_errores, conteo_exito_item


# 3. Interfaz de Usuario
archivos_subidos = st.file_uploader(
    "Selecciona las Órdenes de Trabajo en formato Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if archivos_subidos:
    if st.button("🚀 Iniciar Auditoría Masiva", type="primary"):
        with st.spinner("Procesando y auditando únicamente campos críticos..."):
            errores, detalle_ots, total_arch, arch_con_error, conteo_exito = auditar_archivos_masivos(archivos_subidos)
            df_errores = pd.DataFrame(errores)
            df_ots = pd.DataFrame(detalle_ots)

        st.success("¡Auditoría finalizada con éxito!")
        st.write("---")

        # 4. Indicadores Clave (KPIs)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Archivos", total_arch)
        col2.metric("Con Errores Críticos", arch_con_error)
        col3.metric("Archivos Correctos", total_arch - arch_con_error)
        col4.metric("Total Alertas Críticas", len(df_errores))

        st.write("---")

        # 5. Gráfico de Cumplimiento Porcentual (%) del Top 15 Ítems
        st.subheader("📊 Cumplimiento de los 15 Ítems Críticos (%)")
        
        data_cumplimiento = []
        for item in ITEMS_CRITICOS_NOMBRES:
            porcentaje = (conteo_exito[item] / total_arch) * 100 if total_arch > 0 else 0
            data_cumplimiento.append({
                "Campo / Ítem Crítico": item,
                "Cumplimiento (%)": round(porcentaje, 1),
                "Estado": "100%" if porcentaje == 100 else f"{round(porcentaje, 1)}%"
            })
        
        df_cumplimiento = pd.DataFrame(data_cumplimiento)
        
        # Color dinámico: Verde si es 100%, Rojo/Tomate si es menor a 100%
        fig_cumplimiento = px.bar(
            df_cumplimiento,
            x="Cumplimiento (%)",
            y="Campo / Ítem Crítico",
            orientation="h",
            text="Estado",
            color="Cumplimiento (%)",
            color_continuous_scale=[(0, "#FF4B4B"), (0.99, "#FF4B4B"), (1, "#28A745")],
            range_x=[0, 105]
        )
        
        fig_cumplimiento.update_traces(textposition="outside")
        fig_cumplimiento.update_layout(
            yaxis={"autorange": "reversed"}, 
            height=550,
            showlegend=False,
            coloraxis_showscale=False
        )
        st.plotly_chart(fig_cumplimiento, use_container_width=True)

        st.write("---")

        # 6. Pestañas para Análisis Detallado y Filtro SIMS
        tab1, tab2 = st.columns(2)

        # SECCIÓN DE ALERTAS Y ERRORES
        st.subheader("🚨 Detalle de Errores Críticos Detectados")
        if not df_errores.empty:
            st.dataframe(df_errores, use_container_width=True)
        else:
            st.success("🎉 No se detectaron errores críticos en los archivos procesados.")

        st.write("---")

        # SECCIÓN Y FILTRO DE OTs / SIMS
        st.subheader("🔎 Localizador de OTs y Registro SIMS")
        st.markdown("Usa los filtros a continuación para **ubicar la OT que sí cuenta con datos en el Informe SIMS** o revisar el estado por archivo:")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filtro_sims = st.selectbox(
                "Filtrar por presencia de datos SIMS (B189):", 
                ["Todos", "Solo con datos SIMS (B189 completa)", "Sin datos SIMS (B189 vacía)"]
            )
        with col_f2:
            filtro_estado = st.selectbox(
                "Filtrar por Estado de OT:", 
                ["Todos", "✅ Sin Errores", "🚨 Con Errores"]
            )

        df_ots_filtrado = df_ots.copy()
        
        if filtro_sims == "Solo con datos SIMS (B189 completa)":
            df_ots_filtrado = df_ots_filtrado[df_ots_filtrado["Tiene Dato SIMS (B189)"] == "Sí"]
        elif filtro_sims == "Sin datos SIMS (B189 vacía)":
            df_ots_filtrado = df_ots_filtrado[df_ots_filtrado["Tiene Dato SIMS (B189)"] == "No"]
            
        if filtro_estado != "Todos":
            df_ots_filtrado = df_ots_filtrado[df_ots_filtrado["Estado OT"] == filtro_estado]

        st.dataframe(df_ots_filtrado, use_container_width=True)

        # 7. Descarga del Informe
        st.subheader("📥 Descargar Reporte Completo")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_cumplimiento.to_excel(writer, index=False, sheet_name="Resumen_Cumplimiento")
            if not df_errores.empty:
                df_errores.to_excel(writer, index=False, sheet_name="Errores_Criticos")
            df_ots.to_excel(writer, index=False, sheet_name="Detalle_OTs_SIMS")
        processed_data = output.getvalue()

        col_dl1, col_dl2 = st.columns(2)
        col_dl1.download_button(
            label="📄 Descargar Reporte en Excel (.xlsx)",
            data=processed_data,
            file_name="Reporte_Auditoria_OTs_SIMS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        col_dl2.download_button(
            label="📝 Descargar Resumen de Errores en CSV (.csv)",
            data=df_errores.to_csv(index=False).encode("utf-8") if not df_errores.empty else b"",
            file_name="Reporte_Errores_Criticos.csv",
            mime="text/csv"
        )
else:
    st.info("👋 Por favor, carga uno o más archivos de Órdenes de Trabajo en formato `.xlsx` arriba para empezar.")
