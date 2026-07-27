import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px

# Configuración de página ancha
st.set_page_config(page_title="Auditor Masivo de OTs", layout="wide")

st.title("📋 Auditor Masivo de Órdenes de Trabajo")
st.markdown("Carga múltiples archivos Excel en lote. La aplicación reportará **únicamente** los campos que hayan quedado vacíos.")
st.write("---")

def auditar_archivos_masivos(lista_archivos):
    reporte_errores = []
    total_archivos = len(lista_archivos)
    archivos_con_errores = 0
    
    for archivo in lista_archivos:
        try:
            wb = openpyxl.load_workbook(archivo, data_only=True)
            hojas_disponibles = wb.sheetnames
            hoja1 = wb[hojas_disponibles[0]]
            hoja2 = wb[hojas_disponibles[1]] if len(hojas_disponibles) > 1 else hoja1
            
            nombre_archivo = archivo.name
            errores_en_este_archivo = 0

            # CONFIGURACIÓN PÁGINA 1
            campos_pagina1 = [
                ("EQUIPO", hoja1, ["G7"]),
                ("HOROMETRO", hoja1, ["G13"]),
                ("TURNO", hoja1, ["G19"]),
                ("ORDEN DE PEDIDO / SALIDA DE BODEGA", hoja1, ["G25"]),
                ("FECHA/HORA INICIO (Fecha)", hoja1, ["X9"]),
                ("FECHA/HORA INICIO (Hora)", hoja1, ["AB9"]),
                ("FECHA/HORA FINAL (Fecha)", hoja1, ["X11"]),
                ("FECHA/HORA FINAL (Hora)", hoja1, ["AB11"]),
                ("UBICACIÓN: TALLER", hoja1, ["R21"]),
                ("UBICACIÓN: TERRENO", hoja1, ["Y21"]),
            ]

            # CONFIGURACIÓN PÁGINA 2
            campos_pagina2 = [
                ("N° PIEZA QUE FALLÓ", hoja2, ["B189"]),
                ("DESCRIPCIÓN DE LA PIEZA", hoja2, ["E189"]),
                ("CANTIDAD", hoja2, ["X189"]),
                ("CÓDIGO SERVICIO", hoja2, ["AA189"]),
                ("N° GRUPO", hoja2, ["AE189"]),
                ("DESCRIPCIÓN DEL GRUPO", hoja2, ["AJ186", "AJ189"]),
                ("¿Llegó al fin de su vida útil?", hoja2, ["AR189"]),
                ("COMENTARIOS", hoja2, ["AU189"]),
                ("RESUMEN ANÁLISIS DE FALLA", hoja2, ["E205", "E211", "E216"]),
                ("VALIDACIÓN DE OT POR JEFE TURNO", hoja2, ["C238", "C244"]),
                ("TECNICO RESPONSABLE", hoja2, ["BD239", "BD243"]),
            ]

            todos_los_campos = [("Página 1", n, h, c) for n, h, c in campos_pagina1] + \
                               [("Página 2", n, h, c) for n, h, c in campos_pagina2]

            for num_pagina, nombre, hoja_obj, lista_celdas in todos_los_campos:
                campo_completado = False
                
                for celda_id in lista_celdas:
                    valor = hoja_obj[celda_id].value
                    texto_limpio = str(valor).strip().lower() if valor is not None else ""
                    
                    if valor is not None and texto_limpio not in ["", "no", "none"]:
                        campo_completado = True
                        break # Con una celda llena, el bloque es válido
                
                # SI NO CUMPLE, LO AGREGAMOS AL REPORTE DE ERRORES
                if not campo_completado:
                    errores_en_este_archivo += 1
                    reporte_errores.append({
                        "Archivo Excel": nombre_archivo,
                        "Página": num_pagina,
                        "Campo Incompleto": nombre,
                        "Celdas que debió revisar": ", ".join(lista_celdas),
                        "Estado": "❌ Vacío (Faltante)"
                    })
            
            if errores_en_este_archivo > 0:
                archivos_con_errores += 1

        except Exception as e:
            reporte_errores.append({
                "Archivo Excel": archivo.name,
                "Página": "Error Técnico",
                "Campo Incompleto": "No se pudo leer el archivo",
                "Celdas que debió revisar": "N/A",
                "Estado": f"⚠️ Error: {str(e)}"
            })
            archivos_con_errores += 1

    return pd.DataFrame(reporte_errores), total_archivos, archivos_con_errores

# NUEVO CARGADOR CON MULTIPLE_FILES ACTIVADO
archivos_cargados = st.file_uploader(
    "📂 Sube una o varias Órdenes de Trabajo al mismo tiempo (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if archivos_cargados:
    with st.spinner("Auditando lote de archivos..."):
        df_errores, cant_total, cant_con_errores = auditar_archivos_masivos(archivos_cargados)
        
        # MÁSTERS KPIs
        cant_perfectos = cant_total - cant_con_errores
        
        st.subheader("📊 Resumen de la Auditoría Masiva")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Archivos Subidos", cant_total)
        m2.metric("Archivos 100% Completos ✅", cant_perfectos)
        m3.metric("Archivos con Errores ❌", cant_con_errores)
        
        st.write("---")
        
        if df_errores.empty:
            st.success("🎉 ¡Excelente noticias! Todos los archivos cargados están rellenados al 100%. No se encontraron celdas vacías.")
            # Gráfico feliz (Todo correcto)
            fig = px.pie(values=[100], names=["100% Correctos"], color_discrete_sequence=["#2ecc71"], hole=0.4)
            st.plotly_chart(fig)
        else:
            col_izq, col_der = st.columns([1.4, 0.6])
            
            with col_izq:
                st.subheader("⚠️ Registro Centralizado de Campos Vacíos")
                st.markdown("La siguiente lista muestra exactamente qué archivo y qué celda requiere corrección:")
                st.dataframe(df_errores, use_container_width=True, height=400)
            
            with col_der:
                st.subheader("📉 Distribución de Archivos")
                datos_grafico = pd.DataFrame({
                    "Condición": ["Sin Errores", "Con Campos Vacíos"],
                    "Cantidad": [cant_perfectos, cant_con_errores]
                })
                fig = px.pie(
                    datos_grafico, 
                    values="Cantidad", 
                    names="Condición", 
                    hole=0.4,
                    color="Condición",
                    color_discrete_map={"Sin Errores": "#2ecc71", "Con Campos Vacíos": "#e74c3c"}
                )
                fig.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=350)
                st.plotly_chart(fig, use_container_width=True)
