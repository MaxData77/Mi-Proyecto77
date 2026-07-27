import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px

# Configuración de página ancha
st.set_page_config(page_title="Auditor de Órdenes de Trabajo", layout="wide")

st.title("📋 Auditor de Órdenes de Trabajo - Plantilla AMT")
st.markdown("Auditoría inteligente por coordenadas. Valida el cumplimiento mínimo de llenado en Página 1 y Página 2.")
st.write("---")

def auditar_plantilla(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    
    # Identificar hojas por posición (Página 1 = Índice 0, Página 2 = Índice 1)
    hojas_disponibles = wb.sheetnames
    hoja1 = wb[hojas_disponibles[0]]
    hoja2 = wb[hojas_disponibles[1]] if len(hojas_disponibles) > 1 else hoja1
    
    reporte = []

    # -------------------------------------------------------------------------
    # CONFIGURACIÓN PÁGINA 1: Se revisa una celda específica por campo
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # CONFIGURACIÓN PÁGINA 2: Grupos de celdas (Basta con que una tenga datos)
    # -------------------------------------------------------------------------
    campos_pagina2 = [
        ("N° PIEZA QUE FALLÓ", hoja2, ["B189"]),
        ("DESCRIPCIÓN DE LA PIEZA", hoja2, ["E189"]),
        ("CANTIDAD", hoja2, ["X189"]),
        ("CÓDIGO SERVICIO", hoja2, ["AA189"]),
        ("N° GRUPO", hoja2, ["AE189"]),
        ("DESCRIPCIÓN DEL GRUPO", hoja2, ["AJ186", "AJ189"]), # Agregada celda base por si acaso
        ("¿Llegó al fin de su vida útil?", hoja2, ["AR189"]),
        ("COMENTARIOS", hoja2, ["AU189"]),
        ("RESUMEN ANÁLISIS DE FALLA", hoja2, ["E205", "E211", "E216"]),
        ("VALIDACIÓN DE OT POR JEFE TURNO", hoja2, ["C238", "C244"]),
        ("TECNICO RESPONSABLE", hoja2, ["BD239", "BD243"]),
    ]

    # Unimos ambas listas para procesarlas uniformemente
    todos_los_campos = [("Página 1", n, h, c) for n, h, c in campos_pagina1] + \
                       [("Página 2", n, h, c) for n, h, c in campos_pagina2]

    for num_pagina, nombre, hoja_obj, lista_celdas in todos_los_campos:
        valores_detectados = []
        campo_completado = False
        
        # Escaneamos cada celda del grupo asignado
        for celda_id in lista_celdas:
            valor = hoja_obj[celda_id].value
            texto_limpio = str(valor).strip().lower() if valor is not None else ""
            
            # Si la celda contiene datos reales (y no está vacía o dice "no")
            if valor is not None and texto_limpio not in ["", "no", "none"]:
                campo_completado = True
                valores_detectados.append(f"[{celda_id}]: {str(valor)}")
            else:
                valores_detectados.append(f"[{celda_id}]: Vacío")
        
        # Construcción del estado final del campo evaluado
        if campo_completado:
            estado = "✅ Completado"
            cumple = 1
        else:
            estado = "❌ Vacío (Faltante)"
            cumple = 0
            
        reporte.append({
            "Página": num_pagina,
            "Campo Obligatorio": nombre,
            "Celdas Mapeadas": ", ".join(lista_celdas),
            "Lectura de Celdas": " | ".join(valores_detectados),
            "Estado": estado,
            "Cumple": cumple
        })

    return pd.DataFrame(reporte)

# Interfaz de carga en Streamlit
archivo_cargado = st.file_uploader("📂 Sube el archivo Excel de la Orden de Trabajo (.xlsx)", type=["xlsx"])

if archivo_cargado is not None:
    with st.spinner("Ejecutando auditoría de celdas..."):
        try:
            df_auditoria = auditar_plantilla(archivo_cargado)
            
            # Cálculos de KPIs globales
            total_campos = len(df_auditoria)
            completados = df_auditoria['Cumple'].sum()
            faltantes = total_campos - completados
            porcentaje = round((completados / total_campos) * 100, 1) if total_campos > 0 else 100
            
            # Despliegue de métricas en tarjetas
            st.subheader("📊 Diagnóstico General de Llenado")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Bloques Evaluados", total_campos)
            m2.metric("Bloques Correctos", completados)
            m3.metric("Bloques Incompletos", faltantes)
            m4.metric("Porcentaje de Cumplimiento", f"{porcentaje}%")
            
            st.write("---")
            
            # Layout de resultados: Tabla técnica a la izquierda, Gráfico interactivo a la derecha
            col_izq, col_der = st.columns([1.4, 0.6])
            
            with col_izq:
                st.subheader("📋 Detalle por Bloque de Información")
                filtro = st.radio("Filtrar registros por:", ["Ver Todo", "Solo Errores / Faltantes"], horizontal=True)
                
                if filtro == "Solo Errores / Faltantes":
                    df_mostrar = df_auditoria[df_auditoria['Cumple'] == 0]
                else:
                    df_mostrar = df_auditoria
                    
                st.dataframe(
                    df_mostrar[["Página", "Campo Obligatorio", "Celdas Mapeadas", "Estado", "Lectura de Celdas"]], 
                    use_container_width=True, 
                    height=450
                )
            
            with col_der:
                st.subheader("📈 Gráfico de Calidad")
                datos_pie = pd.DataFrame({
                    "Estado": ["Correctos", "Faltantes"],
                    "Cantidad": [completados, faltantes]
                })
                fig = px.pie(
                    datos_pie, 
                    values="Cantidad", 
                    names="Estado", 
                    hole=0.4,
                    color="Estado",
                    color_discrete_map={"Correctos": "#2ecc71", "Faltantes": "#e74c3c"}
                )
                fig.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=350)
                st.plotly_chart(fig, use_container_width=True)
                
        except Exception as e:
            st.error(f"Error al leer las pestañas del documento: {e}. Revisa que el archivo contenga al menos las dos páginas reglamentarias.")
