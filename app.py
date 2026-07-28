import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px

# Configuración de página ancha
st.set_page_config(page_title="Auditor Masivo de OTs", layout="wide")

st.title("📋 Auditor Masivo de Órdenes de Trabajo")
st.markdown("Carga múltiples archivos Excel en lote. Clasificación automática con motor de alertas y reglas de negocio avanzadas.")
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
            # 1. LISTA UNIFICADA DE CAMPOS VACÍOS (PÁGINA 1 Y PÁGINA 2)
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
                
                # --- PÁGINA 1: INFORMACIÓN DEL TRABAJO ---
                ("Página 1", "HORA INICIO ACTIVIDAD", hoja1, ["B42"], "🚨 CRÍTICO"),
                ("Página 1", "HORA TERMINO ACTIVIDAD", hoja1, ["F42"], "🚨 CRÍTICO"),
                ("Página 1", "Nº ORDEN SERVICIO", hoja1, ["J42"], "🚨 CRÍTICO"),
                ("Página 1", "CÓDIGO COMPONENTE SMCS", hoja1, ["Q42"], "⚠️ NO CRÍTICO"),
                ("Página 1", "CÓDIGO MODIFICADOR", hoja1, ["T42"], "⚠️ NO CRÍTICO"),
                ("Página 1", "CÓDIGO TRABAJO", hoja1, ["W42"], "⚠️ NO CRÍTICO"),
                ("Página 1", "TIPO TAREA", hoja1, ["BO42"], "⚠️ NO CRÍTICO"),
                ("Página 1", "TAREA PRINCIPAL", hoja1, ["BV42"], "⚠️ NO CRÍTICO"),

                # --- PÁGINA 2 ---
                ("Página 2", "DESCRIPCIÓN DE LA PIEZA", hoja2, ["E189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "CANTIDAD", hoja2, ["X189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "CÓDIGO SERVICIO", hoja2, ["AA189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "N° GRUPO", hoja2, ["AE189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "DESCRIPCIÓN DEL GRUPO", hoja2, ["AJ186", "AJ189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "¿Llegó al fin de su vida útil?", hoja2, ["AR189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "COMENTARIOS", hoja2, ["AU189"], "⚠️ NO CRÍTICO"),
                ("Página 2", "VALIDACIÓN OT: NOMBRE JEFE TURNO", hoja2, ["C238"], "🚨 CRÍTICO"),
                ("Página 2", "VALIDACIÓN OT: RUT JEFE TURNO", hoja2, ["C244"], "🚨 CRÍTICO"),
                ("Página 2", "TECNICO RESPONSABLE: NOMBRE", hoja2, ["BD239"], "🚨 CRÍTICO"),
                ("Página 2", "TECNICO RESPONSABLE: RUT", hoja2, ["BD243"], "🚨 CRÍTICO"),
            ]

            # Procesamos validación base de celdas vacías
            for num_pagina, nombre, hoja_obj, lista_celdas, criticidad in definicion_campos:
                campo_completado = False
                for celda_id in lista_celdas:
                    valor = hoja_obj[celda_id].value
                    texto_limpio = str(valor).strip().lower() if valor is not None else ""
                    if valor is not None and texto_limpio not in ["", "no", "none"]:
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
            # 2. MOTOR DE ALERTAS INTELIGENTES (VALORES PROHIBIDOS ESPECÍFICOS)
            # -------------------------------------------------------------------------
            val_z42 = hoja1["Z42"].value
            val_ao42 = hoja1["AO42"].value
            val_ar42 = hoja1["AR42"].value
            val_bh42 = hoja1["BH42"].value

            alertas_especificas = [
                ("Página 1", "DESCRIPCIÓN DEL SÍNTOMA", "Z42", val_z42, "SIN INFORMACION", "Contiene texto prohibido 'SIN INFORMACION'", "🚨 CRÍTICO", "igual_texto"),
                ("Página 1", "CÓDIGO SÍNTOMA", "AO42", val_ao42, "156", "Alerta: Se detectó el código restringido '156'", "🚨 CRÍTICO", "igual_texto"),
                ("Página 1", "DESCRIPCIÓN DE LA CAUSA", "AR42", val_ar42, "OTROS", "Contiene texto no permitido 'OTROS'", "🚨 CRÍTICO", "igual_texto"),
                ("Página 1", "CÓDIGO CAUSA CRÍTICO", "BH42", val_bh42, ["6.6", "6,6", "7.1", "7,1"], "Contiene código de falla crítico", "🚨 CRÍTICO", "en_lista")
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
            # 3. REGLA CONDICIONAL AVANZADA (PÁGINA 2 - PALABRA "CAMBIO")
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
            
            # Si B189 está vacío, evaluamos si es crítico o advertencia según la palabra "cambio"
            if val_b189 is None or txt_b189 in ["", "no", "none"]:
                errores_en_este_archivo += 1
                crit_final = "🚨 CRÍTICO" if contiene_cambio else "⚠️ NO CRÍTICO"
                msg_final = f"Obligatorio rellenar por palabra 'cambio' detectada en {celda_origen_cambio}" if contiene_cambio else "Celda vacía u omitida"
                
                reporte_errores.append({
                    "Archivo Excel": nombre_archivo,
                    "Página": "Página 2",
                    "Campo / Alerta": "N° PIEZA QUE FALLÓ",
                    "Celdas Mapeadas": "B189",
                    "Detalle del Error": msg_final,
                    "Criticidad": crit_final
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
