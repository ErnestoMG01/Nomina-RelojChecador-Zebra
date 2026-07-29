import streamlit as st
import pandas as pd
import io
import re
import json
import os
from datetime import datetime

st.set_page_config(page_title="Sistema automatico de nomina", layout="wide")

HORARIOS_FILE = "horarios_empleados.json"
FLEXIBLES_FILE = "horarios_flexibles.json"

def cargar_horarios():
    if os.path.exists(HORARIOS_FILE):
        with open(HORARIOS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def guardar_horarios(dict_horarios):
    with open(HORARIOS_FILE, "w", encoding="utf-8") as f:
        json.dump(dict_horarios, f, ensure_ascii=False, indent=4)

def cargar_horarios_flex():
    if os.path.exists(FLEXIBLES_FILE):
        with open(FLEXIBLES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def guardar_horarios_flex(dict_flex):
    with open(FLEXIBLES_FILE, "w", encoding="utf-8") as f:
        json.dump(dict_flex, f, ensure_ascii=False, indent=4)

def calcular_minutos_diferencia(hora_real_str, hora_teorica_str):
    try:
        t_real = datetime.strptime(hora_real_str, "%H:%M")
        t_teorica = datetime.strptime(hora_teorica_str, "%H:%M")
        return (t_real - t_teorica).total_seconds() / 60
    except:
        return 0

horarios_guardados = cargar_horarios()
horarios_flexibles = cargar_horarios_flex()

st.title("Sistema automatico de nómina")

tab_administrador, tab_trabajadores = st.tabs(["Vista Administrador", "Vista Trabajadores"])

# =========================================================================
# 1. PESTAÑA: ADMINISTRADOR
# =========================================================================
with tab_administrador:
    st.header("Panel de Administracion y Carga de Datos")
    
    uploaded_file = st.file_uploader("Sube el reporte del checador (.xls, .xlsx, .csv)", type=["xls", "xlsx", "csv"], key="admin_uploader")
    
    if uploaded_file is not None:
        try:
            try:
                content = uploaded_file.getvalue().decode('utf-8')
            except UnicodeDecodeError:
                content = uploaded_file.getvalue().decode('latin1', errors='ignore')
            
            lines = content.splitlines()
            
            if not any("Reporte de Eventos" in l for l in lines):
                try:
                    xls_tabs = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
                except Exception:
                    xls_tabs = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='openpyxl')
                
                lines = []
                for _, df_hoja in xls_tabs.items():
                    for index, row in df_hoja.iterrows():
                        line_str = ",".join([str(x).strip() if pd.notna(x) else "" for x in row.tolist()])
                        lines.append(line_str)
            
            asistencia_start = -1
            for i, line in enumerate(lines):
                if "Reporte de Eventos de Asistencia" in line:
                    asistencia_start = i
                    break
            
            if asistencia_start == -1:
                st.error("No se encontro la seccion 'Reporte de Eventos de Asistencia'.")
                st.stop()
                
            lines = lines[asistencia_start:]
            
            start_idx = -1
            for i, line in enumerate(lines):
                if "ID:" in line and "Nombre:" in line:
                    start_idx = i - 1
                    break
                    
            if start_idx == -1:
                st.error("No se pudo encontrar la estructura de empleados.")
                st.stop()
                
            days_line = lines[start_idx].strip().split(',')
            days = [d for d in days_line if d.strip() != '']
            
            empleados_detectados = []
            empleados_raw_data = {}
            
            for i in range(start_idx + 1, len(lines)):
                line = lines[i]
                if "ID:" in line and "Nombre:" in line:
                    header_parts = line.strip().split(',')
                    try:
                        nombre_idx = header_parts.index('Nombre:')
                        nombre = ""
                        for val in header_parts[nombre_idx+1:]:
                            if val.strip() != "":
                                nombre = val.strip().replace('*', '')
                                break
                    except ValueError:
                        continue
                        
                    if i + 1 < len(lines):
                        times_parts = lines[i+1].strip().split(',')
                        empleados_detectados.append(nombre)
                        empleados_raw_data[nombre] = (times_parts, days)
            
            # --- CONFIGURACION DE HORARIOS ---
            st.subheader("Configuracion de Horarios Individuales")
            st.caption("Marca 'Horario Flexible' para habilitar la modificacion diaria en la seccion de auditoria.")
            
            horarios_lista = []
            for emp in empleados_detectados:
                datos_h = horarios_guardados.get(emp, {"Entrada_Turno1": "09:00", "Salida_Turno1": "13:00", "Entrada_Turno2": "15:00", "Salida_Turno2": "19:00", "Flexible": False})
                horarios_lista.append({
                    "PERSONAL": emp,
                    "Horario Flexible": datos_h.get("Flexible", False),
                    "Entrada 1": datos_h.get("Entrada_Turno1", "09:00"),
                    "Salida 1": datos_h.get("Salida_Turno1", "13:00"),
                    "Entrada 2 (Opc)": datos_h.get("Entrada_Turno2", ""),
                    "Salida 2 (Opc)": datos_h.get("Salida_Turno2", "")
                })
            
            df_editor = pd.DataFrame(horarios_lista)
            edited_horarios = st.data_editor(df_editor, use_container_width=True, key="editor_horarios")
            
            nuevo_dict_horarios = {}
            for _, row in edited_horarios.iterrows():
                nuevo_dict_horarios[row["PERSONAL"]] = {
                    "Flexible": bool(row["Horario Flexible"]),
                    "Entrada_Turno1": str(row["Entrada 1"]).strip() if pd.notna(row["Entrada 1"]) else "",
                    "Salida_Turno1": str(row["Salida 1"]).strip() if pd.notna(row["Salida 1"]) else "",
                    "Entrada_Turno2": str(row["Entrada 2 (Opc)"]).strip() if pd.notna(row["Entrada 2 (Opc)"]) else "",
                    "Salida_Turno2": str(row["Salida 2 (Opc)"]).strip() if pd.notna(row["Salida 2 (Opc)"]) else ""
                }
            guardar_horarios(nuevo_dict_horarios)
            
            # --- PROCESAMIENTO MATEMATICO ---
            historial_punches_global = {}
            data_nomina_base = []
            
            for nombre, (times_parts, days) in empleados_raw_data.items():
                config = nuevo_dict_horarios.get(nombre, {})
                es_flexible = config.get("Flexible", False)
                horarios_flex_emp = horarios_flexibles.get(nombre, {})
                
                dias_trabajados = 0
                retardos = 0
                total_horas_extra = 0
                linhas_tabla_auditoria = []
                
                for day_idx, punches in enumerate(times_parts):
                    if day_idx < len(days) and punches.strip():
                        found_times = re.findall(r'\d{2}:\d{2}', punches)
                        
                        if found_times:
                            dias_trabajados += 1
                            dia_limpio = str(days[day_idx]).split('.')[0]
                            dia_str = f"Dia {dia_limpio}"
                            
                            entrada_flex = horarios_flex_emp.get(dia_str, {}).get("Entrada", "")
                            salida_flex = horarios_flex_emp.get(dia_str, {}).get("Salida", "")
                            
                            linhas_tabla_auditoria.append({
                                "Dia del Periodo": dia_str,
                                "Checadas Reales": "  -  ".join(found_times),
                                "Entrada Asignada": entrada_flex if es_flexible else "",
                                "Salida Asignada": salida_flex if es_flexible else ""
                            })
                            
                            first_punch = found_times[0]
                            last_punch = found_times[-1]
                            
                            if not es_flexible:
                                if config.get("Entrada_Turno1"):
                                    dif_minutos = calcular_minutos_diferencia(first_punch, config["Entrada_Turno1"])
                                    if dif_minutos >= 7:
                                        retardos += 1
                                
                                salida_oficial = config["Salida_Turno2"] if config.get("Salida_Turno2") else config["Salida_Turno1"]
                                if salida_oficial:
                                    minutos_extra_hoy = calcular_minutos_diferencia(last_punch, salida_oficial)
                                    if minutos_extra_hoy >= 60:
                                        total_horas_extra += int(minutos_extra_hoy // 60)
                            else:
                                # Logica estricta para el dia flexible editado
                                if entrada_flex:
                                    dif_minutos = calcular_minutos_diferencia(first_punch, entrada_flex)
                                    if dif_minutos >= 7:
                                        retardos += 1
                                if salida_flex:
                                    minutos_extra_hoy = calcular_minutos_diferencia(last_punch, salida_flex)
                                    if minutos_extra_hoy >= 60:
                                        total_horas_extra += int(minutos_extra_hoy // 60)
                                        
                historial_punches_global[nombre] = linhas_tabla_auditoria
                
                # Regla de Negocio: 13 Trabajados + 2 Descanso = 15 Dias
                faltas_calculadas = retardos // 3
                dias_descanso_fijos = 2
                
                data_nomina_base.append({
                    'PERSONAL': nombre,
                    'DIAS TRABAJADOS': dias_trabajados,
                    'DIAS DESCANSO': dias_descanso_fijos,
                    'TOTAL DIAS PAGADOS': (dias_trabajados + dias_descanso_fijos) - faltas_calculadas,
                    'RETARDOS': retardos,
                    'FALTAS': faltas_calculadas if faltas_calculadas > 0 else 0,
                    'PERMISOS': "",
                    'HORAS EXTRA': total_horas_extra if total_horas_extra > 0 else 0,
                    'OBSERVACIONES': f"-{faltas_calculadas} dia(s) por retardos" if faltas_calculadas > 0 else ""
                })

            # --- AUDITORIA Y EDICION FLEXIBLE ---
            st.markdown("---")
            st.subheader("Buscador y Verificador de Horarios Reales")
            empleado_a_verificar = st.selectbox("Selecciona un empleado para auditar sus registros:", ["-- Seleccionar --"] + sorted(empleados_detectados))
            
            if empleado_a_verificar != "-- Seleccionar --":
                es_flex = nuevo_dict_horarios.get(empleado_a_verificar, {}).get("Flexible", False)
                datos_auditoria = historial_punches_global.get(empleado_a_verificar, [])
                df_auditoria = pd.DataFrame(datos_auditoria)
                
                if not df_auditoria.empty:
                    if es_flex:
                        st.info("Modo Flexible Activado: Puedes asignar una hora de entrada y salida especifica para cada dia. El calculo se actualizara de inmediato al dar clic fuera de la celda.")
                        
                        edited_auditoria = st.data_editor(
                            df_auditoria,
                            column_config={
                                "Dia del Periodo": st.column_config.TextColumn(disabled=True),
                                "Checadas Reales": st.column_config.TextColumn(disabled=True),
                                "Entrada Asignada": st.column_config.TextColumn("Entrada Asignada (HH:MM)"),
                                "Salida Asignada": st.column_config.TextColumn("Salida Asignada (HH:MM)"),
                            },
                            use_container_width=True,
                            key=f"flex_editor_{empleado_a_verificar}"
                        )
                        
                        # Detectar y guardar cambios del horario flexible
                        cambios = False
                        dict_empleado_flex = horarios_flexibles.get(empleado_a_verificar, {})
                        
                        for _, row in edited_auditoria.iterrows():
                            dia = str(row["Dia del Periodo"])
                            ent = str(row["Entrada Asignada"]).strip() if pd.notna(row["Entrada Asignada"]) else ""
                            sal = str(row["Salida Asignada"]).strip() if pd.notna(row["Salida Asignada"]) else ""
                            
                            if ent in ["nan", "None"]: ent = ""
                            if sal in ["nan", "None"]: sal = ""
                            
                            if dict_empleado_flex.get(dia, {}).get("Entrada", "") != ent or dict_empleado_flex.get(dia, {}).get("Salida", "") != sal:
                                if dia not in dict_empleado_flex:
                                    dict_empleado_flex[dia] = {}
                                dict_empleado_flex[dia]["Entrada"] = ent
                                dict_empleado_flex[dia]["Salida"] = sal
                                cambios = True
                                
                        if cambios:
                            horarios_flexibles[empleado_a_verificar] = dict_empleado_flex
                            guardar_horarios_flex(horarios_flexibles)
                            st.rerun() # Recarga automatica para reflejar calculos
                            
                    else:
                        df_auditoria = df_auditoria.drop(columns=["Entrada Asignada", "Salida Asignada"])
                        st.dataframe(df_auditoria, use_container_width=True, hide_index=True)
                else:
                    st.warning("Este empleado no cuenta con registros.")
            
            # --- TABLA INTERACTIVA DE NOMINA FINAL ---
            st.markdown("---")
            st.subheader("Nomina Final (Editable)")
            st.caption("Calculo base de 15 dias (Trabajados + Descansos). Puedes modificar los permisos, faltas o dias totales antes de descargar.")
            
            df_nomina_pre = pd.DataFrame(data_nomina_base)
            
            df_nomina_editada = st.data_editor(
                df_nomina_pre,
                column_config={
                    "DIAS TRABAJADOS": st.column_config.NumberColumn("DIAS TRABAJADOS (Editar)"),
                    "DIAS DESCANSO": st.column_config.NumberColumn("DIAS DESCANSO (Editar)"),
                    "TOTAL DIAS PAGADOS": st.column_config.NumberColumn("TOTAL PAGADOS (Editar)"),
                    "RETARDOS": st.column_config.NumberColumn("RETARDOS (Editar)", min_value=0),
                    "FALTAS": st.column_config.NumberColumn("FALTAS (Editar)", min_value=0),
                    "PERMISOS": st.column_config.TextColumn("PERMISOS (Editar)"),
                    "OBSERVACIONES": st.column_config.TextColumn("OBSERVACIONES (Editar)")
                },
                use_container_width=True,
                key="editor_nomina_final"
            )
            
            st.session_state["df_nomina_procesada"] = df_nomina_editada
            
            df_excel = df_nomina_editada.replace({0: "", "0": "", 0.0: ""})
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_excel.to_excel(writer, index=False, sheet_name='Nomina')
                
            st.download_button(
                label="Descargar Nomina Procesada (.xlsx)",
                data=output.getvalue(),
                file_name="Nomina_Automatizada_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Error procesando el flujo: {e}")

# =========================================================================
# 2. PESTAÑA: TRABAJADORES
# =========================================================================
with tab_trabajadores:
    st.header("Consulta de Asistencias para Trabajadores")
    
    if "df_nomina_procesada" not in st.session_state:
        st.info("La informacion estara disponible cuando Administracion verifique los datos.")
    else:
        df_datos = st.session_state["df_nomina_procesada"]
        lista_personal = sorted(df_datos["PERSONAL"].unique())
        
        trabajador_seleccionado = st.selectbox("Busca y selecciona tu nombre completo:", ["-- Selecciona un nombre --"] + lista_personal)
        
        if trabajador_seleccionado != "-- Selecciona un nombre --":
            fila_empleado = df_datos[df_datos["PERSONAL"] == trabajador_seleccionado].iloc[0]
            
            st.markdown("---")
            st.subheader(f"Resumen Validado: {trabajador_seleccionado}")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                dias_t = int(fila_empleado['DIAS TRABAJADOS'] if pd.notna(fila_empleado['DIAS TRABAJADOS']) else 0)
                dias_d = int(fila_empleado['DIAS DESCANSO'] if pd.notna(fila_empleado['DIAS DESCANSO']) else 0)
                st.metric(label="Dias Totales Pagados", value=(dias_t + dias_d))
            
            with col2:
                retardos_totales = int(fila_empleado['RETARDOS'] if pd.notna(fila_empleado['RETARDOS']) else 0)
                st.metric(label="Retardos Acumulados", value=retardos_totales)
                
            with col3:
                faltas_totales = int(fila_empleado['FALTAS'] if pd.notna(fila_empleado['FALTAS']) else 0)
                retardos_restantes = retardos_totales % 3
                
                subtexto = f"({retardos_restantes}/3 para proxima falta)" if retardos_restantes > 0 else "Sin acumulacion"
                st.metric(label="Faltas Totales", value=faltas_totales, delta=subtexto, delta_color="inverse")
                
            with col4:
                he_valor = int(fila_empleado['HORAS EXTRA'] if pd.notna(fila_empleado['HORAS EXTRA']) else 0)
                st.metric(label="Horas Extras Aprobadas", value=f"{he_valor} hrs")
                
            if pd.notna(fila_empleado['OBSERVACIONES']) and str(fila_empleado['OBSERVACIONES']).strip() != "":
                st.warning(f"Observaciones de Administracion: {fila_empleado['OBSERVACIONES']}")
            if pd.notna(fila_empleado['PERMISOS']) and str(fila_empleado['PERMISOS']).strip() != "":
                st.info(f"Permisos Registrados: {fila_empleado['PERMISOS']}")
                
            st.markdown("---")
            st.caption("Nota: Las horas extras se contabilizan unicamente por horas completas cumplidas en el mismo dia. Cada 3 retardos generan 1 falta automatica.")