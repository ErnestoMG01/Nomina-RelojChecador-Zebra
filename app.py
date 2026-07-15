import streamlit as st
import pandas as pd
import io
import re
import json
import os
from datetime import datetime

st.set_page_config(page_title="Control de Nómina y Asistencias", page_icon="📊", layout="wide")

HORARIOS_FILE = "horarios_empleados.json"

def cargar_horarios():
    if os.path.exists(HORARIOS_FILE):
        with open(HORARIOS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def guardar_horarios(dict_horarios):
    with open(HORARIOS_FILE, "w", encoding="utf-8") as f:
        json.dump(dict_horarios, f, ensure_ascii=False, indent=4)

def calcular_minutos_diferencia(hora_real_str, hora_teorica_str):
    try:
        t_real = datetime.strptime(hora_real_str, "%H:%M")
        t_teorica = datetime.strptime(hora_teorica_str, "%H:%M")
        return (t_real - t_teorica).total_seconds() / 60
    except:
        return 0

horarios_guardados = cargar_horarios()

st.title("📊 Sistema Inteligente de Asistencias y Nómina")

tab_administrador, tab_trabajadores = st.tabs(["🔒 Vista Administrador", "👥 Vista Trabajadores"])

# =========================================================================
# 1. PESTAÑA: ADMINISTRADOR
# =========================================================================
with tab_administrador:
    st.header("Panel de Administración y Carga de Datos")
    
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
                st.error("❌ No se encontró la sección 'Reporte de Eventos de Asistencia'.")
                st.stop()
                
            lines = lines[asistencia_start:]
            
            start_idx = -1
            for i, line in enumerate(lines):
                if "ID:" in line and "Nombre:" in line:
                    start_idx = i - 1
                    break
                    
            if start_idx == -1:
                st.error("❌ No se pudo encontrar la estructura de empleados.")
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
            
            # --- SUB-SECCIÓN: GESTIÓN DE HORARIOS ---
            st.subheader("⚙️ Configuración y Memoria de Horarios Individuales")
            
            horarios_lista = []
            for emp in empleados_detectados:
                datos_h = horarios_guardados.get(emp, {"Entrada_Turno1": "09:00", "Salida_Turno1": "13:00", "Entrada_Turno2": "15:00", "Salida_Turno2": "19:00"})
                horarios_lista.append({
                    "PERSONAL": emp,
                    "Entrada 1": datos_h.get("Entrada_Turno1", "09:00"),
                    "Salida 1": datos_h.get("Salida_Turno1", "13:00"),
                    "Entrada 2 (Opcional)": datos_h.get("Entrada_Turno2", ""),
                    "Salida 2 (Opcional)": datos_h.get("Salida_Turno2", "")
                })
            
            df_editor = pd.DataFrame(horarios_lista)
            edited_df = st.data_editor(df_editor, use_container_width=True, key="editor_horarios")
            
            nuevo_dict_horarios = {}
            for _, row in edited_df.iterrows():
                nuevo_dict_horarios[row["PERSONAL"]] = {
                    "Entrada_Turno1": str(row["Entrada 1"]).strip(),
                    "Salida_Turno1": str(row["Salida 1"]).strip(),
                    "Entrada_Turno2": str(row["Entrada 2 (Opcional)"]).strip(),
                    "Salida_Turno2": str(row["Salida 2 (Opcional)"]).strip()
                }
            guardar_horarios(nuevo_dict_horarios)
            
            # --- NUEVA SUB-SECCIÓN: AUDITORÍA DE CHECADAS ORDENADAS ---
            st.markdown("---")
            st.subheader("🔍 Buscador y Verificador de Horarios Reales")
            st.markdown("Selecciona un trabajador para auditar sus marcajes del reloj ordenados de forma cronológica por día.")
            
            empleado_a_verificar = st.selectbox("Selecciona un empleado para corroborar sus datos:", ["-- Seleccionar --"] + sorted(empleados_detectados))
            
            # Procesamiento de la data fina para visualización
            historial_punches_global = {}
            data_nomina_final = []
            
            for nombre, (times_parts, days) in empleados_raw_data.items():
                config = nuevo_dict_horarios.get(nombre, {})
                dias_trabajados = 0
                retardos = 0
                total_minutos_extra = 0
                linhas_tabla_auditoria = []
                
                for day_idx, punches in enumerate(times_parts):
                    if day_idx < len(days) and punches.strip():
                        found_times = re.findall(r'\d{2}:\d{2}', punches)
                        
                        if found_times:
                            dias_trabajados += 1
                            # Agregar al registro ordenado
                            dia_limpio = str(days[day_idx]).split('.')[0] # Quitar el decimal .0 si existe
                            linhas_tabla_auditoria.append({
                                "Día del Periodo": f"Día {dia_limpio}",
                                "Checadas Registradas (Ordenadas)": "  ➔  ".join(found_times)
                            })
                            
                            # Evaluar Retardo
                            first_punch = found_times[0]
                            if config.get("Entrada_Turno1"):
                                dif_minutos = calcular_minutos_diferencia(first_punch, config["Entrada_Turno1"])
                                if dif_minutos >= 7:
                                    retardos += 1
                            
                            # Evaluar Horas Extra
                            last_punch = found_times[-1]
                            salida_oficial = config["Salida_Turno2"] if config.get("Salida_Turno2") else config["Salida_Turno1"]
                            if salida_oficial:
                                minutos_extra = calcular_minutos_diferencia(last_punch, salida_oficial)
                                if minutos_extra > 0:
                                    total_minutos_extra += minutos_extra
                                    
                historial_punches_global[nombre] = pd.DataFrame(linhas_tabla_auditoria)
                
                # REGLA DE NEGOCIO: 3 Retardos = 1 Falta
                faltas_por_retardos = retardos // 3
                horas_extra_decimal = round(total_minutos_extra / 60, 1) if total_minutos_extra > 0 else 0
                
                data_nomina_final.append({
                    'PERSONAL': nombre,
                    'DIAS TRABAJADOS': dias_trabajados,
                    'RETARDOS': retardos,
                    'FALTAS': faltas_por_retardos if faltas_por_retardos > 0 else 0,
                    'PERMISOS': "",
                    'HORAS EXTRA': horas_extra_decimal,
                    'DESCANSO TRABAJADO': "",
                    'OBSERVACIONES': f"{retardos} retardos registrados" if retardos > 0 else ""
                })
            
            # Mostrar la auditoría si se selecciona un nombre
            if empleado_a_verificar != "-- Seleccionar --":
                df_auditoria = historial_punches_global.get(empleado_a_verificar, pd.DataFrame())
                if not df_auditoria.empty:
                    st.dataframe(df_auditoria, use_container_width=True, hide_index=True)
                else:
                    st.warning("Este empleado no cuenta con registros de asistencia en el periodo.")
            
            # --- VISTA GENERAL DE LA NÓMINA ---
            st.markdown("---")
            st.subheader("📋 Formato de Nómina Calculado (Con regla de retardos aplicada)")
            
            df_nomina = pd.DataFrame(data_nomina_final)
            st.session_state["df_nomina_procesada"] = df_nomina
            
            df_presentacion = df_nomina.replace({0: "", "0": "", 0.0: ""})
            st.dataframe(df_presentacion, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_nomina.to_excel(writer, index=False, sheet_name='Nómina')
                
            st.download_button(
                label="📥 Descargar Formato Nómina.xlsx",
                data=output.getvalue(),
                file_name="Nomina_Automatizada_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Error procesando el flujo administrativo: {e}")

# =========================================================================
# 2. PESTAÑA: TRABAJADORES
# =========================================================================
with tab_trabajadores:
    st.header("👥 Consulta de Asistencias para Trabajadores")
    
    if "df_nomina_procesada" not in st.session_state:
        st.info("💡 La información de asistencia estará disponible aquí tan pronto como el administrador suba el reporte correspondiente en la pestaña anterior.")
    else:
        df_datos = st.session_state["df_nomina_procesada"]
        lista_personal = sorted(df_datos["PERSONAL"].unique())
        
        trabajador_seleccionado = st.selectbox("🔎 Busca y selecciona tu nombre completo:", ["-- Selecciona un nombre --"] + lista_personal)
        
        if trabajador_seleccionado != "-- Selecciona un nombre --":
            fila_empleado = df_datos[df_datos["PERSONAL"] == trabajador_seleccionado].iloc[0]
            
            st.markdown("---")
            st.subheader(f"Resumen de: {trabajador_seleccionado}")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric(label="📅 Días Laborados", value=int(fila_empleado['DIAS TRABAJADOS']))
            
            with col2:
                retardos_totales = int(fila_empleado['RETARDOS'])
                st.metric(label="⚠️ Retardos Acumulados", value=retardos_totales)
                
            with col3:
                # Mostrar equivalencia exacta de faltas por penalización
                faltas_totales = int(fila_empleado['FALTAS'])
                retardos_restantes = retardos_totales % 3
                
                subtexto_ayuda = f"Equivale a {faltas_totales} Falta(s)" if faltas_totales > 0 else "Sin faltas aplicadas"
                if retardos_restantes > 0 and faltas_totales > 0:
                    subtexto_ayuda += f" (+{retardos_restantes} retardo acumulado)"
                elif retardos_restantes > 0:
                    subtexto_ayuda += f" ({retardos_restantes}/3 para falta)"
                    
                st.metric(label="❌ Faltas (Por Retardos)", value=faltas_totales, delta=subtexto_ayuda, delta_color="inverse")
                
            with col4:
                he_valor = fila_empleado['HORAS EXTRA']
                st.metric(label="⏰ Horas Extras Registradas", value=f"{he_valor} hrs")
                
            st.markdown("---")
            st.caption("Nota informativa: Cada 3 retardos acumulados se genera automáticamente 1 falta en tu registro de nómina quincenal.")