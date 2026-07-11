import streamlit as st
import pandas as pd
import io
import re
import json
import os
from datetime import datetime, time, timedelta

st.set_page_config(page_title="Control de Nómina y Asistencias", page_icon="📊", layout="wide")

# ARCHIVO LOCAL PARA GUARDAR LA MEMORIA DE LOS HORARIOS
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
        diferencia = (t_real - t_teorica).total_seconds() / 60
        return diferencia
    except:
        return 0

# Cargar la base de datos de horarios en memoria
horarios_guardados = cargar_horarios()

# --- DISEÑO DE LA INTERFAZ ---
st.title("📊 Sistema Inteligente de Asistencias y Nómina")

# Separación clara de roles mediante pestañas nativas
tab_administrador, tab_trabajadores = st.tabs(["🔒 Vista Administrador", "👥 Vista Trabajadores"])

# =========================================================================
# 1. PESTAÑA: ADMINISTRADOR
# =========================================================================
with tab_administrador:
    st.header("Panel de Administración y Carga de Datos")
    st.markdown("Sube el archivo completo **1_report.xls** para procesar la información del periodo.")
    
    uploaded_file = st.file_uploader("Sube el reporte del checador (.xls, .xlsx, .csv)", type=["xls", "xlsx", "csv"], key="admin_uploader")
    
    if uploaded_file is not None:
        try:
            # Lectura híbrida robusta del archivo del checador
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
            
            # Aislar sección correcta
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
            
            # Recolectar lista de empleados únicos primero para la configuración de horarios
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
            
            # --- SUB-SECCIÓN: GESTIÓN DE HORARIOS CON MEMORIA ---
            st.subheader("⚙️ Configuración y Memoria de Horarios Individuales")
            st.markdown("Define las horas oficiales de entrada y salida. El sistema recordará estos valores para los siguientes archivos.")
            
            # Construir DataFrame para el st.data_editor interactivo
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
            
            # Mostrar tabla editable tipo Excel
            edited_df = st.data_editor(df_editor, use_container_width=True, key="editor_horarios")
            
            # Guardar automáticamente los cambios realizados en la tabla
            nuevo_dict_horarios = {}
            for _, row in edited_df.iterrows():
                nuevo_dict_horarios[row["PERSONAL"]] = {
                    "Entrada_Turno1": str(row["Entrada 1"]).strip(),
                    "Salida_Turno1": str(row["Salida 1"]).strip(),
                    "Entrada_Turno2": str(row["Entrada 2 (Opcional)"]).strip(),
                    "Salida_Turno2": str(row["Salida 2 (Opcional)"]).strip()
                }
            guardar_horarios(nuevo_dict_horarios)
            
            # --- CÁLCULO DE NÓMINA CON PARÁMETROS CONFIGURADOS ---
            data_nomina_final = []
            
            for nombre, (times_parts, days) in empleados_raw_data.items():
                config = nuevo_dict_horarios.get(nombre, {})
                
                dias_trabajados = 0
                retardos = 0
                total_minutos_extra = 0
                
                for day_idx, punches in enumerate(times_parts):
                    if day_idx < len(days) and punches.strip():
                        found_times = re.findall(r'\d{2}:\d{2}', punches)
                        
                        if found_times:
                            dias_trabajados += 1
                            
                            # 1. Evaluar Retardo (Comparar primera checada del día con la Entrada 1)
                            first_punch = found_times[0]
                            if config.get("Entrada_Turno1"):
                                dif_minutos = calcular_minutos_diferencia(first_punch, config["Entrada_Turno1"])
                                # Si pasa de los 7 minutos de tolerancia
                                if dif_minutos >= 7:
                                    retardos += 1
                            
                            # 2. Evaluar Horas Extra (Comparar última checada con la Salida Oficial definitiva)
                            last_punch = found_times[-1]
                            # Identificar cuál es su salida oficial al final de su jornada
                            salida_oficial = config["Salida_Turno2"] if config.get("Salida_Turno2") else config["Salida_Turno1"]
                            
                            if salida_oficial:
                                minutos_extra = calcular_minutos_diferencia(last_punch, salida_oficial)
                                # Solo sumamos si el tiempo es positivo (se quedó más tiempo)
                                if minutos_extra > 0:
                                    total_minutos_extra += minutos_extra
                
                # Convertir minutos totales acumulados a formato horas legibles (ej. 4.5 horas)
                horas_extra_decimal = round(total_minutos_extra / 60, 1) if total_minutos_extra > 0 else 0
                
                data_nomina_final.append({
                    'PERSONAL': nombre,
                    'DIAS TRABAJADOS': dias_trabajados,
                    'RETARDOS': retardos,
                    'FALTAS': "",
                    'PERMISOS': "",
                    'HORAS EXTRA': horas_extra_decimal,
                    'DESCANSO TRABAJADO': "",
                    'OBSERVACIONES': ""
                })
                
            df_nomina = pd.DataFrame(data_nomina_final)
            
            # Guardar el DataFrame procesado en el estado de Streamlit para transferirlo a la pestaña de trabajadores
            st.session_state["df_nomina_procesada"] = df_nomina
            
            # Limpieza visual de ceros para presentación
            df_presentacion = df_nomina.replace({0: "", "0": "", 0.0: ""})
            
            st.subheader("📋 Formato de Nómina Calculado")
            st.dataframe(df_presentacion, use_container_width=True)
            
            # Descargar archivo final Excel
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
# 2. PESTAÑA: TRABAJADORES (Interfaz Sencilla)
# =========================================================================
with tab_trabajadores:
    st.header("👥 Consulta de Asistencias para Trabajadores")
    st.markdown("Selecciona tu nombre para revisar de forma transparente tu resumen de asistencias del periodo actual.")
    
    # Validar si el administrador ya procesó los datos previamente en la pestaña 1
    if "df_nomina_procesada" not in st.session_state:
        st.info("💡 La información de asistencia estará disponible aquí tan pronto como el administrador suba el reporte correspondiente en la pestaña anterior.")
    else:
        df_datos = st.session_state["df_nomina_procesada"]
        lista_personal = sorted(df_datos["PERSONAL"].unique())
        
        # Selector amigable e intuitivo
        trabajador_seleccionado = st.selectbox("🔎 Busca y selecciona tu nombre completo:", ["-- Selecciona un nombre --"] + lista_personal)
        
        if trabajador_seleccionado != "-- Selecciona un nombre --":
            # Filtrar los datos correspondientes únicamente al empleado seleccionado
            fila_empleado = df_datos[df_datos["PERSONAL"] == trabajador_seleccionado].iloc[0]
            
            st.markdown("---")
            st.subheader(f"Resumen de: {trabajador_seleccionado}")
            
            # Creación de tarjetas visuales (Métricas) ultra limpias
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric(label="📅 Días Laborados", value=int(fila_empleado['DIAS TRABAJADOS']))
            
            with col2:
                retardos_valor = int(fila_empleado['RETARDOS']) if fila_empleado['RETARDOS'] != "" else 0
                st.metric(label="⚠️ Retardos Acumulados", value=retardos_valor, 
                          delta="Tolerancia: Minuto 7" if retardos_valor > 0 else "¡Excelente puntualidad!")
                
            with col3:
                he_valor = fila_empleado['HORAS EXTRA'] if fila_empleado['HORAS EXTRA'] != "" else 0
                st.metric(label="⏰ Horas Extras Registradas", value=f"{he_valor} hrs")
                
            st.markdown("---")
            st.caption("Nota informativa: Si notas alguna discrepancia con tus checadas reales, acude con el encargado de Administración para su revisión manual.")