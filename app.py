import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io

st.set_page_config(page_title="Pharmadvisor | Validador Metodológico", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# MÓDULO DE SEGURIDAD Y CONTRASEÑA
# ==========================================
PASSWORD_CORPORATIVA = "Pharmadvisor2026*"

def verificar_password():
    if "password_correcta" not in st.session_state:
        st.session_state["password_correcta"] = False

    if st.session_state["password_correcta"]:
        return True

    st.markdown("""
        <style>
        .stApp { background-color: #F8F9FA; color: #2C3E50; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .login-card { background: #FFFFFF; padding: 40px; border-radius: 12px; border: 1px solid #E2E8F0; max-width: 450px; margin: 80px auto; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.05); }
        .login-title { color: #E6007E; font-size: 26px; font-weight: bold; margin-bottom: 10px; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("""
        <div class="login-card">
            <div class="login-title">Pharm<span style="color: #2C3E50;">Advisor</span></div>
            <p style="color: #718096; font-size: 14px;">Acceso Restringido - Validador Metodológico Multicliente</p>
        </div>
    """, unsafe_allow_html=True)
    
    col_l1, col_l2, col_l3 = st.columns([1, 1.2, 1])
    with col_l2:
        input_pass = st.text_input("Ingrese la Contraseña Corporativa", type="password", key="pwd_input")
        if st.button("Iniciar Sesión", use_container_width=True):
            if input_pass == PASSWORD_CORPORATIVA:
                st.session_state["password_correcta"] = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta. Por favor verifique con el administrador.")
    return False

if not verificar_password():
    st.stop()

# ==========================================
# APLICACIÓN PRINCIPAL (ESTILO ESSITY / LEUKOACADEMY)
# ==========================================
st.markdown("""
    <style>
    .stApp { background-color: #F4F6F9; color: #2C3E50; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .ph-header { display: flex; justify-content: space-between; align-items: center; padding-bottom: 15px; border-bottom: 2px solid #E2E8F0; margin-bottom: 25px; background: #FFFFFF; padding: 20px 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
    .ph-title { color: #E6007E; font-size: 26px; font-weight: bold; margin: 0; }
    .section-divider { margin-top: 40px; margin-bottom: 40px; border: 0; height: 1px; background: #CBD5E0; }
    div[data-testid="stMetric"] { background-color: #FFFFFF; border: 1px solid #E2E8F0; padding: 15px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="ph-header">
        <div>
            <h1 class="ph-title">Validador Metodológico | Programa Leukoacademy</h1>
            <span style="color: #718096; font-size: 13px;">Auditoría Integral de Registro, Calidad y Muestras | Essity & Pharmadvisor</span>
        </div>
        <div style="text-align: right;">
            <span style="color: #E6007E; font-weight: bold; font-size: 20px;">Pharm<span style="color: #2C3E50;">Advisor</span></span>
        </div>
    </div>
""", unsafe_allow_html=True)

st.sidebar.subheader("1. Carga de Reportes")
uploaded_file = st.sidebar.file_uploader("Cargar Listado de Visitas (.xlsx)", type=["xlsx"], key="visitas_validador")
uploaded_impactos = st.sidebar.file_uploader("Cargar Impactos Promocionales (.xlsx)", type=["xlsx"], key="impactos_validador")

st.sidebar.markdown("---")
st.sidebar.subheader("2. Parámetros del Informe Gerencial")
cliente_input = st.sidebar.text_input("Nombre del Cliente / Cuenta", value="Essity / Leukoacademy", key="input_cliente_rep")
lineas_input = st.sidebar.text_input("Líneas / Segmento del Informe", value="Línea Institucional y Médica", key="input_lineas_rep")

@st.cache_data
def cargar_datos_hoja(source, hoja_nombre):
    try:
        df = pd.read_excel(source, sheet_name=hoja_nombre)
        df.columns = df.columns.astype(str).str.strip()
        return df
    except Exception as e:
        return None

@st.cache_data
def procesar_archivo_impactos(file):
    if file is None: return None
    try:
        df_raw = pd.read_excel(file, sheet_name='Consulta', header=None)
        row9 = df_raw.iloc[9].values
        row10 = df_raw.iloc[10].values
        
        specialties = []
        curr_spec = "GENERAL"
        for val in row9:
            if pd.notna(val) and str(val).strip() != '': 
                curr_spec = str(val).strip()
            specialties.append(curr_spec)
            
        records = []
        current_region = "DESCONOCIDA"
        
        for idx in range(11, len(df_raw)):
            col0 = df_raw.iloc[idx, 0]
            col1 = df_raw.iloc[idx, 1]
            if pd.isna(col0) or pd.isna(col1): continue
            col0_str = str(col0).strip()
            
            if col0_str.lower().startswith("region") or col0_str.lower().startswith("región"):
                current_region = col0_str.replace("Región", "").replace("Region", "").strip()
                continue
            if col0_str.startswith("DISTRITO"): continue
            
            rep_text = col0_str
            producto = str(col1).strip()
            coordinacion = "COORDINACION CL" if "CL" in current_region.upper() or "CL" in rep_text.upper() or "ANTIOQUIA" in rep_text.upper() or "COSTA" in rep_text.upper() or "SANTANDER" in rep_text.upper() else "COORDINACION PS"
            
            for c_idx in range(2, len(row10), 3):
                if c_idx >= df_raw.shape[1]: break
                spec = specialties[c_idx]
                imp_esp = df_raw.iloc[idx, c_idx]
                imp_real = df_raw.iloc[idx, c_idx + 1]
                
                if pd.notna(imp_esp) and str(imp_esp).strip() != '-' and str(imp_esp).strip() != '':
                    try:
                        esp_val = float(imp_esp)
                        real_val = float(imp_real) if pd.notna(imp_real) and str(imp_real).strip() != '-' else 0.0
                        records.append({'Región': coordinacion, 'Representante_Raw': rep_text, 'Producto': producto, 'Especialidad': spec, 'Imp_Esp': esp_val, 'Imp_Real': real_val})
                    except: 
                        pass
        return pd.DataFrame(records)
    except Exception as e: 
        return None

df_visitas = None
source_file = uploaded_file if uploaded_file is not None else ('listado_visitas_2026-09-07_11-44-08.xlsx' if os.path.exists('listado_visitas_2026-09-07_11-44-08.xlsx') else None)

if source_file is not None:
    try:
        xls = pd.ExcelFile(source_file)
        hoja_seleccionada = st.sidebar.selectbox("Seleccione la Hoja de Visitas", options=xls.sheet_names, key="select_hoja_excel")
        df_visitas = cargar_datos_hoja(source_file, hoja_seleccionada)
    except Exception as e:
        df_visitas = None

df_impactos = procesar_archivo_impactos(uploaded_impactos)

if df_visitas is not None:
    def buscar_columna(df, posibles_nombres):
        cols_lower = {c.lower().strip(): c for c in df.columns}
        for p in posibles_nombres:
            if p.lower().strip() in cols_lower:
                return cols_lower[p.lower().strip()]
        return None

    col_reg = buscar_columna(df_visitas, ['Región', 'Region', 'Coordinación', 'Coordinacion'])
    col_ciclo = buscar_columna(df_visitas, ['Ciclo', 'Periodo', 'Mes'])
    col_lin = buscar_columna(df_visitas, ['Línea', 'Linea', 'Estrategia'])
    col_rep = buscar_columna(df_visitas, ['Representante', 'Asesor', 'Ejecutivo'])
    col_visita = buscar_columna(df_visitas, ['Cod. visita', 'Cod visita', 'Código visita', 'Codigo visita', 'Id visita', 'Visita'])
    col_fecha = buscar_columna(df_visitas, ['Fecha visita', 'Fecha', 'Date'])
    col_hora = buscar_columna(df_visitas, ['Hora digitado', 'Hora', 'Time'])
    col_med = buscar_columna(df_visitas, ['Nombre farmacia o cliente', 'Farmacia', 'Médicos', 'Medicos', 'Cliente', 'Institución', 'Institucion', 'Nombre'])
    col_obj = buscar_columna(df_visitas, ['Objetivo', 'Obj', 'Objetivos'])
    col_com = buscar_columna(df_visitas, ['Comentario', 'Comentarios', 'Observación', 'Observacion'])
    col_acomp = buscar_columna(df_visitas, ['Visita acompañada', 'Visita acompanada', 'Acompañada', 'Acompañamiento'])
    col_muestra = buscar_columna(df_visitas, ['Material entregado', 'Muestra', 'Muestras'])
    col_cant_muestra = buscar_columna(df_visitas, ['Cantidad de material', 'Cantidad', 'Cant'])
    col_impacto_prod = buscar_columna(df_visitas, ['Impactos', 'Impacto', 'Producto'])

    if not col_visita: col_visita = df_visitas.columns[0]
    if not col_rep: col_rep = df_visitas.columns[0]
    if not col_reg: col_reg = df_visitas.columns[0]
    if not col_com: 
        df_visitas['Comentario'] = '-'
        col_com = 'Comentario'
    if not col_obj: 
        df_visitas['Objetivo'] = '-'
        col_obj = 'Objetivo'

    st.sidebar.markdown("---")
    st.sidebar.subheader("3. Filtros en Cascada")
    
    regiones = sorted(df_visitas[col_reg].dropna().unique()) if col_reg else []
    selected_regiones = st.sidebar.multiselect("Región / Coordinación", options=regiones, default=regiones, key="filtro_reg")
    df_f1 = df_visitas[df_visitas[col_reg].isin(selected_regiones)] if col_reg and selected_regiones else df_visitas
    
    ciclos = sorted(df_f1[col_ciclo].dropna().unique()) if col_ciclo else []
    selected_ciclos = st.sidebar.multiselect("Ciclo", options=ciclos, default=ciclos, key="filtro_ciclo")
    df_f2 = df_f1[df_f1[col_ciclo].isin(selected_ciclos)] if ciclos and selected_ciclos else df_f1
    
    lineas = sorted(df_f2[col_lin].dropna().unique()) if col_lin else []
    selected_lineas = st.sidebar.multiselect("Línea Estratégica", options=lineas, default=lineas, key="filtro_lin")
    df_f3 = df_f2[df_f2[col_lin].isin(selected_lineas)] if col_lin and selected_lineas else df_f2
    
    representantes = sorted(df_f3[col_rep].dropna().unique()) if col_rep else []
    selected_reps = st.sidebar.multiselect("Representante", options=representantes, default=representantes, key="filtro_rep")
    df_filtered = df_f3[df_f3[col_rep].isin(selected_reps)] if col_rep and selected_reps else df_f3
    
    df_unique = df_filtered.drop_duplicates(subset=[col_visita]).copy()
    
    df_unique['Comentario_Clean'] = df_unique[col_com].fillna('').astype(str).str.strip().str.lower()
    df_unique['Rep_Comentario_Count'] = df_unique.groupby([col_rep, 'Comentario_Clean'])[col_visita].transform('count')
    
    def check_repetido_individual(row):
        com = row['Comentario_Clean']
        if com in ['', 'nan', 'none', '-']: return False
        return row['Rep_Comentario_Count'] > 1

    df_unique['Es_Repetido'] = df_unique.apply(check_repetido_individual, axis=1)

    # ==========================================
    # 1. DISTRIBUCIÓN DE VISITAS ÚNICAS POR FRANJA HORARIA
    # ==========================================
    st.subheader("1. Distribución de Visitas Únicas por Franja Horaria")
    st.markdown("<span style='color: #718096;'>Auditoría global de puntualidad y franja de registro en visitas únicas.</span>", unsafe_allow_html=True)
    st.markdown("---")

    def parse_hour_slot(hora_str):
        try:
            dt = pd.to_datetime(str(hora_str), format='%H:%M:%S', errors='coerce')
            if pd.isna(dt): return 'No Registrada'
            h = dt.hour
            if 5 <= h < 12: return '1. Mañana (05:00 - 12:00)'
            elif 12 <= h < 18: return '2. Tarde (12:00 - 18:00)'
            else: return '3. Noche / Madrugada (18:00 - 05:00)'
        except: return 'No Registrada'

    if col_hora and col_hora in df_unique.columns:
        df_unique['Franja_Horaria'] = df_unique[col_hora].apply(parse_hour_slot)
        total_visitas = len(df_unique)
        nocturnas_fuera = len(df_unique[df_unique['Franja_Horaria'] == '3. Noche / Madrugada (18:00 - 05:00)'])
        pct_nocturnas = (nocturnas_fuera / total_visitas * 100) if total_visitas > 0 else 0

        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric("Total Visitas Únicas", f"{total_visitas:,}")
        kpi2.metric("Registros en Noche / Madrugada", f"{nocturnas_fuera:,}")
        kpi3.metric("% Alerta Acumulación Nocturna", f"{pct_nocturnas:.1f}%")
        st.markdown("---")

        franja_counts = df_unique['Franja_Horaria'].value_counts().reset_index()
        franja_counts.columns = ['Franja Horaria', 'Total Visitas']
        franja_counts = franja_counts.sort_values(by='Franja Horaria')
        color_map = {'1. Mañana (05:00 - 12:00)': '#2ECC71', '2. Tarde (12:00 - 18:00)': '#3182CE', '3. Noche / Madrugada (18:00 - 05:00)': '#E53E3E', 'No Registrada': '#A0AEC0'}

        fig_pie = px.pie(franja_counts, names='Franja Horaria', values='Total Visitas', hole=0.5, template='plotly_white', title="Semáforo de Oportunidad en el Registro de Visitas", color='Franja Horaria', color_discrete_map=color_map)
        fig_pie.update_traces(textinfo='percent+value', textposition='inside', insidetextorientation='horizontal', textfont_size=13)
        fig_pie.update_layout(paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF', height=450, legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5), margin=dict(t=50, b=80, l=20, r=20))
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("No se detectó la columna de hora de registro para la distribución de franjas horarias.")

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ==========================================
    # 2. COMPARATIVA DEL % DE COPIADO INTERNO POR COORDINACIÓN
    # ==========================================
    st.subheader("2. Comparativa del % de Copiado Interno por Coordinación")
    st.markdown("<span style='color: #718096;'>Comparativo del índice de autorrepetición agrupado por cada región o coordinación.</span>", unsafe_allow_html=True)
    st.markdown("---")

    df_reg_copia = df_unique.groupby(col_reg).agg(Total_Visitas=(col_visita, 'count'), Comentarios_Repetidos=('Es_Repetido', lambda x: int(x.sum()))).reset_index()
    df_reg_copia.columns = ['Región', 'Total_Visitas', 'Comentarios_Repetidos']
    df_reg_copia['Pct_Copia'] = (df_reg_copia['Comentarios_Repetidos'] / df_reg_copia['Total_Visitas'] * 100).round(1)
    df_reg_copia = df_reg_copia.sort_values(by='Pct_Copia', ascending=False)

    fig_reg_copia = px.bar(df_reg_copia, x='Región', y='Pct_Copia', text='Pct_Copia', template='plotly_white', title="Índice de Copiado Interno por Coordinación / Región", color='Pct_Copia', color_continuous_scale=[[0.0, '#2ECC71'], [0.05, '#2ECC71'], [0.30, '#3182CE'], [0.31, '#E53E3E'], [1.0, '#9B2C2C']], range_color=[0, 100])
    fig_reg_copia.update_traces(texttemplate='%{text}%', textposition='outside', textfont_size=12)
    fig_reg_copia.update_layout(paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF', height=450, xaxis_title="Coordinación / Región", yaxis_title="Índice de Copiado Interno (%)", yaxis=dict(range=[0, 110]), margin=dict(t=50, b=80, l=40, r=20))
    st.plotly_chart(fig_reg_copia, use_container_width=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ==========================================
    # 3. ÍNDICE DE AUTORREPETICIÓN POR REPRESENTANTE
    # ==========================================
    st.subheader("3. Auditoría de Calidad: Índice de Autorrepetición por Representante")
    st.markdown("<span style='color: #718096;'>Evaluación estricta de cuántas veces cada representante recicla sus propios comentarios entre sus visitas.</span>", unsafe_allow_html=True)
    st.markdown("---")
    
    total_comentarios = len(df_unique)
    comentarios_repetidos = int(df_unique['Es_Repetido'].sum())
    pct_copia = (comentarios_repetidos / total_comentarios * 100) if total_comentarios > 0 else 0
    
    ckpi1, ckpi2, ckpi3 = st.columns(3)
    ckpi1.metric("Visitas Únicas Evaluadas", f"{total_comentarios:,}")
    ckpi2.metric("Comentarios Autorrepetidos", f"{comentarios_repetidos:,}")
    ckpi3.metric("% Índice Global de Autorrepetición", f"{pct_copia:.1f}%")
    st.markdown("---")
    
    rep_copia = df_unique.groupby(col_rep).agg(Total_Visitas=(col_visita, 'count'), Comentarios_Repetidos=('Es_Repetido', lambda x: int(x.sum()))).reset_index()
    rep_copia.columns = ['Representante', 'Total_Visitas', 'Comentarios_Repetidos']
    rep_copia['Pct_Copia'] = (rep_copia['Comentarios_Repetidos'] / rep_copia['Total_Visitas'] * 100).round(1)
    rep_copia = rep_copia.sort_values(by='Pct_Copia', ascending=True)
    
    altura_grafico = max(450, len(rep_copia) * 25)
    
    fig_bar_copia = px.bar(
        rep_copia, x='Pct_Copia', y='Representante', text='Pct_Copia', template='plotly_white',
        title="<b>Índice de Autorrepetición (%) por Representante</b>",
        color='Pct_Copia', color_continuous_scale=[[0.0, '#2ECC71'], [0.05, '#2ECC71'], [0.30, '#3182CE'], [0.31, '#E53E3E'], [1.0, '#9B2C2C']],
        range_color=[0, 100], orientation='h'
    )
    fig_bar_copia.update_traces(texttemplate='%{text}%', textposition='outside', textfont_size=11)
    fig_bar_copia.update_layout(
        paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF', height=altura_grafico,
        xaxis_title="Índice de Autorrepetición (%)", yaxis_title="Representante",
        xaxis=dict(range=[0, 115]), margin=dict(t=50, b=50, l=150, r=20)
    )
    st.plotly_chart(fig_bar_copia, use_container_width=True)
    
    with st.expander("Ver listado de visitas únicas con comentarios repetidos"):
        cols_ver = [c for c in [col_reg, col_rep, col_fecha, col_med, col_com] if c is not None]
        st.dataframe(df_unique[df_unique['Es_Repetido'] == True][cols_ver].head(50), use_container_width=True, hide_index=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ==========================================
    # 4. AUDITORÍA DE CALIDAD METODOLÓGICA
    # ==========================================
    st.subheader("4. Auditoría de Calidad Metodológica (Técnica de Ventas - Visitas Únicas)")
    st.markdown("<span style='color: #718096;'>Evaluación inteligente de la Fase 2 (Comentarios) y Fase 1/3 (Objetivos) con penalización automática a Alerta ante registros con copy-paste.</span>", unsafe_allow_html=True)
    st.markdown("---")
    
    df_audit_tec = df_unique.copy()
    df_audit_tec['Com_Text'] = df_audit_tec[col_com].fillna('').astype(str).str.strip()
    df_audit_tec['Obj_Text'] = df_audit_tec[col_obj].fillna('').astype(str).str.strip()
    
    palabras_prohibidas_com = ['', '-', 'nan', 'none', 'nat', '0', 'ok', 'bien', 'excelente', 'sin novedad', 'atendió bien']
    
    def calificar_y_justificar_comentario_flexible(txt, es_rep):
        t_low = txt.lower()
        if t_low in palabras_prohibidas_com or len(txt) < 8:
            return "Alerta: Vacío o Genérico", "El comentario está vacío o usa expresiones genéricas."
        if es_rep:
            return "Alerta: Penalizado por Copy-Paste", "El texto está repetido idénticamente en múltiples visitas, se invalida."
        palabras_alta_calidad = ['acepta', 'indiferente', 'objeción', 'objecion', 'escepticismo', 'evasivo', 'acuerdo', 'compromiso', 'diferencia', 'valor', 'claro', 'dudas', 'explica', 'explicó', 'habla', 'habló', 'revisa', 'revisó', 'conoce', 'conoció', 'prescribe', 'prescribirá']
        if any(w in t_low for w in palabras_alta_calidad):
            return "Alta Calidad (Técnica Aplicada)", "El comentario evidencia de forma sólida la Fase 2."
        else:
            return "Regular (Superficial)", "El texto relata la visita pero carece de profundidad."

    palabras_actividades = ['entregar', 'visitar', 'saludar', 'llamar', 'dejar', 'muestra', 'material', 'obsequio']
    
    def calificar_y_justificar_objetivo(txt):
        t_low = txt.lower()
        if t_low in palabras_prohibidas_com or len(txt) < 8:
            return "Alerta: Sin Objetivo Definido", "El campo de objetivo está vacío."
        elif any(t_low.startswith(act) for act in palabras_actividades):
            return "Alerta: Confunde Actividad", "Describe una tarea logística en lugar de un comportamiento SMART."
        elif any(w in t_low for w in ['iniciar', 'reiniciar', 'aumentar', 'sostener', 'mantener', 'reemplazar', 'posicionar', 'evaluar', 'prescripción', 'uso']):
            return "Alta Calidad (SMART)", "El objetivo está formulado correctamente."
        else:
            return "Regular (Poco Específico)", "El objetivo menciona una intención pero carece de precisión."

    res_com = [calificar_y_justificar_comentario_flexible(row['Com_Text'], row['Es_Repetido']) for _, row in df_audit_tec.iterrows()]
    df_audit_tec['Calidad_Comentario'] = [r[0] for r in res_com]
    df_audit_tec['Justificacion_Comentario'] = [r[1] for r in res_com]

    res_obj = df_audit_tec['Obj_Text'].apply(calificar_y_justificar_objetivo)
    df_audit_tec['Calidad_Objetivo'] = [r[0] for r in res_obj]
    df_audit_tec['Justificacion_Objetivo'] = [r[1] for r in res_obj]
    
    def calificacion_global(row):
        c = row['Calidad_Comentario']
        o = row['Calidad_Objetivo']
        if 'Alerta' in c or 'Alerta' in o: return "Riesgo Metodológico (Alerta)"
        elif 'Regular' in c or 'Regular' in o: return "En Proceso de Apropiación"
        else: return "Visita Sobresaliente (Dominada)"

    df_audit_tec['Estado_Metodologico'] = df_audit_tec.apply(calificacion_global, axis=1)
    
    total_v_audit = len(df_audit_tec)
    sobresalientes = len(df_audit_tec[df_audit_tec['Estado_Metodologico'] == "Visita Sobresaliente (Dominada)"])
    alertas = len(df_audit_tec[df_audit_tec['Estado_Metodologico'] == "Riesgo Metodológico (Alerta)"])
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Visitas Únicas Evaluadas", f"{total_v_audit:,}")
    m2.metric("Visitas Sobresalientes", f"{sobresalientes:,}", f"{(sobresalientes/total_v_audit*100):.1f}%")
    m3.metric("Visitas en Alerta (Riesgo)", f"{alertas:,}", f"{(alertas/total_v_audit*100):.1f}%", delta_color
