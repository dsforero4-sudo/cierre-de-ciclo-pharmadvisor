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
    m3.metric("Visitas en Alerta (Riesgo)", f"{alertas:,}", f"{(alertas/total_v_audit*100):.1f}%", delta_color="inverse")
    st.markdown("---")
    
    df_rep_metodo = df_audit_tec.groupby([col_rep, 'Estado_Metodologico'], as_index=False).agg(Total=(col_visita, 'count'))
    df_rep_metodo.columns = ['Representante', 'Estado_Metodologico', 'Total']
    df_rep_totales = df_rep_metodo.groupby('Representante', as_index=False).agg(Total_Rep=('Total', 'sum'))
    df_rep_metodo = pd.merge(df_rep_metodo, df_rep_totales, on='Representante')
    df_rep_metodo['Porcentaje'] = (df_rep_metodo['Total'] / df_rep_metodo['Total_Rep'] * 100).round(1)
    
    def formato_etiqueta(row):
        if row['Porcentaje'] >= 5.0: return f"{row['Total']} ({row['Porcentaje']}%)"
        return ""

    df_rep_metodo['Texto_Barra'] = df_rep_metodo.apply(formato_etiqueta, axis=1)
    
    fig_metodo = px.bar(
        df_rep_metodo, x='Total', y='Representante', color='Estado_Metodologico', barmode='stack', text='Texto_Barra', template='plotly_white',
        title="<b>Adopción de la Técnica de Ventas por Representante</b>",
        color_discrete_map={'Visita Sobresaliente (Dominada)': '#2ECC71', 'En Proceso de Apropiación': '#3182CE', 'Riesgo Metodológico (Alerta)': '#E53E3E'},
        orientation='h'
    )
    fig_metodo.update_traces(textposition='inside', insidetextanchor='middle', textfont_size=11)
    fig_metodo.update_layout(
        paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF', height=max(450, len(representantes)*25),
        xaxis_title="Cantidad de Visitas Únicas", yaxis_title="Representante",
        margin=dict(t=50, b=50, l=150, r=40)
    )
    st.plotly_chart(fig_metodo, use_container_width=True)
    
    with st.expander("🔍 Ver detalle completo de auditoría (por Visita Única) con filtros de calidad y representante"):
        col_f_niv, col_f_rep, col_f_btn = st.columns([1.5, 1.5, 1])
        niveles_disponibles = ['Todos'] + sorted(df_audit_tec['Estado_Metodologico'].unique().tolist())
        with col_f_niv:
            filtro_nivel_sel = st.selectbox("Filtrar por Nivel Metodológico", options=niveles_disponibles, key="select_filtro_nivel")
            
        reps_disponibles_audit = ['Todos'] + sorted(df_audit_tec[col_rep].dropna().unique().tolist())
        with col_f_rep:
            filtro_rep_sel = st.selectbox("Filtrar por Representante", options=reps_disponibles_audit, key="select_filtro_rep_audit")
            
        df_tabla_filtrada = df_audit_tec.copy()
        if filtro_nivel_sel != 'Todos': df_tabla_filtrada = df_tabla_filtrada[df_tabla_filtrada['Estado_Metodologico'] == filtro_nivel_sel]
        if filtro_rep_sel != 'Todos': df_tabla_filtrada = df_tabla_filtrada[df_tabla_filtrada[col_rep] == filtro_rep_sel]
            
        cols_tabla = [c for c in [col_rep, col_visita, col_fecha, col_med, col_com, 'Calidad_Comentario', 'Justificacion_Comentario', col_obj, 'Calidad_Objetivo', 'Justificacion_Objetivo', 'Estado_Metodologico'] if c is not None]
        
        with col_f_btn:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            def convertir_df_a_excel(df_export):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Detalle_Filtrado')
                return output.getvalue()

            excel_filtrado = convertir_df_a_excel(df_tabla_filtrada[cols_tabla])
            st.download_button(
                label="📥 Descargar Tabla en Excel",
                data=excel_filtrado,
                file_name=f"Detalle_Visitas_Essity_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_excel_inline",
                use_container_width=True
            )

        st.dataframe(df_tabla_filtrada[cols_tabla], use_container_width=True, hide_index=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ==========================================
    # 5. FOCALIZACIÓN DE ACOMPAÑAMIENTOS POR COORDINACIÓN
    # ==========================================
    st.subheader("5. Focalización de Acompañamientos por Coordinación")
    st.markdown("<span style='color: #718096;'>Conteo de visitas con acompañamiento real de coordinador registradas en el ciclo.</span>", unsafe_allow_html=True)
    st.markdown("---")

    if col_acomp and col_acomp in df_unique.columns:
        df_acompañadas = df_unique[df_unique[col_acomp].notna() & (df_unique[col_acomp].astype(str).str.strip() != '')].copy()
        total_acomp = len(df_acompañadas)

        if total_acomp > 0:
            df_coord_acomp = df_acompañadas.groupby(col_reg).agg(Total_Acompañadas=(col_visita, 'count')).reset_index()
            df_coord_acomp.columns = ['Región', 'Total_Acompañadas']
            
            fig_acomp = px.bar(df_coord_acomp, x='Región', y='Total_Acompañadas', text='Total_Acompañadas', template='plotly_white', title="Visitas con Acompañamiento Real por Coordinación", color='Total_Acompañadas', color_continuous_scale='Blues')
            fig_acomp.update_traces(texttemplate='%{text}', textposition='outside', textfont_size=12)
            fig_acomp.update_layout(paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF', height=420, xaxis_title="Coordinación / Región", yaxis_title="Cantidad de Visitas Acompañadas", margin=dict(t=50, b=50, l=40, r=20))
            st.plotly_chart(fig_acomp, use_container_width=True)
        else:
            st.info("No se registraron visitas con acompañamiento bajo los filtros seleccionados.")
    else:
        st.info("No se detectó la columna de 'Visita acompañada' en el archivo.")

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ==========================================
    # 6. AUDITORÍA DE IMPACTOS PROMOCIONALES (REALIZADOS VS ESPERADOS)
    # ==========================================
    st.subheader("6. Auditoría de Impactos Promocionales: Realizados vs. Esperados")
    st.markdown("<span style='color: #718096;'>Comparativo de impactos realizados frente a los esperados por producto (Top 10).</span>", unsafe_allow_html=True)
    st.markdown("---")

    if df_impactos is not None:
        df_imp_filtrado = df_impactos[df_impactos['Región'].isin(selected_regiones)].copy() if 'Región' in df_impactos.columns else df_impactos
        df_imp_agrupado = df_imp_filtrado.groupby('Producto', as_index=False)[['Imp_Esp', 'Imp_Real']].sum()

        if len(df_imp_agrupado) > 0:
            df_imp_agrupado['Imp_Esp'] = df_imp_agrupado['Imp_Esp'].round(0).astype(int)
            df_imp_agrupado['Imp_Real'] = df_imp_agrupado['Imp_Real'].round(0).astype(int)
            df_imp_agrupado['Pct_Cumplimiento'] = (df_imp_agrupado['Imp_Real'] / df_imp_agrupado['Imp_Esp'] * 100).fillna(0).round(1)
            df_imp_agrupado['Imp_Pendiente'] = (df_imp_agrupado['Imp_Esp'] - df_imp_agrupado['Imp_Real']).apply(lambda x: max(0, x))
            df_imp_agrupado = df_imp_agrupado.sort_values(by=['Imp_Esp', 'Pct_Cumplimiento'], ascending=[True, True]).tail(10)
            df_imp_agrupado['Texto_Real'] = df_imp_agrupado['Imp_Real'].astype(str) + " (" + df_imp_agrupado['Pct_Cumplimiento'].astype(str) + "%)"
            df_imp_agrupado['Texto_Total'] = df_imp_agrupado['Imp_Esp'].astype(str)

            fig_impactos = go.Figure()
            fig_impactos.add_trace(go.Bar(y=df_imp_agrupado['Producto'], x=df_imp_agrupado['Imp_Real'], name='Realizados', orientation='h', marker=dict(color='#2ECC71'), text=df_imp_agrupado['Texto_Real'], textposition='inside'))
            fig_impactos.add_trace(go.Bar(y=df_imp_agrupado['Producto'], x=df_imp_agrupado['Imp_Pendiente'], name='Brecha Pendiente', orientation='h', marker=dict(color='#3182CE'), text=df_imp_agrupado['Texto_Total'], textposition='outside'))
            fig_impactos.update_layout(template='plotly_white', title="Top Impactos Promocionales: Realizados vs Esperados", paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF', height=450, barmode='stack', xaxis_title="Impactos Esperados", yaxis_title="Producto")
            st.plotly_chart(fig_impactos, use_container_width=True)
        else:
            st.info("No hay registros de impactos bajo los filtros seleccionados.")
    else:
        st.info("Cargue el archivo de Impactos Promocionales en la barra lateral para visualizar esta sección.")

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ==========================================
    # 7. AUDITORÍA AVANZADA DE MUESTRAS MÉDICAS Y RECEPTIVIDAD
    # ==========================================
    st.subheader("7. Auditoría Avanzada de Muestras Médicas y Receptividad por Producto")
    st.markdown("<span style='color: #718096;'>Análisis de unidades entregadas y receptividad o actitud médica registrada.</span>", unsafe_allow_html=True)
    st.markdown("---")

    if col_muestra and col_cant_muestra and col_muestra in df_filtered.columns:
        df_muestras = df_filtered.dropna(subset=[col_muestra]).copy()
        df_muestras['Muestra_Clean'] = df_muestras[col_muestra].astype(str).str.strip()
        excluidos_muestras = ['', '-', 'nan', 'none', 'nat', '0', 'false', 'no', 'sin muestra', 'ninguna']
        df_muestras = df_muestras[~df_muestras['Muestra_Clean'].str.lower().isin(excluidos_muestras)]
        df_muestras['Cantidad_Num'] = pd.to_numeric(df_muestras[col_cant_muestra], errors='coerce').fillna(1)

        df_m_agrup = df_muestras.groupby('Muestra_Clean', as_index=False).agg(Total_Unidades=('Cantidad_Num', 'sum')).sort_values(by='Total_Unidades', ascending=False).head(10)

        if len(df_m_agrup) > 0:
            fig_muestras = px.bar(df_m_agrup, x='Total_Unidades', y='Muestra_Clean', text='Total_Unidades', template='plotly_white', title="Top Muestras Médicas / Materiales Entregados", color='Total_Unidades', color_continuous_scale='Purples', orientation='h')
            fig_muestras.update_traces(texttemplate='%{text}', textposition='outside', textfont_size=11)
            fig_muestras.update_layout(paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF', height=420, xaxis_title="Total Unidades Entregadas", yaxis_title="Material / Muestra", margin=dict(t=50, b=50, l=150, r=20))
            st.plotly_chart(fig_muestras, use_container_width=True)
        else:
            st.info("No hay registros de muestras médicas válidas.")
    else:
        st.info("Columnas de muestras médicas no identificadas automáticamente.")

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
    st.subheader("Generación de Reporte Ejecutivo Gerencial Essity")
    st.markdown("<span style='color: #718096;'>Genera un informe analítico completo incrustando las visualizaciones comerciales para directivos.</span>", unsafe_allow_html=True)
    
    if st.button("Generar Informe Ejecutivo con Gráficas"):
        html_chart_copia = fig_bar_copia.to_html(full_html=False, include_plotlyjs='cdn')
        html_chart_metodo = fig_metodo.to_html(full_html=False, include_plotlyjs='cdn')
        
        pct_alertas = (alertas / total_v_audit * 100) if total_v_audit > 0 else 0
        
        rep_resumen = df_audit_tec.groupby(col_rep).agg(Visitas=(col_visita, 'count'), Alertas=('Estado_Metodologico', lambda x: (x == 'Riesgo Metodológico (Alerta)').sum()), Copia=('Es_Repetido', 'sum')).reset_index()
        rep_resumen.columns = ['Representante', 'Visitas', 'Alertas', 'Copia']
        rep_resumen['Pct_Riesgo'] = (rep_resumen['Alertas'] / rep_resumen['Visitas'] * 100).round(1)
        criticos = rep_resumen.sort_values(by='Pct_Riesgo', ascending=False).head(5)
        
        tabla_html_rows = ""
        for _, r in criticos.iterrows():
            tabla_html_rows += "<tr><td>" + str(r['Representante']) + "</td><td>" + str(r['Visitas']) + "</td><td>" + str(r['Alertas']) + "</td><td>" + str(r['Pct_Riesgo']) + "%</td></tr>"
        
        css_estilos = """
        <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; color: #2C3E50; margin: 40px; line-height: 1.6; background-color: #F8F9FA; }
        h1 { color: #E6007E; border-bottom: 3px solid #E6007E; padding-bottom: 8px; font-size: 24px; }
        h2 { color: #2D3346; margin-top: 40px; border-bottom: 1px solid #BDC3C7; padding-bottom: 5px; font-size: 18px; }
        .metrics-container { display: flex; justify-content: space-between; margin-bottom: 25px; }
        .metric-card { background: #FFFFFF; border-left: 4px solid #E6007E; padding: 15px; width: 22%; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .metric-title { font-size: 11px; color: #7F8C8D; text-transform: uppercase; font-weight: bold; }
        .metric-value { font-size: 18px; color: #2C3E50; font-weight: bold; margin-top: 5px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 25px; background: #FFFFFF; }
        th, td { border: 1px solid #DDDDDD; padding: 10px; text-align: left; font-size: 12px; }
        th { background-color: #2D3346; color: white; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .chart-container { margin: 30px 0; background: #FFFFFF; padding: 20px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .recommendation-box { background: #FFF5F9; border-left: 4px solid #E6007E; padding: 20px; border-radius: 4px; margin-top: 30px; }
        </style>
        """

        card_1 = "<div class='metric-card'><div class='metric-title'>Visitas Únicas</div><div class='metric-value'>" + f"{total_v_audit:,}" + "</div></div>"
        card_2 = "<div class='metric-card'><div class='metric-title'>Índice de Clonación</div><div class='metric-value'>" + f"{pct_copia:.1f}%" + "</div></div>"
        card_3 = "<div class='metric-card'><div class='metric-title'>Visitas Sobresalientes</div><div class='metric-value'>" + f"{sobresalientes:,} ({((sobresalientes/total_v_audit*100) if total_v_audit > 0 else 0):.1f}%)" + "</div></div>"
        card_4 = "<div class='metric-card'><div class='metric-title'>Riesgo / Alertas</div><div class='metric-value'>" + f"{alertas:,} ({pct_alertas:.1f}%)" + "</div></div>"

        reporte_html = (
            "<!DOCTYPE html><html><head><meta charset='utf-8'>" + css_estilos + "</head><body>"
            "<h1>ESSITY / LEUKOACADEMY | INFORME EJECUTIVO DE VALIDACIÓN METODOLÓGICA</h1>"
            "<p><b>Cliente / Cuenta:</b> " + str(cliente_input) + " | <b>Líneas / Segmento:</b> " + str(lineas_input) + "</p>"
            "<p><b>Fecha de Emisión:</b> " + pd.Timestamp.now().strftime('%Y-%m-%d %H:%M') + " | <b>Segmento:</b> Programa Leukoacademy</p>"
            "<h2>1. Resumen Ejecutivo del Ciclo</h2>"
            "<div class='metrics-container'>" + card_1 + card_2 + card_3 + card_4 + "</div>"
            "<h2>2. Hallazgos Analíticos y Visuales del Ciclo</h2>"
            "<div class='chart-container'><b>Gráfica de Autorrepetición por Representante</b><br>" + html_chart_copia + "</div>"
            "<div class='chart-container'><b>Gráfica de Calidad Metodológica</b><br>" + html_chart_metodo + "</div>"
            "<h2>3. Representantes con Mayor Oportunidad de Acompañamiento (Top Riesgos)</h2>"
            "<table><tr><th>Representante</th><th>Visitas Evaluadas</th><th>Registros en Alerta</th><th>% de Riesgo Metodológico</th></tr>" + tabla_html_rows + "</table>"
            "<div class='recommendation-box'>"
            "<h3 style='margin-top:0; color: #E6007E;'>4. Recomendaciones de Acción</h3>"
            "<ol>"
            "<li><b>Retroalimentación 1 a 1:</b> Programar sesiones de coaching con los asesores identificados con mayores índices de clonación en el programa Leukoacademy.</li>"
            "<li><b>Alineación en Objetivos:</b> Reforzar que los objetivos redactados reflejen comportamientos de valor y no tareas rutinarias.</li>"
            "</ol></div></body></html>"
        )
        
        st.success("¡Informe ejecutivo Essity con gráficas generado exitosamente!")
        st.download_button(
            label="📥 Descargar Informe Ejecutivo Completo (HTML / Imprimible a PDF)",
            data=reporte_html,
            file_name=f"Informe_Gerencial_Essity_{pd.Timestamp.now().strftime('%Y%m%d')}.html",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("Impresión a PDF: Abre el archivo descargado en tu navegador web, presiona Ctrl + P (o Cmd + P en Mac) y selecciona 'Guardar como PDF'.")

else:
    st.info("Por favor carga el archivo de visitas en la barra lateral para iniciar el análisis del programa Leukoacademy.")
