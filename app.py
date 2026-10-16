import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io

st.set_page_config(
    page_title="Pharmadvisor | Validador Metodológico",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
        .stApp { background: #0E1118 !important; color: #FFFFFF !important; }
        .login-card { 
            background: #1B2440; backdrop-filter: blur(12px); 
            padding: 40px; border-radius: 12px; 
            border: 1px solid rgba(61, 214, 255, 0.3); 
            max-width: 450px; margin: 80px auto; text-align: center; 
            box-shadow: 0 10px 30px rgba(0,0,0,0.8); 
        }
        .login-title { color: #FF2D8F; font-size: 26px; font-weight: bold; margin-bottom: 10px; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("""
        <div class="login-card">
            <div class="login-title">Pharm<span style="color: #3DD6FF;">Advisor</span></div>
            <p style="color: #B7BCC9; font-size: 14px;">Acceso Restringido - Validador Metodológico Multicliente</p>
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
# APLICACIÓN PRINCIPAL (ESTILO OFICIAL ADVISING)
# ==========================================
st.markdown("""
    <style>
    html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"], .main, .block-container {
        background-color: #0E1118 !important;
        color: #FFFFFF !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    section[data-testid="stSidebar"] {
        border-right: 3px solid #FF2D8F !important;
        background-color: #0E1118 !important;
    }
    section[data-testid="stSidebar"] .stSubheader, 
    section[data-testid="stSidebar"] label, 
    section[data-testid="stSidebar"] span {
        color: #FFFFFF !important;
    }
    
    .ph-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 24px 30px;
        border-bottom: 2px solid #FF2D8F;
        margin-bottom: 30px;
        background: linear-gradient(90deg, rgba(10, 88, 255, 0.2) 0%, rgba(122, 44, 255, 0.2) 50%, rgba(255, 45, 143, 0.2) 100%);
        border-radius: 10px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.6);
    }
    .ph-title {
        color: #FF2D8F;
        font-size: 26px;
        font-weight: bold;
        margin: 0;
        text-shadow: 0 2px 4px rgba(0,0,0,0.5);
    }
    
    h2 {
        color: #3DD6FF !important;
        font-size: 20px !important;
        font-weight: 700 !important;
        margin-top: 25px !important;
    }
    p, span, label {
        color: #B7BCC9 !important;
    }
    
    .section-divider {
        margin-top: 40px;
        margin-bottom: 40px;
        border: 0;
        height: 1px;
        background: rgba(61, 214, 255, 0.3);
    }
    
    div[data-testid="stMetric"] {
        background-color: #1B2440 !important;
        border-left: 4px solid #FF2D8F !important;
        border-top: 1px solid rgba(61, 214, 255, 0.2) !important;
        border-right: 1px solid rgba(61, 214, 255, 0.2) !important;
        border-bottom: 1px solid rgba(61, 214, 255, 0.2) !important;
        padding: 18px !important;
        border-radius: 10px !important;
        box-shadow: 0 8px 20px rgba(0,0,0,0.5) !important;
    }
    div[data-testid="stMetric"] label {
        color: #B7BCC9 !important;
        font-size: 13px !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #3DD6FF !important;
        font-size: 28px !important;
        font-weight: bold !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="ph-header">
        <div>
            <h1 class="ph-title">Informe de calidad de la visita - Pharmadvisor</h1>
            <span style="color: #B7BCC9; font-size: 13px;">Auditoría Integral de Registro, Calidad y Muestras | Essity & Pharmadvisor</span>
        </div>
        <div style="text-align: right;">
            <span style="color: #FF2D8F; font-weight: bold; font-size: 20px;">Pharm<span style="color: #3DD6FF;">Advisor</span></span>
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
    if file is None: 
        return None
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
            if pd.isna(col0) or pd.isna(col1): 
                continue
            col0_str = str(col0).strip()
            
            if col0_str.lower().startswith("region") or col0_str.lower().startswith("región"):
                current_region = col0_str.replace("Región", "").replace("Region", "").strip()
                continue
            if col0_str.startswith("DISTRITO"): 
                continue
            
            rep_text = col0_str
            producto = str(col1).strip()
            coordinacion = "COORDINACION CL" if "CL" in current_region.upper() or "CL" in rep_text.upper() or "ANTIOQUIA" in rep_text.upper() or "COSTA" in rep_text.upper() or "SANTANDER" in rep_text.upper() else "COORDINACION PS"
            
            for c_idx in range(2, len(row10), 3):
                if c_idx >= df_raw.shape[1]: 
                    break
                spec = specialties[c_idx]
                imp_esp = df_raw.iloc[idx, c_idx]
                imp_real = df_raw.iloc[idx, c_idx + 1]
                
                if pd.notna(imp_esp) and str(imp_esp).strip() != '-' and str(imp_esp).strip() != '':
                    try:
                        esp_val = float(imp_esp)
                        real_val = float(imp_real) if pd.notna(imp_real) and str(imp_real).strip() != '-' else 0.0
                        records.append({
                            'Región': coordinacion, 
                            'Representante_Raw': rep_text, 
                            'Producto': producto, 
                            'Especialidad': spec, 
                            'Imp_Esp': esp_val, 
                            'Imp_Real': real_val
                        })
                    except: 
                        pass
        return pd.DataFrame(records)
    except Exception as e: 
        return None

df_visitas = None
default_path = 'listado_visitas_2026-09-07_11-44-08.xlsx'
source_file = uploaded_file if uploaded_file is not None else (default_path if os.path.exists(default_path) else None)

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

    if not col_visita: 
        col_visita = df_visitas.columns[0]
    if not col_rep: 
        col_rep = df_visitas.columns[0]
    if not col_reg: 
        col_reg = df_visitas.
