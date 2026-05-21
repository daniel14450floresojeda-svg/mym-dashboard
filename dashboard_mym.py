import streamlit as st
import pandas as pd
import os


# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(
    page_title="STARK BI - Muñoz Marchesi",
    layout="wide",
    initial_sidebar_state="expanded"
)


ARCHIVO_HISTORIAL = "historial_bi.xlsx"


# 2. ESTILOS VISUALES GENERALES — PREMIUM DARK
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=Exo+2:wght@300;400;500;600;700&display=swap');


    :root {
        --bg-deep:      #050B18;
        --bg-panel:     #0A1628;
        --bg-card:      #0D1F3C;
        --border:       #1A3558;
        --accent-blue:  #00B4FF;
        --text-bright:  #E8F4FF;
        --text-mid:     #7FA8CC;
    }


    .stApp {
        background-color: var(--bg-deep) !important;
        background-image: linear-gradient(180deg, #050B18 0%, #07111F 100%);
        font-family: 'Exo 2', sans-serif !important;
        color: var(--text-bright) !important;
    }


    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #060E1D 0%, #080F1A 100%) !important;
        border-right: 1px solid var(--border) !important;
    }


    .sidebar-logo {
        background: linear-gradient(135deg, #0A1628 0%, #0D2044 100%);
        border-bottom: 1px solid var(--border);
        padding: 18px 20px;
        margin-bottom: 8px;
    }
    .sidebar-logo h1 {
        font-family: 'Rajdhani', sans-serif !important;
        font-size: 1.8rem !important;
        font-weight: 700 !important;
        color: var(--accent-blue) !important;
        letter-spacing: 3px !important;
        margin: 0 !important;
    }
    .sidebar-logo .subtitle {
        font-size: 0.62rem;
        color: var(--text-mid);
        text-transform: uppercase;
    }


    .sb-section {
        font-family: 'Rajdhani', sans-serif;
        font-size: 0.68rem;
        font-weight: 600;
        color: #3A5A7A;
        letter-spacing: 2px;
        text-transform: uppercase;
        padding: 14px 20px 6px 20px;
        border-top: 1px solid #0F2030;
        margin-top: 10px;
    }


    .main-header {
        padding: 28px 0 6px 0;
        border-bottom: 1px solid var(--border);
        margin-bottom: 24px;
    }
    .main-header h1 {
        font-family: 'Rajdhani', sans-serif !important;
        font-size: 2rem !important;
        font-weight: 700 !important;
    }


    .kpi-wrap {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 1px;
        background: var(--border);
        border: 1px solid var(--border);
        border-radius: 10px;
        overflow: hidden;
        margin-bottom: 28px;
    }
    .kpi-card { background: var(--bg-card); padding: 18px 22px; }
    .kpi-label { font-size: 0.68rem; color: var(--text-mid); text-transform: uppercase; }
    .kpi-value { font-family: 'Rajdhani', sans-serif; font-size: 1.7rem; font-weight: 700; color: var(--accent-blue); }


    .section-title { font-family: 'Rajdhani', sans-serif; font-size: 1.1rem; font-weight: 700; text-transform: uppercase; margin: 28px 0 16px 0; }
    .stButton > button { background: linear-gradient(135deg, #0D2044 0%, #0A3060 100%) !important; border: 1px solid var(--accent-blue) !important; color: var(--accent-blue) !important; font-family: 'Rajdhani', sans-serif !important; text-transform: uppercase !important; width: 100% !important; }
    .footer { text-align: center; font-size: 0.68rem; color: #3A5A7A; padding: 30px 0 10px 0; border-top: 1px solid #0F1E30; margin-top: 40px; text-transform: uppercase; }
</style>
""", unsafe_allow_html=True)


# ── SIDEBAR INTERFAZ — CONTROL DE PALANCAS EN VIVO ──
with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
        <h1>STARK BI</h1>
        <div class="subtitle">Módulo de Postventa</div>
    </div>
    """, unsafe_allow_html=True)
   
    st.markdown('<div class="sb-section">Sección de Carga de Datos</div>', unsafe_allow_html=True)
    archivos_cargados = st.file_uploader("Subí tu reporte de ventas xlsx (.xlsx)", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed")


    st.markdown('<div class="sb-section">1. Margen Comercial Global</div>', unsafe_allow_html=True)
    margen_repuestos = st.number_input("Margen Repuestos (%)", value=30, step=5)
    coeficiente_costo = 1 / (1 + (margen_repuestos / 100))


    st.markdown('<div class="sb-section">2. Dotación Activa (Mecánicos)</div>', unsafe_allow_html=True)
   
    # Podés modificar los valores acá en vivo y todo el tablero se reajusta sin romperse
    with st.expander("MYM - Kawasaki (Alta Gama)"):
        cf_kawasaki = st.number_input("Fijo KAW", value=4800000, step=100000, key="cf_kaw")
        cv_kawasaki = st.number_input("Variable KAW", value=250000, step=50000, key="cv_kaw")
        emp_kawasaki = st.number_input("Mecánicos KAW", value=3, step=1, key="emp_kaw") # Modificable en vivo
       
    with st.expander("MYM - Triumph (Alta Gama)"):
        cf_triumph = st.number_input("Fijo TRI", value=2800000, step=100000, key="cf_tri")
        cv_triumph = st.number_input("Variable TRI", value=600000, step=50000, key="cv_tri")
        emp_triumph = st.number_input("Mecánicos TRI", value=1, step=1, key="emp_tri")
       
    with st.expander("MYM - Corrientes (Estándar)"):
        cf_corrientes = st.number_input("Fijo COR", value=2100000, step=100000, key="cf_cor")
        cv_corrientes = st.number_input("Variable COR", value=400000, step=50000, key="cv_cor")
        emp_corrientes = st.number_input("Mecánicos COR", value=5, step=1, key="emp_cor")


    with st.expander("MYM - Resistencia (Estándar)"):
        cf_resistencia = st.number_input("Fijo RES", value=3200000, step=100000, key="cf_res")
        cv_resistencia = st.number_input("Variable RES", value=600000, step=50000, key="cv_res")
        emp_resistencia = st.number_input("Mecánicos RES", value=6, step=1, key="emp_res")


    cf_brenas, cv_brenas, emp_brenas = 1200000, 200000, 2
    cf_charata, cv_charata, emp_charata = 1100000, 150000, 2
    cf_angela, cv_angela, emp_angela = 1300000, 250000, 2


    st.markdown('<div class="sb-section">Período Comercial</div>', unsafe_allow_html=True)
    mes_sel = st.selectbox("Mes", ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"], index=4, label_visibility="collapsed")
    anio_sel = st.selectbox("Año", [2025,2026,2027,2028], index=1, label_visibility="collapsed")
    periodo_str = f"{mes_sel} {anio_sel}"


    st.markdown("<br>", unsafe_allow_html=True)
    ejecutar_guardado = st.button("💾  Registrar Período Actual")


# ── PROCESAMIENTO MATEMÁTICO CONTABLE REAL ──
sucursales_lista = ["MYM - KAWASAKI","MYM - TRIUMPH","MYM - CORRIENTES","MYM - RESISTENCIA","MYM - LAS BREÑAS","MYM - CHARATA","MYM - VILLA ANGELA"]
palabras_clave = ["KAW", "TRIUMPH", "CORRIENTES", "RESISTENCIA", "BREÑAS", "CHARATA", "ANGELA"]


motos_mes  = [0] * 7
ingresos_totales = [0.0] * 7
costos_mercaderia = [0.0] * 7
costos_fijos     = [cf_kawasaki, cf_triumph, cf_corrientes, cf_resistencia, cf_brenas, cf_charata, cf_angela]
costos_variables = [cv_kawasaki, cv_triumph, cv_corrientes, cv_resistencia, cv_brenas, cv_charata, cv_angela]
dotacion         = [emp_kawasaki, emp_triumph, emp_corrientes, emp_resistencia, emp_brenas, emp_charata, emp_angela]


if archivos_cargados:
    try:
        for archivo in archivos_cargados:
            df_excel = pd.read_excel(archivo, engine='openpyxl')