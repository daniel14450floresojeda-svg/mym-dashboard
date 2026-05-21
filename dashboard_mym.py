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

# 2. ESTILOS VISUALES GENERALES
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=Exo+2:wght@300;400;500;600;700&display=swap');
    :root { --bg-deep: #050B18; --bg-panel: #0A1628; --bg-card: #0D1F3C; --border: #1A3558; --accent-blue: #00B4FF; --text-bright: #E8F4FF; --text-mid: #7FA8CC; }
    .stApp { background-color: var(--bg-deep) !important; background-image: linear-gradient(180deg, #050B18 0%, #07111F 100%); font-family: 'Exo 2', sans-serif !important; color: var(--text-bright) !important; }
    [data-testid="stSidebar"] { background: linear-gradient(180deg, #060E1D 0%, #080F1A 100%) !important; border-right: 1px solid var(--border) !important; }
    .sidebar-logo { background: linear-gradient(135deg, #0A1628 0%, #0D2044 100%); border-bottom: 1px solid var(--border); padding: 18px 20px; margin-bottom: 8px; }
    .sidebar-logo h1 { font-family: 'Rajdhani', sans-serif !important; font-size: 1.8rem !important; font-weight: 700 !important; color: var(--accent-blue) !important; letter-spacing: 3px !important; margin: 0 !important; }
    .sidebar-logo .subtitle { font-size: 0.62rem; color: var(--text-mid); text-transform: uppercase; }
    .sb-section { font-family: 'Rajdhani', sans-serif; font-size: 0.68rem; font-weight: 600; color: #3A5A7A; letter-spacing: 2px; text-transform: uppercase; padding: 14px 20px 6px 20px; border-top: 1px solid #0F2030; margin-top: 10px; }
    .main-header { padding: 28px 0 6px 0; border-bottom: 1px solid var(--border); margin-bottom: 24px; }
    .kpi-wrap { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1px; background: var(--border); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; margin-bottom: 28px; }
    .kpi-card { background: var(--bg-card); padding: 18px 22px; }
    .kpi-label { font-size: 0.68rem; color: var(--text-mid); text-transform: uppercase; }
    .kpi-value { font-family: 'Rajdhani', sans-serif; font-size: 1.7rem; font-weight: 700; color: var(--accent-blue); }
    .section-title { font-family: 'Rajdhani', sans-serif; font-size: 1.1rem; font-weight: 700; text-transform: uppercase; margin: 28px 0 16px 0; }
    .footer { text-align: center; font-size: 0.68rem; color: #3A5A7A; padding: 30px 0 10px 0; border-top: 1px solid #0F1E30; margin-top: 40px; text-transform: uppercase; }
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR ──
with st.sidebar:
    st.markdown('<div class="sidebar-logo"><h1>STARK BI</h1><div class="subtitle">Módulo de Postventa</div></div>', unsafe_allow_html=True)
    archivos_cargados = st.file_uploader("Subí reporte de ventas", type=["xlsx"], accept_multiple_files=True)
    margen_repuestos = st.number_input("Margen Repuestos (%)", value=30, step=5)
    coeficiente_costo = 1 / (1 + (margen_repuestos / 100))
    
    st.markdown('<div class="sb-section">Dotación Activa (Mecánicos)</div>', unsafe_allow_html=True)
    cf_kawasaki = st.number_input("Fijo KAW", value=4800000, step=100000)
    cv_kawasaki = st.number_input("Variable KAW", value=250000, step=50000)
    emp_kawasaki = st.number_input("Mecánicos KAW", value=3, step=1)
    
    # Valores por defecto para otras sucursales (puedes ajustar)
    cf_triumph, cv_triumph, emp_triumph = 2800000, 600000, 1
    cf_otros, cv_otros, emp_otros = 1200000, 200000, 2
    
    ejecutar_guardado = st.button("💾 Registrar Período")

# ── PROCESAMIENTO INTELIGENTE ──
if archivos_cargados:
    lista_dfs = []
    for archivo in archivos_cargados:
        df_temp = pd.read_excel(archivo, engine='openpyxl')
        df_temp.columns = df_temp.columns.astype(str).str.strip().str.lower()
        lista_dfs.append(df_temp)
    
    df = pd.concat(lista_dfs, ignore_index=True)
    
    # LIMPIEZA AUTOMÁTICA DE SUCURSALES
    col_suc = next((c for c in df.columns if 'sucursal' in c), None)
    if col_suc:
        df[col_suc] = df[col_suc].astype(str).str.strip().str.upper()
        # Generar lista única automáticamente
        sucursales_detectadas = df[col_suc].unique().tolist()
        
        st.subheader("Selección de Unidad")
        sucursal_sel = st.selectbox("Unidad de Negocio:", sucursales_detectadas)
        df_filtrado = df[df[col_suc] == sucursal_sel]
        
        st.dataframe(df_filtrado)
    else:
        st.error("No se encontró la columna 'Sucursal' en el archivo.")
else:
    st.info("Subí un archivo para comenzar.")

st.markdown('<div class="footer">Patented · Stark-BI © 2026 — Módulo de Control Muñoz Marchesi</div>', unsafe_allow_html=True)