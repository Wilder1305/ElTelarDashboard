# main.py — Dashboard ElTelar 2024-2025
# Uso columna 'Trimestre' directamente
# Mejoras estéticas y organizativas con código completo e integral

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import io
import os
import unicodedata
import re
from pathlib import Path

# -------------------------
# Configuración de la app y estilo
# -------------------------
st.set_page_config(page_title="Dashboard ElTelar 2024-2025", layout="wide", initial_sidebar_state="expanded")

# Colores personalizados para KPIs
COLOR_POS = "#4CAF50"
COLOR_NEG = "#F44336"

# Aplicar estilo general
st.markdown(
    """
    <style>
    .big-font {
        font-size:22px !important;
        font-weight: bold !important;
    }
    .metric-label {
        color: #333333;
    }
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------
# Helpers
# -------------------------
def strip_accents(s: str) -> str:
    if not isinstance(s, str): return s
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def normalize_colname(s: str) -> str:
    if not isinstance(s, str): return s
    s2 = strip_accents(s)
    s2 = s2.strip().lower()
    s2 = s2.replace('%', 'pct').replace('/', '_')
    s2 = re.sub(r'[^0-9a-zA-Z_ ]', '_', s2)
    s2 = s2.replace(' ', '_')
    while '__' in s2: s2 = s2.replace('__', '_')
    s2 = s2.strip('_')
    return s2

def normalize_columns_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_colname(c) for c in df.columns]
    return df

def normalize_text_series(s: pd.Series) -> pd.Series:
    return s.fillna('').astype(str).apply(strip_accents).str.strip().str.upper()

def find_column(df: pd.DataFrame, candidates: list):
    if df is None or df.empty: return None
    cols = df.columns.tolist()
    for cand in candidates:
        cand_norm = normalize_colname(cand)
        for col in cols:
            if col == cand_norm:
                return col
        for col in cols:
            if cand_norm in col:
                return col
    for col in cols:
        for cand in candidates:
            if cand.lower() in col:
                return col
    return None

# -------------------------
# Funciones de carga
# -------------------------
@st.cache_data(show_spinner=True)
def load_all_excel_sheets(path_or_url):
    try:
        xls = pd.ExcelFile(path_or_url)
    except Exception as e:
        st.warning(f"No se pudo abrir {path_or_url}: {e}")
        return pd.DataFrame()
    dfs = []
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet)
            if df is None or df.shape[0] == 0:
                continue
            df = df.copy()
            df['__source_file__'] = Path(path_or_url).name
            df['__source_sheet__'] = str(sheet)
            dfs.append(df)
        except Exception as e:
            st.warning(f"No se pudo leer hoja {sheet} de {path_or_url}: {e}")
    if not dfs:
        return pd.DataFrame()
    big = pd.concat(dfs, ignore_index=True, sort=False)
    if '__source_sheet__' not in big.columns:
        big['__source_sheet__'] = ''
    return big

@st.cache_data(show_spinner=True)
def cargar_todo_local_o_remoto():
    hist_file = "https://github.com/Wilder1305/ElTelarDashboard/raw/refs/heads/main/ElTelar_Matriculas_2024.xlsx    "
    prop_file = "https://github.com/Wilder1305/ElTelarDashboard/raw/refs/heads/main/Propuesta_Programacion_T2_2025_ElTelar.xlsx"
    port_file = "https://github.com/Wilder1305/ElTelarDashboard/raw/refs/heads/main/Portafolio%20Activo%202025-T2%20y%20Espacios.xlsx"
    df_hist = load_all_excel_sheets(hist_file)
    df_prop = load_all_excel_sheets(prop_file)
    df_port = load_all_excel_sheets(port_file)
    return df_hist, df_prop, df_port

# -------------------------
# Normalización y armonización usando columna 'trimestre'
# -------------------------
def armonizar_historico(df_raw):
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()
    df = normalize_columns_df(df_raw)
    if '__source_sheet__' not in df.columns:
        df['__source_sheet__'] = ''

    inscritos_col = find_column(df, ['total matriculas', 'total_matriculas', 'inscritos', 'matriculas', 'asistentes'])
    if inscritos_col:
        df['total_matriculas'] = pd.to_numeric(df[inscritos_col], errors='coerce').fillna(0).astype(int)
    else:
        df['total_matriculas'] = 0

    capacidad_col = find_column(df, ['cap mx e', 'cap_mx_e', 'capacidad', 'cupos', 'cupo'])
    if capacidad_col:
        df['cap_mx_e'] = pd.to_numeric(df[capacidad_col], errors='coerce')
    else:
        df['cap_mx_e'] = np.nan

    espacio_col = find_column(df, ['aula', 'salon', 'espacio', 'sede', 'sala', 'nombre_del_espacio'])
    if espacio_col:
        df['espacio'] = df[espacio_col].fillna('DESCONOCIDO').astype(str)
    else:
        df['espacio'] = 'DESCONOCIDO'

    producto_col = find_column(df, ['descripcion_evento', 'programa', 'tipo_evento', 'plan_de_estudios', 'producto'])
    if producto_col:
        df['producto'] = df[producto_col].astype(str)
    else:
        df['producto'] = 'DESCONOCIDO'

    fecha_inicio_col = find_column(df, ['fecha_inicio', 'fecha de inicio', 'start_date', 'date'])
    if fecha_inicio_col:
        df['fecha_inicio'] = pd.to_datetime(df[fecha_inicio_col], errors='coerce')
    else:
        df['fecha_inicio'] = pd.NaT

    trimestre_col = find_column(df, ['trimestre'])
    if trimestre_col:
        df['trimestre'] = pd.to_numeric(df[trimestre_col], errors='coerce')
    else:
        df['trimestre'] = pd.NA

    df['ocupacion'] = np.nan
    if 'cap_mx_e' in df.columns and 'total_matriculas' in df.columns:
        mask = (~df['cap_mx_e'].isna()) & (df['cap_mx_e'] != 0)
        if mask.any():
            df.loc[mask, 'ocupacion'] = (df.loc[mask, 'total_matriculas'] / df.loc[mask, 'cap_mx_e']) * 100

    df['espacio'] = normalize_text_series(df['espacio'])
    df['producto'] = normalize_text_series(df['producto'])
    df = df.drop_duplicates().reset_index(drop=True)
    df['fuente'] = 'historico_2024'
    return df

def armonizar_propuesta(df_raw):
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()
    df = normalize_columns_df(df_raw)
    if '__source_sheet__' not in df.columns:
        df['__source_sheet__'] = ''

    inscritos_col = find_column(df, [
        'expected_matriculas', 'expectedmatriculas', 'inscritos', 'total_matriculas', 'matriculas',
        'matricula_esperada', 'matriculas_propuesta', 'matriculados', 'matricula esperada'
    ])
    if inscritos_col:
        df['total_matriculas'] = pd.to_numeric(df[inscritos_col], errors='coerce').fillna(0).astype(int)
    else:
        df['total_matriculas'] = 0

    capacidad_col = find_column(df, ['space_capacity', 'spacecapacity', 'capacidad', 'cap_mx', 'cap_mx_e'])
    if capacidad_col:
        df['cap_mx_e'] = pd.to_numeric(df[capacidad_col], errors='coerce')
    else:
        df['cap_mx_e'] = np.nan

    espacio_col = find_column(df, ['assigned_space', 'space', 'espacio', 'aula', 'salon'])
    if espacio_col:
        df['espacio'] = df[espacio_col].fillna('DESCONOCIDO').astype(str)
    else:
        df['espacio'] = 'DESCONOCIDO'

    horario_col = find_column(df, ['horario', 'hora', 'time', 'franja', 'horario_raw'])
    if horario_col:
        df['horario_raw'] = df[horario_col].astype(str)
    else:
        dia_col = find_column(df, ['dia', 'day'])
        hora_inicio_col = find_column(df, ['hora_inicio', 'start_time'])
        if dia_col and hora_inicio_col:
            df['horario_raw'] = df[dia_col].astype(str) + ' ' + df[hora_inicio_col].astype(str)
        else:
            df['horario_raw'] = ''

    producto_col = find_column(df, ['plan_de_estudios', 'programa', 'tipo_evento', 'producto', 'descripcion_evento'])
    if producto_col:
        df['producto'] = df[producto_col].astype(str)
    else:
        df['producto'] = 'DESCONOCIDO'

    fecha_col = find_column(df, ['fecha_inicio', 'start_date', 'date'])
    if fecha_col:
        df['fecha_inicio'] = pd.to_datetime(df[fecha_col], errors='coerce')
    else:
        df['fecha_inicio'] = pd.NaT

    trimestre_col = find_column(df, ['trimestre'])
    if trimestre_col:
        df['trimestre'] = pd.to_numeric(df[trimestre_col], errors='coerce')
    else:
        df['trimestre'] = pd.NA

    df['ocupacion'] = np.nan
    if 'cap_mx_e' in df.columns and 'total_matriculas' in df.columns:
        mask = (~df['cap_mx_e'].isna()) & (df['cap_mx_e'] != 0)
        if mask.any():
            df.loc[mask, 'ocupacion'] = (df.loc[mask, 'total_matriculas'] / df.loc[mask, 'cap_mx_e']) * 100

    df['espacio'] = normalize_text_series(df['espacio'])
    df['producto'] = normalize_text_series(df['producto'])
    df = df.drop_duplicates().reset_index(drop=True)
    df['fuente'] = 'propuesta_2025'
    return df

def sanitize_periodos(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy()

    # Año
    if 'año' in df.columns:
        df['año'] = pd.to_numeric(df['año'], errors='coerce')
    else:
        df['año'] = pd.NA

    # Usar fecha_inicio para completar año
    if 'fecha_inicio' in df.columns:
        missing_year = df['año'].isna()
        if missing_year.any():
            years_from_date = pd.to_numeric(df.loc[missing_year, 'fecha_inicio'].dt.year, errors='coerce')
            df.loc[missing_year, 'año'] = years_from_date

    # Extraer año desde '__source_file__' si falta
    missing_year = df['año'].isna()
    if missing_year.any() and '__source_file__' in df.columns:
        extracted = pd.to_numeric(df.loc[missing_year, '__source_file__'].astype(str).str.extract(r'(\d{4})')[0], errors='coerce')
        df.loc[missing_year, 'año'] = extracted

    # Sanitizar trimestre
    if 'trimestre' in df.columns:
        df['trimestre'] = pd.to_numeric(df['trimestre'], errors='coerce')
    else:
        df['trimestre'] = pd.NA

    # Extraer desde 'periodo' si falta trimestre
    if 'periodo' in df.columns:
        missing_tr = df['trimestre'].isna()
        if missing_tr.any():
            extracted = pd.to_numeric(df.loc[missing_tr, 'periodo'].astype(str).str.extract(r'[Tt]([1-4])')[0], errors='coerce')
            df.loc[missing_tr, 'trimestre'] = extracted

    # Extraer desde '__source_sheet__' si falta trimestre
    if '__source_sheet__' in df.columns:
        missing_tr = df['trimestre'].isna()
        if missing_tr.any():
            extracted = pd.to_numeric(df.loc[missing_tr, '__source_sheet__'].astype(str).str.extract(r'[Tt]([1-4])')[0], errors='coerce')
            df.loc[missing_tr, 'trimestre'] = extracted

    # Valores fuera de rango a NaN
    df.loc[~df['trimestre'].isin([1, 2, 3, 4]), 'trimestre'] = pd.NA

    try:
        df['trimestre'] = df['trimestre'].astype('Int64')
    except Exception:
        pass
    try:
        df['año'] = df['año'].astype('Int64')
    except Exception:
        pass

    return df

# KPIs y visualizaciones

def kpi_totales(df):
    return int(df['total_matriculas'].sum()) if (df is not None and not df.empty and 'total_matriculas' in df.columns) else 0

def kpi_variacion(hist, prop):
    th = kpi_totales(hist)
    tp = kpi_totales(prop)
    diff = tp - th
    pct = (diff / th * 100) if th != 0 else np.nan
    return th, tp, diff, pct

def kpi_ocupacion_promedio(df):
    if df is None or df.empty or 'ocupacion' not in df.columns:
        return np.nan
    vals = df['ocupacion'].dropna()
    return float(vals.mean()) if not vals.empty else np.nan

def top_productos(df, topn=10):
    if df is None or df.empty or 'producto' not in df.columns:
        return pd.Series(dtype=int)
    s = df.groupby('producto')['total_matriculas'].sum().sort_values(ascending=False).head(topn)
    return s

def alertas_sobrecupo(df):
    if df is None or df.empty or 'cap_mx_e' not in df.columns:
        return pd.DataFrame()
    mask = (~df['cap_mx_e'].isna()) & (df['total_matriculas'] > df['cap_mx_e'])
    return df[mask].copy()

def grafica_serie_temporal(df_hist, df_prop):
    def agg(df):
        if df is None or df.empty: return pd.DataFrame(columns=['año','trimestre','total_matriculas'])
        if 'año' not in df.columns or 'trimestre' not in df.columns:
            return pd.DataFrame(columns=['año','trimestre','total_matriculas'])
        return df.groupby(['año','trimestre'])['total_matriculas'].sum().reset_index()
    h = agg(df_hist)
    p = agg(df_prop)
    if h.empty and p.empty:
        fig, ax = plt.subplots(figsize=(10,5))
        ax.text(0.5,0.5,"No hay datos para la serie temporal con los filtros actuales", ha='center', va='center')
        ax.axis('off')
        return fig
    fig, ax = plt.subplots(figsize=(10,5))
    if not h.empty:
        h = h.sort_values(['año','trimestre'])
        h['periodo'] = h['año'].astype(str) + "-T" + h['trimestre'].astype(str)
        ax.plot(h['periodo'], h['total_matriculas'], marker='o', label='Histórico')
    if not p.empty:
        p = p.sort_values(['año','trimestre'])
        p['periodo'] = p['año'].astype(str) + "-T" + p['trimestre'].astype(str)
        ax.plot(p['periodo'], p['total_matriculas'], marker='o', label='Propuesta')
    ax.set_title('Serie temporal: matrículas por trimestre')
    ax.set_ylabel('Total matrículas')
    ax.set_xlabel('Periodo')
    plt.xticks(rotation=45)
    ax.legend()
    plt.tight_layout()
    return fig

def grafica_barras_apiladas(df_hist, df_prop, topn=8):
    df_all = pd.concat([df_hist, df_prop], ignore_index=True, sort=False)
    if df_all.empty:
        fig, ax = plt.subplots(figsize=(10,5))
        ax.text(0.5,0.5,"No hay datos para barras apiladas", ha='center', va='center')
        ax.axis('off')
        return fig
    top = top_productos(df_all, topn).index.tolist()
    if not top:
        fig, ax = plt.subplots(figsize=(10,5))
        ax.text(0.5,0.5,"No hay productos para mostrar en barras apiladas", ha='center', va='center')
        ax.axis('off')
        return fig
    df_f = df_all[df_all['producto'].isin(top)]
    agg = df_f.groupby(['trimestre','producto','fuente'])['total_matriculas'].sum().reset_index()
    if agg.empty:
        fig, ax = plt.subplots(figsize=(10,5))
        ax.text(0.5,0.5,"No hay datos para barras apiladas", ha='center', va='center')
        ax.axis('off')
        return fig
    fig = px.bar(agg, x='trimestre', y='total_matriculas', color='producto', facet_col='fuente', barmode='stack',
                 title=f'Matrículas por Producto y Trimestre (Top {topn})')
    return fig

def grafica_heatmap_ocupacion(df_filtrado):
    if df_filtrado is None or df_filtrado.empty:
        fig, ax = plt.subplots(figsize=(10,5))
        ax.text(0.5,0.5,"No hay datos para heatmap con los filtros actuales", ha='center', va='center')
        ax.axis('off')
        return fig
    espacio_col = find_column(df_filtrado, ['aula', 'salon', 'espacio', 'sede', 'sala', 'nombre_del_espacio'])
    horario_col = find_column(df_filtrado, ['horario', 'hora_inicio', 'hora', 'franja', 'horario_raw'])
    ocup_col = find_column(df_filtrado, ['ocupacion', 'ocupacion_horario', 'occupancy_pct', 'ocupacion_media', 'ocupacion_calculada'])
    if ocup_col is None and ('total_matriculas' in df_filtrado.columns and 'cap_mx_e' in df_filtrado.columns):
        df_filtrado = df_filtrado.copy()
        mask = (~df_filtrado['cap_mx_e'].isna()) & (df_filtrado['cap_mx_e'] != 0)
        if mask.any():
            df_filtrado.loc[mask, 'ocupacion_calculada'] = (df_filtrado.loc[mask, 'total_matriculas'] / df_filtrado.loc[mask, 'cap_mx_e']) * 100
            ocup_col = 'ocupacion_calculada'
    if espacio_col is None or horario_col is None or ocup_col is None:
        fig, ax = plt.subplots(figsize=(10,5))
        msg_parts = []
        if espacio_col is None: msg_parts.append("columna de espacio (ej. Aula/Salón)")
        if horario_col is None: msg_parts.append("columna de horario (ej. Horario / Hora inicio)")
        if ocup_col is None: msg_parts.append("columna de ocupación o capacidad+matriculas")
        col_list = ", ".join(df_filtrado.columns[:50])
        ax.text(0.02, 0.5, "No se encontraron las columnas necesarias:\n- " + "\n- ".join(msg_parts) + "\n\nColumnas disponibles:\n" + col_list, va='center', ha='left', fontsize=10)
        ax.axis('off')
        return fig
    df = df_filtrado.copy()
    df['horario_label'] = df[horario_col].astype(str).fillna('')
    df['espacio_label'] = df[espacio_col].astype(str).fillna('DESCONOCIDO')
    try:
        heat_df = df.pivot_table(index='espacio_label', columns='horario_label', values=ocup_col, aggfunc='mean')
    except Exception as e:
        fig, ax = plt.subplots(figsize=(10,5))
        ax.text(0.5,0.5,f"Error al construir tabla para heatmap: {e}", ha='center', va='center')
        ax.axis('off')
        return fig
    if heat_df.empty:
        fig, ax = plt.subplots(figsize=(10,5))
        ax.text(0.5,0.5,"Heatmap: no hay combinaciones espacio-horario con datos", ha='center', va='center')
        ax.axis('off')
        return fig
    fig, ax = plt.subplots(figsize=(12, max(4, min(0.4*len(heat_df.index), 20))))
    sns.heatmap(heat_df, cmap='YlGnBu', linewidths=0.3, ax=ax, cbar_kws={'label':'% Ocupación'}, annot=False)
    ax.set_title("Heatmap de ocupación por espacio y horario")
    ax.set_xlabel("Horario")
    ax.set_ylabel("Espacio")
    plt.tight_layout()
    return fig

def grafica_barras_horizontales_top_productos(df_filtrado, topn=10):
    top = top_productos(df_filtrado, topn)
    if top.empty:
        fig, ax = plt.subplots(figsize=(8,4))
        ax.text(0.5,0.5,"No hay datos para Top productos", ha='center', va='center')
        ax.axis('off')
        return fig
    fig, ax = plt.subplots(figsize=(8, max(4, 0.4*len(top))))
    top.sort_values().plot(kind='barh', ax=ax)
    ax.set_xlabel('Total Matrículas')
    ax.set_ylabel('Producto')
    ax.set_title(f'Top {len(top)} Productos por Matriculados')
    plt.tight_layout()
    return fig

def grafica_burbujas_oferta(portafolio_df):
    if portafolio_df is None or portafolio_df.empty:
        # Crear un DataFrame mínimo para evitar errores
        df = pd.DataFrame({
            "Categoria": ["Sin datos"],
            "Programa": ["N/A"],
            "Valor": [1]
        })
    else:
        df = portafolio_df.copy()

        # Normalizar columnas (si tienes una función, puedes usarla aquí)
        df.columns = df.columns.str.lower()

        # Buscar columnas adecuadas
        cat = None
        prog = None
        val = None

        for col in df.columns:
            if cat is None and df[col].dtype == 'object':
                cat = col
            elif prog is None and df[col].dtype == 'object' and col != cat:
                prog = col
            elif val is None and pd.api.types.is_numeric_dtype(df[col]):
                val = col

        # Si no se encuentran columnas, crear columnas ficticias
        if cat is None:
            df["categoria"] = "Sin categoría"
            cat = "categoria"
        if prog is None:
            df["programa"] = "Programa"
            prog = "programa"
        if val is None:
            df["valor"] = 1
            val = "valor"
        # Agrupar y graficar
        resumen = df.groupby([cat, prog])[val].sum().reset_index()
        fig = px.scatter(resumen, x=cat, y=prog, size=val, color=cat,
                        title="Distribución por categoría y programa",
                        size_max=60)
        return fig



def descargar_datos_csv(df):
    return df.to_csv(index=False).encode('utf-8')

def descargar_figura_png_matplotlib(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    return buf

# -------------------------
# Main
# -------------------------
def main():
    st.title("Dashboard Interactivo ElTelar 2024-2025")
    with st.spinner("Cargando archivos (buscando en /mnt/data o usando fallback remoto)..."):
        df_hist_raw, df_prop_raw, df_port_raw = cargar_todo_local_o_remoto()
    st.sidebar.header("Filtros y Opciones")
    st.sidebar.markdown("Personaliza los filtros para analizar datos históricos y propuesta.")

    hist = armonizar_historico(df_hist_raw)
    prop = armonizar_propuesta(df_prop_raw)

    if prop is not None and not prop.empty:
        if 'año' not in prop.columns or prop['año'].isna().all():
            prop['año'] = 2025
        if 'trimestre' not in prop.columns or prop['trimestre'].isna().all():
            prop['trimestre'] = "PRUEBA"

    if hist is not None and not hist.empty:
        if 'año' not in hist.columns or hist['año'].isna().all():
            hist['año'] = 2024

    hist = sanitize_periodos(hist)
    prop = sanitize_periodos(prop)

    portafolio = df_port_raw.copy()
    if not portafolio.empty:
        portafolio = normalize_columns_df(portafolio)

    datos = pd.concat([hist, prop], ignore_index=True, sort=False)
    datos = sanitize_periodos(datos)

    if datos.empty:
        st.warning("No se cargaron datos. Verifica que los archivos Excel existan en /mnt/data o que las URLs remotas sean accesibles.")
        st.stop()

    conteo_hist = len(hist) if hist is not None else 0
    conteo_prop = len(prop) if prop is not None else 0

    st.sidebar.write(f"Filas cargadas — Histórico: {conteo_hist}  |  Propuesta: {conteo_prop}  |  Total: {len(datos)}")

    hojas_hist = sorted(df_hist_raw['__source_sheet__'].dropna().unique().tolist()) if (df_hist_raw is not None and '__source_sheet__' in df_hist_raw.columns) else []
    st.sidebar.write(f"Hojas histórico: {hojas_hist}")

    try:
        trimestres_unicos = sorted([int(x) for x in pd.Series(datos.get('trimestre', pd.Series(dtype=float))).dropna().unique().tolist()])
    except Exception:
        trimestres_unicos = list(pd.Series(datos.get('trimestre', pd.Series(dtype=float))).dropna().unique().tolist())

    st.sidebar.write(f"Trimestres detectados: {trimestres_unicos}")

    fuentes = sorted(datos['fuente'].dropna().unique().tolist()) if 'fuente' in datos.columns else []
    anos = sorted([int(x) for x in pd.Series(datos.get('año', pd.Series(dtype=float))).dropna().unique().tolist()]) if 'año' in datos.columns else []
    trimestres = trimestres_unicos if trimestres_unicos else []
    productos = sorted(datos['producto'].dropna().unique().tolist()) if 'producto' in datos.columns else []
    espacios = sorted(datos['espacio'].dropna().unique().tolist()) if 'espacio' in datos.columns else []
    niveles = sorted(datos.get('nivel', pd.Series(dtype=str)).dropna().unique().tolist()) if 'nivel' in datos.columns else []
    estados = sorted(datos.get('estado', pd.Series(dtype=str)).dropna().unique().tolist()) if 'estado' in datos.columns else []

    fuente_sel = st.sidebar.multiselect("Fuente", options=fuentes, default=fuentes)
    ano_sel = st.sidebar.multiselect("Año", options=anos, default=anos)
    trimestre_sel = st.sidebar.multiselect("Trimestre", options=trimestres, default=trimestres)
    producto_sel = st.sidebar.multiselect("Producto", options=productos, default=productos[:10] if productos else productos)
    espacio_sel = st.sidebar.multiselect("Espacio", options=espacios, default=espacios[:10] if espacios else espacios)
    nivel_sel = st.sidebar.multiselect("Nivel / Escuela", options=niveles, default=niveles if niveles else [])
    estado_sel = st.sidebar.multiselect("Estado", options=estados, default=estados if estados else [])

    if not fuente_sel: fuente_sel = fuentes.copy()
    if 'año' in datos.columns and not ano_sel: ano_sel = anos.copy()
    if 'trimestre' in datos.columns and not trimestre_sel: trimestre_sel = trimestres.copy()
    if 'producto' in datos.columns and not producto_sel: producto_sel = productos.copy()
    if 'espacio' in datos.columns and not espacio_sel: espacio_sel = espacios.copy()
    if 'nivel' in datos.columns and not nivel_sel: nivel_sel = niveles.copy()
    if 'estado' in datos.columns and not estado_sel: estado_sel = estados.copy()

    mask = pd.Series(True, index=datos.index)
    mask &= datos['fuente'].isin(fuente_sel)
    mask &= datos['año'].isin(ano_sel)
    mask &= datos['trimestre'].isin(trimestre_sel)
    mask &= datos['producto'].isin(producto_sel)
    mask &= datos['espacio'].isin(espacio_sel)
    if 'nivel' in datos.columns:
        mask &= datos['nivel'].isin(nivel_sel)
    if 'estado' in datos.columns:
        mask &= datos['estado'].isin(estado_sel)

    df_filtrado = datos[mask].copy()
    st.sidebar.write(f"Filas después de filtrar: {len(df_filtrado)}")

    hist_for_kpi = hist.copy() if hist is not None else pd.DataFrame()
    prop_for_kpi = prop.copy() if prop is not None else pd.DataFrame()

    if not hist_for_kpi.empty and 'año' in hist_for_kpi.columns and ano_sel:
        hist_for_kpi = hist_for_kpi[hist_for_kpi['año'].isin(ano_sel)]
    if not hist_for_kpi.empty and 'trimestre' in hist_for_kpi.columns and trimestre_sel:
        hist_for_kpi = hist_for_kpi[hist_for_kpi['trimestre'].isin(trimestre_sel)]

    st.write("Filas propuesta después del filtro para KPI:", len(prop_for_kpi))
    st.write("Suma total_matriculas en prop_for_kpi:", prop_for_kpi['total_matriculas'].sum())

    th, tp, diff, pct = kpi_variacion(hist_for_kpi, prop_for_kpi)

    st.markdown("## KPIs principales")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Matrículas Histórico (filtradas por periodo)", f"{th:,}")
    k2.metric("Matrículas Propuesta (filtradas por periodo)", f"{tp:,}")
    k3.metric("Variación absoluta", f"{diff:,}",
              delta_color="normal" if diff >= 0 else "inverse")
    k4.metric("Variación %", f"{pct:.1f}%" if not np.isnan(pct) else "N/A",
              delta_color="normal" if pct >= 0 else "inverse")

    ocup_hist = kpi_ocupacion_promedio(hist_for_kpi)
    ocup_prop = kpi_ocupacion_promedio(prop_for_kpi)
    ocup_hist_txt = f"{ocup_hist:.1f}%" if not np.isnan(ocup_hist) else "N/A"
    ocup_prop_txt = f"{ocup_prop:.1f}%" if not np.isnan(ocup_prop) else "N/A"
    st.markdown(f"Ocupación promedio — Histórico: {ocup_hist_txt}  |  Propuesta: {ocup_prop_txt}")

    st.markdown("### Alertas de sobrecupo (filtradas)")
    sob = alertas_sobrecupo(df_filtrado)
    if sob is None or sob.empty:
        st.success("No se detectan alertas de sobrecupo con los filtros actuales.")
    else:
        display_cols = [c for c in ['año', 'trimestre', 'producto', 'espacio', 'total_matriculas', 'cap_mx_e'] if c in sob.columns]
        st.dataframe(sob[display_cols].sort_values(['año', 'trimestre'], ascending=False).reset_index(drop=True))

    st.markdown("### Visualizaciones")
    fig1 = grafica_serie_temporal(hist_for_kpi, prop_for_kpi)
    st.pyplot(fig1)
    fig2 = grafica_barras_apiladas(hist_for_kpi, prop_for_kpi, topn=8)
    st.plotly_chart(fig2, use_container_width=True)
    fig3 = grafica_heatmap_ocupacion(df_filtrado)
    st.pyplot(fig3)
    fig4 = grafica_barras_horizontales_top_productos(df_filtrado, topn=10)
    st.pyplot(fig4)
    fig5 = grafica_burbujas_oferta(portafolio)
    if fig5 is not None:
        st.plotly_chart(fig5, use_container_width=True)
    st.markdown("### Exportar")
    csv = descargar_datos_csv(df_filtrado)
    st.download_button("Descargar datos filtrados (CSV)", data=csv, file_name="eltelar_datos_filtrados.csv", mime="text/csv")
    buf_png = descargar_figura_png_matplotlib(fig1)
    st.download_button("Descargar serie temporal (PNG)", data=buf_png, file_name="serie_temporal.png", mime="image/png")

    st.markdown("---")
    st.write("Instrucciones: utiliza los filtros en la barra lateral. Si alguna visualización indica que faltan columnas, revisa las primeras columnas mostradas en el mensaje para ver cómo vienen con tus datos y actualiza los mapeos si lo deseas.")

if __name__ == "__main__":
    main()
