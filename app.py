import streamlit as st
import pandas as pd
import openpyxl
import re
import warnings
import unicodedata
import plotly.express as px

# Configuración de la página de Streamlit
st.set_page_config(
    page_title="Auditoría de Faltantes - Postobón Jamundí",
    page_icon="🥤",
    layout="wide"
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# --- FUNCIONES DE LIMPIEZA Y FORMATO ---
def remover_tildes_y_raros(texto):
    if pd.isna(texto): return ""
    txt = str(texto).upper().strip()
    txt = unicodedata.normalize('NFD', txt)
    txt = ''.join(c for c in txt if unicodedata.category(c) != 'Mn')
    txt = re.sub(r'[^A-Z0-9]', ' ', txt)
    return re.sub(r'\s+', ' ', txt).strip()

def normalizar_sap(val):
    if pd.isna(val): return ""
    txt = str(val).strip().split('.')[0]
    txt = re.sub(r'\D', '', txt)
    if txt in ['1', ''] or not txt: return '1'
    return txt

def limpiar_monto(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).replace('$', '').replace('-', '0').strip()
    if not val_str or val_str.lower() in ['none', 'nan']: return 0.0
    val_str = val_str.replace('.', '').replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

def convertir_a_fecha(val):
    if pd.isna(val): return pd.NaT
    dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
    return dt

def formatear_fecha_ui(val_dt):
    if pd.isna(val_dt): return 'Sin Fecha'
    return val_dt.strftime('%d/%m/%Y')


# --- CARGA Y PROCESAMIENTO ROBUSTO ---
def cargar_y_limpiar(file_obj):
    if file_obj is None: return pd.DataFrame()
    filename = file_obj.name
    df_raw = pd.DataFrame()
    
    if filename.endswith('.csv'):
        encodings = ['utf-8-sig', 'latin1', 'cp1252', 'utf-8']
        for enc in encodings:
            try:
                df_raw = pd.read_csv(file_obj, encoding=enc, sep=None, engine='python', header=None)
                break
            except Exception:
                continue
    else:
        wb = openpyxl.load_workbook(file_obj, data_only=True)
        sheet = wb.active
        df_raw = pd.DataFrame(list(sheet.values))
        
    if df_raw.empty: return pd.DataFrame()
        
    df_raw = df_raw.dropna(how='all').reset_index(drop=True)
    
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = [str(c).upper().replace('\n', ' ').strip() for c in row if pd.notna(c)]
        if any("VALOR" in c for c in row_str) and any("SALDO" in c for c in row_str):
            header_idx = idx
            break
            
    if header_idx is None: header_idx = 0
    
    df = df_raw.iloc[header_idx + 1:].copy()
    headers = [str(c).upper().replace('\n', ' ').strip() if pd.notna(c) else f"COL_{i}" for i, c in enumerate(df_raw.iloc[header_idx].values)]
    df.columns = headers
    
    df = df.loc[:, df.columns.notna()].copy()
    df = df.loc[:, ~df.columns.duplicated(keep='first')].copy()
    
    # Búsqueda flexible de columnas
    col_client_code = next((c for c in df.columns if 'CODIGO' in c and 'CLIENTE' in c), None)
    if not col_client_code: col_client_code = next((c for c in df.columns if 'CLIENTE' in c), None)
    
    col_client_name = next((c for c in df.columns if 'NOMBRE' in c and 'CLIENTE' in c), None)
    
    col_trans_code = next((c for c in df.columns if 'TRANSPORTADOR' in c and ('CODIGO' in c or 'ZONA' in c)), None)
    col_trans_name = next((c for c in df.columns if 'TRANSPORTADOR' in c and 'NOMBRE' in c), None)
    
    col_fecha = next((c for c in df.columns if 'FECHA' in c and ('GENER' in c or 'TR' in c or 'DOC' in c or 'FALT' in c)), None)
    if not col_fecha: col_fecha = next((c for c in df.columns if 'FECHA' in c), None)
    
    col_valor = next((c for c in df.columns if 'VALOR' in c and ('FALT' in c or 'ORIG' in c)), None)
    if not col_valor: col_valor = next((c for c in df.columns if 'VALOR' in c), None)
    
    col_abono = next((c for c in df.columns if 'ABONO' in c), None)
    col_saldo = next((c for c in df.columns if 'SALDO' in c), None)
    
    df['Valor_Faltante_Num'] = df[col_valor].apply(limpiar_monto) if col_valor else 0.0
    df['Abonos_Num'] = df[col_abono].apply(limpiar_monto) if col_abono else 0.0
    
    if col_saldo:
        df['Saldo_Num'] = df[col_saldo].apply(limpiar_monto)
    else:
        df['Saldo_Num'] = df['Valor_Faltante_Num'] - df['Abonos_Num']
        
    nombres_finales, saps_finales = [], []
    for _, row in df.iterrows():
        c_code = normalizar_sap(row[col_client_code]) if col_client_code else ''
        c_name = remover_tildes_y_raros(row[col_client_name]) if col_client_name else ''
        t_code = normalizar_sap(row[col_trans_code]) if col_trans_code else ''
        t_name = remover_tildes_y_raros(row[col_trans_name]) if col_trans_name else ''
        
        if c_code == '1' or c_name in ['TRANSPORTADOR', '1', '']:
            final_name = t_name if t_name else 'TRANSPORTADOR'
            final_sap = t_code if t_code else c_code
        else:
            final_name = c_name if c_name else (t_name if t_name else 'CLIENTE')
            final_sap = c_code if c_code else t_code
            
        nombres_finales.append(final_name)
        saps_finales.append(final_sap)
        
    df['Nombre_Deudor_OK'] = nombres_finales
    df['Deudor_SAP_OK'] = saps_finales
    
    df['Fecha_DT'] = df[col_fecha].apply(convertir_a_fecha) if col_fecha else pd.NaT
    df['Fecha_UI'] = df['Fecha_DT'].apply(formatear_fecha_ui)
    
    df = df[(df['Valor_Faltante_Num'] > 0) | (df['Abonos_Num'] > 0) | (df['Saldo_Num'] > 0)].copy()
    return df


# --- AUDITORÍA Y CLASIFICACIÓN (JAMUNDÍ) ---
def procesar_archivos(file_hist, file_anexo):
    if file_anexo is None:
        return "⚠️ Por favor sube el Anexo del día para realizar la auditoría.", None, None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_anexo = cargar_y_limpiar(file_anexo)
    df_hist = cargar_y_limpiar(file_hist) if file_hist is not None else pd.DataFrame()
    
    saldados = df_anexo[df_anexo['Saldo_Num'] == 0].copy()
    con_saldo = df_anexo[df_anexo['Saldo_Num'] > 0].copy()
    
    if not con_saldo['Fecha_DT'].isna().all():
        max_dt = con_saldo['Fecha_DT'].max()
        mes_actual = max_dt.month
        anio_actual = max_dt.year
    else:
        mes_actual, anio_actual = 10, 2026
        
    es_mes_actual = (con_saldo['Fecha_DT'].dt.month == mes_actual) & (con_saldo['Fecha_DT'].dt.year == anio_actual)
    nuevos = con_saldo[es_mes_actual].copy()
    pendientes = con_saldo.drop(nuevos.index).copy()
    
    # CRUCE INTELIGENTE Y FLEXIBLE PARA OMITIDOS
    omitidos = pd.DataFrame()
    if not df_hist.empty:
        # Filtramos los que tenían saldo en el histórico
        pendientes_hist = df_hist[df_hist['Saldo_Num'] > 0].copy()
        
        # Creamos conjuntos de identificación flexibles usando SAP y Monto exacto
        claves_anexo_actual = set(zip(con_saldo['Deudor_SAP_OK'].astype(str), con_saldo['Valor_Faltante_Num'].round(2)))
        
        omitidos_list = []
        for _, row in pendientes_hist.iterrows():
            clave_hist = (str(row['Deudor_SAP_OK']), round(row['Valor_Faltante_Num'], 2))
            # Si el registro estaba en el histórico pero NO aparece en el anexo de hoy, ¡cayó en omisión!
            if clave_hist not in claves_anexo_actual and row['Saldo_Num'] > 0:
                omitidos_list.append(row)
                
        if omitidos_list:
            omitidos = pd.DataFrame(omitidos_list).drop_duplicates(subset=['Deudor_SAP_OK', 'Valor_Faltante_Num'])
            
    def preparar_df_ui(df_in):
        if df_in.empty:
            return pd.DataFrame(columns=['Nombre Deudor / Transportador', 'SAP', 'Fecha (DD/MM/AAAA)', 'Valor Faltante ($)', 'Abonos ($)', 'Saldo Pendiente ($)'])
        out = df_in[['Nombre_Deudor_OK', 'Deudor_SAP_OK', 'Fecha_UI', 'Valor_Faltante_Num', 'Abonos_Num', 'Saldo_Num']].copy()
        out.columns = ['Nombre Deudor / Transportador', 'SAP', 'Fecha (DD/MM/AAAA)', 'Valor Faltante ($)', 'Abonos ($)', 'Saldo Pendiente ($)']
        out['Valor Faltante ($)'] = out['Valor Faltante ($)'].apply(lambda x: f"${int(round(x)):,}")
        out['Abonos ($)'] = out['Abonos ($)'].apply(lambda x: f"${int(round(x)):,}")
        out['Saldo Pendiente ($)'] = out['Saldo Pendiente ($)'].apply(lambda x: f"${int(round(x)):,}")
        return out
        
    df_nuevos_ui = preparar_df_ui(nuevos)
    df_saldados_ui = preparar_df_ui(saldados)
    df_pendientes_ui = preparar_df_ui(pendientes)
    df_omitidos_ui = preparar_df_ui(omitidos)
    
    tot_nuevos = nuevos['Saldo_Num'].sum() if not nuevos.empty else 0
    tot_saldados = saldados['Abonos_Num'].sum() if not saldados.empty else 0
    tot_pendientes = pendientes['Saldo_Num'].sum() if not pendientes.empty else 0
    tot_omitidos = omitidos['Saldo_Num'].sum() if not omitidos.empty else 0
    
    m_str = f"## 🥤 **AUDITORÍA DE FALTANTES - POSTOBÓN JAMUNDÍ**\n\n"
    m_str += f"- 🔴 **Nuevos Faltantes (Mes Actual):** {len(nuevos)} registro(s) | **${int(tot_nuevos):,}**\n"
    m_str += f"- 🟢 **Faltantes Saldados / Pagados:** {len(saldados)} registro(s) | **${int(tot_saldados):,}**\n"
    m_str += f"- 🟣 **Pendientes Activos (Meses Anteriores):** {len(pendientes)} registro(s) | **${int(tot_pendientes):,}**\n"
    m_str += f"- ⚠️ **Omitidos en Anexo (Revisar Negligencia):** {len(omitidos)} registro(s) | **${int(tot_omitidos):,}**\n"
    
    categorias = ['Nuevos Faltantes', 'Faltantes Saldados', 'Pendientes Activos', 'Omitidos en Anexo']
    cantidades = [len(nuevos), len(saldados), len(pendientes), len(omitidos)]
    montos = [tot_nuevos, tot_saldados, tot_pendientes, tot_omitidos]
    colores = {'Nuevos Faltantes': '#EF553B', 'Faltantes Saldados': '#00CC96', 'Pendientes Activos': '#AB63FA', 'Omitidos en Anexo': '#FFA15A'}
    
    df_summary = pd.DataFrame({
        'Categoría': categorias,
        'Cantidad': cantidades,
        'Monto': montos
    })
    
    fig_pie = px.pie(
        df_summary, values='Cantidad', names='Categoría',
        title="<b>Distribución por Cantidad de Casos</b>",
        color='Categoría',
        color_discrete_map=colores,
        hole=0.4
    )
    
    fig_bar = px.bar(
        df_summary, x='Categoría', y='Monto', text_auto='.0f',
        title="<b>Impacto Financiero Real ($ COP)</b>",
        color='Categoría',
        color_discrete_map=colores
    )
    fig_bar.update_layout(showlegend=False, yaxis_title="Monto ($)")
    
    return m_str, fig_pie, fig_bar, df_nuevos_ui, df_saldados_ui, df_pendientes_ui, df_omitidos_ui


# --- INTERFAZ EN STREAMLIT ---
st.markdown("# 🥤 **POSTOBÓN - SISTEMA DE AUDITORÍA DE FALTANTES EN CAJA (JAMUNDÍ)**")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Carga de Archivos")
    file_hist = st.file_uploader("Histórico Maestro (Obligatorio para detectar omitidos)", type=['csv', 'xlsx', 'xlsm'])
    file_anexo = st.file_uploader("Anexo del Día Actual (.xlsm / .xlsx / .csv)", type=['xlsm', 'xlsx', 'csv'])

with col2:
    st.subheader("Estado del Proceso")
    ejecutar = st.button("🚀 EJECUTAR AUDITORÍA", type="primary")

if ejecutar:
    if file_anexo is None:
        st.warning("⚠️ Por favor sube el Anexo del día para realizar la auditoría.")
    else:
        with st.spinner("Procesando datos y cruzando históricos de Jamundí..."):
            m_str, fig_pie, fig_bar, df_nuevos, df_saldados, df_pendientes, df_omitidos = procesar_archivos(file_hist, file_anexo)
            
        st.markdown("---")
        st.markdown(m_str)
        
        gcol1, gcol2 = st.columns(2)
        with gcol1:
            st.plotly_chart(fig_pie, use_container_width=True)
        with gcol2:
            st.plotly_chart(fig_bar, use_container_width=True)
            
        st.markdown("## 📄 **Detalle Completo de Registros**")
        
        tab1, tab2, tab3, tab4 = st.tabs(["🔴 Nuevos Faltantes", "🟢 Faltantes Saldados / Pagados", "🟣 Pendientes Activos", "⚠️ Omitidos en Anexo"])
        
        with tab1:
            st.dataframe(df_nuevos, use_container_width=True)
        with tab2:
            st.dataframe(df_saldados, use_container_width=True)
        with tab3:
            st.dataframe(df_pendientes, use_container_width=True)
        with tab4:
            if df_omitidos.empty:
                st.info("No hay registros omitidos detectados o no se cargó el archivo de histórico maestro correctamente.")
            else:
                st.markdown("### ⚠️ Atención: Estos faltantes estaban en el histórico con saldo pendiente, pero **desaparecieron** del anexo de hoy. ¡Aquí aparece el caso del Grupo Surtipanes u otras omisiones del cajero!")
                st.dataframe(df_omitidos, use_container_width=True)
else:
    st.info("💡 Sube el **Histórico Maestro** y el **Anexo del día**, luego presiona **EJECUTAR AUDITORÍA**.")
