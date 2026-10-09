import streamlit as st
import pandas as pd
import openpyxl
import re
import warnings
import unicodedata
import plotly.express as px
import csv

# Configuración de la página de Streamlit
st.set_page_config(
    page_title="Auditoría de Faltantes - Postobón",
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
    if pd.isna(val): return "1"
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


# --- CARGAR HISTÓRICO CSV ---
def cargar_historico_csv(file_obj):
    if file_obj is None: return pd.DataFrame()
    try:
        rows = []
        content = file_obj.getvalue().decode('latin1')
        reader = csv.reader(content.splitlines(), delimiter=';')
        for r in reader:
            if any(str(cell).strip() for cell in r):
                rows.append(r)
            
        if not rows: return pd.DataFrame()
        
        header_idx = -1
        for idx, r in enumerate(rows):
            row_str = " ".join([str(c).upper() for c in r])
            if "VALOR" in row_str or "IMPORTE" in row_str or "SALDO" in row_str:
                header_idx = idx
                break
                
        if header_idx == -1: header_idx = 2
        
        headers = [str(c).upper().replace('\n', ' ').strip() for c in rows[header_idx]]
        data_rows = rows[header_idx + 1:]
        
        if len(data_rows) > 5000:
            data_rows = data_rows[-5000:]
            
        df = pd.DataFrame(data_rows)
        df.columns = headers[:len(df.columns)]
        df = df.loc[:, df.columns.notna()].copy()
        df = df.loc[:, ~df.columns.duplicated(keep='first')].copy()
        
        col_sap = next((c for c in df.columns if 'CLIENTE' in c or 'SAP' in c or 'DUEDOR' in c), df.columns[2] if len(df.columns) > 2 else None)
        col_nombre = next((c for c in df.columns if 'NOMBRE' in c or 'TITULAR' in c), df.columns[3] if len(df.columns) > 3 else None)
        col_fecha = next((c for c in df.columns if 'FECHA' in c), df.columns[4] if len(df.columns) > 4 else None)
        col_valor = next((c for c in df.columns if 'VALOR' in c or 'IMPORTE' in c), df.columns[7] if len(df.columns) > 7 else None)
        col_abono = next((c for c in df.columns if 'ABONO' in c), df.columns[8] if len(df.columns) > 8 else None)
        col_saldo = next((c for c in df.columns if 'SALDO' in c), df.columns[9] if len(df.columns) > 9 else None)
        
        df['Deudor_SAP_OK'] = df[col_sap].apply(normalizar_sap) if col_sap else '1'
        df['Nombre_Deudor_OK'] = df[col_nombre].apply(remover_tildes_y_raros) if col_nombre else 'CLIENTE'
        df['Valor_Faltante_Num'] = df[col_valor].apply(limpiar_monto) if col_valor else 0.0
        df['Abonos_Num'] = df[col_abono].apply(limpiar_monto) if col_abono else 0.0
        df['Saldo_Num'] = df[col_saldo].apply(limpiar_monto) if col_saldo else (df['Valor_Faltante_Num'] - df['Abonos_Num'])
        df['Fecha_DT'] = df[col_fecha].apply(convertir_a_fecha) if col_fecha else pd.NaT
        df['Fecha_UI'] = df['Fecha_DT'].apply(formatear_fecha_ui)
        
        return df[(df['Saldo_Num'] > 0) & (df['Deudor_SAP_OK'] != '') & (df['Fecha_DT'].notna()) & (df['Fecha_DT'].dt.year >= 2026)].copy()
    except Exception as e:
        return pd.DataFrame()


# --- CARGAR ANEXO EXCEL ---
def cargar_anexo_excel(file_obj):
    if file_obj is None: return pd.DataFrame()
    try:
        wb = openpyxl.load_workbook(file_obj, data_only=True)
        sheet = wb.active
        df_raw = pd.DataFrame(list(sheet.values))
        df_raw = df_raw.dropna(how='all').reset_index(drop=True)
        
        header_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = [str(c).upper().replace('\n', ' ').strip() for c in row if pd.notna(c)]
            if any("VALOR" in c or "IMPORTE" in c for c in row_str) and any("FECHA" in c for c in row_str):
                header_idx = idx
                break
                
        df = df_raw.iloc[header_idx + 1:].copy()
        headers = [str(c).upper().replace('\n', ' ').strip() if pd.notna(c) else f"COL_{i}" for i, c in enumerate(df_raw.iloc[header_idx].values)]
        df.columns = headers[:len(df.columns)]
        df = df.loc[:, df.columns.notna()].copy()
        df = df.loc[:, ~df.columns.duplicated(keep='first')].copy()
        
        col_client_code = next((c for c in df.columns if 'CODIGO' in c and 'CLIENTE' in c), None)
        col_client_name = next((c for c in df.columns if 'NOMBRE' in c and 'CLIENTE' in c), None)
        col_trans_code = next((c for c in df.columns if 'DUEDOR SAP' in c or 'SAP' in c), None)
        col_trans_name = next((c for c in df.columns if 'TITULAR' in c or 'TRANSPORTADOR' in c), None)
        col_fecha = next((c for c in df.columns if 'FECHA' in c), None)
        col_valor = next((c for c in df.columns if 'VALOR' in c or 'IMPORTE' in c), None)
        col_abono = next((c for c in df.columns if 'ABONO' in c), None)
        col_saldo = next((c for c in df.columns if 'SALDO' in c), None)
        
        df['Valor_Faltante_Num'] = df[col_valor].apply(limpiar_monto) if col_valor else 0.0
        df['Abonos_Num'] = df[col_abono].apply(limpiar_monto) if col_abono else 0.0
        df['Saldo_Num'] = df[col_saldo].apply(limpiar_monto) if col_saldo else (df['Valor_Faltante_Num'] - df['Abonos_Num'])
        
        nombres_finales, saps_finales = [], []
        for _, row in df.iterrows():
            c_name = remover_tildes_y_raros(row[col_client_name]) if col_client_name and pd.notna(row[col_client_name]) else ''
            c_code = normalizar_sap(row[col_trans_code]) if col_trans_code and pd.notna(row[col_trans_code]) else '1'
            nombres_finales.append(c_name if c_name else 'CLIENTE')
            saps_finales.append(c_code)
            
        df['Nombre_Deudor_OK'] = nombres_finales
        df['Deudor_SAP_OK'] = saps_finales
        df['Fecha_DT'] = df[col_fecha].apply(convertir_a_fecha) if col_fecha else pd.NaT
        df['Fecha_UI'] = df['Fecha_DT'].apply(formatear_fecha_ui)
        
        return df[(df['Valor_Faltante_Num'] > 0) & (df['Fecha_DT'].notna())].copy()
    except Exception as e:
        return pd.DataFrame()


# --- MOTOR DE AUDITORÍA Y CLASIFICACIÓN ---
def procesar_archivos(file_hist, file_anexo):
    if file_anexo is None:
        return "⚠️ Por favor sube el Anexo del día para realizar la auditoría.", None, None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_anexo = cargar_anexo_excel(file_anexo)
    df_hist = cargar_historico_csv(file_hist) if file_hist is not None else pd.DataFrame()
    
    if df_anexo.empty:
        return "⚠️ El anexo cargado no tiene registros válidos.", None, None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    if not df_anexo['Fecha_DT'].isna().all():
        max_dt = df_anexo['Fecha_DT'].max()
        mes_actual = max_dt.month
        anio_actual = max_dt.year
    else:
        mes_actual, anio_actual = 10, 2026
        
    # 1. Saldados (Saldo == 0)
    saldados = df_anexo[df_anexo['Saldo_Num'] == 0].copy()
    
    # 2. Con saldo pendiente en el anexo de hoy
    con_saldo_anexo = df_anexo[df_anexo['Saldo_Num'] > 0].copy()
    
    # Nuevos Faltantes (Mes actual en curso: Octubre 2026)
    es_mes_actual = (con_saldo_anexo['Fecha_DT'].dt.month == mes_actual) & (con_saldo_anexo['Fecha_DT'].dt.year == anio_actual)
    nuevos = con_saldo_anexo[es_mes_actual].copy()
    
    # Pendientes Activos (Meses anteriores que siguen vigentes en el anexo de hoy)
    pendientes = con_saldo_anexo.drop(nuevos.index).copy()
    
    # 3. Omitidos en Anexo (Estrictamente del mes actual que estaban en el histórico con saldo y faltan hoy)
    omitidos = pd.DataFrame()
    if not df_hist.empty:
        pendientes_hist = df_hist[
            (df_hist['Saldo_Num'] > 0) & 
            (df_hist['Fecha_DT'].dt.month == mes_actual) & 
            (df_hist['Fecha_DT'].dt.year == anio_actual)
        ].copy()
        
        montos_presentes_hoy = set(df_anexo['Valor_Faltante_Num'].round(2))
        
        omitidos_list = []
        for _, row in pendientes_hist.iterrows():
            monto_hist = round(row['Valor_Faltante_Num'], 2)
            if monto_hist not in montos_presentes_hoy and row['Saldo_Num'] > 0:
                omitidos_list.append(row)
                
        if omitidos_list:
            omitidos = pd.DataFrame(omitidos_list).drop_duplicates(subset=['Deudor_SAP_OK', 'Valor_Faltante_Num'])
            
    def preparar_df_ui(df_in):
        if df_in.empty:
            return pd.DataFrame(columns=['Nombre Deudor / Transportador', 'SAP', 'Fecha (DD/MM/AAAA)', 'Valor Faltante ($)', 'Abonos ($)', 'Saldo Pendiente ($)'])
        out = df_in[['Nombre_Deudor_OK', 'Deudor_SAP_OK', 'Fecha_UI', 'Valor_Faltante_Num', 'Abonos_Num', 'Saldo_Num']].copy()
        out.columns = ['Nombre Deudor / Transportador', 'SAP', 'Fecha (DD/MM/AAAA)', 'Valor Faltante ($)', 'Abonos ($)', 'Saldo Pendiente ($)']
        out['Valor Faltante ($)'] = out['Valor Faltante ($)'].apply(lambda x: f"${int(round(limpiar_monto(x))):,}")
        out['Abonos ($)'] = out['Abonos ($)'].apply(lambda x: f"${int(round(limpiar_monto(x))):,}")
        out['Saldo Pendiente ($)'] = out['Saldo Pendiente ($)'].apply(lambda x: f"${int(round(limpiar_monto(x))):,}")
        return out
        
    df_nuevos_ui = preparar_df_ui(nuevos)
    df_saldados_ui = preparar_df_ui(saldados)
    df_pendientes_ui = preparar_df_ui(pendientes)
    df_omitidos_ui = preparar_df_ui(omitidos)
    
    tot_nuevos = nuevos['Saldo_Num'].sum() if not nuevos.empty else 0
    tot_saldados = saldados['Abonos_Num'].sum() if not saldados.empty else 0
    tot_pendientes = pendientes['Saldo_Num'].sum() if not pendientes.empty else 0
    tot_omitidos = omitidos['Saldo_Num'].sum() if not omitidos.empty else 0
    
    m_str = f"## 🥤 **AUDITORÍA DE FALTANTES - POSTOBÓN**\n\n"
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
st.markdown("# 🥤 **POSTOBÓN - SISTEMA DE AUDITORÍA DE FALTANTES EN CAJA**")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Carga de Archivos")
    file_hist = st.file_uploader("Histórico Maestro (.csv)", type=['csv'])
    file_anexo = st.file_uploader("Anexo del Día Actual (.xlsx / .xlsm)", type=['xlsx', 'xlsm', 'csv'])

with col2:
    st.subheader("Estado del Proceso")
    ejecutar = st.button("🚀 EJECUTAR AUDITORÍA", type="primary")

if ejecutar:
    if file_anexo is None:
        st.warning("⚠️ Por favor sube el Anexo del día para realizar la auditoría.")
    else:
        with st.spinner("Procesando auditoría de Jamundí..."):
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
                st.info("No hay registros omitidos detectados en el periodo actual.")
            else:
                st.markdown("### ⚠️ Atención: Estos faltantes del mes actual estaban en el histórico con saldo pendiente, pero fueron omitidos en el anexo de hoy.")
                st.dataframe(df_omitidos, use_container_width=True)
else:
    st.info("💡 Sube el **Histórico Maestro (.csv)** y el **Anexo del día**, luego presiona **EJECUTAR AUDITORÍA**.")
