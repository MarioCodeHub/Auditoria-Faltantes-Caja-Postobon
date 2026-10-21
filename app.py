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
    page_title="Auditoría de Faltantes - Jamundí",
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
    if isinstance(val, (int, float)):
        try:
            return pd.to_datetime('1899-12-30') + pd.to_timedelta(val, unit='D')
        except:
            pass
    dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
    return dt

def formatear_fecha_ui(val_dt):
    if pd.isna(val_dt): return 'Sin Fecha'
    return val_dt.strftime('%d/%m/%Y')


# --- CARGAR HISTÓRICO MAESTRO (.CSV) ---
def cargar_historico_csv(file_obj):
    if file_obj is None: return pd.DataFrame()
    try:
        content = file_obj.getvalue().decode('latin1')
        lines = content.splitlines()
        reader = csv.reader(lines, delimiter=';')
        rows = [r for r in reader if any(str(cell).strip() for cell in r)]
            
        if not rows: return pd.DataFrame()
        
        header_idx = -1
        for idx, r in enumerate(rows):
            row_str = " ".join([str(c).upper() for c in r])
            if "IMPORTE" in row_str or "VALOR" in row_str or "DUEDOR SAP" in row_str or "FECHA" in row_str or "CLIENTE" in row_str:
                header_idx = idx
                break
                
        if header_idx == -1: header_idx = 0
        data_rows = rows[header_idx + 1:]
        
        if len(data_rows) > 1000:
            data_rows = data_rows[-1000:]
        
        registros = []
        for r in data_rows:
            if len(r) < 5: continue
            
            sap_raw = r[2] if len(r) > 2 and r[2].replace('.','').isdigit() else (r[3] if len(r) > 3 else "1")
            titular_raw = r[3] if len(r) > 3 and not r[3].replace('.','').isdigit() else (r[4] if len(r) > 4 else "")
            fecha_raw = r[4] if len(r) > 4 else (r[3] if len(r) > 3 else "")
            valor_raw = r[7] if len(r) > 7 else (r[5] if len(r) > 5 else "0")
            abono_raw = r[8] if len(r) > 8 else (r[6] if len(r) > 6 else "0")
            saldo_raw = r[9] if len(r) > 9 else (r[7] if len(r) > 7 else "0")
            motivo_raw = r[11] if len(r) > 11 else (r[8] if len(r) > 8 else "")
            
            val_monto = limpiar_monto(valor_raw)
            val_abono = limpiar_monto(abono_raw)
            val_saldo = limpiar_monto(saldo_raw) if str(saldo_raw).strip() != '-' else (val_monto - val_abono)
            
            if val_monto <= 0: continue
            f_dt = convertir_a_fecha(fecha_raw)
            nombre_limpio = remover_tildes_y_raros(titular_raw) if titular_raw else "RESPONSABLE / ZONA"
            
            registros.append({
                'Deudor_SAP_OK': normalizar_sap(sap_raw),
                'Nombre_Deudor_OK': nombre_limpio,
                'Valor_Faltante_Num': val_monto,
                'Abonos_Num': val_abono,
                'Saldo_Num': max(0.0, val_saldo),
                'Fecha_DT': f_dt,
                'Fecha_UI': formatear_fecha_ui(f_dt),
                'Motivo_OK': str(motivo_raw).strip()
            })
            
        df = pd.DataFrame(registros)
        return df if not df.empty else pd.DataFrame()
    except Exception as e:
        return pd.DataFrame()


# --- CARGAR ANEXO EXCEL DIARIO (JAMUNDÍ) ---
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
            if any("VALOR" in c or "IMPORTE" in c or "FALTANTE" in c for c in row_str) and any("FECHA" in c for c in row_str):
                header_idx = idx
                break
                
        df = df_raw.iloc[header_idx + 1:].copy()
        headers = [str(c).upper().replace('\n', ' ').strip() if pd.notna(c) else f"COL_{i}" for i, c in enumerate(df_raw.iloc[header_idx].values)]
        df.columns = headers[:len(df.columns)]
        df = df.loc[:, df.columns.notna()].copy()
        df = df.loc[:, ~df.columns.duplicated(keep='first')].copy()
        
        col_client_name = next((c for c in df.columns if 'NOMBRE' in c and 'CLIENTE' in c), None)
        col_trans_code = next((c for c in df.columns if 'DUEDOR SAP' in c or 'SAP' in c or 'CODIGO' in c), None)
        col_trans_name = next((c for c in df.columns if 'TITULAR' in c or 'TRANSPORTADOR' in c or 'RESPONSABLE' in c), None)
        col_fecha = next((c for c in df.columns if 'FECHA' in c), None)
        col_valor = next((c for c in df.columns if 'VALOR' in c or 'IMPORTE' in c or 'FALTANTE' in c), None)
        col_abono = next((c for c in df.columns if 'ABONO' in c), None)
        col_saldo = next((c for c in df.columns if 'SALDO' in c), None)
        col_motivo = next((c for c in df.columns if 'MOTIVO' in c), None)
        
        df['Valor_Faltante_Num'] = df[col_valor].apply(limpiar_monto) if col_valor else 0.0
        df['Abonos_Num'] = df[col_abono].apply(limpiar_monto) if col_abono else 0.0
        df['Saldo_Num'] = df[col_saldo].apply(limpiar_monto) if col_saldo else (df['Valor_Faltante_Num'] - df['Abonos_Num'])
        
        nombres_finales, saps_finales, fechas_dt, motivos_finales = [], [], [], []
        
        for _, row in df.iterrows():
            c_name = ""
            if col_trans_name and pd.notna(row[col_trans_name]):
                c_name = str(row[col_trans_name]).strip()
            elif col_client_name and pd.notna(row[col_client_name]):
                c_name = str(row[col_client_name]).strip()
                
            c_code = normalizar_sap(row[col_trans_code]) if col_trans_code and pd.notna(row[col_trans_code]) else '1'
            motivo = str(row[col_motivo]).strip() if col_motivo and pd.notna(row[col_motivo]) else ''
            
            f_val = pd.NaT
            if col_fecha and pd.notna(row[col_fecha]):
                f_val = convertir_a_fecha(row[col_fecha])
            
            nombres_finales.append(remover_tildes_y_raros(c_name) if c_name and "POSTOBON" not in c_name.upper() else 'RESPONSABLE / ZONA')
            saps_finales.append(c_code)
            fechas_dt.append(f_val)
            motivos_finales.append(motivo)
            
        df['Nombre_Deudor_OK'] = nombres_finales
        df['Deudor_SAP_OK'] = saps_finales
        df['Fecha_DT'] = fechas_dt
        df['Motivo_OK'] = motivos_finales
        
        df = df[df['Valor_Faltante_Num'] > 0].copy()
        
        fechas_validas_temp = df['Fecha_DT'].dropna()
        if not fechas_validas_temp.empty:
            max_f = fechas_validas_temp.max()
            df['Fecha_DT'] = df['Fecha_DT'].fillna(max_f)
            
        df['Fecha_UI'] = df['Fecha_DT'].apply(formatear_fecha_ui)
        return df
    except Exception as e:
        return pd.DataFrame()


# --- MOTOR DE AUDITORÍA JAMUNDÍ CON CAZA DE TRAMPAS ---
def procesar_archivos(file_hist, file_anexo):
    if file_anexo is None:
        return "⚠️ Por favor sube el Anexo del día para realizar la auditoría.", None, None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_anexo = cargar_anexo_excel(file_anexo)
    df_hist = cargar_historico_csv(file_hist) if file_hist is not None else pd.DataFrame()
    
    if df_anexo.empty:
        return "⚠️ El anexo cargado no tiene registros válidos.", None, None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    fechas_validas = df_anexo['Fecha_DT'].dropna()
    fecha_dia_actual = fechas_validas.max() if not fechas_validas.empty else pd.Timestamp('2026-10-20')

    # 1. CATALOGAR TODO EL HISTÓRICO PENDIENTE (Saldo > 0)
    mapa_hist = {}
    if not df_hist.empty:
        for idx, r in df_hist.iterrows():
            if r['Saldo_Num'] > 0.0:
                f_str = r['Fecha_DT'].strftime('%Y-%m-%d') if pd.notna(r['Fecha_DT']) else ''
                key = (str(r['Deudor_SAP_OK']), f_str, str(r['Nombre_Deudor_OK']))
                
                if key not in mapa_hist:
                    mapa_hist[key] = []
                mapa_hist[key].append(r)

    # 2. CATALOGAR LO QUE LLEGÓ EN EL ANEXO DE HOY
    mapa_anexo = {}
    for idx, r in df_anexo.iterrows():
        f_dt = r['Fecha_DT']
        f_str = f_dt.strftime('%Y-%m-%d') if pd.notna(f_dt) else ''
        key = (str(r['Deudor_SAP_OK']), f_str, str(r['Nombre_Deudor_OK']))
        
        if key not in mapa_anexo:
            mapa_anexo[key] = []
        mapa_anexo[key].append(r)

    # 3. CAZAR OMITIDOS (Ej: la trampa de los 13 pesos)
    omitidos_list = []
    for key, lista_h in mapa_hist.items():
        cant_hist = len(lista_h)
        cant_anexo = len(mapa_anexo.get(key, []))
        
        if cant_hist > cant_anexo:
            for reg_h in lista_h[cant_anexo:]:
                omitidos_list.append({
                    'Deudor_SAP_OK': key[0],
                    'Nombre_Deudor_OK': key[2],
                    'Fecha_DT': reg_h['Fecha_DT'],
                    'Fecha_UI': reg_h['Fecha_UI'],
                    'Valor_Faltante_Num': reg_h['Valor_Faltante_Num'],
                    'Abonos_Num': reg_h['Abonos_Num'],
                    'Saldo_Num': reg_h['Saldo_Num'],
                    'Motivo_OK': reg_h['Motivo_OK']
                })
        elif key not in mapa_anexo:
            for reg_h in lista_h:
                omitidos_list.append({
                    'Deudor_SAP_OK': key[0],
                    'Nombre_Deudor_OK': key[2],
                    'Fecha_DT': reg_h['Fecha_DT'],
                    'Fecha_UI': reg_h['Fecha_UI'],
                    'Valor_Faltante_Num': reg_h['Valor_Faltante_Num'],
                    'Abonos_Num': reg_h['Abonos_Num'],
                    'Saldo_Num': reg_h['Saldo_Num'],
                    'Motivo_OK': reg_h['Motivo_OK']
                })

    omitidos = pd.DataFrame(omitidos_list) if omitidos_list else pd.DataFrame(columns=df_anexo.columns)

    # 4. PROCESAR SALDADOS Y ACTIVOS DEL ANEXO
    saldados_list = []
    activos_anexo = []

    for _, row in df_anexo.iterrows():
        monto = row['Valor_Faltante_Num']
        abono = row['Abonos_Num']
        saldo = row['Saldo_Num']
        
        if saldo <= 0.01 or (abono >= monto and monto > 0):
            row_saldado = row.copy()
            row_saldado['Abonos_Num'] = monto
            row_saldado['Saldo_Num'] = 0.0
            saldados_list.append(row_saldado)
        else:
            activos_anexo.append(row)

    saldados = pd.DataFrame(saldados_list) if saldados_list else pd.DataFrame()

    nuevos_list = []
    pendientes_list = []

    for row in activos_anexo:
        f_dt = row['Fecha_DT']
        sap_r = str(row['Deudor_SAP_OK'])
        f_str_r = f_dt.strftime('%Y-%m-%d') if pd.notna(f_dt) else ''
        nombre_r = str(row['Nombre_Deudor_OK'])
        key = (sap_r, f_str_r, nombre_r)
        
        v_inicial = row['Valor_Faltante_Num']
        abono_calc = row['Abonos_Num']
        saldo_actual = row['Saldo_Num']
        
        if key in mapa_hist and len(mapa_hist[key]) > 0:
            reg_h = mapa_hist[key][0]
            if reg_h['Valor_Faltante_Num'] > saldo_actual:
                v_inicial = reg_h['Valor_Faltante_Num']
                abono_calc = v_inicial - saldo_actual

        row_final = row.copy()
        row_final['Valor_Faltante_Num'] = v_inicial
        row_final['Abonos_Num'] = round(abono_calc, 2)
        row_final['Saldo_Num'] = round(saldo_actual, 2)

        if pd.notna(f_dt):
            f_dt_norm = pd.to_datetime(f_dt.date())
            f_corte_norm = pd.to_datetime(fecha_dia_actual.date())
            
            if f_dt_norm < f_corte_norm:
                pendientes_list.append(row_final)
            else:
                nuevos_list.append(row_final)
        else:
            nuevos_list.append(row_final)

    nuevos = pd.DataFrame(nuevos_list) if nuevos_list else pd.DataFrame(columns=df_anexo.columns)
    pendientes = pd.DataFrame(pendientes_list) if pendientes_list else pd.DataFrame(columns=df_anexo.columns)
            
    def preparar_df_limpio(df_in):
        if df_in is None or df_in.empty:
            return pd.DataFrame(columns=['Nombre Deudor', 'SAP', 'Fecha', 'Valor Faltante Inicial ($)', 'Abonos Realizados ($)', 'Saldo Pendiente Actual ($)', 'Motivo'])
        
        out = pd.DataFrame()
        out['Nombre Deudor'] = df_in['Nombre_Deudor_OK']
        out['SAP'] = df_in['Deudor_SAP_OK']
        out['Fecha'] = df_in['Fecha_UI']
        out['Valor Faltante Inicial ($)'] = df_in['Valor_Faltante_Num'].apply(lambda x: f"${int(round(x)):,}")
        out['Abonos Realizados ($)'] = df_in['Abonos_Num'].apply(lambda x: f"${int(round(x)):,}")
        out['Saldo Pendiente Actual ($)'] = df_in['Saldo_Num'].apply(lambda x: f"${int(round(x)):,}")
        out['Motivo'] = df_in['Motivo_OK'] if 'Motivo_OK' in df_in.columns else ''
        return out
        
    df_nuevos_ui = preparar_df_limpio(nuevos)
    df_saldados_ui = preparar_df_limpio(saldados)
    df_pendientes_ui = preparar_df_limpio(pendientes)
    df_omitidos_ui = preparar_df_limpio(omitidos)
    
    tot_nuevos = nuevos['Saldo_Num'].sum() if not nuevos.empty else 0
    tot_saldados = saldados['Abonos_Num'].sum() if not saldados.empty else 0
    tot_pendientes = pendientes['Saldo_Num'].sum() if not pendientes.empty else 0
    tot_omitidos = omitidos['Saldo_Num'].sum() if not omitidos.empty else 0
    
    m_str = f"## 🥤 **AUDITORÍA DE FALTANTES - JAMUNDÍ**\n\n"
    m_str += f"- 🔴 **Nuevos Faltantes (Día Actual - {formatear_fecha_ui(fecha_dia_actual)}):** {len(nuevos)} registro(s) | **${int(tot_nuevos):,}**\n"
    m_str += f"- 🟢 **Faltantes Saldados / Pagados:** {len(saldados)} registro(s) | **${int(tot_saldados):,}**\n"
    m_str += f"- 🟣 **Pendientes Activos (Días Anteriores):** {len(pendientes)} registro(s) | **${int(tot_pendientes):,}**\n"
    m_str += f"- ⚠️ **Omitidos en Anexo:** {len(omitidos)} registro(s) | **${int(tot_omitidos):,}**\n"
    
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
st.markdown("# 🥤 **POSTOBÓN JAMUNDÍ - SISTEMA DE AUDITORÍA DE FALTANTES**")

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
        with st.spinner("Procesando auditoría y cazando trampas..."):
            m_str, fig_pie, fig_bar, df_nuevos, df_saldados, df_pendientes, df_omitidos = procesar_archivos(file_hist, file_anexo)
            
        st.markdown("---")
        st.markdown(m_str)
        
        gcol1, gcol2 = st.columns(2)
        with gcol1:
            st.plotly_chart(fig_pie, use_container_width=True)
        with gcol2:
            st.plotly_chart(fig_bar, use_container_width=True)
            
        st.markdown("## 📄 **Detalle de Registros (Listos para Copiar y Pegar)**")
        
        tab1, tab2, tab3, tab4 = st.tabs(["🔴 Nuevos Faltantes", "🟢 Faltantes Saldados / Pagados", "🟣 Pendientes Activos", "⚠️ Omitidos en Anexo"])
        
        with tab1:
            st.dataframe(df_nuevos, use_container_width=True)
        with tab2:
            if df_saldados.empty:
                st.info("No se detectaron faltantes saldados/pagados en este corte.")
            else:
                st.dataframe(df_saldados, use_container_width=True)
        with tab3:
            st.dataframe(df_pendientes, use_container_width=True)
        with tab4:
            if df_omitidos.empty:
                st.info("¡Excelente! No hay registros omitidos en este anexo.")
            else:
                st.dataframe(df_omitidos, use_container_width=True)
else:
    st.info("💡 Sube el **Histórico Maestro (.csv)** y el **Anexo del día**, luego presiona **EJECUTAR AUDITORÍA**.")
