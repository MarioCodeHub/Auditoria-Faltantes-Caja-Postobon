# 1. Instalación de librerías
!pip install -q gradio plotly openpyxl pandas

import pandas as pd
import openpyxl
import re
import warnings
import unicodedata
import gradio as gr
import plotly.express as px

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
    if not val_str or val_str.lower() in ['none', 'nan', '']: return 0.0
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


# --- CARGA Y PROCESAMIENTO DEL ARCHIVO ---
def cargar_y_limpiar(file_obj):
    if file_obj is None: return pd.DataFrame()
    filename = file_obj.name
    if filename.endswith('.csv'):
        encodings = ['utf-8-sig', 'latin1', 'cp1252', 'utf-8']
        for enc in encodings:
            try:
                df_raw = pd.read_csv(filename, encoding=enc, sep=None, engine='python', header=None)
                break
            except Exception:
                continue
    else:
        wb = openpyxl.load_workbook(filename, data_only=True)
        sheet = wb.active
        df_raw = pd.DataFrame(list(sheet.values))

    df_raw = df_raw.dropna(how='all').reset_index(drop=True)

    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = [str(c).upper().replace('\n', ' ').strip() for c in row if pd.notna(c)]
        if any("VALOR DEL FALTANTE" in c or "VALOR" in c for c in row_str) and any("SALDO" in c for c in row_str):
            header_idx = idx
            break

    if header_idx is None: header_idx = 5 

    df = df_raw.iloc[header_idx + 1:].copy()
    headers = [str(c).upper().replace('\n', ' ').strip() if pd.notna(c) else f"COL_{i}" for i, c in enumerate(df_raw.iloc[header_idx].values)]
    df.columns = headers

    df = df.loc[:, df.columns.notna()].copy()
    df = df.loc[:, ~df.columns.duplicated(keep='first')].copy()

    col_client_code = next((c for c in df.columns if 'CODIGO' in c and 'CLIENTE' in c), None)
    col_client_name = next((c for c in df.columns if 'NOMBRE' in c and 'CLIENTE' in c), None)
    col_trans_code = next((c for c in df.columns if 'CODIGO' in c and 'TRANSPORTADOR' in c), None)
    col_trans_name = next((c for c in df.columns if 'NOMBRE' in c and 'TRANSPORTADOR' in c), None)
    col_fecha = next((c for c in df.columns if 'FECHA' in c), None)
    col_valor = next((c for c in df.columns if 'VALOR' in c and 'FALTANTE' in c), None)
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


# --- AUDITORÍA Y CLASIFICACIÓN AJUSTADA ---
def procesar_archivos(file_hist, file_anexo):
    if file_anexo is None:
        return "⚠️ Por favor sube el Anexo del día para realizar la auditoría.", None, None, None, None, None

    df_anexo = cargar_y_limpiar(file_anexo)

    # 1. SALDADOS / PAGADOS: Todo lo que tiene Saldo == 0
    saldados = df_anexo[df_anexo['Saldo_Num'] == 0].copy()

    # De los que tienen saldo pendiente (> 0):
    con_saldo = df_anexo[df_anexo['Saldo_Num'] > 0].copy()

    # Identificar el mes y año más reciente del reporte
    if not con_saldo['Fecha_DT'].isna().all():
        max_dt = con_saldo['Fecha_DT'].max()
        mes_actual = max_dt.month
        anio_actual = max_dt.year
    else:
        mes_actual, anio_actual = 10, 2026

    # 2. NUEVOS FALTANTES: 
    # Todos los registros del mes/año actual que tengan saldo pendiente (incluso si tienen abonos parciales)
    es_mes_actual = (con_saldo['Fecha_DT'].dt.month == mes_actual) & (con_saldo['Fecha_DT'].dt.year == anio_actual)
    nuevos = con_saldo[es_mes_actual].copy()

    # 3. PENDIENTES ACTIVOS: 
    # Únicamente los faltantes que vienen arrastrados de meses pasados
    pendientes = con_saldo.drop(nuevos.index).copy()

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

    tot_nuevos = nuevos['Saldo_Num'].sum() if not nuevos.empty else 0
    tot_saldados = saldados['Abonos_Num'].sum() if not saldados.empty else 0
    tot_pendientes = pendientes['Saldo_Num'].sum() if not pendientes.empty else 0

    m_str = f"## 📊 **AUDITORÍA DE FALTANTES - POSTOBÓN JAMUNDÍ**\n\n"
    m_str += f"- 🔴 **Nuevos Faltantes (Mes Actual con/sin abonos):** {len(nuevos)} registro(s) | **${int(tot_nuevos):,}**\n"
    m_str += f"- 🟢 **Faltantes Saldados / Pagados:** {len(saldados)} registro(s) | **${int(tot_saldados):,}**\n"
    m_str += f"- ⚪ **Pendientes Activos (Meses Anteriores):** {len(pendientes)} registro(s) | **${int(tot_pendientes):,}**"

    df_summary = pd.DataFrame({
        'Categoría': ['Nuevos Faltantes', 'Faltantes Saldados', 'Pendientes Activos'],
        'Cantidad': [len(nuevos), len(saldados), len(pendientes)],
        'Monto': [tot_nuevos, tot_saldados, tot_pendientes]
    })

    fig_pie = px.pie(
        df_summary, values='Cantidad', names='Categoría',
        title="<b>Distribución por Cantidad de Casos</b>",
        color='Categoría',
        color_discrete_map={'Nuevos Faltantes': '#EF553B', 'Faltantes Saldados': '#00CC96', 'Pendientes Activos': '#AB63FA'},
        hole=0.4
    )

    fig_bar = px.bar(
        df_summary, x='Categoría', y='Monto', text_auto=',.0f',
        title="<b>Impacto Financiero Real ($ COP)</b>",
        color='Categoría',
        color_discrete_map={'Nuevos Faltantes': '#EF553B', 'Faltantes Saldados': '#00CC96', 'Pendientes Activos': '#AB63FA'}
    )
    fig_bar.update_layout(showlegend=False, yaxis_title="Monto ($)")

    return m_str, fig_pie, fig_bar, df_nuevos_ui, df_saldados_ui, df_pendientes_ui


# --- INTERFAZ GRADIO ---
with gr.Blocks(theme=gr.themes.Soft(primary_hue="emerald")) as app:
    gr.Markdown("# 🥤 **POSTOBÓN - SISTEMA DE AUDITORÍA DE FALTANTES EN CAJA**")
    
    with gr.Row():
        with gr.Column(scale=1):
            file_hist = gr.File(label="1. Histórico Maestro (Opcional)", file_types=['.csv', '.xlsx', '.xlsm'])
            file_anexo = gr.File(label="2. Anexo del Día Actual (.xlsm / .xlsx)", file_types=['.xlsm', '.xlsx', '.csv'])
            btn_run = gr.Button("🚀 EJECUTAR AUDITORÍA", variant="primary")
        
        with gr.Column(scale=2):
            output_msg = gr.Markdown("*Sube los archivos y presiona Ejecutar.*")
            
    with gr.Row():
        plot_pie = gr.Plot(label="Resumen por Casos")
        plot_bar = gr.Plot(label="Impacto Monetario")

    gr.Markdown("## 📋 **Detalle Completo de Registros**")
    
    with gr.Tabs():
        with gr.TabItem("🔴 Nuevos Faltantes"):
            df_out_nuevos = gr.Dataframe(interactive=False)
        with gr.TabItem("🟢 Faltantes Saldados / Pagados"):
            df_out_saldados = gr.Dataframe(interactive=False)
        with gr.TabItem("⚪ Pendientes Activos"):
            df_out_pendientes = gr.Dataframe(interactive=False)

    btn_run.click(
        fn=procesar_archivos,
        inputs=[file_hist, file_anexo],
        outputs=[output_msg, plot_pie, plot_bar, df_out_nuevos, df_out_saldados, df_out_pendientes]
    )

app.launch(share=True, debug=False)
