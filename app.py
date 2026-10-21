# --- MOTOR DE AUDITORÍA JAMUNDÍ INFALIBLE (SIN DEPENDER DE NOMBRES) ---
def procesar_archivos(file_hist, file_anexo):
    if file_anexo is None:
        return "⚠️ Por favor sube el Anexo del día para realizar la auditoría.", None, None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_anexo = cargar_anexo_excel(file_anexo)
    df_hist = cargar_historico_csv(file_hist) if file_hist is not None else pd.DataFrame()
    
    if df_anexo.empty:
        return "⚠️ El anexo cargado no tiene registros válidos.", None, None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    fechas_validas = df_anexo['Fecha_DT'].dropna()
    fecha_dia_actual = fechas_validas.max() if not fechas_validas.empty else pd.Timestamp('2026-10-20')

    # 1. CATALOGAR HISTÓRICO PENDIENTE CON LLAVE PURA (SAP + FECHA)
    mapa_hist = {}
    if not df_hist.empty:
        for idx, r in df_hist.iterrows():
            if r['Saldo_Num'] > 0.0:
                f_str = r['Fecha_DT'].strftime('%Y-%m-%d') if pd.notna(r['Fecha_DT']) else ''
                # Llave ciega a nombres: solo importa el SAP y la Fecha del faltante
                key = (str(r['Deudor_SAP_OK']), f_str)
                if key not in mapa_hist:
                    mapa_hist[key] = []
                mapa_hist[key].append(r)

    # 2. CATALOGAR ANEXO DE HOY CON LA MISMA LLAVE
    mapa_anexo = {}
    for idx, r in df_anexo.iterrows():
        f_dt = r['Fecha_DT']
        f_str = f_dt.strftime('%Y-%m-%d') if pd.notna(f_dt) else ''
        key = (str(r['Deudor_SAP_OK']), f_str)
        if key not in mapa_anexo:
            mapa_anexo[key] = []
        mapa_anexo[key].append(r)

    # 3. CAZAR OMITIDOS REALES (Estaban en histórico con saldo y NO están en el anexo de hoy)
    omitidos_list = []
    for key, lista_h in mapa_hist.items():
        cant_hist = len(lista_h)
        cant_anexo = len(mapa_anexo.get(key, []))
        
        if cant_hist > cant_anexo:
            for reg_h in lista_h[cant_anexo:]:
                omitidos_list.append({
                    'Deudor_SAP_OK': reg_h['Deudor_SAP_OK'],
                    'Nombre_Deudor_OK': reg_h['Nombre_Deudor_OK'],
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
                    'Deudor_SAP_OK': reg_h['Deudor_SAP_OK'],
                    'Nombre_Deudor_OK': reg_h['Nombre_Deudor_OK'],
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
        key = (sap_r, f_str_r)
        
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
