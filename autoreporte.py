import pandas as pd
import os
import zipfile
import streamlit as st
import shutil

# 📌 Función para cargar datos
def cargar_datos(uploaded_file):
    """Carga el archivo de Excel subido por el usuario."""
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)  # Cargamos el DataFrame
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors='coerce', dayfirst=True)
        return df
    return None


# 📌 Función para generar reportes por cajero
def generar_reporte_por_cajero(df, cajero, fecha_inicio, fecha_fin, carpeta_reportes, observacion_final):
    """Genera un reporte individual para un cajero en un rango de fechas con formato mejorado."""
    
    # Se suma un día a la fecha_fin para incluir el último día completo
    fecha_fin_incl = fecha_fin + pd.Timedelta(days=1)
    df_filtrado = df[(df["CAJERO"] == cajero) & (df["FECHA"] >= fecha_inicio) & (df["FECHA"] < fecha_fin_incl)]
    columnas_reporte = ["FECHA", "FALTANTE", "SOBRANTE", "CANT.TK", "DATOSPLANI", "OBSERVACIONES", "CANT.NC", "MONTO.NC", "CANTANUL", "MONTOANUL"]
    
    if df_filtrado.empty:
        df_reporte = pd.DataFrame(columns=columnas_reporte)
        df_reporte.loc[0] = ["SIN DATOS"] + [0] * (len(columnas_reporte) - 1)
    else:
        columnas_existentes = [col for col in columnas_reporte if col in df_filtrado.columns]
        df_reporte = df_filtrado[columnas_existentes].copy()
        total_fila = df_reporte.select_dtypes(include=['number']).sum()
        total_fila["FECHA"] = "TOTAL"
        total_fila["OBSERVACIONES"] = observacion_final  # Agregar observación final
        df_reporte = pd.concat([df_reporte, pd.DataFrame(total_fila).T], ignore_index=True)
    
    nombre_salida = f"{carpeta_reportes}/Reporte_{cajero.replace(' ', '_')}_{fecha_inicio.date()}_al_{fecha_fin.date()}.xlsx"
    with pd.ExcelWriter(nombre_salida, engine='xlsxwriter') as writer:
        df_reporte.to_excel(writer, sheet_name="Reporte", index=False, startrow=2, startcol=1)
        workbook = writer.book
        worksheet = writer.sheets["Reporte"]
        
        # 📌 Aplicar formato al encabezado
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 2, 'bg_color': '#FFDDC1'})
        worksheet.write('B1', '')  # Dejar celda B1 vacía
        worksheet.merge_range('B2:K2', f"REPORTE DE RENDIMIENTO DE CAJA DE {cajero.upper()} DEL {fecha_inicio.date()} AL {fecha_fin.date()}", header_format)
        worksheet.write_row('B3', columnas_reporte, header_format)
        
        # 📌 Ajustar automáticamente el ancho de las columnas
        for col_num, col_name in enumerate(columnas_reporte, start=1):
            max_len = max(df_reporte[col_name].astype(str).map(len).max(), len(col_name)) + 2
            worksheet.set_column(col_num, col_num, max_len)
    
    print(f"✅ Reporte generado para {cajero}: {nombre_salida}")
    return nombre_salida

# 📌 Aplicación Streamlit
st.title("Generador de Reportes de Cajeros")

uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])
rango_fechas = st.date_input("Selecciona el rango de fechas", [], min_value=None, max_value=None)
nombre_zip = st.text_input("Nombre del archivo ZIP (sin extensión)", "reportes_cajeros")
observacion_final = st.text_area("Observación final para los reportes")

df = cargar_datos(uploaded_file)
if df is not None and len(rango_fechas) == 2:
    st.write("Vista previa de los datos:")
    st.dataframe(df.head())
    
    fecha_inicio, fecha_fin = pd.Timestamp(rango_fechas[0]), pd.Timestamp(rango_fechas[1])
    # Se suma un día a la fecha_fin para incluir el último día en los filtros
    fecha_fin_incl = fecha_fin + pd.Timedelta(days=1)
    
    # 📌 Filtrar por tienda o sucursal si existe la columna
    if "SUCU" in df.columns:
        sucursales_disponibles = df["SUCU"].dropna().unique().tolist()
        sucursal_seleccionada = st.selectbox("Selecciona la sucursal", ["Todas"] + sucursales_disponibles)
        if sucursal_seleccionada != "Todas":
            df = df[df["SUCU"] == sucursal_seleccionada]
    
    # 📌 Crear DataFrame filtrado por el rango de fechas (incluyendo el último día)
    df_periodo = df[(df["FECHA"] >= fecha_inicio) & (df["FECHA"] < fecha_fin_incl)]
    
    # 📌 Top 10 Cajeros con más Faltantes en el periodo seleccionado (ya existente)
    top10_faltantes = (
        df_periodo.groupby("CAJERO")["FALTANTE"]
        .sum()
        .reset_index()
        .sort_values(by="FALTANTE", ascending=False)
        .head(10)
    )
    st.subheader("Top 10 Cajeros con más Faltantes en el periodo seleccionado")
    st.dataframe(top10_faltantes)
    
    # 📌 Top 10 de PROM.TK (Tickeo)
    top10_prom_tk = df_periodo.groupby("CAJERO")["PROM.TK"].mean().reset_index()
    top10_prom_tk_max = top10_prom_tk.sort_values(by="PROM.TK", ascending=False).head(10)
    top10_prom_tk_min = top10_prom_tk.sort_values(by="PROM.TK", ascending=True).head(10)
    st.subheader("Top 10 Cajeros con mayor PROM.TK (Tickeo)")
    st.dataframe(top10_prom_tk_max)
    st.subheader("Top 10 Cajeros con menor PROM.TK (Tickeo)")
    st.dataframe(top10_prom_tk_min)
    
    # 📌 Top 10 de PROM.COB (Cobro)
    top10_prom_cob = df_periodo.groupby("CAJERO")["PROM.COB"].mean().reset_index()
    top10_prom_cob_max = top10_prom_cob.sort_values(by="PROM.COB", ascending=False).head(10)
    top10_prom_cob_min = top10_prom_cob.sort_values(by="PROM.COB", ascending=True).head(10)
    st.subheader("Top 10 Cajeros con mayor PROM.COB (Cobro)")
    st.dataframe(top10_prom_cob_max)
    st.subheader("Top 10 Cajeros con menor PROM.COB (Cobro)")
    st.dataframe(top10_prom_cob_min)
    
    # 📌 Top 10 de promedio de PROM.ITEMS (Items)
    top10_cant_item = df_periodo.groupby("CAJERO")["PROM.ITEMS"].mean().reset_index()
    top10_cant_item_max = top10_cant_item.sort_values(by="PROM.ITEMS", ascending=False).head(10)
    top10_cant_item_min = top10_cant_item.sort_values(by="PROM.ITEMS", ascending=True).head(10)
    st.subheader("Top 10 Cajeros con mayor promedio de PROM.ITEMS")
    st.dataframe(top10_cant_item_max)
    st.subheader("Top 10 Cajeros con menor promedio de PROM.ITEMS")
    st.dataframe(top10_cant_item_min)
    
    # ──────────────────────────────────────────────────────────────
    # 📌 Nuevos cálculos solicitados:
    
    # 1) Promedio de Anulados por ticket = sum(CANTANUL) / sum(CANT.TK)
    agg_anulados = df_periodo.groupby("CAJERO").agg({"CANT.TK": "sum", "CANTANUL": "sum"}).reset_index()
    agg_anulados["Promedio_Anulados_x_Ticket"] = agg_anulados.apply(
        lambda row: row["CANT.TK"] / row["CANTANUL"] if row["CANTANUL"] != 0 else 0, axis=1
    )
    top10_prom_anulados_max = agg_anulados.sort_values(by="Promedio_Anulados_x_Ticket", ascending=False).head(10)
    agg_anulados_nonzero = agg_anulados[agg_anulados["Promedio_Anulados_x_Ticket"] != 0]
    top10_prom_anulados_min = agg_anulados_nonzero.sort_values(by="Promedio_Anulados_x_Ticket", ascending=True).head(10)
    st.subheader("Top 10 Cajeros con mayor Promedio de Anulados por Ticket")
    st.dataframe(top10_prom_anulados_max)
    st.subheader("Top 10 Cajeros con menor Promedio de Anulados por Ticket (excluyendo 0)")
    st.dataframe(top10_prom_anulados_min)
    
    # 2) Promedio de montos de anulados por ticket = sum(MONTOANUL) / sum(CANT.TK)
    agg_monto_anulados = df_periodo.groupby("CAJERO").agg({"MONTOANUL": "sum", "CANT.TK": "sum"}).reset_index()
    agg_monto_anulados["Promedio_Monto_Anulados_x_Ticket"] = agg_monto_anulados.apply(
        lambda row: row["MONTOANUL"] / row["CANT.TK"] if row["CANT.TK"] != 0 else 0, axis=1
    )
    top10_prom_monto_anulados_max = agg_monto_anulados.sort_values(by="Promedio_Monto_Anulados_x_Ticket", ascending=False).head(10)
    agg_monto_anulados_nonzero = agg_monto_anulados[agg_monto_anulados["Promedio_Monto_Anulados_x_Ticket"] != 0]
    top10_prom_monto_anulados_min = agg_monto_anulados_nonzero.sort_values(by="Promedio_Monto_Anulados_x_Ticket", ascending=True).head(10)
    st.subheader("Top 10 Cajeros con mayor Promedio de Montos de Anulados por Ticket")
    st.dataframe(top10_prom_monto_anulados_max)
    st.subheader("Top 10 Cajeros con menor Promedio de Montos de Anulados por Ticket (excluyendo 0)")
    st.dataframe(top10_prom_monto_anulados_min)
    
    # 3) Promedio de notas de crédito por ticket = sum(CANT.NC) / sum(CANT.TK)
    agg_nc = df_periodo.groupby("CAJERO").agg({"CANT.TK": "sum", "CANT.NC": "sum"}).reset_index()
    agg_nc["Promedio_Notas_Credito_x_Ticket"] = agg_nc.apply(
        lambda row: row["CANT.TK"] / row["CANT.NC"] if row["CANT.NC"] != 0 else 0, axis=1
    )
    top10_prom_nc_max = agg_nc.sort_values(by="Promedio_Notas_Credito_x_Ticket", ascending=False).head(10)
    agg_nc_nonzero = agg_nc[agg_nc["Promedio_Notas_Credito_x_Ticket"] != 0]
    top10_prom_nc_min = agg_nc_nonzero.sort_values(by="Promedio_Notas_Credito_x_Ticket", ascending=True).head(10)
    st.subheader("Top 10 Cajeros con mayor Promedio de Notas de Crédito por Ticket")
    st.dataframe(top10_prom_nc_max)
    st.subheader("Top 10 Cajeros con menor Promedio de Notas de Crédito por Ticket (excluyendo 0)")
    st.dataframe(top10_prom_nc_min)
    
    # 4) Promedio de montos de nota de crédito por ticket = sum(MONTO.NC) / sum(CANT.TK)
    agg_monto_nc = df_periodo.groupby("CAJERO").agg({"MONTO.NC": "sum", "CANT.TK": "sum"}).reset_index()
    agg_monto_nc["Promedio_Monto_NC_x_Ticket"] = agg_monto_nc.apply(
        lambda row: row["MONTO.NC"] / row["CANT.TK"] if row["CANT.TK"] != 0 else 0, axis=1
    )
    top10_prom_monto_nc_max = agg_monto_nc.sort_values(by="Promedio_Monto_NC_x_Ticket", ascending=False).head(10)
    agg_monto_nc_nonzero = agg_monto_nc[agg_monto_nc["Promedio_Monto_NC_x_Ticket"] != 0]
    top10_prom_monto_nc_min = agg_monto_nc_nonzero.sort_values(by="Promedio_Monto_NC_x_Ticket", ascending=True).head(10)
    st.subheader("Top 10 Cajeros con mayor Promedio de Montos de Nota de Crédito por Ticket")
    st.dataframe(top10_prom_monto_nc_max)
    st.subheader("Top 10 Cajeros con menor Promedio de Montos de Nota de Crédito por Ticket (excluyendo 0)")
    st.dataframe(top10_prom_monto_nc_min)
    
    # ──────────────────────────────────────────────────────────────
    # 📌 Selección de cajeros específicos para generar reportes
    cajeros_disponibles = df["CAJERO"].dropna().unique().tolist()
    cajeros_seleccionados = st.multiselect("Selecciona los cajeros para generar reportes", cajeros_disponibles, default=cajeros_disponibles)
    
    if st.button("Generar Reportes"):
        carpeta_reportes = "reportes_cajeros"
        
        # 📌 Limpiar carpeta de reportes antes de generar nuevos archivos
        if os.path.exists(carpeta_reportes):
            shutil.rmtree(carpeta_reportes)
        os.makedirs(carpeta_reportes, exist_ok=True)
        
        for cajero in cajeros_seleccionados:
            generar_reporte_por_cajero(df, cajero, fecha_inicio, fecha_fin, carpeta_reportes, observacion_final)
        
        zip_filename = f"{nombre_zip}.zip"
        with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(carpeta_reportes):
                for file in files:
                    zipf.write(os.path.join(root, file), file)
        
        with open(zip_filename, "rb") as f:
            st.download_button("Descargar Reportes", f, file_name=zip_filename)
        
        st.success("✅ Todos los reportes han sido generados y están listos para descargar.")
