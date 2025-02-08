import pandas as pd
import os
import zipfile
import streamlit as st
import shutil

# 📌 Función para cargar datos
def cargar_datos(uploaded_file):
    """Carga el archivo de Excel subido por el usuario."""
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name="Hoja1")
        df.rename(columns={"CANT ": "CANT"}, inplace=True)
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors='coerce', dayfirst=True)
        return df
    return None

# 📌 Función para generar reportes por cajero
def generar_reporte_por_cajero(df, cajero, fecha_inicio, fecha_fin, carpeta_reportes, observacion_final):
    """Genera un reporte individual para un cajero en un rango de fechas con formato mejorado."""
    # Renombrar las columnas en el DataFrame
    renombrar_columnas = {
        "DATOS_PLANILLA": "DATOS PLANI",
        "CANT": "CANTNC",
        "NC": "MONTONC",
        "CANTIA_NUL": "CANT.ANUL",
        "MONTO_ANUL": "MONTO.ANUL"
    }
    
    df = df.rename(columns=renombrar_columnas)
    
    df_filtrado = df[(df["CAJERO"] == cajero) & (df["FECHA"] >= fecha_inicio) & (df["FECHA"] <= fecha_fin)]
    columnas_reporte = ["FECHA", "FALTANTE", "DATOS PLANI", "OBSERVACIONES", "CANTNC", "MONTONC", "CANT.ANUL", "MONTO.ANUL"]
    
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
    
    # 📌 Filtrar por tienda o sucursal si existe la columna
    if "SUCU" in df.columns:
        sucursales_disponibles = df["SUCU"].dropna().unique().tolist()
        sucursal_seleccionada = st.selectbox("Selecciona la sucursal", ["Todas"] + sucursales_disponibles)
        if sucursal_seleccionada != "Todas":
            df = df[df["SUCU"] == sucursal_seleccionada]
    
    # 📌 Selección de cajeros específicos
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
