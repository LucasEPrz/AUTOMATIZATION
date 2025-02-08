import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import io

# Subir archivo Excel
archivo_subido = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

df = archivo_subido

if archivo_subido:  # ✅ Asegura que todo lo que usa 'df' está dentro de este bloque
    # Guardar el archivo en el servidor temporalmente
    with open("archivo_temporal.xlsx", "wb") as f:
        f.write(archivo_subido.getbuffer())

    # Leer el archivo desde el sistema de archivos
    df = pd.read_excel("archivo_temporal.xlsx", engine="openpyxl")

    # Mostrar los primeros registros
    st.write("Vista previa de los datos:")
    st.dataframe(df.head())

    # Renombrar las columnas con espacios y nombres incorrectos
    df.rename(columns={
        'Monto               ': 'Monto1',
        'Monto               .1': 'Monto2',
        'Monto               .2': 'Monto3'
    }, inplace=True)

    # Eliminar espacios en blanco de los nombres de las columnas
    df.columns = df.columns.str.strip()

    # Crear las columnas necesarias
    df['Faltante'] = df['Diferencia Total'].apply(lambda x: x if x > 0 else 0)
    df['Sobrante'] = df['Diferencia Total'].apply(lambda x: -x if x < 0 else 0)
    df['N. Credito'] = df['N. Credito'].abs()

    # Convertir columnas de decimales a enteros
    df['Cant.N1'] = df['Cant.N1'].fillna(0).astype(int)
    df['Cant.N2'] = df['Cant.N2'].fillna(0).astype(int)
    df['CantN3'] = df['CantN3'].fillna(0).astype(int)

    # Sumar las columnas monto1, monto2, monto3
    df['Total_Monto'] = df['Monto1'] + df['Monto2'] + df['Monto3']
    df['Total_Cantidad'] = df['Cant.N1'] + df['Cant.N2'] + df['CantN3']

    # Concatenar valores en "Observaciones"
    df['Observaciones'] = df['Nove_1'].fillna('').astype(str) + ' ' + df['Cant.N1'].astype(str) + '; ' + \
                          df['Nove_2'].fillna('').astype(str) + ' ' + df['Cant.N2'].astype(str) + '; ' + \
                          df['Nove_3'].fillna('').astype(str) + ' ' + df['CantN3'].astype(str) + ';'

    # Aplicar condiciones a 'Observaciones'
    df['Observaciones'] += np.where(df['Cant.de Correccion Sobrante'] > 0,
                                    " Hizo corrección SOBRANTE de caja por: " + df['SumaCorreccion'].astype(str) + ";", "")
    
    df['Observaciones'] += np.where(df['Cant de Correccion Faltantes'] > 0,
                                    " Hizo corrección de FALTANTE de caja por: " + df['SumaCorreccion'].astype(str) + ";", "")
    
    df['Observaciones'] += np.where(df['MERCADO PAGO QR'] > 0,
                                    " Hizo COBROS CON MP16 por: " + df['MERCADO PAGO QR'].astype(str) + ";", "")
    
    df['Observaciones'] += np.where(df['VOUCHER'] > 0,
                                    " Hizo USO DEL MEDIO VOUCHER por: " + df['VOUCHER'].astype(str) + ";", "")

    # Crear DataFrame 'Cajas'
    Cajas = df[['Nombre', 'POS', 'Fh_Caja', 'Faltante', 'Sobrante', 'Cantidad Tiquets',
                'Observaciones', 'Cant NC', 'N. Credito', 'Total_Monto', 'Total_Cantidad',
                'ID. Documento', 'Promed.Tiqueo', 'Items', 'Promed Cobro',
                'Cantidad de Correccion Caja', 'Debito ant', 'Credito ant', 'Tarjeta Naranja',
                'MERCADO PAGO QR']].copy()

    # Agregar columnas adicionales
    Cajas['Tarjeta'] = ''
    Cajas['Suc'] = ''
    Cajas['Datos Planilla'] = ''
    Cajas['MP16'] = ''
    Cajas['Estado'] = ''

    # Generar las columnas "MET_ANT" y "MP16"
    Cajas['MET_ANT'] = np.where(
        (Cajas['Debito ant'] + Cajas['Credito ant'] + Cajas['Tarjeta Naranja']) > 0, "SI", "NO"
    )
    
    Cajas['MP16'] = np.where(Cajas['MERCADO PAGO QR'] > 0, "SI", "NO")

    # Eliminar columnas innecesarias
    Cajas.drop(columns=['Debito ant', 'Credito ant', 'Tarjeta Naranja', 'MERCADO PAGO QR'], inplace=True)

    # Reordenar las columnas en el DataFrame final
    Cajas = Cajas[['ID. Documento', 'Nombre', 'Tarjeta', 'Estado', 'POS', 'Suc', 'Fh_Caja',
                   'Faltante', 'Sobrante', 'Datos Planilla', 'Cantidad Tiquets', 'Promed.Tiqueo', 'Items', 'Promed Cobro',
                   'Cantidad de Correccion Caja', 'MP16', 'MET_ANT', 'Observaciones', 'Cant NC',
                   'N. Credito', 'Total_Cantidad', 'Total_Monto']]

    # Definir la columna 'Suc' con base en POS
    conditions = [
        Cajas['POS'].isin([34, 32, 29]),
        Cajas['POS'].isin([7, 13, 8]),
        Cajas['POS'].isin([110, 111, 112]),
        Cajas['POS'].isin([40, 36, 38]),
        Cajas['POS'].isin([46, 44]),
        Cajas['POS'].isin([50, 52, 51])
    ]

    values = [1, 2, 3, 4, 5, 6]

    Cajas['Suc'] = np.select(conditions, values, default=np.nan)

    # Convertir el DataFrame a un archivo de Excel en memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        Cajas.to_excel(writer, index=False, sheet_name='Datos_Procesados')
    output.seek(0)

    # Botón para descargar el archivo procesado
    st.download_button(
        label="Descargar archivo procesado",
        data=output,
        file_name="Cajas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )