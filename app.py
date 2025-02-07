import subprocess
import sys

# Asegurar la instalación de openpyxl antes de ejecutar el código
def install_packages():
    required_packages = ["openpyxl"]
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Instalando {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

install_packages()

import streamlit as st
import pandas as pd
import numpy as np
import io

# Configuración de la App
st.title("Procesador de Datos de Caja")
st.write("Sube un archivo de Excel para iniciar el procesamiento de datos.")

# Subir archivo Excel
archivo_subido = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if archivo_subido:
    df = pd.read_excel(archivo_subido)
    
    # Renombrar las columnas con espacios y nombres incorrectos
    df.rename(columns={
        'Monto               ': 'Monto1',
        'Monto               .1': 'Monto2',
        'Monto               .2': 'Monto3'
    }, inplace=True)
    
    # Eliminar espacios en blanco de los nombres de las columnas
    df.columns = df.columns.str.strip()
    
    # Crear la columna 'Faltante' para los valores positivos
    df['Faltante'] = df['Diferencia Total'].apply(lambda x: x if x > 0 else 0)
    
    # Crear la columna 'Sobrante' para los valores negativos convertidos a positivos
    df['Sobrante'] = df['Diferencia Total'].apply(lambda x: -x if x < 0 else 0)
    
    # Convertir valores negativos en 'N. Credito' a positivos
    df['N. Credito'] = df['N. Credito'].abs()
    
    # Convertir las columnas de decimales a enteros, manejando NaN si es necesario
    df['Cant.N1'] = df['Cant.N1'].fillna(0).astype(int)
    df['Cant.N2'] = df['Cant.N2'].fillna(0).astype(int)
    df['CantN3'] = df['CantN3'].fillna(0).astype(int)
    
    # Sumar las columnas monto1, monto2, monto3
    df['Total_Monto'] = df['Monto1'] + df['Monto2'] + df['Monto3']
    
    # Sumar las columnas Cant.N1, Cant.N2, CantN3
    df['Total_Cantidad'] = df['Cant.N1'] + df['Cant.N2'] + df['CantN3']
    
    # Concatenar valores en "Observaciones"
    df['Observaciones'] = df['Nove_1'].fillna('').astype(str) + ' ' + df['Cant.N1'].astype(str) + '; ' + \
                          df['Nove_2'].fillna('').astype(str) + ' ' + df['Cant.N2'].astype(str) + '; ' + \
                          df['Nove_3'].fillna('').astype(str) + ' ' + df['CantN3'].astype(str) + ';'
    
    # Concatenación condicional para correcciones de caja
    df['Observaciones'] += np.where(df['Cant.de Correccion Sobrante'] > 0,
                                    " Hizo corrección SOBRANTE de caja por: " + df['SumaCorreccion'].astype(str) + ";", "")
    
    df['Observaciones'] += np.where(df['Cant de Correccion Faltantes'] > 0,
                                    " Hizo corrección de FALTANTE de caja por: " + df['SumaCorreccion'].astype(str) + ";", "")
    
    # Concatenación condicional para cobros con Mercado Pago
    df['Observaciones'] += np.where(df['MERCADO PAGO QR'] > 0,
                                    " Hizo COBROS CON MP16 por: " + df['MERCADO PAGO QR'].astype(str) + ";", "")
    
    # Concatenación condicional para uso de VOUCHER
    df['Observaciones'] += np.where(df['VOUCHER'] > 0,
                                    " Hizo USO DEL MEDIO VOUCHER por: " + df['VOUCHER'].astype(str) + ";", "")
    
    # Seleccionar las columnas necesarias para el DataFrame final
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
    
    # Generar la columna "MET_ANT"
    Cajas['MET_ANT'] = np.where(
        (Cajas['Debito ant'] + Cajas['Credito ant'] + Cajas['Tarjeta Naranja']) > 0, "SI", "NO"
    )
    
    # Generar la columna "MP16"
    Cajas['MP16'] = np.where(Cajas['MERCADO PAGO QR'] > 0, "SI", "NO")
    
    # Eliminar columnas usadas en el cálculo
    Cajas.drop(columns=['Debito ant', 'Credito ant', 'Tarjeta Naranja', 'MERCADO PAGO QR'], inplace=True)

    # Reordenar las columnas en el DataFrame final
    Cajas = Cajas[['ID. Documento', 'Nombre', 'Tarjeta', 'Estado', 'POS', 'Suc', 'Fh_Caja',
                   'Faltante', 'Sobrante', 'Datos Planilla', 'Cantidad Tiquets', 'Promed.Tiqueo', 'Items', 'Promed Cobro',
                   'Cantidad de Correccion Caja', 'MP16', 'MET_ANT', 'Observaciones', 'Cant NC',
                   'N. Credito', 'Total_Cantidad', 'Total_Monto']]
    
    # Agregar columnas adicionales
    Cajas['Tarjeta'] = ''
    Cajas['Suc'] = ''
    Cajas['Datos Planilla'] = ''
    Cajas['MP16'] = ''
    Cajas['Estado'] = ''
    
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
