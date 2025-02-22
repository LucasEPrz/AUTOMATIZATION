import streamlit as st
import pandas as pd
import numpy as np
import io

st.title("Procesamiento de Datos - Cajas")
st.write("Sube el archivo Excel para procesar los datos.")

# Cargar archivo Excel con el widget de Streamlit
uploaded_file = st.file_uploader("Selecciona un archivo Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Leer la hoja n°1 (índice 0) y saltar las primeras 5 filas (la fila 6 se toma como encabezado)
    df = pd.read_excel(uploaded_file, sheet_name=0, skiprows=5, header=0)

    # Renombrar columnas de monto
    renombres = {
        'Monto               ': 'Monto1',
        'Monto               .1': 'Monto2',
        'Monto               .2': 'Monto3'
    }
    df = df.rename(columns=renombres)
    
    # Eliminar la primera columna por nombre ("Suma.               ")
    df = df.drop(columns=['Suma.               '])
    
    # Eliminar espacios en blanco de los encabezados
    df.columns = df.columns.str.strip()

    # Crear las columnas "Faltante" y "Sobrante" a partir de "Total Diferencia"
    df['Faltante'] = df['Total Diferencia'].apply(lambda x: x if x > 0 else 0)
    df['Sobrante'] = df['Total Diferencia'].apply(lambda x: -x if x < 0 else 0)

    # Convertir valores negativos en "N. Credito" a positivos
    df['N. Credito'] = df['N. Credito'].abs()

    # Convertir columnas de cantidades a enteros
    df['Cant.N1'] = df['Cant.N1'].fillna(0).astype(int)
    df['Cant.N2'] = df['Cant.N2'].fillna(0).astype(int)
    df['CantN3'] = df['CantN3'].fillna(0).astype(int)

    # Sumar las columnas de montos y cantidades
    df['Total_Monto'] = df['Monto1'] + df['Monto2'] + df['Monto3']
    df['Total_Cantidad'] = df['Cant.N1'] + df['Cant.N2'] + df['CantN3']

    # Concatenar valores en "Observaciones"
    df['Observaciones'] = (
        df['Nove_1'].fillna('').astype(str) + ' ' + df['Cant.N1'].astype(str) + '; ' +
        df['Nove_2'].fillna('').astype(str) + ' ' + df['Cant.N2'].astype(str) + '; ' +
        df['Nove_3'].fillna('').astype(str) + ' ' + df['CantN3'].astype(str) + ';'
    )

    # Concatenación condicional para correcciones de caja
    df['Observaciones'] += np.where(
        df['Cant.de Correccion Sobrante'] > 0,
        " Hizo corrección SOBRANTE de caja por: " + df['Suma Correcciones'].astype(str) + ";",
        ""
    )
    df['Observaciones'] += np.where(
        df['Cant de Correccion Faltantes'] > 0,
        " Hizo corrección de FALTANTE de caja por: " + df['Suma Correcciones'].astype(str) + ";",
        ""
    )

    # Concatenación condicional para cobros con Mercado Pago y uso de Voucher
    df['Observaciones'] += np.where(
        df['Mercado Pago QR'] > 0,
        " Hizo COBROS CON MP16 por: " + df['Mercado Pago QR'].astype(str) + ";",
        ""
    )
    df['Observaciones'] += np.where(
        df['Voucher'] > 0,
        " Hizo USO DEL MEDIO VOUCHER por: " + df['Voucher'].astype(str) + ";",
        ""
    )

    # Crear el DataFrame final "Cajas" con las columnas necesarias
    Cajas = df[['Nombre', 'POS', 'Caja', 'Faltante', 'Sobrante', 'Cantidad Tiquets',
                'Observaciones', 'Cant NC', 'N. Credito', 'Total_Monto', 'Total_Cantidad',
                'ID. Documento', 'Promed.Tiqueo', 'Items', 'Promed Cobro',
                'Cant de Correccion Caja', 'Debito ANT', 'Credito ANT', 'Tarjeta Naranja ANT',
                'Mercado Pago QR']].copy()

    # Agregar columnas adicionales
    Cajas['Tarjeta'] = ''
    Cajas['Suc'] = ''
    Cajas['Datos Planilla'] = ''
    Cajas['MP16'] = ''
    Cajas['Estado'] = ''

    # Condición 1: Generar la columna "MET_ANT"
    Cajas['MET_ANT'] = np.where(
        (Cajas['Debito ANT'] + Cajas['Credito ANT'] + Cajas['Tarjeta Naranja ANT']) > 0,
        "SI",
        "NO"
    )

    # Condición 3: Agregar observación si se usó método antiguo
    Cajas['Observaciones'] += np.where(
        Cajas['MET_ANT'] == "SI",
        " SE HIZO USO DEL METODO ANTIGUO;",
        ""
    )

    # Condición 2: Generar la columna "MP16"
    Cajas['MP16'] = np.where(
        Cajas['Mercado Pago QR'] > 0,
        "SI",
        "NO"
    )

    # Agregar condicional para Faltante mayor a $1000
    Cajas['Observaciones'] += np.where(
        Cajas['Faltante'] > 1000,
        " FALTANTE por " + Cajas['Faltante'].astype(str) + ";",
        ""
    )

    # Agregar condicional para Sobrante mayor a $1000
    Cajas['Observaciones'] += np.where(
        Cajas['Sobrante'] > 1000,
        " SOBRANTE por " + Cajas['Sobrante'].astype(str) + ";",
        ""
    )

    # Eliminar las columnas usadas en el cálculo para no duplicar información
    Cajas.drop(columns=['Debito ANT', 'Credito ANT', 'Tarjeta Naranja ANT', 'Mercado Pago QR'], inplace=True)

    # Reordenar las columnas en el DataFrame final
    Cajas = Cajas[['ID. Documento', 'Nombre', 'Tarjeta', 'Estado', 'Datos Planilla', 'POS', 'Suc', 'Caja',
                   'Faltante', 'Sobrante', 'Cantidad Tiquets', 'Promed.Tiqueo', 'Items', 'Promed Cobro',
                   'Cant de Correccion Caja', 'MP16', 'MET_ANT', 'Observaciones', 'Cant NC',
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

    st.write("Vista previa de los datos procesados:")
    st.dataframe(Cajas.head())

    # Convertir el DataFrame final a un archivo Excel en memoria
    towrite = io.BytesIO()
    with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
        Cajas.to_excel(writer, index=False, sheet_name='Cajas')
    towrite.seek(0)

    # Botón de descarga
    st.download_button(
        label="Descargar Excel procesado",
        data=towrite,
        file_name="Cajas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
