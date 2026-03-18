import streamlit as st
import pandas as pd
import re
import io



# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Limpiador de Finiquitos", layout="wide")
st.title("🚜 Procesador de Finiquitos de Obra de SICOM")
st.write("Sube tu archivo .xlsm o .xlsx para limpiar los datos y aplicar formato automáticamente.")

def limpiador_sicom(archivo_entrada,filas_a_saltar):

    #--------------1. LECTURA DEL ARCHIVO ---------------------------------------------------

    df = pd.read_excel(archivo_entrada,skiprows=filas_a_saltar)
    #Exploración del dataframe
    #df.head(5)
    #-INICIO DE CÓDIGO DE LIMPIEZA--------------------------------------------------------------------#
    #--------BORRADO DE FILAS VACÍAS --------------------------------------------#
    df_finiquito_auditoria = df[df['CONCEPTO'].notna() &
        (df['CONCEPTO']!= 'N/A')
    ].copy()
    #------CONVERTIR COLUMNAS DE CANTIDAD Y DE IMPORTE A NÚMERO-----------------#
    df_finiquito_auditoria = df_finiquito_auditoria.rename(columns={
        'PROYECTO': 'CANTIDAD PROYECTO',
        'REAL': 'CANTIDAD REAL',
        'ADITIVAS':'CANTIDAD ADITIVAS',
        'DIFERENCIA': 'DIFERENCIA CANTIDAD',

        'PROYECTO.1': 'IMPORTE PROYECTO',
        'ADITIVAS.1': 'IMPORTE ADITIVAS',
        'REAL.1': 'IMPORTE REAL',
        'DIFERENCIA.1': 'DIFERENCIA IMPORTE'
      })
    #-------CONVERTIR COLUMNAS VOLEST_ A NÚMERO--------------------------------#
    # Identificar las columnas de estimaciones usando una expresión regular
    # Buscamos columnas que contengan 'VOLEST_' o las nuevas columnas de cantidad
    patron_columnas = r'VOLEST_?|CANTIDAD PROYECTO|CANTIDAD ADITIVAS|CANTIDAD REAL|DIFERENCIA CANTIDAD'
    columnas_cantidad_estimacion = [col for col in df_finiquito_auditoria.columns if re.search(patron_columnas, col)]

    #--------Aplicar la limpieza a cada una de esas columnas-------------------#
    for col in columnas_cantidad_estimacion:
        # Aseguramos que tratamos la columna como texto, quitamos espacios y hacemos el reemplazo
        df_finiquito_auditoria[col] = (
            df_finiquito_auditoria[col]
            .astype(str) # Convierte todos los valores a tipo string
            .str.strip() #String Method que quita espacios
            .str.replace(r'(.*)-$', r'-\1', regex=True) #Cambia el último carácter "-" al principio
        )

    #-----Convertir a numérico (los valores que no se puedan convertir serán NaN)
        df_finiquito_auditoria[col] = pd.to_numeric(df_finiquito_auditoria[col], errors='coerce') #coerce hace que los valores que no se puedan covertir sean NaN

    # Opcional: Llenar los NaN con 0 si es necesario para tus cálculos
    # df_finiquito_auditoria[columnas_cantidad_estimacion] = df_finiquito_auditoria[columnas_cantidad_estimacion].fillna(0)

    print(f"Se procesaron {len(columnas_cantidad_estimacion)} columnas de cantidad.")

    #---------CONVERTIR COLUMNAS IMPEST_ A NÚMERO--------------------------------
    # Identificar las columnas de estimaciones usando una expresión regular
    # Buscamos columnas que contengan 'IMPEST_' o las nuevas columnas de importe
    patron_columnas = r'IMPEST_?|PRECIO UNITARIO|IMPORTE PROYECTO|IMPORTE ADITIVAS|IMPORTE REAL|DIFERENCIA IMPORTE'
    columnas_importe_estimacion = [col for col in df_finiquito_auditoria.columns if re.search(patron_columnas, col)]
    # Aplicar la limpieza a cada una de esas columnas
    for col in columnas_importe_estimacion:
        # Aseguramos que tratamos la columna como texto, quitamos espacios y hacemos el reemplazo
        cleaned_string_values = (
            df_finiquito_auditoria[col]
            .astype(str) # Convierte todos los valores a tipo string
            .str.strip() #String Method que quita espacios
            .str.replace('$', '', regex=False)
            .str.replace(',', '', regex=False)
        )

        # Convertir a numérico (los valores que no se puedan convertir serán NaN)
        df_finiquito_auditoria[col] = pd.to_numeric(cleaned_string_values, errors='coerce')
        # Llenar los NaN con 0 si es necesario para tus cálculos
        df_finiquito_auditoria[col] = df_finiquito_auditoria[col].fillna(0)

    print(f"Se procesaron {len(columnas_importe_estimacion)} columnas de importe")
    #---------------------FIN DEL CÓDIGO DE LIMPIEZA-----------------------------------------------------#

    #---------------------INICIO DE EXPORTAR A EXCEL CON FORMATO-----------------------------------------#

    # Crear un buffer en memoria
    output = io.BytesIO()

    # Crear un objeto ExcelWriter usando xlsxwriter engine, escribiendo al buffer
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    # Escribir el DataFrame al archivo Excel
    df_finiquito_auditoria.to_excel(writer, sheet_name='Finiquito', index=False)

    # Acceder al objeto workbook y worksheet de xlsxwriter
    workbook  = writer.book
    worksheet = writer.sheets['Finiquito']

    # Establecer un ancho de columna predeterminado (por ejemplo, 20 píxeles)
    # Puedes ajustar este valor según la legibilidad deseada
    for i, col in enumerate(df_finiquito_auditoria.columns):
        # Ajustar el ancho de columna automáticamente basado en el contenido
        # o establecer un ancho fijo. Aquí se establece un ancho fijo de 20.
        worksheet.set_column(i, i, 15)

    # Establecer un alto de fila predeterminado (por ejemplo, 20 píxeles)
    # Puedes ajustar este valor según la legibilidad deseada
    worksheet.set_default_row(15)

    # Cerrar el objeto ExcelWriter para guardar el archivo en el buffer
    writer.close()

    print("DataFrame exportado a 'finiquito_limpio_final_formato.xlsx' con formato exitosamente.")
    #----------------------------FIN DE EXPORTACIÓN A EXCEL CON FORMATO---------------------------------#
    return output.getvalue()

    # --- INTERFAZ DE USUARIO ---
archivo = st.file_uploader("Selecciona el archivo de finiquito", type=["xlsx", "xlsm"])
filas_a_saltar = st.number_input("Número de filas a saltar al inicio del archivo:", min_value=0, value=0)

if archivo:
    if st.button("🚀 Iniciar Limpieza"):
        try:
            resultado_binario = limpiador_sicom(archivo, filas_a_saltar)

            st.success("¡Limpieza terminada con éxito!")

            st.download_button(
                label="📥 Descargar Finiquito Limpio",
                data=resultado_binario,
                file_name="finiquito_limpio_final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Hubo un error al procesar el archivo: {e}")
