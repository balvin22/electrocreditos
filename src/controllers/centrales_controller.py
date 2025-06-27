import pandas as pd

# Asegúrate de que esta ruta sea la correcta para tu archivo de correcciones
ruta_archivo_correcciones = 'c:/Users/usuario/Desktop/Reporte LV/Cédulas a revisar.xlsx'

try:
    print("Leyendo la hoja 'R05' del archivo de correcciones...")
    
    # Leemos solo las columnas que necesitamos
    df_r05 = pd.read_excel(
        ruta_archivo_correcciones, 
        sheet_name='R05',
        usecols=['MCNTIPCRU2', 'MCNNUMCRU2', 'MCNFECHA']
    )
    print("Lectura completada.")

    # --- INICIO DE LA TRANSFORMACIÓN ---
    print("\nIniciando transformación de los datos de la hoja 'R05'...")

    # 1. CONVERSIÓN DE FECHA
    # Convertimos la columna 'MCNFECHA' a un objeto de fecha y luego al formato YYYYMMDD
    # errors='coerce' convertirá cualquier fecha inválida en un valor nulo (NaT)
    df_r05['FECHA_FORMATEADA'] = pd.to_datetime(df_r05['MCNFECHA'], format='%d/%m/%Y', errors='coerce').dt.strftime('%Y%m%d')

    # 2. Preparar columnas para la concatenación
    df_r05['MCNTIPCRU2'] = df_r05['MCNTIPCRU2'].astype(str).str.strip()
    df_r05['MCNNUMCRU2'] = df_r05['MCNNUMCRU2'].astype(str).str.strip()

    # 3. Crear la llave base (ej: 'DF' + '3' = 'DF3')
    df_r05['llave_base'] = df_r05['MCNTIPCRU2'] + df_r05['MCNNUMCRU2']

    # 4. Creamos TRES DataFrames temporales
    df_base = pd.DataFrame({'LLAVE_R05': df_r05['llave_base'], 'FECHA_NUEVA': df_r05['FECHA_FORMATEADA']})
    df_c1 = pd.DataFrame({'LLAVE_R05': df_r05['llave_base'] + 'C1', 'FECHA_NUEVA': df_r05['FECHA_FORMATEADA']})
    df_c2 = pd.DataFrame({'LLAVE_R05': df_r05['llave_base'] + 'C2', 'FECHA_NUEVA': df_r05['FECHA_FORMATEADA']})

    # 5. Unimos los tres DataFrames en la tabla final
    tabla_consulta_r05 = pd.concat([df_base, df_c1, df_c2], ignore_index=True)
    
    # Eliminamos filas donde la fecha no se pudo convertir
    tabla_consulta_r05.dropna(subset=['FECHA_NUEVA'], inplace=True)
    
    # --- FIN DE LA TRANSFORMACIÓN ---

    print("\n¡Transformación completada exitosamente!")
    print("Así se ven las primeras filas de la nueva tabla de consulta 'R05':")
    print(tabla_consulta_r05)

    # Opcional: Guardar esta tabla en un Excel para revisarla
    # tabla_consulta_r05.to_excel("Resultado_R05.xlsx", index=False)

except FileNotFoundError:
    print(f"Error: No se encontró el archivo de correcciones en: {ruta_archivo_correcciones}")
except Exception as e:
    print(f"Ocurrió un error: {e}")