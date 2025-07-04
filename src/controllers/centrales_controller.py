import pandas as pd
import io
import numpy as np

# --- La configuración de columnas que definiste se mantiene igual ---
configuracion = {
    "ANALISIS": {
        "usecols": ["direccion", "barrio", "nomciudad", "totcuotas", "valorcuota", "diasatras", "cuotaspag","cedula","saldofac","tipo","numero"],
        "rename_map": {
            "direccion": "Direccion",
            "barrio": "Barrio",
            "nomciudad": "Nombre_Ciudad",
            "totcuotas": "Total_Cuotas",
            "valorcuota": "Valor_Cuota",
            "diasatras": "Dias_Atraso",
            "cuotaspag": "Cuotas_Pagadas",
            "cedula" : "Cedula_Cliente",
            "saldofac":"Saldo_Factura",
            "tipo":"Tipo_Credito",
            "numero":"Numero_Credito"
        }
    },
    "R91": {
        "usecols": ["VINNOMBRE", "VENNOMBRE", "MCDZONA", "MCDVINCULA", "MCDNUMCRU1", "MCDTIPCRU1","VENOMBRE","VENCODIGO",
                    "COBNOMBRE", "MCDCCOSTO", "CCONOMBRE", "META_INTER", "META_DC_AL", "META_DC_AT", "META_ATRAS"],
        "rename_map": {
            "MCDTIPCRU1": "Tipo_Credito",
            "MCDNUMCRU1": "Numero_Credito",
            "MCDVINCULA" : "Cedula_Cliente",
            "VINNOMBRE": "Nombre_Cliente",
            "MCDZONA" : "Zona",
            "COBNOMBRE" : "Nombre_Cobrador",
            "VENOMBRE" : "Nombre_Vendedor",
            "VENCODIGO" : "Codigo_Vendedor",
            "CCONOMBRE" : "Centro_Costos",
            "MCDCCOSTO" : "Codigo_Centro_Costos",
            "META_INTER" : "Meta_Intereses",
            "META_DC_AL" : "Meta_DC_Al_Dia",
            "META_DC_AT" : "Meta_DC_Atraso",
            "META_ATRAS" : "Meta_Atraso"
        }
    },
    "VENCIMIENTOS": {
        "usecols": ["MCNVINCULA", "VINTELEFO3", "SALDODOC", "VENCE", "VINTELEFON", "MCNCUOCRU1"],
        "rename_map": {
            "MCNVINCULA": "Cedula_Cliente",
            "VINTELEFO3": "Celular",
            "VINTELEFON" : "Telefono",
            "SALDODOC": "Saldo_Factura",
            "MCNCUOCRU1": "Cuota_Vigente",
            "VENCE": "Fecha_Vencimiento",
        }
    },
    "R03":{
        "usecols": ["CODEUDOR1","NOMBRE1","VINTELEFON","CIUNOMBRE1","CODEUDOR2","NOMBRE2","VINTELEFO2","CIUNOMBRE2","CEDULA"],
        "rename_map": {
            "CODEUDOR1": "Codeudor1",
            "NOMBRE1": "Nombre_Codeudor1",
            "VINTELEFON": "Telefono_Codeudor1",
            "CIUNOMBRE1": "Ciudad_Codeudor1",
            "CODEUDOR2": "Codeudor2",
            "NOMBRE2": "Nombre_Codeudor2",
            "VINTELEFO2": "Telefono_Codeudor2",
            "CIUNOMBRE2": "Ciudad_Codeudor2",
            "CEDULA": "Cedula_Cliente"
        }
    },
    "CRTMPCONSULTA1":{
        "usecols":["CORREO","FECHA_FACT","TIPO_DOCUM","NUMERO_DOC","IDENTIFICA"],
        "rename_map":{
            "CORREO": "Correo",
            "FECHA_FACT":"Fecha_Facturada",
            "TIPO_DOCUM":"Tipo_Credito",
            "NUMERO_DOC":"Numero_Credito",
            "IDENTIFICA":"Cedula_Cliente"
        }
    },
    "FNZ003":{
        "usecols":["CREDITO","CONCEPTO","SALDO"],
        "rename_map":{
            "CREDITO":"Credito",
            "CONCEPTO":"Concepto",
            "SALDO":"Saldo"
        }
    }
}

# --- Lista de archivos a procesar ---
# ruta_base = 'C:/Users/usuario/Desktop/JUNIO/'
ruta_base = '/home/balvin/dev/electrocreditos/JUNIO/'
archivos_a_procesar = [
    ruta_base + "ANALISIS ARP GENERAL 0506INICIAL.XLS",
    ruta_base + "ANALISIS FNS GENERAL 0506INICIAL.XLS",
    ruta_base + "R91 ARP JUNIO.XLSX",
    ruta_base + "R91 FS JUNIO.XLSX",
    ruta_base + "VENCIMIENTOS ARP JUNIO.XLSX",
    ruta_base + "VENCIMIENTOS FNS JUNIO.XLSX",
    ruta_base + "R03 2025 FNS.xlsx",
    ruta_base + "R03 2025 ARP.xlsx",
    ruta_base + "CRTMPCONSULTA1.xlsx",
    ruta_base + "FNZ003 A 20 JUN.XLSX"
]

# --- Diccionario para agrupar dataframes por tipo ---
# --- Diccionario para agrupar dataframes por tipo (usando las claves estandarizadas) ---
dataframes_por_tipo = {key: [] for key in configuracion.keys()}

# --- Proceso de Lectura y Agrupación ---
print("--- 🔄 Iniciando lectura de archivos ---")
for ruta_archivo in archivos_a_procesar:
    try:
        nombre_archivo = ruta_archivo.split('/')[-1]
        tipo_archivo_actual = None
        
        # ✨ MEJORA: Bucle dinámico para detectar el tipo de archivo.
        # Esto hace que no necesites añadir más 'elif' en el futuro.
        for tipo in configuracion.keys():
            if tipo in nombre_archivo.replace(" ", "_"): # Reemplaza espacios para coincidir mejor
                 tipo_archivo_actual = tipo
                 break
        
        if not tipo_archivo_actual:
            print(f"⚠️  Archivo '{nombre_archivo}' omitido (no coincide con ningún tipo de configuración).")
            continue

        df = pd.read_excel(ruta_archivo, engine='xlrd' if ruta_archivo.upper().endswith('.XLS') else 'openpyxl')
        config = configuracion[tipo_archivo_actual]
        
        # Filtrar solo las columnas que existen en el archivo para evitar errores
        columnas_a_usar = [col for col in config["usecols"] if col in df.columns]
        df_filtrado = df[columnas_a_usar]
        
        if tipo_archivo_actual == "R03":
            df_filtrado = df_filtrado.replace('.', 'SIN CODEUDOR').fillna('SIN CODEUDOR')
            
        df_renombrado = df_filtrado.rename(columns=config["rename_map"])
        dataframes_por_tipo[tipo_archivo_actual].append(df_renombrado)
        print(f"✅ Archivo '{nombre_archivo}' procesado como tipo '{tipo_archivo_actual}'.")
        
    except Exception as e:
        print(f"❌ Error procesando el archivo '{ruta_archivo}': {e}")

# --- Consolidación, Cruce y Exportación ---
# Se verifica que existan datos en R91 para poder construir el reporte
if dataframes_por_tipo.get("R91"):
    # 1. Consolidar todos los dataframes de forma segura usando .get()
    print("\n---  consolidating dataframes ---")
    analisis_df = pd.concat(dataframes_por_tipo.get("ANALISIS", []), ignore_index=True)
    r91_df = pd.concat(dataframes_por_tipo.get("R91", []), ignore_index=True)
    vencimientos_df = pd.concat(dataframes_por_tipo.get("VENCIMIENTOS", []), ignore_index=True)
    r03_df = pd.concat(dataframes_por_tipo.get("R03", []), ignore_index=True)
    crtmp_df = pd.concat(dataframes_por_tipo.get("CRTMPCONSULTA1", []), ignore_index=True)
    fnz003_df = pd.concat(dataframes_por_tipo.get("FNZ003", []), ignore_index=True)

    # 2. Construir el reporte base desde R91 y crear llaves principales
    print("🔗 Creando el reporte base y llaves principales...")
    reporte_final = r91_df.copy()
    reporte_final['Numero_Credito'] = pd.to_numeric(reporte_final['Numero_Credito'], errors='coerce').astype('Int64')
    reporte_final['Credito'] = reporte_final['Tipo_Credito'].astype(str) + '-' + reporte_final['Numero_Credito'].astype(str)
    reporte_final['Empresa'] = np.where(reporte_final['Tipo_Credito'] == 'DF', 'FINANSUEÑOS', 'ARPESOD')

    # --- 3. PREPARAR Y UNIR SALDO CAPITAL ---
    print("🔍 Preparando y uniendo Saldo Capital...")

    # Mapa para ARPESOD desde ANALISIS
    if not analisis_df.empty:
        analisis_df['Credito'] = analisis_df['Tipo_Credito'].astype(str) + '-' + pd.to_numeric(analisis_df['Numero_Credito'], errors='coerce').astype('Int64').astype(str)
        mapa_saldo_arpesod = analisis_df.drop_duplicates('Credito').set_index('Credito')['Saldo_Factura']
        reporte_final = reporte_final.merge(mapa_saldo_arpesod.rename('Saldo_Capital_Arpesod'), on='Credito', how='left')
    
    # Mapa para FINANSUEÑOS desde FNZ003
    if not fnz003_df.empty:
        conceptos_deseados = ['CAPITAL', 'ABONO DIF TASA']
        fnz003_filtrado = fnz003_df[fnz003_df['Concepto'].isin(conceptos_deseados)]
        mapa_saldo_finansueños = fnz003_filtrado.groupby('Credito')['Saldo'].sum()
        reporte_final = reporte_final.merge(mapa_saldo_finansueños.rename('Saldo_Capital_Finansueños'), on='Credito', how='left')

    # CREAR LA COLUMNA FINAL 'Saldo_Capital'
    reporte_final['Saldo_Capital'] = np.where(
        reporte_final['Empresa'] == 'ARPESOD',
        reporte_final.get('Saldo_Capital_Arpesod'),
        reporte_final.get('Saldo_Capital_Finansueños')
    )

    
    reporte_final['Saldo_Capital'] = pd.to_numeric(reporte_final['Saldo_Capital'], errors='coerce').fillna(0).astype(int)
    reporte_final = reporte_final.drop(columns=['Saldo_Capital_Arpesod', 'Saldo_Capital_Finansueños'], errors='ignore')

    # ✨✨✨ AÑADIR ESTE BLOQUE NUEVO ✨✨✨
# Unir la información adicional de los archivos de ANÁLISIS
    if not analisis_df.empty:
      print("🔗 Uniendo información adicional del reporte de Análisis...")
    # Seleccionamos las columnas que queremos añadir desde analisis_df
      columnas_analisis_a_unir = [
        'Credito', 'Direccion', 'Barrio', 'Nombre_Ciudad', 
        'Total_Cuotas', 'Valor_Cuota', 'Dias_Atraso', 'Cuotas_Pagadas'
    ]
    
    # Nos aseguramos de no duplicar la llave y tomamos el primer registro por crédito
    info_analisis = analisis_df[columnas_analisis_a_unir].drop_duplicates('Credito')
    
    # Hacemos el merge con el reporte final usando la llave 'Credito'
    reporte_final = pd.merge(reporte_final, info_analisis, on='Credito', how='left')
# --- FIN DEL BLOQUE NUEVO ---
    # --- 3B. PREPARAR Y UNIR SALDO AVALES Y SALDO INTERÉS CORRIENTE ---
    print("📊 Calculando Saldo de Avales e Interés Corriente...")
    if not fnz003_df.empty:
        # --- Cálculo de Saldo Avales ---
        fnz003_avales = fnz003_df[fnz003_df['Concepto'] == 'AVAL']
        mapa_saldo_avales = fnz003_avales.groupby('Credito')['Saldo'].sum()
        reporte_final = reporte_final.merge(mapa_saldo_avales.rename('Saldo_Avales_Finansueños'), on='Credito', how='left')
        reporte_final['Saldo_Avales_Finansueños'] = reporte_final['Saldo_Avales_Finansueños'].fillna(0).astype(int)

        # --- Cálculo de Saldo Interés Corriente ---
        fnz003_interes = fnz003_df[fnz003_df['Concepto'] == 'INTERES CORRIENTE']
        mapa_saldo_interes = fnz003_interes.groupby('Credito')['Saldo'].sum()
        reporte_final = reporte_final.merge(mapa_saldo_interes.rename('Saldo_Interes_Finansueños'), on='Credito', how='left')
        reporte_final['Saldo_Interes_Finansueños'] = reporte_final['Saldo_Interes_Finansueños'].fillna(0).astype(int)

    # Crear columnas finales con la lógica condicional
    reporte_final['Saldo_Avales'] = np.where(
        reporte_final['Empresa'] == 'FINANSUEÑOS',
        reporte_final.get('Saldo_Avales_Finansueños'),
        'NO APLICA'
    )
    reporte_final['Saldo_Interes_Corriente'] = np.where(
        reporte_final['Empresa'] == 'FINANSUEÑOS',
        reporte_final.get('Saldo_Interes_Finansueños'),
        'NO APLICA'
    )

    # --- Validación de Avales Negativos ---
    avales_numericos = pd.to_numeric(reporte_final['Saldo_Avales'], errors='coerce')
    if (avales_numericos < 0).any():
        print("\n⚠️ ALERTA: Se encontraron saldos de avales negativos para FINANSUEÑOS.")
        # Opcional: mostrar los créditos con avales negativos
        # print(reporte_final[avales_numericos < 0][['Credito', 'Saldo_Avales']])
        
    # Limpiar columnas temporales
    reporte_final = reporte_final.drop(columns=['Saldo_Avales_Finansueños', 'Saldo_Interes_Finansueños'], errors='ignore')

    # --- 4. UNIR EL RESTO DE LA INFORMACIÓN ---
    print("🔗 Uniendo el resto de la información...")
    if not vencimientos_df.empty:
        vencimientos_df_map = vencimientos_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')
        reporte_final = pd.merge(reporte_final, vencimientos_df_map, on='Cedula_Cliente', how='left', suffixes=('', '_Vencimientos'))
    if not r03_df.empty:
        r03_df_map = r03_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')
        reporte_final = pd.merge(reporte_final, r03_df_map, on='Cedula_Cliente', how='left', suffixes=('', '_R03'))

    # --- 5. OTRAS TRANSFORMACIONES ---
    # (El resto de tus transformaciones y limpieza se mantienen igual)
    for col_cuota in ['Cuotas_Pagadas', 'Cuota_Vigente']:
        if col_cuota in reporte_final.columns:
            reporte_final[col_cuota] = pd.to_numeric(reporte_final[col_cuota], errors='coerce')
            # Esta es una regla de negocio muy específica, se mantiene como está.
            reporte_final[col_cuota] = np.where(reporte_final[col_cuota] > 99, reporte_final[col_cuota] % 100, reporte_final[col_cuota])
            reporte_final[col_cuota] = reporte_final[col_cuota].astype('Int64')
    if 'Codigo_Vendedor' in reporte_final.columns:
        reporte_final['Codigo_Vendedor'] = reporte_final['Codigo_Vendedor'].astype(str)

    cols_a_borrar = [col for col in reporte_final.columns if col.endswith(('_Vencimientos', '_R03'))]
    reporte_final = reporte_final.drop(columns=cols_a_borrar)

    print("📅 Formateando fechas a DD/MM/YY...")
    for col_fecha in ['Fecha_Vencimiento', 'Fecha_Facturada']:
        if col_fecha in reporte_final.columns:
            # Convierte a datetime, los errores se convertirán en NaT (Not a Time)
            reporte_final[col_fecha] = pd.to_datetime(reporte_final[col_fecha], errors='coerce')
            # Formatea a DD/MM/YY y reemplaza los valores nulos (NaT) con un string vacío
            reporte_final[col_fecha] = reporte_final[col_fecha].dt.strftime('%d/%m/%y').fillna('')

    
    print("\n--- 📊 Vista Previa del Reporte Final ---")
    print(reporte_final.head())
    print(f"\n--- Total de registros: {len(reporte_final)} ---")
    
    # --- Guardar el DataFrame final ---
    nombre_archivo_salida = 'Reporte_Consolidado_Final.xlsx'
    reporte_final.to_excel(nombre_archivo_salida, index=False, sheet_name='Reporte Consolidado')

    print(f"\n✨ ¡Éxito! El reporte final se ha guardado como '{nombre_archivo_salida}' ✨")
else:
    print("\n❌ No se encontraron archivos R91 para construir el reporte base. Proceso detenido.")