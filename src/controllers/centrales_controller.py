import pandas as pd
import io
import numpy as np

# --- La configuración de columnas que definiste se mantiene igual ---
configuracion = {
    "ANALISIS": {
        "usecols": ["dirección", "barrio", "nomciudad", "totcuotas", "valorcuota", "díasatras", "cuotaspag","cedula"],
        "rename_map": {
            "dirección": "Direccion",
            "barrio": "Barrio",
            "nomciudad": "Nombre_Ciudad",
            "totcuotas": "Total_Cuotas",
            "valorcuota": "Valor_Cuota",
            "díasatras": "Dias_Atraso",
            "cuotaspag": "Cuotas_Pagadas",
            "cedula" : "Cedula_Cliente"
        }
    },
    "R91": {
        "usecols": ["VINNOMBRE", "VENNOMBRE", "MCDZONA", "MCDVINCULA", "MCDNUMCRU1", "MCDTIPCRU1","VENOMBRE","VENCODIGO",
                    "COBNOMBRE", "MCDCCOSTO", "CCONOMBRE", "META_INTERES", "DC AL DIA", "DC ATRASO", "META_ATRASO"],
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
            "META_INTERES" : "Interes_Mora",
            "DC AL DIA" : "DC_Al_Dia",
            "DC ATRASO" : "DC_Atraso",
            "META_ATRASO" : "Meta_Atraso"
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
        "usecols":["CORREO","FECHA_FACT","TIPO_DOCUM","NUMERO_DOC"],
        "rename_map":{
            "CORREO": "Correo",
            "FECHA_FACT":"Fecha_facturada",
            "TIPO_DOCUM":"Tipo_Credito",
            "NUMERO_DOC":"Numero_Credito"
        }
    }
}

# --- Lista de archivos a procesar ---
# ruta_base = '/home/balvin/dev/electrocreditos/JUNIO/'
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
    ruta_base + "CRTMPCONSULTA1.xlsx"
]

# --- Diccionario para agrupar dataframes por tipo ---
dataframes_por_tipo = {
    "ANALISIS": [],
    "R91": [],
    "VENCIMIENTOS": [],
    "R03": [],
    "CRTMPCONSULTA1":[]
}

# --- Proceso de Lectura y Agrupación ---
# --- Proceso de Lectura y Agrupación ---
for ruta_archivo in archivos_a_procesar:
    try:
        nombre_archivo = ruta_archivo.split('/')[-1]
        tipo_archivo_actual = ""
        if "ANALISIS" in nombre_archivo: tipo_archivo_actual = "ANALISIS"
        elif "R91" in nombre_archivo: tipo_archivo_actual = "R91"
        elif "VENCIMIENTOS" in nombre_archivo: tipo_archivo_actual = "VENCIMIENTOS"
        elif "R03" in nombre_archivo: tipo_archivo_actual = "R03"
        elif "CRTMPCONSULTA1" in nombre_archivo: tipo_archivo_actual = "CRTMPCONSULTA1" # <-- Añadido
        else: continue
        
        if ruta_archivo.upper().endswith('.XLSX'):
            df = pd.read_excel(ruta_archivo)
        elif ruta_archivo.upper().endswith('.XLS'):
            df = pd.read_excel(ruta_archivo, engine='xlrd')
        
        config = configuracion[tipo_archivo_actual]
        columnas_existentes = [col for col in config["usecols"] if col in df.columns]
        df_filtrado = df[columnas_existentes]

        if tipo_archivo_actual == "R03":
            df_filtrado = df_filtrado.replace('.', 'SIN CODEUDOR')
            df_filtrado = df_filtrado.fillna('SIN CODEUDOR')
        
        df_renombrado = df_filtrado.rename(columns=config["rename_map"])
        
        dataframes_por_tipo[tipo_archivo_actual].append(df_renombrado)
        
        print(f"✅ Archivo '{nombre_archivo}' procesado y agrupado.")

    except Exception as e:
        print(f"❌ Error procesando el archivo '{ruta_archivo}': {e}")

# --- Consolidación, Cruce y Exportación ---
if any(dataframes_por_tipo.values()):
    # 1. Consolidar dataframes para cada tipo
    analisis_df = pd.concat(dataframes_por_tipo["ANALISIS"], ignore_index=True) if dataframes_por_tipo["ANALISIS"] else pd.DataFrame()
    r91_df = pd.concat(dataframes_por_tipo["R91"], ignore_index=True) if dataframes_por_tipo["R91"] else pd.DataFrame()
    vencimientos_df = pd.concat(dataframes_por_tipo["VENCIMIENTOS"], ignore_index=True) if dataframes_por_tipo["VENCIMIENTOS"] else pd.DataFrame()
    r03_df = pd.concat(dataframes_por_tipo["R03"], ignore_index=True) if dataframes_por_tipo["R03"] else pd.DataFrame()
    crtmp_df = pd.concat(dataframes_por_tipo["CRTMPCONSULTA1"], ignore_index=True) if dataframes_por_tipo["CRTMPCONSULTA1"] else pd.DataFrame() # <-- Añadido

    # 2. Eliminar duplicados en cada dataframe consolidado
    if not analisis_df.empty:
        analisis_df = analisis_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')
    if not r91_df.empty:
        r91_df = r91_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')
    if not vencimientos_df.empty:
        vencimientos_df = vencimientos_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')
    if not r03_df.empty:
        r03_df = r03_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')

    # 3. Cruzar (merge) las tablas base
    print("\n🔗 Cruzando información de los reportes...")
    reporte_final = pd.DataFrame()
    list_of_base_dfs = [df for df in [analisis_df, r91_df, vencimientos_df] if not df.empty]
    
    if list_of_base_dfs:
        reporte_final = list_of_base_dfs[0]
        for df_to_merge in list_of_base_dfs[1:]:
            reporte_final = pd.merge(reporte_final, df_to_merge, on="Cedula_Cliente", how="outer")

    # 4. Unir la información del R03 con una unión "left"
    if not r03_df.empty and not reporte_final.empty:
        reporte_final = pd.merge(reporte_final, r03_df, on="Cedula_Cliente", how="left")
    
    # 5. Transformaciones principales
    for col_cuota in ['Cuotas_Pagadas', 'Cuota_Vigente']:
        if col_cuota in reporte_final.columns:
            reporte_final[col_cuota] = pd.to_numeric(reporte_final[col_cuota], errors='coerce')
            reporte_final[col_cuota] = np.where(
                reporte_final[col_cuota] > 99,
                reporte_final[col_cuota] % 100,
                reporte_final[col_cuota]
            )
            reporte_final[col_cuota] = reporte_final[col_cuota].astype('Int64')

    if 'Codigo_Vendedor' in reporte_final.columns:
        reporte_final['Codigo_Vendedor'] = reporte_final['Codigo_Vendedor'].astype(str)

    if 'Tipo_Credito' in reporte_final.columns and 'Numero_Credito' in reporte_final.columns:
        reporte_final['Numero_Credito'] = pd.to_numeric(reporte_final['Numero_Credito'], errors='coerce').astype('Int64')
        reporte_final['Credito'] = reporte_final['Tipo_Credito'].astype(str) + '-' + reporte_final['Numero_Credito'].astype(str)
        reporte_final['Empresa'] = np.where(
            reporte_final['Credito'].str.startswith('DF'), 
            'FINANSUEÑOS', 
            'ARPESOD'
        )

    # --- 6. PREPARAR Y UNIR DATOS DE CRTMPCONSULTA1 ---
    if not crtmp_df.empty:
        print("🔗 Cruzando información de CRTMPCONSULTA1...")
        # Crear la misma llave 'Credito' en crtmp_df
        crtmp_df['Numero_Credito'] = pd.to_numeric(crtmp_df['Numero_Credito'], errors='coerce').astype('Int64')
        crtmp_df['Credito'] = crtmp_df['Tipo_Credito'].astype(str) + '-' + crtmp_df['Numero_Credito'].astype(str)

        # Seleccionar solo las columnas nuevas y la llave para evitar duplicados
        columnas_para_unir = ['Credito', 'Correo', 'Fecha_facturada']
        crtmp_df_filtrado = crtmp_df[[col for col in columnas_para_unir if col in crtmp_df.columns]]
        
        # Eliminar duplicados en la llave de crtmp_df antes de unir
        crtmp_df_filtrado = crtmp_df_filtrado.drop_duplicates(subset=['Credito'], keep='first')

        # Unir con el reporte final usando la llave 'Credito'
        if not reporte_final.empty:
            reporte_final = pd.merge(reporte_final, crtmp_df_filtrado, on="Credito", how="left")

    print("\n--- Vista Previa del Reporte Consolidado y Cruzado ---")
    print(reporte_final.head())
    
    print("\n--- Estructura del Reporte Final ---")
    reporte_final.info()

    # --- Guardar el DataFrame final en un archivo Excel ---
    nombre_archivo_salida = 'Reporte_Consolidado_Final.xlsx'
    reporte_final.to_excel(nombre_archivo_salida, index=False, sheet_name='Reporte Consolidado')
    
    print(f"\n✨ ¡Éxito! El reporte final se ha guardado como '{nombre_archivo_salida}'")

else:
    print("\nNo se procesó ningún archivo.")