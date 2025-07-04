import pandas as pd
import numpy as np

configuracion = {
    "ANALISIS": {
        "usecols": ["direccion", "barrio", "nomciudad", "totcuotas", "valorcuota", "diasatras", "cuotaspag","cedula","saldofac","tipo","numero"],
        "rename_map": { "direccion": "Direccion",
                        "barrio": "Barrio",
                        "nomciudad": "Nombre_Ciudad",
                        "totcuotas": "Total_Cuotas",
                        "valorcuota": "Valor_Cuota",
                        "diasatras": "Dias_Atraso",
                        "cuotaspag": "Cuotas_Pagadas",
                        "cedula" : "Cedula_Cliente",
                        "saldofac":"Saldo_Factura",
                        "tipo":"Tipo_Credito",
                        "numero":"Numero_Credito" }
    },
    "R91": {
        "usecols": ["VINNOMBRE", "VENNOMBRE", "MCDZONA", "MCDVINCULA", "MCDNUMCRU1", "MCDTIPCRU1","VENOMBRE","VENCODIGO", "COBNOMBRE", "MCDCCOSTO", "CCONOMBRE", "META_INTER", "META_DC_AL", "META_DC_AT", "META_ATRAS"],
        "rename_map": { "MCDTIPCRU1": "Tipo_Credito",
                       "MCDNUMCRU1": "Numero_Credito",
                       "MCDVINCULA" : "Cedula_Cliente",
                       "VINNOMBRE": "Nombre_Cliente",
                       "MCDZONA" : "Zona",
                       "COBNOMBRE" : "Nombre_Cobrador",
                       "VENNOMBRE" : "Nombre_Vendedor",
                       "VENCODIGO" : "Codigo_Vendedor",
                       "CCONOMBRE" : "Centro_Costos",
                       "MCDCCOSTO" : "Codigo_Centro_Costos",
                       "META_INTER" : "Meta_Intereses",
                       "META_DC_AL" : "Meta_DC_Al_Dia",
                       "META_DC_AT" : "Meta_DC_Atraso",
                       "META_ATRAS" : "Meta_Atraso" }
    },
    "VENCIMIENTOS": {
        "usecols": ["MCNVINCULA", "VINTELEFO3", "SALDODOC", "VENCE", "VINTELEFON", "MCNCUOCRU1"],
        "rename_map": { "MCNVINCULA": "Cedula_Cliente",
                       "VINTELEFO3": "Celular",
                       "VINTELEFON" : "Telefono",
                       "SALDODOC": "Valor_Cuota_Vigente",
                       "MCNCUOCRU1": "Cuota_Vigente",
                       "VENCE": "Fecha_Vencimiento" }
    },
    "R03":{
        "usecols": ["CODEUDOR1","NOMBRE1","VINTELEFON","CIUNOMBRE1","CODEUDOR2","NOMBRE2","VINTELEFO2","CIUNOMBRE2","CEDULA"],
        "rename_map": { "CODEUDOR1": "Codeudor1",
                       "NOMBRE1": "Nombre_Codeudor1",
                       "VINTELEFON": "Telefono_Codeudor1",
                       "CIUNOMBRE1": "Ciudad_Codeudor1",
                       "CODEUDOR2": "Codeudor2",
                       "NOMBRE2": "Nombre_Codeudor2",
                       "VINTELEFO2": "Telefono_Codeudor2",
                       "CIUNOMBRE2": "Ciudad_Codeudor2",
                       "CEDULA": "Cedula_Cliente" }
    },
    "CRTMPCONSULTA1":{
        "usecols":["CORREO","FECHA_FACT","TIPO_DOCUM","NUMERO_DOC","IDENTIFICA"],
        "rename_map":{ "CORREO": "Correo",
                       "FECHA_FACT":"Fecha_Facturada",
                       "TIPO_DOCUM":"Tipo_Credito",
                       "NUMERO_DOC":"Numero_Credito",
                       "IDENTIFICA":"Cedula_Cliente" }
    },
    "FNZ003":{
        "usecols":["CREDITO","CONCEPTO","SALDO"],
        "rename_map":{ "CREDITO":"Credito",
                       "CONCEPTO":"Concepto",
                       "SALDO":"Saldo" }
    },
      "MATRIZ_CARTERA": {
        "skiprows": 2, 
        "header": None, 
        "new_names": [
            'Zona', 'Cobrador', 'telefono_cobrador', 'Regional', 'Gestor', 'gestor_telefono',
            'call_center_1_30_dias', 'call_center_nombre_1_30', 'call_center_telefono_1_30',
            'call_center_31_90_dias', 'call_center_nombre_31_90', 'call_center_telefono_31_90',
            'call_center_91_360_dias', 'call_center_nombre_91_360', 'call_center_telefono_91_360'
        ],
        "merge_on": "Zona" 
      },
      "ASESORES": {
        "sheets": [
            {
                "sheet_name": "ASESORES",
                "usecols": ["NOMBRE ASESOR", "MOVIL ASESOR", "LIDER ZONA", "MOVIL LIDER"],
                "rename_map": {
                    "NOMBRE ASESOR": "Nombre_Vendedor",
                    "MOVIL ASESOR": "Movil_Asesor",
                    "LIDER ZONA": "Lider_Zona",
                    "MOVIL LIDER": "Movil_Lider"
                },
                "merge_on": "Nombre_Vendedor"
            },
            {
                "sheet_name": "Centro Costos",
                "usecols": ["CENTRO DE COSTOS", "ACTIVO"],
                "rename_map": {
                    "CENTRO DE COSTOS": "Codigo_Centro_Costos",
                    "ACTIVO": "Activo_Centro_Costos"
                },
                "merge_on": "Codigo_Centro_Costos"
            }
        ]
    }
}

# --- Lista de archivos a procesar ---
ruta_base = 'C:/Users/usuario/Desktop/JUNIO/'
# ruta_base = '/home/balvin/dev/electrocreditos/JUNIO/'
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
    ruta_base + "FNZ003 A 20 JUN.XLSX",
    ruta_base + "MATRIZ DE CARTERA.xlsx",
    ruta_base + "ASESORES ACTIVOS.xlsx"
]

# --- 2. LECTURA Y PROCESAMIENTO DE ARCHIVOS ---
dataframes_por_tipo = {key: [] for key in configuracion.keys()}
print("--- 🔄 Iniciando lectura de archivos ---")
for ruta_archivo in archivos_a_procesar:
    try:
        nombre_archivo = ruta_archivo.split('/')[-1]
        tipo_archivo_actual = None
        
        nombre_base = nombre_archivo.split('.')[0].upper().replace(" ", "_")
        palabras_en_nombre = set(nombre_base.split('_'))
        
        for tipo in configuracion.keys():
            palabras_en_clave = set(tipo.split('_'))
            if palabras_en_clave.issubset(palabras_en_nombre):
                 tipo_archivo_actual = tipo
                 break
        
        if not tipo_archivo_actual:
            print(f"⚠️  Archivo '{nombre_archivo}' omitido (no coincide con ningún tipo de configuración).")
            continue

        config = configuracion[tipo_archivo_actual]
        
        # El bucle ahora maneja 3 tipos de configuraciones: normal, new_names y sheets
        if "sheets" in config:
            for sheet_config in config["sheets"]:
                sheet_name = sheet_config["sheet_name"]
                df_hoja = pd.read_excel(ruta_archivo, sheet_name=sheet_name, engine='openpyxl')
                columnas_a_usar = [col for col in sheet_config["usecols"] if col in df_hoja.columns]
                df_filtrado = df_hoja[columnas_a_usar]
                df_renombrado = df_filtrado.rename(columns=sheet_config["rename_map"])
                dataframes_por_tipo[tipo_archivo_actual].append({ "data": df_renombrado, "config": sheet_config })
        
        elif "new_names" in config:
            df = pd.read_excel(ruta_archivo, header=config.get("header"), skiprows=config.get("skiprows"), names=config.get("new_names"))
            dataframes_por_tipo[tipo_archivo_actual].append(df)
        
        else:
            df = pd.read_excel(ruta_archivo, engine='xlrd' if ruta_archivo.upper().endswith('.XLS') else 'openpyxl')
            columnas_a_usar = [col for col in config["usecols"] if col in df.columns]
            df_filtrado = df[columnas_a_usar]
            if tipo_archivo_actual == "R03":
                df_filtrado = df_filtrado.replace('.', 'SIN CODEUDOR').fillna('SIN CODEUDOR')
            df_renombrado = df_filtrado.rename(columns=config["rename_map"])
            dataframes_por_tipo[tipo_archivo_actual].append(df_renombrado)

        print(f"✅ Archivo '{nombre_archivo}' procesado como tipo '{tipo_archivo_actual}'.")
        
    except Exception as e:
        print(f"❌ Error procesando el archivo '{ruta_archivo}': {e}")

# --- 3. CONSOLIDACIÓN Y CONSTRUCCIÓN DEL REPORTE BASE ---
if dataframes_por_tipo.get("R91"):
    print("\n--- Consolidando DataFrames ---")

    def safe_concat(key):
        items = dataframes_por_tipo.get(key, [])
        if not items: return pd.DataFrame()
        if isinstance(items[0], dict):
            df_list = [item["data"] for item in items if "data" in item]
        else:
            df_list = items
        return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

    # Cargar todos los dataframes
    analisis_df = safe_concat("ANALISIS")
    r91_df = safe_concat("R91")
    vencimientos_df = safe_concat("VENCIMIENTOS")
    r03_df = safe_concat("R03")
    crtmp_df = safe_concat("CRTMPCONSULTA1")
    fnz003_df = safe_concat("FNZ003")
    matriz_cartera_df = safe_concat("MATRIZ_CARTERA")
    asesores_sheets = dataframes_por_tipo.get("ASESORES", [])

    # --- 4. CREACIÓN Y PREPARACIÓN DEL REPORTE FINAL ---
    print("🔗 Creando el reporte base y la llave principal 'Credito'...")
    reporte_final = r91_df.copy()
    reporte_final['Numero_Credito'] = pd.to_numeric(reporte_final['Numero_Credito'], errors='coerce').astype('Int64')
    reporte_final['Credito'] = reporte_final['Tipo_Credito'].astype(str) + '-' + reporte_final['Numero_Credito'].astype(str)
    reporte_final['Empresa'] = np.where(reporte_final['Tipo_Credito'] == 'DF', 'FINANSUEÑOS', 'ARPESOD')

    # --- 5. UNIÓN SECUENCIAL DE TODOS LOS ARCHIVOS ---
    print("🔍 Uniendo todos los archivos al reporte base...")

    # Función auxiliar para unir DataFrames sin duplicar columnas
    def merge_sin_duplicados(df_left, df_right, on_key):
        if df_right.empty:
            return df_left
        # Columnas en df_right que ya existen en df_left (excluyendo la llave)
        cols_a_quitar = [col for col in df_right.columns if col in df_left.columns and col not in on_key]
        df_right_limpio = df_right.drop(columns=cols_a_quitar)
        return pd.merge(df_left, df_right_limpio, on=on_key, how='left')

    # Unir ANALISIS
    if not analisis_df.empty:
        analisis_df['Numero_Credito'] = pd.to_numeric(analisis_df['Numero_Credito'], errors='coerce').astype('Int64')
        analisis_df['Credito'] = analisis_df['Tipo_Credito'].astype(str) + '-' + analisis_df['Numero_Credito'].astype(str)
        reporte_final = merge_sin_duplicados(reporte_final, analisis_df, on_key=['Credito'])

    # Unir VENCIMIENTOS
    if not vencimientos_df.empty:
        vencimientos_df_limpio = vencimientos_df.drop_duplicates(subset=['Cedula_Cliente'])
        reporte_final = merge_sin_duplicados(reporte_final, vencimientos_df_limpio, on_key=['Cedula_Cliente'])
        
    # Unir R03 (Codeudores)
    if not r03_df.empty:
        r03_df_limpio = r03_df.drop_duplicates(subset=['Cedula_Cliente'])
        reporte_final = merge_sin_duplicados(reporte_final, r03_df_limpio, on_key=['Cedula_Cliente'])

    # Unir CRTMPCONSULTA1
    if not crtmp_df.empty:
        crtmp_df['Numero_Credito'] = pd.to_numeric(crtmp_df['Numero_Credito'], errors='coerce').astype('Int64')
        crtmp_df['Credito'] = crtmp_df['Tipo_Credito'].astype(str) + '-' + crtmp_df['Numero_Credito'].astype(str)
        reporte_final = merge_sin_duplicados(reporte_final, crtmp_df, on_key=['Credito'])

    # Unir FNZ003 (Pivoteado)
    if not fnz003_df.empty:
        fnz003_pivot = fnz003_df.pivot_table(index='Credito', columns='Concepto', values='Saldo', aggfunc='first').reset_index()
        reporte_final = merge_sin_duplicados(reporte_final, fnz003_pivot, on_key=['Credito'])

    # Unir MATRIZ DE CARTERA
    if not matriz_cartera_df.empty:
        matriz_limpia = matriz_cartera_df.drop_duplicates(subset=['Zona'])
        reporte_final = merge_sin_duplicados(reporte_final, matriz_limpia, on_key=['Zona'])

    # Unir hojas de ASESORES
    for item in asesores_sheets:
        info_df = item["data"]
        merge_key = item["config"]["merge_on"]
        info_df_limpia = info_df.drop_duplicates(subset=[merge_key])
        reporte_final = merge_sin_duplicados(reporte_final, info_df_limpia, on_key=[merge_key])

    # --- 6. TRANSFORMACIONES Y LIMPIEZA FINAL ---
    print("🧹 Realizando transformaciones y limpieza final...")
    for col_cuota in ['Cuotas_Pagadas', 'Cuota_Vigente']:
        if col_cuota in reporte_final.columns:
            reporte_final[col_cuota] = pd.to_numeric(reporte_final[col_cuota], errors='coerce').astype('Int64')

    print("📅 Formateando fechas a DD/MM/YY...")
    for col_fecha in ['Fecha_Vencimiento', 'Fecha_Facturada']:
        if col_fecha in reporte_final.columns:
            reporte_final[col_fecha] = pd.to_datetime(reporte_final[col_fecha], errors='coerce').dt.strftime('%d/%m/%y').fillna('')

    # --- 7. EXPORTACIÓN DEL REPORTE FINAL ---
    print("\n--- 📊 Vista Previa del Reporte Final ---")
    print(reporte_final.head())
    print(f"\n--- Total de registros: {len(reporte_final)} ---")
    
    columnas_finales = sorted(reporte_final.columns.tolist())
    print(f"\n--- Columnas finales en el reporte ({len(columnas_finales)}) ---")
    print(columnas_finales)

    nombre_archivo_salida = 'Reporte_Consolidado_Final.xlsx'
    reporte_final.to_excel(nombre_archivo_salida, index=False, sheet_name='Reporte Consolidado')

    print(f"\n✨ ¡Éxito! El reporte final se ha guardado como '{nombre_archivo_salida}' ✨")
else:
    print("\n❌ No se encontraron archivos R91 para construir el reporte base. Proceso detenido.")