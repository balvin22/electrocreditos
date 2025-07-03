import pandas as pd
import io
import numpy as np

# --- La configuración de columnas que definiste se mantiene igual ---
configuracion = {
    "ANALISIS": {
        "usecols": ["dirección", "barrio", "nomciudad", "totcuotas", "valorcuota", "díasatras", "cuotaspag","cedula","saldofac","tipo","numero"],
        "rename_map": {
            "dirección": "Direccion",
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
            "INTERES" : "Intereses",
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
        "usecols":["CORREO","FECHA_FACT","TIPO_DOCUM","NUMERO_DOC","IDENTIFICA"],
        "rename_map":{
            "CORREO": "Correo",
            "FECHA_FACT":"Fecha_Facturada",
            "TIPO_DOCUM":"Tipo_Credito",
            "NUMERO_DOC":"Numero_Credito",
            "IDENTIFICA":"Cedula_Cliente"
        }
    },
    "FNZ003 A 20 JUN":{
        "usecols":["CREDITO","CONCEPTO","VALOR"],
        "rename_map":{
            "CREDITO":"Credito",
            "CONCEPTO":"Concepto",
            "VALOR":"Valor"
        }
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
    ruta_base + "FNZ003 A 20 JUN.xlsx"
]

# --- Diccionario para agrupar dataframes por tipo ---
dataframes_por_tipo = {
    "ANALISIS": [],
    "R91": [],
    "VENCIMIENTOS": [],
    "R03": [],
    "CRTMPCONSULTA1":[],
    "FNZ003 A 20 JUN":[]
}

# --- Proceso de Lectura y Agrupación ---
for ruta_archivo in archivos_a_procesar:
    try:
        nombre_archivo = ruta_archivo.split('/')[-1]
        tipo_archivo_actual = ""
        if "ANALISIS" in nombre_archivo: tipo_archivo_actual = "ANALISIS"
        elif "R91" in nombre_archivo: tipo_archivo_actual = "R91"
        elif "VENCIMIENTOS" in nombre_archivo: tipo_archivo_actual = "VENCIMIENTOS"
        elif "R03" in nombre_archivo: tipo_archivo_actual = "R03"
        elif "CRTMPCONSULTA1" in nombre_archivo: tipo_archivo_actual = "CRTMPCONSULTA1"
        # ✨ CORRECCIÓN: Se añade la condición para leer el archivo FNZ003 ✨
        elif "FNZ003" in nombre_archivo: tipo_archivo_actual = "FNZ003"
        else: continue

        df = pd.read_excel(ruta_archivo, engine='xlrd' if ruta_archivo.upper().endswith('.XLS') else 'openpyxl')
        config = configuracion[tipo_archivo_actual]
        columnas_a_usar = [col for col in config["usecols"] if col in df.columns]
        df_filtrado = df[columnas_a_usar]
        if tipo_archivo_actual == "R03":
            df_filtrado = df_filtrado.replace('.', 'SIN CODEUDOR').fillna('SIN CODEUDOR')
        df_renombrado = df_filtrado.rename(columns=config["rename_map"])
        dataframes_por_tipo[tipo_archivo_actual].append(df_renombrado)
        print(f"✅ Archivo '{nombre_archivo}' procesado y agrupado.")
    except Exception as e:
        print(f"❌ Error procesando el archivo '{ruta_archivo}': {e}")

# --- Consolidación, Cruce y Exportación ---
if dataframes_por_tipo["R91"]:
    # 1. Consolidar todos los dataframes
    analisis_df = pd.concat(dataframes_por_tipo["ANALISIS"], ignore_index=True) if dataframes_por_tipo["ANALISIS"] else pd.DataFrame()
    r91_df = pd.concat(dataframes_por_tipo["R91"], ignore_index=True)
    vencimientos_df = pd.concat(dataframes_por_tipo["VENCIMIENTOS"], ignore_index=True) if dataframes_por_tipo["VENCIMIENTOS"] else pd.DataFrame()
    r03_df = pd.concat(dataframes_por_tipo["R03"], ignore_index=True) if dataframes_por_tipo["R03"] else pd.DataFrame()
    crtmp_df = pd.concat(dataframes_por_tipo["CRTMPCONSULTA1"], ignore_index=True) if dataframes_por_tipo["CRTMPCONSULTA1"] else pd.DataFrame()
    fnz003_df = pd.concat(dataframes_por_tipo["FNZ003"], ignore_index=True) if dataframes_por_tipo["FNZ003"] else pd.DataFrame()

    # 2. Construir el reporte base desde R91 y crear llaves principales
    print("\n🔗 Creando el reporte base y llaves principales...")
    reporte_final = r91_df.copy()
    if 'Tipo_Credito' in reporte_final.columns and 'Numero_Credito' in reporte_final.columns:
        reporte_final['Numero_Credito'] = pd.to_numeric(reporte_final['Numero_Credito'], errors='coerce').astype('Int64')
        reporte_final['Credito'] = reporte_final['Tipo_Credito'].astype(str) + '-' + reporte_final['Numero_Credito'].astype(str)
        reporte_final['Empresa'] = np.where(reporte_final['Credito'].str.startswith('DF'), 'FINANSUEÑOS', 'ARPESOD')

    # --- 3. PREPARAR Y UNIR SALDO CAPITAL ---
    print("🔍 Preparando y uniendo Saldo Capital...")

    # Mapa para ARPESOD desde ANALISIS
    if not analisis_df.empty and all(col in analisis_df.columns for col in ['Tipo_Credito', 'Numero_Credito', 'Saldo_Factura']):
        analisis_df['Credito'] = analisis_df['Tipo_Credito'].astype(str) + '-' + pd.to_numeric(analisis_df['Numero_Credito'], errors='coerce').astype('Int64').astype(str)
        mapa_saldo_arpesod = analisis_df[['Credito', 'Saldo_Factura']].rename(columns={'Saldo_Factura': 'Saldo_Capital_Arpesod'}).drop_duplicates()
        reporte_final = pd.merge(reporte_final, mapa_saldo_arpesod, on='Credito', how='left')
    
    # Mapa para FINANSUEÑOS desde FNZ003
    if not fnz003_df.empty:
        conceptos_deseados = ['CAPITAL', 'ABONO DIF TASA']
        fnz003_filtrado = fnz003_df[fnz003_df['Concepto'].isin(conceptos_deseados)]
        mapa_saldo_finansueños = fnz003_filtrado.groupby('Credito')['Valor'].sum().reset_index().rename(columns={'Valor': 'Saldo_Capital_Finansueños'})
        reporte_final = pd.merge(reporte_final, mapa_saldo_finansueños, on='Credito', how='left')

    # CREAR LA COLUMNA FINAL 'Saldo_Capital'
    reporte_final['Saldo_Capital'] = np.where(
        reporte_final['Empresa'] == 'ARPESOD',
        reporte_final.get('Saldo_Capital_Arpesod'),
        reporte_final.get('Saldo_Capital_Finansueños')
    )
    # Limpiar columnas temporales
    reporte_final = reporte_final.drop(columns=['Saldo_Capital_Arpesod', 'Saldo_Capital_Finansueños'], errors='ignore')

    # --- 4. UNIR EL RESTO DE LA INFORMACIÓN ---
    print("🔗 Uniendo el resto de la información...")
    # Unir VENCIMIENTOS Y R03 por Cedula_Cliente
    if not vencimientos_df.empty:
        vencimientos_df_map = vencimientos_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')
        reporte_final = pd.merge(reporte_final, vencimientos_df_map, on='Cedula_Cliente', how='left', suffixes=('', '_Vencimientos'))
    if not r03_df.empty:
        r03_df_map = r03_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')
        reporte_final = pd.merge(reporte_final, r03_df_map, on='Cedula_Cliente', how='left', suffixes=('', '_R03'))

    # ... Aquí puedes añadir el merge de CRTMPCONSULTA1 si aún lo necesitas ...
    
    # --- 5. OTRAS TRANSFORMACIONES ---
    for col_cuota in ['Cuotas_Pagadas', 'Cuota_Vigente']:
        if col_cuota in reporte_final.columns:
            reporte_final[col_cuota] = pd.to_numeric(reporte_final[col_cuota], errors='coerce')
            reporte_final[col_cuota] = np.where(reporte_final[col_cuota] > 99, reporte_final[col_cuota] % 100, reporte_final[col_cuota])
            reporte_final[col_cuota] = reporte_final[col_cuota].astype('Int64')
    if 'Codigo_Vendedor' in reporte_final.columns:
        reporte_final['Codigo_Vendedor'] = reporte_final['Codigo_Vendedor'].astype(str)

    # Limpiar columnas duplicadas por los merges, manteniendo las del reporte base (izquierda)
    cols_a_borrar = [col for col in reporte_final.columns if col.endswith('_Vencimientos') or col.endswith('_R03')]
    reporte_final = reporte_final.drop(columns=cols_a_borrar)
    
    print("\n--- Vista Previa del Reporte Final ---")
    print(reporte_final.head())
    print(f"\n--- Total de registros en el reporte final: {len(reporte_final)} ---")
    print("\n--- Estructura del Reporte Final ---")
    reporte_final.info()

    # --- Guardar el DataFrame final en un archivo Excel ---
    nombre_archivo_salida = 'Reporte_Consolidado_Final.xlsx'
    reporte_final.to_excel(nombre_archivo_salida, index=False, sheet_name='Reporte Consolidado')

    print(f"\n✨ ¡Éxito! El reporte final se ha guardado como '{nombre_archivo_salida}'")
else:
    print("\n❌ No se encontraron archivos R91 para construir el reporte base. Proceso detenido.")