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
}

# --- Lista de archivos a procesar ---
ruta_base = 'c:/Users/usuario/Desktop/JUNIO/'
archivos_a_procesar = [
    ruta_base + "ANALISIS ARP GENERAL 0506INICIAL.XLS",
    ruta_base + "ANALISIS FNS GENERAL 0506INICIAL.XLS",
    ruta_base + "R91 ARP JUNIO.XLSX",
    ruta_base + "R91 FS JUNIO.XLSX",
    ruta_base + "VENCIMIENTOS ARP JUNIO.XLSX",
    ruta_base + "VENCIMIENTOS FNS JUNIO.XLSX"
]

# --- Diccionario para agrupar dataframes por tipo ---
dataframes_por_tipo = {
    "ANALISIS": [],
    "R91": [],
    "VENCIMIENTOS": []
}

# --- Proceso de Lectura y Agrupación ---
for ruta_archivo in archivos_a_procesar:
    try:
        nombre_archivo = ruta_archivo.split('/')[-1]
        tipo_archivo_actual = ""
        if "ANALISIS" in nombre_archivo: tipo_archivo_actual = "ANALISIS"
        elif "R91" in nombre_archivo: tipo_archivo_actual = "R91"
        elif "VENCIMIENTOS" in nombre_archivo: tipo_archivo_actual = "VENCIMIENTOS"
        else: continue
        
        if ruta_archivo.upper().endswith('.XLSX'):
            df = pd.read_excel(ruta_archivo)
        elif ruta_archivo.upper().endswith('.XLS'):
            df = pd.read_excel(ruta_archivo, engine='xlrd')
        
        config = configuracion[tipo_archivo_actual]
        columnas_existentes = [col for col in config["usecols"] if col in df.columns]
        df_filtrado = df[columnas_existentes]
        df_renombrado = df_filtrado.rename(columns=config["rename_map"])
        
        dataframes_por_tipo[tipo_archivo_actual].append(df_renombrado)
        
        print(f"✅ Archivo '{nombre_archivo}' procesado y agrupado.")

    except Exception as e:
        print(f"❌ Error procesando el archivo '{ruta_archivo}': {e}")

# --- Consolidación, Cruce y Exportación ---
if any(dataframes_por_tipo.values()):
    analisis_df = pd.concat(dataframes_por_tipo["ANALISIS"], ignore_index=True) if dataframes_por_tipo["ANALISIS"] else pd.DataFrame()
    r91_df = pd.concat(dataframes_por_tipo["R91"], ignore_index=True) if dataframes_por_tipo["R91"] else pd.DataFrame()
    vencimientos_df = pd.concat(dataframes_por_tipo["VENCIMIENTOS"], ignore_index=True) if dataframes_por_tipo["VENCIMIENTOS"] else pd.DataFrame()

    if not analisis_df.empty:
        analisis_df = analisis_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')
    if not r91_df.empty:
        r91_df = r91_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')
    if not vencimientos_df.empty:
        vencimientos_df = vencimientos_df.drop_duplicates(subset=['Cedula_Cliente'], keep='first')

    print("\n🔗 Cruzando información de los reportes...")
    
    reporte_final = pd.DataFrame()
    list_of_dfs = [df for df in [analisis_df, r91_df, vencimientos_df] if not df.empty]
    
    if list_of_dfs:
        reporte_final = list_of_dfs[0]
        for df_to_merge in list_of_dfs[1:]:
            reporte_final = pd.merge(reporte_final, df_to_merge, on="Cedula_Cliente", how="outer")

    # --- INICIO DE LAS MODIFICACIONES ---

    # ✨ LÓGICA CORREGIDA PARA LAS CUOTAS ✨
    for col_cuota in ['Cuotas_Pagadas', 'Cuota_Vigente']:
        if col_cuota in reporte_final.columns:
            # Primero, asegurar que la columna sea numérica, convirtiendo errores en Nulo (NaN)
            reporte_final[col_cuota] = pd.to_numeric(reporte_final[col_cuota], errors='coerce')
            
            # Aplicar la corrección solo a números mayores a 99
            reporte_final[col_cuota] = np.where(
                reporte_final[col_cuota] > 99,              # Condición: si el valor tiene más de 2 dígitos
                reporte_final[col_cuota] % 100,             # Verdadero: se obtienen los últimos 2 dígitos (ej: 2014 -> 14)
                reporte_final[col_cuota]                    # Falso: se conserva el valor original (ej: 6 -> 6)
            )
            # Finalmente, convertir a tipo entero para un formato limpio
            reporte_final[col_cuota] = reporte_final[col_cuota].astype('Int64')

    # Ajustar 'Codigo_Vendedor' para que sea de tipo string
    if 'Codigo_Vendedor' in reporte_final.columns:
        reporte_final['Codigo_Vendedor'] = reporte_final['Codigo_Vendedor'].astype(str)

    # Crear la columna 'Credito' (asegurando que no tenga decimales)
    if 'Tipo_Credito' in reporte_final.columns and 'Numero_Credito' in reporte_final.columns:
        reporte_final['Numero_Credito'] = pd.to_numeric(reporte_final['Numero_Credito'], errors='coerce').astype('Int64')
        reporte_final['Credito'] = reporte_final['Tipo_Credito'].astype(str) + '-' + reporte_final['Numero_Credito'].astype(str)
        reporte_final['Empresa'] = np.where(
            reporte_final['Credito'].str.startswith('DF'), 
            'FINANSUEÑOS', 
            'ARPESOD'
        )
    
    # --- FIN DE LAS MODIFICACIONES ---

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