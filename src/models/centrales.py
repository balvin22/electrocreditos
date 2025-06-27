import pandas as pd
import numpy as np

# --- PASO 1: DEFINICIONES INICIALES (Sin cambios) ---
print("--- PASO 1: Definiendo parámetros iniciales ---")
input_file_path = 'c:/Users/usuario/Desktop/Reporte LV/DATA MARZO FS.TXT'
ruta_archivo_correcciones = 'c:/Users/usuario/Desktop/Reporte LV/Cédulas a revisar.xlsx' # Asegúrate que el nombre sea el correcto

colspecs = [
    (0, 1),(1, 12),(30, 75),(12, 30),(76, 84),(84, 92),
    (92, 94),(107, 109),(109, 110),(188, 199),(199, 210),
    (210, 221),(221, 232),(232, 243),(243, 246),(246, 249),
    (249, 252),(263, 271),(271, 279),(577, 597),(625, 685),
    (685, 745),(445, 457),(75, 76),(185, 188),(105, 106),
    (110, 118),(118, 120),(120, 128),(137, 138),(138, 146),
    (252, 255),(255, 263)
]
names = [
    "TIPO DE IDENTIFICACION", "NUMERO DE IDENTIFICACION", "NOMBRE COMPLETO",
    "NUMERO DE LA CUENTA U OBLIGACION", "FECHA APERTURA", "FECHA VENCIMIENTO",
    "RESPONSABLE", "NOVEDAD", "ESTADO ORIGEN DE LA CUENTA", "VALOR INICIAL",
    "VALOR SALDO DEUDA", "VALOR DISPONIBLE", "V CUOTA MENSUAL",
    "VALOR SALDO MORA", "TOTAL CUOTAS", "CUOTAS CANCELADAS", "CUOTAS EN MORA",
    "FECHA LIMITE DE PAGO", "FECHA DE PAGO", "CIUDAD CORRESPONDENCIA",
    "DIRECCION DE CORRESPONDENCIA", "CORREO ELECTRONICO", "CELULAR",
    "SITUACION DEL TITULAR", "EDAD DE MORA", "FORMA DE PAGO",
    "FECHA ESTADO ORIGEN", "ESTADO DE LA CUENTA", "FECHA ESTADO DE LA CUENTA",
    "ADJETIVO", "FECHA DE ADJETIVO","CLAUSULA DE PERMANENCIA","FECHA CLAUSULA DE PERMANENCIA"
]

try:
    # --- PASO 2: LECTURA Y CARGA INICIAL DEL ARCHIVO PRINCIPAL ---
    print("\n--- PASO 2: Leyendo archivo principal (DATA MARZO FS.TXT) ---")
    df = pd.read_fwf(
        input_file_path, colspecs=colspecs, names=names, encoding='latin-1',
        skiprows=1, skipfooter=1, engine='python'
    )
    if not df.empty:
        df = df.iloc[:-1]
    
    df['NUMERO DE IDENTIFICACION'] = df['NUMERO DE IDENTIFICACION'].astype(str).str.strip()

    # --- PASO 3: CORRECCIONES DESDE ARCHIVO EXCEL (EN ORDEN LÓGICO) ---
    print("\n--- PASO 3: Corrigiendo y enriqueciendo datos desde archivo Excel ---")
    try:
        # A. Corregir Cédulas
        print("   A. Corrigiendo 'NUMERO DE IDENTIFICACION'...")
        df_cedulas_malas = pd.read_excel(ruta_archivo_correcciones, sheet_name='Cedulas a corregir')
        mapa_correcciones = pd.Series(df_cedulas_malas['CEDULA CORRECTA'].astype(str).str.strip().values, index=df_cedulas_malas['CEDULA MAL'].astype(str).str.strip()).to_dict()
        df['NUMERO DE IDENTIFICACION'] = df['NUMERO DE IDENTIFICACION'].replace(mapa_correcciones)

        # B. Actualizar campos que dicen 'CORREGIR'
        print("   B. Actualizando campos 'CORREGIR' desde la hoja 'Vinculado'...")
        df_vinculado = pd.read_excel(ruta_archivo_correcciones, sheet_name='Vinculado')
        df_vinculado = df_vinculado.set_index(df_vinculado['CODIGO'].astype(str).str.strip())
        mapa_columnas_vinc = {'NOMBRE COMPLETO':'NOMBRE', 'DIRECCION DE CORRESPONDENCIA':'DIRECCI', 'CORREO ELECTRONICO':'VINEMAIL', 'CELULAR':'TELEFONO'}
        for col_df, col_vinc in mapa_columnas_vinc.items():
            mascara = df[col_df].astype(str).str.strip().str.contains('CORREGIR', case=False, na=False)
            if mascara.any():
                for index in df[mascara].index:
                    id_a_buscar = df.at[index, 'NUMERO DE IDENTIFICACION']
                    if id_a_buscar in df_vinculado.index:
                        df.at[index, col_df] = df_vinculado.at[id_a_buscar, col_vinc]

        # C. Actualizar Tipos de Identificación
        print("   C. Actualizando 'TIPO DE IDENTIFICACION'...")
        df['TIPO DE IDENTIFICACION'] = 1 # Valor por defecto
        df_tipos = pd.read_excel(ruta_archivo_correcciones, sheet_name='Tipos de identificacion')
        # *** CORRECCIÓN DEL ERROR DE NOMBRE DE COLUMNA ***
        columna_con_el_tipo_nuevo = 'CODIGO DATA' 
        mapa_tipos = pd.Series(df_tipos[columna_con_el_tipo_nuevo].values, index=df_tipos['CEDULA CORRECTA'].astype(str).str.strip()).to_dict()
        nuevos_tipos = df['NUMERO DE IDENTIFICACION'].map(mapa_tipos)
        df['TIPO DE IDENTIFICACION'].update(nuevos_tipos.dropna())
        
        print("Correcciones desde Excel completadas.")
    except Exception as e:
        print(f"   ¡ERROR DURANTE CORRECCIONES EXCEL! Detalle: {e}")

    # --- PASO 4: PROCESAMIENTO DE FNZ001 ---
    print("\n--- PASO 4: Procesando datos de FNZ001 para actualizar 'VALOR INICIAL' ---")
    try:
        df['NUMERO DE LA CUENTA U OBLIGACION'] = df['NUMERO DE LA CUENTA U OBLIGACION'].astype(str).str.replace(' ', '').str.zfill(18)
        
        df_fnz = pd.read_excel(ruta_archivo_correcciones, sheet_name='FNZ001', usecols=['DSM_TP', 'DSM_NUM', 'VLR_FNZ'])
        df_fnz['llave_base'] = df_fnz['DSM_TP'].astype(str).str.strip() + df_fnz['DSM_NUM'].astype(str).str.strip()
        
        df_base = pd.DataFrame({'FACTURA': df_fnz['llave_base'], 'VALOR': df_fnz['VLR_FNZ']})
        df_c1 = pd.DataFrame({'FACTURA': df_fnz['llave_base'] + 'C1', 'VALOR': df_fnz['VLR_FNZ']})
        df_c2 = pd.DataFrame({'FACTURA': df_fnz['llave_base'] + 'C2', 'VALOR': df_fnz['VLR_FNZ']})
        
        tabla_factura_valor = pd.concat([df_base, df_c1, df_c2], ignore_index=True)
        tabla_factura_valor['FACTURA'] = tabla_factura_valor['FACTURA'].astype(str).str.zfill(18)

        df = pd.merge(df, tabla_factura_valor, left_on='NUMERO DE LA CUENTA U OBLIGACION', right_on='FACTURA', how='left')
        
        df['VALOR INICIAL'] = np.where(df['VALOR'].notna(), df['VALOR'], df['VALOR INICIAL'])
        
        df = df.drop(columns=['FACTURA', 'VALOR'])
        print("Actualización de 'VALOR INICIAL' desde FNZ001 completada.")
        
    except Exception as e:
        print(f"   ¡ERROR! Ocurrió un error al procesar la hoja FNZ001: {e}")

    # --- PASO 5: LIMPIEZA GENERAL Y VALIDACIONES ---
    print("\n--- PASO 5: Realizando limpieza general y validaciones ---")
    letter_replacements = {'Ñ':'N','Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','Ü':'U','Ÿ':'Y','Â':'A','Ã':'A','š':'S','©':'C','ñ':'N','á':'A','é':'E','í':'I','ó':'O','ú':'U','ü':'U','ÿ':'Y','â':'A','ã':'A'}
    chars_to_remove = ['@','°','|','¬','¡','“','#','$','%','&','/','(',')','=','‘','\\','¿','+','~','´´','´','[','{','^','_','.',':',',',';','<','>','Æ','±']
    
    # *** CORRECCIÓN DEL ERROR DE TIPO DE DATO (.str) ***
    for col in df.columns:
        # Solo aplicar limpieza de texto a columnas que son de tipo 'object' (string)
        if df[col].dtype == 'object' and col != 'CORREO ELECTRONICO':
            df[col] = df[col].astype(str) # Aseguramos que sea string
            for old, new in letter_replacements.items():
                df[col] = df[col].str.replace(old, new, regex=False)
            for char in chars_to_remove:
                df[col] = df[col].str.replace(char, '', regex=False)

    df['FECHA APERTURA'] = df['FECHA APERTURA'].astype(str)
    df['FECHA VENCIMIENTO'] = df['FECHA VENCIMIENTO'].astype(str)
    condicion_fecha_invalida = df['FECHA VENCIMIENTO'] < df['FECHA APERTURA']
    df.loc[condicion_fecha_invalida, 'FECHA VENCIMIENTO'] = df['FECHA APERTURA']
    
    columnas_numericas = ["VALOR INICIAL", "VALOR SALDO DEUDA", "VALOR DISPONIBLE", "V CUOTA MENSUAL", "VALOR SALDO MORA"]
    for col in columnas_numericas:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        if col != 'VALOR DISPONIBLE':
            df.loc[df[col] < 10000, col] = 0
    df['VALOR DISPONIBLE'] = 0
    for col in columnas_numericas:
        df[col] = df[col].astype(int)
    print("Limpieza y validaciones completadas.")

    # --- PASO 6: TRANSFORMACIONES Y FORMATEO FINAL ---
    print("\n--- PASO 6: Aplicando transformaciones y formato final ---")
    df['NOMBRE COMPLETO'] = df['NOMBRE COMPLETO'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip().str.upper()
    
    replacements_map = {'1118291452':'FANDINO LAYNE ASTRID', '10255294581':'MARTINEZ MUNOZ JOSE MANUEL', '25559122':'RAMIREZ DE CASTRO MARIA ESTELLA'}
    for id_number, new_name in replacements_map.items():
        df.loc[df['NUMERO DE IDENTIFICACION'] == id_number, 'NOMBRE COMPLETO'] = new_name
    
    df['NOMBRE COMPLETO'] = df['NOMBRE COMPLETO'].str.ljust(45)
    df['NUMERO DE IDENTIFICACION'] = df['NUMERO DE IDENTIFICACION'].str.zfill(11)
    print("Formato final aplicado.")

    # --- PASO 7: GUARDAR ARCHIVO FINAL ---
    output_filename = "Data_Marzo_Final_Corregido.xlsx"
    df.to_excel(output_filename, index=False)
    print(f"\n¡PROCESO COMPLETADO! Se ha generado el archivo final '{output_filename}'.")

except Exception as e:
    print(f"\nHa ocurrido un error inesperado en el proceso principal: {e}")