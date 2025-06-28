import pandas as pd
import numpy as np

# --- PASO 1: DEFINICIONES INICIALES ---
print("--- PASO 1: Definiendo parámetros iniciales ---")
input_file_path = 'c:/Users/usuario/Desktop/Reporte LV/DATA MARZO FS.TXT'
ruta_archivo_correcciones = 'c:/Users/usuario/Desktop/Reporte LV/Cédulas a revisar.xlsx'

# Definición limpia de las posiciones y nombres para evitar caracteres ocultos
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
    # --- PASO 2: LECTURA Y PREPARACIÓN INICIAL ---
    print("\n--- PASO 2: Leyendo y preparando el archivo principal ---")
    df = pd.read_fwf(
        input_file_path, colspecs=colspecs, names=names, encoding='latin-1',
        skiprows=1, skipfooter=1, engine='python'
    )
    
    # *** LÍNEA DE SEGURIDAD: Limpiamos cualquier espacio extra en los nombres de las columnas ***
    df.columns = df.columns.str.strip()

    print("--- DIAGNÓSTICO: Columnas justo después de la lectura ---")
    print(df.columns.tolist())

    df['NUMERO DE IDENTIFICACION'] = df['NUMERO DE IDENTIFICACION'].astype(str).str.strip()

    # --- PASO 3: CORRECCIONES DE DATOS ---
    print("\n--- PASO 3: Iniciando correcciones desde archivo Excel ---")
    # (El código de corrección va aquí como antes, sin cambios en su lógica)
    try:
        # A. Corregir Cédulas
        df_cedulas = pd.read_excel(ruta_archivo_correcciones, sheet_name='Cedulas a corregir')
        mapa_cedulas = pd.Series(df_cedulas['CEDULA CORRECTA'].astype(str).str.strip().values, index=df_cedulas['CEDULA MAL'].astype(str).str.strip()).to_dict()
        df['NUMERO DE IDENTIFICACION'] = df['NUMERO DE IDENTIFICACION'].replace(mapa_cedulas)

        # B. Actualizar campos 'CORREGIR'
        df_vinculado = pd.read_excel(ruta_archivo_correcciones, sheet_name='Vinculado').set_index(pd.read_excel(ruta_archivo_correcciones, sheet_name='Vinculado')['CODIGO'].astype(str).str.strip())
        mapa_vinc = {'NOMBRE COMPLETO':'NOMBRE', 'DIRECCION DE CORRESPONDENCIA':'DIRECCI', 'CORREO ELECTRONICO':'VINEMAIL', 'CELULAR':'TELEFONO'}
        for col_df, col_vinc in mapa_vinc.items():
            if col_df in df.columns:
                mascara = df[col_df].astype(str).str.strip().str.contains('CORREGIR', case=False, na=False)
                if mascara.any():
                    ids_a_buscar = df.loc[mascara, 'NUMERO DE IDENTIFICACION']
                    valores_nuevos = ids_a_buscar.map(df_vinculado[col_vinc])
                    df[col_df] = valores_nuevos.combine_first(df[col_df])
            else:
                print(f"ADVERTENCIA: La columna '{col_df}' no se encontró para la actualización de 'CORREGIR'.")

        # C. Actualizar Tipos de Identificación
        df['TIPO DE IDENTIFICACION'] = 1
        df_tipos = pd.read_excel(ruta_archivo_correcciones, sheet_name='Tipos de identificacion')
        mapa_tipos = pd.Series(df_tipos['CODIGO DATA'].values, index=df_tipos['CEDULA CORRECTA'].astype(str).str.strip()).to_dict()
        df['TIPO DE IDENTIFICACION'] = df['NUMERO DE IDENTIFICACION'].map(mapa_tipos).combine_first(df['TIPO DE IDENTIFICACION'])
    except Exception as e:
        print(f"   ¡ERROR DURANTE CORRECCIONES EXCEL! Detalle: {e}")

    # --- PASO 4 y 5 (se mantienen igual) ---
    # ... (El código de los pasos 4 y 5 va aquí sin cambios) ...
    
    # --- PASO 6: TRANSFORMACIONES Y FORMATEO FINAL ---
    print("\n--- PASO 6: Aplicando formato final ---")

    # *** LÍNEA DE DIAGNÓSTICO: Imprimimos las columnas justo antes del bloque que falla ***
    print("\n--- DIAGNÓSTICO: Columnas ANTES de formatear la ciudad ---")
    print(df.columns.tolist())

    # A. Formato de CIUDAD DE CORRESPONDENCIA
    print("   Formateando 'CIUDAD DE CORRESPONDENCIA'...")
    col_ciudad = 'CIUDAD DE CORRESPONDENCIA'
    df[col_ciudad] = df[col_ciudad].astype(str).fillna('')
    mascara_reemplazo_ciudad = (df[col_ciudad].str.strip() == '') | (df[col_ciudad].str.strip().str.upper() == 'N/A') | (df[col_ciudad].str.strip() == '0')
    df.loc[mascara_reemplazo_ciudad, col_ciudad] = 'POPAYAN'
    df[col_ciudad] = df[col_ciudad].str.ljust(20)

    # B. Resto de los formatos finales
    print("   Aplicando otros formatos finales...")
    # (El resto del código de formato va aquí)
    df['NOMBRE COMPLETO'] = df['NOMBRE COMPLETO'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip().str.upper()
    replacements_map = {'1118291452':'FANDINO LAYNE ASTRID', '10255294581':'MARTINEZ MUNOZ JOSE MANUEL', '25559122':'RAMIREZ DE CASTRO MARIA ESTELLA'}
    for id_number, new_name in replacements_map.items():
        df.loc[df['NUMERO DE IDENTIFICACION'] == id_number, 'NOMBRE COMPLETO'] = new_name

    col_celular = 'CELULAR'
    df[col_celular] = df[col_celular].astype(str).str.replace(r'\D', '', regex=True)
    es_fijo_valido = (df[col_celular].str.len() == 7)
    es_celular_valido = (df[col_celular].str.len() == 10) & (df[col_celular].str.startswith('3'))
    df.loc[~(es_fijo_valido | es_celular_valido), col_celular] = '0'
    
    col_email = 'CORREO ELECTRONICO'
    df[col_email] = df[col_email].astype(str).str.strip()
    placeholders_a_eliminar = ['CORREGIR', 'PENDIENTE', 'NOTIENE', 'SINC', 'NN@', 'AAA@']
    for placeholder in placeholders_a_eliminar:
        df.loc[df[col_email].str.contains(placeholder, case=False, na=False), col_email] = ''
    email_regex_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    df.loc[~df[col_email].str.match(email_regex_pattern, na=False), col_email] = ''
    
    df['NOMBRE COMPLETO'] = df['NOMBRE COMPLETO'].str.ljust(45)
    df['DIRECCION DE CORRESPONDENCIA'] = df['DIRECCION DE CORRESPONDENCIA'].astype(str).str.ljust(60)
    df['CORREO ELECTRONICO'] = df[col_email].str.ljust(60)
    df['CELULAR'] = df[col_celular].str.zfill(12)
    df['NUMERO DE IDENTIFICACION'] = df['NUMERO DE IDENTIFICACION'].str.zfill(11)
    print("Formato final aplicado.")

    # --- PASO 7: GUARDAR ARCHIVO FINAL ---
    # ... (El código del paso 7 va aquí) ...

except Exception as e:
    print(f"\nHa ocurrido un error inesperado en el proceso principal: {e}")