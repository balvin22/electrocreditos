import pandas as pd
import numpy as np

class ReportService:
    """
    Servicio encargado de toda la lógica de negocio para procesar y consolidar
    los reportes de cartera.
    """
    def __init__(self, config):
        self.config = config

    def _load_dataframes(self, file_paths):
        """Lee todos los archivos de Excel y los convierte en DataFrames según la configuración."""
        dataframes_por_tipo = {key: [] for key in self.config.keys()}
        print("--- 🔄 Iniciando lectura de archivos ---")
        for ruta_archivo in file_paths:
            try:
                nombre_archivo = ruta_archivo.split('/')[-1]
                tipo_archivo_actual = self._get_file_type(nombre_archivo)
                
                if not tipo_archivo_actual:
                    print(f"⚠️  Archivo '{nombre_archivo}' omitido.")
                    continue

                config_actual = self.config[tipo_archivo_actual]
                
                if "sheets" in config_actual:
                    # Lógica para archivos con múltiples hojas
                    xls = pd.ExcelFile(ruta_archivo, engine='openpyxl')
                    for sheet_config in config_actual["sheets"]:
                        if sheet_config["sheet_name"] in xls.sheet_names:
                            df_hoja = pd.read_excel(xls, sheet_name=sheet_config["sheet_name"])
                            df_hoja.columns = df_hoja.columns.str.strip()
                            columnas_a_usar = [col for col in sheet_config["usecols"] if col in df_hoja.columns]
                            df_filtrado = df_hoja[columnas_a_usar].rename(columns=sheet_config["rename_map"])
                            dataframes_por_tipo[tipo_archivo_actual].append({"data": df_filtrado, "config": sheet_config})
                elif "new_names" in config_actual:
                    # Lógica para archivos con formato especial (ej. MATRIZ_CARTERA)
                    df = pd.read_excel(ruta_archivo, header=config_actual.get("header"), skiprows=config_actual.get("skiprows"), names=config_actual.get("new_names"))
                    dataframes_por_tipo[tipo_archivo_actual].append(df)
                else:
                    # Lógica para archivos estándar
                    engine = 'xlrd' if ruta_archivo.upper().endswith('.XLS') else 'openpyxl'
                    df = pd.read_excel(ruta_archivo, engine=engine)
                    df.columns = df.columns.str.strip()
                    columnas_a_usar = [col for col in config_actual["usecols"] if col in df.columns]
                    df_filtrado = df[columnas_a_usar]
                    if tipo_archivo_actual == "R03":
                        df_filtrado = df_filtrado.replace('.', 'SIN CODEUDOR').fillna('SIN CODEUDOR')
                    df_renombrado = df_filtrado.rename(columns=config_actual["rename_map"])
                    dataframes_por_tipo[tipo_archivo_actual].append(df_renombrado)
                
                print(f"✅ Archivo '{nombre_archivo}' procesado como tipo '{tipo_archivo_actual}'.")
            except Exception as e:
                print(f"❌ Error procesando el archivo '{ruta_archivo}': {e}")
        
        return dataframes_por_tipo

    def _get_file_type(self, filename):
        """Determina el tipo de archivo basado en su nombre y las claves de configuración."""
        nombre_base = filename.split('.')[0].upper().replace(" ", "_")
        palabras_en_nombre = set(nombre_base.split('_'))
        for tipo in self.config.keys():
            palabras_en_clave = set(tipo.split('_'))
            if palabras_en_clave.issubset(palabras_en_nombre):
                return tipo
        return None

    def _create_credit_key(self, df):
        """Crea una llave única 'Credito' combinando Tipo y Número de crédito."""
        if 'Numero_Credito' in df.columns and 'Tipo_Credito' in df.columns:
            df['Numero_Credito'] = pd.to_numeric(df['Numero_Credito'], errors='coerce').astype('Int64')
            df['Credito'] = df['Tipo_Credito'].astype(str) + '-' + df['Numero_Credito'].astype(str).str.replace('<NA>', '', regex=False)
        return df

    def _calculate_balances(self, reporte_df, fnz003_df):
        """Calcula los diferentes saldos y los agrega al reporte."""
        print("📊 Calculando saldos...")
        reporte_df['Saldo_Capital'] = np.where(reporte_df['Empresa'] == 'ARPESOD', reporte_df.get('Saldo_Factura'), np.nan)
        if not fnz003_df.empty:
            fnz003_df['Credito'] = fnz003_df['Credito'].astype(str)
            mapa_capital = fnz003_df[fnz003_df['Concepto'].isin(['CAPITAL', 'ABONO DIF TASA'])].groupby('Credito')['Saldo'].sum()
            mapa_avales = fnz003_df[fnz003_df['Concepto'] == 'AVAL'].groupby('Credito')['Saldo'].sum()
            mapa_interes = fnz003_df[fnz003_df['Concepto'] == 'INTERES CORRIENTE'].groupby('Credito')['Saldo'].sum()
            
            reporte_df['Saldo_Capital'] = reporte_df['Credito'].map(mapa_capital).combine_first(reporte_df['Saldo_Capital'])
            reporte_df['Saldo_Avales'] = reporte_df['Credito'].map(mapa_avales)
            reporte_df['Saldo_Interes_Corriente'] = reporte_df['Credito'].map(mapa_interes)
        
        reporte_df['Saldo_Capital'] = pd.to_numeric(reporte_df['Saldo_Capital'], errors='coerce').fillna(0).astype(int)
        reporte_df['Saldo_Avales'] = np.where(reporte_df['Empresa'] == 'FINANSUEÑOS', reporte_df.get('Saldo_Avales').fillna(0).astype(int), 'NO APLICA')
        reporte_df['Saldo_Interes_Corriente'] = np.where(reporte_df['Empresa'] == 'FINANSUEÑOS', reporte_df.get('Saldo_Interes_Corriente').fillna(0).astype(int), 'NO APLICA')
        return reporte_df

    def _finalize_report(self, reporte_df):
        """Realiza la limpieza y transformaciones finales del reporte."""
        print("🧹 Realizando transformaciones y limpieza final...")
        cols_a_borrar = [col for col in reporte_df.columns if any(sufijo in col for sufijo in ['_Venc', '_R03', '_Analisis'])]
        reporte_df = reporte_df.drop(columns=cols_a_borrar, errors='ignore')
        
        print("📅 Formateando fechas a DD/MM/YY...")
        for col_fecha in ['Fecha_Vencimiento', 'Fecha_Facturada']:
            if col_fecha in reporte_df.columns:
                reporte_df[col_fecha] = pd.to_datetime(reporte_df[col_fecha], errors='coerce').dt.strftime('%d/%m/%y').fillna('')
        return reporte_df


    def generate_consolidated_report(self, file_paths):
        """
        Orquesta todo el proceso de ETL: cargar, transformar y consolidar los datos.
        """
        dataframes_por_tipo = self._load_dataframes(file_paths)

        # Función auxiliar para concatenar DataFrames de forma segura
        def safe_concat(items):
            if not items: return pd.DataFrame()
            df_list = [item["data"] if isinstance(item, dict) else item for item in items]
            return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

        analisis_df = safe_concat(dataframes_por_tipo.get("ANALISIS", []))
        r91_df = safe_concat(dataframes_por_tipo.get("R91", []))
        vencimientos_df = safe_concat(dataframes_por_tipo.get("VENCIMIENTOS", []))
        r03_df = safe_concat(dataframes_por_tipo.get("R03", []))
        crtmp_df = safe_concat(dataframes_por_tipo.get("CRTMPCONSULTA1", []))
        fnz003_df = safe_concat(dataframes_por_tipo.get("FNZ003", []))
        matriz_cartera_df = safe_concat(dataframes_por_tipo.get("MATRIZ_CARTERA", []))
        
        if r91_df.empty:
            print("\n❌ No se encontraron archivos R91 para construir el reporte base. Proceso detenido.")
            return None

        print("\n🔗 Creando el reporte base y estandarizando llaves...")
        reporte_final = self._create_credit_key(r91_df.copy())
        analisis_df = self._create_credit_key(analisis_df)
        vencimientos_df = self._create_credit_key(vencimientos_df)
        crtmp_df = self._create_credit_key(crtmp_df)

        reporte_final['Empresa'] = np.where(reporte_final['Tipo_Credito'] == 'DF', 'FINANSUEÑOS', 'ARPESOD')
        
        # Estandarizar llaves antes de cruzar
        for df, col_name in [(reporte_final, 'Cedula_Cliente'), (vencimientos_df, 'Cedula_Cliente'), (r03_df, 'Cedula_Cliente')]:
             if col_name in df.columns:
                df[col_name] = df[col_name].astype(str).str.strip()

        # --- CRUCE SECUENCIAL DE DATOS ---
        print("🔍 Uniendo información de todos los archivos...")
        if not analisis_df.empty:
            reporte_final = pd.merge(reporte_final, analisis_df.drop_duplicates('Credito'), on='Credito', how='left', suffixes=('', '_Analisis'))
        if not vencimientos_df.empty:
            reporte_final = pd.merge(reporte_final, vencimientos_df.drop_duplicates('Credito'), on='Credito', how='left', suffixes=('', '_Venc'))
        if not r03_df.empty:
            reporte_final = pd.merge(reporte_final, r03_df.drop_duplicates('Cedula_Cliente'), on='Cedula_Cliente', how='left', suffixes=('', '_R03'))
        if not matriz_cartera_df.empty:
            reporte_final['Zona'] = reporte_final['Zona'].astype(str).str.strip()
            matriz_cartera_df['Zona'] = matriz_cartera_df['Zona'].astype(str).str.strip()
            reporte_final = pd.merge(reporte_final, matriz_cartera_df.drop_duplicates('Zona'), on='Zona', how='left')
        if not crtmp_df.empty:
            reporte_final = pd.merge(reporte_final, crtmp_df[['Credito', 'Correo', 'Fecha_Facturada']].drop_duplicates('Credito'), on='Credito', how='left')

        # Cruce de datos de ASESORES
        for item in dataframes_por_tipo.get("ASESORES", []):
            info_df = item["data"]
            merge_key = item["config"]["merge_on"]
            if not info_df.empty and merge_key in reporte_final.columns:
                info_df[merge_key] = info_df[merge_key].astype(str).str.strip()
                reporte_final[merge_key] = reporte_final[merge_key].astype(str).str.strip()
                reporte_final = pd.merge(reporte_final, info_df.drop_duplicates(merge_key), on=merge_key, how='left')
        
        reporte_final = self._calculate_balances(reporte_final, fnz003_df)
        reporte_final = self._finalize_report(reporte_final)

        return reporte_final