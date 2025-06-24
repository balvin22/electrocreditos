import pandas as pd
import numpy as np
from typing import Dict, List
from src.models.data_models import FinancieroProcessingConfig

class FinancieroProcessor:
    def __init__(self, config: FinancieroProcessingConfig = None):
        self.config = config if config else FinancieroProcessingConfig()
    
    def validate_input_file(self, file_path: str) -> bool:
        """Valida que el archivo Excel contenga todas las hojas requeridas."""
        try:
            with pd.ExcelFile(file_path) as xls:
                sheets = xls.sheet_names
                missing_sheets = [sheet for sheet in self.config.required_sheets if sheet not in sheets]
                
                if missing_sheets:
                    raise ValueError(f"Faltan hojas requeridas: {', '.join(missing_sheets)}")
                    
            return True
        except Exception as e:
            raise ValueError(f"Error al validar archivo: {str(e)}")

    def load_and_filter_data(self, file_path: str) -> Dict[str, pd.DataFrame]:
        """Carga y filtra los datos del archivo Excel según la configuración."""
        try:
            dfs = pd.read_excel(file_path, sheet_name=None)
            
            # Verificar que todas las hojas requeridas estén presentes
            self.validate_input_file(file_path)
            
            filtered_data = {}
            
            for sheet_name, columns in self.config.sheet_columns.items():
                if sheet_name not in dfs:
                    raise ValueError(f"Hoja '{sheet_name}' no encontrada en el archivo")
                
                # Filtrar columnas
                df = dfs[sheet_name][columns].copy()
                
                # Renombrar columnas si es necesario
                if sheet_name in self.config.rename_columns:
                    df.rename(columns=self.config.rename_columns[sheet_name], inplace=True)
                
                filtered_data[sheet_name] = df
            
            return filtered_data
        except Exception as e:
            raise ValueError(f"Error al cargar datos: {str(e)}")

    @staticmethod
    def convert_columns_to_string(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
        """Convierte columnas específicas a tipo string."""
        for col in columns:
            if col in df.columns:
                df[col] = df[col].astype(str)
        return df

    @staticmethod
    def count_accounts(df: pd.DataFrame, id_column: str, count_column_name: str) -> pd.DataFrame:
        """Cuenta las cuentas por ID y devuelve un DataFrame con los resultados."""
        counts = df[id_column].value_counts().reset_index()
        counts.columns = [id_column, count_column_name]
        return counts

    @staticmethod
    def merge_dataframes(main_df: pd.DataFrame, to_merge: pd.DataFrame, 
                        left_on: str, right_on: str, how: str = 'left') -> pd.DataFrame:
        """Realiza un merge entre DataFrames con manejo de errores."""
        try:
            return main_df.merge(to_merge, how=how, left_on=left_on, right_on=right_on)
        except Exception as e:
            raise ValueError(f"Error al fusionar DataFrames: {str(e)}")

    def process_payment_data(self, payment_df: pd.DataFrame, dfs: Dict[str, pd.DataFrame], 
                           payment_type: str) -> pd.DataFrame:
        """Procesa los datos de pagos (Efecty o Bancolombia)."""
        if payment_type not in ['efecty', 'bancolombia']:
            raise ValueError("Tipo de pago debe ser 'efecty' o 'bancolombia'")
       # Convertir columnas de fecha (si existen) para eliminar la parte horaria
       # Formatear fechas en DD/MM/YYYY (para columnas conocidas)
        if payment_type == 'bancolombia' and 'Fecha' in payment_df.columns:
          payment_df['Fecha'] = pd.to_datetime(payment_df['Fecha']).dt.strftime('%d/%m/%Y')
    
        if payment_type == 'efecty' and 'Fecha' in payment_df.columns:
          payment_df['Fecha'] = pd.to_datetime(payment_df['Fecha']).dt.strftime('%d/%m/%Y')
         
        
        merge_conf = self.config.merge_config[payment_type]
        
        # Fusionar con empleados
        result_df = self.merge_dataframes(
            payment_df, 
            dfs['EMPLEADOS ACTUALES'], 
            *merge_conf['empleados']
        )
        
        # Fusionar con AC FS
        result_df = self.merge_dataframes(
            result_df, 
            dfs['AC FS'], 
            *merge_conf['ac_fs']
        )
        
        # Fusionar con AC ARP
        result_df = self.merge_dataframes(
            result_df, 
            dfs['AC ARP'], 
            *merge_conf['ac_arp']
        )
        
        # Llenar valores faltantes
        result_df['EMPLEADO'] = result_df['ESTADO_EMPLEADO'].fillna('NO')
        result_df['CARTERA EN FINANSUEÑOS'] = result_df['FACTURA_FS'].fillna('SIN CARTERA')
        result_df['CARTERA EN ARPESOD'] = result_df['FACTURA_ARP'].fillna('SIN CARTERA')
        
        # Contar cuentas FS
        conteo_fs = self.count_accounts(dfs['AC FS'], 'CEDULA_FS', 'CANTIDAD CUENTAS FS')
        result_df = self.merge_dataframes(
            result_df, 
            conteo_fs, 
            merge_conf['ac_fs'][0], 
            'CEDULA_FS'
        )
        result_df['CANTIDAD CUENTAS FS'] = result_df['CANTIDAD CUENTAS FS'].fillna(0).astype(int)

        # Contar cuentas ARP
        conteo_arp = self.count_accounts(dfs['AC ARP'], 'CEDULA_ARP', 'CANTIDAD CUENTAS ARP')
        result_df = self.merge_dataframes(
            result_df, 
            conteo_arp, 
            merge_conf['ac_arp'][0], 
            'CEDULA_ARP'
        )
        result_df['CANTIDAD CUENTAS ARP'] = result_df['CANTIDAD CUENTAS ARP'].fillna(0).astype(int)

        # --- INICIO DEL BLOQUE LÓGICO CORREGIDO ---

        # 1. AHORA SÍ, calculamos el total de cuentas
        result_df['TOTAL_CUENTAS'] = result_df['CANTIDAD CUENTAS FS'] + result_df['CANTIDAD CUENTAS ARP']

        # 2. Determinamos la factura final con la lógica de "Mas de una cartera"
        factura_original = np.where(
            result_df['CARTERA EN FINANSUEÑOS'] != 'SIN CARTERA',
            result_df['CARTERA EN FINANSUEÑOS'],
            result_df['CARTERA EN ARPESOD']
        )
        result_df['FACTURA FINAL'] = np.where(
            result_df['TOTAL_CUENTAS'] > 1,
            'Mas de una cartera',
            factura_original
        )

        # 3. Guardamos las columnas originales para identificar cada pago de forma única
        columnas_pago_original = list(payment_df.columns)

        # 4. Eliminamos las filas duplicadas generadas por el merge, manteniendo la primera (ya corregida)
        result_df.drop_duplicates(subset=columnas_pago_original, keep='first', inplace=True)
        
        # 5. Preparamos la tabla de saldos para la consulta
        df_fs_saldos = dfs['AC FS'][['FACTURA_FS', 'SALDO_FS','CENTRO_COSTO_FS']].rename(columns={'FACTURA_FS': 'FACTURA', 'SALDO_FS': 'SALDO','CENTRO_COSTO_FS': 'CENTRO COSTO'})
        df_arp_saldos = dfs['AC ARP'][['FACTURA_ARP', 'SALDO_ARP','CENTRO_COSTO_ARP']].rename(columns={'FACTURA_ARP': 'FACTURA', 'SALDO_ARP': 'SALDO', 'CENTRO_COSTO_ARP': 'CENTRO COSTO'})
        df_saldos_unificados = pd.concat([df_fs_saldos, df_arp_saldos], ignore_index=True).drop_duplicates(subset='FACTURA')

        # 6. Fusionamos para traer los saldos al DataFrame ya limpio
        result_df['FACTURA FINAL'] = result_df['FACTURA FINAL'].astype(str)
        result_df = self.merge_dataframes(
            result_df, 
            df_saldos_unificados, 
            'FACTURA FINAL', 
            'FACTURA'
        )

        # 7. Renombramos y limpiamos las columnas de saldo y valor
        result_df = result_df.rename(columns={'SALDO': 'SALDOS'})
        result_df['SALDOS'] = result_df['SALDOS'].fillna(0).astype(float)
        result_df['Valor'] = result_df['Valor'].fillna(0).astype(float)
        
        # 8. Forzamos el saldo a 0 para los casos de "Mas de una cartera", ya que no se puede asignar uno específico
        result_df.loc[result_df['FACTURA FINAL'] == 'Mas de una cartera', 'SALDOS'] = 0
        
        # --- FIN DEL BLOQUE LÓGICO CORREGIDO ---

        # Validar saldo final
        result_df['VALIDACION ULTIMO SALDO'] = np.where(
            (result_df['SALDOS'] - result_df['Valor']) <= 0,
            'pago total',
            (result_df['SALDOS'] - result_df['Valor']).astype(str)
        )
        
        right_key_column = self.config.merge_config[payment_type]['casa_cobranza'][1] 
        # Creamos una versión sin duplicados de la tabla de casa de cobranza
        casa_cobranza_sin_duplicados = dfs['CASA DE COBRANZA'].drop_duplicates(subset=[right_key_column])
        # Fusionar con casa de cobranza
        result_df = self.merge_dataframes(
            result_df, 
            casa_cobranza_sin_duplicados,
            *merge_conf['casa_cobranza']
         )

        result_df['CASA COBRANZA'] = result_df['CASA COBRANZA'].fillna('SIN CASA DE COBRANZA')
    
        
        
        # Fusionar con codeudores
        result_df = self.merge_dataframes(
            result_df, 
            dfs['CODEUDORES'], 
            *merge_conf['codeudores']
         )
        # Asegurarse de que la columna 'CODEUDOR' sea de tipo string
        result_df['CODEUDOR'] = result_df['CODEUDOR'].fillna('SIN CODEUDOR')
        
        # Eliminar columnas innecesarias
        columns_to_drop = [
            'vincedula', 'FACTURA', 'DOCUMENTO_CODEUDOR', 'FACTURA_x', 'FACTURA_y',
            'ESTADO_EMPLEADO', 'CEDULA_FS', 'CEDULA_FS_x', 'CEDULA_ARP_x',
            'CEDULA_FS_y', 'FACTURA_FS', 'CEDULA_ARP', 'CEDULA_ARP_y',
            'FACTURA_ARP', 'SALDO_FS', 'SALDO_ARP', 'CENTRO_COSTO_FS','CENTRO_COSTO_ARP',
         ]
        
        # Eliminar columnas que no existen en el DataFrame
        result_df = result_df.drop(columns=[col for col in columns_to_drop if col in result_df.columns])
        print(result_df)
        
        # Renombrar columnas finales
        result_df = result_df.rename(columns={'FACTURA FINAL': 'Documento Cartera','CENTRO COSTO': 'C. Costo',
                                              'CASA COBRANZA': 'Casa cobranza',
                                              'EMPLEADO': 'Empleado','CANTIDAD CUENTAS ARP': 'Cuentas ARP',
                                              'CANTIDAD CUENTAS FS': 'Cuentas FS'
                                              })
        # Asegurarse de que las columnas estén en el orden correcto
        condiciones = [
           result_df['Casa cobranza'] != 'SIN CASA DE COBRANZA',
           result_df['CODEUDOR'] != 'SIN CODEUDOR',
           result_df['VALIDACION ULTIMO SALDO'] == 'pago total'
           ]
        # Define los valores a asignar según las condiciones
        valores = [
           result_df['Casa cobranza'],
            'Codeudor:'+ result_df['CODEUDOR'],
            'Pago total'
         ]
        
        # Asigna la novedad según las condiciones
        result_df['Novedad'] = np.select(condiciones, valores, default='Sin novedad')
        mascara_sin_cartera_y_codeudor = (
        (result_df['Documento Cartera'] == 'SIN CARTERA') &
        (result_df['Novedad'].str.contains('Codeudor:')))

       # Extrae el valor del codeudor desde la novedad
        result_df.loc[mascara_sin_cartera_y_codeudor, 'Documento Cartera'] = (
        result_df.loc[mascara_sin_cartera_y_codeudor, 'Novedad']
       .str.extract(r'Codeudor:\s*(\S+)')[0])
        
        condiciones_empresa = [
         # Caso 1: Documento cartera NO es 'SIN CARTERA' y empieza con 'DF'
          (result_df['Documento Cartera'] != 'SIN CARTERA') & result_df['Documento Cartera'].str.startswith('DF'),
    
         # Caso 2: Documento cartera NO es 'SIN CARTERA' y NO empieza con 'DF'
          (result_df['Documento Cartera'] != 'SIN CARTERA') & ~result_df['Documento Cartera'].str.startswith('DF'),
    
         # Caso 3: Documento cartera es 'SIN CARTERA', pero CODEUDOR empieza con 'DF'
          (result_df['Documento Cartera'] == 'SIN CARTERA') & (result_df['CODEUDOR'] != 'SIN CODEUDOR') & result_df['CODEUDOR'].str.startswith('DF'),
    
         # Caso 4: Documento cartera es 'SIN CARTERA', CODEUDOR no empieza con 'DF'
          (result_df['Documento Cartera'] == 'SIN CARTERA') & (result_df['CODEUDOR'] != 'SIN CODEUDOR') & ~result_df['CODEUDOR'].str.startswith('DF'),
           ]
        valores_empresa = [
          'Finansueños',  # Caso 1
          'Arpesod',      # Caso 2
          'Finansueños',  # Caso 3
          'Arpesod'       # Caso 4
          ]
        
        # Asigna la columna 'Empresa' según las condiciones
        result_df['Empresa'] = np.select(condiciones_empresa, valores_empresa, default='')
        
        
        result_df['Valor Aplicar'] = np.where(
           (result_df['VALIDACION ULTIMO SALDO'] == 'pago total') & (result_df['SALDOS'] != 0),
           result_df['SALDOS'],
           result_df['Valor']
           )
        
        # Calcular el valor de anticipos
        diferencia_aprovechamiento= result_df['Valor'] - result_df['SALDOS']
        result_df['Valor Aprovechamientos'] = np.where(
        (diferencia_aprovechamiento > 0) & (diferencia_aprovechamiento <= 10000),
        diferencia_aprovechamiento,
         0  
         )
        
        # Calcular el valor de anticipos
        diferencia_anticipo = result_df['Valor'] -result_df['Valor Aplicar']
        result_df['Valor Anticipos'] = np.where(
            diferencia_anticipo >= 10000,
            diferencia_anticipo,0
        )
        
        return result_df
    
    # Guardar resultados en un archivo Excel con formato y resaltado condicional
    def save_results_to_excel(self, df_bancolombia: pd.DataFrame, df_efecty: pd.DataFrame, 
                              output_file: str) -> bool:
        """Guarda los resultados en un archivo Excel con manejo de errores."""
        COLUMN_ORDER_EFECTY = [
        'No', 'Identificación', 'Valor', 'N° de Autorización', 'Fecha',
        'Documento Cartera', 'C. Costo', 'Empresa', 'Valor Aplicar',
        'Valor Anticipos', 'Valor Aprovechamientos', 'Casa cobranza',
        'Empleado', 'Novedad','Cuentas ARP', 'Cuentas FS','SALDOS','VALIDACION ULTIMO SALDO']
        COLUMN_ORDER_BANCOLOMBIA = [
        'No.', 'Fecha', 'Detalle 1', 'Detalle 2', 'Referencia 1', 'Referencia 2',
        'Valor', 'Documento Cartera', 'C. Costo', 'Empresa', 'Valor Aplicar',
        'Valor Anticipos', 'Valor Aprovechamientos', 'Casa cobranza',
        'Empleado', 'Novedad', 'Cuentas ARP', 'Cuentas FS','SALDOS','VALIDACION ULTIMO SALDO']

        def resaltar_condicional(row):
          arp = row['Cuentas ARP']
          fs = row['Cuentas FS']
    
         # Condiciones de resaltado
          if (arp >= 2) or (fs >= 2):
            return ['background-color: lightcoral', 'background-color: lightcoral']
          elif arp == 1 and fs == 1:
            return ['background-color: lightcoral', 'background-color: lightcoral']
          else:
            return [''] *len(row)
        
        def resaltar_empleados_y_duplicados(df):
        # Crear un DataFrame de estilos vacío
          styles = pd.DataFrame('', index=df.index, columns=df.columns)
        
        # Resaltar celda donde Empleado == 'SI'
          empleado_mask = df['Empleado'].str.upper().str.strip() == 'SI'
          styles.loc[empleado_mask, 'Empleado'] = 'background-color: lightblue'
        
        # Resaltar documentos de cartera duplicados
          dup_mask = df.duplicated('Documento Cartera', keep=False) & (df['Documento Cartera'] != 'SIN CARTERA')
          styles.loc[dup_mask, 'Documento Cartera'] = 'background-color: yellow'
        
          return styles
        
        def limpiar_referencia(valor):
            try:
                if pd.isnull(valor) or valor == '':
                    return ''
                return str(int(float(valor)))
            except (ValueError, TypeError):
                return ''
        try:
            # Limpiar campo 'Referencia 2' si existe en el DataFrame de Bancolombia
            if 'Referencia 1' in df_bancolombia.columns:
                df_bancolombia['Referencia 1'] = df_bancolombia['Referencia 1'].apply(limpiar_referencia)
            if 'Referencia 2' in df_bancolombia.columns:
                df_bancolombia['Referencia 2'] = df_bancolombia['Referencia 2'].apply(limpiar_referencia)
            
            df_efecty = df_efecty.reindex(columns=COLUMN_ORDER_EFECTY)
            df_bancolombia = df_bancolombia.reindex(columns=COLUMN_ORDER_BANCOLOMBIA)
            
            styled_efecty = (df_efecty.style
                         .apply(resaltar_condicional, axis=1, subset=['Cuentas ARP', 'Cuentas FS'])
                         .apply(resaltar_empleados_y_duplicados, axis=None))
        
            styled_bancolombia = (df_bancolombia.style
                             .apply(resaltar_condicional, axis=1, subset=['Cuentas ARP', 'Cuentas FS'])
                             .apply(resaltar_empleados_y_duplicados, axis=None))

            # Escribir ambos DataFrames a un archivo Excel
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                styled_bancolombia.to_excel(writer, sheet_name='Bancolombia', index=False)
                styled_efecty.to_excel(writer, sheet_name='Efecty', index=False)

            return True
        except Exception as e:
            raise ValueError(f"Error al guardar archivo Excel: {str(e)}")
