import os
import pandas as pd
import numpy as np
from typing import Dict, Tuple, List
from .data_models import DataProcessingConfig

class DataProcessor:
    def __init__(self, config: DataProcessingConfig = None):
        self.config = config if config else DataProcessingConfig()
    
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
        
        # Determinar factura final
        result_df['FACTURA FINAL'] = np.where(
            result_df['CARTERA EN FINANSUEÑOS'] != 'SIN CARTERA',
            result_df['CARTERA EN FINANSUEÑOS'],
            result_df['CARTERA EN ARPESOD']
        )
        
        # Unificar saldos
        df_fs_saldos = dfs['AC FS'][['FACTURA_FS', 'SALDO_FS']].rename(columns={'FACTURA_FS': 'FACTURA', 'SALDO_FS': 'SALDO'})
        df_arp_saldos = dfs['AC ARP'][['FACTURA_ARP', 'SALDO_ARP']].rename(columns={'FACTURA_ARP': 'FACTURA', 'SALDO_ARP': 'SALDO'})
        df_saldos_unificados = pd.concat([df_fs_saldos, df_arp_saldos], ignore_index=True).drop_duplicates(subset='FACTURA')
        
        result_df['FACTURA FINAL'] = result_df['FACTURA FINAL'].astype(str)
        result_df = self.merge_dataframes(
            result_df, 
            df_saldos_unificados, 
            'FACTURA FINAL', 
            'FACTURA'
        )
        result_df = result_df.rename(columns={'SALDO': 'SALDOS'})
        result_df['SALDOS'] = result_df['SALDOS'].fillna(0).astype(float)
        result_df['Valor'] = result_df['Valor'].fillna(0).astype(float)
        
        # Validar saldo final
        result_df['VALIDACION ULTIMO SALDO'] = np.where(
            (result_df['SALDOS'] - result_df['Valor']) <= 0,
            'pago total',
            (result_df['SALDOS'] - result_df['Valor']).astype(str)
        )
        
        # Fusionar con casa de cobranza
        result_df = self.merge_dataframes(
            result_df, 
            dfs['CASA DE COBRANZA'], 
            *merge_conf['casa_cobranza']
        )
        result_df['CASA COBRANZA'] = result_df['CASA COBRANZA'].fillna('SIN CASA DE COBRANZA')
        
        # Fusionar con codeudores
        result_df = self.merge_dataframes(
            result_df, 
            dfs['CODEUDORES'], 
            *merge_conf['codeudores']
        )
        result_df['CODEUDOR'] = result_df['CODEUDOR'].fillna('SIN CODEUDOR')
        
        # Eliminar columnas innecesarias
        columns_to_drop = [
            'vincedula', 'FACTURA', 'DOCUMENTO_CODEUDOR', 'FACTURA_x', 'FACTURA_y',
            'ESTADO_EMPLEADO', 'CEDULA_FS', 'CEDULA_FS_x', 'CEDULA_ARP_x',
            'CEDULA_FS_y', 'FACTURA_FS', 'CEDULA_ARP', 'CEDULA_ARP_y',
            'FACTURA_ARP', 'SALDO_FS', 'SALDO_ARP'
        ]
        
        result_df = result_df.drop(columns=[col for col in columns_to_drop if col in result_df.columns])
        
        return result_df

    def save_results_to_excel(self, df_bancolombia: pd.DataFrame, df_efecty: pd.DataFrame, 
                            output_file: str) -> bool:
        """Guarda los resultados en un archivo Excel con manejo de errores."""
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df_bancolombia.to_excel(writer, sheet_name='Bancolombia', index=False)
                df_efecty.to_excel(writer, sheet_name='Efecty', index=False)
            
            # Verificar que el archivo se creó correctamente
            if os.path.exists(output_file):
                with pd.ExcelFile(output_file) as xls:
                    if not all(sheet in xls.sheet_names for sheet in ['Bancolombia', 'Efecty']):
                        raise ValueError("No se crearon todas las hojas en el archivo de salida")
                
                return True
            else:
                raise ValueError("No se pudo crear el archivo de salida")
                
        except Exception as e:
            raise ValueError(f"Error al guardar resultados: {str(e)}")