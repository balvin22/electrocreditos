import pandas as pd
from typing import Dict, List
from src.models.data_models import AnticiposConfig
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

class AnticiposOnlineProcessor:
    def __init__(self, config: AnticiposConfig = None):
        self.config = config if config else AnticiposConfig()
        
    def validate_input_file(self, file_path:str) -> bool:
        try:
            with pd.ExcelFile(file_path) as xls:
                sheets = xls.sheet_names
                missing_sheets = [sheet for sheet in self.config.required_sheets if sheet not in sheets]
                
                if missing_sheets:
                    raise ValueError(f"Faltan hojas requeridas: \n {', '.join(missing_sheets)}")
            return True
        except Exception as e:
            raise ValueError(f"Error al validar el archivo: \n {str(e)}")
        
    def load_and_filter_data(self, file_path: str) -> Dict[str, pd.DataFrame]:
        try:
            #Cargar el archivo Excel y validar las hojas
            dfs = pd.read_excel(file_path,sheet_name=None)
            #Validar las hojas requeridas
            self.validate_input_file(file_path)
            filtered_data={}
            
            for sheet_name,columns in self.config.sheet_columns.items():
                if sheet_name not in dfs:
                    raise ValueError(f"Hoja {sheet_name} no encontrada en el archivo.\n")
                df = dfs[sheet_name][columns].copy()
                # Renombrar columnas si es necesario
                if sheet_name in self.config.rename_columns:
                    df.rename(columns= self.config.rename_columns[sheet_name], inplace=True)
                filtered_data[sheet_name] = df
            return filtered_data
        except Exception as e:
            raise ValueError(f"Error al cargar datos:\n {str(e)}")
    @staticmethod
    def convert_columns_to_string(df: pd.DataFrame, columns: List[str])-> pd.DataFrame:
        for col in columns:
            if col in df.columns:
                df[col] = df[col].astype(str)
                return df
    @staticmethod
    def merge_dataframes(main_df:pd.DataFrame,to_merge:pd.DataFrame, left_on:str, right_on:str,
                         how:str = 'left') -> pd.DataFrame:
        try:
            return main_df.merge(to_merge, how = how, left_on=left_on, right_on=right_on)
        except Exception as e:
            raise ValueError(f"Error al fusionar DataFrames: {str(e)}")
    @staticmethod
    def save_formatted_excel(final_df: pd.DataFrame, output_filename: str):
        """
        Guarda un DataFrame en un archivo Excel con formato profesional.
        - Cabeceras con fondo de color y texto en negrita.
        - Autoajuste del ancho de las columnas.
        - Bordes en todas las celdas.
        """
        try:
            with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
                # 1. Guardar el DataFrame en una hoja específica sin el índice
                final_df.to_excel(writer, sheet_name='Reporte Consolidado', index=False)

                # 2. Acceder al libro y la hoja de trabajo de openpyxl
                workbook = writer.book
                worksheet = writer.sheets['Reporte Consolidado']

                # 3. Definir los estilos a utilizar
                header_font = Font(bold=True, color="FFFFFF", name='Calibri')
                header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                header_alignment = Alignment(horizontal='center', vertical='center')
                
                thin_border_side = Side(style='thin', color='000000')
                cell_border = Border(left=thin_border_side, 
                                     right=thin_border_side, 
                                     top=thin_border_side, 
                                     bottom=thin_border_side)

                # 4. Aplicar formato a la fila de la cabecera
                for cell in worksheet[1]: # La fila 1 es la cabecera
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment
                    cell.border = cell_border

                # 5. Aplicar formato al resto de las celdas y autoajustar columnas
                for row in worksheet.iter_rows(min_row=2): # Empezar desde la segunda fila
                    for cell in row:
                        cell.border = cell_border

                # 6. Autoajustar el ancho de las columnas
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
        
        except Exception as e:
            # Levantar una excepción si algo sale mal durante el guardado
            raise ValueError(f"Error al guardar el archivo Excel formateado:\n {str(e)}")
  
                
                
           
       


