import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
import pandas as pd
import numpy as np
import os
from typing import Dict, Tuple, Optional, List
from dataclasses import dataclass

# ==============================================
# CLASES DE DATOS PARA MEJOR ESTRUCTURACIÓN
# ==============================================

@dataclass
class AppConfig:
    title: str = "Procesador de Reportes Financieros"
    geometry: str = "700x400"
    bg_color: str = "#f0f0f0"
    accent_color: str = "#4b6cb7"
    secondary_color: str = "#2d3747"
    text_color: str = "#333333"
    output_filename: str = "reporte_financiero.xlsx"

@dataclass
class DataProcessingConfig:
    required_sheets: List[str] = None
    sheet_columns: Dict[str, Dict[str, List[str]]] = None
    rename_columns: Dict[str, Dict[str, str]] = None
    merge_config: Dict[str, Dict] = None

    def __post_init__(self):
        if self.required_sheets is None:
            self.required_sheets = [
                'AC FS', 'AC ARP', 'CODEUDORES', 
                'CASA DE COBRANZA', 'EMPLEADOS ACTUALES', 
                'PAGOS BANCOLOMBIA', 'PAGOS EFECTY'
            ]
        
        if self.sheet_columns is None:
            self.sheet_columns = {
                'AC FS': ['CEDULA', 'FACTURA', 'saldofac', 'ccosto'],
                'AC ARP': ['CEDULA', 'FACTURA', 'saldofac', 'ccosto'],
                'CODEUDORES': ['CODEUDOR', 'FACTURA'],
                'CASA DE COBRANZA': ['FACTURA', 'cobra'],
                'EMPLEADOS ACTUALES': ['vincedula', 'ACTIVO'],
                'PAGOS BANCOLOMBIA': ['No.', 'Fecha', 'Detalle 1', 'Detalle 2', 'Referencia 1', 'Referencia 2', 'Valor'],
                'PAGOS EFECTY': ['No', 'Identificación', 'Valor', 'N° de Autorización', 'Fecha']
            }
        
        if self.rename_columns is None:
            self.rename_columns = {
                'AC FS': {'CEDULA': 'CEDULA_FS', 'FACTURA': 'FACTURA_FS', 'saldofac': 'SALDO_FS', 'ccosto': 'CENTRO COSTO FS'},
                'AC ARP': {'CEDULA': 'CEDULA_ARP', 'FACTURA': 'FACTURA_ARP', 'saldofac': 'SALDO_ARP', 'ccosto': 'CENTRO COSTO ARP'},
                'CODEUDORES': {'FACTURA': 'CODEUDOR', 'CODEUDOR': 'DOCUMENTO_CODEUDOR'},
                'CASA DE COBRANZA': {'cobra': 'CASA COBRANZA'},
                'EMPLEADOS ACTUALES': {'ACTIVO': 'ESTADO_EMPLEADO'}
            }
        
        if self.merge_config is None:
            self.merge_config = {
                'efecty': {
                    'empleados': ('Identificación', 'vincedula'),
                    'ac_fs': ('Identificación', 'CEDULA_FS'),
                    'ac_arp': ('Identificación', 'CEDULA_ARP'),
                    'casa_cobranza': ('FACTURA FINAL', 'FACTURA'),
                    'codeudores': ('Identificación', 'DOCUMENTO_CODEUDOR')
                },
                'bancolombia': {
                    'empleados': ('Referencia 1', 'vincedula'),
                    'ac_fs': ('Referencia 1', 'CEDULA_FS'),
                    'ac_arp': ('Referencia 1', 'CEDULA_ARP'),
                    'casa_cobranza': ('CARTERA EN ARPESOD', 'FACTURA'),
                    'codeudores': ('Referencia 1', 'DOCUMENTO_CODEUDOR')
                }
            }

# ==============================================
# FUNCIONES DE UTILIDAD
# ==============================================

def validate_input_file(file_path: str, required_sheets: List[str]) -> bool:
    """Valida que el archivo Excel contenga todas las hojas requeridas."""
    try:
        with pd.ExcelFile(file_path) as xls:
            sheets = xls.sheet_names
            missing_sheets = [sheet for sheet in required_sheets if sheet not in sheets]
            
            if missing_sheets:
                raise ValueError(f"Faltan hojas requeridas: {', '.join(missing_sheets)}")
                
        return True
    except Exception as e:
        raise ValueError(f"Error al validar archivo: {str(e)}")

def load_and_filter_data(file_path: str, config: DataProcessingConfig) -> Dict[str, pd.DataFrame]:
    """Carga y filtra los datos del archivo Excel según la configuración."""
    try:
        dfs = pd.read_excel(file_path, sheet_name=None)
        
        # Verificar que todas las hojas requeridas estén presentes
        validate_input_file(file_path, config.required_sheets)
        
        filtered_data = {}
        
        for sheet_name, columns in config.sheet_columns.items():
            if sheet_name not in dfs:
                raise ValueError(f"Hoja '{sheet_name}' no encontrada en el archivo")
            
            # Filtrar columnas
            df = dfs[sheet_name][columns].copy()
            
            # Renombrar columnas si es necesario
            if sheet_name in config.rename_columns:
                df.rename(columns=config.rename_columns[sheet_name], inplace=True)
                
            
            filtered_data[sheet_name] = df
        
        return filtered_data
    except Exception as e:
        raise ValueError(f"Error al cargar datos: {str(e)}")

def convert_columns_to_string(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    """Convierte columnas específicas a tipo string."""
    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(str)
    return df

def count_accounts(df: pd.DataFrame, id_column: str, count_column_name: str) -> pd.DataFrame:
    """Cuenta las cuentas por ID y devuelve un DataFrame con los resultados."""
    counts = df[id_column].value_counts().reset_index()
    counts.columns = [id_column, count_column_name]
    return counts

def merge_dataframes(main_df: pd.DataFrame, to_merge: pd.DataFrame, 
                    left_on: str, right_on: str, how: str = 'left') -> pd.DataFrame:
    """Realiza un merge entre DataFrames con manejo de errores."""
    try:
        return main_df.merge(to_merge, how=how, left_on=left_on, right_on=right_on)
    except Exception as e:
        raise ValueError(f"Error al fusionar DataFrames: {str(e)}")

def process_payment_data(payment_df: pd.DataFrame, dfs: Dict[str, pd.DataFrame], 
                         config: DataProcessingConfig, payment_type: str) -> pd.DataFrame:
    """Procesa los datos de pagos (Efecty o Bancolombia)."""
    if payment_type not in ['efecty', 'bancolombia']:
        raise ValueError("Tipo de pago debe ser 'efecty' o 'bancolombia'")
    # Formato de fechas
    if payment_type == 'bancolombia' and 'Fecha' in payment_df.columns:
          payment_df['Fecha'] = pd.to_datetime(payment_df['Fecha']).dt.strftime('%d/%m/%Y')
    
    if payment_type == 'efecty' and 'Fecha' in payment_df.columns:
          payment_df['Fecha'] = pd.to_datetime(payment_df['Fecha']).dt.strftime('%d/%m/%Y')
    
    merge_conf = config.merge_config[payment_type]
    
    # Fusionar con empleados
    result_df = merge_dataframes(
        payment_df, 
        dfs['EMPLEADOS ACTUALES'], 
        *merge_conf['empleados']
    )
    
    # Fusionar con AC FS
    result_df = merge_dataframes(
        result_df, 
        dfs['AC FS'], 
        *merge_conf['ac_fs']
    )
    
    # Fusionar con AC ARP
    result_df = merge_dataframes(
        result_df, 
        dfs['AC ARP'], 
        *merge_conf['ac_arp']
    )
    
    # Llenar valores faltantes
    result_df['EMPLEADO'] = result_df['ESTADO_EMPLEADO'].fillna('NO')
    result_df['CARTERA EN FINANSUEÑOS'] = result_df['FACTURA_FS'].fillna('SIN CARTERA')
    result_df['CARTERA EN ARPESOD'] = result_df['FACTURA_ARP'].fillna('SIN CARTERA')
    
    # Contar cuentas FS
    conteo_fs = count_accounts(dfs['AC FS'], 'CEDULA_FS', 'CANTIDAD CUENTAS FS')
    result_df = merge_dataframes(
        result_df, 
        conteo_fs, 
        merge_conf['ac_fs'][0], 
        'CEDULA_FS'
    )
    result_df['CANTIDAD CUENTAS FS'] = result_df['CANTIDAD CUENTAS FS'].fillna(0).astype(int)
    
    # Contar cuentas ARP
    conteo_arp = count_accounts(dfs['AC ARP'], 'CEDULA_ARP', 'CANTIDAD CUENTAS ARP')
    result_df = merge_dataframes(
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
    result_df = merge_dataframes(
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
    result_df = merge_dataframes(
        result_df, 
        dfs['CASA DE COBRANZA'], 
        *merge_conf['casa_cobranza']
    )
    result_df['CASA COBRANZA'] = result_df['CASA COBRANZA'].fillna('SIN CASA DE COBRANZA')
    
    # Fusionar con codeudores
    result_df = merge_dataframes(
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

def save_results_to_excel(df_bancolombia: pd.DataFrame, df_efecty: pd.DataFrame, 
                         output_file: str, status_callback) -> bool:
    """Guarda los resultados en un archivo Excel con manejo de errores."""
    try:
        status_callback("Guardando resultados...")
        
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

# ==============================================
# FUNCIÓN PRINCIPAL DE PROCESAMIENTO
# ==============================================

def process_file(file_path: str, status_callback, progress_callback) -> bool:
    """Función principal que orquesta todo el procesamiento."""
    try:
        config = DataProcessingConfig()
        
        status_callback("Validando archivo...")
        progress_callback(10)
        
        # Cargar y filtrar datos
        status_callback("Cargando y filtrando datos...")
        dfs = load_and_filter_data(file_path, config)
        progress_callback(20)
        
        # Convertir columnas a string
        status_callback("Preparando datos...")
        string_columns = {
            'PAGOS BANCOLOMBIA': ['Referencia 1', 'Referencia 2'],
            'PAGOS EFECTY': ['Identificación'],
            'EMPLEADOS ACTUALES': ['vincedula'],
            'AC FS': ['CEDULA_FS', 'FACTURA_FS'],
            'AC ARP': ['CEDULA_ARP', 'FACTURA_ARP'],
            'CASA DE COBRANZA': ['FACTURA'],
            'CODEUDORES': ['DOCUMENTO_CODEUDOR']
        }
        
        for sheet, cols in string_columns.items():
            if sheet in dfs:
                dfs[sheet] = convert_columns_to_string(dfs[sheet], cols)
        
        progress_callback(30)
        
        # Procesar pagos de Bancolombia
        status_callback("Procesando pagos de Bancolombia...")
        df_bancolombia = process_payment_data(dfs['PAGOS BANCOLOMBIA'], dfs, config, 'bancolombia')
        progress_callback(60)
        
        # Procesar pagos de Efecty
        status_callback("Procesando pagos de Efecty...")
        df_efecty = process_payment_data(dfs['PAGOS EFECTY'], dfs, config, 'efecty')
        progress_callback(80)
        
        # Solicitar ubicación para guardar
        status_callback("Solicitando ubicación para guardar...")
        output_file = filedialog.asksaveasfilename(
            title="Guardar archivo procesado",
            initialdir=os.path.expanduser('~'),
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
            initialfile=AppConfig.output_filename
        )
        
        if not output_file:
            status_callback("Operación cancelada por el usuario")
            messagebox.showinfo("Información", "Guardado cancelado")
            return False
        
        # Guardar resultados
        if save_results_to_excel(df_bancolombia, df_efecty, output_file, status_callback):
            status_callback("Proceso completado con éxito")
            messagebox.showinfo("Éxito", f"¡Archivo generado exitosamente en:\n{output_file}!")
            return True
        
    except Exception as e:
        status_callback("Error en el procesamiento")
        messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")
        return False
    finally:
        progress_callback(100)

# ==============================================
# INTERFAZ GRÁFICA
# ==============================================

class FinancialReportApp:
    def __init__(self, root):
        self.root = root
        self.config = AppConfig()
        self.setup_ui()
        
    def setup_ui(self):
        """Configura la interfaz de usuario."""
        self.root.title(self.config.title)
        self.root.geometry(self.config.geometry)
        self.root.resizable(False, False)
        self.root.configure(bg=self.config.bg_color)
        
        # Configurar estilo
        style = ttk.Style()
        style.theme_use('clam')
        
        # Fuentes
        title_font = Font(family="Helvetica", size=16, weight="bold")
        button_font = Font(family="Arial", size=12)
        label_font = Font(family="Arial", size=10)
        
        # Marco principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.configure(style='Card.TFrame')
        
        # Estilo para el marco
        style.configure('Card.TFrame', background=self.config.bg_color, borderwidth=2, relief="groove")
        
        # Título
        title_label = ttk.Label(
            main_frame, 
            text=self.config.title, 
            font=title_font,
            background=self.config.bg_color,
            foreground=self.config.secondary_color
        )
        title_label.pack(pady=(0, 20))
        
        # Descripción
        desc_label = ttk.Label(
            main_frame,
            text="Esta herramienta procesa archivos Excel con información financiera\ny genera un reporte consolidado.",
            font=label_font,
            background=self.config.bg_color,
            foreground=self.config.text_color,
            justify=tk.CENTER
        )
        desc_label.pack(pady=(0, 30))
        
        # Botón para seleccionar archivo
        self.select_button = ttk.Button(
            main_frame,
            text="Seleccionar Archivo Excel",
            command=self.select_file,
            style='Accent.TButton'
        )
        self.select_button.pack(pady=(0, 20), ipadx=10, ipady=5)
        
        # Barra de progreso
        self.progress_bar = ttk.Progressbar(
            main_frame,
            orient=tk.HORIZONTAL,
            length=300,
            mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 20))
        
        # Información de estado
        self.status_label = ttk.Label(
            main_frame,
            text="Esperando selección de archivo...",
            font=label_font,
            background=self.config.bg_color,
            foreground=self.config.text_color
        )
        self.status_label.pack()
        
        # Configurar estilo para el botón de acento
        style.configure('Accent.TButton', font=button_font, foreground='white', background=self.config.accent_color)
        style.map('Accent.TButton',
                background=[('active', self.config.secondary_color), ('pressed', self.config.secondary_color)])
        
        # Pie de página
        footer_label = ttk.Label(
            main_frame,
            text="© 2023 Departamento Financiero",
            font=label_font,
            background=self.config.bg_color,
            foreground=self.config.text_color
        )
        footer_label.pack(side=tk.BOTTOM, pady=(20, 0))
    
    def select_file(self):
        """Maneja la selección de archivo."""
        filetypes = [
            ('Archivos Excel', '*.xlsx'),
            ('Todos los archivos', '*.*')
        ]
        
        file_path = filedialog.askopenfilename(
            title="Selecciona el archivo Excel a procesar",
            initialdir=os.path.expanduser('~'),
            filetypes=filetypes
        )
        
        if file_path:
            self.progress_bar['value'] = 0
            self.root.after(100, lambda: self.process_file_async(file_path))
    
    def process_file_async(self, file_path):
        """Ejecuta el procesamiento del archivo en segundo plano."""
        try:
            process_file(
                file_path,
                self.update_status,
                self.update_progress
            )
        except Exception as e:
            self.update_status("Error en el procesamiento")
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{str(e)}")
    
    def update_status(self, message):
        """Actualiza el mensaje de estado."""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def update_progress(self, value):
        """Actualiza la barra de progreso."""
        self.progress_bar['value'] = value
        self.root.update_idletasks()

# ==============================================
# EJECUCIÓN DE LA APLICACIÓN
# ==============================================

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = FinancialReportApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Error inesperado en la aplicación: {str(e)}")
        messagebox.showerror("Error Crítico", f"Ocurrió un error inesperado:\n{str(e)}")
    finally:
        print("Aplicación finalizada")