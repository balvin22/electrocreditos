import os
from tkinter import filedialog, messagebox
from typing import Callable
from models.data_processor import DataProcessor  # Cambiado a import absoluto
from models.data_models import DataProcessingConfig

class MainController:
    def __init__(self, view):
        self.view = view
        self.data_processor = DataProcessor()
        
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
            self.view.progress_bar['value'] = 0
            self.view.root.after(100, lambda: self.process_file(file_path))
    
    def process_file(self, file_path: str):
        """Orquesta el procesamiento del archivo."""
        try:
            self.view.update_status("Validando archivo...")
            self.view.update_progress(10)
            
            # Cargar y filtrar datos
            self.view.update_status("Cargando y filtrando datos...")
            dfs = self.data_processor.load_and_filter_data(file_path)
            self.view.update_progress(20)
            
            # Convertir columnas a string
            self.view.update_status("Preparando datos...")
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
                    dfs[sheet] = self.data_processor.convert_columns_to_string(dfs[sheet], cols)
            
            self.view.update_progress(30)
            
            # Procesar pagos de Bancolombia
            self.view.update_status("Procesando pagos de Bancolombia...")
            df_bancolombia = self.data_processor.process_payment_data(
                dfs['PAGOS BANCOLOMBIA'], dfs, 'bancolombia'
            )
            self.view.update_progress(60)
            
            # Procesar pagos de Efecty
            self.view.update_status("Procesando pagos de Efecty...")
            df_efecty = self.data_processor.process_payment_data(
                dfs['PAGOS EFECTY'], dfs, 'efecty'
            )
            self.view.update_progress(80)
            
            # Solicitar ubicación para guardar
            self.view.update_status("Solicitando ubicación para guardar...")
            output_file = filedialog.asksaveasfilename(
                title="Guardar archivo procesado",
                initialdir=os.path.expanduser('~'),
                defaultextension=".xlsx",
                filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
                initialfile=self.data_processor.config.output_filename
            )
            
            if not output_file:
                self.view.update_status("Operación cancelada por el usuario")
                messagebox.showinfo("Información", "Guardado cancelado")
                return False
            
            # Guardar resultados
            if self.data_processor.save_results_to_excel(df_bancolombia, df_efecty, output_file):
                self.view.update_status("Proceso completado con éxito")
                messagebox.showinfo("Éxito", f"¡Archivo generado exitosamente en:\n{output_file}!")
                return True
            
        except Exception as e:
            self.view.update_status("Error en el procesamiento")
            messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")
            return False
        finally:
            self.view.update_progress(100)