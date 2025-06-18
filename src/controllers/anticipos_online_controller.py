import os
import pandas as pd
from tkinter import filedialog, messagebox
from src.models.anticipos_online_processor import AnticiposOnlineProcessor

class MainController:
    def __init__(self,view):
        self.view = view
        self.data_processor = AnticiposOnlineProcessor()
    
    def select_file(self):
        filetypes = [
            ('Archivos Excel','*xlsx'),
            ('Todos los archivos','*.*')
        ]
        
        file_path = filedialog.askopenfilename(
            title = 'Selecciona el archivo Excel a procesar',
            initialdir= os.path.expanduser('~'),
            filetypes=filetypes
        )    
        if file_path:
            self.view.progress_bar['value'] = 0
            self.view.root.after(100, lambda: self.process_file(file_path))
    
    def process_file(self, file_path: str):
        try:
            self.view.update_status("Validando archivo...")
            self.view.update_progress(10)
            
            self.view.update_status("Cargando y filtrando datos...")
            dfs = self.data_processor.load_and_filter_data(file_path)    
            self.view.update_progress(20)
            
            self.view.update_status("Asegurando compatibilidad de tipos (CEDULA)...")
            
            df_online = dfs['ONLINE']
            df_ac_fs = dfs['AC FS']
            df_ac_arp = dfs['AC ARP']
            
            df_online['CEDULA'] = df_online['CEDULA'].astype(str).str.strip()
            df_ac_fs['CEDULA'] = df_ac_fs['CEDULA'].astype(str).str.strip()
            df_ac_arp['CEDULA'] = df_ac_arp['CEDULA'].astype(str).str.strip()
            
            self.view.update_progress(30)
            
            self.view.update_status("Fusionando datos de 'AC FS'...") 
            merged_df = pd.merge(
            left=df_online,
            right=df_ac_fs,
            on='CEDULA',  # La columna clave para la unión
            how='left'    # Tipo de unión
            )
            self.view.update_status(50)
            
            self.view.update_status("Fusionando datos de 'AC ARP'...")
            merged_df = pd.merge(
                left = merged_df,
                right= df_ac_arp,
                on = 'CEDULA',
                how = 'left'
            )
            final_df = merged_df
            self.view.update_status(70)
            
            self.view.update_status("Guardando el reporte final...")
            output_filename = self.data_processor.config.output_filename
            self.data_processor.save_formatted_excel(final_df, output_filename)
            self.view.update_progress(90)
            messagebox.showinfo("Éxito", f"El archivo '{output_filename}' ha sido generado exitosamente.")
            return True
             
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar el archivo:\n {str(e)}")
            self.view.update_status("Error al procesar el archivo.\n")
            return False
        finally:
            self.view.update_progress(100)
            self.view.update_status("Proceso completado.")    