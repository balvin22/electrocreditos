import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
import pandas as pd
import numpy as np
import os

# Función para procesar el archivo (se mantiene igual)
def procesar_archivo(file_path):
    try:
        update_status("Procesando datos...")
        df = pd.read_excel(file_path, sheet_name=None)

        df_ac_fs = df['AC FS']
        df_ac_arp = df['AC ARP']
        df_codeudores = df['CODEUDORES']
        df_casa_cobranza = df['CASA DE COBRANZA']
        df_empleados = df['EMPLEADOS ACTUALES']
        df_pagos_bancolombia = df['PAGOS BANCOLOMBIA']
        df_pagos_efecty = df['PAGOS EFECTY']

        df_pagos_bancolombia_filtrado = df_pagos_bancolombia[['No.','Fecha','Detalle 1','Detalle 2','Referencia 1','Referencia 2','Valor']].copy()
        df_pagos_efecty_filtrado = df_pagos_efecty[['No','Identificación','Valor','N° de Autorización','Fecha']].copy()

        df_ac_fs_filtrado = df_ac_fs[['CEDULA','FACTURA','saldofac','ccosto']].copy().rename(columns={'CEDULA': 'CEDULA_FS', 'FACTURA': 'FACTURA_FS', 'saldofac': 'SALDO_FS','ccosto':'CENTRO COSTO FS'})
        df_ac_arp_filtrado = df_ac_arp[['CEDULA','FACTURA', 'saldofac','ccosto']].copy().rename(columns={'CEDULA': 'CEDULA_ARP', 'FACTURA': 'FACTURA_ARP','saldofac': 'SALDO_ARP','ccosto':'CENTRO COSTO ARP'})
        df_empleados_filtrado = df_empleados[['vincedula','ACTIVO']].copy().rename(columns={'ACTIVO': 'ESTADO_EMPLEADO'})
        df_casa_cobranza_filtrado = df_casa_cobranza[['FACTURA','cobra']].copy().rename(columns={'cobra': 'CASA COBRANZA'})
        df_codeudores_filtrado = df_codeudores[['CODEUDOR','FACTURA']].copy().rename(columns={'FACTURA': 'CODEUDOR','CODEUDOR': 'DOCUMENTO_CODEUDOR'})
    
        

        df_pagos_bancolombia_filtrado['Referencia 1'] = df_pagos_bancolombia_filtrado['Referencia 1'].astype(str)
        df_pagos_bancolombia_filtrado['Referencia 2'] = df_pagos_bancolombia_filtrado['Referencia 2'].astype(str)
        df_pagos_efecty_filtrado['Identificación'] = df_pagos_efecty_filtrado['Identificación'].astype(str)
        df_empleados_filtrado['vincedula'] = df_empleados_filtrado['vincedula'].astype(str)
        df_ac_fs_filtrado['CEDULA_FS'] = df_ac_fs_filtrado['CEDULA_FS'].astype(str)
        df_ac_fs_filtrado['FACTURA_FS'] = df_ac_fs_filtrado['FACTURA_FS'].astype(str)
        df_ac_arp_filtrado['CEDULA_ARP'] = df_ac_arp_filtrado['CEDULA_ARP'].astype(str)
        df_ac_arp_filtrado['FACTURA_ARP'] = df_ac_arp_filtrado['FACTURA_ARP'].astype(str)
        df_casa_cobranza_filtrado['FACTURA'] = df_casa_cobranza_filtrado['FACTURA'].astype(str)
        df_codeudores_filtrado['DOCUMENTO_CODEUDOR'] = df_codeudores_filtrado['DOCUMENTO_CODEUDOR'].astype(str)
        
        df_resultado_efecty = df_pagos_efecty_filtrado.merge(df_empleados_filtrado, how='left', left_on='Identificación', right_on='vincedula')
        df_resultado_efecty = df_resultado_efecty.merge(df_ac_fs_filtrado, how='left', left_on='Identificación', right_on='CEDULA_FS')
        df_resultado_efecty = df_resultado_efecty.merge(df_ac_arp_filtrado, how='left', left_on='Identificación', right_on='CEDULA_ARP')
        

        df_resultado_bancolombia = df_pagos_bancolombia_filtrado.merge(df_empleados_filtrado, how='left', left_on='Referencia 1', right_on='vincedula')
        df_resultado_bancolombia = df_resultado_bancolombia.merge(df_ac_fs_filtrado, how='left', left_on='Referencia 1', right_on='CEDULA_FS')
        df_resultado_bancolombia = df_resultado_bancolombia.merge(df_ac_arp_filtrado, how='left', left_on='Referencia 1', right_on='CEDULA_ARP')
        
        df_resultado_efecty['EMPLEADO'] = df_resultado_efecty['ESTADO_EMPLEADO'].fillna('NO')
        df_resultado_efecty['CARTERA EN FINANSUEÑOS'] = df_resultado_efecty['FACTURA_FS'].fillna('SIN CARTERA')
        df_resultado_efecty['CARTERA EN ARPESOD'] = df_resultado_efecty['FACTURA_ARP'].fillna('SIN CARTERA')

        df_resultado_bancolombia['EMPLEADO'] = df_resultado_bancolombia['ESTADO_EMPLEADO'].fillna('NO')
        df_resultado_bancolombia['CARTERA EN FINANSUEÑOS'] = df_resultado_bancolombia['FACTURA_FS'].fillna('SIN CARTERA')
        df_resultado_bancolombia['CARTERA EN ARPESOD'] = df_resultado_bancolombia['FACTURA_ARP'].fillna('SIN CARTERA')

        conteo_cuentas_fs = df_ac_fs_filtrado['CEDULA_FS'].value_counts().reset_index()
        conteo_cuentas_fs.columns = ['CEDULA_FS', 'CANTIDAD CUENTAS FS']
        
        df_resultado_efecty = df_resultado_efecty.merge(conteo_cuentas_fs, how='left', left_on='Identificación', right_on='CEDULA_FS')
        df_resultado_efecty['CANTIDAD CUENTAS FS'] = df_resultado_efecty['CANTIDAD CUENTAS FS'].fillna(0).astype(int)
        
        df_resultado_bancolombia = df_resultado_bancolombia.merge(conteo_cuentas_fs, how='left', left_on='Referencia 1', right_on='CEDULA_FS')
        df_resultado_bancolombia['CANTIDAD CUENTAS FS'] = df_resultado_bancolombia['CANTIDAD CUENTAS FS'].fillna(0).astype(int)

        conteo_cuentas_arp = df_ac_arp_filtrado['CEDULA_ARP'].value_counts().reset_index()
        conteo_cuentas_arp.columns = ['CEDULA_ARP', 'CANTIDAD CUENTAS ARP']
        
        df_resultado_efecty = df_resultado_efecty.merge(conteo_cuentas_arp, how='left', left_on='Identificación', right_on='CEDULA_ARP')
        df_resultado_efecty['CANTIDAD CUENTAS ARP'] = df_resultado_efecty['CANTIDAD CUENTAS ARP'].fillna(0).astype(int)
        
        df_resultado_bancolombia = df_resultado_bancolombia.merge(conteo_cuentas_arp, how='left', left_on='Referencia 1', right_on='CEDULA_ARP')
        df_resultado_bancolombia['CANTIDAD CUENTAS ARP'] = df_resultado_bancolombia['CANTIDAD CUENTAS ARP'].fillna(0).astype(int)

        df_resultado_efecty['FACTURA FINAL'] = np.where(
            df_resultado_efecty['CARTERA EN FINANSUEÑOS'] != 'SIN CARTERA',
            df_resultado_efecty['CARTERA EN FINANSUEÑOS'],
            df_resultado_efecty['CARTERA EN ARPESOD']
        )
           
        df_resultado_bancolombia['FACTURA FINAL'] = np.where(
            df_resultado_bancolombia['CARTERA EN FINANSUEÑOS'] != 'SIN CARTERA',
            df_resultado_bancolombia['CARTERA EN FINANSUEÑOS'],
            df_resultado_bancolombia['CARTERA EN ARPESOD']
        )
        
        df_fs_saldos = df_ac_fs_filtrado[['FACTURA_FS', 'SALDO_FS']].rename(columns={'FACTURA_FS': 'FACTURA', 'SALDO_FS': 'SALDO'})
        df_arp_saldos = df_ac_arp_filtrado[['FACTURA_ARP', 'SALDO_ARP']].rename(columns={'FACTURA_ARP': 'FACTURA', 'SALDO_ARP': 'SALDO'})
        df_saldos_unificados = pd.concat([df_fs_saldos, df_arp_saldos], ignore_index=True).drop_duplicates(subset='FACTURA')
        
        df_resultado_efecty['FACTURA FINAL'] = df_resultado_efecty['FACTURA FINAL'].astype(str)
        df_resultado_efecty = df_resultado_efecty.merge(df_saldos_unificados, how='left', left_on='FACTURA FINAL', right_on='FACTURA')
        df_resultado_efecty = df_resultado_efecty.rename(columns={'SALDO': 'SALDOS'})
        df_resultado_efecty['SALDOS'] = df_resultado_efecty['SALDOS'].fillna(0).astype(float)
        df_resultado_efecty['Valor'] = df_resultado_efecty['Valor'].fillna(0).astype(float)
        
        df_resultado_bancolombia['FACTURA FINAL'] = df_resultado_bancolombia['FACTURA FINAL'].astype(str)
        df_resultado_bancolombia = df_resultado_bancolombia.merge(df_saldos_unificados, how='left', left_on='FACTURA FINAL', right_on='FACTURA')
        df_resultado_bancolombia = df_resultado_bancolombia.rename(columns={'SALDO': 'SALDOS'})
        df_resultado_bancolombia['SALDOS'] = df_resultado_bancolombia['SALDOS'].fillna(0).astype(float)
        df_resultado_bancolombia['Valor'] = df_resultado_bancolombia['Valor'].fillna(0).astype(float)

        df_resultado_bancolombia['VALIDACION ULTIMO SALDO'] = np.where(
            (df_resultado_bancolombia['SALDOS'] - df_resultado_bancolombia['Valor']) <= 0,
            'pago total',
            (df_resultado_bancolombia['SALDOS'] - df_resultado_bancolombia['Valor']).astype(str)
        )
        
        df_resultado_efecty['VALIDACION ULTIMO SALDO'] = np.where(
            (df_resultado_efecty['SALDOS'] - df_resultado_efecty['Valor']) <= 0,
            'pago total',
            (df_resultado_efecty['SALDOS'] - df_resultado_efecty['Valor']).astype(str)
        )
        
        df_resultado_efecty = df_resultado_efecty.merge(df_casa_cobranza_filtrado, how='left', left_on='FACTURA FINAL', right_on='FACTURA')
        df_resultado_efecty['CASA COBRANZA'] = df_resultado_efecty['CASA COBRANZA'].fillna('SIN CASA DE COBRANZA')

        df_resultado_bancolombia = df_resultado_bancolombia.merge(df_casa_cobranza_filtrado, how='left', left_on='CARTERA EN ARPESOD', right_on='FACTURA')
        df_resultado_bancolombia['CASA COBRANZA'] = df_resultado_bancolombia['CASA COBRANZA'].fillna('SIN CASA DE COBRANZA')
        
        df_resultado_efecty = df_resultado_efecty.merge(df_codeudores_filtrado, how='left', left_on='Identificación', right_on='DOCUMENTO_CODEUDOR')
        df_resultado_efecty['CODEUDOR'] = df_resultado_efecty['CODEUDOR'].fillna('SIN CODEUDOR')

        df_resultado_bancolombia = df_resultado_bancolombia.merge(df_codeudores_filtrado, how='left', left_on = 'Referencia 1', right_on = 'DOCUMENTO_CODEUDOR')
        df_resultado_bancolombia['CODEUDOR'] = df_resultado_bancolombia['CODEUDOR'].fillna('SIN CODEUDOR')

        columnas_a_eliminar = ['vincedula','FACTURA','DOCUMENTO_CODEUDOR','FACTURA_x','FACTURA_y',
                               'ESTADO_EMPLEADO', 'CEDULA_FS','CEDULA_FS_x','CEDULA_ARP_x',
                               'CEDULA_FS_y','FACTURA_FS', 'CEDULA_ARP','CEDULA_ARP_y',
                               'FACTURA_ARP','SALDO_FS', 'SALDO_ARP']
        
        df_resultado_bancolombia = df_resultado_bancolombia.drop(columns=[col for col in columnas_a_eliminar if col in df_resultado_bancolombia.columns])
        df_resultado_efecty = df_resultado_efecty.drop(columns=[col for col in columnas_a_eliminar if col in df_resultado_efecty.columns])
        
        df_resultado_bancolombia 
        print(df_resultado_efecty)
        
        update_status("Solicitando ubicación para guardar...")
        
        output_file = filedialog.asksaveasfilename(
            title="Guardar archivo procesado",
            initialdir=os.path.expanduser('~'),
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
            initialfile="reporte_financiero.xlsx"
        )
        
        if not output_file:
            update_status("Operación cancelada por el usuario")
            messagebox.showinfo("Información", "Guardado cancelado")
            return
            
        update_status(f"Guardando archivo en {output_file}...")
        
        # VERIFICACIÓN ANTES DE GUARDAR
        print("\n=== VERIFICACIÓN DE DATAFRAMES ===")
        print(f"DataFrame Bancolombia - Filas: {len(df_resultado_bancolombia)}, Columnas: {len(df_resultado_bancolombia.columns)}")
        print(f"DataFrame Efecty - Filas: {len(df_resultado_efecty)}, Columnas: {len(df_resultado_efecty.columns)}")
        print("\nPrimeras filas de Efecty:")
        print(df_resultado_efecty.head(2))
        
        # GUARDADO MEJORADO
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Guardar Bancolombia primero
                df_resultado_bancolombia.to_excel(
                    writer, 
                    sheet_name='Bancolombia', 
                    index=False,
                    startrow=0,
                    startcol=0
                )
                
                # Guardar Efecty después
                df_resultado_efecty.to_excel(
                    writer, 
                    sheet_name='Efecty', 
                    index=False,
                    startrow=0,
                    startcol=0
                )
                
            # Verificación post-guardado
            if os.path.exists(output_file):
                # Leer el archivo guardado para verificar
                with pd.ExcelFile(output_file) as xls:
                    sheets = xls.sheet_names
                    print(f"\nHojas en el archivo guardado: {sheets}")
                    
                    if 'Efecty' not in sheets:
                        raise Exception("La hoja Efecty no se creó correctamente")
                        
                update_status("Proceso completado con éxito")
                messagebox.showinfo("Éxito", f"¡Archivo generado exitosamente en:\n{output_file}!")
            else:
                raise Exception("No se pudo crear el archivo de salida")
                
        except Exception as save_error:
            raise Exception(f"Error al guardar: {str(save_error)}")
        
    except Exception as e:
        update_status("Error en el procesamiento")
        # Mostrar detalles técnicos en consola
        import traceback
        traceback.print_exc()
        messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")
    finally:
        progress_bar.stop()
        progress_bar['value'] = 0

def seleccionar_archivo():
    filetypes = [
        ('Archivos Excel', '*.xlsx'),
        ('Todos los archivos', '*.*')
    ]
    archivo = filedialog.askopenfilename(
        title="Selecciona el archivo Excel a procesar",
        initialdir=os.path.expanduser('~'),
        filetypes=filetypes
    )
    if archivo:
        progress_bar.start(10)
        root.after(100, lambda: procesar_archivo(archivo))

def update_status(message):
    status_label.config(text=message)
    root.update_idletasks()

# Configuración de la ventana principal
root = tk.Tk()
root.title("Procesador de Reportes Financieros")
root.geometry("500x350")
root.resizable(False, False)

# Establecer icono (opcional)
try:
    root.iconbitmap(default='icon.ico')  # Puedes crear o descargar un icono
except:
    pass

# Configurar estilo
style = ttk.Style()
style.theme_use('clam')  # Puedes probar con 'alt', 'default', 'vista', 'xpnative'

# Configurar colores
bg_color = "#f0f0f0"
accent_color = "#4b6cb7"
secondary_color = "#2d3747"
text_color = "#333333"

root.configure(bg=bg_color)

# Fuentes
title_font = Font(family="Helvetica", size=16, weight="bold")
button_font = Font(family="Arial", size=12)
label_font = Font(family="Arial", size=10)

# Marco principal
main_frame = ttk.Frame(root, padding="20")
main_frame.pack(fill=tk.BOTH, expand=True)
main_frame.configure(style='Card.TFrame')

# Estilo para el marco
style.configure('Card.TFrame', background=bg_color, borderwidth=2, relief="groove")

# Título
title_label = ttk.Label(
    main_frame, 
    text="Procesador de Reportes Financieros", 
    font=title_font,
    background=bg_color,
    foreground=secondary_color
)
title_label.pack(pady=(0, 20))


# Descripción
desc_label = ttk.Label(
    main_frame,
    text="Esta herramienta procesa archivos Excel con información financiera\ny genera un reporte consolidado.",
    font=label_font,
    background=bg_color,
    foreground=text_color,
    justify=tk.CENTER
)
desc_label.pack(pady=(0, 30))

# Botón para seleccionar archivo
select_button = ttk.Button(
    main_frame,
    text="Seleccionar Archivo Excel",
    command=seleccionar_archivo,
    style='Accent.TButton'
)
select_button.pack(pady=(0, 20), ipadx=10, ipady=5)

# Barra de progreso
progress_bar = ttk.Progressbar(
    main_frame,
    orient=tk.HORIZONTAL,
    length=300,
    mode='indeterminate'
)
progress_bar.pack(pady=(0, 20))

# Información de estado
status_label = ttk.Label(
    main_frame,
    text="Esperando selección de archivo...",
    font=label_font,
    background=bg_color,
    foreground=text_color
)
status_label.pack()

# Configurar estilo para el botón de acento
style.configure('Accent.TButton', font=button_font, foreground='white', background=accent_color)
style.map('Accent.TButton',
          background=[('active', secondary_color), ('pressed', secondary_color)])

# Pie de página
footer_label = ttk.Label(
    main_frame,
    text="© 2023 Departamento Financiero",
    font=label_font,
    background=bg_color,
    foreground=text_color
)
footer_label.pack(side=tk.BOTTOM, pady=(20, 0))

# Actualizar etiqueta de estado durante el procesamiento
def update_status(message):
    status_label.config(text=message)
    root.update_idletasks()

 # Ejecutar la aplicación
try:
    root.mainloop()
except Exception as e:
    print(f"Error inesperado en la aplicación: {str(e)}")
    messagebox.showerror("Error Crítico", f"Ocurrió un error inesperado:\n{str(e)}")
finally:
    # Cualquier limpieza necesaria
    print("Aplicación finalizada")
