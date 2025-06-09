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

        df_pagos_bancolombia_filtrado = df_pagos_bancolombia[['No.','Fecha','Detalle 1','Detalle 2','Referencia 1','Referencia 2','Valor']].copy()

        df_ac_fs_filtrado = df_ac_fs[['CEDULA','FACTURA','saldofac']].copy().rename(columns={'CEDULA': 'CEDULA_FS', 'FACTURA': 'FACTURA_FS', 'saldofac': 'SALDO_FS'})
        df_ac_arp_filtrado = df_ac_arp[['CEDULA','FACTURA', 'saldofac']].copy().rename(columns={'CEDULA': 'CEDULA_ARP', 'FACTURA': 'FACTURA_ARP','saldofac': 'SALDO_ARP'})
        df_empleados_filtrado = df_empleados[['vincedula','ACTIVO']].copy().rename(columns={'ACTIVO': 'ESTADO_EMPLEADO'})
        df_casa_cobranza_filtrado = df_casa_cobranza[['FACTURA','cobra']].copy().rename(columns={'cobra': 'CASA COBRANZA'})
        df_codeudores_filtrado = df_codeudores[['CODEUDOR','FACTURA']].copy().rename(columns={'FACTURA': 'CODEUDOR','CODEUDOR': 'DOCUMENTO_CODEUDOR'})    

        df_pagos_bancolombia_filtrado['Referencia 1'] = df_pagos_bancolombia_filtrado['Referencia 1'].astype(str)
        df_pagos_bancolombia_filtrado['Referencia 2'] = df_pagos_bancolombia_filtrado['Referencia 2'].astype(str)
        df_empleados_filtrado['vincedula'] = df_empleados_filtrado['vincedula'].astype(str)
        df_ac_fs_filtrado['CEDULA_FS'] = df_ac_fs_filtrado['CEDULA_FS'].astype(str)
        df_ac_fs_filtrado['FACTURA_FS'] = df_ac_fs_filtrado['FACTURA_FS'].astype(str)
        df_ac_arp_filtrado['CEDULA_ARP'] = df_ac_arp_filtrado['CEDULA_ARP'].astype(str)
        df_ac_arp_filtrado['FACTURA_ARP'] = df_ac_arp_filtrado['FACTURA_ARP'].astype(str)
        df_casa_cobranza_filtrado['FACTURA'] = df_casa_cobranza_filtrado['FACTURA'].astype(str)
        df_codeudores_filtrado['DOCUMENTO_CODEUDOR'] = df_codeudores_filtrado['DOCUMENTO_CODEUDOR'].astype(str)

        df_resultado = df_pagos_bancolombia_filtrado.merge(df_empleados_filtrado, how='left', left_on='Referencia 1', right_on='vincedula')
        df_resultado = df_resultado.merge(df_ac_fs_filtrado, how='left', left_on='Referencia 1', right_on='CEDULA_FS')
        df_resultado = df_resultado.merge(df_ac_arp_filtrado, how='left', left_on='Referencia 1', right_on='CEDULA_ARP')

        df_resultado['EMPLEADO'] = df_resultado['ESTADO_EMPLEADO'].fillna('NO')
        df_resultado['CARTERA EN FINANSUEÑOS'] = df_resultado['FACTURA_FS'].fillna('SIN CARTERA')
        df_resultado['CARTERA EN ARPESOD'] = df_resultado['FACTURA_ARP'].fillna('SIN CARTERA')

        conteo_cuentas_fs = df_ac_fs_filtrado['CEDULA_FS'].value_counts().reset_index()
        conteo_cuentas_fs.columns = ['CEDULA_FS', 'CANTIDAD CUENTAS FS']
        df_resultado = df_resultado.merge(conteo_cuentas_fs, how='left', left_on='Referencia 1', right_on='CEDULA_FS')
        df_resultado['CANTIDAD CUENTAS FS'] = df_resultado['CANTIDAD CUENTAS FS'].fillna(0).astype(int)

        conteo_cuentas_arp = df_ac_arp_filtrado['CEDULA_ARP'].value_counts().reset_index()
        conteo_cuentas_arp.columns = ['CEDULA_ARP', 'CANTIDAD CUENTAS ARP']
        df_resultado = df_resultado.merge(conteo_cuentas_arp, how='left', left_on='Referencia 1', right_on='CEDULA_ARP')
        df_resultado['CANTIDAD CUENTAS ARP'] = df_resultado['CANTIDAD CUENTAS ARP'].fillna(0).astype(int)

        df_resultado['FACTURA FINAL'] = np.where(
            df_resultado['CARTERA EN FINANSUEÑOS'] != 'SIN CARTERA',
            df_resultado['CARTERA EN FINANSUEÑOS'],
            df_resultado['CARTERA EN ARPESOD']
        )
        df_fs_saldos = df_ac_fs_filtrado[['FACTURA_FS', 'SALDO_FS']].rename(columns={'FACTURA_FS': 'FACTURA', 'SALDO_FS': 'SALDO'})
        df_arp_saldos = df_ac_arp_filtrado[['FACTURA_ARP', 'SALDO_ARP']].rename(columns={'FACTURA_ARP': 'FACTURA', 'SALDO_ARP': 'SALDO'})
        df_saldos_unificados = pd.concat([df_fs_saldos, df_arp_saldos], ignore_index=True).drop_duplicates(subset='FACTURA')
        df_resultado['FACTURA FINAL'] = df_resultado['FACTURA FINAL'].astype(str)
        df_resultado = df_resultado.merge(df_saldos_unificados, how='left', left_on='FACTURA FINAL', right_on='FACTURA')
        df_resultado = df_resultado.rename(columns={'SALDO': 'SALDOS'})
        df_resultado['SALDOS'] = df_resultado['SALDOS'].fillna(0).astype(float)
        df_resultado['Valor'] = df_resultado['Valor'].fillna(0).astype(float)

        df_resultado['VALIDACION ULTIMO SALDO'] = np.where(
            (df_resultado['SALDOS'] - df_resultado['Valor']) <= 0,
            'pago total',
            (df_resultado['SALDOS'] - df_resultado['Valor']).astype(str)
        )

        df_resultado = df_resultado.merge(df_casa_cobranza_filtrado, how='left', left_on='CARTERA EN ARPESOD', right_on='FACTURA')
        df_resultado['CASA COBRANZA'] = df_resultado['CASA COBRANZA'].fillna('SIN CASA DE COBRANZA')

        df_resultado = df_resultado.merge(df_codeudores_filtrado, how='left', left_on = 'Referencia 1', right_on = 'DOCUMENTO_CODEUDOR')
        df_resultado['CODEUDOR'] = df_resultado['CODEUDOR'].fillna('SIN CODEUDOR')

        columnas_a_eliminar = ['vincedula','FACTURA','DOCUMENTO_CODEUDOR','FACTURA_x','FACTURA_y',
                               'ESTADO_EMPLEADO', 'CEDULA_FS','CEDULA_FS_x','CEDULA_ARP_x',
                               'CEDULA_FS_y','FACTURA_FS', 'CEDULA_ARP','CEDULA_ARP_y',
                               'FACTURA_ARP','SALDO_FS', 'SALDO_ARP']
        df_resultado = df_resultado.drop(columns=[col for col in columnas_a_eliminar if col in df_resultado.columns])
        
        update_status("Solicitando ubicación para guardar...")
        
        # Solicitar al usuario dónde guardar
        output_file = filedialog.asksaveasfilename(
            title="Guardar archivo procesado",
            initialdir=os.path.expanduser('~'),
            defaultextension=".xlsx",
            filetypes=[
                ("Archivos Excel", "*.xlsx"),
                ("Todos los archivos", "*.*")
            ],
            initialfile="reporte_financiero.xlsx"
        )
        
        if not output_file:  # Usuario canceló
            update_status("Operación cancelada por el usuario")
            messagebox.showinfo("Información", "Guardado cancelado")
            return
            
        update_status(f"Guardando archivo en {output_file}...")
        df_resultado.to_excel(output_file, index=False)
        
        update_status("Proceso completado con éxito")
        messagebox.showinfo("Éxito", f"¡Archivo generado exitosamente en:\n{output_file}!")
        
    except Exception as e:
        update_status("Error en el procesamiento")
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

# Icono o imagen (opcional)
# Puedes agregar una imagen aquí si lo deseas

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
root.mainloop()
