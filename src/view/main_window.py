import tkinter as tk
from tkinter import ttk
from tkinter.font import Font
from model.data_models import AppConfig

class MainWindow:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.config = AppConfig()
        self.setup_ui()
        
    def setup_ui(self):
        """Configura la interfaz de usuario."""
        self.root.title(self.config.title)
        self.root.geometry(self.config.geometry)
        self.root.resizable(False, False)
        self.root.configure(bg=self.config.bg_color)
        
        # Configurar estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Fuentes
        self.title_font = Font(family="Helvetica", size=16, weight="bold")
        self.button_font = Font(family="Arial", size=12)
        self.label_font = Font(family="Arial", size=10)
        
        # Marco principal
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.main_frame.configure(style='Card.TFrame')
        
        # Estilo para el marco
        self.style.configure('Card.TFrame', background=self.config.bg_color, borderwidth=2, relief="groove")
        
        # Título
        self.title_label = ttk.Label(
            self.main_frame, 
            text=self.config.title, 
            font=self.title_font,
            background=self.config.bg_color,
            foreground=self.config.secondary_color
        )
        self.title_label.pack(pady=(0, 20))
        
        # Descripción
        self.desc_label = ttk.Label(
            self.main_frame,
            text="Esta herramienta procesa archivos Excel con información financiera\ny genera un reporte consolidado.",
            font=self.label_font,
            background=self.config.bg_color,
            foreground=self.config.text_color,
            justify=tk.CENTER
        )
        self.desc_label.pack(pady=(0, 30))
        
        # Botón para seleccionar archivo
        self.select_button = ttk.Button(
            self.main_frame,
            text="Seleccionar Archivo Excel",
            command=self.controller.select_file,
            style='Accent.TButton'
        )
        self.select_button.pack(pady=(0, 20), ipadx=10, ipady=5)
        
        # Barra de progreso
        self.progress_bar = ttk.Progressbar(
            self.main_frame,
            orient=tk.HORIZONTAL,
            length=300,
            mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 20))
        
        # Información de estado
        self.status_label = ttk.Label(
            self.main_frame,
            text="Esperando selección de archivo...",
            font=self.label_font,
            background=self.config.bg_color,
            foreground=self.config.text_color
        )
        self.status_label.pack()
        
        # Configurar estilo para el botón de acento
        self.style.configure('Accent.TButton', font=self.button_font, foreground='white', background=self.config.accent_color)
        self.style.map('Accent.TButton',
                     background=[('active', self.config.secondary_color), ('pressed', self.config.secondary_color)])
        
        # Pie de página
        self.footer_label = ttk.Label(
            self.main_frame,
            text="© 2023 Departamento Financiero",
            font=self.label_font,
            background=self.config.bg_color,
            foreground=self.config.text_color
        )
        self.footer_label.pack(side=tk.BOTTOM, pady=(20, 0))
    
    def update_status(self, message):
        """Actualiza el mensaje de estado."""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def update_progress(self, value):
        """Actualiza la barra de progreso."""
        self.progress_bar['value'] = value
        self.root.update_idletasks()