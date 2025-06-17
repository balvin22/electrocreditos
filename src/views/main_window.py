import tkinter as tk
from tkinter import ttk
from tkinter.font import Font
from models.data_models import AppConfig

class MainWindow:
    def __init__(self, root, controller):
        self.root = root
        self.financiero_controller = controller
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
        
        
        # Frame para contener los botones en una disposición horizontal
        self.buttons_frame = ttk.Frame(self.main_frame)
        self.buttons_frame.pack(pady=(0, 20))
        
        # Primer botón de acción
        self.action1_button = ttk.Button(
            self.buttons_frame,
            text="Reporte Financiero",
            command=self.financiero_controller.select_file,  # Puedes cambiar esto a la acción específica
            style='Accent.TButton'
        )
        self.action1_button.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)
        
        # Segundo botón de acción
        self.action2_button = ttk.Button(
            self.buttons_frame,
            text="Anticipos Online",
            command=self.financiero_controller.select_file,  # Debes implementar este método en el controlador
            style='Accent.TButton'
        )
        self.action2_button.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)
        
        # Barra de progreso (inicialmente oculta)
        self.progress_bar = ttk.Progressbar(
            self.main_frame,
            orient=tk.HORIZONTAL,
            length=300,
            mode='determinate'
        )
        # No la empaquetamos todavía, se mostrará cuando sea necesario
        
        # Información de estado
        self.status_label = ttk.Label(
            self.main_frame,
            text="Seleccione una opción para comenzar...",
            font=self.label_font,
            background=self.config.bg_color,
            foreground=self.config.text_color
        )
        self.status_label.pack(pady=(10, 0))
        
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
        self.progress_bar.pack(pady=(0, 20))
        self.root.update_idletasks()
        