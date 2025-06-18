import tkinter as tk
import sys
from pathlib import Path

# Añade el directorio src al path de Python
sys.path.append(str(Path(__file__).resolve().parent.parent))

from src.controllers.financiero_controller import MainController as FinancieroController
from src.controllers.anticipos_online_controller import MainController as AnticiposController
from views.main_window import MainWindow

def main():
    try:
        root = tk.Tk()
        controller_financiero = FinancieroController(None)  # Pasamos None temporalmente
        controller_anticipos = AnticiposController(None)  # Placeholder para el controlador de anticipos si es necesario
        view= MainWindow(root, controller_financiero, controller_anticipos)
        controller_anticipos.view = view
        controller_financiero.view = view  # Establecemos la vista después
        root.mainloop()
    except Exception as e:
        print(f"Error inesperado en la aplicación: {str(e)}")
    finally:
        print("Aplicación finalizada")

if __name__ == "__main__":
    main()