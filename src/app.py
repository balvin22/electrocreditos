import tkinter as tk
import sys
from pathlib import Path

# Añade el directorio src al path de Python
sys.path.append(str(Path(__file__).resolve().parent.parent))



from controller.main_controller import MainController
from view.main_window import MainWindow

def main():
    try:
        root = tk.Tk()
        controller = MainController(None)  # Pasamos None temporalmente
        view = MainWindow(root, controller)
        controller.view = view  # Establecemos la vista después
        root.mainloop()
    except Exception as e:
        print(f"Error inesperado en la aplicación: {str(e)}")
    finally:
        print("Aplicación finalizada")

if __name__ == "__main__":
    main()