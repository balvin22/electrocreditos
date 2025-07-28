import os
import shutil
import uuid
from fastapi import APIRouter, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from src.controllers.cifin_contoller import CifinController

# Crea el router para las rutas de CIFIN
router = APIRouter()
# Crea una instancia del controlador para usarla en las rutas
cifin_controller = CifinController()

def cleanup_temp_folder(folder_path: str):
    """Función de limpieza para ejecutar en segundo plano."""
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)
        print(f"Limpieza: Carpeta temporal {folder_path} eliminada.")

@router.post("/cifin/process", tags=["CIFIN"])
async def procesar_reporte_cifin(
    background_tasks: BackgroundTasks,
    archivo_plano: UploadFile = File(...),
    archivo_correcciones: UploadFile = File(...)
):
    """
    Recibe los archivos de CIFIN, los pasa al controlador para su procesamiento,
    y devuelve el reporte final.
    """
    temp_dir = f"temp_{uuid.uuid4().hex}"
    os.makedirs(temp_dir)

    try:
        # Guarda los archivos subidos temporalmente
        plano_path = os.path.join(temp_dir, archivo_plano.filename)
        corrections_path = os.path.join(temp_dir, archivo_correcciones.filename)
        output_filename = f"Resultado_CIFIN_{os.path.splitext(archivo_plano.filename)[0]}.xlsx"
        output_path = os.path.join(temp_dir, output_filename)

        with open(plano_path, "wb") as f:
            shutil.copyfileobj(archivo_plano.file, f)
        with open(corrections_path, "wb") as f:
            shutil.copyfileobj(archivo_correcciones.file, f)

        # Llama al controlador para que haga todo el trabajo
        cifin_controller.process_files(plano_path, corrections_path, output_path)

        # Registra la tarea de limpieza para después de enviar la respuesta
        background_tasks.add_task(cleanup_temp_folder, temp_dir)

        # Devuelve el archivo procesado
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        cleanup_temp_folder(temp_dir)
        raise HTTPException(status_code=500, detail=str(e))