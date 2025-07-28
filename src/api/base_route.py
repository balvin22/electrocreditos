import os
import shutil
import uuid
from typing import List, Optional
from fastapi import APIRouter, UploadFile, File, HTTPException, BackgroundTasks, Form
from fastapi.responses import FileResponse
from src.controllers.base_controller import BaseMensualController

router = APIRouter()
base_mensual_controller = BaseMensualController()

def cleanup_temp_folder(folder_path: str):
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)
        print(f"Limpieza: Carpeta temporal {folder_path} eliminada.")

@router.post("/base-mensual/process", tags=["Base Mensual"])
async def procesar_reporte_base_mensual(
    background_tasks: BackgroundTasks,
    # Parámetros para las fechas (opcionales)
    start_date: Optional[str] = Form(None),
    end_date: Optional[str] = Form(None),
    # Parámetros para los archivos
    ANALISIS: List[UploadFile] = File(...),
    R91: List[UploadFile] = File(...),
    VENCIMIENTOS: List[UploadFile] = File(...),
    R03: List[UploadFile] = File(...),
    SC04: UploadFile = File(...),
    CRTMPCONSULTA1: UploadFile = File(...),
    FNZ003: UploadFile = File(...),
    MATRIZ_CARTERA: UploadFile = File(...),
    METAS_FRANJAS: UploadFile = File(...),
    ASESORES: UploadFile = File(...),
    DESEMBOLSOS_FINANSUEÑOS: UploadFile = File(...)
):
    temp_dir = f"temp_{uuid.uuid4().hex}"
    os.makedirs(temp_dir)
    
    try:
        # Diccionario para guardar las rutas de los archivos guardados
        rutas_archivos = {}

        # Función auxiliar para guardar archivos y organizar rutas
        def guardar_archivo(files, key):
            rutas_guardadas = []
            file_list = files if isinstance(files, list) else [files]
            for file in file_list:
                path = os.path.join(temp_dir, file.filename)
                with open(path, "wb") as f:
                    shutil.copyfileobj(file.file, f)
                rutas_guardadas.append(path)
            rutas_archivos[key] = rutas_guardadas
        
        # Guardar todos los archivos subidos
        guardar_archivo(ANALISIS, "ANALISIS")
        guardar_archivo(R91, "R91")
        guardar_archivo(VENCIMIENTOS, "VENCIMIENTOS")
        guardar_archivo(R03, "R03")
        guardar_archivo(SC04, "SC04")
        guardar_archivo(CRTMPCONSULTA1, "CRTMPCONSULTA1")
        guardar_archivo(FNZ003, "FNZ003")
        guardar_archivo(MATRIZ_CARTERA, "MATRIZ_CARTERA")
        guardar_archivo(METAS_FRANJAS, "METAS_FRANJAS")
        guardar_archivo(ASESORES, "ASESORES")
        guardar_archivo(DESEMBOLSOS_FINANSUEÑOS, "DESEMBOLSOS_FINANSUEÑOS")

        output_filename = "Reporte_Consolidado_Final.xlsx"
        output_path = os.path.join(temp_dir, output_filename)

        # Llama al controlador con el diccionario de rutas y las fechas
        base_mensual_controller.process_files(
            rutas_archivos=rutas_archivos,
            output_path=output_path,
            start_date=start_date,
            end_date=end_date
        )

        background_tasks.add_task(cleanup_temp_folder, temp_dir)

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        cleanup_temp_folder(temp_dir)
        raise HTTPException(status_code=500, detail=str(e))