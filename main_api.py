from fastapi import FastAPI
from src.api import datacredito_route
from src.api import cifin_route
from src.api import base_route


app = FastAPI(
    title="API de Procesamiento de Reportes Financieros",
    description="Procesa archivos de Datacrédito"
)

# Incluye las rutas de Datacrédito en la aplicación principal
app.include_router(datacredito_route.router, prefix="/api/v1")
app.include_router(cifin_route.router, prefix="/api/v1")
app.include_router(base_route.router, prefix="/api/v1")


@app.get("/", tags=["Root"])
def read_root():
    return {"message": "Bienvenido a la API de Procesamiento de Reportes"}