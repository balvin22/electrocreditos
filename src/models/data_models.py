from dataclasses import dataclass
from typing import Dict, List, Optional

@dataclass
class AppConfig:
    title: str = "Procesador de Reportes Financieros"
    geometry: str = "700x400"
    bg_color: str = "#f0f0f0"
    accent_color: str = "#4b6cb7"
    secondary_color: str = "#2d3747"
    text_color: str = "#333333"
    output_filename: str = "reporte_financiero.xlsx"

@dataclass
class DataProcessingConfig:
    required_sheets: List[str] = None
    sheet_columns: Dict[str, Dict[str, List[str]]] = None
    rename_columns: Dict[str, Dict[str, str]] = None
    merge_config: Dict[str, Dict] = None
    output_filename: str = "reporte_financiero.xlsx"

    def __post_init__(self):
        if self.required_sheets is None:
            self.required_sheets = [
                'AC FS', 'AC ARP', 'CODEUDORES', 
                'CASA DE COBRANZA', 'EMPLEADOS ACTUALES', 
                'PAGOS BANCOLOMBIA', 'PAGOS EFECTY'
            ]
        
        if self.sheet_columns is None:
            self.sheet_columns = {
                'AC FS': ['CEDULA', 'FACTURA', 'saldofac', 'ccosto'],
                'AC ARP': ['CEDULA', 'FACTURA', 'saldofac', 'ccosto'],
                'CODEUDORES': ['CODEUDOR', 'FACTURA'],
                'CASA DE COBRANZA': ['FACTURA', 'cobra'],
                'EMPLEADOS ACTUALES': ['vincedula', 'ACTIVO'],
                'PAGOS BANCOLOMBIA': ['No.', 'Fecha', 'Detalle 1', 'Detalle 2', 'Referencia 1', 'Referencia 2', 'Valor'],
                'PAGOS EFECTY': ['No', 'Identificación', 'Valor', 'N° de Autorización', 'Fecha']
            }
        
        if self.rename_columns is None:
            self.rename_columns = {
                'AC FS': {'CEDULA': 'CEDULA_FS', 'FACTURA': 'FACTURA_FS', 'saldofac': 'SALDO_FS', 'ccosto': 'CENTRO COSTO FS'},
                'AC ARP': {'CEDULA': 'CEDULA_ARP', 'FACTURA': 'FACTURA_ARP', 'saldofac': 'SALDO_ARP', 'ccosto': 'CENTRO COSTO ARP'},
                'CODEUDORES': {'FACTURA': 'CODEUDOR', 'CODEUDOR': 'DOCUMENTO_CODEUDOR'},
                'CASA DE COBRANZA': {'cobra': 'CASA COBRANZA'},
                'EMPLEADOS ACTUALES': {'ACTIVO': 'ESTADO_EMPLEADO'}
            }
        
        if self.merge_config is None:
            self.merge_config = {
                'efecty': {
                    'empleados': ('Identificación', 'vincedula'),
                    'ac_fs': ('Identificación', 'CEDULA_FS'),
                    'ac_arp': ('Identificación', 'CEDULA_ARP'),
                    'casa_cobranza': ('FACTURA FINAL', 'FACTURA'),
                    'codeudores': ('Identificación', 'DOCUMENTO_CODEUDOR')
                },
                'bancolombia': {
                    'empleados': ('Referencia 1', 'vincedula'),
                    'ac_fs': ('Referencia 1', 'CEDULA_FS'),
                    'ac_arp': ('Referencia 1', 'CEDULA_ARP'),
                    'casa_cobranza': ('CARTERA EN ARPESOD', 'FACTURA'),
                    'codeudores': ('Referencia 1', 'DOCUMENTO_CODEUDOR')
                }
            }