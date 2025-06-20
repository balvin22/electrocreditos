from dataclasses import dataclass
from typing import Dict, List

@dataclass
class AppConfig:
    title: str = "Procesador de Reportes Financieros"
    geometry: str = "900x450"
    bg_color: str = "#f0f0f0"
    accent_color: str = "#4b6cb7"
    secondary_color: str = "#2d3747"
    text_color: str = "#333333"
    output_filename: str = "reporte_financiero.xlsx"

@dataclass
class FinancieroProcessingConfig:
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
                'AC FS': {'CEDULA': 'CEDULA_FS', 'FACTURA': 'FACTURA_FS', 'saldofac': 'SALDO_FS', 'ccosto': 'CENTRO_COSTO_FS'},
                'AC ARP': {'CEDULA': 'CEDULA_ARP', 'FACTURA': 'FACTURA_ARP', 'saldofac': 'SALDO_ARP', 'ccosto': 'CENTRO_COSTO_ARP'},
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
                    'casa_cobranza': ('FACTURA FINAL', 'FACTURA'),
                    'codeudores': ('Referencia 1', 'DOCUMENTO_CODEUDOR')
                }
            }
        # if self.column_order_efecty is None:
        #     self.columnn_order_efecty =[
        #         'No', 'Identificación', 'Valor', 'N° de Autorización', 'Fecha',
        #         'Documento Cartera', 'C. Costo', 'Empresa', 'Valor Aplicar',
        #         'Valor Anticipos', 'Valor Aprovechamientos', 'Casa cobranza',
        #         'Empleado', 'Novedad', 'Cuentas ARP', 'Cuentas FS', 'SALDOS', 'VALIDACION ULTIMO SALDO'
        #     ]
        # if self.column_order_bancolombia is None:
        #     self.column_order_bancolomnia = [
        #         'No.', 'Fecha', 'Detalle 1', 'Detalle 2', 'Referencia 1', 'Referencia 2',
        #         'Valor', 'Documento Cartera', 'C. Costo', 'Empresa', 'Valor Aplicar',
        #         'Valor Anticipos', 'Valor Aprovechamientos', 'Casa cobranza',
        #         'Empleado', 'Novedad', 'Cuentas ARP', 'Cuentas FS', 'SALDOS', 'VALIDACION ULTIMO SALDO'
        #     ]        
            
            
@dataclass
class AnticiposConfig:
    required_sheets: List[str] = None
    sheet_columns: Dict[str, List[str]] = None                                                                                                                                                                   
    rename_columns: Dict[str, str] = None
    merge_config: Dict[str, Dict[str, str]] = None
    output_filename: str = "reporte_anticipos.xlsx"
    column_order_fs: List[str] = None
    column_order_arp: List[str] = None

    def __post_init__(self):
        if self.required_sheets is None:
            self.required_sheets = ['ONLINE','AC FS','AC ARP']

        if self.sheet_columns is None:
            self.sheet_columns = {
                'ONLINE': ['MCNTIPCRU1', 'MCNNUMCRU1', 'MCNVINCULA', 'VINNOMBRE','SALDODOC',],
                'AC FS': ['cobra', 'ccosto','FACTURA','CEDULA','saldofac'],
                'AC ARP': ['cobra', 'ccosto','FACTURA','CEDULA','saldofac']
            }

        if self.rename_columns is None:
            self.rename_columns = {
                'ONLINE': {
                    'MCNTIPCRU1': 'TIPO_RECIBO',
                    'MCNNUMCRU1': 'No',
                    'MCNVINCULA': 'CEDULA',
                    'VINNOMBRE': 'NOMBRE',
                    'SALDODOC': 'VALOR'
                },
                'AC FS': {
                    'saldofac':'ULTIMO_SALDO_FS',
                    'cobra': 'ZONA_COBRADOR_FS',
                    'ccosto': 'CENTRO_COSTO_FS',
                    'FACTURA': 'FACTURA_FS',
                    'CEDULA': 'CEDULA'  
                },
                'AC ARP': {
                    'saldofac':'ULTIMO_SALDO_ARP',
                    'cobra': 'ZONA_COBRADOR_ARP',
                    'ccosto': 'CENTRO_COSTO_ARP',
                    'FACTURA': 'FACTURA_ARP',
                    'CEDULA': 'CEDULA'  
                }
            }
        if self.column_order_fs is None:
            self.column_order_fs = [
                'ITEM',  'TIPO_RECIBO', 'No', 'CEDULA', 'NOMBRE', 'CENTRO_COSTO_FS', 
                'VALOR', 'FACTURA_FS', 'ZONA_COBRADOR_FS','OBSERVACIONES','CUENTAS_FS',
                'ULTIMO_SALDO_FS','VALOR_POSITIVO','RESTA_SALDO'
            ]

        if self.column_order_arp is None:
            self.column_order_arp = [
                'ITEM',  'TIPO_RECIBO', 'No', 'CEDULA', 'NOMBRE', 'CENTRO_COSTO_ARP', 
                'VALOR', 'FACTURA_ARP', 'ZONA_COBRADOR_ARP','OBSERVACIONES','CUENTAS_ARP',
                'ULTIMO_SALDO_ARP','VALOR_POSITIVO','RESTA_SALDO'
            ]
            