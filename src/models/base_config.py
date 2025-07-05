# -*- coding: utf-8 -*-
"""
Este archivo centraliza toda la configuración del proyecto, incluyendo:
- La estructura y reglas de mapeo para cada tipo de archivo.
- Las rutas de los archivos de entrada.
"""

# --- 1. CONFIGURACIÓN DE PROCESAMIENTO POR TIPO DE ARCHIVO ---
configuracion = {
    "ANALISIS": {
        "usecols": ["direccion", "barrio", "nomciudad", "totcuotas", "valorcuota", "diasatras", "cuotaspag","cedula","saldofac","tipo","numero"],
        "rename_map": { "direccion": "Direccion", "barrio": "Barrio", "nomciudad": "Nombre_Ciudad", "totcuotas": "Total_Cuotas", "valorcuota": "Valor_Cuota", "diasatras": "Dias_Atraso", "cuotaspag": "Cuotas_Pagadas", "cedula" : "Cedula_Cliente", "saldofac":"Saldo_Factura", "tipo":"Tipo_Credito", "numero":"Numero_Credito" }
    },
    "R91": {
        "usecols": ["VINNOMBRE", "VENNOMBRE", "MCDZONA", "MCDVINCULA", "MCDNUMCRU1", "MCDTIPCRU1","VENOMBRE","VENCODIGO", "COBNOMBRE", "MCDCCOSTO", "CCONOMBRE", "META_INTER", "META_DC_AL", "META_DC_AT", "META_ATRAS"],
        "rename_map": { "MCDTIPCRU1": "Tipo_Credito", "MCDNUMCRU1": "Numero_Credito", "MCDVINCULA" : "Cedula_Cliente", "VINNOMBRE": "Nombre_Cliente", "MCDZONA" : "Zona", "COBNOMBRE" : "Nombre_Cobrador", "VENNOMBRE" : "Nombre_Vendedor", "VENCODIGO" : "Codigo_Vendedor", "CCONOMBRE" : "Centro_Costos", "MCDCCOSTO" : "Codigo_Centro_Costos", "META_INTER" : "Meta_Intereses", "META_DC_AL" : "Meta_DC_Al_Dia", "META_DC_AT" : "Meta_DC_Atraso", "META_ATRAS" : "Meta_Atraso" }
    },
    "VENCIMIENTOS": {
        "usecols": ["MCNVINCULA", "VINTELEFO3", "SALDODOC", "VENCE", "VINTELEFON", "MCNCUOCRU1","MCNTIPCRU1","MCNNUMCRU1"],
        "rename_map": {"MCNTIPCRU1":"Tipo_Credito", "MCNNUMCRU1":"Numero_Credito", "MCNVINCULA": "Cedula_Cliente", "VINTELEFO3": "Celular", "VINTELEFON" : "Telefono", "SALDODOC": "Valor_Cuota_Vigente", "MCNCUOCRU1": "Cuota_Vigente", "VENCE": "Fecha_Vencimiento" }
    },
    "R03":{
        "usecols": ["CODEUDOR1","NOMBRE1","VINTELEFON","CIUNOMBRE1","CODEUDOR2","NOMBRE2","VINTELEFO2","CIUNOMBRE2","CEDULA"],
        "rename_map": { "CODEUDOR1": "Codeudor1", "NOMBRE1": "Nombre_Codeudor1", "VINTELEFON": "Telefono_Codeudor1", "CIUNOMBRE1": "Ciudad_Codeudor1", "CODEUDOR2": "Codeudor2", "NOMBRE2": "Nombre_Codeudor2", "VINTELEFO2": "Telefono_Codeudor2", "CIUNOMBRE2": "Ciudad_Codeudor2", "CEDULA": "Cedula_Cliente" }
    },
    "CRTMPCONSULTA1":{
        "usecols":["CORREO","FECHA_FACT","TIPO_DOCUM","NUMERO_DOC","IDENTIFICA"],
        "rename_map":{ "CORREO": "Correo", "FECHA_FACT":"Fecha_Facturada", "TIPO_DOCUM":"Tipo_Credito", "NUMERO_DOC":"Numero_Credito", "IDENTIFICA":"Cedula_Cliente" }
    },
    "FNZ003":{
        "usecols":["CREDITO","CONCEPTO","SALDO"],
        "rename_map":{ "CREDITO":"Credito", "CONCEPTO":"Concepto", "SALDO":"Saldo" }
    },
      "MATRIZ_CARTERA": {
        "skiprows": 2, "header": None, 
        "new_names": ['Zona', 'Cobrador', 'telefono_cobrador', 'Regional', 'Gestor', 'gestor_telefono', 'call_center_1_30_dias', 'call_center_nombre_1_30', 'call_center_telefono_1_30', 'call_center_31_90_dias', 'call_center_nombre_31_90', 'call_center_telefono_31_90', 'call_center_91_360_dias', 'call_center_nombre_91_360', 'call_center_telefono_91_360'],
        "merge_on": "Zona" 
      },
      "ASESORES": {
        "sheets": [
            { "sheet_name": "ASESORES", "usecols": ["NOMBRE ASESOR", "MOVIL ASESOR", "LIDER ZONA", "MOVIL LIDER"], "rename_map": { "NOMBRE ASESOR": "Nombre_Vendedor", "MOVIL ASESOR": "Movil_Asesor", "LIDER ZONA": "Lider_Zona", "MOVIL LIDER": "Movil_Lider" }, "merge_on": "Nombre_Vendedor" },
            { "sheet_name": "Centro Costos", "usecols": ["CENTRO DE COSTOS", "ACTIVO"], "rename_map": { "CENTRO DE COSTOS": "Codigo_Centro_Costos", "ACTIVO": "Activo_Centro_Costos" }, "merge_on": "Codigo_Centro_Costos" }
        ]
    }
}

# --- 2. LISTA DE ARCHIVOS A PROCESAR ---
# Asegúrate de que esta ruta sea correcta para tu sistema
ruta_base = 'C:/Users/usuario/Desktop/JUNIO/' 
archivos_a_procesar = [
    ruta_base + "ANALISIS ARP GENERAL 0506INICIAL.XLS",
    ruta_base + "ANALISIS FNS GENERAL 0506INICIAL.XLS",
    ruta_base + "R91 ARP JUNIO.XLSX",
    ruta_base + "R91 FS JUNIO.XLSX",
    ruta_base + "VENCIMIENTOS ARP JUNIO.XLSX",
    ruta_base + "VENCIMIENTOS FNS JUNIO.XLSX",
    ruta_base + "R03 2025 FNS.xlsx",
    ruta_base + "R03 2025 ARP.xlsx",
    ruta_base + "CRTMPCONSULTA1.xlsx",
    ruta_base + "FNZ003 A 20 JUN.XLSX",
    ruta_base + "MATRIZ DE CARTERA.xlsx",
    ruta_base + "ASESORES ACTIVOS.xlsx"
]