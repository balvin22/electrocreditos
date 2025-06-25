import pandas as pd

# Se definen las especificaciones de las columnas (posiciones de inicio y fin) y los nombres de las columnas.
# He analizado tus archivos para determinar estas posiciones, pero puede que necesites ajustarlas si algo no se
# alinea perfectamente.
input_file_path = 'c:/Users/usuario/Desktop/Reporte LV/DATA MARZO FS.TXT'
colspecs = [
    (0, 1), (1, 12), (30, 75), (12, 30), (75, 76), (76, 86), (86, 122),
    (122, 153), (153, 163), (163, 174), (174, 185), (185, 196), (196, 206),
    (206, 217), (217, 221), (221, 225), (225, 229), (229, 237), (237, 245),
    (245, 275), (275, 335), (335, 435), (435, 445), (445, 455),
    (455, 465), (465, 473), (473, 476), (476, 484), (484, 492),
    (492, 522), (522, 530)
]

names = [
    "TIPO DE IDENTIFICACION", "NUMERO DE IDENTIFICACION", "NOMBRE COMPLETO",
    "NUMERO DE LA CUENTA U OBLIGACION", "FECHA APERTURA", "FECHA VENCIMIENTO",
    "RESPONSABLE", "NOVEDAD", "ESTADO ORIGEN DE LA CUENTA", "VALOR INICIAL",
    "VALOR SALDO DEUDA", "VALOR DISPONIBLE", "V CUOTA MENSUAL",
    "VALOR SALDO MORA", "TOTAL CUOTAS", "CUOTAS CANCELADAS", "CUOTAS EN MORA",
    "FECHA LIMITE DE PAGO", "FECHA DE PAGO", "CIUDAD CORRESPONDENCIA",
    "DIRECCION DE CORRESPONDENCIA", "CORREO ELECTRONICO", "CELULAR",
    "SITUACION DEL TITULAR", "EDAD DE MORA", "FORMA DE PAGO",
    "FECHA ESTADO ORIGEN", "ESTADO DE LA CUENTA", "FECHA ESTADO DE LA CUENTA",
    "ADJETIVO", "FECHA DE ADJETIVO"
]

# Leer el archivo de ancho fijo, omitiendo la primera fila (encabezado)
try:
    df = pd.read_fwf(
        input_file_path,
        colspecs=colspecs,
        names=names,
        encoding='latin-1',  # Se utiliza la codificación 'latin-1' para evitar errores con caracteres especiales
        skiprows=1
    )

    # Guardar el DataFrame en un archivo de Excel
    output_filename = "Data_marzo_fs_convertido.xlsx"
    df.to_excel(output_filename, index=False)

    print(f"¡El archivo se ha convertido y guardado como '{output_filename}' exitosamente!")

except FileNotFoundError:
    print("Error: No se encontró el archivo 'DATA MARZO FS.TXT'. Asegúrate de que el archivo esté en el mismo directorio que el script.")
except Exception as e:
    print(f"Ha ocurrido un error inesperado: {e}")