import csv
import openpyxl
from datetime import datetime


def actualizar_archivo_excel(archivo_excel, archivo_csv, columnas, estados_validos, tipos_incidencia_validos, columna_fecha_vencimiento):
    """
    Actualiza un archivo Excel y crea un archivo CSV con los cambios realizados.

    Args:
        archivo_excel (str): Ruta del archivo Excel a actualizar.
        archivo_csv (str): Ruta del archivo CSV a crear.
        columnas (dict): Diccionario que mapea las columnas de interés en el archivo Excel con los nombres correspondientes en el archivo CSV.
        estados_validos (dict): Diccionario que mapea los estados válidos en el archivo Excel a los estados correspondientes en el archivo CSV.
        tipos_incidencia_validos (list): Lista de tipos de incidencia válidos en el archivo Excel.
        columna_fecha_vencimiento (str): Nombre de la columna que contiene la fecha de vencimiento en el archivo Excel.
    """
    # Función para formatear la fecha de vencimiento
    def formatear_fecha_vencimiento(fecha_vencimiento):
        if isinstance(fecha_vencimiento, datetime):
            return fecha_vencimiento.strftime('%d-%m-%Y')
        return fecha_vencimiento

    # Función para formatear el estado
    def formatear_estado(estado):
        return estados_validos.get(estado, estado)

    # Función para formatear el tipo de incidencia
    def formatear_tipo_incidencia(tipo_incidencia):
        if tipo_incidencia not in tipos_incidencia_validos:
            raise ValueError(f'Tipo de incidencia no válida: {tipo_incidencia}')
        if tipo_incidencia == 'Error':
            return 'bug'
        elif tipo_incidencia == 'Consulta':
            return 'tarea'
        elif tipo_incidencia == 'Solicitud de mejora':
            return 'subtarea'

    try:
        # Validar la existencia de los archivos
        if not (archivo_excel and archivo_csv):
            raise ValueError("Debe proporcionar las rutas de archivo válidas.")
        
        # Leer el archivo Excel
        with openpyxl.load_workbook(archivo_excel) as workbook:
            # Seleccionar la hoja de trabajo
            worksheet = workbook.active

            # Buscar las cabeceras de las columnas
            header_row = next(worksheet.iter_rows(min_row=1, max_row=1))
            header = [cell.value for cell in header_row]

            # Obtener los índices de las columnas de interés
            indice_columnas = {}
            for columna, nombre_columna in columnas.items():
                if nombre_columna not in header:
                    raise ValueError(f'Cabecera no encontrada: {nombre_columna}')
                indice_columnas[columna] = header.index(nombre_columna)

            # Crear archivo CSV
            with open(archivo_csv, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)

                # Escribir cabecera
                writer.writerow([columnas[columna] for columna in columnas])

                # Iterar sobre las filas y hacer los cambios necesarios
                for row in worksheet.iter_rows(min_row=2):
                    # Cambiar el formato de la fecha de vencimiento
                    fecha_vencimiento_actualizada = row[indice_columnas[columna_fecha_vencimiento]].value
                    row[indice_columnas[columna_fecha_vencimiento]].value = formatear_fecha_vencimiento(fecha_vencimiento_actualizada)

                    # Cambiar el estado de la tarea
                    estado = row[indice_columnas['estado']].value
                    row[indice_columnas['estado']].value = formatear_estado(estado)

                    # Cambiar el tipo de incidencia si es válido
                    tipo_incidencia = row[indice_columnas['tipo_incidencia']].value
                    if tipo_incidencia:
                        row[indice_columnas['tipo_incidencia']].value = formatear_tipo_incidencia(tipo_incidencia)

                    # Escribir fila actualizada en archivo CSV
                    writer.writerow([row[indice_columnas[columna]].value for columna in columnas])

        print(f"Archivo {archivo_csv} actualizado con éxito!")

    except FileNotFoundError:
        print(f"Error: El archivo {archivo_excel} no existe.")
    except ValueError as error:
        print(f"Error: {error}")


# Pedir al usuario las rutas de los archivos y las columnas de interés
archivo_excel = input("Ingresa la ruta del archivo Excel a actualizar: ")
archivo_csv = input("Ingresa la ruta del archivo CSV a crear: ")
columnas = {
    'estado': 'Estado',
    'tipo_incidencia': 'Tipo de Incidencia',
    'fecha_vencimiento': 'Fecha de vencimiento'
}
estados_validos = {
    'En progreso': 'En curso',
    'Cerrada': 'Cerrado',
    'Abierta': 'Pendiente'
}
tipos_incidencia_validos = ['Error', 'Consulta', 'Solicitud de mejora']
columna_fecha_vencimiento = 'fecha_vencimiento'
actualizar_archivo_excel(archivo_excel, archivo_csv, columnas,
                        estados_validos, tipos_incidencia_validos, columna_fecha_vencimiento)
