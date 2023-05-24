import csv
import openpyxl
from datetime import datetime


def formatear_fecha_fin(fecha_fin):
    """
    Formatea la fecha de vencimiento.

    Args:
        fecha_fin (datetime): Fecha de vencimiento.

    Returns:
        str: Fecha formateada.
    """
    if isinstance(fecha_fin, datetime):
        return fecha_fin.strftime('%d-%m-%Y')
    return fecha_fin


def formatear_estado(estado, estados_validos):
    """
    Formatea el estado.

    Args:
        estado (str): Estado a formatear.
        estados_validos (dict): Diccionario que mapea los estados válidos.

    Returns:
        str: Estado formateado.
    """
    return estados_validos.get(estado, estado)


def formatear_tipo_incidencia(tipo_incidencia, prioridad_usuario, tipos_incidencia_validos):
    """
    Formatea el tipo de incidencia.

    Args:
        tipo_incidencia (str): Tipo de incidencia a formatear.
        prioridad_usuario (str): Prioridad proporcionada por el usuario.
        tipos_incidencia_validos (list): Lista de tipos de incidencia válidos.

    Returns:
        str: Tipo de incidencia formateado.
    """
    if tipo_incidencia not in tipos_incidencia_validos:
        raise ValueError(f'Tipo de incidencia no válida: {tipo_incidencia}')
    if tipo_incidencia == 'Error':
        return 'bug'
    elif tipo_incidencia == 'Consulta':
        return 'tarea'
    elif tipo_incidencia == 'Solicitud de mejora':
        return 'subtarea'
    elif tipo_incidencia == 'Requerimiento':
        return prioridad_usuario


def actualizar_archivo_excel(archivo_excel, archivo_csv, columnas, estados_validos, tipos_incidencia_validos):
    """
    Actualiza un archivo Excel y crea un archivo CSV con los cambios realizados.

    Args:
        archivo_excel (str): Ruta del archivo Excel a actualizar.
        archivo_csv (str): Ruta del archivo CSV a crear.
        columnas (dict): Diccionario que mapea las columnas de interés en el archivo Excel con los nombres correspondientes en el archivo CSV.
        estados_validos (dict): Diccionario que mapea los estados válidos en el archivo Excel a los estados correspondientes en el archivo CSV.
        tipos_incidencia_validos (list): Lista de tipos de incidencia válidos en el archivo Excel.
    """
    try:
        # Validar la existencia de los archivos
        if not (archivo_excel and archivo_csv):
            raise ValueError("Debe proporcionar las rutas de archivo válidas.")

        # Cargar el archivo Excel
        workbook = openpyxl.load_workbook(archivo_excel)

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

        # Pedir al usuario la columna a modificar
        columnas_modificar = list(columnas.values())

        # Pedir al usuario la prioridad
        prioridad_usuario = input("Ingresa la prioridad: ")

        # Crear archivo CSV
        with open(archivo_csv, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)

            # Escribir cabecera
            writer.writerow([columnas[columna] for columna in columnas])

            # Iterar sobre las filas y hacer los cambios necesarios
            for row in worksheet.iter_rows(min_row=2):

                # Iterar sobre las columnas y aplicar los cambios necesarios
                for columna_modificar in columnas_modificar:
                    # Obtener el valor actualizado de la columna
                    valor_actualizado = row[indice_columnas[columna_modificar]].value

                    # Realizar el formateo correspondiente según la columna
                    if columna_modificar == 'fecha_vencimiento':
                        row[indice_columnas[columna_modificar]].value = formatear_fecha_fin(valor_actualizado)
                    elif columna_modificar == 'estado':
                        row[indice_columnas[columna_modificar]].value = formatear_estado(valor_actualizado, estados_validos)
                    elif columna_modificar == 'tipo_incidencia':
                        if valor_actualizado:
                            row[indice_columnas[columna_modificar]].value = formatear_tipo_incidencia(valor_actualizado, prioridad_usuario, tipos_incidencia_validos)
                    elif columna_modificar == 'prioridad':
                        row[indice_columnas[columna_modificar]].value = prioridad_usuario

                # Escribir fila actualizada en archivo CSV
                writer.writerow([row[indice_columnas[columna]].value for columna in columnas])

        workbook.save(archivo_excel)
        workbook.close()

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
    'fecha_vencimiento': 'Fecha de vencimiento',
    'prioridad': 'Prioridad'
}
estados_validos = {
    'En progreso': 'En curso',
    'Cerrada': 'Cerrado',
    'Abierta': 'Pendiente'
}
tipos_incidencia_validos = ['Error', 'Consulta', 'Solicitud de mejora', 'Requerimiento']

actualizar_archivo_excel(archivo_excel, archivo_csv, columnas, estados_validos, tipos_incidencia_validos)
