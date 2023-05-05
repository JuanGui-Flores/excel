import csv
import openpyxl
from datetime import datetime


def actualizar_archivo_excel(archivo_excel, archivo_csv, columnas, estados_validos, tipos_incidencia_validos, columna_fecha):
    
    # Función para formatear la fecha
    def formatear_fecha(fecha):
        if isinstance(fecha, datetime):
            return fecha.strftime('%d%m%Y')
        return fecha

    # Función para formatear el estado
    def formatear_estado(estado):
        return estados_validos.get(estado, estado)

    # Función para formatear el tipo de incidencia
    def formatear_tipo_incidencia(tipo_incidencia):
        if tipo_incidencia not in tipos_incidencia_validos:
            raise ValueError(
                f'Tipo de incidencia no válida: {tipo_incidencia}')
        return tipo_incidencia

    try:
        # Leer el archivo Excel
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

        # Crear archivo CSV
        with open(archivo_csv, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)

            # Escribir cabecera
            writer.writerow([columnas[columna] for columna in columnas])

            # Iterar sobre las filas y hacer los cambios necesarios
            for row in worksheet.iter_rows(min_row=2):
                
                # Cambiar el formato de la fecha
                fecha_actualizada = row[indice_columnas[columna_fecha]].value
                row[indice_columnas[columna_fecha]
                    ].value = formatear_fecha(fecha_actualizada)

                # Cambiar el estado de la tarea
                estado = row[indice_columnas['estado']].value
                row[indice_columnas['estado']].value = formatear_estado(estado)

                # Cambiar el tipo de incidencia si es válido
                tipo_incidencia = row[indice_columnas['tipo_incidencia']].value
                if tipo_incidencia:
                    row[indice_columnas['tipo_incidencia']
                        ].value = formatear_tipo_incidencia(tipo_incidencia)

                # Escribir fila actualizada en archivo CSV
                writer.writerow(
                    [row[indice_columnas[columna]].value for columna in columnas])

        workbook.save()
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
    'fecha': 'Fecha',
    'estado': 'Estado',
    'tipo_incidencia': 'Tipo de incidencia',
    'descripcion': 'Descripción'
}
estados_validos = {
    'En progreso': 'En curso',
    'Cerrada': 'Completada',
    'Abierta': 'Pendiente'
}
tipos_incidencia_validos = ['Error', 'Consulta', 'Solicitud de mejora']
columna_fecha = 'fecha'
actualizar_archivo_excel(archivo_excel, archivo_csv, columnas,
                        estados_validos, tipos_incidencia_validos, columna_fecha)
