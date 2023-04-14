import argparse
import openpyxl
from datetime import datetime

# Definir argumentos de línea de comandos
parser = argparse.ArgumentParser(description='Actualizar archivo Excel')
parser.add_argument('-i', '--input', type=str, help='Archivo de entrada (formato .xlsx)', required=True)
parser.add_argument('-o', '--output', type=str, help='Archivo de salida (formato .csv)', required=True)
args = parser.parse_args()

# Configuración
archivoExcel = args.input
archivoCSV = args.output
columnaFechaVencimiento = 'Fecha de vencimiento'
columnaEstado = 'Estado'
valorEstadoProduccion = 'PRODUCCION'
valorEstadoCerrado = 'CERRADO'
columnaTipoIncidencia = 'Tipo de Incidencia'
tiposIncidencia = [
    'Funcionalidad Planificada',
    'Tarea Planificada',
    'Tarea No Planificada'
]

# Función para formatear la fecha
def formatearFecha(fecha):
    return fecha.strftime('%d%m%Y')

# Función para formatear el estado
def formatearEstado(estado):
    if estado == valorEstadoProduccion:
        return valorEstadoCerrado
    return estado

# Función para formatear el tipo de incidencia
def formatearTipoIncidencia(tipoIncidencia):
    if tipoIncidencia not in tiposIncidencia:
        raise ValueError(f'Tipo de incidencia no válida: {tipoIncidencia}')
    return tipoIncidencia

# Función para actualizar una fila de datos
def actualizarFila(fila):

    # Cambiar el formato de la fecha
    fechaActualizada = fila[worksheet[columnaFechaVencimiento].column_letter + str(fila.row)].value
    if isinstance(fechaActualizada, datetime):
        fila[worksheet[columnaFechaVencimiento].column_letter + str(fila.row)].value = formatearFecha(fechaActualizada)

    # Cambiar el estado de la tarea
    estado = fila[worksheet[columnaEstado].column_letter + str(fila.row)].value
    fila[worksheet[columnaEstado].column_letter + str(fila.row)].value = formatearEstado(estado)

    # Cambiar el tipo de incidencia si es válido
    tipoIncidencia = fila[worksheet[columnaTipoIncidencia].column_letter + str(fila.row)].value
    if tipoIncidencia:
        fila[worksheet[columnaTipoIncidencia].column_letter + str(fila.row)].value = formatearTipoIncidencia(tipoIncidencia)

    return fila

try:
    # Leer el archivo Excel
    workbook = openpyxl.load_workbook(archivoExcel)

    # Seleccionar la hoja de trabajo
    sheetName = workbook.sheetnames[0]
    worksheet = workbook[sheetName]

    # Iterar sobre las filas y hacer los cambios necesarios
    for row in worksheet.iter_rows(min_row=2):
        actualizarFila(row)

    # Guardar el archivo Excel actualizado
    workbook.save(archivoCSV)
    print(f"Archivo {archivoCSV} actualizado con éxito!")

except FileNotFoundError:
    print(f"Error: El archivo {archivoExcel} no existe.")
except Exception as e:
    print(f"Error al actualizar el archivo: {e}")
