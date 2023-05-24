# Actualización de Archivo Excel

Este proyecto consiste en una utilidad para actualizar un archivo Excel y crear un archivo CSV con los cambios realizados. Permite formatear columnas específicas, como la fecha, el estado, el tipo de incidencia y la prioridad.


## Instalación

1. Clona o descarga este repositorio en tu máquina local.
2. Asegúrate de tener Python instalado en tu sistema.
3. Instala las dependencias necesarias ejecutando el siguiente comando: pip install openpyxl


## Uso

1. Ejecuta el script `pyython <nombre del archivo>.py` en tu entorno de Python.
2. Se te pedirá que ingreses la ruta del archivo Excel a actualizar y la ruta del archivo CSV a crear.
3. A continuación, se te solicitará ingresar el nombre de la columna que deseas modificar.
4. Ingresa la prioridad requerida cuando se te solicite.
5. El script procesará los cambios y actualizará el archivo Excel original. Además, creará un archivo CSV con los cambios realizados.


## Configuración

Puedes personalizar la configuración del script editando el archivo `actualizar_archivo_excel.py`:

- `columnas`: Mapea las columnas de interés en el archivo Excel con los nombres correspondientes en el archivo CSV.
- `estados_validos`: Define los estados válidos en el archivo Excel y sus equivalentes en el archivo CSV.
- `tipos_incidencia_validos`: Define los tipos de incidencia válidos en el archivo Excel.
- `prioridad`: define el tipo de prioridad que se le da a una tarea.

Asegúrate de guardar los cambios y reiniciar el script después de realizar modificaciones en la configuración.


## Contribución

1. Realiza un fork de este repositorio.
2. Crea una rama con tu nueva funcionalidad: `git checkout -b`.
3. Realiza tus cambios y realiza confirmaciones significativas: `git commit -am 'Añadir nueva funcionalidad'`.
4. Empuja tu rama a tu repositorio remoto.
5. Abre un pull request en este repositorio.
