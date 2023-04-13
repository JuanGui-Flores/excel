// Importa la biblioteca SheetJS
import * as XLSX from 'xlsx';

// Función para cambiar el formato de la fecha
function cambiarFormatoFecha(fecha) {
  const dia = ("0" + fecha.getDate()).slice(-2);
  const mes = ("0" + (fecha.getMonth() + 1)).slice(-2);
  const año = fecha.getFullYear();
  return `${dia}/${mes}/${año}`;
}

// Función para cambiar el valor del Estado
function cambiarValorEstado(estado) {
  if (estado === "Planificada") {
    return "Funcionalidad Planificada";
  } else {
    return "Funcionalidad No Planificada";
  }
}

// Función para cambiar el valor del Tipo de Incidencia
function cambiarValorTipoIncidencia(tipo) {
  if (tipo === "Tarea Planificada") {
    return "Tarea Planificada";
  } else {
    return "Tarea No Planificada";
  }
}

// Función principal
function procesarArchivo(event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = (event) => {
    const arraybuffer = event.target.result;

    // Convierte el archivo Excel en un objeto de hoja de cálculo
    const workbook = XLSX.read(arraybuffer, { type: 'array' });

    // Obtiene la primera hoja de cálculo
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    // Obtiene los datos de la hoja de cálculo como una matriz de objetos
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Itera sobre cada fila de datos
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Cambia el formato de la fecha
      const fecha = new Date(row[0]);
      row[0] = cambiarFormatoFecha(fecha);

      // Cambia el valor del Estado
      row[1] = cambiarValorEstado(row[1]);

      // Cambia el valor del Tipo de Incidencia
      row[2] = cambiarValorTipoIncidencia(row[2]);
    }

    // Convierte los datos modificados en un objeto de hoja de cálculo
    const new_workbook = XLSX.utils.book_new();
    const new_worksheet = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(new_workbook, new_worksheet, "Datos modificados");

    // Genera un nombre aleatorio para el archivo Excel modificado
    const timestamp = new Date().getTime();
    const fileName = `datos-modificados-${timestamp}.xlsx`;

    // Guarda el archivo Excel modificado en la carpeta "Archivos Modificados"
    XLSX.writeFile(new_workbook, `Work/${fileName}`);
  };

  reader.readAsArrayBuffer(file);
}

// Agrega un escucha de eventos al botón "Procesar Archivo"
const btnProcesarArchivo = document.getElementById("btnProcesarArchivo");
btnProcesarArchivo.addEventListener("click", procesarArchivo);
