function doGet() {
  return HtmlService.createTemplateFromFile('Mesas')
    .evaluate()
    .setTitle('🤍 Maria Claudia y Gustavo 🤍')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function obtenerDatosIniciación() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Data_Mesas');
  const datos = hoja.getDataRange().getValues();
  datos.shift(); 
  
  return datos.map((fila, index) => {
    // IMPORTANTE:
    // fila[0] = Columna A (Nombre)
    // fila[1] = Columna B (Número de Pases) -> Usado para el contador
    // fila[2] = Columna C (Número de Mesa) -> Usado para ubicar en el mapa
    
    let mesaNum = parseInt(fila[2]); 
    if (isNaN(mesaNum) || mesaNum < 1 || mesaNum > 10) {
      mesaNum = 0;
    }

    return {
      nombre: fila[0],
      pases: parseInt(fila[1]) || 1, // <--- Aquí tomamos los pases de la Columna B
      mesa: mesaNum,
      fila: index + 2
    };
  }).filter(g => g.nombre && g.nombre.toString().trim() !== "");
}

function guardarTodo(asignaciones) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Data_Mesas'); 
  
  asignaciones.forEach(item => {
    // Escribimos el resultado en la Columna C (Número 3)
    hoja.getRange(item.fila, 3).setValue(item.mesa > 0 ? item.mesa : "");
  });
  
  SpreadsheetApp.flush();
  return "¡Distribución guardada en la columna C de Data_Mesas!";
}

function resetearExcel(filas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Data_Mesas');
  filas.forEach(f => hoja.getRange(f, 3).setValue(""));
  return "Mesas reseteadas.";
}

