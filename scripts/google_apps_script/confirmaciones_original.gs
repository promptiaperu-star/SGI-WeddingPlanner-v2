function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Respuestas") || ss.getSheets()[0];
  
  var fecha = new Date();
  // El script ahora lee los datos enviados por la URL
  var nombre = e.parameter.p_nombre || "Desconocido";
  var asiste = e.parameter.p_asiste || "Si";
  var pases  = e.parameter.p_pases  || "1";
  var acomp  = e.parameter.p_acomp  || "Ninguno";

  sheet.appendRow([fecha, nombre, asiste, pases, acomp]);
  
  return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
}

