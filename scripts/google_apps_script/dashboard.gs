function obtenerDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaBodas = ss.getSheetByName("Bodas");
  const hojaConfirmaciones = ss.getSheetByName("Confirmaciones");

  if (!hojaBodas || !hojaConfirmaciones) {
    return {
      error: true,
      mensaje: "Faltan hojas requeridas: Bodas o Confirmaciones"
    };
  }

  const bodas = leerBodas_(hojaBodas);
  const confirmaciones = leerConfirmaciones_(hojaConfirmaciones);
  const bodasActivas = bodas.map(boda => construirResumenBoda_(boda, confirmaciones));

  return {
    fechaActualizacion: formatearFechaCompleta_(new Date()),
    totalBodas: bodasActivas.length,
    bodasActivas: bodasActivas,
    alertas: construirAlertas_(bodasActivas),
    resumen: construirResumenGeneral_(bodasActivas)
  };
}

function leerBodas_(hoja) {
  const data = hoja.getDataRange().getValues();
  data.shift();

  return data.map(row => ({
    codigo: row[0],
    nombre: row[1],
    fechaBoda: row[2],
    ciudad: row[3],
    estado: row[4],
    inicioConfirmacion: row[5],
    cierreConfirmacion: row[6],
    totalPases: Number(row[7]) || 0
  })).filter(boda => boda.codigo);
}

function leerConfirmaciones_(hoja) {
  const data = hoja.getDataRange().getValues();
  data.shift();

  return data.map(row => ({
    fecha: row[0],
    codigoBoda: row[1],
    invitado: row[2],
    asiste: normalizarTexto_(row[3]),
    pases: Number(row[4]) || 0,
    pasesLiberados: Number(row[5]) || 0
  }));
}

function construirResumenBoda_(boda, confirmaciones) {
  const respuestas = confirmaciones.filter(r => r.codigoBoda === boda.codigo);

  const siAsisten = respuestas
    .filter(r => r.asiste === "si")
    .reduce((sum, r) => sum + r.pases, 0);

  const noAsisten = respuestas
    .filter(r => r.asiste === "no")
    .reduce((sum, r) => sum + r.pasesLiberados, 0);

  const pasesLiberados = respuestas
    .reduce((sum, r) => sum + r.pasesLiberados, 0);

  const pendientes = Math.max(
    boda.totalPases - siAsisten - noAsisten - pasesLiberados,
    0
  );

  const avance = boda.totalPases > 0
    ? Math.round((siAsisten / boda.totalPases) * 100)
    : 0;

  return {
    codigo: boda.codigo,
    nombre: boda.nombre,
    fechaBoda: formatearFechaCompleta_(boda.fechaBoda),
    ciudad: boda.ciudad,
    estado: boda.estado,
    inicioConfirmacion: formatearFechaCompleta_(boda.inicioConfirmacion),
    cierreConfirmacion: formatearFechaCompleta_(boda.cierreConfirmacion),
    totalPases: boda.totalPases,
    siAsisten: siAsisten,
    noAsisten: noAsisten,
    pendientes: pendientes,
    pasesLiberados: pasesLiberados,
    avance: avance,
    diasCierre: calcularDiasHasta_(boda.cierreConfirmacion),
    evolucion: construirEvolucion_(respuestas)
  };
}

function construirResumenGeneral_(bodas) {
  return bodas.reduce((acc, boda) => {
    acc.invitados += boda.totalPases;
    acc.confirmados += boda.siAsisten;
    acc.pendientes += boda.pendientes;
    acc.noAsistiran += boda.noAsisten;
    acc.pasesLiberados += boda.pasesLiberados;
    return acc;
  }, {
    invitados: 0,
    confirmados: 0,
    pendientes: 0,
    noAsistiran: 0,
    pasesLiberados: 0
  });
}

function construirAlertas_(bodas) {
  const alertas = [];

  bodas.forEach(boda => {
    const etiqueta = `${boda.codigo} - ${boda.nombre}`;

    if (boda.pendientes > 0) {
      alertas.push(`${etiqueta} tiene ${boda.pendientes} pases pendientes de confirmar.`);
    }

    if (boda.diasCierre <= 7 && boda.diasCierre >= 0) {
      alertas.push(`${etiqueta} cierra confirmaciones en ${boda.diasCierre} días.`);
    }

    if (boda.avance < 50 && boda.estado !== "No iniciado") {
      alertas.push(`${etiqueta} tiene avance menor al 50%.`);
    }
  });

  return alertas;
}

function construirEvolucion_(respuestas) {
  const agrupado = {};

  respuestas
    .filter(r => r.asiste === "si")
    .forEach(r => {
      const fecha = formatearFechaCorta_(r.fecha);
      agrupado[fecha] = (agrupado[fecha] || 0) + r.pases;
    });

  return Object.keys(agrupado).map(fecha => ({
    fecha: fecha,
    pases: agrupado[fecha]
  }));
}

function calcularDiasHasta_(fecha) {
  if (!fecha) return 0;

  const hoy = new Date();
  const objetivo = new Date(fecha);

  return Math.ceil((objetivo - hoy) / (1000 * 60 * 60 * 24));
}

function formatearFechaCorta_(fecha) {
  if (fecha instanceof Date) {
    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM");
  }

  return String(fecha || "");
}

function formatearFechaCompleta_(fecha) {
  if (fecha instanceof Date) {
    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }

  return String(fecha || "");
}

function normalizarTexto_(valor) {
  return String(valor || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function probarDashboard() {
  const data = obtenerDashboard();
  Logger.log(JSON.stringify(data, null, 2));
}

function doGet() {
  const data = obtenerDashboard();

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}