/****************************************************
 * SGI v3.1 - Backend Multi-Boda
 ****************************************************/

function obtenerDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaBodas = ss.getSheetByName("Bodas");
  const hojaInvitados = ss.getSheetByName("Invitados");
  const hojaConfirmaciones = ss.getSheetByName("Confirmaciones");

  if (!hojaBodas || !hojaInvitados || !hojaConfirmaciones) {
    return {
      error: true,
      mensaje: "Faltan hojas requeridas: Bodas, Invitados o Confirmaciones."
    };
  }

  const bodas = leerBodas_(hojaBodas);
  const invitados = leerInvitados_(hojaInvitados);
  const confirmaciones = leerConfirmaciones_(hojaConfirmaciones);

  const bodasActivas = bodas.map(boda =>
    construirResumenBoda_(boda, invitados, confirmaciones)
  );

  return {
    fechaActualizacion: formatearFechaCompleta_(new Date()),
    totalBodas: bodasActivas.length,
    bodasActivas: bodasActivas,
    alertas: construirAlertas_(bodasActivas),
    resumen: construirResumenGeneral_(bodasActivas)
  };
}

/****************************************************
 * LECTURA DE HOJAS
 ****************************************************/

function leerBodas_(hoja) {
  const data = hoja.getDataRange().getValues();
  data.shift();

  return data
    .map(row => ({
      codigo: row[0],
      nombre: row[1],
      fechaBoda: row[2],
      ciudad: row[3],
      inicioConfirmacion: row[4],
      cierreConfirmacion: row[5],
      estado: row[6],
      fotoUrl: row[7],
      tallyUrl: row[8],
      htmlUrl: row[9],
      observaciones: row[10]
    }))
    .filter(boda => boda.codigo);
}

function leerInvitados_(hoja) {
  const data = hoja.getDataRange().getValues();
  data.shift();

  return data
    .map(row => ({
      codigoBoda: row[0],
      idInvitado: row[1],
      invitadoPrincipal: row[2],
      pases: Number(row[3]) || 0,
      celular: row[4],
      categoria: row[5],
      quienInvita: row[6],
      observaciones: row[7]
    }))
    .filter(invitado =>
      invitado.codigoBoda &&
      invitado.idInvitado &&
      invitado.invitadoPrincipal
    );
}

function leerConfirmaciones_(hoja) {
  const data = hoja.getDataRange().getValues();
  data.shift();

  return data
    .map(row => ({
      fecha: row[0],
      codigoBoda: row[1],
      invitado: row[2],
      asiste: normalizarTexto_(row[3]),
      pases: Number(row[4]) || 0,
      pasesLiberados: Number(row[5]) || 0,
      acompanantes: row[6],
      ip: row[7],
      ultimaActualizacion: row[8]
    }))
    .filter(r => r.codigoBoda && r.invitado);
}

/****************************************************
 * DASHBOARD
 ****************************************************/

function construirResumenBoda_(boda, invitados, confirmaciones) {
  const invitadosBoda = invitados.filter(i => i.codigoBoda === boda.codigo);
  const respuestas = confirmaciones.filter(r => r.codigoBoda === boda.codigo);

  const totalPases = invitadosBoda.reduce((sum, i) => sum + i.pases, 0);

  const siAsisten = respuestas
    .filter(r => r.asiste === "si")
    .reduce((sum, r) => sum + r.pases, 0);

  const noAsisten = respuestas
    .filter(r => r.asiste === "no")
    .reduce((sum, r) => sum + r.pasesLiberados, 0);

  const pasesLiberados = respuestas
    .reduce((sum, r) => sum + r.pasesLiberados, 0);

  const pendientes = Math.max(
    totalPases - siAsisten - pasesLiberados,
    0
  );

  const avance = totalPases > 0
    ? Math.round((siAsisten / totalPases) * 100)
    : 0;

  return {
    codigo: boda.codigo,
    nombre: boda.nombre,
    fechaBoda: formatearFechaCompleta_(boda.fechaBoda),
    ciudad: boda.ciudad,
    estado: boda.estado,
    inicioConfirmacion: formatearFechaCompleta_(boda.inicioConfirmacion),
    cierreConfirmacion: formatearFechaCompleta_(boda.cierreConfirmacion),
    fotoUrl: boda.fotoUrl,
    tallyUrl: boda.tallyUrl,
    htmlUrl: boda.htmlUrl,
    totalInvitadosPrincipales: invitadosBoda.length,
    totalPases: totalPases,
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

    if (boda.totalPases > 0 && boda.pendientes > 0) {
      alertas.push(`${etiqueta} tiene ${boda.pendientes} pases pendientes de confirmar.`);
    }

    if (boda.totalPases > 0 && boda.diasCierre <= 7 && boda.diasCierre >= 0) {
      alertas.push(`${etiqueta} cierra confirmaciones en ${boda.diasCierre} días.`);
    }

    if (
      boda.totalPases > 0 &&
      boda.avance < 50 &&
      normalizarTexto_(boda.estado) !== "no iniciado"
    ) {
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

/****************************************************
 * CONFIRMACIÓN DE INVITADOS
 ****************************************************/

function buscarInvitadoConfirmacion(codigoBoda, idInvitado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaInvitados = ss.getSheetByName("Invitados");
  const hojaBodas = ss.getSheetByName("Bodas");

  if (!hojaInvitados || !hojaBodas) {
    return {
      error: true,
      mensaje: "Faltan hojas Invitados o Bodas."
    };
  }

  const invitados = leerInvitados_(hojaInvitados);
  const bodas = leerBodas_(hojaBodas);

  const invitado = invitados.find(i =>
    String(i.codigoBoda).trim() === String(codigoBoda).trim() &&
    String(i.idInvitado).trim() === String(idInvitado).trim()
  );

  if (!invitado) {
    return {
      error: true,
      mensaje: "No se encontró el invitado."
    };
  }

  const boda = bodas.find(b =>
    String(b.codigo).trim() === String(codigoBoda).trim()
  );

  return {
    error: false,
    codigoBoda: codigoBoda,
    idInvitado: idInvitado,
    novios: boda ? boda.nombre : "",
    fotoUrl: boda ? boda.fotoUrl : "",
    fechaBoda: boda ? formatearFechaCompleta_(boda.fechaBoda) : "",
    ciudad: boda ? boda.ciudad : "",
    invitadoPrincipal: invitado.invitadoPrincipal,
    pasesAsignados: invitado.pases
  };
}

function registrarConfirmacionInvitado(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Confirmaciones");

  if (!hoja) {
    return {
      error: true,
      mensaje: "No existe la hoja Confirmaciones."
    };
  }

  const codigoBoda = data.codigoBoda;
  const invitado = data.invitadoPrincipal;
  const asiste = String(data.asiste || "").toUpperCase();
  const pasesAsignados = Number(data.pasesAsignados) || 0;

  const pasesConfirmados = asiste === "NO"
    ? 0
    : Number(data.pasesConfirmados) || 0;

  const pasesLiberados = asiste === "NO"
    ? pasesAsignados
    : Math.max(pasesAsignados - pasesConfirmados, 0);

  hoja.appendRow([
    new Date(),
    codigoBoda,
    invitado,
    asiste,
    pasesConfirmados,
    pasesLiberados,
    data.acompanantes || "",
    "",
    new Date()
  ]);

  return {
    error: false,
    mensaje: "Confirmación registrada correctamente."
  };
}

/****************************************************
 * WEB APP
 ****************************************************/

function doGet(e) {
  const accion = e && e.parameter && e.parameter.accion
    ? e.parameter.accion
    : "dashboard";

  let data;

  if (accion === "dashboard") {
    data = obtenerDashboard();

  } else if (accion === "buscarInvitado") {
    data = buscarInvitadoConfirmacion(
      e.parameter.boda,
      e.parameter.id
    );

  } else if (e && e.parameter && e.parameter.boda && e.parameter.id) {
    data = buscarInvitadoConfirmacion(
      e.parameter.boda,
      e.parameter.id
    );

  } else {
    data = {
      error: true,
      mensaje: "Acción no reconocida."
    };
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let payload = {};

  try {
    payload = JSON.parse(e.postData.contents);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: true,
        mensaje: "Payload inválido."
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = registrarConfirmacionInvitado(payload);

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/****************************************************
 * UTILIDADES
 ****************************************************/

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

/****************************************************
 * PRUEBAS
 ****************************************************/

function probarDashboard() {
  const data = obtenerDashboard();
  Logger.log(JSON.stringify(data, null, 2));
}

function probarBuscarInvitado() {
  const data = buscarInvitadoConfirmacion("B001", "B001-001");
  Logger.log(JSON.stringify(data, null, 2));
}

function probarRegistroConfirmacion() {
  const data = registrarConfirmacionInvitado({
    codigoBoda: "B001",
    idInvitado: "B001-001",
    invitadoPrincipal: "Carlos Pérez",
    asiste: "SI",
    pasesAsignados: 3,
    pasesConfirmados: 2,
    acompanantes: "Acompañante Prueba"
  });

  Logger.log(JSON.stringify(data, null, 2));
}