/****************************************************
 * SGI v3.2 - Backend Multi-Boda Operativo
 ****************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SGI")
    .addItem("Generar enlaces y mensajes", "generarLinksYMensajes")
    .addSeparator()
    .addItem("Actualizar Centro de Envíos", "actualizarCentroEnvios")
    .addItem("Abrir WhatsApp del invitado actual", "abrirWhatsappInvitadoActual")
    .addItem("Marcar invitado actual como enviado", "marcarInvitadoActualEnviado")
    .addToUi();
}

function doGet(e) {
  const accion = e && e.parameter && e.parameter.accion
    ? String(e.parameter.accion).trim()
    : "dashboard";

  let data;

  try {
    if (accion === "dashboard") {
      data = obtenerDashboard();
    } else if (accion === "buscarInvitado") {
      data = buscarInvitadoConfirmacion(e.parameter.boda, e.parameter.id);
    } else if (accion === "listarInvitadosBoda") {
      data = listarInvitadosBoda(e.parameter.boda);
    } else if (accion === "listarEnviosManual") {
      data = listarEnviosManual(e.parameter.boda);
    } else if (accion === "reporteInvitados") {
      data = reporteInvitados(e.parameter.boda);
    } else if (accion === "descargarListaInvitadosXlsx") {
      data = descargarListaInvitadosXlsx(e.parameter.boda);
    } else if (accion === "descargarReporteInvitadosXlsx") {
      data = descargarReporteInvitadosXlsx(e.parameter.boda);
    } else if (accion === "marcarEnviadoManual") {
      data = marcarEnviadoManual(e.parameter.boda, e.parameter.id);
    } else if (accion === "listarMesas") {
      data = listarMesas_(e.parameter.boda);
    } else if (accion === "listarInvitadosMesas") {
      data = listarInvitadosMesas_(e.parameter.boda);
    } else if (accion === "guardarAsignacionMesas") {
      data = guardarAsignacionMesas_(e.parameter.boda, e.parameter.payload);
    } else if (accion === "reiniciarAsignacionMesas") {
      data = reiniciarAsignacionMesas_(e.parameter.boda);
    } else if (accion === "reporteMesas") {
      data = reporteMesas_(e.parameter.boda);
    } else if (e && e.parameter && e.parameter.boda && e.parameter.id) {
      data = buscarInvitadoConfirmacion(e.parameter.boda, e.parameter.id);
    } else {
      data = { error: true, mensaje: "Acción no reconocida." };
    }
  } catch (error) {
    data = { error: true, mensaje: error.toString() };
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
    return jsonOutput_({
      error: true,
      mensaje: "Payload inválido."
    });
  }

  let data;

  try {
    if (payload.accion === "marcarEnviadoManual") {
      data = marcarEnviadoManual(payload.codigoBoda, payload.idInvitado);

    } else if (payload.accion === "guardarAsignacionMesas") {
      data = guardarAsignacionMesas_(payload.boda, payload.payload);

    } else {
      data = registrarConfirmacionInvitado(payload);
    }

  } catch (error) {
    data = {
      error: true,
      mensaje: error.toString()
    };
  }

  return jsonOutput_(data);
}

function jsonOutput_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function obtenerDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shBodas = ss.getSheetByName("Bodas");
  const shInv = ss.getSheetByName("Invitados");
  const shConf = ss.getSheetByName("Confirmaciones");

  if (!shBodas || !shInv || !shConf) {
    return { error: true, mensaje: "Faltan hojas requeridas: Bodas, Invitados o Confirmaciones." };
  }

  const bodasData = shBodas.getDataRange().getValues();
  const invData = shInv.getDataRange().getValues();
  const confData = shConf.getDataRange().getValues();

  const hB = bodasData[0];
  const hI = invData[0];
  const hC = confData[0];

  const idxB = obtenerIndicesSGI_(hB);
  const idxI = obtenerIndicesSGI_(hI);
  const idxC = obtenerIndicesSGI_(hC);

  const bodas = bodasData.slice(1)
    .filter(r => r[idxB.codigo_boda])
    .map(b => {
      const codigo = String(b[idxB.codigo_boda]).trim();

      const invitados = invData.slice(1)
        .filter(r => String(r[idxI.codigo_boda]).trim() === codigo);

      const confirmaciones = confData.slice(1)
        .filter(r => String(r[idxC.codigo_boda]).trim() === codigo);

      const totalPases = invitados.reduce((s, r) => s + (Number(r[idxI.pases]) || 0), 0);

      let siAsisten = 0;
      let noAsisten = 0;
      let noUtilizados = 0;

      confirmaciones.forEach(r => {
        const idInvitado = String(r[idxC.id_invitado] || "").trim();
        const nombre = String(r[idxC.invitado] || "").trim();
        const asiste = String(r[idxC.asiste] || "").trim().toUpperCase();
        const pasesConfirmados = Number(r[idxC.pases]) || 0;
        const pasesNoUtilizados = Number(r[idxC.pases_noutilizados]) || 0;

        const inv = invitados.find(i => {
         const idInvLista = String(i[idxI.id_invitado] || "").trim();
         const nombreLista = String(i[idxI.invitado_principal] || "").trim().toUpperCase();

         return idInvitado
          ? idInvLista === idInvitado
          : nombreLista === nombre.toUpperCase();
        });

        const pasesAsignados = inv ? (Number(inv[idxI.pases]) || 0) : pasesConfirmados;

        if (asiste === "SI") {
          siAsisten += pasesConfirmados;
          noUtilizados += pasesNoUtilizados || Math.max(pasesAsignados - pasesConfirmados, 0);
        }

        if (asiste === "NO") {
          noAsisten += pasesAsignados;
        }
      });

      const pendientes = Math.max(totalPases - siAsisten - noAsisten - noUtilizados, 0);
      const liberados = noAsisten + noUtilizados;
      const avance = totalPases ? Math.round((siAsisten / totalPases) * 100) : 0;

      const estadoCalculado =
       avance >= 100 ? "Confirmación completa" :
       pendientes > 0 && Number(calcularDiasCierreSGI_(b[idxB.cierre_confirmacion])) <= 0 ? "Cierre vencido con pendientes" :
       pendientes > 0 ? "Esperando respuestas" :
       "Sin pendientes";

      const evolucion = construirEvolucionSGI_(confirmaciones, idxC);
      const nombreBoda = b[idxB.novios] || "";
      const alertas = construirAlertasSGI_(codigo, nombreBoda, totalPases, siAsisten, noAsisten, pendientes, noUtilizados);

      return {
        codigo,
        nombre: b[idxB.novios] || "",
        estado: b[idxB.estado] || estadoCalculado,
        fechaBoda: formatearFechaLargaSGI_(b[idxB.fecha_boda]),
        ciudad: b[idxB.ciudad] || "",
        fechaEnvio: formatearFechaLargaSGI_(b[idxB.inicio_confirmacion]),
        cierreConfirmacion: formatearFechaLargaSGI_(b[idxB.cierre_confirmacion]),
        diasCierre: calcularDiasCierreSGI_(b[idxB.cierre_confirmacion]),
        fotoUrl: b[idxB.foto_url] || "",
        totalPases,
        siAsisten,
        noAsisten,
        pendientes,
        noUtilizados,
        pasesLiberados: liberados,
        liberados,
        avance,
        evolucion,
        alertas
      };
    });

  return {
  error: false,
  bodasActivas: bodas,
  totalBodas: bodas.length,
  alertasCriticas: bodas.flatMap(b => b.alertas || [])
  };
}

function construirEvolucionSGI_(confirmaciones, idxC) {
  const acumulado = {};

  confirmaciones.forEach(r => {
    const asiste = String(r[idxC.asiste] || "").trim().toUpperCase();
    if (asiste !== "SI") return;

    const fecha = r[idxC.fecha];
    const fechaTxt = fecha instanceof Date
      ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM")
      : String(fecha || "").substring(0, 5);

    acumulado[fechaTxt] = (acumulado[fechaTxt] || 0) + (Number(r[idxC.pases]) || 0);
  });

  let suma = 0;

  return Object.keys(acumulado).sort().map(f => {
    suma += acumulado[f];
    return {
      fecha: f,
      pases: suma
    };
  });
}

function construirAlertasSGI_(codigo, nombreBoda, totalPases, siAsisten, noAsisten, pendientes, noUtilizados) {
  const alertas = [];
  const etiquetaBoda = `Boda ${codigo} - ${nombreBoda}`;

  if (pendientes > 0) {
    alertas.push(`${etiquetaBoda}: quedan ${pendientes} pases pendientes de responder.`);
  }

  if (noUtilizados > 0) {
    alertas.push(`${etiquetaBoda}: hay ${noUtilizados} pases no utilizados por invitados que sí asistirán.`);
  }

  if ((noAsisten + noUtilizados) > 0) {
    alertas.push(`${etiquetaBoda}: pases liberados disponibles: ${noAsisten + noUtilizados}.`);
  }

  if (totalPases > 0 && siAsisten / totalPases < 0.6) {
    alertas.push(`${etiquetaBoda}: avance de confirmación menor al 60%.`);
  }

  return alertas;
}

function calcularDiasCierreSGI_(fechaCierre) {
  if (!(fechaCierre instanceof Date)) return "";
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const cierre = new Date(fechaCierre);
  cierre.setHours(0, 0, 0, 0);
  return Math.max(Math.ceil((cierre - hoy) / (1000 * 60 * 60 * 24)), 0);
}

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
      observaciones: row[7],
      linkConfirmacion: row[8],
      mensajeWhatsapp: row[9],
      enviadoWhatsapp: row[10],
      fechaEnvio: row[11],
      urlWhatsapp: row[12],
      canalEnvio: row[13],
      idApiWhatsapp: row[14],
      estadoApi: row[15]
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

  const pendientes = Math.max(totalPases - siAsisten - pasesLiberados, 0);

  const avance = totalPases > 0
    ? Math.min(Math.round((siAsisten / totalPases) * 100), 100)
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
   data.codigoBoda || "",
   data.invitado || "",
   data.asiste || "",
   data.pases || 0,
   data.pasesNoUtilizados || 0,
   data.acompanantes || "",
   data.ip || "",
   new Date(),
   data.idInvitado || ""
  ]);

  return {
    error: false,
    mensaje: "Confirmación registrada correctamente."
  };
}

function generarLinksYMensajes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaInvitados = ss.getSheetByName("Invitados");
  const hojaBodas = ss.getSheetByName("Bodas");
  const hojaPlantillas = ss.getSheetByName("Plantillas");
  const hojaConfiguracion = ss.getSheetByName("Configuración");

  if (!hojaInvitados || !hojaBodas || !hojaPlantillas || !hojaConfiguracion) {
    throw new Error("Faltan hojas requeridas: Invitados, Bodas, Plantillas o Configuración.");
  }

  const config = obtenerConfiguracionSGI_(hojaConfiguracion);
  const urlConfirmacionBase = config.UrlConfirmacion;
  const paisWhatsapp = String(config.PaisWhatsapp || "51").replace(/\D/g, "");
  const modoEnvio = String(config.ModoEnvio || "MANUAL").toUpperCase();

  if (!urlConfirmacionBase) {
    throw new Error("Falta UrlConfirmacion en la hoja Configuración.");
  }

  const bodas = obtenerBodasSGI_(hojaBodas);
  const plantillas = obtenerPlantillasSGI_(hojaPlantillas);

  const data = hojaInvitados.getDataRange().getValues();
  const headers = data[0];
  const idx = obtenerIndicesSGI_(headers);

  let actualizados = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const codigoBoda = String(row[idx.codigo_boda] || "").trim();
    const idInvitado = String(row[idx.id_invitado] || "").trim();
    const invitado = String(row[idx.invitado_principal] || "").trim();

    if (!codigoBoda || !idInvitado || !invitado) continue;

    const boda = bodas[codigoBoda];
    const plantilla = plantillas[codigoBoda];

    if (!boda) {
      hojaInvitados.getRange(i + 1, idx.estado_api + 1).setValue("ERROR: boda no encontrada");
      continue;
    }

    if (!plantilla) {
      hojaInvitados.getRange(i + 1, idx.estado_api + 1).setValue("ERROR: plantilla activa no encontrada");
      continue;
    }

    const linkConfirmacion = `${urlConfirmacionBase}?boda=${encodeURIComponent(codigoBoda)}&id=${encodeURIComponent(idInvitado)}`;

    const mensaje = construirMensajeDesdePlantillaSGI_(plantilla.mensaje, {
     INVITADO: invitado,
     NOVIOS: boda.novios,
     PASES: row[idx.pases],
     CIUDAD: boda.ciudad,

     FECHA_BODA: formatearFechaLargaSGI_(boda.fechaBoda),
     FECHA_CIERRE: formatearFechaLargaSGI_(boda.cierreConfirmacion),

     EMOJI_SALUDO: String.fromCodePoint(0x1F90D),      // 🤍
     EMOJI_FECHA: String.fromCodePoint(0x1F4C5),       // 📅
     EMOJI_UBICACION: String.fromCodePoint(0x1F4CD),   // 📍
     EMOJI_CEREMONIA: String.fromCodePoint(0x26EA),   // ⛪
     EMOJI_FIESTA: String.fromCodePoint(0x1F38A),      // 🎊

     LINK_CONFIRMACION: linkConfirmacion
  });

    const celular = construirCelularWhatsappSGI_(row[idx.celular], paisWhatsapp);
    const textoWhatsapp = encodeURIComponent(mensaje);

    const urlWhatsapp = celular
     ? `https://web.whatsapp.com/send?phone=${celular}&text=${textoWhatsapp}`
     : "";

    hojaInvitados.getRange(i + 1, idx.link_confirmacion + 1).setValue(linkConfirmacion);
    hojaInvitados.getRange(i + 1, idx.mensaje_whatsapp + 1).setValue(mensaje);
    hojaInvitados.getRange(i + 1, idx.url_whatsapp + 1).setValue(urlWhatsapp);
    hojaInvitados.getRange(i + 1, idx.canal_envio + 1).setValue(modoEnvio);
    hojaInvitados.getRange(i + 1, idx.estado_api + 1).setValue("PENDIENTE");

    actualizados++;
  }

  SpreadsheetApp.getUi().alert(`Proceso terminado. Invitados actualizados: ${actualizados}`);
}

function listarInvitadosBoda(codigoBoda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInvitados = ss.getSheetByName("Invitados");

  if (!hojaInvitados) {
    return { error: true, mensaje: "No existe la hoja Invitados." };
  }

  const invitados = leerInvitados_(hojaInvitados)
    .filter(i => String(i.codigoBoda).trim() === String(codigoBoda).trim())
    .map(i => ({
      idInvitado: i.idInvitado,
      invitadoPrincipal: i.invitadoPrincipal,
      pases: i.pases,
      celular: i.celular,
      categoria: i.categoria,
      quienInvita: i.quienInvita,
      observaciones: i.observaciones
    }));

  return {
    error: false,
    codigoBoda: codigoBoda,
    total: invitados.length,
    invitados: invitados
  };
}

function listarEnviosManual(codigoBoda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInvitados = ss.getSheetByName("Invitados");

  if (!hojaInvitados) {
    return { error: true, mensaje: "No existe la hoja Invitados." };
  }

  const invitados = leerInvitados_(hojaInvitados)
    .filter(i => String(i.codigoBoda).trim() === String(codigoBoda).trim());

  const lista = invitados.map(i => {
    const enviado = normalizarTexto_(i.enviadoWhatsapp) === "si";

    return {
      idInvitado: i.idInvitado,
      invitadoPrincipal: i.invitadoPrincipal,
      celular: i.celular,
      pases: i.pases,
      enviado: enviado,
      fechaEnvio: formatearFechaCompleta_(i.fechaEnvio),
      canalEnvio: i.canalEnvio || "MANUAL",
      estadoApi: i.estadoApi || "PENDIENTE",
      mensajeWhatsapp: i.mensajeWhatsapp || "",
      urlWhatsapp: i.urlWhatsapp || ""
    };
  });

  const pendientes = lista.filter(i => !i.enviado);
  const enviados = lista.filter(i => i.enviado);

  return {
    error: false,
    codigoBoda: codigoBoda,
    total: lista.length,
    enviados: enviados.length,
    pendientes: pendientes.length,
    siguiente: pendientes.length ? pendientes[0] : null,
    invitados: lista
  };
}

function marcarEnviadoManual(codigoBoda, idInvitado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInvitados = ss.getSheetByName("Invitados");

  if (!hojaInvitados) {
    return { error: true, mensaje: "No existe la hoja Invitados." };
  }

  const data = hojaInvitados.getDataRange().getValues();
  const headers = data[0];
  const idx = obtenerIndicesSGI_(headers);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (
      String(row[idx.codigo_boda] || "").trim() === String(codigoBoda).trim() &&
      String(row[idx.id_invitado] || "").trim() === String(idInvitado).trim()
    ) {
      hojaInvitados.getRange(i + 1, idx.enviado_whatsapp + 1).setValue("SI");
      hojaInvitados.getRange(i + 1, idx.fecha_envio + 1).setValue(new Date());
      hojaInvitados.getRange(i + 1, idx.canal_envio + 1).setValue("MANUAL");
      hojaInvitados.getRange(i + 1, idx.estado_api + 1).setValue("ENVIADO_MANUAL");

      return {
        error: false,
        mensaje: "Invitado marcado como enviado.",
        codigoBoda: codigoBoda,
        idInvitado: idInvitado
      };
    }
  }

  return { error: true, mensaje: "No se encontró el invitado." };
}

function reporteInvitados(codigoBoda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaInvitados = ss.getSheetByName("Invitados");
  const hojaConfirmaciones = ss.getSheetByName("Confirmaciones");

  if (!hojaInvitados || !hojaConfirmaciones) {
    return { error: true, mensaje: "Faltan hojas Invitados o Confirmaciones." };
  }

  const invitados = leerInvitados_(hojaInvitados)
    .filter(i => String(i.codigoBoda).trim() === String(codigoBoda).trim());

  const confirmaciones = leerConfirmaciones_(hojaConfirmaciones)
    .filter(c => String(c.codigoBoda).trim() === String(codigoBoda).trim());

  const confirmacionesPorInvitado = {};

  confirmaciones.forEach(c => {
    const key = c.idInvitado
      ? String(c.idInvitado).trim()
      : normalizarTexto_(c.invitado);

    confirmacionesPorInvitado[key] = c;
  });

  const reporte = invitados.map(i => {
    const keyInvitado = i.idInvitado
      ? String(i.idInvitado).trim()
      : normalizarTexto_(i.invitadoPrincipal);

    const conf = confirmacionesPorInvitado[keyInvitado];

    if (!conf) {
      return {
        categoria: i.categoria || "",
        invitadoPrincipal: i.invitadoPrincipal,
        pasesAsignados: i.pases,
        estado: "PENDIENTE",
        pasesConfirmados: 0,
        pasesNoUtilizados: 0,
        acompañantes: "",
        fechaConfirmacion: ""
      };
    }

    return {
      categoria: i.categoria || "",
      invitadoPrincipal: i.invitadoPrincipal,
      pasesAsignados: i.pases,
      estado: conf.asiste === "SI" ? "SI ASISTE" : "NO ASISTE",
      pasesConfirmados: conf.pases,
      pasesNoUtilizados: conf.pasesNoUtilizados || 0,
      acompañantes: conf.acompanantes,
      fechaConfirmacion: formatearFechaCompleta_(conf.fecha)
    };
  });

  reporte.sort((a, b) => {
    const catA = String(a.categoria || "").toUpperCase();
    const catB = String(b.categoria || "").toUpperCase();

    if (catA !== catB) return catA.localeCompare(catB);

    return String(a.invitadoPrincipal || "")
      .localeCompare(String(b.invitadoPrincipal || ""));
  });

  return {
    error: false,
    codigoBoda: codigoBoda,
    total: reporte.length,
    reporte: reporte
  };
}

function descargarListaInvitadosXlsx(codigoBoda) {
  const data = listarInvitadosBoda(codigoBoda);
  if (data.error) return data;

  return generarXlsxBase64_(
    `Lista_Invitados_${codigoBoda}.xlsx`,
    "Lista Invitados",
    ["Invitado principal", "Celular", "Pases", "Categoría", "Quién invita", "Observaciones"],
    data.invitados.map(i => [
      i.invitadoPrincipal || "",
      i.celular || "",
      i.pases || 0,
      i.categoria || "",
      i.quienInvita || "",
      i.observaciones || ""
    ])
  );
}

function descargarReporteInvitadosXlsx(codigoBoda) {
  const data = reporteInvitados(codigoBoda);
  if (data.error) return data;

  return generarXlsxBase64_(
    `Reporte_Invitados_${codigoBoda}.xlsx`,
    "Reporte Invitados",
    ["Invitado principal", "Pases asignados", "Estado", "Pases confirmados", "No utilizados", "Acompañantes", "Fecha confirmación"],
    data.reporte.map(i => [
      i.invitadoPrincipal || "",
      i.pasesAsignados || 0,
      i.estado || "",
      i.pasesConfirmados || 0,
      i.pasesLiberados || 0,
      i.acompanantes || "",
      i.fechaConfirmacion || ""
    ])
  );
}

function generarXlsxBase64_(nombreArchivo, nombreHoja, encabezados, filas) {
  const archivoTemp = SpreadsheetApp.create(nombreArchivo.replace(".xlsx", ""));
  const ssTemp = SpreadsheetApp.openById(archivoTemp.getId());
  const hoja = ssTemp.getSheets()[0];

  hoja.setName(nombreHoja);
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);

  if (filas.length > 0) {
    hoja.getRange(2, 1, filas.length, encabezados.length).setValues(filas);
  }

  hoja.getRange(1, 1, 1, encabezados.length).setFontWeight("bold");
  hoja.autoResizeColumns(1, encabezados.length);

  SpreadsheetApp.flush();

  const url = `https://docs.google.com/spreadsheets/d/${archivoTemp.getId()}/export?format=xlsx`;
  const token = ScriptApp.getOAuthToken();

  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  }).getBlob();

  DriveApp.getFileById(archivoTemp.getId()).setTrashed(true);

  return {
    error: false,
    nombreArchivo: nombreArchivo,
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    base64: Utilities.base64Encode(blob.getBytes())
  };
}

function actualizarCentroEnvios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPanel = ss.getSheetByName("Panel_Envios");
  const hojaInvitados = ss.getSheetByName("Invitados");

  if (!hojaPanel || !hojaInvitados) {
    throw new Error("Faltan hojas requeridas: Panel_Envios o Invitados.");
  }

  const codigoBoda = String(hojaPanel.getRange("B3").getValue() || "").trim();
  const modo = String(hojaPanel.getRange("E3").getValue() || "MANUAL").trim().toUpperCase();

  if (!codigoBoda) {
    SpreadsheetApp.getUi().alert("Debes ingresar el código de boda en B3.");
    return;
  }

  const data = hojaInvitados.getDataRange().getValues();
  const headers = data[0];
  const idx = obtenerIndicesSGI_(headers);

  let pendientes = 0;
  let invitadoActual = null;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const rowCodigoBoda = String(row[idx.codigo_boda] || "").trim();
    const enviado = normalizarTextoSGI_(row[idx.enviado_whatsapp]);

    if (rowCodigoBoda !== codigoBoda) continue;

    if (enviado !== "si") {
      pendientes++;

      if (!invitadoActual) {
        invitadoActual = {
          invitado: row[idx.invitado_principal],
          celular: row[idx.celular],
          pases: row[idx.pases],
          idInvitado: row[idx.id_invitado],
          codigoBoda: row[idx.codigo_boda]
        };
      }
    }
  }

  hojaPanel.getRange("H3").setValue(pendientes);
  hojaPanel.getRange("A7:G7").clearContent();

  if (!invitadoActual) {
    hojaPanel.getRange("A7").setValue("No hay invitados pendientes para esta boda.");
    hojaPanel.getRange("D7").setValue("COMPLETO");
    SpreadsheetApp.getUi().alert("No hay invitados pendientes para esta boda.");
    return;
  }

  hojaPanel.getRange("A7").setValue(invitadoActual.invitado);
  hojaPanel.getRange("B7").setValue(invitadoActual.celular);
  hojaPanel.getRange("C7").setValue(invitadoActual.pases);
  hojaPanel.getRange("D7").setValue("PENDIENTE");
  hojaPanel.getRange("E7").setValue(modo === "API" ? "ENVIAR API" : "ABRIR WHATSAPP");
  hojaPanel.getRange("F7").setValue(invitadoActual.idInvitado);
  hojaPanel.getRange("G7").setValue(invitadoActual.codigoBoda);

  hojaPanel.hideColumns(6, 2);
}

function abrirWhatsappInvitadoActual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPanel = ss.getSheetByName("Panel_Envios");

  const codigoBoda = String(hojaPanel.getRange("G7").getValue() || "").trim();
  const idInvitado = String(hojaPanel.getRange("F7").getValue() || "").trim();
  const modo = String(hojaPanel.getRange("E3").getValue() || "MANUAL").trim().toUpperCase();

  if (!codigoBoda || !idInvitado) {
    SpreadsheetApp.getUi().alert("No hay invitado actual cargado. Ejecuta primero: Actualizar Centro de Envíos.");
    return;
  }

  if (modo === "API") {
    SpreadsheetApp.getUi().alert("Modo API aún no está activo. Por ahora usa modo MANUAL.");
    return;
  }

  const envio = listarEnviosManual(codigoBoda).invitados.find(i => i.idInvitado === idInvitado);

  if (!envio || !envio.urlWhatsapp) {
    SpreadsheetApp.getUi().alert("El invitado no tiene URL WhatsApp generada. Ejecuta primero: Generar enlaces y mensajes.");
    return;
  }

  const html = HtmlService.createHtmlOutput(`
    <script>
      window.open("${envio.urlWhatsapp}", "_blank");
      google.script.host.close();
    </script>
  `);

  SpreadsheetApp.getUi().showModalDialog(html, "Abriendo WhatsApp...");
}

function marcarInvitadoActualEnviado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPanel = ss.getSheetByName("Panel_Envios");

  const codigoBoda = String(hojaPanel.getRange("G7").getValue() || "").trim();
  const idInvitado = String(hojaPanel.getRange("F7").getValue() || "").trim();

  const resultado = marcarEnviadoManual(codigoBoda, idInvitado);

  if (resultado.error) {
    SpreadsheetApp.getUi().alert(resultado.mensaje);
    return;
  }

  actualizarCentroEnvios();
  SpreadsheetApp.getUi().alert("Invitado marcado como enviado.");
}

function obtenerConfiguracionSGI_(hoja) {
  const data = hoja.getDataRange().getValues();
  const config = {};

  for (let i = 1; i < data.length; i++) {
    const clave = String(data[i][0] || "").trim();
    const valor = data[i][1];
    if (clave) config[clave] = valor;
  }

  return config;
}

function obtenerBodasSGI_(hoja) {
  const data = hoja.getDataRange().getValues();
  const headers = data[0];
  const idx = obtenerIndicesSGI_(headers);

  const bodas = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const codigo = String(row[idx.codigo_boda] || "").trim();
    if (!codigo) continue;

    bodas[codigo] = {
     codigo: codigo,
     novios: row[idx.novios],
     fechaBoda: row[idx.fecha_boda],
     ciudad: row[idx.ciudad],
     cierreConfirmacion: row[idx.cierre_confirmacion],
     fotoUrl: idx.foto_url !== undefined ? row[idx.foto_url] : "",
     htmlUrl: idx.html_url !== undefined ? row[idx.html_url] : ""
    };
  }

  return bodas;
}

function obtenerPlantillasSGI_(hoja) {
  const data = hoja.getDataRange().getValues();
  const headers = data[0];
  const idx = obtenerIndicesSGI_(headers);

  const plantillas = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const codigo = String(row[idx.codigo_boda] || "").trim();
    const activo = normalizarTextoSGI_(row[idx.activo]);

    if (!codigo || activo !== "si") continue;

    plantillas[codigo] = {
      codigo: codigo,
      mensaje: row[idx.mensaje_whatsapp]
    };
  }

  return plantillas;
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

function formatearFechaLargaSGI_(fecha) {
  if (!(fecha instanceof Date)) return String(fecha || "");

  const dias = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  const meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

  const diaSemana = dias[fecha.getDay()];
  const dia = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd");
  const mes = meses[fecha.getMonth()];
  const anio = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy");

  return `${diaSemana}, ${dia} de ${mes} del ${anio}`;
}

function normalizarTexto_(valor) {
  return String(valor || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function normalizarTextoSGI_(valor) {
  return normalizarTexto_(valor);
}

function normalizarHeaderSGI_(valor) {
  return String(valor || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "_");
}

function obtenerIndicesSGI_(headers) {
  const idx = {};

  headers.forEach((h, i) => {
    const key = normalizarHeaderSGI_(h);
    idx[key] = i;
  });

  return idx;
}

function construirMensajeDesdePlantillaSGI_(plantilla, valores) {
  let mensaje = String(plantilla || "");

  Object.keys(valores).forEach(clave => {
    const regex = new RegExp(`{{${clave}}}`, "g");
    mensaje = mensaje.replace(regex, valores[clave] || "");
  });

  return mensaje;
}

function construirCelularWhatsappSGI_(celular, pais) {
  let numero = String(celular || "").replace(/\D/g, "");

  if (!numero) return "";

  // Si ya viene con código internacional, no agregar código de país
  if (numero.length > 9) {
    return numero;
  }

  // Si es celular peruano de 9 dígitos, agregar 51
  if (numero.length === 9) {
    return "51" + numero;
  }

  // Caso no estándar: devolver limpio para revisión manual
  return numero;
}

function listarMesas_(codigoBoda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Mesas");
  if (!hoja) return { error: true, mensaje: "No existe la hoja Mesas." };

  const data = hoja.getDataRange().getValues();
  const headers = data.shift();
  const idx = obtenerIndicesSGI_(headers);

  const mesas = data
    .filter(r => String(r[idx.codigo_boda]).trim() === String(codigoBoda).trim())
    .map(r => ({
      codigoBoda: r[idx.codigo_boda],
      mesa: r[idx.mesa],
      capacidad: Number(r[idx.capacidad]) || 10,
      ubicacionX: Number(r[idx.ubicacion_x]) || 0,
      ubicacionY: Number(r[idx.ubicacion_y]) || 0,
      tipoMesa: r[idx.tipo_mesa] || ""
    }));

  return { error: false, mesas };
}

function listarInvitadosMesas_(codigoBoda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaConfirmaciones = ss.getSheetByName("Confirmaciones");
  const hojaAsignacion = ss.getSheetByName("Asignacion_Mesas");

  if (!hojaConfirmaciones) {
    return { error: true, mensaje: "No existe la hoja Confirmaciones." };
  }

  const dataConf = hojaConfirmaciones.getDataRange().getValues();
  const headersConf = dataConf[0];
  const idxConf = obtenerIndicesSGI_(headersConf);

  const asignaciones = {};
  if (hojaAsignacion && hojaAsignacion.getLastRow() > 1) {
    const dataAsig = hojaAsignacion.getDataRange().getValues();
    const headersAsig = dataAsig[0];
    const idxAsig = obtenerIndicesSGI_(headersAsig);

    for (let i = 1; i < dataAsig.length; i++) {
      const row = dataAsig[i];
      if (String(row[idxAsig.codigo_boda]).trim() !== String(codigoBoda).trim()) continue;
      asignaciones[String(row[idxAsig.id_invitado]).trim()] = String(row[idxAsig.mesa] || "").trim();
    }
  }

  const invitados = [];

  for (let i = 1; i < dataConf.length; i++) {
    const row = dataConf[i];
    if (String(row[idxConf.codigo_boda]).trim() !== String(codigoBoda).trim()) continue;
    if (normalizarTextoSGI_(row[idxConf.asiste]) !== "si") continue;

    const idBase = String(row[idxConf.invitado] || "").trim();
    const nombreBase = String(row[idxConf.invitado] || "").trim();
    const acompanantes = String(row[idxConf.acompanantes] || "").trim();

    invitados.push({
      idInvitado: `${codigoBoda}_${idBase}_principal`,
      nombreInvitado: nombreBase,
      tipo: acompanantes ? "pareja" : "principal",
      grupoOrigen: nombreBase,
      mesa: asignaciones[`${codigoBoda}_${idBase}_principal`] || ""
    });

    if (acompanantes) {
      acompanantes.split(",").map(a => a.trim()).filter(Boolean).forEach((nombre, index) => {
        invitados.push({
          idInvitado: `${codigoBoda}_${idBase}_acompanante_${index + 1}`,
          nombreInvitado: nombre,
          tipo: "acompanante",
          grupoOrigen: nombreBase,
          mesa: asignaciones[`${codigoBoda}_${idBase}_acompanante_${index + 1}`] || ""
        });
      });
    }
  }

  const total = invitados.length;
  const asignados = invitados.filter(i => i.mesa).length;

  return {
    error: false,
    total,
    asignados,
    pendientes: total - asignados,
    avance: total ? Math.round((asignados / total) * 100) : 0,
    invitados
  };
}

function guardarAsignacionMesas_(codigoBoda, payloadEntrada) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Asignacion_Mesas");

  if (!hoja) {
    return {
      error: true,
      mensaje: "No existe la hoja Asignacion_Mesas."
    };
  }

  let payload = payloadEntrada;

  if (typeof payloadEntrada === "string") {
    payload = JSON.parse(payloadEntrada || "[]");
  }

  if (!Array.isArray(payload)) {
    return {
      error: true,
      mensaje: "Payload inválido. No es una lista."
    };
  }

  const data = hoja.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(codigoBoda).trim()) {
      hoja.deleteRow(i + 1);
    }
  }

  const filas = payload
    .map(item => [
      codigoBoda,
      item.mesa || "",
      item.idInvitado || item.id_invitado || "",
      item.nombreInvitado || item.nombre_invitado || item.nombre || "",
      item.tipo || "",
      item.grupoOrigen || item.grupo_origen || "",
      new Date()
    ])
    .filter(fila => fila[1] && fila[2] && fila[3]);

  if (filas.length > 0) {
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 7).setValues(filas);
  }

  return {
    error: false,
    mensaje: "Distribución de mesas guardada correctamente.",
    registros: filas.length
  };
}

function reiniciarAsignacionMesas_(codigoBoda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Asignacion_Mesas");
  if (!hoja) return { error: true, mensaje: "No existe la hoja Asignacion_Mesas." };

  const data = hoja.getDataRange().getValues();
  const headers = data[0];
  const idx = obtenerIndicesSGI_(headers);

  let eliminados = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idx.codigo_boda]).trim() === String(codigoBoda).trim()) {
      hoja.deleteRow(i + 1);
      eliminados++;
    }
  }

  return {
    error: false,
    mensaje: "Asignación de mesas reiniciada.",
    eliminados
  };
}

function reporteMesas_(codigoBoda) {
  const mesasResp = listarMesas_(codigoBoda);
  const invitadosResp = listarInvitadosMesas_(codigoBoda);

  if (mesasResp.error) return mesasResp;
  if (invitadosResp.error) return invitadosResp;

  if (invitadosResp.pendientes > 0) {
    return {
      error: true,
      mensaje: `No se puede generar reporte. Faltan ${invitadosResp.pendientes} invitados por asignar.`
    };
  }

  const reporte = mesasResp.mesas.map(m => ({
    mesa: m.mesa,
    capacidad: m.capacidad,
    tipoMesa: m.tipoMesa,
    invitados: invitadosResp.invitados
      .filter(i => String(i.mesa).trim() === String(m.mesa).trim())
      .map(i => ({
        nombreInvitado: i.nombreInvitado,
        tipo: i.tipo,
        grupoOrigen: i.grupoOrigen
      }))
  }));

  return {
    error: false,
    codigoBoda,
    totalInvitados: invitadosResp.total,
    reporte
  };
}



function probarDashboard() {
  Logger.log(JSON.stringify(obtenerDashboard(), null, 2));
}

function probarBuscarInvitado() {
  Logger.log(JSON.stringify(buscarInvitadoConfirmacion("B001", "B001-001"), null, 2));
}

function probarListarInvitadosBoda() {
  Logger.log(JSON.stringify(listarInvitadosBoda("B001"), null, 2));
}

function probarListarEnviosManual() {
  Logger.log(JSON.stringify(listarEnviosManual("B001"), null, 2));
}

function probarReporteInvitados() {
  Logger.log(JSON.stringify(reporteInvitados("B001"), null, 2));
}