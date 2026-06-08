/**
 * Dashboard Ejecutivo SGI v3
 * Obtiene indicadores generales para el panel UX.
 */

function obtenerDashboard() {

  const resultado = {
    fechaActualizacion: new Date(),
    totalBodas: 0,
    bodasActivas: [],
    alertas: [],
    resumen: {
      invitados: 0,
      confirmados: 0,
      pendientes: 0,
      noAsistiran: 0,
      pasesLiberados: 0
    }
  };

  return resultado;

}