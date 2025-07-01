const reportService = require('../services/reportService');

/**
 * Genera un reporte basado en una solicitud y formulario
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
*/

const generateReport = async (req, res) => {
  try {
    const { solicitudId, formNumber } = req.body;
    console.log("[INICIO] generateReport: solicitudId=", solicitudId, "formNumber=", formNumber);
  
    if (!solicitudId || !formNumber) {
      console.error('[ERROR] Parámetros faltantes: solicitudId o formNumber');
      return res.status(400).json({ error: 'Los parámetros solicitudId y formNumber son requeridos' });
    }

    // Iniciar tiempo de ejecución para medir rendimiento
    const startTime = Date.now();

    try {
      console.log("[PASO] Llamando a reportService.generateReport...");
      const result = await reportService.generateReport(solicitudId, formNumber);
      const endTime = Date.now();
      console.log(`[OK] Reporte generado en ${(endTime - startTime)/1000} segundos`);
      res.status(200).json(result);
    } catch (error) {
      const failTime = Date.now();
      console.error('[ERROR] Error al generar el reporte:', error.message);
      console.error('[ERROR] Stack:', error.stack);
      console.error(`[ERROR] Tiempo transcurrido: ${(failTime - startTime)/1000} segundos`);
      res.status(500).json({ 
        error: 'Error al generar el reporte', 
        details: error.message,
        stack: error.stack
      });
    }
  } catch (error) {
    console.error('[ERROR] Error en el controlador generateReport:', error.message);
    console.error('[ERROR] Stack:', error.stack);
    res.status(500).json({ 
      error: 'Error en el controlador', 
      details: error.message 
    });
  }
};

/**
 * Genera un reporte para descarga o edición
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
*/

const downloadReport = async (req, res) => {
  try {
    const { solicitudId, formNumber, mode } = req.body;
    
    if (!solicitudId || !formNumber) {
      return res.status(400).json({ error: 'Los parámetros solicitudId y formNumber son requeridos' });
    }

    // Usar el servicio de reportes para generar el informe en el modo especificado
    const result = await reportService.downloadReport(solicitudId, formNumber, mode);
    res.status(200).json(result);
  } catch (error) {
    console.error('Error al generar el informe para descarga/edición:', error);
    res.status(500).json({ error: 'Error al generar el informe para descarga/edición' });
  }
};

module.exports = {
  generateReport,
  downloadReport
};