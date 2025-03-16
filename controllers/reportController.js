const reportService = require('../services/reportService');

/**
 * Genera un reporte basado en una solicitud y formulario
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
*/

const generateReport = async (req, res) => {
  try {
    const { solicitudId, formNumber } = req.body;
    console.log("Datos recibidos en generateReport:");
    console.log("solicitudId:", solicitudId);
    console.log("formNumber:", formNumber);
  
    if (!solicitudId || !formNumber) {
      console.error('Error: Los parámetros solicitudId y formNumber son requeridos');
      return res.status(400).json({ error: 'Los parámetros solicitudId y formNumber son requeridos' });
    }

    // Usar el servicio de reportes para generar el informe
    const result = await reportService.generateReport(solicitudId, formNumber);
    res.status(200).json(result);
  } catch (error) {
    console.error('Error al generar el informe:', error);
    res.status(500).json({ error: 'Error al generar el informe' });
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