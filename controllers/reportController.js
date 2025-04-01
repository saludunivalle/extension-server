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
      return res.status(400).json({ error: 'Los parámetros solicitudId y formNumber son requeridos' });
    }

    // Iniciar tiempo de ejecución para medir rendimiento
    const startTime = Date.now();

    try {
      // Usar el servicio de reportes para generar el informe
      const result = await reportService.generateReport(solicitudId, formNumber);
      
      const endTime = Date.now();
      console.log(`✅ Reporte generado en ${(endTime - startTime)/1000} segundos`);
      
      res.status(200).json(result);
    } catch (error) {
      console.error('Error al generar el reporte:', error);
      
      // Si es un error de cuota excedida, devolver un mensaje específico
      if (error.message?.includes('Quota exceeded') || 
          error.code === 429 || 
          (error.response && error.response.status === 429)) {
        return res.status(200).json({
          success: false,
          quotaExceeded: true,
          message: 'No se pudo generar el reporte debido a límites de la API de Google. Por favor, intente más tarde.',
          errorDetails: 'API Quota exceeded',
          fallbackLink: null // Aquí podrías proporcionar un enlace a una plantilla genérica
        });
      }
      
      // Para otros errores, mantener el 500
      res.status(500).json({ 
        error: 'Error al generar el reporte', 
        details: error.message
      });
    }
  } catch (error) {
    console.error('Error en el controlador generateReport:', error);
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

// Añadir este nuevo método
const previewReport = async (req, res) => {
  try {
    const { solicitudId, formNumber } = req.body;
    
    if (!solicitudId || !formNumber) {
      return res.status(400).json({ error: 'Los parámetros solicitudId y formNumber son requeridos' });
    }
    
    // Obtener los datos sin transformar
    const reportConfig = require(`../reportConfigs/report${formNumber}Config.js`);
    
    // Obtener datos de la solicitud (usando caché para evitar problemas de cuota)
    const { getDataWithCache } = require('../utils/cacheUtils');
    const solicitudData = await getDataWithCache(
      `preview_${solicitudId}`,
      async () => {
        // Simplificar para evitar múltiples llamadas a Sheets
        return {
          id_solicitud: solicitudId,
          // Datos mínimos para prueba
          nombre_actividad: "Actividad de ejemplo para previsualización",
          fecha_solicitud: new Date().toISOString().slice(0, 10),
          // Puedes añadir más datos de prueba aquí
        };
      }
    );
    
    // Transformar datos usando la configuración del reporte
    const transformedData = await reportConfig.transformData(solicitudData);
    
    // Devolver los datos para previsualización
    res.status(200).json({
      success: true,
      message: 'Datos de previsualización del reporte',
      reportData: transformedData,
      config: {
        title: reportConfig.title,
        templateId: reportConfig.templateId,
        // Otros detalles de configuración
      }
    });
  } catch (error) {
    console.error('Error al generar previsualización:', error);
    res.status(500).json({ error: 'Error al generar previsualización', details: error.message });
  }
};

module.exports = {
  generateReport,
  downloadReport,
  previewReport
};