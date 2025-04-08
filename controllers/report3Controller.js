const reportService = require('../services/reportService');

/**
 * Genera el reporte para el formulario 3
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const generateReport3 = async (req, res) => {
  try {
    const { solicitudId } = req.body;
    
    if (!solicitudId) {
      return res.status(400).json({ 
        error: 'El parámetro solicitudId es requerido' 
      });
    }
    
    // El tipo de formulario es 3 (fijo para este controlador)
    const formNumber = 3;
    
    // Registrar tiempo de inicio para medición de rendimiento
    const startTime = Date.now();
    
    // Generar el reporte usando el servicio genérico
    const result = await reportService.generateReport(solicitudId, formNumber);
    
    // Calcular tiempo de generación
    const endTime = Date.now();
    console.log(`✅ Reporte 3 generado en ${(endTime - startTime)/1000} segundos`);
    
    res.status(200).json(result);
  } catch (error) {
    console.error('Error al generar el reporte 3:', error);
    res.status(500).json({ 
      error: 'Error al generar el reporte 3',
      details: error.message 
    });
  }
};

module.exports = {
  generateReport3
};