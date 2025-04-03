const reportService = require('../services/reportService');

/**
 * Genera el reporte para el formulario 2
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const generateReport2 = async (req, res) => {
  try {
    const { solicitudId, formData: clientFormData } = req.body; // Recibir datos del frontend
    
    if (!solicitudId) {
      return res.status(400).json({ 
        error: 'El parámetro solicitudId es requerido' 
      });
    }
    
    // El tipo de formulario es 2 (fijo para este controlador)
    const formNumber = 2;
    
    // Registrar tiempo de inicio para medición de rendimiento
    const startTime = Date.now();
    
    // MODIFICACIÓN: Obtener datos directamente de Google Sheets pero
    // inyectar un objeto de reemplazo si están vacíos
    let result;
    
    try {
      // Intentar obtener datos normalmente
      const reportConfig = require('../reportConfigs/report2Config.js');
      const datosSolicitud = await reportService.getSolicitudData(solicitudId, reportConfig.sheetDefinitions);
      
      // AÑADIR ESTE LOG CRÍTICO:
      console.log("🔍 DATOS DE SOLICITUD OBTENIDOS:", JSON.stringify(datosSolicitud).substring(0, 500) + "...");
      
      // Si los datos están vacíos o parece haber un problema, usar los datos del cliente
      if (!datosSolicitud || Object.keys(datosSolicitud).length === 0 || !datosSolicitud.nombre_actividad) {
        console.log("⚠️ Usando datos enviados por el cliente como respaldo");
 
        // Generar el reporte directamente con estos datos
        result = await reportService.generateReportWithData(solicitudId, formNumber, datosHardcoded);
      } else {
        // Usar el flujo normal
        result = await reportService.generateReport(solicitudId, formNumber);
      }
    } catch (error) {
      console.error("Error obteniendo datos, usando respaldo:", error);
      result = await reportService.generateReportWithData(solicitudId, formNumber, datosEmergencia);
    }
    
    // Calcular tiempo de generación
    const endTime = Date.now();
    console.log(`✅ Reporte 2 generado en ${(endTime - startTime)/1000} segundos`);
    
    res.status(200).json(result);
  } catch (error) {
    console.error('Error al generar el reporte 2:', error);
    res.status(500).json({ 
      error: 'Error al generar el reporte 2',
      details: error.message 
    });
  }
};

module.exports = {
  generateReport2
};