const sheetsService = require('./sheetsService');
const driveService = require('./driveService');

class ReportGenerationService {
  /**
   * Genera un reporte basado en la configuración del tipo de reporte
   * @param {String} solicitudId - ID de la solicitud
   * @param {Number} formNumber - Número de formulario (1-4)
   * @returns {Promise<Object>} Resultado con link al reporte generado
   */
  async generateReport(solicitudId, formNumber) {
    try {
      // Validar parámetros
      if (!solicitudId || !formNumber) {
        throw new Error('Los parámetros solicitudId y formNumber son requeridos');
      }

      // Cargar la configuración específica del reporte
      const reportConfig = this.loadReportConfig(formNumber);
      
      if (!reportConfig) {
        throw new Error(`No se encontró configuración para el formulario ${formNumber}`);
      }

      // Obtener datos de la solicitud usando la configuración de hojas del reporte
      const solicitudData = await this.getSolicitudData(solicitudId, reportConfig.sheetDefinitions);
      
      // AÑADIR ESTE CÓDIGO PARA DIAGNÓSTICO
      console.log("🔍 DATOS RECIBIDOS PARA TRANSFORMAR:");
      console.log("- ID Solicitud:", solicitudId);
      console.log("- Datos solicitud:", JSON.stringify(solicitudData).substring(0, 300) + "...");
      
      // Verificar específicamente campos críticos
      const camposCriticos = [
        'nombre_actividad', 'ingresos_cantidad', 'ingresos_vr_unit', 
        'subtotal_gastos', 'imprevistos_3', 'gastos'
      ];
      
      camposCriticos.forEach(campo => {
        console.log(`- ${campo}: ${solicitudData[campo] || 'NO ENCONTRADO'}`);
      });

      // Procesar datos adicionales si el reporte lo requiere (como gastos)
      const additionalData = await this.processAdditionalData(solicitudId, reportConfig);

      // Combinar datos de la solicitud con datos adicionales
      const combinedData = { ...solicitudData, ...additionalData };

      // Transformar datos utilizando la función específica de la configuración del reporte
      const transformedData = reportConfig.transformData(combinedData);

      // Generar el reporte usando el servicio de Drive
      const reportLink = await driveService.generateReport(
        formNumber,
        solicitudId,
        transformedData
      );

      return { link: reportLink };
    } catch (error) {
      console.error('Error al generar el informe:', error);
      throw error;
    }
  }

  /**
   * Carga la configuración específica de un reporte
   * @param {Number} formNumber - Número de formulario
   * @returns {Object} Configuración del reporte
   */
  loadReportConfig(formNumber) {
    try {
      // Cargar dinámicamente la configuración del reporte
      return require(`../reportConfigs/report${formNumber}Config.js`);
    } catch (error) {
      console.error(`Error al cargar configuración del reporte ${formNumber}:`, error);
      return null;
    }
  }

  /**
   * Obtiene datos de una solicitud
   * @param {String} solicitudId - ID de la solicitud
   * @param {Object} sheetDefinitions - Definiciones de hojas para obtener datos
   * @returns {Promise<Object>} Datos de la solicitud
   */
  async getSolicitudData(solicitudId, sheetDefinitions) {
    try {
      return await sheetsService.getSolicitudData(solicitudId, sheetDefinitions);
    } catch (error) {
      console.error('Error al obtener datos de la solicitud:', error);
      throw new Error('Error al obtener datos de la solicitud');
    }
  }

  /**
   * Procesa datos adicionales requeridos por el reporte
   * @param {String} solicitudId - ID de la solicitud
   * @param {Object} reportConfig - Configuración del reporte
   * @returns {Promise<Object>} Datos adicionales procesados
   */
  async processAdditionalData(solicitudId, reportConfig) {
    // Si el reporte no requiere datos adicionales, retornar objeto vacío
    if (!reportConfig.requiresAdditionalData) {
      return {};
    }

    try {
      // Procesar gastos si el reporte lo requiere
      if (reportConfig.requiresGastos) {
        return await this.processGastosData(solicitudId);
      }
      
      // Otros tipos de datos adicionales pueden ser procesados aquí

      return {};
    } catch (error) {
      console.error('Error al procesar datos adicionales:', error);
      return {};
    }
  }

  /**
   * Procesa datos de gastos para una solicitud
   * @param {String} solicitudId - ID de la solicitud
   * @returns {Promise<Object>} Datos de gastos procesados
   */
  async processGastosData(solicitudId) {
    try {
      const client = sheetsService.getClient();
      
      // Obtener gastos y conceptos
      const gastosResponse = await client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'GASTOS!A2:F500'
      });
      
      // Procesar los datos - SOLUCIÓN MEJORADA
      const gastosRows = gastosResponse.data.values || [];
      
      // Filtrar gastos de la solicitud actual y convertir a formato de objeto
      const solicitudGastos = gastosRows
        .filter(row => row[1] === solicitudId.toString())
        .map(row => ({
          id_conceptos: row[0]?.toString() || '',
          id_solicitud: row[1]?.toString() || '',
          cantidad: parseFloat(row[2]) || 0,
          valor_unit: parseFloat(row[3]) || 0,
          valor_total: parseFloat(row[4]) || 0,
          concepto_padre: row[5]?.toString() || ''
        }));
      
      console.log(`🔍 GASTOS procesados para solicitud ${solicitudId}: ${solicitudGastos.length}`);
      if (solicitudGastos.length > 0) {
        console.log(`🔍 Primer GASTO: ${JSON.stringify(solicitudGastos[0])}`);
      }
      
      return { GASTOS: solicitudGastos };
    } catch (error) {
      console.error('Error al procesar gastos:', error);
      return { GASTOS: [] };
    }
  }
}

module.exports = new ReportGenerationService();