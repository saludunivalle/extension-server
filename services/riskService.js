/**
 * Service for handling risk data processing in reports
 */
const sheetsService = require('./sheetsService');
const { generateRows, insertDynamicRows } = require('./dynamicRows/risksGenerator');

class RiskService {
  /**
   * Process risk data for a specific solicitud
   * @param {String} solicitudId - ID of the solicitud
   * @returns {Promise<Object>} Structured risk data
   */
  async processRiskData(solicitudId) {
    try {
      console.log(`📊 Procesando datos de riesgos para solicitud ${solicitudId}`);
      
      if (!solicitudId) {
        throw new Error('ID de solicitud es requerido para procesar riesgos');
      }
      
      // Obtener todos los riesgos de la solicitud
      const riesgos = await sheetsService.getRisksBySolicitud(solicitudId);
      
      if (!riesgos || riesgos.length === 0) {
        console.log(`⚠️ No se encontraron riesgos para la solicitud ${solicitudId}`);
        return { riesgos: [], riesgosPorCategoria: {} };
      }
      
      console.log(`📋 Procesando ${riesgos.length} riesgos encontrados`);
      
      // Clasificar riesgos por categoría
      const riesgosPorCategoria = {
        diseno: [],
        locacion: [],
        desarrollo: [],
        cierre: [],
        otros: []
      };
      
      // Normalizar datos y agrupar por categoría
      riesgos.forEach(riesgo => {
        // Normalizar el objeto de riesgo
        const riesgoNormalizado = {
          id_riesgo: riesgo.id_riesgo,
          nombre_riesgo: riesgo.nombre_riesgo,
          aplica: riesgo.aplica || 'No',
          mitigacion: riesgo.mitigacion || '',
          id_solicitud: riesgo.id_solicitud,
          categoria: riesgo.categoria?.toLowerCase() || 'otros',
          
          // Campos adicionales para el mapeo de plantilla
          id: riesgo.id_riesgo,
          descripcion: riesgo.nombre_riesgo,
          impacto: riesgo.aplica === 'Sí' || riesgo.aplica === 'Si' ? 'Alto' : 'Bajo',
          probabilidad: riesgo.aplica === 'Sí' || riesgo.aplica === 'Si' ? 'Alta' : 'Baja',
          estrategia: riesgo.mitigacion || 'No especificado'
        };
        
        // Determinar la categoría
        let categoriaAsignada = 'otros';
        const cat = riesgo.categoria?.toLowerCase() || '';
        
        if (cat.includes('dise')) {
          categoriaAsignada = 'diseno';
        } else if (cat.includes('loca')) {
          categoriaAsignada = 'locacion';
        } else if (cat.includes('desa')) {
          categoriaAsignada = 'desarrollo';
        } else if (cat.includes('cier')) {
          categoriaAsignada = 'cierre';
        }
        
        // Añadir a la categoría correspondiente
        riesgosPorCategoria[categoriaAsignada].push(riesgoNormalizado);
      });
      
      // Generar estadísticas de riesgos por categoría
      const stats = {};
      Object.keys(riesgosPorCategoria).forEach(cat => {
        stats[cat] = riesgosPorCategoria[cat].length;
      });
      
      console.log(`📈 Estadísticas de riesgos por categoría:`, stats);
      
      return {
        riesgos: riesgos,
        riesgosPorCategoria: riesgosPorCategoria,
        stats: stats
      };
    } catch (error) {
      console.error(`❌ Error al procesar riesgos:`, error);
      return { riesgos: [], riesgosPorCategoria: {}, error: error.message };
    }
  }
  
  /**
   * Generate dynamic rows for risks in a specific category
   * @param {Array} riesgos - Risk data
   * @param {String} categoria - Category of risks
   * @param {String} insertLocation - Location to insert rows
   * @returns {Object} Dynamic rows data
   */
  generateRiskRows(riesgos, categoria, insertLocation) {
    try {
      return generateRows(riesgos, categoria, insertLocation);
    } catch (error) {
      console.error(`❌ Error al generar filas dinámicas para riesgos:`, error);
      return null;
    }
  }
  
  /**
   * Insert dynamic risk rows into a Google Sheet
   * @param {String} fileId - ID of the Google Sheet
   * @param {Object} dynamicRowsData - Processed data with rows to insert
   * @returns {Promise<Boolean>} Success status
   */
  async insertDynamicRows(fileId, dynamicRowsData) {
    try {
      return await insertDynamicRows(fileId, dynamicRowsData);
    } catch (error) {
      console.error(`❌ Error al insertar filas dinámicas de riesgos:`, error);
      return false;
    }
  }
}

module.exports = new RiskService();