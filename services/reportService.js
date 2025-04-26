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
    console.log('⚡️ INICIO DE GENERACIÓN DE REPORTE:');
    console.log(`📋 Datos recibidos: solicitudId=${solicitudId}, formNumber=${formNumber}`);
    
    try {
      // Verificar servicios
      if (!sheetsService || typeof sheetsService.getClient !== 'function') {
        throw new Error('El servicio sheetsService no está configurado correctamente');
      }
      
      if (!driveService || typeof driveService.generateReport !== 'function') {
        throw new Error('El servicio driveService no está configurado correctamente');
      }
      
      console.log('✅ Verificación de servicios: OK');
      
      // Validar parámetros
      if (!solicitudId || typeof solicitudId !== 'string') {
        throw new Error('solicitudId inválido');
      }
  
      const formNum = parseInt(formNumber, 10);
      if (isNaN(formNum)) {
        throw new Error('formNumber debe ser numérico');
      }
      
      console.log('✅ Validación de parámetros: OK');
  
      // Cargar la configuración específica del reporte
      console.log(`🔄 Cargando configuración para formulario ${formNum}...`);
      const reportConfig = this.loadReportConfig(formNum);
      
      if (!reportConfig) {
        throw new Error(`No se encontró configuración para el formulario ${formNum}`);
      }
      
      if (!reportConfig.transformData || typeof reportConfig.transformData !== 'function') {
        throw new Error(`La configuración del formulario ${formNum} no tiene un método transformData válido`);
      }
      
      console.log('✅ Configuración cargada correctamente:', {
        título: reportConfig.title || 'Sin título',
        tieneTransformData: !!reportConfig.transformData,
        requiereDatosAdicionales: reportConfig.requiresAdditionalData || false,
        requiereGastos: reportConfig.requiresGastos || false
      });
  
      // Obtener datos de la solicitud usando la configuración de hojas del reporte
      console.log(`🔄 Obteniendo datos de solicitud ${solicitudId}...`);
      const solicitudData = await this.getSolicitudData(solicitudId, reportConfig.sheetDefinitions);
      console.log(`✅ Datos de solicitud obtenidos:`, {
        tieneData: !!solicitudData,
        camposRecibidos: Object.keys(solicitudData).length,
        muestraData: {
          nombre_actividad: solicitudData.nombre_actividad,
          fecha_solicitud: solicitudData.fecha_solicitud,
          // Agregar otros campos importantes según el tipo de formulario
        }
      });
  
      // Procesar datos adicionales si el reporte lo requiere (como gastos)
      console.log(`🔄 Procesando datos adicionales...`);
      const additionalData = await this.processAdditionalData(solicitudId, reportConfig);
      console.log(`✅ Datos adicionales procesados:`, {
        tieneData: !!additionalData,
        camposRecibidos: Object.keys(additionalData).length
      });
  
      // Combinar datos de la solicitud con datos adicionales
      const combinedData = { ...solicitudData, ...additionalData };
      console.log(`✅ Datos combinados: ${Object.keys(combinedData).length} campos totales`);
  
      // Transformar datos utilizando la función específica de la configuración del reporte
      console.log(`🔄 Transformando datos para el reporte...`);
      const transformedData = reportConfig.transformData(combinedData);
      console.log(`✅ Datos transformados: ${Object.keys(transformedData).length} campos`);
  
      // Generar el reporte usando el servicio de Drive
      console.log(`🔄 Generando reporte en Drive...`);
      const reportLink = await driveService.generateReport(
        formNum,
        solicitudId,
        transformedData
      );
      
      if (!reportLink) {
        throw new Error('No se recibió un enlace válido del servicio de Drive');
      }
      
      console.log(`✅ Reporte generado exitosamente. Link: ${reportLink}`);
      return { link: reportLink };
    } catch (error) {
      console.error(`❌ ERROR al generar informe para solicitud ${solicitudId}, formulario ${formNumber}:`, error.message);
      console.error('📚 Stack de error:', error.stack);
      throw error;
    }
  }

  /**
   * Genera un reporte para descarga con un modo específico (view, edit)
   * @param {String} solicitudId - ID de la solicitud 
   * @param {Number} formNumber - Número de formulario (1-4)
   * @param {String} mode - Modo de acceso al documento (view, edit)
   * @returns {Promise<Object>} Resultado con link al reporte generado
   */
  async downloadReport(solicitudId, formNumber, mode = 'view') {
    console.log(`⚡️ INICIO DE GENERACIÓN DE REPORTE (MODO ${mode})`);
    try {
      // Validación básica
      if (!solicitudId || !formNumber) {
        throw new Error('solicitudId y formNumber son requeridos');
      }
      
      // Convertir a números por seguridad
      const formNum = parseInt(formNumber, 10);
      
      // Cargar config, obtener datos y generar como en el método principal
      const reportConfig = this.loadReportConfig(formNum);
      
      if (!reportConfig) {
        throw new Error(`No se encontró configuración para el formulario ${formNum} (modo ${mode})`);
      }
      
      console.log('✅ Configuración cargada correctamente para download/edit');
      
      const solicitudData = await this.getSolicitudData(solicitudId, reportConfig.sheetDefinitions);
      const additionalData = await this.processAdditionalData(solicitudId, reportConfig);
      const combinedData = { ...solicitudData, ...additionalData };
      const transformedData = reportConfig.transformData(combinedData);
      
      // Generar usando el modo especificado
      console.log(`🔄 Generando reporte para ${mode}...`);
      const reportLink = await driveService.generateReport(
        formNum,
        solicitudId,
        transformedData,
        mode
      );
      
      console.log(`✅ Reporte generado exitosamente para ${mode}. Link: ${reportLink}`);
      return { link: reportLink };
    } catch (error) {
      console.error(`❌ ERROR al generar informe para ${mode}:`, error.message);
      console.error('📚 Stack de error:', error.stack);
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
      let reportConfig;
      
      switch (formNumber) {
        case 1:
          reportConfig = require('../reportConfigs/report1Config');
          break;
        case 2:
          reportConfig = require('../reportConfigs/report2Config');
          break;
        case 3:
          reportConfig = require('../reportConfigs/report3Config');
          break;
        case 4:
          reportConfig = require('../reportConfigs/report4Config');
          break;
        default:
          throw new Error(`Número de formulario inválido: ${formNumber}`);
      }
      
      console.log(`Configuración cargada para formulario ${formNumber}: ${reportConfig.title || 'Sin título'}`);
      return reportConfig;
    } catch (error) {
      console.error(`Error al cargar configuración para formulario ${formNumber}:`, error);
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
      console.log(`🔄 Obteniendo datos desde Sheets para solicitud ${solicitudId}...`);
      if (!sheetsService || typeof sheetsService.getSolicitudData !== 'function') {
        throw new Error('El servicio sheetsService no está configurado correctamente');
      }
      
      const data = await sheetsService.getSolicitudData(solicitudId, sheetDefinitions);
      
      // Verificar si obtuvimos datos
      if (!data || typeof data !== 'object' || Object.keys(data).length === 0) {
        console.warn(`⚠️ No se encontraron datos para la solicitud ${solicitudId}`);
        return {};
      }
      
      console.log(`✅ Datos obtenidos correctamente: ${Object.keys(data).length} tablas`);
      return data;
    } catch (error) {
      console.error(`❌ Error al obtener datos de la solicitud ${solicitudId}:`, error.message);
      console.error('📚 Stack de error:', error.stack);
      throw new Error(`Error al obtener datos de la solicitud: ${error.message}`);
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
      console.log('ℹ️ Este reporte no requiere datos adicionales');
      return {};
    }

    try {
      // Procesar gastos si el reporte lo requiere
      if (reportConfig.requiresGastos) {
        console.log(`🔄 Procesando gastos para solicitud ${solicitudId}...`);
        const datos = await this.processGastosData(solicitudId);
        console.log(`✅ Gastos procesados: ${datos.gastos?.length || 0} registros`);
        return datos;
      }
      
      // Procesar riesgos si el reporte lo requiere (formulario 3)
      if (reportConfig.requiresRiesgos) {
        console.log(`🔄 Procesando riesgos para solicitud ${solicitudId}...`);
        const datos = await this.processRiesgosData(solicitudId);
        console.log(`✅ Riesgos procesados: ${datos.riesgos?.length || 0} registros`);
        return datos;
      }
      
      // Otros tipos de datos adicionales pueden ser procesados aquí
      console.log('ℹ️ No hay procesamiento adicional definido');
      return {};
    } catch (error) {
      console.error(`❌ Error al procesar datos adicionales:`, error.message);
      console.error('📚 Stack de error:', error.stack);
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
      if (!sheetsService || typeof sheetsService.getClient !== 'function') {
        throw new Error('sheetsService no está configurado correctamente');
      }
      
      const client = sheetsService.getClient();
      console.log(`🔄 Obteniendo datos de gastos desde Sheets para solicitud ${solicitudId}...`);
      
      // Verificar si tenemos acceso al ID de la hoja
      if (!sheetsService.spreadsheetId) {
        throw new Error('No se encontró el ID de la hoja de cálculo');
      }
      
      // Obtener gastos y conceptos
      const gastosResponse = await client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'GASTOS!A2:F500'
      });
      
      const conceptosResponse = await client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'CONCEPTO$!A2:F500'
      });
      
      // Procesar los datos
      const gastosRows = gastosResponse.data.values || [];
      const conceptosRows = conceptosResponse.data.values || [];
      
      console.log(`ℹ️ Registros obtenidos: ${gastosRows.length} gastos, ${conceptosRows.length} conceptos`);
      
      // Filtrar gastos de la solicitud actual
      const solicitudGastos = gastosRows.filter(row => row[1] === solicitudId);
      console.log(`✅ Encontrados ${solicitudGastos.length} gastos para la solicitud ${solicitudId}`);
      
      // Crear un mapa de conceptos para obtener nombres
      const conceptosMap = new Map();
      conceptosRows.forEach(row => {
        conceptosMap.set(row[0], {
          nombre: row[1] || row[0],
          es_padre: row[2] === 'true' || row[2] === 'TRUE'
        });
      });
      
      // Procesar gastos normales y dinámicos separadamente
      const gastosNormales = [];
      const gastosDinamicos = [];
      
      solicitudGastos.forEach(row => {
        const idConcepto = row[0];
        const cantidad = parseFloat(row[2]) || 0;
        const valorUnit = parseFloat(row[3]) || 0;
        const valorTotal = parseFloat(row[4]) || 0;
        
        // Formatear valores monetarios
        const valorUnit_formatted = new Intl.NumberFormat('es-CO', {
          style: 'currency',
          currency: 'COP',
          minimumFractionDigits: 0
        }).format(valorUnit);
        
        const valorTotal_formatted = new Intl.NumberFormat('es-CO', {
          style: 'currency',
          currency: 'COP',
          minimumFractionDigits: 0
        }).format(valorTotal);
        
        // Obtener nombre del concepto
        const concepto = conceptosMap.get(idConcepto)?.nombre || idConcepto;
        
        // Crear objeto de gasto
        const gastoObj = {
          id: idConcepto,
          concepto,
          cantidad,
          valorUnit,
          valorTotal,
          valorUnit_formatted,
          valorTotal_formatted,
          descripcion: concepto
        };
        
        // Determinar si es un gasto dinámico (empieza con 15.)
        if (idConcepto.startsWith('15.')) {
          gastosDinamicos.push(gastoObj);
        } else {
          gastosNormales.push(gastoObj);
        }
      });
      
      console.log(`📊 Gastos procesados: ${gastosNormales.length} normales, ${gastosDinamicos.length} dinámicos`);
      
      return { 
        gastos: solicitudGastos,
        gastosNormales,
        gastosDinamicos,
        gastosFormateados: {
          normales: gastosNormales,
          dinamicos: gastosDinamicos
        }
      };
    } catch (error) {
      console.error(`❌ Error al procesar gastos para solicitud ${solicitudId}:`, error.message);
      console.error('📚 Stack de error:', error.stack);
      return {};
    }
  }

  /**
   * Procesa datos de riesgos para una solicitud
   * @param {String} solicitudId - ID de la solicitud
   * @returns {Promise<Object>} Datos de riesgos procesados
   */
  async processRiesgosData(solicitudId) {
    try {
      if (!sheetsService || typeof sheetsService.getClient !== 'function') {
        throw new Error('sheetsService no está configurado correctamente');
      }
      
      const client = sheetsService.getClient();
      console.log(`🔄 Obteniendo datos de riesgos desde Sheets para solicitud ${solicitudId}...`);
      
      // Verificar si tenemos acceso al ID de la hoja
      if (!sheetsService.spreadsheetId) {
        throw new Error('No se encontró el ID de la hoja de cálculo');
      }
      
      // Obtener riesgos desde Google Sheets
      const riesgosResponse = await client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'RIESGOS!A2:F500'
      });
      
      // Procesar los datos de riesgos
      const riesgosRows = riesgosResponse.data.values || [];
      console.log(`ℹ️ Registros obtenidos: ${riesgosRows.length} riesgos`);
      
      // Filtrar riesgos de la solicitud actual
      const solicitudRiesgos = riesgosRows.filter(row => row[4] === solicitudId);
      console.log(`✅ Encontrados ${solicitudRiesgos.length} riesgos para la solicitud ${solicitudId}`);
      
      // Si no hay riesgos, retornar objeto vacío
      if (solicitudRiesgos.length === 0) {
        return { riesgos: [], riesgosPorCategoria: {} };
      }
      
      // Procesar riesgos para crear objetos con datos normalizados
      const riesgos = [];
      const riesgosPorCategoria = {
        diseno: [],
        locacion: [],
        desarrollo: [],
        cierre: [],
        otros: []
      };
      
      solicitudRiesgos.forEach(row => {
        // Extraer datos básicos
        const id = row[0] || '';
        const nombreRiesgo = row[1] || '';
        const aplica = row[2] || 'No';
        const mitigacion = row[3] || '';
        const idSolicitud = row[4] || '';
        const categoria = (row[5] || 'otros').toLowerCase();
        
        // Crear objeto normalizado
        const riesgoObj = {
          id_riesgo: id,
          nombre_riesgo: nombreRiesgo,
          aplica: aplica,
          mitigacion: mitigacion,
          id_solicitud: idSolicitud,
          categoria: categoria,
          
          // Campos adicionales para el templateMapper
          id: id,
          descripcion: nombreRiesgo,
          impacto: aplica === 'Sí' || aplica === 'Si' ? 'Alto' : 'Bajo',
          probabilidad: aplica === 'Sí' || aplica === 'Si' ? 'Alta' : 'Baja',
          estrategia: mitigacion
        };
        
        // Añadir a la lista principal
        riesgos.push(riesgoObj);
        
        // Clasificar por categoría
        let categoriaAsignada = 'otros';
        
        if (categoria.includes('dise')) {
          categoriaAsignada = 'diseno';
        } else if (categoria.includes('loca')) {
          categoriaAsignada = 'locacion';
        } else if (categoria.includes('desa')) {
          categoriaAsignada = 'desarrollo';
        } else if (categoria.includes('cier')) {
          categoriaAsignada = 'cierre';
        }
        
        // Añadir a la categoría asignada
        riesgosPorCategoria[categoriaAsignada].push(riesgoObj);
      });
      
      console.log(`📊 Riesgos procesados y categorizados:`);
      Object.keys(riesgosPorCategoria).forEach(cat => {
        console.log(`- ${cat}: ${riesgosPorCategoria[cat].length} riesgos`);
      });
      
      return { 
        riesgos: riesgos,
        riesgosPorCategoria: riesgosPorCategoria
      };
    } catch (error) {
      console.error(`❌ Error al procesar riesgos para solicitud ${solicitudId}:`, error.message);
      console.error('📚 Stack de error:', error.stack);
      return { riesgos: [], riesgosPorCategoria: {} };
    }
  }
}

module.exports = new ReportGenerationService();