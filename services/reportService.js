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
      const additionalData = await this.processAdditionalData(solicitudId, reportConfig, solicitudData);
      console.log(`✅ Datos adicionales procesados:`, {
        tieneData: !!additionalData,
        camposRecibidos: Object.keys(additionalData).length
      });
  
            // APLANAR los datos anidados por hoja en un solo objeto antes de combinar
      console.log(`🔄 Aplanando datos anidados por hoja...`);
      
      // LOG INMEDIATO: Verificar estructura de datos
      console.log(`🔍 [DEBUG] Estructura de solicitudData recibida:`, JSON.stringify(solicitudData, null, 2));
      console.log(`🔍 [DEBUG] Tipo de solicitudData:`, typeof solicitudData);
      console.log(`🔍 [DEBUG] Keys de solicitudData:`, Object.keys(solicitudData || {}));
      console.log(`🔍 [DEBUG] formNum:`, formNum);
      
      let flattenedSolicitudData = {};
      if (typeof solicitudData === 'object' && solicitudData !== null) {
        // NUEVO: Para reporte 2, priorizar SOLICITUDES2 sobre SOLICITUDES
        if (formNum === 2) {
          console.log(`🔄 [REPORTE 2] Procesando con prioridad a SOLICITUDES2...`);
          
          // Primero procesar SOLICITUDES (datos básicos)
          if (solicitudData.SOLICITUDES) {
            console.log(`📋 Agregando datos básicos de SOLICITUDES:`, solicitudData.SOLICITUDES);
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES };
          }
          
          // Luego procesar SOLICITUDES2 (sobreescribir campos críticos)
          if (solicitudData.SOLICITUDES2) {
            console.log(`📋 Sobreescribiendo con datos críticos de SOLICITUDES2:`, solicitudData.SOLICITUDES2);
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES2 };
          }
          
          // Procesar otras hojas si existen
          Object.keys(solicitudData).forEach(hoja => {
            if (hoja !== 'SOLICITUDES' && hoja !== 'SOLICITUDES2' && 
                typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
              flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
            }
          });
          
          console.log(`🔍 [REPORTE 2] Campos críticos después del aplanado:`, {
            id_solicitud: flattenedSolicitudData.id_solicitud,
            nombre_actividad: flattenedSolicitudData.nombre_actividad,
            fecha_solicitud: flattenedSolicitudData.fecha_solicitud,
            ingresos_cantidad: flattenedSolicitudData.ingresos_cantidad,
            ingresos_vr_unit: flattenedSolicitudData.ingresos_vr_unit,
            total_ingresos: flattenedSolicitudData.total_ingresos
          });
        } else if (formNum === 3) {
          console.log(`🔄 [REPORTE 3] Procesando con prioridad a SOLICITUDES3...`);
          
          // Primero procesar SOLICITUDES (datos básicos)
          if (solicitudData.SOLICITUDES) {
            console.log(`📋 Agregando datos básicos de SOLICITUDES:`, solicitudData.SOLICITUDES);
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES };
          }
          
          // Luego procesar SOLICITUDES3 (sobreescribir campos críticos)
          if (solicitudData.SOLICITUDES3) {
            console.log(`📋 Sobreescribiendo con datos críticos de SOLICITUDES3:`, solicitudData.SOLICITUDES3);
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES3 };
          }
          
          // Procesar otras hojas si existen (como RIESGOS)
          Object.keys(solicitudData).forEach(hoja => {
            if (hoja !== 'SOLICITUDES' && hoja !== 'SOLICITUDES3' && 
                typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
              flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
            }
          });
          
          console.log(`🔍 [REPORTE 3] Campos críticos después del aplanado:`, {
            id_solicitud: flattenedSolicitudData.id_solicitud,
            nombre_actividad: flattenedSolicitudData.nombre_actividad,
            fecha_solicitud: flattenedSolicitudData.fecha_solicitud,
            nombre_solicitante: flattenedSolicitudData.nombre_solicitante,
            proposito: flattenedSolicitudData.proposito,
            comentario: flattenedSolicitudData.comentario,
            programa: flattenedSolicitudData.programa
          });
        } else if (formNum === 4) {
          console.log(`🔄 [REPORTE 4] Procesando con prioridad a SOLICITUDES4...`);
          
          // Solo tomar fecha_solicitud y nombre_dependencia de SOLICITUDES
          if (solicitudData.SOLICITUDES) {
            const solicitudesBasicas = {
              fecha_solicitud: solicitudData.SOLICITUDES.fecha_solicitud,
              nombre_dependencia: solicitudData.SOLICITUDES.nombre_dependencia,
              id_solicitud: solicitudData.SOLICITUDES.id_solicitud,
              nombre_actividad: solicitudData.SOLICITUDES.nombre_actividad,
              nombre_solicitante: solicitudData.SOLICITUDES.nombre_solicitante
            };
            console.log(`📋 Agregando datos básicos limitados de SOLICITUDES:`, solicitudesBasicas);
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudesBasicas };
          }
          
          // Luego procesar SOLICITUDES4 (todos los campos específicos del formulario 4)
          if (solicitudData.SOLICITUDES4) {
            console.log(`📋 Agregando datos específicos de SOLICITUDES4:`, solicitudData.SOLICITUDES4);
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES4 };
          }
          
          // Procesar otras hojas si existen
          Object.keys(solicitudData).forEach(hoja => {
            if (hoja !== 'SOLICITUDES' && hoja !== 'SOLICITUDES4' && 
                typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
              flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
            }
          });
          
          console.log(`🔍 [REPORTE 4] Campos críticos después del aplanado:`, {
            id_solicitud: flattenedSolicitudData.id_solicitud,
            nombre_actividad: flattenedSolicitudData.nombre_actividad,
            fecha_solicitud: flattenedSolicitudData.fecha_solicitud,
            nombre_dependencia: flattenedSolicitudData.nombre_dependencia,
            descripcionPrograma: flattenedSolicitudData.descripcionPrograma,
            beneficiosTangibles: flattenedSolicitudData.beneficiosTangibles,
            particulares: flattenedSolicitudData.particulares,
            dofaDebilidades: flattenedSolicitudData.dofaDebilidades
          });
        } else {
          // Para otros reportes, usar la lógica original
          Object.keys(solicitudData).forEach(hoja => {
            if (typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
              flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
            }
          });
        }
        
        // Si no hay datos anidados, usar los datos tal como vienen
        if (Object.keys(flattenedSolicitudData).length === 0) {
          flattenedSolicitudData = solicitudData;
        }
      } else {
        flattenedSolicitudData = solicitudData;
      }
      
      console.log(`✅ Datos aplanados: ${Object.keys(flattenedSolicitudData).length} campos`);
      console.log(`🔍 Campos críticos aplanados:`, {

        observaciones_cambios: flattenedSolicitudData.observaciones_cambios,
        tipo: flattenedSolicitudData.tipo,
        modalidad: flattenedSolicitudData.modalidad
      });

      // Combinar datos aplanados con datos adicionales
      let combinedData = { ...flattenedSolicitudData, ...additionalData };

      // --- FIX para reporte 2: pasar SOLICITUDES2 como objeto plano ---
      if (formNum === 2 && solicitudData.SOLICITUDES2) {
        combinedData.SOLICITUDES2 = solicitudData.SOLICITUDES2;
        console.log(`🔍 [REPORTE 2] SOLICITUDES2 preservado como objeto separado:`, combinedData.SOLICITUDES2);
      }
      // --- FIX para reporte 3: pasar SOLICITUDES y SOLICITUDES3 como objetos separados ---
      if (formNum === 3) {
        if (solicitudData.SOLICITUDES) {
          combinedData.SOLICITUDES = solicitudData.SOLICITUDES;
          console.log(`🔍 [REPORTE 3] SOLICITUDES preservado como objeto separado:`, combinedData.SOLICITUDES);
        }
        if (solicitudData.SOLICITUDES3) {
          combinedData.SOLICITUDES3 = solicitudData.SOLICITUDES3;
          console.log(`🔍 [REPORTE 3] SOLICITUDES3 preservado como objeto separado:`, combinedData.SOLICITUDES3);
        }
      }
      // --- FIX para reporte 4: pasar SOLICITUDES y SOLICITUDES4 como objetos separados ---
      if (formNum === 4) {
        if (solicitudData.SOLICITUDES) {
          combinedData.SOLICITUDES = solicitudData.SOLICITUDES;
          console.log(`🔍 [REPORTE 4] SOLICITUDES preservado como objeto separado:`, combinedData.SOLICITUDES);
        }
        if (solicitudData.SOLICITUDES4) {
          combinedData.SOLICITUDES4 = solicitudData.SOLICITUDES4;
          console.log(`🔍 [REPORTE 4] SOLICITUDES4 preservado como objeto separado:`, combinedData.SOLICITUDES4);
        }
      }
      // --- FIN FIX ---

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
        transformedData,
        additionalData.riesgosPorCategoria // Pasar riesgos categorizados para procesamiento dinámico
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
      const additionalData = await this.processAdditionalData(solicitudId, reportConfig, solicitudData);
      
      // APLANAR los datos anidados por hoja en un solo objeto antes de combinar
      let flattenedSolicitudData = {};
      if (typeof solicitudData === 'object' && solicitudData !== null) {
        // NUEVO: Para reporte 2, priorizar SOLICITUDES2 sobre SOLICITUDES
        if (formNum === 2) {
          console.log(`🔄 [REPORTE 2 - DOWNLOAD] Procesando con prioridad a SOLICITUDES2...`);
          
          // Primero procesar SOLICITUDES (datos básicos)
          if (solicitudData.SOLICITUDES) {
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES };
          }
          
          // Luego procesar SOLICITUDES2 (sobreescribir campos críticos)
          if (solicitudData.SOLICITUDES2) {
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES2 };
          }
          
          // Procesar otras hojas si existen
          Object.keys(solicitudData).forEach(hoja => {
            if (hoja !== 'SOLICITUDES' && hoja !== 'SOLICITUDES2' && 
                typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
              flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
            }
          });
        } else if (formNum === 3) {
          console.log(`🔄 [REPORTE 3 - DOWNLOAD] Procesando con prioridad a SOLICITUDES3...`);
          
          // Primero procesar SOLICITUDES (datos básicos)
          if (solicitudData.SOLICITUDES) {
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES };
          }
          
          // Luego procesar SOLICITUDES3 (sobreescribir campos críticos)
          if (solicitudData.SOLICITUDES3) {
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES3 };
          }
          
          // Procesar otras hojas si existen (como RIESGOS)
          Object.keys(solicitudData).forEach(hoja => {
            if (hoja !== 'SOLICITUDES' && hoja !== 'SOLICITUDES3' && 
                typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
              flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
            }
          });
        } else if (formNum === 4) {
          console.log(`🔄 [REPORTE 4 - DOWNLOAD] Procesando con prioridad a SOLICITUDES4...`);
          
          // Solo tomar fecha_solicitud y nombre_dependencia de SOLICITUDES
          if (solicitudData.SOLICITUDES) {
            const solicitudesBasicas = {
              fecha_solicitud: solicitudData.SOLICITUDES.fecha_solicitud,
              nombre_dependencia: solicitudData.SOLICITUDES.nombre_dependencia,
              id_solicitud: solicitudData.SOLICITUDES.id_solicitud,
              nombre_actividad: solicitudData.SOLICITUDES.nombre_actividad,
              nombre_solicitante: solicitudData.SOLICITUDES.nombre_solicitante
            };
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudesBasicas };
          }
          
          // Luego procesar SOLICITUDES4 (todos los campos específicos del formulario 4)
          if (solicitudData.SOLICITUDES4) {
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES4 };
          }
          
          // Procesar otras hojas si existen
          Object.keys(solicitudData).forEach(hoja => {
            if (hoja !== 'SOLICITUDES' && hoja !== 'SOLICITUDES4' && 
                typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
              flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
            }
          });
        } else {
          // Para otros reportes, usar la lógica original
          Object.keys(solicitudData).forEach(hoja => {
            if (typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
              flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
            }
          });
        }
        
        if (Object.keys(flattenedSolicitudData).length === 0) {
          flattenedSolicitudData = solicitudData;
        }
      } else {
        flattenedSolicitudData = solicitudData;
      }
      
      let combinedData = { ...flattenedSolicitudData, ...additionalData };
      
      // Preservar datos originales para reporte 2, 3 y 4
      if (formNum === 2 && solicitudData.SOLICITUDES2) {
        combinedData.SOLICITUDES2 = solicitudData.SOLICITUDES2;
      }
      if (formNum === 3) {
        if (solicitudData.SOLICITUDES) {
          combinedData.SOLICITUDES = solicitudData.SOLICITUDES;
        }
        if (solicitudData.SOLICITUDES3) {
          combinedData.SOLICITUDES3 = solicitudData.SOLICITUDES3;
        }
      }
      if (formNum === 4) {
        if (solicitudData.SOLICITUDES) {
          combinedData.SOLICITUDES = solicitudData.SOLICITUDES;
        }
        if (solicitudData.SOLICITUDES4) {
          combinedData.SOLICITUDES4 = solicitudData.SOLICITUDES4;
        }
      }
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
   * Genera un archivo Excel local listo para descargar
   * @param {String} solicitudId - ID de la solicitud
   * @param {Number} formNumber - Número de formulario (1-4)
   * @param {String} mode - Modo de acceso al documento (view, edit)
   * @returns {Promise<{filePath: string, fileName: string}>}
   */
  async downloadReportFile(solicitudId, formNumber, mode = 'view') {
    // Cargar config, obtener datos y generar como en el método principal
    const formNum = parseInt(formNumber, 10);
    const reportConfig = this.loadReportConfig(formNum);
    if (!reportConfig) {
      throw new Error(`No se encontró configuración para el formulario ${formNum} (modo ${mode})`);
    }
    const solicitudData = await this.getSolicitudData(solicitudId, reportConfig.sheetDefinitions);
    const additionalData = await this.processAdditionalData(solicitudId, reportConfig, solicitudData);
    // Aplanar datos
    let flattenedSolicitudData = {};
    if (typeof solicitudData === 'object' && solicitudData !== null) {
      // NUEVO: Para reporte 2, priorizar SOLICITUDES2 sobre SOLICITUDES
      if (formNum === 2) {
        console.log(`🔄 [REPORTE 2 - FILE] Procesando con prioridad a SOLICITUDES2...`);
        
        // Primero procesar SOLICITUDES (datos básicos)
        if (solicitudData.SOLICITUDES) {
          flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES };
        }
        
        // Luego procesar SOLICITUDES2 (sobreescribir campos críticos)
        if (solicitudData.SOLICITUDES2) {
          flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES2 };
        }
        
        // Procesar otras hojas si existen
        Object.keys(solicitudData).forEach(hoja => {
          if (hoja !== 'SOLICITUDES' && hoja !== 'SOLICITUDES2' && 
              typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
          }
        });
      } else if (formNum === 3) {
        console.log(`🔄 [REPORTE 3 - FILE] Procesando con prioridad a SOLICITUDES3...`);
        
        // Primero procesar SOLICITUDES (datos básicos)
        if (solicitudData.SOLICITUDES) {
          flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES };
        }
        
        // Luego procesar SOLICITUDES3 (sobreescribir campos críticos)
        if (solicitudData.SOLICITUDES3) {
          flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES3 };
        }
        
        // Procesar otras hojas si existen (como RIESGOS)
        Object.keys(solicitudData).forEach(hoja => {
          if (hoja !== 'SOLICITUDES' && hoja !== 'SOLICITUDES3' && 
              typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
          }
        });
      } else if (formNum === 4) {
        console.log(`🔄 [REPORTE 4 - FILE] Procesando con prioridad a SOLICITUDES4...`);
        
        // Solo tomar fecha_solicitud y nombre_dependencia de SOLICITUDES
        if (solicitudData.SOLICITUDES) {
          const solicitudesBasicas = {
            fecha_solicitud: solicitudData.SOLICITUDES.fecha_solicitud,
            nombre_dependencia: solicitudData.SOLICITUDES.nombre_dependencia,
            id_solicitud: solicitudData.SOLICITUDES.id_solicitud,
            nombre_actividad: solicitudData.SOLICITUDES.nombre_actividad,
            nombre_solicitante: solicitudData.SOLICITUDES.nombre_solicitante
          };
          flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudesBasicas };
        }
        
        // Luego procesar SOLICITUDES4 (todos los campos específicos del formulario 4)
        if (solicitudData.SOLICITUDES4) {
          flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData.SOLICITUDES4 };
        }
        
        // Procesar otras hojas si existen
        Object.keys(solicitudData).forEach(hoja => {
          if (hoja !== 'SOLICITUDES' && hoja !== 'SOLICITUDES4' && 
              typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
          }
        });
      } else {
        // Para otros reportes, usar la lógica original
        Object.keys(solicitudData).forEach(hoja => {
          if (typeof solicitudData[hoja] === 'object' && solicitudData[hoja] !== null) {
            flattenedSolicitudData = { ...flattenedSolicitudData, ...solicitudData[hoja] };
          }
        });
      }
      
      if (Object.keys(flattenedSolicitudData).length === 0) {
        flattenedSolicitudData = solicitudData;
      }
    } else {
      flattenedSolicitudData = solicitudData;
    }
    
    let combinedData = { ...flattenedSolicitudData, ...additionalData };
    
    // Preservar datos originales para reporte 2, 3 y 4
    if (formNum === 2 && solicitudData.SOLICITUDES2) {
      combinedData.SOLICITUDES2 = solicitudData.SOLICITUDES2;
    }
    if (formNum === 3) {
      if (solicitudData.SOLICITUDES) {
        combinedData.SOLICITUDES = solicitudData.SOLICITUDES;
      }
      if (solicitudData.SOLICITUDES3) {
        combinedData.SOLICITUDES3 = solicitudData.SOLICITUDES3;
      }
    }
    if (formNum === 4) {
      if (solicitudData.SOLICITUDES) {
        combinedData.SOLICITUDES = solicitudData.SOLICITUDES;
      }
      if (solicitudData.SOLICITUDES4) {
        combinedData.SOLICITUDES4 = solicitudData.SOLICITUDES4;
      }
    }
    const transformedData = reportConfig.transformData(combinedData);
    // Generar archivo Excel local usando driveService
    const { filePath, fileName } = await driveService.generateLocalExcelReport(
      formNum,
      solicitudId,
      transformedData,
      mode
    );
    return { filePath, fileName };
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
   * @param {Object} solicitudData - Datos ya obtenidos de getSolicitudData (opcional)
   * @returns {Promise<Object>} Datos adicionales procesados
   */
  async processAdditionalData(solicitudId, reportConfig, solicitudData = null) {
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
        // Pasar solicitudData para usar datos ya obtenidos si están disponibles
        const datos = await this.processRiesgosData(solicitudId, solicitudData);
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

      const normalizeConceptId = (value) => String(value || '').trim();
      const toDotId = (value) => normalizeConceptId(value).replace(/,/g, '.');
      const toCommaId = (value) => normalizeConceptId(value).replace(/\./g, ',');
      
      // Crear un mapa de conceptos para obtener nombres
      const conceptosMap = new Map();
      conceptosRows.forEach(row => {
        const rawId = normalizeConceptId(row[0]);
        if (!rawId) return;

        const conceptData = {
          nombre: row[1] || row[0],
          es_padre: row[2] === 'true' || row[2] === 'TRUE'
        };

        // Guardar variantes para evitar fallos por formato 14.1 vs 14,1
        conceptosMap.set(rawId, conceptData);
        conceptosMap.set(toDotId(rawId), conceptData);
        conceptosMap.set(toCommaId(rawId), conceptData);
      });
      
      // Procesar gastos normales y dinámicos separadamente
      const gastosNormales = [];
      const gastosDinamicos = [];
      
      solicitudGastos.forEach(row => {
        const idConcepto = normalizeConceptId(row[0]);
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

        // Hijos del concepto 14 son dinámicos: 14.1, 14.2, 14,1, etc.
        const idConceptoDot = toDotId(idConcepto);
        const isGastoDinamico = /^14(\.\d+)+$/.test(idConceptoDot);
        
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

        if (isGastoDinamico) {
          gastosDinamicos.push(gastoObj);
          console.log(`🧩 Gasto dinámico detectado para reporte: ${idConcepto}`);
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
   * Procesa datos de riesgos para una solicitud utilizando datos ya obtenidos
   * @param {String} solicitudId - ID de la solicitud
   * @param {Object} solicitudData - Datos ya obtenidos de getSolicitudData (opcional)
   * @returns {Promise<Object>} Datos de riesgos procesados
   */
  async processRiesgosData(solicitudId, solicitudData = null) {
    try {
      let riesgosRows = [];
      let solicitudRiesgos = [];

      // Si no se proporcionan datos previos, hacer la llamada a Google Sheets
      if (!solicitudData || !solicitudData.RIESGOS) {
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
        riesgosRows = riesgosResponse.data.values || [];
        console.log(`ℹ️ Registros obtenidos: ${riesgosRows.length} riesgos`);
        solicitudRiesgos = riesgosRows.filter(row => row[4] === solicitudId); // Columna E (índice 4) es id_solicitud
        console.log(`✅ Encontrados ${solicitudRiesgos.length} riesgos para la solicitud ${solicitudId}`);
      } else {
        // Usar datos ya obtenidos - verificar si los datos de RIESGOS ya están en formato de objeto
        console.log(`🔄 Usando datos de riesgos ya obtenidos para solicitud ${solicitudId}...`);
        if (Array.isArray(solicitudData.RIESGOS)) {
          // Si RIESGOS es un array de objetos, filtrar por solicitud correcta
          solicitudRiesgos = solicitudData.RIESGOS.filter(riesgo => {
            // Verificar id_solicitud en formato de objeto
            return riesgo.id_solicitud === solicitudId;
          });
          console.log(`✅ Encontrados ${solicitudRiesgos.length} riesgos pre-procesados para la solicitud ${solicitudId}`);
        } else if (solicitudData.RIESGOS && typeof solicitudData.RIESGOS === 'object') {
          // Si es un objeto único, verificar que pertenezca a la solicitud correcta
          if (solicitudData.RIESGOS.id_solicitud === solicitudId) {
            solicitudRiesgos = [solicitudData.RIESGOS];
            console.log(`✅ Encontrado 1 riesgo pre-procesado para la solicitud ${solicitudId}`);
          } else {
            console.log(`⚠️ El riesgo pre-procesado no pertenece a la solicitud ${solicitudId}, haciendo consulta directa...`);
            // Hacer la consulta directa como fallback
            if (!sheetsService || typeof sheetsService.getClient !== 'function') {
              throw new Error('sheetsService no está configurado correctamente');
            }
            
            const client = sheetsService.getClient();
            const riesgosResponse = await client.spreadsheets.values.get({
              spreadsheetId: sheetsService.spreadsheetId,
              range: 'RIESGOS!A2:F500'
            });
            
            riesgosRows = riesgosResponse.data.values || [];
            solicitudRiesgos = riesgosRows.filter(row => row[4] === solicitudId);
            console.log(`✅ Encontrados ${solicitudRiesgos.length} riesgos (consulta directa) para la solicitud ${solicitudId}`);
          }
        } else {
          console.log(`⚠️ No se encontraron datos de RIESGOS válidos, realizando consulta directa...`);
          // Hacer la consulta directa como fallback
          if (!sheetsService || typeof sheetsService.getClient !== 'function') {
            throw new Error('sheetsService no está configurado correctamente');
          }
          
          const client = sheetsService.getClient();
          const riesgosResponse = await client.spreadsheets.values.get({
            spreadsheetId: sheetsService.spreadsheetId,
            range: 'RIESGOS!A2:F500'
          });
          
          riesgosRows = riesgosResponse.data.values || [];
          solicitudRiesgos = riesgosRows.filter(row => row[4] === solicitudId);
          console.log(`✅ Encontrados ${solicitudRiesgos.length} riesgos (consulta directa fallback) para la solicitud ${solicitudId}`);
        }
      }

      if (solicitudRiesgos.length === 0) {
        console.log(`⚠️ No se encontraron riesgos para la solicitud ${solicitudId}.`);
        return { riesgos: [], riesgosPorCategoria: {} };
      }

      const riesgos = [];
      const riesgosPorCategoria = {
        diseno: [], locacion: [], desarrollo: [], cierre: [], otros: []
      };

      solicitudRiesgos.forEach(row => {
        let id, nombreRiesgo, aplica, mitigacion, idSolicitud, categoria;

        // Verificar si es un objeto ya mapeado o un array de valores
        if (Array.isArray(row)) {
          // Formato de fila (array): [id_riesgo, nombre_riesgo, aplica, mitigacion, id_solicitud, categoria]
          id = row[0] || '';
          nombreRiesgo = row[1] || '';
          aplica = row[2] || 'No';
          mitigacion = row[3] || '';
          idSolicitud = row[4] || '';
          categoria = (row[5] || 'otros').trim().toLowerCase();
        } else if (typeof row === 'object') {
          // Formato de objeto ya mapeado
          id = row.id_riesgo || '';
          nombreRiesgo = row.nombre_riesgo || '';
          aplica = row.aplica || 'No';
          mitigacion = row.mitigacion || '';
          idSolicitud = row.id_solicitud || '';
          categoria = (row.categoria || 'otros').trim().toLowerCase();
        } else {
          console.warn('Formato de riesgo no reconocido:', row);
          return; // Saltar este elemento
        }

        const riesgoObj = {
          id_riesgo: id, nombre_riesgo: nombreRiesgo, aplica: aplica, mitigacion: mitigacion,
          id_solicitud: idSolicitud, categoria: categoria,
          // Campos adicionales para el templateMapper
          id: id, descripcion: nombreRiesgo,
          impacto: (aplica === 'Sí aplica' || aplica === 'Si aplica') ? 'Alto' : 'Bajo',
          probabilidad: (aplica === 'Sí aplica' || aplica === 'Si aplica') ? 'Alta' : 'Baja',
          estrategia: mitigacion
        };

        riesgos.push(riesgoObj);

        // Clasificar
        let categoriaAsignada = 'otros'; // Asignación por defecto
        if (categoria.includes('dise')) {
          categoriaAsignada = 'diseno';
        } else if (categoria.includes('loca')) {
          categoriaAsignada = 'locacion';
        } else if (categoria.includes('desa')) {
          categoriaAsignada = 'desarrollo';
        } else if (categoria.includes('cier')) {
          categoriaAsignada = 'cierre';
        }
        // Si ninguno coincide, permanece en 'otros'

        riesgosPorCategoria[categoriaAsignada].push(riesgoObj);
        // console.log(`Riesgo "${nombreRiesgo}" (Cat: ${categoria}) asignado a: ${categoriaAsignada}`); // Log detallado opcional
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
      return { riesgos: [], riesgosPorCategoria: {} }; // Retornar estructura vacía en caso de error
    }
  }
}

module.exports = new ReportGenerationService();