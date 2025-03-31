const sheetsService = require('./sheetsService');
const driveService = require('./driveService');
const dateUtils = require('../utils/dateUtils');

/**
 * Servicio para la generación y administración de reportes
*/

class ReportService {
  /**
   * Formatea partes de una fecha en formato día, mes, año
   * @param {string} date - Fecha en cualquier formato reconocible por Date()
   * @returns {Object} Objeto con día, mes y año formateados
    */
  formatDateParts(date) {
    return dateUtils.getDateParts(date);
  }

  /**
   * Transforma los datos para adaptarlos al formato requerido por las plantillas
   * @param {Object} formData - Datos del formulario base
   * @param {Number} formNumber - Número de formulario (1-4)
   * @param {Object} allSheetData - Datos de todas las hojas relacionadas
   * @returns {Object} Datos transformados listos para la plantilla
   */
  transformDataForTemplate(formData, formNumber, allSheetData) {
    try {
      // Verificar los datos antes de transformarlos
      verificarDatosSolicitud(formData);
      
      // Corregir campos mal mapeados si es necesario
      const formDataCorregido = corregirCamposMalMapeados(formData);

      console.log("Datos brutos recibidos para transformación:", formDataCorregido);
  
      const safeFormData = formDataCorregido || {};
      let transformedData = {...safeFormData};

      // Combinar con datos de SOLICITUDES3 siempre que estén disponibles
    if (allSheetData.SOLICITUDES3) {
      console.log("Combinando datos con SOLICITUDES3");
      transformedData = {
        ...transformedData,
        ...allSheetData.SOLICITUDES3
      };
    }
  
      // Caso especial para formulario 2
      if (formNumber === 2 && allSheetData.SOLICITUDES2) {
        transformedData = {
          ...transformedData,
          ...allSheetData.SOLICITUDES2
        };
        return transformedData;
      }
  
      // Procesamiento de fechas
      if (formDataCorregido.fecha_solicitud) {
        const { dia, mes, anio } = this.formatDateParts(formDataCorregido.fecha_solicitud);
        transformedData.dia = dia;
        transformedData.mes = mes;
        transformedData.anio = anio;
      } else {
        // Valores por defecto
        transformedData.dia = '';
        transformedData.mes = '';
        transformedData.anio = '';
      }
  
      // Cálculo de becas totales
      const becasTotal = 
        Number(formDataCorregido.becas_convenio || 0) +
        Number(formDataCorregido.becas_estudiantes || 0) +
        Number(formDataCorregido.becas_docentes || 0) +
        Number(formDataCorregido.becas_egresados || 0) +
        Number(formDataCorregido.becas_funcionarios || 0) +
        Number(formDataCorregido.becas_otros || 0);
        transformedData.becas_total = becasTotal;
  
      // Procesamiento de casillas de selección para tipo de actividad
      const tipo = formDataCorregido.tipo || '';
      const tipoData = {
        tipo_curso: tipo === 'Curso' ? 'X' : '',
        tipo_congreso: tipo === 'Congreso' ? 'X' : '',
        tipo_conferencia: tipo === 'Conferencia' ? 'X' : '',
        tipo_simposio: tipo === 'Simposio' ? 'X' : '',
        tipo_diplomado: tipo === 'Diplomado' ? 'X' : '',
        tipo_otro: tipo === 'Otro' ? 'X' : '',
        tipo_otro_cual: tipo === 'Otro' ? formDataCorregido.otro_tipo || '' : '',
      };
  
      // Procesamiento de casillas de selección para modalidad
      const modalidad = formDataCorregido.modalidad || '';
      const modalidadData = {
        modalidad_presencial: modalidad === 'Presencial' ? 'X' : '',
        modalidad_semipresencial: modalidad === 'Semipresencial' ? 'X' : '',
        modalidad_virtual: modalidad === 'Virtual' ? 'X' : '',
        modalidad_mixta: modalidad === 'Mixta' ? 'X' : '',
        modalidad_todas: modalidad === 'Todas las anteriores' ? 'X' : '',
      };
  
      // Procesamiento de casillas de selección para periodicidad - CORREGIDO
      const periodicidad = (formDataCorregido.periodicidad_oferta || '').toLowerCase();
      console.log(`DEBUG: Procesando periodicidad_oferta: "${periodicidad}" (original: "${formDataCorregido.periodicidad_oferta}")`);
      
      const periodicidadData = {
        periodicidad_anual: periodicidad === 'anual' ? 'X' : '',
        periodicidad_semestral: periodicidad === 'semestral' ? 'X' : '',
        periodicidad_permanente: periodicidad === 'permanente' ? 'X' : '',
      };
      
      // Log para verificación
      console.log("DEBUG: Valores marcados en periodicidad:", 
        JSON.stringify(periodicidadData, null, 2));
  
      // Procesamiento de casillas de selección para organización - CORREGIDO
      const organizacion = formDataCorregido.organizacion_actividad || '';
      console.log(`DEBUG: Procesando organizacion_actividad: "${organizacion}"`);
      
      const organizacionData = {
        organizacion_ofi_ext: organizacion === 'ofi_ext' ? 'X' : '',
        organizacion_unidad_acad: organizacion === 'unidad_acad' ? 'X' : '',
        organizacion_otro: organizacion === 'otro_act' ? 'X' : '',
        organizacion_otro_cual: organizacion === 'otro_act' ? formDataCorregido.otro_tipo_act || '' : '',
      };
      
      // Log para verificación
      console.log("DEBUG: Valores marcados en organización:", 
        JSON.stringify(organizacionData, null, 2));
      
      // NUEVO: Procesamiento detallado de certificados
      const certificadoData = {
        certificado_asistencia: '',
        certificado_aprobacion: '',
        certificado_no_otorga: '',
        porcentaje_asistencia_minima: '',
        metodo_control_asistencia: '',
        calificacion_minima: '',
        escala_calificacion: '',
        metodo_evaluacion: '',
        registro_calificacion_participante: '',
        razon_no_certificado_texto: ''
      };
      
      // Marcar con X según el valor seleccionado y añadir información adicional
      const certificadoTipo = formDataCorregido.certificado_solicitado || '';
      if (certificadoTipo === 'De asistencia') {
        certificadoData.certificado_asistencia = 'X';
        certificadoData.porcentaje_asistencia_minima = formDataCorregido.porcentaje_asistencia_minima || '';
        certificadoData.metodo_control_asistencia = formDataCorregido.metodo_control_asistencia || '';
        certificadoData.registro_calificacion_participante = formDataCorregido.registro_calificacion_participante || '';
      } 
      else if (certificadoTipo === 'De aprobación') {
        certificadoData.certificado_aprobacion = 'X';
        certificadoData.calificacion_minima = formDataCorregido.calificacion_minima || '';
        certificadoData.escala_calificacion = formDataCorregido.escala_calificacion || '';
        certificadoData.metodo_evaluacion = formDataCorregido.metodo_evaluacion || '';
        certificadoData.registro_calificacion_participante = formDataCorregido.registro_calificacion_participante || '';
      } 
      else if (certificadoTipo === 'No otorga certificado') {
        certificadoData.certificado_no_otorga = 'X';
        certificadoData.razon_no_certificado_texto = formDataCorregido.razon_no_certificado || '';
      }
  
      // Procesamiento para extensión solidaria - MEJORADO
      const extensionSolidaria = (formDataCorregido.extension_solidaria || '').toString().toLowerCase();
      console.log(`DEBUG: Procesando extension_solidaria: "${extensionSolidaria}" (original: "${formDataCorregido.extension_solidaria}")`);
      
      // Manejo de valores numéricos y cadenas
      const esExtensionSi = extensionSolidaria === 'si' || extensionSolidaria === '1' || extensionSolidaria === 1;
      const esExtensionNo = extensionSolidaria === 'no' || extensionSolidaria === '0' || extensionSolidaria === 0;
      
      const extensionSolidariaData = {
        extension_solidaria_si: esExtensionSi ? 'X' : '',
        extension_solidaria_no: esExtensionNo ? 'X' : '',
        costo_extension_solidaria: esExtensionSi ? formDataCorregido.costo_extension_solidaria || '' : ''
      };
      
      // Log para verificación
      console.log("DEBUG: Valores marcados en extensión solidaria:", 
        JSON.stringify(extensionSolidariaData, null, 2));
  
      // Combinar todos los datos transformados
      return {
        ...transformedData,
        ...tipoData,
        ...modalidadData,
        ...periodicidadData,
        ...organizacionData,
        ...certificadoData,  // Añadir los datos de certificados
        ...extensionSolidariaData
      };
    } catch (error) {
      console.error("Error en transformDataForTemplate:", error);
      throw new Error("Error transformando datos para plantilla");
    }
  }

  /**
   * Definición de hojas necesarias para los reportes
   * @returns {Object} Estructura de hojas y sus configuraciones
   */
  get reportSheetDefinitions() {
    return {
      SOLICITUDES: {
        range: 'SOLICITUDES!A2:AU', // Ampliamos el rango para incluir todas las columnas
        fields: [
          'id_solicitud', 'fecha_solicitud', 'nombre_actividad', 'nombre_solicitante', 'dependencia_tipo',
          'nombre_escuela', 'nombre_departamento', 'nombre_seccion', 'nombre_dependencia', 'introduccion',
          'objetivo_general', 'objetivos_especificos', 'justificacion', 'metodologia', 'tipo', 'otro_tipo',
          'modalidad', 'horas_trabajo_presencial', 'horas_sincronicas', 'total_horas', 'programCont',
          'dirigidoa', 'creditos', 'cupo_min', 'cupo_max', 'nombre_coordinador', 'correo_coordinador',
          'tel_coordinador', 'pefil_competencia', 'formas_evaluacion', 'certificado_solicitado',
          'calificacion_minima', 'razon_no_certificado', 'valor_inscripcion', 'becas_convenio',
          'becas_estudiantes', 'becas_docentes', 'becas_egresados', 'becas_funcionarios', 'becas_otros',
          'becas_total', 'periodicidad_oferta', 'organizacion_actividad', 'otro_tipo_act',
          // Campos adicionales detectados:
          'extension_solidaria', 'costo_extension_solidaria', 'personal_externo', 'pieza_grafica',
          // Campos específicos para certificados:
          'porcentaje_asistencia_minima', 'metodo_control_asistencia', 'escala_calificacion',
          'metodo_evaluacion', 'registro_calificacion_participante'
        ]
      },
      // El resto de las hojas permanece igual
      SOLICITUDES2: {
        range: 'SOLICITUDES2!A2:CL',
        fields: sheetsService.fieldDefinitions.SOLICITUDES2
      },
      SOLICITUDES3: {
        range: 'SOLICITUDES3!A2:AC',
        fields: sheetsService.fieldDefinitions.SOLICITUDES3
      },
      SOLICITUDES4: {
        range: 'SOLICITUDES4!A2:BK',
        fields: sheetsService.fieldDefinitions.SOLICITUDES4
      }
    };
  }

  /**
   * Genera un reporte basado en datos de una solicitud
   * @param {String} solicitudId - ID de la solicitud
   * @param {Number} formNumber - Número de formulario (1-4)
   * @returns {Promise<String>} URL del reporte generado
   */
  async generateReport(solicitudId, formNumber) {
    try {
      // Validar parámetros
      if (!solicitudId || !formNumber) {
        throw new Error('Los parámetros solicitudId y formNumber son requeridos');
      }

      // Usar el getter de definiciones de hojas
      const hojas = this.reportSheetDefinitions;

      // Obtener datos de la solicitud de Google Sheets
      const solicitudData = await this.getSolicitudData(solicitudId, hojas);

      // Transformar datos para la plantilla
      console.log('Transformando datos para la plantilla...');
      const transformedData = this.transformDataForTemplate(
        solicitudData.SOLICITUDES,
        parseInt(formNumber),
        solicitudData
      );

      // Generar el reporte usando el servicio de Drive
      console.log(`Generando reporte para el formulario ${formNumber}...`);
      const reportLink = await driveService.generateReport(
        parseInt(formNumber),
        solicitudId,
        transformedData
      );

      return {
        message: `Informe generado exitosamente para el formulario ${formNumber}`,
        link: reportLink
      };
    } catch (error) {
      console.error('Error al generar el informe:', error.message);
      throw new Error(`Error al generar informe: ${error.message}`);
    }
  }

  /**
   * Genera un reporte para descargar o editar
   * @param {String} solicitudId - ID de la solicitud
   * @param {Number} formNumber - Número de formulario (1-4)
   * @param {String} mode - Modo de visualización ('view' o 'edit')
   * @returns {Promise<String>} URL del reporte generado
   */

  async downloadReport(solicitudId, formNumber, mode = 'view') {
    try {
      // Validar parámetros
      if (!solicitudId || !formNumber) {
        throw new Error('Los parámetros solicitudId y formNumber son requeridos');
      }

      // Usar el mismo getter aquí también
      const hojas = this.reportSheetDefinitions;

      // Obtener datos de la solicitud
      const solicitudData = await this.getSolicitudData(solicitudId, hojas);

      // Transformar los datos para ajustarse a los marcadores de la plantilla
      const transformedData = this.transformDataForTemplate(
        solicitudData.SOLICITUDES,
        parseInt(formNumber),
        solicitudData
      );

      // Generar reporte con el modo especificado
      const reportLink = await driveService.generateReport(
        parseInt(formNumber),
        solicitudId,
        transformedData,
        mode
      );

      return {
        message: `Informe generado exitosamente para el formulario ${formNumber}`,
        link: reportLink
      };
    } catch (error) {
      console.error('Error al generar el informe para descarga/edición:', error.message);
      throw new Error(`Error al generar informe para descarga: ${error.message}`);
    }
  }

  /**
   * Método auxiliar para obtener datos de una solicitud
   * @param {String} solicitudId - ID de la solicitud
   * @param {Object} hojas - Definición de hojas y campos
   * @returns {Promise<Object>} Datos de la solicitud
   */
  
  async getSolicitudData(solicitudId, hojas) {
    try {
      // Usar el servicio de hojas para obtener los datos
      return await sheetsService.getSolicitudData(solicitudId, hojas);
    } catch (error) {
      console.error('Error al obtener los datos de la solicitud:', error.message);
      throw new Error('Error al obtener los datos de la solicitud');
    }
  }
}

// En reportService.js, añade esta función para verificar datos de entrada antes de transformarlos

const verificarDatosSolicitud = (formData) => {
  console.log("=========== VERIFICACIÓN DE DATOS DE SOLICITUD ===========");
  console.log(`Periodicidad: "${formData.periodicidad_oferta}"`);
  console.log(`Organización: "${formData.organizacion_actividad}"`);
  console.log(`Extensión Solidaria: "${formData.extension_solidaria}"`);
  
  // Verificar si hay valores mezclados
  if (formData.periodicidad_oferta === 'ofi_ext' || 
      formData.periodicidad_oferta === 'unidad_acad' || 
      formData.periodicidad_oferta === 'otro_act') {
    console.warn("⚠️ ADVERTENCIA: El valor de periodicidad_oferta parece ser un valor de organización");
  }
  
  if (formData.organizacion_actividad === 'si' || formData.organizacion_actividad === 'no') {
    console.warn("⚠️ ADVERTENCIA: El valor de organizacion_actividad parece ser un valor de extensión solidaria");
  }
  
  if (formData.extension_solidaria === 'anual' || 
      formData.extension_solidaria === 'semestral' || 
      formData.extension_solidaria === 'permanente') {
    console.warn("⚠️ ADVERTENCIA: El valor de extension_solidaria parece ser un valor de periodicidad");
  }
  
  console.log("===========================================================");
};

// En reportService.js, añade esta función para corregir campos mal mapeados

const corregirCamposMalMapeados = (formData) => {
  const dataCopia = {...formData};
  
  // Si periodicidad tiene un valor de organización, intercambiarlos
  if (['ofi_ext', 'unidad_acad', 'otro_act'].includes(dataCopia.periodicidad_oferta)) {
    console.log("🔄 Corrigiendo mapeo: intercambiando periodicidad y organización");
    const temp = dataCopia.periodicidad_oferta;
    dataCopia.periodicidad_oferta = dataCopia.organizacion_actividad;
    dataCopia.organizacion_actividad = temp;
  }
  
  // Si organización tiene un valor de extensión solidaria, intercambiarlos
  if (['si', 'no'].includes(dataCopia.organizacion_actividad) && 
      !['si', 'no'].includes(dataCopia.extension_solidaria)) {
    console.log("🔄 Corrigiendo mapeo: intercambiando organización y extensión solidaria");
    const temp = dataCopia.organizacion_actividad;
    dataCopia.organizacion_actividad = dataCopia.extension_solidaria;
    dataCopia.extension_solidaria = temp;
  }
  
  // Normalizar periodicidad (asegurarse de que sea anual, semestral o permanente)
  if (!['anual', 'semestral', 'permanente'].includes(dataCopia.periodicidad_oferta?.toLowerCase())) {
    console.log(`⚠️ Valor de periodicidad no reconocido: "${dataCopia.periodicidad_oferta}"`);
    // Establecer un valor por defecto
    dataCopia.periodicidad_oferta = 'semestral';
  }
  
  // NUEVO: Normalizar valores numéricos de extensión solidaria
  if (dataCopia.extension_solidaria === 0 || dataCopia.extension_solidaria === '0') {
    console.log("🔄 Corrigiendo extensión solidaria: '0' -> 'no'");
    dataCopia.extension_solidaria = 'no';
  } else if (dataCopia.extension_solidaria === 1 || dataCopia.extension_solidaria === '1') {
    console.log("🔄 Corrigiendo extensión solidaria: '1' -> 'si'");
    dataCopia.extension_solidaria = 'si';
  }
  
  return dataCopia;
};

module.exports = new ReportService();