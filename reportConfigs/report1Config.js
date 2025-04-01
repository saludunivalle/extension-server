const dateUtils = require('../utils/dateUtils');

/**
 * Configuración específica para el reporte del Formulario 1 - Aprobación
 */
const report1Config = {
  title: 'Formulario de Aprobación - F-05-MP-05-01-01',
  templateId: '1xsz9YSnYEOng56eNKGV9it9EgTn0mZw1', // ID correcto de la plantilla del formulario 1
  requiresAdditionalData: false, 
  requiresGastos: false,
  
  // Definición de hojas necesarias para este reporte
  sheetDefinitions: {
    SOLICITUDES: {
      range: 'SOLICITUDES!A2:AU',
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
        'extension_solidaria', 'costo_extension_solidaria', 'personal_externo', 'pieza_grafica',
        'porcentaje_asistencia_minima', 'metodo_control_asistencia', 'escala_calificacion',
        'metodo_evaluacion', 'registro_calificacion_participante'
      ]
    },
    SOLICITUDES3: {
      range: 'SOLICITUDES3!A2:AC',
      fields: [] // Definir los campos necesarios para este reporte
    }
  },
  
  /**
   * Función específica para transformar datos del formulario 1
   * @param {Object} allData - Datos de la solicitud y adicionales
   * @returns {Object} - Datos transformados para la plantilla
   */
  transformData: function(allData) {
    // Obtener datos de la solicitud
    const formData = allData.SOLICITUDES || {};
    
    // Verificar y corregir datos
    const formDataCorregido = corregirCamposMalMapados(formData);
    
    // Inicializar objeto de resultado
    let transformedData = {...formDataCorregido};
    
    // Procesar fecha
    if (formDataCorregido.fecha_solicitud) {
      const { dia, mes, anio } = formatDateParts(formDataCorregido.fecha_solicitud);
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
    
    // Procesamiento de tipo de actividad
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
    
    // Procesamiento de modalidad
    const modalidad = formDataCorregido.modalidad || '';
    const modalidadData = {
      modalidad_presencial: modalidad === 'Presencial' ? 'X' : '',
      modalidad_semipresencial: modalidad === 'Semipresencial' ? 'X' : '',
      modalidad_virtual: modalidad === 'Virtual' ? 'X' : '',
      modalidad_mixta: modalidad === 'Mixta' ? 'X' : '',
      modalidad_todas: modalidad === 'Todas las anteriores' ? 'X' : '',
    };
    
    // Procesamiento de periodicidad
    const periodicidad = (formDataCorregido.periodicidad_oferta || '').toLowerCase();
    const periodicidadData = {
      periodicidad_anual: periodicidad === 'anual' ? 'X' : '',
      periodicidad_semestral: periodicidad === 'semestral' ? 'X' : '',
      periodicidad_permanente: periodicidad === 'permanente' ? 'X' : '',
    };
    
    // Procesamiento de organización
    const organizacion = formDataCorregido.organizacion_actividad || '';
    const organizacionData = {
      organizacion_ofi_ext: organizacion === 'ofi_ext' ? 'X' : '',
      organizacion_unidad_acad: organizacion === 'unidad_acad' ? 'X' : '',
      organizacion_otro: organizacion === 'otro_act' ? 'X' : '',
      organizacion_otro_cual: organizacion === 'otro_act' ? formDataCorregido.otro_tipo_act || '' : '',
    };
    
    // Procesamiento de certificado
    const certificadoData = procesarCertificado(formDataCorregido);
    
    // Procesamiento de extensión solidaria
    const extensionSolidaria = (formDataCorregido.extension_solidaria || '').toString().toLowerCase();
    const esExtensionSi = extensionSolidaria === 'si' || extensionSolidaria === '1' || extensionSolidaria === 1;
    const esExtensionNo = extensionSolidaria === 'no' || extensionSolidaria === '0' || extensionSolidaria === 0;
    
    const extensionSolidariaData = {
      extension_solidaria_si: esExtensionSi ? 'X' : '',
      extension_solidaria_no: esExtensionNo ? 'X' : '',
      costo_extension_solidaria: esExtensionSi ? formDataCorregido.costo_extension_solidaria || '' : ''
    };
    
    // Combinar todos los datos transformados
    return {
      ...transformedData,
      ...tipoData,
      ...modalidadData,
      ...periodicidadData,
      ...organizacionData,
      ...certificadoData,
      ...extensionSolidariaData
    };
  }
};

/**
 * Función auxiliar para formatear partes de fecha
 * @param {String} date - Fecha en cualquier formato
 * @returns {Object} - Partes de la fecha formateadas
 */
function formatDateParts(date) {
  try {
    const dateObj = new Date(date);
    if (isNaN(dateObj.getTime())) {
      // Si no es una fecha válida, intentar parsear formatos comunes
      if (date.includes('/')) {
        const parts = date.split('/');
        return {
          dia: parts[0].padStart(2, '0'),
          mes: parts[1].padStart(2, '0'),
          anio: parts[2]
        };
      } else if (date.includes('-')) {
        const parts = date.split('-');
        if (parts[0].length === 4) {
          // Formato YYYY-MM-DD
          return {
            dia: parts[2].padStart(2, '0'),
            mes: parts[1].padStart(2, '0'),
            anio: parts[0]
          };
        } else {
          // Formato DD-MM-YYYY
          return {
            dia: parts[0].padStart(2, '0'),
            mes: parts[1].padStart(2, '0'),
            anio: parts[2]
          };
        }
      }
    } else {
      // Es una fecha válida
      return {
        dia: dateObj.getDate().toString().padStart(2, '0'),
        mes: (dateObj.getMonth() + 1).toString().padStart(2, '0'),
        anio: dateObj.getFullYear().toString()
      };
    }
  } catch (error) {
    console.error('Error al formatear fecha:', error);
    // Valor por defecto
    return { dia: '', mes: '', anio: '' };
  }
}

/**
 * Función para procesar datos de certificado
 * @param {Object} formData - Datos del formulario
 * @returns {Object} - Datos procesados de certificado
 */
function procesarCertificado(formData) {
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
  
  const certificadoTipo = formData.certificado_solicitado || '';
  
  if (certificadoTipo === 'De asistencia') {
    certificadoData.certificado_asistencia = 'X';
    certificadoData.porcentaje_asistencia_minima = formData.porcentaje_asistencia_minima || '';
    certificadoData.metodo_control_asistencia = formData.metodo_control_asistencia || '';
    certificadoData.registro_calificacion_participante = formData.registro_calificacion_participante || '';
  } 
  else if (certificadoTipo === 'De aprobación') {
    certificadoData.certificado_aprobacion = 'X';
    certificadoData.calificacion_minima = formData.calificacion_minima || '';
    certificadoData.escala_calificacion = formData.escala_calificacion || '';
    certificadoData.metodo_evaluacion = formData.metodo_evaluacion || '';
    certificadoData.registro_calificacion_participante = formData.registro_calificacion_participante || '';
  } 
  else if (certificadoTipo === 'No otorga certificado') {
    certificadoData.certificado_no_otorga = 'X';
    certificadoData.razon_no_certificado_texto = formData.razon_no_certificado || '';
  }
  
  return certificadoData;
}

/**
 * Función para corregir campos mal mapeados
 * @param {Object} formData - Datos del formulario
 * @returns {Object} - Datos corregidos
 */
function corregirCamposMalMapados(formData) {
  const dataCopia = {...formData};
  
  // Si periodicidad tiene un valor de organización, intercambiarlos
  if (['ofi_ext', 'unidad_acad', 'otro_act'].includes(dataCopia.periodicidad_oferta)) {
    const temp = dataCopia.periodicidad_oferta;
    dataCopia.periodicidad_oferta = dataCopia.organizacion_actividad;
    dataCopia.organizacion_actividad = temp;
  }
  
  // Si organización tiene un valor de extensión solidaria, intercambiarlos
  if (['si', 'no'].includes(dataCopia.organizacion_actividad) && 
      !['si', 'no'].includes(dataCopia.extension_solidaria)) {
    const temp = dataCopia.organizacion_actividad;
    dataCopia.organizacion_actividad = dataCopia.extension_solidaria;
    dataCopia.extension_solidaria = temp;
  }
  
  // Normalizar periodicidad
  if (!['anual', 'semestral', 'permanente'].includes(dataCopia.periodicidad_oferta?.toLowerCase())) {
    dataCopia.periodicidad_oferta = 'semestral';
  }
  
  // Normalizar valores numéricos de extensión solidaria
  if (dataCopia.extension_solidaria === 0 || dataCopia.extension_solidaria === '0') {
    dataCopia.extension_solidaria = 'no';
  } else if (dataCopia.extension_solidaria === 1 || dataCopia.extension_solidaria === '1') {
    dataCopia.extension_solidaria = 'si';
  }
  
  return dataCopia;
}

module.exports = report1Config;