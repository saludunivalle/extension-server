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
      console.log("Datos brutos recibidos para transformación:", formData);
  
      const safeFormData = formData || {};
      let transformedData = {...safeFormData};
  
      // Caso especial para formulario 2
      if (formNumber === 2 && allSheetData.SOLICITUDES2) {
        transformedData = {
          ...transformedData,
          ...allSheetData.SOLICITUDES2
        };
        return transformedData;
      }
  
      // Procesamiento de fechas
      if (formData.fecha_solicitud) {
        const { dia, mes, anio } = this.formatDateParts(formData.fecha_solicitud);
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
        Number(formData.becas_convenio || 0) +
        Number(formData.becas_estudiantes || 0) +
        Number(formData.becas_docentes || 0) +
        Number(formData.becas_egresados || 0) +
        Number(formData.becas_funcionarios || 0) +
        Number(formData.becas_otros || 0);
        transformedData.becas_total = becasTotal;
  
      // Procesamiento de casillas de selección para tipo de actividad
      const tipo = formData.tipo || '';
      const tipoData = {
        tipo_curso: tipo === 'Curso' ? 'X' : '',
        tipo_taller: tipo === 'Taller' ? 'X' : '',
        tipo_seminario: tipo === 'Seminario' ? 'X' : '',
        tipo_diplomado: tipo === 'Diplomado' ? 'X' : '',
        tipo_programa: tipo === 'Programa' ? 'X' : '',
        otro_cual: tipo === 'Otro' ? formData.otro_tipo || '' : '',
      };
  
      // Procesamiento de casillas de selección para modalidad
      const modalidad = formData.modalidad || '';
      const modalidadData = {
        modalidad_presencial: modalidad === 'Presencial' ? 'X' : '',
        modalidad_tecnologia: modalidad === 'Presencialidad asistida por Tecnología' ? 'X' : '',
        modalidad_virtual: modalidad === 'Virtual' ? 'X' : '',
        modalidad_mixta: modalidad === 'Mixta' ? 'X' : '',
        modalidad_todas: modalidad === 'Todas las anteriores' ? 'X' : '',
      };
  
      // Procesamiento de casillas de selección para periodicidad
      const periodicidad = formData.periodicidad_oferta || '';
      const periodicidadData = {
        per_anual: periodicidad === 'Anual' ? 'X' : '',
        per_semestral: periodicidad === 'Semestral' ? 'X' : '',
        per_permanente: periodicidad === 'Permanente' ? 'X' : '',
      };
  
      // Procesamiento de casillas de selección para organización
      const organizacion = formData.organizacion_actividad || '';
      const organizacionData = {
        oficina_extension: organizacion === 'Oficina de Extensión' ? 'X' : '',
        unidad_acad: organizacion === 'Unidad Académica' ? 'X' : '',
        otro_organizacion: organizacion === 'Otro' ? 'X' : '',
        cual_otro: organizacion === 'Otro' ? formData.otro_tipo_act || '' : '',
      };
  
      // Combinar todos los datos transformados
      return {
        ...transformedData,
        ...tipoData,
        ...modalidadData,
        ...periodicidadData,
        ...organizacionData,
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
        ]
      },
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

module.exports = new ReportService();