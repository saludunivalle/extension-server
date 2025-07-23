const dateUtils = require('../utils/dateUtils');

/**
 * Configuración específica para el reporte del Formulario 4 - Mercadeo Relacional
 */
const report4Config = {
  title: 'Formulario 4 - Mercadeo Relacional',
  templateId: '1FTC7Vq3O4ultexRPXYrJKOpL9G0071-0',
  requiresAdditionalData: false,
  requiresGastos: false,
  
  // Definición de hojas necesarias para este reporte
  sheetDefinitions: {
    SOLICITUDES: {
      range: 'SOLICITUDES!A2:AU',
      fields: [
        'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'nombre_solicitante',
        'dependencia_tipo', 'nombre_escuela', 'nombre_departamento', 'nombre_seccion', 'nombre_dependencia'
      ]
    },
    SOLICITUDES4: {
      range: 'SOLICITUDES4!A2:BZ',
      fields: [
        'id_solicitud', 'descripcionPrograma', 'identificacionNecesidades', 
        'atributosBasicos', 'atributosDiferenciadores', 'competencia', 'programa',
        'programasSimilares', 'estrategiasCompetencia', 'personasInteresChecked', 'personasMatriculadasChecked',
        'otroInteres', 'innovacion', 'solicitudExterno', 'interesSondeo', 'otroMercadeo',
        'llamadas', 'encuestas', 'webinar', 'pautas_redes', 'otroEstrategias',
        'preregistroFisico', 'preregistroGoogle', 'preregistroOtro', 'gremios', 'sectores_empresariales',
        'politicas_publicas', 'otros_mesas_trabajo', 'focusGroup', 'desayunosTrabajo', 'almuerzosTrabajo',
        'openHouse', 'ferias_colegios', 'ferias_empresarial', 'otros_mercadeo', 'valorEconomico',
        'modalidadPresencial', 'modalidadVirtual', 'modalidadSemipresencial', 'traslados_docente', 'modalidad_asistida_tecnologia',
        'beneficiosTangibles', 'beneficiosIntangibles', 'particulares', 'colegios', 'empresas',
        'egresados', 'colaboradores', 'otros_publicos_potenciales', 'tendenciasActuales',
        'dofaDebilidades', 'dofaOportunidades', 'dofaFortalezas', 'dofaAmenazas',
        'paginaWeb', 'facebook', 'instagram', 'linkedin', 'correo', 'prensa', 'boletin',
        'llamadas_redes', 'otro_canal'
      ]
    }
  },
  
  transformData: function(allData) {
    console.log("🔄 [REPORT4] Iniciando transformación de datos desde Google Sheets");
    console.log("🔍 [REPORT4] Estructura de datos recibida:", {
      keys: Object.keys(allData || {}),
      hasSolicitudes: !!allData.SOLICITUDES,
      hasSolicitudes4: !!allData.SOLICITUDES4,
      solicitudes_count: allData.SOLICITUDES ? Object.keys(allData.SOLICITUDES).length : 0,
      solicitudes4_count: allData.SOLICITUDES4 ? Object.keys(allData.SOLICITUDES4).length : 0
    });

    try {
      // 1. EXTRAER DATOS DE GOOGLE SHEETS
      const solicitudData = allData.SOLICITUDES || {};
      const formData = allData.SOLICITUDES4 || {};
      
      console.log("📋 [REPORT4] Datos de SOLICITUDES:", {
        id_solicitud: solicitudData.id_solicitud,
        nombre_actividad: solicitudData.nombre_actividad,
        fecha_solicitud: solicitudData.fecha_solicitud,
        nombre_solicitante: solicitudData.nombre_solicitante,
        nombre_dependencia: solicitudData.nombre_dependencia
      });
      
      console.log("📋 [REPORT4] Datos de SOLICITUDES4 (muestra):", {
        id_solicitud: formData.id_solicitud,
        descripcionPrograma: formData.descripcionPrograma,
        beneficiosTangibles: formData.beneficiosTangibles,
        particulares: formData.particulares,
        dofaDebilidades: formData.dofaDebilidades
      });

      // 2. MAPEO EXPLÍCITO DE CAMPOS
      const result = {};

      // CAMPOS BASE (SOLO CAMPOS ESPECÍFICOS DE SOLICITUDES)
      result.id_solicitud = solicitudData.id_solicitud || '';
      result.nombre_actividad = solicitudData.nombre_actividad || '';
      result.fecha_solicitud = solicitudData.fecha_solicitud || '';
      result.nombre_solicitante = solicitudData.nombre_solicitante || '';
      result.nombre_dependencia = solicitudData.nombre_dependencia || '';

      // PASO 1 - ACTIVIDADES DE MERCADEO RELACIONAL
      result.descripcionPrograma = formData.descripcionPrograma || '';
      result.identificacionNecesidades = formData.identificacionNecesidades || '';

      // PASO 2 - VALOR ECONÓMICO DE LOS PROGRAMAS
      result.atributosBasicos = formData.atributosBasicos || '';
      result.atributosDiferenciadores = formData.atributosDiferenciadores || '';
      result.competencia = formData.competencia || '';
      result.programa = formData.programa || '';
      result.programasSimilares = formData.programasSimilares || '';
      result.estrategiasCompetencia = formData.estrategiasCompetencia || '';

      // PASO 3 - MODALIDAD DE EJECUCIÓN
      result.personasInteresChecked = formData.personasInteresChecked || 'No';
      result.personasInteresadas = formData.personasInteresadas || '0';
      result.personasMatriculadasChecked = formData.personasMatriculadasChecked || 'No';
      result.personasMatriculadas = formData.personasMatriculadas || '0';
      result.otroInteresChecked = formData.otroInteresChecked || 'No';
      result.otroInteres = formData.otroInteres || '';
      result.innovacion = formData.innovacion || 'No';
      result.solicitudExterno = formData.solicitudExterno || 'No';
      result.interesSondeo = formData.interesSondeo || 'No';
      result.otroMercadeoChecked = formData.otroMercadeoChecked || 'No';
      result.otroMercadeo = formData.otroMercadeo || '';
      result.llamadas = formData.llamadas || 'No';
      result.encuestas = formData.encuestas || 'No';
      result.webinar = formData.webinar || 'No';
      result.pautas_redes = formData.pautas_redes || 'No';
      result.otroEstrategiasChecked = formData.otroEstrategiasChecked || 'No';
      result.otroEstrategias = formData.otroEstrategias || '';
      result.preregistroFisico = formData.preregistroFisico || 'No';
      result.preregistroGoogle = formData.preregistroGoogle || 'No';
      result.preregistroOtroChecked = formData.preregistroOtroChecked || 'No';
      result.preregistroOtro = formData.preregistroOtro || '';
      result.observaciones = formData.observaciones || '';

      // PASO 4 - BENEFICIOS OFRECIDOS
      result.gremios = formData.gremios || 'No';
      result.sectores_empresariales = formData.sectores_empresariales || 'No';
      result.politicas_publicas = formData.politicas_publicas || 'No';
      result.otros_mesas_trabajoChecked = formData.otros_mesas_trabajoChecked || 'No';
      result.otros_mesas_trabajo = formData.otros_mesas_trabajo || '';
      result.focusGroup = formData.focusGroup || 'No';
      result.desayunosTrabajo = formData.desayunosTrabajo || 'No';
      result.almuerzosTrabajo = formData.almuerzosTrabajo || 'No';
      result.openHouse = formData.openHouse || 'No';
      result.ferias_colegios = formData.ferias_colegios || 'No';
      result.ferias_empresarial = formData.ferias_empresarial || 'No';
      result.otros_mercadeoChecked = formData.otros_mercadeoChecked || 'No';
      result.otros_mercadeo = formData.otros_mercadeo || '';
      result.valorEconomico = formData.valorEconomico || '';
      result.modalidadPresencial = formData.modalidadPresencial || 'No';
      result.modalidadVirtual = formData.modalidadVirtual || 'No';
      result.modalidadSemipresencial = formData.modalidadSemipresencial || 'No';
      result.traslados_docente = formData.traslados_docente || 'No';
      result.modalidad_asistida_tecnologia = formData.modalidad_asistida_tecnologia || 'No';

      // PASO 5 - DOFA DEL PROGRAMA
      result.beneficiosTangibles = formData.beneficiosTangibles || '';
      result.beneficiosIntangibles = formData.beneficiosIntangibles || '';
      
      // Modalidad del programa (públicos potenciales)
      result.particulares = formData.particulares || 'No';
      result.colegios = formData.colegios || 'No';
      result.empresas = formData.empresas || 'No';
      result.egresados = formData.egresados || 'No';
      result.colaboradores = formData.colaboradores || 'No';
      result.otros_publicos_potencialesChecked = formData.otros_publicos_potencialesChecked || 'No';
      result.otros_publicos_potenciales = formData.otros_publicos_potenciales || '';
      
      // Tendencias
      result.tendenciasActuales = formData.tendenciasActuales || '';
      
      // DOFA
      result.dofaDebilidades = formData.dofaDebilidades || '';
      result.dofaOportunidades = formData.dofaOportunidades || '';
      result.dofaFortalezas = formData.dofaFortalezas || '';
      result.dofaAmenazas = formData.dofaAmenazas || '';
      
      // Canales de divulgación
      result.paginaWeb = formData.paginaWeb || 'No';
      result.facebook = formData.facebook || 'No';
      result.instagram = formData.instagram || 'No';
      result.linkedin = formData.linkedin || 'No';
      result.correo = formData.correo || 'No';
      result.prensa = formData.prensa || 'No';
      result.boletin = formData.boletin || 'No';
      result.llamadas_redes = formData.llamadas_redes || 'No';
      result.otro_canalChecked = formData.otro_canalChecked || 'No';
      result.otro_canal = formData.otro_canal || '';

      // 3. PROCESAR FECHA
      try {
        if (result.fecha_solicitud) {
          const dateParts = dateUtils.formatDateParts(result.fecha_solicitud);
          result.dia = dateParts.dia;
          result.mes = dateParts.mes;
          result.anio = dateParts.anio;
          result.fecha_solicitud = `${dateParts.dia}/${dateParts.mes}/${dateParts.anio}`;
        } else {
          const fechaActual = new Date();
          result.dia = fechaActual.getDate().toString().padStart(2, '0');
          result.mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
          result.anio = fechaActual.getFullYear().toString();
          result.fecha_solicitud = `${result.dia}/${result.mes}/${result.anio}`;
        }
      } catch (error) {
        console.error('Error al procesar fecha:', error);
        const fechaActual = new Date();
        result.dia = fechaActual.getDate().toString().padStart(2, '0');
        result.mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
        result.anio = fechaActual.getFullYear().toString();
        result.fecha_solicitud = `${result.dia}/${result.mes}/${result.anio}`;
      }

      // 4. FORMATEAR CAMPOS DE PREREGISTRO
      const preregistroOpciones = [];
      if (result.preregistroFisico === 'Sí') preregistroOpciones.push('Físico');
      if (result.preregistroGoogle === 'Sí') preregistroOpciones.push('Google');
      if (result.preregistroOtroChecked === 'Sí' && result.preregistroOtro) {
        preregistroOpciones.push(result.preregistroOtro);
      }
      result.preregistro = preregistroOpciones.length > 0 ? preregistroOpciones.join(', ') : 'No';

      // 5. NORMALIZAR CAMPOS CHECKBOX A "Sí/No"
      const camposCheckbox = [
        'personasInteresChecked', 'personasMatriculadasChecked', 'otroInteresChecked', 'otroMercadeoChecked',
        'otroEstrategiasChecked', 'preregistroFisico', 'preregistroGoogle', 'preregistroOtroChecked',
        'gremios', 'sectores_empresariales', 'politicas_publicas', 'otros_mesas_trabajoChecked',
        'focusGroup', 'desayunosTrabajo', 'almuerzosTrabajo', 'openHouse', 'ferias_colegios', 'ferias_empresarial',
        'otros_mercadeoChecked', 'modalidadPresencial', 'modalidadVirtual', 'modalidadSemipresencial',
        'traslados_docente', 'modalidad_asistida_tecnologia',
        'particulares', 'colegios', 'empresas', 'egresados', 'colaboradores', 'otros_publicos_potencialesChecked',
        'paginaWeb', 'facebook', 'instagram', 'linkedin', 'correo', 'prensa', 'boletin',
        'llamadas_redes', 'otro_canalChecked', 'innovacion', 'solicitudExterno', 'interesSondeo',
        'llamadas', 'encuestas', 'webinar', 'pautas_redes'
      ];

      camposCheckbox.forEach(campo => {
        const valor = result[campo];
        const esAfirmativo = 
          valor === true || 
          valor === 'Sí' || 
          valor === 'Si' ||
          valor === 'sí' ||
          valor === 'si' ||
          valor === '1' ||
          valor === 1;
        
        result[campo] = esAfirmativo ? 'Sí' : 'No';
      });

      // 6. PROCESAR CAMPOS CONDICIONALES
      const camposCondicionales = [
        { checkbox: 'otroInteresChecked', campo: 'otroInteres' },
        { checkbox: 'otroMercadeoChecked', campo: 'otroMercadeo' },
        { checkbox: 'otroEstrategiasChecked', campo: 'otroEstrategias' },
        { checkbox: 'otros_mesas_trabajoChecked', campo: 'otros_mesas_trabajo' },
        { checkbox: 'otros_mercadeoChecked', campo: 'otros_mercadeo' },
        { checkbox: 'otros_publicos_potencialesChecked', campo: 'otros_publicos_potenciales' },
        { checkbox: 'otro_canalChecked', campo: 'otro_canal' }
      ];

      camposCondicionales.forEach(({ checkbox, campo }) => {
        if (result[checkbox] !== 'Sí') {
          result[campo] = 'No';
        } else if (!result[campo] || result[campo].trim() === '') {
          result[campo] = 'Sin especificar';
        }
      });

      // 7. LIMPIAR PLACEHOLDERS {{...}}
      Object.keys(result).forEach(key => {
        const value = result[key];
        if (typeof value === 'string' && (value.includes('{{') || value.includes('}}'))) {
          console.warn(`[REPORT4] Limpiando placeholder en ${key}: "${value}"`);
          result[key] = '';
        }
      });

      // 8. VALORES POR DEFECTO PARA CAMPOS CRÍTICOS
      if (!result.descripcionPrograma || result.descripcionPrograma.trim() === '') {
        result.descripcionPrograma = 'No especificado';
      }
      if (!result.identificacionNecesidades || result.identificacionNecesidades.trim() === '') {
        result.identificacionNecesidades = 'No especificado';
      }
      if (!result.beneficiosTangibles || result.beneficiosTangibles.trim() === '') {
        result.beneficiosTangibles = 'No especificado';
      }
      if (!result.beneficiosIntangibles || result.beneficiosIntangibles.trim() === '') {
        result.beneficiosIntangibles = 'No especificado';
      }
      if (!result.tendenciasActuales || result.tendenciasActuales.trim() === '') {
        result.tendenciasActuales = 'No especificado';
      }
      if (!result.dofaDebilidades || result.dofaDebilidades.trim() === '') {
        result.dofaDebilidades = 'No especificado';
      }
      if (!result.dofaOportunidades || result.dofaOportunidades.trim() === '') {
        result.dofaOportunidades = 'No especificado';
      }
      if (!result.dofaFortalezas || result.dofaFortalezas.trim() === '') {
        result.dofaFortalezas = 'No especificado';
      }
      if (!result.dofaAmenazas || result.dofaAmenazas.trim() === '') {
        result.dofaAmenazas = 'No especificado';
      }

      console.log("✅ [REPORT4] Transformación completada. Campos críticos:", {
        id_solicitud: result.id_solicitud,
        nombre_actividad: result.nombre_actividad,
        fecha_solicitud: result.fecha_solicitud,
        nombre_dependencia: result.nombre_dependencia,
        descripcionPrograma: result.descripcionPrograma?.substring(0, 50) + '...',
        beneficiosTangibles: result.beneficiosTangibles?.substring(0, 50) + '...',
        particulares: result.particulares,
        dofaDebilidades: result.dofaDebilidades?.substring(0, 50) + '...',
        total_campos: Object.keys(result).length
      });
      
      // VERIFICACIÓN FINAL: Log de campos críticos para debugging
      console.log("🔍 [REPORT4] VERIFICACIÓN FINAL - Comparación de fechas:");
      console.log(`- fecha_solicitud desde SOLICITUDES: ${solicitudData.fecha_solicitud}`);
      console.log(`- fecha_solicitud en resultado final: ${result.fecha_solicitud}`);
      console.log("🔍 [REPORT4] VERIFICACIÓN FINAL - ID de solicitud:");
      console.log(`- id_solicitud desde SOLICITUDES: ${solicitudData.id_solicitud}`);
      console.log(`- id_solicitud desde SOLICITUDES4: ${formData.id_solicitud}`);
      console.log(`- id_solicitud en resultado final: ${result.id_solicitud}`);

      return result;

    } catch (error) {
      console.error('❌ [REPORT4] Error en transformación de datos:', error);
      return {
        error: true,
        message: error.message,
        id_solicitud: '',
        nombre_actividad: 'Error al cargar',
        descripcionPrograma: 'Error al cargar datos',
        dia: new Date().getDate().toString().padStart(2, '0'),
        mes: (new Date().getMonth() + 1).toString().padStart(2, '0'),
        anio: new Date().getFullYear().toString(),
        fecha_solicitud: new Date().toLocaleDateString('es-CO')
      };
    }
  },
  
  sheetsConfig: {
    sheetName: 'Formulario4',
    dataRange: 'A1:BZ100'
  },
  footerText: 'Universidad del Valle - Extensión y Proyección Social - Mercadeo Relacional',
  watermark: false
};

module.exports = report4Config;