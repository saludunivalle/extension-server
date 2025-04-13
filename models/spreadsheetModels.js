/**
 * Modelos para las hojas de Google Sheets
 * Define la estructura de datos para cada hoja de cálculo
*/

class SpreadsheetModel {
    constructor(name, fields, columnMappings = {}) {
      this.name = name;         // Nombre de la hoja en Google Sheets
      this.fields = fields;     // Campos/columnas en orden
      this.columnMappings = columnMappings; // Mapeo de pasos a columnas
    }
  
    /**
     * Valida un objeto de datos contra el esquema del modelo
     * @param {Object} data - Datos a validar
     * @returns {Object} - Resultado de validación {isValid, errors}
    */

    validate(data) {
      const errors = [];
      const requiredFields = this.fields.filter(f => 
        typeof f === 'object' && f.required).map(f => f.name || f);
      
      // Verificar campos requeridos
      for (const field of requiredFields) {
        if (!data[field]) {
          errors.push(`Campo requerido faltante: ${field}`);
        }
      }
      
      return {
        isValid: errors.length === 0,
        errors
      };
    }
  
    /**
     * Convierte un objeto de datos a un array según el orden de campos
     * @param {Object} data - Datos a convertir
     * @returns {Array} - Array ordenado según los campos
    */

    toArray(data) {
      return this.fields.map(field => {
        const fieldName = typeof field === 'object' ? field.name : field;
        return data[fieldName] || '';
      });
    }
  }
  
  // Definición del modelo SOLICITUDES
  const SOLICITUDES = new SpreadsheetModel(
    'SOLICITUDES',
    [
      'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'nombre_solicitante', 'dependencia_tipo',
      'nombre_escuela', 'nombre_departamento', 'nombre_seccion', 'nombre_dependencia',
      'introduccion', 'objetivo_general', 'objetivos_especificos', 'justificacion', 'metodologia', 'tipo',
      'otro_tipo', 'modalidad', 'horas_trabajo_presencial', 'horas_sincronicas', 'total_horas',
      'programCont', 'dirigidoa', 'creditos', 'cupo_min', 'cupo_max', 'nombre_coordinador',
      'correo_coordinador', 'tel_coordinador', 'pefil_competencia', 'formas_evaluacion',
      'certificado_solicitado', 'calificacion_minima', 'razon_no_certificado', 'valor_inscripcion',
      'becas_convenio', 'becas_estudiantes', 'becas_docentes', 'becas_egresados', 'becas_funcionarios',
      'becas_total', 'becas_otros', 'periodicidad_oferta', 'fechas_actividad', 'organizacion_actividad'
    ],
    {
      1: ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'],
      2: ['J', 'K', 'L', 'M', 'N'],
      3: ['O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y'],
      4: ['Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH'],
      5: ['AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU']
    }
  );
  
  // Actualizar la definición del modelo SOLICITUDES2 para que coincida con las columnas reales
  const SOLICITUDES2 = new SpreadsheetModel(
    'SOLICITUDES2',
    [
      'id_solicitud', // A
      'nombre_actividad', // B 
      'fecha_solicitud', // C 
      'ingresos_cantidad', // D
      'ingresos_vr_unit', // E
      'total_ingresos', // F 
      'subtotal_gastos', // G 
      'imprevistos_3', // H
      'total_gastos_imprevistos', // I
      'fondo_comun_porcentaje', // J
      'fondo_comun', // K
      'facultad_instituto', // L (Previously missing porcentaje)
      'escuela_departamento_porcentaje', // M
      'escuela_departamento', // N
      'total_recursos', // O
      'observaciones', // P (Assuming this is next)
      'responsable_financiero' // Q (Assuming this is next)
    ],
    {
      1: ['B', 'C'], // Step 1: nombre_actividad, fecha_solicitud
      2: ['D', 'E', 'F', 'G', 'H', 'I'], // Step 2: Ingresos y Gastos Totales
      3: ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q'] // Step 3: Aportes y Resumen (J to Q)
    }
  );
  
  // Definición del modelo SOLICITUDES3
  const SOLICITUDES3 = new SpreadsheetModel(
    'SOLICITUDES3',
    [
      'id_solicitud', 'proposito', 'comentario', 'programa', 'fecha_solicitud', 'nombre_solicitante', 'aplicaDiseno1', 'aplicaDiseno2', 
      'aplicaDiseno3', 'aplicaDiseno4', 'aplicaLocacion1', 'aplicaLocacion2', 'aplicaLocacion3', 'aplicaLocacion4', 'aplicaLocacion5', 
      'aplicaDesarrollo1', 'aplicaDesarrollo2', 'aplicaDesarrollo3', 'aplicaDesarrollo4', 'aplicaDesarrollo5', 'aplicaDesarrollo6', 
      'aplicaDesarrollo7', 'aplicaDesarrollo8', 'aplicaDesarrollo9', 'aplicaDesarrollo10', 'aplicaDesarrollo11', 'aplicaCierre1', 
      'aplicaCierre2', 'aplicaCierre3'
    ],
    {
      1: ['B', 'C', 'D', 'E', 'F'],
      2: ['G', 'H', 'I', 'J'],
      3: ['K', 'L', 'M', 'N', 'O'],
      4: ['P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'],
      5: ['AA', 'AB', 'AC']
    }
  );
  
  // Definición del modelo SOLICITUDES4
  const SOLICITUDES4 = new SpreadsheetModel(
    'SOLICITUDES4',
    [
      'id_solicitud', 'descripcionPrograma', 'identificacionNecesidades', 'atributosBasicos', 
          'atributosDiferenciadores', 'competencia', 'programa', 'programasSimilares', 
          'estrategiasCompetencia', 'personasInteresChecked', 'personasMatriculadasChecked', 'otroInteres', 
          'innovacion', 'solicitudExterno', 'interesSondeo', 'otroMercadeo','llamadas', 'encuestas', 'webinar', 
          'pautas_redes', 'otroEstrategias', 'preregistroFisico', 'preregistroGoogle', 'preregistroOtro',
          'gremios', 'sectores_empresariales', 'politicas_publicas', 'otros_mesas_trabajo', 'focusGroup', 
          'desayunosTrabajo', 'almuerzosTrabajo', 'openHouse', 'ferias_colegios', 'ferias_empresarial', 'otros_mercadeo',
          'valorEconomico', 'modalidadPresencial', 'modalidadVirtual', 'modalidadSemipresencial', 
          'traslados_docente', 'modalidad_asistida_tecnologia', 'beneficiosTangibles', 'beneficiosIntangibles', 
          'particulares', 'colegios', 'empresas', 'egresados', 'colaboradores', 'otros_publicos_potenciales', 
          'tendenciasActuales', 'dofaDebilidades', 'dofaOportunidades', 'dofaFortalezas', 'dofaAmenazas', 
          'paginaWeb', 'facebook', 'instagram', 'linkedin', 'correo', 'prensa', 'boletin', 'llamadas', 
          'otro_canal'
    ],
    {
      1: ['B', 'C'],
      2: ['D', 'E', 'F', 'G', 'H', 'I'],
      3: ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X'],
      4: ['Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO'],
      5: ['AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK']
    }
  );
  
  // Definición del modelo GASTOS
  const GASTOS = new SpreadsheetModel(
    'GASTOS',
    [
      'id_conceptos', 'id_solicitud', 'cantidad', 'valor_unit', 'valor_total', 'concepto_padre'
    ],
    {
      1: ['B', 'C', 'D', 'E', 'F', 'G'],
      2: ['H', 'I', 'J', 'K']
    }
  );

  // Definición del modelo RIESGOS
  const RIESGOS = new SpreadsheetModel(
    'RIESGOS',
    [
      'id_riesgo', 'nombre_riesgo', 'aplica', 'mitigacion', 'id_solicitud', 'categoria'
    ],
    {
      1: ['A', 'B', 'C', 'D', 'E', 'F']
    }
  );
  
  // Definición del modelo ETAPAS
  const ETAPAS = new SpreadsheetModel(
    'ETAPAS',
    [
      'id_solicitud', 'id_usuario', 'fecha', 'name', 'etapa_actual', 'estado', 'nombre_actividad', 'paso', 'estado_formularios'
    ]
  );
  
  // Definición del modelo USUARIOS
  const USUARIOS = new SpreadsheetModel(
    'USUARIOS',
    [
      'id', 'email', 'name'
    ]
  );
  
  module.exports = {
    SpreadsheetModel,
    SOLICITUDES,
    SOLICITUDES2,
    SOLICITUDES3,
    SOLICITUDES4,
    GASTOS,
    RIESGOS,
    ETAPAS,
    USUARIOS,
    // Método auxiliar para obtener todos los modelos como objeto
    getModels() {
      return {
        SOLICITUDES: this.SOLICITUDES,
        SOLICITUDES2: this.SOLICITUDES2,
        SOLICITUDES3: this.SOLICITUDES3,
        SOLICITUDES4: this.SOLICITUDES4,
        GASTOS: this.GASTOS,
        RIESGOS: this.RIESGOS,
        ETAPAS: this.ETAPAS,
        USUARIOS: this.USUARIOS,
      };
    }
  };