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
      'id_solicitud', 'fecha_solicitud', 'nombre_actividad', 'nombre_solicitante', 'dependencia_tipo',
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
  
  // Definición del modelo SOLICITUDES2
  const SOLICITUDES2 = new SpreadsheetModel(
    'SOLICITUDES2',
    [
        'id_solicitud', 'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos', 
        'costos_personal_cantidad', 'costos_personal_vr_unit', 'total_costos_personal', 
        'personal_universidad_cantidad', 'personal_universidad_vr_unit', 'total_personal_universidad', 
        'honorarios_docentes_cantidad', 'honorarios_docentes_vr_unit', 'total_honorarios_docentes',
        'otro_personal_cantidad', 'otro_personal_vr_unit', 'total_otro_personal', 
        'materiales_sumi_cantidad', 'materiales_sumi_vr_unit', 'total_materiales_sumi',
        'gastos_alojamiento_cantidad', 'gastos_alojamiento_vr_unit', 'total_gastos_alojamiento',
        'gastos_alimentacion_cantidad', 'gastos_alimentacion_vr_unit', 'total_gastos_alimentacion',
        'gastos_transporte_cantidad', 'gastos_transporte_vr_unit', 'total_gastos_transporte',
        'equipos_alquiler_compra_cantidad', 'equipos_alquiler_compra_vr_unit', 'total_equipos_alquiler_compra',
        'dotacion_participantes_cantidad', 'dotacion_participantes_vr_unit', 'total_dotacion_participantes',
        'carpetas_cantidad', 'carpetas_vr_unit', 'total_carpetas',
        'libretas_cantidad', 'libretas_vr_unit', 'total_libretas',
        'lapiceros_cantidad', 'lapiceros_vr_unit', 'total_lapiceros',
        'memorias_cantidad', 'memorias_vr_unit', 'total_memorias',
        'marcadores_papel_otros_cantidad', 'marcadores_papel_otros_vr_unit', 'total_marcadores_papel_otros',
        'impresos_cantidad', 'impresos_vr_unit', 'total_impresos',
        'labels_cantidad', 'labels_vr_unit', 'total_labels',
        'certificados_cantidad', 'certificados_vr_unit', 'total_certificados',
        'escarapelas_cantidad', 'escarapelas_vr_unit', 'total_escarapelas',
        'fotocopias_cantidad', 'fotocopias_vr_unit', 'total_fotocopias',
        'estacion_cafe_cantidad', 'estacion_cafe_vr_unit', 'total_estacion_cafe',
        'transporte_mensaje_cantidad', 'transporte_mensaje_vr_unit', 'total_transporte_mensaje',
        'refrigerios_cantidad', 'refrigerios_vr_unit', 'total_refrigerios',
        'infraestructura_fisica_cantidad', 'infraestructura_fisica_vr_unit', 'total_infraestructura_fisica',
        'gastos_generales_cantidad', 'gastos_generales_vr_unit', 'total_gastos_generales',
        'infraestructura_universitaria_cantidad', 'infraestructura_universitaria_vr_unit', 
        'total_infraestructura_universitaria', 
        'imprevistos_cantidad', 'imprevistos_vr_unit', 'total_imprevistos',
        'costos_administrativos_cantidad', 'costos_administrativos_vr_unit', 'total_costos_administrativos',
        'gastos_extras_cantidad', 'gastos_extras_vr_unit', 'total_gastos_extras',
        'subtotal_gastos',
        'imprevistos_3%',
        'total_gastos_imprevistos',
        'fondo_comun_porcentaje','facultadad_instituto_porcentaje','escuela_departamento_porcentaje', 'total_recursos'
    ],
    {
      1: ['B', 'C'],
      2: ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI','CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS','CT', 'CU', 'CV'],
      3: ['CS', 'CT', 'CU','CV']
    }
  );
  
  // Definición del modelo SOLICITUDES3
  const SOLICITUDES3 = new SpreadsheetModel(
    'SOLICITUDES3',
    [
      'id_solicitud', 'proposito', 'comentario', 'fecha', 'elaboradoPor', 'aplicaDiseno1', 'aplicaDiseno2', 
      'aplicaDiseno3', 'aplicaLocacion1', 'aplicaLocacion2', 'aplicaLocacion3', 'aplicaDesarrollo1', 
      'aplicaDesarrollo2', 'aplicaDesarrollo3', 'aplicaDesarrollo4', 'aplicaCierre1', 'aplicaCierre2', 
      'aplicaOtros1', 'aplicaOtros2'  
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
          'estrategiasCompetencia', 'personasInteres', 'personasMatriculadas', 'otroInteres', 
          'innovacion', 'solicitudExterno', 'interesSondeo', 'llamadas', 'encuestas', 'webinar', 
          'preregistro', 'mesasTrabajo', 'focusGroup', 'desayunosTrabajo', 'almuerzosTrabajo', 'openHouse', 
          'valorEconomico', 'modalidadPresencial', 'modalidadVirtual', 'modalidadSemipresencial', 
          'otraModalidad', 'beneficiosTangibles', 'beneficiosIntangibles', 'particulares', 'colegios', 
          'empresas', 'egresados', 'colaboradores', 'otros_publicos_potenciales', 'tendenciasActuales', 
          'dofaDebilidades', 'dofaOportunidades', 'dofaFortalezas', 'dofaAmenazas', 'paginaWeb', 
          'facebook', 'instagram', 'linkedin', 'correo', 'prensa', 'boletin', 'llamadas_redes', 'otro_canal'
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
        ETAPAS: this.ETAPAS,
        USUARIOS: this.USUARIOS,
      };
    }
  };