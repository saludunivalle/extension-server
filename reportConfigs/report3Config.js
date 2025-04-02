const dateUtils = require('../utils/dateUtils');

/**
 * Configuración específica para el reporte del Formulario 3 - Matriz de Riesgos
 * Implementación siguiendo el patrón declarativo que funciona en report1 y report2
 */
const report3Config = {
  title: 'F-08-MP-05-01-01 - Riesgos Potenciales',
  templateId: '1WoPUZYusNl2u3FpmZ1qiO5URBUqHIwKF', 
  showHeader: true,
  
  // Definición de hojas necesarias para este reporte
  sheetDefinitions: {
    SOLICITUDES3: {
      range: 'SOLICITUDES3!A2:CL',
      fields: [
        'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'nombre_solicitante',
        'proposito', 'comentario', 'programa',
        'aplicaDiseno1', 'aplicaDiseno2', 'aplicaDiseno3', 'aplicaDiseno4',
        'aplicaLocacion1', 'aplicaLocacion2', 'aplicaLocacion3', 'aplicaLocacion4', 'aplicaLocacion5',
        'aplicaDesarrollo1', 'aplicaDesarrollo2', 'aplicaDesarrollo3', 'aplicaDesarrollo4', 
        'aplicaDesarrollo5', 'aplicaDesarrollo6', 'aplicaDesarrollo7', 'aplicaDesarrollo8', 
        'aplicaDesarrollo9', 'aplicaDesarrollo10', 'aplicaDesarrollo11',
        'aplicaCierre1', 'aplicaCierre2', 'aplicaCierre3'
      ]
    }
  },
  
  /**
   * Transforma los datos para el reporte 3 de Matriz de Riesgos
   * @param {Object} allData - Datos de la solicitud
   * @returns {Object} - Datos transformados para la plantilla
   */
  transformData: function(allData) {
    console.log("🔄 Iniciando transformación para reporte 3 - Matriz de Riesgos");
    
    // Extraer datos de la solicitud
    const solicitudData = allData.SOLICITUDES3 || {};
    
    // PASO 1: Inicializar objeto de resultado con TODOS los campos posibles
    let transformedData = {};
    
    // Lista de campos básicos que podrían estar en la plantilla
    const allBasicFields = [
      // Datos de identificación
      'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'nombre_solicitante',
      'dia', 'mes', 'anio', 'programa',
      
      // Propósito y comentario
      'proposito', 'comentario',
      
      // Campos para matriz de riesgos - Diseño
      'aplicaDiseno1', 'aplicaDiseno2', 'aplicaDiseno3', 'aplicaDiseno4',
      
      // Campos para matriz de riesgos - Locaciones
      'aplicaLocacion1', 'aplicaLocacion2', 'aplicaLocacion3', 'aplicaLocacion4', 'aplicaLocacion5',
      
      // Campos para matriz de riesgos - Desarrollo
      'aplicaDesarrollo1', 'aplicaDesarrollo2', 'aplicaDesarrollo3', 'aplicaDesarrollo4', 
      'aplicaDesarrollo5', 'aplicaDesarrollo6', 'aplicaDesarrollo7', 'aplicaDesarrollo8', 
      'aplicaDesarrollo9', 'aplicaDesarrollo10', 'aplicaDesarrollo11',
      
      // Campos para matriz de riesgos - Cierre
      'aplicaCierre1', 'aplicaCierre2', 'aplicaCierre3'
    ];
    
    // Inicializar todos los campos con cadena vacía
    allBasicFields.forEach(field => {
      transformedData[field] = '';
    });
    
    // PASO 2: Copiar datos de la solicitud
    Object.keys(solicitudData).forEach(key => {
      if (solicitudData[key] !== undefined && solicitudData[key] !== null) {
        transformedData[key] = solicitudData[key];
      }
    });
    
    // PASO 3: Procesar fecha
    if (solicitudData.fecha_solicitud) {
      try {
        const { dia, mes, anio } = dateUtils.formatDateParts(solicitudData.fecha_solicitud);
        transformedData.dia = dia;
        transformedData.mes = mes;
        transformedData.anio = anio;
      } catch (error) {
        console.error("Error al procesar fecha:", error);
        // Usar fecha actual como respaldo
        const today = new Date();
        transformedData.dia = today.getDate().toString().padStart(2, '0');
        transformedData.mes = (today.getMonth() + 1).toString().padStart(2, '0');
        transformedData.anio = today.getFullYear().toString();
      }
    }
    
    // PASO 4: Convertir campos de riesgo de formato Sí/No a "Sí aplica"/"No aplica"
    const camposRiesgo = [
      // Diseño
      'aplicaDiseno1', 'aplicaDiseno2', 'aplicaDiseno3', 'aplicaDiseno4',
      // Locaciones
      'aplicaLocacion1', 'aplicaLocacion2', 'aplicaLocacion3', 'aplicaLocacion4', 'aplicaLocacion5',
      // Desarrollo
      'aplicaDesarrollo1', 'aplicaDesarrollo2', 'aplicaDesarrollo3', 'aplicaDesarrollo4', 
      'aplicaDesarrollo5', 'aplicaDesarrollo6', 'aplicaDesarrollo7', 'aplicaDesarrollo8', 
      'aplicaDesarrollo9', 'aplicaDesarrollo10', 'aplicaDesarrollo11',
      // Cierre
      'aplicaCierre1', 'aplicaCierre2', 'aplicaCierre3'
    ];
    
    camposRiesgo.forEach(campo => {
      const valor = transformedData[campo];
      // Normalizar el valor (puede venir como string "Sí"/"No" o como booleano)
      if (valor === true || valor === 'true' || valor === 'Sí' || valor === 'Si' || valor === 'si' || valor === 'sí') {
        transformedData[campo] = 'Sí aplica';
      } else {
        transformedData[campo] = 'No aplica';
      }
    });
    
    // PASO 5: Datos de respaldo para campos críticos
    if (!transformedData.proposito) transformedData.proposito = 'No especificado';
    if (!transformedData.comentario) transformedData.comentario = 'No especificado';
    if (!transformedData.programa) transformedData.programa = 'No especificado';
    
    // PASO 6: Limpieza final - eliminar marcadores sin reemplazar
    Object.keys(transformedData).forEach(key => {
      const value = transformedData[key];
      if (typeof value === 'string' && (
        value.includes('{{') || 
        value.includes('}}') || 
        value === 'undefined' || 
        value === 'null'
      )) {
        transformedData[key] = '';
      }
    });
    
    console.log("✅ Transformación completada para reporte 3");
    
    return transformedData;
  },
  
  // Configuración para Google Sheets
  sheetsConfig: {
    sheetName: 'Formulario3',
    dataRange: 'A1:Z100'
  },
  
  footerText: 'Universidad del Valle - Extensión y Proyección Social - Matriz de Riesgos',
  watermark: false
};

module.exports = report3Config;