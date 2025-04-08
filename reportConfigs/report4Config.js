const dateUtils = require('../utils/dateUtils');

/**
 * Configuración específica para el reporte del Formulario 4
 */
const report4Config = {
  title: 'Formulario 4 - [Nombre del formulario]',
  templateId: '1FTC7Vq3O4ultexRPXYrJKOpL9G0071-0', // Use the correct template ID from driveService.js
  requiresAdditionalData: false,
  requiresGastos: false,
  
  // Definición de hojas necesarias para este reporte
  sheetDefinitions: {
    SOLICITUDES: {
      range: 'SOLICITUDES!A2:AU',
      fields: [
        'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'nombre_solicitante'
        // Add other fields you need from the SOLICITUDES sheet
      ]
    },
    SOLICITUDES4: {
      range: 'SOLICITUDES4!A2:AC',
      fields: [
        'id_solicitud'
        // Add other fields specific to form 4
      ]
    }
  },
  
  transformData: function(allData) {
    try {
      console.log("Datos recibidos en transformData para formulario 4:", allData);
      
      // Extraer datos de las fuentes
      const solicitudData = allData.SOLICITUDES || {}; 
      const formData = allData.SOLICITUDES4 || {};
      
      // Crear un objeto combinado con todas las fuentes
      const combinedData = {
        ...solicitudData,
        ...formData
      };
      
      // Crear un objeto nuevo para los datos transformados
      const transformedData = {};
      
      // Copiar datos base del formulario
      Object.keys(combinedData).forEach(key => {
        if (combinedData[key] !== undefined && combinedData[key] !== null) {
          transformedData[key] = combinedData[key];
        }
      });
      
      // Formatear fecha
      try {
        if (combinedData.fecha_solicitud) {
          const dateParts = dateUtils.formatDateParts(combinedData.fecha_solicitud);
          transformedData.dia = dateParts.dia;
          transformedData.mes = dateParts.mes;
          transformedData.anio = dateParts.anio;
          transformedData.fecha_solicitud = `${dateParts.dia}/${dateParts.mes}/${dateParts.anio}`;
        } else {
          // Valores por defecto
          const fechaActual = new Date();
          transformedData.dia = fechaActual.getDate().toString().padStart(2, '0');
          transformedData.mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
          transformedData.anio = fechaActual.getFullYear().toString();
        }
      } catch (error) {
        console.error('Error al procesar la fecha:', error);
        const fechaActual = new Date();
        transformedData.dia = fechaActual.getDate().toString().padStart(2, '0');
        transformedData.mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
        transformedData.anio = fechaActual.getFullYear().toString();
      }
      
      // Agregar transformaciones específicas para form 4 aquí
      
      // Valores por defecto para campos críticos
      // Add default values for critical fields if needed
      
      console.log("⭐ DATOS TRANSFORMADOS FINALES - FORM 4:", transformedData);
      return transformedData;
    } catch (error) {
      console.error('Error en transformación de datos para formulario 4:', error);
      return {
        error: true,
        message: error.message,
        dia: new Date().getDate().toString().padStart(2, '0'),
        mes: (new Date().getMonth() + 1).toString().padStart(2, '0'),
        anio: new Date().getFullYear().toString(),
        fecha_solicitud: new Date().toLocaleDateString('es-CO')
      };
    }
  },
  
  // Configuración adicional específica para Google Sheets
  sheetsConfig: {
    sheetName: 'Formulario4',
    dataRange: 'A1:Z100'
  },
  footerText: 'Universidad del Valle - Extensión y Proyección Social',
  watermark: false
};

module.exports = report4Config;