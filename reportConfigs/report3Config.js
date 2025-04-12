const dateUtils = require('../utils/dateUtils');

/**
 * Configuración específica para el reporte del Formulario 3 - Matriz de Riesgos
 */
const report3Config = {
  title: 'F-08-MP-05-01-01 - Riesgos Potenciales',
  templateId: '1WoPUZYusNl2u3FpmZ1qiO5URBUqHIwKF', // Reemplazar con el ID correcto de la plantilla
  requiresAdditionalData: false,
  requiresGastos: false,
  
  // Definición de hojas necesarias para este reporte
  sheetDefinitions: {
    SOLICITUDES: {
      range: 'SOLICITUDES!A2:AU',
      fields: [
        'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'nombre_solicitante'
      ]
    },
    SOLICITUDES3: {
      range: 'SOLICITUDES3!A2:AC',
      fields: [
        'id_solicitud', 'proposito', 'comentario', 'programa', 'fecha_solicitud', 'nombre_solicitante',
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
      ]
    }
  },
  
  transformData: function(allData) {
    try {
      console.log("🔎 DATOS RECIBIDOS EN TRANSFORMDATA PARA FORMULARIO 3:");
      console.log("- DATOS COMPLETOS:", JSON.stringify(allData, null, 2));
      
      // Extraer datos de las fuentes
      const solicitudData = allData.SOLICITUDES || {}; 
      const formData = allData.SOLICITUDES3 || {};
      
      console.log("- SOLICITUDES:", JSON.stringify(solicitudData, null, 2));
      console.log("- SOLICITUDES3:", JSON.stringify(formData, null, 2));
      console.log("- Campo 'programa':", formData.programa);
      console.log("- Campos diseño:", {
        aplicaDiseno1: formData.aplicaDiseno1,
        aplicaDiseno2: formData.aplicaDiseno2,
        aplicaDiseno3: formData.aplicaDiseno3,
        aplicaDiseno4: formData.aplicaDiseno4
      });
      
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
      
      // Convertir campos de riesgo a "Sí aplica" o "No aplica"
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
        // Valores considerados afirmativos
        if (
          valor === true || 
          valor === 'true' || 
          valor === 'TRUE' || 
          valor === 'Sí' || 
          valor === 'Si' || 
          valor === 'SI' || 
          valor === 'si' || 
          valor === 'sí' || 
          valor === 'SÍ' ||
          valor === 'S' ||
          valor === 's' ||
          valor === 'Y' || 
          valor === 'y' ||
          valor === 'Yes' ||
          valor === 'yes'
        ) {
          transformedData[campo] = 'Sí aplica';
        } else {
          // Si no es afirmativo o está vacío, considerarlo como "No aplica"
          transformedData[campo] = 'No aplica';
        }
        
        // Imprimir para depuración
        console.log(`Campo ${campo}: valor original = "${valor}", transformado = "${transformedData[campo]}"`);
      });
      
      // Valores por defecto para campos críticos
      if (!transformedData.proposito) transformedData.proposito = 'No especificado';
      if (!transformedData.comentario) transformedData.comentario = 'No especificado';
      if (!transformedData.programa) transformedData.programa = 'No especificado';
      
      console.log("⭐ DATOS TRANSFORMADOS FINALES - FORM 3:", transformedData);
      return transformedData;
    } catch (error) {
      console.error('Error en transformación de datos para formulario 3:', error);
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
    sheetName: 'Formulario3',
    dataRange: 'A1:Z100'
  },
  footerText: 'Universidad del Valle - Extensión y Proyección Social - Matriz de Riesgos',
  watermark: false
};

module.exports = report3Config;