const dateUtils = require('../utils/dateUtils');
const { generateRiskRows } = require('../services/dynamicRows'); // Asegúrate que esta ruta sea correcta

/**
 * Configuración específica para el reporte del Formulario 3 - Matriz de Riesgos
 */
const report3Config = {
  title: 'F-08-MP-05-01-01 - Riesgos Potenciales',
  templateId: '1Tq-V2BSoe17-xjOeWeqaq4Hm6bU1dLx0TG-bcIhnS_4', // Reemplazar con el ID correcto de la plantilla
  requiresAdditionalData: true,
  requiresGastos: false,
  requiresRiesgos: true, // Nuevo indicador para procesar riesgos

  // Definición de hojas necesarias para este reporte
  sheetDefinitions: { // <-- Solo un nivel de sheetDefinitions
    SOLICITUDES: {
      range: 'SOLICITUDES!A2:AU', // Puedes ajustar AU si necesitas más o menos columnas de SOLICITUDES
      fields: [
        'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'nombre_solicitante' // Ajusta los campos que realmente necesitas de SOLICITUDES
        // ... otros campos si son necesarios
      ]
    },
    SOLICITUDES3: { // <-- Debe estar al mismo nivel que SOLICITUDES y RIESGOS
      range: 'SOLICITUDES3!A2:AC', // Asegúrate que AC es la última columna necesaria
      fields: [
        'id_solicitud', 'proposito', 'comentario', 'programa', 'fecha_solicitud', 'nombre_solicitante',
        'aplicaDiseno1', 'aplicaDiseno2', 'aplicaDiseno3', 'aplicaDiseno4',
        'aplicaLocacion1', 'aplicaLocacion2', 'aplicaLocacion3', 'aplicaLocacion4', 'aplicaLocacion5',
        'aplicaDesarrollo1', 'aplicaDesarrollo2', 'aplicaDesarrollo3', 'aplicaDesarrollo4',
        'aplicaDesarrollo5', 'aplicaDesarrollo6', 'aplicaDesarrollo7', 'aplicaDesarrollo8',
        'aplicaDesarrollo9', 'aplicaDesarrollo10', 'aplicaDesarrollo11',
        'aplicaCierre1', 'aplicaCierre2', 'aplicaCierre3'
      ]
    },
    RIESGOS: {
      range: 'RIESGOS!A2:F', // Asegúrate que F es la última columna necesaria
      fields: [
        'id_riesgo', 'nombre_riesgo', 'aplica', 'mitigacion', 'id_solicitud', 'categoria'
      ]
    }
  },

  transformData: function(allData) {
    try {
      console.log("🔎 DATOS RECIBIDOS EN TRANSFORMDATA PARA FORMULARIO 3:");
      console.log("- DATOS COMPLETOS:", Object.keys(allData));

      // Extraer datos de las fuentes
      const solicitudData = allData.SOLICITUDES || {};
      const formData = allData.SOLICITUDES3 || {}; // Ahora debería funcionar

      // Obtener datos de riesgos si están disponibles
      const riesgosData = allData.riesgos || [];
      const riesgosPorCategoria = allData.riesgosPorCategoria || {};

      console.log("- SOLICITUDES:", JSON.stringify(solicitudData, null, 2));
      console.log("- SOLICITUDES3:", JSON.stringify(formData, null, 2)); // Verifica que ahora tenga datos
      console.log("- RIESGOS:", riesgosData.length);
      console.log("- RIESGOS POR CATEGORÍA:", Object.keys(riesgosPorCategoria));

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
        // Priorizar fecha de SOLICITUDES3 si existe, sino usar la de SOLICITUDES
        const fechaFuente = combinedData.fecha_solicitud; // combinedData ya prioriza SOLICITUDES3 sobre SOLICITUDES si ambos existen
        if (fechaFuente) {
          // --- FIX: Use the correct function name ---
          const dateParts = dateUtils.getDateParts(fechaFuente); 
          // --- End FIX ---
          transformedData.dia = dateParts.dia;
          transformedData.mes = dateParts.mes;
          transformedData.anio = dateParts.anio;
          // Optionally keep the original formatted date if needed elsewhere, 
          // but ensure dia, mes, anio are prioritized for placeholders
          // transformedData.fecha_solicitud = `${dateParts.dia}/${dateParts.mes}/${dateParts.anio}`; 
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
          (typeof valor === 'string' && ['true', 'sí', 'si', 's', 'yes', 'y'].includes(valor.toLowerCase()))
        ) {
          transformedData[campo] = 'Sí aplica';
        } else {
          // Si no es afirmativo o está vacío, considerarlo como "No aplica"
          transformedData[campo] = 'No aplica';
        }
      });

      // Valores por defecto para campos críticos
      if (!transformedData.proposito) transformedData.proposito = 'No especificado';
      if (!transformedData.comentario) transformedData.comentario = 'No especificado';
      if (!transformedData.programa) transformedData.programa = 'No especificado';

      // Procesar RIESGOS DINÁMICOS para cada categoría
      if (riesgosPorCategoria && Object.keys(riesgosPorCategoria).length > 0) {
        console.log("📊 Procesando riesgos dinámicos por categoría");

        // Definición de las posiciones de inserción para cada categoría
        const posiciones = {
          diseno: 'B18:H18',
          locacion: 'B24:H24',
          desarrollo: 'B35:H35', // Ajusta si la plantilla cambió
          cierre: 'B38:H38',   // Ajusta si la plantilla cambió
          otros: 'B41:H41'    // Ajusta si la plantilla cambió
        };

        // Generar filas dinámicas para cada categoría
        Object.keys(riesgosPorCategoria).forEach(categoria => {
          const riesgosCat = riesgosPorCategoria[categoria] || [];
          if (riesgosCat.length > 0) {
            const insertarEn = posiciones[categoria] || posiciones.otros; // Usar 'otros' como fallback

            console.log(`🔄 Generando filas dinámicas para ${riesgosCat.length} riesgos de categoría ${categoria} en ${insertarEn}`);

            // Usar generateRiskRows para esta categoría específica
            // Asegúrate que generateRiskRows exista y funcione como se espera
            const dynamicRowsData = generateRiskRows(riesgosCat, categoria, insertarEn);

            if (dynamicRowsData) {
              // Usar un nombre de campo único para cada categoría
              const fieldName = `__FILAS_DINAMICAS_${categoria.toUpperCase()}__`;
              transformedData[fieldName] = dynamicRowsData;

              console.log(`✅ Configuración de filas dinámicas para ${categoria} completada: ${dynamicRowsData.rows.length} filas`);
            } else {
              console.log(`❌ Error al generar filas dinámicas para categoría ${categoria}`);
            }
          } else {
            console.log(`⚠️ No hay riesgos para la categoría ${categoria}`);
          }
        });
      } else if (riesgosData && riesgosData.length > 0) {
         console.log("📊 Procesando riesgos sin categorización (se asignarán a 'otros')");

        // Si tenemos riesgos pero no están categorizados, los ponemos todos como 'otros'
        const dynamicRowsData = generateRiskRows(riesgosData, 'otros', 'B41:H41'); // Ajusta el rango si es necesario

        if (dynamicRowsData) {
          transformedData['__FILAS_DINAMICAS_OTROS__'] = dynamicRowsData;
          console.log(`✅ Configuración de filas dinámicas para riesgos sin categoría completada: ${dynamicRowsData.rows.length} filas`);
        }
      } else {
        console.log("⚠️ No se encontraron riesgos para generar filas dinámicas");
      }


      console.log("⭐ DATOS TRANSFORMADOS FINALES - FORM 3:", Object.keys(transformedData));
      return transformedData;
    } catch (error) {
      console.error('Error en transformación de datos para formulario 3:', error);
      // Devolver un objeto con valores por defecto y el error
      return {
        error: true,
        message: error.message,
        dia: new Date().getDate().toString().padStart(2, '0'),
        mes: (new Date().getMonth() + 1).toString().padStart(2, '0'),
        anio: new Date().getFullYear().toString(),
        fecha_solicitud: new Date().toLocaleDateString('es-CO') // O un formato predeterminado
        // ... otros campos con valores por defecto si es necesario
      };
    }
  },


  // Configuración adicional específica para Google Sheets
  sheetsConfig: {
    sheetName: 'Formulario3', // Nombre de la hoja donde podrían estar estos datos (si aplica)
    dataRange: 'A1:AC100' // Rango amplio por si acaso
  },
  footerText: 'Universidad del Valle - Extensión y Proyección Social - Matriz de Riesgos',
  watermark: false
};

module.exports = report3Config;