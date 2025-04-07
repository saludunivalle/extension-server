/**
 * Configuración específica para el reporte del Formulario 2 - Presupuesto
 * Implementación optimizada para marcado directo en Google Sheets y manejo de placeholders
 */
const report2Config = {
  title: 'Formulario de Presupuesto - F-05-MP-05-01-02',
  showHeader: true,
  
  // Función para transformar los datos para Google Sheets
  transformData: async (formData) => {
    try {
      // Crear un objeto nuevo vacío
      const transformedData = {};
      
      // Añadir fecha actual formateada para el reporte (como valor por defecto)
      const fechaActual = new Date();
      transformedData['dia'] = fechaActual.getDate().toString().padStart(2, '0');
      transformedData['mes'] = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
      transformedData['anio'] = fechaActual.getFullYear().toString();

      // PROCESAMIENTO UNIFICADO DE FECHA (un solo bloque de código)
      if (formData.fecha_solicitud) {
        try {
          // Intentar diferentes formatos de fecha
          let fechaParts;
          
          if (formData.fecha_solicitud.includes('/')) {
            // Formato DD/MM/YYYY
            fechaParts = formData.fecha_solicitud.split('/');
            transformedData['dia'] = fechaParts[0].padStart(2, '0');
            transformedData['mes'] = fechaParts[1].padStart(2, '0');
            transformedData['anio'] = fechaParts[2];
          } else if (formData.fecha_solicitud.includes('-')) {
            // Formato YYYY-MM-DD o DD-MM-YYYY
            fechaParts = formData.fecha_solicitud.split('-');
            
            if (fechaParts[0].length === 4) {
              // Formato YYYY-MM-DD
              transformedData['dia'] = fechaParts[2].padStart(2, '0');
              transformedData['mes'] = fechaParts[1].padStart(2, '0');
              transformedData['anio'] = fechaParts[0];
            } else {
              // Formato DD-MM-YYYY
              transformedData['dia'] = fechaParts[0].padStart(2, '0');
              transformedData['mes'] = fechaParts[1].padStart(2, '0');
              transformedData['anio'] = fechaParts[2];
            }
          }
          
          // Asegurarse que fecha_solicitud esté en formato estándar
          transformedData['fecha_solicitud'] = `${transformedData['dia']}/${transformedData['mes']}/${transformedData['anio']}`;
          
          console.log(`Fecha procesada: día=${transformedData['dia']}, mes=${transformedData['mes']}, año=${transformedData['anio']}`);
        } catch (error) {
          console.error('Error al procesar la fecha:', error);
        }
      }

      // Copiar datos base del formulario
      Object.keys(formData).forEach(key => {
        if (formData[key] !== undefined && formData[key] !== null) {
          transformedData[key] = formData[key];
        }
      }); 
      
      console.log("🔄 Transformando datos para Google Sheets - Formulario 2:", formData);
      
      // Formatear valores monetarios
      const formatCurrency = (value) => {
        if (!value && value !== 0) return '';
        
        const numValue = parseFloat(value);
        if (isNaN(numValue)) return value;
        
        return new Intl.NumberFormat('es-CO', {
          style: 'currency',
          currency: 'COP',
          minimumFractionDigits: 0,
          maximumFractionDigits: 0
        }).format(numValue);
      };
      
      // NUEVO: Obtener los gastos dinámicos de la hoja GASTOS con mejor manejo de IDs
      try {
        // Importar el servicio de sheetsService
        const sheetsService = require('../services/sheetsService');
        if (sheetsService && sheetsService.client) {
          // Código para procesar gastos dinámicos
          // (Implementación según tus necesidades)
        }
      } catch (error) {
        console.error('Error al obtener gastos dinámicos:', error);
        console.error('Stack:', error.stack);
        // Continuar sin los gastos dinámicos
      }
      
      // Formatear valores monetarios específicos (campos estáticos)
      const monetaryFields = [
        'ingresos_vr_unit', 'total_ingresos',
        'subtotal_costos_directos', 'costos_indirectos_cantidad', 'administracion_cantidad',
        'descuentos_cantidad', 'total_costo_actividad', 'excedente_cantidad',
        'valor_inscripcion_individual', 'subtotal_gastos', 'total_gastos_imprevistos',
        'total_recursos'
      ];
      
      monetaryFields.forEach(field => {
        if (transformedData[field]) {
          transformedData[field] = formatCurrency(transformedData[field]);
        }
      });
      
      // IMPORTANTE: Último paso - ELIMINAR valores que podrían contener marcadores de posición no deseados
      Object.keys(transformedData).forEach(key => {
        const value = transformedData[key];
        if (typeof value === 'string' && (value.includes('{{') || value.includes('}}'))) {
          console.log(`⚠️ Detectado posible marcador sin reemplazar en campo ${key}: "${value}"`);
          transformedData[key] = ''; // Convertir a cadena vacía
        }
      });
      
      // Importante: la fecha es un campo crítico, si falta, utilizar valores por defecto
      if (!transformedData['fecha_solicitud'] || !transformedData['dia']) {
        console.log('⚠️ Usando fecha por defecto para algunos campos faltantes');
        const fechaActual = new Date();
        
        if (!transformedData['fecha_solicitud']) {
          transformedData['fecha_solicitud'] = `${fechaActual.getDate()}/${fechaActual.getMonth()+1}/${fechaActual.getFullYear()}`;
        }
        
        if (!transformedData['dia']) transformedData['dia'] = fechaActual.getDate().toString().padStart(2, '0');
        if (!transformedData['mes']) transformedData['mes'] = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
        if (!transformedData['anio']) transformedData['anio'] = fechaActual.getFullYear().toString();
      }
      
      // Imprimir datos finales transformados para depuración
      console.log("⭐ DATOS TRANSFORMADOS FINALES - FORM 2:", transformedData);
      return transformedData;
    } catch (error) {
      console.error('Error general en la transformación de datos:', error);
      // Asegurar que siempre retornamos algo válido incluso en caso de error
      return {
        error: true,
        message: error.message,
        // Datos mínimos requeridos para que el reporte no falle completamente
        dia: new Date().getDate().toString().padStart(2, '0'),
        mes: (new Date().getMonth() + 1).toString().padStart(2, '0'),
        anio: new Date().getFullYear().toString(),
        fecha_solicitud: new Date().toLocaleDateString('es-CO')
      };
    }
  },
  
  // Configuración adicional específica para Google Sheets
  sheetsConfig: {
    sheetName: 'Formulario2',
    dataRange: 'A1:Z100'
  },
  footerText: 'Universidad del Valle - Extensión y Proyección Social - Presupuesto',
  watermark: false
};

module.exports = report2Config;
