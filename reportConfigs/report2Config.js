const dateUtils = require('../utils/dateUtils');

/**
 * Configuración específica para el reporte del Formulario 2 - Presupuesto
 */
const report2Config = {
  title: 'Formulario de Presupuesto - F-05-MP-05-01-02',
  templateId: '1JY-4IfJqEWLqZ_wrq_B_bfIlI9MeVzgF',
  requiresAdditionalData: true, 
  requiresGastos: true,
  
  // Definición de hojas necesarias para este reporte
  sheetDefinitions: {
    SOLICITUDES2: {
      range: 'SOLICITUDES2!A2:CL',
      fields: [
        'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'ingresos_cantidad', 
        'ingresos_vr_unit', 'total_ingresos', 'subtotal_gastos', 'imprevistos_3%',
        'total_gastos_imprevistos', 'fondo_comun_porcentaje', 'facultadad_instituto_porcentaje',
        'escuela_departamento_porcentaje', 'total_recursos', 'subtotal_costos_directos',
        'costos_indirectos_cantidad', 'administracion_cantidad', 'descuentos_cantidad',
        'total_costo_actividad', 'excedente_cantidad', 'valor_inscripcion_individual'
      ]
    },
    SOLICITUDES: {
      range: 'SOLICITUDES!A2:F',
      fields: ['id_solicitud', 'fecha_solicitud', 'nombre_actividad', 'nombre_solicitante']
    },
    GASTOS: {
      range: 'GASTOS!A2:F',
      fields: ['id_conceptos', 'id_solicitud', 'cantidad', 'valor_unit', 'valor_total', 'concepto_padre']
    }
  },
  
  transformData: function(allData) {
    try {
      console.log("Datos recibidos en transformData:", allData);
      
      // Extraer datos de las fuentes
      const solicitudData = allData.SOLICITUDES || {}; 
      const formData = allData.SOLICITUDES2 || {};
      const gastosData = allData.GASTOS || [];
      
      console.log("Datos de SOLICITUDES:", solicitudData);
      console.log("Datos de SOLICITUDES2:", formData);
      console.log("Datos de GASTOS:", gastosData);
      
      // Crear un objeto combinado con todas las fuentes
      const combinedData = {
        ...solicitudData,
        ...formData
      };
      
      // Crear un objeto nuevo para los datos transformados
      const transformedData = {};
      
      // PRE-INICIALIZAR PLACEHOLDERS DE GASTOS
      // Lista completa de IDs de gastos (formato con coma para la plantilla)
      const conceptosGastos = [
        '1', '1,1', '1,2', '1,3', '2', '3', '4', '5', '6', '7', '7,1', '7,2', 
        '7,3', '7,4', '7,5', '8', '8,1', '8,2', '8,3', '8,4', '9', '9,1', '9,2', 
        '9,3', '10', '11', '12', '13', '14', '15'
      ];
      
      // Inicializar todos los placeholders de gastos con valores por defecto
      conceptosGastos.forEach(concepto => {
        transformedData[`gasto_${concepto}_cantidad`] = '0';
        transformedData[`gasto_${concepto}_valor_unit`] = '$0';
        transformedData[`gasto_${concepto}_valor_total`] = '$0';
        transformedData[`gasto_${concepto}_descripcion`] = '';
      });
      
      // Añadir fecha actual formateada para el reporte (como valor por defecto)
      const fechaActual = new Date();
      transformedData['dia'] = fechaActual.getDate().toString().padStart(2, '0');
      transformedData['mes'] = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
      transformedData['anio'] = fechaActual.getFullYear().toString();

      // PROCESAMIENTO UNIFICADO DE FECHA (un solo bloque de código)
      if (combinedData.fecha_solicitud) {
        try {
          // Intentar diferentes formatos de fecha
          let fechaParts;
          
          if (combinedData.fecha_solicitud.includes('/')) {
            // Formato DD/MM/YYYY
            fechaParts = combinedData.fecha_solicitud.split('/');
            transformedData['dia'] = fechaParts[0].padStart(2, '0');
            transformedData['mes'] = fechaParts[1].padStart(2, '0');
            transformedData['anio'] = fechaParts[2];
          } else if (combinedData.fecha_solicitud.includes('-')) {
            // Formato YYYY-MM-DD o DD-MM-YYYY
            fechaParts = combinedData.fecha_solicitud.split('-');
            
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
        } catch (error) {
          console.error('Error al procesar la fecha:', error);
        }
      }

      // Copiar datos base del formulario
      Object.keys(combinedData).forEach(key => {
        if (combinedData[key] !== undefined && combinedData[key] !== null) {
          transformedData[key] = combinedData[key];
        }
      }); 
      
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
      
      // PROCESAMIENTO DE GASTOS - CORREGIDO
      if (Array.isArray(gastosData) && gastosData.length > 0) {
        console.log(`Procesando ${gastosData.length} gastos`);
        
        // Procesar cada gasto y asignarlo a su placeholder correspondiente
        gastosData.forEach(gasto => {
          // Extraer datos del gasto
          const idConcepto = gasto.id_conceptos || '';
          const cantidad = parseFloat(gasto.cantidad) || 0;
          const valorUnit = parseFloat(gasto.valor_unit) || 0;
          const valorTotal = parseFloat(gasto.valor_total) || 0;
          
          // Convertir el ID de concepto (si tiene punto, cambiarlo por coma)
          // por ejemplo: 7.1 -> 7,1 para que coincida con el formato de la plantilla
          const placeholderId = idConcepto.replace(/\./g, ',');
          
          // Asignar valores a los placeholders correspondientes
          transformedData[`gasto_${placeholderId}_cantidad`] = cantidad.toString();
          transformedData[`gasto_${placeholderId}_valor_unit`] = formatCurrency(valorUnit);
          transformedData[`gasto_${placeholderId}_valor_total`] = formatCurrency(valorTotal);
          transformedData[`gasto_${placeholderId}_descripcion`] = gasto.concepto_padre || `Concepto ${idConcepto}`;
          
          console.log(`✅ Asignado gasto con ID ${idConcepto} (${placeholderId}) a placeholders`);
        });
      } else {
        console.log('⚠️ No se encontraron gastos para procesar');
      }
      
      // IMPORTANTE: Eliminar marcadores no reemplazados
      Object.keys(transformedData).forEach(key => {
        const value = transformedData[key];
        if (typeof value === 'string' && (value.includes('{{') || value.includes('}}'))) {
          console.log(`⚠️ Detectado posible marcador en ${key}: "${value}"`);
          transformedData[key] = '';
        }
      });
      
      // Garantizar valores por defecto
      if (!transformedData['fecha_solicitud'] || !transformedData['dia']) {
        console.log('⚠️ Usando fecha por defecto para campos faltantes');
        const fechaActual = new Date();
        
        if (!transformedData['fecha_solicitud']) {
          transformedData['fecha_solicitud'] = `${fechaActual.getDate()}/${fechaActual.getMonth()+1}/${fechaActual.getFullYear()}`;
        }
        
        if (!transformedData['dia']) transformedData['dia'] = fechaActual.getDate().toString().padStart(2, '0');
        if (!transformedData['mes']) transformedData['mes'] = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
        if (!transformedData['anio']) transformedData['anio'] = fechaActual.getFullYear().toString();
      }
      
      console.log("⭐ DATOS TRANSFORMADOS FINALES - FORM 2:", transformedData);
      return transformedData;
    } catch (error) {
      console.error('Error en transformación de datos:', error);
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
    sheetName: 'Formulario2',
    dataRange: 'A1:Z100'
  },
  footerText: 'Universidad del Valle - Extensión y Proyección Social - Presupuesto',
  watermark: false
};

module.exports = report2Config;
