const dateUtils = require('../utils/dateUtils');

/**
 * Configuración específica para el reporte del Formulario 2 - Presupuesto
 */
const report2Config = {
  title: 'Formulario de Presupuesto - F-05-MP-05-01-02',
  templateId: '1JY-4IfJqEWLqZ_wrq_B_bfIlI9MeVzgF', // Replace with actual template ID
  requiresAdditionalData: false,
  requiresGastos: true, // Budget form likely needs expense data
  
  // Definición de hojas necesarias para este reporte
  sheetDefinitions: {
    SOLICITUDES2: {
      range: 'SOLICITUDES2!A2:Z',
      fields: [
        'id_solicitud', 'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos',
        'subtotal_gastos', 'imprevistos_3', 'imprevistos_3%', 'total_gastos_imprevistos',
        'fondo_comun_porcentaje', 'facultadad_instituto_porcentaje',
        'escuela_departamento_porcentaje', 'total_recursos',
        'subtotal_costos_directos', 'costos_indirectos_cantidad', 
        'administracion_cantidad', 'descuentos_cantidad',
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

      // PROCESAMIENTO DE FECHA
      if (combinedData.fecha_solicitud) {
        try {
          // Formatear usando la utilidad de fechas
          const dateParts = dateUtils.formatDateParts(combinedData.fecha_solicitud);
          transformedData.dia = dateParts.dia;
          transformedData.mes = dateParts.mes;
          transformedData.anio = dateParts.anio;
          transformedData.fecha_solicitud = `${dateParts.dia}/${dateParts.mes}/${dateParts.anio}`;
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
      
      // Formatear valores monetarios específicos
      ['ingresos_vr_unit', 'total_ingresos', 'subtotal_gastos', 
       'imprevistos_3', 'total_gastos_imprevistos', 'total_recursos',
       'costos_indirectos_cantidad', 'administracion_cantidad', 
       'descuentos_cantidad', 'total_costo_actividad', 'excedente_cantidad',
       'valor_inscripcion_individual'].forEach(field => {
        if (transformedData[field]) {
          transformedData[field + '_formatted'] = formatCurrency(transformedData[field]);
        }
      });
      
      // PROCESAMIENTO DE GASTOS
      // Filtrar gastos para esta solicitud
      if (gastosData && gastosData.length > 0) {
        const gastosFiltrados = gastosData.filter(g => g.id_solicitud === combinedData.id_solicitud);
        console.log(`Procesando ${gastosFiltrados.length} gastos para solicitud ${combinedData.id_solicitud}`);
        
        gastosFiltrados.forEach(gasto => {
          const idConcepto = gasto.id_conceptos;
          const idConComa = idConcepto.replace(/\./g, ','); // Convertir '1.1' a '1,1'
          const placeholderId = `gasto_${idConComa}`;
          const cantidad = gasto.cantidad || 0;
          const valorUnit = gasto.valor_unit || 0;
          const valorTotal = gasto.valor_total || 0;
          
          // Asignar valores a sus respectivos placeholders
          transformedData[`${placeholderId}_cantidad`] = cantidad.toString();
          transformedData[`${placeholderId}_valor_unit`] = valorUnit.toString();
          transformedData[`${placeholderId}_valor_unit_formatted`] = formatCurrency(valorUnit);
          transformedData[`${placeholderId}_valor_total`] = valorTotal.toString();
          transformedData[`${placeholderId}_valor_total_formatted`] = formatCurrency(valorTotal);
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
