const dateUtils = require('../utils/dateUtils');

/**
 * Configuración específica para el reporte del Formulario 2 - Presupuesto
 */
const report2Config = {
  title: 'Formulario de Presupuesto - F-05-MP-05-01-02',
  templateId: '1JY-4IfJqEWLqZ_wrq_B_bfIlI9MeVzgF',
  requiresAdditionalData: true,
  requiresGastos: true, // Budget form needs expense data
  
  // Definición de hojas necesarias para este reporte
  sheetDefinitions: {
    SOLICITUDES2: {
      range: 'SOLICITUDES2!A2:M',
      fields: [
        'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos',
        'subtotal_gastos', 'imprevistos_3', 'imprevistos_3%', 'total_gastos_imprevistos',
        'fondo_comun_porcentaje', 'fondo_comun', 'facultad_instituto_porcentaje', 'facultad_instituto',
        'escuela_departamento_porcentaje', 'escuela_departamento', 'total_recursos',
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
    // Get data from both sheets
    const solicitudData = allData.SOLICITUDES || {};
    const formData = allData.SOLICITUDES2 || {};
    
    // Combine the data
    const combinedData = {
      ...solicitudData,
      ...formData
    };
    
    try {
      console.log("Datos recibidos en transformData:", allData);
      
      // Extraer datos de las fuentes
      const solicitudData = allData.SOLICITUDES || {}; 
      const formData = allData.SOLICITUDES2 || {};
      
      // Fix: Ensure we extract gastos data from all possible sources
      const gastosData = Array.isArray(allData.GASTOS) ? allData.GASTOS : [];
      const gastosFromAdditional = allData.gastosNormales || [];
      const gastosDinamicos = allData.gastosDinamicos || [];
      
      // FIELD TYPE VALIDATION FUNCTIONS
      const isDateLike = (value) => {
        if (typeof value !== 'string') return false;
        // Check for common date patterns: yyyy-mm-dd, dd/mm/yyyy, etc.
        return /\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(value);
      };
      
      const isNumeric = (value) => {
        if (value === undefined || value === null) return false;
        return !isNaN(parseFloat(value)) && isFinite(value.toString().replace(/,/g, ''));
      };
      
      // Log diagnostic data for debugging
      console.log("🔍 DIAGNÓSTICO DE DATOS:");
      console.log("- SOLICITUDES.nombre_actividad:", solicitudData.nombre_actividad);
      console.log("- SOLICITUDES.fecha_solicitud:", solicitudData.fecha_solicitud);
      console.log("- SOLICITUDES2.nombre_actividad:", formData.nombre_actividad);
      console.log("- SOLICITUDES2.fecha_solicitud:", formData.fecha_solicitud);
      console.log("- SOLICITUDES2.ingresos_cantidad:", formData.ingresos_cantidad);
      console.log("- SOLICITUDES2.ingresos_vr_unit:", formData.ingresos_vr_unit);
      console.log("- SOLICITUDES2.total_ingresos:", formData.total_ingresos);
      
      // Create a copy to avoid modifying the original data
      const formDataCorregido = { ...formData };
      
      // CASE 1: Date in quantity field - FIX THE MAIN ISSUE
      if (isDateLike(formDataCorregido.ingresos_cantidad) && !isDateLike(formDataCorregido.fecha_solicitud)) {
        console.log("⚠️ CORRECCIÓN: Fecha detectada en campo ingresos_cantidad");
        
        // Move date to correct field
        formDataCorregido.fecha_solicitud = formDataCorregido.ingresos_cantidad;
        
        // Check for quantity value in other fields
        if (isNumeric(formDataCorregido.ingresos_vr_unit)) {
          formDataCorregido.ingresos_cantidad = formDataCorregido.ingresos_vr_unit;
          
          // Try to find unit price in another field
          if (isNumeric(formDataCorregido.total_ingresos)) {
            const cantidad = parseFloat(formDataCorregido.ingresos_cantidad);
            const total = parseFloat(formDataCorregido.total_ingresos);
            // Calculate unit price if possible
            formDataCorregido.ingresos_vr_unit = cantidad > 0 ? (total / cantidad).toString() : '0';
          } else {
            formDataCorregido.ingresos_vr_unit = '0';
          }
        } else {
          // Default safe values
          formDataCorregido.ingresos_cantidad = '0';
        }
        
        console.log("✅ VALORES CORREGIDOS:");
        console.log(`- fecha_solicitud: "${formDataCorregido.fecha_solicitud}"`);
        console.log(`- ingresos_cantidad: "${formDataCorregido.ingresos_cantidad}"`);
        console.log(`- ingresos_vr_unit: "${formDataCorregido.ingresos_vr_unit}"`);
      }
      
      // CASE 2: Use data from SOLICITUDES2 if available, otherwise fallback to SOLICITUDES
      // Priorize data from SOLICITUDES2 sheet over SOLICITUDES
      if (formDataCorregido.nombre_actividad) {
        console.log("ℹ️ Usando nombre_actividad de tabla SOLICITUDES2");
      } else if (solicitudData.nombre_actividad) {
        console.log("ℹ️ Usando nombre_actividad de tabla SOLICITUDES");
        formDataCorregido.nombre_actividad = solicitudData.nombre_actividad;
      }

      if (formDataCorregido.fecha_solicitud) {
        console.log("ℹ️ Usando fecha_solicitud de tabla SOLICITUDES2");
      } else if (solicitudData.fecha_solicitud) {
        console.log("ℹ️ Usando fecha_solicitud de tabla SOLICITUDES");
        formDataCorregido.fecha_solicitud = solicitudData.fecha_solicitud;
      }
      
      if (!formDataCorregido.nombre_solicitante && solicitudData.nombre_solicitante) {
        console.log("ℹ️ Usando nombre_solicitante de tabla SOLICITUDES");
        formDataCorregido.nombre_solicitante = solicitudData.nombre_solicitante;
      }
      
      // CASE 3: Ensure numeric fields contain valid numbers
      const numericFields = [
        'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos', 
        'subtotal_gastos', 'imprevistos_3', 'total_gastos_imprevistos',
        'fondo_comun_porcentaje', 'facultad_instituto_porcentaje', 
        'escuela_departamento_porcentaje'
      ];
      
      numericFields.forEach(field => {
        // Check if field exists and is not a valid number
        if (formDataCorregido[field] !== undefined && !isNumeric(formDataCorregido[field])) {
          console.log(`⚠️ CORRECCIÓN: Campo ${field} no es numérico "${formDataCorregido[field]}", estableciendo a 0`);
          formDataCorregido[field] = '0';
        } else if (formDataCorregido[field] === undefined) {
          formDataCorregido[field] = '0';
        }
      });
      
      // CASE 4: Calculate missing values where possible
      if (isNumeric(formDataCorregido.ingresos_cantidad) && isNumeric(formDataCorregido.ingresos_vr_unit)) {
        const cantidad = parseFloat(formDataCorregido.ingresos_cantidad);
        const valorUnit = parseFloat(formDataCorregido.ingresos_vr_unit);
        const totalCalculado = cantidad * valorUnit;
        
        if (!isNumeric(formDataCorregido.total_ingresos) || parseFloat(formDataCorregido.total_ingresos) === 0) {
          console.log(`ℹ️ Calculando total_ingresos: ${cantidad} × ${valorUnit} = ${totalCalculado}`);
          formDataCorregido.total_ingresos = totalCalculado.toString();
        }
      }
      
      // Combine all data with corrected values
      const datosCorregidos = {
        ...solicitudData,
        ...formDataCorregido
      };
      
      // Now continue with the rest of the transformation
      const transformedData = {};
      
      // Pre-initialize placeholders for expenses
      const conceptosGastos = [
        '1', '1,1', '1,2', '1,3', '2', '3', '4', '5', '6', '7', '7,1', '7,2', 
        '7,3', '7,4', '7,5', '8', '8,1', '8,2', '8,3', '8,4', '9', '9,1', '9,2', 
        '9,3', '10', '11', '12', '13', '14', '15'
      ];
      
      conceptosGastos.forEach(concepto => {
        transformedData[`gasto_${concepto}_cantidad`] = '0';
        transformedData[`gasto_${concepto}_valor_unit`] = '$0';
        transformedData[`gasto_${concepto}_valor_total`] = '$0';
        transformedData[`gasto_${concepto}_descripcion`] = '';
      });
      
      // Copy all corrected data to the result
      Object.keys(datosCorregidos).forEach(key => {
        if (datosCorregidos[key] !== undefined && datosCorregidos[key] !== null) {
          transformedData[key] = datosCorregidos[key];
        }
      });
      
      // Process date formatting
      const fechaActual = new Date();
      try {
        const fechaStr = transformedData.fecha_solicitud;
        if (fechaStr) {
          let fechaProcesada;
          
          // Try to parse the date string based on format
          if (fechaStr.includes('/')) {
            // Format: dd/mm/yyyy
            const [dia, mes, anio] = fechaStr.split('/');
            fechaProcesada = new Date(parseInt(anio), parseInt(mes) - 1, parseInt(dia));
          } else if (fechaStr.includes('-')) {
            // Format: yyyy-mm-dd or dd-mm-yyyy
            const parts = fechaStr.split('-');
            if (parts[0].length === 4) {
              // yyyy-mm-dd
              fechaProcesada = new Date(fechaStr);
            } else {
              // dd-mm-yyyy
              const [dia, mes, anio] = parts;
              fechaProcesada = new Date(parseInt(anio), parseInt(mes) - 1, parseInt(dia));
            }
          } else {
            // Try standard Date parsing
            fechaProcesada = new Date(fechaStr);
          }
          
          if (!isNaN(fechaProcesada.getTime())) {
            // Valid date, extract parts
            transformedData.dia = fechaProcesada.getDate().toString().padStart(2, '0');
            transformedData.mes = (fechaProcesada.getMonth() + 1).toString().padStart(2, '0');
            transformedData.anio = fechaProcesada.getFullYear().toString();
            
            // Log successful date extraction
            console.log(`✅ Fecha procesada correctamente: dia=${transformedData.dia}, mes=${transformedData.mes}, anio=${transformedData.anio}`);
          } else {
            // Invalid date, use current date
            console.log(`⚠️ Fecha inválida: "${fechaStr}", usando fecha actual`);
            transformedData.dia = fechaActual.getDate().toString().padStart(2, '0');
            transformedData.mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
            transformedData.anio = fechaActual.getFullYear().toString();
          }
        } else {
          // No date available, use current date
          console.log('ℹ️ No hay fecha_solicitud, usando fecha actual');
          transformedData.dia = fechaActual.getDate().toString().padStart(2, '0');
          transformedData.mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
          transformedData.anio = fechaActual.getFullYear().toString();
        }
      } catch (error) {
        console.error('Error al procesar la fecha:', error);
        transformedData.dia = fechaActual.getDate().toString().padStart(2, '0');
        transformedData.mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
        transformedData.anio = fechaActual.getFullYear().toString();
      }
      
      // Format currency values
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
      
      // Format specific monetary fields
      const monetaryFields = [
        'ingresos_vr_unit', 'total_ingresos', 'subtotal_gastos', 
        'imprevistos_3', 'total_gastos_imprevistos', 'total_recursos',
        'costos_indirectos_cantidad', 'administracion_cantidad', 
        'descuentos_cantidad', 'total_costo_actividad', 'excedente_cantidad',
        'valor_inscripcion_individual'
      ];
      
      // Ensure ingresos fields are properly populated
      if (isNumeric(transformedData.ingresos_cantidad) && isNumeric(transformedData.ingresos_vr_unit)) {
        const cantidad = parseFloat(transformedData.ingresos_cantidad);
        const valorUnit = parseFloat(transformedData.ingresos_vr_unit);
        const totalCalculado = cantidad * valorUnit;
        
        // Update formatted value and ensure total_ingresos is calculated
        transformedData.ingresos_cantidad_formatted = cantidad.toString();
        transformedData.ingresos_vr_unit_formatted = formatCurrency(valorUnit);
        transformedData.total_ingresos = totalCalculado.toString();
        transformedData.total_ingresos_formatted = formatCurrency(totalCalculado);
        
        console.log(`✅ Valores de ingresos procesados: cantidad=${cantidad}, valorUnit=${valorUnit}, total=${totalCalculado}`);
      } else {
        console.log("⚠️ Valores de ingresos inválidos o ausentes, estableciendo valores por defecto");
        transformedData.ingresos_cantidad = transformedData.ingresos_cantidad || '0';
        transformedData.ingresos_vr_unit = transformedData.ingresos_vr_unit || '0';
        transformedData.total_ingresos = transformedData.total_ingresos || '0';
      }
      
      monetaryFields.forEach(field => {
        if (transformedData[field]) {
          transformedData[field + '_formatted'] = formatCurrency(transformedData[field]);
        }
      });
      
      // PROCESAMIENTO DE GASTOS
      // First check gastosFromAdditional which comes from processGastosData
      if (gastosFromAdditional && gastosFromAdditional.length > 0) {
        console.log(`Procesando ${gastosFromAdditional.length} gastos normales de datos adicionales`);
        
        gastosFromAdditional.forEach(gasto => {
          // La plantilla usa formato con coma (1,1)
          const idConComa = gasto.id.replace(/\./g, ',');
          const placeholderId = `gasto_${idConComa}`;
          
          // Asignar valores a sus respectivos placeholders
          transformedData[`${placeholderId}_cantidad`] = gasto.cantidad.toString();
          transformedData[`${placeholderId}_valor_unit`] = gasto.valorUnit.toString();
          transformedData[`${placeholderId}_valor_unit_formatted`] = gasto.valorUnit_formatted;
          transformedData[`${placeholderId}_valor_total`] = gasto.valorTotal.toString();
          transformedData[`${placeholderId}_valor_total_formatted`] = gasto.valorTotal_formatted;
          transformedData[`${placeholderId}_descripcion`] = gasto.descripcion || gasto.concepto;
        });
      }
      // Then check raw gastosData from GASTOS sheet
      else if (gastosData && gastosData.length > 0) {
        console.log(`Procesando ${gastosData.length} gastos desde hoja GASTOS`);
        
        // Filter gastos that match the solicitudId
        const gastosFiltrados = gastosData.filter(g => g.id_solicitud === datosCorregidos.id_solicitud);
        console.log(`Procesando ${gastosFiltrados.length} gastos para solicitud ${datosCorregidos.id_solicitud}`);
        
        gastosFiltrados.forEach(gasto => {
          const idConcepto = gasto.id_conceptos;
          const idConComa = idConcepto.replace(/\./g, ','); // Convert '1.1' to '1,1'
          const placeholderId = `gasto_${idConComa}`;
          const cantidad = parseFloat(gasto.cantidad) || 0;
          const valorUnit = parseFloat(gasto.valor_unit) || 0;
          const valorTotal = parseFloat(gasto.valor_total) || cantidad * valorUnit;
          
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
      
      // Process dynamic expenses if available
      if (gastosDinamicos && gastosDinamicos.length > 0) {
        console.log(`Procesando ${gastosDinamicos.length} gastos dinámicos para el reporte`);
        
        // Add special field for dynamic expenses
        transformedData['__GASTOS_DINAMICOS__'] = {
          insertarEn: 'E45', // Adjust this to the correct insertion point in your template
          gastos: gastosDinamicos
        };
      }
      
      // IMPORTANTE: Asegurarse que los valores de subtotal_gastos e imprevistos son correctos
      // En caso de que estos valores vengan desplazados, calcularlos basados en gastos
      // Calcular subtotal_gastos si no existe o es inválido
      if (!transformedData.subtotal_gastos || isNaN(parseFloat(transformedData.subtotal_gastos))) {
        let subtotalCalculado = 0;
        
        // Sumar todos los gastos normales
        conceptosGastos.forEach(concepto => {
          const valorTotal = parseFloat(transformedData[`gasto_${concepto}_valor_total`]) || 0;
          subtotalCalculado += valorTotal;
        });
        
        // Sumar gastos dinámicos si existen
        if (gastosDinamicos && gastosDinamicos.length > 0) {
          gastosDinamicos.forEach(gasto => {
            subtotalCalculado += parseFloat(gasto.valorTotal) || 0;
          });
        }
        
        transformedData.subtotal_gastos = subtotalCalculado.toString();
        console.log(`✏️ Recalculado subtotal_gastos: ${subtotalCalculado}`);
      }
      
      // Calcular imprevistos_3 como 3% del subtotal_gastos
      const subtotalGastos = parseFloat(transformedData.subtotal_gastos) || 0;
      // Usar exactamente 3% para imprevistos, independientemente del valor en imprevistos_3%
      const imprevistos3 = subtotalGastos * 0.03;
      transformedData.imprevistos_3 = imprevistos3.toString();
      transformedData['imprevistos_3%'] = '3'; // Fijar el porcentaje en 3%
      
      // Calcular total_gastos_imprevistos como suma del subtotal_gastos + imprevistos_3
      const totalGastosImprevistos = subtotalGastos + imprevistos3;
      transformedData.total_gastos_imprevistos = totalGastosImprevistos.toString();
      
      // Calcular los valores monetarios a partir de los porcentajes para el fondo común, facultad e instituto, y escuela
      const totalIngresos = parseFloat(transformedData.total_ingresos) || 0;
      
      // Obtener porcentajes (usar valores por defecto si no existen)
      const fondoComunPorcentaje = parseFloat(transformedData.fondo_comun_porcentaje) || 30;
      const facultadInstitutoPorcentaje = 5; // Fijo en 5%
      const escuelaDepartamentoPorcentaje = parseFloat(transformedData.escuela_departamento_porcentaje) || 0;
      
      // Calcular valores monetarios
      const fondoComun = totalIngresos * (fondoComunPorcentaje / 100);
      const facultadInstituto = totalIngresos * (facultadInstitutoPorcentaje / 100);
      const escuelaDepartamento = totalIngresos * (escuelaDepartamentoPorcentaje / 100);
      
      // Calcular total de recursos
      const totalRecursos = fondoComun + facultadInstituto + escuelaDepartamento;
      
      // Asignar valores calculados
      transformedData.fondo_comun = fondoComun.toString();
      transformedData.facultad_instituto = facultadInstituto.toString();
      transformedData.escuela_departamento = escuelaDepartamento.toString();
      transformedData.total_recursos = totalRecursos.toString();
      
      // Guardar también el porcentaje de facultad_instituto para referencia
      transformedData.facultad_instituto_porcentaje = facultadInstitutoPorcentaje.toString();
      
      console.log(`✅ Cálculos de gastos: subtotal=${subtotalGastos}, imprevistos(3%)=${imprevistos3}, total=${totalGastosImprevistos}`);
      console.log(`✅ Cálculos de aportes: fondo_comun(${fondoComunPorcentaje}%)=${fondoComun}, facultad(${facultadInstitutoPorcentaje}%)=${facultadInstituto}, escuela(${escuelaDepartamentoPorcentaje}%)=${escuelaDepartamento}, total=${totalRecursos}`);
      
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
        message: error.message
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

// Helper function if dateUtils.getDateParts is not available
function getDateParts(date) {
  try {
    const dateObj = new Date(date);
    
    if (isNaN(dateObj.getTime())) {
      return { dia: '', mes: '', anio: '' };
    }
    
    return {
      dia: dateObj.getDate().toString().padStart(2, '0'),
      mes: (dateObj.getMonth() + 1).toString().padStart(2, '0'),
      anio: dateObj.getFullYear().toString()
    };
  } catch (error) {
    console.error('Error en getDateParts:', error);
    return { dia: '', mes: '', anio: '' };
  }
}

module.exports = report2Config;
