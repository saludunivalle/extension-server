const dateUtils = require('../utils/dateUtils');
const { generateExpenseRows } = require('../services/dynamicRows');

/**
 * Configuración específica para el reporte del Formulario 2 - Presupuesto
 */
const report2Config = {
  title: 'Formulario de Presupuesto - F-05-MP-05-01-02',
  templateId: '1nWY2gYKtJuXQnGLsdN7RID_2QmrHKRtOwcQsaTsOOm8',
  requiresAdditionalData: true,
  requiresGastos: true, // Budget form needs expense data
  
  // Definición de hojas necesarias para este reporte
  sheetDefinitions: {
    SOLICITUDES2: {
      range: 'SOLICITUDES2!A2:R',
      fields: [
        'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos',
        'subtotal_gastos', 'imprevistos_3', 'total_gastos_imprevistos', 'diferencia',
        'fondo_comun_porcentaje', 'fondo_comun', 'facultad_instituto_porcentaje', 'facultad_instituto',
        'escuela_departamento_porcentaje', 'escuela_departamento', 'total_recursos', 'observaciones'
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
    // === LOGS DE DEPURACIÓN CRÍTICOS ===
    console.log('🔥 [REPORT2CONFIG] transformData ejecutándose...');
    console.log('🔥 [REPORT2CONFIG] allData keys:', Object.keys(allData || {}));
    console.log('🔥 [REPORT2CONFIG] allData.SOLICITUDES2:', allData.SOLICITUDES2);
    console.log('🔥 [REPORT2CONFIG] allData.SOLICITUDES:', allData.SOLICITUDES);
    console.log('🔥 [REPORT2CONFIG] estructura completa allData:', JSON.stringify(allData, null, 2));
    
    // ===== NUEVA LÓGICA: TRABAJAR CON DATOS APLANADOS =====
    // Los datos vienen directamente en allData, no anidados por hoja
    
    // CORRECCIÓN CRÍTICA: Los campos están intercambiados en el mapeo
    // Lo que llega como "nombre_actividad" es realmente la fecha
    // Lo que llega como "fecha_solicitud" es realmente el nombre
    const nombre_actividad_real = allData.fecha_solicitud || ''; // ¡CORREGIDO!
    const fecha_solicitud_real = allData.nombre_actividad || ''; // ¡CORREGIDO!
    
    console.log('🔥 [CORRECCION] nombre_actividad_real:', nombre_actividad_real);
    console.log('🔥 [CORRECCION] fecha_solicitud_real:', fecha_solicitud_real);
    
    // Crear objeto base con valores corregidos
    const transformedData = {
      // Campos corregidos
      id_solicitud: allData.id_solicitud || '',
      nombre_actividad: nombre_actividad_real,
      fecha_solicitud: fecha_solicitud_real,
      
      // Campos de ingresos
      ingresos_cantidad: allData.ingresos_cantidad || '0',
      ingresos_vr_unit: allData.ingresos_vr_unit || '0',
      total_ingresos: allData.total_ingresos || '0',
      
      // Campos de gastos
      subtotal_gastos: allData.subtotal_gastos || '0',
      imprevistos_3: allData.imprevistos_3 || '0',
      total_gastos_imprevistos: allData.total_gastos_imprevistos || '0',
      diferencia: allData.diferencia || '0',
      
      // Campos de aportes
      fondo_comun_porcentaje: allData.fondo_comun_porcentaje || '30',
      fondo_comun: allData.fondo_comun || '0',
      facultad_instituto_porcentaje: allData.facultad_instituto_porcentaje || '5',
      facultad_instituto: allData.facultad_instituto || '0',
      escuela_departamento_porcentaje: allData.escuela_departamento_porcentaje || '0',
      escuela_departamento: allData.escuela_departamento || '0',
      total_recursos: allData.total_recursos || '0',
      
      // Otros campos
      observaciones: allData.observaciones || '',
      nombre_solicitante: allData.nombre_solicitante || '',
      
      // Datos de gastos
      gastos: allData.gastos || [],
      gastosNormales: allData.gastosNormales || [],
      gastosDinamicos: allData.gastosDinamicos || [],
      gastosFormateados: allData.gastosFormateados || {}
    };
    
    console.log('🔥 [DESPUES_CORRECCION] transformedData inicial:', {
      id_solicitud: transformedData.id_solicitud,
      nombre_actividad: transformedData.nombre_actividad,
      fecha_solicitud: transformedData.fecha_solicitud,
      ingresos_cantidad: transformedData.ingresos_cantidad,
      ingresos_vr_unit: transformedData.ingresos_vr_unit,
      total_ingresos: transformedData.total_ingresos
    });
    
    // ===== CONTINUAR CON LA LÓGICA EXISTENTE =====
    // Ahora que tenemos los datos corregidos, continuar con el resto de la lógica
    
    const gastosFromAdditional = allData.gastosNormales || [];
    const gastosDinamicos = allData.gastosDinamicos || [];
    
    // Helper para priorizar SOLICITUDES2
    const getField = (field) => {
      if (transformedData[field] !== undefined && transformedData[field] !== null && transformedData[field] !== '') return transformedData[field];
      return '';
    };

    // Construir objeto combinado explícitamente
    const combinedData = {
      id_solicitud: getField('id_solicitud'),
      nombre_actividad: getField('nombre_actividad'),
      fecha_solicitud: getField('fecha_solicitud'),
      ingresos_cantidad: getField('ingresos_cantidad'),
      ingresos_vr_unit: getField('ingresos_vr_unit'),
      total_ingresos: getField('total_ingresos'),
      subtotal_gastos: getField('subtotal_gastos'),
      imprevistos_3: getField('imprevistos_3'),
      total_gastos_imprevistos: getField('total_gastos_imprevistos'),
      diferencia: getField('diferencia'),
      fondo_comun_porcentaje: getField('fondo_comun_porcentaje'),
      fondo_comun: getField('fondo_comun'),
      facultad_instituto_porcentaje: getField('facultad_instituto_porcentaje'),
      facultad_instituto: getField('facultad_instituto'),
      escuela_departamento_porcentaje: getField('escuela_departamento_porcentaje'),
      escuela_departamento: getField('escuela_departamento'),
      total_recursos: getField('total_recursos'),
      observaciones: getField('observaciones'),
      nombre_solicitante: getField('nombre_solicitante'),
    };

    // Agregar gastos y dinámicos al objeto combinado
    combinedData.gastos = transformedData.gastos;
    combinedData.gastosNormales = gastosFromAdditional;
    combinedData.gastosDinamicos = gastosDinamicos;
    combinedData.gastosFormateados = transformedData.gastosFormateados;

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
    console.log("- SOLICITUDES.nombre_actividad:", transformedData.nombre_actividad);
    console.log("- SOLICITUDES.fecha_solicitud:", transformedData.fecha_solicitud);
    console.log("- SOLICITUDES2.nombre_actividad:", transformedData.nombre_actividad);
    console.log("- SOLICITUDES2.fecha_solicitud:", transformedData.fecha_solicitud);
    console.log("- SOLICITUDES2.ingresos_cantidad:", transformedData.ingresos_cantidad);
    console.log("- SOLICITUDES2.ingresos_vr_unit:", transformedData.ingresos_vr_unit);
    console.log("- SOLICITUDES2.total_ingresos:", transformedData.total_ingresos);
    
    // Create a copy to avoid modifying the original data
    const formDataCorregido = { ...transformedData };
    
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
    } else if (transformedData.nombre_actividad) {
      console.log("ℹ️ Usando nombre_actividad de tabla SOLICITUDES");
      formDataCorregido.nombre_actividad = transformedData.nombre_actividad;
    }

    if (formDataCorregido.fecha_solicitud) {
      console.log("ℹ️ Usando fecha_solicitud de tabla SOLICITUDES2");
    } else if (transformedData.fecha_solicitud) {
      console.log("ℹ️ Usando fecha_solicitud de tabla SOLICITUDES");
      formDataCorregido.fecha_solicitud = transformedData.fecha_solicitud;
    }
    
    if (!formDataCorregido.nombre_solicitante && transformedData.nombre_solicitante) {
      console.log("ℹ️ Usando nombre_solicitante de tabla SOLICITUDES");
      formDataCorregido.nombre_solicitante = transformedData.nombre_solicitante;
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
      ...transformedData,
      ...formDataCorregido
    };
    
    // 1. Lista de campos de SOLICITUDES2
    const allSolicitud2Fields = report2Config.sheetDefinitions.SOLICITUDES2.fields;

    // 2. Asegura que todos los campos estén presentes y priorizados
    allSolicitud2Fields.forEach(field => {
      if (formDataCorregido[field] === undefined || formDataCorregido[field] === '') {
        formDataCorregido[field] = transformedData[field] || '';
      }
    });

    // 3. Crea el objeto final solo con los campos requeridos
    const datosCorregidosFinal = {};
    allSolicitud2Fields.forEach(field => {
      // Prioriza el valor de formDataCorregido (SOLICITUDES2 corregido), si no, usa el de transformedData (SOLICITUDES)
      datosCorregidosFinal[field] = (formDataCorregido[field] !== undefined && formDataCorregido[field] !== '')
        ? formDataCorregido[field]
        : (transformedData[field] || '');
    });

    // LOG ESPECIAL PARA nombre_actividad
    console.log('🟢 Valor final de nombre_actividad:', datosCorregidosFinal['nombre_actividad']);
    
    // Now continue with the rest of the transformation
    const transformedDataFinal = { ...combinedData };
    
    // Pre-initialize placeholders for expenses
    const conceptosGastos = [
      '1', '1,1', '1,2', '1,3',
      '2', '2,1', '2,2', '2,3',
      '3', '3,1', '3,2',
      '4', '4,1', '4,2', '4,3', '4,4',
      '5', '5,1', '5,2', '5,3',
      '6', '6,1', '6,2',
      '7', '7,1', '7,2', '7,3',
      '8' // Gastos dinámicos 8,1; 8,2; ...
    ];
    
    conceptosGastos.forEach(concepto => {
      transformedDataFinal[`gasto_${concepto}_cantidad`] = '0';
      transformedDataFinal[`gasto_${concepto}_valor_unit`] = '$0';
      transformedDataFinal[`gasto_${concepto}_valor_total`] = '$0';
      transformedDataFinal[`gasto_${concepto}_descripcion`] = '';
    });
    
    // Copy all corrected data to the result
    Object.keys(datosCorregidosFinal).forEach(key => {
      if (datosCorregidosFinal[key] !== undefined && datosCorregidosFinal[key] !== null) {
        transformedDataFinal[key] = datosCorregidosFinal[key];
      }
    });

    // LOG FINAL para depuración de nombre_actividad en el objeto transformado
    console.log('🟢 Valor de nombre_actividad en transformedData:', transformedDataFinal['nombre_actividad']);
    
    // Process date formatting
    const fechaActual = new Date();
    try {
      const fechaStr = transformedDataFinal.fecha_solicitud;
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
          transformedDataFinal.dia = fechaProcesada.getDate().toString().padStart(2, '0');
          transformedDataFinal.mes = (fechaProcesada.getMonth() + 1).toString().padStart(2, '0');
          transformedDataFinal.anio = fechaProcesada.getFullYear().toString();
          
          // Log successful date extraction
          console.log(`✅ Fecha procesada correctamente: dia=${transformedDataFinal.dia}, mes=${transformedDataFinal.mes}, anio=${transformedDataFinal.anio}`);
        } else {
          // Invalid date, use current date
          console.log(`⚠️ Fecha inválida: "${fechaStr}", usando fecha actual`);
          transformedDataFinal.dia = fechaActual.getDate().toString().padStart(2, '0');
          transformedDataFinal.mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
          transformedDataFinal.anio = fechaActual.getFullYear().toString();
        }
      } else {
        // No date available, use current date
        console.log('ℹ️ No hay fecha_solicitud, usando fecha actual');
        transformedDataFinal.dia = fechaActual.getDate().toString().padStart(2, '0');
        transformedDataFinal.mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
        transformedDataFinal.anio = fechaActual.getFullYear().toString();
      }
    } catch (error) {
      console.error('Error al procesar la fecha:', error);
      transformedDataFinal.dia = fechaActual.getDate().toString().padStart(2, '0');
      transformedDataFinal.mes = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
      transformedDataFinal.anio = fechaActual.getFullYear().toString();
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
    if (isNumeric(transformedDataFinal.ingresos_cantidad) && isNumeric(transformedDataFinal.ingresos_vr_unit)) {
      const cantidad = parseFloat(transformedDataFinal.ingresos_cantidad);
      const valorUnit = parseFloat(transformedDataFinal.ingresos_vr_unit);
      const totalCalculado = cantidad * valorUnit;
      
      // Update formatted value and ensure total_ingresos is calculated
      transformedDataFinal.ingresos_cantidad_formatted = cantidad.toString();
      transformedDataFinal.ingresos_vr_unit_formatted = formatCurrency(valorUnit);
      transformedDataFinal.total_ingresos = totalCalculado.toString();
      transformedDataFinal.total_ingresos_formatted = formatCurrency(totalCalculado);
      
      console.log(`✅ Valores de ingresos procesados: cantidad=${cantidad}, valorUnit=${valorUnit}, total=${totalCalculado}`);
    } else {
      console.log("⚠️ Valores de ingresos inválidos o ausentes, estableciendo valores por defecto");
      transformedDataFinal.ingresos_cantidad = transformedDataFinal.ingresos_cantidad || '0';
      transformedDataFinal.ingresos_vr_unit = transformedDataFinal.ingresos_vr_unit || '0';
      transformedDataFinal.total_ingresos = transformedDataFinal.total_ingresos || '0';
    }
    
    monetaryFields.forEach(field => {
      if (transformedDataFinal[field]) {
        transformedDataFinal[field + '_formatted'] = formatCurrency(transformedDataFinal[field]);
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
        transformedDataFinal[`${placeholderId}_cantidad`] = gasto.cantidad.toString();
        transformedDataFinal[`${placeholderId}_valor_unit`] = gasto.valorUnit.toString();
        transformedDataFinal[`${placeholderId}_valor_unit_formatted`] = gasto.valorUnit_formatted;
        transformedDataFinal[`${placeholderId}_valor_total`] = gasto.valorTotal.toString();
        transformedDataFinal[`${placeholderId}_valor_total_formatted`] = gasto.valorTotal_formatted;
        transformedDataFinal[`${placeholderId}_descripcion`] = gasto.descripcion || gasto.concepto;
      });
    }
    // Then check raw gastosData from GASTOS sheet
    else if (transformedDataFinal.gastos && transformedDataFinal.gastos.length > 0) {
      console.log(`Procesando ${transformedDataFinal.gastos.length} gastos desde hoja GASTOS`);
      
      // Filter gastos that match the solicitudId
      const gastosFiltrados = transformedDataFinal.gastos.filter(g => g.id_solicitud === datosCorregidosFinal.id_solicitud);
      console.log(`Procesando ${gastosFiltrados.length} gastos para solicitud ${datosCorregidosFinal.id_solicitud}`);
      
      gastosFiltrados.forEach(gasto => {
        const idConcepto = gasto.id_conceptos;
        const idConComa = idConcepto.replace(/\./g, ','); // Convert '1.1' to '1,1'
        const placeholderId = `gasto_${idConComa}`;
        const cantidad = parseFloat(gasto.cantidad) || 0;
        const valorUnit = parseFloat(gasto.valor_unit) || 0;
        const valorTotal = parseFloat(gasto.valor_total) || cantidad * valorUnit;
        
        // Asignar valores a sus respectivos placeholders
        transformedDataFinal[`${placeholderId}_cantidad`] = cantidad.toString();
        transformedDataFinal[`${placeholderId}_valor_unit`] = valorUnit.toString();
        transformedDataFinal[`${placeholderId}_valor_unit_formatted`] = formatCurrency(valorUnit);
        transformedDataFinal[`${placeholderId}_valor_total`] = valorTotal.toString();
        transformedDataFinal[`${placeholderId}_valor_total_formatted`] = formatCurrency(valorTotal);
      });
    } else {
      console.log('⚠️ No se encontraron gastos para procesar');
    }
    
    // Process dynamic expenses if available
    if (gastosDinamicos && gastosDinamicos.length > 0) {
      console.log(`Procesando ${gastosDinamicos.length} gastos dinámicos para incluir en el reporte`);
      
      // Use the dynamicRows service to generate the correct structure
      const dynamicRowsData = generateExpenseRows(gastosDinamicos);
      
      if (dynamicRowsData) {
        // IMPORTANT: Ensure dynamic rows are inserted at row 43
        console.log(`⚠️ CRÍTICO: Configurando inserción de filas dinámicas a partir de la fila 43`);
        
        // Add special field for dynamic rows with complete structure
        transformedDataFinal['__FILAS_DINAMICAS__'] = {
          gastos: dynamicRowsData.gastos,
          rows: dynamicRowsData.rows,
          insertarEn: "A42:AK42", // Template row range (usar fila 42 como template)
          insertStartRow: 43 // FIXED VALUE: Always insert at row 43
        };
        
        console.log(`✅ Configuración de filas dinámicas completada con insertStartRow=43`);
        console.log(`Estructura __FILAS_DINAMICAS__ generada:`, JSON.stringify(transformedDataFinal['__FILAS_DINAMICAS__'], null, 2));
      } else {
        console.log(`⚠️ No se pudo generar la estructura de filas dinámicas`);
      }
    }
    
    // IMPORTANTE: Asegurarse que los valores de subtotal_gastos e imprevistos son correctos
    // En caso de que estos valores vengan desplazados, calcularlos basados en gastos
    // Calcular subtotal_gastos si no existe o es inválido
    if (!transformedDataFinal.subtotal_gastos || isNaN(parseFloat(transformedDataFinal.subtotal_gastos))) {
      let subtotalCalculado = 0;
      
      // Sumar todos los gastos normales
      conceptosGastos.forEach(concepto => {
        const valorTotal = parseFloat(transformedDataFinal[`gasto_${concepto}_valor_total`]) || 0;
        subtotalCalculado += valorTotal;
      });
      
      // Sumar gastos dinámicos si existen
      if (gastosDinamicos && gastosDinamicos.length > 0) {
        gastosDinamicos.forEach(gasto => {
          subtotalCalculado += parseFloat(gasto.valorTotal) || 0;
        });
      }
      
      transformedDataFinal.subtotal_gastos = subtotalCalculado.toString();
      console.log(`✏️ Recalculado subtotal_gastos: ${subtotalCalculado}`);
    }
    
    // Calcular imprevistos_3 como 3% del subtotal_gastos
    const subtotalGastos = parseFloat(transformedDataFinal.subtotal_gastos) || 0;
    // Usar exactamente 3% para imprevistos, independientemente del valor en imprevistos_3%
    const imprevistos3 = subtotalGastos * 0.03;
    transformedDataFinal.imprevistos_3 = imprevistos3.toString();
    transformedDataFinal['imprevistos_3%'] = '3'; // Fijar el porcentaje en 3%
    
    // Calcular total_gastos_imprevistos como suma del subtotal_gastos + imprevistos_3
    const totalGastosImprevistos = subtotalGastos + imprevistos3;
    transformedDataFinal.total_gastos_imprevistos = totalGastosImprevistos.toString();
    
    // Calcular la diferencia (Ingresos - Gastos)
    const totalIngresos = parseFloat(transformedDataFinal.total_ingresos) || 0;
    const diferencia = totalIngresos - totalGastosImprevistos;
    transformedDataFinal.diferencia = diferencia.toString();
    
    // Calcular los valores monetarios a partir de los porcentajes para el fondo común, facultad e instituto, y escuela
    // Obtener porcentajes (usar valores por defecto si no existen)
    const fondoComunPorcentaje = parseFloat(transformedDataFinal.fondo_comun_porcentaje) || 30;
    const facultadInstitutoPorcentaje = parseFloat(transformedDataFinal.facultad_instituto_porcentaje) || 5; // Ahora editable
    const escuelaDepartamentoPorcentaje = parseFloat(transformedDataFinal.escuela_departamento_porcentaje) || 0;
    
    // Calcular valores monetarios
    const fondoComun = totalIngresos * (fondoComunPorcentaje / 100);
    const facultadInstituto = totalIngresos * (facultadInstitutoPorcentaje / 100);
    const escuelaDepartamento = totalIngresos * (escuelaDepartamentoPorcentaje / 100);
    
    // Calcular total de recursos
    const totalRecursos = fondoComun + facultadInstituto + escuelaDepartamento;
    
    // Asignar valores calculados
    transformedDataFinal.fondo_comun = fondoComun.toString();
    transformedDataFinal.facultad_instituto = facultadInstituto.toString();
    transformedDataFinal.escuela_departamento = escuelaDepartamento.toString();
    transformedDataFinal.total_recursos = totalRecursos.toString();
    
    // Guardar también el porcentaje de facultad_instituto para referencia
    transformedDataFinal.facultad_instituto_porcentaje = facultadInstitutoPorcentaje.toString();
    
    // Asegurar que el campo observaciones esté presente
    if (!transformedDataFinal.observaciones) {
      transformedDataFinal.observaciones = '';
    }
    
    console.log(`✅ Cálculos de gastos: subtotal=${subtotalGastos}, imprevistos(3%)=${imprevistos3}, total=${totalGastosImprevistos}, diferencia=${diferencia}`);
    console.log(`✅ Cálculos de aportes: fondo_comun(${fondoComunPorcentaje}%)=${fondoComun}, facultad(${facultadInstitutoPorcentaje}%)=${facultadInstituto}, escuela(${escuelaDepartamentoPorcentaje}%)=${escuelaDepartamento}, total=${totalRecursos}`);
    
    // IMPORTANTE: Eliminar marcadores no reemplazados
    Object.keys(transformedDataFinal).forEach(key => {
      const value = transformedDataFinal[key];
      if (typeof value === 'string' && (value.includes('{{') || value.includes('}}'))) {
        console.log(`⚠️ Detectado posible marcador en ${key}: "${value}"`);
        transformedDataFinal[key] = '';
      }
    });
    
    // Garantizar valores por defecto
    if (!transformedDataFinal['fecha_solicitud'] || !transformedDataFinal['dia']) {
      console.log('⚠️ Usando fecha por defecto para campos faltantes');
      const fechaActual = new Date();
      
      if (!transformedDataFinal['fecha_solicitud']) {
        transformedDataFinal['fecha_solicitud'] = `${fechaActual.getDate()}/${fechaActual.getMonth()+1}/${fechaActual.getFullYear()}`;
      }
      
      if (!transformedDataFinal['dia']) transformedDataFinal['dia'] = fechaActual.getDate().toString().padStart(2, '0');
      if (!transformedDataFinal['mes']) transformedDataFinal['mes'] = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
      if (!transformedDataFinal['anio']) transformedDataFinal['anio'] = fechaActual.getFullYear().toString();
    }
    
    console.log("⭐ DATOS TRANSFORMADOS FINALES - FORM 2:", transformedDataFinal);
    return transformedDataFinal;
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
