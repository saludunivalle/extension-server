const dateUtils = require('../utils/dateUtils');

/**
 * Configuración específica para el reporte del Formulario 2 - Presupuesto
 * Implementación siguiendo el patrón exacto de report1Config que funciona correctamente
 */
const formatCurrency = (value) => {
  if (!value && value !== 0) return '0';
  const numValue = parseFloat(value);
  if (isNaN(numValue)) return '0';
  
  // Formato esperado por la plantilla: valor numérico con símbolo de moneda
  return `$${Math.round(numValue).toLocaleString('es-CO')}`;
};

const report2Config = {
  title: 'Formulario de Presupuesto - F-06-MP-05-01-01',
  templateId: '1JY-4IfJqEWLqZ_wrq_B_bfIlI9MeVzgF',
  showHeader: true,
  requiresAdditionalData: true,
  requiresGastos: true,
  
  // Definición de hojas necesarias para este reporte (copiando el patrón de report1)
  sheetDefinitions: {
    SOLICITUDES2: {
      range: 'SOLICITUDES2!A2:CL',
      fields: [
        'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'nombre_solicitante',
        'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos', 'subtotal_gastos',
        'imprevistos_3', 'total_gastos_imprevistos', 'fondo_comun_porcentaje',
        'facultadad_instituto_porcentaje', 'escuela_departamento_porcentaje',
        'total_recursos', 'observaciones', 'responsable_financiero'
      ]
    },
    GASTOS: {
      range: 'GASTOS!A2:F500',
      fields: [
        'id_conceptos', 'id_solicitud', 'cantidad', 'valor_unit', 'valor_total', 'concepto_padre'
      ]
    }
  },
  
  /**
   * Transforma los datos para el reporte 2 siguiendo el modelo de report1
   * @param {Object} allData - Datos de la solicitud y adicionales
   * @returns {Object} - Datos transformados para la plantilla
   */
  transformData: function(allData) {
    console.log("🔄 Iniciando transformación para reporte 2 con patrón de report1");
    
    // Obtener datos de la solicitud
    const solicitudData = allData.SOLICITUDES2 || {};
    const gastosData = allData.GASTOS || [];
    
    // PASO 1: Inicializar el objeto con TODOS los campos posibles
    let transformedData = {};
    
    // Lista completa de campos básicos que podrían aparecer en la plantilla
    const allBasicFields = [
      // Datos de identificación
      'id_solicitud', 'nombre_actividad', 'fecha_solicitud', 'nombre_solicitante',
      'dia', 'mes', 'anio',
      
      // Ingresos
      'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos',
      
      // Gastos
      'subtotal_gastos', 'imprevistos_3', 'imprevistos_3%', 'total_gastos_imprevistos',
      
      // Distribución
      'fondo_comun_porcentaje', 'facultadad_instituto_porcentaje', 
      'escuela_departamento_porcentaje', 'total_recursos',
      
      // Otros
      'observaciones', 'responsable_financiero', 'responsable_firma'
    ];
    
    // Inicializar todos los campos básicos
    allBasicFields.forEach(field => {
      transformedData[field] = '';
    });
    
    // PASO 2: Inicializar todos los campos de gastos con valores predeterminados
    // Lista de todos los ID de gastos que podrían aparecer en la plantilla
    const gastosIDs = [
      '1', '1,1', '1,2', '1,3', '2', '3', '4', '5', '6', '7', '7,1', '7,2', 
      '7,3', '7,4', '7,5', '8', '8,1', '8,2', '8,3', '8,4', '9', '9,1', '9,2', 
      '9,3', '10', '11', '12', '13', '14', '15'
    ];
    
    // Crear todas las variantes de campos para los gastos
    gastosIDs.forEach(id => {
      transformedData[`gasto_${id}_cantidad`] = '0';
      transformedData[`gasto_${id}_valor_unit`] = '$0';
      transformedData[`gasto_${id}_valor_total`] = '$0';
      transformedData[`gasto_${id}_descripcion`] = '';
    });
    
    // PASO 3: Copiar los datos de la solicitud al resultado
    Object.keys(solicitudData).forEach(key => {
      if (solicitudData[key] !== undefined && solicitudData[key] !== null) {
        transformedData[key] = solicitudData[key];
      }
    });
    
    // PASO 4: Procesar fecha siguiendo el mismo patrón que report1
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
    
    // PASO 5: Procesar campos monetarios con especial atención a los de gastos
    [
      'ingresos_vr_unit', 
      'total_ingresos', 
      'subtotal_gastos', // Asegurar que esté incluido
      'imprevistos_3', 
      'total_gastos_imprevistos'
    ].forEach(field => {
      if (transformedData[field]) {
        transformedData[field] = formatCurrency(transformedData[field]);
      } else {
        // Asignar valores por defecto si falta alguno de estos campos críticos
        if (field === 'subtotal_gastos') {
          transformedData[field] = '$0';
        } else if (field === 'imprevistos_3') {
          // Si no hay imprevistos y sí hay subtotal, calcular el 3%
          if (transformedData.subtotal_gastos) {
            const subtotal = parseFloat(transformedData.subtotal_gastos.replace(/[$,.]/g, '')) || 0;
            transformedData[field] = formatCurrency(subtotal * 0.03);
          } else {
            transformedData[field] = '$0';
          }
        } else if (field === 'total_gastos_imprevistos') {
          // Si no hay total pero sí subtotal e imprevistos, calcular la suma
          const subtotal = parseFloat(transformedData.subtotal_gastos?.replace(/[$,.]/g, '')) || 0;
          const imprevistos = parseFloat(transformedData.imprevistos_3?.replace(/[$,.]/g, '')) || 0;
          transformedData[field] = formatCurrency(subtotal + imprevistos);
        } else {
          transformedData[field] = '$0';
        }
      }
    });

    // PASO 6: Verificar campos críticos y asignar valores por defecto
    if (!transformedData.fondo_comun_porcentaje) transformedData.fondo_comun_porcentaje = '30';
    if (!transformedData.facultadad_instituto_porcentaje) transformedData.facultadad_instituto_porcentaje = '5';
    if (!transformedData.escuela_departamento_porcentaje) transformedData.escuela_departamento_porcentaje = '0';
    if (!transformedData['imprevistos_3%']) transformedData['imprevistos_3%'] = '3';

    // NUEVO PASO: Verificar y recalcular valores monetarios relacionados
    // 1. Extraer valores numéricos (quitando símbolos de moneda y separadores)
    const getNumericValue = (value) => {
      if (!value) return 0;
      return parseFloat(String(value).replace(/[$,.]/g, '')) || 0;
    };

    // 2. Recalcular los valores monetarios en base al subtotal_gastos
    const subtotalGastos = getNumericValue(transformedData.subtotal_gastos);
    const imprevistosPorc = parseFloat(transformedData['imprevistos_3%'] || 3) / 100;
    const imprevistosValor = subtotalGastos * imprevistosPorc;
    const totalGastosImprevistos = subtotalGastos + imprevistosValor;

    // 3. Actualizar los campos con los valores recalculados
    transformedData.subtotal_gastos = formatCurrency(subtotalGastos);
    transformedData.imprevistos_3 = formatCurrency(imprevistosValor);
    transformedData.total_gastos_imprevistos = formatCurrency(totalGastosImprevistos);

    // Log para verificar los valores
    console.log("🔢 VERIFICACIÓN DE CÁLCULOS:");
    console.log(`- subtotal_gastos: ${transformedData.subtotal_gastos}`);
    console.log(`- imprevistos (${transformedData['imprevistos_3%']}%): ${transformedData.imprevistos_3}`);
    console.log(`- total_gastos_imprevistos: ${transformedData.total_gastos_imprevistos}`);
    
    // PASO 7: Procesar gastos - VERSIÓN MEJORADA
    console.log(`Procesando ${gastosData.length} gastos encontrados para la solicitud`);
    console.log(`Datos de GASTOS: ${JSON.stringify(gastosData.slice(0, 2))}`); // Mostrar primeros 2 gastos para debug

    // Agrupar gastos por ID para procesar los totales correctamente
    const gastosPorID = {};

    // Procesamiento mejorado de gastos
    gastosData.forEach(gasto => {
      // Validar que tenemos id_conceptos
      if (!gasto || !gasto.id_conceptos) {
        console.log("⚠️ Gasto sin ID de concepto detectado:", gasto);
        return;
      }
      
      // Normalizar id_conceptos a string y garantizar formato
      const idConcepto = gasto.id_conceptos.toString();
      
      // Crear MÚLTIPLES FORMATOS para garantizar compatibilidad
      // 1. Formato con coma (1,1) - formato principal para la plantilla
      const idConComa = idConcepto.replace(/\./g, ',');
      // 2. Formato con punto (1.1) - formato alternativo
      const idConPunto = idConcepto.replace(/,/g, '.');
      // 3. Formato con underscore (1_1) - otro formato alternativo
      const idConUnderscore = idConcepto.replace(/[,\.]/g, '_');
      
      // Valores numéricos
      const cantidad = parseFloat(gasto.cantidad) || 0;
      const valorUnit = parseFloat(gasto.valor_unit) || 0;
      const valorTotal = parseFloat(gasto.valor_total) || 0;
      
      // Registrar lo que estamos procesando para depuración
      console.log(`🔄 Procesando gasto: ID=${idConcepto}, cantidad=${cantidad}, valor=${valorUnit}`);
      
      // Agregar TODOS los formatos posibles para asegurar compatibilidad
      // Formato principal con coma
      transformedData[`gasto_${idConComa}_cantidad`] = cantidad.toString();
      transformedData[`gasto_${idConComa}_valor_unit`] = formatCurrency(valorUnit);
      transformedData[`gasto_${idConComa}_valor_total`] = formatCurrency(valorTotal);
      
      // Formato con punto
      transformedData[`gasto_${idConPunto}_cantidad`] = cantidad.toString();
      transformedData[`gasto_${idConPunto}_valor_unit`] = formatCurrency(valorUnit);
      transformedData[`gasto_${idConPunto}_valor_total`] = formatCurrency(valorTotal);
      
      // También agregar versión sin prefijo para mayor compatibilidad
      transformedData[`${idConComa}_cantidad`] = cantidad.toString();
      transformedData[`${idConComa}_valor_unit`] = formatCurrency(valorUnit);
      transformedData[`${idConComa}_valor_total`] = formatCurrency(valorTotal);

      // NUEVO: Formato con guion bajo (para placeholders sin coma)
      const idSinComa = idConComa.replace(/,/g, '_');
      transformedData[`gasto_${idSinComa}_cantidad`] = cantidad.toString();
      transformedData[`gasto_${idSinComa}_valor_unit`] = formatCurrency(valorUnit);
      transformedData[`gasto_${idSinComa}_valor_total`] = formatCurrency(valorTotal);
      
      // NUEVO: Versión simplificada para IDs principales
      if (idConComa.length === 1 || idConComa === '1,1' || idConComa === '1,2' || idConComa === '1,3') {
        const idSimple = idConComa.split(',')[0];  // Tomar solo el primer número
        transformedData[`gasto_${idSimple}_cantidad`] = cantidad.toString();
        transformedData[`gasto_${idSimple}_valor_unit`] = formatCurrency(valorUnit);
        transformedData[`gasto_${idSimple}_valor_total`] = formatCurrency(valorTotal);
      }
      
      // Guardar para agrupar por categoría padre
      const idPadre = gasto.concepto_padre || idConComa.split(',')[0];
      if (!gastosPorID[idPadre]) {
        gastosPorID[idPadre] = {
          total: 0,
          items: []
        };
      }
      
      gastosPorID[idPadre].total += valorTotal;
      gastosPorID[idPadre].items.push(gasto);
    });

    // Al final del procesamiento, verificar qué campos se generaron
    console.log("✅ Campos generados para gastos:", 
      Object.keys(transformedData)
        .filter(key => key.includes('gasto_'))
        .slice(0, 10) // mostrar solo los primeros 10 para no saturar logs
    );
    
    // PASO 8: Usar el mismo enfoque que report1 para validar campos finales
    // Verificar campos críticos como ingresos y gastos
    const camposCriticos = [
      'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos',
      'subtotal_gastos', 'imprevistos_3', 'total_gastos_imprevistos'
    ];
    
    camposCriticos.forEach(campo => {
      if (!transformedData[campo]) {
        if (campo.includes('cantidad') || campo === 'ingresos_cantidad') {
          transformedData[campo] = '0';
        } else if (campo.includes('vr_unit') || campo.includes('total')) {
          transformedData[campo] = '$0';
        }
      }
    });
    
    // Datos de prueba adicionales para depuración
    transformedData.nombre_actividad = transformedData.nombre_actividad || "ACTIVIDAD DE PRUEBA";
    
    // PASO 9: Eliminación final de placeholders sin reemplazar (como en report1)
    Object.keys(transformedData).forEach(key => {
      const value = transformedData[key];
      if (typeof value === 'string' && (
        value.includes('{{') || 
        value.includes('}}') || 
        value === 'undefined' || 
        value === 'null'
      )) {
        // Reemplazar con cadena vacía en lugar de dejar el placeholder
        transformedData[key] = '';
      }
    });

    // Verificar datos finales antes de retornar
    console.log("✅ Datos finales:\n", 
      Object.entries(transformedData)
        .filter(([key]) => key.includes('gasto_') || key.includes('ingresos_'))
        .map(([key, value]) => `${key}: ${value}`)
        .join('\n')
    );

    console.log("✅ Transformación completada para reporte 2");
    
    return transformedData;
  },
  
  sheetsConfig: {
    sheetName: 'Formulario2',
    dataRange: 'A1:Z100'
  },
  footerText: 'Universidad del Valle - Extensión y Proyección Social - Presupuesto',
  watermark: false,
  
  // Mantener procesamiento de filas dinámicas
  processDynamicRows: async (spreadsheetId, data, sheetsApi) => {
    try {
      // Verificar si existen filas dinámicas para procesar
      if (data['__FILAS_DINAMICAS__']) {
        const filasDinamicas = data['__FILAS_DINAMICAS__'];
        const insertarEn = filasDinamicas.insertarEn || 'E45:AK45';
        const gastos = filasDinamicas.gastos || [];
        
        if (gastos.length > 0) {
          // Preparar los valores para todas las filas
          const values = gastos.map(gasto => {
            // Preparar el array
            const rowData = new Array(37).fill('');
            
            // Colocar valores en las posiciones correctas
            rowData[0] = gasto.id_concepto || '';
            rowData[1] = gasto.descripcion || '';
            rowData[23] = gasto.cantidad?.toString() || '';
            rowData[25] = gasto.valor_unit_formatted || '';
            rowData[28] = gasto.valor_total_formatted || '';
            
            return rowData;
          });
          
          // Insertar las filas una a una
          const match = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(insertarEn);
          if (match) {
            const startRow = parseInt(match[2]);
            
            for (let i = 0; i < values.length; i++) {
              const rowNum = startRow + i;
              const range = `E${rowNum}:AK${rowNum}`;
              
              await sheetsApi.spreadsheets.values.update({
                spreadsheetId,
                range,
                valueInputOption: 'USER_ENTERED',
                resource: { values: [values[i]] }
              });
            }
            
            console.log(`✅ Insertadas ${values.length} filas de gastos dinámicos`);
          }
        }
        
        // Limpiar la propiedad especial
        delete data['__FILAS_DINAMICAS__'];
      }
    } catch (error) {
      console.error('Error al procesar filas dinámicas:', error);
    }
    
    return data;
  }
};

module.exports = report2Config;