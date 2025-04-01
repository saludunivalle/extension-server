/**
 * Configuración específica para el reporte del Formulario 2 - Presupuesto
 * Implementación optimizada para marcado directo en Google Sheets y manejo de placeholders
 */
export const report2Config = {
  title: 'Formulario de Presupuesto - F-05-MP-05-01-02',
  showHeader: true,
  
  // Función para transformar los datos para Google Sheets
  transformData: async (formData) => {
    // Crear un objeto nuevo vacío
    const transformedData = {};
    
    // Añadir fecha actual formateada para el reporte
    const fechaActual = new Date();
    transformedData['dia'] = fechaActual.getDate().toString().padStart(2, '0');
    transformedData['mes'] = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
    transformedData['anio'] = fechaActual.getFullYear().toString();

    // Si hay una fecha de solicitud, usar esa fecha en lugar de la actual
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
    
    // Procesar la fecha para extraer día, mes y año
    if (formData.fecha_solicitud) {
      try {
        // Intentar diferentes formatos de fecha
        let fechaParts;
        
        if (formData.fecha_solicitud.includes('/')) {
          // Formato DD/MM/YYYY
          fechaParts = formData.fecha_solicitud.split('/');
          transformedData['dia'] = fechaParts[0];
          transformedData['mes'] = fechaParts[1];
          transformedData['anio'] = fechaParts[2];
        } else if (formData.fecha_solicitud.includes('-')) {
          // Formato YYYY-MM-DD o DD-MM-YYYY
          fechaParts = formData.fecha_solicitud.split('-');
          
          if (fechaParts[0].length === 4) {
            // Formato YYYY-MM-DD
            transformedData['dia'] = fechaParts[2];
            transformedData['mes'] = fechaParts[1];
            transformedData['anio'] = fechaParts[0];
          } else {
            // Formato DD-MM-YYYY
            transformedData['dia'] = fechaParts[0];
            transformedData['mes'] = fechaParts[1];
            transformedData['anio'] = fechaParts[2];
          }
        } else {
          // Si no se puede parsear, generar fecha actual
          const fechaActual = new Date();
          transformedData['dia'] = fechaActual.getDate().toString().padStart(2, '0');
          transformedData['mes'] = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
          transformedData['anio'] = fechaActual.getFullYear().toString();
        }
        
        console.log(`Fecha procesada: día=${transformedData['dia']}, mes=${transformedData['mes']}, año=${transformedData['anio']}`);
      } catch (error) {
        console.error('Error al procesar la fecha:', error);
        // En caso de error, usar la fecha actual
        const fechaActual = new Date();
        transformedData['dia'] = fechaActual.getDate().toString().padStart(2, '0');
        transformedData['mes'] = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
        transformedData['anio'] = fechaActual.getFullYear().toString();
      }
    } else {
      // Si no hay fecha_solicitud, usar la fecha actual
      const fechaActual = new Date();
      transformedData['dia'] = fechaActual.getDate().toString().padStart(2, '0');
      transformedData['mes'] = (fechaActual.getMonth() + 1).toString().padStart(2, '0');
      transformedData['anio'] = fechaActual.getFullYear().toString();
    }
    
    console.log("🔄 Transformando datos para Google Sheets - Formulario 2:", formData);
    
    // Formatear valores monetarios (función igual que antes)
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
      // Importar el servicio de sheetsService dinámicamente
      const sheetsService = require('../services/sheetsService');
      
      // Obtener también datos de conceptos para tener las descripciones
      const conceptosResponse = await sheetsService.client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'CONCEPTO$!A2:B500'
      });
      
      const conceptosRows = conceptosResponse.data.values || [];
      const conceptosMap = {};
      
      // Crear un mapa de ID de concepto a descripción
      conceptosRows.forEach(row => {
        conceptosMap[row[0]] = row[1] || `Concepto ${row[0]}`;
      });
      
      // Obtener todos los gastos para esta solicitud
      const gastosResponse = await sheetsService.client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'GASTOS!A2:F500' // Ampliado para más datos
      });
      
      const gastosRows = gastosResponse.data.values || [];
      const solicitudGastos = gastosRows.filter(row => row[1] === formData.id_solicitud);
      
      console.log(`Encontrados ${solicitudGastos.length} gastos para la solicitud ${formData.id_solicitud}`);
      
      // IMPORTANTE: Procesar TODOS los gastos, independientemente de su formato
      solicitudGastos.forEach(gasto => {
        const idConcepto = gasto[0]; // id_conceptos (puede ser "1", "1.3", "10", etc.)
        const cantidad = parseFloat(gasto[2]) || 0;
        const valorUnit = parseFloat(gasto[3]) || 0;
        const valorTotal = parseFloat(gasto[4]) || 0;
        
        // Obtener la descripción del concepto
        const descripcion = conceptosMap[idConcepto] || `Concepto ${idConcepto}`;
        
        console.log(`Procesando gasto: ID=${idConcepto}, Cantidad=${cantidad}, Unit=${valorUnit}, Total=${valorTotal}`);
        
        // Generar variables para cada formato de ID
        
        // 1. Formato original con punto (si existe)
        transformedData[`gasto_${idConcepto}_cantidad`] = cantidad.toString();
        transformedData[`gasto_${idConcepto}_valor_unit`] = formatCurrency(valorUnit);
        transformedData[`gasto_${idConcepto}_valor_total`] = formatCurrency(valorTotal);
        transformedData[`gasto_${idConcepto}_descripcion`] = descripcion;
        
        // 2. Formato con coma (para compatibilidad con Excel)
        const idConceptoConComa = idConcepto.replace('.', ',');
        if (idConcepto !== idConceptoConComa) {
          transformedData[`gasto_${idConceptoConComa}_cantidad`] = cantidad.toString();
          transformedData[`gasto_${idConceptoConComa}_valor_unit`] = formatCurrency(valorUnit);
          transformedData[`gasto_${idConceptoConComa}_valor_total`] = formatCurrency(valorTotal);
          transformedData[`gasto_${idConceptoConComa}_descripcion`] = descripcion;
        }
      });
      
      // NUEVO: Generar conjunto de todos los IDs de gastos encontrados para verificar
      const idsEncontrados = solicitudGastos.map(gasto => gasto[0]);
      console.log(`IDs de gastos encontrados: ${idsEncontrados.join(', ')}`);
      
      // NUEVO: Verificar específicamente qué campos de gastos están disponibles
      const camposGastosGenerados = Object.keys(transformedData).filter(
        key => key.startsWith('gasto_')
      );
      
      console.log(`Se generaron ${camposGastosGenerados.length} campos de gastos:`);
      camposGastosGenerados.forEach(campo => {
        console.log(`- ${campo}: ${transformedData[campo]}`);
      });

      // Agregar información sobre filas dinámicas que se insertarán en el reporte
      transformedData['__FILAS_DINAMICAS__'] = {
        insertarEn: 'E44:AK44',
        gastos: solicitudGastos.map(gasto => {
          const idConcepto = gasto[0];
          const cantidad = parseFloat(gasto[2]) || 0;
          const valorUnit = parseFloat(gasto[3]) || 0;
          const valorTotal = parseFloat(gasto[4]) || 0;
          const descripcion = conceptosMap[idConcepto] || `Concepto ${idConcepto}`;
          
          return {
            id_concepto: idConcepto,
            descripcion: descripcion,
            cantidad: cantidad,
            valor_unit: valorUnit,
            valor_total: valorTotal,
            valor_unit_formatted: formatCurrency(valorUnit),
            valor_total_formatted: formatCurrency(valorTotal)
          };
        })
      };

      // Registrar cuántas filas dinámicas se están enviando
      console.log(`✅ Preparadas ${solicitudGastos.length} filas dinámicas para insertar en E44:AK44`);
      
    } catch (error) {
      console.error('Error al obtener gastos dinámicos:', error);
      console.error('Stack:', error.stack);
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
    
    // Verificar qué campos están presentes y cuáles no para depuración
    const camposFaltantes = [
      'nombre_actividad', 'fecha_solicitud', 'dia', 'mes', 'anio',
      'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos',
      'subtotal_gastos', 'imprevistos_3%', 'total_gastos_imprevistos',
      'fondo_comun_porcentaje', 'facultadad_instituto_porcentaje', 
      'escuela_departamento_porcentaje', 'total_recursos',
      // También verificar algunos campos de gastos
      'gasto_1_cantidad', 'gasto_1_valor_unit', 'gasto_1_valor_total',
      'gasto_1,2_cantidad', 'gasto_1,2_valor_unit', 'gasto_1,2_valor_total'
    ];
    
    console.log('🔍 VERIFICACIÓN DE CAMPOS CRÍTICOS:');
    camposFaltantes.forEach(campo => {
      if (transformedData[campo] === undefined || transformedData[campo] === '') {
        console.log(`❌ FALTA: ${campo}`);
      } else {
        console.log(`✅ OK: ${campo} = "${transformedData[campo]}"`);
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
    
    // Procesar TODOS los IDs de gastos, no solo los específicos
    const camposGastos = Object.keys(transformedData).filter(
      key => key.startsWith('gasto_') && key.endsWith('_cantidad')
    );

    // Extraer todos los IDs únicos de los campos de gastos
    const todosLosIds = new Set();
    camposGastos.forEach(campo => {
      // Extraer el ID del patrón 'gasto_ID_cantidad'
      const partes = campo.split('_');
      if (partes.length >= 3) {
        // Extraer todo lo que está entre 'gasto_' y '_cantidad'
        const idCompleto = campo.substring(6, campo.length - 9);
        todosLosIds.add(idCompleto);
      }
    });

    console.log(`🔍 Procesando ${todosLosIds.size} IDs de gastos`);

    // Procesar cada ID y asegurar que exista en ambos formatos (punto y coma)
    todosLosIds.forEach(id => {
      // Si el ID contiene un punto, crear también la versión con coma
      if (id.includes('.')) {
        const idConComa = id.replace('.', ',');
        
        // Verificar si existe la versión con punto pero falta la versión con coma
        if (transformedData[`gasto_${id}_cantidad`] && !transformedData[`gasto_${idConComa}_cantidad`]) {
          console.log(`🔄 Generando formato con coma para ID: ${id} -> ${idConComa}`);
          transformedData[`gasto_${idConComa}_cantidad`] = transformedData[`gasto_${id}_cantidad`];
          transformedData[`gasto_${idConComa}_valor_unit`] = transformedData[`gasto_${id}_valor_unit`];
          transformedData[`gasto_${idConComa}_valor_total`] = transformedData[`gasto_${id}_valor_total`];
          transformedData[`gasto_${idConComa}_descripcion`] = transformedData[`gasto_${id}_descripcion`] || '';
        }
      }
      // Si el ID contiene una coma, crear también la versión con punto
      else if (id.includes(',')) {
        const idConPunto = id.replace(',', '.');
        
        // Verificar si existe la versión con coma pero falta la versión con punto
        if (transformedData[`gasto_${id}_cantidad`] && !transformedData[`gasto_${idConPunto}_cantidad`]) {
          console.log(`🔄 Generando formato con punto para ID: ${id} -> ${idConPunto}`);
          transformedData[`gasto_${idConPunto}_cantidad`] = transformedData[`gasto_${id}_cantidad`];
          transformedData[`gasto_${idConPunto}_valor_unit`] = transformedData[`gasto_${id}_valor_unit`];
          transformedData[`gasto_${idConPunto}_valor_total`] = transformedData[`gasto_${id}_valor_total`];
          transformedData[`gasto_${idConPunto}_descripcion`] = transformedData[`gasto_${id}_descripcion`] || '';
        }
      }
    });

    // Para compatibilidad con el código anterior, mantener la verificación de los IDs específicos
    const idsEspecificos = ['1.3', '1.2', '10'];
    idsEspecificos.forEach(id => {
      const idConComa = id.replace('.', ',');
      
      // Verificar si existen los campos para estos IDs específicos
      if (!transformedData[`gasto_${id}_cantidad`] && !transformedData[`gasto_${idConComa}_cantidad`]) {
        console.warn(`⚠️ ID específico ${id} no encontrado en los datos procesados`);
      } else {
        console.log(`✅ ID específico ${id} procesado correctamente`);
      }
    });
    
    // Imprimir datos finales transformados para depuración
    console.log("⭐ DATOS TRANSFORMADOS FINALES - FORM 2:", transformedData);
    return transformedData;
  },
  
  // Configuración adicional específica para Google Sheets
  sheetsConfig: {
    sheetName: 'Formulario2',
    dataRange: 'A1:Z100'
  },
  footerText: 'Universidad del Valle - Extensión y Proyección Social - Presupuesto',
  watermark: false
};

// Ejemplo de código a implementar en el servicio que genera los reportes
// Este código detectaría la propiedad __FILAS_DINAMICAS__ y procesaría las filas dinámicas

// Si se detectan filas dinámicas, insertarlas en la posición especificada
if (data['__FILAS_DINAMICAS__']) {
  const filasDinamicas = data['__FILAS_DINAMICAS__'];
  const insertarEn = filasDinamicas.insertarEn || 'E44:AK44';
  const gastos = filasDinamicas.gastos || [];
  
  if (gastos.length > 0) {
    console.log(`Insertando ${gastos.length} filas dinámicas en ${insertarEn}`);
    
    // Código para insertar las filas en la posición especificada
    // Esto dependerá de si estás usando Google Sheets API, Excel.js u otra biblioteca
    
    // Por ejemplo, si estás usando Google Sheets API:
    await sheetsApi.spreadsheets.values.update({
      spreadsheetId: reporteId,
      range: insertarEn,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: gastos.map(gasto => [
          gasto.descripcion,
          gasto.cantidad.toString(),
          gasto.valor_unit_formatted,
          gasto.valor_total_formatted
          // Agrega más columnas según sea necesario
        ])
      }
    });
  }
  
  // Eliminar la propiedad especial para que no cause problemas
  delete data['__FILAS_DINAMICAS__'];
}