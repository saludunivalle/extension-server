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
    
    // Generar campos dinámicos para cada concepto y subconcepto
    try {
      // Importar el servicio de sheetsService dinámicamente (para evitar problemas con ES modules)
      const sheetsService = require('../services/sheetsService');
      
      // Obtener todos los gastos para esta solicitud
      const gastosResponse = await sheetsService.client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'GASTOS!A2:F'
      });
      
      const gastosRows = gastosResponse.data.values || [];
      const solicitudGastos = gastosRows.filter(row => row[1] === formData.id_solicitud);
      
      console.log(`Encontrados ${solicitudGastos.length} gastos para la solicitud ${formData.id_solicitud}`);
      
      // Organizar gastos por categoría principal
      const gastosPorCategoria = {};
      const gastosPlanos = {};
      
      solicitudGastos.forEach(gasto => {
        const idConcepto = gasto[0]; // id_conceptos
        const cantidad = parseFloat(gasto[2]) || 0;
        const valorUnit = parseFloat(gasto[3]) || 0;
        const valorTotal = parseFloat(gasto[4]) || 0;
        const conceptoPadre = gasto[5]; // concepto_padre
        
        // Determinar si es un concepto principal o subconcepto
        const esPrincipal = !idConcepto.includes('.');
        const categoriaId = esPrincipal ? idConcepto : idConcepto.split('.')[0];
        
        // Inicializar la categoría si no existe
        if (!gastosPorCategoria[categoriaId]) {
          gastosPorCategoria[categoriaId] = {
            principal: null,
            subconceptos: []
          };
        }
        
        // Guardar el gasto en su lugar correcto
        if (esPrincipal) {
          gastosPorCategoria[categoriaId].principal = {
            id: idConcepto,
            cantidad,
            valorUnit, 
            valorTotal
          };
        } else {
          gastosPorCategoria[categoriaId].subconceptos.push({
            id: idConcepto,
            cantidad,
            valorUnit,
            valorTotal
          });
        }
        
        // También guardar en formato plano para referencia directa
        gastosPlanos[idConcepto] = {
          cantidad,
          valorUnit,
          valorTotal
        };
      });
      
      // Generar campos dinámicos para cada concepto y subconcepto
      const conceptosOrdenados = Object.keys(gastosPorCategoria).sort((a, b) => parseInt(a) - parseInt(b));
      
      conceptosOrdenados.forEach(categoriaId => {
        const categoria = gastosPorCategoria[categoriaId];
        
        // Concepto principal
        if (categoria.principal) {
          const idConcepto = categoria.principal.id;
          transformedData[`gasto_${idConcepto}_cantidad`] = categoria.principal.cantidad.toString();
          transformedData[`gasto_${idConcepto}_valor_unit`] = formatCurrency(categoria.principal.valorUnit);
          transformedData[`gasto_${idConcepto}_valor_total`] = formatCurrency(categoria.principal.valorTotal);
        }
        
        // Subconceptos
        categoria.subconceptos.forEach(subconcepto => {
          const idSubConcepto = subconcepto.id;
          transformedData[`gasto_${idSubConcepto}_cantidad`] = subconcepto.cantidad.toString();
          transformedData[`gasto_${idSubConcepto}_valor_unit`] = formatCurrency(subconcepto.valorUnit);
          transformedData[`gasto_${idSubConcepto}_valor_total`] = formatCurrency(subconcepto.valorTotal);
        });
      });
      
      console.log(`Generados ${Object.keys(gastosPlanos).length} campos de gastos dinámicos`);
      
    } catch (error) {
      console.error('Error al obtener gastos dinámicos:', error);
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