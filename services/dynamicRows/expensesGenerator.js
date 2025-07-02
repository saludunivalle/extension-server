/**
 * Expenses (Gastos) Generator
 * Processes expense data and prepares it for insertion as dynamic rows
 */

const fs = require('fs');
const path = require('path');
const templateMapper = require('./templateMapperGastos');
const { google } = require('googleapis');
const { jwtClient } = require('../../config/google');

// Load template configuration
let templateConfig;
try {
  const templateFile = path.join(__dirname, '../../templates/expenses.json');
  templateConfig = JSON.parse(fs.readFileSync(templateFile, 'utf8'));
  console.log('Template configuration loaded successfully for expenses');
} catch (error) {
  console.error('Error loading template configuration for expenses:', error);
  templateConfig = {
    templateRow: {
      range: "A42:AK42",
      copyStyle: true,
      insertStartRow: 43
    },
    columns: {
      id: { column: "E" },
      description: { column: "F", span: 17 },
      quantity: { column: "X" },
      unitValue: { column: "Z" },
      totalValue: { column: "AC" }
    }
  };
}

/**
 * Dynamic rows generator for expenses
 */

/**
 * Generates dynamic rows for expense data
 * @param {Array|Object} expenses - Array of expense objects or a single expense object
 * @returns {Object} Object with rows data for Google Sheets API
 */
const generateRows = (expenses) => {
  try {
    if (!expenses) {
      console.log('⚠️ No expenses provided to generateRows');
      return null;
    }
    
    // Ensure expenses is an array
    const expensesArray = Array.isArray(expenses) ? expenses : [expenses];
    
    if (expensesArray.length === 0) {
      console.log('⚠️ Empty expenses array provided to generateRows');
      return null;
    }
    
    console.log(`🔄 Generating rows for ${expensesArray.length} expenses`);
    
    // Normalize property names to ensure compatibility with different formats
    const normalizedExpenses = expensesArray.map(gasto => {
      // Log original data for debugging
      console.log(`🔍 DATOS ORIGINALES del gasto:`, {
        id: gasto.id,
        concepto: gasto.concepto,
        descripcion: gasto.descripcion,
        cantidad: gasto.cantidad,
        valorUnit: gasto.valorUnit,
        valorTotal: gasto.valorTotal
      });
      
      // Create a normalized expense object
      const normalized = {
        id: gasto.id || gasto.id_concepto || gasto.id_conceptos || '',
        descripcion: gasto.descripcion || gasto.concepto || gasto.name || '',
        cantidad: parseFloat(gasto.cantidad) || 0,
        valorUnit: parseFloat(gasto.valorUnit || gasto.valor_unit) || 0,
        valorTotal: parseFloat(gasto.valorTotal || gasto.valor_total) || 0,
        valorUnit_formatted: gasto.valorUnit_formatted || gasto.valor_unit_formatted || '',
        valorTotal_formatted: gasto.valorTotal_formatted || gasto.valor_total_formatted || ''
      };
      
      // Log normalized data for debugging
      console.log(`✅ DATOS NORMALIZADOS del gasto:`, normalized);
      
      return normalized;
    });
    
    // Map each expense to the format expected by the template
    const rows = normalizedExpenses.map(gasto => {
      const row = templateMapper.createRow(gasto);
      console.log(`Fila mapeada para ${gasto.id}: ${JSON.stringify(row)}`);
      return row;
    });
    
    // Return formatted object for API
    return {
      gastos: normalizedExpenses,
      rows: rows,
      templateRow: templateMapper.defaultInsertLocation,
      insertStartRow: 43 // Fixed default value - updated to match template
    };
  } catch (error) {
    console.error('Error al generar filas dinámicas para gastos:', error);
    console.error('Stack trace:', error.stack);
    return null;
  }
};

/**
 * Versión optimizada de insertDynamicRows en expensesGenerator.js
 * con soporte para formato de ID con coma y corrección de celdas combinadas
 */
const insertDynamicRows = async (fileId, dynamicRowsData) => {
  try {
    if (!dynamicRowsData || !dynamicRowsData.gastos || dynamicRowsData.gastos.length === 0) {
      console.log('No hay datos para insertar filas dinámicas');
      return false;
    }
    
    // Verificar que tenemos acceso a las dependencias necesarias
    if (!google || !jwtClient) {
      console.error('Dependencias no disponibles: google o jwtClient');
      return false;
    }
    
    // Inicializar API de Google Sheets
    const sheets = google.sheets({version: 'v4', auth: jwtClient});
    
    // Extraer información esencial para inserción
    const rowsData = dynamicRowsData.gastos; // Datos originales de gastos
    const templateRange = dynamicRowsData.insertarEn || "A42:AK42"; // Rango de template
    const insertStartRow = dynamicRowsData.insertStartRow || 43; // Fila para comenzar inserción
    
    console.log(`Insertando ${rowsData.length} filas dinámicas en la hoja ${fileId}`);
    console.log(`Configuración: templateRange=${templateRange}, insertStartRow=${insertStartRow}`);
    
    // 1. PASO 1: Obtener información del documento
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: fileId,
      includeGridData: true,
      ranges: [templateRange]
    });
    
    // Verificar que obtuvimos datos del documento
    if (!spreadsheet.data || !spreadsheet.data.sheets || spreadsheet.data.sheets.length === 0) {
      console.error('No se pudo obtener información de la hoja');
      return false;
    }
    
    // Obtener ID de la primera hoja
    const sheetId = spreadsheet.data.sheets[0].properties.sheetId;
    console.log(`ID de la hoja: ${sheetId}`);
    
    // Extraer información sobre celdas combinadas
    const merges = spreadsheet.data.sheets[0].merges || [];
    console.log(`Encontradas ${merges.length} regiones combinadas en la hoja`);
    
    // 2. PASO 2: Insertar filas vacías
    const insertRowsRequest = {
      spreadsheetId: fileId,
      resource: {
        requests: [{
          insertDimension: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: insertStartRow - 1, // Convertir a 0-based
              endIndex: (insertStartRow - 1) + rowsData.length
            },
            inheritFromBefore: true
          }
        }]
      }
    };
    
    await sheets.spreadsheets.batchUpdate(insertRowsRequest);
    console.log(`✅ ${rowsData.length} filas vacías insertadas en la posición ${insertStartRow}`);
    
    // 3. PASO 3: Extraer información del rango template
    const rangeMatch = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(templateRange);
    if (!rangeMatch) {
      throw new Error(`Formato de rango inválido: ${templateRange}`);
    }
    
    const startCol = rangeMatch[1];
    const endCol = rangeMatch[3];
    const templateRowNum = parseInt(rangeMatch[2]);
    
    // Función auxiliar para convertir columna a índice
    const colToIndex = (col) => {
      let index = 0;
      for (let i = 0; i < col.length; i++) {
        index = index * 26 + col.charCodeAt(i) - 64;
      }
      return index - 1; // 0-based
    };
    
    const startColIndex = colToIndex(startCol);
    const endColIndex = colToIndex(endCol) + 1;
    
    // 4. PASO 4: Copiar formato de la fila template
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: fileId,
      resource: {
        requests: [{
          copyPaste: {
            source: {
              sheetId: sheetId,
              startRowIndex: templateRowNum - 1, // Convertir a 0-based
              endRowIndex: templateRowNum,
              startColumnIndex: startColIndex,
              endColumnIndex: endColIndex
            },
            destination: {
              sheetId: sheetId,
              startRowIndex: insertStartRow - 1, // Convertir a 0-based
              endRowIndex: (insertStartRow - 1) + rowsData.length,
              startColumnIndex: startColIndex,
              endColumnIndex: endColIndex
            },
            pasteType: 'PASTE_FORMAT',
            pasteOrientation: 'NORMAL'
          }
        }]
      }
    });
    
    console.log(`✅ Formato copiado correctamente`);
    
    // 5. PASO 5: Recrear celdas combinadas en las nuevas filas
    // Identificar las celdas combinadas que afectan a la fila de plantilla
    const templateRowMerges = merges.filter(merge => {
      return merge.startRowIndex === templateRowNum - 1 && merge.endRowIndex === templateRowNum;
    });
    
    if (templateRowMerges.length > 0) {
      console.log(`Encontradas ${templateRowMerges.length} celdas combinadas en la fila plantilla`);
      
      const mergeCellRequests = [];
      
      // Para cada fila insertada
      for (let i = 0; i < rowsData.length; i++) {
        // Para cada combinación de celdas en la plantilla
        templateRowMerges.forEach(merge => {
          mergeCellRequests.push({
            mergeCells: {
              range: {
                sheetId: sheetId,
                startRowIndex: insertStartRow - 1 + i,
                endRowIndex: insertStartRow + i,
                startColumnIndex: merge.startColumnIndex,
                endColumnIndex: merge.endColumnIndex
              },
              mergeType: 'MERGE_ALL'
            }
          });
        });
      }
      
      if (mergeCellRequests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: fileId,
          resource: {
            requests: mergeCellRequests
          }
        });
        console.log(`✅ Creadas ${mergeCellRequests.length} celdas combinadas en las nuevas filas`);
      }
    }
    
    // 6. PASO 6: Insertar datos en las celdas
    const valueRequests = [];
    
    for (let i = 0; i < rowsData.length; i++) {
      const rowIndex = insertStartRow + i;
      const gasto = rowsData[i];
      
      // Manejar ambos formatos de ID (con coma o con punto)
      let idValue = gasto.id || gasto.id_conceptos || `8,${i+1}`;
      
      // Convertir formato a coma si viene con punto
      if (idValue.includes('.')) {
        idValue = idValue.replace('.', ',');
      }
      
      // Extraer valores con manejo de diferentes nombres de propiedades
      const descripcionValue = gasto.descripcion || gasto.concepto || '';
      const cantidadValue = gasto.cantidad?.toString() || '0';
      const valorUnitValue = gasto.valorUnit_formatted || gasto.valor_unit_formatted || `$${gasto.valorUnit || gasto.valor_unit || 0}`;
      const valorTotalValue = gasto.valorTotal_formatted || gasto.valor_total_formatted || `$${gasto.valorTotal || gasto.valor_total || 0}`;
      
      // Log para depuración
      console.log(`Fila ${rowIndex}, ID: ${idValue}, Descripción: ${descripcionValue}, Cantidad: ${cantidadValue}, ValorUnit: ${valorUnitValue}, ValorTotal: ${valorTotalValue}`);
      
      // Añadir solicitudes para cada celda
      valueRequests.push({
        range: `E${rowIndex}`, // ID
        values: [[idValue]]
      });
      
      valueRequests.push({
        range: `F${rowIndex}`, // Descripción
        values: [[descripcionValue]]
      });
      
      valueRequests.push({
        range: `X${rowIndex}`, // Cantidad
        values: [[cantidadValue]]
      });
      
      valueRequests.push({
        range: `Z${rowIndex}`, // Valor unitario
        values: [[valorUnitValue]]
      });
      
      valueRequests.push({
        range: `AC${rowIndex}`, // Valor total
        values: [[valorTotalValue]]
      });
    }
    
    // Enviar solicitud de actualización de valores
    if (valueRequests.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: fileId,
        resource: {
          valueInputOption: 'USER_ENTERED',
          data: valueRequests
        }
      });
      console.log(`✅ Datos insertados correctamente en las celdas`);
    } else {
      console.log('⚠️ No hay datos para insertar');
    }
    
    return true;
  } catch (error) {
    console.error('Error al insertar filas dinámicas:', error);
    if (error.response && error.response.data) {
      console.error('Detalles del error Google Sheets API:', error.response.data);
    }
    return false;
  }
};

module.exports = {
  generateRows,
  insertDynamicRows
};
