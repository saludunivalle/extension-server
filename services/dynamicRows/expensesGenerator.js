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
      range: "A43:AK43",
      copyStyle: true,
      insertStartRow: 44
    },
    columns: {
      id: { column: "E" },
      description: { column: "F", span: 17 },
      quantity: { column: "W:Y" },
      unitValue: { column: "Z" },
      totalValue: { column: "AC" }
    }
  };
}

/**
 * Dynamic rows generator for expenses
 */

function parseConceptIdParts(idValue) {
  return String(idValue || '')
    .trim()
    .replace(/,/g, '.')
    .split('.')
    .map((segment) => {
      const n = parseInt(segment, 10);
      return Number.isNaN(n) ? 0 : n;
    });
}

function compareDynamicConceptIds(a, b) {
  const aParts = parseConceptIdParts(a?.id || a?.id_conceptos || a?.id_concepto);
  const bParts = parseConceptIdParts(b?.id || b?.id_conceptos || b?.id_concepto);
  const maxLen = Math.max(aParts.length, bParts.length);

  for (let i = 0; i < maxLen; i++) {
    const av = aParts[i] || 0;
    const bv = bParts[i] || 0;
    if (av !== bv) {
      return av - bv;
    }
  }

  return 0;
}

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
    
    const sortedExpenses = [...normalizedExpenses].sort(compareDynamicConceptIds);

    // Map each expense to the format expected by the template
    const rows = sortedExpenses.map(gasto => {
      const row = templateMapper.createRow(gasto);
      console.log(`Fila mapeada para ${gasto.id}: ${JSON.stringify(row)}`);
      return row;
    });
    
    // Return formatted object for API
    return {
      gastos: sortedExpenses,
      rows: rows,
      templateRow: templateMapper.defaultInsertLocation,
      insertStartRow: 44 // Fixed default value - place rows before SUB TOTAL
    };
  } catch (error) {
    console.error('Error al generar filas dinámicas para gastos:', error);
    console.error('Stack trace:', error.stack);
    return null;
  }
};

/**
 * Detects insertion anchor rows directly from sheet content.
 * It locates concept 14 and inserts children before SUB TOTAL GASTOS.
 */
const detectExpenseInsertLocation = async (sheets, spreadsheetId) => {
  try {
    const valuesResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: 'A1:AK200'
    });

    const values = valuesResponse?.data?.values || [];
    let concept14Row = null;
    let subtotalRow = null;

    for (let i = 0; i < values.length; i++) {
      const row = values[i] || [];
      const idCell = String(row[4] || '').trim().replace(',', '.'); // Columna E
      const descCell = String(row[5] || '').trim(); // Columna F

      if (!concept14Row) {
        const idMatches14 = idCell === '14' || /^14(\.0+)?$/.test(idCell);
        const descLooks14 = /^14\s/.test(descCell) || /COSTOS ADMINISTRATIVOS DEL PROYECTO/i.test(descCell);
        if (idMatches14 || descLooks14) {
          concept14Row = i + 1;
        }
      }

      if (!subtotalRow && /SUB\s*TOTAL\s*GASTOS/i.test(descCell)) {
        subtotalRow = i + 1;
      }
    }

    if (concept14Row) {
      return {
        templateRange: `A${concept14Row}:AK${concept14Row}`,
        insertStartRow: concept14Row + 1,
        subtotalRow
      };
    }

    return null;
  } catch (error) {
    console.warn('⚠️ No fue posible detectar ancla dinámica para gastos:', error.message);
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
    let templateRange = dynamicRowsData.insertarEn || "A43:AK43"; // Rango de template
    let insertStartRow = dynamicRowsData.insertStartRow || 44; // Fila para comenzar inserción

    const detectedLocation = await detectExpenseInsertLocation(sheets, fileId);
    if (detectedLocation) {
      templateRange = detectedLocation.templateRange;
      insertStartRow = detectedLocation.insertStartRow;
      console.log('✅ Ancla de gastos detectada automáticamente:', detectedLocation);
    }
    
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
            inheritFromBefore: false
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
    
    const getFallbackMergeRanges = () => {
      const fallback = [];
      const configured = [
        templateConfig?.columns?.description?.column,
        templateConfig?.columns?.quantity?.column,
        templateConfig?.columns?.unitValue?.column,
        templateConfig?.columns?.totalValue?.column
      ].filter(Boolean);

      configured.forEach((columnRange) => {
        const [start, end] = String(columnRange).split(':').map((v) => v.trim());
        if (start && end) {
          fallback.push({
            startColumnIndex: colToIndex(start),
            endColumnIndex: colToIndex(end) + 1
          });
        }
      });

      if (fallback.length === 0) {
        return [
          { startColumnIndex: colToIndex('F'), endColumnIndex: colToIndex('V') + 1 },
          { startColumnIndex: colToIndex('W'), endColumnIndex: colToIndex('Y') + 1 },
          { startColumnIndex: colToIndex('Z'), endColumnIndex: colToIndex('AB') + 1 },
          { startColumnIndex: colToIndex('AC'), endColumnIndex: colToIndex('AK') + 1 }
        ];
      }

      return fallback;
    };

    // 5. PASO 5: Recrear celdas combinadas en las nuevas filas
    // Identificar las celdas combinadas que afectan a la fila de plantilla
    const templateRowMerges = merges.filter(merge => {
      return merge.startRowIndex === templateRowNum - 1 && merge.endRowIndex === templateRowNum;
    });
    const fallbackMergeRanges = getFallbackMergeRanges();

    const mergeRangesByKey = new Map();
    templateRowMerges.forEach((merge) => {
      const key = `${merge.startColumnIndex}-${merge.endColumnIndex}`;
      mergeRangesByKey.set(key, {
        startColumnIndex: merge.startColumnIndex,
        endColumnIndex: merge.endColumnIndex
      });
    });
    fallbackMergeRanges.forEach((merge) => {
      const key = `${merge.startColumnIndex}-${merge.endColumnIndex}`;
      if (!mergeRangesByKey.has(key)) {
        mergeRangesByKey.set(key, merge);
      }
    });

    const mergeRangesToApply = [...mergeRangesByKey.values()];
    if (mergeRangesToApply.length > 0) {
      console.log(`Aplicando ${mergeRangesToApply.length} rangos combinados en filas dinámicas`);

      const mergeCellRequests = [];
      for (let i = 0; i < rowsData.length; i++) {
        mergeRangesToApply.forEach((merge) => {
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

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: fileId,
        resource: {
          requests: mergeCellRequests
        }
      });
      console.log(`✅ Creadas ${mergeCellRequests.length} celdas combinadas en las nuevas filas`);
    }

    // 5.1 PASO EXTRA: centrar solo campos numéricos de gastos dinámicos
    const alignmentRequests = [];
    for (let i = 0; i < rowsData.length; i++) {
      const rowStart = insertStartRow - 1 + i;
      const rowEnd = rowStart + 1;
      [
        { start: 'W', end: 'Y' },
        { start: 'Z', end: 'AB' },
        { start: 'AC', end: 'AK' }
      ].forEach((segment) => {
        alignmentRequests.push({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: rowStart,
              endRowIndex: rowEnd,
              startColumnIndex: colToIndex(segment.start),
              endColumnIndex: colToIndex(segment.end) + 1
            },
            cell: {
              userEnteredFormat: {
                horizontalAlignment: 'CENTER',
                verticalAlignment: 'MIDDLE'
              }
            },
            fields: 'userEnteredFormat(horizontalAlignment,verticalAlignment)'
          }
        });
      });
    }

    if (alignmentRequests.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: fileId,
        resource: {
          requests: alignmentRequests
        }
      });
      console.log(`✅ Alineación centrada aplicada en ${alignmentRequests.length} rangos numéricos dinámicos`);
    }

    // 5.2 PASO EXTRA: formato numérico sin símbolo de moneda para dinámicos
    const numberFormatRequests = [];
    for (let i = 0; i < rowsData.length; i++) {
      const rowStart = insertStartRow - 1 + i;
      const rowEnd = rowStart + 1;

      numberFormatRequests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: rowStart,
            endRowIndex: rowEnd,
            startColumnIndex: colToIndex('W'),
            endColumnIndex: colToIndex('Y') + 1
          },
          cell: {
            userEnteredFormat: {
              numberFormat: {
                type: 'NUMBER',
                pattern: '#,##0'
              }
            }
          },
          fields: 'userEnteredFormat.numberFormat'
        }
      });

      ['Z', 'AC'].forEach((startCol) => {
        const endCol = startCol === 'Z' ? 'AB' : 'AK';
        numberFormatRequests.push({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: rowStart,
              endRowIndex: rowEnd,
              startColumnIndex: colToIndex(startCol),
              endColumnIndex: colToIndex(endCol) + 1
            },
            cell: {
              userEnteredFormat: {
                numberFormat: {
                  type: 'NUMBER',
                  pattern: '#,##0'
                }
              }
            },
            fields: 'userEnteredFormat.numberFormat'
          }
        });
      });
    }

    if (numberFormatRequests.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: fileId,
        resource: {
          requests: numberFormatRequests
        }
      });
      console.log(`✅ Formato numérico sin moneda aplicado en ${numberFormatRequests.length} rangos dinámicos`);
    }
    
    // 6. PASO 6: Insertar datos en las celdas
    const valueRequests = [];
    
    for (let i = 0; i < rowsData.length; i++) {
      const rowIndex = insertStartRow + i;
      const gasto = rowsData[i];
      
      // Manejar ambos formatos de ID (con coma o con punto)
      let idValue = gasto.id || gasto.id_conceptos || `14,${i+1}`;
      
      // Convertir formato a coma si viene con punto
      if (idValue.includes('.')) {
        idValue = idValue.replace('.', ',');
      }
      
      // Extraer valores con manejo de diferentes nombres de propiedades
      const descripcionValue = gasto.descripcion || gasto.concepto || '';
      const cantidadRaw = parseFloat(gasto.cantidad);
      const valorUnitRaw = parseFloat(gasto.valorUnit ?? gasto.valor_unit);
      const valorTotalRaw = parseFloat(gasto.valorTotal ?? gasto.valor_total);

      const cantidadValue = Number.isNaN(cantidadRaw) ? 0 : cantidadRaw;
      const valorUnitValue = Number.isNaN(valorUnitRaw) ? 0 : valorUnitRaw;
      const valorTotalValue = Number.isNaN(valorTotalRaw)
        ? cantidadValue * valorUnitValue
        : valorTotalRaw;
      
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
        range: `W${rowIndex}`, // Cantidad
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
