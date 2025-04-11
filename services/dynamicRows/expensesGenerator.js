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
      range: "A44:AK44",
      copyStyle: true,
      insertStartRow: 45
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
 * Generate formatted rows from expense data
 * @param {Array} gastos - Array of expense objects
 * @param {String} insertLocation - Optional custom insert location
 * @returns {Object} Formatted data for dynamic rows
 */
const generateRows = (gastos, insertLocation = null) => {
  try {
    if (!gastos || !Array.isArray(gastos) || gastos.length === 0) {
      console.log('No hay gastos para generar filas dinámicas');
      return null;
    }
    
    console.log(`Analizando ${gastos.length} gastos para generar filas dinámicas`);
    
    // Before filtering, log all expense IDs for debugging
    console.log('IDs de gastos disponibles:', gastos.map(g => g.id_concepto || g.id).join(', '));
    
    // Filter only dynamic expenses (those with ID starting with 15.)
    const gastosDinamicos = gastos.filter(gasto => {
      const id = gasto.id_concepto || gasto.id || '';
      const isDynamic = typeof id === 'string' && id.startsWith('15.');
      
      // Log result for each item for debugging
      if (isDynamic) {
        console.log(`✅ Gasto dinámico identificado: ID=${id}`);
      }
      
      return isDynamic;
    });
    
    if (gastosDinamicos.length === 0) {
      console.log('⚠️ No hay gastos dinámicos para procesar (IDs que empiecen con 15.)');
      return null;
    }
    
    console.log(`Procesando ${gastosDinamicos.length} gastos dinámicos`);
    
    // Map each expense to the format expected by the template
    const rows = gastosDinamicos.map(gasto => {
      const row = templateMapper.createRow(gasto);
      console.log(`Fila mapeada para ${gasto.id || gasto.id_concepto}: ${JSON.stringify(row)}`);
      return row;
    });
    
    // CRITICAL: Get template configuration details
    // Always use the fixed value 45 for insertStartRow as required
    const templateRow = templateConfig?.templateRow?.range || "A44:AK44";
    const insertStartRow = 45; // FIXED VALUE: Always insert at row 45 as required
    
    // Default insert location from template config - use template row as default
    const defaultInsert = insertLocation || templateRow;
    
    console.log(`⚠️ CRÍTICO: Filas dinámicas se insertarán en la fila ${insertStartRow}`);
    console.log(`Configuración de inserción: Template=${templateRow}, Fila inicial=${insertStartRow}`);
    
    return {
      insertarEn: defaultInsert,
      gastos: gastosDinamicos,
      rows: rows,
      templateConfig: {
        ...templateConfig,
        templateRow: {
          ...templateConfig?.templateRow,
          insertStartRow: insertStartRow // Force row 45 as insert position
        }
      },
      insertStartRow: insertStartRow // Explicitly include the start row
    };
  } catch (error) {
    console.error('Error al generar filas dinámicas:', error);
    return null;
  }
};

/**
 * Insert dynamic expense rows into a Google Sheet
 * @param {String} fileId - ID of the Google Sheet
 * @param {Object} dynamicRowsData - Processed data with rows to insert
 * @returns {Promise<Boolean>} Success status
 */
const insertDynamicRows = async (fileId, dynamicRowsData) => {
  try {
    if (!dynamicRowsData || !dynamicRowsData.rows || dynamicRowsData.rows.length === 0) {
      console.log('No hay datos para insertar filas dinámicas');
      return false;
    }
    
    // CRITICAL: Force the insertStartRow to be 45 always
    const insertStartRow = 45; // FIXED VALUE: Always insert at row 45 as required
    
    // Get template configuration
    const config = {
      ...dynamicRowsData.templateConfig || templateConfig,
      templateRow: {
        ...dynamicRowsData.templateConfig?.templateRow || templateConfig?.templateRow,
        insertStartRow: insertStartRow // Force row 45
      }
    };
    
    // Use the configured template row range
    const templateRange = config?.templateRow?.range || "A44:AK44";
    const copyStyle = config?.templateRow?.copyStyle !== false; // Default to true
    
    const rowsData = dynamicRowsData.gastos; // Use original data for insertion
    console.log(`⚠️ CRÍTICO: Insertando ${rowsData.length} filas dinámicas en la posición ${insertStartRow} (fila 45)`);
    console.log(`Basadas en template ${templateRange}`);
    
    // Initialize Google Sheets API
    const sheets = google.sheets({version: 'v4', auth: jwtClient});
    
    // Step 1: Get the spreadsheet info to identify sheet ID
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: fileId,
      fields: 'sheets.properties'
    });
    
    if (!spreadsheet.data.sheets || spreadsheet.data.sheets.length === 0) {
      throw new Error('No sheets found in the spreadsheet');
    }
    
    // Use the first sheet ID
    const sheetId = spreadsheet.data.sheets[0].properties.sheetId;
    const sheetName = spreadsheet.data.sheets[0].properties.title;
    console.log(`Target sheet: "${sheetName}" (ID: ${sheetId})`);
    
    // Step 2: Insert empty rows at the specified position (always 45)
    console.log(`⚠️ CRÍTICO: Insertando filas vacías en la posición EXACTA ${insertStartRow}`);
    await insertEmptyRows(sheets, fileId, insertStartRow, rowsData.length);
    console.log(`✅ Insertadas ${rowsData.length} filas vacías en la posición ${insertStartRow}`);
    
    // Step 3: If style copying is enabled, copy the template row styles to each new row
    if (copyStyle) {
      console.log(`Copiando estilos desde la fila template (44) a las filas dinámicas comenzando en la fila ${insertStartRow}`);
      await copyTemplateRowStyles(sheets, fileId, templateRange, insertStartRow, rowsData.length);
      console.log(`✅ Copiados estilos y celdas combinadas a ${rowsData.length} filas`);
    }
    
    // Step 4: Now insert the data into specific cells based on the column configuration
    console.log(`Insertando datos en las celdas correspondientes para ${rowsData.length} filas dinámicas`);
    const dataMapping = config?.dataMapping || {};
    const columnDefs = config?.columns || {};
    
    // Create separate batch requests for each row to avoid issues with merged cells
    for (let i = 0; i < rowsData.length; i++) {
      const rowNum = insertStartRow + i;
      const gasto = rowsData[i];
      const requests = [];
      
      console.log(`Llenando datos para fila ${rowNum}: ${gasto.id || gasto.id_concepto} - ${gasto.descripcion || ''}`);
      
      // Prepare data updates for each column defined in the template
      Object.keys(columnDefs).forEach(colKey => {
        const columnInfo = columnDefs[colKey];
        const dataKey = dataMapping[colKey]; // Get the corresponding key in the gasto object
        let valueToInsert = '';
        
        // Get the value to insert from the gasto object
        if (dataKey && gasto[dataKey] !== undefined) {
          valueToInsert = gasto[dataKey]?.toString() || '';
        } else if (gasto[colKey] !== undefined) { // Fallback to direct key match
          valueToInsert = gasto[colKey]?.toString() || '';
        }
        
        // Handle column range (e.g., "F:V")
        let startColumnLetter, endColumnLetter;
        if (columnInfo.column.includes(':')) {
          [startColumnLetter, endColumnLetter] = columnInfo.column.split(':');
        } else {
          startColumnLetter = columnInfo.column;
          
          // If span is defined, calculate the end column
          if (columnInfo.span) {
            // Simple span calculation based on ASCII codes
            const startCode = startColumnLetter.charCodeAt(0);
            endColumnLetter = String.fromCharCode(startCode + columnInfo.span - 1);
          } else {
            endColumnLetter = startColumnLetter;
          }
        }
        
        // Construct the cell range for this value
        const cellRange = `${startColumnLetter}${rowNum}`;
        
        // Add this cell update to the batch
        requests.push({
          range: cellRange,
          values: [[valueToInsert]]
        });
        
        console.log(`  - Campo "${colKey}" -> Celda ${cellRange}: "${valueToInsert}"`);
      });
      
      // Execute batch update for this row
      if (requests.length > 0) {
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: fileId,
          resource: {
            valueInputOption: 'USER_ENTERED',
            data: requests
          }
        });
      }
    }
    
    console.log(`✅ Insertados datos en ${rowsData.length} filas dinámicas en el reporte`);
    return true;
  } catch (error) {
    console.error('Error al insertar filas dinámicas:', error);
    if (error.response && error.response.data) {
      console.error('Google API Error:', JSON.stringify(error.response.data, null, 2));
    }
    return false;
  }
};

/**
 * Insert empty rows at specified position
 * @param {Object} sheets - Google Sheets API client
 * @param {String} fileId - Spreadsheet ID
 * @param {Number} startRowIndex - 1-based row index where to insert
 * @param {Number} rowCount - Number of rows to insert
 */
async function insertEmptyRows(sheets, fileId, startRowIndex, rowCount) {
  try {
    // CRITICAL: Always force insertion at row 45 for dynamic expenses
    if (startRowIndex !== 45) {
      console.warn(`⚠️ ADVERTENCIA: Se intentó insertar en posición ${startRowIndex}, forzando a fila 45`);
      startRowIndex = 45;
    }
    
    console.log(`⚠️ CRÍTICO: Insertando ${rowCount} filas vacías en la posición ${startRowIndex} (1-based, fila 45)`);
    
    // Get the spreadsheet to find the first sheet ID
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: fileId,
      fields: 'sheets.properties'
    });
    
    if (!spreadsheet.data.sheets || spreadsheet.data.sheets.length === 0) {
      throw new Error('No sheets found in the spreadsheet');
    }
    
    // Get sheet details
    const sheet = spreadsheet.data.sheets[0];
    const sheetId = sheet.properties.sheetId;
    const sheetName = sheet.properties.title;
    
    console.log(`Insertando filas en hoja "${sheetName}" (ID: ${sheetId})`);
    
    // Convert from 1-based to 0-based index for Google Sheets API
    let zeroBasedIndex = startRowIndex - 1; // For row 45, this should be 44
    console.log(`Índice convertido a 0-based: ${zeroBasedIndex} (debe ser 44 para insertar en fila 45)`);
    
    // Validate zero-based index is 44 (which corresponds to row 45)
    if (zeroBasedIndex !== 44) {
      console.error(`⚠️ ERROR CRÍTICO: Índice 0-based incorrecto (${zeroBasedIndex}), debe ser 44 para insertar en fila 45`);
      // Force to correct index
      zeroBasedIndex = 44;
    }
    
    // Create the insert dimension request
    const request = {
      insertDimension: {
        range: {
          sheetId: sheetId,
          dimension: 'ROWS',
          startIndex: zeroBasedIndex, // Should be 44 for row 45
          endIndex: zeroBasedIndex + rowCount
        },
        inheritFromBefore: true
      }
    };
    
    // Log the request details
    console.log('Solicitud de inserción:', JSON.stringify(request, null, 2));
    
    // Execute the request
    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: fileId,
      resource: {
        requests: [request]
      }
    });
    
    console.log(`✅ Filas insertadas correctamente en la posición ${startRowIndex} (1-based)`);
    return true;
  } catch (error) {
    console.error('Error inserting empty rows:', error);
    console.error('Error details:', error.response?.data || error.message);
    throw error;
  }
}

/**
 * Copy template row styles to newly inserted rows
 * @param {Object} sheets - Google Sheets API client
 * @param {String} fileId - Spreadsheet ID
 * @param {String} templateRange - Template row range (e.g., "A44:AK44")
 * @param {Number} startRowIndex - 1-based row index where rows were inserted
 * @param {Number} rowCount - Number of rows inserted
 */
async function copyTemplateRowStyles(sheets, fileId, templateRange, startRowIndex, rowCount) {
  try {
    // Get the spreadsheet details including merged cells info
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: fileId,
      fields: 'sheets.properties,sheets.merges,sheets.data.rowData.values.effectiveFormat'
    });
    
    if (!spreadsheet.data.sheets || spreadsheet.data.sheets.length === 0) {
      throw new Error('No sheets found in the spreadsheet');
    }
    
    // Use the first sheet ID
    const sheetId = spreadsheet.data.sheets[0].properties.sheetId;
    
    // Parse the template range
    const match = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(templateRange);
    if (!match) {
      throw new Error(`Invalid template range format: ${templateRange}`);
    }
    
    const [_, startCol, startRow, endCol, endRow] = match;
    const templateRowIndex = parseInt(startRow) - 1; // Convert to 0-based
    
    // Find merged cells in the template row
    const templateMerges = [];
    if (spreadsheet.data.sheets[0].merges) {
      // Log all merges for debugging
      console.log(`Found ${spreadsheet.data.sheets[0].merges.length} total merged cells in sheet`);
      
      spreadsheet.data.sheets[0].merges.forEach(mergeInfo => {
        // Check if the merge range intersects with our template row
        if (mergeInfo.startRowIndex <= templateRowIndex && 
            mergeInfo.endRowIndex > templateRowIndex) {
          templateMerges.push({
            startColumnIndex: mergeInfo.startColumnIndex,
            endColumnIndex: mergeInfo.endColumnIndex,
            // Keep track of the column span for this merge
            columnSpan: mergeInfo.endColumnIndex - mergeInfo.startColumnIndex
          });
          
          console.log(`Template merge found: columns ${mergeInfo.startColumnIndex} to ${mergeInfo.endColumnIndex}`);
        }
      });
    }
    
    console.log(`Found ${templateMerges.length} merged cell ranges in template row`);
    
    // Create copy paste and merge requests for each target row
    const requests = [];
    
    // For each new row
    for (let i = 0; i < rowCount; i++) {
      const targetRowIndex = startRowIndex - 1 + i; // Convert to 0-based
      
      // 1. First copy all formatting and styles
      requests.push({
        copyPaste: {
          source: {
            sheetId: sheetId,
            startRowIndex: templateRowIndex,
            endRowIndex: templateRowIndex + 1,
            startColumnIndex: columnToIndex(startCol),
            endColumnIndex: columnToIndex(endCol) + 1
          },
          destination: {
            sheetId: sheetId,
            startRowIndex: targetRowIndex,
            endRowIndex: targetRowIndex + 1,
            startColumnIndex: columnToIndex(startCol),
            endColumnIndex: columnToIndex(endCol) + 1
          },
          pasteType: 'PASTE_FORMAT',
          pasteOrientation: 'NORMAL'
        }
      });
      
      // 2. Create the same merged cells in the destination row
      templateMerges.forEach(merge => {
        requests.push({
          mergeCells: {
            range: {
              sheetId: sheetId,
              startRowIndex: targetRowIndex,
              endRowIndex: targetRowIndex + 1,
              startColumnIndex: merge.startColumnIndex,
              endColumnIndex: merge.endColumnIndex
            },
            mergeType: 'MERGE_ALL'
          }
        });
      });
    }
    
    // Execute the batch update
    if (requests.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: fileId,
        resource: { requests }
      });
      
      console.log(`✅ Successfully copied template styles and merged cells to ${rowCount} rows`);
    }
    
    return true;
  } catch (error) {
    console.error('Error copying template styles:', error);
    throw error;
  }
}

/**
 * Extract column range from a cell range like "A44:AK44"
 * @param {String} range - Cell range
 * @returns {Array} Array with start and end column letters
 */
function extractColumnRange(range) {
  const match = /([A-Z]+)\d+:([A-Z]+)\d+/.exec(range);
  if (!match) {
    return ['A', 'Z']; // Default fallback
  }
  return [match[1], match[2]];
}

/**
 * Convert column letter to index (0-based)
 * @param {String} column - Column letter (e.g., "A", "BC")
 * @returns {Number} Column index (0-based)
 */
function columnToIndex(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result - 1; // Convert to 0-based index
}

module.exports = {
  generateRows,
  insertDynamicRows
};
