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
    
    // Filter to only include dynamic expenses (ID starts with '15.')
    const gastosDinamicos = expensesArray.filter(gasto => {
      // Check all possible ID field names to catch all dynamic expenses
      const id = gasto.id_conceptos || gasto.id_concepto || gasto.id || '';
      const isDynamic = typeof id === 'string' && id.startsWith('15.');
      
      // Log result for each item for debugging
      if (isDynamic) {
        console.log(`✅ Gasto dinámico identificado: ID=${id}, Descripción=${gasto.name || gasto.descripcion || gasto.concepto || ''}`);
      } else {
        console.log(`ℹ️ Gasto no dinámico: ID=${id}`);
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
      console.log(`Fila mapeada para ${gasto.id || gasto.id_conceptos || gasto.id_concepto}: ${JSON.stringify(row)}`);
      return row;
    });
    
    // Return formatted object for API
    return {
      gastos: gastosDinamicos,
      rows: rows,
      templateRow: templateMapper.defaultInsertLocation,
      // Fixed value for insert location - critical for proper insertion in reports
      insertStartRow: 45 
    };
  } catch (error) {
    console.error('Error al generar filas dinámicas para gastos:', error);
    console.error('Stack trace:', error.stack);
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
    
    // Get template configuration
    const config = dynamicRowsData.templateConfig || templateConfig;
    
    // Use the configured template row range
    const templateRange = config?.templateRow?.range || "A44:AK44";
    const insertStartRow = config?.templateRow?.insertStartRow || 45;
    const copyStyle = config?.templateRow?.copyStyle !== false; // Default to true
    
    const rowsData = dynamicRowsData.gastos; // Use original data for insertion
    console.log(`Insertando ${rowsData.length} filas dinámicas basadas en template ${templateRange}`);
    
    // Initialize Google Sheets API
    const sheets = google.sheets({version: 'v4', auth: jwtClient});
    
    // 1. First, we need to add enough empty rows
    await insertEmptyRows(sheets, fileId, insertStartRow, rowsData.length);
    console.log(`✅ Insertadas ${rowsData.length} filas vacías en la posición ${insertStartRow}`);
    
    // 2. If we need to copy styles, copy the template row to each new row
    if (copyStyle) {
      await copyTemplateRowStyles(sheets, fileId, templateRange, insertStartRow, rowsData.length);
      console.log(`✅ Copiados estilos del template para ${rowsData.length} filas`);
    }
    
    // 3. Now insert the data into the specific starting cells for each field
    const dataMapping = config?.dataMapping || {};
    const columnDefs = config?.columns || {};
    const requests = [];

    for (let i = 0; i < rowsData.length; i++) {
      const rowNum = insertStartRow + i;
      const gasto = rowsData[i];

      // Prepare data updates for each column defined in the template
      Object.keys(columnDefs).forEach(colKey => {
        const columnInfo = columnDefs[colKey];
        const dataKey = dataMapping[colKey]; // Get the corresponding key in the gasto object
        let valueToInsert = '';

        if (dataKey && gasto[dataKey] !== undefined) {
          valueToInsert = gasto[dataKey]?.toString() || '';
        } else if (gasto[colKey] !== undefined) { // Fallback to direct key match
          valueToInsert = gasto[colKey]?.toString() || '';
        }
        
        // Extract the starting column letter (e.g., F from F:V)
        const startColumnLetter = columnInfo.column.split(':')[0];
        const cellRange = `${startColumnLetter}${rowNum}`;

        // Add update request for this specific cell
        requests.push({
          range: cellRange,
          values: [[valueToInsert]]
        });
      });
    }
    
    // Execute batch update for cell values
    if (requests.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: fileId,
        resource: {
          valueInputOption: 'USER_ENTERED',
          data: requests
        }
      });
      console.log(`✅ Insertados datos en ${rowsData.length} filas dinámicas en el reporte`);
    } else {
        console.log('No data update requests generated.');
    }

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
    // Get the spreadsheet to find the first sheet ID
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: fileId,
      fields: 'sheets.properties'
    });
    
    if (!spreadsheet.data.sheets || spreadsheet.data.sheets.length === 0) {
      throw new Error('No sheets found in the spreadsheet');
    }
    
    // Use the first sheet ID
    const sheetId = spreadsheet.data.sheets[0].properties.sheetId;
    
    // Insert rows (need to convert from 1-based to 0-based index)
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: fileId,
      resource: {
        requests: [{
          insertDimension: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: startRowIndex - 1, // Convert to 0-based
              endIndex: startRowIndex - 1 + rowCount
            },
            inheritFromBefore: true
          }
        }]
      }
    });
    
    return true;
  } catch (error) {
    console.error('Error inserting empty rows:', error);
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
      fields: 'sheets.properties,sheets.merges'
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
      spreadsheet.data.sheets[0].merges.forEach(mergeInfo => {
        if (mergeInfo.startRowIndex === templateRowIndex && 
            mergeInfo.endRowIndex === templateRowIndex + 1) {
          templateMerges.push(mergeInfo);
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
