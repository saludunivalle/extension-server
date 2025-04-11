/**
 * Risks (Riesgos) Generator
 * Processes risk data and prepares it for insertion as dynamic rows
 */

const fs = require('fs');
const path = require('path');
const templateMapper = require('./templateMapperRiesgos');
const { google } = require('googleapis');
const { jwtClient } = require('../../config/google');

// Load template configuration
let templateConfig;
try {
  const templateFile = path.join(__dirname, '../../templates/risks.json');
  templateConfig = JSON.parse(fs.readFileSync(templateFile, 'utf8'));
  console.log('Template configuration loaded successfully for risks');
} catch (error) {
  console.error('Error loading template configuration for risks:', error);
  templateConfig = {
    templateRow: {
      range: "B44:H44",
      copyStyle: true,
      insertStartRow: 45
    },
    columns: {
      id: { column: "B" },
      description: { column: "C", span: 2 },
      impact: { column: "E" },
      probability: { column: "F" },
      strategy: { column: "G" }
    }
  };
}

/**
 * Generate formatted rows from risk data
 * @param {Array} riesgos - Array of risk objects
 * @param {String} insertLocation - Optional custom insert location
 * @returns {Object} Formatted data for dynamic rows
 */
const generateRows = (riesgos, insertLocation = null) => {
  if (!riesgos || !Array.isArray(riesgos) || riesgos.length === 0) {
    console.log('No hay riesgos para generar filas dinámicas');
    return null;
  }
  
  console.log(`Generando ${riesgos.length} filas dinámicas para riesgos`);
  
  // Map each risk to the format expected by the template
  const rows = riesgos.map(riesgo => templateMapper.createRow(riesgo));
  
  // Default insert location from template config if available
  const defaultInsert = templateConfig?.templateRow?.range || templateMapper.defaultInsertLocation;
  
  return {
    insertarEn: insertLocation || defaultInsert,
    riesgos: riesgos,
    rows: rows,
    templateConfig: templateConfig // Include template config for styling
  };
};

/**
 * Insert dynamic risk rows into a Google Sheet
 * @param {String} fileId - ID of the Google Sheet
 * @param {Object} dynamicRowsData - Processed data with rows to insert
 * @returns {Promise<Boolean>} Success status
 */
const insertDynamicRows = async (fileId, dynamicRowsData) => {
  try {
    if (!dynamicRowsData || !dynamicRowsData.rows || dynamicRowsData.rows.length === 0) {
      console.log('No hay datos para insertar filas dinámicas de riesgos');
      return false;
    }
    
    // Get template configuration
    const config = dynamicRowsData.templateConfig || templateConfig;
    
    // Use the configured template row range
    const templateRange = config?.templateRow?.range || "B44:H44";
    const insertStartRow = config?.templateRow?.insertStartRow || 45;
    const copyStyle = config?.templateRow?.copyStyle !== false; // Default to true
    
    const rowsData = dynamicRowsData.riesgos; // Use original data for insertion
    console.log(`Insertando ${rowsData.length} filas dinámicas de riesgos basadas en template ${templateRange}`);
    
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
      const riesgo = rowsData[i];

      // Prepare data updates for each column defined in the template
      Object.keys(columnDefs).forEach(colKey => {
        const columnInfo = columnDefs[colKey];
        const dataKey = dataMapping[colKey]; // Get the corresponding key in the riesgo object
        let valueToInsert = '';

        if (dataKey && riesgo[dataKey] !== undefined) {
          valueToInsert = riesgo[dataKey]?.toString() || '';
        } else if (riesgo[colKey] !== undefined) { // Fallback to direct key match
          valueToInsert = riesgo[colKey]?.toString() || '';
        }
        
        // Extract the starting column letter (e.g., C from C:D)
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
      console.log(`✅ Insertados datos en ${rowsData.length} filas dinámicas de riesgos en el reporte`);
    } else {
        console.log('No data update requests generated for risks.');
    }

    return true;
  } catch (error) {
    console.error('Error al insertar filas dinámicas de riesgos:', error);
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
 * @param {String} templateRange - Template row range (e.g., "B44:H44")
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
    
    console.log(`Found ${templateMerges.length} merged cell ranges in template row for risks`);
    
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
 * Extract column range from a cell range like "B44:H44"
 * @param {String} range - Cell range
 * @returns {Array} Array with start and end column letters
 */
function extractColumnRange(range) {
  const match = /([A-Z]+)\d+:([A-Z]+)\d+/.exec(range);
  if (!match) {
    return ['B', 'H']; // Default fallback
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
