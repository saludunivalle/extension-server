/**
 * Risks (Riesgos) Generator
 * Processes risk data and prepares it for insertion as dynamic rows
 */

const fs = require('fs');
const path = require('path');
const templateMapper = require('./templateMapperRiesgos');
const { google } = require('googleapis');
const { jwtClient } = require('../../config/google');

// Cargar configuración de template base desde risks.json
let baseTemplateConfig;
try {
  const templateFile = path.join(__dirname, '../../templates/risks.json');
  baseTemplateConfig = JSON.parse(fs.readFileSync(templateFile, 'utf8'));
  console.log('Base template configuration loaded successfully for risks');
} catch (error) {
  console.error('Error loading base template configuration for risks:', error);
  // Fallback configuration if file load fails
  baseTemplateConfig = {
    templateRow: { range: "B43:D43", copyStyle: true, insertStartRow: 44 },
    columns: {
      descripcion: { column: "B" },
      aplica: { column: "C" },
      mitigacion: { column: "D" }
    },
    dataMapping: {
      descripcion: "nombre_riesgo",
      aplica: "aplica",
      mitigacion: "mitigacion"
    }
  };
}

// Configuración específica para cada categoría de riesgos (B, C, D columns)
const categoriaConfig = {
  diseno: {
    templateRow: {
      range: "B17:D17",      // Template row is 17
      insertStartRow: 18,  // Insert below row 17 (start at 18)
      copyStyle: true
    }
  },
  locacion: {
    templateRow: {
      range: "B24:D24",      // Template row is 24
      insertStartRow: 25,  // Insert below row 24 (start at 25)
      copyStyle: true
    }
  },
  desarrollo: {
    templateRow: {
      range: "B35:D35",      // Template row is 35
      insertStartRow: 36,  // Insert below row 35 (start at 36)
      copyStyle: true
    }
  },
  cierre: {
    templateRow: {
      range: "B39:D39",      // Template row is 39
      insertStartRow: 40,  // Insert below row 39 (start at 40)
      copyStyle: true
    }
  },
  otros: {
    templateRow: {
      range: "B43:D43",      // Default template row is 43
      insertStartRow: 44,  // Default insert below row 43 (start at 44)
      copyStyle: true
      // This will be adjusted if 'cierre' rows exist
    }
  }
};

/**
 * Generate formatted rows from risk data
 * @param {Array} riesgos - Array of risk objects
 * @param {String} categoria - Risk category (diseno, locacion, desarrollo, cierre, otros)
 * @param {String | null} insertLocation - Optional custom insert location (overrides category config)
 * @param {Object} prevCategories - Information about previously inserted categories { cierre: { count: number, insertStartRow: number } }
 * @returns {Object | null} Formatted data for dynamic rows or null
 */
const generateRows = (riesgos, categoria = null, insertLocation = null, prevCategories = {}) => {
  if (!riesgos || !Array.isArray(riesgos) || riesgos.length === 0) {
    console.log(`No hay riesgos para generar filas dinámicas ${categoria ? `para categoría ${categoria}` : ''}`);
    return null;
  }
  
  console.log(`Generando ${riesgos.length} filas dinámicas para riesgos ${categoria ? `de categoría ${categoria}` : ''}`);
  
  // Deep copy base config to avoid modification issues
  let templateConfig = JSON.parse(JSON.stringify(baseTemplateConfig)); 
  let usingCategoryConfig = false;

  // Apply category-specific config if available and no custom location is given
  if (!insertLocation && categoria && categoriaConfig[categoria.toLowerCase()]) {
    const catKey = categoria.toLowerCase();
    templateConfig.templateRow = { 
      ...templateConfig.templateRow, 
      ...categoriaConfig[catKey].templateRow 
    };
    usingCategoryConfig = true;
    console.log(`Usando configuración específica para categoría: ${categoria}`);
    
    // --- SPECIAL CASE: Adjust 'otros' based on 'cierre' ---
    if (catKey === 'otros' && prevCategories.cierre && prevCategories.cierre.count > 0) {
      const cierreCount = prevCategories.cierre.count;
      // cierre.insertStartRow is the 1-based index where cierre *started* inserting.
      const cierreInsertStartRow = prevCategories.cierre.insertStartRow; 
      
      // Calculate where 'otros' should start inserting
      const otrosInsertStartRow = cierreInsertStartRow + cierreCount; // Start right after the last cierre row
      // The template row to copy style/range from is the one *before* the insertion point
      const otrosTemplateRowIndex = otrosInsertStartRow - 1; 

      // Update the template config for 'otros'
      templateConfig.templateRow.range = `B${otrosTemplateRowIndex}:D${otrosTemplateRowIndex}`;
      templateConfig.templateRow.insertStartRow = otrosInsertStartRow;
      
      console.log(`✅ Ajustada posición de "otros" para ir después de ${cierreCount} filas de cierre. Nueva fila de inserción: ${otrosInsertStartRow}`);
    }
    // --- END SPECIAL CASE ---

  } else if (insertLocation) {
    // Handle custom insert location if provided
    const match = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(insertLocation);
    if (match) {
      const startCol = match[1];
      const startRow = parseInt(match[2]);
      const endCol = match[3];
      // Update template config based on custom location
      templateConfig.templateRow.range = `${startCol}${startRow}:${endCol}${startRow}`; // Assume single template row
      templateConfig.templateRow.insertStartRow = startRow + 1; // Insert below
      console.log(`Usando ubicación de inserción personalizada: ${insertLocation} -> fila ${startRow + 1}`);
    } else {
       console.warn(`Formato de insertLocation inválido: ${insertLocation}. Usando configuración por defecto.`);
       // Optionally fallback to category or base config here
       if (categoria && categoriaConfig[categoria.toLowerCase()]) {
         templateConfig.templateRow = { ...templateConfig.templateRow, ...categoriaConfig[categoria.toLowerCase()].templateRow };
         usingCategoryConfig = true;
       }
    }
  } else if (categoria) {
      console.warn(`Categoría '${categoria}' no encontrada en categoriaConfig. Usando configuración base.`);
  }

  // Map each risk to the row format using the mapper
  // Ensure templateMapper is correctly loaded and createRow exists
  if (!templateMapper || typeof templateMapper.createRow !== 'function') {
     console.error("❌ templateMapperRiesgos no está cargado o no tiene createRow");
     return null;
  }
  const rows = riesgos.map(riesgo => templateMapper.createRow(riesgo));
  
  // Final insertion details
  const finalInsertarEn = templateConfig.templateRow.range;
  const finalInsertStartRow = templateConfig.templateRow.insertStartRow;
  
  console.log(`Configuración final: insertarEn=${finalInsertarEn}, insertStartRow=${finalInsertStartRow}`);
  
  return {
    insertarEn: finalInsertarEn,
    riesgos: riesgos, // Original risk data
    rows: rows,       // Formatted rows for insertion
    templateConfig: templateConfig, // Effective template config used
    insertStartRow: finalInsertStartRow, // Final 1-based start row for insertion
    count: riesgos.length // Number of rows generated for this category
  };
};

/**
 * Insert dynamic risk rows into a Google Sheet
 * @param {String} fileId - ID of the Google Sheet
 * @param {Object} dynamicRowsData - Processed data from generateRows
 * @returns {Promise<Boolean>} Success status
 */
const insertDynamicRows = async (fileId, dynamicRowsData) => {
  try {
    if (!dynamicRowsData || !dynamicRowsData.rows || dynamicRowsData.rows.length === 0) {
      console.log('No hay datos válidos para insertar filas dinámicas de riesgos');
      return false;
    }
    
    // Extract necessary info from dynamicRowsData
    const config = dynamicRowsData.templateConfig || baseTemplateConfig; // Fallback just in case
    const templateRange = dynamicRowsData.insertarEn; // e.g., "B17:D17"
    const insertStartRow = dynamicRowsData.insertStartRow; // e.g., 18 (1-based)
    const copyStyle = config?.templateRow?.copyStyle !== false; // Default to true
    const rowsToInsert = dynamicRowsData.rows; // The array of [desc, aplica, mitig]
    const originalRiesgosData = dynamicRowsData.riesgos; // Original data if needed for complex mapping (less likely now)
    const rowCount = rowsToInsert.length;

    console.log(`Insertando ${rowCount} filas dinámicas de riesgos basadas en template ${templateRange} desde la fila ${insertStartRow}`);
    
    // Initialize Google Sheets API
    const sheets = google.sheets({version: 'v4', auth: jwtClient});
    
    // 1. Insert empty rows first
    await insertEmptyRows(sheets, fileId, insertStartRow, rowCount);
    console.log(`✅ Insertadas ${rowCount} filas vacías en la posición ${insertStartRow}`);
    
    // 2. Copy styles if required
    if (copyStyle) {
      // Pass the correct range (B:D)
      await copyTemplateRowStyles(sheets, fileId, templateRange, insertStartRow, rowCount);
      console.log(`✅ Copiados estilos del template ${templateRange} para ${rowCount} filas`);
    }
    
    // 3. Insert the actual data using batchUpdate
    // Determine the range where data will be inserted
    const startColLetter = templateRange.match(/([A-Z]+)/)[1]; // Should be 'B'
    const endColLetter = templateRange.match(/:([A-Z]+)/)[1];   // Should be 'D'
    const dataInsertRange = `${startColLetter}${insertStartRow}:${endColLetter}${insertStartRow + rowCount - 1}`;
    
    console.log(`Insertando datos en el rango: ${dataInsertRange}`);

    await sheets.spreadsheets.values.update({
        spreadsheetId: fileId,
        range: dataInsertRange,
        valueInputOption: 'USER_ENTERED', // Or 'RAW' if no formulas/parsing needed
        resource: {
            values: rowsToInsert // Use the pre-formatted rows directly
        }
    });

    console.log(`✅ Insertados datos en ${rowCount} filas dinámicas de riesgos en el reporte`);

    return true;
  } catch (error) {
    console.error('❌ Error al insertar filas dinámicas de riesgos:', error);
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
            inheritFromBefore: true // Inherit styles from row above
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
 * @param {String} templateRange - Template row range (e.g., "B17:D17")
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
    
    // --- ADJUSTMENT: Calculate correct column indices for B:D ---
    const startColIndex = columnToIndex(startCol); // Should be 1 for 'B'
    const endColIndex = columnToIndex(endCol);     // Should be 3 for 'D'
    // --- END ADJUSTMENT ---

    if (startColIndex === null || endColIndex === null) {
        throw new Error(`Invalid column letters in range: ${templateRange}`);
    }

    // Find merged cells in the template row
    const templateMerges = [];
    if (spreadsheet.data.sheets[0].merges) {
      spreadsheet.data.sheets[0].merges.forEach(mergeInfo => {
        // Check if merge is within the template row and within B:D columns
        if (mergeInfo.sheetId === sheetId &&
            mergeInfo.startRowIndex === templateRowIndex && 
            mergeInfo.endRowIndex === templateRowIndex + 1 &&
            mergeInfo.startColumnIndex >= startColIndex &&
            mergeInfo.endColumnIndex <= endColIndex + 1) { // End index is exclusive
          templateMerges.push(mergeInfo);
        }
      });
    }
    console.log(`Found ${templateMerges.length} merged cell ranges in template row ${templateRange}`);
    
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
            startColumnIndex: startColIndex,
            endColumnIndex: endColIndex + 1
          },
          destination: {
            sheetId: sheetId,
            startRowIndex: targetRowIndex,
            endRowIndex: targetRowIndex + 1,
            startColumnIndex: startColIndex,
            endColumnIndex: endColIndex + 1
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
 * Convert column letter to index (0-based)
 * @param {String} column - Column letter (e.g., "A", "BC")
 * @returns {Number|null} Column index (0-based) or null if invalid
 */
function columnToIndex(column) {
  if (!column || typeof column !== 'string') return null;
  const colLetter = column.toUpperCase();
  let result = 0;
  for (let i = 0; i < colLetter.length; i++) {
    const charCode = colLetter.charCodeAt(i);
    if (charCode < 65 || charCode > 90) return null;
    result = result * 26 + (charCode - 64);
  }
  return result - 1; // Convert to 0-based index
}

module.exports = {
  generateRows,
  insertDynamicRows
};