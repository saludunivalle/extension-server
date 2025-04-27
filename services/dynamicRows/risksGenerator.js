/**
 * Risks (Riesgos) Generator
 * Processes risk data and prepares it for insertion as dynamic rows
 */

const fs = require('fs');
const path = require('path');
const templateMapper = require('./templateMapperRiesgos');
const { google } = require('googleapis');
const { jwtClient } = require('../../config/google');

// Cargar configuración completa desde risks.json
let fullTemplateConfig;
try {
  const templateFile = path.join(__dirname, '../../templates/risks.json');
  fullTemplateConfig = JSON.parse(fs.readFileSync(templateFile, 'utf8'));
  console.log('Full template configuration loaded successfully for risks');
} catch (error) {
  console.error('Error loading full template configuration for risks:', error);
  // Fallback configuration if file load fails (ajustar según sea necesario)
  fullTemplateConfig = {
    templateRow: { range: "B43:E43", copyStyle: true, insertStartRow: 44 },
    columns: {
      descripcion: { column: "B" },
      aplica: { column: "C" },
      mitigacion: { column: "D:E", span: 2 }
    },
    mergedCells: { // Rangos por defecto Columna A
      diseno: { startRow: 14, endRow: 17 },
      locacion: { startRow: 19, endRow: 24 },
      desarrollo: { startRow: 25, endRow: 35 },
      cierre: { startRow: 36, endRow: 38 },
      otros: { startRow: 39, endRow: 39 }
    },
    dataMapping: { /* ... */ }
  };
}

// Configuración específica para la *inserción* de filas (B:E) y copia de estilos
const categoriaConfigInsercion = {
  diseno: { templateRow: { range: "B17:E17", insertStartRow: 18, copyStyle: true } },     // Template fila 17, inserta @ 18
  locacion: { templateRow: { range: "B24:E24", insertStartRow: 25, copyStyle: true } },   // Template fila 24, inserta @ 25
  desarrollo: { templateRow: { range: "B35:E35", insertStartRow: 36, copyStyle: true } }, // Template fila 35, inserta @ 36
  cierre: { templateRow: { range: "B38:E38", insertStartRow: 39, copyStyle: true } },     // CORREGIDO: Template fila 38, inserta @ 39
  otros: { templateRow: { range: "B39:E39", insertStartRow: 40, copyStyle: true } }      // CORREGIDO: Template fila 39, inserta @ 40 (default, se ajusta si cierre tiene filas)
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
    // Devolvemos una estructura mínima para que se pueda procesar la columna A
    const catKey = categoria?.toLowerCase();
    const defaultMergeInfo = fullTemplateConfig.mergedCells[catKey] || fullTemplateConfig.mergedCells.otros;
    const insertionInfo = categoriaConfigInsercion[catKey] || categoriaConfigInsercion.otros;

    // Ajuste para 'otros' si 'cierre' existe
    let finalInsertStartRow = insertionInfo.templateRow.insertStartRow;
    if (catKey === 'otros' && prevCategories.cierre && prevCategories.cierre.count > 0) {
        finalInsertStartRow = prevCategories.cierre.startRow + prevCategories.cierre.count;
    }

    return {
        insertarEn: insertionInfo.templateRow.range, // Rango B:E para copiar estilo
        riesgos: [],
        rows: [],
        templateConfig: insertionInfo, // Config de inserción B:E
        insertStartRow: finalInsertStartRow, // Dónde se *insertarían* las filas
        count: 0,
        categoria: categoria,
        // Añadir info para merge de columna A
        columnaAMergeConfig: {
            startRow: defaultMergeInfo.startRow,
            defaultEndRow: defaultMergeInfo.endRow
        }
    };
  }

  console.log(`Generando ${riesgos.length} filas dinámicas para riesgos ${categoria ? `de categoría ${categoria}` : ''}`);

  // Usar la configuración de inserción específica de la categoría
  let templateConfig = JSON.parse(JSON.stringify(fullTemplateConfig)); // Copia profunda
  let insertionConfig = {}; // Config para B:E
  const catKey = categoria?.toLowerCase();

  if (!insertLocation && catKey && categoriaConfigInsercion[catKey]) {
    insertionConfig = JSON.parse(JSON.stringify(categoriaConfigInsercion[catKey])); // Copia profunda
    console.log(`Usando configuración de inserción específica para categoría: ${categoria}`);

    // --- SPECIAL CASE: Adjust 'otros' based on 'cierre' ---
    if (catKey === 'otros' && prevCategories.cierre && prevCategories.cierre.count > 0) {
      const cierreCount = prevCategories.cierre.count;
      const cierreStartRow = prevCategories.cierre.startRow; // Fila donde empezó cierre
      const otrosInsertStartRow = cierreStartRow + cierreCount; // Fila donde empieza otros
      const otrosTemplateRowIndex = otrosInsertStartRow - 1; // Fila plantilla para otros

      // Actualizar la config de *inserción* para 'otros'
      insertionConfig.templateRow.range = `B${otrosTemplateRowIndex}:E${otrosTemplateRowIndex}`;
      insertionConfig.templateRow.insertStartRow = otrosInsertStartRow;

      console.log(`✅ Ajustada posición de inserción de "otros" para ir después de ${cierreCount} filas de cierre. Nueva fila de inserción: ${otrosInsertStartRow}`);
    }
    // --- END SPECIAL CASE ---

  } else if (insertLocation) {
    // Manejar ubicación personalizada (menos probable para riesgos, pero mantenido)
    // ... (código existente para parsear insertLocation y ajustar insertionConfig) ...
    console.warn("Ubicación personalizada no recomendada para riesgos categorizados, puede afectar la combinación de Columna A.");
    // Usar config por defecto si falla el parseo
    if (!insertionConfig.templateRow) {
        insertionConfig = JSON.parse(JSON.stringify(categoriaConfigInsercion.otros));
    }
  } else {
     console.warn(`Categoría '${categoria}' no encontrada en categoriaConfigInsercion. Usando configuración de inserción 'otros'.`);
     insertionConfig = JSON.parse(JSON.stringify(categoriaConfigInsercion.otros));
  }

  // Mapear datos usando templateMapper (que usa B:E)
  if (!templateMapper || typeof templateMapper.createRow !== 'function') {
     console.error("❌ templateMapperRiesgos no está cargado o no tiene createRow");
     return null;
  }
  const rows = riesgos.map(riesgo => templateMapper.createRow(riesgo));

  // Detalles finales de inserción
  const finalInsertarEn = insertionConfig.templateRow.range; // Rango B:E
  const finalInsertStartRow = insertionConfig.templateRow.insertStartRow; // Fila 1-based

  // Información para la combinación de la columna A
  const defaultMergeInfo = fullTemplateConfig.mergedCells[catKey] || fullTemplateConfig.mergedCells.otros;

  console.log(`Configuración final: insertarEn=${finalInsertarEn}, insertStartRow=${finalInsertStartRow}`);

  return {
    insertarEn: finalInsertarEn, // Rango B:E para copiar estilo
    riesgos: riesgos,
    rows: rows,
    templateConfig: insertionConfig, // Config de inserción B:E
    insertStartRow: finalInsertStartRow, // Dónde se insertan las filas
    count: riesgos.length,
    categoria: categoria,
    // Añadir info para merge de columna A
    columnaAMergeConfig: {
        startRow: defaultMergeInfo.startRow,
        defaultEndRow: defaultMergeInfo.endRow
    }
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
    // Validar datos de entrada
    if (!dynamicRowsData || !dynamicRowsData.categoria || !dynamicRowsData.columnaAMergeConfig) {
      console.log('Datos inválidos para insertar filas dinámicas de riesgos (faltan datos de categoría o config merge A)');
      return false;
    }

    const {
        rows: rowsToInsert,
        insertarEn: templateRange, // Rango B:E
        insertStartRow, // Fila 1-based donde insertar
        count: rowCount,
        categoria,
        columnaAMergeConfig,
        templateConfig // Config de inserción B:E
    } = dynamicRowsData;

    const copyStyle = templateConfig?.templateRow?.copyStyle !== false;

    // Initialize Google Sheets API
    const sheets = google.sheets({version: 'v4', auth: jwtClient});

    // Get sheet ID
    const spreadsheet = await sheets.spreadsheets.get({
        spreadsheetId: fileId,
        fields: 'sheets.properties,sheets.merges' // Necesitamos merges para unmerge
    });
    if (!spreadsheet.data.sheets || spreadsheet.data.sheets.length === 0) {
        throw new Error('No sheets found in the spreadsheet');
    }
    const sheetId = spreadsheet.data.sheets[0].properties.sheetId;
    const currentMerges = spreadsheet.data.sheets[0].merges || [];


    if (rowCount > 0) {
        console.log(`Insertando ${rowCount} filas dinámicas de riesgos (${categoria}) basadas en template ${templateRange} desde la fila ${insertStartRow}`);

        // 1. Insert empty rows first
        await insertEmptyRows(sheets, fileId, sheetId, insertStartRow, rowCount);
        console.log(`✅ Insertadas ${rowCount} filas vacías en la posición ${insertStartRow}`);

        // 2. Copy styles (B:E) and recreate D:E merge if required
        if (copyStyle) {
            await copyTemplateRowStyles(sheets, fileId, sheetId, currentMerges, templateRange, insertStartRow, rowCount);
            console.log(`✅ Copiados estilos del template ${templateRange} y recreada combinación D:E para ${rowCount} filas`);
        }

        // 3. Insert the actual data (B:D)
        const startColLetter = templateRange.match(/([A-Z]+)/)[1]; // 'B'
        // --- CORRECCIÓN: El mapeador ya genera 4 columnas (B, C, D, E), D y E se combinan después ---
        // --- Necesitamos insertar en B:E ---
        const endColLetter = templateRange.match(/:([A-Z]+)/)[1];   // 'E'
        const dataInsertRange = `${startColLetter}${insertStartRow}:${endColLetter}${insertStartRow + rowCount - 1}`;

        console.log(`Insertando datos en el rango: ${dataInsertRange}`);

        await sheets.spreadsheets.values.update({
            spreadsheetId: fileId,
            range: dataInsertRange,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: rowsToInsert // El mapeador ya genera las 4 columnas
            }
        });
        console.log(`✅ Insertados datos en ${rowCount} filas dinámicas de riesgos`);

    } else {
        console.log(`No hay filas dinámicas para insertar en categoría ${categoria}. Solo se ajustará la columna A.`);
    }

    // 4. Handle Column A Merging (SIEMPRE se llama, incluso si rowCount es 0)
    await handleColumnAMerges(
        sheets,
        fileId,
        sheetId,
        currentMerges, // Pasar merges actuales para unmerge
        categoria,
        columnaAMergeConfig.startRow,
        columnaAMergeConfig.defaultEndRow,
        insertStartRow, // Fila donde *empezaron* las inserciones (o donde empezarían)
        rowCount
    );

    return true;
  } catch (error) {
    console.error(`❌ Error al insertar filas dinámicas de riesgos (${dynamicRowsData?.categoria}):`, error);
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
 * @param {Number} sheetId - ID of the sheet
 * @param {Number} startRowIndex - 1-based row index where to insert
 * @param {Number} rowCount - Number of rows to insert
 */
async function insertEmptyRows(sheets, fileId, sheetId, startRowIndex, rowCount) {
  if (rowCount <= 0) return;
  try {
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
            inheritFromBefore: true // Herencia simple, el estilo se copia después
          }
        }]
      }
    });
  } catch (error) {
    console.error('Error inserting empty rows:', error);
    throw error; // Re-lanzar para que insertDynamicRows lo capture
  }
}

/**
 * Copy template row styles (B:E) and recreate merges (like D:E)
 * @param {Object} sheets - Google Sheets API client
 * @param {String} fileId - Spreadsheet ID
 * @param {Number} sheetId - Sheet ID
 * @param {Array} currentMerges - Array of existing merge objects from spreadsheet metadata
 * @param {String} templateRange - Template row range (e.g., "B17:E17")
 * @param {Number} insertStartRow - 1-based row index where rows were inserted
 * @param {Number} rowCount - Number of rows inserted
 */
async function copyTemplateRowStyles(sheets, fileId, sheetId, currentMerges, templateRange, insertStartRow, rowCount) {
  if (rowCount <= 0) return;
  try {
    // Parse the template range (B:E)
    const match = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(templateRange);
    if (!match) throw new Error(`Invalid template range format: ${templateRange}`);

    const [_, startCol, startRow, endCol, endRow] = match;
    const templateRowIndex = parseInt(startRow) - 1; // 0-based
    const startColIndex = columnToIndex(startCol); // e.g., 1 for 'B'
    const endColIndex = columnToIndex(endCol);     // e.g., 4 for 'E'

    if (startColIndex === null || endColIndex === null) {
        throw new Error(`Invalid column letters in range: ${templateRange}`);
    }

    // Find merges within the template row's columns (B:E)
    const templateMergesToRecreate = [];
    if (currentMerges) {
      currentMerges.forEach(mergeInfo => {
        // Check if the merge is exactly the template row and within B:E columns
        if (mergeInfo.sheetId === sheetId &&
            mergeInfo.startRowIndex === templateRowIndex &&
            mergeInfo.endRowIndex === templateRowIndex + 1 && // Merge is only one row high
            mergeInfo.startColumnIndex >= startColIndex &&
            mergeInfo.endColumnIndex <= endColIndex + 1) { // End index is exclusive
          templateMergesToRecreate.push(mergeInfo);
        }
      });
    }
    console.log(`Found ${templateMergesToRecreate.length} merge(s) to recreate in template row ${templateRange} (e.g., D:E)`);

    const requests = [];

    // 1. Copy Format Request (B:E)
    requests.push({
      copyPaste: {
        source: {
          sheetId: sheetId,
          startRowIndex: templateRowIndex,
          endRowIndex: templateRowIndex + 1,
          startColumnIndex: startColIndex,
          endColumnIndex: endColIndex + 1 // Exclusive
        },
        destination: {
          sheetId: sheetId,
          startRowIndex: insertStartRow - 1, // 0-based start of new rows
          endRowIndex: insertStartRow - 1 + rowCount,
          startColumnIndex: startColIndex,
          endColumnIndex: endColIndex + 1 // Exclusive
        },
        pasteType: 'PASTE_FORMAT', // Only copy formatting
        pasteOrientation: 'NORMAL'
      }
    });

    // 2. Recreate Merges (like D:E) for each new row
    for (let i = 0; i < rowCount; i++) {
      const targetRowIndex = insertStartRow - 1 + i;
      templateMergesToRecreate.forEach(merge => {
        requests.push({
          mergeCells: {
            range: {
              sheetId: sheetId,
              startRowIndex: targetRowIndex,
              endRowIndex: targetRowIndex + 1,
              // Calculate relative column indices based on the template merge
              startColumnIndex: merge.startColumnIndex,
              endColumnIndex: merge.endColumnIndex
            },
            mergeType: 'MERGE_ALL' // Or MERGE_COLUMNS if appropriate
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

  } catch (error) {
    console.error('Error copying template styles or recreating merges:', error);
    throw error; // Re-lanzar
  }
}


/**
 * Handle merging cells in column A for a specific category section.
 * Unmerges existing overlaps first, then creates the new merge.
 * Handles special cases for 'cierre' and 'otros'.
 * @param {Object} sheets - Google Sheets API client
 * @param {String} fileId - Spreadsheet ID
 * @param {Number} sheetId - Sheet ID
 * @param {Array} currentMerges - Array of existing merge objects from spreadsheet metadata
 * @param {String} categoria - Risk category key (lowercase)
 * @param {Number} sectionStartRow - The defined starting row for this category's A-column cell (1-based, e.g., 14 for diseno)
 * @param {Number} sectionDefaultEndRow - The defined default ending row if no dynamic rows (1-based, e.g., 17 for diseno)
 * @param {Number} dynamicInsertStartRow - The row where dynamic rows *started* inserting (1-based). Can be 0 if no rows inserted.
 * @param {Number} dynamicRowCount - Number of dynamic rows inserted for this category.
 */
async function handleColumnAMerges(sheets, fileId, sheetId, currentMerges, categoria, sectionStartRow, sectionDefaultEndRow, dynamicInsertStartRow, dynamicRowCount) {
  const requests = [];
  const categoriaKey = categoria.toLowerCase();

  // --- Special Case: "otros" with no dynamic rows ---
  if (categoriaKey === 'otros' && dynamicRowCount === 0) {
    console.log(`Categoría 'otros' sin filas dinámicas. No se realizarán cambios en Columna A.`);
    // IMPORTANTE: No hacer nada. No borrar, no combinar, no descombinar.
    // Si previamente existía una combinación A39:A39, se deja como está o se maneja manualmente.
    // Opcionalmente, podríamos descombinar explícitamente A39 si siempre debe quedar descombinado en este caso.
    // await unmergeColumnACells(sheets, fileId, sheetId, currentMerges, sectionStartRow, sectionStartRow); // Descomentar si se quiere descombinar A39
    return; // Salir de la función para 'otros' sin filas
  }

  // --- Normal Case / Cierre / Otros con filas ---

  // Calculate the final end row for the merge (1-based)
  let finalEndRow;
  if (dynamicRowCount > 0) {
    // If there are dynamic rows, always merge down to the last one
    finalEndRow = dynamicInsertStartRow + dynamicRowCount - 1;
  } else {
    // No dynamic rows (and not 'otros'), use the default end row
    finalEndRow = sectionDefaultEndRow;
  }

  // Ensure start row is not after end row
  if (finalEndRow < sectionStartRow) {
      console.warn(`La fila final calculada (${finalEndRow}) es menor que la inicial (${sectionStartRow}) para ${categoria}. No se combinará la columna A.`);
      return; // Do nothing
  }

  console.log(`Preparando combinación para Columna A (${categoria}): A${sectionStartRow}:A${finalEndRow}`);

  // 1. Unmerge existing overlaps in Column A for the *entire potential range*
  // Esto es importante para evitar errores al intentar combinar celdas ya combinadas de forma diferente.
  await unmergeColumnACells(sheets, fileId, sheetId, currentMerges, sectionStartRow, finalEndRow);

  // 2. Create the new merge request for Column A
  requests.push({
    mergeCells: {
      range: {
        sheetId: sheetId,
        startRowIndex: sectionStartRow - 1,   // 0-based start
        endRowIndex: finalEndRow,             // 0-based end (exclusive, so finalEndRow works)
        startColumnIndex: 0,                  // Column A
        endColumnIndex: 1                     // Exclusive
      },
      mergeType: 'MERGE_COLUMNS' // Merge vertically in the column
    }
  });


  // Execute batch update if there are requests
  if (requests.length > 0) {
    try {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: fileId,
        resource: { requests }
      });
      console.log(`✅ Combinación de Columna A para categoría ${categoria} (A${sectionStartRow}:A${finalEndRow}) completada.`);
    } catch (error) {
      console.error(`Error en batchUpdate para ${categoria} (Columna A):`, error);
      if (error.response && error.response.data) {
        console.error('Google API Error:', JSON.stringify(error.response.data, null, 2));
      }
      // Decide if you want to throw or just log
    }
  }
}


/**
 * Unmerge any existing merged cells in column A that overlap with the target range.
 * @param {Object} sheets - Google Sheets API client
 * @param {String} fileId - Spreadsheet ID
 * @param {Number} sheetId - Sheet ID
 * @param {Array} currentMerges - Array of existing merge objects from spreadsheet metadata
 * @param {Number} targetStartRow - Start row of the range to check (1-based)
 * @param {Number} targetEndRow - End row of the range to check (1-based)
 */
async function unmergeColumnACells(sheets, fileId, sheetId, currentMerges, targetStartRow, targetEndRow) {
  const requests = [];
  if (!currentMerges || currentMerges.length === 0) {
      console.log("No hay combinaciones existentes para verificar.");
      return; // No merges exist
  }


  // Convert target range to 0-based for comparison
  const targetStartRowIndex = targetStartRow - 1;
  const targetEndRowIndex = targetEndRow; // Exclusive

  console.log(`Buscando combinaciones existentes en Col A para descombinar en rango ${targetStartRow}-${targetEndRow}`);

  for (const merge of currentMerges) {
    // Check if it's in the correct sheet and is a Column A merge
    if (merge.sheetId === sheetId && merge.startColumnIndex === 0 && merge.endColumnIndex === 1) {
      // Check for overlap: !(merge ends before target starts || merge starts after target ends)
      const overlaps = !(merge.endRowIndex <= targetStartRowIndex || merge.startRowIndex >= targetEndRowIndex);

      if (overlaps) {
        requests.push({
          unmergeCells: {
            range: { // Use the exact range of the merge to unmerge
              sheetId: sheetId,
              startRowIndex: merge.startRowIndex,
              endRowIndex: merge.endRowIndex,
              startColumnIndex: merge.startColumnIndex,
              endColumnIndex: merge.endColumnIndex
            }
          }
        });
        console.log(`-> Deshaciendo combinación existente A${merge.startRowIndex + 1}:A${merge.endRowIndex -1}`);
      }
    }
  }

  // Execute unmerge requests if any
  if (requests.length > 0) {
    try {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: fileId,
        resource: { requests }
      });
      console.log(`✅ Deshizo ${requests.length} combinaciones existentes en columna A para el rango ${targetStartRow}-${targetEndRow}`);
    } catch (error) {
      console.error('Error al deshacer combinaciones existentes en Columna A:', error);
      if (error.response && error.response.data) {
        console.error('Google API Error (unmerge):', JSON.stringify(error.response.data, null, 2));
      }
      // Log the error but continue, maybe the merge will still work
    }
  } else {
      console.log("No se encontraron combinaciones superpuestas en Col A para deshacer.");
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
    if (charCode < 65 || charCode > 90) return null; // Invalid char
    result = result * 26 + (charCode - 64);
  }
  return result - 1; // Convert to 0-based index
}

module.exports = {
  generateRows,
  insertDynamicRows
};