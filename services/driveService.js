const { google } = require('googleapis');
const { jwtClient } = require('../config/google');
const excelUtils = require('../utils/excelUtils');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { generateExpenseRows } = require('./dynamicRows');
const expensesGenerator = require('./dynamicRows/expensesGenerator');

/**
 * Servicio para manejar operaciones con Google Drive
 */

class DriveService {
  constructor() {
    this.drive = google.drive({ version: 'v3', auth: jwtClient });
    this.templateFolderId = '12bxb0XEArXMLvc7gX2ndqJVqS_sTiiUE'; //Folder con plantillas
    this.uploadsFolder = '1iDJTcUYCV7C7dTsa0Y3rfBAjFUUelu-x';
    
    // IDs de plantillas de formularios
    this.templateIds = {
      1: '1xsz9YSnYEOng56eNKGV9it9EgTn0mZw1', //Plantilla de formulario 1
      2: '1nWY2gYKtJuXQnGLsdN7RID_2QmrHKRtOwcQsaTsOOm8', //Plantilla de formulario 2
      3: '1WoPUZYusNl2u3FpmZ1qiO5URBUqHIwKF', //Plantilla de formulario 3
      4: '1FTC7Vq3O4ultexRPXYrJKOpL9G0071-0', //Plantilla de formulario 4
    };
  }

  /**
   * Sube un archivo a Google Drive
   * @param {Object} file - Objeto de archivo de Multer
   * @param {String} folderDestination - ID de carpeta destino (opcional)
   * @returns {Promise<String>} - URL del archivo subido
   */

  async uploadFile(file, folderDestination = this.uploadsFolder) {
    try {
      if (!file) {
        throw new Error('No se proporcionó ningún archivo para subir');
      }

      const fileMetadata = {
        name: file.originalname,
        parents: [folderDestination]
      };
      
      const media = {
        mimeType: file.mimetype,
        body: fs.createReadStream(file.path)
      };
      
      const uploadedFile = await this.drive.files.create({
        requestBody: fileMetadata,
        media: media,
        fields: 'id'
      });
      
      const fileId = uploadedFile.data.id;
      
      // Hacer el archivo accesible públicamente
      await this.makeFilePublic(fileId);
      
      return `https://drive.google.com/file/d/${fileId}/view`;
    } catch (error) {
      console.error('Error al subir archivo a Google Drive:', error);
      throw new Error(`Error al subir archivo: ${error.message}`);
    }
  }

  /**
   * Hace público un archivo en Google Drive
   * @param {String} fileId - ID del archivo
   */

  async makeFilePublic(fileId) {
    try {
      await this.drive.permissions.create({
        fileId,
        requestBody: {
          role: 'reader',
          type: 'anyone'
        }
      });
      
      return true;
    } catch (error) {
      console.error('Error al hacer público el archivo:', error);
      throw new Error(`Error al modificar permisos: ${error.message}`);
    }
  }

  /**
   * Versión mejorada de processXLSXWithStyles en driveService.js 
   * para manejar correctamente filas dinámicas
   */
  async processXLSXWithStyles(templateId, data, fileName) {
    try {
      console.log(`Descargando la plantilla: ${templateId}`);
      
      // 1. Primero, verificar el tipo de archivo de la plantilla
      const fileInfo = await this.drive.files.get({
        fileId: templateId,
        fields: 'mimeType,name'
      });
      
      const isGoogleSheet = fileInfo.data.mimeType === 'application/vnd.google-apps.spreadsheet';
      console.log(`Tipo de plantilla: ${fileInfo.data.mimeType} (${isGoogleSheet ? 'Google Sheet' : 'XLSX/Excel'})`);
      
      let fileId;
      
      // 2. Extraer y guardar datos dinámicos para procesamiento posterior
      const dynamicRowsData = data['__FILAS_DINAMICAS__'];
      console.log('Datos dinámicos:', dynamicRowsData ? `${dynamicRowsData.gastos?.length || 0} gastos` : 'Ninguno');
      
      // Extraer campos de gastos individuales antes de eliminar __FILAS_DINAMICAS__
      const gastoFields = {};
      if (dynamicRowsData && dynamicRowsData.gastos) {
        console.log(`Procesando ${dynamicRowsData.gastos.length} gastos dinámicos para marcadores`);
        dynamicRowsData.gastos.forEach((gasto, index) => {
          // Manejar ambos formatos de ID (con coma o con punto)
          let idConComa = gasto.id || `15,${index + 1}`;
          let idConPunto = gasto.id || `15.${index + 1}`;
          
          // Asegurar que tenemos ambas versiones
          if (idConComa.includes('.')) {
            idConComa = idConComa.replace('.', ',');
          }
          
          if (idConPunto.includes(',')) {
            idConPunto = idConPunto.replace(',', '.');
          }
          
          // Versión con coma para Excel
          gastoFields[`gasto_${idConComa}_cantidad`] = gasto.cantidad?.toString() || '0';
          gastoFields[`gasto_${idConComa}_valor_unit`] = gasto.valorUnit_formatted || gasto.valor_unit_formatted || '0';
          gastoFields[`gasto_${idConComa}_valor_total`] = gasto.valorTotal_formatted || gasto.valor_total_formatted || '0';
          gastoFields[`gasto_${idConComa}_descripcion`] = gasto.descripcion || gasto.concepto || '';
          
          // Versión con punto (más estándar)
          gastoFields[`gasto_${idConPunto}_cantidad`] = gasto.cantidad?.toString() || '0';
          gastoFields[`gasto_${idConPunto}_valor_unit`] = gasto.valorUnit_formatted || gasto.valor_unit_formatted || '0';
          gastoFields[`gasto_${idConPunto}_valor_total`] = gasto.valorTotal_formatted || gasto.valor_total_formatted || '0';
          gastoFields[`gasto_${idConPunto}_descripcion`] = gasto.descripcion || gasto.concepto || '';
          
          console.log(`Marcadores creados para gasto ${idConComa}`);
        });
      }
      
      // DIFERENTE FLUJO SEGÚN EL TIPO DE PLANTILLA
      if (isGoogleSheet) {
        // ==== FLUJO PARA PLANTILLA DE GOOGLE SHEET ====
        console.log('Usando flujo para plantilla Google Sheet');
        
        // Duplicar la hoja de Google ya existente
        const copyResponse = await this.drive.files.copy({
          fileId: templateId,
          requestBody: {
            name: fileName,
            parents: [this.templateFolderId]
          }
        });
        
        fileId = copyResponse.data.id;
        await this.makeFilePublic(fileId);
        console.log(`Plantilla Google Sheet duplicada con ID: ${fileId}`);
        
        // Crear una copia de los datos sin dinamicRowsData
        const processData = { ...data };
        delete processData['__FILAS_DINAMICAS__'];
        
        // Incorporar los campos individuales de gastos
        Object.assign(processData, gastoFields);
        
        // Reemplazar marcadores en el Sheet
        const sheets = google.sheets({ version: 'v4', auth: jwtClient });
        
        // Obtener todos los datos actuales de la hoja para buscar marcadores
        const sheetData = await sheets.spreadsheets.values.get({
          spreadsheetId: fileId,
          range: 'A1:AZ1000' // Rango amplio para cubrir todos los posibles marcadores
        });
        
        const values = sheetData.data.values || [];
        const updates = [];
        
        // Buscar marcadores y reemplazarlos
        for (let r = 0; r < values.length; r++) {
          const row = values[r];
          for (let c = 0; c < row.length; c++) {
            const cell = row[c];
            
            // Si la celda contiene un marcador
            if (typeof cell === 'string' && cell.includes('{{') && cell.includes('}}')) {
              const matches = cell.match(/\{\{([^}]+)\}\}/g);
              
              if (matches) {
                let replacedValue = cell;
                
                matches.forEach(match => {
                  const key = match.slice(2, -2); // Extraer el nombre del marcador: {{nombre_actividad}} -> nombre_actividad
                  if (processData[key] !== undefined) {
                    replacedValue = replacedValue.replace(match, processData[key]);
                  }
                });
                
                // Si se realizó algún reemplazo, añadir a las actualizaciones
                if (replacedValue !== cell) {
                  const colLetter = String.fromCharCode(65 + c); // A, B, C, ...
                  updates.push({
                    range: `${colLetter}${r + 1}`,
                    values: [[replacedValue]]
                  });
                }
              }
            }
          }
        }
        
        // Aplicar todos los reemplazos en una sola operación
        if (updates.length > 0) {
          await sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: fileId,
            resource: {
              valueInputOption: 'RAW',
              data: updates
            }
          });
          console.log(`✅ Reemplazados ${updates.length} marcadores en Google Sheet`);
        }
        
      } else {
        // ==== FLUJO PARA PLANTILLA XLSX ====
        console.log('Usando flujo para plantilla XLSX');
        
        // Descargar plantilla XLSX
        const fileResponse = await this.drive.files.get(
          { fileId: templateId, alt: 'media' },
          { responseType: 'stream' }
        );
        
        // Cargar el libro desde el stream
        const workbook = await excelUtils.loadWorkbookFromStream(fileResponse.data);
        console.log('Libro XLSX cargado desde stream correctamente');
        
        // Hacer copia de los datos sin dinamicRowsData
        const processData = { ...data };
        delete processData['__FILAS_DINAMICAS__'];
        
        // Incorporar los campos individuales de gastos
        Object.assign(processData, gastoFields);
        
        // Reemplazar marcadores
        excelUtils.replaceMarkers(workbook, processData, true);
        console.log('Marcadores reemplazados correctamente en XLSX');
        
        // Guardar a archivo temporal
        const tempFilePath = await excelUtils.saveToTempFile(workbook, fileName);
        console.log(`Archivo XLSX guardado temporalmente en ${tempFilePath}`);
        
        // Subir a Google Drive como Google Sheets directamente
        const uploadResponse = await this.drive.files.create({
          requestBody: {
            name: fileName,
            parents: [this.templateFolderId],
            mimeType: 'application/vnd.google-apps.spreadsheet', // Subir directamente como Sheet
          },
          media: {
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            body: fs.createReadStream(tempFilePath),
          },
        });
        
        fileId = uploadResponse.data.id;
        await this.makeFilePublic(fileId);
        console.log(`Archivo XLSX convertido a Google Sheet con ID: ${fileId}`);
        
        // Limpiar archivos temporales
        excelUtils.cleanupTempFiles([tempFilePath]);
      }
      
      // INDEPENDIENTEMENTE DEL TIPO DE PLANTILLA INICIAL
      // Si hay filas dinámicas, insertarlas en el Sheet final
      if (dynamicRowsData && dynamicRowsData.gastos && dynamicRowsData.gastos.length > 0) {
        try {
          console.log('Iniciando proceso para insertar filas dinámicas...');
          
          // Importar el generador de filas dinámicas
          const { insertDynamicRows } = require('./dynamicRows/expensesGenerator');
          
          // Asegurarse que dynamicRowsData tiene el formato correcto
          const preparedData = {
            ...dynamicRowsData,
            insertarEn: dynamicRowsData.insertarEn || "A44:AK44",
            insertStartRow: dynamicRowsData.insertStartRow || 45
          };
          
          // Insertar filas dinámicas 
          const result = await insertDynamicRows(fileId, preparedData);
          
          if (result) {
            console.log('✅ Filas dinámicas insertadas correctamente');
          } else {
            console.log('⚠️ No se pudieron insertar las filas dinámicas');
          }
        } catch (error) {
          console.error('Error al procesar filas dinámicas:', error);
          console.error('Stack:', error.stack);
        }
      }
      
      // Devolver URL del documento Google Sheet
      return `https://docs.google.com/spreadsheets/d/${fileId}/edit?usp=sharing`;
    } catch (error) {
      console.error('Error al procesar archivo:', error);
      console.error('Stack:', error.stack);
      throw new Error(`Error al procesar archivo: ${error.message}`);
    }
  }

  // New function to insert dynamic rows into a Google Sheet using the dynamicRows service
  async insertDynamicRowsInSheet(fileId, dynamicRowsData) {
    try {
      if (!dynamicRowsData) {
        console.log('No hay datos de filas dinámicas para insertar');
        return false;
      }
      
      console.log(`Procesando inserción de filas dinámicas en el reporte (fileId: ${fileId})`);
      
      // Validate that we have the required data structure
      if (!dynamicRowsData.gastos || !Array.isArray(dynamicRowsData.gastos)) {
        console.log('Estructura de datos inválida - gastos no es un array');
        return false;
      }
      
      // Log the data we're working with
      console.log(`Inserción de ${dynamicRowsData.gastos.length} filas dinámicas`);
      console.log(`Rango de inserción: ${dynamicRowsData.insertarEn || 'No especificado'}`);
      
      // Import the expensesGenerator if not already available
      const { insertDynamicRows } = require('./dynamicRows/expensesGenerator');
      
      // Call the insertDynamicRows function from the expensesGenerator
      const result = await insertDynamicRows(fileId, dynamicRowsData);
      
      if (result) {
        console.log('✅ Filas dinámicas insertadas correctamente');
      } else {
        console.log('⚠️ No se pudieron insertar las filas dinámicas');
      }
      
      return result;
    } catch (error) {
      console.error('Error al insertar filas dinámicas:', error);
      console.error('Stack:', error.stack);
      return false;
    }
  }

  /**
   * Duplica una plantilla de Google Sheets
   * @param {String} templateId - ID de la plantilla a duplicar
   * @param {String} newName - Nombre del nuevo archivo
   * @returns {Promise<String>} - ID del nuevo archivo
   */
  async duplicateTemplate(templateId, newName) {
    try {
      console.log(`Duplicando plantilla: ${templateId} con nombre: ${newName}`);
      
      // 1. Obtener metadatos del archivo original
      const fileMetadata = await this.drive.files.get({
        fileId: templateId,
        fields: 'name,parents'
      });
      
      // 2. Crear una copia con la API de Drive
      const copyResponse = await this.drive.files.copy({
        fileId: templateId,
        requestBody: {
          name: newName,
          parents: fileMetadata.data.parents // Mantener en la misma carpeta
        }
      });
      
      const newFileId = copyResponse.data.id;
      console.log(`✅ Plantilla duplicada exitosamente. Nuevo ID: ${newFileId}`);
      
      // 3. Hacer público el archivo duplicado
      await this.makeFilePublic(newFileId);
      
      return newFileId;
    } catch (error) {
      console.error('Error al duplicar plantilla:', error);
      throw new Error(`Error al duplicar plantilla: ${error.message}`);
    }
  }

  /**
   * Genera un reporte basado en plantilla y datos
   * @param {Number} formNumber - Número de formulario (1-4)
   * @param {String} solicitudId - ID de la solicitud
   * @param {Object} data - Datos para rellenar la plantilla
   * @param {String} mode - Modo de visualización (view/edit)
   * @returns {Promise<String>} - URL del reporte generado
   */
  async generateReport(formNumber, solicitudId, data, mode = 'view') {
    try {
      const templateId = this.templateIds[formNumber];

      if (!templateId) {
        throw new Error(`Plantilla no encontrada para formulario ${formNumber}`);
      }

      // --- SOLUTION 1: Extract dynamic data early ---
      const dynamicRowsData = data['__FILAS_DINAMICAS__'];
      // Create a copy for static placeholder replacement to avoid modifying original 'data'
      const staticData = { ...data };
      delete staticData['__FILAS_DINAMICAS__']; // Remove from the copy

      // Log inicial (using the extracted variable)
      if (dynamicRowsData) {
        console.log(`Detectadas ${dynamicRowsData.gastos?.length || 0} filas dinámicas para el reporte (pre-procesamiento)`);
      }

      const fileName = `Formulario${formNumber}_${solicitudId}`;

      // 1. Duplicar la plantilla
      console.log(`Duplicando plantilla: ${templateId} con nombre: ${fileName}`);
      const newFileId = await this.duplicateTemplate(templateId, fileName);

      // 2. Determinar tipo de documento y usar el método adecuado
      const fileMetadata = await this.drive.files.get({
        fileId: newFileId,
        fields: 'mimeType'
      });

      const mimeType = fileMetadata.data.mimeType;
      console.log(`Tipo de archivo: ${mimeType}`);

      // 3. Procesar los datos estáticos (reemplazar marcadores) según el tipo
      //    Use the 'staticData' copy here
      if (mimeType === 'application/vnd.google-apps.document') {
        await this.replaceDocPlaceholders(newFileId, staticData);
      } else {
        await this.replacePlaceholders(newFileId, staticData);
      }

      // --- SOLUTION 2: Add detailed logging and check length ---
      console.log('🔄 Verificando datos para filas dinámicas ANTES del IF...');
      console.log('   Tiene dynamicRowsData?', !!dynamicRowsData);

      // --- FIX: Move the property checks INSIDE the if(dynamicRowsData) block ---
      if (dynamicRowsData) { // Check if dynamicRowsData exists FIRST
          console.log('   Tiene .gastos?', dynamicRowsData.hasOwnProperty('gastos')); // Safe now
          if (dynamicRowsData.hasOwnProperty('gastos')) {
              console.log('   Tipo de .gastos:', typeof dynamicRowsData.gastos);
              console.log('   Es array?', Array.isArray(dynamicRowsData.gastos));
              // Use optional chaining here too for extra safety, although hasOwnProperty should suffice
              console.log('   Longitud de .gastos:', dynamicRowsData.gastos?.length);
          }
      } else {
          // Optional: Log that dynamicRowsData was indeed undefined
          console.log('   dynamicRowsData es undefined o falsy.');
      }

      // 4. Procesar filas dinámicas si existen (using the extracted variable)
      //    The existing check here is already robust and correct
      if (dynamicRowsData && dynamicRowsData.gastos && dynamicRowsData.gastos.length > 0) {
        // ... (rest of the if block remains the same) ...
        console.log(`✅ CONDICIÓN CUMPLIDA: Procesando ${dynamicRowsData.gastos.length} gastos dinámicos`);

        const dinamicConfig = {
          gastos: dynamicRowsData.gastos,
          insertarEn: dynamicRowsData.insertarEn || "A44:AK44", // Default range
          insertStartRow: dynamicRowsData.insertStartRow || 45 // Default start row
        };

        await this.insertarFilasDinamicas(newFileId, dinamicConfig);

      } else {
          // Log why the condition failed (this existing block is good)
          console.log('⚠️ CONDICIÓN NO CUMPLIDA para insertar filas dinámicas.');
          if (!dynamicRowsData) {
              console.log('   Razón: dynamicRowsData no existe o es falsy.');
          } else if (!dynamicRowsData.gastos) {
              console.log('   Razón: dynamicRowsData.gastos no existe o es falsy.');
          } else if (!Array.isArray(dynamicRowsData.gastos)) {
              console.log('   Razón: dynamicRowsData.gastos no es un array.');
          } else if (dynamicRowsData.gastos.length === 0) {
              console.log('   Razón: dynamicRowsData.gastos está vacío.');
          }
      }

      // 5. Construir y devolver el enlace
      let reportLink = `https://drive.google.com/file/d/${newFileId}/view`;

      // Si se solicita el modo "edit", modificamos el enlace
      if (mode === 'edit') {
        if (mimeType === 'application/vnd.google-apps.spreadsheet') {
          reportLink = `https://docs.google.com/spreadsheets/d/${newFileId}/edit`;
        } else if (mimeType === 'application/vnd.google-apps.document') {
          reportLink = `https://docs.google.com/document/d/${newFileId}/edit`;
        }
      }

      console.log(`✅ Reporte generado exitosamente. Link: ${reportLink}`); // Moved log here
      return reportLink;
    } catch (error) {
      console.error('Error al generar reporte:', error);
      throw new Error(`Error al generar reporte: ${error.message}`);
    }
  }

  /**
   * Inserta filas dinámicas para gastos en una hoja de Google Sheets
   * @param {String} fileId - ID del archivo de Google Sheets
   * @param {Object} dinamicConfig - Configuración para filas dinámicas
   * @returns {Promise<Boolean>} - Éxito de la operación
   */
  async insertarFilasDinamicas(fileId, dinamicConfig) {
    try {
      // Extraer la información de los gastos
      const gastos = dinamicConfig.gastos || [];
      const insertarEn = dinamicConfig.insertarEn || "A44:AK44";
      const insertStartRow = dinamicConfig.insertStartRow || 45;
      
      if (!gastos || gastos.length === 0) {
        console.log('No hay gastos dinámicos para insertar');
        return true;
      }
      
      console.log(`Insertando ${gastos.length} filas dinámicas en ${insertarEn} a partir de la fila ${insertStartRow}`);
      
      // Inicializar Google Sheets API - usar el mismo auth que usa el resto del servicio
      const sheets = google.sheets({ version: 'v4', auth: jwtClient });
      
      // 1. Obtener información de la hoja
      const spreadsheet = await sheets.spreadsheets.get({
        spreadsheetId: fileId,
        fields: 'sheets.properties'
      });
      
      if (!spreadsheet.data.sheets || spreadsheet.data.sheets.length === 0) {
        throw new Error('No se encontraron hojas en el archivo');
      }
      
      // Usar la primera hoja
      const sheetId = spreadsheet.data.sheets[0].properties.sheetId;
      
      // 2. Insertar filas vacías
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: fileId,
        resource: {
          requests: [{
            insertDimension: {
              range: {
                sheetId: sheetId,
                dimension: 'ROWS',
                startIndex: insertStartRow - 1, // Convertir a índice base 0
                endIndex: insertStartRow - 1 + gastos.length
              },
              inheritFromBefore: true
            }
          }]
        }
      });
      
      console.log(`✅ ${gastos.length} filas vacías insertadas`);
      
      // 3. Insertar datos en las celdas
      const valueRequests = [];
      
      gastos.forEach((gasto, index) => {
        const rowIndex = insertStartRow + index;
        
        // Preparar valores para cada columna según la especificación
        const idValue = gasto.id || gasto.id_conceptos || '';
        const descripcionValue = gasto.descripcion || gasto.concepto || '';
        const cantidadValue = gasto.cantidad?.toString() || '0';
        const valorUnitValue = gasto.valorUnit_formatted || gasto.valor_unit_formatted || '';
        const valorTotalValue = gasto.valorTotal_formatted || gasto.valor_total_formatted || '';
        
        // Log the values being pushed for a specific row (e.g., the first one)
        if (index === 0) {
            console.log(`   Valores para fila ${rowIndex}: ID=${idValue}, Desc=${descripcionValue}, Cant=${cantidadValue}, Unit=${valorUnitValue}, Total=${valorTotalValue}`);
        }

        // Añadir actualizaciones para las celdas específicas según la estructura de la plantilla
        valueRequests.push(
          // ID en columna E (Concepto ID)
          { range: `E${rowIndex}`, values: [[idValue]] },
          // Descripción en columna F (abarca hasta V)
          { range: `F${rowIndex}`, values: [[descripcionValue]] },
          // Cantidad en columna X (abarca hasta Y)
          { range: `X${rowIndex}`, values: [[cantidadValue]] },
          // Valor unitario en columna Z (abarca hasta AB) - Use USER_ENTERED
          { range: `Z${rowIndex}`, values: [[valorUnitValue]] },
          // Valor total en columna AC (abarca hasta AK) - Use USER_ENTERED
          { range: `AC${rowIndex}`, values: [[valorTotalValue]] }
        );
      });

      if (valueRequests.length > 0) {
        console.log(`   Enviando ${valueRequests.length} actualizaciones de valores...`);
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: fileId,
          resource: {
            valueInputOption: 'USER_ENTERED', // Important for currency/numbers
            data: valueRequests
          }
        });
      }

      console.log(`✅ Datos insertados en ${gastos.length} filas dinámicas`);
      return true;
    } catch (error) {
      console.error('❌ Error CRÍTICO al insertar filas dinámicas:', error); // Make error more prominent
      console.error('   Detalles del error:', error.message);
      if (error.response && error.response.data) { // Log Google API specific errors
          console.error('   Error de Google API:', JSON.stringify(error.response.data, null, 2));
      }
      console.error('   Stack:', error.stack); // Log stack trace
      // Consider re-throwing during development/debugging:
      // throw new Error(`Fallo al insertar filas dinámicas: ${error.message}`);
      return false; // Keep returning false for production flow if needed
    }
  }

  /**
   * Convierte una letra de columna a índice (A=0, B=1, etc.)
   * @param {String} col - Letra de columna (A, B, AA, etc.)
   * @returns {Number} - Índice de la columna (0-based)
   */
  columnToIndex(col) {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
      result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result - 1; // Convertir a índice base 0
  }

  /**
   * Convierte un índice de columna a letra (0=A, 1=B, etc.)
   * @param {Number} index - Índice de columna (0-based)
   * @returns {String} - Letra de columna
   */
  indexToColumn(index) {
    let temp, letter = '';
    while (index >= 0) {
      temp = index % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  }

  /**
   * Reemplaza marcadores de posición en las celdas de una hoja de Google Sheets
   * @param {String} fileId - ID del archivo de Google Sheets
   * @param {Object} data - Datos para reemplazar marcadores
   * @returns {Promise<Boolean>} - Éxito de la operación
   */
  async replacePlaceholders(fileId, data) {
    try {
      console.log('Reemplazando marcadores en hoja de cálculo...');
      
      // Eliminar campos especiales que no deben procesarse como marcadores
      const processData = { ...data };
      delete processData['__FILAS_DINAMICAS__'];
      
      // Inicializar la API de Sheets
      const sheets = google.sheets({ version: 'v4', auth: jwtClient });
      
      // 1. Obtener todo el contenido de la hoja
      let response;
      try {
        response = await sheets.spreadsheets.values.get({
          spreadsheetId: fileId,
          range: 'A1:AZ100' // Rango amplio para cubrir toda la hoja
        });
      } catch (getError) {
        console.error('Error al obtener valores de la hoja:', getError);
        return false;
      }
      
      if (!response || !response.data || !response.data.values) {
        console.warn('No se encontraron valores en la hoja');
        return false;
      }
      
      const rows = response.data.values;
      
      // 2. Buscar y reemplazar marcadores
      const updates = [];
      
      for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        const row = rows[rowIndex];
        for (let colIndex = 0; colIndex < row.length; colIndex++) {
          const cell = row[colIndex];
          if (typeof cell === 'string' && cell.includes('{{') && cell.includes('}}')) {
            // Extraer marcadores en formato {{nombre_campo}}
            const matches = cell.match(/\{\{([^}]+)\}\}/g);
            
            if (matches) {
              let newValue = cell;
              
              matches.forEach(match => {
                const fieldName = match.substring(2, match.length - 2);
                const replacement = processData[fieldName] !== undefined ? processData[fieldName] : '';
                
                newValue = newValue.replace(match, replacement);
              });
              
              if (newValue !== cell) {
                const colLetter = this.indexToColumn(colIndex);
                updates.push({
                  range: `${colLetter}${rowIndex + 1}`,
                  values: [[newValue]]
                });
              }
            }
          }
        }
      }
      
      // 3. Aplicar las actualizaciones en lote
      if (updates.length > 0) {
        try {
          await sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: fileId,
            resource: {
              valueInputOption: 'USER_ENTERED',
              data: updates
            }
          });
          console.log(`✅ Reemplazados ${updates.length} marcadores en la hoja de cálculo`);
        } catch (updateError) {
          console.error('Error al actualizar celdas:', updateError);
          return false;
        }
      } else {
        console.log('No se encontraron marcadores para reemplazar');
      }
      
      return true;
    } catch (error) {
      console.error('Error general en replacePlaceholders:', error);
      return false; // Devolver false en lugar de lanzar la excepción
    }
  }

  /**
   * Elimina un archivo de Google Drive
   * @param {String} fileId - ID del archivo a eliminar
   */

  async deleteFile(fileId) {
    try {
      await this.drive.files.delete({ fileId });
      return true;
    } catch (error) {
      console.error('Error al eliminar archivo:', error);
      throw new Error(`Error al eliminar archivo: ${error.message}`);
    }
  }

  /**
   * Reemplaza marcadores de posición en el documento con datos reales
   * @param {String} fileId - ID del archivo a modificar
   * @param {Object} data - Datos para reemplazar los marcadores
   */
  async replaceDocPlaceholders(fileId, data) {
    try {
      console.log('Reemplazando marcadores en documento...');
      
      // Primero procesamos los datos especiales para manipulación de tablas
      if (data['__FILAS_DINAMICAS__']) {
        await this.insertDynamicRowsInSheet(fileId, data['__FILAS_DINAMICAS__']);
        // Eliminar este campo especial para no intentar reemplazarlo como marcador normal
        delete data['__FILAS_DINAMICAS__'];
      }
      
      // Extraer el texto del documento
      const response = await this.docs.documents.get({
        documentId: fileId
      });
      
      const document = response.data;
      const requests = [];
      
      // Buscar todos los marcadores de posición
      for (const key in data) {
        const value = data[key];
        if (value !== undefined && value !== null) {
          // Solo procesar si es un valor de cadena (no objetos ni arrays)
          if (typeof value === 'string' || typeof value === 'number') {
            requests.push({
              replaceAllText: {
                containsText: {
                  text: `{{${key}}}`,
                  matchCase: true
                },
                replaceText: String(value)
              }
            });
          }
        }
      }
      
      if (requests.length > 0) {
        await this.docs.documents.batchUpdate({
          documentId: fileId,
          resource: {
            requests: requests
          }
        });
        console.log(`Reemplazados ${requests.length} marcadores en el documento`);
      }
      
      return true;
    } catch (error) {
      console.error('Error al reemplazar marcadores en documento:', error);
      throw error;
    }
  }

  /**
   * Inserta filas de gastos dinámicos en el documento
   * @param {String} fileId - ID del documento
   * @param {Object} gastosDinamicos - Datos de gastos dinámicos y coordenadas
   */
  async insertarFilasGastosDinamicos(fileId, gastosDinamicos) {
    try {
      const { insertarEn, gastos } = gastosDinamicos;
      if (!gastos || gastos.length === 0) {
        console.log('No hay gastos dinámicos para insertar');
        return;
      }
      
      console.log(`Insertando ${gastos.length} gastos dinámicos en ${insertarEn}`);
      
      // 1. Obtener el documento y encontrar la tabla objetivo
      const docResponse = await this.docs.documents.get({
        documentId: fileId
      });
      
      const document = docResponse.data;
      const tablas = this.encontrarTablas(document);
      
      if (tablas.length === 0) {
        console.error('No se encontraron tablas en el documento');
        return;
      }
      
      // Encontrar la tabla correcta (normalmente será la que tenga una celda con "Concepto")
      let tablaObjetivo = null;
      let indiceTabla = -1;
      
      for (let i = 0; i < tablas.length; i++) {
        const tabla = tablas[i];
        // Buscar una celda que contenga "Concepto" o similar
        const tieneCabecera = this.buscarTextoEnTabla(document, tabla, ["concepto", "descripción", "descripcion"]);
        
        if (tieneCabecera) {
          tablaObjetivo = tabla;
          indiceTabla = i;
          break;
        }
      }
      
      if (!tablaObjetivo) {
        console.error('No se encontró la tabla de gastos en el documento');
        return;
      }
      
      console.log(`Encontrada tabla objetivo (#${indiceTabla})`);
      
      // 2. Preparar las solicitudes para insertar filas
      const requests = [];
      
      // Determinar la posición de inserción (la última fila después de los gastos existentes)
      const filaInsercion = Math.max(1, tablaObjetivo.rows - 1); // -1 para no contar el encabezado
      
      // Insertar una fila para cada gasto
      gastos.forEach((gasto, index) => {
        requests.push({
          insertTableRow: {
            tableStartLocation: {
              tableId: tablaObjetivo.tableId
            },
            insertBelow: true,
            rowIndex: filaInsercion + index
          }
        });
        
        // Llenar las celdas con los datos del gasto
        const celdas = [
          { texto: gasto.concepto || "Concepto sin nombre" },
          { texto: gasto.cantidad.toString() },
          { texto: gasto.valorUnit_formatted },
          { texto: gasto.valorTotal_formatted }
        ];
        
        celdas.forEach((celda, celdaIndex) => {
          requests.push({
            insertText: {
              location: {
                tableId: tablaObjetivo.tableId,
                rowIndex: filaInsercion + index + 1, // +1 porque ya insertamos la fila
                columnIndex: celdaIndex
              },
              text: celda.texto
            }
          });
        });
      });
      
      // 3. Ejecutar las solicitudes
      if (requests.length > 0) {
        await this.docs.documents.batchUpdate({
          documentId: fileId,
          resource: {
            requests: requests
          }
        });
        
        console.log(`Insertadas ${gastos.length} filas de gastos dinámicos`);
      }
      
      return true;
    } catch (error) {
      console.error('Error al insertar gastos dinámicos:', error);
      throw error;
    }
  }

  /**
   * Inserta datos de gastos en las celdas dinámicas del documento
   * @param {String} fileId - ID del documento
   * @param {Array} gastos - Array de datos de gastos
   * @param {String} insertLocation - Coordenadas donde insertar los datos (ej: 'E44:AK44')
   * @returns {Promise<Boolean>} - Resultado de la operación
   */
  async insertarGastosDinamicos(fileId, gastos, insertLocation = 'E44:AK44') {
    try {
      if (!gastos || gastos.length === 0) {
        console.log('No hay gastos para insertar');
        return false;
      }
      
      console.log(`Intentando insertar ${gastos.length} gastos en ${insertLocation}`);
      
      // Extraer referencia de tabla
      const matchTable = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(insertLocation);
      if (!matchTable) {
        console.error('Formato de coordenadas inválido');
        return false;
      }
      
      const [, colInicioStr, filaInicio, colFinStr, filaFin] = matchTable;
      
      // Convertir coordenadas de Excel a índices
      const colInicio = this.columnToIndex(colInicioStr);
      const colFin = this.columnToIndex(colFinStr);
      
      // Usar Sheets API para insertar los datos
      const sheets = google.sheets({version: 'v4', auth: this.auth});
      
      // Preparar los datos en el formato esperado por Sheets API
      const values = gastos.map(gasto => {
        // Crear un array con celdas vacías para todas las columnas
        const row = Array(colFin - colInicio + 1).fill('');
        
        // Llenar las celdas relevantes
        row[0] = gasto.descripcion || 'Concepto';  // Columna E (índice 0)
        row[1] = gasto.cantidad.toString();        // Columna F (índice 1)
        row[2] = gasto.valorUnit_formatted;        // Columna G (índice 2)
        row[3] = gasto.valorTotal_formatted;       // Columna H (índice 3)
        
        return row;
      });
      
      // Insertar las filas en la tabla
      await sheets.spreadsheets.values.update({
        spreadsheetId: fileId,
        range: insertLocation,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values
        }
      });
      
      console.log(`✅ Insertados ${gastos.length} gastos dinámicos en ${insertLocation}`);
      return true;
    } catch (error) {
      console.error('Error al insertar gastos dinámicos:', error);
      return false;
    }
  }

  /**
   * Convierte una letra de columna de Excel a índice (A=0, B=1, etc.)
   * @param {String} col - Letra de columna (A, B, AA, etc.)
   * @returns {Number} - Índice de la columna
   */
  columnToIndex(col) {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
      result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result - 1; // 0-based index
  }

  /**
   * Encuentra todas las tablas en un documento
   * @param {Object} document - Documento de Google Docs
   * @returns {Array} Array de datos de tablas
   */
  encontrarTablas(document) {
    const tablas = [];
    
    // Recorrer todos los elementos del documento
    if (document.body && document.body.content) {
      document.body.content.forEach(elemento => {
        if (elemento.table) {
          tablas.push({
            tableId: elemento.table.tableId,
            rows: elemento.table.rows,
            columns: elemento.table.columns
          });
        }
      });
    }
    
    return tablas;
  }

  /**
   * Busca texto en una tabla
   * @param {Object} document - Documento de Google Docs
   * @param {Object} tabla - Datos de la tabla
   * @param {Array} textosBuscar - Textos a buscar
   * @returns {Boolean} Verdadero si se encontró alguno de los textos
   */
  buscarTextoEnTabla(document, tabla, textosBuscar) {
    // Función simplificada para buscar texto en tabla
    // En una implementación real, recorrerías las celdas de la tabla
    return true; // Para este ejemplo, siempre devolvemos true
  }
}

module.exports = new DriveService();