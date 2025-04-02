const { google } = require('googleapis');
const { jwtClient } = require('../config/google');
const excelUtils = require('../utils/excelUtils');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const os = require('os');

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
      2: '1JY-4IfJqEWLqZ_wrq_B_bfIlI9MeVzgF', //Plantilla de formulario 2
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
   * Procesa un archivo XLSX sustituyendo marcadores con datos
   * @param {String} templateId - ID de la plantilla
   * @param {Object} data - Datos para reemplazar marcadores
   * @param {String} fileName - Nombre del archivo resultante
   * @returns {Promise<String>} - URL del archivo resultante
   */

  async processXLSXWithStyles(templateId, data, fileName) {
    try {
      console.log(`Descargando la plantilla: ${templateId}`);
      // Determinar el número de formulario basado en el templateId
      let formNumber = null;
      for (const [key, value] of Object.entries(this.templateIds)) {
        if (value === templateId) {
          formNumber = key;
          break;
        }
      }
      
      const fileResponse = await this.drive.files.get(
        { fileId: templateId, alt: 'media' },
        { responseType: 'stream' }
      );
  
      // Cargar el libro desde el stream
      const workbook = await excelUtils.loadWorkbookFromStream(fileResponse.data);
      console.log('Libro cargado desde stream correctamente');
  
      // Store the dynamic rows data for later processing
      const dynamicRowsData = data['__FILAS_DINAMICAS__'];
      
      // Extraer campos de gastos individuales antes de eliminar __FILAS_DINAMICAS__
      const gastoFields = {};
      if (dynamicRowsData && dynamicRowsData.gastos) {
        dynamicRowsData.gastos.forEach((gasto, index) => {
          // Crear marcadores individuales para cada gasto (con coma y con punto)
          const idConComa = gasto.id?.replace('.', ',') || `1,${index + 1}`;
          const idConPunto = gasto.id?.replace(',', '.') || `1.${index + 1}`;
          
          // Versión con coma para Excel
          gastoFields[`gasto_${idConComa}_cantidad`] = gasto.cantidad?.toString() || '0';
          gastoFields[`gasto_${idConComa}_valor_unit`] = gasto.valorUnit_formatted || '0';
          gastoFields[`gasto_${idConComa}_valor_total`] = gasto.valorTotal_formatted || '0';
          gastoFields[`gasto_${idConComa}_descripcion`] = gasto.descripcion || '';
          
          // Versión con punto (más estándar)
          gastoFields[`gasto_${idConPunto}_cantidad`] = gasto.cantidad?.toString() || '0';
          gastoFields[`gasto_${idConPunto}_valor_unit`] = gasto.valorUnit_formatted || '0';
          gastoFields[`gasto_${idConPunto}_valor_total`] = gasto.valorTotal_formatted || '0';
          gastoFields[`gasto_${idConPunto}_descripcion`] = gasto.descripcion || '';
        });
      }
      
      // Remove special field to avoid processing it as a normal marker
      delete data['__FILAS_DINAMICAS__'];
  
      // Incorporar los campos individuales de gastos al objeto data
      Object.assign(data, gastoFields);
      
      // Reemplazar marcadores
      excelUtils.replaceMarkers(workbook, data);
      console.log('Marcadores reemplazados correctamente');
  
      // Subir a Google Drive primero para obtener el ID
      const tempFilePath = await excelUtils.saveToTempFile(workbook, fileName);
      console.log(`Archivo guardado temporalmente en ${tempFilePath}`);
  
      const uploadResponse = await this.drive.files.create({
        requestBody: {
          name: fileName,
          parents: [this.templateFolderId],
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
        media: {
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          body: fs.createReadStream(tempFilePath),
        },
      });
  
      const fileId = uploadResponse.data.id;
      await this.makeFilePublic(fileId);
      
      // Procesar filas dinámicas si existe el método y está configurado
      if (formNumber) {
        try {
          // Usar import dinámico para evitar errores si el archivo no existe
          const reportConfig = await import(`../reportConfigs/report${formNumber}Config.js`)
            .catch(err => {
              console.log(`Configuración no encontrada para formulario ${formNumber}`);
              return null;
            });
            
          if (reportConfig && typeof reportConfig.processDynamicRows === 'function') {
            const sheets = google.sheets({version: 'v4', auth: this.jwtClient});
            console.log(`Procesando filas dinámicas para formulario ${formNumber}...`);
            await reportConfig.processDynamicRows(fileId, data, sheets);
          } else {
            // Usar el método genérico si no hay configuración específica
            if (dynamicRowsData && dynamicRowsData.gastos && dynamicRowsData.gastos.length > 0) {
              await this.insertDynamicRowsInSheet(fileId, dynamicRowsData);
            }
          }
        } catch (dynamicRowsError) {
          console.error('Error al procesar filas dinámicas:', dynamicRowsError);
          // Continuar con la generación aunque falle
        }
      }
  
      // Insertar filas dinámicas si es necesario
      if (dynamicRowsData && dynamicRowsData.gastos && dynamicRowsData.gastos.length > 0) {
        await this.insertDynamicRowsInSheet(fileId, dynamicRowsData);
      }
  
      // Limpiar archivos temporales
      excelUtils.cleanupTempFiles([tempFilePath]);
  
      return `https://drive.google.com/file/d/${fileId}/view`;
    } catch (error) {
      console.error('Error al procesar archivo XLSX:', error);
      throw new Error(`Error al procesar archivo XLSX: ${error.message}`);
    }
  }

  // New function to insert dynamic rows into a Google Sheet
  async insertDynamicRowsInSheet(fileId, dynamicRowsData) {
    try {
      const insertLocation = dynamicRowsData.insertarEn || 'E45:AK45'; // Start at row 45 (after example row 44)
      const gastos = dynamicRowsData.gastos || [];
      
      if (gastos.length === 0) {
        console.log('No hay gastos para insertar');
        return;
      }
      
      console.log(`Insertando ${gastos.length} gastos dinámicos en ${insertLocation}`);
      
      // Initialize Google Sheets API
      const sheets = google.sheets({version: 'v4', auth: jwtClient});
      
      // Parse the insertion location to get row and column info
      const match = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(insertLocation);
      if (!match) {
        console.error('Formato de ubicación inválido:', insertLocation);
        return;
      }
      
      const [_, startCol, startRow] = match;
      const startRowNum = parseInt(startRow);
      
      // Prepare the data for each row
      const values = gastos.map((gasto, index) => {
        // Create a row with enough cells to span from E to AK (37 columns)
        const row = new Array(37).fill('');
        
        // Fill in the specific cells according to the requirements:
        // ID in column E (index 0)
        row[0] = gasto.id_concepto || `15.${index + 1}`;
        
        // Description in columns F to V (indices 1-21)
        row[1] = gasto.descripcion || '';
        
        // Quantity in columns X to Z (indices 23-25)
        row[23] = gasto.cantidad?.toString() || '0';
        
        // Unit value in columns Z to AB (indices 25-27)
        row[25] = gasto.valor_unit_formatted || '0';
        
        // Total value in columns AC to AK (indices 28-36)
        row[28] = gasto.valor_total_formatted || '0';
        
        return row;
      });
      
      // Insert data into each row, starting at the row after the example (row 45)
      for (let i = 0; i < gastos.length; i++) {
        const rowNum = startRowNum + i;
        await sheets.spreadsheets.values.update({
          spreadsheetId: fileId,
          range: `E${rowNum}:AK${rowNum}`,
          valueInputOption: 'USER_ENTERED',
          resource: {
            values: [values[i]]
          }
        });
      }
      
      console.log(`✅ Insertados ${gastos.length} gastos dinámicos en el reporte`);
      return true;
    } catch (error) {
      console.error('Error al insertar filas dinámicas:', error);
      return false;
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
      
      const fileName = `Formulario${formNumber}_${solicitudId}`;
      const reportLink = await this.processXLSXWithStyles(templateId, data, fileName);
      
      // Si se solicita el modo "edit", modificamos el enlace
      if (mode === 'edit') {
        return reportLink.replace('/view', '/view?usp=drivesdk&edit=true');
      }
      
      return reportLink;
    } catch (error) {
      console.error('Error al generar reporte:', error);
      throw new Error(`Error al generar reporte: ${error.message}`);
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
  async replacePlaceholders(fileId, data) {
    try {
      console.log("📊 INICIO DE REEMPLAZO DE PLACEHOLDERS");

      // Combinar los datos originales con los valores de prueba
      const mergedData = { ...data, ...testData };
      
      // PASO 2: Analizar la plantilla para encontrar todos los placeholders
      const docResponse = await this.docs.documents.get({
        documentId: fileId
      });
      
      // Extraer todos los placeholders del documento para ver qué formato usan
      const docText = JSON.stringify(docResponse.data);
      const placeholderRegex = /\{\{([^}]+)\}\}/g;
      const placeholdersFound = [];
      let match;
      
      while ((match = placeholderRegex.exec(docText)) !== null) {
        const placeholder = match[0];
        const fieldName = match[1].trim();
        placeholdersFound.push({ placeholder, fieldName });
      }
      
      console.log("🔍 PLACEHOLDERS ENCONTRADOS:", placeholdersFound.length);
      // Mostrar solo los primeros 10 para no saturar los logs
      placeholdersFound.slice(0, 10).forEach(p => {
        console.log(`- ${p.placeholder} => ${mergedData[p.fieldName] || "NO TIENE VALOR"}`);
      });
      
      // PASO 3: Procesar cada placeholder con un enfoque más directo
      const requests = [];
      
      placeholdersFound.forEach(p => {
        // Buscar el valor usando el nombre de campo exacto
        let value = mergedData[p.fieldName] || '';
        
        // Si no se encuentra, probar con variantes (sin espacios, sin caracteres especiales)
        if (!value) {
          // Variante 1: sin espacios
          const fieldNoSpaces = p.fieldName.replace(/\s+/g, '');
          if (mergedData[fieldNoSpaces]) {
            value = mergedData[fieldNoSpaces];
            console.log(`  ✅ Encontrado valor en variante sin espacios: ${fieldNoSpaces}`);
          }
          
          // Variante 2: con puntos en lugar de comas
          else if (p.fieldName.includes(',')) {
            const fieldWithDots = p.fieldName.replace(/,/g, '.');
            if (mergedData[fieldWithDots]) {
              value = mergedData[fieldWithDots];
              console.log(`  ✅ Encontrado valor en variante con puntos: ${fieldWithDots}`);
            }
          }
          
          // Variante 3: sin caracteres especiales
          else if (p.fieldName.includes('%') || p.fieldName.includes(',')) {
            const fieldSimplified = p.fieldName.replace(/[%,]/g, '');
            if (mergedData[fieldSimplified]) {
              value = mergedData[fieldSimplified];
              console.log(`  ✅ Encontrado valor en variante simplificada: ${fieldSimplified}`);
            }
          }
        }
        
        // Una vez tenemos el valor, crear la solicitud de reemplazo
        if (value !== '') {
          console.log(`  🔄 Reemplazando ${p.placeholder} con "${value}"`);
          requests.push({
            replaceAllText: {
              containsText: { text: p.placeholder, matchCase: true },
              replaceText: value
            }
          });
        } else {
          console.log(`  ⚠️ No se encontró valor para ${p.placeholder}`);
        }
      });
      
      // PASO 4: Ejecutar todos los reemplazos en una sola operación batch
      if (requests.length > 0) {
        console.log(`📝 Ejecutando ${requests.length} reemplazos...`);
        await this.docs.documents.batchUpdate({
          documentId: fileId,
          resource: { requests }
        });
        console.log("✅ Reemplazos completados");
      } else {
        console.warn("⚠️ No se generaron solicitudes de reemplazo");
      }
      
      return fileId;
    } catch (error) {
      console.error("❌ Error al reemplazar placeholders:", error);
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

// Agregar en services/sheetsService.js
const cache = {};

async function getDataWithCache(key, fetcher, ttlMinutes = 5) {
  if (cache[key] && (Date.now() - cache[key].timestamp) < ttlMinutes * 60 * 1000) {
    console.log(`Usando datos en caché para ${key}`);
    return cache[key].data;
  }
  
  const data = await fetcher();
  cache[key] = { data, timestamp: Date.now() };
  return data;
}