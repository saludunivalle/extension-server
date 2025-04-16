const { google } = require('googleapis');
const { jwtClient } = require('../config/google')
const models = require('../models/spreadsheetModels');
const { GoogleAPIError } = require('../middleware/errorHandler');

/**
 * Servicio para manejar operaciones con Google Sheets
*/

class SheetsService {
  constructor() {
    this.spreadsheetId = process.env.SPREADSHEET_ID || '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';
    this.client = this.getClient();
    this.models = models.getModels(); // Usar los modelos definidos
  }

  async saveUserIfNotExists(userId, email, name) {
    try {
      const userCheckRange = 'USUARIOS!A2:A';
      const userCheckResponse = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: userCheckRange,
      });
  
      const existingUsers = userCheckResponse.data.values ? userCheckResponse.data.values.flat() : [];
      
      // Verificar si el usuario ya existe
      if (!existingUsers.includes(userId)) {
        const userRange = 'USUARIOS!A2:C2';
        const userValues = [[userId, email, name]];
  
        await this.client.spreadsheets.values.append({
          spreadsheetId: this.spreadsheetId,
          range: userRange,
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS',
          resource: { values: userValues },
        });
        return true; // Usuario nuevo añadido
      }
      return false; // Usuario ya existía
    } catch (error) {
      console.error('Error al guardar usuario:', error);
      throw new Error('Error al guardar usuario');
    }
  }

  // Método para obtener el cliente de Google Sheets  
  getClient() {
    return google.sheets({ version: 'v4', auth: jwtClient });
  }

   /**
    * Mapeos de columnas para diferentes formularios
   */
  //Definicion para cada hoja del sheets
  get fieldDefinitions() {
    const result = {};
    Object.entries(this.models).forEach(([name, model]) => {
      result[name] = model.fields;
    });
    return result;
  }

  // Obtener los mapeos de columnas de un modelo
  get columnMappings() {
    const result = {};
    Object.entries(this.models).forEach(([name, model]) => {
      result[name] = model.columnMappings;
    });
    return result;
  }

  //Obtención de ID en las hojas
  async getLastId(sheetName) {
    try {
      const response = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: `${sheetName}!A2:A`
      });

      const rows = response.data.values || [];
      const lastId = rows
        .map(row => parseInt(row[0], 10))
        .filter(id => !isNaN(id))
        .reduce((max, id) => Math.max(max, id), 0);

      return lastId;
    } catch (error) {
      console.error(`Error al obtener el último ID de ${sheetName}:`, error);
      throw new Error(`Error al obtener el último ID de ${sheetName}`);
    }
  }

  //Encuentra o crea una solicitud
  async findOrCreateRequestRow(sheetName, idSolicitud) {
    // Añadir caché para evitar solicitudes duplicadas
    const cacheKey = `${sheetName}_${idSolicitud}`;
    if (this.rowCache && this.rowCache[cacheKey]) {
      return this.rowCache[cacheKey];
    }
    
    try {
      const response = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: `${sheetName}!A:A`
      });

      const rows = response.data.values || [];
      let rowIndex = rows.findIndex((row) => row[0] === idSolicitud.toString());

      // Si no existe, crear una nueva fila
      if (rowIndex === -1) {
        rowIndex = rows.length + 1;
        await this.client.spreadsheets.values.append({
          spreadsheetId: this.spreadsheetId,
          range: `${sheetName}!A${rowIndex}`,
          valueInputOption: 'RAW',
          resource: { values: [[idSolicitud]] }
        });
      } else {
        rowIndex += 1; // Ajustar para que sea 1-based
      }

      const result = rowIndex;

      // Guardar en caché
      if (!this.rowCache) this.rowCache = {};
      this.rowCache[cacheKey] = result;
      
      return result;
    } catch (error) {
      console.error(`Error al buscar/crear fila en ${sheetName}:`, error);
      throw new Error(`Error al buscar/crear fila en ${sheetName}`);
    }
  }

  //Actualiza el progreso de una solicitud
  async updateRequestProgress(params) {
    const { sheetName, rowIndex, startColumn, endColumn, values } = params;
    
    try {
      return await this.client.spreadsheets.values.update({
        spreadsheetId: this.spreadsheetId,
        range: `${sheetName}!${startColumn}${rowIndex}:${endColumn}${rowIndex}`,
        valueInputOption: 'RAW',
        resource: { values: [values] }
      });
    } catch (error) {
      console.error('Error al actualizar progreso:', error);
      throw new Error('Error al actualizar progreso');
    }
  }

  /**
   * Obtiene todos los datos de una solicitud
  */
  async getSolicitudData(solicitudId,  definicionesHojas) {
    try {
      const hojas = definicionesHojas || {
        SOLICITUDES: {
          range: 'SOLICITUDES!A2:AU',
          fields: this.fieldDefinitions.SOLICITUDES
        },
        SOLICITUDES2: {
          range: 'SOLICITUDES2!A2:CL',
          fields: this.fieldDefinitions.SOLICITUDES2
        },
        // Agregar estas hojas:
        SOLICITUDES3: {
          range: 'SOLICITUDES3!A2:AC',
          fields: this.fieldDefinitions.SOLICITUDES3
        },
        SOLICITUDES4: {
          range: 'SOLICITUDES4!A2:BK',
          fields: this.fieldDefinitions.SOLICITUDES4
        },
        GASTOS: {
          range: 'GASTOS!A2:F',
          fields: this.fieldDefinitions.GASTOS
        }
      };

      const resultados = {};
      let solicitudEncontrada = false;

      for (let [hoja, { range, fields }] of Object.entries(hojas)) {
        const response = await this.client.spreadsheets.values.get({
          spreadsheetId: this.spreadsheetId,
          range
        });

        const rows = response.data.values || [];
        const solicitudData = rows.find(row => row[0] === solicitudId);

        if (solicitudData) {
          solicitudEncontrada = true;
          resultados[hoja] = fields.reduce((acc, field, index) => {
            acc[field] = solicitudData[index] || '';
            return acc;
          }, {});
        }
      }

      if (!solicitudEncontrada) {
        return { message: 'La solicitud no existe aún en Google Sheets', data: {} };
      }

      return resultados;
    } catch (error) {
      console.error('Error al obtener datos de solicitud:', error);
      throw new GoogleAPIError('Error al obtener datos de solicitud');
    }
  }

  async saveGastos(idSolicitud, gastos, actualizarConceptos = true) { // Keep the actualizarConceptos parameter if needed elsewhere, but logic relies on it being true here
    try {
      const conceptosValidos = new Set();
      const conceptosSolicitudValidos = new Set();
      // const conceptosPadre = new Map(); // Might not be needed if we trust frontend data

      // 1. Get existing concepts to avoid duplicates for the specific solicitud
      const conceptosResponse = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: 'CONCEPTO$!A2:F' // Read ID, Desc, Padre, Nombre, Tipo, SolicitudID
      });

      (conceptosResponse.data.values || []).forEach(row => {
        const conceptoId = String(row[0]);
        const solicitudIdConcepto = String(row[5] || ''); // Solicitud ID from CONCEPTO$ sheet

        // Add to sets for checking existence
        conceptosValidos.add(conceptoId);
        if (solicitudIdConcepto) {
            conceptosSolicitudValidos.add(`${conceptoId}:${solicitudIdConcepto}`);
        } else {
            // Handle concepts that might not have a solicitud ID (older data?)
             conceptosSolicitudValidos.add(`${conceptoId}:`); // Or handle as appropriate
        }
        // conceptosPadre.set(conceptoId, {descripcion: row[1] || '', esPadre: row[2] === 'true' || row[2] === 'TRUE'});
      });

      let nuevosConceptos = [];
      if (actualizarConceptos) {
        // 2. Identify ONLY the new concepts for THIS solicitud that need to be added to CONCEPTO$
        nuevosConceptos = gastos
          .filter(gasto => {
            // Check if this specific concept for this specific solicitud already exists
            // Ensure gasto.id_conceptos and idSolicitud are strings for comparison
            const key = `${String(gasto.id_conceptos)}:${String(idSolicitud)}`;
            return !conceptosSolicitudValidos.has(key);
          })
          .map(gasto => {
            // Directly use the data sent from the frontend for the new concept row
            // Ensure boolean es_padre is converted to string 'true'/'false'
            const esPadreStr = (typeof gasto.es_padre === 'boolean') ? (gasto.es_padre ? 'true' : 'false') : 'false';

            return [
              String(gasto.id_conceptos), // Col A: id_conceptos
              gasto.descripcion || gasto.nombre_conceptos || String(gasto.id_conceptos), // Col B: descripcion (use provided description/name)
              esPadreStr, // Col C: es_padre (use provided boolean, converted to string)
              gasto.nombre_conceptos || gasto.descripcion || String(gasto.id_conceptos), // Col D: nombre_conceptos (use provided name/label)
              gasto.tipo || 'estándar', // Col E: tipo (use provided type, default to estándar)
              String(idSolicitud) // Col F: id_solicitud
            ];
          });

        // 3. Guardar nuevos conceptos en CONCEPTO$ if any exist
        if (nuevosConceptos.length > 0) {
          console.log(`Adding ${nuevosConceptos.length} new concepts to CONCEPTO$ for solicitud ${idSolicitud}`);
          await this.client.spreadsheets.values.append({
            spreadsheetId: this.spreadsheetId,
            range: 'CONCEPTO$!A2:F', // Append to the end
            valueInputOption: 'RAW', // Use RAW to prevent Sheets from interpreting values
            insertDataOption: 'INSERT_ROWS',
            resource: { values: nuevosConceptos }
          });

          // 4. Update the set of known concepts locally to avoid trying to add them again if saveGastos is called multiple times quickly
          nuevosConceptos.forEach(conceptoRow => {
            const idConcepto = conceptoRow[0];
            const idSol = conceptoRow[5];
            conceptosValidos.add(idConcepto);
            conceptosSolicitudValidos.add(`${idConcepto}:${idSol}`);
          });
        } else {
           console.log(`No new concepts to add to CONCEPTO$ for solicitud ${idSolicitud}`);
        }
      } else {
         console.log("Skipping concept update based on actualizarConceptos flag.");
      }


      // 5. Prepare rows for GASTOS sheet (always do this for all valid expenses)
      const rows = gastos
        // Optional: Add filtering here again if you want to be extra sure only valid expenses go to GASTOS
        // .filter(gasto => (parseFloat(gasto.cantidad) || 0) > 0 && (parseFloat(gasto.valor_unit) || 0) > 0)
        .map(gasto => {
          const cantidad = parseFloat(gasto.cantidad) || 0;
          const valor_unit = parseFloat(gasto.valor_unit) || 0;
          // Use COMMA as the separator for consistency when determining concepto_padre for GASTOS
          const conceptoPadre = String(gasto.id_conceptos).includes(',') ?
            String(gasto.id_conceptos).split(',')[0] : String(gasto.id_conceptos);

          return [
            String(gasto.id_conceptos), // Col A: id_conceptos
            String(idSolicitud),        // Col B: id_solicitud
            cantidad,                   // Col C: cantidad
            valor_unit,                 // Col D: valor_unit
            cantidad * valor_unit,      // Col E: valor_total
            conceptoPadre               // Col F: concepto_padre
          ];
        });

      // 6. Insert into GASTOS if there are rows to insert
      if (rows.length === 0) {
        console.log(`No valid expense rows to insert into GASTOS for solicitud ${idSolicitud}`);
        return false; // Indicate nothing was saved if no valid rows
      }

      console.log(`Inserting ${rows.length} rows into GASTOS for solicitud ${idSolicitud}`);
      await this.client.spreadsheets.values.append({
        spreadsheetId: this.spreadsheetId,
        range: 'GASTOS!A2:F', // Append to the end
        valueInputOption: 'USER_ENTERED', // Use USER_ENTERED for GASTOS if formulas/formatting are needed
        insertDataOption: 'INSERT_ROWS',
        resource: { values: rows }
      });

      return true; // Indicate success
    } catch (error) {
      console.error("Error in saveGastos:", error);
      // Log more details if available
      if (error.response && error.response.data) {
        console.error("Google API Error Details:", error.response.data.error);
      }
      throw new Error('Error guardando gastos'); // Re-throw for the controller to catch
    }
  }

  /**
   * Obtiene todos los riesgos asociados a una solicitud
   * @param {string} solicitudId - ID de la solicitud 
   * @returns {Promise<Array>} - Lista de riesgos
   */
  async getRisksBySolicitud(solicitudId) {
    try {
      const response = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: 'RIESGOS!A2:F'
      });
  
      const rows = response.data.values || [];
      const riesgos = rows
        .filter(row => row[4] === solicitudId.toString())
        .map(row => {
          return {
            id_riesgo: row[0],
            nombre_riesgo: row[1],
            aplica: row[2] || 'No',
            mitigacion: row[3] || '',
            id_solicitud: row[4],
            categoria: row[5] || 'General'
          };
        });
  
      return riesgos;
    } catch (error) {
      console.error('Error al obtener riesgos por solicitud:', error);
      throw new Error('Error al obtener riesgos por solicitud');
    }
  }
  
  /**
   * Guarda un nuevo riesgo
   * @param {Object} riskData - Datos del riesgo a guardar
   * @returns {Promise<boolean>} - Resultado de la operación
   */
  async saveRisk(riskData) {
    try {
      await this.client.spreadsheets.values.append({
        spreadsheetId: this.spreadsheetId,
        range: 'RIESGOS!A2:F',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: {
          values: [[
            riskData.id_riesgo,
            riskData.nombre_riesgo,
            riskData.aplica,
            riskData.mitigacion,
            riskData.id_solicitud,
            riskData.categoria
          ]]
        }
      });
      
      return true;
    } catch (error) {
      console.error('Error al guardar riesgo:', error);
      throw new Error('Error al guardar riesgo');
    }
  }
  
  /**
   * Guarda múltiples riesgos de una vez
   * @param {Array<Object>} risksData - Lista de riesgos a guardar
   * @returns {Promise<boolean>} - Resultado de la operación
   */
  async saveBulkRisks(risksData) {
    try {
      const values = risksData.map(risk => [
        risk.id_riesgo,
        risk.nombre_riesgo,
        risk.aplica,
        risk.mitigacion,
        risk.id_solicitud,
        risk.categoria
      ]);
      
      await this.client.spreadsheets.values.append({
        spreadsheetId: this.spreadsheetId,
        range: 'RIESGOS!A2:F',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values }
      });
      
      return true;
    } catch (error) {
      console.error('Error al guardar riesgos en bloque:', error);
      throw new Error('Error al guardar riesgos en bloque');
    }
  }
  
  /**
   * Obtiene un riesgo por su ID
   * @param {string} riskId - ID del riesgo a buscar
   * @returns {Promise<Object|null>} - Datos del riesgo o null si no existe
   */
  async getRiskById(riskId) {
    try {
      const response = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: 'RIESGOS!A2:F'
      });
  
      const rows = response.data.values || [];
      const riskRow = rows.find(row => row[0] === riskId.toString());
      
      if (!riskRow) return null;
      
      return {
        id_riesgo: riskRow[0],
        nombre_riesgo: riskRow[1],
        aplica: riskRow[2] || 'No',
        mitigacion: riskRow[3] || '',
        id_solicitud: riskRow[4],
        categoria: riskRow[5] || 'General'
      };
    } catch (error) {
      console.error(`Error al obtener riesgo con ID ${riskId}:`, error);
      throw new Error(`Error al obtener riesgo con ID ${riskId}`);
    }
  }
  
  /**
   * Actualiza un riesgo existente
   * @param {Object} riskData - Datos actualizados del riesgo
   * @returns {Promise<boolean>} - Resultado de la operación
   */
  async updateRisk(riskData) {
    try {
      // Obtener todos los riesgos para encontrar la fila a actualizar
      const response = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: 'RIESGOS!A2:F'
      });
  
      const rows = response.data.values || [];
      const rowIndex = rows.findIndex(row => row[0] === riskData.id_riesgo.toString());
      
      if (rowIndex === -1) {
        throw new Error(`Riesgo con ID ${riskData.id_riesgo} no encontrado`);
      }
      
      // Calcular la fila en Sheets (índice base 0 + 2 para tener en cuenta el encabezado)
      const sheetRowIndex = rowIndex + 2;
      
      // Actualizar la fila
      await this.client.spreadsheets.values.update({
        spreadsheetId: this.spreadsheetId,
        range: `RIESGOS!A${sheetRowIndex}:F${sheetRowIndex}`,
        valueInputOption: 'RAW',
        resource: {
          values: [[
            riskData.id_riesgo,
            riskData.nombre_riesgo,
            riskData.aplica,
            riskData.mitigacion,
            riskData.id_solicitud,
            riskData.categoria
          ]]
        }
      });
      
      return true;
    } catch (error) {
      console.error('Error al actualizar riesgo:', error);
      throw new Error('Error al actualizar riesgo');
    }
  }
  
  /**
   * Elimina un riesgo por su ID
   * @param {string} riskId - ID del riesgo a eliminar
   * @returns {Promise<boolean>} - Resultado de la operación
   */
  async deleteRisk(riskId) {
    try {
      // Obtener todos los riesgos
      const response = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: 'RIESGOS!A2:F'
      });
  
      const rows = response.data.values || [];
      const newRows = rows.filter(row => row[0] !== riskId.toString());
      
      // Si las filas son iguales, no se encontró el riesgo
      if (rows.length === newRows.length) {
        throw new Error(`Riesgo con ID ${riskId} no encontrado`);
      }
      
      // Limpiar la hoja y volver a escribir todas las filas menos la eliminada
      await this.client.spreadsheets.values.clear({
        spreadsheetId: this.spreadsheetId,
        range: 'RIESGOS!A2:F'
      });
      
      if (newRows.length > 0) {
        await this.client.spreadsheets.values.append({
          spreadsheetId: this.spreadsheetId,
          range: 'RIESGOS!A2',
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS',
          resource: { values: newRows }
        });
      }
      
      return true;
    } catch (error) {
      console.error('Error al eliminar riesgo:', error);
      throw new Error('Error al eliminar riesgo');
    }
  }
  

}

module.exports = new SheetsService();