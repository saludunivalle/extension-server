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
          range: 'SOLICITUDES!A2:AV',
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

  async saveGastos(idSolicitud, gastos, actualizarConceptos = true) {
    try {
      console.log(`🔄 Iniciando saveGastos para solicitud ${idSolicitud}. actualizarConceptos=${actualizarConceptos}`);
      console.log(`   Recibidos ${gastos.length} gastos del frontend.`);

      const conceptosExistentes = new Map(); // Cambiado a Map para almacenar fila e índice
      const conceptosSolicitudValidos = new Set();

      // 1. Obtener conceptos existentes
      console.log('   Paso 1: Obteniendo conceptos existentes de CONCEPTO$...');
      const conceptosResponse = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: 'CONCEPTO$!A2:F' // Obtener ID, Desc, Padre, Nombre, Tipo, SolicitudID
      });

      const existingConceptRows = conceptosResponse.data.values || [];
      console.log(`   Encontrados ${existingConceptRows.length} conceptos existentes en total.`);

      // Mapear los conceptos existentes con su índice de fila para actualización
      existingConceptRows.forEach((row, index) => {
        const conceptoId = String(row[0] || '').trim(); // Col A: id_conceptos
        const solicitudIdConcepto = String(row[5] || '').trim(); // Col F: id_solicitud

        if (conceptoId && solicitudIdConcepto) {
          // Normalizar formato a punto
          const normalizedId = conceptoId.replace(/,/g, '.');
          const key = `${normalizedId}:${solicitudIdConcepto}`;
          
          // Almacenar el índice de fila (base 0, ajustar a 1-based para API sheets)
          conceptosExistentes.set(key, {
            rowIndex: index + 2, // +2 porque índice es 0-based y hay que considerar encabezado
            data: row
          });
          
          conceptosSolicitudValidos.add(key);
        }
      });
      
      console.log(`   Total de conceptos existentes mapeados: ${conceptosExistentes.size}`);

      // 2. Preparar todos los conceptos para guardar o actualizar
      console.log('   Paso 2: Preparando conceptos para guardado/actualización...');

      // Solicitudes de actualización para conceptos existentes
      const updateRequests = [];
      // Conceptos nuevos para añadir
      const nuevosConceptos = [];

      // Procesar cada gasto para decidir si actualizar o insertar
      gastos.forEach(gasto => {
        // Normalizar ID del concepto
        const idConceptoStr = String(gasto.id_conceptos || '').trim();
        if (!idConceptoStr) {
          console.warn(`      -> Omitiendo gasto sin id_conceptos:`, gasto);
          return; // Continuar con el siguiente gasto
        }

        const normalizedId = idConceptoStr.replace(/,/g, '.');
        const esDinamico = normalizedId.startsWith('8.') || idConceptoStr.startsWith('8,');
        const esPadre = typeof gasto.es_padre === 'boolean' 
          ? gasto.es_padre 
          : !(normalizedId.includes('.')); // Si no tiene punto después de normalizar, es un concepto padre
        
        const esPadreStr = esPadre ? 'true' : 'false';
        const descripcion = gasto.descripcion || gasto.nombre_conceptos || normalizedId;
        const nombreConcepto = esDinamico ? descripcion : (gasto.nombre_conceptos || descripcion);
        const tipo = esDinamico ? 'particular' : 'estándar';
        
        // Datos del concepto formateados para sheets
        const conceptoData = [
          normalizedId,                            // Col A: id_conceptos (normalizado con puntos)
          descripcion,                            // Col B: descripcion
          esPadreStr,                             // Col C: es_padre
          nombreConcepto,                         // Col D: nombre_conceptos
          tipo,                                   // Col E: tipo
          String(idSolicitud)                     // Col F: id_solicitud
        ];

        // Verificar si este concepto ya existe
        const key = `${normalizedId}:${String(idSolicitud)}`;
        
        if (conceptosExistentes.has(key)) {
          // El concepto ya existe - ACTUALIZAR
          const existingInfo = conceptosExistentes.get(key);
          const rowIndex = existingInfo.rowIndex;
          
          console.log(`      * ACTUALIZANDO concepto existente: ID=${normalizedId}, Fila=${rowIndex}`);
          
          // Añadir request para actualizar toda la fila
          updateRequests.push({
            range: `CONCEPTO$!A${rowIndex}:F${rowIndex}`,
            values: [conceptoData]
          });
        } else {
          // Concepto nuevo - INSERTAR
          console.log(`      * NUEVO concepto: ID=${normalizedId}, Tipo=${tipo}, EsPadre=${esPadreStr}`);
          nuevosConceptos.push(conceptoData);
        }
      });

      // 3. Actualizar conceptos existentes (si hay)
      if (updateRequests.length > 0) {
        console.log(`   Paso 3: Actualizando ${updateRequests.length} conceptos existentes...`);
        try {
          await this.client.spreadsheets.values.batchUpdate({
            spreadsheetId: this.spreadsheetId,
            resource: {
              valueInputOption: 'RAW',
              data: updateRequests
            }
          });
          console.log(`   ✅ ${updateRequests.length} conceptos actualizados en CONCEPTO$.`);
        } catch (updateError) {
          console.error(`   ❌ ERROR al actualizar conceptos en CONCEPTO$:`, updateError.message);
          // No lanzamos error para intentar continuar con nuevos conceptos
        }
      }

      // 4. Añadir nuevos conceptos (si hay)
      if (nuevosConceptos.length > 0) {
        console.log(`   Paso 4: Añadiendo ${nuevosConceptos.length} nuevos conceptos...`);
        try {
          await this.client.spreadsheets.values.append({
            spreadsheetId: this.spreadsheetId,
            range: 'CONCEPTO$!A2:F',
            valueInputOption: 'RAW',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: nuevosConceptos }
          });
          console.log(`   ✅ ${nuevosConceptos.length} nuevos conceptos añadidos a CONCEPTO$.`);
        } catch (appendError) {
          console.error(`   ❌ ERROR al añadir nuevos conceptos a CONCEPTO$:`, appendError.message);
          // No lanzamos error para intentar continuar con gastos
        }
      }

      // 5. Manejar la hoja GASTOS
      // Primero, eliminamos los gastos existentes para esta solicitud
      console.log(`   Paso 5: Eliminando gastos existentes para solicitud ${idSolicitud}...`);
      await this.deleteGastosBySolicitud(idSolicitud);

      // 6. Preparar e insertar nuevos gastos
      console.log('   Paso 6: Preparando filas para la hoja GASTOS...');
      const rowsGastos = gastos
        .filter(gasto => {
          const cantidad = parseFloat(gasto.cantidad) || 0;
          const valor_unit = parseFloat(gasto.valor_unit) || 0;
          const isValid = cantidad > 0 || valor_unit > 0;
          return isValid;
        })
        .map(gasto => {
          const cantidad = parseFloat(gasto.cantidad) || 0;
          const valor_unit = parseFloat(gasto.valor_unit) || 0;
          const idConceptoStr = String(gasto.id_conceptos || '').trim();
          const normalizedId = idConceptoStr.replace(/,/g, '.');
          const conceptoPadre = normalizedId.includes('.') ?
            normalizedId.split('.')[0] : normalizedId;

          return [
            normalizedId,                        // Col A: id_conceptos (normalizado)
            String(idSolicitud),                // Col B: id_solicitud
            cantidad,                           // Col C: cantidad
            valor_unit,                         // Col D: valor_unit
            cantidad * valor_unit,              // Col E: valor_total (Calculado siempre)
            conceptoPadre                       // Col F: concepto_padre
          ];
        });

      if (rowsGastos.length > 0) {
        console.log(`   Paso 7: Guardando ${rowsGastos.length} filas en GASTOS...`);
        try {
          await this.client.spreadsheets.values.append({
            spreadsheetId: this.spreadsheetId,
            range: 'GASTOS!A2:F',
            valueInputOption: 'RAW',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: rowsGastos }
          });
          console.log(`   ✅ ${rowsGastos.length} filas guardadas en GASTOS.`);
        } catch (gastosError) {
          console.error(`   ❌ ERROR al guardar filas en GASTOS:`, gastosError.message);
          throw gastosError;
        }
      } else {
        console.log(`   No hay filas válidas para insertar en GASTOS.`);
      }

      return true;
    } catch (error) {
      console.error(`❌ Error CRÍTICO en saveGastos para solicitud ${idSolicitud}:`, error);
      throw error;
    }
  }

  /**
   * Elimina todas las filas de la hoja GASTOS que coincidan con un id_solicitud.
   * @param {string} idSolicitud - El ID de la solicitud cuyos gastos se eliminarán.
   * @returns {Promise<boolean>} - True si la operación fue exitosa o no había nada que borrar, false si hubo un error.
   */
  async deleteGastosBySolicitud(idSolicitud) {
    try {
        const sheetName = 'GASTOS';
        const range = `${sheetName}!A:F`;

        const response = await this.client.spreadsheets.values.get({
            spreadsheetId: this.spreadsheetId,
            range: range,
            majorDimension: 'ROWS',
        });

        const rows = response.data.values || [];
        if (rows.length < 2) {
            console.log(`      -> No hay datos en ${sheetName} para eliminar.`);
            return true;
        }

        const deleteRequests = [];
        for (let i = rows.length - 1; i >= 1; i--) {
            const row = rows[i];
            if (row && row[1] === String(idSolicitud)) {
                deleteRequests.push({
                    deleteDimension: {
                        range: {
                            sheetId: await this.getSheetIdByName(sheetName),
                            dimension: 'ROWS',
                            startIndex: i,
                            endIndex: i + 1
                        }
                    }
                });
            }
        }

        if (deleteRequests.length > 0) {
            console.log(`      -> Preparando para eliminar ${deleteRequests.length} filas de ${sheetName} para solicitud ${idSolicitud}.`);
            await this.client.spreadsheets.batchUpdate({
                spreadsheetId: this.spreadsheetId,
                resource: { requests: deleteRequests }
            });
            console.log(`      -> ${deleteRequests.length} filas eliminadas exitosamente.`);
        } else {
            console.log(`      -> No se encontraron filas para eliminar en ${sheetName} para solicitud ${idSolicitud}.`);
        }

        return true;
    } catch (error) {
        console.error(`      -> Error al eliminar gastos por solicitud ${idSolicitud}:`, error.message);
        return false;
    }
  }

  /**
   * Obtiene el ID numérico de una hoja por su nombre. Cachea el resultado.
   * @param {string} sheetName - Nombre de la hoja.
   * @returns {Promise<number|null>} - ID numérico de la hoja o null si no se encuentra.
   */
  async getSheetIdByName(sheetName) {
      if (!this.sheetIdsCache) {
          this.sheetIdsCache = new Map();
      }
      if (this.sheetIdsCache.has(sheetName)) {
          return this.sheetIdsCache.get(sheetName);
      }

      try {
          const response = await this.client.spreadsheets.get({
              spreadsheetId: this.spreadsheetId,
              fields: 'sheets(properties(sheetId,title))'
          });
          const sheets = response.data.sheets || [];
          for (const sheet of sheets) {
              if (sheet.properties && sheet.properties.title === sheetName) {
                  const sheetId = sheet.properties.sheetId;
                  this.sheetIdsCache.set(sheetName, sheetId);
                  return sheetId;
              }
          }
          console.error(`Sheet ID not found for name: ${sheetName}`);
          return null;
      } catch (error) {
          console.error(`Error getting sheet ID for name ${sheetName}:`, error.message);
          return null;
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