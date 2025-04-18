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

  async saveGastos(idSolicitud, gastos, actualizarConceptos = true) {
    try {
      console.log(`🔄 Iniciando saveGastos para solicitud ${idSolicitud}. actualizarConceptos=${actualizarConceptos}`);
      console.log(`   Recibidos ${gastos.length} gastos del frontend.`);

      const conceptosValidos = new Set();
      const conceptosSolicitudValidos = new Set();

      // 1. Get existing concepts
      console.log('   Paso 1: Obteniendo conceptos existentes de CONCEPTO$...');
      const conceptosResponse = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: 'CONCEPTO$!A2:F' // Read ID, Desc, Padre, Nombre, Tipo, SolicitudID
      });

      const existingConceptRows = conceptosResponse.data.values || [];
      console.log(`   Encontrados ${existingConceptRows.length} conceptos existentes en total.`);

      existingConceptRows.forEach(row => {
        // Normalizar formato de ID para asegurar consistencia (reemplazar comas y puntos)
        const conceptoId = String(row[0] || '').trim(); // Col A: id_conceptos
        const solicitudIdConcepto = String(row[5] || '').trim(); // Col F: id_solicitud

        // Solo añadir si ambos IDs son válidos para la comparación
        if (conceptoId && solicitudIdConcepto) {
            // Crear clave normalizada para buscar coincidencias independiente del formato
            // Normalizar formato: convertir tanto puntos como comas a un formato común (usamos punto)
            const normalizedId = conceptoId.replace(/,/g, '.');
            const key = `${normalizedId}:${solicitudIdConcepto}`;
            conceptosSolicitudValidos.add(key);
            conceptosValidos.add(normalizedId); // También mantener un set de IDs normalizados
        }
      });
      console.log(`   Total de claves únicas 'concepto:solicitud' cacheadas: ${conceptosSolicitudValidos.size}`);

      // Siempre actualizar conceptos, independientemente del parámetro actualizarConceptos
      console.log('   Paso 2: Identificando NUEVOS conceptos para esta solicitud...');
      
      // 2. Identify ALL concepts (not just new ones) to ensure CONCEPTO$ is updated
      const todoConceptos = gastos
        .filter(gasto => {
          // Asegurarse de que el gasto tiene un ID válido
          const idConceptoGasto = String(gasto.id_conceptos || '').trim();
          if (!idConceptoGasto) {
            console.warn(`      -> Omitiendo gasto sin id_conceptos:`, gasto);
            return false; // Ignorar gastos sin ID
          }
          return true; // Incluir todos los gastos con ID válido
        })
        .map(gasto => {
          // 3. Formatear la fila para CONCEPTO$
          // Normalizar formato para asegurar consistencia (unificar comas y puntos a punto)
          const idConceptoStr = String(gasto.id_conceptos || '').trim();
          const normalizedId = idConceptoStr.replace(/,/g, '.');
          
          // Determinar si es un gasto dinámico (verificando AMBOS formatos: punto y coma)
          const esDinamico = normalizedId.startsWith('15.') || idConceptoStr.startsWith('15,');
          
          // Determinar si es padre (concepto principal sin subíndice)
          const esPadre = typeof gasto.es_padre === 'boolean' 
            ? gasto.es_padre 
            : !(normalizedId.includes('.')); // Si no tiene punto después de normalizar, es un concepto padre
          
          const esPadreStr = esPadre ? 'true' : 'false';
          
          // Usar descripción o nombre de concepto proporcionado, o el ID como último recurso
          const descripcion = gasto.descripcion || gasto.nombre_conceptos || normalizedId;
          
          // Para gastos dinámicos, el nombre_conceptos debe ser igual a descripción
          const nombreConcepto = esDinamico ? descripcion : (gasto.nombre_conceptos || descripcion);
          
          // Tipo debe ser "particular" para gastos extras, "estándar" para el resto
          const tipo = esDinamico ? 'particular' : 'estándar';
          
          console.log(`      * Preparando concepto: ID=${normalizedId}, Tipo=${tipo}, EsPadre=${esPadreStr}, Desc=${descripcion}`);

          return [
            normalizedId,                            // Col A: id_conceptos (normalizado con puntos)
            descripcion,                            // Col B: descripcion
            esPadreStr,                             // Col C: es_padre
            nombreConcepto,                         // Col D: nombre_conceptos
            tipo,                                   // Col E: tipo
            String(idSolicitud)                     // Col F: id_solicitud
          ];
        });

      // Filtrar conceptos que ya existen en CONCEPTO$
      const nuevosConceptos = todoConceptos.filter(conceptoRow => {
        const idConcepto = conceptoRow[0]; // Normalizado con puntos
        const key = `${idConcepto}:${String(idSolicitud)}`;
        return !conceptosSolicitudValidos.has(key);
      });

      console.log(`   Identificados ${nuevosConceptos.length} nuevos conceptos para añadir a CONCEPTO$.`);

      // 4. Guardar nuevos conceptos en CONCEPTO$ (siempre, eliminando la condición redundante)
      if (nuevosConceptos.length > 0) {
        console.log(`   Paso 4: Guardando ${nuevosConceptos.length} nuevos conceptos en CONCEPTO$...`);
        try {
          await this.client.spreadsheets.values.append({
            spreadsheetId: this.spreadsheetId,
            range: 'CONCEPTO$!A2:F', // Append to the end
            valueInputOption: 'RAW',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: nuevosConceptos }
          });
          console.log(`   ✅ ${nuevosConceptos.length} nuevos conceptos añadidos a CONCEPTO$.`);

          // Actualizar el caché local para evitar añadirlos de nuevo inmediatamente
          nuevosConceptos.forEach(conceptoRow => {
            const idConcepto = conceptoRow[0];
            const idSol = conceptoRow[5];
            const key = `${idConcepto}:${idSol}`;
            conceptosValidos.add(idConcepto);
            conceptosSolicitudValidos.add(key);
          });

        } catch (appendError) {
          console.error(`   ❌ ERROR al añadir nuevos conceptos a CONCEPTO$:`, appendError.message);
          throw appendError; // Re-lanzar para asegurar que se detecte el problema
        }
      } else {
         console.log(`   No hay nuevos conceptos para añadir a CONCEPTO$ para la solicitud ${idSolicitud}.`);
      }

      // 5. Prepare rows for GASTOS sheet (always do this for all valid expenses)
      console.log('   Paso 5: Preparando filas para la hoja GASTOS...');
      const rowsGastos = gastos
        .filter(gasto => {
            const cantidad = parseFloat(gasto.cantidad) || 0;
            const valor_unit = parseFloat(gasto.valor_unit) || 0;
            const isValid = cantidad > 0 || valor_unit > 0; // Guardar si hay cantidad o valor unitario
            return isValid;
        })
        .map(gasto => {
          const cantidad = parseFloat(gasto.cantidad) || 0;
          const valor_unit = parseFloat(gasto.valor_unit) || 0;
          const idConceptoStr = String(gasto.id_conceptos || '').trim();
          // Normalizar para consistencia en GASTOS también
          const normalizedId = idConceptoStr.replace(/,/g, '.');
          // Para concepto_padre, extraer la parte antes del punto 
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
      console.log(`   Preparadas ${rowsGastos.length} filas para GASTOS.`);

      // 6. Insert into GASTOS if there are rows to insert
      if (rowsGastos.length > 0) {
        console.log(`   Paso 6: Guardando ${rowsGastos.length} filas en GASTOS...`);
        try {
            console.log(`      -> Buscando y eliminando gastos existentes para solicitud ${idSolicitud} en GASTOS...`);
            const deleteSuccess = await this.deleteGastosBySolicitud(idSolicitud);
            if (deleteSuccess) {
                console.log(`      -> Gastos anteriores eliminados de GASTOS.`);
            } else {
                console.warn(`      -> No se pudieron eliminar gastos anteriores o no existían.`);
            }

            await this.client.spreadsheets.values.append({
                spreadsheetId: this.spreadsheetId,
                range: 'GASTOS!A2:F', // Append to the end
                valueInputOption: 'RAW',
                insertDataOption: 'INSERT_ROWS',
                resource: { values: rowsGastos }
            });
            console.log(`   ✅ ${rowsGastos.length} filas guardadas en GASTOS.`);
            return true; // Indicar éxito general
        } catch (gastosError) {
            console.error(`   ❌ ERROR al guardar filas en GASTOS:`, gastosError.message);
            throw gastosError; // Re-lanzar para que el controlador lo maneje
        }
      } else {
        console.log(`   No hay filas válidas para insertar en GASTOS para la solicitud ${idSolicitud}.`);
        return true;
      }

    } catch (error) {
      console.error(`❌ Error CRÍTICO en saveGastos para solicitud ${idSolicitud}:`, error);
      throw error; // Re-lanzar el error para que el controlador lo capture
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