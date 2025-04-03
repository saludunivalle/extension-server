const { google } = require('googleapis');
const { jwtClient } = require('../config/google')
const models = require('../models/spreadsheetModels');
const { GoogleAPIError } = require('../middleware/errorHandler');
const { withExponentialBackoff } = require('../utils/apiUtils');
const { getDataWithCache } = require('../utils/cacheUtils');

/**
 * Servicio para manejar operaciones con Google Sheets
*/

class SheetsService {
  constructor() {
    this.spreadsheetId = process.env.SPREADSHEET_ID || '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';
    this.client = this.getClient();
    this.models = models.getModels(); // Usar los modelos definidos
    this.cache = new Map();
    this.cacheTTL = 600000; // Aumentado a 10 minutos para reducir peticiones
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
    const cacheKey = `${sheetName}_${idSolicitud}`;
    
    // Verificar caché primero
    if (this.cache.has(cacheKey)) {
      const {value, timestamp} = this.cache.get(cacheKey);
      if (Date.now() - timestamp < this.cacheTTL) {
        return value;
      }
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
      // Al final, guardar en caché:
      this.cache.set(cacheKey, {
        value: result,
        timestamp: Date.now()
      });
      
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
  async getSolicitudData(solicitudId, definicionesHojas) {
    try {
      // Usar caché para esta operación frecuente
      const cacheKey = `solicitud_${solicitudId}`;
      
      // Verificar caché local primero
      if (this.cache.has(cacheKey)) {
        const {value, timestamp} = this.cache.get(cacheKey);
        if (Date.now() - timestamp < this.cacheTTL) {
          console.log(`✅ Usando caché local para solicitud ${solicitudId}`);
          return value;
        }
      }
      
      return await getDataWithCache(cacheKey, async () => {
        const hojas = definicionesHojas || {
          SOLICITUDES: {
            range: 'SOLICITUDES!A2:AU',
            fields: this.fieldDefinitions.SOLICITUDES
          },
          SOLICITUDES2: {
            range: 'SOLICITUDES2!A2:CL',
            fields: this.fieldDefinitions.SOLICITUDES2
          },
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

        // Preparar rangos para batchGet - consolidar múltiples llamadas en una
        const ranges = Object.values(hojas).map(h => h.range);
        
        try {
          // Usar batchGetValues para obtener todos los datos de una vez
          const valueRanges = await withExponentialBackoff(() => 
            this.batchGetValues(ranges), 2);
          
          const resultados = {};
          let solicitudEncontrada = false;
          
          // Procesar resultados
          valueRanges.forEach((valueRange, index) => {
            const hojaName = Object.keys(hojas)[index];
            const rows = valueRange.values || [];
            const fields = hojas[hojaName].fields;
            
            const solicitudData = rows.find(row => row[0] && row[0].toString() === solicitudId.toString());
            
            if (solicitudData) {
              solicitudEncontrada = true;
              resultados[hojaName] = fields.reduce((acc, field, i) => {
                acc[field] = solicitudData[i] || '';
                return acc;
              }, {});
            }
          });
          
          if (!solicitudEncontrada) {
            return { message: 'La solicitud no existe aún en Google Sheets', data: {} };
          }
          
          // Guardar en caché local también
          this.cache.set(cacheKey, {
            value: resultados,
            timestamp: Date.now()
          });
          
          return resultados;
        } catch (error) {
          // Si hay error de cuota, intentar usar caché expirado
          if (this.isQuotaError(error)) {
            console.warn(`⚠️ Error de cuota al obtener datos de solicitud ${solicitudId}`);
            
            // Intentar usar caché expirado si existe
            if (this.cache.has(cacheKey)) {
              console.log('Usando datos expirados de caché como respaldo');
              return this.cache.get(cacheKey).value;
            }
          }
          throw error;
        }
      }, 5); // TTL de 5 minutos (aumentado de 2)
    } catch (error) {
      console.error('Error al obtener datos de solicitud:', error);
      throw error;
    }
  }

  async saveGastos(idSolicitud, gastos) {
    try {
      // Clave de caché para conceptos
      const conceptosCacheKey = 'conceptos_all';
      let conceptosRows;
      
      // Intentar obtener conceptos de la caché primero
      if (this.cache.has(conceptosCacheKey)) {
        const {value, timestamp} = this.cache.get(conceptosCacheKey);
        if (Date.now() - timestamp < this.cacheTTL) {
          conceptosRows = value;
          console.log('✅ Usando conceptos desde caché local');
        }
      }
      
      // Si no están en caché, obtenerlos de la API
      if (!conceptosRows) {
        try {
          const conceptosResponse = await this.client.spreadsheets.values.get({
            spreadsheetId: this.spreadsheetId,
            range: 'CONCEPTO$!A2:F'
          });
          
          conceptosRows = conceptosResponse.data.values || [];
          
          // Guardar en caché para futuras solicitudes
          this.cache.set(conceptosCacheKey, {
            value: conceptosRows,
            timestamp: Date.now()
          });
        } catch (error) {
          // En caso de error de cuota, continuar con un arreglo vacío
          if (this.isQuotaError(error)) {
            console.warn('⚠️ Error de cuota al obtener conceptos, continuando con valores por defecto');
            conceptosRows = [];
          } else {
            throw error;
          }
        }
      }
      
      // Resto del código de saveGastos...
      const conceptosValidos = new Set();
      const conceptosSolicitudValidos = new Set();
      const conceptosPadre = new Map();
      
      // Procesar conceptos existentes
      (conceptosRows || []).forEach(row => {
        const concepto = String(row[0]);
        const descripcion = row[1] || '';
        const esPadre = row[2] === 'true' || row[2] === 'TRUE';
        const idSolicitudConcepto = row[5] || '';
        
        conceptosValidos.add(concepto);
        conceptosSolicitudValidos.add(`${concepto}:${idSolicitudConcepto}`);
        conceptosPadre.set(concepto, {descripcion, esPadre});
        
        // Si es concepto padre, agregar subconceptos
        if (esPadre) {
          for (let i = 1; i <= 10; i++) {
            conceptosValidos.add(`${concepto}.${i}`);
          }
        }
      });

      // 3. Identificar TODOS los nuevos conceptos (tanto regulares como sub)
      const nuevosConceptos = gastos
        .filter(gasto => {
          // Verificar si ya existe, sin filtrar por tipo
          return !conceptosSolicitudValidos.has(`${gasto.id_conceptos}:${idSolicitud}`);
        })
        .map(gasto => {
          const esSubconcepto = gasto.id_conceptos.includes('.');
          const conceptoPadre = esSubconcepto ? gasto.id_conceptos.split('.')[0] : '';
          
          return [
            gasto.id_conceptos.toString(),
            gasto.concepto || (esSubconcepto ? `Subconcepto de ${conceptoPadre}` : gasto.id_conceptos),
            esSubconcepto ? 'false' : 'true',
            conceptoPadre,  // nombre_conceptos 
            "gasto_dinamico", // tipo
            idSolicitud  // id_solicitud
          ];
        });
  
      // 4. Guardar nuevos conceptos en CONCEPTO$
      if (nuevosConceptos.length > 0) {
        await this.client.spreadsheets.values.append({
          spreadsheetId: this.spreadsheetId,
          range: 'CONCEPTO$!A2:F',
          valueInputOption: 'RAW',
          resource: { values: nuevosConceptos }
        });
  
        // 5. Actualizar conceptosValidos con los nuevos
        nuevosConceptos.forEach(concepto => {
          const idConcepto = concepto[0];
          conceptosValidos.add(idConcepto);
          conceptosSolicitudValidos.add(`${idConcepto}:${idSolicitud}`);
        });
      }
  
      // Preparar filas válidas
      const rows = gastos
        .filter(gasto => gasto.id_conceptos && conceptosValidos.has(String(gasto.id_conceptos)))
        .map(gasto => {
          const cantidad = parseFloat(gasto.cantidad) || 0;
          const valor_unit = parseFloat(gasto.valor_unit) || 0;
          const conceptoPadre = gasto.id_conceptos.includes('.') ? 
            gasto.id_conceptos.split('.')[0] : gasto.id_conceptos;
          
          return [
            gasto.id_conceptos.toString(),
            idSolicitud.toString(),
            cantidad,
            valor_unit,
            cantidad * valor_unit,
            conceptoPadre 
          ];
        })
  
      if (rows.length === 0) return false;
  
      // Insertar en GASTOS
      await this.client.spreadsheets.values.append({
        spreadsheetId: this.spreadsheetId,
        range: 'GASTOS!A2:F',
        valueInputOption: 'USER_ENTERED',
        resource: { values: rows }
      });
  
      return true;
    } catch (error) {
      // Añadir manejo específico para errores de cuota aquí...
      console.error("Error en saveGastos:", error);
      throw new Error('Error guardando gastos');
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

  /**
   * NUEVA FUNCIÓN: Obtiene múltiples rangos de datos en una sola llamada API
   * @param {Array<string>} ranges - Array de rangos a obtener (ej. ['SOLICITUDES!A2:Z', 'GASTOS!A2:F']) 
   * @returns {Promise<Array>} - Datos de todos los rangos solicitados
   */
  async batchGetValues(ranges) {
    try {
      // Verificar si hay demasiados rangos (máximo recomendado: 20)
      if (ranges.length > 20) {
        console.warn(`⚠️ batchGetValues recibió ${ranges.length} rangos, el máximo recomendado es 20. Particionando...`);
        
        const results = [];
        // Procesar en grupos de 20 rangos
        for (let i = 0; i < ranges.length; i += 20) {
          const batchRanges = ranges.slice(i, i + 20);
          const batchResults = await this.batchGetValues(batchRanges);
          results.push(...batchResults);
        }
        return results;
      }

      console.log(`🔄 Solicitando ${ranges.length} rangos en una sola llamada API`);
      
      const response = await this.client.spreadsheets.values.batchGet({
        spreadsheetId: this.spreadsheetId,
        ranges: ranges,
        majorDimension: 'ROWS',
        valueRenderOption: 'UNFORMATTED_VALUE'
      });
      
      return response.data.valueRanges || [];
    } catch (error) {
      if (this.isQuotaError(error)) {
        console.error('🚨 Error de cuota en batchGetValues:', error.message);
        throw new Error('Cuota API excedida: intente más tarde');
      }
      
      console.error('Error en batchGetValues:', error);
      throw new Error('Error al obtener valores en lote');
    }
  }

  /**
   * NUEVA FUNCIÓN: Verifica si un error es por exceso de cuota
   * @param {Error} error - El error a verificar
   * @returns {boolean} - True si es error de cuota
   */
  isQuotaError(error) {
    return error.code === 429 || 
           (error.response && error.response.status === 429) ||
           error.message?.includes('Quota exceeded');
  }

}

module.exports = new SheetsService();