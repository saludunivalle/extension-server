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

      return rowIndex;
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
          range: 'GASTOS!A2:E',
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

  async saveGastos(idSolicitud, gastos) {
    try {
      // Validar conceptos
      const conceptosResponse = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: 'CONCEPTO$!A2:A'
      });
  
      const conceptosValidos = new Set(
        (conceptosResponse.data.values || []).flat().map(String)
      );
  
      // Preparar filas válidas
      const rows = gastos
        .filter(gasto => gasto.id_conceptos && conceptosValidos.has(String(gasto.id_conceptos)))
        .map(gasto => {
          const cantidad = parseFloat(gasto.cantidad) || 0;
          const valor_unit = parseFloat(gasto.valor_unit) || 0;
          return [
            gasto.id_conceptos.toString(),
            idSolicitud.toString(),
            cantidad,
            valor_unit,
            cantidad * valor_unit
          ];
        });
  
      if (rows.length === 0) return false;
  
      // Insertar en GASTOS
      await this.client.spreadsheets.values.append({
        spreadsheetId: this.spreadsheetId,
        range: 'GASTOS!A2:E',
        valueInputOption: 'USER_ENTERED',
        resource: { values: rows }
      });
  
      return true;
    } catch (error) {
      console.error("Error en saveGastos:", error);
      throw new Error('Error guardando gastos');
    }
  }

}

module.exports = new SheetsService();