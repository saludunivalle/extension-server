const sheetsService = require('./sheetsService');
const path = require('path');
const fs = require('fs');

class ProgressStateService {
  constructor() {
    // Sin inicialización de Redis
    console.log('Initialized ProgressStateService with session storage only');
  }

  async getProgress(solicitudId, session) {
    // Si hay datos en la sesión, usarlos
    if (session?.progressState && session.progressState.id_solicitud === solicitudId) {
      return session.progressState;
    }
    
    // Si no hay datos en la sesión, obtener de Google Sheets
    return this.loadFromSheets(solicitudId);
  }

  async setProgress(solicitudId, progressData, session) {
    try {
      // Guardar en la sesión
      if (session) {
        session.progressState = {
          ...progressData,
          id_solicitud: solicitudId
        };
      }
      
      // Guardar en Google Sheets
      await this.saveToSheets(solicitudId, progressData);
      return true;
    } catch (error) {
      console.error(`Error setting progress for ${solicitudId}:`, error);
      return false;
    }
  }

  async saveToSheets(solicitudId, progressData) {
    try {
      const { etapa_actual, paso, estado, estado_formularios } = progressData;

      // Obtener los datos actuales de ETAPAS
      const client = sheetsService.getClient();
      const etapasResponse = await client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'ETAPAS!A:I'
      });
      const etapasRows = etapasResponse.data.values || [];

      // Buscar la fila que corresponde al id_solicitud
      let filaEtapas = etapasRows.findIndex(row => row[0] === solicitudId.toString());

      if (filaEtapas === -1) {
        console.log(`No se encontró la solicitud con ID ${solicitudId} en ETAPAS.`);
        return false;
      }

      filaEtapas += 1; // Ajustar índice a 1-based para Google Sheets

      // Actualizar la fila en Google Sheets
      await client.spreadsheets.values.update({
        spreadsheetId: sheetsService.spreadsheetId,
        range: `ETAPAS!E${filaEtapas}:I${filaEtapas}`,
        valueInputOption: 'RAW',
        resource: {
          values: [[
            etapa_actual,
            estado,
            etapasRows[filaEtapas - 1][6] || 'N/A', // Mantener nombre_actividad
            paso,
            JSON.stringify(estado_formularios)
          ]]
        }
      });

      console.log(`✅ Progreso guardado en Google Sheets para ${solicitudId}`);
      return true;

    } catch (error) {
      console.error(`Error saving progress to Google Sheets for ${solicitudId}:`, error);
      // Lógica de reintento con retroceso exponencial
      if (retryCount < 3) {
        const delay = Math.pow(2, retryCount) * 1000; // 1s, 2s, 4s
        console.log(`Retrying in ${delay}ms... (attempt ${retryCount + 1})`);
        await new Promise(resolve => setTimeout(resolve, delay));
        return this.saveToSheets(solicitudId, progressData, retryCount + 1);
      } else {
        console.error('Max retries reached, failing.');
        throw error;
      }
    }
  }

  async loadFromSheets(solicitudId) {
    try {
      // Obtener los datos actuales de ETAPAS
      const client = sheetsService.getClient();
      const etapasResponse = await client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'ETAPAS!A:I'
      });
      const etapasRows = etapasResponse.data.values || [];

      // Buscar la fila que corresponde al id_solicitud
      const filaEtapas = etapasRows.find(row => row[0] === solicitudId.toString());

      if (!filaEtapas) {
        console.log(`No se encontró la solicitud con ID ${solicitudId} en ETAPAS.`);
        return { // Valores por defecto
          etapa_actual: 1,
          paso: 1,
          estado: 'En progreso',
          estado_formularios: {
            "1": "En progreso", "2": "En progreso",
            "3": "En progreso", "4": "En progreso"
          },
          version: 0
        };
      }

      const etapa_actual = parseInt(filaEtapas[4]) || 1;
      const estado = filaEtapas[5] || 'En progreso';
      const paso = parseInt(filaEtapas[7]) || 1;
      const estado_formularios = filaEtapas[8] ? JSON.parse(filaEtapas[8]) : {
        "1": "En progreso", "2": "En progreso",
        "3": "En progreso", "4": "En progreso"
      };
      const version = parseInt(filaEtapas[9]) || 0;

      return {
        etapa_actual,
        paso,
        estado,
        estado_formularios,
        version
      };

    } catch (error) {
      console.error(`Error loading progress from Google Sheets for ${solicitudId}:`, error);
      return { // Valores por defecto
        etapa_actual: 1,
        paso: 1,
        estado: 'En progreso',
        estado_formularios: {
          "1": "En progreso", "2": "En progreso",
          "3": "En progreso", "4": "En progreso"
        },
        version: 0
      };
    }
  }
}

module.exports = new ProgressStateService();