const { google } = require('googleapis');
const { jwtClient } = require('../config/google');
const sheetsService = require('../services/sheetsService');
const reportService = require('../services/reportService');

/**
 * Obtiene el último ID de una hoja específica
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
*/

const getLastId = async (req, res) => {
  const { sheetName } = req.query;
  try {
    // Usar el servicio de sheets en lugar de la función global
    const lastId = await sheetsService.getLastId(sheetName);
    res.status(200).json({ lastId });
  } catch (error) {
    console.error(`Error al obtener el último ID de ${sheetName}:`, error);
    res.status(500).json({ error: `Error al obtener el último ID de ${sheetName}` });
  }
};

/**
 * Obtiene datos de programas y oficinas desde la hoja de cálculo
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
*/

const getProgramasYOficinas = async (req, res) => {
  try {
    const spreadsheetId = sheetsService.spreadsheetId;
    const client = sheetsService.getClient();

    const response = await client.spreadsheets.values.get({
      spreadsheetId,
      range: 'PROGRAMAS!A2:K500',
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron datos en la hoja de Google Sheets' });
    }

    const programas = [];
    const oficinas = new Set();

    rows.forEach(row => {
      if (row[4] || row[5] || row[6] || row[7]) {
        programas.push({
          Programa: row[0],
          Snies: row[1],
          Sede: row[2],
          Facultad: row[3],
          Escuela: row[4],
          Departamento: row[5],
          Sección: row[6] || 'General',
          PregradoPosgrado: row[7],
        });
      }

      if (row[9]) {
        oficinas.add(row[9]);
      }
    });

    res.status(200).json({
      programas,
      oficinas: Array.from(oficinas),
    });
  } catch (error) {
    console.error('Error al obtener datos de la hoja de Google Sheets:', error);
    res.status(500).json({ error: 'Error al obtener datos de la hoja de Google Sheets' });
  }
};

/**
 * Obtiene datos de una solicitud específica
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
*/

const getSolicitud = async (req, res) => {
  try {
    const { id_solicitud } = req.query;
    
    if (!id_solicitud) {
      return res.status(400).json({ error: 'El ID de la solicitud es requerido' });
    }

    // Usar el servicio de hojas
    const hojas = sheetsService.reportSheetDefinitions || reportService.reportSheetDefinitions;
      
    const resultados = await sheetsService.getSolicitudData(id_solicitud, hojas);

    res.status(200).json(resultados);
  } catch (error) {
    console.error('Error al obtener los datos de la solicitud:', error);
    res.status(500).json({ error: 'Error al obtener los datos de la solicitud' });
  }
};

module.exports = {
  getLastId,
  getProgramasYOficinas,
  getSolicitud
};