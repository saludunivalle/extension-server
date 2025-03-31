const sheetsService = require('../services/sheetsService');

/**
 * Obtiene el progreso de una solicitud
 */
const getProgress = async (req, res) => {
  try {
    const { id_solicitud } = req.params;
    
    if (!id_solicitud) {
      return res.status(400).json({
        success: false,
        error: 'ID de solicitud requerido'
      });
    }
    
    // Obtener progreso desde la sesión si existe
    let progress = req.session.progressState;
    
    // Si no está en la sesión, obtener desde Google Sheets
    if (!progress) {
      const client = sheetsService.getClient();
      const etapasResponse = await client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'ETAPAS!A:I'
      });
      
      const etapasRows = etapasResponse.data.values || [];
      const filaActual = etapasRows.find(row => row[0] === id_solicitud);
      
      if (!filaActual) {
        return res.status(404).json({
          success: false,
          error: `No se encontró la solicitud con ID ${id_solicitud}`
        });
      }
      
      // Extraer datos de la fila
      let estadoFormularios = {};
      if (filaActual[8]) {
        try {
          estadoFormularios = JSON.parse(filaActual[8]);
        } catch (e) {
          estadoFormularios = {
            "1": "En progreso", "2": "En progreso",
            "3": "En progreso", "4": "En progreso"
          };
        }
      } else {
        estadoFormularios = {
          "1": "En progreso", "2": "En progreso",
          "3": "En progreso", "4": "En progreso"
        };
      }
      
      progress = {
        etapa_actual: parseInt(filaActual[4]) || 1,
        paso: parseInt(filaActual[7]) || 1,
        estado: filaActual[5] || 'En progreso',
        estado_formularios: estadoFormularios
      };
    }
    
    res.status(200).json({
      success: true,
      data: progress
    });
  } catch (error) {
    console.error('Error al obtener progreso:', error);
    res.status(500).json({
      success: false,
      error: 'Error al obtener progreso'
    });
  }
};

/**
 * Actualiza el progreso de una solicitud
 */
const updateProgress = async (req, res) => {
  try {
    const { id_solicitud } = req.params;
    const { etapa_actual, paso, estado, estado_formularios } = req.body;
    
    if (!id_solicitud) {
      return res.status(400).json({
        success: false,
        error: 'ID de solicitud requerido'
      });
    }
    
    // Crear el objeto de progreso
    const progressData = {
      etapa_actual: etapa_actual || 1,
      paso: paso || 1,
      estado: estado || 'En progreso',
      estado_formularios: estado_formularios || {
        "1": "En progreso", "2": "En progreso",
        "3": "En progreso", "4": "En progreso"
      }
    };
    
    // Actualizar en la sesión
    req.session.progressState = progressData;
    
    // Actualizar en Google Sheets
    const client = sheetsService.getClient();
    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A:I'
    });
    
    const etapasRows = etapasResponse.data.values || [];
    let filaEtapas = etapasRows.findIndex(row => row[0] === id_solicitud);
    
    if (filaEtapas === -1) {
      return res.status(404).json({
        success: false,
        error: `No se encontró la solicitud con ID ${id_solicitud}`
      });
    }
    
    filaEtapas += 1; // Ajustar índice a 1-based para Google Sheets
    
    // Actualizar la hoja
    await client.spreadsheets.values.update({
      spreadsheetId: sheetsService.spreadsheetId,
      range: `ETAPAS!E${filaEtapas}:I${filaEtapas}`,
      valueInputOption: 'RAW',
      resource: {
        values: [[
          progressData.etapa_actual,
          progressData.estado,
          etapasRows[filaEtapas - 1][6] || 'N/A',
          progressData.paso,
          JSON.stringify(progressData.estado_formularios)
        ]]
      }
    });
    
    res.status(200).json({
      success: true,
      message: 'Progreso actualizado correctamente',
      data: progressData
    });
  } catch (error) {
    console.error('Error al actualizar progreso:', error);
    res.status(500).json({
      success: false,
      error: 'Error al actualizar progreso'
    });
  }
};

module.exports = {
  getProgress,
  updateProgress
};