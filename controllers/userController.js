const sheetsService = require('../services/sheetsService');

/**
 * Guarda un usuario en Google Sheets si no existe previamente
 * @param {Object} req - Objeto de solicitud Express 
 * @param {Object} res - Objeto de respuesta Express
*/

const saveUser = async (req, res) => {
  try {
    console.log('Datos recibidos en saveUser:', req.body);
    const { id, email, name } = req.body;
    
    // Validación de datos
    if (!id || !email || !name) {
      return res.status(400).json({ 
        success: false,
        error: 'Se requieren los campos id, email y name' 
      });
    }

    // Usar el servicio de sheets
    const client = sheetsService.getClient();
    
    // Verificar si el usuario ya existe
    const userCheckRange = 'USUARIOS!A2:A';
    const userCheckResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: userCheckRange,
    });

    const existingUsers = userCheckResponse.data.values ? userCheckResponse.data.values.flat() : [];
    
    // Si el usuario no existe, guardarlo
    if (!existingUsers.includes(id)) {
      const userRange = 'USUARIOS!A2:C2';
      const userValues = [[id, email, name]];

      await client.spreadsheets.values.append({
        spreadsheetId: sheetsService.spreadsheetId,
        range: userRange,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: userValues },
      });
    }

    res.status(200).json({ 
      success: true,
      message: 'Usuario guardado correctamente',
      userInfo: { id, email, name }
    });
  } catch (error) {
    console.error('Error al guardar el usuario:', error);
    res.status(500).json({ 
      success: false,
      error: 'Error al guardar el usuario',
      details: error.message
    });
  }
};

module.exports = {
  saveUser
};