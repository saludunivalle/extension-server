const sheetsService = require('../services/sheetsService');

/**
 * Guarda un usuario en Google Sheets si no existe previamente
 * @param {Object} req - Objeto de solicitud Express 
 * @param {Object} res - Objeto de respuesta Express
*/

const saveUser = async (req, res) => {
  try {
    console.log('Datos recibidos en saveUser:', req.body);
    const { id, email, name, role } = req.body;
    
    // Validación de datos
    if (!id || !email || !name) {
      return res.status(400).json({ 
        success: false,
        error: 'Se requieren los campos id, email y name' 
      });
    }

    // Usar el servicio de sheets
    const client = sheetsService.getClient();
    
    // Verificar si el usuario ya existe y recuperar su rol actual
    const userCheckRange = 'USUARIOS!A2:D';
    const userCheckResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: userCheckRange,
    });

    const rows = userCheckResponse.data.values || [];
    const normalizedId = String(id).trim();
    const matchingRows = rows.filter((row) => String(row[0] || '').trim() === normalizedId);

    const existingUsers = rows.map((row) => String(row[0] || '').trim());
    console.log('Usuarios existentes en la hoja:', existingUsers);

    let resolvedRole = '';

    if (matchingRows.length > 0) {
      const preferredRow = matchingRows.find(
        (row) => String(row[3] || '').trim().toLowerCase() === 'admin'
      ) || matchingRows[0];

      resolvedRole = String(preferredRow[3] || '').trim().toLowerCase() === 'admin' ? 'admin' : '';
    }

    const normalizedRole = (role || '').toString().trim().toLowerCase() === 'admin' ? 'admin' : '';

    // Si el usuario no existe, guardarlo
    if (!existingUsers.includes(normalizedId)) {
      const userRange = 'USUARIOS!A2:D2';
      const userValues = [[id, email, name, normalizedRole]];

      await client.spreadsheets.values.append({
        spreadsheetId: sheetsService.spreadsheetId,
        range: userRange,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: userValues },
      });

      resolvedRole = normalizedRole;
    }

    res.status(200).json({ 
      success: true,
      message: 'Usuario guardado correctamente',
      userInfo: { id, email, name, role: resolvedRole }
    });
    console.log(`Usuario autenticado: ${email} (ID: ${id}, Rol en hoja: ${resolvedRole || 'sin rol'})`);
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