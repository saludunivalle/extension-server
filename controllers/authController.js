const { oAuth2Client } = require('../config/google');
const sheetsService = require('../services/sheetsService');

/**
 * Autentica al usuario mediante Google OAuth2
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */

const googleAuth = async (req, res) => {
  try {
    const { token } = req.body;
    const ticket = await oAuth2Client.verifyIdToken({
      idToken: token,
      audience: process.env.GOOGLE_CLIENT_ID,
    });

    const payload = ticket.getPayload();
    const userId = payload['sub'];
    const userEmail = payload['email'];
    const userName = payload['name'];

    // Usar sheetsService para verificar si el usuario existe
    const client = sheetsService.getClient();
    const userCheckRange = 'USUARIOS!A2:D';
    const userCheckResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: userCheckRange,
    });


    let userRole = '';
    let userExists = false;

    if (userCheckResponse.data.values) {
      const normalizedUserId = String(userId).trim();
      const matchingRows = userCheckResponse.data.values.filter(
        (row) => String(row[0] || '').trim() === normalizedUserId
      );

      if (matchingRows.length > 0) {
        userExists = true;

        // Si hay duplicados, priorizar el registro con rol admin.
        const preferredRow = matchingRows.find(
          (row) => String(row[3] || '').trim().toLowerCase() === 'admin'
        ) || matchingRows[0];

        userRole = String(preferredRow[3] || '').trim().toLowerCase() === 'admin' ? 'admin' : '';
      }
    }

    
    if (!userExists) {
      userRole = '';
      const userRange = 'USUARIOS!A2:D2';
      const userValues = [[userId, userEmail, userName, userRole]]; // Agrega un rol predeterminado si no se proporciona
      console.log('Usuario no encontrado, guardando nuevo usuario:', { userId, userEmail, userName, userRole });
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
      userInfo: {
        id: userId,
        email: userEmail,
        name: userName,
        role: userRole || '' // Devuelve el rol del usuario si está disponible
      }
    });
    console.log(`Usuario autenticado: ${userEmail} (ID: ${userId}, Rol: ${userRole || 'sin rol'})`);
  } catch (error) {
    console.error('Error al autenticar con Google:', error);
    res.status(500).json({ error: 'Error al autenticar con Google', success: false });
  }
};

module.exports = {
  googleAuth
};