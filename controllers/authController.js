const { oAuth2Client } = require('../config/google');
const sheetsService = require('../services/sheetsService');
const { generateSessionToken } = require('../utils/jwtUtils');

/**
 * Autentica al usuario mediante Google OAuth2 y genera un token de sesión de larga duración (30 días).
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
      const normalizedUserEmail = String(userEmail || '').trim().toLowerCase();
      const allRows = userCheckResponse.data.values;

      // 1. Buscar coincidencias exactas por ID de usuario
      const matchingById = allRows
        .map((row, idx) => ({ row, rowIndex: idx + 2 }))
        .filter((item) => String(item.row[0] || '').trim() === normalizedUserId);

      if (matchingById.length > 0) {
        userExists = true;
        const preferred = matchingById.find(
          (item) => String(item.row[3] || '').trim().toLowerCase() === 'admin'
        ) || matchingById[0];

        userRole = String(preferred.row[3] || '').trim().toLowerCase();
      } else {
        // 2. Si no coincide por ID, buscar por correo electrónico (p.ej. pre-registrado como director)
        const matchingByEmail = allRows
          .map((row, idx) => ({ row, rowIndex: idx + 2 }))
          .filter(
            (item) => String(item.row[1] || '').trim().toLowerCase() === normalizedUserEmail
          );

        if (matchingByEmail.length > 0) {
          userExists = true;
          const matched = matchingByEmail[0];
          userRole = String(matched.row[3] || '').trim().toLowerCase();

          // Vincular el Google ID y actualizar nombre en la hoja
          try {
            await client.spreadsheets.values.update({
              spreadsheetId: sheetsService.spreadsheetId,
              range: `USUARIOS!A${matched.rowIndex}:C${matched.rowIndex}`,
              valueInputOption: 'RAW',
              resource: {
                values: [[userId, userEmail, userName || matched.row[2] || '']],
              },
            });
            sheetsService.invalidateUsersCache();
            console.log(`Director/Usuario vinculado por correo: ${userEmail} -> Google ID ${userId}`);
          } catch (syncErr) {
            console.warn('No se pudo actualizar el ID de Google en USUARIOS:', syncErr.message);
          }
        }
      }
    }

    if (!userExists) {
      userRole = 'solicitante';
      const userRange = 'USUARIOS!A2:D2';
      const userValues = [[userId, userEmail, userName, userRole]];
      console.log('Usuario no encontrado, guardando nuevo usuario:', { userId, userEmail, userName, userRole });
      await client.spreadsheets.values.append({
        spreadsheetId: sheetsService.spreadsheetId,
        range: userRange,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: userValues },
      });
    }

    const userInfo = {
      id: userId,
      email: userEmail,
      name: userName,
      role: userRole || 'solicitante'
    };

    // Generar token de sesión persistente (30 días)
    const sessionToken = generateSessionToken(userInfo);

    res.status(200).json({
      success: true,
      token: sessionToken,
      userInfo
    });
    console.log(`Usuario autenticado con sesión de 30 días: ${userEmail} (ID: ${userId}, Rol: ${userInfo.role})`);
  } catch (error) {
    console.error('Error al autenticar con Google:', error);
    res.status(500).json({ error: 'Error al autenticar con Google', success: false });
  }
};

/**
 * Valida la sesión actual del usuario usando el token provisto.
 * @param {Object} req - Objeto de solicitud Express (con req.user inyectado por verifyToken)
 * @param {Object} res - Objeto de respuesta Express
 */
const verifySession = async (req, res) => {
  try {
    return res.status(200).json({
      success: true,
      userInfo: req.user
    });
  } catch (error) {
    console.error('Error al verificar sesión:', error);
    return res.status(500).json({ error: 'Error al verificar sesión', success: false });
  }
};

module.exports = {
  googleAuth,
  verifySession
};
