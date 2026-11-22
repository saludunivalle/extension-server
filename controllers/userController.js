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

      resolvedRole = String(preferredRow[3] || '').trim().toLowerCase() || 'solicitante';
    }

    const normalizedRole = (role || '').toString().trim().toLowerCase() || 'solicitante';

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

/**
 * Obtiene todos los usuarios que tienen rol de director
 */
const getDirectors = async (req, res) => {
  try {
    const directors = await sheetsService.getDirectorsList();
    return res.status(200).json({
      success: true,
      directors,
    });
  } catch (error) {
    console.error('Error al obtener lista de directores:', error);
    return res.status(500).json({
      success: false,
      error: 'Error al obtener la lista de directores',
      details: error.message,
    });
  }
};

/**
 * Agrega un nuevo usuario con rol de director
 */
const addDirector = async (req, res) => {
  try {
    const { email, name } = req.body;
    if (!email || !String(email).trim()) {
      return res.status(400).json({
        success: false,
        error: 'El correo electrónico es obligatorio',
      });
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(String(email).trim())) {
      return res.status(400).json({
        success: false,
        error: 'El formato del correo electrónico no es válido',
      });
    }

    const director = await sheetsService.addDirectorUser({ email, name });
    return res.status(201).json({
      success: true,
      message: 'Director agregado exitosamente',
      director,
    });
  } catch (error) {
    console.error('Error al agregar director:', error);
    return res.status(400).json({
      success: false,
      error: error.message || 'Error al agregar el director',
    });
  }
};

/**
 * Modifica el correo y/o nombre de un director
 */
const updateDirector = async (req, res) => {
  try {
    const { currentEmail, newEmail, name } = req.body;
    if (!currentEmail || !newEmail) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere el correo actual y el nuevo correo',
      });
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(String(newEmail).trim())) {
      return res.status(400).json({
        success: false,
        error: 'El formato del nuevo correo electrónico no es válido',
      });
    }

    const updated = await sheetsService.updateDirectorEmail({
      currentEmail,
      newEmail,
      name,
    });

    return res.status(200).json({
      success: true,
      message: 'Director actualizado exitosamente',
      director: updated,
    });
  } catch (error) {
    console.error('Error al actualizar director:', error);
    return res.status(400).json({
      success: false,
      error: error.message || 'Error al actualizar el director',
    });
  }
};

/**
 * Elimina a un usuario con rol de director
 */
const deleteDirector = async (req, res) => {
  try {
    const email = req.params.email || req.body?.email;
    if (!email) {
      return res.status(400).json({
        success: false,
        error: 'El correo del director a eliminar es obligatorio',
      });
    }

    await sheetsService.deleteDirectorUser(email);
    return res.status(200).json({
      success: true,
      message: `Director con correo ${email} eliminado correctamente`,
    });
  } catch (error) {
    console.error('Error al eliminar director:', error);
    return res.status(400).json({
      success: false,
      error: error.message || 'Error al eliminar el director',
    });
  }
};

module.exports = {
  saveUser,
  getDirectors,
  addDirector,
  updateDirector,
  deleteDirector,
};