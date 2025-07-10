const { oAuth2Client } = require('../config/google');

/**
 * Middleware para verificar la autenticación mediante token de Google
 * Añade la información del usuario a req.user si la autenticación es exitosa
*/

const verifyToken = async (req, res, next) => {
  try {
    // Obtener el token de los headers
    const authHeader = req.headers.authorization;
    
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ 
        success: false,
        error: 'No se proporcionó token de autenticación' 
      });
    }

    const token = authHeader.split(' ')[1];

    // Verificar el token con Google
    const ticket = await oAuth2Client.verifyIdToken({
      idToken: token,
      audience: process.env.GOOGLE_CLIENT_ID
    });

    const payload = ticket.getPayload();
    
    // Añadir la información del usuario a la solicitud
    req.user = {
      id: payload.sub,
      email: payload.email,
      name: payload.name
    };
    
    next();
  } catch (error) {
    console.error('Error de autenticación:', error);
    res.status(401).json({ 
      success: false,
      error: 'Token inválido o expirado'
    });
  }
};

/**
 * Middleware para verificar si el usuario es administrador
 * Debe usarse después de verifyToken
*/

const isAdmin = (req, res, next) => {
  try {
    // Asegurar que verifyToken fue ejecutado primero
    if (!req.user) {
      return res.status(401).json({ 
        success: false,
        error: 'Usuario no autenticado' 
      });
    }
    
    // Implementar lógica para verificar si el usuario es administrador
    // Por ejemplo, verificar contra una lista de correos administrativos
    const adminEmails = process.env.ADMIN_EMAILS?.split(',') || [];
    
    if (adminEmails.includes(req.user.email)) {
      next();
    } else {
      res.status(403).json({
        success: false,
        error: 'No tienes permisos para acceder a este recurso'
      });
    }
  } catch (error) {
    console.error('Error de autorización:', error);
    res.status(500).json({ 
      success: false,
      error: 'Error de autorización'
    });
  }
};

module.exports = {
  verifyToken,
  isAdmin
};