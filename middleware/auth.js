const { oAuth2Client } = require('../config/google');
const sheetsService = require('../services/sheetsService');
const { verifySessionToken } = require('../utils/jwtUtils');

/**
 * Middleware para verificar la autenticación mediante token de sesión (JWT) o token de Google.
 * Añade la información del usuario a req.user si la autenticación es exitosa.
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

    let userId = null;
    let userEmail = null;
    let userName = null;
    let tokenRole = null;

    // 1. Intentar verificar con nuestro token JWT de sesión persistente (30 días)
    try {
      const decoded = verifySessionToken(token);
      userId = decoded.id;
      userEmail = decoded.email;
      userName = decoded.name;
      tokenRole = decoded.role;
    } catch (jwtError) {
      // 2. Si no es un token propio, intentar como Google ID Token para compatibilidad
      try {
        const ticket = await oAuth2Client.verifyIdToken({
          idToken: token,
          audience: process.env.GOOGLE_CLIENT_ID
        });
        const payload = ticket.getPayload();
        userId = payload.sub;
        userEmail = payload.email;
        userName = payload.name;
      } catch (googleError) {
        console.warn('Error de autenticación (Token inválido o expirado):', jwtError.message);
        return res.status(401).json({ 
          success: false,
          error: 'Token inválido o expirado' 
        });
      }
    }

    // Añadir la información del usuario a la solicitud
    req.user = {
      id: userId,
      email: userEmail,
      name: userName,
      role: tokenRole || ''
    };

    // Consultar o refrescar el rol desde USUARIOS en Google Sheets
    try {
      const users = await sheetsService.getUsers();
      const userRow = users.find(row => String(row[0] || '').trim() === String(userId).trim());
      if (userRow?.[3]) {
        req.user.role = String(userRow[3]).trim().toLowerCase();
      } else if (!req.user.role) {
        req.user.role = 'solicitante';
      }
    } catch (roleError) {
      console.warn('No se pudo consultar el rol en Sheets; se usará el rol del token:', roleError.message);
      if (!req.user.role) {
        req.user.role = 'solicitante';
      }
    }
    
    next();
  } catch (error) {
    console.error('Error de autenticación general:', error);
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
    if (!req.user) {
      return res.status(401).json({ 
        success: false,
        error: 'Usuario no autenticado' 
      });
    }
    
    const adminEmails = process.env.ADMIN_EMAILS?.split(',') || [];
    
    if (adminEmails.includes(req.user.email) || req.user.role === 'admin') {
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

/**
 * Middleware para verificar si el usuario tiene permiso de Revisor 1 (o admin)
 * Debe usarse después de verifyToken
 */
const isRevisor1 = (req, res, next) => {
  try {
    if (!req.user) {
      return res.status(401).json({
        success: false,
        error: 'Usuario no autenticado'
      });
    }

    const role = String(req.user.role || '').trim().toLowerCase();
    const adminEmails = process.env.ADMIN_EMAILS?.split(',') || [];
    const isUserAdmin = adminEmails.includes(req.user.email) || role === 'admin';
    const isUserRevisor1 = role === 'revisor_1' || role === 'revisor1';

    if (isUserRevisor1 || isUserAdmin) {
      return next();
    }

    return res.status(403).json({
      success: false,
      error: 'Acceso denegado: este módulo está disponible únicamente para usuarios con permiso Revisor 1.'
    });
  } catch (error) {
    console.error('Error de autorización de Revisor 1:', error);
    return res.status(500).json({
      success: false,
      error: 'Error de autorización'
    });
  }
};

module.exports = {
  verifyToken,
  isAdmin,
  isRevisor1
};
