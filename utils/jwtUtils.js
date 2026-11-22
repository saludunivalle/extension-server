const jwt = require('jsonwebtoken');

const JWT_SECRET = process.env.JWT_SECRET || process.env.SESSION_SECRET || 'extension_jwt_secret_default_key';
const JWT_EXPIRES_IN = process.env.JWT_EXPIRES_IN || '30d';

/**
 * Genera un token JWT de sesión para el usuario.
 * @param {Object} payload Datos del usuario (id, email, name, role)
 * @param {string} expiresIn Duración del token (por defecto '30d')
 * @returns {string} Token JWT
 */
const generateSessionToken = (payload, expiresIn = JWT_EXPIRES_IN) => {
  return jwt.sign(
    {
      id: payload.id,
      email: payload.email,
      name: payload.name,
      role: payload.role || 'solicitante'
    },
    JWT_SECRET,
    { expiresIn }
  );
};

/**
 * Verifica un token JWT de sesión.
 * @param {string} token
 * @returns {Object} Payload decodificado
 */
const verifySessionToken = (token) => {
  return jwt.verify(token, JWT_SECRET);
};

module.exports = {
  generateSessionToken,
  verifySessionToken,
  JWT_SECRET,
  JWT_EXPIRES_IN
};
