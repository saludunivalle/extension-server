const { google } = require('googleapis');
const { config } = require('dotenv');
config();

/**
 * Configuración para las APIs de Google (Sheets y Drive)
 * Proporciona clientes autenticados para las diferentes servicios
 */

// Ámbitos de permisos requeridos
const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets', 
  'https://www.googleapis.com/auth/drive'
];

// Obtener credenciales del archivo .env
const {
  client_id,
  client_secret,
  client_email,
  private_key
} = process.env;

/**
 * Cliente OAuth2 para autenticación de usuarios
 */
const oAuth2Client = new google.auth.OAuth2(
  client_id,
  client_secret,
  process.env.REDIRECT_URI || undefined
);

/**
 * Cliente JWT para autenticación de servicio
 */
const jwtClient = new google.auth.JWT(
  process.env.GOOGLE_CLIENT_EMAIL,
  null,
  (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
  SCOPES
);

// Iniciar la autenticación de servicio
jwtClient.authorize((err, tokens) => {
  if (err) {
    console.error('Error al autorizar JWT:', err);
    return;
  }
  console.log('Conexión exitosa usando JWT!');
});

/**
 * Obtiene un token de acceso mediante el código de autorización
 * @param {string} code - Código de autorización 
 * @returns {Promise<Object>} Tokens de acceso
 */

const getAccessToken = async (code) => {
  try {
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    console.log('Tokens obtenidos:', tokens);
    return tokens;
  } catch (error) {
    console.error('Error al obtener token de acceso:', error);
    throw error;
  }
};

module.exports = {
  oAuth2Client,
  jwtClient,
  getAccessToken
};