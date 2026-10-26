require('dotenv').config();

const http = require('http');
const crypto = require('crypto');
const { google } = require('googleapis');

const port = Number(process.env.DRIVE_OAUTH_PORT || 3002);
const redirectUri = process.env.GOOGLE_DRIVE_REDIRECT_URI ||
  process.env.DRIVE_OAUTH_REDIRECT_URI ||
  `http://localhost:${port}/oauth2callback`;
const clientId = process.env.GOOGLE_CLIENT_ID || process.env.client_id;
const clientSecret = process.env.GOOGLE_CLIENT_SECRET || process.env.client_secret;
const scopes = ['https://www.googleapis.com/auth/drive'];
const state = crypto.randomBytes(24).toString('hex');

if (!clientId || !clientSecret) {
  console.error('Faltan GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET (o client_id/client_secret) en .env');
  process.exitCode = 1;
} else {
  const oauthClient = new google.auth.OAuth2(clientId, clientSecret, redirectUri);
  const authorizationUrl = oauthClient.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: scopes,
    state
  });

  const server = http.createServer(async (request, response) => {
    const requestUrl = new URL(request.url, redirectUri);
    if (requestUrl.pathname !== new URL(redirectUri).pathname) {
      response.writeHead(404);
      response.end('Ruta no encontrada');
      return;
    }

    if (requestUrl.searchParams.get('state') !== state) {
      response.writeHead(400);
      response.end('Estado OAuth invalido');
      server.close();
      process.exitCode = 1;
      return;
    }

    const error = requestUrl.searchParams.get('error');
    if (error) {
      response.writeHead(400);
      response.end(`Autorizacion cancelada: ${error}`);
      server.close();
      process.exitCode = 1;
      return;
    }

    try {
      const code = requestUrl.searchParams.get('code');
      if (!code) throw new Error('Google no devolvio un codigo de autorizacion');
      const { tokens } = await oauthClient.getToken(code);
      if (!tokens.refresh_token) {
        throw new Error('Google no devolvio refresh_token. Revoca el acceso de la app y vuelve a ejecutar el script.');
      }

      response.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
      response.end('Autorizacion completada. Puedes cerrar esta ventana.');
      console.log('\nAutorizacion completada.');
      console.log('Agrega esta variable al .env del backend (no la compartas ni la subas al repositorio):');
      console.log(`GOOGLE_DRIVE_REFRESH_TOKEN=${tokens.refresh_token}`);
      console.log('\nReinicia el backend despues de guardarla.');
      server.close();
    } catch (exchangeError) {
      response.writeHead(500);
      response.end('No fue posible completar la autorizacion');
      console.error('\nError al intercambiar el codigo OAuth:', exchangeError.message);
      server.close();
      process.exitCode = 1;
    }
  });

  server.on('error', error => {
    console.error(`No se pudo iniciar el callback local en ${redirectUri}:`, error.message);
    process.exitCode = 1;
  });

  server.listen(port, '127.0.0.1', () => {
    console.log('Abre esta ruta en el navegador para autorizar la cuenta institucional:\n');
    console.log(authorizationUrl);
    console.log(`\nEsperando el callback en ${redirectUri}`);
    console.log('Asegurate de que esta URI este registrada en el OAuth Client de Google.');
  });
}
