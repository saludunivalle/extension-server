const { google } = require('googleapis');
const { config } = require('dotenv');
config();

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'];

const {
  client_id,
  client_secret,
  client_email,
  private_key
} = process.env;

const oAuth2Client = new google.auth.OAuth2(
  client_id,
  client_secret,
);

const jwtClient = new google.auth.JWT(
  client_email,
  null,
  private_key.replace(/\\n/g, '\n'),
  SCOPES
);

jwtClient.authorize((err, tokens) => {
  if (err) {
    console.error('Error authorizing JWT:', err);
    return;
  }
  console.log('Successfully connected using JWT!');
});

const getAccessToken = async (code) => {
  try {
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    console.log('Tokens acquired:', tokens);
    return tokens;
  } catch (error) {
    console.error('Error getting access token:', error);
    throw error;
  }
};

module.exports = {
  oAuth2Client,
  jwtClient,
  getAccessToken
};
