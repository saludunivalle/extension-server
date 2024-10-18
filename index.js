// Importaciones y Configuraciones Iniciales
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const { google } = require('googleapis');
const { config } = require('dotenv');
const multer = require('multer');
const fs = require('fs');
const { jwtClient, oAuth2Client } = require('./google');
config();

const drive = google.drive({ version: 'v3', auth: jwtClient });

// Configuración de multer para almacenar los archivos temporalmente
const upload = multer({ dest: '/tmp/uploads/' });

// Configuración del Servidor
const app = express();
const PORT = process.env.PORT || 3001;

app.use(bodyParser.json());
app.use(cors());

// Middleware de CORS y body-parser
app.use(bodyParser.json());
app.use(cors());

// Ruta Raíz
app.get('/', (req, res) => {
  res.send('El servidor de extensión está funcionando correctamente');
});

// ==========================================================================
// Funciones Utilitarias
// ==========================================================================
const getSpreadsheet = () => google.sheets({ version: 'v4', auth: jwtClient });
const SPREADSHEET_ID = '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';

// ==========================================================================
// Rutas de Autenticación y Usuarios
// ==========================================================================

// Ruta para manejar el inicio de sesión con Google y guardar el usuario en la hoja de Google Sheets
app.post('/auth/google', async (req, res) => {
  try {
    const { token } = req.body;

    const ticket = await oAuth2Client.verifyIdToken({
      idToken: token,
      audience: process.env.CLIENT_ID,
    });

    const payload = ticket.getPayload();
    const userId = payload['sub'];
    const userEmail = payload['email'];
    const userName = payload['name'];

    console.log('Datos del usuario obtenidos del token:', { userId, userEmail, userName });

    const sheets = getSpreadsheet();

    // Comprobar si el usuario ya existe en la hoja
    const userCheckRange = 'USUARIOS!A2:A';
    const userCheckResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: userCheckRange,
    });

    const existingUsers = userCheckResponse.data.values ? userCheckResponse.data.values.flat() : [];
    if (!existingUsers.includes(userId)) {
      const userRange = 'USUARIOS!A2:C2';
      const userValues = [[userId, userEmail, userName]];

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: userRange,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: {
          values: userValues,
        },
      });
      console.log('Usuario guardado correctamente.');
    } else {
      console.log('El usuario ya existe en la hoja de USUARIOS.');
    }

    res.status(200).json({
      success: true,
      userInfo: {
        id: userId,
        email: userEmail,
        name: userName
      }
    });
  } catch (error) {
    console.error('Error al autenticar con Google:', error);
    res.status(500).json({ error: 'Error al autenticar con Google', success: false });
  }
});

// Ruta para guardar el usuario en la hoja de Google Sheets
app.post('/saveUser', async (req, res) => {
  try {
    const { id, email, name } = req.body;
    const sheets = getSpreadsheet();

    // Comprobar si el usuario ya existe en la hoja
    const userCheckRange = 'USUARIOS!A2:A';
    const userCheckResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: userCheckRange,
    });

    const existingUsers = userCheckResponse.data.values ? userCheckResponse.data.values.flat() : [];
    if (!existingUsers.includes(id)) {
      const userRange = 'USUARIOS!A2:C2';
      const userValues = [[id, email, name]];

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: userRange,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: {
          values: userValues,
        },
      });
      console.log('Usuario guardado correctamente.');
    } else {
      console.log('El usuario ya existe en la hoja de USUARIOS.');
    }

    res.status(200).json({ success: true });
  } catch (error) {
    console.error('Error al guardar usuario:', error);
    res.status(500).json({ error: 'Error al guardar usuario', success: false });
  }
});

// ==========================================================================
// Rutas de Formularios y Datos
// ==========================================================================

// Ruta para obtener el último ID de la hoja de Google Sheets
app.get('/getLastId', async (req, res) => {
  try {
    const { sheetName } = req.query;
    const sheets = getSpreadsheet();

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A:A`,
    });

    const lastRow = response.data.values.length;
    const lastId = lastRow > 1 ? response.data.values[lastRow - 1][0] : 0;

    res.status(200).json({ lastId });
  } catch (error) {
    console.error('Error al obtener el último ID:', error);
    res.status(500).json({ error: 'Error al obtener el último ID', success: false });
  }
});

// Ruta para guardar el progreso del formulario
app.post('/saveProgress', async (req, res) => {
  try {
    const { id_usuario, formData, activeStep } = req.body;

    console.log("Recibido del frontend:", formData);
    const sheets = getSpreadsheet();

    let idSolicitud = formData.id_solicitud;

    // Obtener el próximo ID si no existe un ID de solicitud
    if (!idSolicitud) {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ETAPAS!A:A',
      });

      const rows = response.data.values;
      idSolicitud = rows && rows.length > 1 ? rows.length : 1;
      formData.id_solicitud = idSolicitud.toString();
    }

    console.log('Datos procesados para guardar:', { idSolicitud, formData, activeStep });

    const nextRow = parseInt(idSolicitud) + 1;

    if (activeStep === 0) {
      const etapasRange = `ETAPAS!A${nextRow}:F${nextRow}`;
      const etapasValues = [
        [
          idSolicitud,
          id_usuario,
          formData.fecha_solicitud || new Date().toLocaleDateString('es-ES'),
          formData.nombre_solicitante || '',
          'En progreso',
          'En progreso'
        ]
      ];

      const solicitudesRange = `SOLICITUDES!A${nextRow}:L${nextRow}`;  // Asegúrate de que el rango cubra todas las columnas necesarias
      const solicitudesValues = [
        [
          idSolicitud,
          formData.introduccion || '',
          formData.objetivo_general || '',
          formData.objetivos_especificos || '',
          formData.justificacion || '',
          formData.descripcion || '',
          formData.alcance || '',
          formData.metodologia || '',
          formData.dirigido_a || '',
          formData.programa_contenidos || '',
          formData.duracion || '',
          formData.certificacion || '',
          formData.recursos || ''
        ]
      ];

      await Promise.all([
        sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: etapasRange,
          valueInputOption: 'RAW',
          resource: { values: etapasValues },
        }),
        sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: solicitudesRange,
          valueInputOption: 'RAW',
          resource: { values: solicitudesValues },
        })
      ]);
    } else {
      return res.status(400).json({ error: 'Paso inválido', status: false });
    }

    res.status(200).json({ success: 'Progreso guardado correctamente', status: true });
  } catch (error) {
    console.error('Error guardando el progreso:', error);
    res.status(500).json({ error: 'Error guardando el progreso', status: false });
  }
});


// ==========================================================================
// Rutas de Manejo de Archivos
// ==========================================================================

// Ruta para manejar la subida de archivos a Google Drive
app.post('/uploadFile', upload.single('matriz_riesgo'), async (req, res) => {
  try {
    const filePath = req.file.path;
    const fileMetadata = {
      name: req.file.originalname,
      parents: ['12bxb0XEArXMLvc7gX2ndqJVqS_sTiiUE'],
    };

    const media = {
      mimeType: req.file.mimetype,
      body: fs.createReadStream(filePath),
    };

    const file = await drive.files.create({
      resource: fileMetadata,
      media: media,
      fields: 'id, webViewLink',
    });

    fs.unlinkSync(filePath);

    res.status(200).json({ fileUrl: file.data.webViewLink });
  } catch (error) {
    console.error('Error al subir el archivo a Google Drive:', error);
    res.status(500).json({ error: 'Error al subir el archivo a Google Drive' });
  }
});

// Ruta para obtener las solicitudes del usuario
app.get('/getRequests', async (req, res) => {
  try {
    const { userId } = req.query;
    const sheets = getSpreadsheet();

    // Obtener etapas activas desde la hoja ETAPAS
    const activeResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `ETAPAS!A2:F`,
    });

    const rows = activeResponse.data.values;
    if (!rows || rows.length === 0) {
      // Si no se encuentran registros, devolver un mensaje adecuado
      return res.status(404).json({ error: 'No se encontraron solicitudes activas o terminadas' });
    }

    // Filtrar las solicitudes activas y terminadas
    const activeRequests = rows.filter(
      (row) => row[1] === userId && row[5] === 'En progreso'
    );

    const completedRequests = rows.filter(
      (row) => row[1] === userId && row[5] === 'Terminado'
    );

    res.status(200).json({
      activeRequests,
      completedRequests,
    });
  } catch (error) {
    console.error('Error al obtener las etapas:', error);
    res.status(500).json({ error: 'Error al obtener las etapas' });
  }
});

app.get('/getProgramasYOficinas', async (req, res) => {
  try {
    const spreadsheetId = '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: 'PROGRAMAS!A2:K500', 
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron datos en la hoja de Google Sheets' });
    }

    const programas = [];
    const oficinas = new Set(); 

    rows.forEach(row => {
      if (row[4] || row[5] || row[6] || row[7]) { 
        programas.push({
          Programa: row[0],
          Snies: row[1],
          Sede: row[2],
          Facultad: row[3],
          Escuela: row[4],
          Departamento: row[5],
          Sección: row[6] || 'General',
          PregradoPosgrado: row[7],
        });
      }

      if (row[9]) { 
        oficinas.add(row[9]);
      }
    });

    res.status(200).json({
      programas,
      oficinas: Array.from(oficinas), 
    });
  } catch (error) {
    console.error('Error al obtener datos de la hoja de Google Sheets:', error);
    res.status(500).json({ error: 'Error al obtener datos de la hoja de Google Sheets' });
  }
});


// ==========================================================================
// Inicializar el Servidor
// ==========================================================================
app.listen(PORT, () => {
  console.log(`Servidor de extensión escuchando en el puerto ${PORT}`);
});
