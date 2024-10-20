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
const app = express();
const upload = multer({ dest: '/tmp/uploads/' });
const PORT = process.env.PORT || 3001;

app.use(bodyParser.json());
app.use(cors());

// Función para conectarse a Google Sheets
const getSpreadsheet = () => google.sheets({ version: 'v4', auth: jwtClient });
const SPREADSHEET_ID = '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';

// ==========================================================================
// Rutas de Autenticación y Usuarios
// ==========================================================================

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

    const sheets = getSpreadsheet();
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
        resource: { values: userValues },
      });
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

app.post('/saveUser', async (req, res) => {
  try {
    const { id, email, name } = req.body;
    const sheets = getSpreadsheet();
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
        resource: { values: userValues },
      });
    }

    res.status(200).json({ success: true });
  } catch (error) {
    console.error('Error al guardar usuario:', error);
    res.status(500).json({ error: 'Error al guardar usuario', success: false });
  }
});

// ==========================================================================
// Rutas de Formularios y Guardado de Progreso
// ==========================================================================

app.post('/guardarProgreso', async (req, res) => {
  const { id_solicitud, formData, paso, hoja, userData } = req.body;

  try {
    const sheets = getSpreadsheet();
    let sheetName = '';
    let columnas = [];

    // Identificar la hoja y las columnas según el formulario
    switch (hoja) {
        case 1:
          sheetName = 'SOLICITUDES';
          columnas = {
            1: ['B', 'C', 'D'], // Columnas para los datos del paso 1
            2: ['E', 'F'], // Columnas para los datos del paso 2
            3: ['G', 'H', 'I'], // Columnas para los datos del paso 3
            4: ['J', 'K'],      // Columnas para los datos del paso 4
            5: ['L', 'M'],      // Columnas para los datos del paso 5
          };
          break;
          case 2:
            sheetName = 'SOLICITUDES2';
            columnas = {
              1: ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'], // Columnas para los datos del paso 1
              2: ['J', 'K', 'L'], // Columnas para los datos del paso 2
              3: ['M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T'], // Columnas para los datos del paso 3
              4: ['U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC'],      // Columnas para los datos del paso 4
              5: ['AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL'],      // Columnas para los datos del paso 5
            };
            break;
          case 3:
            sheetName = 'SOLICITUDES3';
            columnas = {
              1: ['B', 'C'], // Columnas para los datos del paso 1 (B a C)
              2: [
                'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 
                'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 
                'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 
                'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 
                'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 
                'CE', 'CF', 'CG', 'CH', 'CI'
              ], // Columnas para los datos del paso 2 (D a CI)
              3: ['CJ', 'CK', 'CL'], // Columnas para los datos del paso 3 (CJ a CL)
            };
            break;

            case 4:
              sheetName = 'SOLICITUDES4';
              columnas = {
                1: ['B', 'C'], // Paso 1 va de B a C
                2: ['D', 'E', 'F', 'G', 'H', 'I'], // Paso 2 va de D a I
                3: ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X'], // Paso 3 va de J a X
                4: ['Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO'], // Paso 4 va de Y a AO
                5: ['AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK'], // Paso 5 va de AP a BK
              };            
            break;
            case 5:
              sheetName = 'SOLICITUDES5';
              columnas = {
                1: ['B', 'C', 'D', 'E', 'F'], // Paso 1 va de B a F
                2: ['G', 'H', 'I', 'J'], // Paso 2 va de G a J
                3: ['K', 'L', 'M', 'N', 'O'], // Paso 3 va de K a O
                4: ['P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'], // Paso 4 va de P a Z
                5: ['AA', 'AB', 'AC'] // Paso 5 va de AA a AC
              };            
            break;
      default:
        return res.status(400).json({ error: 'Hoja no válida' });
    }

    // Buscar el id_solicitud en la primera columna (columna A) en la hoja correspondiente
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A:A`,
    });

    const rows = response.data.values || [];
    let fila = rows.findIndex((row) => row[0] === id_solicitud.toString());

    // Si no existe, agrega una nueva fila
    if (fila === -1) {
      fila = rows.length + 1;
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!A${fila}`,
        valueInputOption: 'RAW',
        resource: { values: [[id_solicitud]] }, // Inserta el id_solicitud en la primera columna
      });
    } else {
      fila += 1; // Ajustar el índice de la fila para trabajar con la API (basada en 1)
    }


    // Actualizar el progreso del formulario
    const columnasPaso = columnas[paso];
    const valores = Object.values(formData);

    // Agrega los logs para verificar las columnas y los valores enviados
    console.log('Columnas esperadas:', columnasPaso.length);
    console.log('Valores enviados:', valores.length);
    console.log('Valores enviados:', valores);

    if (valores.length !== columnasPaso.length) {
      return res.status(400).json({ error: 'Número de columnas no coincide con los valores' });
    }

    const columnaInicial = columnasPaso[0]; // Columna inicial
    const columnaFinal = columnasPaso[columnasPaso.length - 1]; // Columna final

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!${columnaInicial}${fila}:${columnaFinal}${fila}`,
      valueInputOption: 'RAW',
      resource: { values: [valores] }, // Actualiza los valores del paso
    });

    // Ahora actualizamos la hoja de ETAPAS
    const etapaActual = `Paso ${paso}`;
    const estado = paso === 5 ? 'Completado' : 'En progreso';
    const fechaActual = new Date().toLocaleDateString();

    // Buscar el id_solicitud en la hoja ETAPAS
    const etapasResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `ETAPAS!A:A`,
    });

    const etapasRows = etapasResponse.data.values || [];
    let filaEtapas = etapasRows.findIndex((row) => row[0] === id_solicitud.toString());

    if (filaEtapas === -1) {
      // Si no existe un registro en ETAPAS para esta solicitud, agregarlo
      filaEtapas = etapasRows.length + 1;
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `ETAPAS!A${filaEtapas}`,
        valueInputOption: 'RAW',
        resource: {
          values: [[id_solicitud, userData.id_usuario, fechaActual, userData.name, etapaActual, estado, formData.dependencia || '']]
        },
      });
    } else {
      filaEtapas += 1; // Ajustar el índice de la fila para trabajar con la API (basada en 1)

      // Actualizar el registro en ETAPAS
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `ETAPAS!B${filaEtapas}:G${filaEtapas}`,
        valueInputOption: 'RAW',
        resource: {
          values: [[userData.id_usuario, fechaActual, userData.name, etapaActual, estado, formData.dependencia || '']],
        },
      });
    }

    res.status(200).json({ success: true });
  } catch (error) {
    console.error('Error al guardar el progreso:', error);
    res.status(500).json({ error: 'Error al guardar el progreso' });
  }
});



// ==========================================================================
// Otras rutas auxiliares
// ==========================================================================




// Ruta para obtener el último ID
app.get('/getLastId', async (req, res) => {
  const { sheetName } = req.query;
  try {
    const sheets = getSpreadsheet();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A:A`, // Columna A contiene el id_solicitud
    });
    const lastRow = response.data.values ? response.data.values.length : 0;
    const lastId = lastRow > 1 ? parseInt(response.data.values[lastRow - 1][0], 10) : 0;
    res.status(200).json({ lastId });
  } catch (error) {
    console.error('Error al obtener el último ID:', error);
    res.status(500).json({ error: 'Error al obtener el último ID' });
  }
});


app.get('/getRequests', async (req, res) => {
  try {
    const { userId } = req.query;
    const sheets = getSpreadsheet();

    const activeResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `ETAPAS!A2:F`,
    });

    const rows = activeResponse.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron solicitudes activas o terminadas' });
    }

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
// Inicializar el servidor
// ==========================================================================
app.listen(PORT, () => {
  console.log(`Servidor de extensión escuchando en el puerto ${PORT}`);
});
