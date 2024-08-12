const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const { google } = require('googleapis');
const { config } = require('dotenv');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { jwtClient, oAuth2Client } = require('./google');
config();

const drive = google.drive({ version: 'v3', auth: jwtClient });

// Configuración de multer para almacenar los archivos temporalmente
const upload = multer({ dest: '/tmp/uploads/' });

const app = express();
const PORT = process.env.PORT || 3001;

app.use(bodyParser.json());
app.use(cors());

// Ruta para la raíz
app.get('/', (req, res) => {
  res.send('El servidor de extensión está funcionando correctamente');
});

// Ruta para manejar el inicio de sesión y guardar el usuario en la hoja de Google Sheets
app.post('/auth/google', async (req, res) => {
  try {
    const { token } = req.body;

    // Decodificar el token JWT para obtener la información del usuario
    const ticket = await oAuth2Client.verifyIdToken({
      idToken: token,
      audience: process.env.CLIENT_ID,
    });

    const payload = ticket.getPayload();
    const userId = payload['sub'];
    const userEmail = payload['email'];
    const userName = payload['name'];

    console.log('Datos del usuario obtenidos del token:', { userId, userEmail, userName });

    const spreadsheetId = '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    // Comprobar si el usuario ya existe en la hoja
    const userCheckRange = 'USUARIOS!A2:A'; 
    const userCheckResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: userCheckRange,
    });

    const existingUsers = userCheckResponse.data.values ? userCheckResponse.data.values.flat() : [];
    
    console.log('Usuarios existentes:', existingUsers);

    if (!existingUsers.includes(userId)) {
      // Guardar la información del usuario si no existe
      const userRange = 'USUARIOS!A2:C2'; // Ajusta para agregar en la siguiente fila vacía
      const userValues = [[userId, userEmail, userName]];
      const userRequest = {
        spreadsheetId,
        range: userRange,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS', // Añadir la fila en lugar de sobrescribir
        resource: {
          values: userValues,
        },
      };

      const appendResponse = await sheets.spreadsheets.values.append(userRequest);
      console.log('Respuesta de la inserción en USUARIOS:', appendResponse);
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

// Ruta para obtener el último ID de la hoja de Google Sheets
app.get('/getLastId', async (req, res) => {
  try {
    const { sheetName } = req.query;
    const spreadsheetId = '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:A`, // Solo la columna A que contiene los IDs
    });

    const lastRow = response.data.values.length;
    const lastId = lastRow > 1 ? response.data.values[lastRow - 1][0] : 0; // Se asegura que no tome la fila de los encabezados

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
    const spreadsheetId = '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    
    // Formato dd/mm/yyyy
    const formatFechaSolicitud = (fecha) => {
      const [year, month, day] = fecha.split('-');
      return `${day}/${month}/${year}`;
    };

    const formattedDate = formatFechaSolicitud(formData.fecha_solicitud);

    console.log('Datos recibidos:', { id_usuario, formData, activeStep }); 

    let range;
    let values = [];

    switch (activeStep) {
      case 0:
        // Guardar en SOLICITUDES y RESPUESTAS (primeras columnas)
        // SOLICITUDES
        const solicitudesRange = `SOLICITUDES!A${parseInt(formData.id_solicitud) + 1}:F${parseInt(formData.id_solicitud) + 1}`;
        const solicitudesValues = [
          [
            formData.id_solicitud,
            id_usuario,
            formattedDate, // fecha_solicitud formateada
            formData.nombre_dependencia, 
            formData.nombre_solicitante,
            'En progreso' // estado
          ]
        ];

        // RESPUESTAS - Primeras columnas
        const respuestasRange = `RESPUESTAS!A${parseInt(formData.id_solicitud) + 1}:E${parseInt(formData.id_solicitud) + 1}`;
        const respuestasValues = [
          [
            formData.id_solicitud,
            formattedDate, // fecha_solicitud formateada
            formData.nombre_actividad,
            formData.nombre_solicitante,
            formData.nombre_dependencia
          ]
        ];

        // Ejecutar ambas actualizaciones en paralelo
        await Promise.all([
          sheets.spreadsheets.values.update({
            spreadsheetId,
            range: solicitudesRange,
            valueInputOption: 'RAW',
            resource: {
              values: solicitudesValues,
            },
          }),
          sheets.spreadsheets.values.update({
            spreadsheetId,
            range: respuestasRange,
            valueInputOption: 'RAW',
            resource: {
              values: respuestasValues,
            },
          })
        ]);
        break;

      case 1:
        // Guardar en RESPUESTAS - Columnas para el paso 1
        range = `RESPUESTAS!F${parseInt(formData.id_solicitud) + 1}:J${parseInt(formData.id_solicitud) + 1}`;
        values = [
          [
            formData.tipo,
            formData.modalidad,
            formData.ofrecido_por,
            formData.unidad_academica,
            formData.ofrecido_para
          ]
        ];
        break;

      case 2:
        // Guardar en RESPUESTAS - Columnas para el paso 2
        range = `RESPUESTAS!K${parseInt(formData.id_solicitud) + 1}:P${parseInt(formData.id_solicitud) + 1}`;
        values = [
          [
            formData.intensidad_horaria,
            formData.horas_modalidad,
            formData.horas_trabj_ind,
            formData.creditos,
            formData.cupo_min,
            formData.cupo_max
          ]
        ];
        break;

      case 3:
        // Guardar en RESPUESTAS - Columnas para el paso 3
        range = `RESPUESTAS!Q${parseInt(formData.id_solicitud) + 1}:X${parseInt(formData.id_solicitud) + 1}`;
        values = [
          [
            formData.nombre_coordinador,
            formData.tel_coordinador,
            formData.profesor_participante,
            formData.formas_evaluacion,
            formData.certificado_solicitado,
            formData.calificacion_minima,
            formData.razon_no_certificado,
            formData.valor_inscripcion
          ]
        ];
        break;

      case 4:
        // Guardar en RESPUESTAS - Columnas para el paso 4
        range = `RESPUESTAS!Y${parseInt(formData.id_solicitud) + 1}:AI${parseInt(formData.id_solicitud) + 1}`;
        values = [
          [
            formData.becas_convenio,
            formData.becas_estudiantes,
            formData.becas_docentes,
            formData.becas_otros,
            formData.becas_total,
            formData.fechas_actividad,
            formData.organizacion_actividad,
            formData.nombre_firma,
            formData.cargo_firma,
            formData.firma,
            formData.anexo_documento
          ]
        ];

        // Actualizar el estado a "Terminado" cuando se complete el último paso
        const updateRange = `SOLICITUDES!F${parseInt(formData.id_solicitud) + 1}`;  // Columna F para etapa_actual
        const updateValues = [['Terminado']];

        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: updateRange,
          valueInputOption: 'RAW',
          resource: {
            values: updateValues,
          },
        });
        break;

      default:
        return res.status(400).json({ error: 'Paso inválido', status: false });
    }

    if (range && values.length > 0) {
      const request = {
        spreadsheetId,
        range,
        valueInputOption: 'RAW',
        resource: {
          values,
        },
      };
      await sheets.spreadsheets.values.update(request);
    }

    res.status(200).json({ success: 'Progreso guardado correctamente', status: true });
  } catch (error) {
    console.error('Error guardando el progreso:', error);
    res.status(500).json({ error: 'Error guardando el progreso', status: false });
  }
});

// Ruta para guardar el usuario en la hoja de Google Sheets
app.post('/saveUser', async (req, res) => {
    try {
      const { id, email, name } = req.body;
  
      const spreadsheetId = '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';
      const sheets = google.sheets({ version: 'v4', auth: jwtClient });
  
      // Comprobar si el usuario ya existe en la hoja
      const userCheckRange = 'USUARIOS!A2:A'; // Comprueba solo la columna de ID de usuario
      const userCheckResponse = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: userCheckRange,
      });
  
      const existingUsers = userCheckResponse.data.values ? userCheckResponse.data.values.flat() : [];
      
      if (!existingUsers.includes(id)) {
        // Guardar la información del usuario si no existe
        const userRange = 'USUARIOS!A2:C2'; // Ajusta para agregar en la siguiente fila vacía
        const userValues = [[id, email, name]];
        const userRequest = {
          spreadsheetId,
          range: userRange,
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS', // Añadir la fila en lugar de sobrescribir
          resource: {
            values: userValues,
          },
        };
  
        const appendResponse = await sheets.spreadsheets.values.append(userRequest);
        console.log('Respuesta de la inserción en USUARIOS:', appendResponse);
      } else {
        console.log('El usuario ya existe en la hoja de USUARIOS.');
      }
  
      res.status(200).json({
        success: true,
      });
  
    } catch (error) {
      console.error('Error al guardar usuario:', error);
      res.status(500).json({ error: 'Error al guardar usuario', success: false });
    }
  }); 

// Ruta para manejar la subida de archivos a Google Drive
app.post('/uploadFile', upload.single('anexo_documento'), async (req, res) => {
  try {
      const filePath = req.file.path;
      const fileMetadata = {
          name: req.file.originalname,
          parents: ['12bxb0XEArXMLvc7gX2ndqJVqS_sTiiUE'], // ID de la carpeta donde deseas guardar los archivos
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

      // Elimina el archivo temporal después de subirlo
      try {
          fs.unlinkSync(filePath);
      } catch (error) {
          console.error('Error al eliminar el archivo temporal:', error);
      }

      // Devuelve el enlace para visualizar el archivo
      res.status(200).json({ fileUrl: file.data.webViewLink });
  } catch (error) {
      console.error('Error al subir el archivo a Google Drive:', error);
      res.status(500).json({ error: 'Error al subir el archivo a Google Drive' });
  }
});

app.get('/getRequests', async (req, res) => {
  try {
    const { userId } = req.query;
    const spreadsheetId = '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    // Obtener solicitudes activas desde la hoja SOLICITUDES
    const activeResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `SOLICITUDES!A2:F`,
    });

    const activeRequests = activeResponse.data.values.filter(
      (row) => row[1] === userId && row[5] === 'En progreso'
    );

    // Obtener solicitudes terminadas desde la hoja SOLICITUDES
    const completedResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `SOLICITUDES!A2:F`,
    });

    const completedRequests = completedResponse.data.values.filter(
      (row) => row[1] === userId && row[5] === 'Terminado'
    );

    res.status(200).json({
      activeRequests,
      completedRequests,
    });
  } catch (error) {
    console.error('Error al obtener las solicitudes:', error);
    res.status(500).json({ error: 'Error al obtener las solicitudes' });
  }
});

app.get('/getFormData', async (req, res) => {
  try {
    const { id_solicitud } = req.query;
    const spreadsheetId = '16XaKQ0UAljlVmKKqB3xXN8L9NQlMoclCUqBPRVxI-sA';
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    // Obtener los detalles de la solicitud desde la hoja RESPUESTAS
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `RESPUESTAS!A${parseInt(id_solicitud) + 1}:AI${parseInt(id_solicitud) + 1}`,
    });

    if (!response.data.values || response.data.values.length === 0) {
      return res.status(404).json({ error: 'No se encontraron datos para la solicitud proporcionada' });
    }

    const formData = response.data.values[0];

    res.status(200).json({
      id_solicitud: formData[0],
      fecha_solicitud: formData[1],
      nombre_actividad: formData[2],
      nombre_solicitante: formData[3],
      nombre_dependencia: formData[4],
      tipo: formData[5],
      modalidad: formData[6],
      ofrecido_por: formData[7],
      unidad_academica: formData[8],
      ofrecido_para: formData[9],
      intensidad_horaria: formData[10],
      horas_modalidad: formData[11],
      horas_trabj_ind: formData[12],
      creditos: formData[13],
      cupo_min: formData[14],
      cupo_max: formData[15],
      nombre_coordinador: formData[16],
      tel_coordinador: formData[17],
      profesor_participante: formData[18],
      formas_evaluacion: formData[19],
      certificado_solicitado: formData[20],
      calificacion_minima: formData[21],
      razon_no_certificado: formData[22],
      valor_inscripcion: formData[23],
      becas_convenio: formData[24],
      becas_estudiantes: formData[25],
      becas_docentes: formData[26],
      becas_otros: formData[27],
      becas_total: formData[28],
      fechas_actividad: formData[29],
      organizacion_actividad: formData[30],
      nombre_firma: formData[31],
      cargo_firma: formData[32],
      firma: formData[33],
      anexo_documento: formData[34]
    });
  } catch (error) {
    console.error('Error al obtener los datos del formulario:', error);
    res.status(500).json({ error: 'Error al obtener los datos del formulario' });
  }
});



app.listen(PORT, () => {
  console.log(`Servidor de extensión escuchando en el puerto ${PORT}`);
});
