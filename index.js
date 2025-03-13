// Importaciones y Configuraciones Iniciales
const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const ExcelJS = require('exceljs');
const path = require('path');
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
const axios = require('axios');

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

app.post('/guardarProgreso', upload.single('pieza_grafica'), async (req, res) => {
  const { id_solicitud, paso, hoja, id_usuario, name, ...formData } = req.body;
  const piezaGrafica = req.file; // Archivo de la pieza gráfica si existe

  console.log('Recibiendo datos para guardar progreso:');
  console.log('Body completo:', req.body);
  console.log('Archivo:', req.file);

  // Convertir hoja a número 
  const parsedHoja = parseInt(hoja, 10);
  if (isNaN(parsedHoja)) {
    console.error('Hoja no válida: no es un número');
    return res.status(400).json({ error: 'Hoja no válida: no es un número' });
  }

  try {
    const sheets = getSpreadsheet();
    let sheetName = '';
    let columnas = [];
    const fechaActual = new Date().toLocaleDateString();

    // Identificar la hoja y las columnas según el formulario
    switch (parsedHoja) {
      case 1:
        sheetName = 'SOLICITUDES';
        columnas = {
          1: ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'],
          2: ['J', 'K', 'L', 'M', 'N'],
          3: ['O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y'],
          4: ['Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH'],
          5: ['AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU']
        };
        break;
      case 2:
        sheetName = 'SOLICITUDES2';
        columnas = {
          1: ['B', 'C'],
          2: ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI'],
          3: ['CJ', 'CK', 'CL']
        };
        break;
      case 3:
        sheetName = 'SOLICITUDES3';
        columnas = {
          1: ['B', 'C', 'D', 'E', 'F'],
          2: ['G', 'H', 'I', 'J'],
          3: ['K', 'L', 'M', 'N', 'O'],
          4: ['P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'],
          5: ['AA', 'AB', 'AC']
        };
        break;
      case 4:
        sheetName = 'SOLICITUDES4';
        columnas = {
          1: ['B', 'C'],
          2: ['D', 'E', 'F', 'G', 'H', 'I'],
          3: ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X'],
          4: ['Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO'],
          5: ['AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK']
        };
        break;
        case 5:
          sheetName = 'GASTOS';
          columnas = {
            1: ['B', 'C', 'D', 'E', 'F', 'G'],
            2: ['H', 'I', 'J', 'K'],
          };
          break;
      default:
        return res.status(400).json({ error: 'Hoja no válida' });
    }

    // Buscar el id_solicitud en la primera columna
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A:A`
    });

    console.log('Filas obtenidas de Google Sheets:', response.data.values);

    const rows = response.data.values || [];
    let fila = rows.findIndex((row) => row[0] === id_solicitud.toString());

    console.log('Índice de la fila para la solicitud:', fila);

    // Si no existe, agrega una nueva fila
    if (fila === -1) {
      console.log('Solicitud no encontrada, agregando nueva fila...');
      fila = rows.length + 1;
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!A${fila}`,
        valueInputOption: 'RAW',
        resource: { values: [[id_solicitud]] }
      });
      console.log('Nueva fila añadida en la posición:', fila);
    } else {
      fila += 1; // Ajustar el índice para la API de Google Sheets (basada en 1)
    }

    // Subir pieza gráfica a Google Drive si existe
    let piezaGraficaUrl = '';
    if (piezaGrafica) {
      const fileMetadata = {
        name: piezaGrafica.originalname,
        parents: ['1iDJTcUYCV7C7dTsa0Y3rfBAjFUUelu-x']
      };
      const media = {
        mimeType: piezaGrafica.mimetype,
        body: fs.createReadStream(piezaGrafica.path)
      };
      try {
        const uploadedFile = await drive.files.create({
          requestBody: fileMetadata,
          media: media,
          fields: 'id'
        });
        const fileId = uploadedFile.data.id;

        // Hacer el archivo accesible públicamente
        await drive.permissions.create({
          fileId,
          requestBody: {
            role: 'reader',
            type: 'anyone'
          }
        });

        piezaGraficaUrl = `https://drive.google.com/file/d/${fileId}/view`;
      } catch (error) {
        console.error('Error al subir la pieza gráfica a Google Drive:', error);
        return res.status(500).json({ error: 'Error al subir la pieza gráfica' });
      }
    }

    // Actualizar el progreso del formulario
    const columnasPaso = columnas[paso];

    // Añadir la URL de la pieza gráfica solo si está presente
    const valores = [...Object.values(formData)];
    if (piezaGraficaUrl) {
      valores.push(piezaGraficaUrl);
    }

    // Asegurarnos de no enviar más valores de los que se esperan para las columnas
    const valoresFinales = valores.slice(0, columnasPaso.length);

    console.log('Columnas esperadas:', columnasPaso.length);
    console.log('Valores enviados:', valoresFinales.length);
    console.log('Valores enviados:', valoresFinales);

    if (valoresFinales.length !== columnasPaso.length) {
      return res.status(400).json({ error: 'Número de columnas no coincide con los valores' });
    }

    const columnaInicial = columnasPaso[0];
    const columnaFinal = columnasPaso[columnasPaso.length - 1];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!${columnaInicial}${fila}:${columnaFinal}${fila}`,
      valueInputOption: 'RAW',
      resource: { values: [valoresFinales] }
    });

    // Actualizar hoja de ETAPAS
    const estadoGlobal = (parsedHoja === 4 && paso === maxPasos[3]) ? 'Completado' : 'En progreso';

    const maxPasos = {
      1: 5,
      2: 2,
      3: 5,
      4: 5
    };

    const etapaActual = (paso === maxPasos[parsedHoja]) ? parsedHoja + 1 : parsedHoja;

    if (etapaActual > 4) etapaActual = 4;

    // Obtener los datos actuales de ETAPAS (columnas A hasta H)
    const etapasResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'ETAPAS!A:H'
    });
    const etapasRows = etapasResponse.data.values || [];

    // Buscar la fila que corresponde al id_solicitud en la columna A de ETAPAS
    let filaEtapas = etapasRows.findIndex(row => row[0] === id_solicitud.toString());

    if (filaEtapas === -1) {
      // Si no se encuentra, usar append para agregar una nueva fila
      filaEtapas = etapasRows.length + 1; // La API es 1-based
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `ETAPAS!A${filaEtapas}:H${filaEtapas}`,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: {
          values: [[
            id_solicitud,
            id_usuario,
            fechaActual,
            name,
            etapaActual,
            estadoGlobal,
            formData.nombre_actividad || 'N/A',
            paso  // Columna H con el paso actual
          ]]
        }
      });
    } else {
      filaEtapas += 1; // Ajustar índice a 1-based
  
  // Calcular valores actualizados
  const estadoGlobal = (parsedHoja === 4 && paso === 5) ? 'Completado' : 'En progreso';
  const etapaActual = paso === 5 ? parsedHoja + 1 : parsedHoja;

  // Actualizar múltiples columnas usando batchUpdate
  const updateRequests = [
    {
      range: `ETAPAS!E${filaEtapas}`, // Columna E: etapa_actual
      values: [[etapaActual]]
    },
    {
      range: `ETAPAS!F${filaEtapas}`, // Columna F: estado
      values: [[estadoGlobal]]
    },
    {
      range: `ETAPAS!H${filaEtapas}`, // Columna H: paso
      values: [[paso]]
    }
  ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    resource: {
      valueInputOption: 'RAW',
      data: updateRequests
    }
  });
}

    // Enviar respuesta final al cliente para indicar éxito
    res.status(200).json({ success: true });
    
  } catch (error) {
    console.error("Error en guardarProgreso:", error);
    res.status(500).json({
      success: false,
      error: 'Error de conexión con Google Sheets',
      details: error.message
    });
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
      range: `${sheetName}!A2:A`, // Columna A contiene los id_solicitud
    });

    const rows = response.data.values || [];
    
    // Buscar el ID más alto en lugar de contar filas
    const lastId = rows
      .map(row => parseInt(row[0], 10)) // Convertir a número
      .filter(id => !isNaN(id)) // Eliminar valores no numéricos
      .reduce((max, id) => Math.max(max, id), 0); // Encontrar el máximo

    res.status(200).json({ lastId });
  } catch (error) {
    console.error('Error al obtener el último ID:', error);
    res.status(500).json({ error: 'Error al obtener el último ID' });
  }
});

app.post('/createNewRequest', async (req, res) => {
  try {
    const { id_solicitud, fecha_solicitud, nombre_actividad, nombre_solicitante, dependencia_tipo, nombre_dependencia } = req.body;

    const sheets = getSpreadsheet();
    const range = 'SOLICITUDES!A2:F2'; // Rango donde se insertarán los datos
    const values = [[id_solicitud, fecha_solicitud, nombre_actividad, nombre_solicitante, dependencia_tipo, nombre_dependencia]];

    // Insertar nueva fila en Google Sheets
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      resource: { values },
    });

    console.log(`✅ Nueva solicitud guardada en Sheets con ID: ${id_solicitud}`);
    res.status(200).json({ success: true, id_solicitud });
  } catch (error) {
    console.error('🚨 Error al crear la nueva solicitud en Sheets:', error);
    res.status(500).json({ error: 'Error al crear la nueva solicitud' });
  }
});


app.get('/getRequests', async (req, res) => {
  try {
    const { userId } = req.query;
    const sheets = getSpreadsheet();

    const activeResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `ETAPAS!A2:H`,
    });

    const rows = activeResponse.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron solicitudes activas o terminadas' });
    }

    const activeRequests = rows.filter(
      (row) => row[1] === userId && row[5] === 'En progreso'
    );

    const completedRequests = rows.filter(
      (row) => row[1] === userId && row[5] === 'Completado'
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

// Extraer solicitudes en progreso
app.get('/getActiveRequests', async (req, res) => {
  try {
    const { userId } = req.query;
    console.log('Obteniendo solicitudes activas para el usuario:', userId);
    const sheets = getSpreadsheet();

    // Obtener datos de ETAPAS
    const etapasResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `ETAPAS!A2:I`,
    });

    const rows = etapasResponse.data.values;

    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron solicitudes activas' });
    }

    // Filtrar solicitudes activas
    const activeRequests = rows
  .filter((row) => row[1] === userId && row[5] === 'En progreso')
  .map((row) => ({
    idSolicitud: row[0],
    formulario: parseInt(row[4]),
    paso: parseInt(row[7]),
    nombre_actividad: row[6],
    // Nuevo campo: estado por formulario
    formulariosCompletados: {
      1: parseInt(row[7]) >= 5,  // 5 pasos para formulario 1
      2: parseInt(row[7]) >= 3,  // 3 pasos para formulario 2
      3: parseInt(row[7]) >= 5,
      4: parseInt(row[7]) >= 5
    }
  }));


    // Obtener datos de SOLICITUDES para buscar nombre_actividad
    const solicitudesResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `SOLICITUDES!A2:D`, // Ajusta la columna de nombre_actividad
    });

    const solicitudesRows = solicitudesResponse.data.values;

    // Combinar datos de ETAPAS y SOLICITUDES
    const combinedRequests = activeRequests.map((request) => {
      const solicitud = solicitudesRows.find((row) => row[0] === request.idSolicitud); // Comparar por idSolicitud
      return {
        ...request,
        nombre_actividad: solicitud ? solicitud[2] : 'Sin nombre', // Ajusta el índice de nombre_actividad
      };
    });

    console.log('Solicitudes activas combinadas:', combinedRequests);

    res.status(200).json(combinedRequests);
  } catch (error) {
    console.error('Error al obtener solicitudes activas:', error);
    res.status(500).json({ error: 'Error al obtener solicitudes activas' });
  }
});


app.get('/getCompletedRequests', async (req, res) => {
  try {
    const { userId } = req.query;
    console.log('Obteniendo solicitudes activas para el usuario:', userId);
    const sheets = getSpreadsheet();

    const etapasResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `ETAPAS!A2:I`, // Asegúrate de incluir la columna de formulario y paso
    });

    const rows = etapasResponse.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron solicitudes activas' });
    }

    const activeRequests = rows.filter((row) => row[1] === userId && row[5] === 'Completado')
      .map((row) => ({
        idSolicitud: row[0], // id_solicitud
        formulario: parseInt(row[4]), // columna para el formulario
        etapa_actual: parseInt(row[4]), // Etapa actual
        paso: parseInt(row[7]), // columna para el paso
        nombre_actividad: row[6] // nombre de la actividad
      }));

      console.log('Solicitudes activas:', activeRequests);

    res.status(200).json(activeRequests);
  } catch (error) {
    console.error('Error al obtener solicitudes activas:', error);
    res.status(500).json({ error: 'Error al obtener solicitudes activas' });
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

app.get('/getSolicitud', async (req, res) => {
  try {
    const { id_solicitud } = req.query;
    const sheets = getSpreadsheet();

    if (!id_solicitud) {
      return res.status(400).json({ error: 'El ID de la solicitud es requerido' });
    }

    // Definir las hojas y el mapeo de columnas a campos
    const hojas = {
      SOLICITUDES: {
        range: 'SOLICITUDES!A2:AQ',
        fields: [
          'id_solicitud', 'fecha_solicitud', 'nombre_actividad', 'nombre_solicitante', 'dependencia_tipo', 
          'nombre_escuela', 'nombre_departamento', 'nombre_seccion', 'nombre_dependencia', 
          'introduccion', 'objetivo_general', 'objetivos_especificos', 'justificacion', 'metodologia', 'tipo', 
          'otro_tipo', 'modalidad', 'horas_trabajo_presencial', 'horas_sincronicas', 'total_horas', 
          'programCont', 'dirigidoa', 'creditos', 'cupo_min', 'cupo_max', 'nombre_coordinador', 
          'correo_coordinador', 'tel_coordinador', 'perfil_competencia', 'formas_evaluacion', 
          'certificado_solicitado', 'calificacion_minima', 'razon_no_certificado', 'valor_inscripcion', 
          'becas_convenio', 'becas_estudiantes', 'becas_docentes', 'becas_egresados', 'becas_funcionarios', 
          'becas_otros', 'becas_total', 'periodicidad_oferta', 'fechas_actividad', 'organizacion_actividad'
        ]
      },
      SOLICITUDES2: {
        range: 'SOLICITUDES2!A2:CL',
        fields: [
          'id_solicitud', 'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos', 
          'costos_personal_cantidad', 'costos_personal_vr_unit', 'total_costos_personal', 
          'personal_universidad_cantidad', 'personal_universidad_vr_unit', 'total_personal_universidad', 
          'honorarios_docentes_cantidad', 'honorarios_docentes_vr_unit', 'total_honorarios_docentes',
          'otro_personal_cantidad', 'otro_personal_vr_unit', 'total_otro_personal', 
          'materiales_sumi_cantidad', 'materiales_sumi_vr_unit', 'total_materiales_sumi',
          'gastos_alojamiento_cantidad', 'gastos_alojamiento_vr_unit', 'total_gastos_alojamiento',
          'gastos_alimentacion_cantidad', 'gastos_alimentacion_vr_unit', 'total_gastos_alimentacion',
          'gastos_transporte_cantidad', 'gastos_transporte_vr_unit', 'total_gastos_transporte',
          'equipos_alquiler_compra_cantidad', 'equipos_alquiler_compra_vr_unit', 'total_equipos_alquiler_compra',
          'dotacion_participantes_cantidad', 'dotacion_participantes_vr_unit', 'total_dotacion_participantes',
          'carpetas_cantidad', 'carpetas_vr_unit', 'total_carpetas',
          'libretas_cantidad', 'libretas_vr_unit', 'total_libretas',
          'lapiceros_cantidad', 'lapiceros_vr_unit', 'total_lapiceros',
          'memorias_cantidad', 'memorias_vr_unit', 'total_memorias',
          'marcadores_papel_otros_cantidad', 'marcadores_papel_otros_vr_unit', 'total_marcadores_papel_otros',
          'impresos_cantidad', 'impresos_vr_unit', 'total_impresos',
          'labels_cantidad', 'labels_vr_unit', 'total_labels',
          'certificados_cantidad', 'certificados_vr_unit', 'total_certificados',
          'escarapelas_cantidad', 'escarapelas_vr_unit', 'total_escarapelas',
          'fotocopias_cantidad', 'fotocopias_vr_unit', 'total_fotocopias',
          'estacion_cafe_cantidad', 'estacion_cafe_vr_unit', 'total_estacion_cafe',
          'transporte_mensaje_cantidad', 'transporte_mensaje_vr_unit', 'total_transporte_mensaje',
          'refrigerios_cantidad', 'refrigerios_vr_unit', 'total_refrigerios',
          'infraestructura_fisica_cantidad', 'infraestructura_fisica_vr_unit', 'total_infraestructura_fisica',
          'gastos_generales_cantidad', 'gastos_generales_vr_unit', 'total_gastos_generales',
          'infraestructura_universitaria_cantidad', 'infraestructura_universitaria_vr_unit', 
          'total_infraestructura_universitaria', 'imprevistos',
          'escuela_departamento_porcentaje', 'total_aportes_univalle'
        ]
      },
      SOLICITUDES4: {
        range: 'SOLICITUDES4!A2:BK',
        fields: [
          'id_solicitud', 'descripcionPrograma', 'identificacionNecesidades', 'atributosBasicos', 
          'atributosDiferenciadores', 'competencia', 'programa', 'programasSimilares', 
          'estrategiasCompetencia', 'personasInteres', 'personasMatriculadas', 'otroInteres', 
          'innovacion', 'solicitudExterno', 'interesSondeo', 'llamadas', 'encuestas', 'webinar', 
          'preregistro', 'mesasTrabajo', 'focusGroup', 'desayunosTrabajo', 'almuerzosTrabajo', 'openHouse', 
          'valorEconomico', 'modalidadPresencial', 'modalidadVirtual', 'modalidadSemipresencial', 
          'otraModalidad', 'beneficiosTangibles', 'beneficiosIntangibles', 'particulares', 'colegios', 
          'empresas', 'egresados', 'colaboradores', 'otros_publicos_potenciales', 'tendenciasActuales', 
          'dofaDebilidades', 'dofaOportunidades', 'dofaFortalezas', 'dofaAmenazas', 'paginaWeb', 
          'facebook', 'instagram', 'linkedin', 'correo', 'prensa', 'boletin', 'llamadas_redes', 'otro_canal'
        ]
      },
      SOLICITUDES3: {
        range: 'SOLICITUDES3!A2:AC',
        fields: [
          'id_solicitud', 'proposito', 'comentario', 'fecha', 'elaboradoPor', 'aplicaDiseno1', 'aplicaDiseno2', 
          'aplicaDiseno3', 'aplicaLocacion1', 'aplicaLocacion2', 'aplicaLocacion3', 'aplicaDesarrollo1', 
          'aplicaDesarrollo2', 'aplicaDesarrollo3', 'aplicaDesarrollo4', 'aplicaCierre1', 'aplicaCierre2', 
          'aplicaOtros1', 'aplicaOtros2'
        ]
      },
      GASTOS: {
        range: 'GASTOS!A2:E',
        fields: [
          'id_conceptos', 'id_solicitud', 'id_gastos', 'cantidad', 'valor_unitario', 'valor_total'
        ]
      },
    };
    let solicitudEncontrada = false;
    const resultados = {};

    // Recorremos cada hoja y buscamos los datos asociados al id_solicitud
    for (let [hoja, { range, fields }] of Object.entries(hojas)) {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range
      });

      const rows = response.data.values || [];

      // Buscar la fila que coincida con el id_solicitud
      const solicitudData = rows.find(row => row[0] === id_solicitud);

      if (solicitudData) {
        // Crear un objeto donde las claves son los nombres de los campos
        solicitudEncontrada = true;
        resultados[hoja] = fields.reduce((acc, field, index) => {
        //const mappedData = fields.reduce((acc, field, index) => {
          acc[field] = solicitudData[index] || ''; // Asigna el valor correspondiente o vacío si no existe
          return acc;
        }, {});

        // Almacenar los datos mapeados de esta hoja dentro del objeto `resultados`
        //resultados[hoja] = mappedData;
      }
    }

    // Si no encontramos la solicitud, devolvemos un objeto vacío en lugar de un error 404
    if (!solicitudEncontrada) {
      return res.status(200).json({ message: 'La solicitud no existe aún en Google Sheets', data: {} });
    }

    // Verificar si se encontraron datos en al menos una hoja
    if (Object.keys(resultados).length === 0) {
      return res.status(404).json({ error: 'No se encontraron datos para esta solicitud' });
    }

    // Devolver todos los datos encontrados
    res.status(200).json(resultados);
  } catch (error) {
    console.error('Error al obtener los datos de la solicitud:', error);
    res.status(500).json({ error: 'Error al obtener los datos de la solicitud' });
  }
});

async function replaceMarkers(templateId, data, fileName) {
  try {
    console.log(`Generando archivo desde la plantilla: ${templateId}`);
    const copiedFile = await drive.files.copy({
      fileId: templateId,
      requestBody: {
        name: fileName,
        parents: [folderId],
      },
    });

    const fileId = copiedFile.data.id;

    // Descargar archivo .xlsx
    const dest = path.resolve(__dirname, `${fileName}.xlsx`);
    const destStream = fs.createWriteStream(dest);
    await drive.files.get(
      { fileId, alt: 'media' },
      { responseType: 'stream' },
      (err, { data }) => {
        if (err) throw new Error('Error al descargar el archivo XLSX');
        data.pipe(destStream);
      }
    );

    await new Promise((resolve) => destStream.on('finish', resolve));
    console.log('Archivo XLSX descargado:', dest);

    // Leer el archivo .xlsx
    const workbook = XLSX.readFile(dest);
    const sheetNames = workbook.SheetNames;

    sheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Reemplazar marcadores
      const updatedRows = rows.map((row) =>
        row.map((cell) =>
          typeof cell === 'string'
            ? Object.keys(data).reduce(
                (updatedCell, marker) => updatedCell.replace(`{{${marker}}}`, data[marker] || ''),
                cell
              )
            : cell
        )
      );

      // Sobrescribir la hoja con los valores actualizados
      workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(updatedRows);
    });

    // Guardar el archivo modificado localmente
    const updatedFilePath = path.resolve(__dirname, `Updated_${fileName}.xlsx`);
    XLSX.writeFile(workbook, updatedFilePath);

    // Subir el archivo modificado a Google Drive
    const updatedFile = await drive.files.create({
      requestBody: {
        name: `Updated_${fileName}`,
        parents: [folderId],
      },
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: fs.createReadStream(updatedFilePath),
      },
    });

    const updatedFileId = updatedFile.data.id;

    // Otorgar permisos públicos al archivo
    await drive.permissions.create({
      fileId: updatedFileId,
      requestBody: {
        role: 'reader',
        type: 'anyone',
      },
    });

    console.log(`Archivo actualizado y subido: ${updatedFileId}`);
    return `https://drive.google.com/file/d/${updatedFileId}/view`;
  } catch (error) {
    console.error('Error al reemplazar marcadores en el archivo XLSX:', error.message);
    throw new Error('Error al reemplazar marcadores en el archivo XLSX');
  }
}

// Función para formatear partes de una fecha
const formatDateParts = (date) => {
  const fecha = new Date(date);
  return {
    dia: fecha.getDate().toString().padStart(2, '0'),
    mes: (fecha.getMonth() + 1).toString().padStart(2, '0'),
    anio: fecha.getFullYear().toString(),
  };
};


// Función para obtener datos desde Google Sheets
const getSolicitudData = async (solicitudId, sheets, spreadsheetId, hojas) => {
  try {
    const resultados = {};

    for (let [hoja, { range, fields }] of Object.entries(hojas)) {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
      });

      const rows = response.data.values || [];

      const solicitudData = rows.find((row) => row[0] === solicitudId);

      if (solicitudData) {
        const mappedData = fields.reduce((acc, field, index) => {
          acc[field] = solicitudData[index] || '';
          return acc;
        }, {});

        resultados[hoja] = mappedData;
      }
    }

    if (Object.keys(resultados).length === 0) {
      throw new Error('No se encontraron datos para esta solicitud');
    }

    return resultados;
  } catch (error) {
    console.error('Error al obtener los datos de la solicitud:', error.message);
    throw new Error('Error al obtener los datos de la solicitud');
  }
};

// Función para transformar datos de la solicitud para casillas de selección
const transformDataForTemplate = (formData) => {
  // Procesar "tipo"
  const tipo = formData.tipo || '';
  const tipoData = {
    tipo_curso: tipo === 'Curso' ? 'X' : '',
    tipo_taller: tipo === 'Taller' ? 'X' : '',
    tipo_seminario: tipo === 'Seminario' ? 'X' : '',
    tipo_diplomado: tipo === 'Diplomado' ? 'X' : '',
    tipo_programa: tipo === 'Programa' ? 'X' : '',
    otro_cual: tipo === 'Otro' ? formData.otro_tipo || '' : '',
  };

  // Procesar "modalidad"
  const modalidad = formData.modalidad || '';
  const modalidadData = {
    modalidad_presencial: modalidad === 'Presencial' ? 'X' : '',
    modalidad_tecnologia: modalidad === 'Presencialidad asistida por Tecnología' ? 'X' : '',
    modalidad_virtual: modalidad === 'Virtual' ? 'X' : '',
    modalidad_mixta: modalidad === 'Mixta' ? 'X' : '',
    modalidad_todas: modalidad === 'Todas las anteriores' ? 'X' : '',
  };

  // Procesar periodicidad
  const periodicidad = formData.periodicidad_oferta || '';
  const periodicidadData = {
    per_anual: periodicidad === 'Anual' ? 'X' : '',
    per_semestral: periodicidad === 'Semestral' ? 'X' : '',
    per_permanente: periodicidad === 'Permanente' ? 'X' : '',
  };

  // Procesar organización de la actividad
  const organizacion = formData.organizacion_actividad || '';
  const organizacionData = {
    oficina_extension: organizacion === 'Oficina de Extensión' ? 'X' : '',
    unidad_acad: organizacion === 'Unidad Académica' ? 'X' : '',
    otro_organizacion: organizacion === 'Otro' ? 'X' : '',
    cual_otro: organizacion === 'Otro' ? formData.otro_tipo_act || '' : '',
  };

  // Combinar todos los datos transformados
  return {
    ...formData,
    ...tipoData,
    ...modalidadData,
    ...periodicidadData,
    ...organizacionData,
  };
};

// Función para procesar y reemplazar marcadores en archivos XLSX
const processXLSXWithStyles = async (templateId, data, fileName, folderId) => {
  try {
    console.log(`Descargando la plantilla: ${templateId}`);
    const fileResponse = await drive.files.get(
      { fileId: templateId, alt: 'media' },
      { responseType: 'stream' }
    );

    const tempFilePath = path.join('/tmp', `${fileName}.xlsx`);
    const writeStream = fs.createWriteStream(tempFilePath);
    await new Promise((resolve, reject) => {
      fileResponse.data.pipe(writeStream);
      fileResponse.data.on('end', resolve);
      fileResponse.data.on('error', reject);
    });

    console.log('Leyendo el archivo descargado...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(tempFilePath);

    console.log('Procesando celdas para reemplazo...');
    workbook.eachSheet((sheet) => {
      sheet.eachRow((row) => {
        row.eachCell((cell) => {
          if (typeof cell.value === 'string') {
            Object.keys(data).forEach((key) => {
              const marker = `{{${key}}}`;
              if (cell.value.includes(marker)) {
                cell.value = cell.value.replace(marker, data[key] || '');
              }
            });
          }
        });
      });
    });

    const updatedFilePath = path.join('/tmp', `updated_${fileName}.xlsx`);
    await workbook.xlsx.writeFile(updatedFilePath);
    console.log('Archivo actualizado con los datos reemplazados.');

    console.log('Subiendo el archivo actualizado a Google Drive...');
    const uploadResponse = await drive.files.create({
      requestBody: {
        name: fileName,
        parents: [folderId],
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: fs.createReadStream(updatedFilePath),
      },
    });

    const fileId = uploadResponse.data.id;
    await drive.permissions.create({
      fileId,
      requestBody: {
        role: 'reader',
        type: 'anyone',
      },
    });

    const link = `https://drive.google.com/file/d/${fileId}/view`;
    console.log(`Archivo subido y disponible en: ${link}`);
    return link;
  } catch (error) {
    console.error('Error al procesar archivo XLSX con estilos:', error.message);
    throw new Error('Error al procesar archivo XLSX con estilos');
  }
};

app.post('/guardarGastos', async (req, res) => {
  try {
    const sheets = getSpreadsheet();
    const { id_solicitud, gastos } = req.body;

    // Validación más flexible
    if (!id_solicitud || !gastos?.length) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere id_solicitud y al menos un concepto'
      });
    }

    // Convertir a string para evitar errores de tipo
    const idSolicitudStr = id_solicitud.toString();

    // Obtener conceptos válidos (desde tu hoja CONCEPTO$)
    const conceptosResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'CONCEPTO$!A2:A' // Solo columna de ids
    });

    const conceptosValidos = new Set(
      (conceptosResponse.data.values || []).flat().map(String)
    );

    console.log("Datos recibidos:", JSON.stringify(gastos, null, 2));
    console.log("Conceptos válidos:", [...conceptosValidos]);

    // Preparar filas válidas
    const rows = [];
    for (const gasto of gastos) {
      // Validar existencia del concepto
      if (!gasto.id_conceptos) {
        console.log('Falta id_conceptos en', gasto);
        continue;
      }

      if (!conceptosValidos.has(String(gasto.id_conceptos))) {
        console.log(`Concepto ${gasto.id_conceptos} no encontrado`);
        continue;
      }

      const cantidad = parseFloat(gasto.cantidad) || 0;
      const valor_unit = parseFloat(gasto.valor_unit) || 0;
      const valor_total = cantidad * valor_unit;

      rows.push([
        gasto.id_conceptos.toString(), // Columna A (CONCEPTO)
        idSolicitudStr, // Columna B (ID_SOLICITUD)
        gasto.cantidad || 0, // Columna C (CANTIDAD)
        gasto.valor_unit || 0, // Columna D (VALOR_UNIT)
        (gasto.cantidad || 0) * (gasto.valor_unit || 0) // Columna E (VALOR_TOTAL)
      ]);
    }

    if (rows.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'Ningún concepto válido para guardar'
      });
    }

    // Insertar en GASTOS
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'GASTOS!A2:E',
      valueInputOption: 'USER_ENTERED',
      resource: { values: rows }
    });

    res.json({ success: true });

  } catch (error) {
    console.error("Error en guardarGastos:", error);
    res.status(500).json({
      success: false,
      error: 'Error de conexión con Google Sheets',
      details: error.message
    });
  }
});


app.post('/generateReport', async (req, res) => {
  try {
    const { solicitudId, formNumber } = req.body;
    console.log("Datos recibidos en generateReport:");
    console.log("solicitudId:", solicitudId);
    console.log("formNumber:", formNumber);
  

    if (!solicitudId || !formNumber) {
      console.error('Error: Los parámetros solicitudId y formNumber son requeridos');
      return res.status(400).json({ error: 'Los parámetros solicitudId y formNumber son requeridos' });
    }

    const folderId = '12bxb0XEArXMLvc7gX2ndqJVqS_sTiiUE';
    const templateIds = {
      1: '1xsz9YSnYEOng56eNKGV9it9EgTn0mZw1',
      2: '1JY-4IfJqEWLqZ_wrq_B_bfIlI9MeVzgF',
      3: '1FTC7Vq3O4ultexRPXYrJKOpL9G0071-0',
      4: '1WoPUZYusNl2u3FpmZ1qiO5URBUqHIwKF',
    };

    const templateId = templateIds[formNumber];
    if (!templateId) {
      console.error(`Formulario no válido: ${formNumber}`);
      return res.status(400).json({ error: 'Número de formulario no válido' });
    }

    const hojas = {
      SOLICITUDES: {
        range: 'SOLICITUDES!A2:AQ',
        fields: [
          'id_solicitud', 'fecha_solicitud', 'nombre_actividad', 'nombre_solicitante', 'dependencia_tipo',
          'nombre_escuela', 'nombre_departamento', 'nombre_seccion', 'nombre_dependencia', 'introduccion',
          'objetivo_general', 'objetivos_especificos', 'justificacion', 'metodologia', 'tipo', 'otro_tipo',
          'modalidad', 'horas_trabajo_presencial', 'horas_sincronicas', 'total_horas', 'programCont',
          'dirigidoa', 'creditos', 'cupo_min', 'cupo_max', 'nombre_coordinador', 'correo_coordinador',
          'tel_coordinador', 'perfil_competencia', 'formas_evaluacion', 'certificado_solicitado',
          'calificacion_minima', 'razon_no_certificado', 'valor_inscripcion', 'becas_convenio',
          'becas_estudiantes', 'becas_docentes', 'becas_egresados', 'becas_funcionarios', 'becas_otros',
          'periodicidad_oferta', 'organizacion_actividad', 'otro_tipo_act',
        ],
      },
      SOLICITUDES2: {
        range: 'SOLICITUDES2!A2:CL',
        fields: [
          'id_solicitud', 'ingresos_cantidad', 'ingresos_vr_unit', 'total_ingresos',
          'costos_personal_cantidad', 'costos_personal_vr_unit', 'total_costos_personal',
        ],
      },
      SOLICITUDES4: {
        range: 'SOLICITUDES4!A2:BK',
        fields: [
          'id_solicitud', 'descripcionPrograma', 'identificacionNecesidades', 'atributosBasicos',
          'atributosDiferenciadores', 'competencia', 'programa', 'programasSimilares',
          'estrategiasCompetencia', 'personasInteres', 'personasMatriculadas', 'otroInteres',
          'innovacion', 'solicitudExterno', 'interesSondeo', 'llamadas', 'encuestas', 'webinar',
          'preregistro', 'mesasTrabajo', 'focusGroup', 'desayunosTrabajo', 'almuerzosTrabajo', 'openHouse',
          'valorEconomico', 'modalidadPresencial', 'modalidadVirtual', 'modalidadSemipresencial',
          'otraModalidad', 'beneficiosTangibles', 'beneficiosIntangibles', 'particulares', 'colegios',
          'empresas', 'egresados', 'colaboradores', 'otros_publicos_potenciales', 'tendenciasActuales',
          'dofaDebilidades', 'dofaOportunidades', 'dofaFortalezas', 'dofaAmenazas', 'paginaWeb',
          'facebook', 'instagram', 'linkedin', 'correo', 'prensa', 'boletin', 'llamadas_redes', 'otro_canal',
        ],
      },
      SOLICITUDES3: {
        range: 'SOLICITUDES3!A2:AC',
        fields: [
          'id_solicitud', 'proposito', 'comentario', 'fecha', 'elaboradoPor', 'aplicaDiseno1', 'aplicaDiseno2',
          'aplicaDiseno3', 'aplicaLocacion1', 'aplicaLocacion2', 'aplicaLocacion3', 'aplicaDesarrollo1',
          'aplicaDesarrollo2', 'aplicaDesarrollo3', 'aplicaDesarrollo4', 'aplicaCierre1', 'aplicaCierre2',
          'aplicaOtros1', 'aplicaOtros2',
        ],
      },
    };

    const sheets = getSpreadsheet();

    // Usar getSolicitudData para obtener los datos
    const solicitudData = await getSolicitudData(solicitudId, sheets, SPREADSHEET_ID, hojas);

    console.log('Transformando datos para la plantilla...');
    const transformedData = transformDataForTemplate(solicitudData.SOLICITUDES);

    console.log(`Generando reporte para el formulario ${formNumber}...`);
    const reportLink = await processXLSXWithStyles(
      templateId,
      transformedData,
      `Formulario${formNumber}_${solicitudId}`,
      folderId
    );

    res.status(200).json({
      message: `Informe generado exitosamente para el formulario ${formNumber}`,
      link: reportLink,
    });
  } catch (error) {
    console.error('Error al generar el informe:', error.message);
    res.status(500).json({ error: 'Error al generar el informe' });
  }
});


app.listen(PORT, () => {
  console.log(`Servidor de extensión escuchando en el puerto ${PORT}`);
});
