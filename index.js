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

app.post('/guardarProgreso', async (req, res) => {
  const { id_solicitud, formData, paso, hoja, userData } = req.body;

  try {
    const sheets = getSpreadsheet();
    let sheetName = '';
    let columnas = [];
    const fechaActual = new Date().toLocaleDateString();

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

    // Definir el estado global de la solicitud
    const estadoGlobal = (hoja === 5 && paso === 5) ? 'Completado' : 'En progreso';
    const etapaActual = hoja; // Guardar solo el número del formulario; // Actualizamos para que sea "Formulario" en lugar de "Paso"

    // Buscar el id_solicitud en la hoja ETAPAS
    const etapasResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `ETAPAS!A:A`, // Suponiendo que los ID de solicitud están en la columna A
    });

    const etapasRows = etapasResponse.data.values || [];
    let filaEtapas = etapasRows.findIndex((row) => row[0] === id_solicitud.toString());

    // Si no existe un registro en ETAPAS para esta solicitud, agregarlo
    if (filaEtapas === -1) {
      filaEtapas = etapasRows.length + 1;
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `ETAPAS!A${filaEtapas}`,
        valueInputOption: 'RAW',
        resource: {
          values: [[id_solicitud, userData.id, fechaActual, userData.name, etapaActual, estadoGlobal, formData.nombre_actividad || '', paso]]
        },
      });
    } else {
      // Si ya existe un registro, actualizamos la fila con el estado y la etapa actual
      filaEtapas += 1; // Ajustamos el índice de la fila para trabajar con la API (basada en 1)

      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `ETAPAS!B${filaEtapas}:H${filaEtapas}`, // Actualiza las columnas correspondientes
        valueInputOption: 'RAW',
        resource: {
          values: [[userData.id, fechaActual, userData.name, etapaActual, estadoGlobal, formData.nombre_actividad || '', paso]],
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

// Extraer solicitudes en progreso
app.get('/getActiveRequests', async (req, res) => {
  try {
    const { userId } = req.query;
    const sheets = getSpreadsheet();

    const etapasResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `ETAPAS!A2:I`, // Asegúrate de incluir la columna de formulario y paso
    });

    const rows = etapasResponse.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron solicitudes activas' });
    }

    const activeRequests = rows.filter((row) => row[1] === userId && row[5] === 'En progreso')
      .map((row) => ({
        idSolicitud: row[0], // id_solicitud
        formulario: parseInt(row[4]), // columna para el formulario
        paso: parseInt(row[7]), // columna para el paso
        nombre_actividad: row[6] // nombre de la actividad
      }));

    res.status(200).json(activeRequests);
  } catch (error) {
    console.error('Error al obtener solicitudes activas:', error);
    res.status(500).json({ error: 'Error al obtener solicitudes activas' });
  }
});


// Extraer solicitudes terminadas
app.get('/getCompletedRequests', async (req, res) => {
  try {
    const { userId } = req.query; // Asegúrate de recibir el id del usuario
    const sheets = getSpreadsheet();

    // Obtener la hoja de ETAPAS
    const etapasResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `ETAPAS!A2:H`, // Ajustar el rango según tu estructura
    });

    const rows = etapasResponse.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron solicitudes terminadas' });
    }

    // Filtrar las solicitudes que están terminadas (completadas)
    const completedRequests = rows.filter((row) => row[1] === userId && row[5] === 'Completado');

    res.status(200).json(completedRequests);
  } catch (error) {
    console.error('Error al obtener solicitudes terminadas:', error);
    res.status(500).json({ error: 'Error al obtener solicitudes terminadas' });
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

// app.get('/getSolicitud', async (req, res) => {
//   try {
//     const { id_solicitud } = req.query;
//     const sheets = getSpreadsheet();

//     // Definir las hojas que queremos consultar
//     const hojas = ['SOLICITUDES', 'SOLICITUDES2', 'SOLICITUDES3', 'SOLICITUDES4', 'SOLICITUDES5'];

//     // Variable para almacenar los datos de todas las hojas
//     const resultados = {};

//     // Recorremos cada hoja y buscamos los datos asociados al id_solicitud
//     for (let hoja of hojas) {
//       const response = await sheets.spreadsheets.values.get({
//         spreadsheetId: SPREADSHEET_ID,
//         range: `${hoja}!A2:CL`, // Ajusta el rango según el número de columnas de cada hoja
//       });

//       const rows = response.data.values || [];
      
//       // Buscar la fila que coincida con el id_solicitud
//       const solicitudData = rows.find(row => row[0] === id_solicitud);

//       if (solicitudData) {
//         // Almacenar los datos de esta hoja dentro del objeto `resultados`
//         resultados[hoja] = solicitudData;
//       }
//     }

//     // Verificar si se encontraron datos en al menos una hoja
//     if (Object.keys(resultados).length === 0) {
//       return res.status(404).json({ error: 'No se encontraron datos para esta solicitud' });
//     }

//     // Devolver todos los datos encontrados
//     res.status(200).json(resultados);
//   } catch (error) {
//     console.error('Error al obtener los datos de la solicitud:', error);
//     res.status(500).json({ error: 'Error al obtener los datos de la solicitud' });
//   }
// });


// ==========================================================================
// Inicializar el servidor
// ==========================================================================

app.get('/getSolicitud', async (req, res) => {
  try {
    const { id_solicitud } = req.query;
    const sheets = getSpreadsheet();

    // Definir las hojas y el mapeo de columnas a campos
    const hojas = {
      SOLICITUDES: {
        range: 'SOLICITUDES!A2:M',
        fields: [
          'id_solicitud', 'introduccion', 'objetivo_general', 'objetivos_especificos', 'justificacion', 
          'descripcion', 'alcance', 'metodologia', 'dirigido_a', 'programa_contenidos', 'duracion', 
          'certificacion', 'recursos'
        ]
      },
      SOLICITUDES2: {
        range: 'SOLICITUDES2!A2:AL',
        fields: [
          'id_solicitud', 'fecha_solicitud', 'nombre_actividad', 'nombre_solicitante', 'dependencia_tipo', 
          'nombre_escuela', 'nombre_departamento', 'nombre_seccion', 'nombre_dependencia', 'tipo', 
          'otro_tipo', 'modalidad', 'horas_trabajo_presencial', 'horas_sincronicas', 'total_horas', 
          'programCont', 'dirigidoa', 'creditos', 'cupo_min', 'cupo_max', 'nombre_coordinador', 
          'correo_coordinador', 'tel_coordinador', 'perfil_competencia', 'formas_evaluacion', 
          'certificado_solicitado', 'calificacion_minima', 'razon_no_certificado', 'valor_inscripcion', 
          'becas_convenio', 'becas_estudiantes', 'becas_docentes', 'becas_egresados', 'becas_funcionarios', 
          'becas_otros', 'becas_total', 'periodicidad_oferta', 'fechas_actividad', 'organizacion_actividad'
        ]
      },
      SOLICITUDES3: {
        range: 'SOLICITUDES3!A2:CL',
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
      SOLICITUDES5: {
        range: 'SOLICITUDES5!A2:AC',
        fields: [
          'id_solicitud', 'proposito', 'comentario', 'fecha', 'elaboradoPor', 'aplicaDiseno1', 'aplicaDiseno2', 
          'aplicaDiseno3', 'aplicaLocacion1', 'aplicaLocacion2', 'aplicaLocacion3', 'aplicaDesarrollo1', 
          'aplicaDesarrollo2', 'aplicaDesarrollo3', 'aplicaDesarrollo4', 'aplicaCierre1', 'aplicaCierre2', 
          'aplicaOtros1', 'aplicaOtros2'
        ]
      }
    };

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
        const mappedData = fields.reduce((acc, field, index) => {
          acc[field] = solicitudData[index] || ''; // Asigna el valor correspondiente o vacío si no existe
          return acc;
        }, {});

        // Almacenar los datos mapeados de esta hoja dentro del objeto `resultados`
        resultados[hoja] = mappedData;
      }
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

// Función para obtener datos desde Google Sheets
async function fetchSheetData(spreadsheetId, ranges) {
  const sheets = getSpreadsheet(); // Asegúrate de que `getSpreadsheet` esté definido en tu archivo
  try {
    const response = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: spreadsheetId,
      ranges: ranges,
    });

    // Mapeamos los datos de cada rango
    const data = {};
    response.data.valueRanges.forEach((valueRange, index) => {
      data[ranges[index]] = valueRange.values || []; // Si no hay valores, devolvemos un arreglo vacío
    });

    return data;
  } catch (error) {
    console.error('Error al obtener datos de Google Sheets:', error.message);
    throw new Error('Error al obtener datos de Google Sheets');
  }
}
 
// app.post('/generateReport', async (req, res) => {
//   try {
//     const { solicitudId } = req.body;

//     if (!solicitudId) {
//       console.error('Error: El parámetro solicitudId es requerido');
//       return res.status(400).json({ error: 'El parámetro solicitudId es requerido' });
//     }

//     console.log(`Procesando solicitud con ID: ${solicitudId}`);

//     // Función para obtener datos desde Google Sheets
//     async function fetchSheetData(spreadsheetId, ranges) {
//       const sheets = getSpreadsheet();
//       try {
//         console.log('Obteniendo datos de Google Sheets...');
//         const response = await sheets.spreadsheets.values.batchGet({
//           spreadsheetId: spreadsheetId,
//           ranges: ranges,
//         });

//         const data = {};
//         response.data.valueRanges.forEach((valueRange, index) => {
//           data[ranges[index]] = valueRange.values || [];
//         });

//         console.log('Datos obtenidos de Google Sheets:', data);
//         return data;
//       } catch (error) {
//         console.error('Error al obtener datos de Google Sheets:', error.message);
//         throw new Error('Error al obtener datos de Google Sheets');
//       }
//     }

//     const ranges = ['SOLICITUDES!A2:M', 'SOLICITUDES2!A2:AL'];
//     let data;

//     try {
//       data = await fetchSheetData(SPREADSHEET_ID, ranges);
//     } catch (error) {
//       console.error('Error al consultar Google Sheets:', error.message);
//       return res.status(500).json({ error: 'Error al consultar datos de Google Sheets' });
//     }

//     const resultados = {};

//     // Procesar datos obtenidos
//     ranges.forEach((range, index) => {
//       const rows = data[range];
//       if (!rows || rows.length === 0) {
//         console.warn(`No se encontraron datos en el rango: ${range}`);
//         return;
//       }

//       const solicitudData = rows.find((row) => row[0] === solicitudId);
//       if (solicitudData) {
//         const fields =
//           index === 0
//             ? ['id_solicitud', 'introduccion', 'objetivo_general', 'objetivos_especificos', 'justificacion', 'descripcion', 'alcance', 'metodologia', 'dirigido_a', 'programa_contenidos', 'duracion', 'certificacion', 'recursos']
//             : ['id_solicitud', 'fecha_solicitud', 'nombre_actividad', 'nombre_solicitante', 'dependencia_tipo', 'nombre_escuela', 'nombre_departamento', 'nombre_seccion', 'nombre_dependencia', 'tipo', 'otro_tipo', 'modalidad', 'horas_trabajo_presencial', 'horas_sincronicas', 'total_horas', 'programCont', 'dirigidoa', 'creditos', 'cupo_min', 'cupo_max', 'nombre_coordinador', 'correo_coordinador', 'tel_coordinador', 'perfil_competencia', 'formas_evaluacion', 'certificado_solicitado', 'calificacion_minima', 'razon_no_certificado', 'valor_inscripcion', 'becas_convenio', 'becas_estudiantes', 'becas_docentes', 'becas_egresados', 'becas_funcionarios', 'becas_otros', 'becas_total', 'periodicidad_oferta', 'fechas_actividad', 'organizacion_actividad'];

//         resultados[range.split('!')[0]] = fields.reduce((acc, field, idx) => {
//           acc[field] = solicitudData[idx] || '';
//           return acc;
//         }, {});
//       }
//     });

//     if (!Object.keys(resultados).length) {
//       console.warn('No se encontraron datos para la solicitud proporcionada');
//       return res.status(404).json({ error: 'No se encontraron datos para esta solicitud' });
//     }

//     console.log('Datos procesados:', resultados);

//     // Generar informe en Google Drive
//     const folderId = '12bxb0XEArXMLvc7gX2ndqJVqS_sTiiUE'; // ID de la carpeta de destino
//     const templateFileId = '1WiNfcR2_hRcvcNFohFyh0BPzLek9o9f0'; // ID de la plantilla

//     async function generateReportInDrive(templateFileId, folderId, fileName) {
//       try {
//         console.log('Generando informe en Google Drive...');
//         const copiedFile = await drive.files.copy({
//           fileId: templateFileId,
//           requestBody: {
//             name: fileName,
//             parents: [folderId],
//           },
//         });

//         const fileId = copiedFile.data.id;

//         console.log(`Archivo generado con ID: ${fileId}`);

//         // Compartir el archivo generado
//         await drive.permissions.create({
//           fileId: fileId,
//           requestBody: {
//             role: 'reader',
//             type: 'anyone',
//           },
//         });

//         console.log(`Archivo compartido públicamente: ${fileId}`);
//         return `https://drive.google.com/file/d/${fileId}/view`;
//       } catch (error) {
//         console.error('Error al generar informe en Google Drive:', error.message);
//         throw new Error('Error al generar informe en Google Drive');
//       }
//     }

//     const fileName = `Reporte_Solicitud_${solicitudId}`;
//     let generatedLink;

//     try {
//       generatedLink = await generateReportInDrive(templateFileId, folderId, fileName);
//     } catch (error) {
//       console.error('Error al generar el informe:', error.message);
//       return res.status(500).json({ error: 'Error al generar el informe' });
//     }

//     if (!generatedLink) {
//       console.error('No se generaron enlaces de informes');
//       return res.status(500).json({ error: 'No se generaron enlaces de informes' });
//     }

//     console.log('Informe generado exitosamente:', generatedLink);

//     res.status(200).json({
//       message: 'Informe generado exitosamente',
//       link: generatedLink,
//     });
//   } catch (error) {
//     console.error('Error al generar los informes:', error.message);
//     res.status(500).json({ error: 'Error al generar los informes' });
//   }
// });

app.post('/generateReport', async (req, res) => {
  try {
    const { solicitudId } = req.body;

    if (!solicitudId) {
      console.error('Error: El parámetro solicitudId es requerido');
      return res.status(400).json({ error: 'El parámetro solicitudId es requerido' });
    }

    const ranges = ['SOLICITUDES!A2:M', 'SOLICITUDES2!A2:AL'];

    // Obtener datos desde Google Sheets
    const fetchSheetData = async (spreadsheetId, ranges) => {
      const sheets = getSpreadsheet();
      try {
        const response = await sheets.spreadsheets.values.batchGet({
          spreadsheetId: spreadsheetId,
          ranges: ranges,
        });

        const data = {};
        response.data.valueRanges.forEach((valueRange, index) => {
          data[ranges[index]] = valueRange.values || [];
        });

        return data;
      } catch (error) {
        console.error('Error al obtener datos de Google Sheets:', error.message);
        throw new Error('Error al obtener datos de Google Sheets');
      }
    };

    let data;
    try {
      data = await fetchSheetData(SPREADSHEET_ID, ranges);
      console.log('Datos obtenidos desde Google Sheets:', data);
    } catch (error) {
      console.error('Error al consultar Google Sheets:', error.message);
      return res.status(500).json({ error: 'Error al consultar datos de Google Sheets' });
    }

    const resultados = {};
    ranges.forEach((range, index) => {
      const rows = data[range];
      const solicitudData = rows.find((row) => row[0] === solicitudId);

      if (solicitudData) {
        const fields =
          index === 0
            ? ['id_solicitud', 'introduccion', 'objetivo_general', 'objetivos_especificos', 'justificacion', 'descripcion', 'alcance', 'metodologia', 'dirigido_a', 'programa_contenidos', 'duracion', 'certificacion', 'recursos']
            : ['id_solicitud', 'fecha_solicitud', 'nombre_actividad', 'nombre_solicitante', 'dependencia_tipo', 'nombre_escuela', 'nombre_departamento', 'nombre_seccion', 'nombre_dependencia', 'tipo', 'otro_tipo', 'modalidad', 'horas_trabajo_presencial', 'horas_sincronicas', 'total_horas', 'programCont', 'dirigidoa', 'creditos', 'cupo_min', 'cupo_max', 'nombre_coordinador', 'correo_coordinador', 'tel_coordinador', 'perfil_competencia', 'formas_evaluacion', 'certificado_solicitado', 'calificacion_minima', 'razon_no_certificado', 'valor_inscripcion', 'becas_convenio', 'becas_estudiantes', 'becas_docentes', 'becas_egresados', 'becas_funcionarios', 'becas_otros', 'becas_total', 'periodicidad_oferta', 'fechas_actividad', 'organizacion_actividad'];

        resultados[range.split('!')[0]] = fields.reduce((acc, field, idx) => {
          acc[field] = solicitudData[idx] || '';
          return acc;
        }, {});
      }
    });

    if (!Object.keys(resultados).length) {
      console.error('Error: No se encontraron datos para la solicitud:', solicitudId);
      return res.status(404).json({ error: 'No se encontraron datos para esta solicitud' });
    }

    const form1TemplateId = '1WiNfcR2_hRcvcNFohFyh0BPzLek9o9f0';
    const form2TemplateId = '1XZDXyMf4TC9PthBal0LPrgLMawHGeFM3';
    const folderId = '12bxb0XEArXMLvc7gX2ndqJVqS_sTiiUE';

    const replaceMarkers = async (templateId, data, fileName) => {
      try {
        const copiedFile = await drive.files.copy({
          fileId: templateId,
          requestBody: {
            name: fileName,
            parents: [folderId],
          },
        });

        const fileId = copiedFile.data.id;

        await drive.permissions.create({
          fileId: fileId,
          requestBody: {
            role: 'reader',
            type: 'anyone',
          },
        });

        const sheets = google.sheets({ version: 'v4', auth: jwtClient });

        const sheetInfo = await sheets.spreadsheets.get({ spreadsheetId: fileId });
        const sheetNames = sheetInfo.data.sheets.map((sheet) => sheet.properties.title);

        for (const sheetName of sheetNames) {
          const range = `${sheetName}!A1:Z100`;
          const response = await sheets.spreadsheets.values.get({
            spreadsheetId: fileId,
            range,
          });

          const values = response.data.values || [];
          const updatedValues = values.map((row) =>
            row.map((cell) =>
              typeof cell === 'string'
                ? Object.keys(data).reduce(
                    (updatedCell, marker) => updatedCell.replace(`{{${marker}}}`, data[marker] || ''),
                    cell
                  )
                : cell
            )
          );

          await sheets.spreadsheets.values.update({
            spreadsheetId: fileId,
            range,
            valueInputOption: 'USER_ENTERED',
            resource: { values: updatedValues },
          });
        }

        return `https://drive.google.com/file/d/${fileId}/view`;
      } catch (error) {
        console.error('Error al reemplazar marcadores en el archivo:', error.message);
        throw new Error('Error al reemplazar marcadores en el archivo');
      }
    };

    let form1Link, form2Link;
    try {
      form1Link = await replaceMarkers(
        form1TemplateId,
        resultados['SOLICITUDES'],
        `Formulario1_Solicitud_${solicitudId}`
      );
      form2Link = await replaceMarkers(
        form2TemplateId,
        resultados['SOLICITUDES2'],
        `Formulario2_Solicitud_${solicitudId}`
      );
    } catch (error) {
      console.error('Error al generar los informes:', error.message);
      return res.status(500).json({ error: 'Error al generar los informes' });
    }

    if (!form1Link || !form2Link) {
      console.error('Error: No se generaron todos los informes');
      return res.status(500).json({ error: 'No se generaron todos los informes' });
    }

    res.status(200).json({
      message: 'Informes generados exitosamente',
      links: [form1Link, form2Link],
    });
  } catch (error) {
    console.error('Error al generar los informes:', error.message);
    res.status(500).json({ error: 'Error al generar los informes' });
  }
});


app.listen(PORT, () => {
  console.log(`Servidor de extensión escuchando en el puerto ${PORT}`);
});
