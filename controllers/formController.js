const sheetsService = require('../services/sheetsService');
const driveService = require('../services/driveService');
const dateUtils = require('../utils/dateUtils');

/**
 * Guarda el progreso de un formulario
*/

const guardarProgreso = async (req, res) => {
  const { id_solicitud, paso, hoja, id_usuario, name, ...formData } = req.body;
  const piezaGrafica = req.file;

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
    // Obtener columnas del servicio de sheets
    const columnas = sheetsService.columnMappings[getSheetName(parsedHoja)];
    if (!columnas) {
      return res.status(400).json({ error: 'Hoja no válida' });
    }

    const sheetName = getSheetName(parsedHoja);
    const fechaActual = dateUtils.getCurrentDate();

    // Encontrar o crear fila para la solicitud
    const fila = await sheetsService.findOrCreateRequestRow(sheetName, id_solicitud);

    // Subir pieza gráfica a Google Drive si existe
    let piezaGraficaUrl = '';
    if (piezaGrafica) {
      try {
        piezaGraficaUrl = await driveService.uploadFile(piezaGrafica);
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

    // Actualizar datos en Google Sheets usando sheetsService
    await sheetsService.updateRequestProgress({
      sheetName,
      rowIndex: fila,
      startColumn: columnaInicial,
      endColumn: columnaFinal,
      values: valoresFinales
    });

    // Actualizar hoja de ETAPAS
    const maxPasos = {
      1: 5,
      2: 3,
      3: 5,
      4: 5
    };

    const estadoGlobal = (parsedHoja === 4 && paso === maxPasos[3]) ? 'Completado' : 'En progreso';
    const etapaActual = (paso === maxPasos[parsedHoja]) ? parsedHoja + 1 : parsedHoja;
    const etapaActualAjustada = etapaActual > 4 ? 4 : etapaActual;

    // Obtener los datos actuales de ETAPAS
    const client = sheetsService.getClient();
    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A:H'
    });
    const etapasRows = etapasResponse.data.values || [];

    // Buscar la fila que corresponde al id_solicitud
    let filaEtapas = etapasRows.findIndex(row => row[0] === id_solicitud.toString());

    if (filaEtapas === -1) {
      // Si no se encuentra, agregar una nueva fila
      filaEtapas = etapasRows.length + 1;
      await client.spreadsheets.values.append({
        spreadsheetId: sheetsService.spreadsheetId,
        range: `ETAPAS!A${filaEtapas}:H${filaEtapas}`,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: {
          values: [[
            id_solicitud,
            id_usuario,
            fechaActual,
            name,
            etapaActualAjustada,
            estadoGlobal,
            formData.nombre_actividad || 'N/A',
            paso
          ]]
        }
      });
    } else {
      filaEtapas += 1; // Ajustar índice a 1-based
      
      // Actualizar múltiples columnas usando batchUpdate
      const updateRequests = [
        {
          range: `ETAPAS!E${filaEtapas}`, // Columna E: etapa_actual
          values: [[etapaActualAjustada]]
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

      await client.spreadsheets.values.batchUpdate({
        spreadsheetId: sheetsService.spreadsheetId,
        resource: {
          valueInputOption: 'RAW',
          data: updateRequests
        }
      });
    }

    res.status(200).json({ success: true });
  } catch (error) {
    console.error("Error en guardarProgreso:", error);
    res.status(500).json({
      success: false,
      error: 'Error de conexión con Google Sheets',
      details: error.message
    });
  }
};

/**
 * Crea una nueva solicitud
*/
const createNewRequest = async (req, res) => {
  try {
    const { id_solicitud, fecha_solicitud, nombre_actividad, nombre_solicitante, dependencia_tipo, nombre_dependencia } = req.body;
    
    const client = sheetsService.getClient();
    const range = 'SOLICITUDES!A2:F2';
    const values = [[id_solicitud, fecha_solicitud, nombre_actividad, nombre_solicitante, dependencia_tipo, nombre_dependencia]];

    await client.spreadsheets.values.append({
      spreadsheetId: sheetsService.spreadsheetId,
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
};

/**
 * Obtiene todas las solicitudes de un usuario
*/
const getRequests = async (req, res) => {
  try {
    const { userId } = req.query;
    const client = sheetsService.getClient();

    const activeResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
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
};

/**
 * Obtiene solicitudes activas de un usuario
*/
const getActiveRequests = async (req, res) => {
  try {
    const { userId } = req.query;
    console.log('Obteniendo solicitudes activas para el usuario:', userId);
    const client = sheetsService.getClient();

    // Obtener datos de ETAPAS
    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
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
    const solicitudesResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: `SOLICITUDES!A2:D`, // Columna nombre_actividad
    });

    const solicitudesRows = solicitudesResponse.data.values;

    // Combinar datos de ETAPAS y SOLICITUDES
    const combinedRequests = activeRequests.map((request) => {
      const solicitud = solicitudesRows?.find((row) => row[0] === request.idSolicitud);
      return {
        ...request,
        nombre_actividad: solicitud ? solicitud[2] : 'Sin nombre', // Índice de nombre_actividad
      };
    });

    console.log('Solicitudes activas combinadas:', combinedRequests);
    res.status(200).json(combinedRequests);
  } catch (error) {
    console.error('Error al obtener solicitudes activas:', error);
    res.status(500).json({ error: 'Error al obtener solicitudes activas' });
  }
};

/**
 * Obtiene solicitudes completadas de un usuario
*/
const getCompletedRequests = async (req, res) => {
  try {
    const { userId } = req.query;
    console.log('Obteniendo solicitudes completadas para el usuario:', userId);
    const client = sheetsService.getClient();

    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: `ETAPAS!A2:I`,
    });

    const rows = etapasResponse.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron solicitudes completadas' });
    }

    const completedRequests = rows.filter((row) => row[1] === userId && row[5] === 'Completado')
      .map((row) => ({
        idSolicitud: row[0],
        formulario: parseInt(row[4]),
        etapa_actual: parseInt(row[4]),
        paso: parseInt(row[7]),
        nombre_actividad: row[6]
      }));

    console.log('Solicitudes completadas:', completedRequests);
    res.status(200).json(completedRequests);
  } catch (error) {
    console.error('Error al obtener solicitudes completadas:', error);
    res.status(500).json({ error: 'Error al obtener solicitudes completadas' });
  }
};

/**
 * Obtiene datos específicos del formulario 2
*/
const getFormDataForm2 = async (req, res) => {
  try {
    const { id_solicitud } = req.query;
    const client = sheetsService.getClient();
    
    const range = 'SOLICITUDES2!A2:CL';
    const response = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range,
    });

    const rows = response.data.values || [];
    const solicitudData = rows.find(row => row[0] === id_solicitud);

    if (!solicitudData) {
      return res.status(404).json({ error: 'Solicitud no encontrada en Formulario 2' });
    }

    // Mapear los campos según la estructura de SOLICITUDES2
    const fields = sheetsService.fieldDefinitions.SOLICITUDES2;

    const formData = {};
    fields.forEach((field, index) => {
      formData[field] = solicitudData[index] || '';
    });

    res.status(200).json(formData);
  } catch (error) {
    console.error('Error al obtener datos del Formulario 2:', error);
    res.status(500).json({ error: 'Error interno del servidor' });
  }
};

/**
 * Guarda información de gastos
*/
const guardarGastos = async (req, res) => {
  try {
    const { id_solicitud, gastos } = req.body;

    if (!id_solicitud || !gastos?.length) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere id_solicitud y al menos un concepto'
      });
    }

    // Usar el servicio de sheetsService para guardar gastos
    const success = await sheetsService.saveGastos(id_solicitud, gastos);
    
    if (success) {
      res.status(200).json({ success: true });
    } else {
      res.status(400).json({ 
        success: false, 
        error: 'No se pudo guardar ningún gasto'
      });
    }
  } catch (error) {
    console.error("Error en guardarGastos:", error);
    res.status(500).json({
      success: false,
      error: 'Error de conexión con Google Sheets',
      details: error.message
    });
  }
};

// Función auxiliar para obtener el nombre de la hoja según el ID
function getSheetName(hojaId) {
  switch (hojaId) {
    case 1: return 'SOLICITUDES';
    case 2: return 'SOLICITUDES2';
    case 3: return 'SOLICITUDES3';
    case 4: return 'SOLICITUDES4';
    case 5: return 'GASTOS';
    default: return null;
  }
}

module.exports = {
  guardarProgreso,
  createNewRequest,
  getRequests,
  getActiveRequests,
  getCompletedRequests,
  getFormDataForm2,
  guardarGastos
};