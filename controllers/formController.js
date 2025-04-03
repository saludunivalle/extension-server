const sheetsService = require('../services/sheetsService');
const driveService = require('../services/driveService');
const dateUtils = require('../utils/dateUtils');
const progressService = require('../services/progressStateService'); // Añadir esta línea

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
    let estadoFormularios = {};
  
    // Obtener los datos actuales de ETAPAS
    const client = sheetsService.getClient();
    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A:I'
    });
    const etapasRows = etapasResponse.data.values || [];

    // Buscar la fila que corresponde al id_solicitud
    const filaExistente = etapasRows.find(row => row[0] === id_solicitud.toString());

    // Inicializar con todos los formularios
    estadoFormularios = {
      "1": "En progreso",
      "2": "En progreso",
      "3": "En progreso",
      "4": "En progreso"
    };
    
    // Si hay datos existentes, sobrescribir con ellos
    if (filaExistente && filaExistente[8]) {
      try {
        const estadoExistente = JSON.parse(filaExistente[8]);
        Object.assign(estadoFormularios, estadoExistente);
      } catch (e) {
        console.error('Error al parsear estado_formularios en guardarProgreso:', e);
        // Mantener el estado inicial
      }
    }
    
    // Marcar formularios anteriores como completados
    for (let i = 1; i < parsedHoja; i++) {
      estadoFormularios[i.toString()] = "Completado";
    }

    // Actualizar estado del formulario actual
    estadoFormularios[parsedHoja] = (paso >= maxPasos[parsedHoja]) ? 'Completado' : 'En progreso';

    const estadoFormulariosJSON = JSON.stringify(estadoFormularios);

    const etapaActual = (paso === maxPasos[parsedHoja]) ? parsedHoja + 1 : parsedHoja;
    const etapaActualAjustada = etapaActual > 4 ? 4 : etapaActual;

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
        },
        {
          range: `ETAPAS!I${filaEtapas}`, 
          values: [[estadoFormulariosJSON]]
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

    // Actualizar el estado en la sesión
    req.session.progressState = {
      etapa_actual: etapaActualAjustada,
      paso: paso,
      estado: estadoGlobal,
      estado_formularios: estadoFormularios
    };

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
      range: `ETAPAS!A2:I`,
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
    console.error('Detalles del error:', error.message);
    console.error('Stack trace:', error.stack);
    res.status(500).json({ error: 'Error al obtener solicitudes activas', details: error.message });
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

/**
 * Obtiene los gastos asociados a una solicitud
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const getGastos = async (req, res) => {
  try {
    const { id_solicitud } = req.query;
    
    if (!id_solicitud) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere id_solicitud'
      });
    }
    
    console.log(`Obteniendo gastos para solicitud ${id_solicitud}`);
    
    // Obtener cliente sheets
    const client = sheetsService.getClient();
    
    // Buscar todos los gastos para esta solicitud
    const response = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'GASTOS!A2:F'
    });
    
    const rows = response.data.values || [];
    const solicitudGastos = rows.filter(row => row[1] === id_solicitud);
    
    console.log(`Se encontraron ${solicitudGastos.length} gastos para la solicitud ${id_solicitud}`);
    
    // Formatear los resultados
    const gastos = solicitudGastos.map(row => ({
      id_conceptos: row[0] || '',
      id_solicitud: row[1] || '',
      cantidad: parseFloat(row[2]) || 0,
      valor_unit: parseFloat(row[3]) || 0,
      valor_total: parseFloat(row[4]) || 0,
      concepto_padre: row[5] || ''
    }));
    
    res.status(200).json({
      success: true,
      data: gastos
    });
    
  } catch (error) {
    console.error('Error al obtener gastos:', error);
    res.status(500).json({
      success: false,
      error: 'Error al obtener gastos',
      details: error.message
    });
  }
};

/**
 * Actualiza el paso máximo para una solicitud
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const actualizarPasoMaximo = async (req, res) => {
  try {
    const { id_solicitud, etapa_actual, paso } = req.body;
    
    // Validación más estricta de parámetros
    if (!id_solicitud) {
      return res.status(400).json({
        success: false,
        error: 'El ID de solicitud es requerido'
      });
    }
    
    // Validar cada campo individualmente para mensajes de error más específicos
    if (etapa_actual === undefined || etapa_actual === null) {
      return res.status(400).json({
        success: false,
        error: 'El campo etapa_actual es requerido'
      });
    }
    
    if (paso === undefined || paso === null) {
      return res.status(400).json({
        success: false,
        error: 'El campo paso es requerido'
      });
    }
    
    // Convertir y validar tipos
    const parsedEtapa = parseInt(etapa_actual);
    const parsedPaso = parseInt(paso);
    
    if (isNaN(parsedEtapa)) {
      return res.status(400).json({
        success: false,
        error: 'etapa_actual debe ser un valor numérico'
      });
    }
    
    if (isNaN(parsedPaso)) {
      return res.status(400).json({
        success: false,
        error: 'paso debe ser un valor numérico'
      });
    }
    
    // Validación de rango
    if (parsedEtapa < 1 || parsedEtapa > 4) {
      return res.status(400).json({
        success: false,
        error: 'etapa_actual debe estar entre 1 y 4'
      });
    }
    
    if (parsedPaso < 1) {
      return res.status(400).json({
        success: false,
        error: 'paso debe ser mayor o igual a 1'
      });
    }
    
    // Obtener los datos actuales de ETAPAS
    const client = sheetsService.getClient();
    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A:I'
    });
    const etapasRows = etapasResponse.data.values || [];
    
    // Buscar la fila que corresponde al id_solicitud
    let filaEtapas = etapasRows.findIndex(row => row[0] === id_solicitud.toString());
    
    if (filaEtapas === -1) {
      return res.status(404).json({
        success: false,
        error: `No se encontró la solicitud con ID ${id_solicitud}`
      });
    }
    
    filaEtapas += 1; // Ajustar índice a 1-based para Google Sheets
    
    // Leer el estado actual de formularios
    let estadoFormularios = {};
    if (etapasRows[filaEtapas - 1][8]) {
      try {
        estadoFormularios = JSON.parse(etapasRows[filaEtapas - 1][8]);
      } catch (e) {
        // Si no es JSON válido, inicializar estructura
        estadoFormularios = {
          "1": "En progreso",
          "2": "En progreso",
          "3": "En progreso",
          "4": "En progreso"
        };
      }
    } else {
      // Inicializar todos los formularios en "En progreso"
      estadoFormularios = {
        "1": "En progreso",
        "2": "En progreso",
        "3": "En progreso",
        "4": "En progreso"
      };
    }
    
    // Actualizar estados según etapa_actual
    // Si el formulario es menor que etapa_actual, se considera completado
    for (let i = 1; i <= 4; i++) {
      if (i < parsedEtapa) {
        estadoFormularios[i.toString()] = "Completado";
      }
    }
    
    // Actualizar la hoja con la nueva información
    await client.spreadsheets.values.update({
      spreadsheetId: sheetsService.spreadsheetId,
      range: `ETAPAS!E${filaEtapas}:I${filaEtapas}`,
      valueInputOption: 'RAW',
      resource: {
        values: [[
          parsedEtapa,
          etapasRows[filaEtapas - 1][5] || 'En progreso', // Mantener el estado global
          etapasRows[filaEtapas - 1][6] || 'N/A', // Mantener el nombre_actividad
          parsedPaso,
          JSON.stringify(estadoFormularios)
        ]]
      }
    });

    // Actualizar el estado en Redis
    const progressData = {
      etapa_actual: parsedEtapa,
      paso: parsedPaso,
      estadoFormularios: estadoFormularios
    };
    await progressService.setProgress(id_solicitud, progressData);

    // Actualizar el estado en la sesión
    req.session.progressState = {
      etapa_actual: parsedEtapa,
      paso: parsedPaso,
      estadoFormularios: estadoFormularios
    };
    
    res.status(200).json({
      success: true,
      message: 'Paso máximo actualizado correctamente'
    });
  } catch (error) {
    // Log más detallado para depuración
    console.error('Error detallado en actualizarPasoMaximo:', {
      mensaje: error.message,
      stack: error.stack,
      cuerpoSolicitud: req.body
    });
    
    res.status(500).json({
      success: false,
      error: 'Error al actualizar paso máximo',
      details: error.message
    });
  }
};

/**
 * Valida si un usuario puede avanzar a un paso/etapa específico
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const validarProgresion = async (req, res) => {
  try {
    const { id_solicitud, etapa_destino, paso_destino } = req.body;

    // Validaciones de parámetros
    if (!id_solicitud || etapa_destino === undefined || paso_destino === undefined) {
      return res.status(400).json({
        success: false,
        error: 'Se requieren id_solicitud, etapa_destino y paso_destino'
      });
    }

    // Convertir parámetros a números
    const etapaDestino = parseInt(etapa_destino);
    const pasoDestino = parseInt(paso_destino);

    if (isNaN(etapaDestino) || isNaN(pasoDestino)) {
      return res.status(400).json({
        success: false,
        error: 'etapa_destino y paso_destino deben ser valores numéricos'
      });
    }

    // Validar rangos
    if (etapaDestino < 1 || etapaDestino > 4) {
      return res.status(400).json({
        success: false,
        error: 'etapa_destino debe estar entre 1 y 4'
      });
    }

    if (pasoDestino < 1) {
      return res.status(400).json({
        success: false,
        error: 'paso_destino debe ser mayor o igual a 1'
      });
    }

    // Obtener datos actuales de la solicitud
    const client = sheetsService.getClient();
    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A2:I'
    });

    const etapasRows = etapasResponse.data.values || [];
    const filaActual = etapasRows.find(row => row[0] === id_solicitud.toString());

    if (!filaActual) {
      return res.status(404).json({
        success: false,
        error: `No se encontró la solicitud con ID ${id_solicitud}`
      });
    }

    // Obtener la etapa actual y estado de formularios
    const etapaActual = parseInt(filaActual[4]) || 1;
    const pasoActual = parseInt(filaActual[7]) || 1;

    let estadoFormularios = {};
    if (filaActual[8]) {
      try {
        estadoFormularios = JSON.parse(filaActual[8]);
      } catch (e) {
        estadoFormularios = {
          "1": "En progreso",
          "2": "En progreso",
          "3": "En progreso",
          "4": "En progreso"
        };
      }
    } else {
      estadoFormularios = {
        "1": "En progreso",
        "2": "En progreso",
        "3": "En progreso",
        "4": "En progreso"
      };
    }

    // Definir pasos máximos para cada formulario
    const maxPasos = {
      1: 5,
      2: 3,
      3: 5,
      4: 5
    };

    // Lógica de validación SIMPLIFICADA
    let puedeAvanzar = true; // Por defecto permitimos la navegación
    let mensaje = '';

    // Verificar el estado de todos los formularios hasta el destino
    const formulariosIniciados = [];
    for (let i = 1; i <= 4; i++) {
      if (estadoFormularios[i.toString()] === 'Completado' || 
          estadoFormularios[i.toString()] === 'En progreso') {
        formulariosIniciados.push(i);
      }
    }

    // ÚNICA RESTRICCIÓN: No permitir saltar a formularios futuros no iniciados
    if (etapaDestino > etapaActual && 
        !formulariosIniciados.includes(etapaDestino)) {
      puedeAvanzar = false;
      mensaje = 'No puede acceder a formularios futuros sin completar los anteriores';
    } else {
      // Permitir navegar a cualquier formulario ya iniciado o completado
      puedeAvanzar = true;
      
      // Mensaje más descriptivo según el caso
      if (etapaDestino < etapaActual) {
        mensaje = 'Navegando a un formulario anterior';
      } else if (etapaDestino === etapaActual) {
        mensaje = 'Continuando en el formulario actual';
      } else {
        mensaje = 'Avanzando al siguiente formulario';
      }
      
      // Log para depuración
      console.log(`Permitiendo navegación: ID=${id_solicitud}, De=${etapaActual} a=${etapaDestino}, formulariosIniciados=${formulariosIniciados.join(',')}`);
    }

    res.status(200).json({
      success: true,
      puedeAvanzar,
      mensaje,
      estado: {
        etapaActual,
        pasoActual,
        estadoFormularios,
        etapaDestino,
        pasoDestino,
        formulariosIniciados
      }
    });

  } catch (error) {
    console.error('Error en validarProgresion:', {
      mensaje: error.message,
      stack: error.stack
    });

    res.status(500).json({
      success: false,
      error: 'Error al validar la progresión',
      details: error.message
    });
  }
};

/**
 * Actualiza de manera centralizada el progreso de una solicitud
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const actualizarProgresoGlobal = async (req, res) => {
  try {
    const { 
      id_solicitud, 
      etapa_actual, 
      paso_actual,
      estadoFormularios,
      estadoGlobal,
      actualizar_formularios_previos = true
    } = req.body;
    
    // Validación básica
    if (!id_solicitud || etapa_actual === undefined || paso_actual === undefined) {
      return res.status(400).json({
        success: false,
        error: 'Se requieren id_solicitud, etapa_actual y paso_actual'
      });
    }
    
    // Convertir a enteros
    const parsedEtapa = parseInt(etapa_actual);
    const parsedPaso = parseInt(paso_actual);
    
    if (isNaN(parsedEtapa) || isNaN(parsedPaso)) {
      return res.status(400).json({
        success: false,
        error: 'etapa_actual y paso_actual deben ser valores numéricos'
      });
    }
    
    // Validar rangos
    if (parsedEtapa < 1 || parsedEtapa > 4) {
      return res.status(400).json({
        success: false,
        error: 'etapa_actual debe estar entre 1 y 4'
      });
    }
    
    if (parsedPaso < 1) {
      return res.status(400).json({
        success: false,
        error: 'paso_actual debe ser mayor o igual a 1'
      });
    }
    
    // Pasos máximos por formulario
    const maxPasos = {
      1: 5,
      2: 3,
      3: 5,
      4: 5
    };
    
    // Obtener datos actuales
    const client = sheetsService.getClient();
    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A:I'
    });
    
    const etapasRows = etapasResponse.data.values || [];
    let filaEtapas = etapasRows.findIndex(row => row[0] === id_solicitud.toString());
    
    if (filaEtapas === -1) {
      return res.status(404).json({
        success: false,
        error: `No se encontró la solicitud con ID ${id_solicitud}`
      });
    }
    
    filaEtapas += 1; // Ajustar a 1-based para Google Sheets
    
    // Recuperar el estado de formularios existente o usar el proporcionado
    let nuevoEstadoFormularios = {};
    
    if (estadoFormularios) {
      // Usar el estado proporcionado
      nuevoEstadoFormularios = typeof estadoFormularios === 'string' 
        ? JSON.parse(estadoFormularios) 
        : estadoFormularios;
    } else {
      // Recuperar estado existente o crear uno nuevo
      if (etapasRows[filaEtapas - 1][8]) {
        try {
          nuevoEstadoFormularios = JSON.parse(etapasRows[filaEtapas - 1][8]);
        } catch (e) {
          nuevoEstadoFormularios = {
            "1": "En progreso",
            "2": "En progreso",
            "3": "En progreso",
            "4": "En progreso"
          };
        }
      } else {
        nuevoEstadoFormularios = {
          "1": "En progreso",
          "2": "En progreso",
          "3": "En progreso",
          "4": "En progreso"
        };
      }
      
      // Actualizar el estado del formulario actual
      nuevoEstadoFormularios[parsedEtapa.toString()] = 
        (parsedPaso >= maxPasos[parsedEtapa]) ? 'Completado' : 'En progreso';
      
      // Marcar formularios anteriores como completados si se solicita
      if (actualizar_formularios_previos) {
        for (let i = 1; i < parsedEtapa; i++) {
          nuevoEstadoFormularios[i.toString()] = 'Completado';
        }
      }
    }
    
    // Determinar el estado global
    const nuevoEstadoGlobal = estadoGlobal || 
      ((parsedEtapa === 4 && parsedPaso >= maxPasos[4]) ? 'Completado' : 'En progreso');
    
    // Actualizar la hoja
    await client.spreadsheets.values.update({
      spreadsheetId: sheetsService.spreadsheetId,
      range: `ETAPAS!E${filaEtapas}:I${filaEtapas}`,
      valueInputOption: 'RAW',
      resource: {
        values: [[
          parsedEtapa,
          nuevoEstadoGlobal,
          etapasRows[filaEtapas - 1][6] || 'N/A', // Mantener nombre_actividad
          parsedPaso,
          JSON.stringify(nuevoEstadoFormularios)
        ]]
      }
    });

    // Actualizar el estado en la sesión
    req.session.progressState = {
      etapa_actual: parsedEtapa,
      paso: parsedPaso,
      estadoFormularios: nuevoEstadoFormularios
    };
    
    res.status(200).json({
      success: true,
      message: 'Progreso actualizado correctamente',
      data: {
        etapa_actual: parsedEtapa,
        paso_actual: parsedPaso,
        estado_global: nuevoEstadoGlobal,
        estado_formularios: nuevoEstadoFormularios
      }
    });
    
  } catch (error) {
    console.error('Error en actualizarProgresoGlobal:', {
      mensaje: error.message,
      stack: error.stack,
      cuerpoSolicitud: req.body
    });
    
    res.status(500).json({
      success: false,
      error: 'Error al actualizar el progreso global',
      details: error.message
    });
  }
};

/**
 * Obtiene el último ID de una hoja de cálculo
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const getLastId = async (req, res) => {
  try {
    const { sheetName } = req.query;

    if (!sheetName) {
      return res.status(400).json({ error: 'El nombre de la hoja es requerido' });
    }

    const client = sheetsService.getClient();
    const range = `${sheetName}!A:A`; // Obtener todos los valores de la columna A
    const response = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range,
    });

    const values = response.data.values || [];
    const lastId = values.length > 0 ? values.length : 0; // El último ID es la cantidad de filas

    res.status(200).json({ lastId });
  } catch (error) {
    console.error('Error al obtener el último ID:', error);
    res.status(500).json({ error: 'Error al obtener el último ID', details: error.message });
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

/**
 * Guarda datos específicos del Formulario 2 Paso 2 en SOLICITUDES2
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const guardarForm2Paso2 = async (req, res) => {
  try {
    // Extraer los datos relevantes
    const { id_solicitud, formData } = req.body;

    if (!id_solicitud) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere id_solicitud'
      });
    }

    // Obtener cliente sheets
    const client = sheetsService.getClient();
    
    console.log('👉 DATOS RECIBIDOS DEL FRONTEND:', formData);
    console.log('🔍 Verificando campos críticos:');
    console.log('- nombre_actividad:', formData.nombre_actividad || 'NO ENCONTRADO');
    console.log('- fecha_solicitud:', formData.fecha_solicitud || 'NO ENCONTRADO');
    
    // SOLUCIÓN: Verificar si faltan campos críticos y obtenerlos de SOLICITUDES
    if (!formData.nombre_actividad || !formData.fecha_solicitud) {
      console.log("⚠️ ALERTA: Faltan campos críticos, intentando recuperarlos de SOLICITUDES");
      
      try {
        const solicitudesResponse = await client.spreadsheets.values.get({
          spreadsheetId: sheetsService.spreadsheetId,
          range: 'SOLICITUDES!A2:D100'
        });
        
        const solicitudesRows = solicitudesResponse.data.values || [];
        const solicitudRow = solicitudesRows.find(row => row[0] === id_solicitud);
        
        if (solicitudRow) {
          console.log("✅ Datos encontrados en SOLICITUDES:", solicitudRow);
          
          // Actualizar formData con los campos de SOLICITUDES
          if (!formData.nombre_actividad && solicitudRow[2]) {
            formData.nombre_actividad = solicitudRow[2];
            console.log(`Campo nombre_actividad recuperado: ${formData.nombre_actividad}`);
          }
          
          if (!formData.fecha_solicitud && solicitudRow[1]) {
            formData.fecha_solicitud = solicitudRow[1];
            console.log(`Campo fecha_solicitud recuperado: ${formData.fecha_solicitud}`);
          }
        } else {
          console.log("❌ No se encontró la solicitud en SOLICITUDES, generando valores por defecto");
          if (!formData.nombre_actividad) {
            formData.nombre_actividad = `Actividad para solicitud ${id_solicitud}`;
          }
          if (!formData.fecha_solicitud) {
            const hoy = new Date();
            formData.fecha_solicitud = `${hoy.getDate()}/${hoy.getMonth()+1}/${hoy.getFullYear()}`;
          }
        }
      } catch (error) {
        console.error("❌ Error al buscar datos en SOLICITUDES:", error);
        // Generar valores por defecto en caso de error
        if (!formData.nombre_actividad) {
          formData.nombre_actividad = `Actividad para solicitud ${id_solicitud}`;
        }
        if (!formData.fecha_solicitud) {
          const hoy = new Date();
          formData.fecha_solicitud = `${hoy.getDate()}/${hoy.getMonth()+1}/${hoy.getFullYear()}`;
        }
      }
    }
    
    // Encontrar o crear fila para la solicitud
    const fila = await sheetsService.findOrCreateRequestRow('SOLICITUDES2', id_solicitud);
    
    // Campos específicos para el paso 2 del formulario 2 - NO MODIFICAR ESTE ORDEN
    const campos = [
      'nombre_actividad',  // Columna B
      'fecha_solicitud',   // Columna C 
      'ingresos_cantidad', // Columna D
      'ingresos_vr_unit',  // Columna E
      'total_ingresos',    // Columna F
      'subtotal_gastos',   // Columna G
      'imprevistos_3%',    // Columna H
      'total_gastos_imprevistos',       // Columna I
      'fondo_comun_porcentaje',         // Columna J
      'facultadad_instituto_porcentaje', // Columna K
      'escuela_departamento_porcentaje', // Columna L
      'total_recursos'                  // Columna M
    ];
    
    // Preparar valores a guardar (extrayéndolos de formData)
    const valores = campos.map(campo => {
      let valor = formData[campo] || '';
      if (campo === 'nombre_actividad' && !valor) {
        valor = `Actividad para solicitud ${id_solicitud}`;
      }
      if (campo === 'fecha_solicitud' && !valor) {
        const hoy = new Date();
        valor = `${hoy.getDate()}/${hoy.getMonth()+1}/${hoy.getFullYear()}`;
      }
      console.log(`Campo: ${campo}, Valor: ${valor}`);
      return valor;
    });
    
    console.log(`Guardando datos en SOLICITUDES2 para solicitud ${id_solicitud}, fila ${fila}`);
    console.log('Valores a guardar:', valores);
    
    // VERIFICACIÓN: Confirmar que los campos nombre_actividad y fecha_solicitud están presentes
    if (!valores[0]) {
      console.error("❌ CRÍTICO: nombre_actividad aún falta después de intentar recuperarlo");
    }
    if (!valores[1]) {
      console.error("❌ CRÍTICO: fecha_solicitud aún falta después de intentar recuperarlo");
    }
    
    // Actualizar datos en Google Sheets - REVISAR QUE EL RANGO SEA CORRECTO
    await client.spreadsheets.values.update({
      spreadsheetId: sheetsService.spreadsheetId,
      range: `SOLICITUDES2!B${fila}:M${fila}`,
      valueInputOption: 'USER_ENTERED', // USER_ENTERED para mejor formato
      resource: {
        values: [valores]
      }
    });
    
    // VERIFICACIÓN ADICIONAL: Leer de vuelta los datos guardados para confirmar
    const verificacionResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: `SOLICITUDES2!B${fila}:C${fila}`
    });
    const datosSalvados = verificacionResponse.data.values?.[0] || [];
    console.log("✅ VERIFICACIÓN DE DATOS GUARDADOS:");
    console.log(`- nombre_actividad guardado: ${datosSalvados[0] || 'NO GUARDADO'}`);
    console.log(`- fecha_solicitud guardado: ${datosSalvados[1] || 'NO GUARDADO'}`);
    
    console.log('✅ Operación completa: Datos del formulario 2 paso 2 guardados');
    
    res.status(200).json({ 
      success: true, 
      message: 'Datos del formulario 2 paso 2 guardados correctamente',
      datosSalvados: {
        nombre_actividad: datosSalvados[0] || '',
        fecha_solicitud: datosSalvados[1] || '',
      }
    });
    
  } catch (error) {
    console.error('❌ Error en guardarForm2Paso2:', error);
    res.status(500).json({
      success: false,
      error: 'Error al guardar los datos del formulario 2',
      details: error.message
    });
  }
};

module.exports = {
  guardarProgreso,
  createNewRequest,
  getRequests,
  getActiveRequests,
  getCompletedRequests,
  getFormDataForm2,
  guardarGastos,
  getGastos,
  actualizarPasoMaximo,  
  validarProgresion,  
  actualizarProgresoGlobal,
  getLastId,
  guardarForm2Paso2
};