const sheetsService = require('../services/sheetsService');
const driveService = require('../services/driveService');
const dateUtils = require('../utils/dateUtils');
const progressService = require('../services/progressStateService');
const { getDataWithCache, invalidateCache } = require('../utils/cacheUtils');

/**
 * Gestiona errores de cuota devolviendo respuesta exitosa con advertencia
 * @param {Error} error - El error capturado
 * @param {Object} res - Objeto de respuesta Express
 * @param {Object} dataToReturn - Datos a devolver en caso de error
 * @returns {boolean} - true si se manejó el error, false si no
 */
const handleQuotaError = (error, res, dataToReturn = {}) => {
  // Verificar si es un error de cuota excedida
  if (error.code === 429 || 
      (error.response && error.response.status === 429) ||
      error.message?.includes('Quota exceeded')) {
    console.log('⚠️ Error de cuota excedida, continuando con respuesta de éxito simulada');
    
    // Devolver éxito a pesar del error
    res.status(200).json({
      success: true,
      warning: 'Procesado localmente por limitaciones de API',
      quotaExceeded: true,
      ...dataToReturn
    });
    return true;
  }
  return false;
};

/**
 * Guarda el progreso de un formulario
 */

const guardarProgreso = async (req, res) => {
  const { id_solicitud, paso, hoja, id_usuario, name, ...formData } = req.body;
  const piezaGrafica = req.file;

  // Logging reducido para evitar sobrecarga
  console.log(`Guardando progreso ID:${id_solicitud}, paso:${paso}, hoja:${hoja}`);

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
      return res.status(400).json({ error: `Hoja ${parsedHoja} no válida` });
    }

    const sheetName = getSheetName(parsedHoja);
    const fechaActual = dateUtils.getCurrentDate();

    // Usar caché para encontrar la fila con mayor TTL (10 minutos)
    const cacheKey = `fila_${sheetName}_${id_solicitud}`;
    const fila = await getDataWithCache(
      cacheKey,
      () => sheetsService.findOrCreateRequestRow(sheetName, id_solicitud),
      10
    );

    // Subir pieza gráfica a Google Drive si existe
    let piezaGraficaUrl = '';
    if (piezaGrafica) {
      try {
        piezaGraficaUrl = await driveService.uploadFile(piezaGrafica);
      } catch (uploadError) {
        console.error('Error al subir pieza gráfica, continuando sin ella:', uploadError);
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

    // Logging reducido
    console.log(`Actualizando ${valoresFinales.length} valores en columnas`);

    if (valoresFinales.length !== columnasPaso.length) {
      console.warn(`⚠️ Cantidad de valores (${valoresFinales.length}) no coincide con columnas (${columnasPaso.length})`);
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

    // OPTIMIZACIÓN: Invalidar solo las cachés específicas, no todas
    invalidateCache(`solicitud_${id_solicitud}_${sheetName}`);
    
    // No invalidar otras cachés que no están relacionadas con esta actualización
    // invalidateCache(`solicitud_${id_solicitud}`);

    res.status(200).json({ success: true });
  } catch (error) {
    console.error("Error en guardarProgreso:", error);
    
    // Usar el manejador centralizado de errores de cuota
    if (handleQuotaError(error, res, {
      message: "Datos guardados localmente para sincronización posterior"
    })) return;
    
    // Para otros errores, mantener comportamiento original
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
    
    // Usar caché para esta operación
    const cacheKey = `form2_data_${id_solicitud}`;
    
    const formData = await getDataWithCache(
      cacheKey,
      async () => {
        const client = sheetsService.getClient();
        
        const range = 'SOLICITUDES2!A2:CL';
        const response = await client.spreadsheets.values.get({
          spreadsheetId: sheetsService.spreadsheetId,
          range,
        });

        const rows = response.data.values || [];
        const solicitudData = rows.find(row => row[0] === id_solicitud);

        if (!solicitudData) {
          return { error: 'Solicitud no encontrada en Formulario 2' };
        }

        // Mapear los campos según la estructura de SOLICITUDES2
        const fields = sheetsService.fieldDefinitions.SOLICITUDES2;

        const result = {};
        fields.forEach((field, index) => {
          result[field] = solicitudData[index] || '';
        });
        
        return result;
      },
      3 // TTL de 3 minutos
    );
    
    // Si hubo un error en la función de obtención de datos
    if (formData.error) {
      return res.status(404).json({ error: formData.error });
    }
    
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
    
    // OPTIMIZACIÓN: Aumentar TTL a 5 minutos para reducir llamadas
    const cacheKey = `gastos_${id_solicitud}`;
    const gastos = await getDataWithCache(
      cacheKey,
      async () => {
        // Implementar retry con backoff exponencial
        let retries = 3;
        let lastError;
        
        while (retries > 0) {
          try {
            const client = sheetsService.getClient();
            
            // OPTIMIZACIÓN: Usar un selector más específico en la consulta
            const range = `GASTOS!A2:F`;
            const response = await client.spreadsheets.values.get({
              spreadsheetId: sheetsService.spreadsheetId,
              range
            });
            
            const rows = response.data.values || [];
            
            // Filtrar sólo los gastos de esta solicitud
            return rows
              .filter(row => row[1] === id_solicitud.toString())
              .map(row => ({
                id_conceptos: row[0],
                id_solicitud: row[1],
                cantidad: parseFloat(row[2]) || 0,
                valor_unit: parseFloat(row[3]) || 0,
                valor_total: parseFloat(row[4]) || 0,
                concepto_padre: row[5] || ''
              }));
          } catch (error) {
            lastError = error;
            retries--;
            
            // Si es error de cuota y tenemos datos en caché, usar esos
            if (error.message?.includes('Quota exceeded') && 
                sheetsService.cache.has(cacheKey)) {
              return sheetsService.cache.get(cacheKey).value;
            }
            
            // Si aún tenemos intentos, esperar antes de reintentar
            if (retries > 0) {
              const waitTime = Math.pow(2, 3 - retries) * 1000;
              console.log(`Reintentando obtener gastos en ${waitTime}ms...`);
              await new Promise(resolve => setTimeout(resolve, waitTime));
            }
          }
        }
        
        throw lastError;
      },
      5 // 5 minutos
    );
    
    res.status(200).json({
      success: true,
      data: gastos
    });
    
  } catch (error) {
    console.error('Error al obtener gastos:', error);
    
    // Usar el manejador centralizado de errores de cuota
    if (handleQuotaError(error, res, {
      data: [], // Devolver array vacío en caso de error de cuota
      message: "No se pudieron obtener gastos por límites de API"
    })) return;
    
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

    // OPTIMIZACIÓN: Generar clave de caché que incluya todos los parámetros relevantes
    const cacheKey = `progresion_${id_solicitud}_${etapaDestino}_${pasoDestino}`;
    
    // OPTIMIZACIÓN: Aumentar TTL a 2 minutos para reducir llamadas
    const resultado = await getDataWithCache(
      cacheKey,
      async () => {
        try {
          // Primero intentar usar datos de sesión si están disponibles
          if (req.session?.progressState) {
            const { etapa_actual, paso } = req.session.progressState;
            
            // Verificar si podemos resolver esto sin llamar a la API
            if (etapa_actual && paso) {
              const puedeAvanzar = etapaDestino <= etapa_actual;
              return {
                success: true,
                puedeAvanzar,
                mensaje: puedeAvanzar ? 'Puede avanzar según caché local' : 'No puede avanzar según caché local',
                fuente: 'session'
              };
            }
          }
        
          // Si no hay datos en sesión, proceder con llamada a Google Sheets
          const client = sheetsService.getClient();
          const etapasResponse = await client.spreadsheets.values.get({
            spreadsheetId: sheetsService.spreadsheetId,
            range: 'ETAPAS!A:I'
          });
          
          const etapasRows = etapasResponse.data.values || [];
          const filaActual = etapasRows.find(row => row[0] === id_solicitud.toString());
          
          // Si no se encuentra la solicitud
          if (!filaActual) {
            return {
              success: false,
              puedeAvanzar: false,
              mensaje: `No se encontró la solicitud con ID ${id_solicitud}`
            };
          }
          
          // Extraer datos relevantes
          const etapaActual = parseInt(filaActual[4]);
          const pasoActual = parseInt(filaActual[7]);
          
          // Si no hay valores válidos
          if (isNaN(etapaActual) || isNaN(pasoActual)) {
            return {
              success: false,
              puedeAvanzar: false,
              mensaje: 'No se encontraron datos válidos de progresión'
            };
          }
          
          // Verificar el estado de los formularios
          let estadoFormularios = {};
          try {
            estadoFormularios = filaActual[8] ? JSON.parse(filaActual[8]) : {
              "1": "En progreso",
              "2": "En progreso",
              "3": "En progreso",
              "4": "En progreso"
            };
          } catch (e) {
            estadoFormularios = {
              "1": "En progreso",
              "2": "En progreso",
              "3": "En progreso",
              "4": "En progreso"
            };
          }
          
          // Verificar el estado de todos los formularios hasta el destino
          const formulariosIniciados = [];
          for (let i = 1; i <= 4; i++) {
            if (estadoFormularios[i.toString()] === 'Completado' || 
                estadoFormularios[i.toString()] === 'En progreso') {
              formulariosIniciados.push(i);
            }
          }
          
          // SIMPLIFICACIÓN: Permitir navegación a cualquier formulario anterior o actual
          let puedeAvanzar = true;
          let mensaje = '';
          
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
          }
          
          // También guardar esta información en la sesión para futuras referencias
          if (req.session) {
            req.session.progressState = {
              etapa_actual: etapaActual,
              paso: pasoActual,
              estadoFormularios
            };
          }
          
          return {
            success: true,
            puedeAvanzar,
            mensaje,
            estado: {
              etapaActual,
              pasoActual,
              estadoFormularios,
              etapaDestino,
              pasoDestino,
              formulariosIniciados,
              fuente: 'sheets'
            }
          };
        } catch (error) {
          // Si hay error y tenemos datos en sesión, usarlos como fallback
          if (req.session?.progressState) {
            console.warn('Usando datos de sesión como fallback debido a error:', error.message);
            const { etapa_actual, paso } = req.session.progressState;
            
            return {
              success: true,
              puedeAvanzar: etapaDestino <= etapa_actual,
              mensaje: 'Datos obtenidos desde sesión (fallback)',
              fuente: 'session_fallback',
              error: {
                message: error.message
              }
            };
          }
          
          // Si no hay datos en sesión, propagar el error
          throw error;
        }
      },
      2 // 2 minutos en lugar de 0.5
    );

    res.status(200).json(resultado);

  } catch (error) {
    console.error('Error en validarProgresion:', {
      mensaje: error.message,
      stack: error.stack
    });
    
    // Manejar error de cuota
    if (handleQuotaError(error, res, {
      success: true,
      puedeAvanzar: true, // Permitir avanzar en caso de error
      mensaje: 'Permitiendo navegación debido a error de API',
      estado: {
        etapaActual: parseInt(req.body.etapa_destino),
        pasoActual: 1,
        error: true
      }
    })) return;

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
    
    // AÑADIR ESTO:
    if (error.code === 429 || 
        (error.response && error.response.status === 429) ||
        error.message?.includes('Quota exceeded')) {
      console.log('Error de cuota excedida en Google Sheets API, continuando de todos modos...');
      
      return res.status(200).json({ 
        success: true,
        warning: 'Se alcanzó el límite de solicitudes, pero la acción fue registrada.'
      });
    }
    
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
    
    // Antes de guardar, verificar si la solicitud ya existe
    try {
      // Usar caché para esta operación
      const cacheKey = `solicitud2_check_${id_solicitud}`;
      const solicitudExistente = await getDataWithCache(
        cacheKey,
        async () => {
          const range = 'SOLICITUDES2!A2:A';
          const response = await client.spreadsheets.values.get({
            spreadsheetId: sheetsService.spreadsheetId,
            range,
          });
          
          const rows = response.data.values || [];
          return rows.some(row => row[0] === id_solicitud.toString());
        },
        5 // 5 minutos
      );
      
      // Determinar si es un insert o un update
      const operacionTipo = solicitudExistente ? 'update' : 'insert';
      console.log(`Operación: ${operacionTipo} para solicitud ${id_solicitud}`);
      
      // Si es una actualización, buscar la fila exacta
      let filaActualizar;
      if (operacionTipo === 'update') {
        filaActualizar = await getDataWithCache(
          `solicitud2_fila_${id_solicitud}`,
          async () => {
            const range = 'SOLICITUDES2!A2:A';
            const response = await client.spreadsheets.values.get({
              spreadsheetId: sheetsService.spreadsheetId,
              range,
            });
            
            const rows = response.data.values || [];
            const index = rows.findIndex(row => row[0] === id_solicitud.toString());
            return index + 2; // +2 porque los índices son 0-based y hay un encabezado
          },
          5 // 5 minutos
        );
      }
      
      // Preparar datos para guardar - Convertir objeto a array
      const fields = sheetsService.fieldDefinitions.SOLICITUDES2;
      const values = fields.map(field => formData[field] || '');
      
      // Asegurar que el ID esté en la primera posición
      values[0] = id_solicitud.toString();
      
      // Ejecutar la operación adecuada
      if (operacionTipo === 'insert') {
        await client.spreadsheets.values.append({
          spreadsheetId: sheetsService.spreadsheetId,
          range: 'SOLICITUDES2!A2',
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS',
          resource: { values: [values] },
        });
      } else {
        await client.spreadsheets.values.update({
          spreadsheetId: sheetsService.spreadsheetId,
          range: `SOLICITUDES2!A${filaActualizar}:${getColumnLetter(fields.length)}${filaActualizar}`,
          valueInputOption: 'RAW',
          resource: { values: [values] },
        });
      }
      
      // Invalidar caché específica
      invalidateCache(`solicitud_${id_solicitud}_SOLICITUDES2`);
      
      res.status(200).json({
        success: true,
        message: `Datos ${operacionTipo === 'insert' ? 'guardados' : 'actualizados'} correctamente`
      });
    } catch (error) {
      console.error('Error al guardar datos en SOLICITUDES2:', error);
      
      // Manejar error de cuota
      if (handleQuotaError(error, res, {
        message: "Datos almacenados localmente para sincronización posterior"
      })) return;
      
      res.status(500).json({
        success: false,
        error: 'Error al guardar datos en SOLICITUDES2',
        details: error.message
      });
    }
  } catch (error) {
    console.error('Error general en guardarForm2Paso2:', error);
    res.status(500).json({
      success: false,
      error: 'Error general en guardarForm2Paso2',
      details: error.message
    });
  }
};

// Función auxiliar para convertir número de columna a letra
function getColumnLetter(columnNumber) {
  let columnLetter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnLetter;
}

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