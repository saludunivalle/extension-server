const sheetsService = require('../services/sheetsService');
const driveService = require('../services/driveService');
const dateUtils = require('../utils/dateUtils');
const progressService = require('../services/progressStateService'); // Añadir esta línea

const ESTADOS_GENERALES = Object.freeze({
  EN_PROCESO: 'En proceso',
  TERMINADO: 'Terminado',
  ENVIADO_REVISION: 'Enviado a revisión',
  APROBADO_PARCIAL: 'Aprobado parcialmente',
  APROBADO: 'Aprobado'
});

const ESTADOS_FORMULARIO = Object.freeze({
  EN_PROGRESO: 'En progreso',
  COMPLETADO: 'Completado',
  ENVIADO_REVISION: 'Enviado a revisión',
  APROBADO: 'Aprobado',
  REQUIERE_CORRECCIONES: 'Requiere correcciones'
});

// Sistema de caché para ETAPAS
const etapasCache = {
  datos: null,
  ultimaActualizacion: null,
  ttl: 60000, // 1 minuto (en lugar de 5)

  valido() {
    return this.datos && (Date.now() - this.ultimaActualizacion < this.ttl);
  },

  async refresh() {
    try {
      console.log('🔄 Actualizando caché de ETAPAS...');
      const client = sheetsService.getClient();
      const response = await client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'ETAPAS!A:K'
      });
      this.datos = response.data.values || [];
      this.ultimaActualizacion = Date.now();
      console.log(`✅ Caché de ETAPAS actualizada con ${this.datos.length} filas`);
      return true;
    } catch (error) {
      console.error('❌ Error al actualizar caché de ETAPAS:', error);
      return false;
    }
  }
};

// Función para obtener datos de ETAPAS usando caché
async function getEtapas(fuerzaRefresh = false) {
  if (fuerzaRefresh || !etapasCache.valido()) {
    await etapasCache.refresh();
  }
  return etapasCache.datos;
}

async function getUserRoleInfo(userId) {
  if (!userId) {
    return { role: '', isAdmin: false };
  }

  const client = sheetsService.getClient();
  const response = await client.spreadsheets.values.get({
    spreadsheetId: sheetsService.spreadsheetId,
    range: 'USUARIOS!A2:D'
  });

  const rows = response.data.values || [];
  const normalizedUserId = String(userId).trim();
  const matchingRows = rows.filter(
    (row) => String(row[0] || '').trim() === normalizedUserId
  );

  const preferredRow = matchingRows.find(
    (row) => String(row[3] || '').trim().toLowerCase() === 'admin'
  ) || matchingRows[0];

  const role = String(preferredRow?.[3] || '').trim().toLowerCase() === 'admin' ? 'admin' : '';

  return {
    role,
    isAdmin: role.toLowerCase() === 'admin'
  };
}

function getDefaultEstadoFormularios() {
  return {
    '1': 'En progreso',
    '2': 'En progreso',
    '3': 'En progreso',
    '4': 'En progreso'
  };
}

function normalizeEstadoFormulario(estadoFormulario) {
  const value = String(estadoFormulario || '').trim().toLowerCase();
  if (value === 'aprobado') {
    return ESTADOS_FORMULARIO.APROBADO;
  }

  if (value === 'enviado a revisión' || value === 'enviado a revision') {
    return ESTADOS_FORMULARIO.ENVIADO_REVISION;
  }

  if (value === 'requiere correcciones') {
    return ESTADOS_FORMULARIO.REQUIERE_CORRECCIONES;
  }

  if (value === 'completado' || value === 'terminado') {
    return ESTADOS_FORMULARIO.COMPLETADO;
  }

  return ESTADOS_FORMULARIO.EN_PROGRESO;
}

function parseEstadoFormularios(rawEstadoFormularios) {
  let parsed = getDefaultEstadoFormularios();

  if (rawEstadoFormularios) {
    try {
      const candidate = typeof rawEstadoFormularios === 'string'
        ? JSON.parse(rawEstadoFormularios)
        : rawEstadoFormularios;
      if (candidate && typeof candidate === 'object') {
        parsed = {
          ...parsed,
          ...candidate
        };
      }
    } catch (error) {
      console.warn('No se pudo parsear estado_formularios, usando valores por defecto.');
    }
  }

  return parsed;
}

function getEstadoGeneralLegacy(estadoGeneral) {
  return estadoGeneral === ESTADOS_GENERALES.EN_PROCESO ? 'En progreso' : 'Completado';
}

function calculateEstadoGeneral(estadoFormularios, estadoGeneralActual = '') {
  const keys = ['1', '2', '3', '4'];
  const normalized = keys.map((key) => normalizeEstadoFormulario(estadoFormularios[key]));

  const allAprobados = normalized.every((estado) => estado === ESTADOS_FORMULARIO.APROBADO);
  if (allAprobados) {
    return ESTADOS_GENERALES.APROBADO;
  }

  const hasAprobados = normalized.some((estado) => estado === ESTADOS_FORMULARIO.APROBADO);
  if (hasAprobados) {
    return ESTADOS_GENERALES.APROBADO_PARCIAL;
  }

  const hasEnRevision = normalized.some((estado) => estado === ESTADOS_FORMULARIO.ENVIADO_REVISION);
  if (hasEnRevision) {
    return ESTADOS_GENERALES.ENVIADO_REVISION;
  }

  const allCompletados = normalized.every(
    (estado) => estado === ESTADOS_FORMULARIO.COMPLETADO
  );
  if (allCompletados) {
    return ESTADOS_GENERALES.TERMINADO;
  }

  const hasInProgress = normalized.some((estado) => estado === ESTADOS_FORMULARIO.EN_PROGRESO);
  if (hasInProgress) {
    return ESTADOS_GENERALES.EN_PROCESO;
  }

  if (estadoGeneralActual === ESTADOS_GENERALES.APROBADO_PARCIAL) {
    return ESTADOS_GENERALES.APROBADO_PARCIAL;
  }

  return ESTADOS_GENERALES.EN_PROCESO;
}

function parseComentariosPorFormulario(rawComentarios) {
  if (!rawComentarios) {
    return {};
  }

  try {
    const parsed = typeof rawComentarios === 'string'
      ? JSON.parse(rawComentarios)
      : rawComentarios;

    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
      return parsed;
    }
  } catch (error) {
    // Compatibilidad: valor antiguo de texto libre.
  }

  return {
    general: String(rawComentarios)
  };
}

function stringifyComentariosPorFormulario(comentarios) {
  if (!comentarios || Object.keys(comentarios).length === 0) {
    return '';
  }

  return JSON.stringify(comentarios);
}

function sanitizeFormularios(formularios) {
  if (!Array.isArray(formularios)) {
    return [];
  }

  const unique = [...new Set(formularios
    .map((f) => parseInt(f, 10))
    .filter((f) => Number.isInteger(f) && f >= 1 && f <= 4))];

  return unique.sort((a, b) => a - b);
}

function getFormulariosByEstado(estadoFormularios, allowedStates) {
  return Object.entries(estadoFormularios)
    .filter(([, estado]) => allowedStates.includes(normalizeEstadoFormulario(estado)))
    .map(([formulario]) => parseInt(formulario, 10));
}

function isFormularioRealizado(estadoFormulario) {
  const normalized = normalizeEstadoFormulario(estadoFormulario);
  return normalized !== ESTADOS_FORMULARIO.EN_PROGRESO;
}

function setEstadoFormulario(estadoFormularios, formularioId, estado) {
  estadoFormularios[String(formularioId)] = estado;
}

function getEstadoGeneralFromRow(row = []) {
  const rowEstadoGeneral = String(row[9] || '').trim();
  if (Object.values(ESTADOS_GENERALES).includes(rowEstadoGeneral)) {
    return rowEstadoGeneral;
  }

  const estadoFormularios = parseEstadoFormularios(row[8]);
  return calculateEstadoGeneral(estadoFormularios);
}

async function userIsAdmin(userId) {
  const { isAdmin } = await getUserRoleInfo(userId);
  return isAdmin;
}

function buildEtapaSummary(row) {
  const estadoFormularios = parseEstadoFormularios(row[8]);
  const comentariosPorFormulario = parseComentariosPorFormulario(row[10] || '');
  return {
    id_solicitud: row[0] || '',
    id_usuario: row[1] || '',
    fecha: row[2] || '',
    name: row[3] || '',
    etapa_actual: parseInt(row[4], 10) || 1,
    estado: row[5] || 'En progreso',
    nombre_actividad: row[6] || '',
    paso: parseInt(row[7], 10) || 1,
    estado_formularios: estadoFormularios,
    estado_general: getEstadoGeneralFromRow(row),
    comentarios: row[10] || '',
    comentarios_por_formulario: comentariosPorFormulario,
    formularios_en_revision: getFormulariosByEstado(estadoFormularios, [ESTADOS_FORMULARIO.ENVIADO_REVISION]),
    formularios_aprobados: getFormulariosByEstado(estadoFormularios, [ESTADOS_FORMULARIO.APROBADO]),
    formularios_con_correcciones: getFormulariosByEstado(estadoFormularios, [ESTADOS_FORMULARIO.REQUIERE_CORRECCIONES])
  };
}

/**
 * Guarda el progreso de un formulario
*/

const guardarProgreso = async (req, res) => {
  // Extraer nombre_actividad explícitamente para asegurar que se capture correctamente
  const { id_solicitud, paso, hoja, id_usuario, name, nombre_actividad, ...restFormData } = req.body;
  // Reconstruir formData incluyendo nombre_actividad explícitamente
  const formData = { nombre_actividad, ...restFormData };

  // Compatibilidad de nombres para el nuevo campo de programa
  // Frontend puede enviar tipo_programa/proceso mientras el modelo usa programa
  if (!formData.programa) {
    formData.programa = formData.tipo_programa || formData.proceso || '';
  }

  if (!formData.entradas_diseño) {
    formData.entradas_diseño = formData.entradas_diseño ||  formData.entrada_diseno || '';
  }

  // Compatibilidad de nombres para valor de inscripción
  if (!formData.valor_inscripcion) {
    formData.valor_inscripcion =
      formData.valorInscripcion ||
      formData.inscripcion ||
      formData.valor_inscripcion_individual ||
      formData.valorInscripcionIndividual ||
      '';
  }


  const piezaGrafica = req.file;

  console.log('Recibiendo datos para guardar progreso:', { id_solicitud, paso, hoja, id_usuario, name });
  console.log('Datos del formulario:', formData);
  console.log('Archivo adjunto:', req.file);

  const parsedPaso = parseInt(paso, 10);
  const parsedHoja = parseInt(hoja, 10);

  if (isNaN(parsedPaso) || isNaN(parsedHoja)) {
    console.error('Paso u Hoja no válida: no son números');
    return res.status(400).json({ error: 'Paso u Hoja no válida' });
  }

  const sheetName = getSheetName(parsedHoja);
  if (!sheetName) {
    return res.status(400).json({ error: 'Hoja no válida' });
  }

  try {
    // Obtener la definición del modelo para la hoja específica
    const model = sheetsService.models[sheetName];
    if (!model) {
      return res.status(400).json({ error: `Modelo no encontrado para la hoja ${sheetName}` });
    }

    // Obtener las columnas específicas para el paso actual
    const columnasPaso = model.columnMappings[parsedPaso];
    if (!columnasPaso || columnasPaso.length === 0) {
      console.warn(`No hay mapeo de columnas definido para la hoja ${sheetName}, paso ${parsedPaso}`);
      // Considerar si se debe devolver un error o simplemente no hacer nada
      return res.status(200).json({ success: true, message: 'No hay columnas definidas para este paso.' });
    }

    const allFields = model.fields;
    const columnaInicialLetra = columnasPaso[0];
    const columnaFinalLetra = columnasPaso[columnasPaso.length - 1];

    // Convertir letras de columna a índices (0-based)
    const colIndexToLetter = (index) => {
      let temp, letter = '';
      while (index >= 0) {
        temp = index % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        index = Math.floor(index / 26) - 1;
      }
      return letter;
    };
    const letterToColIndex = (letter) => {
      let index = 0, length = letter.length;
      for (let i = 0; i < length; i++) {
        index = index * 26 + letter.charCodeAt(i) - 64;
      }
      return index - 1;
    };

    const columnaInicialIndex = letterToColIndex(columnaInicialLetra);
    const columnaFinalIndex = letterToColIndex(columnaFinalLetra);

    // Crear un array de valores ordenado según las columnas definidas en allFields
    // Incluirá celdas vacías para las columnas no relevantes para este paso
    const valoresOrdenados = [];
    for (let i = columnaInicialIndex; i <= columnaFinalIndex; i++) {
      const fieldName = allFields[i]; // Obtener el nombre del campo esperado para esta columna
      // Buscar el valor correspondiente en formData
      const value = formData.hasOwnProperty(fieldName) ? (formData[fieldName] ?? '').toString() : ''; 
      valoresOrdenados.push(value);
    }
    
    // Manejar pieza gráfica si existe y si la columna está en el rango actual
    const piezaGraficaFieldName = 'pieza_grafica'; 
    const piezaGraficaColIndex = allFields.indexOf(piezaGraficaFieldName);
    let piezaGraficaUrl = '';
    if (piezaGrafica) {
      try {
        piezaGraficaUrl = await driveService.uploadFile(piezaGrafica);
        // Si la columna de pieza gráfica está dentro del rango de este paso, actualizar el valor
        if (piezaGraficaColIndex >= columnaInicialIndex && piezaGraficaColIndex <= columnaFinalIndex) {
          valoresOrdenados[piezaGraficaColIndex - columnaInicialIndex] = piezaGraficaUrl;
        }
      } catch (error) {
        console.error('Error al subir la pieza gráfica a Google Drive:', error);
        // Considerar si devolver un error o continuar sin la URL
        // return res.status(500).json({ error: 'Error al subir la pieza gráfica' });
      }
    }

    console.log(`Actualizando ${sheetName}, Paso ${parsedPaso}`);
    console.log(`  Rango de columnas: ${columnaInicialLetra} a ${columnaFinalLetra} (Índices ${columnaInicialIndex} a ${columnaFinalIndex})`);
    console.log('  Valores ordenados para enviar:', valoresOrdenados);

    // DEBUG: Log específico para paso 2 y metodología
    if (parsedPaso === 2) {
      console.log('🔍 DEBUG PASO 2:');
      console.log('  allFields completo:', allFields);
      console.log('  metodologia en posición:', allFields.indexOf('metodologia'));
      console.log('  formData recibido:', formData);
      console.log('  Mapeo detallado por índice:');
      for (let i = columnaInicialIndex; i <= columnaFinalIndex; i++) {
        const fieldName = allFields[i];
        const value = formData.hasOwnProperty(fieldName) ? formData[fieldName] : '[NO ENCONTRADO]';
        console.log(`    Índice ${i} -> Campo '${fieldName}' -> Valor: '${value}'`);
      }
    }

    // DEBUG: Log específico para paso 3 y cupos
    if (parsedPaso === 3) {
      console.log('🔍 DEBUG PASO 3:');
      console.log('  allFields completo:', allFields);
      console.log('  cupo_min en posición:', allFields.indexOf('cupo_min'));
      console.log('  cupo_max en posición:', allFields.indexOf('cupo_max'));
      console.log('  formData recibido:', formData);
      console.log('  Campos específicos:');
      console.log('    cupo_min:', formData.cupo_min);
      console.log('    cupo_max:', formData.cupo_max);
      console.log('  Mapeo detallado por índice:');
      for (let i = columnaInicialIndex; i <= columnaFinalIndex; i++) {
        const fieldName = allFields[i];
        const value = formData.hasOwnProperty(fieldName) ? formData[fieldName] : '[NO ENCONTRADO]';
        console.log(`    Índice ${i} -> Campo '${fieldName}' -> Valor: '${value}'`);
      }
    }

    // DEBUG: Log específico para paso 5 y campos AU/AV
    if (parsedPaso === 5) {
      console.log('🔍 DEBUG PASO 5:');
      console.log('  allFields completo:', allFields);
      console.log('  pieza_grafica en posición:', allFields.indexOf('pieza_grafica'));
      console.log('  observaciones_cambios en posición:', allFields.indexOf('observaciones_cambios'));
      console.log('  formData recibido:', formData);
      console.log('  Campos específicos:');
      console.log('    pieza_grafica:', formData.pieza_grafica);
      console.log('    observaciones_cambios:', formData.observaciones_cambios);
      console.log('  Mapeo detallado por índice:');
      for (let i = columnaInicialIndex; i <= columnaFinalIndex; i++) {
        const fieldName = allFields[i];
        const value = formData.hasOwnProperty(fieldName) ? formData[fieldName] : '[NO ENCONTRADO]';
        console.log(`    Índice ${i} -> Campo '${fieldName}' -> Valor: '${value}'`);
      }
    }

    // Encontrar o crear fila para la solicitud
    const rowInfo = await sheetsService.findOrCreateRequestRow(sheetName, id_solicitud);
    if (!rowInfo || !rowInfo.rowIndex) {
        throw new Error('No se pudo obtener el índice de la fila.');
    }

    // Garantizar que el id_solicitud siempre quede en la columna A de SOLICITUDES2
    if (sheetName === 'SOLICITUDES2') {
      await sheetsService.getClient().spreadsheets.values.update({
        spreadsheetId: sheetsService.spreadsheetId,
        range: `SOLICITUDES2!A${rowInfo.rowIndex}`,
        valueInputOption: 'RAW',
        resource: { values: [[String(id_solicitud)]] }
      });
    }

    console.log(`  Fila encontrada/creada: ${rowInfo.rowIndex}`);

    // Actualizar datos en Google Sheets usando sheetsService
    await sheetsService.updateRequestProgress({
      sheetName,
      rowIndex: rowInfo.rowIndex,
      startColumn: columnaInicialLetra, // Usar letra de columna inicial
      endColumn: columnaFinalLetra,     // Usar letra de columna final
      values: valoresOrdenados         // Usar el array ordenado
    });

    console.log(`✅ Progreso guardado para ${sheetName}, Solicitud ${id_solicitud}, Paso ${parsedPaso}`);

    // --- Lógica para actualizar ETAPAS ---
    const maxPasos = { 1: 5, 2: 3, 3: 6, 4: 5 };
    let estadoFormularios = getDefaultEstadoFormularios();

    const client = sheetsService.getClient();
    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A:K'
    });
    const etapasRows = etapasResponse.data.values || [];
    const filaExistenteIndex = etapasRows.findIndex(row => row[0] === id_solicitud.toString());
    const filaExistente = filaExistenteIndex !== -1 ? etapasRows[filaExistenteIndex] : null;
    const estadoGeneralActual = filaExistente ? String(filaExistente[9] || '').trim() : '';
    const comentariosActuales = filaExistente ? String(filaExistente[10] || '') : '';

    if (filaExistente) {
      estadoFormularios = parseEstadoFormularios(filaExistente[8]);
    }

    for (let i = 1; i < parsedHoja; i++) {
      if (normalizeEstadoFormulario(estadoFormularios[i.toString()]) === ESTADOS_FORMULARIO.EN_PROGRESO) {
        estadoFormularios[i.toString()] = ESTADOS_FORMULARIO.COMPLETADO;
      }
    }
    estadoFormularios[parsedHoja.toString()] = (parsedPaso >= maxPasos[parsedHoja])
      ? ESTADOS_FORMULARIO.COMPLETADO
      : ESTADOS_FORMULARIO.EN_PROGRESO;

    const estadoFormulariosJSON = JSON.stringify(estadoFormularios);
    const estadoGeneral = calculateEstadoGeneral(estadoFormularios, estadoGeneralActual);
    const estadoGlobal = getEstadoGeneralLegacy(estadoGeneral);
    const etapaActual = (parsedPaso >= maxPasos[parsedHoja] && parsedHoja < 4) ? parsedHoja + 1 : parsedHoja;
    const etapaActualAjustada = etapaActual > 4 ? 4 : etapaActual;
    
    // Usar directamente el nombre_actividad de formData si existe (ya que lo extrajimos explícitamente al inicio)
    // y está en el formulario 1, paso 1
    let nombreActividadActual;
    
    if (parsedHoja === 1 && parsedPaso === 1 && formData.nombre_actividad) {
      // Usar el valor actualizado de nombre_actividad
      nombreActividadActual = formData.nombre_actividad;
      console.log(`⭐ Actualizando nombre de actividad en ETAPAS: "${nombreActividadActual}"`);
    } else {
      // Usar el valor existente o el que viene en la solicitud
      nombreActividadActual = filaExistente ? 
        (filaExistente[6] || formData.nombre_actividad || 'N/A') : 
        (formData.nombre_actividad || 'N/A');
    }

    if (filaExistenteIndex === -1) {
      const fechaActual = dateUtils.getCurrentDate();
      const nuevaFila = [
        id_solicitud,
        id_usuario || 'N/A',
        fechaActual,
        name || 'N/A',
        etapaActualAjustada,
        estadoGlobal,
        nombreActividadActual,
        parsedPaso,
        estadoFormulariosJSON,
        estadoGeneral,
        ''
      ];
      await client.spreadsheets.values.append({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'ETAPAS!A:K',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: [nuevaFila] }
      });
    } else {
      const filaEtapas = filaExistenteIndex + 1; // Ajustar índice a 1-based
      const updateRequests = [
        { range: `ETAPAS!E${filaEtapas}`, values: [[etapaActualAjustada]] },
        { range: `ETAPAS!F${filaEtapas}`, values: [[estadoGlobal]] },
        { range: `ETAPAS!H${filaEtapas}`, values: [[parsedPaso]] },
        { range: `ETAPAS!I${filaEtapas}`, values: [[estadoFormulariosJSON]] },
        { range: `ETAPAS!J${filaEtapas}`, values: [[estadoGeneral]] },
        { range: `ETAPAS!K${filaEtapas}`, values: [[comentariosActuales]] },
        { range: `ETAPAS!B${filaEtapas}`, values: [[id_usuario || filaExistente[1]]] }, // Actualizar si viene, si no mantener
        { range: `ETAPAS!G${filaEtapas}`, values: [[nombreActividadActual]] }, // Actualizar nombre actividad
        { range: `ETAPAS!D${filaEtapas}`, values: [[name || filaExistente[3]]] } // Actualizar si viene, si no mantener
      ];
      await client.spreadsheets.values.batchUpdate({
        spreadsheetId: sheetsService.spreadsheetId,
        resource: { data: updateRequests, valueInputOption: 'RAW' }
      });
    }
    // --- Fin lógica ETAPAS ---

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
    // 1. Primero obtener el último ID existente en SOLICITUDES
    const lastId = await sheetsService.getLastId('SOLICITUDES');
    const id_solicitud = lastId + 1;
    
    // 2. Extraer resto de datos del body (sin id_solicitud)
    const { fecha_solicitud, nombre_actividad, nombre_solicitante, dependencia_tipo, nombre_dependencia } = req.body;
    
    // Asegurar que nombre_actividad tenga un valor, con un mensaje claro si no se proporciona
    const activityName = nombre_actividad || 'Actividad sin nombre';
    
    const client = sheetsService.getClient();
    const fechaActual = dateUtils.getCurrentDate();
    
    // 3. Crear entrada en SOLICITUDES
    const range = 'SOLICITUDES!A2:F2';
    // Orden correcto según modelo SOLICITUDES: A=id, B=nombre_actividad, C=fecha_solicitud
    const values = [[id_solicitud, activityName, fecha_solicitud, nombre_solicitante, dependencia_tipo, nombre_dependencia]];

    await client.spreadsheets.values.append({
      spreadsheetId: sheetsService.spreadsheetId,
      range,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      resource: { values },
    });
    
    // 4. Crear entrada en ETAPAS con el mismo ID
    const estadoFormularios = getDefaultEstadoFormularios();
    
    // Usar el mismo nombre de actividad para ETAPAS que fue guardado en SOLICITUDES
    const nuevaFila = [
      id_solicitud,
      req.body.id_usuario || 'N/A',
      fechaActual,
      req.body.name || 'N/A',
      1,                          // etapa_actual inicial
      'En progreso',              // estado inicial
      activityName,              // Usar la misma variable validada, no un valor por defecto
      1,                          // paso inicial
      JSON.stringify(estadoFormularios),
      ESTADOS_GENERALES.EN_PROCESO,
      ''
    ];
    
    await client.spreadsheets.values.append({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A:K',
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      resource: { values: [nuevaFila] }
    });

    console.log(`✅ Nueva solicitud "${activityName}" guardada en Sheets con ID: ${id_solicitud}`);
    
    // 5. Solo ahora enviamos la respuesta exitosa
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

    if (!userId) {
      return res.status(400).json({ error: 'Se requiere userId' });
    }

    const client = sheetsService.getClient();
    const { role, isAdmin } = await getUserRoleInfo(userId);

    const activeResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A2:K',
    });

    const rows = activeResponse.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron solicitudes activas o terminadas' });
    }

    const canAccessRow = (row) => isAdmin || row[1] === userId;
    const visibleRows = rows.filter((row) => canAccessRow(row));

    const activeRequests = visibleRows.filter((row) => getEstadoGeneralFromRow(row) === ESTADOS_GENERALES.EN_PROCESO);
    const completedRequests = visibleRows.filter((row) => getEstadoGeneralFromRow(row) !== ESTADOS_GENERALES.EN_PROCESO);
    const adminReviewRequests = isAdmin
      ? visibleRows.filter((row) => {
        if (row[1] === userId) return false;
        const estadoFormularios = parseEstadoFormularios(row[8]);
        return getFormulariosByEstado(estadoFormularios, [ESTADOS_FORMULARIO.ENVIADO_REVISION]).length > 0;
      })
      : [];

    res.status(200).json({
      role,
      isAdmin,
      activeRequests,
      completedRequests,
      adminReviewRequests
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
    console.log('Obteniendo solicitudes activas para:', userId);
    
    if (!userId) {
      return res.status(400).json({ error: 'Se requiere el ID de usuario' });
    }
    
    const client = sheetsService.getClient();
    const { isAdmin } = await getUserRoleInfo(userId);

    // 1. Obtener datos de ETAPAS con manejo de errores mejorado
    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A2:K',
    }).catch(err => {
      console.error(`Error al acceder a ETAPAS: ${err.message}`);
      throw new Error(`Error al acceder a ETAPAS: ${err.message}`);
    });

    const etapasRows = etapasResponse.data.values || [];
    
    // 2. Filtrar y mapear con validación robusta de datos
    const activeRequests = etapasRows
      .filter(row => {
        const estadoGeneral = getEstadoGeneralFromRow(row);
        return row.length >= 9 && 
               (isAdmin || row[1] === userId) && 
               estadoGeneral === ESTADOS_GENERALES.EN_PROCESO;
      })
      .map(row => {
        // Añadir chequeos de valores
        const etapa = row[4] ? parseInt(row[4], 10) : 1;
        const paso = row[7] ? parseInt(row[7], 10) : 1;
        const estadoGeneral = getEstadoGeneralFromRow(row);
        
        // Parsing seguro de JSON
        const estadoFormularios = parseEstadoFormularios(row[8]);
        
        return {
          idSolicitud: row[0] || 'N/A',
          id_usuario: row[1] || 'N/A',
          nombre_usuario: row[3] || 'N/A',
          fecha_solicitud: row[2] || '',
          formulario: isNaN(etapa) ? 1 : etapa,
          paso: isNaN(paso) ? 1 : paso,
          estadoFormularios,
          estado_general: estadoGeneral,
          comentarios: row[10] || '',
          nombre_actividad: row[6]?.trim() || 'Sin nombre'
        };
      });

    res.status(200).json(activeRequests);
  } catch (error) {
    console.error('Error en getActiveRequests:', error);
    res.status(500).json({ 
      error: 'Error al obtener solicitudes activas',
      details: error.message 
    });
  }
};

/**
 * Obtiene solicitudes completadas de un usuario
*/
const getCompletedRequests = async (req, res) => {
  try {
    const { userId } = req.query;

    if (!userId) {
      return res.status(400).json({ error: 'Se requiere userId' });
    }

    console.log('Obteniendo solicitudes completadas para el usuario:', userId);
    const client = sheetsService.getClient();
    const { isAdmin } = await getUserRoleInfo(userId);

    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A2:K',
    });

    const rows = etapasResponse.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'No se encontraron solicitudes completadas' });
    }

    const completedRequests = rows
      .filter((row) => {
        const estadoGeneral = getEstadoGeneralFromRow(row);
        return (isAdmin || row[1] === userId) && estadoGeneral !== ESTADOS_GENERALES.EN_PROCESO;
      })
      .map((row) => ({
        idSolicitud: row[0],
        id_usuario: row[1] || 'N/A',
        fecha_solicitud: row[2] || '',
        nombre_usuario: row[3] || 'N/A',
        formulario: parseInt(row[4]),
        etapa_actual: parseInt(row[4]),
        paso: parseInt(row[7]),
        estado_general: getEstadoGeneralFromRow(row),
        comentarios: row[10] || '',
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
    console.log('📝 Payload en guardarGastos:', JSON.stringify(req.body, null,2));
    const { id_solicitud, gastos, actualizarConceptos = true } = req.body;

    if (!id_solicitud || !gastos?.length) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere id_solicitud y al menos un concepto'
      });
    }

    // Usar el servicio de sheetsService para guardar gastos - pasar el tercer parámetro
    const success = await sheetsService.saveGastos(id_solicitud, gastos, actualizarConceptos);
    
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
      range: 'ETAPAS!A:K'
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
    
    const rowActual = etapasRows[filaEtapas - 1];
    const estadoFormularios = parseEstadoFormularios(rowActual[8]);
    
    // Actualizar estados según etapa_actual
    // Si el formulario es menor que etapa_actual, se considera completado
    for (let i = 1; i <= 4; i++) {
      if (i < parsedEtapa) {
        estadoFormularios[i.toString()] = ESTADOS_FORMULARIO.COMPLETADO;
      }
    }
    
    const estadoGeneral = calculateEstadoGeneral(estadoFormularios, String(rowActual[9] || '').trim());
    const estadoLegacy = getEstadoGeneralLegacy(estadoGeneral);

    // Actualizar la hoja con la nueva información
    await client.spreadsheets.values.update({
      spreadsheetId: sheetsService.spreadsheetId,
      range: `ETAPAS!E${filaEtapas}:J${filaEtapas}`,
      valueInputOption: 'RAW',
      resource: {
        values: [[
          parsedEtapa,
          estadoLegacy,
          rowActual[6] || 'N/A', // Mantener el nombre_actividad
          parsedPaso,
          JSON.stringify(estadoFormularios),
          estadoGeneral
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
      message: 'Paso máximo actualizado correctamente',
      data: {
        estado_general: estadoGeneral
      }
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
    const id_usuario = req.body.id_usuario || 'N/A';
    const name = req.body.name || 'N/A';

    // Validaciones básicas
    if (!id_solicitud || etapa_destino === undefined || paso_destino === undefined) {
      return res.status(400).json({
        success: false,
        error: 'Se requieren id_solicitud, etapa_destino y paso_destino'
      });
    }

    const client = sheetsService.getClient();
    
    // Obtener datos de ETAPAS con manejo de creación de fila
    const etapasResponse = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A:K'
    });

    let etapasRows = etapasResponse.data.values || [];
    let filaExistente = etapasRows.find(row => row[0] === id_solicitud.toString());

    // Si no existe la solicitud, crear registro inicial
    if (!filaExistente) {
      console.log(`No se encontró registro para solicitud ${id_solicitud}. Creando registro inicial...`);
      const fechaActual = dateUtils.getCurrentDate();
      const estadoFormularios = getDefaultEstadoFormularios();
      
      const nuevaFila = [
        id_solicitud,
        req.body.id_usuario || 'N/A',  // Usar datos del request
        fechaActual,
        req.body.name || 'N/A',        // Usar datos del request
        1,                             // etapa_actual inicial
        'En progreso',
        req.body.nombre_actividad || 'Nueva solicitud', 
        1,                             // paso inicial
        JSON.stringify(estadoFormularios),
        ESTADOS_GENERALES.EN_PROCESO,
        ''
      ];

      try {
        await client.spreadsheets.values.append({
          spreadsheetId: sheetsService.spreadsheetId,
          range: 'ETAPAS!A:K',
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS',
          resource: { values: [nuevaFila] }
        });
        
        console.log(`✅ Registro inicial creado para solicitud ${id_solicitud}`);
        
        // Actualizar lista de filas para continuar el proceso
        filaExistente = nuevaFila;
        
        // También guardar en progressState si está disponible
        if (req.session) {
          req.session.progressState = {
            etapa_actual: 1,
            paso: 1,
            estado: 'En progreso',
            estado_formularios: estadoFormularios,
            id_solicitud: id_solicitud
          };
        }
      } catch (appendError) {
        console.error(`Error al crear registro inicial para solicitud ${id_solicitud}:`, appendError);
        return res.status(500).json({
          success: false,
          error: 'Error al crear registro inicial para la solicitud',
          details: appendError.message
        });
      }
    }

    // Continuar con el resto de la lógica normal de validarProgresion
    // Extraer datos actuales
    const etapaActual = parseInt(filaExistente[4]) || 1;
    const pasoActual = parseInt(filaExistente[7]) || 1;
    
    // Extraer estado de formularios con manejo de errores
    const estadoFormularios = parseEstadoFormularios(filaExistente[8]);
    
    // Resto de la lógica de validación...
    // Simplificamos la validación para permitir navegación flexible
    const formulariosIniciados = Object.entries(estadoFormularios)
      .filter(([_, estado]) => isFormularioRealizado(estado) || normalizeEstadoFormulario(estado) === ESTADOS_FORMULARIO.EN_PROGRESO)
      .map(([num, _]) => parseInt(num));
      
    let puedeAvanzar = true;
    let mensaje = '';
    
    // ÚNICA RESTRICCIÓN: No permitir saltar a formularios futuros no iniciados
    if (etapa_destino > etapaActual && !formulariosIniciados.includes(etapa_destino)) {
      puedeAvanzar = false;
      mensaje = 'No puede acceder a formularios futuros sin completar los anteriores';
    } else {
      mensaje = etapa_destino < etapaActual 
        ? 'Navegando a un formulario anterior' 
        : etapa_destino === etapaActual 
          ? 'Continuando en el formulario actual' 
          : 'Avanzando al siguiente formulario';
    }
    
    return res.status(200).json({
      success: true,
      puedeAvanzar,
      mensaje,
      estado: {
        etapaActual,
        pasoActual,
        estadoGeneral: getEstadoGeneralFromRow(filaExistente),
        comentarios: filaExistente[10] || '',
        estadoFormularios,
        etapaDestino: parseInt(etapa_destino),
        pasoDestino: parseInt(paso_destino),
        formulariosIniciados
      }
    });
    
  } catch (error) {
    console.error('Error en validarProgresion:', error);
    res.status(500).json({
      success: false,
      error: 'Error al validar progresión',
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
      range: 'ETAPAS!A:K'
    });
    
    const etapasRows = etapasResponse.data.values || [];
    let filaEtapas = etapasRows.findIndex(row => row[0] === id_solicitud.toString());
    
    if (filaEtapas === -1) {
      return res.status(404).json({
        success: false,
        error: `No se encontró la solicitud con ID ${id_solicitud}`
      });
    }
    
    filaEtapas += 1; // Ajustar índice a 1-based para Google Sheets
    
    // Recuperar el estado de formularios existente o usar el proporcionado
    let nuevoEstadoFormularios = {};
    const rowActual = etapasRows[filaEtapas - 1];
    
    if (estadoFormularios) {
      // Usar el estado proporcionado
      nuevoEstadoFormularios = parseEstadoFormularios(estadoFormularios);
    } else {
      // Recuperar estado existente o crear uno nuevo
      nuevoEstadoFormularios = parseEstadoFormularios(rowActual[8]);
      
      // Actualizar el estado del formulario actual
      nuevoEstadoFormularios[parsedEtapa.toString()] = 
        (parsedPaso >= maxPasos[parsedEtapa]) ? ESTADOS_FORMULARIO.COMPLETADO : ESTADOS_FORMULARIO.EN_PROGRESO;
      
      // Marcar formularios anteriores como completados si se solicita
      if (actualizar_formularios_previos) {
        for (let i = 1; i < parsedEtapa; i++) {
          nuevoEstadoFormularios[i.toString()] = ESTADOS_FORMULARIO.COMPLETADO;
        }
      }
    }
    
    // Determinar el estado global legacy (col F) y estado general (col J)
    const estadoGeneralActual = String(rowActual[9] || '').trim();
    const nuevoEstadoGeneral = Object.values(ESTADOS_GENERALES).includes(estadoGlobal)
      ? estadoGlobal
      : calculateEstadoGeneral(nuevoEstadoFormularios, estadoGeneralActual);
    const nuevoEstadoGlobal = getEstadoGeneralLegacy(nuevoEstadoGeneral);
    
    // Actualizar la hoja
    await client.spreadsheets.values.update({
      spreadsheetId: sheetsService.spreadsheetId,
      range: `ETAPAS!E${filaEtapas}:J${filaEtapas}`,
      valueInputOption: 'RAW',
      resource: {
        values: [[
          parsedEtapa,
          nuevoEstadoGlobal,
          rowActual[6] || 'N/A', // Mantener nombre_actividad
          parsedPaso,
          JSON.stringify(nuevoEstadoFormularios),
          nuevoEstadoGeneral
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
        estado_general: nuevoEstadoGeneral,
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
    
    // Utilizar el método mejorado de sheetsService para obtener el último ID
    const lastId = await sheetsService.getLastId(sheetName);
    
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
 * Guarda datos del Paso 1 del Formulario 2 en SOLICITUDES2
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const guardarForm2Paso1 = async (req, res) => {
  try {
    const { id_solicitud, nombre_actividad, fecha_solicitud } = req.body;

    if (!id_solicitud) {
      return res.status(400).json({ error: 'El ID de solicitud es obligatorio' });
    }

    console.log("📝 Datos recibidos para guardar en SOLICITUDES2 (Paso 1):", req.body);

    // Buscar la fila de la solicitud en SOLICITUDES2
    const resultado = await sheetsService.findOrCreateRequestRow('SOLICITUDES2', id_solicitud);
    
    if (!resultado || !resultado.rowIndex) {
      return res.status(404).json({ error: 'No se pudo encontrar o crear la fila para la solicitud' });
    }

    // Mapear los campos al modelo correspondiente en Sheets
    const modelo = sheetsService.models.SOLICITUDES2;
    const modelFields = modelo.fields || [];
    const updateValues = [];
    const columnas = [];
    
    // Datos básicos del paso 1
    const campos = ['nombre_actividad', 'fecha_solicitud'];

    // Mapear cada campo al modelo y añadir a la actualización
    campos.forEach(campo => {
      const fieldIndex = modelFields.indexOf(campo);
      if (fieldIndex !== -1) {
        const colIndex = fieldIndex + 1; // Columnas 1-based
        columnas.push(colIndex);
        updateValues.push(req.body[campo] !== undefined ? req.body[campo].toString() : '');
      }
    });

    // Si hay campos para actualizar
    if (columnas.length > 0) {
      // Calcular columna de inicio y fin para la actualización
      const startColumn = Math.min(...columnas);
      const endColumn = Math.max(...columnas);
      
      // Crear un array lleno de valores vacíos para todas las columnas en el rango
      const rangeValues = Array(endColumn - startColumn + 1).fill('');
      
      // Colocar los valores en las posiciones correctas
      columnas.forEach((col, index) => {
        rangeValues[col - startColumn] = updateValues[index];
      });
      
      // Actualizar la hoja
      await sheetsService.updateRequestProgress({
        sheetName: 'SOLICITUDES2',
        rowIndex: resultado.rowIndex,
        startColumn,
        endColumn,
        values: rangeValues
      });

      console.log(`✅ Datos Paso 1 guardados en SOLICITUDES2 para solicitud ${id_solicitud}`);
      
      // Actualizar el progreso global
      await progressService.updateRequestProgress(id_solicitud, 2, 1);
      
      return res.status(200).json({ success: true, message: 'Datos básicos guardados correctamente' });
    } else {
      return res.status(400).json({ error: 'No hay campos válidos para actualizar' });
    }
  } catch (error) {
    console.error('Error al guardar datos del Formulario 2 Paso 1:', error);
    return res.status(500).json({ error: 'Error al guardar datos básicos', details: error.message });
  }
};

/**
 * Guarda datos específicos del Formulario 2 Paso 2 en SOLICITUDES2
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const guardarForm2Paso2 = async (req, res) => {
  try {
    const { id_solicitud, formData } = req.body;
    
    if (!id_solicitud) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere id_solicitud'
      });
    }

    console.log('📝 DATOS RECIBIDOS:', formData);
    
    // 1. Buscar la fila de la solicitud
    const resultado = await sheetsService.findOrCreateRequestRow('SOLICITUDES2', id_solicitud);
    if (!resultado) {
      return res.status(404).json({ 
        error: 'No se pudo encontrar o crear la fila para la solicitud' 
      });
    }
    
    // 2. Preparar datos para guardar - CORREGIDO EL MAPEO
    const campos = [
      'nombre_actividad',
      'fecha_solicitud',
      'ingresos_cantidad',
      'ingresos_vr_unit',
      'total_ingresos',
      'subtotal_gastos',
      'imprevistos_3', // IMPORTANTE: Este valor se guardará en la columna 'imprevistos_3%'
      'total_gastos_imprevistos',
      'fondo_comun_porcentaje',
      'fondo_comun',
      'facultad_instituto',
      'escuela_departamento_porcentaje',
      'escuela_departamento',
      'total_recursos'
    ];
    
    // 3. Crear array de valores
    const valores = [];
    
    // Mapear campos del formData a sus valores
    campos.forEach(campo => {
      // Para el caso especial de imprevistos_3
      if (campo === 'imprevistos_3' && formData['imprevistos_3%'] !== undefined) {
        valores.push(formData['imprevistos_3%'] || '');
      }
      // Para el resto de campos
      else {
        valores.push(formData[campo] || '');
      }
    });

    // 4. Actualizar en Google Sheets
    await sheetsService.client.spreadsheets.values.update({
      spreadsheetId: sheetsService.spreadsheetId,
      range: `SOLICITUDES2!B${resultado.rowIndex}:M${resultado.rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [valores]
      }
    });
    
    // 5. Actualizar el progreso global IMPORTANTE
    try {
      // Actualizar estado de formularios
      const etapasResponse = await sheetsService.client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'ETAPAS!A:K'
      });
      
      const etapasRows = etapasResponse.data.values || [];
      let filaEtapa = etapasRows.findIndex(row => row[0] === id_solicitud.toString());
      
      if (filaEtapa !== -1) {
        filaEtapa += 1;
        
        const rowActual = etapasRows[filaEtapa - 1];
        const estadoFormularios = parseEstadoFormularios(rowActual[8]);
        
        // Actualizar formulario 2 como "En progreso"
        estadoFormularios['2'] = ESTADOS_FORMULARIO.EN_PROGRESO;
        const estadoGeneral = calculateEstadoGeneral(estadoFormularios, String(rowActual[9] || '').trim());
        const estadoLegacy = getEstadoGeneralLegacy(estadoGeneral);
        
        // Guardar en ETAPAS
        await sheetsService.client.spreadsheets.values.batchUpdate({
          spreadsheetId: sheetsService.spreadsheetId,
          resource: {
            valueInputOption: 'RAW',
            data: [
              { range: `ETAPAS!F${filaEtapa}`, values: [[estadoLegacy]] },
              { range: `ETAPAS!I${filaEtapa}`, values: [[JSON.stringify(estadoFormularios)]] },
              { range: `ETAPAS!J${filaEtapa}`, values: [[estadoGeneral]] }
            ]
          }
        });
        
        // Actualizar en sesión
        if (req.session) {
          req.session.progressState = {
            ...req.session.progressState,
            estado_formularios: estadoFormularios
          };
        }
      }
    } catch (error) {
      console.error("Error al actualizar estado de progreso:", error);
      // Continuar aunque haya error en el progreso
    }
    
    console.log('✅ Datos del formulario 2 paso 2 guardados');
    res.status(200).json({ 
      success: true, 
      message: 'Datos del formulario 2 paso 2 guardados correctamente'
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

/**
 * Guarda los datos del paso 3 del formulario 2 (aportes y resumen financiero)
 * @param {Object} req - Solicitud HTTP
 * @param {Object} res - Respuesta HTTP
 */
const guardarForm2Paso3 = async (req, res) => {
  try {
    const { 
      id_solicitud, 
      id_usuario, // Capturar explícitamente id_usuario del request
      name,       // Capturar explícitamente el nombre del request
      fondo_comun_porcentaje, 
      fondo_comun,
      facultad_instituto_porcentaje, // Campo editable
      facultad_instituto, 
      escuela_departamento_porcentaje, 
      escuela_departamento,
      total_recursos,
      ingresos_cantidad,
      ingresos_vr_unit,
      total_ingresos,
      subtotal_gastos,
      imprevistos_3,
      total_gastos_imprevistos,
      observaciones // Campo de observaciones
    } = req.body;

    console.log('Datos recibidos en guardarForm2Paso3:', req.body);

    if (!id_solicitud) {
      return res.status(400).json({ error: 'ID de solicitud no proporcionado' });
    }

    // Establecer valores por defecto si no se proporcionan
    const fondoComunPorcentaje = parseFloat(fondo_comun_porcentaje) || 0;
    const facultadInstitutoPorcentaje = parseFloat(facultad_instituto_porcentaje) || 5; // Ahora editable
    const escuelaDepartamentoPorcentaje = parseFloat(escuela_departamento_porcentaje) || 0;

    // Recalcular valores si es necesario
    let totalIngresos = parseFloat(total_ingresos) || 0;
    if (!totalIngresos && ingresos_cantidad && ingresos_vr_unit) {
      totalIngresos = parseFloat(ingresos_cantidad) * parseFloat(ingresos_vr_unit);
    }

    // Calcular valores derivados
    const fondoComun = parseFloat(fondo_comun) || (totalIngresos * fondoComunPorcentaje / 100);
    const facultadInstitutoValor = parseFloat(facultad_instituto) || (totalIngresos * facultadInstitutoPorcentaje / 100); // Usar porcentaje editable
    const escuelaDepartamentoValor = parseFloat(escuela_departamento) || (totalIngresos * escuelaDepartamentoPorcentaje / 100);
    const totalRecursosValor = parseFloat(total_recursos) || (fondoComun + facultadInstitutoValor + escuelaDepartamentoValor);
    
    // Asegurar que observaciones tenga valor
    const observacionesValor = observaciones || '';

    try {
      // 1. Encontrar o crear la fila para esta solicitud
      const rowInfo = await sheetsService.findOrCreateRequestRow('SOLICITUDES2', id_solicitud);
      if (!rowInfo || !rowInfo.rowIndex) {
        throw new Error(`No se pudo encontrar o crear fila para solicitud ${id_solicitud}`);
      }

      // 2. Actualizar la hoja SOLICITUDES2 con los valores calculados
      // Solo actualizar las columnas K-R que corresponden al paso 3
      const valoresAportes = [
        fondoComunPorcentaje.toString(), // Columna K: fondo_comun_porcentaje
        fondoComun.toString(), // Columna L: fondo_comun
        facultadInstitutoPorcentaje.toString(), // Columna M: facultad_instituto_porcentaje (editable)
        facultadInstitutoValor.toString(), // Columna N: facultad_instituto
        escuelaDepartamentoPorcentaje.toString(), // Columna O: escuela_departamento_porcentaje
        escuelaDepartamentoValor.toString(), // Columna P: escuela_departamento
        totalRecursosValor.toString(), // Columna Q: total_recursos
        observacionesValor // Columna R: observaciones
      ];

      console.log(`Actualizando datos de aportes en SOLICITUDES2 para la solicitud ${id_solicitud}: `, valoresAportes);

      // 4. Actualizar solo las columnas K a R (correspondiente al paso 3)
      await sheetsService.updateRequestProgress({
        sheetName: 'SOLICITUDES2',
        rowIndex: rowInfo.rowIndex,
        startColumn: 'K',    // Columna inicial para el paso 3
        endColumn: 'R',    // Columna final para el paso 3
        values: valoresAportes
      });
      
      // 5. Actualizar el progreso global de la solicitud
      await progressService.updateRequestProgress(id_solicitud, 2, 3); // Formulario 2, Paso 3 (completado)
      
      // 7. Actualizar estado de formularios en ETAPAS
      const etapasResponse = await sheetsService.getClient().spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'ETAPAS!A:K'
      });
      const etapasRows = etapasResponse.data.values || [];
      const filaEtapaIndex = etapasRows.findIndex(row => row[0] === id_solicitud.toString());
      if (filaEtapaIndex !== -1) {
        const filaEtapa = filaEtapaIndex + 1; // Ajustar a 1-based para Sheets
        
        // Obtener valores actuales para preservarlos
        const currentUserId = etapasRows[filaEtapaIndex][1] || 'N/A';
        const currentUserName = etapasRows[filaEtapaIndex][3] || 'N/A';
        
        // Usar los valores del request o mantener los existentes
        const userIdToUpdate = id_usuario || currentUserId;
        const nameToUpdate = name || currentUserName;
        
        let estadoFormularios = {
          '1': ESTADOS_FORMULARIO.COMPLETADO,
          '2': ESTADOS_FORMULARIO.COMPLETADO,
          '3': ESTADOS_FORMULARIO.EN_PROGRESO,
          '4': ESTADOS_FORMULARIO.EN_PROGRESO
        };
        try {
          if (etapasRows[filaEtapaIndex][8]) {
            const parsedEstado = JSON.parse(etapasRows[filaEtapaIndex][8]);
            if (parsedEstado) {
              estadoFormularios = parsedEstado;
              estadoFormularios['2'] = ESTADOS_FORMULARIO.COMPLETADO; // Marcar formulario 2 como completado
            }
          }
        } catch (e) {
          console.error("Error al parsear estado_formularios:", e);
        }

        const estadoGeneral = calculateEstadoGeneral(estadoFormularios, String(etapasRows[filaEtapaIndex][9] || '').trim());
        const estadoLegacy = getEstadoGeneralLegacy(estadoGeneral);
        
        // Actualizar etapa, estado formulario, y PRESERVAR id_usuario y name
        await sheetsService.getClient().spreadsheets.values.batchUpdate({
          spreadsheetId: sheetsService.spreadsheetId,
          resource: {
            data: [
              { range: `ETAPAS!B${filaEtapa}`, values: [[userIdToUpdate]] }, // Preservar id_usuario
              { range: `ETAPAS!D${filaEtapa}`, values: [[nameToUpdate]] },   // Preservar name
              { range: `ETAPAS!E${filaEtapa}`, values: [[3]] }, // Avanzar a etapa 3
              { range: `ETAPAS!F${filaEtapa}`, values: [[estadoLegacy]] },
              { range: `ETAPAS!I${filaEtapa}`, values: [[JSON.stringify(estadoFormularios)]] },
              { range: `ETAPAS!J${filaEtapa}`, values: [[estadoGeneral]] }
            ],
            valueInputOption: 'RAW'
          }
        });
      }
      console.log('✅ Datos del formulario 2 paso 3 guardados correctamente');
      res.status(200).json({
        success: true,
        message: 'Datos de aportes guardados correctamente'
      });
    } catch (sheetsError) {
      console.error('❌ Error al guardar datos en Google Sheets:', sheetsError);
      res.status(500).json({ 
        error: 'Error al guardar datos en Google Sheets',
        details: sheetsError.message 
      });
    }
  } catch (error) {
    console.error('❌ Error en guardarForm2Paso3:', error);
    res.status(500).json({ 
      error: 'Error al guardar datos del paso 3',
      details: error.message 
    });
  }
};

/**
 * Usuario normal: envía a revisión uno o varios formularios ya realizados.
 */
const enviarSolicitudRevision = async (req, res) => {
  try {
    const { id_solicitud, userId, formularios } = req.body;

    if (!id_solicitud || !userId) {
      return res.status(400).json({
        success: false,
        error: 'Se requieren id_solicitud y userId'
      });
    }

    const client = sheetsService.getClient();
    const response = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A2:K'
    });

    const rows = response.data.values || [];
    const filaIndex = rows.findIndex((row) => String(row[0] || '').trim() === String(id_solicitud).trim());

    if (filaIndex === -1) {
      return res.status(404).json({
        success: false,
        error: `No se encontró la solicitud con ID ${id_solicitud}`
      });
    }

    const fila = rows[filaIndex];
    const isAdmin = await userIsAdmin(userId);
    if (!isAdmin && String(fila[1] || '').trim() !== String(userId).trim()) {
      return res.status(403).json({
        success: false,
        error: 'No tienes permisos para enviar esta solicitud a revisión'
      });
    }

    const estadoFormularios = parseEstadoFormularios(fila[8]);
    const formulariosSolicitados = sanitizeFormularios(formularios);
    const formulariosRealizados = getFormulariosByEstado(estadoFormularios, [
      ESTADOS_FORMULARIO.COMPLETADO,
      ESTADOS_FORMULARIO.REQUIERE_CORRECCIONES,
      ESTADOS_FORMULARIO.ENVIADO_REVISION,
      ESTADOS_FORMULARIO.APROBADO
    ]);

    const formulariosAEnviar = formulariosSolicitados.length > 0
      ? formulariosSolicitados
      : formulariosRealizados.filter((formularioId) => {
        const estado = normalizeEstadoFormulario(estadoFormularios[String(formularioId)]);
        return estado !== ESTADOS_FORMULARIO.APROBADO && estado !== ESTADOS_FORMULARIO.ENVIADO_REVISION;
      });

    if (formulariosAEnviar.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'No hay formularios listos para enviar a revisión'
      });
    }

    const invalidos = [];
    formulariosAEnviar.forEach((formularioId) => {
      const estadoActual = normalizeEstadoFormulario(estadoFormularios[String(formularioId)]);
      if (estadoActual === ESTADOS_FORMULARIO.EN_PROGRESO || estadoActual === ESTADOS_FORMULARIO.APROBADO) {
        invalidos.push(formularioId);
        return;
      }

      setEstadoFormulario(estadoFormularios, formularioId, ESTADOS_FORMULARIO.ENVIADO_REVISION);
    });

    if (invalidos.length > 0) {
      return res.status(400).json({
        success: false,
        error: `Los formularios ${invalidos.join(', ')} no se pueden enviar a revisión en su estado actual`
      });
    }

    const estadoGeneral = calculateEstadoGeneral(estadoFormularios, String(fila[9] || '').trim());
    const estadoLegacy = getEstadoGeneralLegacy(estadoGeneral);

    const sheetRow = filaIndex + 2;
    await client.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetsService.spreadsheetId,
      resource: {
        valueInputOption: 'RAW',
        data: [
          { range: `ETAPAS!F${sheetRow}`, values: [[estadoLegacy]] },
          { range: `ETAPAS!I${sheetRow}`, values: [[JSON.stringify(estadoFormularios)]] },
          { range: `ETAPAS!J${sheetRow}`, values: [[estadoGeneral]] }
        ]
      }
    });

    return res.status(200).json({
      success: true,
      message: 'Formularios enviados a revisión',
      data: {
        id_solicitud: String(id_solicitud),
        formularios_enviados: formulariosAEnviar,
        estado_formularios: estadoFormularios,
        estado_general: estadoGeneral
      }
    });
  } catch (error) {
    console.error('Error en enviarSolicitudRevision:', error);
    return res.status(500).json({
      success: false,
      error: 'Error al enviar la solicitud a revisión',
      details: error.message
    });
  }
};

/**
 * Admin: obtiene solicitudes de otros usuarios con al menos un formulario en revisión.
 */
const getSolicitudesRevisionAdmin = async (req, res) => {
  try {
    const { userId } = req.query;

    if (!userId) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere userId'
      });
    }

    const isAdmin = await userIsAdmin(userId);
    if (!isAdmin) {
      return res.status(403).json({
        success: false,
        error: 'Solo un administrador puede consultar solicitudes en revisión'
      });
    }

    const client = sheetsService.getClient();
    const response = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A2:K'
    });

    const rows = response.data.values || [];
    const solicitudes = rows
      .filter((row) => {
        if (String(row[1] || '').trim() === String(userId).trim()) {
          return false;
        }

        const estadoFormularios = parseEstadoFormularios(row[8]);
        const formulariosEnRevision = getFormulariosByEstado(estadoFormularios, [ESTADOS_FORMULARIO.ENVIADO_REVISION]);
        return formulariosEnRevision.length > 0;
      })
      .map((row) => buildEtapaSummary(row));

    return res.status(200).json({
      success: true,
      data: solicitudes
    });
  } catch (error) {
    console.error('Error en getSolicitudesRevisionAdmin:', error);
    return res.status(500).json({
      success: false,
      error: 'Error al obtener solicitudes en revisión',
      details: error.message
    });
  }
};

/**
 * Admin: aprueba parcialmente formularios (maximo 3 por llamada).
 */
const aprobarSolicitudAdmin = async (req, res) => {
  try {
    const { id_solicitud, userId, formularios } = req.body;

    if (!id_solicitud || !userId) {
      return res.status(400).json({
        success: false,
        error: 'Se requieren id_solicitud y userId'
      });
    }

    const isAdmin = await userIsAdmin(userId);
    if (!isAdmin) {
      return res.status(403).json({
        success: false,
        error: 'Solo un administrador puede aprobar solicitudes'
      });
    }

    const client = sheetsService.getClient();
    const response = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A2:K'
    });

    const rows = response.data.values || [];
    const filaIndex = rows.findIndex((row) => String(row[0] || '').trim() === String(id_solicitud).trim());
    if (filaIndex === -1) {
      return res.status(404).json({
        success: false,
        error: `No se encontró la solicitud con ID ${id_solicitud}`
      });
    }

    const fila = rows[filaIndex];
    const estadoFormularios = parseEstadoFormularios(fila[8]);
    const formulariosAprobar = sanitizeFormularios(formularios);

    if (formulariosAprobar.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'Debes enviar al menos un formulario para aprobar'
      });
    }

    if (formulariosAprobar.length > 3) {
      return res.status(400).json({
        success: false,
        error: 'Solo puedes aprobar maximo 3 formularios por llamada'
      });
    }

    const invalidos = [];
    formulariosAprobar.forEach((formularioId) => {
      const estadoActual = normalizeEstadoFormulario(estadoFormularios[String(formularioId)]);
      if (estadoActual !== ESTADOS_FORMULARIO.ENVIADO_REVISION) {
        invalidos.push(formularioId);
        return;
      }

      setEstadoFormulario(estadoFormularios, formularioId, ESTADOS_FORMULARIO.APROBADO);
    });

    if (invalidos.length > 0) {
      return res.status(400).json({
        success: false,
        error: `Los formularios ${invalidos.join(', ')} deben estar en estado Enviado a revisión para aprobarse`
      });
    }

    const estadoGeneral = calculateEstadoGeneral(estadoFormularios, String(fila[9] || '').trim());
    const estadoLegacy = getEstadoGeneralLegacy(estadoGeneral);
    const sheetRow = filaIndex + 2;

    await client.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetsService.spreadsheetId,
      resource: {
        valueInputOption: 'RAW',
        data: [
          { range: `ETAPAS!F${sheetRow}`, values: [[estadoLegacy]] },
          { range: `ETAPAS!I${sheetRow}`, values: [[JSON.stringify(estadoFormularios)]] },
          { range: `ETAPAS!J${sheetRow}`, values: [[estadoGeneral]] }
        ]
      }
    });

    return res.status(200).json({
      success: true,
      message: 'Formularios aprobados correctamente',
      data: {
        id_solicitud: String(id_solicitud),
        formularios_aprobados: formulariosAprobar,
        estado_formularios: estadoFormularios,
        estado_general: estadoGeneral
      }
    });
  } catch (error) {
    console.error('Error en aprobarSolicitudAdmin:', error);
    return res.status(500).json({
      success: false,
      error: 'Error al aprobar solicitud',
      details: error.message
    });
  }
};

/**
 * Admin: aprueba todos los formularios de una solicitud.
 */
const aprobarSolicitudCompletaAdmin = async (req, res) => {
  try {
    const { id_solicitud, userId } = req.body;

    if (!id_solicitud || !userId) {
      return res.status(400).json({
        success: false,
        error: 'Se requieren id_solicitud y userId'
      });
    }

    const isAdmin = await userIsAdmin(userId);
    if (!isAdmin) {
      return res.status(403).json({
        success: false,
        error: 'Solo un administrador puede aprobar la solicitud completa'
      });
    }

    const client = sheetsService.getClient();
    const response = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A2:K'
    });

    const rows = response.data.values || [];
    const filaIndex = rows.findIndex((row) => String(row[0] || '').trim() === String(id_solicitud).trim());
    if (filaIndex === -1) {
      return res.status(404).json({
        success: false,
        error: `No se encontró la solicitud con ID ${id_solicitud}`
      });
    }

    const fila = rows[filaIndex];
    const estadoFormularios = parseEstadoFormularios(fila[8]);
    const formulariosPendientes = ['1', '2', '3', '4']
      .filter((key) => normalizeEstadoFormulario(estadoFormularios[key]) === ESTADOS_FORMULARIO.EN_PROGRESO)
      .map((key) => parseInt(key, 10));

    if (formulariosPendientes.length > 0) {
      return res.status(400).json({
        success: false,
        error: `No puedes aprobar solicitud completa: formularios en progreso ${formulariosPendientes.join(', ')}`
      });
    }

    ['1', '2', '3', '4'].forEach((key) => {
      estadoFormularios[key] = ESTADOS_FORMULARIO.APROBADO;
    });

    const sheetRow = filaIndex + 2;
    await client.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetsService.spreadsheetId,
      resource: {
        valueInputOption: 'RAW',
        data: [
          { range: `ETAPAS!F${sheetRow}`, values: [['Completado']] },
          { range: `ETAPAS!I${sheetRow}`, values: [[JSON.stringify(estadoFormularios)]] },
          { range: `ETAPAS!J${sheetRow}`, values: [[ESTADOS_GENERALES.APROBADO]] }
        ]
      }
    });

    return res.status(200).json({
      success: true,
      message: 'Solicitud aprobada completamente',
      data: {
        id_solicitud: String(id_solicitud),
        formularios_aprobados: [1, 2, 3, 4],
        estado_formularios: estadoFormularios,
        estado_general: ESTADOS_GENERALES.APROBADO
      }
    });
  } catch (error) {
    console.error('Error en aprobarSolicitudCompletaAdmin:', error);
    return res.status(500).json({
      success: false,
      error: 'Error al aprobar solicitud completa',
      details: error.message
    });
  }
};

/**
 * Admin: envía correcciones por formulario.
 */
const enviarCorreccionesAdmin = async (req, res) => {
  try {
    const { id_solicitud, userId, formularios, comentarios, comentarios_por_formulario } = req.body;

    if (!id_solicitud || !userId) {
      return res.status(400).json({
        success: false,
        error: 'Se requieren id_solicitud y userId'
      });
    }

    const isAdmin = await userIsAdmin(userId);
    if (!isAdmin) {
      return res.status(403).json({
        success: false,
        error: 'Solo un administrador puede enviar correcciones'
      });
    }

    const client = sheetsService.getClient();
    const response = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A2:K'
    });

    const rows = response.data.values || [];
    const filaIndex = rows.findIndex((row) => String(row[0] || '').trim() === String(id_solicitud).trim());
    if (filaIndex === -1) {
      return res.status(404).json({
        success: false,
        error: `No se encontró la solicitud con ID ${id_solicitud}`
      });
    }

    const fila = rows[filaIndex];
    const estadoFormularios = parseEstadoFormularios(fila[8]);
    const comentariosMap = parseComentariosPorFormulario(fila[10] || '');
    const formulariosCorreccion = sanitizeFormularios(formularios);

    if (formulariosCorreccion.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'Debes enviar al menos un formulario para correccion'
      });
    }

    const invalidos = [];
    formulariosCorreccion.forEach((formularioId) => {
      const estadoActual = normalizeEstadoFormulario(estadoFormularios[String(formularioId)]);
      if (estadoActual === ESTADOS_FORMULARIO.EN_PROGRESO || estadoActual === ESTADOS_FORMULARIO.APROBADO) {
        invalidos.push(formularioId);
        return;
      }

      setEstadoFormulario(estadoFormularios, formularioId, ESTADOS_FORMULARIO.REQUIERE_CORRECCIONES);

      const comentarioFormulario = comentarios_por_formulario && typeof comentarios_por_formulario === 'object'
        ? String(comentarios_por_formulario[String(formularioId)] || '').trim()
        : '';
      const comentarioGeneral = String(comentarios || '').trim();
      const comentarioFinal = comentarioFormulario || comentarioGeneral;

      if (comentarioFinal) {
        comentariosMap[String(formularioId)] = comentarioFinal;
      }
    });

    if (invalidos.length > 0) {
      return res.status(400).json({
        success: false,
        error: `Los formularios ${invalidos.join(', ')} no se pueden enviar a correccion en su estado actual`
      });
    }

    const estadoGeneral = calculateEstadoGeneral(estadoFormularios, String(fila[9] || '').trim());
    const estadoLegacy = getEstadoGeneralLegacy(estadoGeneral);
    const comentariosSerializados = stringifyComentariosPorFormulario(comentariosMap);

    const sheetRow = filaIndex + 2;
    await client.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetsService.spreadsheetId,
      resource: {
        valueInputOption: 'RAW',
        data: [
          { range: `ETAPAS!F${sheetRow}`, values: [[estadoLegacy]] },
          { range: `ETAPAS!I${sheetRow}`, values: [[JSON.stringify(estadoFormularios)]] },
          { range: `ETAPAS!J${sheetRow}`, values: [[estadoGeneral]] },
          { range: `ETAPAS!K${sheetRow}`, values: [[comentariosSerializados]] }
        ]
      }
    });

    return res.status(200).json({
      success: true,
      message: 'Correcciones enviadas correctamente',
      data: {
        id_solicitud: String(id_solicitud),
        formularios_con_correccion: formulariosCorreccion,
        estado_formularios: estadoFormularios,
        estado_general: estadoGeneral,
        comentarios_por_formulario: comentariosMap
      }
    });
  } catch (error) {
    console.error('Error en enviarCorreccionesAdmin:', error);
    return res.status(500).json({
      success: false,
      error: 'Error al enviar correcciones',
      details: error.message
    });
  }
};

/**
 * Usuario/Admin: obtiene estado general y comentarios de una solicitud.
 */
const getEstadoRevisionSolicitud = async (req, res) => {
  try {
    const { id_solicitud, userId } = req.query;

    if (!id_solicitud || !userId) {
      return res.status(400).json({
        success: false,
        error: 'Se requieren id_solicitud y userId'
      });
    }

    const client = sheetsService.getClient();
    const response = await client.spreadsheets.values.get({
      spreadsheetId: sheetsService.spreadsheetId,
      range: 'ETAPAS!A2:K'
    });

    const rows = response.data.values || [];
    const fila = rows.find((row) => String(row[0] || '').trim() === String(id_solicitud).trim());

    if (!fila) {
      return res.status(404).json({
        success: false,
        error: `No se encontró la solicitud con ID ${id_solicitud}`
      });
    }

    const isAdmin = await userIsAdmin(userId);
    if (!isAdmin && String(fila[1] || '').trim() !== String(userId).trim()) {
      return res.status(403).json({
        success: false,
        error: 'No tienes permisos para consultar esta solicitud'
      });
    }

    return res.status(200).json({
      success: true,
      data: buildEtapaSummary(fila)
    });
  } catch (error) {
    console.error('Error en getEstadoRevisionSolicitud:', error);
    return res.status(500).json({
      success: false,
      error: 'Error al obtener estado de revisión de la solicitud',
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
  guardarForm2Paso1,
  guardarForm2Paso2,
  guardarForm2Paso3,
  enviarSolicitudRevision,
  getSolicitudesRevisionAdmin,
  aprobarSolicitudAdmin,
  aprobarSolicitudCompletaAdmin,
  enviarCorreccionesAdmin,
  getEstadoRevisionSolicitud
};