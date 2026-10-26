const fs = require('fs');
const sheetsService = require('../services/sheetsService');
const driveService = require('../services/driveService');
const emailService = require('../services/emailService');

const STAGES = Object.freeze({
  SOLICITUD: 'solicitud',
  REVISION_1: 'revision_1',
  REVISION_2: 'revision_2',
  PRESUPUESTO: 'presupuesto',
  FIRMAS: 'firmas',
  DECANATO: 'decanato',
  FINALIZADO: 'finalizado'
});

const ROLE_STAGE = Object.freeze({
  revisor_1: STAGES.REVISION_1,
  revisor_2: STAGES.REVISION_2,
  presupuesto: STAGES.PRESUPUESTO,
  decanato: STAGES.DECANATO
});

const DECISION_ROLE_STAGE = Object.freeze({
  revisor_1: STAGES.REVISION_1,
  revisor_2: STAGES.REVISION_2,
  presupuesto: STAGES.PRESUPUESTO
});

const OBSERVATION_STAGES = new Set(Object.values(STAGES).filter(stage => stage !== STAGES.FINALIZADO));

const now = () => new Date().toISOString();
const text = value => String(value ?? '').trim();

async function readRows(sheet) {
  const response = await sheetsService.client.spreadsheets.values.get({
    spreadsheetId: sheetsService.spreadsheetId,
    range: `${sheet}!A2:Z`
  });
  return response.data.values || [];
}

function rowObject(fields, row, index) {
  return fields.reduce((obj, field, i) => {
    obj[field] = row[i] ?? '';
    return obj;
  }, { _row: index + 2 });
}

async function append(sheet, fields, data) {
  await sheetsService.client.spreadsheets.values.append({
    spreadsheetId: sheetsService.spreadsheetId,
    range: `${sheet}!A2:${String.fromCharCode(64 + Math.min(fields.length, 26))}`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    resource: {
      values: [[
        ...fields.map(field => {
          const value = data[field];
          return value !== null && typeof value === 'object' ? JSON.stringify(value) : value ?? '';
        })
      ]]
    }
  });
}

async function update(sheet, fields, rowNumber, data) {
  await sheetsService.client.spreadsheets.values.update({
    spreadsheetId: sheetsService.spreadsheetId,
    range: `${sheet}!A${rowNumber}:${String.fromCharCode(64 + Math.min(fields.length, 26))}${rowNumber}`,
    valueInputOption: 'RAW',
    resource: { values: [fields.map(field => data[field] ?? '')] }
  });
}

async function findStage(id) {
  const rows = await readRows('ETAPAS');
  const fields = sheetsService.fieldDefinitions.ETAPAS;
  const index = rows.findIndex(row => text(row[0]) === text(id));
  return index < 0 ? null : rowObject(fields, rows[index], index);
}

async function findUserEmail(userId) {
  const fields = sheetsService.fieldDefinitions.USUARIOS;
  const rows = (await readRows('USUARIOS')).map((row, i) => rowObject(fields, row, i));
  return text(rows.find(user => text(user.id_usuario) === text(userId))?.correo_usuario);
}

async function findRoleEmail(role) {
  const emails = await findRoleEmails(role);
  return emails[0] || '';
}

async function findRoleEmails(role) {
  const fields = sheetsService.fieldDefinitions.USUARIOS;
  const rows = (await readRows('USUARIOS')).map((row, i) => rowObject(fields, row, i));
  return rows
    .filter(user => text(user.role).toLowerCase() === role)
    .map(user => text(user.correo_usuario))
    .filter(Boolean);
}

function parseRecipients(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value.flatMap(parseRecipients);
  const raw = text(value);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed.flatMap(parseRecipients);
  } catch (error) {
    // Accept comma/semicolon-separated addresses for multipart clients.
  }
  return raw.split(/[;,]/).map(item => item.trim()).filter(Boolean);
}

function renderEmailTemplate(template, stage) {
  return text(template)
    .replaceAll('{{id_solicitud}}', text(stage.id_solicitud))
    .replaceAll('{{nombre_solicitante}}', text(stage.nombre_solicitante))
    .replaceAll('{{etapa}}', STAGES.DECANATO)
    .replaceAll('{{fecha}}', now());
}

async function notifyRecipients({ stage, event, recipients, subject, body }) {
  const uniqueRecipients = [...new Set(recipients.map(item => text(item).toLowerCase()).filter(Boolean))];
  if (!uniqueRecipients.length) {
    return notifySignatureEvent({ stage, event, recipient: '', subject, body });
  }
  const notifications = [];
  for (const recipient of uniqueRecipients) {
    notifications.push(await notifySignatureEvent({ stage, event, recipient, subject, body }));
  }
  return notifications;
}

async function logHistory(id, stage, action, previous, next, req, comment = '') {
  const fields = sheetsService.fieldDefinitions.HISTORIAL_ETAPAS;
  await append('HISTORIAL_ETAPAS', fields, {
    id_historial: await sheetsService.getLastId('HISTORIAL_ETAPAS') + 1,
    id_solicitud: id,
    etapa: stage,
    accion: action,
    estado_anterior: previous,
    estado_nuevo: next,
    rol_actor: req.user.role,
    id_usuario_actor: req.user.id,
    fecha: now(),
    comentario: comment
  });
}

function assertRole(req, allowed) {
  const role = text(req.user.role).toLowerCase() || 'solicitante';
  if (!allowed.includes(role) && role !== 'admin') {
    const error = new Error('El usuario no tiene el rol requerido');
    error.status = 403;
    throw error;
  }
  return role;
}

async function updateStage(stage, patch) {
  const fields = sheetsService.fieldDefinitions.ETAPAS;
  const nextStage = patch.etapa_codigo ?? stage.etapa_codigo;
  const data = {
    ...stage,
    ...patch,
    etapa_anterior: patch.etapa_anterior ?? (nextStage !== stage.etapa_codigo ? stage.etapa_codigo : stage.etapa_anterior),
    updated_at: now(),
    version: Number(stage.version || 0) + 1
  };
  await update('ETAPAS', fields, stage._row, data);
  return data;
}

function normalizeObservationStatus(value) {
  const status = text(value).toLowerCase();
  return ['pendiente', 'subsanada', 'anulada'].includes(status) ? status : 'pendiente';
}

function resetReturnedForms(stage) {
  let states;
  try {
    states = stage.estado_formularios ? JSON.parse(stage.estado_formularios) : {};
  } catch (error) {
    states = {};
  }
  const updated = { ...states };
  let changed = false;
  for (const [formulario, estado] of Object.entries(updated)) {
    if (text(estado).toLowerCase().replace('ó', 'o') === 'enviado a revision') {
      updated[formulario] = 'Completado';
      changed = true;
    }
  }
  return changed ? JSON.stringify(updated) : stage.estado_formularios;
}

function parseObservationStore(raw) {
  if (!raw) return {};
  try {
    const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
      if (parsed.etapas && typeof parsed.etapas === 'object') return parsed.etapas;
      return Object.fromEntries(Object.entries(parsed).filter(([, value]) => Array.isArray(value)));
    }
  } catch (error) {
    console.warn('No se pudo interpretar ETAPAS.comentarios como JSON:', error.message);
  }
  return {};
}

function observationId() {
  return `obs-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function observationStage(stage, requestedStage) {
  const selectedStage = text(requestedStage) || text(stage.etapa_codigo);
  if (!OBSERVATION_STAGES.has(selectedStage)) {
    const error = new Error('La etapa indicada no permite observaciones');
    error.status = 400;
    throw error;
  }
  if (selectedStage !== text(stage.etapa_codigo)) {
    const error = new Error('Las observaciones deben pertenecer a la etapa actual de la solicitud');
    error.status = 409;
    throw error;
  }
  return selectedStage;
}

async function saveObservationStore(stage, store) {
  const fields = sheetsService.fieldDefinitions.ETAPAS;
  const comments = JSON.stringify(store);
  return updateStage(stage, { comentarios: comments });
}

async function getObservationsForRequest(stage) {
  const store = parseObservationStore(stage.comentarios);
  return Object.entries(store).flatMap(([etapa, observations]) =>
    observations.map(item => ({ ...item, id_solicitud: stage.id_solicitud, etapa: item.etapa || etapa }))
  );
}

async function appendObservation(stage, req, payload) {
  const observation = text(payload.observacion || payload.texto);
  if (!observation) {
    const error = new Error('La observación no puede estar vacía');
    error.status = 400;
    throw error;
  }
  const data = {
    id_observacion: observationId(),
    id_solicitud: stage.id_solicitud,
    etapa: observationStage(stage, payload.etapa),
    formulario: text(payload.formulario),
    paso: text(payload.paso),
    observacion: observation,
    estado: 'pendiente',
    id_usuario_autor: req.user.id,
    nombre_autor: req.user.name,
    fecha_creacion: now(),
    fecha_actualizacion: now()
  };
  const store = parseObservationStore(stage.comentarios);
  if (!Array.isArray(store[data.etapa])) store[data.etapa] = [];
  store[data.etapa].push(data);
  const updatedStage = await saveObservationStore(stage, store);
  Object.assign(stage, updatedStage);
  return data;
}

async function recordNotification(id, event, recipient, subject, body) {
  const fields = sheetsService.fieldDefinitions.NOTIFICACIONES;
  const data = {
    id_notificacion: await sheetsService.getLastId('NOTIFICACIONES') + 1,
    id_solicitud: id,
    evento: event,
    estado: 'pendiente',
    intentos: 1,
    fecha_creacion: now(),
    fecha_envio: '',
    ultimo_error: ''
  };

  try {
    await emailService.send({ to: recipient, subject, text: body });
    data.estado = 'enviada';
    data.fecha_envio = now();
  } catch (error) {
    data.estado = 'fallida';
    data.ultimo_error = error.message;
    console.error(`No se pudo enviar notificación ${event} para ${id}:`, error);
  }

  await append('NOTIFICACIONES', fields, data);
  return data;
}

async function notifySignatureEvent({ stage, event, recipient, subject, body }) {
  if (!recipient) {
    console.warn(`No se envió ${event}: no se encontró correo destinatario para ${stage.id_solicitud}`);
    const fields = sheetsService.fieldDefinitions.NOTIFICACIONES;
    const data = {
      id_notificacion: await sheetsService.getLastId('NOTIFICACIONES') + 1,
      id_solicitud: stage.id_solicitud,
      evento: event,
      estado: 'sin_destinatario',
      intentos: 0,
      fecha_creacion: now(),
      fecha_envio: '',
      ultimo_error: 'No se encontró correo destinatario'
    };
    await append('NOTIFICACIONES', fields, data);
    return data;
  }
  return recordNotification(stage.id_solicitud, event, recipient, subject, body);
}

const getQueue = async (req, res, next) => {
  try {
    const role = assertRole(req, [...Object.keys(ROLE_STAGE), 'director']);
    const fields = sheetsService.fieldDefinitions.ETAPAS;
    const stages = (await readRows('ETAPAS')).map((row, i) => rowObject(fields, row, i));
    const result = stages
      .filter(item => (text(item.rol_responsable) === role ||
        (['revisor_1', 'revisor_2'].includes(role) && item.etapa_codigo === STAGES.DECANATO)) &&
        ['pendiente', 'en_revision', 'devuelta'].includes(text(item.estado_etapa)));
    if (role === 'director') {
      const signed = (await readRows('FIRMAS')).map((row, i) => rowObject(sheetsService.fieldDefinitions.FIRMAS, row, i))
        .filter(item => item.tipo_firma === 'director_escuela' && item.estado === 'firmada')
        .map(item => text(item.id_solicitud));
      result.push(...stages.filter(item => item.etapa_codigo === STAGES.FIRMAS &&
        !signed.includes(text(item.id_solicitud)) &&
        ['pendiente', 'en_revision'].includes(text(item.estado_etapa))));
      result.push(...stages.filter(item => signed.includes(text(item.id_solicitud)) &&
        !result.some(current => current.id_solicitud === item.id_solicitud)));
    }
    res.json({ success: true, role, data: result });
  } catch (error) { next(error); }
};

const getConsultationQueue = async (req, res, next) => {
  try {
    const role = assertRole(req, ['revisor_1', 'revisor_2', 'presupuesto', 'director']);
    const fields = sheetsService.fieldDefinitions.ETAPAS;
    const stages = (await readRows('ETAPAS')).map((row, i) => rowObject(fields, row, i));
    const cortes = (await readRows('CORTES_DECANATO')).map((row, i) =>
      rowObject(sheetsService.fieldDefinitions.CORTES_DECANATO, row, i));
    const latestCorteByRequest = new Map();
    cortes.forEach((corte) => {
      const id = text(corte.id_solicitud);
      const current = latestCorteByRequest.get(id);
      if (!current || Number(corte.id_corte) > Number(current.id_corte)) latestCorteByRequest.set(id, corte);
    });
    let data = [];
    if (['revisor_1', 'revisor_2'].includes(role)) {
      data = stages;
    } else if (role === 'director') {
      const signatures = (await readRows('FIRMAS')).map((row, i) => rowObject(sheetsService.fieldDefinitions.FIRMAS, row, i));
      const ids = new Set(signatures.filter(item => item.tipo_firma === 'director_escuela').map(item => text(item.id_solicitud)));
      data = stages.filter(item => item.etapa_codigo === STAGES.FIRMAS || ids.has(text(item.id_solicitud)));
    } else {
      const history = (await readRows('HISTORIAL_ETAPAS')).map((row, i) =>
        rowObject(sheetsService.fieldDefinitions.HISTORIAL_ETAPAS, row, i));
      const reviewedIds = new Set(history
        .filter(item => text(item.rol_actor).toLowerCase() === role)
        .map(item => text(item.id_solicitud)));
      data = stages.filter(item => reviewedIds.has(text(item.id_solicitud)) ||
        text(item.etapa_anterior).toLowerCase() === STAGES.PRESUPUESTO);
    }
    res.json({
      success: true,
      role,
      data: data.map(item => ({
        ...item,
        ...(latestCorteByRequest.has(text(item.id_solicitud))
          ? { corte_decanato: latestCorteByRequest.get(text(item.id_solicitud)) }
          : {}),
        puede_actuar: role === 'revisor_1'
          ? item.etapa_codigo === STAGES.REVISION_1
          : role === 'revisor_2'
            ? item.etapa_codigo === STAGES.REVISION_2
            : role === 'presupuesto'
              ? item.etapa_codigo === STAGES.PRESUPUESTO
              : item.etapa_codigo === STAGES.FIRMAS && !item.usuario_asignado,
        motivo_sin_permiso: 'La consulta es de solo lectura fuera de la etapa asignada'
      }))
    });
  } catch (error) { next(error); }
};

const getPendingSignatures = async (req, res, next) => {
  try {
    assertRole(req, ['director']);
    const stages = (await readRows('ETAPAS')).map((row, i) => rowObject(sheetsService.fieldDefinitions.ETAPAS, row, i))
      .filter(stage => stage.etapa_codigo === STAGES.FIRMAS && ['pendiente', 'en_revision'].includes(text(stage.estado_etapa)));
    const signatures = (await readRows('FIRMAS')).map((row, i) => rowObject(sheetsService.fieldDefinitions.FIRMAS, row, i));
    const data = stages.map(stage => {
      const signed = signatures.filter(item => text(item.id_solicitud) === text(stage.id_solicitud) && item.estado === 'firmada');
      return { ...stage, firma_solicitante: signed.some(item => item.tipo_firma === 'solicitante'), firma_director: signed.some(item => item.tipo_firma === 'director_escuela') };
    }).filter(stage => !stage.firma_director);
    res.json({ success: true, data });
  } catch (error) { next(error); }
};

const take = async (req, res, next) => {
  try {
    const role = assertRole(req, Object.keys(ROLE_STAGE));
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    const canTakeDecanato = stage.etapa_codigo === STAGES.DECANATO &&
      ['revisor_1', 'revisor_2'].includes(role);
    if (stage.rol_responsable !== role && !canTakeDecanato) {
      return res.status(403).json({ success: false, error: 'La solicitud no pertenece a esta cola' });
    }
    if (stage.usuario_asignado && stage.usuario_asignado !== req.user.id) return res.status(409).json({ success: false, error: 'La solicitud ya fue tomada por otro usuario' });
    const updated = await updateStage(stage, { usuario_asignado: req.user.id, estado_etapa: 'en_revision', fecha_inicio: stage.fecha_inicio || now() });
    await logHistory(stage.id_solicitud, stage.etapa_codigo, 'tomar', stage.estado_etapa, 'en_revision', req);
    res.json({ success: true, message: 'Solicitud tomada', data: updated });
  } catch (error) { next(error); }
};

const decide = async (req, res, next) => {
  try {
    const role = assertRole(req, Object.keys(DECISION_ROLE_STAGE));
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    if (stage.etapa_codigo !== DECISION_ROLE_STAGE[role]) return res.status(409).json({ success: false, error: 'La solicitud no está en la etapa correspondiente' });
    if (stage.usuario_asignado && stage.usuario_asignado !== req.user.id && role !== 'admin') return res.status(403).json({ success: false, error: 'La solicitud fue tomada por otro usuario' });
    const action = text(req.body.accion).toLowerCase();
    const comment = text(req.body.comentario);
    if (!['aprobar', 'devolver'].includes(action)) return res.status(400).json({ success: false, error: 'Acción inválida; el cierre se registra en Decanato' });
    if (action === 'devolver' && !comment) return res.status(400).json({ success: false, error: 'El comentario es obligatorio' });
    if (action === 'devolver') {
      const inlineObservations = Array.isArray(req.body.observaciones) ? req.body.observaciones : [];
      const existingObservations = parseObservationStore(stage.comentarios)[stage.etapa_codigo] || [];
      const hasObservation = inlineObservations.length > 0 ||
        existingObservations.some(item => normalizeObservationStatus(item.estado) === 'pendiente');
      if (!hasObservation) {
        return res.status(400).json({ success: false, error: 'Debe incluir al menos una observación para devolver la solicitud' });
      }
    }

    const nextStage = action === 'devolver' ?
      (stage.etapa_codigo === STAGES.PRESUPUESTO ? STAGES.REVISION_2 : STAGES.SOLICITUD) :
      ({ revision_1: STAGES.REVISION_2, revision_2: STAGES.PRESUPUESTO, presupuesto: STAGES.FIRMAS }[stage.etapa_codigo] || STAGES.FINALIZADO);
    const nextRole = nextStage === STAGES.SOLICITUD ? 'solicitante' :
      nextStage === STAGES.REVISION_2 ? 'revisor_2' :
      nextStage === STAGES.PRESUPUESTO ? 'presupuesto' :
      nextStage === STAGES.FIRMAS ? 'solicitante/director' : '';
    const newState = action === 'devolver' ? 'devuelta' : 'pendiente';
    const observations = [];
    if (Array.isArray(req.body.observaciones)) {
      for (const item of req.body.observaciones) {
        observations.push(await appendObservation(stage, req, {
          ...item,
          etapa: item.etapa || stage.etapa_codigo
        }));
      }
    }
    await append('REVISIONES', sheetsService.fieldDefinitions.REVISIONES, {
      id_revision: await sheetsService.getLastId('REVISIONES') + 1,
      id_solicitud: stage.id_solicitud, etapa: stage.etapa_codigo, rol_revisor: role,
      id_usuario_revisor: req.user.id, estado: action === 'aprobar' ? 'aprobada' : action,
      observaciones: comment, fecha_inicio: stage.fecha_inicio || now(), fecha_respuesta: now(),
      comentarios_por_formulario: req.body.comentarios_por_formulario || '', numero_intento: req.body.numero_intento || 1,
      id_usuario_toma: stage.usuario_asignado || req.user.id, fecha_toma: stage.fecha_inicio || now()
    });
    const updated = await updateStage(stage, {
      etapa_codigo: nextStage,
      etapa_actual: nextStage,
      rol_responsable: nextRole,
      usuario_asignado: '',
      estado_etapa: newState,
      fecha_fin: now(),
      comentarios: stage.comentarios
    });
    await logHistory(stage.id_solicitud, stage.etapa_codigo, action, stage.estado_etapa, newState, req, comment);
    res.json({
      success: true,
      data: {
        ...updated,
        observaciones: observations
      }
    });
  } catch (error) { next(error); }
};

const getObservations = async (req, res, next) => {
  try {
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    const observations = await getObservationsForRequest(stage);
    res.json({ success: true, data: observations });
  } catch (error) { next(error); }
};

const addObservation = async (req, res, next) => {
  try {
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    const role = text(req.user.role).toLowerCase() || 'solicitante';
    const allowed = ['revisor_1', 'revisor_2', 'presupuesto', 'decanato', 'solicitante', 'director', 'admin'];
    if (!allowed.includes(role)) return res.status(403).json({ success: false, error: 'Rol no autorizado para agregar observaciones' });
    const observation = await appendObservation(stage, req, {
      ...req.body,
      etapa: req.body.etapa || stage.etapa_codigo
    });
    res.status(201).json({ success: true, data: observation });
  } catch (error) { next(error); }
};

const updateObservation = async (req, res, next) => {
  try {
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    const store = parseObservationStore(stage.comentarios);
    let current;
    let currentStage;
    for (const [stageName, observations] of Object.entries(store)) {
      const found = observations.find(item => text(item.id_observacion) === text(req.params.idObservacion));
      if (found) {
        current = { ...found, id_solicitud: stage.id_solicitud, etapa: found.etapa || stageName };
        currentStage = stageName;
        break;
      }
    }
    if (!current) return res.status(404).json({ success: false, error: 'Observación no encontrada' });
    const role = text(req.user.role).toLowerCase() || 'solicitante';
    const isOwner = text(stage.id_usuario) === text(req.user.id);
    const isAuthor = text(current.id_usuario_autor) === text(req.user.id);
    const isStageReviewer = ROLE_STAGE[role] === currentStage;
    if (role !== 'admin' && !isAuthor && !isOwner && !isStageReviewer) {
      return res.status(403).json({ success: false, error: 'No puedes modificar esta observación' });
    }
    if (isOwner && !isAuthor && !isStageReviewer && req.body.observacion !== undefined) {
      return res.status(403).json({ success: false, error: 'El solicitante solo puede actualizar el estado de la observación' });
    }
    const next = {
      ...current,
      observacion: req.body.observacion === undefined ? current.observacion : text(req.body.observacion),
      estado: req.body.estado === undefined ? current.estado : normalizeObservationStatus(req.body.estado),
      fecha_actualizacion: now()
    };
    if (!next.observacion) return res.status(400).json({ success: false, error: 'La observación no puede estar vacía' });
    if (isOwner && !isAuthor && !isStageReviewer && next.estado !== 'subsanada') {
      return res.status(403).json({ success: false, error: 'El solicitante solo puede marcar la observación como subsanada' });
    }
    const observations = store[currentStage];
    const itemIndex = observations.findIndex(item => text(item.id_observacion) === text(req.params.idObservacion));
    observations[itemIndex] = {
      ...observations[itemIndex],
      observacion: next.observacion,
      estado: next.estado,
      fecha_actualizacion: next.fecha_actualizacion
    };
    await saveObservationStore(stage, store);
    const pendingObservations = Object.values(store)
      .filter(Array.isArray)
      .flat()
      .filter(item => normalizeObservationStatus(item.estado) === 'pendiente');
    if (isOwner && pendingObservations.length === 0) {
      const estadoFormularios = resetReturnedForms(stage);
      if (estadoFormularios !== stage.estado_formularios) {
        await updateStage(stage, { estado_formularios: estadoFormularios });
      }
    }
    res.json({ success: true, data: next });
  } catch (error) { next(error); }
};

const sign = async (req, res, next) => {
  try {
    const role = text(req.user.role).toLowerCase() || 'solicitante';
    if (!['solicitante', 'director'].includes(role)) return res.status(403).json({ success: false, error: 'Rol no autorizado para firmar' });
    const stage = await findStage(req.params.idSolicitud);
    if (!stage || stage.etapa_codigo !== STAGES.FIRMAS) return res.status(409).json({ success: false, error: 'La solicitud no está pendiente de firmas' });
    if (role === 'solicitante' && text(stage.id_usuario) !== text(req.user.id)) return res.status(403).json({ success: false, error: 'Solo el solicitante puede registrar esta firma' });
    const type = role === 'director' ? 'director_escuela' : 'solicitante';
    const fields = sheetsService.fieldDefinitions.FIRMAS;
    const existing = (await readRows('FIRMAS')).map((row, i) => rowObject(fields, row, i))
      .find(item => text(item.id_solicitud) === text(req.params.idSolicitud) && item.tipo_firma === type && item.estado === 'firmada');
    if (existing) return res.status(409).json({ success: false, error: 'Esta firma ya fue registrada' });
    if (!req.file) return res.status(400).json({ success: false, error: 'Debe cargar la imagen de la firma' });
    if (!/^image\/(png|jpeg|webp)$/.test(req.file.mimetype)) return res.status(400).json({ success: false, error: 'La firma debe ser una imagen PNG, JPEG o WEBP' });

    const extension = req.file.originalname.split('.').pop().toLowerCase();
    const evidence = await driveService.uploadFile(
      req.file,
      process.env.DRIVE_UPLOADS_FOLDER,
      `firma-${req.params.idSolicitud}-${type}-${req.user.id}-${Date.now()}.${extension}`
    );
    fs.unlink(req.file.path, () => {});
    const signature = {
      id_firma: await sheetsService.getLastId('FIRMAS') + 1, id_solicitud: req.params.idSolicitud,
      tipo_firma: type, id_firmante: req.user.id, nombre_firmante: req.user.name,
      estado: 'firmada', fecha_firma: now(), evidencia_firma: evidence,
      observaciones: text(req.body.observaciones), fecha_solicitud_firma: text(req.body.fecha_solicitud_firma) || now()
    };
    await append('FIRMAS', fields, signature);
    const signatures = (await readRows('FIRMAS')).map((row, i) => rowObject(fields, row, i))
      .filter(item => text(item.id_solicitud) === text(req.params.idSolicitud) && item.estado === 'firmada');
    const bothSigned = signatures.some(item => item.tipo_firma === 'solicitante') && signatures.some(item => item.tipo_firma === 'director_escuela');
    const updated = bothSigned ? await updateStage(stage, {
      etapa_codigo: STAGES.DECANATO, etapa_actual: STAGES.DECANATO, rol_responsable: 'decanato',
      estado_etapa: 'pendiente', usuario_asignado: '', fecha_fin: now()
    }) : await updateStage(stage, { estado_etapa: 'en_revision' });
    await logHistory(stage.id_solicitud, STAGES.FIRMAS, 'firmar', stage.estado_etapa, bothSigned ? 'completas' : 'parcial', req);

    const recipient = type === 'solicitante' ? await findRoleEmail('director') : await findUserEmail(stage.id_usuario);
    const notification = await notifySignatureEvent({
      stage,
      event: 'firma_registrada',
      recipient,
      subject: `Firma registrada - solicitud ${stage.id_solicitud}`,
      body: `Se registró una firma en la solicitud ${stage.id_solicitud}. ${bothSigned ? 'La solicitud pasó a Decanato.' : 'Aún falta una firma.'}`
    });
    let decanatoNotifications = [];
    if (bothSigned) {
      const defaultSubject = `Solicitud enviada a Decanato - ${stage.id_solicitud}`;
      const defaultBody = `La solicitud ${stage.id_solicitud} de ${stage.nombre_solicitante || 'un solicitante'} ya cuenta con ambas firmas y está pendiente de revisión por Decanato.`;
      const subject = text(req.body.asunto_decanato) || defaultSubject;
      const body = renderEmailTemplate(text(req.body.cuerpo_decanato) || defaultBody, stage);
      const recipients = [
        ...(await findRoleEmails('decanato')),
        ...parseRecipients(req.body.destinatarios_decanato)
      ];
      decanatoNotifications = await notifyRecipients({
        stage,
        event: 'solicitud_enviada_decanato',
        recipients,
        subject,
        body
      });
    }
    res.json({
      success: true,
      data: { firma: signature, firmas: signatures, etapa: updated, notificacion: notification, notificaciones_decanato: decanatoNotifications }
    });
  } catch (error) {
    if (req.file?.path) fs.unlink(req.file.path, () => {});
    next(error);
  }
};

const remindSignature = async (req, res, next) => {
  try {
    const role = assertRole(req, ['solicitante', 'director', 'admin']);
    const stage = await findStage(req.params.idSolicitud);
    if (!stage || stage.etapa_codigo !== STAGES.FIRMAS) return res.status(409).json({ success: false, error: 'La solicitud no está pendiente de firmas' });
    const target = text(req.body.tipo_firma);
    if (!['solicitante', 'director_escuela'].includes(target)) return res.status(400).json({ success: false, error: 'tipo_firma debe ser solicitante o director_escuela' });
    if (role === 'solicitante' && target !== 'director_escuela') return res.status(403).json({ success: false, error: 'El solicitante solo puede recordar al director' });
    if (role === 'director' && target !== 'solicitante') return res.status(403).json({ success: false, error: 'El director solo puede recordar al solicitante' });
    const recipient = target === 'solicitante' ? await findUserEmail(stage.id_usuario) : await findRoleEmail('director');
    const notification = await notifySignatureEvent({
      stage,
      event: target === 'solicitante' ? 'recordatorio_firma_solicitante' : 'recordatorio_firma_director',
      recipient,
      subject: `Recordatorio de firma - solicitud ${stage.id_solicitud}`,
      body: `Hay una firma pendiente para la solicitud ${stage.id_solicitud}.`
    });
    res.json({ success: true, data: notification });
  } catch (error) { next(error); }
};

const createCorte = async (req, res, next) => {
  try {
    const role = assertRole(req, ['decanato', 'revisor_1', 'revisor_2']);
    const { id_solicitud, estado_corte = 'finalizado', resultado = '', observaciones = '', resolucion = '' } = req.body;
    if (!id_solicitud || !resultado) return res.status(400).json({ success: false, error: 'Se requieren id_solicitud y resultado' });
    const stage = await findStage(id_solicitud);
    if (!stage || stage.etapa_codigo !== STAGES.DECANATO) return res.status(409).json({ success: false, error: 'La solicitud no está en Decanato' });
    if (stage.usuario_asignado && stage.usuario_asignado !== req.user.id && role !== 'admin') {
      return res.status(403).json({ success: false, error: 'La solicitud fue tomada por otro usuario' });
    }
    if (resultado.toLowerCase() === 'aprobada' && !req.file) return res.status(400).json({ success: false, error: 'La resolución PDF es obligatoria para una decisión aprobada' });
    if (req.file && req.file.mimetype !== 'application/pdf') return res.status(400).json({ success: false, error: 'La resolución debe ser un archivo PDF' });
    let resolution = text(resolucion);
    if (req.file) {
      resolution = await driveService.uploadFile(
        req.file,
        process.env.DRIVE_UPLOADS_FOLDER,
        `resolucion-${id_solicitud}-${Date.now()}.pdf`
      );
      fs.unlink(req.file.path, () => {});
    }
    const data = {
      id_corte: await sheetsService.getLastId('CORTES_DECANATO') + 1,
      fecha_corte: now(), id_solicitud, estado_corte,
      id_usuario_envia: req.user.id, fecha_envio: now(), resultado, observaciones,
      resolucion, estado_final: resultado
    };
    await append('CORTES_DECANATO', sheetsService.fieldDefinitions.CORTES_DECANATO, data);
    const updated = await updateStage(stage, {
      etapa_codigo: STAGES.FINALIZADO, etapa_actual: STAGES.FINALIZADO, rol_responsable: '',
      estado_etapa: 'finalizado', fecha_fin: now(), comentarios: observaciones
    });
    await logHistory(id_solicitud, STAGES.DECANATO, 'registrar_decision', stage.estado_etapa, 'finalizado', req, observaciones);
    res.status(201).json({ success: true, data: { ...data, etapa: updated } });
  } catch (error) {
    if (req.file?.path) fs.unlink(req.file.path, () => {});
    next(error);
  }
};

module.exports = {
  getQueue,
  getConsultationQueue,
  getPendingSignatures,
  take,
  decide,
  getObservations,
  addObservation,
  updateObservation,
  sign,
  remindSignature,
  createCorte
};
