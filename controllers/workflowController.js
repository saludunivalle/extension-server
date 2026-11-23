const fs = require('fs');
const sheetsService = require('../services/sheetsService');
const driveService = require('../services/driveService');
const emailService = require('../services/emailService');
const workflowNotificationService = require('../services/workflowNotificationService');
const { createDecanatoAttachments, generateFormPdfBuffer, getSignatureBuffer } = require('../services/workflowAttachmentService');

const STAGES = Object.freeze({
  SOLICITUD: 'solicitud',
  REVISION_1: 'revision_1',
  REVISION_2: 'revision_2',
  PRESUPUESTO: 'presupuesto',
  FIRMAS: 'firmas',
  LISTA_PARA_DECANATO: 'lista_para_decanato',
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

const rowsCache = new Map();
const ROWS_CACHE_TTL = 15000; // 15 segundos de caché en memoria

function invalidateRowsCache(sheet) {
  if (sheet) {
    rowsCache.delete(sheet);
  } else {
    rowsCache.clear();
  }
}

async function readRows(sheet) {
  const cached = rowsCache.get(sheet);
  const currentTime = Date.now();
  if (cached && currentTime < cached.expiresAt) {
    return cached.data;
  }
  const response = await sheetsService.client.spreadsheets.values.get({
    spreadsheetId: sheetsService.spreadsheetId,
    range: `${sheet}!A2:Z`
  });
  const data = response.data.values || [];
  rowsCache.set(sheet, { data, expiresAt: currentTime + ROWS_CACHE_TTL });
  return data;
}

async function readMultipleSheets(sheetNames) {
  const currentTime = Date.now();
  const missing = sheetNames.filter(name => {
    const cached = rowsCache.get(name);
    return !cached || currentTime >= cached.expiresAt;
  });

  if (missing.length > 0) {
    try {
      const ranges = missing.map(name => `${name}!A2:Z`);
      const response = await sheetsService.client.spreadsheets.values.batchGet({
        spreadsheetId: sheetsService.spreadsheetId,
        ranges
      });
      (response.data.valueRanges || []).forEach((rangeObj, index) => {
        const sheetName = missing[index];
        const data = rangeObj.values || [];
        rowsCache.set(sheetName, { data, expiresAt: currentTime + ROWS_CACHE_TTL });
      });
    } catch (batchError) {
      console.warn('Error en batchGet, reintentando individualmente:', batchError.message);
      for (const name of missing) {
        await readRows(name);
      }
    }
  }

  const result = {};
  for (const name of sheetNames) {
    result[name] = rowsCache.get(name)?.data || [];
  }
  return result;
}

async function readDependencyByRequest() {
  const fields = sheetsService.fieldDefinitions.SOLICITUDES;
  const rows = await readRows('SOLICITUDES');
  return new Map(
    rows.map((row, index) => {
      const solicitud = rowObject(fields, row, index);
      return [text(solicitud.id_solicitud), text(solicitud.nombre_dependencia)];
    }),
  );
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
  invalidateRowsCache(sheet);
}

async function update(sheet, fields, rowNumber, data) {
  await sheetsService.client.spreadsheets.values.update({
    spreadsheetId: sheetsService.spreadsheetId,
    range: `${sheet}!A${rowNumber}:${String.fromCharCode(64 + Math.min(fields.length, 26))}${rowNumber}`,
    valueInputOption: 'RAW',
    resource: { values: [fields.map(field => data[field] ?? '')] }
  });
  invalidateRowsCache(sheet);
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

async function notifyRecipients({ stage, event, recipients, subject, body, html, attachments = [] }) {
  const uniqueRecipients = [...new Set(recipients.map(item => text(item).toLowerCase()).filter(Boolean))];
  if (!uniqueRecipients.length) {
    return notifySignatureEvent({ stage, event, recipient: '', subject, body, html, attachments });
  }
  const notifications = [];
  for (const recipient of uniqueRecipients) {
    notifications.push(await notifySignatureEvent({ stage, event, recipient, subject, body, html, attachments }));
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
  let role = text(req.user.role).toLowerCase() || 'solicitante';
  if (role === 'director_escuela') role = 'director';
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
  return ['pendiente', 'corregida', 'subsanada', 'anulada'].includes(status) ? status : 'pendiente';
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

function getObservationsList(comentarios, idSolicitud) {
  const store = parseObservationStore(comentarios);
  return Object.entries(store).flatMap(([etapa, observations]) =>
    (Array.isArray(observations) ? observations : []).map(item => ({
      ...item,
      id_solicitud: text(idSolicitud),
      etapa: item.etapa || etapa
    }))
  );
}

async function getObservationsForRequest(stage) {
  return getObservationsList(stage.comentarios, stage.id_solicitud);
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

async function recordNotification(id, event, recipient, subject, body, attachments = []) {
  return workflowNotificationService.recordNotification({
    idSolicitud: id,
    event,
    recipient,
    subject,
    text: body,
    attachments
  });
}

async function notifySignatureEvent({ stage, event, recipient, subject, body, html, attachments = [] }) {
  return workflowNotificationService.recordNotification({
    idSolicitud: stage.id_solicitud,
    event,
    recipient,
    subject,
    text: body,
    html,
    attachments
  });
}

function buildUserDirectory(users = []) {
  return new Map(
    users.map((user) => [
      text(user.id_usuario),
      {
        email: text(user.correo_usuario || user.email),
        name: text(user.nombre_usuario || user.name || user.nombre),
      },
    ])
  );
}

function enrichHistoryEntries(entries = [], userById) {
  const enriched = entries.map((entry) => {
    const actor = userById.get(text(entry.id_usuario_actor));
    return {
      ...entry,
      email_actor: actor?.email || 'Sin email disponible',
      nombre_actor: actor?.name || actor?.email || 'Usuario sin nombre',
    };
  });
  return enriched.sort((a, b) => {
    const dateA = a?.fecha ? new Date(a.fecha).getTime() : 0;
    const dateB = b?.fecha ? new Date(b.fecha).getTime() : 0;
    if (!isNaN(dateA) && !isNaN(dateB) && dateA !== dateB) {
      return dateB - dateA;
    }
    const idA = Number(a?.id_historial) || 0;
    const idB = Number(b?.id_historial) || 0;
    if (idA && idB && idA !== idB) {
      return idB - idA;
    }
    return 0;
  });
}

function buildSignatureSummary(requestSignatures = [], defaultDirectorName = '') {
  const solSig = requestSignatures.find(
    (signature) =>
      text(signature.tipo_firma).toLowerCase() === 'solicitante' &&
      text(signature.estado).toLowerCase() === 'firmada'
  );
  const dirSig = requestSignatures.find(
    (signature) =>
      ['director_escuela', 'director'].includes(text(signature.tipo_firma).toLowerCase()) &&
      text(signature.estado).toLowerCase() === 'firmada'
  );
  return {
    firma_solicitante: Boolean(solSig),
    firma_director: Boolean(dirSig),
    firmas_completas: Boolean(solSig && dirSig),
    nombre_solicitante: solSig ? text(solSig.nombre_firmante) : '',
    nombre_director: dirSig ? text(dirSig.nombre_firmante) : defaultDirectorName,
    firma_solicitante_info: solSig
      ? {
          nombre_firmante: text(solSig.nombre_firmante),
          fecha_firma: text(solSig.fecha_firma),
          evidencia_firma: text(solSig.evidencia_firma),
        }
      : null,
    firma_director_info: dirSig
      ? {
          nombre_firmante: text(dirSig.nombre_firmante),
          fecha_firma: text(dirSig.fecha_firma),
          evidencia_firma: text(dirSig.evidencia_firma),
        }
      : null,
  };
}

const getQueue = async (req, res, next) => {
  try {
    const role = assertRole(req, [...Object.keys(ROLE_STAGE), 'director']);
    const batch = await readMultipleSheets(['ETAPAS', 'SOLICITUDES', 'FIRMAS', 'HISTORIAL_ETAPAS', 'USUARIOS']);
    const fields = sheetsService.fieldDefinitions.ETAPAS;
    const stages = batch['ETAPAS'].map((row, i) => rowObject(fields, row, i));
    const dependencyByRequest = new Map(
      batch['SOLICITUDES'].map((row, index) => {
        const solicitud = rowObject(sheetsService.fieldDefinitions.SOLICITUDES, row, index);
        return [text(solicitud.id_solicitud), text(solicitud.nombre_dependencia)];
      })
    );
    const solicitanteByRequest = new Map(
      batch['SOLICITUDES'].map((row, index) => {
        const solicitud = rowObject(sheetsService.fieldDefinitions.SOLICITUDES, row, index);
        return [text(solicitud.id_solicitud), text(solicitud.nombre_solicitante)];
      })
    );
    const signatures = batch['FIRMAS'].map((row, i) =>
      rowObject(sheetsService.fieldDefinitions.FIRMAS, row, i));
    const signaturesByRequest = new Map();
    signatures.forEach((signature) => {
      const id = text(signature.id_solicitud);
      const current = signaturesByRequest.get(id) || [];
      current.push(signature);
      signaturesByRequest.set(id, current);
    });
    const history = batch['HISTORIAL_ETAPAS'].map((row, i) =>
      rowObject(sheetsService.fieldDefinitions.HISTORIAL_ETAPAS, row, i));
    const users = batch['USUARIOS'].map((row, i) =>
      rowObject(sheetsService.fieldDefinitions.USUARIOS, row, i));
    const defaultDirectorUser = users.find(u =>
      ['director', 'director_escuela'].includes(text(u.role).toLowerCase()) ||
      text(u.role).toLowerCase().includes('director')
    );
    const defaultDirectorName = defaultDirectorUser ? text(defaultDirectorUser.nombre_usuario) : '';
    const userById = buildUserDirectory(users);
    const historyByRequest = new Map();
    enrichHistoryEntries(history, userById).forEach((entry) => {
      const id = text(entry.id_solicitud);
      const current = historyByRequest.get(id) || [];
      current.push(entry);
      historyByRequest.set(id, current);
    });
    const result = stages
      .filter(item => (text(item.rol_responsable) === role ||
        (['revisor_1', 'revisor_2'].includes(role) && item.etapa_codigo === STAGES.DECANATO) ||
        (['presupuesto', 'revisor_2'].includes(role) && item.etapa_codigo === STAGES.LISTA_PARA_DECANATO)) &&
        ['pendiente', 'en_revision', 'devuelta'].includes(text(item.estado_etapa)));
    if (role === 'director') {
      const signed = signatures
        .filter(item => ['director_escuela', 'director'].includes(text(item.tipo_firma).toLowerCase()) && text(item.estado).toLowerCase() === 'firmada')
        .map(item => text(item.id_solicitud));
      result.push(...stages.filter(item => item.etapa_codigo === STAGES.FIRMAS &&
        !signed.includes(text(item.id_solicitud)) &&
        ['pendiente', 'en_revision'].includes(text(item.estado_etapa))));
      result.push(...stages.filter(item => signed.includes(text(item.id_solicitud)) &&
        !result.some(current => current.id_solicitud === item.id_solicitud)));
    }
    res.json({
      success: true,
      role,
      data: result.map((item) => {
        let storedNotes = '';
        try {
          const parsed = typeof item.comentarios === 'string' ? JSON.parse(item.comentarios) : item.comentarios;
          storedNotes = parsed?.notas_presupuesto_pdf || parsed?.notas_decanato_pdf || '';
        } catch (e) {}
        const sigSummary = buildSignatureSummary(signaturesByRequest.get(text(item.id_solicitud)) || [], defaultDirectorName);
        return {
          ...item,
          notas_presupuesto_pdf: storedNotes,
          nombre_dependencia:
            dependencyByRequest.get(text(item.id_solicitud)) ||
            text(item.nombre_dependencia),
          nombre_solicitante:
            sigSummary.nombre_solicitante ||
            solicitanteByRequest.get(text(item.id_solicitud)) ||
            text(item.nombre_solicitante),
          nombre_director:
            sigSummary.nombre_director ||
            defaultDirectorName,
          ...sigSummary,
          historial: historyByRequest.get(text(item.id_solicitud)) || [],
          observaciones: getObservationsList(item.comentarios, item.id_solicitud)
        };
      })
    });
  } catch (error) { next(error); }
};

const getConsultationQueue = async (req, res, next) => {
  try {
    const role = assertRole(req, ['revisor_1', 'revisor_2', 'presupuesto', 'director', 'admin']);
    const batch = await readMultipleSheets(['ETAPAS', 'SOLICITUDES', 'CORTES_DECANATO', 'FIRMAS', 'HISTORIAL_ETAPAS', 'USUARIOS']);
    const fields = sheetsService.fieldDefinitions.ETAPAS;
    const stages = batch['ETAPAS'].map((row, i) => rowObject(fields, row, i));
    const dependencyByRequest = new Map(
      batch['SOLICITUDES'].map((row, index) => {
        const solicitud = rowObject(sheetsService.fieldDefinitions.SOLICITUDES, row, index);
        return [text(solicitud.id_solicitud), text(solicitud.nombre_dependencia)];
      })
    );
    const solicitanteByRequest = new Map(
      batch['SOLICITUDES'].map((row, index) => {
        const solicitud = rowObject(sheetsService.fieldDefinitions.SOLICITUDES, row, index);
        return [text(solicitud.id_solicitud), text(solicitud.nombre_solicitante)];
      })
    );
    const cortes = batch['CORTES_DECANATO'].map((row, i) =>
      rowObject(sheetsService.fieldDefinitions.CORTES_DECANATO, row, i));
    const latestCorteByRequest = new Map();
    cortes.forEach((corte) => {
      const id = text(corte.id_solicitud);
      const current = latestCorteByRequest.get(id);
      if (!current || Number(corte.id_corte) > Number(current.id_corte)) latestCorteByRequest.set(id, corte);
    });
    const signatures = batch['FIRMAS'].map((row, i) =>
      rowObject(sheetsService.fieldDefinitions.FIRMAS, row, i));
    const history = batch['HISTORIAL_ETAPAS'].map((row, i) =>
      rowObject(sheetsService.fieldDefinitions.HISTORIAL_ETAPAS, row, i));
    const signaturesByRequest = new Map();
    signatures.forEach((signature) => {
      const id = text(signature.id_solicitud);
      const current = signaturesByRequest.get(id) || [];
      current.push(signature);
      signaturesByRequest.set(id, current);
    });
    const historyByRequest = new Map();
    history.forEach((entry) => {
      const id = text(entry.id_solicitud);
      const current = historyByRequest.get(id) || [];
      current.push(entry);
      historyByRequest.set(id, current);
    });
    const userFields = sheetsService.fieldDefinitions.USUARIOS;
    const users = batch['USUARIOS'].map((row, i) => rowObject(userFields, row, i));
    const defaultDirectorUser = users.find(u =>
      ['director', 'director_escuela'].includes(text(u.role).toLowerCase()) ||
      text(u.role).toLowerCase().includes('director')
    );
    const defaultDirectorName = defaultDirectorUser ? text(defaultDirectorUser.nombre_usuario) : '';
    const userById = buildUserDirectory(users);
    let data = [];
    if (role === 'admin' || ['revisor_1', 'revisor_2'].includes(role)) {
      data = stages;
    } else if (role === 'director') {
      const ids = new Set(signatures
        .filter(item => ['director_escuela', 'director'].includes(text(item.tipo_firma).toLowerCase()) && text(item.estado).toLowerCase() === 'firmada')
        .map(item => text(item.id_solicitud)));
      data = stages.filter(item => item.etapa_codigo === STAGES.FIRMAS || ids.has(text(item.id_solicitud)));
    } else {
      const reviewedIds = new Set(history
        .filter(item => text(item.rol_actor).toLowerCase() === role)
        .map(item => text(item.id_solicitud)));
      data = stages.filter(item => (
        // La solicitud actualmente asignada a Presupuesto o Lista para Decanato debe aparecer,
        // aunque todavía no tenga historial de una acción de ese rol.
        text(item.etapa_codigo).toLowerCase() === STAGES.PRESUPUESTO
        || text(item.etapa_codigo).toLowerCase() === STAGES.LISTA_PARA_DECANATO
        || text(item.rol_responsable).toLowerCase() === role
        || reviewedIds.has(text(item.id_solicitud))
        || text(item.etapa_anterior).toLowerCase() === STAGES.PRESUPUESTO
      ));
    }
    res.json({
      success: true,
      role,
      data: data.map(item => {
        let storedNotes = '';
        try {
          const parsed = typeof item.comentarios === 'string' ? JSON.parse(item.comentarios) : item.comentarios;
          storedNotes = parsed?.notas_presupuesto_pdf || parsed?.notas_decanato_pdf || '';
        } catch (e) {}
        const sigSummary = buildSignatureSummary(signaturesByRequest.get(text(item.id_solicitud)) || [], defaultDirectorName);
        return {
          ...item,
          notas_presupuesto_pdf: storedNotes,
          nombre_dependencia:
            dependencyByRequest.get(text(item.id_solicitud)) ||
            text(item.nombre_dependencia),
          nombre_solicitante:
            sigSummary.nombre_solicitante ||
            solicitanteByRequest.get(text(item.id_solicitud)) ||
            text(item.nombre_solicitante),
          nombre_director:
            sigSummary.nombre_director ||
            defaultDirectorName,
          ...sigSummary,
          historial: enrichHistoryEntries(historyByRequest.get(text(item.id_solicitud)) || [], userById),
          ...(latestCorteByRequest.has(text(item.id_solicitud))
            ? { corte_decanato: latestCorteByRequest.get(text(item.id_solicitud)) }
            : {}),
          observaciones: getObservationsList(item.comentarios, item.id_solicitud),
          puede_actuar: role === 'admin'
          ? false
          : role === 'revisor_1'
          ? item.etapa_codigo === STAGES.REVISION_1
          : role === 'revisor_2'
            ? (item.etapa_codigo === STAGES.REVISION_2 || item.etapa_codigo === STAGES.LISTA_PARA_DECANATO)
            : role === 'presupuesto'
              ? (item.etapa_codigo === STAGES.PRESUPUESTO || item.etapa_codigo === STAGES.LISTA_PARA_DECANATO)
              : role === 'director'
                ? item.etapa_codigo === STAGES.FIRMAS
                : item.etapa_codigo === STAGES.FIRMAS,
        motivo_sin_permiso: role === 'admin'
          ? 'La vista global de administración es de solo lectura'
          : 'La consulta es de solo lectura fuera de la etapa asignada'
      };
    })
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
    const canTakeListaDecanato = stage.etapa_codigo === STAGES.LISTA_PARA_DECANATO &&
      ['presupuesto', 'revisor_2'].includes(role);
    if (stage.rol_responsable !== role && !canTakeDecanato && !canTakeListaDecanato) {
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
    const validActions = ['aprobar', 'devolver', 'aprobar_firmas', 'enviar_revision_2'];
    if (!validActions.includes(action)) return res.status(400).json({ success: false, error: 'Acción inválida; el cierre se registra en Decanato' });

    const storeObservations = Object.values(parseObservationStore(stage.comentarios)).flat();
    const pendingObservations = storeObservations.filter(item => normalizeObservationStatus(item.estado) === 'pendiente');

    if (action === 'devolver') {
      const inlineObservations = Array.isArray(req.body.observaciones) ? req.body.observaciones : [];
      const hasObservation = inlineObservations.length > 0 ||
        pendingObservations.length > 0 ||
        Boolean(comment);
      if (!hasObservation) {
        return res.status(400).json({ success: false, error: 'Debe incluir al menos una observación o motivo para devolver la solicitud' });
      }
    }

    let nextStage;
    let nextRole;
    let newState = action === 'devolver' ? 'devuelta' : 'pendiente';

    if (action === 'devolver') {
      if (role === 'revisor_2') {
        nextStage = STAGES.REVISION_1;
        nextRole = 'revisor_1';
      } else {
        nextStage = STAGES.SOLICITUD;
        nextRole = 'solicitante';
      }
    } else {
      if (role === 'revisor_2') {
        if (pendingObservations.length > 0) {
          return res.status(400).json({
            success: false,
            error: 'No se puede aprobar la solicitud porque aún hay observaciones pendientes. Todas las observaciones deben estar marcadas como subsanadas antes de aprobar.'
          });
        }
        nextStage = STAGES.REVISION_1;
        nextRole = 'revisor_1';
        newState = 'pendiente';
      } else if (role === 'revisor_1') {
        const isFirmas = action === 'aprobar_firmas' || Boolean(req.body.enviar_firmas) || (stage.etapa_anterior === STAGES.REVISION_2 && action !== 'enviar_revision_2');
        if (isFirmas) {
          if (pendingObservations.length > 0) {
            return res.status(400).json({
              success: false,
              error: 'No se puede avanzar a firmas porque aún hay observaciones pendientes por subsanar.'
            });
          }
          nextStage = STAGES.FIRMAS;
          nextRole = 'solicitante/director';
          newState = 'pendiente';
        } else {
          nextStage = STAGES.REVISION_2;
          nextRole = 'revisor_2';
          newState = 'pendiente';
        }
      } else {
        nextStage = ({
          firmas: STAGES.PRESUPUESTO,
          presupuesto: STAGES.LISTA_PARA_DECANATO,
          lista_para_decanato: STAGES.DECANATO
        }[stage.etapa_codigo] || STAGES.FINALIZADO);
        nextRole = nextStage === STAGES.PRESUPUESTO ? 'presupuesto' :
          nextStage === STAGES.LISTA_PARA_DECANATO ? 'presupuesto' :
          nextStage === STAGES.DECANATO ? 'decanato' : '';
        newState = 'pendiente';
      }
    }
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
    let comentariosActualizados = stage.comentarios;
    if (stage.etapa_codigo === STAGES.PRESUPUESTO && action === 'aprobar') {
      let store = {};
      try {
        store = typeof stage.comentarios === 'string' ? JSON.parse(stage.comentarios || '{}') : (stage.comentarios || {});
      } catch {
        store = {};
      }
      const incomingNotes = text(req.body.notas_pdf ?? req.body.notasPdf ?? req.body.notas ?? '');
      store.notas_presupuesto_pdf = incomingNotes;
      store.incluir_historial_pdf = true;
      comentariosActualizados = JSON.stringify(store);
    }

    const updated = await updateStage(stage, {
      etapa_codigo: nextStage,
      etapa_actual: nextStage,
      rol_responsable: nextRole,
      usuario_asignado: '',
      estado_etapa: newState,
      fecha_fin: now(),
      comentarios: comentariosActualizados
    });
    await logHistory(stage.id_solicitud, stage.etapa_codigo, action, stage.estado_etapa, newState, req, comment);

    let notifications = [];
    // "enviar_revision_2" es la acción que usa la interfaz cuando Revisor 1
    // remite explícitamente la solicitud a la bandeja de Revisor 2. También
    // debe disparar la notificación de nueva solicitud.
    if (action === 'aprobar' || action === 'enviar_revision_2') {
      try {
        if (nextStage === STAGES.DECANATO) {
          // Ya no se envía correo a Decanato; el paso se gestiona mediante la carga en plataforma externa
        } else {
          notifications = await workflowNotificationService.notifyStageApproved({
            stage: updated,
            previousStage: stage.etapa_codigo,
            nextStage,
            nextRole,
            comment,
            approvedBy: req.user
          });
        }
      } catch (notifyErr) {
        console.warn('Error al enviar notificaciones de aprobación:', notifyErr.message);
      }
    } else if (action === 'devolver') {
      try {
        notifications = await workflowNotificationService.notifyStageReturned({
          stage: updated,
          fromStage: stage.etapa_codigo,
          toStage: nextStage,
          returnRole: nextRole,
          comment,
          observations: observations.length > 0 ? observations : (req.body.observaciones || []),
          returnedBy: req.user
        });
      } catch (notifyErr) {
        console.warn('Error al enviar notificaciones de devolución:', notifyErr.message);
      }
    }

    res.json({
      success: true,
      data: {
        ...updated,
        observaciones: observations,
        notificaciones: notifications
      }
    });
  } catch (error) { next(error); }
};

const getObservations = async (req, res, next) => {
  try {
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    const observations = await getObservationsForRequest(stage);
    res.json({ success: true, data: observations, etapa_actual: stage.etapa_codigo });
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
    const isStageReviewer = ROLE_STAGE[role] === currentStage || ROLE_STAGE[role] === stage.etapa_codigo;
    const isRevisor2InStage = role === 'revisor_2' && stage.etapa_codigo === STAGES.REVISION_2;
    if (role !== 'admin' && !isAuthor && !isOwner && !isStageReviewer && !isRevisor2InStage) {
      return res.status(403).json({ success: false, error: 'No puedes modificar esta observación' });
    }
    if (isOwner && !isAuthor && !isStageReviewer && (stage.etapa_codigo !== STAGES.SOLICITUD || stage.estado_etapa !== 'devuelta')) {
      return res.status(403).json({ success: false, error: 'No se pueden modificar observaciones una vez reenviada la solicitud' });
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
    if (isOwner && !isAuthor && !isStageReviewer && !['corregida', 'pendiente'].includes(next.estado)) {
      return res.status(403).json({ success: false, error: 'El solicitante solo puede alternar la observación entre pendiente y corregida' });
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
    const rawRole = text(req.user.role).toLowerCase() || 'solicitante';
    const isDirector = rawRole === 'director' || rawRole === 'director_escuela' || rawRole.includes('director');
    const isAdmin = rawRole === 'admin';
    if (!isDirector && rawRole !== 'solicitante' && !isAdmin) return res.status(403).json({ success: false, error: 'Rol no autorizado para firmar' });
    const stage = await findStage(req.params.idSolicitud);
    if (!stage || stage.etapa_codigo !== STAGES.FIRMAS) return res.status(409).json({ success: false, error: 'La solicitud no está pendiente de firmas' });
    if (!isDirector && !isAdmin && text(stage.id_usuario) !== text(req.user.id)) return res.status(403).json({ success: false, error: 'Solo el solicitante puede registrar esta firma' });
    const type = isDirector ? 'director_escuela' : 'solicitante';
    const fields = sheetsService.fieldDefinitions.FIRMAS;
    const requestSignatures = (await readRows('FIRMAS')).map((row, i) => rowObject(fields, row, i))
      .filter(item => text(item.id_solicitud) === text(req.params.idSolicitud) && text(item.estado).toLowerCase() === 'firmada');
    const existing = requestSignatures.find(item => {
      const itemType = text(item.tipo_firma).toLowerCase();
      return type === 'director_escuela' ? ['director_escuela', 'director'].includes(itemType) : itemType === 'solicitante';
    });
    if (existing) return res.status(409).json({ success: false, error: 'Esta firma ya fue registrada' });
    if (!isDirector && !isAdmin && !requestSignatures.some(item => ['director_escuela', 'director'].includes(text(item.tipo_firma).toLowerCase()))) {
      return res.status(409).json({ success: false, error: 'Primero debe registrarse la firma del director de escuela' });
    }
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
      .filter(item => text(item.id_solicitud) === text(req.params.idSolicitud) && text(item.estado).toLowerCase() === 'firmada');
    const bothSigned = signatures.some(item => text(item.tipo_firma).toLowerCase() === 'solicitante') && signatures.some(item => ['director_escuela', 'director'].includes(text(item.tipo_firma).toLowerCase()));
    const updated = await updateStage(stage, { estado_etapa: 'en_revision' });
    await logHistory(stage.id_solicitud, STAGES.FIRMAS, 'firmar', stage.estado_etapa, bothSigned ? 'completas' : 'parcial', req);

    let notifications = [];
    try {
      notifications = await workflowNotificationService.notifySignatureRegistered({
        stage: updated,
        signature,
        bothSigned,
        signerUser: req.user
      });
    } catch (notifErr) {
      console.warn('Error al enviar notificaciones de firma registrada:', notifErr.message);
    }

    res.json({
      success: true,
      data: {
        firma: signature,
        firmas: signatures,
        etapa: updated,
        notificacion: notifications[0] || null,
        notificaciones: notifications,
        notificaciones_decanato: []
      }
    });
  } catch (error) {
    if (req.file?.path) fs.unlink(req.file.path, () => {});
    next(error);
  }
};

const deleteSignature = async (req, res, next) => {
  try {
    const rawRole = text(req.user.role).toLowerCase() || 'solicitante';
    const isDirector = rawRole === 'director' || rawRole === 'director_escuela' || rawRole.includes('director');
    const isAdmin = rawRole === 'admin';
    if (!isDirector && rawRole !== 'solicitante' && !isAdmin) {
      return res.status(403).json({ success: false, error: 'Rol no autorizado para eliminar firmas' });
    }
    const stage = await findStage(req.params.idSolicitud);
    if (!stage || stage.etapa_codigo !== STAGES.FIRMAS) {
      return res.status(409).json({ success: false, error: 'Solo se pueden eliminar firmas mientras la solicitud esté en la etapa de firmas' });
    }

    let targetType = text(req.params.tipoFirma || req.query.tipo_firma).toLowerCase();
    if (['director', 'director_escuela'].includes(targetType) || (!targetType && isDirector)) {
      targetType = 'director_escuela';
    } else {
      targetType = 'solicitante';
    }

    if (!isDirector && !isAdmin) {
      if (text(stage.id_usuario) !== text(req.user.id)) {
        return res.status(403).json({ success: false, error: 'Solo el solicitante dueño de la solicitud puede eliminar su firma' });
      }
      targetType = 'solicitante';
    }

    const fields = sheetsService.fieldDefinitions.FIRMAS;
    const requestSignatures = (await readRows('FIRMAS'))
      .map((row, i) => rowObject(fields, row, i))
      .filter(item => {
        if (text(item.id_solicitud) !== text(req.params.idSolicitud)) return false;
        if (text(item.estado).toLowerCase() !== 'firmada') return false;
        const itemType = text(item.tipo_firma).toLowerCase();
        if (targetType === 'director_escuela') {
          return ['director_escuela', 'director'].includes(itemType);
        }
        return itemType === 'solicitante';
      });

    if (requestSignatures.length === 0) {
      return res.status(404).json({ success: false, error: 'No se encontró una firma registrada activa para eliminar' });
    }

    const targetSignature = requestSignatures[requestSignatures.length - 1];
    await update('FIRMAS', fields, targetSignature._row, {
      ...targetSignature,
      estado: 'anulada'
    });

    const label = targetType === 'director_escuela' ? 'Firma del Director de Escuela eliminada' : 'Firma del Solicitante eliminada';
    await logHistory(stage.id_solicitud, STAGES.FIRMAS, 'eliminar_firma', stage.estado_etapa, stage.estado_etapa, req, label);

    res.json({
      success: true,
      message: 'Firma eliminada exitosamente'
    });
  } catch (error) {
    next(error);
  }
};

const STAGE_LABELS_FOR_PDF = {
  solicitud: 'Solicitud',
  revision_1: 'Revisión 1',
  'revision 1': 'Revisión 1',
  'revisión 1': 'Revisión 1',
  revision_2: 'Revisión 2',
  'revision 2': 'Revisión 2',
  'revisión 2': 'Revisión 2',
  firmas: 'Firmas',
  presupuesto: 'Presupuesto',
  lista_para_decanato: 'Lista para Decanato',
  'lista para decanato': 'Lista para Decanato',
  decanato: 'Decanato',
  'en decanato': 'Decanato',
  finalizado: 'Finalizado'
};

const ACTION_LABELS_FOR_PDF = {
  aprobar: 'Aprobado',
  devolver: 'Devuelto con observaciones',
  tomar: 'Tomado en revisión',
  firmar: 'Firmado',
  eliminar_firma: 'Firma eliminada',
  enviar_revision: 'Enviado a Revisión 1',
  reenviar_correcciones: 'Correcciones reenviadas a revisión',
  enviar_presupuesto: 'Enviado a Presupuesto',
  enviar_decanato: 'Enviado a Decanato',
  cargar_plataforma: 'Cargado en plataforma externa y remitido a Decanato',
  registrar_decision: 'Decisión registrada'
};

const ROLE_LABELS_FOR_PDF = {
  solicitante: 'Solicitante',
  revisor_1: 'Revisor 1',
  revisor_2: 'Revisor 2',
  presupuesto: 'Presupuesto',
  director: 'Director',
  director_escuela: 'Director',
  decanato: 'Decanato',
  admin: 'Administrador'
};

function formatHistoryDateStr(value) {
  const raw = String(value || '').trim();
  if (!raw) return 'Fecha no disponible';
  const normalizedRaw = raw.includes(' ') && !raw.includes('T') ? raw.replace(' ', 'T') : raw;
  const parsed = new Date(normalizedRaw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString('es-CO', {
    timeZone: 'America/Bogota',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true
  });
}

function formatHistoryEntryForPdf(entry) {
  const dateStr = formatHistoryDateStr(entry.fecha);
  const rawStage = String(entry.etapa || '').trim().toLowerCase();
  const rawAction = String(entry.accion || '').trim().toLowerCase();
  const rawRole = String(entry.rol_actor || '').trim().toLowerCase();

  const stage = STAGE_LABELS_FOR_PDF[rawStage] || entry.etapa || 'Etapa';
  const action = ACTION_LABELS_FOR_PDF[rawAction] || entry.accion || 'Movimiento';
  const role = ROLE_LABELS_FOR_PDF[rawRole] || entry.rol_actor || '';
  const actor = entry.nombre_actor || entry.id_usuario_actor || 'Sistema';
  const roleStr = role ? ` (${role})` : '';
  const comment = entry.comentario ? ` | Obs: ${entry.comentario}` : '';

  return `- [${dateStr}] Etapa: ${stage} - Acción: ${action} (por: ${actor}${roleStr})${comment}`;
}

async function getHistoryForSolicitud(idSolicitud) {
  try {
    const [historyRows, userRows] = await Promise.all([
      readRows('HISTORIAL_ETAPAS'),
      readRows('USUARIOS')
    ]);
    const userFields = sheetsService.fieldDefinitions.USUARIOS;
    const historyFields = sheetsService.fieldDefinitions.HISTORIAL_ETAPAS;

    const userById = new Map();
    userRows.forEach((row, i) => {
      const u = rowObject(userFields, row, i);
      userById.set(text(u.id_usuario), {
        name: text(u.name || u.nombre || u.nombre_usuario),
        email: text(u.correo_usuario || u.email),
        role: text(u.role || u.rol)
      });
    });

    const matching = historyRows
      .map((row, i) => rowObject(historyFields, row, i))
      .filter(item => text(item.id_solicitud) === text(idSolicitud));

    return enrichHistoryEntries(matching, userById);
  } catch (error) {
    console.error('Error al obtener historial para solicitud:', error);
    return [];
  }
}

async function buildFullNotesAndHistoryText(stage, customNotas = '') {
  let userNotes = text(customNotas);
  if (!userNotes && stage?.comentarios) {
    try {
      const store = typeof stage.comentarios === 'string' ? JSON.parse(stage.comentarios) : stage.comentarios;
      userNotes = text(store?.notas_presupuesto_pdf || store?.notas_decanato_pdf || '');
    } catch (e) {}
  }
  userNotes = userNotes.replace(/\n*--- REGISTRO DEL PASO RECIENTE \(HISTORIAL\) ---[\s\S]*$/, '').trim();

  const isListaOrLater = ['lista_para_decanato', 'decanato', 'finalizado'].includes(String(stage?.etapa_codigo || '').toLowerCase());

  let notesSection = '';
  if (userNotes) {
    notesSection = `OBSERVACIONES / NOTAS DE PRESUPUESTO:\n${userNotes}`;
  }

  let historySection = '';
  if (isListaOrLater || userNotes) {
    const historyEntries = await getHistoryForSolicitud(stage.id_solicitud);
    if (historyEntries.length > 0) {
      const historyLines = historyEntries
        .slice()
        .sort((a, b) => {
          const dateA = a?.fecha ? new Date(a.fecha).getTime() : 0;
          const dateB = b?.fecha ? new Date(b.fecha).getTime() : 0;
          return dateA - dateB;
        })
        .map(entry => formatHistoryEntryForPdf(entry));

      historySection = `REGISTRO INSTITUCIONAL DE MOVIMIENTOS (HISTORIAL):\n${historyLines.join('\n')}`;
    }
  }

  return {
    userNotes: notesSection.trim(),
    historyText: historySection.trim(),
    toString() {
      return [this.userNotes, this.historyText].filter(Boolean).join('\n\n');
    }
  };
}

async function prepareDecanatoAttachments(stage, options = {}) {
  const fields = sheetsService.fieldDefinitions.FIRMAS;
  const signatures = (await readRows('FIRMAS')).map((row, i) => rowObject(fields, row, i))
    .filter(item => text(item.id_solicitud) === text(stage.id_solicitud) && item.estado === 'firmada');
  const fullNotes = await buildFullNotesAndHistoryText(stage, options.notasPdf);
  return createDecanatoAttachments({ stage, signatures, notasPdf: fullNotes });
}

async function dispatchDecanatoNotifications({ stage, updated, body = {} }) {
  let attachments = [];
  let tempDir = null;
  const notasPdf = text(body.notas_pdf || body.notasPdf || body.notas || '');
  try {
    const result = await prepareDecanatoAttachments(updated, { notasPdf });
    attachments = result.attachments;
    tempDir = result.tempDir;
  } catch (attErr) {
    console.warn('No se pudieron adjuntar los PDFs firmados al notificar a Decanato:', attErr.message);
  }
  try {
    const recipients = [
      ...(await findRoleEmails('decanato')),
      ...parseRecipients(body.destinatarios_decanato)
    ];
    const customSubject = text(body.asunto_decanato);
    const customBody = body.cuerpo_decanato ? renderEmailTemplate(body.cuerpo_decanato, stage) : '';
    const notifications = await workflowNotificationService.notifyDecanatoSubmission({
      stage: updated,
      recipients,
      customSubject,
      customBody,
      attachments
    });
    return { notifications, adjuntos: attachments.map(item => item.filename) };
  } finally {
    if (tempDir) {
      await fs.promises.rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}

const sendToPresupuesto = async (req, res, next) => {
  try {
    const role = assertRole(req, ['revisor_1', 'revisor_2', 'admin']);
    const stage = await findStage(req.params.idSolicitud);
    if (!stage || stage.etapa_codigo !== STAGES.FIRMAS) {
      return res.status(409).json({ success: false, error: 'La solicitud no está en la etapa de firmas' });
    }
    const fields = sheetsService.fieldDefinitions.FIRMAS;
    const signatures = (await readRows('FIRMAS')).map((row, i) => rowObject(fields, row, i))
      .filter(item => text(item.id_solicitud) === text(req.params.idSolicitud) && item.estado === 'firmada');
    const bothSigned = signatures.some(item => item.tipo_firma === 'solicitante') && signatures.some(item => item.tipo_firma === 'director_escuela');
    if (!bothSigned) {
      return res.status(409).json({ success: false, error: 'La solicitud aún no cuenta con las dos firmas' });
    }

    const recipients = [
      ...(await findRoleEmails('presupuesto')),
      ...parseRecipients(req.body.destinatarios_presupuesto || req.body.destinatarios_decanato)
    ];
    const customSubject = text(req.body.asunto_presupuesto || req.body.asunto_decanato);
    const customBody = renderEmailTemplate(text(req.body.cuerpo_presupuesto || req.body.cuerpo_decanato), stage);

    const { attachments, tempDir } = await createDecanatoAttachments({ stage, signatures });
    try {
      const updated = await updateStage(stage, {
        etapa_codigo: STAGES.PRESUPUESTO,
        etapa_actual: STAGES.PRESUPUESTO,
        rol_responsable: 'presupuesto',
        estado_etapa: 'pendiente',
        usuario_asignado: '',
        fecha_fin: now()
      });

      await logHistory(stage.id_solicitud, STAGES.FIRMAS, 'enviar_presupuesto', stage.estado_etapa, 'presupuesto', req);

      let notifications = [];
      try {
        notifications = await workflowNotificationService.notifyPresupuestoSubmission({
          stage: updated,
          recipients,
          customSubject,
          customBody,
          attachments
        });
      } catch (e) {
        console.warn('Error al notificar avance a presupuesto:', e.message);
      }

      res.json({
        success: true,
        data: {
          etapa: updated,
          notificaciones_presupuesto: notifications,
          adjuntos: attachments.map(item => item.filename),
          rol: role
        }
      });
    } finally {
      await fs.promises.rm(tempDir, { recursive: true, force: true });
    }
  } catch (error) { next(error); }
};

const sendToDecanato = async (req, res, next) => {
  try {
    const stage = await findStage(req.params.idSolicitud);
    if (stage && stage.etapa_codigo === STAGES.FIRMAS) {
      return sendToPresupuesto(req, res, next);
    }

    const role = assertRole(req, ['presupuesto']);
    if (!stage || stage.etapa_codigo !== STAGES.PRESUPUESTO) {
      return res.status(409).json({ success: false, error: 'La solicitud no está en la etapa de presupuesto' });
    }
    if (stage.usuario_asignado && stage.usuario_asignado !== req.user.id && role !== 'admin') {
      return res.status(403).json({ success: false, error: 'La solicitud fue tomada por otro usuario' });
    }

    const notasPdf = text(req.body.notas_pdf || req.body.notasPdf || req.body.notas);
    const comment = text(req.body.comentario) || (notasPdf ? `Notas en PDFs para Decanato: ${notasPdf}` : 'Aprobado en Presupuesto y remitido a Decanato');

    // Guardar notas_presupuesto_pdf en ETAPAS.comentarios para persistencia permanente
    let comentariosActualizados = stage.comentarios;
    if (notasPdf) {
      const store = parseObservationStore(stage.comentarios);
      store.notas_presupuesto_pdf = notasPdf;
      comentariosActualizados = JSON.stringify(store);
    }

    await append('REVISIONES', sheetsService.fieldDefinitions.REVISIONES, {
      id_revision: await sheetsService.getLastId('REVISIONES') + 1,
      id_solicitud: stage.id_solicitud,
      etapa: stage.etapa_codigo,
      rol_revisor: role,
      id_usuario_revisor: req.user.id,
      estado: 'aprobada',
      observaciones: comment,
      fecha_inicio: stage.fecha_inicio || now(),
      fecha_respuesta: now(),
      comentarios_por_formulario: req.body.comentarios_por_formulario || '',
      numero_intento: req.body.numero_intento || 1,
      id_usuario_toma: stage.usuario_asignado || req.user.id,
      fecha_toma: stage.fecha_inicio || now()
    });

    const updated = await updateStage(stage, {
      etapa_codigo: STAGES.DECANATO,
      etapa_actual: STAGES.DECANATO,
      rol_responsable: 'decanato',
      usuario_asignado: '',
      estado_etapa: 'pendiente',
      fecha_fin: now(),
      comentarios: comentariosActualizados
    });

    await logHistory(stage.id_solicitud, STAGES.PRESUPUESTO, 'enviar_decanato', stage.estado_etapa, 'decanato', req, comment);

    let notifications = [];
    let adjuntos = [];
    try {
      const dispatched = await dispatchDecanatoNotifications({ stage, updated, body: { ...req.body, notas_pdf: notasPdf } });
      notifications = dispatched.notifications;
      adjuntos = dispatched.adjuntos;
    } catch (notifyErr) {
      console.warn('Error al notificar envío a Decanato:', notifyErr.message);
    }

    res.json({
      success: true,
      data: {
        etapa: updated,
        notificaciones_decanato: notifications,
        adjuntos,
        rol: role
      }
    });
  } catch (error) { next(error); }
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
    let notifications = [];
    try {
      notifications = await workflowNotificationService.notifySignatureReminder({
        stage,
        target,
        requestedBy: req.user
      });
    } catch (notifErr) {
      console.warn('Error al enviar recordatorio de firma:', notifErr.message);
    }
    res.json({ success: true, data: notifications[0] || null, notificaciones: notifications });
  } catch (error) { next(error); }
};

const createCorte = async (req, res, next) => {
  try {
    const role = assertRole(req, ['decanato', 'revisor_1', 'revisor_2']);
    const { id_solicitud, estado_corte = 'finalizado', resultado = '', observaciones = '', resolucion = '' } = req.body;
    if (!id_solicitud || !resultado) return res.status(400).json({ success: false, error: 'Se requieren id_solicitud y resultado' });
    if (!text(resolucion)) return res.status(400).json({ success: false, error: 'El nombre o número de resolución es obligatorio' });
    const stage = await findStage(id_solicitud);
    if (!stage || stage.etapa_codigo !== STAGES.DECANATO) return res.status(409).json({ success: false, error: 'La solicitud no está en Decanato' });
    if (stage.usuario_asignado && stage.usuario_asignado !== req.user.id && role !== 'admin') {
      return res.status(403).json({ success: false, error: 'La solicitud fue tomada por otro usuario' });
    }
    if (req.file && req.file.mimetype !== 'application/pdf') return res.status(400).json({ success: false, error: 'La resolución debe ser un archivo PDF' });
    const resolutionName = text(resolucion);
    let resolutionFile = '';
    if (req.file) {
      resolutionFile = await driveService.uploadFile(
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
      resolucion: resolutionName, archivo_resolucion: resolutionFile, estado_final: resultado
    };
    await append('CORTES_DECANATO', sheetsService.fieldDefinitions.CORTES_DECANATO, data);
    const updated = await updateStage(stage, {
      etapa_codigo: STAGES.FINALIZADO, etapa_actual: STAGES.FINALIZADO, rol_responsable: '',
      estado_etapa: 'finalizado', fecha_fin: now(), comentarios: observaciones
    });
    await logHistory(id_solicitud, STAGES.DECANATO, 'registrar_decision', stage.estado_etapa, 'finalizado', req, observaciones);

    let notificationCorte = null;
    try {
      notificationCorte = await workflowNotificationService.notifyDecanatoResolution({
        stage: updated,
        resultado,
        resolucion: resolutionName,
        observaciones
      });
    } catch (corteNotifyErr) {
      console.warn('Error al notificar resolución de Decanato:', corteNotifyErr.message);
    }

    res.status(201).json({ success: true, data: { ...data, etapa: updated, notificacion: notificationCorte } });
  } catch (error) {
    if (req.file?.path) fs.unlink(req.file.path, () => {});
    next(error);
  }
};

const deleteObservation = async (req, res, next) => {
  try {
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    const store = parseObservationStore(stage.comentarios);
    let foundStage = null;
    let itemIndex = -1;
    let currentItem = null;

    for (const [stageName, observations] of Object.entries(store)) {
      const idx = observations.findIndex(item => text(item.id_observacion) === text(req.params.idObservacion));
      if (idx !== -1) {
        foundStage = stageName;
        itemIndex = idx;
        currentItem = observations[idx];
        break;
      }
    }

    if (!currentItem) {
      return res.status(404).json({ success: false, error: 'Observación no encontrada' });
    }

    const role = text(req.user.role).toLowerCase();
    const isAuthor = text(currentItem.id_usuario_autor) === text(req.user.id);
    const isStageReviewer = ROLE_STAGE[role] === foundStage;
    const isRevisor2InStage = role === 'revisor_2' && stage.etapa_codigo === STAGES.REVISION_2;

    if (role !== 'admin' && !isAuthor && !isStageReviewer && !isRevisor2InStage) {
      return res.status(403).json({ success: false, error: 'No tienes permisos para eliminar esta observación' });
    }

    store[foundStage].splice(itemIndex, 1);
    await saveObservationStore(stage, store);

    res.json({ success: true, message: 'Observación eliminada correctamente' });
  } catch (error) { next(error); }
};

const addObservationComment = async (req, res, next) => {
  try {
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    const role = text(req.user.role).toLowerCase() || 'solicitante';
    const allowed = ['revisor_1', 'revisor_2', 'presupuesto', 'decanato', 'solicitante', 'director', 'admin'];
    if (!allowed.includes(role)) return res.status(403).json({ success: false, error: 'Rol no autorizado' });

    const commentText = text(req.body.texto || req.body.comentario);
    if (!commentText) return res.status(400).json({ success: false, error: 'El comentario no puede estar vacío' });

    const store = parseObservationStore(stage.comentarios);
    let targetObs = null;
    for (const [, obsList] of Object.entries(store)) {
      if (Array.isArray(obsList)) {
        const found = obsList.find(o => text(o.id_observacion) === text(req.params.idObservacion));
        if (found) {
          targetObs = found;
          break;
        }
      }
    }
    if (!targetObs) return res.status(404).json({ success: false, error: 'Observación no encontrada' });

    if (!Array.isArray(targetObs.comentarios)) {
      targetObs.comentarios = [];
    }

    const newComment = {
      id_comentario: `com-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      id_usuario_autor: req.user.id,
      nombre_autor: req.user.name || req.user.nombre || 'Usuario',
      rol_autor: role,
      texto: commentText,
      fecha_creacion: now()
    };
    targetObs.comentarios.push(newComment);
    targetObs.fecha_actualizacion = now();

    await saveObservationStore(stage, store);

    res.status(201).json({ success: true, data: newComment, observacion: targetObs });
  } catch (error) { next(error); }
};

const deleteObservationComment = async (req, res, next) => {
  try {
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    const store = parseObservationStore(stage.comentarios);
    let found = false;
    for (const [, obsList] of Object.entries(store)) {
      if (Array.isArray(obsList)) {
        const targetObs = obsList.find(o => text(o.id_observacion) === text(req.params.idObservacion));
        if (targetObs && Array.isArray(targetObs.comentarios)) {
          const idx = targetObs.comentarios.findIndex(c => text(c.id_comentario) === text(req.params.idComentario));
          if (idx !== -1) {
            const comment = targetObs.comentarios[idx];
            if (req.user.role !== 'admin' && text(comment.id_usuario_autor) !== text(req.user.id)) {
              return res.status(403).json({ success: false, error: 'No autorizado para eliminar este comentario' });
            }
            targetObs.comentarios.splice(idx, 1);
            targetObs.fecha_actualizacion = now();
            found = true;
            break;
          }
        }
      }
    }
    if (!found) return res.status(404).json({ success: false, error: 'Comentario no encontrado' });
    await saveObservationStore(stage, store);
    res.json({ success: true, message: 'Comentario eliminado' });
  } catch (error) { next(error); }
};

const getFormPdf = async (req, res, next) => {
  try {
    const { idSolicitud, formNumber } = req.params;
    const stage = await findStage(idSolicitud);
    if (!stage) {
      return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    }
    const formNum = parseInt(formNumber, 10);
    if (![1, 2, 3, 4].includes(formNum)) {
      return res.status(400).json({ success: false, error: 'Número de formulario inválido' });
    }

    const fields = sheetsService.fieldDefinitions.FIRMAS;
    const signatures = (await readRows('FIRMAS')).map((row, i) => rowObject(fields, row, i))
      .filter(item => text(item.id_solicitud) === text(idSolicitud) && item.estado === 'firmada');

    const customNotas = text(req.query.notas || req.query.notas_pdf);
    const fullNotes = await buildFullNotesAndHistoryText(stage, customNotas);
    const pdfBuffer = await generateFormPdfBuffer({
      stage,
      formNumber: formNum,
      signatures,
      notasPdf: fullNotes || undefined
    });

    const isDownload = req.query.download === '1' || req.query.download === 'true';
    const disposition = isDownload ? 'attachment' : 'inline';
    const filename = `Formulario${formNum}_${idSolicitud}_firmado.pdf`;

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `${disposition}; filename="${filename}"`);
    return res.send(pdfBuffer);
  } catch (error) {
    console.error('Error al generar PDF del formulario:', error);
    next(error);
  }
};

const getSignatureImage = async (req, res, next) => {
  try {
    const { idSolicitud, tipoFirma } = req.params;
    let normalizedType = text(tipoFirma).toLowerCase();
    if (normalizedType === 'director') normalizedType = 'director_escuela';
    if (!['director_escuela', 'solicitante'].includes(normalizedType)) {
      return res.status(400).json({ success: false, error: 'Tipo de firma inválido' });
    }

    const fields = sheetsService.fieldDefinitions.FIRMAS;
    const signatures = (await readRows('FIRMAS')).map((row, i) => rowObject(fields, row, i))
      .filter((item) => {
        if (text(item.id_solicitud) !== text(idSolicitud)) return false;
        if (text(item.estado).toLowerCase() !== 'firmada') return false;
        const itemType = text(item.tipo_firma).toLowerCase();
        return normalizedType === 'director_escuela'
          ? ['director_escuela', 'director'].includes(itemType)
          : itemType === 'solicitante';
      });
    const signature = signatures[signatures.length - 1];
    if (!signature?.evidencia_firma) {
      return res.status(404).json({ success: false, error: 'Firma no encontrada' });
    }

    const buffer = await getSignatureBuffer(signature);
    if (!buffer || buffer.length === 0) {
      return res.status(404).json({ success: false, error: 'No se pudo cargar la imagen de la firma' });
    }

    let contentType = 'image/png';
    if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xd8) {
      contentType = 'image/jpeg';
    } else if (buffer.length >= 12 && buffer.subarray(0, 4).toString('ascii') === 'RIFF' && buffer.subarray(8, 12).toString('ascii') === 'WEBP') {
      contentType = 'image/webp';
    }

    res.setHeader('Content-Type', contentType);
    res.setHeader('Cache-Control', 'private, max-age=300');
    return res.send(buffer);
  } catch (error) {
    next(error);
  }
};

const saveDecanatoNotes = async (req, res, next) => {
  try {
    const role = assertRole(req, ['presupuesto', 'revisor_2', 'admin']);
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    if (![STAGES.PRESUPUESTO, STAGES.LISTA_PARA_DECANATO].includes(stage.etapa_codigo) && role !== 'admin') {
      return res.status(409).json({ success: false, error: 'La solicitud no está en la etapa requerida' });
    }
    const notasPdf = text(req.body.notas_pdf ?? req.body.notasPdf ?? req.body.notas);
    let store = {};
    try {
      store = typeof stage.comentarios === 'string' ? JSON.parse(stage.comentarios || '{}') : (stage.comentarios || {});
    } catch {
      store = {};
    }
    store.notas_presupuesto_pdf = notasPdf;
    if (req.body.incluir_historial !== undefined) {
      store.incluir_historial_pdf = Boolean(req.body.incluir_historial);
    }
    const updated = await updateStage(stage, { comentarios: JSON.stringify(store) });
    res.json({
      success: true,
      message: 'Notas para los PDF guardadas correctamente',
      data: {
        id_solicitud: stage.id_solicitud,
        notas_pdf: notasPdf,
        incluir_historial: store.incluir_historial_pdf
      }
    });
  } catch (error) { next(error); }
};

const confirmPlataformaDecanato = async (req, res, next) => {
  try {
    const role = assertRole(req, ['presupuesto', 'revisor_2', 'admin']);
    const stage = await findStage(req.params.idSolicitud);
    if (!stage) return res.status(404).json({ success: false, error: 'Solicitud no encontrada' });
    if (![STAGES.LISTA_PARA_DECANATO, STAGES.PRESUPUESTO].includes(stage.etapa_codigo) && role !== 'admin') {
      return res.status(409).json({ success: false, error: 'La solicitud no está en etapa Lista para Decanato' });
    }

    const notasPdf = text(req.body.notas_pdf ?? req.body.notasPdf ?? req.body.notas);
    let comentariosActualizados = stage.comentarios;
    if (notasPdf) {
      let store = {};
      try {
        store = typeof stage.comentarios === 'string' ? JSON.parse(stage.comentarios || '{}') : (stage.comentarios || {});
      } catch {
        store = {};
      }
      store.notas_presupuesto_pdf = notasPdf;
      comentariosActualizados = JSON.stringify(store);
    }

    const comment = text(req.body.comentario) || (notasPdf ? `Cargado en plataforma externa. Notas: ${notasPdf}` : 'Documentos en PDF cargados en plataforma externa y solicitud remitida a Decanato');

    await append('REVISIONES', sheetsService.fieldDefinitions.REVISIONES, {
      id_revision: await sheetsService.getLastId('REVISIONES') + 1,
      id_solicitud: stage.id_solicitud,
      etapa: stage.etapa_codigo,
      rol_revisor: role,
      id_usuario_revisor: req.user.id,
      estado: 'aprobada',
      observaciones: comment,
      fecha_inicio: stage.fecha_inicio || now(),
      fecha_respuesta: now(),
      comentarios_por_formulario: '',
      numero_intento: 1,
      id_usuario_toma: stage.usuario_asignado || req.user.id,
      fecha_toma: stage.fecha_inicio || now()
    });

    const updated = await updateStage(stage, {
      etapa_codigo: STAGES.DECANATO,
      etapa_actual: STAGES.DECANATO,
      rol_responsable: 'decanato',
      usuario_asignado: '',
      estado_etapa: 'pendiente',
      fecha_fin: now(),
      comentarios: comentariosActualizados
    });

    await logHistory(stage.id_solicitud, stage.etapa_codigo, 'cargar_plataforma', stage.estado_etapa, 'pendiente', req, comment);

    res.json({
      success: true,
      message: 'Solicitud remitida a Decanato correctamente tras confirmar la subida a plataforma',
      data: {
        etapa: updated,
        rol: role
      }
    });
  } catch (error) { next(error); }
};

module.exports = {
  getQueue,
  getConsultationQueue,
  sendToDecanato,
  sendToPresupuesto,
  getPendingSignatures,
  take,
  decide,
  getObservations,
  addObservation,
  updateObservation,
  deleteObservation,
  addObservationComment,
  deleteObservationComment,
  sign,
  deleteSignature,
  getSignatureImage,
  remindSignature,
  createCorte,
  getFormPdf,
  saveDecanatoNotes,
  confirmPlataformaDecanato
};
