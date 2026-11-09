const sheetsService = require('./sheetsService');
const emailService = require('./emailService');

const STAGE_LABELS = Object.freeze({
  solicitud: 'Solicitud / Diligenciamiento',
  revision_1: 'Revisión 1',
  revision_2: 'Revisión 2',
  presupuesto: 'Presupuesto',
  firmas: 'Firmas',
  decanato: 'Decanato',
  finalizado: 'Finalizado'
});

const ROLE_LABELS = Object.freeze({
  solicitante: 'Solicitante',
  revisor_1: 'Revisor 1',
  revisor_2: 'Revisor 2',
  presupuesto: 'Presupuesto',
  director: 'Director de Escuela',
  decanato: 'Decanato',
  admin: 'Administrador'
});

const now = () => new Date().toISOString();
const text = value => String(value ?? '').trim();

function rowObject(fields, row, index) {
  return fields.reduce((obj, field, i) => {
    obj[field] = row[i] ?? '';
    return obj;
  }, { _row: index + 2 });
}

async function readRows(sheet) {
  const response = await sheetsService.client.spreadsheets.values.get({
    spreadsheetId: sheetsService.spreadsheetId,
    range: `${sheet}!A2:Z`
  });
  return response.data.values || [];
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

class WorkflowNotificationService {
  constructor() {
    this.stageLabels = STAGE_LABELS;
    this.roleLabels = ROLE_LABELS;
  }

  async getUser(userId) {
    if (!userId) return null;
    const fields = sheetsService.fieldDefinitions.USUARIOS;
    const rows = (await readRows('USUARIOS')).map((row, i) => rowObject(fields, row, i));
    return rows.find(user => text(user.id_usuario) === text(userId)) || null;
  }

  async getUserEmail(userId) {
    if (!userId) return '';
    const user = await this.getUser(userId);
    return text(user?.correo_usuario || user?.email);
  }

  async getRoleEmails(role) {
    if (!role) return [];
    const normalizedRole = text(role).toLowerCase();
    const fields = sheetsService.fieldDefinitions.USUARIOS;
    const rows = (await readRows('USUARIOS')).map((row, i) => rowObject(fields, row, i));
    const emails = rows
      .filter(user => text(user.role).toLowerCase() === normalizedRole)
      .map(user => text(user.correo_usuario || user.email))
      .filter(Boolean);
    return [...new Set(emails)];
  }

  async getStage(idSolicitud) {
    if (!idSolicitud) return null;
    const rows = await readRows('ETAPAS');
    const fields = sheetsService.fieldDefinitions.ETAPAS;
    const index = rows.findIndex(row => text(row[0]) === text(idSolicitud));
    return index < 0 ? null : rowObject(fields, rows[index], index);
  }

  formatObservations(observations, comment = '') {
    const lines = [];
    if (comment && text(comment)) {
      lines.push(`Comentario general: ${text(comment)}`);
    }
    if (Array.isArray(observations) && observations.length > 0) {
      const items = observations
        .map((obs, idx) => {
          if (typeof obs === 'string' && obs.trim()) {
            return `• ${obs.trim()}`;
          }
          if (obs && typeof obs === 'object') {
            const context = [
              obs.formulario ? `Formulario ${obs.formulario}` : '',
              obs.paso ? `Paso ${obs.paso}` : ''
            ].filter(Boolean).join(' - ');
            const obsText = obs.observacion || obs.texto || '';
            if (!obsText) return '';
            return `• ${context ? `[${context}] ` : ''}${obsText}`;
          }
          return '';
        })
        .filter(Boolean);

      if (items.length > 0) {
        lines.push('Observaciones registradas:\n' + items.join('\n'));
      }
    }
    return lines.join('\n\n');
  }

  buildHtmlTemplate({
    title,
    badgeText,
    badgeColor = '#2563eb',
    recipientName = '',
    message,
    details = [],
    callToActionText = 'Ingrese a la plataforma para gestionar o revisar la solicitud.'
  }) {
    const detailRows = details
      .filter(item => item && item.label && item.value)
      .map(item => `
        <tr>
          <td style="padding: 8px 12px; font-weight: 600; color: #374151; width: 35%; border-bottom: 1px solid #e5e7eb; vertical-align: top;">
            ${item.label}:
          </td>
          <td style="padding: 8px 12px; color: #111827; border-bottom: 1px solid #e5e7eb; vertical-align: top;">
            ${String(item.value).replace(/\n/g, '<br/>')}
          </td>
        </tr>
      `)
      .join('');

    return `
      <!DOCTYPE html>
      <html lang="es">
      <head>
        <meta charset="utf-8">
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f3f4f6; margin: 0; padding: 20px; color: #1f2937; }
          .container { max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
          .header { background-color: #BA0C2F; padding: 24px; text-align: center; color: #ffffff; }
          .header h1 { margin: 0; font-size: 20px; font-weight: 700; letter-spacing: 0.5px; }
          .header p { margin: 6px 0 0 0; font-size: 13px; opacity: 0.9; }
          .content { padding: 24px; }
          .badge { display: inline-block; padding: 4px 12px; font-size: 12px; font-weight: 700; text-transform: uppercase; border-radius: 9999px; color: #ffffff; background-color: ${badgeColor}; margin-bottom: 16px; }
          .message { font-size: 15px; line-height: 1.6; color: #374151; margin-bottom: 20px; }
          .details-card { background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 6px; overflow: hidden; margin-bottom: 24px; }
          .details-table { width: 100%; border-collapse: collapse; font-size: 14px; }
          .cta { background-color: #eff6ff; border-left: 4px solid #2563eb; padding: 14px 16px; border-radius: 4px; font-size: 14px; color: #1e40af; margin-bottom: 20px; }
          .footer { background-color: #f9fafb; border-top: 1px solid #e5e7eb; padding: 16px 24px; text-align: center; font-size: 12px; color: #6b7280; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>Sistema de Gestión de Extensión</h1>
            <p>Facultad de Salud · Universidad del Valle</p>
          </div>
          <div class="content">
            <span class="badge">${badgeText}</span>
            <h2 style="margin: 0 0 16px 0; font-size: 18px; color: #111827;">${title}</h2>
            ${recipientName ? `<p style="font-size: 14px; color: #6b7280; margin: 0 0 12px 0;">Estimado(a) <strong>${recipientName}</strong>,</p>` : ''}
            <div class="message">${message.replace(/\n/g, '<br/>')}</div>
            ${details.length > 0 ? `
              <div class="details-card">
                <table class="details-table">
                  <tbody>
                    ${detailRows}
                  </tbody>
                </table>
              </div>
            ` : ''}
            ${callToActionText ? `<div class="cta">${callToActionText}</div>` : ''}
          </div>
          <div class="footer">
            Este es un mensaje institucional automático del Sistema de Gestión de Solicitudes de Extensión.<br/>
            Por favor no responda a este correo.
          </div>
        </div>
      </body>
      </html>
    `;
  }

  async recordNotification({
    idSolicitud,
    event,
    recipient,
    subject,
    text: bodyText,
    html,
    attachments = []
  }) {
    const fields = sheetsService.fieldDefinitions.NOTIFICACIONES;
    const cleanRecipient = text(recipient).toLowerCase();

    if (!cleanRecipient) {
      console.warn(`[NOTIFICACIONES] No se envió ${event}: falta correo destinatario para solicitud ${idSolicitud}`);
      const emptyData = {
        id_notificacion: await sheetsService.getLastId('NOTIFICACIONES') + 1,
        id_solicitud: idSolicitud,
        evento: event,
        estado: 'sin_destinatario',
        intentos: 0,
        fecha_creacion: now(),
        fecha_envio: '',
        ultimo_error: 'No se encontró correo destinatario'
      };
      await append('NOTIFICACIONES', fields, emptyData);
      return emptyData;
    }

    const data = {
      id_notificacion: await sheetsService.getLastId('NOTIFICACIONES') + 1,
      id_solicitud: idSolicitud,
      evento: event,
      estado: 'pendiente',
      intentos: 1,
      fecha_creacion: now(),
      fecha_envio: '',
      ultimo_error: ''
    };

    try {
      await emailService.send({
        to: cleanRecipient,
        subject,
        text: bodyText,
        html,
        attachments
      });
      data.estado = 'enviada';
      data.fecha_envio = now();
      console.log(`[NOTIFICACIONES] Correo '${event}' enviado exitosamente a ${cleanRecipient} (Solicitud #${idSolicitud})`);
    } catch (error) {
      data.estado = 'fallida';
      data.ultimo_error = error.message;
      console.error(`[NOTIFICACIONES] Falló envío '${event}' a ${cleanRecipient} (Solicitud #${idSolicitud}):`, error.message);
    }

    try {
      await append('NOTIFICACIONES', fields, data);
    } catch (sheetError) {
      console.error(`[NOTIFICACIONES] Error al registrar notificación en hoja:`, sheetError.message);
    }

    return data;
  }

  /**
   * Notifica cuando el solicitante envía o reenvía la solicitud a revisión.
   * Actores:
   * 1. Solicitante (confirma que su solicitud avanzó a revisión)
   * 2. Revisores de la etapa destino (ej. revisor_1 o revisor_2)
   */
  async notifyRequestSubmitted({
    idSolicitud,
    stage = {},
    applicantId,
    applicantName = '',
    activityName = '',
    targetStage = 'revision_1',
    targetRole = 'revisor_1',
    isResubmission = false
  }) {
    const fullStage = stage.nombre_solicitante ? stage : (await this.getStage(idSolicitud)) || stage;
    const finalApplicantId = applicantId || fullStage.id_usuario;
    const finalApplicantName = applicantName || fullStage.nombre_solicitante || 'Solicitante';
    const finalActivityName = activityName || fullStage.nombre_actividad || 'Actividad de extensión';
    const stageLabel = this.stageLabels[targetStage] || targetStage;
    const actionLabel = isResubmission ? 'reenviada tras subsanar observaciones' : 'enviada a revisión';

    const notifications = [];

    // 1. Notificar al Solicitante
    const applicantEmail = await this.getUserEmail(finalApplicantId);
    const applicantSubject = `[Extensión] Solicitud #${idSolicitud} ${isResubmission ? 'reenviada' : 'enviada'} a ${stageLabel}`;
    const applicantMessage = `Su solicitud #${idSolicitud} ha sido ${actionLabel} y se encuentra en la etapa de ${stageLabel}. En breve el equipo revisor continuará con su evaluación.`;
    const applicantHtml = this.buildHtmlTemplate({
      title: `Solicitud ${isResubmission ? 'Reenviada' : 'Enviada'} a ${stageLabel}`,
      badgeText: isResubmission ? 'Reenvío de Solicitud' : 'Envío a Revisión',
      badgeColor: '#2563eb',
      recipientName: finalApplicantName,
      message: applicantMessage,
      details: [
        { label: 'Número de Solicitud', value: `#${idSolicitud}` },
        { label: 'Actividad', value: finalActivityName },
        { label: 'Etapa Actual', value: stageLabel },
        { label: 'Responsable', value: this.roleLabels[targetRole] || targetRole },
        { label: 'Fecha de Envío', value: new Date().toLocaleString('es-CO') }
      ],
      callToActionText: 'Puede consultar el progreso y los comentarios de su solicitud desde el portal web en cualquier momento.'
    });

    const notifApplicant = await this.recordNotification({
      idSolicitud,
      event: isResubmission ? 'solicitud_reenviada_revision' : 'solicitud_enviada_revision',
      recipient: applicantEmail,
      subject: applicantSubject,
      text: `${applicantMessage}\n\nSolicitud: #${idSolicitud}\nActividad: ${finalActivityName}\nEtapa: ${stageLabel}`,
      html: applicantHtml
    });
    notifications.push(notifApplicant);

    // 2. Notificar a los Revisores del rol destino (ej. revisor_1 o revisor_2)
    const reviewerEmails = await this.getRoleEmails(targetRole);
    const reviewerSubject = `[Extensión] Nueva solicitud para ${stageLabel} - #${idSolicitud}`;
    const reviewerMessage = `La solicitud #${idSolicitud} ("${finalActivityName}"), presentada por ${finalApplicantName}, se encuentra en su bandeja de ${stageLabel} pendiente de revisión.`;
    const reviewerHtml = this.buildHtmlTemplate({
      title: `Nueva Solicitud en Bandeja de ${stageLabel}`,
      badgeText: 'Pendiente de Revisión',
      badgeColor: '#059669',
      recipientName: this.roleLabels[targetRole] || 'Equipo Evaluador',
      message: reviewerMessage,
      details: [
        { label: 'Número de Solicitud', value: `#${idSolicitud}` },
        { label: 'Actividad', value: finalActivityName },
        { label: 'Solicitante', value: finalApplicantName },
        { label: 'Etapa', value: stageLabel },
        { label: 'Tipo de Evento', value: isResubmission ? 'Reenvío con observaciones subsanadas' : 'Primera postulación' },
        { label: 'Fecha', value: new Date().toLocaleString('es-CO') }
      ],
      callToActionText: 'Por favor ingrese a su cola de revisión en el sistema para tomar la solicitud y emitir concepto.'
    });

    if (reviewerEmails.length === 0) {
      notifications.push(await this.recordNotification({
        idSolicitud,
        event: isResubmission ? 'solicitud_reenviada_notif_revisor' : 'solicitud_pendiente_revision',
        recipient: '',
        subject: reviewerSubject,
        text: reviewerMessage,
        html: reviewerHtml
      }));
    } else {
      for (const email of reviewerEmails) {
        notifications.push(await this.recordNotification({
          idSolicitud,
          event: isResubmission ? 'solicitud_reenviada_notif_revisor' : 'solicitud_pendiente_revision',
          recipient: email,
          subject: reviewerSubject,
          text: `${reviewerMessage}\n\nSolicitud: #${idSolicitud}\nActividad: ${finalActivityName}\nSolicitante: ${finalApplicantName}`,
          html: reviewerHtml
        }));
      }
    }

    return notifications;
  }

  /**
   * Notifica cuando un revisor aprueba una etapa y la solicitud avanza.
   * Transiciones:
   * - revision_1 -> revision_2: Solicitante (avanzó a R2) y revisores R2 (nueva solicitud)
   * - revision_2 -> presupuesto: Solicitante (avanzó a Presupuesto) y Presupuesto (nueva solicitud)
   * - presupuesto -> firmas: Solicitante (firmar) y Director (firmar)
   */
  async notifyStageApproved({
    stage = {},
    previousStage = '',
    nextStage = '',
    nextRole = '',
    comment = '',
    approvedBy = {},
    attachments = []
  }) {
    const idSolicitud = stage.id_solicitud;
    const fullStage = stage.nombre_solicitante ? stage : (await this.getStage(idSolicitud)) || stage;
    const applicantId = fullStage.id_usuario;
    const applicantName = fullStage.nombre_solicitante || 'Solicitante';
    const activityName = fullStage.nombre_actividad || 'Actividad de extensión';
    const prevLabel = this.stageLabels[previousStage] || previousStage;
    const nextLabel = this.stageLabels[nextStage] || nextStage;

    const notifications = [];
    const applicantEmail = await this.getUserEmail(applicantId);

    // Caso 1: Avance a FIRMAS (desde presupuesto)
    if (nextStage === 'firmas') {
      // 1.1 Solicitante (debe registrar su firma)
      const solSubject = `[Extensión] Solicitud #${idSolicitud} aprobada en Revisión 2 - Pendiente de Firmas`;
      const solMsg = `¡Buenas noticias! Su solicitud #${idSolicitud} ("${activityName}") ha sido aprobada en la etapa de Revisión 2 y ha avanzado a la etapa de Firmas. Por favor ingrese al sistema para registrar su firma institucional.`;
      const solHtml = this.buildHtmlTemplate({
        title: 'Solicitud Lista para Firmas',
        badgeText: 'Firma Requerida',
        badgeColor: '#16a34a',
        recipientName: applicantName,
        message: solMsg,
        details: [
          { label: 'Número de Solicitud', value: `#${idSolicitud}` },
          { label: 'Actividad', value: activityName },
          { label: 'Etapa Anterior', value: prevLabel },
          { label: 'Etapa Actual', value: nextLabel },
          comment ? { label: 'Comentarios de Aprobación', value: comment } : null
        ].filter(Boolean),
        callToActionText: 'Ingrese a la plataforma para cargar o registrar su firma y continuar el trámite hacia Presupuesto.'
      });
      notifications.push(await this.recordNotification({
        idSolicitud,
        event: 'solicitud_avanza_firmas_solicitante',
        recipient: applicantEmail,
        subject: solSubject,
        text: `${solMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}\n${comment ? `Comentario: ${comment}` : ''}`,
        html: solHtml
      }));

      // 1.2 Director de Escuela (requiere su firma de aval)
      const directorEmails = await this.getRoleEmails('director');
      const dirSubject = `[Extensión] Solicitud #${idSolicitud} pendiente de firma de Director`;
      const dirMsg = `La solicitud de extensión #${idSolicitud} ("${activityName}"), presentada por ${applicantName}, ha completado las revisiones de Extensión y se encuentra en etapa de Firmas. Se requiere su firma de aval como Director de Escuela.`;
      const dirHtml = this.buildHtmlTemplate({
        title: 'Firma de Director de Escuela Requerida',
        badgeText: 'Firma Pendiente',
        badgeColor: '#2563eb',
        recipientName: 'Director(a) de Escuela',
        message: dirMsg,
        details: [
          { label: 'Número de Solicitud', value: `#${idSolicitud}` },
          { label: 'Actividad', value: activityName },
          { label: 'Solicitante', value: applicantName },
          { label: 'Etapa', value: 'Firmas' }
        ],
        callToActionText: 'Por favor ingrese a su bandeja de firmas en el sistema para avalar la solicitud.'
      });

      if (directorEmails.length === 0) {
        notifications.push(await this.recordNotification({
          idSolicitud,
          event: 'solicitud_avanza_firmas_director',
          recipient: '',
          subject: dirSubject,
          text: dirMsg,
          html: dirHtml
        }));
      } else {
        for (const email of directorEmails) {
          notifications.push(await this.recordNotification({
            idSolicitud,
            event: 'solicitud_avanza_firmas_director',
            recipient: email,
            subject: dirSubject,
            text: `${dirMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}\nSolicitante: ${applicantName}`,
            html: dirHtml
          }));
        }
      }

      return notifications;
    }

    // Caso 2: Avances ordinarios (revision_1 -> revision_2 o revision_2 -> presupuesto)
    // 2.1 Notificar al Solicitante
    const solSubject = `[Extensión] Solicitud #${idSolicitud} aprobada en ${prevLabel} (Avanza a ${nextLabel})`;
    const solMsg = `Le informamos que su solicitud #${idSolicitud} ("${activityName}") ha sido aprobada en la etapa de ${prevLabel} y ha avanzado a la etapa de ${nextLabel}.`;
    const solHtml = this.buildHtmlTemplate({
      title: `Avance a ${nextLabel}`,
      badgeText: 'Etapa Aprobada',
      badgeColor: '#16a34a',
      recipientName: applicantName,
      message: solMsg,
      details: [
        { label: 'Número de Solicitud', value: `#${idSolicitud}` },
        { label: 'Actividad', value: activityName },
        { label: 'Etapa Aprobada', value: prevLabel },
        { label: 'Nueva Etapa', value: nextLabel },
        { label: 'Nuevo Responsable', value: this.roleLabels[nextRole] || nextRole },
        comment ? { label: 'Comentarios', value: comment } : null
      ].filter(Boolean),
      callToActionText: 'Puede consultar el seguimiento de su solicitud en el sistema.'
    });
    notifications.push(await this.recordNotification({
      idSolicitud,
      event: 'etapa_aprobada_solicitante',
      recipient: applicantEmail,
      subject: solSubject,
      text: `${solMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}\n${comment ? `Comentario: ${comment}` : ''}`,
      html: solHtml
    }));

    // 2.2 Notificar al nuevo rol responsable (ej. revisor_2 o presupuesto)
    const nextRoleEmails = await this.getRoleEmails(nextRole);
    const eventName = nextRole === 'presupuesto' ? 'solicitud_asignada_presupuesto' : 'solicitud_asignada_revisor';
    const roleSubject = `[Extensión] Solicitud #${idSolicitud} disponible para ${nextLabel}`;
    const roleMsg = `La solicitud #${idSolicitud} ("${activityName}"), presentada por ${applicantName}, fue aprobada en ${prevLabel} y se encuentra disponible en su bandeja de ${nextLabel} para su evaluación.`;
    const roleHtml = this.buildHtmlTemplate({
      title: `Nueva Solicitud en ${nextLabel}`,
      badgeText: 'Asignación de Solicitud',
      badgeColor: '#2563eb',
      recipientName: this.roleLabels[nextRole] || 'Equipo Evaluador',
      message: roleMsg,
      details: [
        { label: 'Número de Solicitud', value: `#${idSolicitud}` },
        { label: 'Actividad', value: activityName },
        { label: 'Solicitante', value: applicantName },
        { label: 'Etapa', value: nextLabel },
        comment ? { label: 'Comentarios de Etapa Previa', value: comment } : null
      ].filter(Boolean),
      callToActionText: 'Por favor ingrese al sistema para tomar la solicitud y realizar la revisión respectiva.'
    });

    if (nextRoleEmails.length === 0) {
      notifications.push(await this.recordNotification({
        idSolicitud,
        event: eventName,
        recipient: '',
        subject: roleSubject,
        text: roleMsg,
        html: roleHtml,
        attachments
      }));
    } else {
      for (const email of nextRoleEmails) {
        notifications.push(await this.recordNotification({
          idSolicitud,
          event: eventName,
          recipient: email,
          subject: roleSubject,
          text: `${roleMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}\nSolicitante: ${applicantName}`,
          html: roleHtml,
          attachments
        }));
      }
    }

    return notifications;
  }

  /**
   * Notifica cuando una solicitud es devuelta con observaciones.
   * Transiciones:
   * - Devolución a solicitante (desde revision_1 o revision_2 hacia 'solicitud')
   * - Devolución desde presupuesto hacia revision_2
   */
  async notifyStageReturned({
    stage = {},
    fromStage = '',
    toStage = '',
    returnRole = '',
    comment = '',
    observations = [],
    returnedBy = {}
  }) {
    const idSolicitud = stage.id_solicitud;
    const fullStage = stage.nombre_solicitante ? stage : (await this.getStage(idSolicitud)) || stage;
    const applicantId = fullStage.id_usuario;
    const applicantName = fullStage.nombre_solicitante || 'Solicitante';
    const activityName = fullStage.nombre_actividad || 'Actividad de extensión';
    const fromLabel = this.stageLabels[fromStage] || fromStage;
    const toLabel = this.stageLabels[toStage] || toStage;

    const formattedObs = this.formatObservations(observations, comment);
    const notifications = [];
    const applicantEmail = await this.getUserEmail(applicantId);

    // Caso 1: Devuelta al solicitante (desde revision_1 o revision_2)
    if (toStage === 'solicitud') {
      const solSubject = `[Extensión] Solicitud #${idSolicitud} devuelta con observaciones (${fromLabel})`;
      const solMsg = `Le informamos que su solicitud #${idSolicitud} ("${activityName}") ha sido devuelta en la etapa de ${fromLabel} debido a que requiere ajustes y subsanación de observaciones.`;
      const solHtml = this.buildHtmlTemplate({
        title: `Solicitud Devuelta con Observaciones`,
        badgeText: 'Requiere Correcciones',
        badgeColor: '#d97706',
        recipientName: applicantName,
        message: solMsg,
        details: [
          { label: 'Número de Solicitud', value: `#${idSolicitud}` },
          { label: 'Actividad', value: activityName },
          { label: 'Etapa que Devuelve', value: fromLabel },
          formattedObs ? { label: 'Observaciones / Motivo', value: formattedObs } : null
        ].filter(Boolean),
        callToActionText: 'Por favor ingrese al sistema, atienda las observaciones señaladas, márquelas como subsanadas y reenvíe la solicitud a revisión.'
      });

      notifications.push(await this.recordNotification({
        idSolicitud,
        event: 'solicitud_devuelta_solicitante',
        recipient: applicantEmail,
        subject: solSubject,
        text: `${solMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}\n\n${formattedObs}\n\nPor favor ingrese al sistema para subsanar las observaciones.`,
        html: solHtml
      }));

      return notifications;
    }

    // Caso 2: Devuelta desde Presupuesto hacia Revisión 2
    if (fromStage === 'presupuesto' && toStage === 'revision_2') {
      // 2.1 Revisor 2 (quien debe ajustar o revisar las observaciones presupuestales)
      const r2Emails = await this.getRoleEmails('revisor_2');
      const r2Subject = `[Extensión] Solicitud #${idSolicitud} devuelta desde Presupuesto a Revisión 2`;
      const r2Msg = `La solicitud #${idSolicitud} ("${activityName}") de ${applicantName} ha sido devuelta desde la etapa de Presupuesto a su bandeja de Revisión 2 para revisión de aspectos presupuestales.`;
      const r2Html = this.buildHtmlTemplate({
        title: 'Devolución desde Presupuesto',
        badgeText: 'Devuelta a Revisión 2',
        badgeColor: '#d97706',
        recipientName: 'Equipo de Revisión 2',
        message: r2Msg,
        details: [
          { label: 'Número de Solicitud', value: `#${idSolicitud}` },
          { label: 'Actividad', value: activityName },
          { label: 'Solicitante', value: applicantName },
          { label: 'Devuelto Por', value: 'Presupuesto' },
          formattedObs ? { label: 'Observaciones de Presupuesto', value: formattedObs } : null
        ].filter(Boolean),
        callToActionText: 'Por favor ingrese al sistema para revisar las observaciones y gestionar las correcciones necesarias.'
      });

      if (r2Emails.length === 0) {
        notifications.push(await this.recordNotification({
          idSolicitud,
          event: 'solicitud_devuelta_revisor_2',
          recipient: '',
          subject: r2Subject,
          text: r2Msg,
          html: r2Html
        }));
      } else {
        for (const email of r2Emails) {
          notifications.push(await this.recordNotification({
            idSolicitud,
            event: 'solicitud_devuelta_revisor_2',
            recipient: email,
            subject: r2Subject,
            text: `${r2Msg}\n\nSolicitud: #${idSolicitud}\n${formattedObs}`,
            html: r2Html
          }));
        }
      }

      // 2.2 Notificar al solicitante que la solicitud está en revisión de ajustes en Revisión 2
      const solSubject = `[Extensión] Solicitud #${idSolicitud} devuelta a Revisión 2`;
      const solMsg = `Le informamos que su solicitud #${idSolicitud} ("${activityName}") ha sido devuelta desde Presupuesto a la etapa de Revisión 2 para ajustes financieros.`;
      const solHtml = this.buildHtmlTemplate({
        title: 'Solicitud Devuelta a Revisión 2',
        badgeText: 'Ajustes en Curso',
        badgeColor: '#d97706',
        recipientName: applicantName,
        message: solMsg,
        details: [
          { label: 'Número de Solicitud', value: `#${idSolicitud}` },
          { label: 'Actividad', value: activityName },
          { label: 'Etapa Actual', value: 'Revisión 2' },
          formattedObs ? { label: 'Observaciones Registradas', value: formattedObs } : null
        ].filter(Boolean),
        callToActionText: 'El equipo de revisión está evaluando los ajustes correspondientes.'
      });

      notifications.push(await this.recordNotification({
        idSolicitud,
        event: 'solicitud_devuelta_aviso_solicitante',
        recipient: applicantEmail,
        subject: solSubject,
        text: `${solMsg}\n\nSolicitud: #${idSolicitud}\n${formattedObs}`,
        html: solHtml
      }));

      return notifications;
    }

    return notifications;
  }

  /**
   * Notifica al solicitante cuando su solicitud ya tiene ambas firmas y es remitida a Decanato.
   */
  async notifyDecanatoSubmissionToApplicant({ stage = {} }) {
    const idSolicitud = stage.id_solicitud;
    const fullStage = stage.nombre_solicitante ? stage : (await this.getStage(idSolicitud)) || stage;
    const applicantId = fullStage.id_usuario;
    const applicantName = fullStage.nombre_solicitante || 'Solicitante';
    const activityName = fullStage.nombre_actividad || 'Actividad de extensión';
    const applicantEmail = await this.getUserEmail(applicantId);

    const subject = `[Extensión] Solicitud #${idSolicitud} enviada a Decanato`;
    const message = `Su solicitud #${idSolicitud} ("${activityName}") cuenta con todas las firmas requeridas (solicitante y director de escuela) y ha sido remitida exitosamente a Decanato para la revisión final y expedición de resolución.`;
    const html = this.buildHtmlTemplate({
      title: 'Solicitud Remitida a Decanato',
      badgeText: 'En Trámite de Decanato',
      badgeColor: '#7c3aed',
      recipientName: applicantName,
      message,
      details: [
        { label: 'Número de Solicitud', value: `#${idSolicitud}` },
        { label: 'Actividad', value: activityName },
        { label: 'Estado', value: 'Pendiente de Resolución de Decanato' },
        { label: 'Fecha de Envío', value: new Date().toLocaleString('es-CO') }
      ],
      callToActionText: 'Le estaremos notificando por este medio cuando Decanato expida la resolución final.'
    });

    return this.recordNotification({
      idSolicitud,
      event: 'solicitud_enviada_decanato_solicitante',
      recipient: applicantEmail,
      subject,
      text: `${message}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}`,
      html
    });
  }

  /**
   * Notifica al solicitante cuando Decanato registra la decisión final (Corte de Decanato).
   */
  async notifyDecanatoResolution({
    stage = {},
    resultado = '',
    resolucion = '',
    observaciones = ''
  }) {
    const idSolicitud = stage.id_solicitud;
    const fullStage = stage.nombre_solicitante ? stage : (await this.getStage(idSolicitud)) || stage;
    const applicantId = fullStage.id_usuario;
    const applicantName = fullStage.nombre_solicitante || 'Solicitante';
    const activityName = fullStage.nombre_actividad || 'Actividad de extensión';
    const applicantEmail = await this.getUserEmail(applicantId);

    const isApproved = String(resultado).toLowerCase().includes('aprob');
    const badgeColor = isApproved ? '#16a34a' : '#dc2626';

    const subject = `[Extensión] Decisión Final Decanato - Solicitud #${idSolicitud} (${resultado})`;
    const message = `Decanato ha registrado la decisión final sobre su solicitud #${idSolicitud} ("${activityName}"). El resultado del trámite es: ${resultado}.`;
    const html = this.buildHtmlTemplate({
      title: 'Decisión Final de Decanato',
      badgeText: `Resultado: ${resultado}`,
      badgeColor,
      recipientName: applicantName,
      message,
      details: [
        { label: 'Número de Solicitud', value: `#${idSolicitud}` },
        { label: 'Actividad', value: activityName },
        { label: 'Resultado', value: resultado },
        resolucion ? { label: 'Resolución / Documento', value: resolucion } : null,
        observaciones ? { label: 'Observaciones de Decanato', value: observaciones } : null,
        { label: 'Fecha', value: new Date().toLocaleString('es-CO') }
      ].filter(Boolean),
      callToActionText: 'Puede ingresar a la plataforma institucional para consultar los detalles y descargar los soportes.'
    });

    return this.recordNotification({
      idSolicitud,
      event: 'solicitud_finalizada_solicitante',
      recipient: applicantEmail,
      subject,
      text: `${message}\n\nSolicitud: #${idSolicitud}\nResultado: ${resultado}\nResolución: ${resolucion}\nObservaciones: ${observaciones}`,
      html
    });
  }

  /**
   * Notifica cuando se registra una firma en la etapa de Firmas.
   * Cubre tanto la firma parcial (falta 1) como la firma completa (ambas listas).
   */
  async notifySignatureRegistered({
    stage = {},
    signature = {},
    bothSigned = false,
    signerUser = {}
  }) {
    const idSolicitud = stage.id_solicitud;
    const fullStage = stage.nombre_solicitante ? stage : (await this.getStage(idSolicitud)) || stage;
    const applicantId = fullStage.id_usuario;
    const applicantName = fullStage.nombre_solicitante || 'Solicitante';
    const activityName = fullStage.nombre_actividad || 'Actividad de extensión';
    const applicantEmail = await this.getUserEmail(applicantId);
    const directorEmails = await this.getRoleEmails('director');

    const signerType = signature.tipo_firma;
    const isSolicitanteSigner = signerType === 'solicitante';
    const signerRoleLabel = isSolicitanteSigner ? 'Solicitante / Coordinador(a)' : 'Director(a) de Escuela';
    const signerName = signature.nombre_firmante || signerUser?.name || signerRoleLabel;

    const notifications = [];

    if (!bothSigned) {
      // --- CASO 1: Firma parcial registrada (falta la otra firma) ---
      if (isSolicitanteSigner) {
        // 1.1 Notificar al Director de Escuela (requiere su aval)
        const dirSubject = `[Extensión] Solicitud #${idSolicitud} firmada por Solicitante - Pendiente de su firma de aval`;
        const dirMsg = `Le informamos que el(la) solicitante **${applicantName}** ha registrado exitosamente su firma en la solicitud de extensión **#${idSolicitud}** ("${activityName}").\n\nPara continuar con el trámite y remitir la solicitud a la dependencia de Presupuesto, se requiere su firma de aval institucional como Director(a) de la Unidad Académica.`;
        const dirHtml = this.buildHtmlTemplate({
          title: 'Firma de Aval Requerida - Director(a) de Escuela',
          badgeText: 'Firma Pendiente (1/2)',
          badgeColor: '#2563eb',
          recipientName: 'Director(a) de Escuela',
          message: dirMsg,
          details: [
            { label: 'Número de Solicitud', value: `#${idSolicitud}` },
            { label: 'Actividad', value: activityName },
            { label: 'Solicitante', value: applicantName },
            { label: 'Firmas Registradas', value: `${applicantName} (Solicitante)` },
            { label: 'Firma Pendiente', value: 'Director(a) de Escuela' },
            { label: 'Fecha de Registro', value: new Date().toLocaleString('es-CO') }
          ],
          callToActionText: 'Por favor ingrese al sistema para revisar los formularios oficiales y estampar su firma de aval.'
        });

        if (directorEmails.length === 0) {
          notifications.push(await this.recordNotification({
            idSolicitud,
            event: 'firma_solicitante_registrada_aviso_director',
            recipient: '',
            subject: dirSubject,
            text: dirMsg,
            html: dirHtml
          }));
        } else {
          for (const email of directorEmails) {
            notifications.push(await this.recordNotification({
              idSolicitud,
              event: 'firma_solicitante_registrada_aviso_director',
              recipient: email,
              subject: dirSubject,
              text: `${dirMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}\nSolicitante: ${applicantName}`,
              html: dirHtml
            }));
          }
        }

        // 1.2 Confirmación al Solicitante
        const solSubject = `[Extensión] Confirmación: Firma registrada en solicitud #${idSolicitud}`;
        const solMsg = `Su firma institucional para la solicitud de extensión **#${idSolicitud}** ("${activityName}") ha sido registrada satisfactoriamente en el sistema.\n\nHemos notificado a la Dirección de Escuela para que proceda con la firma de aval correspondiente. Una vez se cuente con ambas firmas, la solicitud continuará su trámite hacia Presupuesto.`;
        const solHtml = this.buildHtmlTemplate({
          title: 'Firma Registrada Exitosamente',
          badgeText: 'Firma Registrada (1/2)',
          badgeColor: '#16a34a',
          recipientName: applicantName,
          message: solMsg,
          details: [
            { label: 'Número de Solicitud', value: `#${idSolicitud}` },
            { label: 'Actividad', value: activityName },
            { label: 'Su Firma', value: 'Registrada correctamente' },
            { label: 'Siguiente Paso', value: 'Pendiente de aval del Director(a) de Escuela' }
          ],
          callToActionText: 'Le estaremos notificando en cuanto el Director(a) registre su firma institucional.'
        });
        notifications.push(await this.recordNotification({
          idSolicitud,
          event: 'firma_solicitante_confirmacion',
          recipient: applicantEmail,
          subject: solSubject,
          text: `${solMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}`,
          html: solHtml
        }));

      } else {
        // Firmó el Director, falta el Solicitante
        // 1.3 Notificar al Solicitante que el Director ya firmó
        const solSubject = `[Extensión] Solicitud #${idSolicitud} avalada por Dirección - Pendiente de su firma`;
        const solMsg = `Le informamos que el(la) Director(a) de Escuela (${signerName}) ha registrado su firma de aval en la solicitud de extensión **#${idSolicitud}** ("${activityName}").\n\nPara completar el proceso y enviar la solicitud a Presupuesto, por favor ingrese a la plataforma y registre su firma institucional.`;
        const solHtml = this.buildHtmlTemplate({
          title: 'Firma de Solicitante Requerida',
          badgeText: 'Firma Pendiente (1/2)',
          badgeColor: '#2563eb',
          recipientName: applicantName,
          message: solMsg,
          details: [
            { label: 'Número de Solicitud', value: `#${idSolicitud}` },
            { label: 'Actividad', value: activityName },
            { label: 'Firma de Aval', value: `Registrada por Dirección de Escuela (${signerName})` },
            { label: 'Firma Pendiente', value: `${applicantName} (Solicitante)` }
          ],
          callToActionText: 'Por favor ingrese al sistema para registrar su firma y remitir la solicitud a Presupuesto.'
        });
        notifications.push(await this.recordNotification({
          idSolicitud,
          event: 'firma_director_registrada_aviso_solicitante',
          recipient: applicantEmail,
          subject: solSubject,
          text: `${solMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}`,
          html: solHtml
        }));

        // 1.4 Confirmación al Director
        const dirSubject = `[Extensión] Aval registrado en solicitud #${idSolicitud}`;
        const dirMsg = `Su firma de aval para la solicitud de extensión **#${idSolicitud}** ("${activityName}") de ${applicantName} ha sido registrada con éxito. Se encuentra pendiente la firma del solicitante para completar el proceso.`;
        const dirHtml = this.buildHtmlTemplate({
          title: 'Aval Registrado Exitosamente',
          badgeText: 'Aval Registrado',
          badgeColor: '#16a34a',
          recipientName: 'Director(a) de Escuela',
          message: dirMsg,
          details: [
            { label: 'Número de Solicitud', value: `#${idSolicitud}` },
            { label: 'Actividad', value: activityName },
            { label: 'Solicitante', value: applicantName },
            { label: 'Estado', value: 'Pendiente de firma del solicitante' }
          ]
        });
        for (const email of directorEmails) {
          notifications.push(await this.recordNotification({
            idSolicitud,
            event: 'firma_director_confirmacion',
            recipient: email,
            subject: dirSubject,
            text: `${dirMsg}\n\nSolicitud: #${idSolicitud}`,
            html: dirHtml
          }));
        }
      }

    } else {
      // --- CASO 2: Ambas firmas completas ---
      // 2.1 Al Solicitante
      const solSubject = `[Extensión] ¡Firmas completadas con éxito! - Solicitud #${idSolicitud}`;
      const solMsg = `Nos complace informarle que la solicitud de extensión **#${idSolicitud}** ("${activityName}") cuenta con ambas firmas institucionales requeridas (Solicitante y Director de Escuela).\n\nLos formularios oficiales en formato PDF han sido consolidados con las firmas digitales correspondientes y la solicitud se encuentra lista para ser remitida a la dependencia de Presupuesto.`;
      const solHtml = this.buildHtmlTemplate({
        title: 'Proceso de Firmas Completado',
        badgeText: 'Firmas Completas (2/2)',
        badgeColor: '#16a34a',
        recipientName: applicantName,
        message: solMsg,
        details: [
          { label: 'Número de Solicitud', value: `#${idSolicitud}` },
          { label: 'Actividad', value: activityName },
          { label: 'Solicitante', value: applicantName },
          { label: 'Firmas Registradas', value: 'Solicitante y Director(a) de Escuela (2 de 2)' },
          { label: 'Estado', value: 'Lista para remisión a Presupuesto' },
          { label: 'Fecha de Finalización de Firmas', value: new Date().toLocaleString('es-CO') }
        ],
        callToActionText: 'La solicitud será remitida a Presupuesto para la revisión financiera. Le notificaremos tan pronto avance.'
      });
      notifications.push(await this.recordNotification({
        idSolicitud,
        event: 'firmas_completas_solicitante',
        recipient: applicantEmail,
        subject: solSubject,
        text: `${solMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}`,
        html: solHtml
      }));

      // 2.2 Al Director de Escuela
      const dirSubject = `[Extensión] Firmas completadas - Solicitud #${idSolicitud} (${activityName})`;
      const dirMsg = `Le informamos que la solicitud de extensión **#${idSolicitud}** ("${activityName}"), presentada por ${applicantName}, cuenta con ambas firmas institucionales (Solicitante y Director de Escuela) y se encuentra lista para su revisión en la dependencia de Presupuesto.`;
      const dirHtml = this.buildHtmlTemplate({
        title: 'Firmas Completadas - Lista para Presupuesto',
        badgeText: 'Firmas Completas (2/2)',
        badgeColor: '#16a34a',
        recipientName: 'Director(a) de Escuela',
        message: dirMsg,
        details: [
          { label: 'Número de Solicitud', value: `#${idSolicitud}` },
          { label: 'Actividad', value: activityName },
          { label: 'Solicitante', value: applicantName },
          { label: 'Firmas', value: 'Ambas firmas registradas satisfactoriamente' },
          { label: 'Próxima Etapa', value: 'Presupuesto' }
        ]
      });
      for (const email of directorEmails) {
        notifications.push(await this.recordNotification({
          idSolicitud,
          event: 'firmas_completas_director',
          recipient: email,
          subject: dirSubject,
          text: `${dirMsg}\n\nSolicitud: #${idSolicitud}`,
          html: dirHtml
        }));
      }
    }

    return notifications;
  }

  /**
   * Notifica el paso de la solicitud desde Firmas hacia Presupuesto:
   * 1. A los destinatarios de Presupuesto y adicionales (con PDFs firmados adjuntos y plantilla HTML).
   * 2. Al Solicitante (confirmando que su solicitud firmada ya está en Presupuesto).
   * 3. Al Director de Escuela (aviso informativo de avance).
   */
  async notifyPresupuestoSubmission({
    stage = {},
    recipients = [],
    customSubject = '',
    customBody = '',
    attachments = []
  }) {
    const idSolicitud = stage.id_solicitud;
    const fullStage = stage.nombre_solicitante ? stage : (await this.getStage(idSolicitud)) || stage;
    const applicantId = fullStage.id_usuario;
    const applicantName = fullStage.nombre_solicitante || 'Solicitante';
    const activityName = fullStage.nombre_actividad || 'Actividad de extensión';
    const applicantEmail = await this.getUserEmail(applicantId);
    const directorEmails = await this.getRoleEmails('director');

    const notifications = [];

    // 1. Correo a la dependencia de Presupuesto y destinatarios designados
    const teamSubject = customSubject || `[Extensión] Solicitud #${idSolicitud} firmada y remitida para revisión presupuestal - ${activityName}`;
    const defaultTeamMsg = `Cordial saludo.\n\nPor medio del presente correo se remite para su conocimiento, evaluación y trámite presupuestal la solicitud de extensión **#${idSolicitud}** ("${activityName}"), presentada por el(la) docente/coordinador(a) **${applicantName}**.\n\nLa solicitud ya cuenta con el aval y las firmas institucionales completas (Solicitante y Director(a) de Escuela). Se adjuntan los formularios oficiales debidamente diligenciados y firmados en formato PDF para su respectiva revisión financiera.`;
    const teamMsg = customBody || defaultTeamMsg;

    const teamHtml = this.buildHtmlTemplate({
      title: 'Solicitud de Extensión para Revisión Presupuestal',
      badgeText: 'Revisión Presupuestal',
      badgeColor: '#2563eb',
      recipientName: 'Equipo de Presupuesto',
      message: teamMsg,
      details: [
        { label: 'Número de Solicitud', value: `#${idSolicitud}` },
        { label: 'Nombre de la Actividad', value: activityName },
        { label: 'Solicitante / Coordinador(a)', value: applicantName },
        { label: 'Estado de Firmas', value: 'Firmas Completas (Solicitante y Director de Escuela)' },
        { label: 'Documentos Adjuntos', value: attachments.length > 0 ? attachments.map(a => a.filename || a.name || 'Formulario PDF').join(', ') : 'Formularios oficiales firmados (PDF)' },
        { label: 'Fecha de Remisión', value: new Date().toLocaleString('es-CO') }
      ],
      callToActionText: 'Por favor ingrese al Sistema de Gestión de Extensión para consultar la información financiera, emitir concepto y continuar el flujo del trámite.'
    });

    const uniqueRecipients = [...new Set(recipients.map(item => text(item).toLowerCase()).filter(Boolean))];
    if (uniqueRecipients.length === 0) {
      notifications.push(await this.recordNotification({
        idSolicitud,
        event: 'solicitud_enviada_presupuesto',
        recipient: '',
        subject: teamSubject,
        text: teamMsg,
        html: teamHtml,
        attachments
      }));
    } else {
      for (const email of uniqueRecipients) {
        notifications.push(await this.recordNotification({
          idSolicitud,
          event: 'solicitud_enviada_presupuesto',
          recipient: email,
          subject: teamSubject,
          text: `${teamMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}\nSolicitante: ${applicantName}`,
          html: teamHtml,
          attachments
        }));
      }
    }

    // 2. Correo al Solicitante
    const solSubject = `[Extensión] Su solicitud #${idSolicitud} ha sido remitida a Presupuesto`;
    const solMsg = `Le informamos que su solicitud de extensión **#${idSolicitud}** ("${activityName}") ha completado satisfactoriamente la etapa de firmas institucionales y ha sido remitida a la dependencia de **Presupuesto** para su revisión financiera y asignación de rubros correspondiente.`;
    const solHtml = this.buildHtmlTemplate({
      title: 'Solicitud Remitida a Presupuesto',
      badgeText: 'En Revisión Presupuestal',
      badgeColor: '#16a34a',
      recipientName: applicantName,
      message: solMsg,
      details: [
        { label: 'Número de Solicitud', value: `#${idSolicitud}` },
        { label: 'Actividad', value: activityName },
        { label: 'Etapa Anterior', value: 'Firmas (Completadas: Solicitante y Director)' },
        { label: 'Etapa Actual', value: 'Presupuesto' },
        { label: 'Fecha de Envío', value: new Date().toLocaleString('es-CO') }
      ],
      callToActionText: 'Puede hacer seguimiento al avance de su trámite ingresando a la plataforma. Le estaremos notificando por este medio tan pronto culmine la revisión presupuestal.'
    });

    notifications.push(await this.recordNotification({
      idSolicitud,
      event: 'solicitud_firmada_avanza_presupuesto_solicitante',
      recipient: applicantEmail,
      subject: solSubject,
      text: `${solMsg}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}`,
      html: solHtml
    }));

    // 3. Correo al Director de Escuela (aviso informativo)
    const dirSubject = `[Extensión] Solicitud #${idSolicitud} (${activityName}) remitida a Presupuesto`;
    const dirMsg = `Le informamos que la solicitud de extensión **#${idSolicitud}** ("${activityName}"), presentada por ${applicantName} y que contó con su aval institucional como Director(a) de Escuela, ha sido remitida a la dependencia de Presupuesto para continuar su proceso.`;
    const dirHtml = this.buildHtmlTemplate({
      title: 'Solicitud Avalada Remitida a Presupuesto',
      badgeText: 'Avanza a Presupuesto',
      badgeColor: '#2563eb',
      recipientName: 'Director(a) de Escuela',
      message: dirMsg,
      details: [
        { label: 'Número de Solicitud', value: `#${idSolicitud}` },
        { label: 'Actividad', value: activityName },
        { label: 'Solicitante', value: applicantName },
        { label: 'Estado', value: 'En revisión por la dependencia de Presupuesto' }
      ]
    });

    for (const email of directorEmails) {
      notifications.push(await this.recordNotification({
        idSolicitud,
        event: 'solicitud_firmada_avanza_presupuesto_director',
        recipient: email,
        subject: dirSubject,
        text: `${dirMsg}\n\nSolicitud: #${idSolicitud}`,
        html: dirHtml
      }));
    }

    return notifications;
  }

  /**
   * Notifica un recordatorio de firma pendiente (del solicitante al director o viceversa).
   */
  async notifySignatureReminder({
    stage = {},
    target = '',
    requestedBy = {}
  }) {
    const idSolicitud = stage.id_solicitud;
    const fullStage = stage.nombre_solicitante ? stage : (await this.getStage(idSolicitud)) || stage;
    const applicantId = fullStage.id_usuario;
    const applicantName = fullStage.nombre_solicitante || 'Solicitante';
    const activityName = fullStage.nombre_actividad || 'Actividad de extensión';
    const applicantEmail = await this.getUserEmail(applicantId);
    const directorEmails = await this.getRoleEmails('director');

    const requesterName = requestedBy?.name || 'Un usuario participante';

    if (target === 'director_escuela') {
      const subject = `[Extensión] Recordatorio: Firma de aval pendiente para la solicitud #${idSolicitud}`;
      const message = `Le recordamos amablemente que la solicitud de extensión **#${idSolicitud}** ("${activityName}"), presentada por el(la) docente **${applicantName}**, se encuentra a la espera de su firma de aval como Director(a) de Escuela.\n\nEste aval es indispensable para que los formularios oficiales puedan consolidarse y remitirse a la dependencia de Presupuesto.`;
      const html = this.buildHtmlTemplate({
        title: 'Recordatorio de Firma de Aval',
        badgeText: 'Firma Pendiente',
        badgeColor: '#ea580c',
        recipientName: 'Director(a) de Escuela',
        message,
        details: [
          { label: 'Número de Solicitud', value: `#${idSolicitud}` },
          { label: 'Actividad', value: activityName },
          { label: 'Solicitante', value: applicantName },
          { label: 'Solicitado por', value: requesterName },
          { label: 'Firma Requerida', value: 'Director(a) de Escuela' }
        ],
        callToActionText: 'Por favor ingrese a su bandeja de firmas en el sistema para avalar la solicitud.'
      });

      const notifications = [];
      if (directorEmails.length === 0) {
        notifications.push(await this.recordNotification({
          idSolicitud,
          event: 'recordatorio_firma_director',
          recipient: '',
          subject,
          text: message,
          html
        }));
      } else {
        for (const email of directorEmails) {
          notifications.push(await this.recordNotification({
            idSolicitud,
            event: 'recordatorio_firma_director',
            recipient: email,
            subject,
            text: `${message}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}`,
            html
          }));
        }
      }
      return notifications;
    } else {
      const subject = `[Extensión] Recordatorio: Firma pendiente en su solicitud #${idSolicitud}`;
      const message = `Le recordamos amablemente que su solicitud de extensión **#${idSolicitud}** ("${activityName}") se encuentra en espera de su firma institucional como Solicitante / Coordinador(a).\n\nUna vez registre su firma, la solicitud podrá continuar su trámite hacia la dependencia de Presupuesto.`;
      const html = this.buildHtmlTemplate({
        title: 'Recordatorio de Firma Institucional',
        badgeText: 'Firma Pendiente',
        badgeColor: '#ea580c',
        recipientName: applicantName,
        message,
        details: [
          { label: 'Número de Solicitud', value: `#${idSolicitud}` },
          { label: 'Actividad', value: activityName },
          { label: 'Firma Requerida', value: `${applicantName} (Solicitante)` },
          { label: 'Etapa Actual', value: 'Firmas' }
        ],
        callToActionText: 'Por favor ingrese a la plataforma institucional para cargar o registrar su firma.'
      });

      return [await this.recordNotification({
        idSolicitud,
        event: 'recordatorio_firma_solicitante',
        recipient: applicantEmail,
        subject,
        text: `${message}\n\nSolicitud: #${idSolicitud}\nActividad: ${activityName}`,
        html
      })];
    }
  }

}

module.exports = new WorkflowNotificationService();
