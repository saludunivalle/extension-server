const fs = require('fs');
const os = require('os');
const path = require('path');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const reportService = require('./reportService');
const driveService = require('./driveService');

const extractDriveFileId = (value) => {
  const raw = String(value || '');
  return raw.match(/\/d\/([a-zA-Z0-9_-]+)/)?.[1]
    || raw.match(/[?&]id=([a-zA-Z0-9_-]+)/)?.[1]
    || raw;
};

const isPng = (buffer) => buffer.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
const isJpeg = (buffer) => buffer[0] === 0xff && buffer[1] === 0xd8;

const getCompletedFormNumbers = (stage) => {
  let states = {};
  try {
    states = typeof stage.estado_formularios === 'string'
      ? JSON.parse(stage.estado_formularios || '{}')
      : (stage.estado_formularios || {});
  } catch {
    states = {};
  }
  return [1, 2, 3, 4].filter((formNumber) => {
    const status = String(states[String(formNumber)] || '').toLowerCase().replace('ó', 'o');
    return ['completado', 'aprobado', 'enviado a revision'].includes(status);
  });
};

async function addSignaturePage(pdfBytes, signatures, requestId, formNumber) {
  const pdf = await PDFDocument.load(pdfBytes);
  const page = pdf.addPage([595, 190]);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  page.drawText(`Firmas de la solicitud ${requestId}`, { x: 36, y: 155, size: 14, font: bold, color: rgb(0.12, 0.18, 0.28) });
  page.drawText(`Formulario ${formNumber}`, { x: 36, y: 135, size: 10, font, color: rgb(0.35, 0.4, 0.48) });

  const positions = [70, 320];
  for (let index = 0; index < 2; index += 1) {
    const signature = signatures[index];
    const x = positions[index];
    page.drawRectangle({ x, y: 38, width: 205, height: 78, borderWidth: 1, borderColor: rgb(0.78, 0.82, 0.88) });
    if (signature?.evidencia_firma) {
      const imageBuffer = await driveService.downloadFileBuffer(extractDriveFileId(signature.evidencia_firma));
      const image = isPng(imageBuffer) ? await pdf.embedPng(imageBuffer) : isJpeg(imageBuffer) ? await pdf.embedJpg(imageBuffer) : null;
      if (image) {
        const scale = Math.min(170 / image.width, 54 / image.height);
        page.drawImage(image, { x: x + 17, y: 55, width: image.width * scale, height: image.height * scale });
      }
    }
    page.drawText(signature?.nombre_firmante || 'Firma pendiente', { x: x + 8, y: 43, size: 9, font: bold, color: rgb(0.12, 0.18, 0.28), maxWidth: 189 });
    page.drawText(signature?.tipo_firma === 'director_escuela' ? 'Director de escuela' : 'Solicitante', { x: x + 8, y: 26, size: 8, font, color: rgb(0.35, 0.4, 0.48) });
  }
  return pdf.save();
}

async function createDecanatoAttachments({ stage, signatures }) {
  const tempDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'decanato-pdfs-'));
  const completedForms = getCompletedFormNumbers(stage);
  const orderedSignatures = [
    signatures.find((item) => item.tipo_firma === 'solicitante'),
    signatures.find((item) => item.tipo_firma === 'director_escuela'),
  ];
  const attachments = [];
  try {
    for (const formNumber of completedForms) {
      let reportId;
      try {
        const reportResult = await reportService.generateReport(String(stage.id_solicitud), formNumber);
        const reportLink = reportResult?.link || reportResult;
        reportId = extractDriveFileId(reportLink);
      } catch (remoteReportError) {
        // Algunas plantillas remotas pueden haber sido movidas o eliminadas.
        // Usamos la plantilla XLSX local como respaldo para no bloquear el envío.
        console.warn(`No fue posible generar el reporte remoto del formulario ${formNumber}. Usando plantilla local: ${remoteReportError.message}`);
        const localReport = await reportService.downloadReportFile(String(stage.id_solicitud), formNumber, 'view');
        reportId = await driveService.uploadXlsxAsGoogleSheet(
          localReport.filePath,
          `Formulario${formNumber}_${stage.id_solicitud}_respaldo`
        );
        await fs.promises.unlink(localReport.filePath).catch(() => {});
      }
      if (!reportId || reportId === '[object Object]') {
        throw new Error(`No se pudo obtener el ID de Drive para el formulario ${formNumber}`);
      }
      const sourcePath = path.join(tempDir, `formulario-${formNumber}-source.pdf`);
      const outputPath = path.join(tempDir, `Formulario${formNumber}_${stage.id_solicitud}_firmado.pdf`);
      await driveService.exportFileAsPdf(reportId, sourcePath);
      const sourceBytes = await fs.promises.readFile(sourcePath);
      const signedBytes = await addSignaturePage(sourceBytes, orderedSignatures, stage.id_solicitud, formNumber);
      await fs.promises.writeFile(outputPath, signedBytes);
      attachments.push({ filename: path.basename(outputPath), path: outputPath, contentType: 'application/pdf' });
    }
    return { attachments, tempDir };
  } catch (error) {
    await fs.promises.rm(tempDir, { recursive: true, force: true });
    throw error;
  }
}

module.exports = { createDecanatoAttachments };
