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

const isPng = (buffer) => buffer && buffer.length >= 8 && buffer.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
const isJpeg = (buffer) => buffer && buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xd8;

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

async function getSignatureBuffer(signature) {
  if (!signature?.evidencia_firma) return null;
  const raw = String(signature.evidencia_firma).trim();
  try {
    if (fs.existsSync(raw)) {
      return await fs.promises.readFile(raw);
    }
  } catch {
    // ignorar error de path local
  }
  const fileId = extractDriveFileId(raw);
  if (fileId) {
    try {
      return await driveService.downloadFileBuffer(fileId);
    } catch (error) {
      console.warn(`Error al descargar archivo de firma ${fileId}:`, error.message);
    }
  }
  return null;
}

async function drawSignatureBox({ page, signature, x, y, maxWidth, maxHeight }) {
  if (!signature) return;
  const imageBuffer = await getSignatureBuffer(signature);
  if (!imageBuffer || imageBuffer.length === 0) return;
  const doc = page.doc;
  const image = isPng(imageBuffer)
    ? await doc.embedPng(imageBuffer)
    : isJpeg(imageBuffer)
    ? await doc.embedJpg(imageBuffer)
    : null;
  if (!image) return;

  const scale = Math.min(maxWidth / image.width, maxHeight / image.height);
  const imgW = image.width * scale;
  const imgH = image.height * scale;
  const posX = x + (maxWidth - imgW) / 2;
  const posY = y + (maxHeight - imgH) / 2;

  page.drawImage(image, {
    x: posX,
    y: posY,
    width: imgW,
    height: imgH,
  });
}

async function addSignaturePageToDoc(pdf, signatures, requestId, formNumber) {
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
      const imageBuffer = await getSignatureBuffer(signature);
      if (imageBuffer) {
        const image = isPng(imageBuffer)
          ? await pdf.embedPng(imageBuffer)
          : isJpeg(imageBuffer)
          ? await pdf.embedJpg(imageBuffer)
          : null;
        if (image) {
          const scale = Math.min(170 / image.width, 54 / image.height);
          page.drawImage(image, { x: x + 17, y: 55, width: image.width * scale, height: image.height * scale });
        }
      }
    }
    page.drawText(signature?.nombre_firmante || 'Firma pendiente', { x: x + 8, y: 43, size: 9, font: bold, color: rgb(0.12, 0.18, 0.28), maxWidth: 189 });
    page.drawText(signature?.tipo_firma === 'director_escuela' ? 'Director de escuela' : 'Solicitante', { x: x + 8, y: 26, size: 8, font, color: rgb(0.35, 0.4, 0.48) });
  }
  return pdf.save();
}

function sanitizePdfPages(pdf) {
  let pages = pdf.getPages();
  while (pages.length > 1) {
    const lastPage = pages[pages.length - 1];
    const contents = lastPage.node.Contents();
    if (!contents || contents.size() === 0) {
      pdf.removePage(pages.length - 1);
      pages = pdf.getPages();
    } else {
      break;
    }
  }
}

async function embedSignaturesInPdf(pdfBytes, orderedSignatures = [], requestId, formNumber) {
  const formNum = parseInt(formNumber, 10);
  const pdf = await PDFDocument.load(pdfBytes);
  sanitizePdfPages(pdf);
  const pages = pdf.getPages();
  if (pages.length === 0) return pdf.save();

  const solicitanteSig = orderedSignatures.find((item) => item?.tipo_firma === 'solicitante');
  const directorSig = orderedSignatures.find((item) => item?.tipo_firma === 'director_escuela');

  if (formNum === 1) {
    // Formulario 1:
    // Solicitante: justo arriba de "Nombre: Coordinador de la Actividad de Extensión"
    // Director: justo arriba de "Nombre: Visto Bueno de la Unidad Académica"
    const targetPage = pages[pages.length - 1];
    if (solicitanteSig) {
      await drawSignatureBox({
        page: targetPage,
        signature: solicitanteSig,
        x: 60,
        y: 128,
        maxWidth: 175,
        maxHeight: 36,
      });
    }
    if (directorSig) {
      await drawSignatureBox({
        page: targetPage,
        signature: directorSig,
        x: 295,
        y: 128,
        maxWidth: 160,
        maxHeight: 36,
      });
    }
    return pdf.save();
  }

  if (formNum === 2) {
    // Formulario 2:
    // Solicitante: justo arriba de "Coordinador de la Actividad de Extensión"
    // Director: justo arriba de "Visto Bueno de la Unidad Académica"
    const targetPage = pages[pages.length - 1];
    if (solicitanteSig) {
      await drawSignatureBox({
        page: targetPage,
        signature: solicitanteSig,
        x: 25,
        y: 47,
        maxWidth: 200,
        maxHeight: 44,
      });
    }
    if (directorSig) {
      await drawSignatureBox({
        page: targetPage,
        signature: directorSig,
        x: 320,
        y: 47,
        maxWidth: 200,
        maxHeight: 44,
      });
    }
    return pdf.save();
  }

  // Formularios 3 y 4: mantener ubicación donde se ubicaba la firma (página de firmas agregada al final)
  return addSignaturePageToDoc(pdf, orderedSignatures, requestId, formNum);
}

async function generateFormPdfBuffer({ stage, formNumber, signatures = [] }) {
  const formNum = parseInt(formNumber, 10);
  const orderedSignatures = [
    signatures.find((item) => item?.tipo_firma === 'solicitante'),
    signatures.find((item) => item?.tipo_firma === 'director_escuela'),
  ].filter(Boolean);

  let reportId;
  try {
    const reportResult = await reportService.generateReport(String(stage.id_solicitud), formNum);
    const reportLink = reportResult?.link || reportResult;
    reportId = extractDriveFileId(reportLink);
  } catch (remoteReportError) {
    console.warn(`No fue posible generar el reporte remoto del formulario ${formNum}. Usando plantilla local: ${remoteReportError.message}`);
    const localReport = await reportService.downloadReportFile(String(stage.id_solicitud), formNum, 'view');
    reportId = await driveService.uploadXlsxAsGoogleSheet(
      localReport.filePath,
      `Formulario${formNum}_${stage.id_solicitud}_respaldo`
    );
    await fs.promises.unlink(localReport.filePath).catch(() => {});
  }

  if (!reportId || reportId === '[object Object]') {
    throw new Error(`No se pudo obtener el ID de Drive para el formulario ${formNum}`);
  }

  const tempSourcePath = path.join(os.tmpdir(), `form-${stage.id_solicitud}-${formNum}-${Date.now()}.pdf`);
  try {
    await driveService.exportFileAsPdf(reportId, tempSourcePath, { formNumber: formNum });
    const sourceBytes = await fs.promises.readFile(tempSourcePath);
    const signedBytes = await embedSignaturesInPdf(sourceBytes, orderedSignatures, stage.id_solicitud, formNum);
    return Buffer.from(signedBytes);
  } finally {
    await fs.promises.unlink(tempSourcePath).catch(() => {});
  }
}

async function createDecanatoAttachments({ stage, signatures }) {
  const tempDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'decanato-pdfs-'));
  const completedForms = getCompletedFormNumbers(stage);
  const attachments = [];
  try {
    for (const formNumber of completedForms) {
      const outputPath = path.join(tempDir, `Formulario${formNumber}_${stage.id_solicitud}_firmado.pdf`);
      const signedBytes = await generateFormPdfBuffer({ stage, formNumber, signatures });
      await fs.promises.writeFile(outputPath, signedBytes);
      attachments.push({ filename: path.basename(outputPath), path: outputPath, contentType: 'application/pdf' });
    }
    return { attachments, tempDir };
  } catch (error) {
    await fs.promises.rm(tempDir, { recursive: true, force: true });
    throw error;
  }
}

module.exports = {
  createDecanatoAttachments,
  generateFormPdfBuffer,
  embedSignaturesInPdf,
  getCompletedFormNumbers,
};
