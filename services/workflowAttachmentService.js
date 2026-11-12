const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');
const { PDFDocument, PDFName, PDFArray, PDFRawStream, StandardFonts, rgb } = require('pdf-lib');
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

const IDENTITY_MATRIX = [1, 0, 0, 1, 0, 0];

const SIGNATURE_FORM_LAYOUTS = {
  1: {
    hasNameRow: true,
    maxWidth: 170,
    maxHeight: 28,
    offsetAboveName: 8,
    offsetAboveTitle: 22,
    fallbackBottomOffset: 58,
    fallbackLeftX: 60,
    fallbackRightX: 295,
    fallbackLeftWidth: 175,
    fallbackRightWidth: 160,
    fallbackY: 220,
  },
  2: {
    hasNameRow: false,
    maxWidth: 185,
    maxHeight: 34,
    offsetAboveTitle: 10,
    fallbackBottomOffset: 12,
    fallbackLeftX: 35,
    fallbackRightX: 325,
    fallbackLeftWidth: 190,
    fallbackRightWidth: 190,
    fallbackY: 60,
  },
};

const normalizePdfText = (value) =>
  String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();

const multiplyMatrices = (m1, m2) => [
  m1[0] * m2[0] + m1[1] * m2[2],
  m1[0] * m2[1] + m1[1] * m2[3],
  m1[2] * m2[0] + m1[3] * m2[2],
  m1[2] * m2[1] + m1[3] * m2[3],
  m1[4] * m2[0] + m1[5] * m2[2] + m2[4],
  m1[4] * m2[1] + m1[5] * m2[3] + m2[5],
];

const applyMatrix = (m, x, y) => [
  m[0] * x + m[2] * y + m[4],
  m[1] * x + m[3] * y + m[5],
];

function decodeStreamBytes(pdf, streamObj) {
  const stream = pdf.context.lookup(streamObj);
  if (!stream || typeof stream.contents === 'undefined') return Buffer.alloc(0);
  const raw = Buffer.from(stream.contents);
  const filterObj = stream.dict?.get(PDFName.of('Filter'));
  const resolvedFilter = filterObj ? pdf.context.lookup(filterObj) : null;
  let isFlate = false;
  if (resolvedFilter instanceof PDFName) {
    isFlate = resolvedFilter.asString() === '/FlateDecode';
  } else if (resolvedFilter instanceof PDFArray) {
    for (let i = 0; i < resolvedFilter.size(); i += 1) {
      const item = pdf.context.lookup(resolvedFilter.get(i));
      if (item instanceof PDFName && item.asString() === '/FlateDecode') {
        isFlate = true;
        break;
      }
    }
  }
  if (!isFlate) return raw;
  try {
    return zlib.inflateSync(raw);
  } catch {
    return raw;
  }
}

const hexToUtf16 = (hex) => {
  const clean = hex.replace(/[^0-9a-fA-F]/g, '');
  if (!clean) return '';
  const padded = clean.length % 2 === 1 ? `0${clean}` : clean;
  let result = '';
  if (padded.length >= 4 && padded.length % 4 === 0) {
    for (let i = 0; i < padded.length; i += 4) {
      const code = parseInt(padded.slice(i, i + 4), 16);
      if (code > 0) result += String.fromCharCode(code);
    }
    return result;
  }
  for (let i = 0; i < padded.length; i += 2) {
    const code = parseInt(padded.slice(i, i + 2), 16);
    if (code > 0) result += String.fromCharCode(code);
  }
  return result;
};

function parseToUnicodeCMap(cmapText) {
  const map = new Map();
  let codeBytes = 2;
  const csMatch = cmapText.match(/begincodespacerange([\s\S]*?)endcodespacerange/);
  if (csMatch) {
    const hexMatch = csMatch[1].match(/<([0-9a-fA-F]+)>/);
    if (hexMatch) {
      codeBytes = Math.max(1, Math.floor(hexMatch[1].length / 2));
    }
  }

  const bfCharRe = /beginbfchar([\s\S]*?)endbfchar/g;
  let m;
  while ((m = bfCharRe.exec(cmapText)) !== null) {
    const pairRe = /<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/g;
    let p;
    while ((p = pairRe.exec(m[1])) !== null) {
      const src = parseInt(p[1], 16);
      codeBytes = Math.max(codeBytes, Math.ceil(p[1].length / 2));
      map.set(src, hexToUtf16(p[2]));
    }
  }

  const bfRangeRe = /beginbfrange([\s\S]*?)endbfrange/g;
  while ((m = bfRangeRe.exec(cmapText)) !== null) {
    const block = m[1];
    const lineRe = /<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(?:<([0-9a-fA-F]+)>|\[([^\]]+)\])/g;
    let r;
    while ((r = lineRe.exec(block)) !== null) {
      const start = parseInt(r[1], 16);
      const end = parseInt(r[2], 16);
      codeBytes = Math.max(codeBytes, Math.ceil(r[1].length / 2));
      if (r[3]) {
        const dstBase = parseInt(r[3], 16);
        for (let code = start; code <= end; code += 1) {
          map.set(code, String.fromCodePoint(dstBase + (code - start)));
        }
      } else if (r[4]) {
        const dstHexes = [...r[4].matchAll(/<([0-9a-fA-F]+)>/g)].map((h) => h[1]);
        for (let i = 0; i <= end - start && i < dstHexes.length; i += 1) {
          map.set(start + i, hexToUtf16(dstHexes[i]));
        }
      }
    }
  }
  return { map, codeBytes };
}

function buildFontDecoders(pdf, resources) {
  const decoders = new Map();
  if (!resources) return decoders;
  const resolvedRes = pdf.context.lookup(resources);
  const fontDict = resolvedRes?.get?.(PDFName.of('Font'))
    ? pdf.context.lookup(resolvedRes.get(PDFName.of('Font')))
    : null;
  if (!fontDict || typeof fontDict.entries !== 'function') return decoders;
  for (const [nameObj, fontRef] of fontDict.entries()) {
    const fontKey = nameObj.asString();
    const font = pdf.context.lookup(fontRef);
    const toUnicodeRef = font?.get?.(PDFName.of('ToUnicode'));
    if (toUnicodeRef) {
      const cmapBytes = decodeStreamBytes(pdf, toUnicodeRef);
      decoders.set(fontKey, parseToUnicodeCMap(cmapBytes.toString('latin1')));
    }
  }
  return decoders;
}

function decodePdfBytesWithFont(bytes, fontDecoder) {
  if (!bytes || bytes.length === 0) return '';
  if (fontDecoder && fontDecoder.map.size > 0) {
    const step = fontDecoder.codeBytes || 1;
    let out = '';
    for (let i = 0; i < bytes.length; i += step) {
      let code = 0;
      for (let b = 0; b < step && i + b < bytes.length; b += 1) {
        code = (code << 8) | bytes[i + b];
      }
      out += fontDecoder.map.get(code) ?? (code >= 32 && code < 127 ? String.fromCharCode(code) : '');
    }
    return out;
  }
  return Buffer.from(bytes).toString('latin1');
}

function tokenizePdfStream(buffer) {
  const tokens = [];
  const len = buffer.length;
  let i = 0;

  const isWs = (b) => b === 0x20 || b === 0x0a || b === 0x0d || b === 0x09 || b === 0x0c || b === 0x00;
  const isDelim = (b) => isWs(b) || b === 0x28 || b === 0x29 || b === 0x3c || b === 0x3e || b === 0x5b || b === 0x5d || b === 0x2f || b === 0x25;

  while (i < len) {
    const b = buffer[i];
    if (isWs(b)) {
      i += 1;
      continue;
    }
    if (b === 0x25) { // % comment
      while (i < len && buffer[i] !== 0x0a && buffer[i] !== 0x0d) i += 1;
      continue;
    }
    if (b === 0x28) { // ( literal string
      i += 1;
      let depth = 1;
      const bytes = [];
      while (i < len && depth > 0) {
        const ch = buffer[i];
        if (ch === 0x5c && i + 1 < len) {
          i += 1;
          const esc = buffer[i];
          if (esc === 0x6e) bytes.push(0x0a);
          else if (esc === 0x72) bytes.push(0x0d);
          else if (esc === 0x74) bytes.push(0x09);
          else if (esc === 0x62) bytes.push(0x08);
          else if (esc === 0x66) bytes.push(0x0c);
          else if (esc >= 0x30 && esc <= 0x37) {
            let oct = String.fromCharCode(esc);
            for (let k = 0; k < 2 && i + 1 < len && buffer[i + 1] >= 0x30 && buffer[i + 1] <= 0x37; k += 1) {
              i += 1;
              oct += String.fromCharCode(buffer[i]);
            }
            bytes.push(parseInt(oct, 8) & 0xff);
          } else if (esc !== 0x0a && esc !== 0x0d) {
            bytes.push(esc);
          }
          i += 1;
        } else if (ch === 0x28) {
          depth += 1;
          bytes.push(ch);
          i += 1;
        } else if (ch === 0x29) {
          depth -= 1;
          if (depth > 0) bytes.push(ch);
          i += 1;
        } else {
          bytes.push(ch);
          i += 1;
        }
      }
      tokens.push({ type: 'str', bytes: Buffer.from(bytes) });
      continue;
    }
    if (b === 0x3c && buffer[i + 1] === 0x3c) {
      i += 2;
      tokens.push({ type: 'op', value: '<<' });
      continue;
    }
    if (b === 0x3e && buffer[i + 1] === 0x3e) {
      i += 2;
      tokens.push({ type: 'op', value: '>>' });
      continue;
    }
    if (b === 0x3c) { // < hex string
      i += 1;
      let hex = '';
      while (i < len && buffer[i] !== 0x3e) {
        const ch = String.fromCharCode(buffer[i]);
        if (/[0-9a-fA-F]/.test(ch)) hex += ch;
        i += 1;
      }
      if (i < len && buffer[i] === 0x3e) i += 1;
      if (hex.length % 2 === 1) hex += '0';
      tokens.push({ type: 'str', bytes: Buffer.from(hex, 'hex') });
      continue;
    }
    if (b === 0x5b) { // [ array
      i += 1;
      tokens.push({ type: 'arr_start' });
      continue;
    }
    if (b === 0x5d) { // ] array end
      i += 1;
      tokens.push({ type: 'arr_end' });
      continue;
    }
    if (b === 0x2f) { // /Name
      const start = i;
      i += 1;
      while (i < len && !isDelim(buffer[i])) i += 1;
      tokens.push({ type: 'name', value: buffer.toString('latin1', start, i) });
      continue;
    }
    const start = i;
    while (i < len && !isDelim(buffer[i])) i += 1;
    if (start === i) {
      i += 1;
      continue;
    }
    const word = buffer.toString('latin1', start, i);
    const num = Number(word);
    if (/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(word) && !Number.isNaN(num)) {
      tokens.push({ type: 'num', value: num });
    } else {
      tokens.push({ type: 'op', value: word });
    }
  }
  return tokens;
}

function collectPageStreamBuffers(pdf, page) {
  const contents = page.node.Contents();
  if (!contents) return [];
  const resolved = pdf.context.lookup(contents);
  if (!resolved) return [];
  if (resolved instanceof PDFArray) {
    const bufs = [];
    for (let i = 0; i < resolved.size(); i += 1) {
      const buf = decodeStreamBytes(pdf, resolved.get(i));
      if (buf.length > 0) bufs.push(buf);
    }
    return bufs;
  }
  if (resolved instanceof PDFRawStream || typeof resolved.contents !== 'undefined') {
    const buf = decodeStreamBytes(pdf, resolved);
    return buf.length > 0 ? [buf] : [];
  }
  return [];
}

function groupFragmentsIntoClusters(fragments) {
  if (fragments.length === 0) return [];
  const sorted = [...fragments].sort((a, b) => b.y - a.y || a.x - b.x);
  const rows = [];
  for (const frag of sorted) {
    let targetRow = rows.find((r) => Math.abs(r.y - frag.y) <= 3.5);
    if (!targetRow) {
      targetRow = { y: frag.y, items: [] };
      rows.push(targetRow);
    }
    targetRow.items.push(frag);
  }
  const clusters = [];
  for (const row of rows) {
    row.items.sort((a, b) => a.x - b.x);
    let current = null;
    for (const item of row.items) {
      const gap = current ? item.x - current.endX : 0;
      if (!current || gap > 38) {
        if (current) clusters.push(current);
        current = {
          text: item.text,
          x: item.x,
          endX: item.endX,
          y: item.y,
          fontSize: item.fontSize,
        };
      } else {
        current.text += gap > Math.max(1.5, item.fontSize * 0.18) && !current.text.endsWith(' ') && !item.text.startsWith(' ') ? ` ${item.text}` : item.text;
        current.endX = Math.max(current.endX, item.endX);
      }
    }
    if (current) clusters.push(current);
  }
  return clusters.map((cluster) => ({
    ...cluster,
    centerX: (cluster.x + cluster.endX) / 2,
    normalized: normalizePdfText(cluster.text),
  }));
}

function extractPageLayoutInfo(pdf, page) {
  const { width: pageWidth, height: pageHeight } = page.getSize();
  const fragments = [];
  const state = { hasImage: false, minDrawnY: null, maxDrawnY: null };

  const recordDrawnPoint = (x, y) => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) return;
    if (y < -5 || y > pageHeight + 5) return;
    state.minDrawnY = state.minDrawnY === null ? y : Math.min(state.minDrawnY, y);
    state.maxDrawnY = state.maxDrawnY === null ? y : Math.max(state.maxDrawnY, y);
  };

  const processStream = (streamBufs, resources, initialCtm, depth = 0) => {
    if (depth > 4) return;
    const fontDecoders = buildFontDecoders(pdf, resources);
    const resolvedRes = resources ? pdf.context.lookup(resources) : null;
    const xObjectDict = resolvedRes?.get?.(PDFName.of('XObject'))
      ? pdf.context.lookup(resolvedRes.get(PDFName.of('XObject')))
      : null;

    let ctm = [...initialCtm];
    const ctmStack = [];
    let inText = false;
    let tm = [...IDENTITY_MATRIX];
    let tlm = [...IDENTITY_MATRIX];
    let leading = 0;
    let fontKey = '';
    let fontSize = 10;

    const recordTextBytes = (bytes) => {
      const decoder = fontDecoders.get(fontKey);
      const decoded = decodePdfBytesWithFont(bytes, decoder);
      if (!decoded || !decoded.trim()) return;
      const combined = multiplyMatrices(tm, ctm);
      const scale = Math.hypot(combined[0], combined[1]) || 1;
      const effFontSize = Math.abs(fontSize * scale) || 10;
      const estWidth = Math.max(effFontSize * 0.48 * decoded.length, 4);
      const x = combined[4];
      const y = combined[5];
      fragments.push({ text: decoded, x, endX: x + estWidth, y, fontSize: effFontSize });
      recordDrawnPoint(x, y);
      tm = multiplyMatrices([1, 0, 0, 1, (estWidth / scale), 0], tm);
    };

    for (const buf of streamBufs) {
      const tokens = tokenizePdfStream(buf);
      const stack = [];
      for (const tok of tokens) {
        if (tok.type === 'arr_start') {
          stack.push({ type: 'arr', items: [] });
          continue;
        }
        if (tok.type === 'arr_end') {
          const collected = [];
          while (stack.length > 0) {
            const top = stack.pop();
            if (top.type === 'arr') {
              stack.push({ type: 'array', value: collected.reverse() });
              break;
            }
            collected.push(top);
          }
          continue;
        }
        if (tok.type !== 'op') {
          stack.push(tok);
          continue;
        }

        const op = tok.value;
        if (op === 'q') {
          ctmStack.push([...ctm]);
        } else if (op === 'Q') {
          if (ctmStack.length > 0) ctm = ctmStack.pop();
        } else if (op === 'cm' && stack.length >= 6) {
          const vals = stack.slice(-6).map((t) => t.value);
          ctm = multiplyMatrices(vals, ctm);
        } else if (op === 'BT') {
          inText = true;
          tm = [...IDENTITY_MATRIX];
          tlm = [...IDENTITY_MATRIX];
        } else if (op === 'ET') {
          inText = false;
        } else if (op === 'Tf' && stack.length >= 2) {
          const sz = stack[stack.length - 1];
          const fn = stack[stack.length - 2];
          if (fn?.type === 'name') fontKey = fn.value;
          if (sz?.type === 'num') fontSize = sz.value;
        } else if (op === 'Tm' && stack.length >= 6) {
          const vals = stack.slice(-6).map((t) => t.value);
          tm = [...vals];
          tlm = [...vals];
        } else if ((op === 'Td' || op === 'TD') && stack.length >= 2) {
          const tx = stack[stack.length - 2].value || 0;
          const ty = stack[stack.length - 1].value || 0;
          if (op === 'TD') leading = -ty;
          tlm = multiplyMatrices([1, 0, 0, 1, tx, ty], tlm);
          tm = [...tlm];
        } else if (op === 'TL' && stack.length >= 1) {
          leading = stack[stack.length - 1].value || 0;
        } else if (op === 'T*') {
          tlm = multiplyMatrices([1, 0, 0, 1, 0, -leading], tlm);
          tm = [...tlm];
        } else if (inText && op === 'Tj' && stack.length >= 1) {
          const s = stack[stack.length - 1];
          if (s?.type === 'str') recordTextBytes(s.bytes);
        } else if (inText && op === "'" && stack.length >= 1) {
          tlm = multiplyMatrices([1, 0, 0, 1, 0, -leading], tlm);
          tm = [...tlm];
          const s = stack[stack.length - 1];
          if (s?.type === 'str') recordTextBytes(s.bytes);
        } else if (inText && op === '"' && stack.length >= 3) {
          tlm = multiplyMatrices([1, 0, 0, 1, 0, -leading], tlm);
          tm = [...tlm];
          const s = stack[stack.length - 1];
          if (s?.type === 'str') recordTextBytes(s.bytes);
        } else if (inText && op === 'TJ' && stack.length >= 1) {
          const arr = stack[stack.length - 1];
          if (arr?.type === 'array') {
            for (const item of arr.value) {
              if (item.type === 'str') {
                recordTextBytes(item.bytes);
              } else if (item.type === 'num') {
                const shift = (-item.value / 1000) * fontSize;
                tm = multiplyMatrices([1, 0, 0, 1, shift, 0], tm);
              }
            }
          }
        } else if ((op === 'm' || op === 'l') && stack.length >= 2) {
          const [px, py] = applyMatrix(ctm, stack[stack.length - 2].value || 0, stack[stack.length - 1].value || 0);
          recordDrawnPoint(px, py);
        } else if (op === 're' && stack.length >= 4) {
          const rx = stack[stack.length - 4].value || 0;
          const ry = stack[stack.length - 3].value || 0;
          const rw = stack[stack.length - 2].value || 0;
          const rh = stack[stack.length - 1].value || 0;
          const [, py1] = applyMatrix(ctm, rx, ry);
          const [, py2] = applyMatrix(ctm, rx + rw, ry + rh);
          if (Math.abs(py2 - py1) < pageHeight - 10) {
            recordDrawnPoint(rx, Math.min(py1, py2));
            recordDrawnPoint(rx, Math.max(py1, py2));
          }
        } else if (op === 'Do' && stack.length >= 1 && xObjectDict) {
          const xName = stack[stack.length - 1];
          if (xName?.type === 'name') {
            const xRef = xObjectDict.get(PDFName.of(xName.value.slice(1)));
            const xObj = xRef ? pdf.context.lookup(xRef) : null;
            const subtype = xObj?.dict?.get?.(PDFName.of('Subtype'));
            if (subtype instanceof PDFName && subtype.asString() === '/Image') {
              state.hasImage = true;
            } else if (subtype instanceof PDFName && subtype.asString() === '/Form') {
              const formMatrixObj = xObj.dict.get(PDFName.of('Matrix'));
              let formMat = [...IDENTITY_MATRIX];
              if (formMatrixObj instanceof PDFArray && formMatrixObj.size() === 6) {
                formMat = [0, 1, 2, 3, 4, 5].map((idx) => formMatrixObj.get(idx).asNumber());
              }
              const childBytes = decodeStreamBytes(pdf, xObj);
              if (childBytes.length > 0) {
                processStream([childBytes], xObj.dict.get(PDFName.of('Resources')) || resources, multiplyMatrices(formMat, ctm), depth + 1);
              }
            }
          }
        }
        stack.length = 0;
      }
    }
  };

  const bufs = collectPageStreamBuffers(pdf, page);
  if (bufs.length > 0) {
    processStream(bufs, page.node.Resources(), IDENTITY_MATRIX, 0);
  }
  const clusters = groupFragmentsIntoClusters(fragments);
  return {
    pageWidth,
    pageHeight,
    clusters,
    hasImage: state.hasImage,
    hasVisibleContent: clusters.length > 0 || state.hasImage,
    minDrawnY: state.minDrawnY,
    maxDrawnY: state.maxDrawnY,
  };
}

function sanitizePdfPages(pdf) {
  let pageCount = pdf.getPageCount();
  while (pageCount > 1) {
    const lastIndex = pageCount - 1;
    const lastPage = pdf.getPage(lastIndex);
    const layout = extractPageLayoutInfo(pdf, lastPage);
    if (!layout.hasVisibleContent) {
      pdf.removePage(lastIndex);
      pageCount = pdf.getPageCount();
    } else {
      break;
    }
  }
}

function pickSignatureAnchors(layout) {
  const { clusters, pageWidth } = layout;
  const coordCandidates = clusters.filter(
    (c) => c.normalized.includes('coordinador') && !c.normalized.startsWith('14.')
  );
  const directorCandidates = clusters.filter(
    (c) => c.normalized.includes('visto bueno')
  );
  const coordAnchor =
    coordCandidates.find((c) => c.normalized.includes('extension')) ||
    coordCandidates.sort((a, b) => a.y - b.y)[0] ||
    null;
  const directorAnchor =
    directorCandidates.find((c) => c.normalized.includes('unidad academica')) ||
    directorCandidates.sort((a, b) => a.y - b.y)[0] ||
    null;

  const nombreCandidates = clusters.filter((c) => c.normalized.startsWith('nombre'));
  const findNameAbove = (anchor) => {
    if (!anchor) return null;
    const isRightSide = anchor.centerX >= pageWidth / 2;
    return nombreCandidates.find(
      (n) =>
        n.y > anchor.y + 3 &&
        n.y <= anchor.y + 28 &&
        (isRightSide ? n.centerX >= pageWidth * 0.4 : n.centerX < pageWidth * 0.6)
    ) || null;
  };

  return {
    coordAnchor,
    directorAnchor,
    coordName: findNameAbove(coordAnchor),
    directorName: findNameAbove(directorAnchor),
  };
}

function resolveSignatureTarget(pdf, formNum) {
  const pageCount = pdf.getPageCount();
  const cfg = SIGNATURE_FORM_LAYOUTS[formNum];
  let bestPageIndex = pageCount - 1;
  let bestLayout = null;
  let bestAnchors = null;

  for (let idx = pageCount - 1; idx >= 0; idx -= 1) {
    const layout = extractPageLayoutInfo(pdf, pdf.getPage(idx));
    const anchors = pickSignatureAnchors(layout);
    if (anchors.coordAnchor || anchors.directorAnchor) {
      bestPageIndex = idx;
      bestLayout = layout;
      bestAnchors = anchors;
      break;
    }
    if (!bestLayout && layout.hasVisibleContent) {
      bestPageIndex = idx;
      bestLayout = layout;
      bestAnchors = anchors;
    }
  }

  if (!bestLayout) {
    bestLayout = extractPageLayoutInfo(pdf, pdf.getPage(bestPageIndex));
    bestAnchors = pickSignatureAnchors(bestLayout);
  }

  const pageWidth = bestLayout.pageWidth || 612;
  const computeX = (anchor, fallbackX, width) => {
    if (!anchor) return fallbackX;
    const centered = anchor.centerX - width / 2;
    return Math.max(18, Math.min(pageWidth - width - 18, centered));
  };
  const computeY = (anchor, nameAnchor) => {
    if (cfg.hasNameRow && nameAnchor) {
      return nameAnchor.y + cfg.offsetAboveName;
    }
    if (anchor) {
      return anchor.y + cfg.offsetAboveTitle;
    }
    if (bestLayout.minDrawnY !== null) {
      return bestLayout.minDrawnY + cfg.fallbackBottomOffset;
    }
    return cfg.fallbackY;
  };

  return {
    page: pdf.getPage(bestPageIndex),
    solicitanteBox: {
      x: computeX(bestAnchors.coordAnchor, cfg.fallbackLeftX, cfg.maxWidth),
      y: computeY(bestAnchors.coordAnchor || bestAnchors.directorAnchor, bestAnchors.coordName || bestAnchors.directorName),
      maxWidth: bestAnchors.coordAnchor ? cfg.maxWidth : cfg.fallbackLeftWidth,
      maxHeight: cfg.maxHeight,
    },
    directorBox: {
      x: computeX(bestAnchors.directorAnchor, cfg.fallbackRightX, cfg.maxWidth),
      y: computeY(bestAnchors.directorAnchor || bestAnchors.coordAnchor, bestAnchors.directorName || bestAnchors.coordName),
      maxWidth: bestAnchors.directorAnchor ? cfg.maxWidth : cfg.fallbackRightWidth,
      maxHeight: cfg.maxHeight,
    },
  };
}

async function embedSignaturesInPdf(pdfBytes, orderedSignatures = [], requestId, formNumber) {
  const formNum = parseInt(formNumber, 10);
  const pdf = await PDFDocument.load(pdfBytes);
  sanitizePdfPages(pdf);
  const pageCount = pdf.getPageCount();
  if (pageCount === 0) return pdf.save();

  const solicitanteSig = orderedSignatures.find((item) => item?.tipo_firma === 'solicitante');
  const directorSig = orderedSignatures.find((item) => item?.tipo_firma === 'director_escuela');

  if (SIGNATURE_FORM_LAYOUTS[formNum]) {
    const { page, solicitanteBox, directorBox } = resolveSignatureTarget(pdf, formNum);
    if (solicitanteSig) {
      await drawSignatureBox({ page, signature: solicitanteSig, ...solicitanteBox });
    }
    if (directorSig) {
      await drawSignatureBox({ page, signature: directorSig, ...directorBox });
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
  getSignatureBuffer,
};
