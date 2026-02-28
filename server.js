/**
 * server.js — QIR Merge Server
 * Self-contained. No external local files needed.
 *
 * Endpoints:
 *   GET  /          health check
 *   POST /merge     merge QIR + certs, return final PDF
 */

const express = require('express');
const fetch   = require('node-fetch');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');

const app  = express();
const PORT = process.env.PORT || 3000;

// CORS — allow requests from local HTML files and any origin
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Parse JSON bodies up to 50MB (base64 PDFs can be large)
app.use(express.json({ limit: '50mb' }));

// ── Health check ──────────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({ status: 'ok', service: 'QIR Merge Server', version: '1.0.0' });
});

// ── Helpers ───────────────────────────────────────────────────

/** Load PDF from { type: 'url'|'base64', value } → Buffer */
async function loadPDFFromSource(src) {
  if (src.type === 'url') {
    console.log(`  Fetching URL: ${src.value.substring(0, 70)}...`);
    const res = await fetch(src.value, { timeout: 30000 });
    if (!res.ok) throw new Error(`HTTP ${res.status} fetching PDF`);
    return Buffer.from(await res.arrayBuffer());
  } else {
    return Buffer.from(src.value, 'base64');
  }
}

/** Build the Index page as a single-page A4 landscape PDF buffer */
async function buildIndexPage({ reportNo, partName, date, inspectionPage, certEntries }) {
  const doc  = await PDFDocument.create();
  const W = 841.89, H = 595.28;   // A4 landscape in points
  const page = doc.addPage([W, H]);

  const fontBold   = await doc.embedFont(StandardFonts.HelveticaBold);
  const fontNormal = await doc.embedFont(StandardFonts.Helvetica);

  const DARK  = rgb(0.10, 0.10, 0.18);
  const RED   = rgb(0.78, 0.29, 0.17);
  const GRAY  = rgb(0.50, 0.50, 0.50);
  const LGRAY = rgb(0.93, 0.93, 0.93);
  const OFF   = rgb(0.96, 0.96, 0.97);

  // Top red bar
  page.drawRectangle({ x:0, y:H-6, width:W, height:6, color:RED });

  // Title
  page.drawText('Quality Inspection Report', {
    x:40, y:H-44, size:22, font:fontBold, color:DARK,
  });
  page.drawText(`${reportNo}   ·   ${partName}   ·   ${date}`, {
    x:40, y:H-64, size:10, font:fontNormal, color:GRAY,
  });

  // INDEX heading
  page.drawText('INDEX', { x:40, y:H-100, size:11, font:fontBold, color:RED });
  page.drawLine({
    start:{x:40, y:H-108}, end:{x:W-40, y:H-108},
    thickness:0.5, color:LGRAY,
  });

  const ENTRY_H = 32;
  const COL_L   = 40;
  const COL_R   = W - 80;
  let   rowY    = H - 118;

  function drawRow(label, pageNum, isSub=false, isHeader=false) {
    page.drawRectangle({
      x:36, y:rowY-ENTRY_H+6, width:W-72, height:ENTRY_H-2,
      color: isHeader ? OFF : rgb(1,1,1),
    });
    const indent   = isSub ? 24 : 0;
    const fontSize = isHeader ? 10 : 9.5;
    const font     = isHeader ? fontBold : fontNormal;
    const color    = isHeader ? DARK : isSub ? GRAY : DARK;

    page.drawText(String(label), {
      x: COL_L + indent, y: rowY - 10,
      size: fontSize, font, color,
    });

    const pgStr = String(pageNum);
    const pgW   = fontBold.widthOfTextAtSize(pgStr, fontSize);
    page.drawText(pgStr, {
      x: COL_R - pgW, y: rowY - 10,
      size: fontSize, font: fontBold,
      color: isHeader ? DARK : RED,
    });

    // Dotted leader line
    if (!isHeader) {
      const labelW = font.widthOfTextAtSize(String(label), fontSize);
      const dStart = COL_L + indent + labelW + 8;
      const dEnd   = COL_R - pgW - 8;
      for (let dx = dStart; dx < dEnd; dx += 4) {
        page.drawCircle({ x:dx, y:rowY-7, size:0.7, color:LGRAY });
      }
    }
    rowY -= ENTRY_H;
  }

  // Fixed sections
  drawRow('Report Header & Part Information', 1,               false, true);
  drawRow('Index',                            2,               false, true);
  drawRow('Part Drawing',                     3,               false, true);
  drawRow('Inspection',                       inspectionPage,  false, true);
  drawRow('Dimensional Inspection',           inspectionPage,  true,  false);
  drawRow('Visual Inspection',                '—',             true,  false);

  // Divider before certs
  rowY -= 6;
  page.drawLine({
    start:{x:40, y:rowY+10}, end:{x:W-40, y:rowY+10},
    thickness:0.5, color:LGRAY,
  });
  rowY -= 10;

  drawRow('Tests & Certificates', certEntries[0]?.startPage || '—', false, true);
  for (const c of certEntries) {
    drawRow(c.label, c.startPage, true, false);
  }

  // Bottom red bar + page number
  page.drawRectangle({ x:0, y:0, width:W, height:4, color:RED });
  const p2W = fontNormal.widthOfTextAtSize('2', 8);
  page.drawText('2', { x:W/2-p2W/2, y:10, size:8, font:fontNormal, color:GRAY });

  return Buffer.from(await doc.save());
}

/** Stamp page numbers at the bottom of every page in a PDF.
 *  Font size is normalized to ~9pt relative to A4 width (595pt landscape = 841pt).
 *  This keeps numbers visually consistent regardless of the cert PDF's page size.
 */
async function stampPageNumbers(pdfBytes, startPageNum) {
  const pdf      = await PDFDocument.load(pdfBytes, { ignoreEncryption:true });
  const font     = await pdf.embedFont(StandardFonts.Helvetica);

  pdf.getPages().forEach((page, i) => {
    const { width, height } = page.getSize();

    // Normalize: target ~9pt on A4 landscape width (841pt). Scale to actual page width.
    const BASE_WIDTH  = 841;
    const TARGET_SIZE = 9;
    const fontSize    = Math.round((width / BASE_WIDTH) * TARGET_SIZE * 10) / 10;
    const barH        = fontSize * 2.4;

    // Subtle light bar at bottom
    page.drawRectangle({
      x:0, y:0, width, height:barH,
      color:rgb(0.96,0.96,0.96),
      opacity:0.85,
    });

    // Page number — centred, dark gray (not red)
    const pgStr = String(startPageNum + i);
    const pgW   = font.widthOfTextAtSize(pgStr, fontSize);
    page.drawText(pgStr, {
      x: width/2 - pgW/2,
      y: barH * 0.28,
      size: fontSize,
      font,
      color: rgb(0.25, 0.25, 0.25),
    });
  });

  return Buffer.from(await pdf.save());
}



// ── POST /merge ───────────────────────────────────────────────
/**
 * Body:
 * {
 *   reportNo, partName, date,
 *   qirSource:    { type: 'url'|'base64', value: '...' }
 *   certificates: [{ label, type: 'url'|'base64', value }, ...]
 * }
 * Returns: PDF binary (application/pdf)
 */
app.post('/merge', async (req, res) => {
  const { reportNo='QIR', partName='', date='', qirSource, certificates=[] } = req.body;

  console.log(`\n── /merge: ${reportNo} — ${certificates.length} cert(s)`);

  try {
    // 1. Load QIR PDF
    console.log('[1] Loading QIR...');
    const qirBytes    = await loadPDFFromSource(qirSource);
    const qirPdf      = await PDFDocument.load(qirBytes, { ignoreEncryption:true });
    const qirPageCount = qirPdf.getPageCount();
    console.log(`    ${qirPageCount} pages`);

    // 2. Load cert PDFs
    console.log('[2] Loading certificates...');
    const certData = [];
    for (const cert of certificates) {
      if (!cert.value?.trim()) continue;
      try {
        const bytes  = await loadPDFFromSource(cert);
        const pdf    = await PDFDocument.load(bytes, { ignoreEncryption:true });
        certData.push({
          label:     cert.label || 'Certificate',
          bytes,
          pageCount: pdf.getPageCount(),
        });
        console.log(`    "${cert.label}" — ${pdf.getPageCount()} page(s)`);
      } catch(e) {
        console.warn(`    Skipped "${cert.label}": ${e.message}`);
      }
    }

    // 3. Calculate page numbers
    // After inserting index as page 2, all original QIR pages shift +1
    // So certs start at: qirPageCount + 1 (index) + 1 (offset) = qirPageCount + 2
    const INSPECTION_PAGE = 4;
    const certStartPage   = qirPageCount + 2;
    let   runningPage     = certStartPage;

    const certEntries = certData.map(c => {
      const entry = {
        label:     c.label,
        startPage: runningPage,
        pageCount: c.pageCount,
        bytes:     c.bytes,
      };
      runningPage += c.pageCount;
      return entry;
    });

    console.log(`[3] Pages: inspection=${INSPECTION_PAGE}, certs start=${certStartPage}`);

    // 4. Build index page
    console.log('[4] Building index page...');
    const indexBytes = await buildIndexPage({
      reportNo, partName, date,
      inspectionPage: INSPECTION_PAGE,
      certEntries,
    });

    // 5. Merge everything
    console.log('[5] Merging...');
    const merged   = await PDFDocument.create();
    const qirPages = await merged.copyPages(qirPdf, qirPdf.getPageIndices());

    // QIR page 1 (header)
    merged.addPage(qirPages[0]);

    // Index as page 2
    const idxPdf    = await PDFDocument.load(indexBytes);
    const [idxPage] = await merged.copyPages(idxPdf, [0]);
    merged.addPage(idxPage);

    // Rest of QIR (part drawing, inspection, conclusion)
    for (let i = 1; i < qirPages.length; i++) {
      merged.addPage(qirPages[i]);
    }

    // Certificates with page numbers
    for (const cert of certEntries) {
      let bytes = await stampPageNumbers(cert.bytes, cert.startPage);
      const cp  = await PDFDocument.load(bytes, { ignoreEncryption:true });
      const pgs = await merged.copyPages(cp, cp.getPageIndices());
      pgs.forEach(p => merged.addPage(p));
      console.log(`    Added "${cert.label}" p.${cert.startPage}`);
    }

    // 6. Send PDF
    const finalBytes = await merged.save();
    const filename   = `QIR-${reportNo}-${date}.pdf`.replace(/[^a-zA-Z0-9\-_.]/g, '_');
    const totalPages = merged.getPageCount();
    const sizeKB     = (finalBytes.length / 1024).toFixed(0);

    console.log(`✓ Done — ${totalPages} pages, ${sizeKB} KB → ${filename}\n`);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(Buffer.from(finalBytes));

  } catch(err) {
    console.error('✗ Merge error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── Start ────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\nQIR Merge Server on port ${PORT}`);
  console.log(`Health:  GET  /`);
  console.log(`Merge:   POST /merge\n`);
});
