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

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

app.use(express.json({ limit: '50mb' }));

app.get('/', (req, res) => {
  res.json({ status: 'ok', service: 'QIR Merge Server', version: '1.0.0' });
});

async function loadPDFFromSource(src) {
  if (src.type === 'url') {
    console.log(`  Fetching: ${src.value.substring(0, 70)}...`);
    const res = await fetch(src.value, { timeout: 30000 });
    if (!res.ok) throw new Error(`HTTP ${res.status} fetching PDF`);
    return Buffer.from(await res.arrayBuffer());
  }
  return Buffer.from(src.value, 'base64');
}

// ── INDEX PAGE — clean black/grey table style ─────────────────
async function buildIndexPage({ reportNo, partName, date, inspectionPage, certEntries }) {
  const doc  = await PDFDocument.create();
  const W = 841.89, H = 595.28;
  const page = doc.addPage([W, H]);

  const fontBold   = await doc.embedFont(StandardFonts.HelveticaBold);
  const fontNormal = await doc.embedFont(StandardFonts.Helvetica);

  const BLACK  = rgb(0.10, 0.10, 0.10);
  const DGRAY  = rgb(0.30, 0.30, 0.30);
  const MGRAY  = rgb(0.55, 0.55, 0.55);
  const LGRAY  = rgb(0.82, 0.82, 0.82);
  const OFFWHT = rgb(0.95, 0.95, 0.95);
  const HDRBG  = rgb(0.88, 0.88, 0.88);
  const WHITE  = rgb(1.00, 1.00, 1.00);

  page.drawRectangle({ x:0, y:0, width:W, height:H, color:WHITE });

  // Title
  page.drawText('Quality Inspection Report', {
    x:40, y:H-52, size:20, font:fontBold, color:BLACK,
  });
  page.drawText(`${reportNo}   ·   ${partName}   ·   ${date}`, {
    x:40, y:H-70, size:9, font:fontNormal, color:MGRAY,
  });
  page.drawLine({
    start:{x:40,y:H-80}, end:{x:W-40,y:H-80},
    thickness:0.5, color:LGRAY,
  });
  page.drawText('TABLE OF CONTENTS', {
    x:40, y:H-97, size:8, font:fontBold, color:MGRAY,
  });

  // Table layout
  const TBL_X  = 40;
  const TBL_W  = W - 80;
  const PAD    = 10;
  const ROW_H  = 24;
  let   rowY   = H - 112;
  const tableTopY = rowY;

  function drawRow(label, pageNum, isSub=false, isHeader=false) {
    const bg = isHeader ? HDRBG : (isSub ? WHITE : OFFWHT);
    page.drawRectangle({ x:TBL_X, y:rowY-ROW_H, width:TBL_W, height:ROW_H, color:bg });
    page.drawLine({ start:{x:TBL_X,y:rowY-ROW_H}, end:{x:TBL_X+TBL_W,y:rowY-ROW_H}, thickness:0.3, color:LGRAY });

    const indent   = isSub ? 18 : 0;
    const fontSize = isHeader ? 8.5 : 8;
    const font     = isHeader ? fontBold : fontNormal;
    const color    = isHeader ? BLACK : isSub ? MGRAY : DGRAY;
    const textY    = rowY - ROW_H + (ROW_H - fontSize) / 2 + 1;

    page.drawText(String(label), { x:TBL_X+PAD+indent, y:textY, size:fontSize, font, color });

    const pgStr = String(pageNum);
    const pgW   = (isHeader ? fontBold : fontNormal).widthOfTextAtSize(pgStr, fontSize);
    page.drawText(pgStr, { x:TBL_X+TBL_W-PAD-pgW, y:textY, size:fontSize, font: isHeader ? fontBold : fontNormal, color: isHeader ? BLACK : DGRAY });

    rowY -= ROW_H;
  }

  // Top border
  page.drawLine({ start:{x:TBL_X,y:tableTopY}, end:{x:TBL_X+TBL_W,y:tableTopY}, thickness:0.6, color:DGRAY });

  drawRow('Section', 'Page', false, true);
  drawRow('Report Header & Part Information', 1);
  drawRow('Index', 2);
  drawRow('Part Drawing', 3);
  drawRow('Inspection', inspectionPage);
  drawRow('Dimensional Inspection', inspectionPage, true);
  drawRow('Visual Inspection', '—', true);
  drawRow('Tests & Certificates', certEntries[0]?.startPage || '—');
  for (const c of certEntries) drawRow(c.label, c.startPage, true);

  // Bottom + side borders
  page.drawLine({ start:{x:TBL_X,y:rowY}, end:{x:TBL_X+TBL_W,y:rowY}, thickness:0.6, color:DGRAY });
  page.drawLine({ start:{x:TBL_X,y:tableTopY}, end:{x:TBL_X,y:rowY}, thickness:0.6, color:DGRAY });
  page.drawLine({ start:{x:TBL_X+TBL_W,y:tableTopY}, end:{x:TBL_X+TBL_W,y:rowY}, thickness:0.6, color:DGRAY });

  // Page number
  const p2W = fontNormal.widthOfTextAtSize('2', 8);
  page.drawText('2', { x:W/2-p2W/2, y:20, size:8, font:fontNormal, color:MGRAY });

  return Buffer.from(await doc.save());
}

// ── HEADING STAMP — text only, no bar, no red line ───────────
async function stampHeading(pdfBytes, label) {
  if (!label?.trim()) return pdfBytes;
  try {
    const pdf  = await PDFDocument.load(pdfBytes, { ignoreEncryption:true });
    const page = pdf.getPages()[0];
    const font = await pdf.embedFont(StandardFonts.HelveticaBold);
    const { width, height } = page.getSize();

    // Normalize font size to appear same across different page sizes
    const fontSize = Math.round((width / 841) * 11 * 10) / 10;

    // Semi-transparent white strip so text is readable over any background
    page.drawRectangle({
      x:0, y:height - fontSize*2.8,
      width, height: fontSize*2.8,
      color: rgb(1,1,1), opacity: 0.75,
    });

    // Label text — dark, centred, no line, no colour
    const textW = font.widthOfTextAtSize(label, fontSize);
    page.drawText(label, {
      x: (width - textW) / 2,
      y: height - fontSize*2.0,
      size: fontSize, font,
      color: rgb(0.10, 0.10, 0.10),
    });

    return Buffer.from(await pdf.save());
  } catch(e) { return pdfBytes; }
}

// ── PAGE NUMBER STAMP ─────────────────────────────────────────
// Normalized font size so numbers appear same physical size on any page size
async function stampPageNumbers(pdfBytes, startPageNum) {
  const pdf  = await PDFDocument.load(pdfBytes, { ignoreEncryption:true });
  const font = await pdf.embedFont(StandardFonts.Helvetica);

  pdf.getPages().forEach((page, i) => {
    const { width } = page.getSize();
    const fontSize = Math.round((width / 841) * 8 * 10) / 10;
    const barH     = fontSize * 2.6;

    page.drawRectangle({ x:0, y:0, width, height:barH, color:rgb(0.96,0.96,0.96), opacity:0.9 });

    const pgStr = String(startPageNum + i);
    const pgW   = font.widthOfTextAtSize(pgStr, fontSize);
    page.drawText(pgStr, { x:width/2-pgW/2, y:barH*0.25, size:fontSize, font, color:rgb(0.20,0.20,0.20) });
  });

  return Buffer.from(await pdf.save());
}

// ── Stamp page numbers directly onto an already-loaded PDFDocument ──
async function stampPageNumbersOnDoc(pdfDoc, pageIndexOffset, startPageNum) {
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const pages = pdfDoc.getPages();

  for (let i = pageIndexOffset; i < pages.length; i++) {
    const page = pages[i];
    const { width } = page.getSize();
    const fontSize = Math.round((width / 841) * 8 * 10) / 10;
    const barH = fontSize * 2.6;
    const finalPageNum = startPageNum + (i - pageIndexOffset);

    page.drawRectangle({ x:0, y:0, width, height:barH, color:rgb(0.96,0.96,0.96), opacity:0.9 });
    const pgStr = String(finalPageNum);
    const pgW   = font.widthOfTextAtSize(pgStr, fontSize);
    page.drawText(pgStr, { x:width/2-pgW/2, y:barH*0.25, size:fontSize, font, color:rgb(0.20,0.20,0.20) });
  }
}

// ── POST /merge ───────────────────────────────────────────────
app.post('/merge', async (req, res) => {
  const { reportNo='QIR', partName='', date='', qirSource, certificates=[] } = req.body;
  console.log(`\n── /merge: ${reportNo} — ${certificates.length} cert(s)`);

  try {
    // 1. Load QIR
    console.log('[1] Loading QIR...');
    const qirBytes     = await loadPDFFromSource(qirSource);
    const qirPdf       = await PDFDocument.load(qirBytes, { ignoreEncryption:true });
    const qirPageCount = qirPdf.getPageCount();
    console.log(`    ${qirPageCount} pages`);

    // 2. Load certs
    console.log('[2] Loading certificates...');
    const certData = [];
    for (const cert of certificates) {
      if (!cert.value?.trim()) continue;
      try {
        const bytes = await loadPDFFromSource(cert);
        const pdf   = await PDFDocument.load(bytes, { ignoreEncryption:true });
        certData.push({ label: cert.label || 'Certificate', bytes, pageCount: pdf.getPageCount() });
        console.log(`    "${cert.label}" — ${pdf.getPageCount()} page(s)`);
      } catch(e) { console.warn(`    Skipped "${cert.label}": ${e.message}`); }
    }

    // 3. Calculate page numbers
    // Final structure:
    //   p.1            QIR page 1 (header)
    //   p.2            Index (inserted)
    //   p.3            QIR page 2 (part drawing)
    //   p.4 .. p.N+1   QIR pages 3..N (inspection etc.)
    //   p.N+2 onwards  Certificates
    const INSPECTION_PAGE = 4;
    const certStartPage   = qirPageCount + 2;
    let   runningPage     = certStartPage;
    const certEntries = certData.map(c => {
      const entry = { ...c, startPage: runningPage };
      runningPage += c.pageCount;
      return entry;
    });
    console.log(`[3] inspection=p.${INSPECTION_PAGE}, certs=p.${certStartPage}`);

    // 4. Build index
    console.log('[4] Building index...');
    const indexBytes = await buildIndexPage({ reportNo, partName, date, inspectionPage: INSPECTION_PAGE, certEntries });

    // 5. Stamp QIR page numbers
    // QIR p.1 → final p.1,  QIR p.2 → final p.3,  QIR p.3 → final p.4, etc.
    console.log('[5] Numbering QIR pages...');
    const qirForStamp = await PDFDocument.load(qirBytes, { ignoreEncryption:true });
    const qirFont     = await qirForStamp.embedFont(StandardFonts.Helvetica);
    qirForStamp.getPages().forEach((page, i) => {
      const finalNum = i === 0 ? 1 : i + 2;  // page 1 stays 1; others shift +1 for index
      const { width } = page.getSize();
      const fontSize = Math.round((width / 841) * 8 * 10) / 10;
      const barH = fontSize * 2.6;
      page.drawRectangle({ x:0, y:0, width, height:barH, color:rgb(0.96,0.96,0.96), opacity:0.9 });
      const pgStr = String(finalNum);
      const pgW   = qirFont.widthOfTextAtSize(pgStr, fontSize);
      page.drawText(pgStr, { x:width/2-pgW/2, y:barH*0.25, size:fontSize, font:qirFont, color:rgb(0.20,0.20,0.20) });
    });
    const qirNumbered = Buffer.from(await qirForStamp.save());

    // 6. Merge
    console.log('[6] Merging...');
    const merged  = await PDFDocument.create();
    const qirFinal = await PDFDocument.load(qirNumbered, { ignoreEncryption:true });
    const qirPages = await merged.copyPages(qirFinal, qirFinal.getPageIndices());

    merged.addPage(qirPages[0]);                                              // p.1 header
    const idxPdf = await PDFDocument.load(indexBytes);
    const [idxPg] = await merged.copyPages(idxPdf, [0]);
    merged.addPage(idxPg);                                                    // p.2 index
    for (let i = 1; i < qirPages.length; i++) merged.addPage(qirPages[i]);   // p.3+ QIR

    for (const cert of certEntries) {
      let   numbered = await stampHeading(cert.bytes, cert.label);
      numbered = await stampPageNumbers(numbered, cert.startPage);
      const cp  = await PDFDocument.load(numbered, { ignoreEncryption:true });
      const pgs = await merged.copyPages(cp, cp.getPageIndices());
      pgs.forEach(p => merged.addPage(p));
      console.log(`    "${cert.label}" p.${cert.startPage}–${cert.startPage+cert.pageCount-1}`);
    }

    const finalBytes = await merged.save();
    const filename   = `QIR-${reportNo}-${date}.pdf`.replace(/[^a-zA-Z0-9\-_.]/g, '_');
    console.log(`✓ ${merged.getPageCount()} pages, ${(finalBytes.length/1024).toFixed(0)} KB\n`);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(Buffer.from(finalBytes));

  } catch(err) {
    console.error('✗ Error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`\nQIR Merge Server on port ${PORT}`);
  console.log(`Health: GET  /`);
  console.log(`Merge:  POST /merge\n`);
});
