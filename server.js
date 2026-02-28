/**
 * server.js
 * QIR PDF Generation Server
 *
 * Single endpoint: POST /generate
 * Receives Clappia webhook → generates QIR PDF → fetches & merges cert PDFs
 * → emails final PDF to submitter + internal team
 *
 * Deploy on Render:
 *   1. Push this folder to a GitHub repo
 *   2. New Web Service on Render → connect repo
 *   3. Build command: npm install
 *   4. Start command: node server.js
 *   5. Add environment variables in Render dashboard
 */

require('dotenv').config();

const express = require('express');
const { generateQIR }     = require('./generateQIR');
const { buildMergedPDF }  = require('./mergePDFs');
const { sendQIREmail }    = require('./sendEmail');

const app  = express();
const PORT = process.env.PORT || 3000;

// Allow requests from any origin (needed for local HTML file testing + Clappia)
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Parse JSON bodies up to 50MB (base64 images can be large)
app.use(express.json({ limit: '50mb' }));

// ── Health check ─────────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({ status: 'ok', service: 'QIR Generator', version: '1.0.0' });
});

// ── Main endpoint ─────────────────────────────────────────────
/**
 * POST /generate
 *
 * Expected body from Clappia REST API action:
 * {
 *   // ── Report header ──
 *   "report_no":        "<<[report_no]>>",
 *   "submission_date":  "<<[$submission_date]>>",
 *   "your_email":       "<<[your_email]>>",       // Clappia submitter email
 *
 *   // ── Part info ──
 *   "part_name":   "<<[part_name]>>",
 *   "rm_grade":    "<<[rm_grade]>>",
 *   "customer":    "<<[customer]>>",
 *   "item_code":   "<<[item_code]>>",
 *   "heat_no":     "<<[heat_no]>>",
 *   "order_qty":   "<<[order_qty]>>",
 *   "remarks":     "<<[remarks]>>",
 *   "conclusion":  "<<[conclusion]>>",
 *
 *   // ── Images (base64 data URLs OR null) ──
 *   "part_drawing": "<<[part_drawi]>>",
 *   "insp_image":   null,
 *
 *   // ── Dimensional rows ──
 *   "sampleCount": 5,
 *   "dimRows": [
 *     {
 *       "index": 1,
 *       "parameter":  "<<[parameter#1]>>",
 *       "specificat": "<<[specificat#1]>>",
 *       "instrument": "<<[instrument#1]>>",
 *       "samples":    ["<<[sample_1#1]>>", "<<[sample_2#1]>>", ...],
 *       "status_1":   "<<[status_1#1]>>",
 *       "qc_photo":   "<<[qc_photo#1]>>"           // base64 or null
 *     },
 *     ...up to N rows
 *   ],
 *
 *   // ── Visual rows ──
 *   "visRows": [
 *     {
 *       "index":     1,
 *       "parameter": "<<[visual_inspection#1]>>",
 *       "status":    "<<[status_visual#1]>>",
 *       "comments":  "<<[comments#1]>>",
 *       "photo":     null
 *     },
 *     ...
 *   ],
 *
 *   // ── Certificate PDFs to merge (pre-signed S3 URLs from Clappia) ──
 *   // Must be fetched within 5 minutes of Clappia sending this webhook.
 *   "certificates": [
 *     { "label": "Mill TC",         "url": "https://s3.amazonaws.com/..." },
 *     { "label": "Hardness Report", "url": "https://s3.amazonaws.com/..." }
 *   ]
 * }
 */
app.post('/generate', async (req, res) => {
  const startTime = Date.now();
  const data = req.body;

  // Basic validation
  if (!data || !data.part_name) {
    return res.status(400).json({
      error: 'Invalid payload — at minimum part_name is required'
    });
  }

  const reportNo = data.report_no || `QIR-${new Date().getFullYear()}-XXXX`;
  const date     = data.submission_date || new Date().toISOString().split('T')[0];
  const filename = `QIR-${reportNo}-${date}.pdf`.replace(/[^a-zA-Z0-9\-_.]/g, '_');

  console.log(`\n── New report request ──────────────────────`);
  console.log(`  Report No:  ${reportNo}`);
  console.log(`  Part:       ${data.part_name}`);
  console.log(`  Submitter:  ${data.your_email || 'not provided'}`);
  console.log(`  Certs:      ${(data.certificates || []).length}`);

  try {

    // ── Step 1: Generate QIR PDF ────────────────────────────
    console.log('\n[1/3] Generating QIR PDF...');
    const qirBuffer = generateQIR(data);
    console.log(`  Done — ${(qirBuffer.length / 1024).toFixed(0)} KB`);

    // ── Step 2: Fetch certs + merge ─────────────────────────
    console.log('\n[2/3] Fetching certificates and merging...');
    const mergedBuffer = await buildMergedPDF(qirBuffer, data.certificates || []);
    console.log(`  Done — merged PDF: ${(mergedBuffer.length / 1024).toFixed(0)} KB`);

    // ── Step 3: Send email ───────────────────────────────────
    console.log('\n[3/3] Sending email...');
    await sendQIREmail(data, mergedBuffer, filename);

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`\n✓ Complete in ${elapsed}s — ${filename}`);
    console.log(`────────────────────────────────────────────\n`);

    res.json({
      success:  true,
      filename,
      elapsed:  `${elapsed}s`,
      pages:    'generated',
      certs:    (data.certificates || []).length,
    });

  } catch (err) {
    console.error(`\n✗ Error processing ${reportNo}:`, err);
    res.status(500).json({
      error:   err.message,
      report:  reportNo,
    });
  }
});

// ── Start ────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\nQIR Server running on port ${PORT}`);
  console.log(`POST http://localhost:${PORT}/generate`);
  console.log(`Health: http://localhost:${PORT}/\n`);

  // Warn about missing config on startup
  if (!process.env.GMAIL_USER)         console.warn('⚠  GMAIL_USER not set');
  if (!process.env.GMAIL_APP_PASSWORD) console.warn('⚠  GMAIL_APP_PASSWORD not set');
  if (!process.env.INTERNAL_EMAILS)    console.warn('⚠  INTERNAL_EMAILS not set');
});

// ── MERGE ENDPOINT ─────────────────────────────────────────────
/**
 * POST /merge
 * Accepts QIR PDF + certs (as URLs or base64), builds index page,
 * merges everything, returns the final PDF as binary response.
 *
 * Body:
 * {
 *   reportNo, partName, date,
 *   qirSource:    { type: 'url'|'base64', value: '...', name?: '...' }
 *   certificates: [{ label, type: 'url'|'base64', value, name? }, ...]
 * }
 */
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const fetch = require('node-fetch');

async function loadPDFFromSource(src) {
  if (src.type === 'url') {
    const res = await fetch(src.value, { timeout: 30000 });
    if (!res.ok) throw new Error(`HTTP ${res.status} fetching PDF`);
    return Buffer.from(await res.arrayBuffer());
  } else {
    return Buffer.from(src.value, 'base64');
  }
}

async function buildIndexPage(opts) {
  const { reportNo, partName, date, inspectionPage, certEntries } = opts;
  const doc = await PDFDocument.create();
  const W = 841.89, H = 595.28;
  const page = doc.addPage([W, H]);
  const fontBold   = await doc.embedFont(StandardFonts.HelveticaBold);
  const fontNormal = await doc.embedFont(StandardFonts.Helvetica);
  const DARK  = rgb(0.10, 0.10, 0.18);
  const RED   = rgb(0.78, 0.29, 0.17);
  const GRAY  = rgb(0.50, 0.50, 0.50);
  const LGRAY = rgb(0.93, 0.93, 0.93);

  page.drawRectangle({ x:0, y:H-6, width:W, height:6, color:RED });
  page.drawText('Quality Inspection Report', { x:40, y:H-44, size:22, font:fontBold, color:DARK });
  page.drawText(`${reportNo}   ·   ${partName}   ·   ${date}`, { x:40, y:H-64, size:10, font:fontNormal, color:GRAY });
  page.drawText('INDEX', { x:40, y:H-100, size:11, font:fontBold, color:RED });
  page.drawLine({ start:{x:40,y:H-108}, end:{x:W-40,y:H-108}, thickness:0.5, color:LGRAY });

  const ENTRY_H = 32, COL_LEFT = 40, COL_PG = W-80;
  let rowY = H-118;

  function drawRow(label, pageNum, isSub=false, isHeader=false) {
    const bgColor  = isHeader ? rgb(0.96,0.96,0.97) : rgb(1,1,1);
    page.drawRectangle({ x:36, y:rowY-ENTRY_H+6, width:W-72, height:ENTRY_H-2, color:bgColor });
    const indent   = isSub ? 24 : 0;
    const fontSize = isHeader ? 10 : 9.5;
    const font     = isHeader ? fontBold : fontNormal;
    const color    = isHeader ? DARK : isSub ? GRAY : DARK;
    page.drawText(label, { x:COL_LEFT+indent, y:rowY-10, size:fontSize, font, color });
    const pgStr = String(pageNum);
    const pgW   = fontBold.widthOfTextAtSize(pgStr, fontSize);
    page.drawText(pgStr, { x:COL_PG-pgW, y:rowY-10, size:fontSize, font:fontBold, color:isHeader?DARK:RED });
    if (!isHeader) {
      const labelW = font.widthOfTextAtSize(label, fontSize);
      for (let dx = COL_LEFT+indent+labelW+8; dx < COL_PG-pgW-8; dx += 4)
        page.drawCircle({ x:dx, y:rowY-7, size:0.7, color:LGRAY });
    }
    rowY -= ENTRY_H;
  }

  drawRow('Report Header & Part Information', 1, false, true);
  drawRow('Index',                            2, false, true);
  drawRow('Part Drawing',                     3, false, true);
  drawRow('Inspection',          inspectionPage, false, true);
  drawRow('Dimensional Inspection', inspectionPage, true, false);
  drawRow('Visual Inspection',   '—', true, false);
  rowY -= 6;
  page.drawLine({ start:{x:40,y:rowY+10}, end:{x:W-40,y:rowY+10}, thickness:0.5, color:LGRAY });
  rowY -= 10;
  drawRow('Tests & Certificates', certEntries[0]?.startPage || '—', false, true);
  for (const c of certEntries) drawRow(c.label, c.startPage, true, false);

  page.drawRectangle({ x:0, y:0, width:W, height:4, color:RED });
  const p2W = fontNormal.widthOfTextAtSize('2', 8);
  page.drawText('2', { x:W/2-p2W/2, y:10, size:8, font:fontNormal, color:GRAY });

  return Buffer.from(await doc.save());
}

async function stampPageNumbers(pdfBytes, startPageNum, label) {
  const pdf      = await PDFDocument.load(pdfBytes, { ignoreEncryption:true });
  const font     = await pdf.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdf.embedFont(StandardFonts.HelveticaBold);
  pdf.getPages().forEach((page, i) => {
    const { width, height } = page.getSize();
    page.drawRectangle({ x:0, y:0, width, height:18, color:rgb(0.97,0.97,0.97), opacity:0.9 });
    if (i===0 && label)
      page.drawText(label, { x:20, y:5, size:7, font, color:rgb(0.55,0.55,0.55) });
    const pgStr = String(startPageNum+i);
    const pgW   = fontBold.widthOfTextAtSize(pgStr, 9);
    page.drawText(pgStr, { x:width/2-pgW/2, y:5, size:9, font:fontBold, color:rgb(0.78,0.29,0.17) });
  });
  return Buffer.from(await pdf.save());
}

async function stampHeadingOnPdf(pdfBytes, label) {
  if (!label?.trim()) return pdfBytes;
  try {
    const pdf  = await PDFDocument.load(pdfBytes, { ignoreEncryption:true });
    const page = pdf.getPages()[0];
    const font = await pdf.embedFont(StandardFonts.HelveticaBold);
    const { width, height } = page.getSize();
    const fontSize=14, barH=28;
    page.drawRectangle({ x:0, y:height-barH, width, height:barH, color:rgb(1,1,1), opacity:0.9 });
    page.drawLine({ start:{x:0,y:height-barH}, end:{x:width,y:height-barH}, thickness:1.5, color:rgb(0.78,0.29,0.17) });
    const textW = font.widthOfTextAtSize(label, fontSize);
    page.drawText(label, { x:(width-textW)/2, y:height-barH+(barH-fontSize)/2+2, size:fontSize, font, color:rgb(0.1,0.1,0.18) });
    return Buffer.from(await pdf.save());
  } catch { return pdfBytes; }
}

app.post('/merge', async (req, res) => {
  const { reportNo='QIR', partName='', date='', qirSource, certificates=[] } = req.body;
  console.log(`\n── Merge request: ${reportNo} — ${certificates.length} cert(s)`);

  try {
    // Load QIR
    console.log('  Loading QIR PDF...');
    const qirBytes    = await loadPDFFromSource(qirSource);
    const qirPdf      = await PDFDocument.load(qirBytes, { ignoreEncryption:true });
    const qirPageCount= qirPdf.getPageCount();
    console.log(`  QIR pages: ${qirPageCount}`);

    // Load certs
    const certData = [];
    for (const cert of certificates) {
      if (!cert.value?.trim()) continue;
      try {
        const bytes     = await loadPDFFromSource(cert);
        const pdf       = await PDFDocument.load(bytes, { ignoreEncryption:true });
        certData.push({ label: cert.label||'Certificate', bytes, pageCount: pdf.getPageCount() });
        console.log(`  Cert "${cert.label}" — ${pdf.getPageCount()} page(s)`);
      } catch(e) { console.warn(`  Skipped cert: ${e.message}`); }
    }

    // Calculate pages
    const INSPECTION_PAGE = 4;
    const certStartPage   = qirPageCount + 2;
    let   runningPage     = certStartPage;
    const certEntries     = certData.map(c => {
      const entry = { label:c.label, startPage:runningPage, pageCount:c.pageCount, bytes:c.bytes };
      runningPage += c.pageCount;
      return entry;
    });

    // Build index
    console.log('  Building index page...');
    const indexBytes = await buildIndexPage({ reportNo, partName, date, inspectionPage:INSPECTION_PAGE, certEntries });

    // Merge
    console.log('  Merging...');
    const merged   = await PDFDocument.create();
    const qirPages = await merged.copyPages(qirPdf, qirPdf.getPageIndices());

    merged.addPage(qirPages[0]);
    const idxPdf    = await PDFDocument.load(indexBytes);
    const [idxPage] = await merged.copyPages(idxPdf, [0]);
    merged.addPage(idxPage);
    for (let i=1; i<qirPages.length; i++) merged.addPage(qirPages[i]);

    for (const cert of certEntries) {
      let bytes = await stampHeadingOnPdf(cert.bytes, cert.label);
      bytes     = await stampPageNumbers(bytes, cert.startPage, cert.label);
      const cp  = await PDFDocument.load(bytes, { ignoreEncryption:true });
      const pgs = await merged.copyPages(cp, cp.getPageIndices());
      pgs.forEach(p => merged.addPage(p));
    }

    const finalBytes = await merged.save();
    const filename   = `QIR-${reportNo}-${date}.pdf`.replace(/[^a-zA-Z0-9\-_.]/g,'_');
    console.log(`  Done — ${merged.getPageCount()} pages, ${(finalBytes.length/1024).toFixed(0)} KB`);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(Buffer.from(finalBytes));

  } catch(err) {
    console.error('Merge error:', err);
    res.status(500).json({ error: err.message });
  }
});
