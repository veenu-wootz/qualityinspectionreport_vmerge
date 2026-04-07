/**
 * mergePDFs.js
 * 1. Builds an index page (page 2 of final PDF)
 * 2. Fetches drawing PDF, stamps it, inserts as page 3 (if present)
 * 3. Stamps page numbers + logo on ALL pages (QIR + drawing + certs)
 * 4. Stamps heading text (no white strip) on first page of each cert
 * 5. Merges everything into one PDF Buffer
 *
 * Final page order:
 *   p.1          → QIR p.1 (Header + Part Info)
 *   p.2          → Index
 *   p.3          → Part Drawing p.1 (if present) ← NEW
 *   p.4..N+2     → remaining QIR pages (Dimensional, Visual, etc.)
 *   p.N+3..end   → Certificates
 *
 * Bulletproof page positioning:
 *   - Uses CropBox if present, falls back to MediaBox
 *   - All coordinates anchored to actual visible area, not assumed (0,0)
 *   - Works correctly for scanned PDFs, portrait, landscape, any size
 */

const fetch = require('node-fetch');
const { PDFDocument, rgb, StandardFonts, degrees } = require('pdf-lib');

// ── Logo — fetched once, reused across all requests ──────────
const LOGO_URL = 'https://res.cloudinary.com/dbwg6zz3l/image/upload/w_300,f_png,q_90/v1773643264/Black_Yellow_kq9kef.png';
let logoPngBytes = null;

async function ensureLogo() {
  if (logoPngBytes) return logoPngBytes;
  try {
    const res = await fetch(LOGO_URL, { timeout: 10000 });
    if (res.ok) {
      logoPngBytes = Buffer.from(await res.arrayBuffer());
      console.log(`  Logo fetched: ${(logoPngBytes.length / 1024).toFixed(0)} KB`);
    }
  } catch(e) {
    console.warn('  Logo fetch failed — will skip logo on pages:', e.message);
  }
  return logoPngBytes;
}

// ── Bulletproof visible-area helper ──────────────────────────
function getVisibleBox(page) {
  try {
    const crop = page.getCropBox();
    if (crop && crop.width > 0 && crop.height > 0) return crop;
  } catch(_) {}
  return page.getMediaBox();
}

// ── Fetch PDF from URL ────────────────────────────────────────
async function fetchPDF(url) {
  const res = await fetch(url, {
    headers: { 'User-Agent': 'QIR-Server/2.0' },
    timeout: 30000,
  });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${String(url).substring(0, 80)}`);
  return Buffer.from(await res.arrayBuffer());
}

// ── Rotation-aware footer stamping ───────────────────────────
// For pages with /Rotate, the viewer rotates the display but raw coordinates
// stay the same. We must draw the footer bar at the raw edge that corresponds
// to the VISUAL bottom, with text/logo rotated to read correctly.
//
// Visual bottom → raw edge mapping:
//   0°   → raw bottom  (y = box.y)
//   90°  → raw left    (x = box.x)           bar is a vertical strip
//   180° → raw top     (y = box.y + height)
//   270° → raw right   (x = box.x + width)   bar is a vertical strip
//
function stampFooterOnPage(page, font, fontBold, logoImg, pgStr, label, barColor, textColor) {
  const box      = getVisibleBox(page);
  const bx       = box.x, by = box.y;
  const W        = box.width, H = box.height;
  const rotation = page.getRotation().angle;

  // Use shorter visual dimension for font sizing regardless of rotation
  const shortSide = rotation === 90 || rotation === 270
    ? Math.min(W, H)   // raw dims may be swapped visually
    : Math.min(W, H);
  const fontSize = Math.round(shortSide * 0.014 * 10) / 10;
  const barH     = fontSize * 2;
  const PAD_R    = barH * 0.8;
  const LOGO_H   = barH * 0.72;
  const LOGO_PAD = barH * 0.14;

  if (rotation === 0) {
    // ── Standard: bar at raw bottom ──────────────────────────
    page.drawRectangle({ x: bx, y: by, width: W, height: barH, color: barColor, opacity: 1.0 });

    if (logoImg) {
      try {
        const logoDims = logoImg.scale(1);
        const s = LOGO_H / logoDims.height;
        page.drawImage(logoImg, { x: bx + LOGO_PAD, y: by + LOGO_PAD, width: logoDims.width * s, height: LOGO_H });
      } catch(e) {}
    }

    const textY = by + barH * 0.28;
    if (label) {
      const labelW = fontBold.widthOfTextAtSize(label, fontSize);
      page.drawText(label, { x: bx + W / 2 - labelW / 2, y: textY, size: fontSize, font: fontBold, color: textColor });
      const pgW = font.widthOfTextAtSize(pgStr, fontSize);
      page.drawText(pgStr, { x: bx + W - PAD_R - pgW, y: textY, size: fontSize, font, color: textColor });
    } else {
      const pgW = font.widthOfTextAtSize(pgStr, fontSize);
      page.drawText(pgStr, { x: bx + W - PAD_R - pgW, y: textY, size: fontSize, font, color: textColor });
    }

  } else if (rotation === 90) {
    // ── 90°: visual bottom = raw LEFT edge ───────────────────
    // Bar is a vertical strip on the left side of the raw page
    page.drawRectangle({ x: bx, y: by, width: barH, height: H, color: barColor, opacity: 1.0 });

    // Logo — rotated 90° CCW to read correctly when page is rotated 90° CW
    if (logoImg) {
      try {
        const logoDims = logoImg.scale(1);
        const s  = LOGO_H / logoDims.height;
        const lw = logoDims.width * s;
        page.drawImage(logoImg, {
          x: bx + LOGO_PAD + LOGO_H, y: by + LOGO_PAD,
          width: lw, height: LOGO_H,
          rotate: degrees(90),
        });
      } catch(e) {}
    }

    const textX = bx + barH * 0.28;
    if (label) {
      const labelW = fontBold.widthOfTextAtSize(label, fontSize);
      page.drawText(label, {
        x: textX, y: by + H / 2 - labelW / 2,
        size: fontSize, font: fontBold, color: textColor, rotate: degrees(90),
      });
      const pgW = font.widthOfTextAtSize(pgStr, fontSize);
      page.drawText(pgStr, {
        x: textX, y: by + PAD_R + pgW,
        size: fontSize, font, color: textColor, rotate: degrees(90),
      });
    } else {
      const pgW = font.widthOfTextAtSize(pgStr, fontSize);
      page.drawText(pgStr, {
        x: textX, y: by + PAD_R + pgW,
        size: fontSize, font, color: textColor, rotate: degrees(90),
      });
    }

  } else if (rotation === 180) {
    // ── 180°: visual bottom = raw TOP edge ───────────────────
    const barY = by + H - barH;
    page.drawRectangle({ x: bx, y: barY, width: W, height: barH, color: barColor, opacity: 1.0 });

    if (logoImg) {
      try {
        const logoDims = logoImg.scale(1);
        const s  = LOGO_H / logoDims.height;
        const lw = logoDims.width * s;
        // Rotated 180° — logo appears at right side reading correctly
        page.drawImage(logoImg, {
          x: bx + W - LOGO_PAD, y: barY + barH - LOGO_PAD,
          width: lw, height: LOGO_H,
          rotate: degrees(180),
        });
      } catch(e) {}
    }

    const textY = barY + barH - barH * 0.28 - fontSize;
    if (label) {
      const labelW = fontBold.widthOfTextAtSize(label, fontSize);
      page.drawText(label, {
        x: bx + W / 2 + labelW / 2, y: textY,
        size: fontSize, font: fontBold, color: textColor, rotate: degrees(180),
      });
      const pgW = font.widthOfTextAtSize(pgStr, fontSize);
      page.drawText(pgStr, {
        x: bx + PAD_R + pgW, y: textY,
        size: fontSize, font, color: textColor, rotate: degrees(180),
      });
    } else {
      const pgW = font.widthOfTextAtSize(pgStr, fontSize);
      page.drawText(pgStr, {
        x: bx + PAD_R + pgW, y: textY,
        size: fontSize, font, color: textColor, rotate: degrees(180),
      });
    }

  } else if (rotation === 270) {
    // ── 270°: visual bottom = raw RIGHT edge ─────────────────
    const barX = bx + W - barH;
    page.drawRectangle({ x: barX, y: by, width: barH, height: H, color: barColor, opacity: 1.0 });

    if (logoImg) {
      try {
        const logoDims = logoImg.scale(1);
        const s  = LOGO_H / logoDims.height;
        const lw = logoDims.width * s;
        page.drawImage(logoImg, {
          x: barX + barH - LOGO_PAD, y: by + H - LOGO_PAD,
          width: lw, height: LOGO_H,
          rotate: degrees(270),
        });
      } catch(e) {}
    }

    const textX = barX + barH - barH * 0.28;
    if (label) {
      const labelW = fontBold.widthOfTextAtSize(label, fontSize);
      page.drawText(label, {
        x: textX, y: by + H / 2 + labelW / 2,
        size: fontSize, font: fontBold, color: textColor, rotate: degrees(270),
      });
      const pgW = font.widthOfTextAtSize(pgStr, fontSize);
      page.drawText(pgStr, {
        x: textX, y: by + H - PAD_R,
        size: fontSize, font, color: textColor, rotate: degrees(270),
      });
    } else {
      const pgW = font.widthOfTextAtSize(pgStr, fontSize);
      page.drawText(pgStr, {
        x: textX, y: by + H - PAD_R,
        size: fontSize, font, color: textColor, rotate: degrees(270),
      });
    }

  } else {
    // Non-standard rotation — fall back to raw bottom, best effort
    page.drawRectangle({ x: bx, y: by, width: W, height: barH, color: barColor, opacity: 1.0 });
    const pgW = font.widthOfTextAtSize(pgStr, fontSize);
    page.drawText(pgStr, { x: bx + W - PAD_R - pgW, y: by + barH * 0.28, size: fontSize, font, color: textColor });
  }
}

// ── Stamp page number + logo on every page of a PDF ──────────
async function stampPageNumbers(pdfBytes, startPageNum, label = null) {
  const pdf       = await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
  const font      = await pdf.embedFont(StandardFonts.Helvetica);
  const fontBold  = await pdf.embedFont(StandardFonts.HelveticaBold);
  const logoBytes = await ensureLogo();

  let logoImg = null;
  if (logoBytes) {
    try { logoImg = await pdf.embedPng(logoBytes); } catch(e) {}
  }

  const barColor  = rgb(0.96, 0.96, 0.96);
  const textColor = rgb(0.20, 0.20, 0.20);

  pdf.getPages().forEach((page, i) => {
    const _box = getVisibleBox(page); const _rot = page.getRotation().angle;
    console.log(`  [stamp] page ${i} rot=${_rot} box={x:${_box.x.toFixed(1)},y:${_box.y.toFixed(1)},w:${_box.width.toFixed(1)},h:${_box.height.toFixed(1)}}`);
    stampFooterOnPage(page, font, fontBold, logoImg,
      String(startPageNum + i), label, barColor, textColor);
  });

  return Buffer.from(await pdf.save());
}

// ── Heading stamp — bold text only, NO white strip ────────────
async function stampHeading(pdfBytes, label) {
  if (!label?.trim()) return pdfBytes;
  try {
    const pdf  = await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
    const page = pdf.getPages()[0];
    const font = await pdf.embedFont(StandardFonts.HelveticaBold);

    const box    = getVisibleBox(page);
    const bx     = box.x, by = box.y;
    const width  = box.width, height = box.height;

    const shortSide = Math.min(width, height);
    const fontSize  = Math.round(shortSide * 0.019 * 10) / 10;
    const topMargin = fontSize * 1.5;

    const textW = font.widthOfTextAtSize(label, fontSize);
    page.drawText(label, {
      x:    bx + (width - textW) / 2,
      y:    by + height - topMargin - fontSize,
      size: fontSize, font,
      color: rgb(0.10, 0.10, 0.10),
    });

    return Buffer.from(await pdf.save());
  } catch(e) { return pdfBytes; }
}

// ── Prepare drawing page: stamp heading + footer, return single-page PDF ──
async function prepareDrawingPage(drawingUrl, pageNum) {
  console.log(`  Fetching drawing PDF: ${String(drawingUrl).substring(0, 70)}...`);
  const rawBytes  = await fetchPDF(drawingUrl);

  const srcPdf    = await PDFDocument.load(rawBytes, { ignoreEncryption: true });
  const singleDoc = await PDFDocument.create();
  const [page1]   = await singleDoc.copyPages(srcPdf, [0]);
  singleDoc.addPage(page1);
  let drawingBytes = Buffer.from(await singleDoc.save());

  drawingBytes = await stampHeading(drawingBytes, 'Part Drawing');
  drawingBytes = await stampPageNumbers(drawingBytes, pageNum);

  console.log(`  Drawing page prepared (final p.${pageNum})`);
  return drawingBytes;
}

// ── Fit external page to A4 landscape ────────────────────────
// Target: 841.89 × 595.28 pt (A4 landscape)
// Rules:
//   - If either dimension exceeds target → scale DOWN uniformly so both fit
//   - If both dimensions within target   → no scaling, just centre
//   - In both cases: set page to target size, white background, content centred
//   - Never scales up (scale capped at 1.0)
// NOTE: kept for future use — not currently called
const TARGET_W = 841.89;
const TARGET_H = 595.28;

async function fitPageToA4Landscape(pdfBytes) {
  const srcPdf  = await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
  const outPdf  = await PDFDocument.create();

  for (let i = 0; i < srcPdf.getPageCount(); i++) {
    const [embedded] = await outPdf.embedPdf(srcPdf, [i]);
    const srcPage = srcPdf.getPages()[i];
    const pageW   = srcPage.getWidth();
    const pageH   = srcPage.getHeight();
    const scale   = Math.min(TARGET_W / pageW, TARGET_H / pageH, 1.0);
    const scaledW = pageW  * scale;
    const scaledH = pageH  * scale;
    const offsetX = (TARGET_W - scaledW) / 2;
    const offsetY = (TARGET_H - scaledH) / 2;
    const newPage = outPdf.addPage([TARGET_W, TARGET_H]);
    newPage.drawRectangle({ x: 0, y: 0, width: TARGET_W, height: TARGET_H, color: rgb(1, 1, 1), opacity: 1.0 });
    newPage.drawPage(embedded, { x: offsetX, y: offsetY, width: scaledW, height: scaledH, opacity: 1.0 });
  }

  return Buffer.from(await outPdf.save());
}

// ── Index page ────────────────────────────────────────────────
async function buildIndexPage({ qirPageCount, hasDrawing, certEntries }) {
  const doc = await PDFDocument.create();
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

  page.drawRectangle({ x: 0, y: 0, width: W, height: H, color: rgb(1, 1, 1) });

  const shortSide    = Math.min(W, H);
  const PG_FONT_SIZE = Math.round(shortSide * 0.014 * 10) / 10;
  const PG_BAR_H     = PG_FONT_SIZE * 2;

  page.drawRectangle({ x: 0, y: 0, width: W, height: PG_BAR_H, color: rgb(0.96, 0.96, 0.96), opacity: 1.0 });

  const logoBytes = await ensureLogo();
  const LOGO_H    = PG_BAR_H * 0.72;
  const LOGO_PAD  = PG_BAR_H * 0.14;
  if (logoBytes) {
    try {
      const logoImg  = await doc.embedPng(logoBytes);
      const logoDims = logoImg.scale(1);
      const scale    = LOGO_H / logoDims.height;
      const lw       = logoDims.width * scale;
      page.drawImage(logoImg, { x: LOGO_PAD, y: LOGO_PAD, width: lw, height: LOGO_H });
    } catch(e) {}
  }

  const p2Str = '2';
  const p2W   = fontNormal.widthOfTextAtSize(p2Str, PG_FONT_SIZE);
  page.drawText(p2Str, {
    x: W - PG_BAR_H * 0.8 - p2W, y: PG_BAR_H * 0.25,
    size: PG_FONT_SIZE, font: fontNormal, color: rgb(0.20, 0.20, 0.20),
  });

  const tocLabel = 'Table of Content';
  const tocW = fontBold.widthOfTextAtSize(tocLabel, 11);
  page.drawText(tocLabel, { x: W / 2 - tocW / 2, y: H - 52, size: 11, font: fontBold, color: BLACK });

  const TBL_X = 40, TBL_W = W - 80, PAD = 10, ROW_H = 24;
  let rowY = H - 72;
  const tableTopY = rowY;

  function drawRow(label, pageNum, isSub = false, isHeader = false) {
    const bg = isHeader ? HDRBG : (isSub ? rgb(1, 1, 1) : OFFWHT);
    page.drawRectangle({ x: TBL_X, y: rowY - ROW_H, width: TBL_W, height: ROW_H, color: bg });
    page.drawLine({ start: { x: TBL_X, y: rowY - ROW_H }, end: { x: TBL_X + TBL_W, y: rowY - ROW_H },
      thickness: 0.3, color: LGRAY });
    const indent   = isSub ? 18 : 0;
    const fontSize = isHeader ? 8.5 : 8;
    const font     = isHeader ? fontBold : fontNormal;
    const color    = isHeader ? BLACK : isSub ? MGRAY : DGRAY;
    const textY    = rowY - ROW_H + (ROW_H - fontSize) / 2 + 1;
    page.drawText(String(label), { x: TBL_X + PAD + indent, y: textY, size: fontSize, font, color });
    const pgStr = String(pageNum);
    const pgW   = (isHeader ? fontBold : fontNormal).widthOfTextAtSize(pgStr, fontSize);
    page.drawText(pgStr, { x: TBL_X + TBL_W - PAD - pgW, y: textY, size: fontSize,
      font: isHeader ? fontBold : fontNormal, color: isHeader ? BLACK : DGRAY });
    rowY -= ROW_H;
  }

  page.drawLine({ start: { x: TBL_X, y: tableTopY }, end: { x: TBL_X + TBL_W, y: tableTopY }, thickness: 0.6, color: DGRAY });

  drawRow('Section', 'Page', false, true);
  drawRow('Part Information', 1);
  drawRow('Content Table', 2);

  let nextQirPage = hasDrawing ? 4 : 3;
  if (hasDrawing) drawRow('Part Drawing', 3);

  if (certEntries._hasDim || certEntries._hasVis) {
    drawRow('Inspection', nextQirPage);
    if (certEntries._hasDim) drawRow('Dimensional Inspection', '', true);
    if (certEntries._hasVis) drawRow('Visual Inspection',      '', true);
  }

  const certStart = qirPageCount + 1 + (hasDrawing ? 1 : 0) + 1;
  drawRow('Tests & Certificates', certEntries.length > 0 ? certStart : '—');
  for (const c of certEntries) drawRow(c.label, c.startPage, true);

  page.drawLine({ start: { x: TBL_X, y: rowY }, end: { x: TBL_X + TBL_W, y: rowY }, thickness: 0.6, color: DGRAY });
  page.drawLine({ start: { x: TBL_X, y: tableTopY }, end: { x: TBL_X, y: rowY }, thickness: 0.6, color: DGRAY });
  page.drawLine({ start: { x: TBL_X + TBL_W, y: tableTopY }, end: { x: TBL_X + TBL_W, y: rowY }, thickness: 0.6, color: DGRAY });

  return Buffer.from(await doc.save());
}

// ── Watermark stamp ───────────────────────────────────────────
async function stampWatermark(pdfBytes) {
  const pdf  = await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
  const font = await pdf.embedFont(StandardFonts.HelveticaBold);

  pdf.getPages().forEach(page => {
    const box      = getVisibleBox(page);
    const W        = box.width, H = box.height;
    const cx       = box.x + W / 2;
    const cy       = box.y + H / 2;

    // Font size: 10% of shorter side — large enough to span diagonally but not crop
    const fontSize = Math.min(W, H) * 0.12;
    const textW    = font.widthOfTextAtSize('UNVERIFIED', fontSize);
    const angle    = 45 * Math.PI / 180;

    // Offset x,y so text centre lands at page centre after 45° rotation
    const x = cx - (textW / 2) * Math.cos(angle) + (fontSize / 2) * Math.sin(angle);
    const y = cy - (textW / 2) * Math.sin(angle) - (fontSize / 2) * Math.cos(angle);

    page.drawText('UNVERIFIED', {
      x, y,
      size:    fontSize,
      font,
      color:   rgb(0.75, 0.75, 0.75),
      opacity: 0.15,
      rotate:  degrees(45),
    });
  });

  return Buffer.from(await pdf.save());
}

// async function stampWatermark(pdfBytes) {
//   const pdf  = await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
//   const font = await pdf.embedFont(StandardFonts.HelveticaBold);

//   pdf.getPages().forEach(page => {
//     const box      = getVisibleBox(page);
//     const cx       = box.x + box.width  / 2;
//     const cy       = box.y + box.height / 2;
//     const fontSize = Math.min(box.width, box.height) * 0.15;
//     const textW    = font.widthOfTextAtSize('UNVERIFIED', fontSize);

//     page.drawText('UNVERIFIED', {
//       x:       cx - textW / 2,
//       y:       cy - fontSize / 2,
//       size:    fontSize,
//       font,
//       color:   rgb(0.75, 0.75, 0.75),
//       opacity: 0.15,
//       rotate:  degrees(45),
//     });
//   });

//   return Buffer.from(await pdf.save());
// }

// ── Main ──────────────────────────────────────────────────────
async function buildMergedPDF(qirBuffer, certs = [], meta = {}) {
  await ensureLogo();

  const qirPdf       = await PDFDocument.load(qirBuffer, { ignoreEncryption: true });
  const qirPageCount = qirPdf.getPageCount();
  const hasDrawing   = !!meta.partDrawingUrl;
  console.log(`  QIR pages: ${qirPageCount}, hasDrawing: ${hasDrawing}`);

  const [drawingResult, ...certResults] = await Promise.allSettled([
    hasDrawing ? fetchPDF(meta.partDrawingUrl) : Promise.resolve(null),
    ...certs.filter(c => c && c.url && c.url.trim()).map(async (cert) => {
      console.log(`    Fetching cert: ${cert.label} — ${String(cert.url).substring(0, 60)}...`);
      const bytes = await fetchPDF(cert.url);
      const pdf   = await PDFDocument.load(bytes, { ignoreEncryption: true });
      return { label: cert.label || 'Certificate', bytes, pageCount: pdf.getPageCount() };
    }),
  ]);

  let drawingRawBytes = null;
  if (hasDrawing) {
    if (drawingResult.status === 'fulfilled' && drawingResult.value) {
      drawingRawBytes = drawingResult.value;
    } else {
      console.error(`  Drawing fetch failed: ${drawingResult.reason?.message}`);
    }
  }

  const certData = [];
  for (const r of certResults) {
    if (r.status === 'fulfilled') certData.push(r.value);
    else console.error(`  Cert fetch failed: ${r.reason?.message}`);
  }

  const drawingPageNum  = 3;
  const qirRemapOffset  = hasDrawing ? 3 : 2;
  const certStartPage   = qirPageCount + 1 + (hasDrawing ? 1 : 0) + 1;
  let   runningPage     = certStartPage;

  const certEntries = certData.map(c => {
    const entry = { ...c, startPage: runningPage };
    runningPage += c.pageCount;
    return entry;
  });
  certEntries._hasDrawing = hasDrawing;
  certEntries._hasDim     = meta.hasDim  || false;
  certEntries._hasVis     = meta.hasVis  || false;

  console.log(`  Page layout: QIR(${qirPageCount}) + Index + ${hasDrawing ? 'Drawing + ' : ''}Certs → total ~${runningPage - 1}`);

  console.log('  Building index page...');
  const indexBytes = await buildIndexPage({ qirPageCount, hasDrawing, certEntries });

  const qirForStamp = await PDFDocument.load(qirBuffer, { ignoreEncryption: true });
  const qirFont     = await qirForStamp.embedFont(StandardFonts.Helvetica);
  const logoBytes   = logoPngBytes;
  let   qirLogoImg  = null;
  if (logoBytes) {
    try { qirLogoImg = await qirForStamp.embedPng(logoBytes); } catch(e) {}
  }

  qirForStamp.getPages().forEach((page, i) => {
    const finalNum  = i === 0 ? 1 : i + qirRemapOffset;
    const box       = getVisibleBox(page);
    const bx        = box.x, by = box.y;
    const width     = box.width, height = box.height;
    const shortSide = Math.min(width, height);
    const fontSize  = Math.round(shortSide * 0.014 * 10) / 10;
    const barH      = fontSize * 2;

    page.drawRectangle({ x: bx, y: by, width, height: barH, color: rgb(0.96, 0.96, 0.96), opacity: 1.0 });

    if (qirLogoImg) {
      try {
        const LOGO_H   = barH * 0.72;
        const LOGO_PAD = barH * 0.14;
        const logoDims = qirLogoImg.scale(1);
        const scale    = LOGO_H / logoDims.height;
        const lw       = logoDims.width * scale;
        page.drawImage(qirLogoImg, { x: bx + LOGO_PAD, y: by + LOGO_PAD, width: lw, height: LOGO_H });
      } catch(e) {}
    }

    const pgStr = String(finalNum);
    const PAD_R = barH * 0.8;
    const pgW   = qirFont.widthOfTextAtSize(pgStr, fontSize);
    page.drawText(pgStr, {
      x: bx + width - PAD_R - pgW,
      y: by + barH * 0.28,
      size: fontSize, font: qirFont, color: rgb(0.20, 0.20, 0.20),
    });
  });
  const qirNumbered = Buffer.from(await qirForStamp.save());

  let drawingStamped = null;
  if (drawingRawBytes) {
    try {
      const srcPdf    = await PDFDocument.load(drawingRawBytes, { ignoreEncryption: true });
      const singleDoc = await PDFDocument.create();
      const [pg1]     = await singleDoc.copyPages(srcPdf, [0]);
      singleDoc.addPage(pg1);
      let drawBytes   = Buffer.from(await singleDoc.save());

      // Fit to A4 landscape (scale down if needed, centre, white background)
      // drawBytes = await fitPageToA4Landscape(drawBytes);

      drawBytes = await stampHeading(drawBytes, 'Part Drawing');
      drawBytes = await stampPageNumbers(drawBytes, drawingPageNum);

      drawingStamped = drawBytes;
      console.log(`  Drawing page stamped as p.${drawingPageNum}`);
    } catch(e) {
      console.error(`  Drawing page preparation failed: ${e.message}`);
    }
  }

  console.log('  Merging...');
  const merged   = await PDFDocument.create();
  const qirFinal = await PDFDocument.load(qirNumbered, { ignoreEncryption: true });
  const qirPages = await merged.copyPages(qirFinal, qirFinal.getPageIndices());

  merged.addPage(qirPages[0]);

  const idxPdf    = await PDFDocument.load(indexBytes);
  const [idxPage] = await merged.copyPages(idxPdf, [0]);
  merged.addPage(idxPage);

  if (drawingStamped) {
    const drawPdf    = await PDFDocument.load(drawingStamped, { ignoreEncryption: true });
    const [drawPage] = await merged.copyPages(drawPdf, [0]);
    merged.addPage(drawPage);
  }

  for (let i = 1; i < qirPages.length; i++) merged.addPage(qirPages[i]);

  // Certs — rotation-aware footer on every page, no heading on first page
  for (const cert of certEntries) {
    // To re-enable scaling, uncomment the next two lines (comment out the third):
    // let bytes = await fitPageToA4Landscape(cert.bytes);
    // bytes     = await stampPageNumbers(bytes, cert.startPage, cert.label);
    let bytes = await stampPageNumbers(cert.bytes, cert.startPage, cert.label);
    const cp  = await PDFDocument.load(bytes, { ignoreEncryption: true });
    const pgs = await merged.copyPages(cp, cp.getPageIndices());
    pgs.forEach(p => merged.addPage(p));
    console.log(`    "${cert.label}" p.${cert.startPage}–${cert.startPage + cert.pageCount - 1}`);
  }
  
  console.log(`  Final: ${merged.getPageCount()} pages`);
  const finalBytes = Buffer.from(await merged.save());
  if (!meta.verifiedBy || meta.verifiedBy === 'Unverified') {
    console.log('  Stamping UNVERIFIED watermark...');
    return await stampWatermark(finalBytes);
  }
  return finalBytes;
}

module.exports = { buildMergedPDF };
