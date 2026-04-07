/**
 * generateQIR.js
 * Generates the QIR PDF using jsPDF + jspdf-autotable.
 * Updated for AppSheet data structure.
 * Images are fetched from AppSheet URLs before drawing.
 * Returns a Buffer.
 *
 * CHANGE: Part Drawing page removed from here.
 *         mergePDFs.js now fetches the drawing PDF and inserts
 *         its first page directly as page 3 of the final document.
 *
 * CHANGE: Part info table — empty fields are skipped entirely (no blank rows).
 * CHANGE: Dim table  — Instrument col hidden if all blank; Photo col hidden if no photos.
 * CHANGE: Vis table  — Photo col hidden if no photos exist.
 */

const { jsPDF }  = require('jspdf');
require('jspdf-autotable');
const fetch      = require('node-fetch');

// ── Constants ────────────────────────────────────────────────
const PW = 297, PH = 210, ML = 10, MR = 10, MT = 12, MB = 12;
const CW    = PW - ML - MR;
const GRAY  = [240, 240, 240];
const BORDER= [180, 180, 180];
const DARK  = [26,  26,  46 ];

// ── Helpers ──────────────────────────────────────────────────

function sectionHeading(doc, label, y) {
  doc.setFontSize(12);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(...DARK);
  doc.text(label, PW / 2, y, { align: 'center' });
  return y + 7;
}

/** Fetch an image URL → base64 data URL. Returns null on failure. */
async function fetchImageAsDataUrl(url) {
  if (!url || !url.trim()) return null;
  try {
    const res = await fetch(url, { timeout: 15000 });
    if (!res.ok) return null;
    const contentType = res.headers.get('content-type') || 'image/jpeg';
    const buffer = await res.arrayBuffer();
    const b64 = Buffer.from(buffer).toString('base64');
    return `data:${contentType};base64,${b64}`;
  } catch (e) {
    console.warn(`  fetchImage failed for ${String(url).substring(0, 60)}: ${e.message}`);
    return null;
  }
}

/** Draw an image inside a cell, preserving aspect ratio, centred. */
function drawImgInCell(doc, src, cx, cy, cw, ch, pad = 1.5) {
  if (!src || !src.startsWith('data:')) return;
  try {
    const props = doc.getImageProperties(src);
    const maxW = cw - pad * 2, maxH = ch - pad * 2;
    let iw = props.width, ih = props.height;
    const scale = Math.min(maxW / iw, maxH / ih, 1);
    iw *= scale; ih *= scale;
    const ox = cx + pad + (maxW - iw) / 2;
    const oy = cy + pad + (maxH - ih) / 2;
    const fmt = src.includes('image/png') ? 'PNG' : 'JPEG';
    doc.addImage(src, fmt, ox, oy, iw, ih);
  } catch (e) { /* skip silently */ }
}

// ── Color helpers ────────────────────────────────────────────

/** Returns [r,g,b] 0-255 from a '#RRGGBB' hex string. */
function hexToRgb(hex) {
  const n = parseInt(hex.replace('#',''), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

/**
 * Status cell fill color.
 * Pass → green, Fail → red, Doubt → orange, else → default (null).
 */
function statusColor(status) {
  const s = String(status || '').trim().toLowerCase();
  if (s === 'pass')  return hexToRgb('#C8F5C8');
  if (s === 'fail')  return hexToRgb('#FADBD8');
  if (s === 'doubt') return hexToRgb('#FDEBD0');
  return null;
}

/**
 * Out-of-range check for a sample value against min/max strings.
 * Returns true (highlight red) only when:
 *   - both min and max are present and numeric
 *   - the sample value is numeric
 *   - the value is strictly outside [min, max]
 * Any text value or missing range → returns false (no highlight).
 */
function isOutOfRange(value, min, max) {
  if (!min && !max) return false;           // no range defined
  const v   = parseFloat(String(value).trim());
  const mn  = parseFloat(String(min).trim());
  const mx  = parseFloat(String(max).trim());
  if (isNaN(v) || isNaN(mn) || isNaN(mx)) return false;  // text value or range
  return v < mn || v > mx;
}

// ── Main export (async — fetches images) ─────────────────────
async function generateQIR(data) {
  // Pre-fetch all images in parallel before drawing
  // NOTE: part_drawing intentionally excluded — handled as PDF page in mergePDFs.js
  console.log('  Pre-fetching images...');
  const [inspImage, logoImg] = await Promise.all([
    fetchImageAsDataUrl(data.insp_image),
    fetchImageAsDataUrl('https://res.cloudinary.com/dbwg6zz3l/image/upload/w_300,f_png,q_90/v1773643264/Black_Yellow_kq9kef.png'),
  ]);

  // Fetch dim row QC photos
  const dimPhotoMap = {};
  await Promise.all(
    (data.dimRows || []).map(async (r, i) => {
      if (r.qc_photo) {
        dimPhotoMap[i] = await fetchImageAsDataUrl(r.qc_photo);
      }
    })
  );

  // Fetch visual row photos
  const visPhotoMap = {};
  await Promise.all(
    (data.visRows || []).map(async (r, i) => {
      if (r.photo) {
        visPhotoMap[i] = await fetchImageAsDataUrl(r.photo);
      }
    })
  );

  console.log('  Images fetched. Building PDF...');

  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  let y = MT;

  // ── PAGE 1: HEADER ──────────────────────────────────────────
  doc.setDrawColor(...BORDER);
  doc.setLineWidth(0.3);
  doc.rect(ML, y, CW, 16);
  doc.line(ML + CW * 0.28, y, ML + CW * 0.28, y + 16);
  doc.line(ML + CW * 0.73, y, ML + CW * 0.73, y + 16);

  if (logoImg) {
    try {
      const logoProp = doc.getImageProperties(logoImg);
      const maxW = CW * 0.22, maxH = 14;
      const scale = Math.min(maxW / logoProp.width, maxH / logoProp.height);
      const lw = logoProp.width * scale, lh = logoProp.height * scale;
      doc.addImage(logoImg, 'JPEG', ML + (CW * 0.28 - lw) / 2, y + (16 - lh) / 2, lw, lh);
    } catch(e) {
      doc.setFontSize(7); doc.setTextColor(150, 150, 150);
      doc.text('[LOGO]', ML + CW * 0.14, y + 9, { align: 'center' });
    }
  } else {
    doc.setFontSize(7); doc.setTextColor(150, 150, 150);
    doc.text('[LOGO]', ML + CW * 0.14, y + 9, { align: 'center' });
  }

  doc.setFontSize(16);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(...DARK);
  doc.text('Quality Inspection Report', ML + CW * 0.505, y + 10, { align: 'center' });

  doc.setFontSize(8);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(60, 60, 60);
  doc.text(`Doc No: ${data.report_no}`,        ML + CW * 0.76, y + 5);
  doc.text(`Date:   ${data.submission_date}`,   ML + CW * 0.76, y + 10);
  doc.text(`By:     ${data.created_by || '—'}`, ML + CW * 0.76, y + 15);

  y += 20;
  doc.setDrawColor(...BORDER);
  doc.setLineWidth(0.3);

  // ── PART INFO TABLE — skip fields with empty values ──────────
  // Helper: returns trimmed string or null if empty
  const v = (val) => (val && String(val).trim()) ? String(val).trim() : null;

  // All possible fields in display order — only non-null ones are shown
  const fields = [
    ['Part Name',   v(data.part_name)   ],
    ['Part No.',    v(data.part_number) ],
    ['Customer',    v(data.customer)    ],
    ['Created By',  v(data.created_by)  ],
    ['Title',       v(data.title)       ],
    ['Item Code',   v(data.item_code)   ],
    ['RM Grade',    v(data.rm_grade)    ],
    ['Heat No.',    v(data.heat_no)     ],
    ['Qty',         v(data.order_qty)   ],
    ['Samples Checked', v(data.samples_checked)],
    ['Verified By', v(data.verified_by)],
  ].filter(([, val]) => val !== null);  // drop pairs where value is empty/null

  // Pack into rows of 3 pairs → 6 cells per row
  const infoRows = [];
  for (let i = 0; i < fields.length; i += 3) {
    const pair0 = fields[i]     || ['', ''];
    const pair1 = fields[i + 1] || ['', ''];
    const pair2 = fields[i + 2] || ['', ''];
    infoRows.push([pair0[0], pair0[1], pair1[0], pair1[1], pair2[0], pair2[1]]);
  }

  doc.autoTable({
    startY: y,
    margin: { left: ML, right: MR },
    tableWidth: CW,
    head: [],
    body: infoRows,
    styles: { fontSize: 8, cellPadding: 2.5, lineColor: BORDER, lineWidth: 0.3 },
    columnStyles: {
      0: { fontStyle: 'bold', fillColor: GRAY, cellWidth: CW * 0.09 },
      2: { fontStyle: 'bold', fillColor: GRAY, cellWidth: CW * 0.09 },
      4: { fontStyle: 'bold', fillColor: GRAY, cellWidth: CW * 0.09 },
    },
  });
  y = doc.lastAutoTable.finalY + 4;

  // Note / Remarks — rendered as a single line below the table, not inside it
  if (data.remarks && data.remarks.trim()) {
    doc.setFontSize(8);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(...DARK);
    doc.text('Note: ', ML, y);
    const noteLabelW = doc.getTextWidth('Note: ');
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(60, 60, 60);
    const noteLines = doc.splitTextToSize(data.remarks.trim(), CW - noteLabelW);
    doc.text(noteLines, ML + noteLabelW, y);
    y += (noteLines.length * 4.5) + 2;
  }

  // Conclusion on page 1 if present
  if (data.conclusion) {
    doc.setFontSize(9);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(...DARK);
    doc.text('Conclusion:', ML, y);
    doc.setFont('helvetica', 'normal');
    const lines = doc.splitTextToSize(data.conclusion, CW - 26);
    doc.text(lines, ML + 26, y);
  }

  // ── PAGE 2: DIMENSIONAL INSPECTION ───────────────────────────
  // NOTE: Part Drawing (page 3 in final PDF) is inserted by mergePDFs.js
  if (data.dimRows && data.dimRows.length > 0) {
    doc.addPage('a4', 'landscape');
    y = MT;
    y = sectionHeading(doc, 'Dimensional Inspection', y);

    const n = data.sampleCount || 5;
    const ROW_H = 14;

    // For each optional column: hide it only if ALL rows have a blank value for it.
    const allBlank = (rows, key) => rows.every(r => !r[key] || !String(r[key]).trim());

    const showSpec       = !allBlank(data.dimRows, 'specificat');
    const showInstrument = !allBlank(data.dimRows, 'instrument');
    const showStatus     = !allBlank(data.dimRows, 'status_1');
    const showDimComment = !allBlank(data.dimRows, 'comment');
    // Photo column: blank means no fetched image for that row
    const showDimPhoto   = data.dimRows.some((_, i) => !!dimPhotoMap[i]);

    // Fixed column widths — omitted cols contribute 0, freeing space for sample cols
    const FIXED_NO_W   = CW * 0.04;
    const FIXED_PAR_W  = CW * 0.11;
    const SPEC_W       = showSpec       ? CW * 0.115 : 0;
    const INSTR_W      = showInstrument ? CW * 0.10  : 0;
    const STATUS_W     = showStatus     ? CW * 0.065 : 0;
    const COMMENT_W    = showDimComment ? CW * 0.18 : 0;
    const PHOTO_W      = showDimPhoto   ? CW * 0.08  : 0;
    const FIXED_TOTAL  = FIXED_NO_W + FIXED_PAR_W + SPEC_W + INSTR_W + STATUS_W + COMMENT_W + PHOTO_W;
    const sColW        = (CW - FIXED_TOTAL) / n;

    // Build header row and body rows with only present columns
    const dimHead = ['No.', 'Parameter'];
    if (showSpec)       dimHead.push('Specification');
    if (showInstrument) dimHead.push('Instrument');
    dimHead.push(...Array.from({ length: n }, (_, i) => `${i + 1}`));
    if (showStatus)     dimHead.push('Status');
    if (showDimComment) dimHead.push('Comments');
    if (showDimPhoto)   dimHead.push('Photo');

    const dimBody = data.dimRows.map(r => {
      const row = [r.index, r.parameter];
      if (showSpec)       row.push(r.specificat   || '');
      if (showInstrument) row.push(r.instrument   || '');
      row.push(...r.samples.slice(0, n).concat(Array(Math.max(0, n - r.samples.length)).fill('')));
      if (showStatus)   row.push(r.status_1 || '');
      if (showDimComment) row.push(r.comment || '');
      if (showDimPhoto) row.push('');
      return row;
    });

    // Column index tracking — shifts as optional cols are added
    let ci = 2;
    const specIdx   = showSpec       ? ci++ : -1;
    const instrIdx  = showInstrument ? ci++ : -1;
    const sampleStart = ci; ci += n;
    const statusIdx = showStatus   ? ci++ : -1;
    const commentIdx = showDimComment ? ci++ : -1;
    const photoIdx  = showDimPhoto ? ci   : -1;

    const dimColStyles = {
      0: { cellWidth: FIXED_NO_W },
      1: { cellWidth: FIXED_PAR_W, halign: 'left' },
      ...Object.fromEntries(Array.from({ length: n }, (_, i) => [sampleStart + i, { cellWidth: sColW }])),
    };
    if (showSpec)       dimColStyles[specIdx]  = { cellWidth: SPEC_W };
    if (showInstrument) dimColStyles[instrIdx]  = { cellWidth: INSTR_W };
    if (showStatus)     dimColStyles[statusIdx] = { cellWidth: STATUS_W };
    if (showDimComment) dimColStyles[commentIdx] = { cellWidth: COMMENT_W, halign: 'left' };
    if (showDimPhoto)   dimColStyles[photoIdx]  = { cellWidth: PHOTO_W };

    doc.autoTable({
      startY: y,
      margin: { left: ML, right: MR },
      tableWidth: CW,
      head: [dimHead],
      body: dimBody,
      styles: {
        fontSize: 7.5, cellPadding: 2, lineColor: BORDER, lineWidth: 0.3,
        halign: 'center', valign: 'middle', minCellHeight: ROW_H,
      },
      headStyles: { fillColor: GRAY, textColor: DARK, fontStyle: 'bold', fontSize: 7 },
      columnStyles: dimColStyles,
      didParseCell: (d) => {
        if (d.section !== 'body') return;
        const raw = d.row.raw;

        // Status column — colour by Pass/Fail/Doubt
        if (showStatus && d.column.index === statusIdx) {
          const col = statusColor(raw[statusIdx]);
          if (col) d.cell.styles.fillColor = col;
        }

        // Sample columns — highlight red if value is out of min/max range
        if (d.column.index >= sampleStart && d.column.index < sampleStart + n) {
          const rowNo   = raw[0];
          const rowData = data.dimRows.find(r => r.index === rowNo);
          if (rowData) {
            const sampleVal = rowData.samples[d.column.index - sampleStart];
            if (isOutOfRange(sampleVal, rowData.min, rowData.max)) {
              d.cell.styles.fillColor = hexToRgb('#FADBD8');
            }
          }
        }
      },
      didDrawCell: (d) => {
        if (showDimPhoto && d.section === 'body' && d.column.index === photoIdx) {
          const img = dimPhotoMap[d.row.index];
          if (img) drawImgInCell(doc, img, d.cell.x, d.cell.y, d.cell.width, d.cell.height);
        }
      },
    });
  }

  // ── PAGE 3: VISUAL INSPECTION ─────────────────────────────────
  if (data.visRows && data.visRows.length > 0) {
    doc.addPage('a4', 'landscape');
    y = MT;
    y = sectionHeading(doc, 'Visual Inspection', y);

    const ROW_H = 14;

    // Hide each column only if ALL rows are blank for that field
    const allBlankV = (rows, key) => rows.every(r => !r[key] || !String(r[key]).trim());

    const showVisStatus   = !allBlankV(data.visRows, 'status');
    const showVisComments = !allBlankV(data.visRows, 'comments');
    const showVisPhoto    = data.visRows.some((_, i) => !!visPhotoMap[i]);

    const STATUS_W  = showVisStatus   ? CW * 0.10 : 0;
    const COMMENT_W = showVisComments ? CW - CW * 0.06 - CW * 0.22 - STATUS_W - (showVisPhoto ? CW * 0.09 : 0) : 0;
    const PHOTO_W   = showVisPhoto    ? CW * 0.09 : 0;

    const visHead = ['No.', 'Parameter'];
    if (showVisStatus)   visHead.push('Status');
    if (showVisComments) visHead.push('Comments');
    if (showVisPhoto)    visHead.push('Photo');

    const visBody = data.visRows.map(r => {
      const row = [r.index, r.parameter];
      if (showVisStatus)   row.push(r.status   || '');
      if (showVisComments) row.push(r.comments || '');
      if (showVisPhoto)    row.push('');
      return row;
    });

    let vi = 2;
    const visStatusIdx  = showVisStatus   ? vi++ : -1;
    const visCommentIdx = showVisComments ? vi++ : -1;
    const visPhotoIdx   = showVisPhoto    ? vi   : -1;

    const visColStyles = {
      0: { cellWidth: CW * 0.06 },
      1: { cellWidth: CW * 0.22, halign: 'left' },
    };
    if (showVisStatus)   visColStyles[visStatusIdx]  = { cellWidth: STATUS_W };
    if (showVisComments) visColStyles[visCommentIdx] = { cellWidth: COMMENT_W, halign: 'left' };
    if (showVisPhoto)    visColStyles[visPhotoIdx]   = { cellWidth: PHOTO_W };

    doc.autoTable({
      startY: y,
      margin: { left: ML, right: MR },
      tableWidth: CW,
      head: [visHead],
      body: visBody,
      styles: {
        fontSize: 8, cellPadding: 2.5, lineColor: BORDER, lineWidth: 0.3,
        halign: 'center', valign: 'middle', minCellHeight: ROW_H,
      },
      headStyles: { fillColor: GRAY, textColor: DARK, fontStyle: 'bold' },
      columnStyles: visColStyles,
      didParseCell: (d) => {
        if (d.section !== 'body') return;
        // Status column — colour by Pass/Fail/Doubt
        if (showVisStatus && d.column.index === visStatusIdx) {
          const col = statusColor(d.row.raw[visStatusIdx]);
          if (col) d.cell.styles.fillColor = col;
        }
      },
      didDrawCell: (d) => {
        if (showVisPhoto && d.section === 'body' && d.column.index === visPhotoIdx) {
          const img = visPhotoMap[d.row.index];
          if (img) drawImgInCell(doc, img, d.cell.x, d.cell.y, d.cell.width, d.cell.height);
        }
      },
    });

    y = doc.lastAutoTable.finalY + 4;

    if (inspImage) {
      if (y > PH - MB - 30) { doc.addPage('a4', 'landscape'); y = MT; }
      try {
        const props = doc.getImageProperties(inspImage);
        const maxW = CW * 0.5, maxH = PH - y - MB - 4;
        let iw = props.width, ih = props.height;
        const s = Math.min(maxW / iw, maxH / ih, 1);
        iw *= s; ih *= s;
        const fmt = inspImage.includes('image/png') ? 'PNG' : 'JPEG';
        doc.addImage(inspImage, fmt, ML + (CW - iw) / 2, y, iw, ih);
      } catch (e) { }
    }
  }

  return Buffer.from(doc.output('arraybuffer'));
}

module.exports = { generateQIR };
