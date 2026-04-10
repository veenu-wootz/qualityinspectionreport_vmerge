/**
 * server.js — QIR Server (AppSheet edition)
 *
 * POST /generate   Receives AppSheet webhook → generates QIR PDF
 *                  → merges test cert PDFs → emails result
 *
 * AppSheet sends:
 * {
 *   Sample: { sample_id, title, part_number, part_name, created_at,
 *             created_by, customer_name, inspection_map },
 *   Related_Inspection: [
 *     { inspection_type, parameter, min_req, max_req, instrument,
 *       sample_1..sample_10, status, qc_photo, comment, test_doc }
 *   ],
 *   exported_by: "user@email.com",
 *   bcc_email:   "a@b.com, c@d.com"
 * }
 */

require('dotenv').config();

const express          = require('express');
const { generateQIR }  = require('./generateQIR');
const { buildMergedPDF } = require('./mergePDFs');
const { sendQIREmail } = require('./sendEmail');
const { uploadToS3 } = require('./awsUpload');
const { addCheckinRow } = require('./appsheetRows');

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

// Log every POST for debugging
app.use((req, res, next) => {
  if (req.method === 'POST') {
    console.log('\nRAW PAYLOAD:', JSON.stringify(req.body, null, 2));
  }
  next();
});

app.get('/', (req, res) => {
  res.json({ status: 'ok', service: 'QIR Generator (AppSheet)', version: '2.0.0' });
});

// ── AppSheet URL builder ──────────────────────────────────────
// AppSheet stores file paths, not full URLs.
// We reconstruct the URL using the pattern AppSheet documents.
const APPSHEET_APP   = process.env.APPSHEET_APP_NAME || 'WootzCheckin2Ayush-832304561';
const APPSHEET_TABLE = 'Inspection';

function appsheetFileUrl(fileName) {
  if (!fileName || !fileName.trim()) return null;
  // If already a full URL, return as-is
  if (fileName.startsWith('http')) return fileName;
  return `https://www.appsheet.com/template/gettablefileurl?appName=${APPSHEET_APP}&tableName=${APPSHEET_TABLE}&fileName=${encodeURIComponent(fileName)}`;
}

// ── Parse AppSheet payload → internal format ──────────────────
function parsePayload(body) {
  const sample = body.Sample || {};
  const rows   = body.Related_Inspection || [];

  const dimRows  = [];
  const visRows  = [];
  const certDocs = [];  // { label, url }

  let dimIdx = 1, visIdx = 1;

  for (const row of rows) {
    const type = (row.inspection_type || '').toLowerCase().trim();

    if (type === 'dimension' || type === 'dimensional') {
      // Build samples array — only include non-empty values
      const samples = [];
      for (let i = 1; i <= 10; i++) {
        const v = row[`sample_${i}`];
        if (v !== undefined) samples.push(v ?? '');
      }
      // Trim trailing empty samples but keep at least 1
      while (samples.length > 1 && samples[samples.length - 1] === '') samples.pop();

      dimRows.push({
        index:      dimIdx++,
        parameter:  row.parameter   || '',
        specificat: `${row.min_req || ''}${row.min_req && row.max_req ? ' – ' : ''}${row.max_req || ''}`,
        instrument: row.instrument  || '',
        min:        row.min_req     || '',
        max:        row.max_req     || '',        
        samples,
        status_1:   row.status      || '',
        qc_photo:   appsheetFileUrl(row.qc_photo),
        comment:    row.comment     || '',
      });

    } else if (type === 'visual') {
      visRows.push({
        index:     visIdx++,
        parameter: row.parameter || '',
        status:    row.status    || '',
        comments:  row.comment   || '',
        photo:     appsheetFileUrl(row.qc_photo),
      });

    } else if (type === 'test' || type === 'certificate' || type === 'report' || type === 'attachment' ) {
      const docUrl = appsheetFileUrl(row.test_doc);
      if (docUrl) {
        certDocs.push({
          label: row.parameter || 'Certificate',
          url:   docUrl,
        });
      }
    }
  }

  // Determine sample count from longest dim row
  const sampleCount = dimRows.reduce((max, r) => Math.max(max, r.samples.length), 5);

  return {
    // Header fields
    report_no:       sample.sample_id    || `QIR-${Date.now()}`,
    submission_date: sample.created_at   || new Date().toISOString().split('T')[0],
    id:              sample.id           || '',
    part_name:       sample.part_name    || '',
    part_number:     sample.part_number  || '',
    project_pocs:    sample.project_pocs || '',
    customer:        sample.customer_name|| '',
    created_by:      sample.created_by   || '',
    created_by_email: sample.created_by_email || '',
    title:           sample.title        || '',
    rm_grade:        sample.rm_grade     || '',
    item_code:       sample.item_code    || '',
    heat_no:         sample.heat_no      || '',
    order_qty:       sample.qty          || '',
    samples_checked: sample.samples_checked || '',
    verified_by:     sample.verified_by  || 'Unverified',
    add_to_checkin:  sample.add_to_checkin === true || sample.add_to_checkin === 'true',
    remarks:         sample.remark       || '',
    timestamp:       sample.timestamp    || '',
    conclusion:      '',
    
    // Drawing / images
    part_drawing:    appsheetFileUrl(sample.inspection_map),
    insp_image:      null,

    // Inspection data
    sampleCount,
    dimRows,
    visRows,

    // Email
    your_email:      body.exported_by || '',
    bcc_email:       body.bcc_email   || '',

    // Cert PDFs — passed separately to buildMergedPDF
    certificates: certDocs,
  };
}

// ── Main endpoint ─────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  const startTime = Date.now();

  try {
    // ── RAW payload from AppSheet ──
    console.log('\n━━━━━━━━━━━━ RAW PAYLOAD ━━━━━━━━━━━━');
    console.log(JSON.stringify(req.body, null, 2));
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

    const data = parsePayload(req.body);

    // ── Parsed result ──
    console.log('━━━━━━━━━━━━ PARSED DATA ━━━━━━━━━━━━');
    console.log(`  report_no:       ${data.report_no}`);

    const filename = `Inspection Report-${data.title}-${data.timestamp}.pdf`;

    const s3FileUrlName = `${data.report_no}-${data.timestamp}.pdf`.replace(/[^a-zA-Z0-9\-_.]/g, '_');

    // 1. Generate QIR PDF (HTML → jsPDF)
    console.log('\n[1/3] Generating QIR PDF...');
    const qirBuffer = await generateQIR(data);
    console.log(`  ${(qirBuffer.length / 1024).toFixed(0)} KB`);

    // 2. Merge certificates
    console.log('\n[2/3] Merging certificates...');
    const mergedBuffer = await buildMergedPDF(qirBuffer, data.certificates, {
      reportNo:       data.report_no,
      partName:       data.part_name,
      date:           data.submission_date,
      partDrawingUrl: data.part_drawing,        // full URL — mergePDFs fetches & inserts p.1
      hasDrawing:     !!data.part_drawing,
      hasDim:         data.dimRows.length  > 0,
      hasVis:         data.visRows.length  > 0,
      verifiedBy:     data.verified_by,
    });
    console.log(`  Merged: ${(mergedBuffer.length / 1024).toFixed(0)} KB`);

    // 3. If verified — upload to Drive and append to Sheet
    let driveUrl = null;
    if (data.verified_by && data.verified_by !== 'Unverified' && data.add_to_checkin) {
      try {
        console.log('\n[3/5] Uploading to AWS_s3...');
        driveUrl = await uploadToS3(mergedBuffer, s3FileUrlName);
 
        console.log('\n[4/5] Appending to Appsheet...');
        await addCheckinRow(data, driveUrl);
      } catch (uploadErr) {
        // Non-fatal — log and continue to email
        console.error('  Appsheet error (non-fatal):', uploadErr.message);
      }
    }
    
    // 4. Send email
    console.log('\n[3/3] Sending email...');
    await sendQIREmail(data, mergedBuffer, filename);

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`\n✓ Done in ${elapsed}s — ${filename}\n`);

    res.json({ success: true, filename, elapsed: `${elapsed}s`, certs: data.certificates.length });

  } catch (err) {
    console.error('✗ Error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`\nQIR Server (AppSheet) on port ${PORT}\n`);
  if (!process.env.SMTP_USER)       console.warn('⚠  SMTP_USER not set');
  if (!process.env.SMTP_PASSWORD)   console.warn('⚠  SMTP_PASSWORD not set');
  if (!process.env.APPSHEET_APP_NAME) console.warn('⚠  APPSHEET_APP_NAME not set (using default)');
});
