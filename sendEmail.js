/**
 * sendEmail.js — for QIR Merge Server
 * Microsoft 365 SMTP. Sends merged PDF to submitter + internal team.
 *
 * Env vars:
 *   SMTP_USER        e.g. technology@wootz.work
 *   SMTP_PASSWORD    Microsoft account password
 *   INTERNAL_EMAILS  e.g. team@wootz.work,manager@wootz.work
 */

const nodemailer = require('nodemailer');
let transporter;

function getTransporter() {
  if (!transporter) {
    // transporter = nodemailer.createTransport({
    //   host: 'smtp.office365.com', port: 587, secure: false,
    //   auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASSWORD },
    //   tls: { ciphers: 'SSLv3' },
    // });
    transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASSWORD },
    });
  }
  return transporter;
}

/**
 * @param {object} meta      { reportNo, partName, date, yourEmail }
 * @param {Buffer} pdfBuffer  merged final PDF
 * @param {string} filename   e.g. "QIR-2025-0001-2025-01-15.pdf"
 */
async function sendQIREmail(meta, pdfBuffer, filename) {
  const { reportNo, partName, date, yourEmail } = meta;

  // Build recipient list
  const recipients = [];
  if (yourEmail?.trim()) recipients.push(yourEmail.trim());
  (process.env.INTERNAL_EMAILS || '').split(',').map(e => e.trim()).filter(Boolean).forEach(e => recipients.push(e));

  if (recipients.length === 0) {
    console.warn('  sendEmail: no recipients — skipping');
    return;
  }

  const to = [...new Set(recipients)].join(', ');

  const subject = `Quality Inspection Report — ${reportNo} | ${partName}`;

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
      <div style="background:#1a1a1a;padding:20px 24px;border-radius:8px 8px 0 0;">
        <h2 style="color:#fff;margin:0;font-size:18px;">Quality Inspection Report</h2>
        <p style="color:#aaa;margin:4px 0 0;font-size:13px;">${reportNo}</p>
      </div>
      <div style="background:#f9f9f9;padding:24px;border:1px solid #e0e0e0;border-top:none;">
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          <tr><td style="padding:6px 0;color:#888;width:120px;">Part Name</td><td style="padding:6px 0;font-weight:600;">${partName || '—'}</td></tr>
          <tr><td style="padding:6px 0;color:#888;">Date</td><td style="padding:6px 0;">${date || '—'}</td></tr>
          ${yourEmail ? `<tr><td style="padding:6px 0;color:#888;">Submitted by</td><td style="padding:6px 0;">${yourEmail}</td></tr>` : ''}
        </table>
      </div>
      <div style="background:#fff;padding:16px 24px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px;">
        <p style="margin:0;font-size:13px;color:#555;">The complete Quality Inspection Report with all certificates is attached.</p>
      </div>
    </div>`;

  const info = await getTransporter().sendMail({
    from: `"QIR System" <${process.env.SMTP_USER}>`,
    to, subject,
    text: `QIR ${reportNo} | ${partName} | ${date}\nSubmitted by: ${yourEmail || '—'}\n\nFull report attached.`,
    html,
    attachments: [{ filename, content: pdfBuffer, contentType: 'application/pdf' }],
  });

  console.log(`  Email → ${to} (${info.messageId})`);
  return info;
}

module.exports = { sendQIREmail };
