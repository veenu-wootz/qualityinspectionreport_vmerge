/**
 * sendEmail.js — QIR (AppSheet edition)
 * Sends merged PDF to exported_by address with bcc_email in BCC.
 *
 * Env vars:
 *   SMTP_USER        e.g. technology@wootz.work  OR  yourgmail@gmail.com
 *   SMTP_PASSWORD    App password
 *   SMTP_SERVICE     'gmail' | 'office365' (default: office365)
 *   INTERNAL_EMAILS  always-CC list, comma-separated (optional)
 */

const nodemailer = require('nodemailer');
let transporter;

function getTransporter() {
  if (!transporter) {
    const service = (process.env.SMTP_SERVICE || 'office365').toLowerCase();

    if (service === 'gmail') {
      transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASSWORD },
      });
    } else {
      // Office 365 / Microsoft 365
      transporter = nodemailer.createTransport({
        host: 'smtp.office365.com', port: 587, secure: false,
        auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASSWORD },
        tls: { ciphers: 'SSLv3' },
      });
    }
  }
  return transporter;
}

/**
 * @param {object} data       full parsed data object from server.js
 * @param {Buffer} pdfBuffer  merged final PDF
 * @param {string} filename
 */
async function sendQIREmail(data, pdfBuffer, filename) {
  const to  = data.your_email || data.exported_by || '';
  const bcc = [
    data.bcc_email || '',
    process.env.INTERNAL_EMAILS || '',
  ].join(',').split(',').map(e => e.trim()).filter(Boolean).join(', ');

  if (!to && !bcc) {
    console.warn('  sendEmail: no recipients — skipping');
    return;
  }

  const subject = `Inspection — ${data.title} | ${data.part_number} ${new Date().toLocaleString('en-GB', { timeZone: 'Asia/Kolkata', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false })}`;

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
      <div style="background:#1a1a2e;padding:20px 24px; border:1px solid #e0e0e0;border-bottom:none; border-radius:8px 8px 0 0;">
        <h2 style="color:#fff;margin:0;font-size:18px;">Quality Inspection Report</h2>
        <p style="color:#aaa;margin:4px 0 0;font-size:13px;">${data.report_no}</p>
      </div>
      <div style="background:#f9f9f9;padding:24px;border:1px solid #e0e0e0;border-top:none;">
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          <tr>
            <td style="padding:6px 0;color:#888;width:130px;">Title</td>
            <td style="padding:6px 0;font-weight:600;">${data.title || '—'}</td>
          </tr>
          <tr>
            <td style="padding:6px 0;color:#888;">Part Number</td>
            <td style="padding:6px 0;">${data.part_number || '—'}</td>
          </tr>
          <tr>
            <td style="padding:6px 0;color:#888;">Customer</td>
            <td style="padding:6px 0;">${data.customer || '—'}</td>
          </tr>
          <tr>
            <td style="padding:6px 0;color:#888;">Date</td>
            <td style="padding:6px 0;">${data.submission_date || '—'}</td>
          </tr>
          <tr>
            <td style="padding:6px 0;color:#888;">Prepared by</td>
            <td style="padding:6px 0;">${data.created_by || '—'}</td>
          </tr>
          <tr>
            <td style="padding:6px 0;color:#888;">Inspections</td>
            <td style="padding:6px 0;">${data.dimRows?.length || 0} dimensional,
              ${data.visRows?.length || 0} visual,
              ${data.certificates?.length || 0} doc/test report(s)</td>
          </tr>
        </table>
      </div>
      <div style="background:#fff;padding:16px 24px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px;">
        <p style="margin:0;font-size:13px;color:#555;">
          Compiled Quality Inspection Report with all the documents is attached.
        </p>
      </div>
    </div>`;

  const info = await getTransporter().sendMail({
    from:        `"Wootz.Checkin" <${process.env.SMTP_USER}>`,
    to:          to || undefined,
    bcc:         bcc || undefined,
    subject,
    text:        `QIR ${data.report_no} | ${data.part_name} | ${data.submission_date}\nFull report attached.`,
    html,
    attachments: [{ filename, content: pdfBuffer, contentType: 'application/pdf' }],
  });

  console.log(`  Email → to:${to} bcc:${bcc} (${info.messageId})`);
  return info;
}

module.exports = { sendQIREmail };
