/**
 * googleDrive.js
 * Uploads a PDF buffer to Google Drive using a service account.
 * Returns a shareable "anyone with link can view" URL.
 *
 * Env vars required:
 *   GOOGLE_SERVICE_ACCOUNT_JSON  — full service account JSON key (as string)
 *   GOOGLE_DRIVE_FOLDER_ID       — target folder ID in Drive
 */

const { google } = require('googleapis');
const { Readable } = require('stream');

function getAuthClient() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!raw) throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON env var not set');

  let credentials;
  try {
    credentials = JSON.parse(raw);
  } catch (e) {
    throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON: ' + e.message);
  }

  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      'https://www.googleapis.com/auth/drive.file',
    ],
  });
}

/**
 * Upload a PDF buffer to Google Drive.
 * @param {Buffer} buffer  — PDF file contents
 * @param {string} filename — desired filename in Drive
 * @returns {string} shareable Drive URL
 */
async function uploadToDrive(buffer, filename) {
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
  if (!folderId) throw new Error('GOOGLE_DRIVE_FOLDER_ID env var not set');

  const auth   = getAuthClient();
  const drive  = google.drive({ version: 'v3', auth });

  // Convert Buffer to a readable stream for the upload
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);

  console.log(`  Uploading "${filename}" to Drive folder ${folderId}...`);

  // Upload the file
  const uploadRes = await drive.files.create({
    requestBody: {
      name:    filename,
      parents: [folderId],
      mimeType: 'application/pdf',
    },
    media: {
      mimeType: 'application/pdf',
      body:     stream,
    },
    fields: 'id, name',
  });

  const fileId = uploadRes.data.id;
  console.log(`  Uploaded: fileId=${fileId}`);

  // Set permission — anyone with the link can view
  await drive.permissions.create({
    fileId,
    requestBody: {
      role: 'reader',
      type: 'anyone',
    },
  });

  const driveUrl = `https://drive.google.com/file/d/${fileId}/view`;
  console.log(`  Drive URL: ${driveUrl}`);
  return driveUrl;
}

module.exports = { uploadToDrive };
