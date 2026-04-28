/**
 * appsheetCheckin.js
 * Adds a row to the AppSheet 'Checkin' table via AppSheet API v2.
 * Runs after Samples edit, only when add_to_checkin is true.
 *
 * Env vars required:
 *   APPSHEET_ACCESS_KEY  — from AppSheet → My Account → Integrations → Access Key
 *   APPSHEET_APP_NAME    — already set (e.g. WootzCheckin2Ayush-832304561)
 */

const fetch = require('node-fetch');

const TABLE_NAME = 'CheckIn';

/**
 * Add one row to the Checkin table via AppSheet API.
 * @param {object} data  — parsed QIR payload
 * @param {string} s3Url — uploaded PDF S3 URL
 */
async function addToCheckin(data, s3Url) {
  const accessKey = process.env.APPSHEET_ACCESS_KEY;
  const appId     = process.env.APPSHEET_APP_NAME;

  if (!accessKey) throw new Error('APPSHEET_ACCESS_KEY env var not set');
  if (!appId)     throw new Error('APPSHEET_APP_NAME env var not set');

  const url = `https://api.appsheet.com/api/v2/apps/${appId}/tables/${TABLE_NAME}/Action`;

  // Tomorrow's date in MM/DD/YYYY format
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const reminderDate = `${tomorrow.getMonth() + 1}/${tomorrow.getDate()}/${tomorrow.getFullYear()}`;

  const body = {
    Action: 'Add',
    Properties: {
      Locale:   'en-US',
      Timezone: 'Asia/Kolkata',
    },
    Rows: [
      {
        'CheckIn ID':           data.report_no  || '',
        'ID':                   data.id         || '',   
        'Status':               'Update',
        'Description':          'Please find the inspection report attached.',
        'Created by':           data.your_email || '',
        'Timestamp':            data.timestamp  || '',
        'Files':                s3Url           || '',
        'Reminder_custom_date': reminderDate,
      },
    ],
  };

  console.log(`  Adding row to AppSheet table '${TABLE_NAME}' for CheckIn ID: ${data.report_no}`);

  const res = await fetch(url, {
    method:  'POST',
    headers: {
      'Content-Type':         'application/json',
      'ApplicationAccessKey': accessKey,
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`AppSheet API error ${res.status}: ${errText}`);
  }

  const result = await res.json();
  console.log(`  Checkin row added successfully`);
  console.log('  AppSheet response:', JSON.stringify(result));
  return result;
}

module.exports = { addToCheckin };
