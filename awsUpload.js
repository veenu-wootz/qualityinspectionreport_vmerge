/**
 * awsUpload.js
 * Uploads a PDF buffer to AWS S3 and returns a permanent public URL.
 * Public access is granted via bucket policy (not ACL).
 *
 * Env vars required:
 *   AWS_ACCESS_KEY_ID      — IAM user access key
 *   AWS_SECRET_ACCESS_KEY  — IAM user secret key
 *   AWS_REGION             — e.g. ap-south-1
 *   AWS_S3_BUCKET_NAME     — target S3 bucket name
 */

const { S3Client, PutObjectCommand } = require('@aws-sdk/client-s3');

function getS3Client() {
  const region = process.env.AWS_REGION;
  if (!region) throw new Error('AWS_REGION env var not set');

  return new S3Client({
    region,
    credentials: {
      accessKeyId:     process.env.AWS_ACCESS_KEY_ID,
      secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
    },
  });
}

/**
 * Upload a PDF buffer to S3.
 * @param {Buffer} buffer   — PDF file contents
 * @param {string} filename — desired filename in S3
 * @returns {string} permanent public URL
 */
async function uploadToS3(buffer, filename) {
  const bucket = process.env.AWS_S3_BUCKET_NAME;
  const region = process.env.AWS_REGION;
  if (!bucket) throw new Error('AWS_S3_BUCKET_NAME env var not set');

  const client = getS3Client();

  console.log(`  Uploading "${filename}" to S3 bucket "${bucket}"...`);

  await client.send(new PutObjectCommand({
    Bucket:      bucket,
    Key:         filename,
    Body:        buffer,
    ContentType: 'application/pdf',
  }));

  const url = `https://${bucket}.s3.${region}.amazonaws.com/${encodeURIComponent(filename)}`;
  console.log(`  S3 URL: ${url}`);
  return url;
}

module.exports = { uploadToS3 };
