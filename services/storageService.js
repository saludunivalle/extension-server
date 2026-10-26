const axios = require('axios');
const fs = require('fs');
const path = require('path');
const { JWT } = require('google-auth-library');

const STORAGE_SCOPE = 'https://www.googleapis.com/auth/devstorage.read_write';

class StorageService {
  constructor() {
    this.bucket = String(process.env.FIREBASE_STORAGE_BUCKET || '').trim()
      .replace(/^gs:\/\//, '')
      .replace(/\/+$/, '');
    this.client = new JWT({
      email: process.env.GOOGLE_CLIENT_EMAIL,
      key: String(process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
      scopes: [STORAGE_SCOPE]
    });
  }

  assertConfigured() {
    if (!this.bucket || !process.env.GOOGLE_CLIENT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) {
      throw new Error('Faltan FIREBASE_STORAGE_BUCKET, GOOGLE_CLIENT_EMAIL o GOOGLE_PRIVATE_KEY');
    }
  }

  async uploadFile(file, objectName) {
    this.assertConfigured();
    if (!file || !file.path) throw new Error('No se proporcionó un archivo para almacenar');

    const safeObjectName = objectName.split('/').map(part => path.basename(part)).join('/');
    const token = await this.client.getAccessToken();
    const content = fs.createReadStream(file.path);
    const url = `https://storage.googleapis.com/upload/storage/v1/b/${encodeURIComponent(this.bucket)}/o`;

    await axios.post(url, content, {
      params: { uploadType: 'media', name: safeObjectName },
      headers: {
        Authorization: `Bearer ${token.token}`,
        'Content-Type': file.mimetype || 'application/octet-stream',
        'Content-Length': file.size
      },
      maxBodyLength: Infinity,
      maxContentLength: Infinity
    });

    return `https://storage.googleapis.com/${this.bucket}/${safeObjectName
      .split('/').map(encodeURIComponent).join('/')}`;
  }
}

module.exports = new StorageService();
