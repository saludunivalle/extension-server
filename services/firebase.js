const { config } = require('dotenv');

config();

module.exports = Object.freeze({
  projectId: process.env.FIREBASE_PROJECT_ID,
  storageBucket: process.env.FIREBASE_STORAGE_BUCKET
});
