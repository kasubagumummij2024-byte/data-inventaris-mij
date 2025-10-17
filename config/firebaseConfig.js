// config/firebaseConfig.js
const admin = require('firebase-admin');

// Pastikan file serviceAccountKey.json ada di dalam folder 'config' ini
const serviceAccount = require('./serviceAccountKey.json');

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

module.exports = db;