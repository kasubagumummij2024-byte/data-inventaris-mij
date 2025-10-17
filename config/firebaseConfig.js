// config/firebaseConfig.js
const admin = require('firebase-admin');

let serviceAccount;

// Cek jika aplikasi berjalan di lingkungan server produksi (seperti Railway)
if (process.env.SERVICE_ACCOUNT_KEY_JSON) {
    // Ambil konfigurasi dari Environment Variable dan parse sebagai JSON
    serviceAccount = JSON.parse(process.env.SERVICE_ACCOUNT_KEY_JSON);
    console.log("Menggunakan service account dari Environment Variable.");
} else {
    // Jika tidak, gunakan file lokal (untuk development di komputer Anda)
    console.log("Menggunakan file serviceAccountKey.json lokal.");
    serviceAccount = require('./serviceAccountKey.json');
}

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

module.exports = db;

