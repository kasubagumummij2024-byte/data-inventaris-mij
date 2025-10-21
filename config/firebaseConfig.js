// config/firebaseConfig.js
const admin = require('firebase-admin');

let serviceAccount;

// =======================================================
// KODE DEBUGGING DIMULAI DI SINI
// =======================================================

// BARIS DEBUG 1: Kita cek apakah variabelnya ada atau tidak sama sekali.
console.log('Mencoba membaca Environment Variable...');
console.log('Isi dari process.env.SERVICE_ACCOUNT_KEY_JSON adalah:', process.env.SERVICE_ACCOUNT_KEY_JSON);

// =======================================================

// Cek jika aplikasi berjalan di lingkungan server produksi (seperti Railway)
if (process.env.SERVICE_ACCOUNT_KEY_JSON) {
    console.log("Variabel Ditemukan! Mencoba mem-parse JSON...");
    try {
        serviceAccount = JSON.parse(process.env.SERVICE_ACCOUNT_KEY_JSON);
        console.log("Berhasil mem-parse JSON dari Environment Variable.");
    } catch (e) {
        console.error("GAGAL mem-parse JSON! Isi variabel kemungkinan besar rusak atau tidak lengkap.", e.message);
        // Paksa aplikasi crash agar kita tahu ada masalah serius
        throw new Error("Gagal mem-parse SERVICE_ACCOUNT_KEY_JSON.");
    }
} else {
    console.log("Variabel TIDAK Ditemukan. Mencoba menggunakan file lokal serviceAccountKey.json...");
    // Jika tidak, gunakan file lokal (untuk development di komputer Anda)
    serviceAccount = require('./serviceAccountKey.json');
}

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

module.exports = db;