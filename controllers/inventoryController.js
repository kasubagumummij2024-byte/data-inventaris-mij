// controllers/inventoryController.js
const db = require('../config/firebaseConfig');
const exceljs = require('exceljs');
const QRCode = require('qrcode');

const inventarisCollection = db.collection('inventaris');
const countersCollection = db.collection('counters');

function getRomanMonth(monthNumber) {
    if (monthNumber < 1 || monthNumber > 12) { return 'XX'; }
    const romanMonths = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
    return romanMonths[monthNumber - 1];
}

// --- FUNGSI HELPER BARU UNTUK STANDARISASI ID LOKASI ---
function standardizeLokasi(lokasiStr) {
    // Jika lokasi kosong atau tidak valid, gunakan 'NOLOK'
    if (!lokasiStr || typeof lokasiStr !== 'string' || lokasiStr.trim() === '') {
        return 'NOLOK'; // 'No Location'
    }
    // Mengubah "P.01 / Lemari A" -> "P.01LEMARIA"
    // Hanya mempertahankan huruf, angka, dan titik.
    return lokasiStr
        .toUpperCase()
        .replace(/[^A-Z0-9\.]/g, '') // Hapus semua karakter kecuali huruf, angka, dan titik
        .trim();
}
// --- AKHIR FUNGSI HELPER BARU ---

// =======================================================
// DATA MASTER LENGKAP BARU (Berdasarkan CSV Anda - FINAL)
// =======================================================
const dataMasterLengkap = {
  'Aset Kantor & Furnitur': {
    'Meja': { kode: 'AKF-01', contoh: 'Meja kerja, meja rapat, meja guru, meja komputer, meja resepsionis' },
    'Kursi': { kode: 'AKF-02', contoh: 'Kursi kerja, kursi tamu, kursi lipat, kursi tunggu, kursi siswa' },
    'Lemari (Kayu/Besi)': { kode: 'AKF-03', contoh: 'Lemari arsip, lemari pakaian, lemari dokumen, filling cabinet' },
    'Etalase / Loker': { kode: 'AKF-04', contoh: 'Etalase kaca, loker siswa, loker staf, rak display' },
    'Sofa': { kode: 'AKF-05', contoh: 'Sofa tamu, sofa ruang kepala, sofa lobi' },
    'Rak / Laci / Papan Tulis': { kode: 'AKF-06', contoh: 'Rak buku, rak dokumen, laci meja, papan tulis putih, papan pengumuman' },
    'Lainnya (Aset Kantor)': { kode: 'AKF-99', contoh: 'Partisi ruangan, karpet kantor, meja altar, backdrop kayu' }
  },
  'Perangkat Elektronik & IT': {
    'Komputer / Laptop': { kode: 'EIT-01', contoh: 'PC Desktop, Laptop, Notebook, All-in-One PC' },
    'Monitor': { kode: 'EIT-02', contoh: 'Monitor LED, Monitor LCD' },
    'Printer / Scanner': { kode: 'EIT-03', contoh: 'Printer Inkjet, Printer Laser, Scanner Flatbed, Printer Dot Matrix' },
    'Proyektor & Layar': { kode: 'EIT-04', contoh: 'Proyektor LCD, Proyektor DLP, Layar Proyektor Tripod, Layar Gantung' },
    'Perangkat Jaringan': { kode: 'EIT-05', contoh: 'Router, Switch Hub, Access Point Wifi, Modem, Kabel LAN' },
    'Server & Penyimpanan': { kode: 'EIT-06', contoh: 'Server Rackmount, Server Tower, NAS (Network Attached Storage)' },
    'Perangkat Audio Visual': { kode: 'EIT-07', contoh: 'Speaker Aktif, Sound System Portable, Mixer Audio, Mikrofon, TV LED' },
    'Perangkat Komunikasi': { kode: 'EIT-08', contoh: 'Telepon PABX, Mesin Fax, Walkie-Talkie' },
    'UPS & Power': { kode: 'EIT-09', contoh: 'UPS, Stabilizer (Stavolt), Power Strip (Stop Kontak)' },
    'Aksesoris & Lainnya (Elektronik)': { kode: 'EIT-99', contoh: 'Keyboard, Mouse, Webcam, Hard Disk Eksternal, Flashdisk' }
  },
  'Alat Tulis Kantor (ATK) & Habis Pakai': {
    'Kertas & Produk Kertas': { kode: 'ATK-01', contoh: 'Kertas HVS (A4, F4, dll.), Amplop, Kertas Foto, Sticky Notes, Buku Tulis' },
    'Alat Tulis': { kode: 'ATK-02', contoh: 'Pulpen, Pensil, Spidol, Stabilo, Penghapus, Tipe-X, Rautan' },
    'Perlengkapan Meja & Arsip': { kode: 'ATK-03', contoh: 'Stapler & Isi, Perforator, Gunting, Cutter, Map, Ordner, Klip, Lem' },
    'Tinta & Toner': { kode: 'ATK-04', contoh: 'Tinta Printer (Botol/Cartridge), Toner Laser' },
    'Baterai': { kode: 'ATK-05', contoh: 'Baterai AA, Baterai AAA, Baterai Kancing' },
    'Lainnya (ATK)': { kode: 'ATK-99', contoh: 'Materai, Stempel, Bak Stempel, Kalkulator' }
  },
  'Perlengkapan Operasional': {
      'Mesin & Peralatan Berat': { kode: 'OPS-01', contoh: 'Mesin fotokopi, mesin jilid, mesin potong kertas, genset' },
      'Peralatan Tangan (Tools)': { kode: 'OPS-02', contoh: 'Bor, gerinda, obeng set, kunci pas set, palu, tang' },
      'Alat Ukur & Pengujian': { kode: 'OPS-03', contoh: 'Multimeter, jangka sorong, timbangan, meteran' },
      'Alat Pelindung Diri (APD)': { kode: 'OPS-04', contoh: 'Helm safety, sarung tangan, kacamata pelindung, sepatu safety, masker' },
      'Tangga & Perancah': { kode: 'OPS-05', contoh: 'Tangga lipat aluminium, tangga multifungsi, scaffolding (jika ada)' },
      'Perlengkapan Pemadam Api': { kode: 'OPS-06', contoh: 'APAR (Tabung Pemadam Api Ringan), Hydrant (jika ada)' },
      'Lainnya (Operasional)': { kode: 'OPS-99', contoh: 'Genset portable, pompa air, troli barang' }
  },
    'Aset Kendaraan': {
    'Kendaraan Roda Empat': { kode: 'KDN-01', contoh: 'Mobil Operasional, Mobil Kepala Sekolah, Minibus Sekolah' },
    'Kendaraan Roda Dua': { kode: 'KDN-02', contoh: 'Motor Dinas' },
    'Kendaraan Khusus': { kode: 'KDN-03', contoh: 'Forklift, Gerobak Dorong' },
    'Aksesoris & Suku Cadang': { kode: 'KDN-04', contoh: 'Ban, Aki, Oli, Dongkrak, Helm' },
    'Lainnya (Kendaraan)': { kode: 'KDN-99', contoh: 'Sepeda (jika ada)' }
  },
  'Perlengkapan Kebersihan & Maintenance': {
    'Alat Kebersihan Manual': { kode: 'KMT-01', contoh: 'Sapu, Pel, Pengki, Sikat, Kemoceng, Wiper Kaca' },
    'Bahan Pembersih': { kode: 'KMT-02', contoh: 'Cairan pembersih, sabun, disinfektan, pewangi ruangan' },
    'Peralatan Kebersihan Khusus': { kode: 'KMT-03', contoh: 'Vacuum cleaner, mesin polisher, mesin penyedot debu' },
    'Tempat Sampah & Aksesori': { kode: 'KMT-04', contoh: 'Sulo, tong sampah, tempat sampah stainless, kantong plastik' },
    'Lainnya (Kebersihan)': { kode: 'KMT-99', contoh: 'Ember, gayung, rak alat kebersihan' }
  },
  'Perlengkapan Pantry & Dapur': {
    'Peralatan Makan & Masak': { kode: 'PAN-01', contoh: 'Piring, mangkuk, sendok, garpu, wajan, panci, pisau' },
    'Perlengkapan Saji': { kode: 'PAN-02', contoh: 'Nampan, tudung saji, troli saji, teko saji' },
    'Wadah Penyimpanan': { kode: 'PAN-03', contoh: 'Toples, wadah kerupuk, termos nasi, kontainer makanan' },
    'Lainnya (Pantry)': { kode: 'PAN-99', contoh: 'Teko air, talenan, serbet dapur' }
  },
  'Perlengkapan Ekstrakurikuler & Laboratorium': {
    'Ekstrakurikuler': { kode: 'EKS-01', contoh: 'Perlengkapan pramuka, drum band, bola futsal, net voli, alat musik' },
    'Laboratorium IPA': { kode: 'EKS-02', contoh: 'Mikroskop, tabung reaksi, alat peraga fisika/biologi/kimia' },
    'Laboratorium Komputer': { kode: 'EKS-03', contoh: 'Komputer lab, Jaringan lab, Meja lab komputer' },
    'Laboratorium Bahasa': { kode: 'EKS-04', contoh: 'Headset lab bahasa, master control lab bahasa' },
    'Lainnya (Ekskul/Lab)': { kode: 'EKS-99', contoh: 'Perlengkapan UKS, Alat Peraga Matematika' }
  },
  'Lain-lain': {
    'Lain-lain': { kode: 'LLN-99', contoh: 'Barang yang tidak termasuk kategori di atas' }
  }
};

const kodeSumberAnggaranMap = { 'BOS KB': 'BOSKB', 'BOS RA': 'BOSRA', 'BOS MI': 'BOSMI', 'BOS MTs': 'BOSMTS', 'BOS MA': 'BOSMA','BOP KB': 'BOPKB', 'BOP RA': 'BOPRA','KAS KB': 'KASKB', 'KAS RA': 'KASRA', 'KAS MI': 'KASMI', 'KAS MTS': 'KASMTS', 'KAS MA': 'KASMA','HIBAH': 'HBH', 'SPONSORSHIP': 'SPS', 'MANDIRI': 'MDR', 'LAIN-LAIN': 'LLN' };

const dropdownOptions = {
    satuan: ['Unit', 'Pcs', 'Set', 'Lusin', 'Box', 'Roll', 'Kg', 'Liter'],
    statusKondisi: ['Baik', 'Perlu Perbaikan', 'Rusak', 'Dalam Proses Perbaikan', 'Habis'],
    warna: ['Hitam', 'Putih', 'Abu-abu', 'Silver', 'Merah', 'Biru', 'Hijau', 'Kuning', 'Coklat', 'Oranye','Ungu','Pink', 'Lainnya'],
    sumberAnggaran: Object.keys(kodeSumberAnggaranMap),
    statusPenghapusan: ['Masih Digunakan', 'Dibuang', 'Dihibahkan', 'Dilelang']
};
// =======================================================

exports.getAllItems = async (req, res) => {
    try {
        const { search, uploadStatus, message, count } = req.query;
        let query = inventarisCollection.orderBy('nomorInventaris', 'desc');
        const snapshot = await query.get();
        let items = [];
        snapshot.forEach(doc => items.push({ id: doc.id, ...doc.data() }));

        if (search) {
            items = items.filter(item =>
                Object.values(item).some(value =>
                    String(value).toLowerCase().includes(search.toLowerCase())
                )
            );
        }
        res.render('index', { items, search: search || '', uploadStatus, message, count, options: dropdownOptions });
    } catch (error) {
        res.status(500).send(error.message);
    }
};

exports.getAddItemForm = (req, res) => {
    res.render('form-tambah', { dataMasterLengkap, options: dropdownOptions });
};

exports.createItem = async (req, res) => {
    try {
        const counterRef = countersCollection.doc('inventoryCounter');
        let newNomorUrut;
        await db.runTransaction(async (t) => {
            const counterDoc = await t.get(counterRef);
            newNomorUrut = (counterDoc.data()?.lastNumber || 0) + 1;
            t.set(counterRef, { lastNumber: newNomorUrut });
        });

        const { kategori, subKategori, sumberAnggaran } = req.body;
        const now = new Date();
        const tahun = now.getFullYear();
        const bulanAngka = now.getMonth() + 1;
        const bulanRomawi = getRomanMonth(bulanAngka);
        const nomorUrutPadded = String(newNomorUrut).padStart(4, '0');
        const kodeAnggaran = kodeSumberAnggaranMap[sumberAnggaran] || 'ERR';
        const kodeSubKategori = dataMasterLengkap[kategori]?.[subKategori]?.kode || 'ERR-SUB';

        // --- REVISI ID BARU ---
        // Ambil noPintuLokasi dari form, standarisasi, dan masukkan ke ID
        const kodeNoPintu = standardizeLokasi(req.body.noPintuLokasi);
        const nomorInventaris = `${nomorUrutPadded}/${kodeNoPintu}/${kodeSubKategori}/${kodeAnggaran}/INV-MIJ/${bulanRomawi}/${tahun}`;
        // --- AKHIR REVISI ---

        const newItem = {
            namaBarang: req.body.namaBarang,
            kategori: req.body.kategori,
            subKategori: req.body.subKategori,
            warna: req.body.warna,
            sumberAnggaran: req.body.sumberAnggaran,
            tahunPerolehan: parseInt(req.body.tahunPerolehan),
            jumlah: parseInt(req.body.jumlah),
            satuan: req.body.satuan,
            nilaiPerolehan: parseFloat(req.body.nilaiPerolehan),
            lokasiFisik: req.body.lokasiFisik,
            noPintuLokasi: req.body.noPintuLokasi || null,
            penanggungJawab: req.body.penanggungJawab || null, // DITAMBAHKAN
            statusKondisi: req.body.statusKondisi,
            
            jumlah_awal: parseInt(req.body.jumlah),
            lokasiFisik_awal: req.body.lokasiFisik,
            noPintuLokasi_awal: req.body.noPintuLokasi || null, // Ini "kunci" ID kita
            penanggungJawab_awal: req.body.penanggungJawab || null, // DITAMBAHKAN
            statusKondisi_awal: req.body.statusKondisi,
            
            statusPenghapusan: 'Masih Digunakan',
            dasarPenghapusan: null,
            tanggalPenghapusan: null,
            
            nomorInventaris, // Menggunakan ID baru yang sudah direvisi
            createdAt: now,
            createdBy: req.user.email,
            updatedAt: now,
            updatedBy: req.user.email,
        };
        
        await inventarisCollection.add(newItem);
        res.redirect('/');
    } catch (error) {
        res.status(500).send(error.message);
    }
};

exports.getEditItemForm = async (req, res) => {
    try {
        const doc = await inventarisCollection.doc(req.params.id).get();
        if (!doc.exists) return res.status(404).send('Barang tidak ditemukan');
        
        let itemData = doc.data();
        if (itemData.tanggalPenghapusan && itemData.tanggalPenghapusan.toDate) {
             const dt = itemData.tanggalPenghapusan.toDate();
             itemData.tanggalPenghapusanFormatted = dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0') + '-' + String(dt.getDate()).padStart(2, '0');
        } else {
             itemData.tanggalPenghapusanFormatted = '';
        }

        res.render('form-edit', {
            item: { id: doc.id, ...itemData },
            dataMasterLengkap,
            options: dropdownOptions
        });
    } catch (error) {
        res.status(500).send(error.message);
    }
};

exports.updateItem = async (req, res) => {
    try {
        const { nomorInventaris, ...restOfBody } = req.body;
        
        let tanggalPenghapusanValue = null;
        if (req.body.tanggalPenghapusan) {
             tanggalPenghapusanValue = new Date(req.body.tanggalPenghapusan);
             if (isNaN(tanggalPenghapusanValue.getTime())) {
                  tanggalPenghapusanValue = null;
             }
        }

        const updatedItem = {
            ...restOfBody,
            tahunPerolehan: parseInt(req.body.tahunPerolehan),
            jumlah: parseInt(req.body.jumlah),
            nilaiPerolehan: parseFloat(req.body.nilaiPerolehan),
            penanggungJawab: req.body.penanggungJawab || null, // DITAMBAHKAN
            noPintuLokasi: req.body.noPintuLokasi || null,
            statusPenghapusan: req.body.statusPenghapusan,
            dasarPenghapusan: req.body.dasarPenghapusan || null,
            tanggalPenghapusan: tanggalPenghapusanValue,
            updatedAt: new Date(),
            updatedBy: req.user.email
        };
        
        delete updatedItem.tanggalPenghapusanFormatted;
        // Kita tidak mengubah 'nomorInventaris' saat update
        await inventarisCollection.doc(req.params.id).update(updatedItem);
        res.redirect('/');
    } catch (error) {
        res.status(500).send(error.message);
    }
};

exports.deleteItem = async (req, res) => {
     try {
        await inventarisCollection.doc(req.params.id).delete();
        res.redirect('/');
    } catch (error) {
        res.status(500).send(error.message);
    }
};

exports.getItemDetail = async (req, res) => {
     try {
        const doc = await inventarisCollection.doc(req.params.id).get();
        if (!doc.exists) return res.status(404).send('Barang tidak ditemukan');
        const itemData = { id: doc.id, ...doc.data() };
        const url = `${req.protocol}://${req.get('host')}/barang/${itemData.id}`;
        const qrCodeDataUrl = await QRCode.toDataURL(url);
        res.render('detail', { item: itemData, qrCodeDataUrl });
    } catch (error) {
        res.status(500).send(error.message);
    }
};

exports.getReferensiKodePage = (req, res) => {
    const referensiData = Object.entries(dataMasterLengkap).map(([kategori, subkategorisObject]) => ({
        namaKategori: kategori,
        subkategoris: Object.entries(subkategorisObject).map(([namaSub, dataSub]) => ({
            namaSubKategori: namaSub,
            kodeSubKategori: dataSub.kode,
            contohBarang: dataSub.contoh
        }))
    }));
    res.render('referensi-kode', { referensiData });
};

exports.downloadReferensiExcel = async (req, res) => {
    try {
        const workbook = new exceljs.Workbook();
        const worksheet = workbook.addWorksheet('Referensi Kode Inventaris');

        worksheet.columns = [
            { header: 'Kategori', key: 'kategori', width: 40 },
            { header: 'Sub Kategori', key: 'subKategori', width: 40 },
            { header: 'Kode Sub Kategori', key: 'kode', width: 20 },
            { header: 'Contoh Barang', key: 'contoh', width: 60 }
        ];

        worksheet.getRow(1).font = { bold: true };
        worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

        Object.entries(dataMasterLengkap).forEach(([kategori, subKategorisObject]) => {
            Object.entries(subKategorisObject).forEach(([namaSub, dataSub]) => {
                worksheet.addRow({
                    kategori: kategori,
                    subKategori: namaSub,
                    kode: dataSub.kode,
                    contoh: dataSub.contoh
                });
            });
        });

        worksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {
            if (rowNumber > 1) {
                row.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
            }
        });

        res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition','attachment; filename=' + 'Referensi_Kode_Inventaris_MIJ.xlsx');

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('Gagal membuat file Excel referensi:', error);
        res.status(500).send("Terjadi kesalahan saat membuat file Excel referensi.");
    }
};

// =======================================================
// FUNGSI LAPORAN LENGKAP
// =======================================================
exports.downloadExcel = async (req, res) => {
     try {
        const snapshot = await inventarisCollection.orderBy('createdAt', 'asc').get();
        let items = [];
        snapshot.forEach(doc => items.push({ id: doc.id, ...doc.data() }));

        const workbook = new exceljs.Workbook();
        const worksheet = workbook.addWorksheet('Data Inventaris Lengkap');

        worksheet.columns = [
            { header: 'QR Code', key: 'qr', width: 15 },
            { header: 'Nomor Inventaris', key: 'nomorInventaris', width: 30 },
            { header: 'Nama Barang', key: 'namaBarang', width: 30 },
            { header: 'Warna', key: 'warna', width: 15 },
            { header: 'Sumber Anggaran', key: 'sumberAnggaran', width: 20 },
            { header: 'Tahun Perolehan', key: 'tahunPerolehan', width: 15 },
            { header: 'Kategori', key: 'kategori', width: 25 },
            { header: 'Sub Kategori', key: 'subKategori', width: 20 },
            { header: 'Kondisi Awal', key: 'statusKondisi_awal', width: 20 },
            { header: 'Kondisi Terkini', key: 'statusKondisi', width: 20 },
            { header: 'Lokasi Awal', key: 'lokasiFisik_awal', width: 30 },
            { header: 'Lokasi Terkini', key: 'lokasiFisik', width: 30 },
            { header: 'No. Pintu Awal', key: 'noPintuLokasi_awal', width: 15 },
            { header: 'Penanggung Jawab Awal', key: 'penanggungJawab_awal', width: 25 }, // DITAMBAHKAN
            { header: 'No. Pintu Terkini', key: 'noPintuLokasi', width: 15 },
            { header: 'Penanggung Jawab Terkini', key: 'penanggungJawab', width: 25 }, // DITAMBAHKAN
            { header: 'Jumlah Awal', key: 'jumlah_awal', width: 15 },
            { header: 'Jumlah Terkini', key: 'jumlah', width: 15 },
            { header: 'Satuan', key: 'satuan', width: 10 },
            { header: 'Nilai Perolehan (Rp)', key: 'nilaiPerolehan', width: 20, style: { numFmt: '"Rp"#,##0' } },
            { header: 'Tanggal Input Awal', key: 'createdAt', width: 20 },
            { header: 'Diinput Oleh', key: 'createdBy', width: 25 },
            { header: 'Tanggal Update Terakhir', key: 'updatedAt', width: 20 },
            { header: 'Diupdate Oleh', key: 'updatedBy', width: 25 },
            { header: 'Status Penghapusan', key: 'statusPenghapusan', width: 20 },
            { header: 'Dasar Penghapusan', key: 'dasarPenghapusan', width: 30 },
            { header: 'Tanggal Penghapusan', key: 'tanggalPenghapusan', width: 20 },
        ];
        
        worksheet.getRow(1).font = { bold: true };

        let rowNumber = 2;
        for (const item of items) {
            let tanggalPenghapusanFormatted = '';
            if (item.tanggalPenghapusan && item.tanggalPenghapusan.toDate) {
                tanggalPenghapusanFormatted = item.tanggalPenghapusan.toDate().toLocaleDateString('id-ID');
            }

            const row = worksheet.addRow({
                ...item,
                createdAt: item.createdAt.toDate ? item.createdAt.toDate().toLocaleString('id-ID') : '',
                updatedAt: item.updatedAt.toDate ? item.updatedAt.toDate().toLocaleString('id-ID') : '',
                updatedBy: item.updatedBy || '',
                tanggalPenghapusan: tanggalPenghapusanFormatted,
                noPintuLokasi_awal: item.noPintuLokasi_awal || '',
                penanggungJawab_awal: item.penanggungJawab_awal || '', // DITAMBAHKAN
                noPintuLokasi: item.noPintuLokasi || '',
                penanggungJawab: item.penanggungJawab || '', // DITAMBAHKAN
            });

            row.height = 80;
            row.eachCell({ includeEmpty: false }, cell => {
                 cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
            });
            row.getCell('nilaiPerolehan').alignment = { vertical: 'middle', horizontal: 'right' };
            row.getCell('qr').alignment = { vertical: 'middle', horizontal: 'center' };


            const url = `${req.protocol}://${req.get('host')}/barang/${item.id}`;
            const qrBuffer = await QRCode.toBuffer(url, { type: 'png', width: 100, margin: 1 });
            const imageId = workbook.addImage({ buffer: qrBuffer, extension: 'png' });
            worksheet.addImage(imageId, {
                tl: { col: 0.1, row: rowNumber - 0.9 },
                ext: { width: 100, height: 100 }
            });

            rowNumber++;
        }

        res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition','attachment; filename=' + 'Laporan_Inventaris_Lengkap_MIJ.xlsx');
        
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error("Gagal men-generate Excel lengkap:", error);
        res.status(500).send(error.message);
    }
};

// =======================================================
// FUNGSI BARU UNTUK CETAK LABEL
// =======================================================
exports.downloadLabelSheet = async (req, res) => {
    try {
        const snapshot = await inventarisCollection.orderBy('createdAt', 'asc').get();
        let items = [];
        snapshot.forEach(doc => items.push({ id: doc.id, ...doc.data() }));

        const workbook = new exceljs.Workbook();
        const worksheet = workbook.addWorksheet('Label QR Code Inventaris');

        worksheet.columns = [
            { header: 'QR Code', key: 'qr', width: 15 },
            { header: 'Nomor Inventaris', key: 'nomorInventaris', width: 35 },
            { header: 'Nama Barang', key: 'namaBarang', width: 30 },
            { header: 'Penanggung Jawab', key: 'penanggungJawab', width: 30 }, // DITAMBAHKAN
        ];
        
        worksheet.getRow(1).font = { bold: true };

        let rowNumber = 2;
        for (const item of items) {
            const row = worksheet.addRow({
                nomorInventaris: item.nomorInventaris,
                namaBarang: item.namaBarang,
                penanggungJawab: item.penanggungJawab || '', // DITAMBAHKAN
            });

            row.height = 80;
            row.getCell('B').alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }; 
            row.getCell('C').alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }; 
            row.getCell('D').alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }; // DITAMBAHKAN

            const url = `${req.protocol}://${req.get('host')}/barang/${item.id}`;
            const qrBuffer = await QRCode.toBuffer(url, { type: 'png', width: 100, margin: 1 });
            const imageId = workbook.addImage({ buffer: qrBuffer, extension: 'png' });
            worksheet.addImage(imageId, {
                tl: { col: 0.1, row: rowNumber - 0.9 },
                ext: { width: 100, height: 100 }
            });

            rowNumber++;
        }

        res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition','attachment; filename=' + 'Label_QRCode_Inventaris_MIJ.xlsx');
        
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error("Gagal men-generate Excel Label QR:", error);
        res.status(500).send(error.message);
    }
};


exports.downloadTemplate = async (req, res) => {
    try {
        const workbook = new exceljs.Workbook();
        const petunjukSheet = workbook.addWorksheet('Petunjuk');
        petunjukSheet.getColumn('A').width = 80; petunjukSheet.getCell('A1').value = 'PETUNJUK PENGISIAN TEMPLATE INVENTARIS MIJ'; petunjukSheet.getCell('A1').font = { bold: true, size: 16 }; petunjukSheet.getCell('A3').value = '1. Jangan mengubah, menghapus, atau menambah kolom (header) di sheet "Data Inventaris".'; petunjukSheet.getCell('A4').value = '2. Isi data mulai dari baris kedua di sheet "Data Inventaris".'; petunjukSheet.getCell('A5').value = '3. Kolom Kategori, Satuan, Status Kondisi, Warna, Sumber Anggaran WAJIB dipilih dari dropdown.'; petunjukSheet.getCell('A6').value = '4. Kolom Sub Kategori WAJIB diisi manual sesuai Kategori yang dipilih (lihat WebApp/Referensi).'; petunjukSheet.getCell('A7').value = '5. Kolom Nilai Perolehan dan Tahun Perolehan, masukkan angka saja (contoh: 5000000 / 2024).'; petunjukSheet.getCell('A8').value = '6. Kolom terkait Penghapusan dikosongkan saat input awal.'; petunjukSheet.getCell('A9').value = '7. No. Pintu Lokasi diisi jika ada, jika tidak kosongkan.'; petunjukSheet.getCell('A10').value = '8. Setelah selesai, simpan file ini dan unggah melalui WebApp Inventaris.';

        const dataSheet = workbook.addWorksheet('Data Inventaris');
        
        dataSheet.columns = [
            { header: 'Nama Barang', key: 'namaBarang', width: 40 },
            { header: 'Penanggung Jawab', key: 'penanggungJawab', width: 30 }, // DITAMBAHKAN
            { header: 'Kategori', key: 'kategori', width: 30 },
            { header: 'Sub Kategori', key: 'subKategori', width: 25 },
            { header: 'Warna', key: 'warna', width: 20 },
            { header: 'Sumber Anggaran', key: 'sumberAnggaran', width: 25 },
            { header: 'Tahun Perolehan', key: 'tahunPerolehan', width: 15 },
            { header: 'Jumlah', key: 'jumlah', width: 15 },
            { header: 'Satuan', key: 'satuan', width: 15 },
            { header: 'Nilai Perolehan (Rp)', key: 'nilaiPerolehan', width: 25 },
            { header: 'Lokasi Fisik', key: 'lokasiFisik', width: 40 },
            { header: 'No. Pintu Lokasi', key: 'noPintuLokasi', width: 20 },
            { header: 'Status Kondisi', key: 'statusKondisi', width: 25 },
            { header: 'Status Penghapusan', key: 'statusPenghapusan', width: 20 },
            { header: 'Dasar Penghapusan', key: 'dasarPenghapusan', width: 30 },
            { header: 'Tanggal Penghapusan (YYYY-MM-DD)', key: 'tanggalPenghapusan', width: 25 },
        ];
        dataSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        dataSheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF203764' } };
        
        const lastRow = 1001;
        // Validasi digeser 1 kolom
        dataSheet.dataValidations.add(`C2:C${lastRow}`, { type: 'list', allowBlank: false, formulae: [`"${Object.keys(dataMasterLengkap).join(',')}"`] });
        dataSheet.dataValidations.add(`E2:E${lastRow}`, { type: 'list', allowBlank: false, formulae: [`"${dropdownOptions.warna.join(',')}"`] });
        dataSheet.dataValidations.add(`F2:F${lastRow}`, { type: 'list', allowBlank: false, formulae: [`"${dropdownOptions.sumberAnggaran.join(',')}"`] });
        dataSheet.dataValidations.add(`I2:I${lastRow}`, { type: 'list', allowBlank: false, formulae: [`"${dropdownOptions.satuan.join(',')}"`] });
        dataSheet.dataValidations.add(`M2:M${lastRow}`, { type: 'list', allowBlank: false, formulae: [`"${dropdownOptions.statusKondisi.join(',')}"`] });
        dataSheet.dataValidations.add(`N2:N${lastRow}`, { type: 'list', allowBlank: true, formulae: [`"${dropdownOptions.statusPenghapusan.join(',')}"`] });

        res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition','attachment; filename=' + 'Template-Inventaris-MIJ-v5.xlsx');
        
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error('Gagal membuat template:', error);
        res.status(500).send(error.message);
    }
};

exports.uploadExcel = async (req, res) => {
    try {
        if (!req.file) { return res.status(400).send('Tidak ada file yang diunggah.'); }
        const workbook = new exceljs.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.getWorksheet('Data Inventaris'); // <-- REVISI DI SINI

        // Tambahkan pengecekan jika sheet tidak ditemukan
        if (!worksheet) {
            throw new Error('Sheet "Data Inventaris" tidak ditemukan di file Excel.');
        }

        const newItems = [];
        
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber > 1) {
                
                // --- 1. Ambil Semua Data Mentah dari Sel ---
                const namaBarang = row.getCell(1).value;
                const penanggungJawab = row.getCell(2).value || null;
                const kategori = row.getCell(3).value;
                const subKategori = row.getCell(4).value;
                const warna = row.getCell(5).value;
                const sumberAnggaran = row.getCell(6).value;
                const tahunPerolehanRaw = row.getCell(7).value;
                const jumlahRaw = row.getCell(8).value;
                const satuan = row.getCell(9).value;
                const nilaiPerolehanRaw = row.getCell(10).value;
                const lokasiFisik = row.getCell(11).value;
                const noPintuLokasi = row.getCell(12).value || null;
                const statusKondisi = row.getCell(13).value;
                const statusPenghapusanExcel = row.getCell(14).value;
                const dasarPenghapusanExcel = row.getCell(15).value;
                const tanggalPenghapusanExcel = row.getCell(16).value;

                // --- 2. BLOK VALIDASI SPESIFIK ---
                // Dapatkan konteks nama barang (jika ada) untuk pesan error
                const errorContext = namaBarang ? `(Barang: ${namaBarang})` : '';

                // Cek Wajib Isi (Nama Barang)
                if (!namaBarang) {
                    throw new Error(`Baris ${rowNumber}: Nama Barang wajib diisi (Kolom A).`);
                }
                
                // Cek Kategori (Wajib + Valid)
                if (!kategori) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Kategori wajib diisi (Kolom C).`);
                }
                if (!dataMasterLengkap[kategori]) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Kategori "${kategori}" tidak valid. Pilih dari daftar.`);
                }

                // Cek Sub Kategori (Wajib + Valid)
                if (!subKategori) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Sub Kategori wajib diisi (Kolom D).`);
              _ }
                const validSubKategoris = dataMasterLengkap[kategori];
                if (!validSubKategoris || !validSubKategoris[subKategori]) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Sub Kategori "${subKategori}" tidak valid untuk Kategori "${kategori}".`);
                }

                // Cek Warna (Wajib + Valid)
                if (!warna) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Warna wajib diisi (Kolom E).`);
                }
                if (!dropdownOptions.warna.includes(warna)) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Warna "${warna}" tidak valid. Pilih dari daftar.`);
                }

                // Cek Sumber Anggaran (Wajib + Valid)
                if (!sumberAnggaran) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Sumber Anggaran wajib diisi (Kolom F).`);
                }
                if (!kodeSumberAnggaranMap[sumberAnggaran]) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Sumber Anggaran "${sumberAnggaran}" tidak valid. Pilih dari daftar.`);
                }

                // Cek Tahun Perolehan (Wajib + Angka 4 Digit)
                if (!tahunPerolehanRaw) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Tahun Perolehan wajib diisi (Kolom G).`);
                }
                const tahunPerolehan = parseInt(tahunPerolehanRaw);
                if (isNaN(tahunPerolehan) || String(tahunPerolehan).length !== 4) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Tahun Perolehan "${tahunPerolehanRaw}" tidak valid. Masukkan angka 4 digit (contoh: 2024).`);
                }

                // Cek Jumlah (Wajib + Angka)
                if (jumlahRaw === null || jumlahRaw === undefined) { // Cek null/undefined, karena 0 bisa jadi valid
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Jumlah wajib diisi (Kolom H).`);
                }
                const jumlah = parseInt(jumlahRaw);
                if (isNaN(jumlah)) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Jumlah "${jumlahRaw}" tidak valid. Masukkan angka saja.`);
                }
                
                // Cek Satuan (Wajib + Valid)
                if (!satuan) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Satuan wajib diisi (Kolom I).`);
E               }
                if (!dropdownOptions.satuan.includes(satuan)) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Satuan "${satuan}" tidak valid. Pilih dari daftar.`);
                }

                // Cek Nilai Perolehan (Wajib + Angka)
                if (nilaiPerolehanRaw === null || nilaiPerolehanRaw === undefined) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Nilai Perolehan (Rp) wajib diisi (Kolom J). Masukkan 0 jika gratis.`);
                }
                const nilaiPerolehan = parseFloat(nilaiPerolehanRaw);
                if (isNaN(nilaiPerolehan)) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Nilai Perolehan "${nilaiPerolehanRaw}" tidak valid. Masukkan angka saja.`);
                }

                // Cek Lokasi Fisik (Wajib)
                if (!lokasiFisik) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Lokasi Fisik wajib diisi (Kolom K).`);
                }

                // Cek Status Kondisi (Wajib + Valid)
                if (!statusKondisi) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Status Kondisi wajib diisi (Kolom M).`);
                }
                if (!dropdownOptions.statusKondisi.includes(statusKondisi)) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Status Kondisi "${statusKondisi}" tidak valid. Pilih dari daftar.`);
      S         }
                
                // --- 3. Jika Lolos, Buat Objek itemData ---
                const itemData = {
                    namaBarang: namaBarang,
                    penanggungJawab: penanggungJawab,
                    kategori: kategori,
                    subKategori: subKategori,
                    warna: warna,
                    sumberAnggaran: sumberAnggaran,
                    tahunPerolehan: tahunPerolehan,
                    jumlah: jumlah,
                    satuan: satuan,
                    nilaiPerolehan: nilaiPerolehan,
                    lokasiFisik: lokasiFisik,
                    noPintuLokasi: noPintuLokasi,
                    statusKondisi: statusKondisi,
                    statusPenghapusanExcel: statusPenghapusanExcel,
                    dasarPenghapusanExcel: dasarPenghapusanExcel,
                    tanggalPenghapusanExcel: tanggalPenghapusanExcel,
          S       };

                // Validasi opsional (Status Penghapusan)
                if (itemData.statusPenghapusanExcel && !dropdownOptions.statusPenghapusan.includes(itemData.statusPenghapusanExcel)) {
                    throw new Error(`Baris ${rowNumber} ${errorContext}: Status Penghapusan "${itemData.statusPenghapusanExcel}" tidak valid (Kolom N).`);
                }

                newItems.push(itemData);
            }
        });

        if (newItems.length === 0) {
            return res.redirect(`/?uploadStatus=error&message=${encodeURIComponent('File Excel kosong atau format tidak sesuai.')}`);
        }

        const batch = db.batch();
        const counterRef = countersCollection.doc('inventoryCounter');
        const counterDoc = await counterRef.get();
        let lastNumber = counterDoc.data()?.lastNumber || 0;
        const now = new Date();

        newItems.forEach(item => {
            lastNumber++;
            const tahun = now.getFullYear();
            const bulanAngka = now.getMonth() + 1;
            const bulanRomawi = getRomanMonth(bulanAngka);
            const kodeAnggaran = kodeSumberAnggaranMap[item.sumberAnggaran] || 'ERR';
            const nomorUrutPadded = String(lastNumber).padStart(4, '0');
            
            const kodeSubKategori = dataMasterLengkap[item.kategori]?.[item.subKategori]?.kode || 'ERR-SUB';

            // --- REVISI ID BARU ---
            // Ambil noPintuLokasi dari item Excel, standarisasi, dan masukkan ke ID
            const kodeNoPintu = standardizeLokasi(item.noPintuLokasi);
            const nomorInventaris = `${nomorUrutPadded}/${kodeNoPintu}/${kodeSubKategori}/${kodeAnggaran}/INV-MIJ/${bulanRomawi}/${tahun}`;
            // --- AKHIR REVISI ---
            
            let tanggalPenghapusanValue = null;
            if (item.tanggalPenghapusanExcel instanceof Date) {
                tanggalPenghapusanValue = item.tanggalPenghapusanExcel;
            } else if (typeof item.tanggalPenghapusanExcel === 'string' && item.tanggalPenghapusanExcel.trim() !== '') {
                 try {
                      tanggalPenghapusanValue = new Date(item.tanggalPenghapusanExcel);
                      if (isNaN(tanggalPenghapusanValue.getTime())) tanggalPenghapusanValue = null;
                 } catch (e) { tanggalPenghapusanValue = null; }
            }

            const docRef = inventarisCollection.doc();
            batch.set(docRef, {
                namaBarang: item.namaBarang,
                penanggungJawab: item.penanggungJawab, // DITAMBAHKAN
                kategori: item.kategori,
                subKategori: item.subKategori,
                warna: item.warna,
                sumberAnggaran: item.sumberAnggaran,
                tahunPerolehan: item.tahunPerolehan,
                jumlah: item.jumlah,
                satuan: item.satuan,
                nilaiPerolehan: item.nilaiPerolehan,
                lokasiFisik: item.lokasiFisik,
                noPintuLokasi: item.noPintuLokasi || null,
                statusKondisi: item.statusKondisi,
                
                jumlah_awal: item.jumlah,
                lokasiFisik_awal: item.lokasiFisik,
                noPintuLokasi_awal: item.noPintuLokasi || null, // Ini "kunci" ID kita
                penanggungJawab_awal: item.penanggungJawab, // DITAMBAHKAN
                statusKondisi_awal: item.statusKondisi,

                statusPenghapusan: item.statusPenghapusanExcel || 'Masih Digunakan',
                dasarPenghapusan: item.dasarPenghapusanExcel || null,
                tanggalPenghapusan: tanggalPenghapusanValue,

                nomorInventaris, // Menggunakan ID baru yang sudah direvisi
                createdAt: now,
                createdBy: `${req.user.email} (via Upload)`,
                updatedAt: now,
                updatedBy: `${req.user.email} (via Upload)`,
            });
        });
        
        batch.set(counterRef, { lastNumber });
        await batch.commit();
        res.redirect(`/?uploadStatus=success&count=${newItems.length}`);

    } catch (error) {
        console.error('Gagal memproses file Excel:', error);
        res.redirect(`/?uploadStatus=error&message=${encodeURIComponent(error.message)}`);
    }
};

