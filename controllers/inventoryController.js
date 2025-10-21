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

const masterKategori = { 'Aset Kantor & Furnitur': ['Meja', 'Kursi', 'Lemari', 'Papan Tulis', 'Lainnya'],'Perangkat Elektronik & IT': ['Komputer', 'Laptop', 'Printer', 'Proyektor', 'Server', 'Jaringan', 'Lainnya'],'ATK & Habis Pakai': ['Kertas', 'Alat Tulis', 'Tinta & Toner', 'Baterai', 'Lainnya'],'Perlengkapan Operasional': ['Mesin', 'Peralatan Tangan', 'Alat Ukur', 'APD', 'Lainnya'],'Aset Kendaraan': ['Mobil', 'Motor', 'Lainnya'],'Kebersihan & Maintenance': ['Alat Kebersihan', 'Bahan Pembersih', 'Lainnya'],'Lain-lain': ['Lain-lain'] };
const kodeKategoriMap = { 'Aset Kantor & Furnitur': 'AKF','Perangkat Elektronik & IT': 'EIT','ATK & Habis Pakai': 'ATK','Perlengkapan Operasional': 'OPS','Aset Kendaraan': 'KDN','Kebersihan & Maintenance': 'KMT','Lain-lain': 'LLN' };
const kodeSumberAnggaranMap = { 'BOS KB': 'BOSKB', 'BOS RA': 'BOSRA', 'BOS MI': 'BOSMI', 'BOS MTs': 'BOSMTS', 'BOS MA': 'BOSMA','BOP KB': 'BOPKB', 'BOP RA': 'BOPRA','KAS KB': 'KASKB', 'KAS RA': 'KASRA', 'KAS MI': 'KASMI', 'KAS MTS': 'KASMTS', 'KAS MA': 'KASMA','HIBAH': 'HBH', 'SPONSORSHIP': 'SPS', 'MANDIRI': 'MDR', 'LAIN-LAIN': 'LLN' };
const dropdownOptions = { satuan: ['Unit', 'Pcs', 'Set', 'Lusin', 'Box', 'Roll', 'Kg', 'Liter'],statusKondisi: ['Baik', 'Perlu Perbaikan', 'Rusak', 'Dalam Proses Perbaikan', 'Habis'],warna: ['Hitam', 'Putih', 'Abu-abu', 'Silver', 'Merah', 'Biru', 'Hijau', 'Kuning', 'Coklat', 'Oranye', 'Lainnya'],sumberAnggaran: Object.keys(kodeSumberAnggaranMap) };

// ... (Fungsi getAllItems, getAddItemForm, createItem, getEditItemForm, updateItem, deleteItem, getItemDetail tidak berubah) ...
exports.getAllItems = async (req, res) => { try { const { search, uploadStatus, message, count } = req.query; let query = inventarisCollection.orderBy('nomorInventaris', 'desc'); const snapshot = await query.get(); let items = []; snapshot.forEach(doc => items.push({ id: doc.id, ...doc.data() })); if (search) { items = items.filter(item => Object.values(item).some(value => String(value).toLowerCase().includes(search.toLowerCase()))); } res.render('index', { items, search: search || '', uploadStatus, message, count }); } catch (error) { res.status(500).send(error.message); } };
exports.getAddItemForm = (req, res) => { res.render('form-tambah', { masterKategori, options: dropdownOptions }); };
exports.createItem = async (req, res) => { try { const counterRef = countersCollection.doc('inventoryCounter'); let newNomorUrut; await db.runTransaction(async (t) => { const counterDoc = await t.get(counterRef); newNomorUrut = (counterDoc.data()?.lastNumber || 0) + 1; t.set(counterRef, { lastNumber: newNomorUrut }); }); const { kategori, sumberAnggaran } = req.body; const now = new Date(); const tahun = now.getFullYear(); const bulanAngka = now.getMonth() + 1; const bulanRomawi = getRomanMonth(bulanAngka); const nomorUrutPadded = String(newNomorUrut).padStart(4, '0'); const kodeKategori = kodeKategoriMap[kategori] || 'ERR'; const kodeAnggaran = kodeSumberAnggaranMap[sumberAnggaran] || 'ERR'; const nomorInventaris = `${nomorUrutPadded}/${kodeKategori}/${kodeAnggaran}/INV-MIJ/${bulanRomawi}/${tahun}`; const newItem = { namaBarang: req.body.namaBarang, kategori: req.body.kategori, subKategori: req.body.subKategori, warna: req.body.warna, sumberAnggaran: req.body.sumberAnggaran, jumlah: parseInt(req.body.jumlah), satuan: req.body.satuan, nilaiPerolehan: parseFloat(req.body.nilaiPerolehan), lokasiFisik: req.body.lokasiFisik, statusKondisi: req.body.statusKondisi, jumlah_awal: parseInt(req.body.jumlah), lokasiFisik_awal: req.body.lokasiFisik, statusKondisi_awal: req.body.statusKondisi, nomorInventaris, createdAt: now, createdBy: req.user.email, updatedAt: now, updatedBy: req.user.email, }; await inventarisCollection.add(newItem); res.redirect('/'); } catch (error) { res.status(500).send(error.message); } };
exports.getEditItemForm = async (req, res) => { try { const doc = await inventarisCollection.doc(req.params.id).get(); if (!doc.exists) return res.status(404).send('Barang tidak ditemukan'); res.render('form-edit', { item: { id: doc.id, ...doc.data() }, masterKategori, options: dropdownOptions }); } catch (error) { res.status(500).send(error.message); } };
exports.updateItem = async (req, res) => { try { const { nomorInventaris, ...restOfBody } = req.body; const updatedItem = { ...restOfBody, jumlah: parseInt(req.body.jumlah), nilaiPerolehan: parseFloat(req.body.nilaiPerolehan), updatedAt: new Date(), updatedBy: req.user.email }; await inventarisCollection.doc(req.params.id).update(updatedItem); res.redirect('/'); } catch (error) { res.status(500).send(error.message); } };
exports.deleteItem = async (req, res) => { try { await inventarisCollection.doc(req.params.id).delete(); res.redirect('/'); } catch (error) { res.status(500).send(error.message); } };
exports.getItemDetail = async (req, res) => { try { const doc = await inventarisCollection.doc(req.params.id).get(); if (!doc.exists) return res.status(404).send('Barang tidak ditemukan'); const itemData = { id: doc.id, ...doc.data() }; const url = `${req.protocol}://${req.get('host')}/barang/${itemData.id}`; const qrCodeDataUrl = await QRCode.toDataURL(url); res.render('detail', { item: itemData, qrCodeDataUrl }); } catch (error) { res.status(500).send(error.message); } };

// =======================================================
// REVISI TOTAL DIMULAI DARI FUNGSI INI
// =======================================================
exports.downloadExcel = async (req, res) => {
    try {
        const snapshot = await inventarisCollection.orderBy('createdAt', 'asc').get();
        let items = [];
        snapshot.forEach(doc => items.push({ id: doc.id, ...doc.data() }));

        const workbook = new exceljs.Workbook();
        const worksheet = workbook.addWorksheet('Data Inventaris dengan QR Code');

        // DIUBAH: Menambahkan kolom 'QR Code' di awal dan menyesuaikan lebar
        worksheet.columns = [
            { header: 'QR Code', key: 'qr', width: 15 }, // Kolom untuk gambar QR
            { header: 'Nomor Inventaris', key: 'nomorInventaris', width: 30 },
            { header: 'Nama Barang', key: 'namaBarang', width: 30 },
            { header: 'Warna', key: 'warna', width: 15 },
            { header: 'Sumber Anggaran', key: 'sumberAnggaran', width: 20 },
            { header: 'Kategori', key: 'kategori', width: 25 },
            { header: 'Sub Kategori', key: 'subKategori', width: 20 },
            { header: 'Lokasi Terkini', key: 'lokasiFisik', width: 30 },
            { header: 'Jumlah Terkini', key: 'jumlah', width: 15 },
            { header: 'Satuan', key: 'satuan', width: 10 },
            { header: 'Kondisi Terkini', key: 'statusKondisi', width: 20 },
        ];
        
        worksheet.getRow(1).font = { bold: true };

        // DIUBAH: Menggunakan loop 'for...of' agar bisa memakai 'await' di dalamnya
        // Kita tidak bisa menggunakan forEach dengan fungsi async
        let rowNumber = 2; // Mulai dari baris ke-2 setelah header
        for (const item of items) {
            // Tambahkan data teks ke baris terlebih dahulu
            worksheet.addRow({
                nomorInventaris: item.nomorInventaris,
                namaBarang: item.namaBarang,
                warna: item.warna,
                sumberAnggaran: item.sumberAnggaran,
                kategori: item.kategori,
                subKategori: item.subKategori,
                lokasiFisik: item.lokasiFisik,
                jumlah: item.jumlah,
                satuan: item.satuan,
                statusKondisi: item.statusKondisi,
            });

            // BARU: Atur tinggi baris agar gambar QR Code muat
            const currentRow = worksheet.getRow(rowNumber);
            currentRow.height = 80;
            currentRow.alignment = { vertical: 'middle' };

            // BARU: Generate QR Code sebagai buffer gambar
            const url = `${req.protocol}://${req.get('host')}/barang/${item.id}`;
            const qrBuffer = await QRCode.toBuffer(url, { 
                type: 'png', 
                width: 100, // Ukuran gambar QR
                margin: 1 
            });

            // BARU: Tambahkan buffer gambar ke workbook
            const imageId = workbook.addImage({
                buffer: qrBuffer,
                extension: 'png',
            });

            // BARU: Sisipkan gambar ke sel yang tepat di kolom 'QR Code'
            // 'tl' adalah top-left corner, 'ext' adalah extent (ukuran)
            worksheet.addImage(imageId, {
                tl: { col: 0.1, row: rowNumber - 0.9 }, // Kolom A, Baris saat ini (dengan sedikit margin)
                ext: { width: 100, height: 100 }
            });

            rowNumber++;
        }

        res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition','attachment; filename=' + 'Laporan_Inventaris_dengan_QR_Code.xlsx');
        
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error("Gagal men-generate Excel dengan QR Code:", error);
        res.status(500).send(error.message);
    }
};

// ... (Fungsi downloadTemplate dan uploadExcel tidak berubah) ...
exports.downloadTemplate = async (req, res) => { try { const workbook = new exceljs.Workbook(); const petunjukSheet = workbook.addWorksheet('Petunjuk'); petunjukSheet.getColumn('A').width = 80; petunjukSheet.getCell('A1').value = 'PETUNJUK PENGISIAN TEMPLATE INVENTARIS MIJ'; petunjukSheet.getCell('A1').font = { bold: true, size: 16 }; petunjukSheet.getCell('A3').value = '1. Jangan mengubah, menghapus, atau menambah kolom (header) di sheet "Data Inventaris".'; petunjukSheet.getCell('A4').value = '2. Isi data mulai dari baris kedua di sheet "Data Inventaris".'; petunjukSheet.getCell('A5').value = '3. Untuk kolom Kategori, Satuan, Status Kondisi, Warna, dan Sumber Anggaran, WAJIB memilih dari daftar dropdown.'; petunjukSheet.getCell('A6').value = '4. Untuk kolom Nilai Perolehan, masukkan angka saja tanpa "Rp" atau titik (contoh: 5000000).'; petunjukSheet.getCell('A7').value = '5. Setelah selesai, simpan file ini dan unggah melalui WebApp Inventaris.'; const dataSheet = workbook.addWorksheet('Data Inventaris'); dataSheet.columns = [ { header: 'Nama Barang', key: 'namaBarang', width: 40 }, { header: 'Kategori', key: 'kategori', width: 30 }, { header: 'Sub Kategori', key: 'subKategori', width: 25 }, { header: 'Warna', key: 'warna', width: 20 }, { header: 'Sumber Anggaran', key: 'sumberAnggaran', width: 25 }, { header: 'Jumlah', key: 'jumlah', width: 15 }, { header: 'Satuan', key: 'satuan', width: 15 }, { header: 'Nilai Perolehan (Rp)', key: 'nilaiPerolehan', width: 25 }, { header: 'Lokasi Fisik', key: 'lokasiFisik', width: 40 }, { header: 'Status Kondisi', key: 'statusKondisi', width: 25 } ]; dataSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } }; dataSheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF203764' } }; const lastRow = 1001; dataSheet.dataValidations.add(`B2:B${lastRow}`, { type: 'list', allowBlank: false, formulae: [`"${Object.keys(masterKategori).join(',')}"`] }); dataSheet.dataValidations.add(`D2:D${lastRow}`, { type: 'list', allowBlank: false, formulae: [`"${dropdownOptions.warna.join(',')}"`] }); dataSheet.dataValidations.add(`E2:E${lastRow}`, { type: 'list', allowBlank: false, formulae: [`"${dropdownOptions.sumberAnggaran.join(',')}"`] }); dataSheet.dataValidations.add(`G2:G${lastRow}`, { type: 'list', allowBlank: false, formulae: [`"${dropdownOptions.satuan.join(',')}"`] }); dataSheet.dataValidations.add(`J2:J${lastRow}`, { type: 'list', allowBlank: false, formulae: [`"${dropdownOptions.statusKondisi.join(',')}"`] }); res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); res.setHeader('Content-Disposition','attachment; filename=' + 'Template-Inventaris-MIJ.xlsx'); await workbook.xlsx.write(res); res.end(); } catch (error) { console.error('Gagal membuat template:', error); res.status(500).send(error.message); } };
exports.uploadExcel = async (req, res) => { try { if (!req.file) { return res.status(400).send('Tidak ada file yang diunggah.'); } const workbook = new exceljs.Workbook(); await workbook.xlsx.load(req.file.buffer); const worksheet = workbook.getWorksheet(1); const newItems = []; worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => { if (rowNumber > 1) { const itemData = { namaBarang: row.getCell(1).value, kategori: row.getCell(2).value, subKategori: row.getCell(3).value, warna: row.getCell(4).value, sumberAnggaran: row.getCell(5).value, jumlah: parseInt(row.getCell(6).value), satuan: row.getCell(7).value, nilaiPerolehan: parseFloat(row.getCell(8).value), lokasiFisik: row.getCell(9).value, statusKondisi: row.getCell(10).value, }; if (!itemData.namaBarang || !itemData.kategori || !itemData.subKategori || !itemData.sumberAnggaran) { throw new Error(`Data tidak valid di baris ${rowNumber}. Nama Barang, Kategori, Sub Kategori, dan Sumber Anggaran wajib diisi.`); } const validSubKategoriList = masterKategori[itemData.kategori]; if (!validSubKategoriList || !validSubKategoriList.includes(itemData.subKategori)) { throw new Error(`Sub Kategori "${itemData.subKategori}" tidak valid untuk Kategori "${itemData.kategori}" di baris ${rowNumber}.`); } newItems.push(itemData); } }); if (newItems.length === 0) { return res.redirect(`/?uploadStatus=error&message=${encodeURIComponent('File Excel kosong atau format tidak sesuai.')}`); } const batch = db.batch(); const counterRef = countersCollection.doc('inventoryCounter'); const counterDoc = await counterRef.get(); let lastNumber = counterDoc.data()?.lastNumber || 0; const now = new Date(); newItems.forEach(item => { lastNumber++; const tahun = now.getFullYear(); const bulanAngka = now.getMonth() + 1; const bulanRomawi = getRomanMonth(bulanAngka); const kodeKategori = kodeKategoriMap[item.kategori] || 'ERR'; const kodeAnggaran = kodeSumberAnggaranMap[item.sumberAnggaran] || 'ERR'; const nomorUrutPadded = String(lastNumber).padStart(4, '0'); const nomorInventaris = `${nomorUrutPadded}/${kodeKategori}/${kodeAnggaran}/INV-MIJ/${bulanRomawi}/${tahun}`; const docRef = inventarisCollection.doc(); batch.set(docRef, { ...item, nomorInventaris, jumlah_awal: item.jumlah, lokasiFisik_awal: item.lokasiFisik, statusKondisi_awal: item.statusKondisi, createdAt: now, createdBy: `${req.user.email} (via Upload)`, updatedAt: now, updatedBy: `${req.user.email} (via Upload)`, }); }); batch.set(counterRef, { lastNumber }); await batch.commit(); res.redirect(`/?uploadStatus=success&count=${newItems.length}`); } catch (error) { console.error('Gagal memproses file Excel:', error); res.redirect(`/?uploadStatus=error&message=${encodeURIComponent(error.message)}`); } };

