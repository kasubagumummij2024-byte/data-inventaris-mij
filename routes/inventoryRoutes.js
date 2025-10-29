// routes/inventoryRoutes.js
const express = require('express');
const router = express.Router();
const inventoryController = require('../controllers/inventoryController');
const { requireAuth } = require('../controllers/authController');
const multer = require('multer');

const upload = multer({ storage: multer.memoryStorage() });

// === RUTE PUBLIK ===
router.get('/', inventoryController.getAllItems);
router.get('/barang/:id', inventoryController.getItemDetail);
router.get('/download-excel', inventoryController.downloadExcel); // Laporan LENGKAP
router.get('/download-template', inventoryController.downloadTemplate);
router.get('/referensi-kode', inventoryController.getReferensiKodePage);
router.get('/download-referensi', inventoryController.downloadReferensiExcel);

// BARU: Rute khusus untuk men-download Laporan Label
router.get('/download-label', inventoryController.downloadLabelSheet);


// === RUTE TERPROTEKSI ===
router.get('/tambah', requireAuth, inventoryController.getAddItemForm);
router.post('/tambah', requireAuth, inventoryController.createItem);
router.get('/edit/:id', requireAuth, inventoryController.getEditItemForm);
router.post('/update/:id', requireAuth, inventoryController.updateItem);
router.post('/hapus/:id', requireAuth, inventoryController.deleteItem);
router.post('/upload-excel', requireAuth, upload.single('excelFile'), inventoryController.uploadExcel);

module.exports = router;