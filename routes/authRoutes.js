// routes/authRoutes.js
const express = require('express');
const router = express.Router();
const authController = require('../controllers/authController');

// Rute untuk menampilkan halaman
// router.get('/register', authController.getRegisterPage); // <-- HAPUS BARIS INI
router.get('/login', authController.getLoginPage);

// Rute untuk memproses form
// router.post('/register', authController.registerUser); // <-- HAPUS BARIS INI
router.post('/login', authController.loginUser);
router.get('/logout', authController.logoutUser);

module.exports = router;