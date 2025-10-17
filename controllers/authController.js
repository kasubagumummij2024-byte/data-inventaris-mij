// controllers/authController.js
const { getAuth, createUserWithEmailAndPassword, signInWithEmailAndPassword } = require('firebase/auth');
const admin = require('firebase-admin');

// PENTING: Ganti dengan konfigurasi asli dari Proyek Firebase Anda
const firebaseClientConfig = {
  apiKey: "AIzaSyD-lA9D95bH7L31Qu9PC5artR8McKD5QQU",
  authDomain: "webapp-data-inventaris.firebaseapp.com",
  projectId: "webapp-data-inventaris",
  storageBucket: "webapp-data-inventaris.firebasestorage.app",
  messagingSenderId: "37876691041",
  appId: "1:37876691041:web:f62d4cc5f3a4f7443a7a20"
};
const { initializeApp } = require('firebase/app');
const firebaseApp = initializeApp(firebaseClientConfig);
const auth = getAuth(firebaseApp);

exports.requireAuth = async (req, res, next) => {
    const token = req.cookies.session;
    if (!token) {
        return res.redirect('/login');
    }
    try {
        req.user = await admin.auth().verifyIdToken(token);
        next();
    } catch (error) {
        res.redirect('/login');
    }
};

exports.checkUser = async (req, res, next) => {
    const token = req.cookies.session;
    if (token) {
        try {
            res.locals.user = await admin.auth().verifyIdToken(token);
        } catch (error) {
            res.locals.user = null;
        }
    } else {
        res.locals.user = null;
    }
    next();
};

exports.getLoginPage = (req, res) => res.render('auth/login', { error: null });

exports.loginUser = async (req, res) => {
    const { email, password } = req.body;
    try {
        const userCredential = await signInWithEmailAndPassword(auth, email, password);
        const idToken = await userCredential.user.getIdToken();
        res.cookie('session', idToken, { httpOnly: true, secure: process.env.NODE_ENV === 'production' });
        res.redirect('/');
    } catch (error) {
        res.render('auth/login', { error: 'Email atau password salah.' });
    }
};

exports.logoutUser = (req, res) => {
    res.clearCookie('session');
    res.redirect('/');
};