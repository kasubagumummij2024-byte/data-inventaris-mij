// index.js
// KODE BARU
if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const cookieParser = require('cookie-parser');

const inventoryRoutes = require('./routes/inventoryRoutes');
const authRoutes = require('./routes/authRoutes');
const { checkUser } = require('./controllers/authController');

require('./config/firebaseConfig');

const app = express();
const port = process.env.PORT || 3000;

// Setup Middleware
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(cookieParser());

// Middleware ini membuat variabel 'user' tersedia di semua halaman
app.use(checkUser); 

// Gunakan Rute
app.use('/', inventoryRoutes);
app.use('/', authRoutes);

app.listen(port, () => {
    console.log(`🚀 Server berjalan di http://localhost:${port}`);
});