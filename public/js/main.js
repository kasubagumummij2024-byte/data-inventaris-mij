// public/js/main.js
document.addEventListener('DOMContentLoaded', function() {
    const kategoriSelect = document.getElementById('kategori');
    const subKategoriSelect = document.getElementById('subKategori');

    // Cek apakah elemen ada di halaman
    if (kategoriSelect && subKategoriSelect) {
        // Ambil data master dari atribut data-* yang kita set di EJS
        const masterKategori = JSON.parse(kategoriSelect.dataset.master);

        function updateSubKategori() {
            const selectedKategori = kategoriSelect.value;
            const subKategoriOptions = masterKategori[selectedKategori] || [];
            
            // Simpan nilai sub kategori yang sedang dipilih (untuk form edit)
            const currentSubKategoriValue = subKategoriSelect.value;

            // Kosongkan pilihan sub kategori
            subKategoriSelect.innerHTML = '<option value="" disabled selected>-- Pilih Sub Kategori --</option>';

            // Isi dengan pilihan yang baru
            subKategoriOptions.forEach(sub => {
                const option = document.createElement('option');
                option.value = sub;
                option.textContent = sub;
                subKategoriSelect.appendChild(option);
            });
            
            // Set kembali nilai yang sebelumnya terpilih jika masih ada di opsi baru
            if (subKategoriOptions.includes(currentSubKategoriValue)) {
                subKategoriSelect.value = currentSubKategoriValue;
            }
        }

        // Panggil fungsi saat halaman dimuat (untuk form edit)
        updateSubKategori();

        // Panggil fungsi setiap kali kategori berubah
        kategoriSelect.addEventListener('change', updateSubKategori);
    }
});