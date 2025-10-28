// public/js/main.js
document.addEventListener('DOMContentLoaded', function() {
    const kategoriSelect = document.getElementById('kategori');
    const subKategoriSelect = document.getElementById('subKategori');

    if (kategoriSelect && subKategoriSelect) {
        // DIUBAH: Ambil data master baru dari atribut data-*
        const dataMasterLengkap = JSON.parse(kategoriSelect.dataset.master);

        function updateSubKategori() {
            const selectedKategori = kategoriSelect.value;
            // DIUBAH: Ambil daftar subkategori dari struktur baru
            const subKategoriObject = dataMasterLengkap[selectedKategori] || {};
            const subKategoriOptions = Object.keys(subKategoriObject); // Ambil nama-nama subkategori

            const currentSubKategoriValue = subKategoriSelect.value;
            subKategoriSelect.innerHTML = '<option value="" disabled selected>-- Pilih Sub Kategori --</option>';

            subKategoriOptions.forEach(sub => {
                const option = document.createElement('option');
                option.value = sub; // Value tetap nama subkategori
                option.textContent = sub;
                subKategoriSelect.appendChild(option);
            });

            // Set kembali nilai yang sebelumnya terpilih (untuk form edit)
            // Cek dataset initialValue jika ada
            const initialValue = subKategoriSelect.dataset.initialValue;
            if (subKategoriOptions.includes(currentSubKategoriValue)) {
                 subKategoriSelect.value = currentSubKategoriValue;
            } else if (initialValue && subKategoriOptions.includes(initialValue)) {
                 // Jika nilai saat ini tidak valid TAPI nilai awal valid, gunakan nilai awal
                 subKategoriSelect.value = initialValue;
                 // Hapus initial value agar tidak terpilih lagi jika kategori diubah lagi
                 delete subKategoriSelect.dataset.initialValue;
            }
        }

        // Simpan nilai awal untuk form edit (hanya jika ada nilai terpilih)
        if (subKategoriSelect.options.length > 0 && subKategoriSelect.options[0].value !== "" && subKategoriSelect.options[0].selected) {
             subKategoriSelect.dataset.initialValue = subKategoriSelect.options[0].value;
        }


        // Panggil saat halaman dimuat (penting untuk form edit)
        updateSubKategori();

        kategoriSelect.addEventListener('change', updateSubKategori);
    }
});