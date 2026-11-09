# Panduan: Web di GitHub Pages + Login & Data di Google Sheet (Google Drive)

Pembagian tugas:
- **GitHub** menyimpan dan menayangkan halaman web (`index.html`).
- **Google Sheet di Google Drive kamu** menyimpan akun login, data barang, dan riwayat (Log).
- **Google Apps Script** menjadi penghubung (API). Login dicek di server ini, bukan di browser.

## A. Siapkan Google Sheet + API (IT, sekali saja)
1. Buka Google Sheets, buat spreadsheet baru di Drive, misalnya **Stock Opname Gudang**.
2. **Extensions → Apps Script**. Hapus isi `Code.gs`, tempel isi file **Code.gs** terbaru.
   Ubah baris `AKUN_IT_AWAL` (username, nama, password minimal 8 karakter).
   File HTML bernama `Index` **tidak diperlukan lagi**. Kalau sudah ada, boleh dihapus.
3. Pilih fungsi **setupAkunIT** → **Run** → izinkan akses. Setelah sukses, **hapus password dari kode** dan simpan.
4. **Deploy → New deployment → Web app**: Execute as **Me**, Who has access **Anyone** → Deploy.
5. Salin URL yang berakhiran `/exec`. Membuka URL itu di browser harus menampilkan `{"ok":true,"app":"Stock Opname Gudang API"}`.

> Setiap kali `Code.gs` diubah: Deploy → Manage deployments → ikon pensil → Version: **New version** → Deploy. URL tidak berubah.

## B. Tayangkan di GitHub Pages
1. Buat repository GitHub (misalnya `stock-opname`).
2. Buka `index.html`, cari baris `const API_URL='';` lalu isi dengan URL `/exec` tadi:
   `const API_URL='https://script.google.com/macros/s/XXXX/exec';`
   (Kalau dikosongkan, halaman login meminta URL itu dan menyimpannya di browser masing-masing.)
3. Upload `index.html` ke repository (Add file → Upload files), lalu **Settings → Pages → Deploy from a branch → main / (root) → Save**.
4. Alamat web: `https://USERNAME.github.io/stock-opname/`. Bagikan alamat ini ke Team Gudang.

Untuk memperbarui halaman nantinya, cukup edit atau upload ulang `index.html` di GitHub. Data tidak ikut berubah karena ada di Google Sheet.

## C. Akun login dari data di Google Drive
Akun disimpan di sheet **Akun** (password berupa hash). Ada dua cara menambah akun:
- **Lewat web**: login sebagai IT → **Kelola akun**.
- **Dari daftar yang sudah ada di Sheet** (banyak akun sekaligus):
  1. Di spreadsheet yang sama buat sheet **Akun_Import** dengan kolom: `username | nama | role | password`
     (role diisi `IT` atau `Gudang`; password minimal 8 karakter).
  2. Di editor Apps Script jalankan fungsi **importAkunDariSheet**.
  3. Akun dibuat, password diubah jadi hash, dan kolom password di Akun_Import diganti `(sudah diimpor)`.
  Baris yang dilewati (kosong, username sudah ada, password terlalu pendek) tercantum di log eksekusi.

## D. Hak akses
| Fitur | Team Gudang | IT |
|---|---|---|
| Login, isi stok fisik, tambah barang, scan barcode | ✔ | ✔ |
| Muat ulang, Export Excel | ✔ | ✔ |
| Edit / hapus barang, Import Excel | – | ✔ |
| Kelola akun, Buat sheet Laporan | – | ✔ |

Pembatasan peran dicek **di server** (Apps Script).

## E. Catatan keamanan
- **Jangan bagikan Google Sheet ke Team Gudang.** Mereka cukup memakai alamat web.
- URL `/exec` ikut terlihat di `index.html` (publik kalau repo publik). Itu tidak berbahaya: tanpa login yang benar tidak ada data yang keluar. Percobaan login salah 5 kali mengunci akun itu 5 menit.
- Gunakan password yang panjang dan unik. Sesi login berlaku maksimal 6 jam.
- Perubahan stok fisik, barang, dan akun tercatat di sheet **Log** (siapa, kapan).
- Kuota Apps Script gratis cukup untuk satu gudang dengan beberapa petugas.

## F. Kalau ada masalah
| Gejala | Cek |
|---|---|
| "Tidak bisa terhubung ke server" | URL `/exec` salah atau internet putus |
| "Respon server tidak valid" | Deployment belum diatur **Anyone**, atau belum dibuat New version setelah mengganti kode |
| "Setup belum dijalankan" | Jalankan `setupAkunIT` dari editor |
| "Username atau password salah" | Cek spasi; akun terkunci 5 menit setelah 5 kali salah |
