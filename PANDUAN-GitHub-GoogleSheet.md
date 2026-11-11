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
5. Di editor pilih fungsi **tampilkanUrlApi** → **Run**. Buka **Execution log**, lalu salin baris `const API_URL='https://script.google.com/macros/s/…/exec';`.
   (Atau salin sendiri URL Web app yang berakhiran `/exec`.) Membuka URL itu di browser harus menampilkan `{"ok":true,"app":"Stock Opname Gudang API"}`.

> Setiap kali `Code.gs` diubah: Deploy → Manage deployments → ikon pensil → Version: **New version** → Deploy. URL tidak berubah.

## B. Tayangkan di GitHub Pages
1. Buat repository GitHub (misalnya `stock-opname`).
2. Buka `index.html` dengan editor teks (atau tombol pensil di GitHub), cari baris `const API_URL='';` di bagian **PENGATURAN SERVER**, lalu tempel baris dari langkah A5 menggantikannya. Contoh:
   `const API_URL='https://script.google.com/macros/s/XXXX/exec';`
   URL ini tertanam di file, jadi **pengguna tidak perlu mengisi apa pun** dan halaman login langsung terhubung.
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
| Login, isi stok fisik (hitungan sendiri), tambah barang, scan barcode | ✔ | ✔ |
| Muat ulang, Export Excel (hitungan yang sedang tampil) | ✔ | ✔ |
| Membuat sesi hitung baru, Clear stok fisik (sesi sendiri) | ✔ | ✔ |
| Melihat sesi akun lain / total semua | – | ✔ (hanya lihat) |
| Reset semua hitungan semua akun | – | ✔ |
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
| "Server belum dikonfigurasi. Hubungi IT." | `API_URL` di `index.html` masih kosong atau bukan URL `/exec` yang valid |
| "Tidak bisa terhubung ke server" | URL `/exec` salah atau internet putus |
| "Respon server tidak valid" | Deployment belum diatur **Anyone**, atau belum dibuat New version setelah mengganti kode |
| "Setup belum dijalankan" | Jalankan `setupAkunIT` dari editor |
| "Username atau password salah" | Cek spasi; akun terkunci 5 menit setelah 5 kali salah |

## G. Hitungan terpisah: per akun dan per sesi
Daftar barang dan **stok sistem** dipakai bersama. **Stok fisik** disimpan terpisah di sheet **Hitung**, per akun dan per **sesi hitung**.

**Sesi hitung**
- Setiap akun otomatis punya **Sesi utama**.
- Tekan **＋ Sesi baru** setiap kali mulai menghitung (misalnya per rak, per lokasi, atau per hari) dan beri nama. Sesi baru kosong dan terpisah dari sesi lain.
- Pilih sesi lewat kotak **Sesi hitung**. Jumlah barang yang sudah dihitung ditampilkan di tiap sesi.
- **Login bersamaan:** dua akun berbeda otomatis terpisah. Kalau **satu akun dipakai di dua perangkat** sekaligus, tekan **＋ Sesi baru** di masing-masing perangkat supaya hitungannya tidak bercampur.
- Browser mengingat sesi terakhir yang dipakai tiap akun, jadi setelah login kamu melanjutkan sesi yang sama.

**Clear stok fisik**
- Tombol **🗑 Clear stok fisik** menghapus semua stok fisik **di sesi yang sedang dipilih** (milik sendiri). Ada konfirmasi dulu, dan sesi lain serta hitungan akun lain tidak terpengaruh.
- Khusus IT: **⚠ Reset semua hitungan** menghapus hitungan semua akun dan semua sesi (untuk memulai opname baru). Perlu konfirmasi dan mengetik `HAPUS`. Buat sheet Laporan atau Export Excel dulu kalau datanya masih dibutuhkan.

**Hak akses**
- **Team Gudang** hanya melihat dan mengubah sesi miliknya. Sesi akun lain ditolak oleh server.
- **IT** bisa melihat sesi akun mana pun dan **Total semua akun & sesi**. Tampilan itu hanya bisa dilihat, jadi kolom stok fisiknya terkunci.
- **Export Excel** mengekspor sesi yang sedang tampil, dengan nama akun dan sesi di nama file.
- **Buat sheet Laporan** (IT) membuat satu kolom `Fisik - <username> / <sesi>` per sesi yang berisi data, ditambah **Total Fisik**, **Selisih**, dan **Status**.

**Pembagian kerja yang disarankan**
Total dan selisih di Laporan menjumlahkan semua sesi. Kalau dua sesi menghitung barang yang sama, angkanya dobel. Bagi area kerja per sesi (per rak atau lokasi). Untuk hitung ulang, bandingkan lewat kolom per sesi, bukan Total Fisik.

**Memperbarui dari versi sebelumnya**
1. Tempel `Code.gs` baru ke Apps Script, lalu **Deploy → Manage deployments → ikon pensil → Version: New version → Deploy**.
2. Ganti `index.html` di GitHub dengan versi baru (isi `API_URL` lagi kalau kamu mengisinya di file).
3. Sheet **Sesi** dibuat otomatis, dan sheet **Hitung** diperbarui otomatis (kolom `sesi` ditambahkan). Hitungan dari versi per-akun sebelumnya otomatis menjadi **Sesi utama** akun itu.
4. Hitungan bersama versi paling lama (satu hitungan untuk semua akun) tidak tampil. Untuk menyalinnya, jalankan fungsi **migrasiHitungLama** dari editor Apps Script.
5. Kolom *Stok Fisik* di file Excel untuk **Import Excel** diabaikan. Hitungan hanya diisi lewat input atau scan.
6. Kalau akun dihapus, hitungannya tetap ada di sheet Hitung dan ikut dijumlahkan di Total dan Laporan.
