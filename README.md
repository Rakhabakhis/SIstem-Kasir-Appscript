# 🧾 Sistem Kasir (POS) - Google Apps Script

Sistem Point of Sale lengkap berbasis Google Apps Script & Google Sheets sebagai database. Cocok untuk warung, kafe, dan bisnis kecil menengah.

---

## ✨ Fitur Lengkap

### 🛒 Kasir & Transaksi
- **Pesanan Baru** — Tambah produk ke keranjang, filter kategori, view grid/list
- **2 Mode Checkout** — Bayar Langsung (cash/QRIS) atau Buat Pesanan (hold/pending)
- **Gabung Pesanan** — Gabungkan item dari pesanan berbeda sebelum pembayaran
- **Varian Produk** — Pilih rasa/varian saat checkout (Bahan Pilihan per grup)
- **Cetak Struk** — Format thermal 80mm, otomatis muncul setelah bayar
- **2 Metode Pembayaran** — Cash & QRIS
- **Produk Gratis** — Opsi harga Rp 0 untuk item bonus/compliment

### 📋 Daftar Pesanan (Hold)
- Lihat semua pesanan yang belum dibayar
- Bayar pesanan hold di hari berbeda — waktu transaksi tercatat saat **pembayaran dilakukan** (bukan saat pesanan dibuat)
- Batalkan pesanan

### 📦 Manajemen Produk
- Tambah, edit, hapus/nonaktifkan produk
- Filter & cari produk
- URL gambar dari Google Drive (otomatis dikonversi ke format thumbnail yang bisa ditampilkan di browser)
- Status Aktif/Nonaktif — produk nonaktif disembunyikan dari kasir

### 🏷️ Manajemen Kategori (Dinamis)
- Tambah, ubah nama, dan hapus kategori secara bebas
- Kategori tersimpan di **sheet master `Kategori`** — tidak perlu ada produk dulu
- Dropdown kategori saat tambah/edit produk selalu **real-time** dari master
- Data validation di Google Sheets otomatis diperbarui setiap ada perubahan kategori
- Rename kategori → semua produk yang menggunakannya ikut diperbarui otomatis

### 🥘 Bahan Baku & Resep
- CRUD bahan baku (nama, grup/brand, satuan, stok, stok minimum)
- Hubungkan produk ke bahan baku via resep (jumlah per porsi)
- **Bahan Tetap** (selalu terpakai) dan **Bahan Pilihan** (user pilih varian)
- Stok otomatis terpotong setiap transaksi selesai
- Stok porsi tampil langsung di kartu produk kasir
- Log perubahan stok lengkap (masuk, keluar, penyesuaian)
- Restok bahan baku dengan keterangan

### 📧 Notifikasi Email Stok Rendah
- Alert otomatis via email ketika stok bahan baku di bawah minimum
- Konfigurasi email tujuan di halaman **Pengaturan**
- Tidak ada batas waktu pengiriman — setiap cek stok yang rendah akan mengirim email
- Detail bahan baku yang hampir habis (nama, stok sisa, satuan, minimum)
- Tombol otorisasi MailApp tersedia di Pengaturan

### 📊 Laporan Penjualan
- Laporan **Harian, Mingguan, Bulanan**
- Statistik: Total Transaksi, Total Pendapatan, Rata-rata/Transaksi, Pajak, Diskon
- Breakdown metode pembayaran (Cash vs QRIS)
- Grafik pendapatan
- Laporan item terjual dengan filter kategori & periode
- Export CSV

### 🕐 Riwayat Transaksi
- Tampilan tabel lengkap dengan pagination
- Filter status (Selesai, Dibatalkan, Pesanan)
- Pencarian & filter tanggal
- Badge **"Hold"** pada transaksi yang dibayar beda hari dari pembuatan pesanan — menampilkan waktu pembayaran aktual
- Detail transaksi per item
- Hapus transaksi
- Stat cards: Total, Selesai, Dibatalkan, Omzet

### ⚙️ Pengaturan
- Nama toko, alamat, telepon
- Toggle pajak on/off + atur persentase
- Manajemen kasir (tambah, edit PIN, role, status)
- Konfigurasi email notifikasi stok rendah
- Otorisasi MailApp

### 👤 Sistem Login
- Multi-kasir dengan PIN masing-masing
- Role: Admin & Kasir
- Kasir dipilih saat proses transaksi

### 📱 Responsif
- Tampilan optimal di desktop & mobile
- Mobile: Floating cart button, overlay sidebar

---

## 🚀 Langkah Setup

### Langkah 1: Buat Google Sheet Baru

1. Buka [Google Sheets](https://sheets.google.com)
2. Klik **"+ Blank"** → beri nama **"Database Kasir"**

### Langkah 2: Buka Apps Script Editor

1. Klik menu **Extensions** → **Apps Script**
2. Hapus isi default di `Code.gs`

### Langkah 3: Buat Semua File

#### File `.gs` (Server-Side)

| No | Nama File | Deskripsi |
|----|-----------|-----------|
| 1 | `Code` | Main entry point & routing |
| 2 | `Utils` | Helper functions & konstanta sheet |
| 3 | `ProdukService` | CRUD produk & kategori dinamis |
| 4 | `TransaksiService` | Transaksi, pesanan hold, pembayaran |
| 5 | `LaporanService` | Generate laporan & statistik |
| 6 | `BahanBakuService` | CRUD bahan baku, resep, log stok |
| 7 | `OtorisasiHelper` | Notifikasi email stok rendah |

> ⚠️ Nama file **TANPA** ekstensi `.gs`

#### File `.html` (Client-Side)

| No | Nama File | Deskripsi |
|----|-----------|-----------|
| 1 | `Index` | Layout utama & shell SPA |
| 2 | `Stylesheet` | CSS & design system |
| 3 | `JavaScript` | Client-side logic |
| 4 | `HalamanKasir` | Halaman POS kasir |
| 5 | `HalamanProduk` | Manajemen produk |
| 6 | `HalamanKategori` | Manajemen kategori |
| 7 | `HalamanBahanBaku` | Manajemen bahan baku & log |
| 8 | `HalamanLaporan` | Laporan penjualan |
| 9 | `HalamanRiwayat` | Riwayat transaksi |
| 10 | `HalamanPengaturan` | Pengaturan & manajemen kasir |

> ⚠️ Nama file **TANPA** ekstensi `.html`

### Langkah 4: Copy-Paste Kode

Salin isi setiap file dari folder proyek ini ke file yang sesuai di GAS Editor.

**Urutan yang disarankan:**
1. `Utils.gs` → `Utils`
2. `Code.gs` → `Code`
3. `ProdukService.gs` → `ProdukService`
4. `TransaksiService.gs` → `TransaksiService`
5. `LaporanService.gs` → `LaporanService`
6. `BahanBakuService.gs` → `BahanBakuService`
7. `OtorisasiHelper.gs` → `OtorisasiHelper`
8. `Stylesheet.html` → `Stylesheet`
9. `JavaScript.html` → `JavaScript`
10. `HalamanKasir.html` → `HalamanKasir`
11. `HalamanProduk.html` → `HalamanProduk`
12. `HalamanKategori.html` → `HalamanKategori`
13. `HalamanBahanBaku.html` → `HalamanBahanBaku`
14. `HalamanLaporan.html` → `HalamanLaporan`
15. `HalamanRiwayat.html` → `HalamanRiwayat`
16. `HalamanPengaturan.html` → `HalamanPengaturan`
17. `Index.html` → `Index`

### Langkah 5: Inisialisasi Database

1. Pilih fungsi **`setupDatabase`** dari dropdown toolbar GAS Editor
2. Klik **▶ Run** → izinkan permission
3. Tunggu dialog **"✅ Database berhasil diinisialisasi!"**
4. Google Sheet akan memiliki **9 sheet** baru:
   - `Produk` — data produk
   - `Transaksi` — riwayat transaksi (+ kolom `Waktu_Pembayaran` di kolom O)
   - `DetailTransaksi` — detail item per transaksi
   - `Pengaturan` — konfigurasi toko
   - `Kasir` — data kasir (default: Admin, PIN: **1234**)
   - `BahanBaku` — data bahan baku
   - `ResepProduk` — resep per produk
   - `LogBahanBaku` — log perubahan stok
   - `Kategori` — master daftar kategori *(buat manual atau jalankan migrasi)*

### Langkah 6: Setup Kategori (Migrasi Data Lama)

Jika sheet `Kategori` belum ada atau ingin mengimpor kategori dari produk yang sudah ada:

1. Pilih fungsi **`migrasiKategoriDariProduk`** → klik **▶ Run**
2. Kategori unik dari kolom Produk akan diimpor ke sheet `Kategori`
3. Jalankan juga **`updateValidasiKategoriSheet`** untuk memperbarui dropdown di Sheets

### Langkah 7: Migrasi Kolom Waktu_Pembayaran (Jika Sheet Sudah Ada)

Jika sheet `Transaksi` sudah ada sebelum update ini:

1. Pilih fungsi **`tambahKolomWaktuPembayaran`** → klik **▶ Run**
2. Kolom ke-15 (`Waktu_Pembayaran`) akan ditambahkan otomatis dengan lebar 165px

### Langkah 8: Deploy sebagai Web App

1. Klik **Deploy** → **New deployment**
2. Ikon ⚙️ → **Web app**
3. Konfigurasi:
   - **Execute as**: Me
   - **Who has access**: Anyone
4. Klik **Deploy** → salin URL
5. Buka URL → Login: **Admin** / PIN: **1234**

---

## 📊 Struktur Database (Google Sheets)

### Sheet: `Produk`
| Kolom | Header | Keterangan |
|-------|--------|------------|
| A | ID_Produk | PRD-0001, dst |
| B | Nama_Produk | Nama produk |
| C | Kategori | Dropdown dinamis dari sheet Kategori |
| D | Harga | Harga jual (Rp), bisa 0 untuk produk gratis |
| E | Stok | Legacy — stok aktual dihitung dari ResepProduk |
| F | Deskripsi | Deskripsi singkat |
| G | URL_Gambar | URL gambar (Google Drive / URL langsung) |
| H | Status | Aktif / Nonaktif |
| I | Tanggal_Dibuat | Timestamp |
| J | Tanggal_Diupdate | Timestamp |

### Sheet: `Transaksi`
| Kolom | Header | Keterangan |
|-------|--------|------------|
| A | ID_Transaksi | TRX-0001, dst |
| B | Tanggal | Waktu pesanan dibuat / transaksi langsung |
| C | Nama_Pelanggan | Opsional |
| D | Jumlah_Item | Total qty |
| E | Subtotal | Sebelum pajak/diskon |
| F | Pajak | Jumlah pajak |
| G | Diskon | Jumlah diskon |
| H | Total | Total bayar |
| I | Jenis_Pembayaran | Cash / QRIS |
| J | Jumlah_Bayar | Uang diterima |
| K | Kembalian | Uang kembali |
| L | Status | Selesai / Pesanan / Dibatalkan |
| M | Kasir | Nama kasir |
| N | Catatan | Catatan tambahan |
| O | **Waktu_Pembayaran** | Timestamp saat pembayaran dilakukan (untuk pesanan hold) |

### Sheet: `Kategori` *(BARU)*
| Kolom | Header | Keterangan |
|-------|--------|------------|
| A | Nama_Kategori | Nama kategori (Makanan, Minuman, dll) |

> Master data kategori. Sumber untuk dropdown di form produk dan validasi kolom C sheet Produk.

### Sheet: `DetailTransaksi`
| Kolom | Header |
|-------|--------|
| A | ID_Detail |
| B | ID_Transaksi |
| C | ID_Produk |
| D | Nama_Produk |
| E | Harga_Satuan |
| F | Jumlah |
| G | Subtotal |
| H | Catatan_Item |

### Sheet: `Pengaturan`
| Key | Contoh Value |
|-----|-------------|
| nama_toko | Warkop Dharma |
| alamat | Jl. Contoh No. 1 |
| telepon | 0812-xxxx-xxxx |
| pajak_persen | 10 |
| pajak_aktif | Ya |
| notif_stok_email_aktif | Ya |
| notif_stok_email_tujuan | admin@email.com |

### Sheet: `Kasir`
| Kolom | Header | Keterangan |
|-------|--------|------------|
| A | ID_Kasir | KSR-0001, dst |
| B | Nama | Nama kasir |
| C | PIN | Min. 4 digit |
| D | Role | Admin / Kasir |
| E | Status | Aktif / Nonaktif |

### Sheet: `BahanBaku`
| Kolom | Header | Keterangan |
|-------|--------|------------|
| A | ID_Bahan | BHN-0001, dst |
| B | Nama_Bahan | Nama + varian (misal: "Nutrisari Jeruk") |
| C | Grup | Brand/kelompok (misal: "Nutrisari") |
| D | Stok | Jumlah stok saat ini |
| E | Satuan | saset / bungkus / gr / kg / ml / liter / pcs / botol / kaleng |
| F | Stok_Minimum | Alert email jika stok ≤ nilai ini |
| G | Keterangan | Catatan opsional |
| H | Status | Aktif / Nonaktif |
| I | Tanggal_Dibuat | Timestamp |
| J | Tanggal_Diupdate | Timestamp |

### Sheet: `ResepProduk`
| Kolom | Header | Keterangan |
|-------|--------|------------|
| A | ID_Resep | RSP-0001, dst |
| B | ID_Produk | FK → Produk |
| C | Nama_Produk | Denormalized |
| D | ID_Bahan | FK → BahanBaku |
| E | Nama_Bahan | Denormalized |
| F | Jumlah_Per_Porsi | Qty bahan per 1 porsi |
| G | Satuan | Satuan bahan |
| H | Tipe | `tetap` / `pilihan` |
| I | Grup_Pilihan | Grup varian (untuk Bahan Pilihan) |

### Sheet: `LogBahanBaku`
| Kolom | Header | Keterangan |
|-------|--------|------------|
| A | ID_Log | LBB-0001, dst |
| B | Tanggal | Timestamp perubahan |
| C | ID_Bahan | FK → BahanBaku |
| D | Nama_Bahan | Denormalized |
| E | Stok_Sebelum | Stok sebelum |
| F | Perubahan | + masuk / − keluar |
| G | Stok_Sesudah | Stok setelah |
| H | Tipe | Masuk / Keluar / Penyesuaian |
| I | Keterangan | Deskripsi |
| J | ID_Transaksi | Referensi transaksi (jika dari penjualan) |
| K | Diupdate_Oleh | Nama kasir/admin |

---

## 🏷️ Cara Kerja Kategori Dinamis

1. **Tambah Kategori** di halaman **Kategori** → tersimpan ke sheet `Kategori`
2. **Dropdown Produk** selalu fetch live dari backend — kategori baru langsung muncul
3. **Rename Kategori** → nama di sheet `Kategori` + semua kolom produk yang pakai nama lama diperbarui otomatis
4. **Hapus Kategori** → dihapus dari master, produk dipindah ke "Lainnya"
5. **Validasi Sheets** → kolom Kategori di sheet Produk otomatis diperbarui menggunakan dropdown yang mereferensi sheet `Kategori`

---

## 🕐 Cara Kerja Timestamp Pesanan Hold

Untuk pesanan yang dibayar di hari berbeda dari pembuatan:

| Kolom | Isi |
|-------|-----|
| `Tanggal` (kolom B) | Waktu pesanan **dibuat** |
| `Waktu_Pembayaran` (kolom O) | Waktu pembayaran **dilakukan** |

Di halaman Riwayat Transaksi:
- Transaksi langsung → tampil waktu transaksi biasa
- Pesanan hold yang dibayar beda hari → tampil **waktu pembayaran** + badge **"Hold"** (hover = lihat waktu pembuatan)

---

## 🥘 Cara Kerja Bahan Baku & Resep

1. **Tambah Bahan Baku** — isi nama, grup, satuan, stok awal, stok minimum
2. **Hubungkan ke Produk** — di form produk, tambahkan Bahan Tetap atau Bahan Pilihan beserta qty per porsi
3. **Stok Otomatis Terpotong** — setiap transaksi selesai, stok bahan berkurang sesuai resep
4. **Stok Porsi di Kasir** — kartu produk menampilkan berapa porsi bisa dibuat dari stok yang ada

### Contoh Resep Multi-Varian (Indomie)
- `BHN-0001` | Indomie Goreng Original | Grup: **Indomie** | 50 bungkus
- `BHN-0002` | Indomie Goreng Soto | Grup: **Indomie** | 40 bungkus

Produk "Indomie Goreng/Rebus" menggunakan **Bahan Pilihan** dengan Grup `Indomie` → user pilih rasa saat checkout.

---

## 📧 Notifikasi Email Stok Rendah

1. Di halaman **Pengaturan**, masukkan email tujuan dan aktifkan notifikasi
2. Klik **Otorisasi MailApp** → ikuti langkah izin
3. Setiap cek stok (saat transaksi selesai), jika ada bahan di bawah minimum → email terkirim otomatis
4. Email berisi daftar lengkap bahan yang hampir habis beserta stok sisa

---

## 🖼️ Gambar Produk dari Google Drive

**Format URL yang didukung:**
- `https://drive.google.com/uc?id=FILE_ID`
- `https://drive.google.com/file/d/FILE_ID/view`
- `https://drive.google.com/open?id=FILE_ID`

Sistem otomatis mengkonversi ke format thumbnail yang bisa ditampilkan di browser:
```
https://drive.google.com/thumbnail?id=FILE_ID&sz=w400
```

> Pastikan file Drive diset **"Anyone with the link can view"**

---

## ⚙️ Kustomisasi

### Mengubah Nama Toko
Halaman **Pengaturan** → isi nama toko, alamat, telepon → Simpan

### Mengubah Pajak
**Pengaturan** → toggle pajak on/off + atur persentase

### Menambah Kasir
**Pengaturan** → **Manajemen Kasir** → **Tambah Kasir** — isi nama, PIN, dan role

### Cetak Struk Thermal
1. Hubungkan printer thermal ke komputer & jadikan default printer
2. Setelah pembayaran, klik **Cetak Struk**
3. Format otomatis disesuaikan untuk kertas 80mm

---

## ❓ FAQ

**Q: Kenapa muncul error "Sheet not found"?**
A: Jalankan `setupDatabase()` dari GAS Editor terlebih dahulu.

**Q: Kategori baru tidak muncul di dropdown produk?**
A: Tambah kategori melalui halaman **Kategori** di aplikasi — bukan langsung di sheet. Dropdown akan otomatis update.

**Q: Kolom Waktu_Pembayaran tidak ada di sheet Transaksi?**
A: Jalankan fungsi `tambahKolomWaktuPembayaran()` dari GAS Editor sekali saja.

**Q: Gambar produk dari Drive tidak muncul?**
A: Pastikan file sudah diset "Anyone with the link can view" dan gunakan URL dalam format yang didukung.

**Q: Notifikasi email tidak terkirim?**
A: Jalankan **Otorisasi MailApp** dari halaman Pengaturan dan ikuti langkah izin. Pastikan email tujuan sudah diisi dan notifikasi diaktifkan.

**Q: Bagaimana cara update setelah deploy?**
A: GAS Editor → **Deploy** → **Manage deployments** → Edit → versi "New version" → Deploy

**Q: Apakah bisa diakses dari HP?**
A: Ya, tampilan responsif untuk mobile dengan floating cart button dan sidebar overlay.

**Q: Data tersimpan dimana?**
A: Semua data tersimpan di Google Sheets yang terhubung dengan Apps Script.

---

## 📝 Lisensi

Free to use. Dibuat dengan ❤️ menggunakan Google Apps Script.
