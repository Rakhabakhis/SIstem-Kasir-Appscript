# 🧾 Sistem Kasir (POS) - Google Apps Script

Sistem Point of Sale lengkap menggunakan Google Apps Script & Google Sheets sebagai database.

## ✨ Fitur

- **Halaman Kasir** - Tambah produk ke keranjang, proses pembayaran
- **Manajemen Produk** - Tambah, edit, hapus produk (CRUD)
- **Bahan Baku & Resep** - Kelola bahan baku, hubungkan ke produk dengan resep, stok otomatis terpotong saat transaksi
- **Laporan Transaksi** - Harian, mingguan, bulanan + breakdown pembayaran
- **Responsive** - Tampilan optimal di desktop & mobile
- **2 Metode Pembayaran** - Cash dan QRIS
- **Login System** - Multi-kasir dengan PIN, bisa tambah/edit kasir
- **Cetak Struk** - Print ke printer thermal (58mm/80mm)
- **Pengaturan Pajak** - Toggle on/off, atur persentase pajak
- **Gambar Produk** - Dari URL online atau Google Drive

---

## 🚀 Langkah-Langkah Setup

### Langkah 1: Buat Google Sheet Baru

1. Buka [Google Sheets](https://sheets.google.com)
2. Klik **"+ Blank"** untuk membuat spreadsheet baru
3. Beri nama: **"Database Kasir"**

### Langkah 2: Buka Apps Script Editor

1. Di Google Sheet, klik menu **Extensions** → **Apps Script**
2. Editor Apps Script akan terbuka di tab baru
3. Hapus isi yang ada di `Code.gs`

### Langkah 3: Buat Semua File

Di sisi kiri editor Apps Script, Anda perlu membuat file berikut:

#### File .gs (Server-Side)

Klik **"+"** → **Script** untuk membuat file .gs baru:

| No | Nama File | Deskripsi |
|----|-----------|-----------|
| 1 | `Code` | Main entry point (sudah ada, ganti isinya) |
| 2 | `Utils` | Helper functions |
| 3 | `ProdukService` | CRUD produk |
| 4 | `TransaksiService` | Proses transaksi |
| 5 | `LaporanService` | Generate laporan |
| 6 | `BahanBakuService` | CRUD bahan baku, resep, dan log |

> ⚠️ **PENTING**: Nama file TANPA ekstensi `.gs` (Apps Script otomatis menambahkan)

#### File .html (Client-Side)

Klik **"+"** → **HTML** untuk membuat file HTML baru:

| No | Nama File | Deskripsi |
|----|-----------|-----------|
| 1 | `Index` | Layout utama & shell SPA |
| 2 | `Stylesheet` | CSS styles |
| 3 | `JavaScript` | Client-side logic |
| 4 | `HalamanKasir` | Halaman POS |
| 5 | `HalamanProduk` | Manajemen produk |
| 6 | `HalamanBahanBaku` | Manajemen bahan baku & log |
| 7 | `HalamanLaporan` | Laporan transaksi |
| 8 | `HalamanPengaturan` | Pengaturan & manajemen kasir |

> ⚠️ **PENTING**: Nama file TANPA ekstensi `.html` (Apps Script otomatis menambahkan)

### Langkah 4: Copy-Paste Kode

Buka setiap file dari folder proyek ini dan copy-paste isinya ke file yang sesuai di Apps Script Editor.

**Urutan yang disarankan:**
1. `Utils.gs` → paste ke file `Utils`
2. `Code.gs` → paste ke file `Code`
3. `ProdukService.gs` → paste ke file `ProdukService`
4. `TransaksiService.gs` → paste ke file `TransaksiService`
5. `LaporanService.gs` → paste ke file `LaporanService`
6. `BahanBakuService.gs` → paste ke file `BahanBakuService`
7. `Stylesheet.html` → paste ke file `Stylesheet`
8. `JavaScript.html` → paste ke file `JavaScript`
9. `HalamanKasir.html` → paste ke file `HalamanKasir`
10. `HalamanProduk.html` → paste ke file `HalamanProduk`
11. `HalamanBahanBaku.html` → paste ke file `HalamanBahanBaku`
12. `HalamanLaporan.html` → paste ke file `HalamanLaporan`
13. `HalamanPengaturan.html` → paste ke file `HalamanPengaturan`
14. `Index.html` → paste ke file `Index`

### Langkah 5: Inisialisasi Database

1. Di Apps Script Editor, pilih fungsi **`setupDatabase`** dari dropdown di toolbar
2. Klik tombol **▶ Run**
3. Sistem akan minta izin - klik **Review Permissions** → **Allow**
4. Tunggu sampai muncul dialog "✅ Database berhasil diinisialisasi!"
5. Kembali ke Google Sheet - Anda akan melihat **8 sheet** baru:
   - Produk (10 produk contoh)
   - Transaksi
   - DetailTransaksi
   - Pengaturan
   - **Kasir** (1 admin default, PIN: **1234**)
   - **BahanBaku** (7 contoh bahan baku)
   - **ResepProduk**
   - **LogBahanBaku**

### Langkah 6: Deploy sebagai Web App

1. Di Apps Script Editor, klik **Deploy** → **New deployment**
2. Klik ikon ⚙️ → pilih **Web app**
3. Isi konfigurasi:
   - **Description**: Sistem Kasir POS
   - **Execute as**: **Me** (email Anda)
   - **Who has access**: **Anyone** (atau **Anyone with Google account** jika ingin lebih aman)
4. Klik **Deploy**
5. Copy URL yang muncul - itu adalah link aplikasi Anda!

### Langkah 7: Buka Aplikasi

1. Paste URL di browser
2. **Login** dengan nama: **Admin**, PIN: **1234**
3. Aplikasi kasir Anda siap digunakan! 🎉

---

## 📊 Struktur Database (Google Sheets)

### Sheet: Produk
| Kolom | Header | Keterangan |
|-------|--------|-------------|
| A | ID_Produk | PRD-0001, PRD-0002, dst |
| B | Nama_Produk | Nama produk |
| C | Kategori | Makanan / Minuman / Snack / Dessert / Paket Hemat |
| D | Harga | Harga jual (Rp) |
| E | Stok | Kolom legacy (stok aktual dihitung otomatis dari BahanBaku + ResepProduk) |
| F | Deskripsi | Deskripsi singkat |
| G | URL_Gambar | URL gambar produk |
| H | Status | Aktif / Nonaktif |
| I | Tanggal_Dibuat | Timestamp |
| J | Tanggal_Diupdate | Timestamp |

### Sheet: Transaksi
| Kolom | Header | Keterangan |
|-------|--------|-------------|
| A | ID_Transaksi | TRX-0001, dst |
| B | Tanggal | Timestamp transaksi |
| C | Nama_Pelanggan | Opsional |
| D | Jumlah_Item | Total qty |
| E | Subtotal | Sebelum pajak/diskon |
| F | Pajak | Jumlah pajak |
| G | Diskon | Jumlah diskon |
| H | Total | Total bayar |
| I | Jenis_Pembayaran | Cash / QRIS |
| J | Jumlah_Bayar | Uang diterima |
| K | Kembalian | Uang kembali |
| L | Status | Selesai/Pending/Dibatalkan |
| M | Kasir | Nama kasir |
| N | Catatan | Catatan tambahan |

### Sheet: DetailTransaksi
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

### Sheet: Pengaturan
| Key | Value |
|-----|-------|
| nama_toko | Toko Saya |
| alamat | Jl. Contoh No. 1 |
| telepon | 0812-xxxx-xxxx |
| pajak_persen | 10 |
| pajak_aktif | Ya |

### Sheet: Kasir
| Kolom | Header | Keterangan |
|-------|--------|-------------|
| A | ID_Kasir | KSR-0001, dst |
| B | Nama | Nama kasir |
| C | PIN | PIN login (min. 4 digit) |
| D | Role | Admin / Kasir |
| E | Status | Aktif / Nonaktif |

### Sheet: BahanBaku
| Kolom | Header | Keterangan |
|-------|--------|-------------|
| A | ID_Bahan | BHN-0001, dst |
| B | Nama_Bahan | Nama lengkap + varian (misal: "Nutrisari Jeruk") |
| C | Grup | Brand/kategori untuk pengelompokan (misal: "Nutrisari") |
| D | Stok | Jumlah stok saat ini |
| E | Satuan | kg / gr / ml / saset / bungkus / pcs / botol / kaleng |
| F | Stok_Minimum | Alert jika stok di bawah nilai ini |
| G | Keterangan | Catatan opsional |
| H | Status | Aktif / Nonaktif |
| I | Tanggal_Dibuat | Timestamp |
| J | Tanggal_Diupdate | Timestamp |

### Sheet: ResepProduk
| Kolom | Header | Keterangan |
|-------|--------|-------------|
| A | ID_Resep | RSP-0001, dst |
| B | ID_Produk | FK ke sheet Produk |
| C | Nama_Produk | Nama produk (denormalized) |
| D | ID_Bahan | FK ke sheet BahanBaku |
| E | Nama_Bahan | Nama bahan (denormalized) |
| F | Jumlah_Per_Porsi | Jumlah bahan yang dibutuhkan per 1 porsi produk |
| G | Satuan | Satuan sesuai bahan baku |

### Sheet: LogBahanBaku
| Kolom | Header | Keterangan |
|-------|--------|-------------|
| A | ID_Log | LBB-0001, dst |
| B | Tanggal | Timestamp perubahan |
| C | ID_Bahan | FK ke BahanBaku |
| D | Nama_Bahan | Nama bahan (denormalized) |
| E | Stok_Sebelum | Stok sebelum perubahan |
| F | Perubahan | Positif = masuk, negatif = keluar |
| G | Stok_Sesudah | Stok setelah perubahan |
| H | Tipe | Masuk / Keluar / Penyesuaian |
| I | Keterangan | Deskripsi perubahan |
| J | ID_Transaksi | Referensi transaksi (jika dari penjualan) |
| K | Diupdate_Oleh | Nama kasir/admin |

---

## 🥘 Fitur Bahan Baku

### Cara Kerja
1. **Tambah Bahan Baku** di menu "Bahan Baku" — isi nama, grup/brand, satuan, dan stok awal
2. **Hubungkan ke Produk** — saat edit produk, tambahkan bahan baku ke bagian "Bahan Baku & Resep" beserta jumlah per porsi
3. **Stok Otomatis Terpotong** — setiap transaksi, stok bahan baku berkurang sesuai resep
4. **Stok Porsi Tampil di Kasir** — produk dengan resep menampilkan berapa porsi yang bisa dibuat

### Contoh Resep
| Produk | Bahan Baku | Jumlah Per Porsi | Artinya |
|--------|-----------|-----------------|---------|
| Indomie Goreng | Indomie Goreng Original | 1 bungkus | 1 transaksi = -1 bungkus |
| Indomie Double | Indomie Goreng Original | 2 bungkus | 1 transaksi = -2 bungkus |
| Nutrisari Jeruk | Nutrisari Jeruk | 1 saset | 1 transaksi = -1 saset |

### Rekomendasi Multi-Varian (contoh Nutrisari)
Setiap varian dibuat sebagai bahan baku terpisah dengan kolom **Grup** yang sama:
- `BHN-0003` | Nutrisari Jeruk | Grup: **Nutrisari** | 100 saset
- `BHN-0004` | Nutrisari Mangga | Grup: **Nutrisari** | 80 saset

---

## ⚙️ Kustomisasi

### Mengubah Nama Toko
Edit nilai di sheet **Pengaturan** baris `nama_toko`

### Mengubah Pajak
Bisa langsung dari halaman **Pengaturan** di aplikasi:
- Toggle pajak on/off
- Atur persentase pajak

### Menambah Kasir Baru
Buka halaman **Pengaturan** → **Manajemen Kasir** → **Tambah Kasir**
- Setiap kasir punya PIN sendiri untuk login
- Kasir bisa dipilih saat melakukan transaksi

### Cetak Struk Thermal
1. Pastikan printer thermal sudah terhubung ke komputer
2. Set printer thermal sebagai default printer
3. Setelah pembayaran, klik tombol **Cetak Struk**
4. Format struk otomatis disesuaikan untuk kertas 80mm

### Gambar Produk
Gunakan URL gambar dari:
- **Google Drive**: Upload gambar → Share → Copy link → Gunakan format `https://drive.google.com/uc?id=FILE_ID`
- **URL Online**: URL langsung ke file gambar dari website manapun

---

## ❓ FAQ

**Q: Kenapa muncul error "Sheet not found"?**
A: Jalankan fungsi `setupDatabase()` terlebih dahulu.

**Q: Bagaimana cara update setelah deploy?**
A: Di Apps Script → Deploy → Manage deployments → Edit → Pilih version "New version" → Deploy

**Q: Apakah bisa diakses dari HP?**
A: Ya! Tampilan sudah responsive untuk mobile.

**Q: Data tersimpan dimana?**
A: Semua data tersimpan di Google Sheets yang sama dengan Apps Script.

---

## 📝 Lisensi

Free to use. Dibuat dengan ❤️ menggunakan Google Apps Script.
