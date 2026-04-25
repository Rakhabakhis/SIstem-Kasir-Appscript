/**
 * =====================================================
 * CODE.GS - Main Entry Point & Routing
 * Sistem Kasir (POS) - Google Apps Script
 * =====================================================
 * 
 * SETUP:
 * 1. Buat Google Sheet baru
 * 2. Buka Extensions > Apps Script
 * 3. Copy semua file .gs dan .html
 * 4. Jalankan setupDatabase() sekali untuk inisialisasi
 * 5. Deploy > New Deployment > Web App
 */

// ==================== WEB APP ENTRY POINT ====================

/**
 * Handler GET request - Menampilkan halaman web app
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  return template.evaluate()
    .setTitle('Sistem Kasir - POS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3144/3144456.png');
}

/**
 * Include HTML partial file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================== DATABASE SETUP ====================

/**
 * Inisialisasi database - JALANKAN SEKALI SAJA!
 * Membuat semua sheet dengan header dan data awal
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // === Sheet 1: Produk ===
  var sheetProduk = getOrCreateSheet(ss, SHEET_PRODUK);
  if (sheetProduk.getLastRow() === 0) {
    sheetProduk.appendRow([
      'ID_Produk', 'Nama_Produk', 'Kategori', 'Harga', 'Stok',
      'Deskripsi', 'URL_Gambar', 'Status', 'Tanggal_Dibuat', 'Tanggal_Diupdate'
    ]);

    // Format header
    formatHeader(sheetProduk);
    sheetProduk.setColumnWidth(1, 100);
    sheetProduk.setColumnWidth(2, 180);
    sheetProduk.setColumnWidth(3, 120);
    sheetProduk.setColumnWidth(4, 100);
    sheetProduk.setColumnWidth(5, 70);
    sheetProduk.setColumnWidth(6, 200);
    sheetProduk.setColumnWidth(7, 300);
  }
  
  // === Sheet 2: Transaksi ===
  var sheetTransaksi = getOrCreateSheet(ss, SHEET_TRANSAKSI);
  if (sheetTransaksi.getLastRow() === 0) {
    sheetTransaksi.appendRow([
      'ID_Transaksi', 'Tanggal', 'Nama_Pelanggan', 'Jumlah_Item',
      'Subtotal', 'Pajak', 'Diskon', 'Total',
      'Jenis_Pembayaran', 'Jumlah_Bayar', 'Kembalian', 'Status', 'Kasir', 'Catatan'
    ]);
    formatHeader(sheetTransaksi);
  }
  
  // === Sheet 3: DetailTransaksi ===
  var sheetDetail = getOrCreateSheet(ss, SHEET_DETAIL_TRANSAKSI);
  if (sheetDetail.getLastRow() === 0) {
    sheetDetail.appendRow([
      'ID_Detail', 'ID_Transaksi', 'ID_Produk', 'Nama_Produk',
      'Harga_Satuan', 'Jumlah', 'Subtotal', 'Catatan_Item'
    ]);
    formatHeader(sheetDetail);
  }
  
  // === Sheet 4: Pengaturan ===
  var sheetPengaturan = getOrCreateSheet(ss, SHEET_PENGATURAN);
  if (sheetPengaturan.getLastRow() === 0) {
    sheetPengaturan.appendRow(['Key', 'Value']);
    
    var settings = [
      ['nama_toko', 'Toko Saya'],
      ['alamat', 'Jl. Contoh No. 1'],
      ['telepon', '0812-xxxx-xxxx'],
      ['pajak_persen', 10],
      ['pajak_aktif', 'Ya'],
      ['id_transaksi_terakhir', 0],
      ['id_produk_terakhir', 0],
      ['id_detail_terakhir', 0],
      ['id_kasir_terakhir', 1]
    ];
    
    for (var i = 0; i < settings.length; i++) {
      sheetPengaturan.appendRow(settings[i]);
    }
    formatHeader(sheetPengaturan);
    sheetPengaturan.setColumnWidth(1, 200);
    sheetPengaturan.setColumnWidth(2, 250);
  }
  
  // === Sheet 5: Kasir ===
  var sheetKasir = getOrCreateSheet(ss, SHEET_KASIR);
  if (sheetKasir.getLastRow() === 0) {
    sheetKasir.appendRow(['ID_Kasir', 'Nama', 'PIN', 'Role', 'Status']);
    sheetKasir.appendRow(['KSR-0001', 'Admin', '1234', 'Admin', 'Aktif']);
    formatHeader(sheetKasir);
    sheetKasir.setColumnWidth(1, 100);
    sheetKasir.setColumnWidth(2, 180);
    sheetKasir.setColumnWidth(3, 80);
    sheetKasir.setColumnWidth(4, 100);
    sheetKasir.setColumnWidth(5, 80);
  }
  
  // === Sheet 6: BahanBaku ===
  var sheetBahanBaku = getOrCreateSheet(ss, SHEET_BAHAN_BAKU);
  if (sheetBahanBaku.getLastRow() === 0) {
    sheetBahanBaku.appendRow([
      'ID_Bahan', 'Nama_Bahan', 'Grup', 'Stok', 'Satuan',
      'Stok_Minimum', 'Keterangan', 'Status', 'Tanggal_Dibuat', 'Tanggal_Diupdate'
    ]);
    var sampleBahan = [
      ['BHN-0001', 'Indomie Goreng Original', 'Indomie', 50, 'bungkus', 10, 'Mie goreng instan', 'Aktif', new Date(), new Date()],
      ['BHN-0002', 'Indomie Soto Ayam', 'Indomie', 30, 'bungkus', 5, 'Mie soto instan', 'Aktif', new Date(), new Date()],
      ['BHN-0003', 'Nutrisari Jeruk', 'Nutrisari', 100, 'saset', 20, 'Minuman serbuk rasa jeruk', 'Aktif', new Date(), new Date()],
      ['BHN-0004', 'Nutrisari Mangga', 'Nutrisari', 80, 'saset', 20, 'Minuman serbuk rasa mangga', 'Aktif', new Date(), new Date()],
      ['BHN-0005', 'Gula Pasir', 'Bahan Dasar', 5000, 'gr', 500, 'Gula pasir putih', 'Aktif', new Date(), new Date()],
      ['BHN-0006', 'Kopi Bubuk', 'Bahan Dasar', 2000, 'gr', 300, 'Kopi robusta', 'Aktif', new Date(), new Date()],
      ['BHN-0007', 'Susu UHT', 'Bahan Dasar', 10000, 'ml', 1000, 'Susu full cream', 'Aktif', new Date(), new Date()]
    ];
    for (var i = 0; i < sampleBahan.length; i++) {
      sheetBahanBaku.appendRow(sampleBahan[i]);
    }
    formatHeader(sheetBahanBaku);
    sheetBahanBaku.setColumnWidth(1, 100);
    sheetBahanBaku.setColumnWidth(2, 200);
    sheetBahanBaku.setColumnWidth(3, 130);
  }

  // === Sheet 7: ResepProduk ===
  var sheetResep = getOrCreateSheet(ss, SHEET_RESEP_PRODUK);
  if (sheetResep.getLastRow() === 0) {
    sheetResep.appendRow([
      'ID_Resep', 'ID_Produk', 'Nama_Produk', 'ID_Bahan', 'Nama_Bahan',
      'Jumlah_Per_Porsi', 'Satuan', 'Tipe_Bahan', 'Grup_Pilihan'
    ]);
    formatHeader(sheetResep);
    sheetResep.setColumnWidth(1, 100);
    sheetResep.setColumnWidth(2, 100);
    sheetResep.setColumnWidth(3, 180);
    sheetResep.setColumnWidth(4, 100);
    sheetResep.setColumnWidth(5, 200);
    sheetResep.setColumnWidth(8, 100);
    sheetResep.setColumnWidth(9, 130);
  } else {
    // Migration: tambah kolom baru jika belum ada
    migrasiResepKolom(sheetResep);
  }

  // === Sheet 8: LogBahanBaku ===
  var sheetLogBahan = getOrCreateSheet(ss, SHEET_LOG_BAHAN_BAKU);
  if (sheetLogBahan.getLastRow() === 0) {
    sheetLogBahan.appendRow([
      'ID_Log', 'Tanggal', 'ID_Bahan', 'Nama_Bahan',
      'Stok_Sebelum', 'Perubahan', 'Stok_Sesudah',
      'Tipe', 'Keterangan', 'ID_Transaksi', 'Diupdate_Oleh'
    ]);
    formatHeader(sheetLogBahan);
  }

  // Tambah counter ID baru ke Pengaturan jika belum ada
  if (!getSetting('id_bahan_terakhir')) setSetting('id_bahan_terakhir', 7);
  if (!getSetting('id_resep_terakhir')) setSetting('id_resep_terakhir', 0);
  if (!getSetting('id_log_bahan_terakhir')) setSetting('id_log_bahan_terakhir', 0);

  // Hapus Sheet1 default jika ada
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // Hapus sheet legacy yang sudah tidak digunakan
  cleanupLegacySheets(ss);
  
  // Set timezone ke WIB (Asia/Jakarta) agar tanggal di sheet sesuai
  fixTimezone();
  
  var setupMessage = '✅ Database berhasil diinisialisasi!\n\nSheet yang dibuat:\n- Produk\n- Transaksi\n- DetailTransaksi\n- Pengaturan\n- Kasir (1 admin default, PIN: 1234)\n- BahanBaku (7 contoh bahan)\n- ResepProduk\n- LogBahanBaku\n\n⏰ Timezone spreadsheet sudah diset ke WIB (Asia/Jakarta).\n\nSekarang Anda bisa deploy sebagai Web App.';
  try {
    // Bisa gagal jika dipanggil dari context tanpa UI (mis. web app / trigger)
    SpreadsheetApp.getUi().alert(setupMessage);
  } catch (uiErr) {
    Logger.log(setupMessage);
    Logger.log('setupDatabase info: UI alert dilewati - ' + uiErr.message);
  }
}

/**
 * Hapus sheet lama/legacy yang sudah tidak dipakai
 */
function cleanupLegacySheets(ss) {
  var legacyNames = ['logstok', 'LogStok', 'LOGSTOK', 'Log_Stok'];
  for (var i = 0; i < legacyNames.length; i++) {
    var sheet = ss.getSheetByName(legacyNames[i]);
    if (sheet && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  }
}

/**
 * =====================================================
 * PERBAIKI TIMEZONE SPREADSHEET KE WIB (ASIA/JAKARTA)
 * Jalankan fungsi ini sekali dari editor Apps Script
 * jika waktu di kolom Tanggal pada sheet tidak sesuai WIB.
 * =====================================================
 */
function fixTimezone() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setSpreadsheetTimeZone('Asia/Jakarta');
    Logger.log('✅ Timezone spreadsheet berhasil diset ke Asia/Jakarta (WIB)');
  } catch(e) {
    Logger.log('❌ Gagal set timezone: ' + e.message);
  }
}


/**
 * Migrasi: tambah kolom Tipe_Bahan dan Grup_Pilihan ke ResepProduk jika belum ada
 * Aman dijalankan berulang kali
 */
function migrasiResepKolom(sheet) {
  try {
    if (!sheet) sheet = getSheetSafe(SHEET_RESEP_PRODUK);
    if (!sheet || sheet.getLastRow() === 0) return;
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var hasTipe = false;
    var hasGrup = false;
    for (var i = 0; i < headerRow.length; i++) {
      if (headerRow[i] === 'Tipe_Bahan') hasTipe = true;
      if (headerRow[i] === 'Grup_Pilihan') hasGrup = true;
    }
    if (!hasTipe) {
      var col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue('Tipe_Bahan');
      // Fill existing rows with 'Tetap' as default
      if (sheet.getLastRow() > 1) {
        var fill = [];
        for (var r = 2; r <= sheet.getLastRow(); r++) fill.push(['Tetap']);
        sheet.getRange(2, col, fill.length, 1).setValues(fill);
      }
      formatHeader(sheet);
    }
    if (!hasGrup) {
      var col2 = sheet.getLastColumn() + 1;
      sheet.getRange(1, col2).setValue('Grup_Pilihan');
      formatHeader(sheet);
    }
  } catch (e) {
    Logger.log('migrasiResepKolom warning: ' + e.message);
  }
}

/**
 * Mendapatkan atau membuat sheet
 */
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Format header row
 */
function formatHeader(sheet) {
  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setBackground('#0D9488');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

// ==================== INIT DATA FOR FRONTEND ====================

/**
 * Mendapatkan data inisialisasi untuk frontend
 */
function getInitData() {
  try {
    var settings = getAllSettings();
    var produk = getProduk();
    var kasirList = getKasirListInternal();
    
    // Kategori dari kolom Kategori di sheet Produk saja
    var kategoriSet = {};
    for (var i = 0; i < produk.length; i++) {
      if (produk[i].kategori) kategoriSet[produk[i].kategori] = true;
    }
    var kategoriList = ['Semua Produk'].concat(Object.keys(kategoriSet).sort());

    // Bahan baku aktif
    var bahanRaw = getSheetDataSafe(SHEET_BAHAN_BAKU);
    var bahanBakuAktif = [];
    for (var b = 0; b < bahanRaw.length; b++) {
      if (bahanRaw[b].Status === 'Aktif') {
        bahanBakuAktif.push({
          id: bahanRaw[b].ID_Bahan,
          nama: bahanRaw[b].Nama_Bahan,
          grup: bahanRaw[b].Grup || '',
          stok: Number(bahanRaw[b].Stok) || 0,
          satuan: bahanRaw[b].Satuan || '',
          stokMinimum: Number(bahanRaw[b].Stok_Minimum) || 0
        });
      }
    }

    // Grup unik dari bahan baku
    var grupSet = {};
    for (var g = 0; g < bahanBakuAktif.length; g++) {
      if (bahanBakuAktif[g].grup) grupSet[bahanBakuAktif[g].grup] = true;
    }
    var grupList = Object.keys(grupSet).sort();

    // Kalkulasi stok dari bahan baku (mendukung Tetap dan Pilihan)
    var stokResult = getStokDariResepSemuaProduk();
    for (var j = 0; j < produk.length; j++) {
      var idP = produk[j].id;
      produk[j].stokDariBahan = stokResult.stok[idP] !== undefined ? stokResult.stok[idP] : 0;
      produk[j].pakaiResep = true;
    }

    // Peta bahan Pilihan per produk (untuk variant picker di kasir)
    // {idProduk: [{grupPilihan, jumlahPerPorsi, satuan}]}
    var resepPilihan = stokResult.pilihanPerProduk || {};

    // Hitung jumlah pesanan pending
    var pesananCount = 0;
    var trxRaw = getSheetDataSafe(SHEET_TRANSAKSI);
    for (var p = 0; p < trxRaw.length; p++) {
      if (trxRaw[p].Status === 'Pesanan') pesananCount++;
    }

    return successResponse({
      settings: settings,
      produk: produk,
      kategori: kategoriList,
      kasirList: kasirList,
      bahanBaku: bahanBakuAktif,
      grupList: grupList,
      resepPilihan: resepPilihan,
      pesananPendingCount: pesananCount
    });
  } catch (e) {
    return errorResponse('Gagal memuat data: ' + e.message);
  }
}

// ==================== KASIR / LOGIN ====================

function loginKasir(nama, pin) {
  try {
    // Coba baca dari sheet Kasir (aman, tidak crash jika sheet belum ada)
    var data = getSheetDataSafe(SHEET_KASIR);
    
    // Jika sheet belum ada atau kosong, gunakan akun admin default
    if (!data || data.length === 0) {
      if (nama === 'Admin' && String(pin) === '1234') {
        return successResponse({ id: 'KSR-0001', nama: 'Admin', role: 'Admin' }, 'Login berhasil! Selamat datang, Admin');
      }
      return errorResponse('Sheet Kasir belum dibuat. Jalankan setupDatabase() atau login sebagai Admin / PIN 1234');
    }
    
    for (var i = 0; i < data.length; i++) {
      if (data[i].Nama === nama && String(data[i].PIN) === String(pin) && data[i].Status === 'Aktif') {
        return successResponse({
          id: data[i].ID_Kasir,
          nama: data[i].Nama,
          role: data[i].Role
        }, 'Login berhasil! Selamat datang, ' + data[i].Nama);
      }
    }
    return errorResponse('Nama atau PIN salah');
  } catch (e) {
    // Fallback: jika ada error apapun, izinkan login admin default
    if (nama === 'Admin' && String(pin) === '1234') {
      return successResponse({ id: 'KSR-0001', nama: 'Admin', role: 'Admin' }, 'Login berhasil (mode darurat)');
    }
    return errorResponse('Gagal login: ' + e.message);
  }
}

function getKasirListInternal() {
  var data = getSheetDataSafe(SHEET_KASIR);
  // Jika sheet belum ada, kembalikan admin default
  if (!data || data.length === 0) {
    return [{ id: 'KSR-0001', nama: 'Admin', role: 'Admin', status: 'Aktif' }];
  }
  var result = [];
  for (var i = 0; i < data.length; i++) {
    result.push({
      id: data[i].ID_Kasir,
      nama: data[i].Nama,
      role: data[i].Role,
      status: data[i].Status,
      _rowIndex: data[i]._rowIndex
    });
  }
  return result;
}

function getKasirList() {
  try {
    return successResponse(getKasirListInternal());
  } catch (e) {
    return successResponse([{ id: 'KSR-0001', nama: 'Admin', role: 'Admin', status: 'Aktif' }]);
  }
}

function getKasirAktif() {
  try {
    var data = getSheetDataSafe(SHEET_KASIR);
    // Jika sheet belum ada, kembalikan admin default
    if (!data || data.length === 0) {
      return successResponse([{ id: 'KSR-0001', nama: 'Admin' }]);
    }
    var result = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i].Status === 'Aktif') {
        result.push({ id: data[i].ID_Kasir, nama: data[i].Nama });
      }
    }
    // Jika tidak ada kasir aktif, tambah admin default
    if (result.length === 0) {
      result.push({ id: 'KSR-0001', nama: 'Admin' });
    }
    return successResponse(result);
  } catch (e) {
    // Selalu kembalikan minimal satu kasir
    return successResponse([{ id: 'KSR-0001', nama: 'Admin' }]);
  }
}

function tambahKasir(data) {
  try {
    if (!data.nama || !data.pin) return errorResponse('Nama dan PIN wajib diisi');
    if (String(data.pin).length < 4) return errorResponse('PIN minimal 4 digit');
    var sheet = getSheet(SHEET_KASIR);
    var idBaru = generateId('KSR', 'id_kasir_terakhir');
    sheet.appendRow([idBaru, data.nama, String(data.pin), data.role || 'Kasir', 'Aktif']);
    return successResponse({ id: idBaru }, 'Kasir "' + data.nama + '" berhasil ditambahkan');
  } catch (e) {
    return errorResponse('Gagal tambah kasir: ' + e.message);
  }
}

function editKasir(data) {
  try {
    if (!data.id) return errorResponse('ID Kasir tidak valid');
    var sheet = getSheet(SHEET_KASIR);
    var allData = getSheetData(SHEET_KASIR);
    for (var i = 0; i < allData.length; i++) {
      if (allData[i].ID_Kasir === data.id) {
        var row = allData[i]._rowIndex;
        if (data.nama) sheet.getRange(row, 2).setValue(data.nama);
        if (data.pin) sheet.getRange(row, 3).setValue(String(data.pin));
        if (data.role) sheet.getRange(row, 4).setValue(data.role);
        if (data.status) sheet.getRange(row, 5).setValue(data.status);
        return successResponse(null, 'Kasir berhasil diupdate');
      }
    }
    return errorResponse('Kasir tidak ditemukan');
  } catch (e) {
    return errorResponse('Gagal edit kasir: ' + e.message);
  }
}

function hapusKasir(id) {
  try {
    var sheet = getSheet(SHEET_KASIR);
    var data = getSheetData(SHEET_KASIR);
    for (var i = 0; i < data.length; i++) {
      if (data[i].ID_Kasir === id) {
        sheet.getRange(data[i]._rowIndex, 5).setValue('Nonaktif');
        return successResponse(null, 'Kasir dinonaktifkan');
      }
    }
    return errorResponse('Kasir tidak ditemukan');
  } catch (e) {
    return errorResponse('Gagal hapus kasir: ' + e.message);
  }
}

function savePengaturan(settings) {
  try {
    if (settings.nama_toko !== undefined) setSetting('nama_toko', settings.nama_toko);
    if (settings.alamat !== undefined) setSetting('alamat', settings.alamat);
    if (settings.telepon !== undefined) setSetting('telepon', settings.telepon);
    if (settings.pajak_aktif !== undefined) setSetting('pajak_aktif', settings.pajak_aktif);
    if (settings.pajak_persen !== undefined) setSetting('pajak_persen', Number(settings.pajak_persen));
    return successResponse(null, 'Pengaturan berhasil disimpan');
  } catch (e) {
    return errorResponse('Gagal simpan: ' + e.message);
  }
}
