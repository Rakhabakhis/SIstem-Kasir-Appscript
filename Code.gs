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
    
    // Data contoh produk
    var sampleProduk = [
      ['PRD-0001', 'Classic Beef Burger', 'Makanan', 35000, 24,
       'Daging sapi asli, keju, sayur', 'https://images.unsplash.com/photo-1568901346375-23c9450c58cd?w=300&h=300&fit=crop', 'Aktif', new Date(), new Date()],
      ['PRD-0002', 'Double Cheese Burger', 'Makanan', 55000, 18,
       'Ekstra keju lumer', 'https://images.unsplash.com/photo-1553979459-d2229ba7433b?w=300&h=300&fit=crop', 'Aktif', new Date(), new Date()],
      ['PRD-0003', 'French Fries Large', 'Snack', 20000, 50,
       'Kentang goreng renyah', 'https://images.unsplash.com/photo-1573080496219-bb080dd4f877?w=300&h=300&fit=crop', 'Aktif', new Date(), new Date()],
      ['PRD-0004', 'Iced Caramel Latte', 'Minuman', 28000, 45,
       'Kopi susu karamel dingin', 'https://images.unsplash.com/photo-1461023058943-07fcbe16d735?w=300&h=300&fit=crop', 'Aktif', new Date(), new Date()],
      ['PRD-0005', 'Iced Lemon Tea', 'Minuman', 15000, 60,
       'Teh lemon segar dingin', 'https://images.unsplash.com/photo-1556679343-c7306c1976bc?w=300&h=300&fit=crop', 'Aktif', new Date(), new Date()],
      ['PRD-0006', 'Iced Americano', 'Minuman', 22000, 40,
       'Kopi hitam dingin', 'https://images.unsplash.com/photo-1517701550927-30cf4ba1dba5?w=300&h=300&fit=crop', 'Aktif', new Date(), new Date()],
      ['PRD-0007', 'Chicken Wings (6 pcs)', 'Makanan', 32000, 30,
       'Sayap ayam goreng krispy', 'https://images.unsplash.com/photo-1567620832903-9fc6debc209f?w=300&h=300&fit=crop', 'Aktif', new Date(), new Date()],
      ['PRD-0008', 'Chocolate Milkshake', 'Minuman', 25000, 35,
       'Milkshake coklat premium', 'https://images.unsplash.com/photo-1572490122747-3968b75cc699?w=300&h=300&fit=crop', 'Aktif', new Date(), new Date()],
      ['PRD-0009', 'Brownies', 'Dessert', 18000, 20,
       'Brownies coklat lembut', 'https://images.unsplash.com/photo-1606313564200-e75d5e30476c?w=300&h=300&fit=crop', 'Aktif', new Date(), new Date()],
      ['PRD-0010', 'Paket Hemat 1', 'Paket Hemat', 45000, 99,
       'Burger + Fries + Drink', 'https://images.unsplash.com/photo-1594212699903-ec8a3eca50f5?w=300&h=300&fit=crop', 'Aktif', new Date(), new Date()]
    ];
    
    for (var i = 0; i < sampleProduk.length; i++) {
      sheetProduk.appendRow(sampleProduk[i]);
    }
    
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
  
  // === Sheet 4: LogStok ===
  var sheetLog = getOrCreateSheet(ss, SHEET_LOG_STOK);
  if (sheetLog.getLastRow() === 0) {
    sheetLog.appendRow([
      'ID_Log', 'Tanggal', 'ID_Produk', 'Nama_Produk',
      'Stok_Sebelum', 'Perubahan', 'Stok_Sesudah', 'Tipe', 'Keterangan', 'Diupdate_Oleh'
    ]);
    formatHeader(sheetLog);
  }
  
  // === Sheet 5: Pengaturan ===
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
      ['id_produk_terakhir', 10],
      ['id_log_terakhir', 0],
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
  
  // === Sheet 6: Kasir ===
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
  
  // Hapus Sheet1 default jika ada
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
  
  // Set timezone ke WIB (Asia/Jakarta) agar tanggal di sheet sesuai
  fixTimezone();
  
  SpreadsheetApp.getUi().alert('✅ Database berhasil diinisialisasi!\n\nSheet yang dibuat:\n- Produk (10 produk contoh)\n- Transaksi\n- DetailTransaksi\n- LogStok\n- Pengaturan\n- Kasir (1 admin default, PIN: 1234)\n\n⏰ Timezone spreadsheet sudah diset ke WIB (Asia/Jakarta).\n\nSekarang Anda bisa deploy sebagai Web App.');
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
    
    return successResponse({
      settings: settings,
      produk: produk,
      kategori: kategoriList,
      kasirList: kasirList
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
