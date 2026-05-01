/**
 * =====================================================
 * UTILS.GS - Helper Functions
 * Sistem Kasir (POS) - Google Apps Script
 * =====================================================
 */

// ==================== KONSTANTA ====================
const SHEET_PRODUK = 'Produk';
const SHEET_TRANSAKSI = 'Transaksi';
const SHEET_DETAIL_TRANSAKSI = 'DetailTransaksi';
const SHEET_PENGATURAN = 'Pengaturan';
const SHEET_KASIR = 'Kasir';
const SHEET_BAHAN_BAKU = 'BahanBaku';
const SHEET_RESEP_PRODUK = 'ResepProduk';
const SHEET_LOG_BAHAN_BAKU = 'LogBahanBaku';
const SHEET_KATEGORI = 'Kategori'; // Master data kategori produk

// ==================== SPREADSHEET ACCESS ====================

/**
 * Mendapatkan Spreadsheet aktif
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Mendapatkan Sheet berdasarkan nama
 */
function getSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" tidak ditemukan. Jalankan setupDatabase() terlebih dahulu.');
  }
  return sheet;
}

/**
 * Mendapatkan Sheet dengan aman (tidak melempar error jika tidak ada)
 * @returns {Sheet|null}
 */
function getSheetSafe(sheetName) {
  var ss = getSpreadsheet();
  return ss.getSheetByName(sheetName) || null;
}

/**
 * Mendapatkan semua data dari sheet sebagai array of objects
 */
function getSheetData(sheetName) {
  var sheet = getSheet(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];
  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    row._rowIndex = i + 1;
    result.push(row);
  }
  
  return result;
}

/**
 * Mendapatkan data dari sheet dengan aman (kembalikan [] jika sheet tidak ada)
 */
function getSheetDataSafe(sheetName) {
  var sheet = getSheetSafe(sheetName);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];
  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    row._rowIndex = i + 1;
    result.push(row);
  }
  
  return result;
}

// ==================== ID GENERATOR ====================

/**
 * Generate ID unik dengan prefix
 * @param {string} prefix - Prefix ID (misal: PRD, TRX, LOG)
 * @param {string} settingKey - Key di sheet Pengaturan untuk counter
 * @returns {string} ID baru
 */
function generateId(prefix, settingKey) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    var lastId = parseInt(getSetting(settingKey)) || 0;
    var newId = lastId + 1;
    setSetting(settingKey, newId);
    
    var paddedNum = ('0000' + newId).slice(-4);
    return prefix + '-' + paddedNum;
  } finally {
    lock.releaseLock();
  }
}

// ==================== PENGATURAN ====================

/**
 * Mendapatkan nilai pengaturan
 */
function getSetting(key) {
  var sheet = getSheetSafe(SHEET_PENGATURAN);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

/**
 * Menyimpan nilai pengaturan
 */
function setSetting(key, value) {
  var sheet = getSheet(SHEET_PENGATURAN);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  // Key tidak ditemukan, tambah baris baru
  sheet.appendRow([key, value]);
}

/**
 * Mendapatkan semua pengaturan sebagai object
 */
function getAllSettings() {
  var sheet = getSheetSafe(SHEET_PENGATURAN);
  // Default settings jika sheet belum ada
  var defaults = {
    nama_toko: 'Toko Saya',
    alamat: '',
    telepon: '',
    pajak_persen: 10,
    pajak_aktif: 'Ya',
    notif_stok_email_aktif: 'Tidak',
    notif_stok_email_tujuan: '',
    notif_stok_email_last_sent: ''
  };
  if (!sheet) return defaults;
  var data = sheet.getDataRange().getValues();
  var settings = {};
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) settings[data[i][0]] = data[i][1];
  }
  // Merge defaults dengan data yang ada
  for (var key in defaults) {
    if (settings[key] === undefined) settings[key] = defaults[key];
  }
  return settings;
}

// ==================== FORMAT HELPERS ====================

/**
 * Format angka ke Rupiah
 */
function formatRupiah(angka) {
  var num = parseInt(angka) || 0;
  return 'Rp ' + num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

/**
 * Format tanggal ke string Indonesia
 */
function formatTanggal(date) {
  if (!date) return '-';
  try {
    var d = (date instanceof Date) ? date : new Date(date);
    if (isNaN(d.getTime())) return String(date);
    return Utilities.formatDate(d, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
  } catch(e) {
    return String(date);
  }
}

/**
 * Format tanggal singkat
 */
function formatTanggalSingkat(date) {
  if (!date) return '-';
  try {
    var d = (date instanceof Date) ? date : new Date(date);
    if (isNaN(d.getTime())) return String(date);
    return Utilities.formatDate(d, 'Asia/Jakarta', 'dd/MM/yyyy');
  } catch(e) {
    return String(date);
  }
}

/**
 * Mendapatkan tanggal hari ini (awal hari) dalam WIB
 */
function getToday() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

/**
 * Mendapatkan timestamp sekarang
 * Selalu return Date object — Google Sheets akan menyimpan dalam UTC,
 * tampilan di sheet tergantung timezone setting spreadsheet.
 * Untuk tampilan yang konsisten, pasang timezone spreadsheet ke Asia/Jakarta
 * melalui: File > Spreadsheet Settings > Time zone > (GMT+07:00) Jakarta
 */
function getNow() {
  return new Date();
}


// ==================== VALIDATION ====================

/**
 * Validasi data tidak kosong
 */
function validateRequired(data, fields) {
  var errors = [];
  for (var i = 0; i < fields.length; i++) {
    var field = fields[i];
    if (!data[field.key] && data[field.key] !== 0) {
      errors.push(field.label + ' wajib diisi');
    }
  }
  return errors;
}

// ==================== RESPONSE HELPERS ====================

/**
 * Response sukses
 */
function successResponse(data, message) {
  return {
    success: true,
    data: data,
    message: message || 'Berhasil'
  };
}

/**
 * Response error
 */
function errorResponse(message) {
  return {
    success: false,
    data: null,
    message: message || 'Terjadi kesalahan'
  };
}

/**
 * Parse daftar email dari string comma/semicolon separated.
 * @returns {Array<string>} email unik dan valid (lowercase)
 */
function parseEmailList(raw) {
  var text = String(raw || '').trim();
  if (!text) return [];
  var parts = text.split(/[;,]/);
  var result = [];
  var seen = {};
  for (var i = 0; i < parts.length; i++) {
    var email = String(parts[i] || '').trim().toLowerCase();
    if (!email) continue;
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) continue;
    if (seen[email]) continue;
    seen[email] = true;
    result.push(email);
  }
  return result;
}
