/**
 * =====================================================
 * PRODUK SERVICE - CRUD Produk
 * Sistem Kasir (POS) - Google Apps Script
 * =====================================================
 */

/**
 * Mendapatkan semua produk aktif
 * @returns {Object} Response dengan array produk
 */
function getProduk() {
  try {
    var data = getSheetDataSafe(SHEET_PRODUK);
    var stokResult = getStokDariResepSemuaProduk();
    var stokMap = stokResult.stok || {};
    var produkAktif = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i].Status === 'Aktif') {
        var stokOtomatis = stokMap[data[i].ID_Produk];
        produkAktif.push({
          id: data[i].ID_Produk,
          nama: data[i].Nama_Produk,
          kategori: data[i].Kategori,
          harga: Number(data[i].Harga),
          stok: stokOtomatis !== undefined ? Number(stokOtomatis) : 0,
          stokDariBahan: stokOtomatis !== undefined ? Number(stokOtomatis) : 0,
          deskripsi: data[i].Deskripsi,
          gambar: data[i].URL_Gambar,
          status: data[i].Status,
          pakaiResep: true,
          _rowIndex: data[i]._rowIndex
        });
      }
    }
    
    return produkAktif; // Tetap return array (dipakai internal di getInitData)
  } catch (e) {
    Logger.log('Error getProduk: ' + e.message);
    return [];
  }
}

/**
 * Mendapatkan semua produk (termasuk nonaktif) untuk manajemen
 */
function getAllProdukAdmin() {
  try {
    var data = getSheetDataSafe(SHEET_PRODUK);
    var stokResult = getStokDariResepSemuaProduk();
    var stokMap = stokResult.stok || {};
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var stokOtomatis = stokMap[data[i].ID_Produk];
      result.push({
        id: data[i].ID_Produk,
        nama: data[i].Nama_Produk,
        kategori: data[i].Kategori,
        harga: Number(data[i].Harga),
        stok: stokOtomatis !== undefined ? Number(stokOtomatis) : 0,
        stokDariBahan: stokOtomatis !== undefined ? Number(stokOtomatis) : 0,
        deskripsi: data[i].Deskripsi,
        gambar: data[i].URL_Gambar,
        status: data[i].Status,
        pakaiResep: true,
        tanggalDibuat: data[i].Tanggal_Dibuat ? String(data[i].Tanggal_Dibuat) : '',
        tanggalDiupdate: data[i].Tanggal_Diupdate ? String(data[i].Tanggal_Diupdate) : '',
        _rowIndex: data[i]._rowIndex
      });
    }
    
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal memuat produk: ' + e.message);
  }
}

/**
 * Mendapatkan produk berdasarkan ID
 */
function getProdukById(id) {
  try {
    var data = getSheetData(SHEET_PRODUK);
    var stokResult = getStokDariResepSemuaProduk();
    var stokMap = stokResult.stok || {};
    
    for (var i = 0; i < data.length; i++) {
      if (data[i].ID_Produk === id) {
        var stokOtomatis = stokMap[data[i].ID_Produk];
        return successResponse({
          id: data[i].ID_Produk,
          nama: data[i].Nama_Produk,
          kategori: data[i].Kategori,
          harga: Number(data[i].Harga),
          stok: stokOtomatis !== undefined ? Number(stokOtomatis) : 0,
          stokDariBahan: stokOtomatis !== undefined ? Number(stokOtomatis) : 0,
          deskripsi: data[i].Deskripsi,
          gambar: data[i].URL_Gambar,
          status: data[i].Status,
          pakaiResep: true,
          _rowIndex: data[i]._rowIndex
        });
      }
    }
    
    return errorResponse('Produk tidak ditemukan');
  } catch (e) {
    return errorResponse('Error: ' + e.message);
  }
}

/**
 * Menambahkan produk baru
 * @param {Object} data - {nama, kategori, harga, deskripsi, gambar}
 */
function tambahProduk(data) {
  try {
    // Validasi
    var errors = validateRequired(data, [
      { key: 'nama', label: 'Nama Produk' },
      { key: 'kategori', label: 'Kategori' },
      { key: 'harga', label: 'Harga' }
    ]);
    
    if (errors.length > 0) {
      return errorResponse(errors.join(', '));
    }
    
    if (Number(data.harga) <= 0) {
      return errorResponse('Harga harus lebih dari 0');
    }
    
    var sheet = getSheet(SHEET_PRODUK);
    var idBaru = generateId('PRD', 'id_produk_terakhir');
    var now = new Date();
    
    sheet.appendRow([
      idBaru,
      data.nama,
      data.kategori,
      Number(data.harga),
      0,
      data.deskripsi || '',
      data.gambar || '',
      'Aktif',
      now,
      now
    ]);
    
    return successResponse({ id: idBaru }, 'Produk "' + data.nama + '" berhasil ditambahkan');
  } catch (e) {
    return errorResponse('Gagal menambahkan produk: ' + e.message);
  }
}

/**
 * Mengedit produk
 * @param {Object} data - {id, nama, kategori, harga, deskripsi, gambar}
 */
function editProduk(data) {
  try {
    var errors = validateRequired(data, [
      { key: 'id', label: 'ID Produk' },
      { key: 'nama', label: 'Nama Produk' },
      { key: 'kategori', label: 'Kategori' },
      { key: 'harga', label: 'Harga' }
    ]);
    
    if (errors.length > 0) {
      return errorResponse(errors.join(', '));
    }
    
    var sheet = getSheet(SHEET_PRODUK);
    var allData = getSheetDataSafe(SHEET_PRODUK);
    
    for (var i = 0; i < allData.length; i++) {
      if (allData[i].ID_Produk === data.id) {
        var rowIndex = allData[i]._rowIndex;
        
        // Update kolom produk
        sheet.getRange(rowIndex, 2).setValue(data.nama);
        sheet.getRange(rowIndex, 3).setValue(data.kategori);
        sheet.getRange(rowIndex, 4).setValue(Number(data.harga));
        sheet.getRange(rowIndex, 5).setValue(0);
        sheet.getRange(rowIndex, 6).setValue(data.deskripsi || '');
        sheet.getRange(rowIndex, 7).setValue(data.gambar || '');
        // Update Status jika dikirim dari frontend
        if (data.status) {
          sheet.getRange(rowIndex, 8).setValue(data.status);
        }
        sheet.getRange(rowIndex, 10).setValue(new Date()); // Tanggal_Diupdate
        
        return successResponse({ id: data.id }, 'Produk "' + data.nama + '" berhasil diupdate');
      }
    }
    
    return errorResponse('Produk dengan ID ' + data.id + ' tidak ditemukan');
  } catch (e) {
    return errorResponse('Gagal mengedit produk: ' + e.message);
  }
}

/**
 * Toggle status produk: Aktif → Nonaktif atau Nonaktif → Aktif
 */
function toggleStatusProduk(id) {
  try {
    var sheet = getSheet(SHEET_PRODUK);
    var data = getSheetDataSafe(SHEET_PRODUK);
    
    for (var i = 0; i < data.length; i++) {
      if (data[i].ID_Produk === id) {
        var rowIndex = data[i]._rowIndex;
        var statusBaru = data[i].Status === 'Aktif' ? 'Nonaktif' : 'Aktif';
        sheet.getRange(rowIndex, 8).setValue(statusBaru);
        sheet.getRange(rowIndex, 10).setValue(new Date());
        return successResponse({ status: statusBaru }, 'Status produk diubah ke ' + statusBaru);
      }
    }
    return errorResponse('Produk tidak ditemukan');
  } catch (e) {
    return errorResponse('Gagal mengubah status: ' + e.message);
  }
}

/**
 * Menghapus produk secara permanen (menghapus baris dari sheet)
 */
function hapusProdukPermanen(id) {
  try {
    var sheet = getSheet(SHEET_PRODUK);
    var data = getSheetDataSafe(SHEET_PRODUK);
    
    for (var i = 0; i < data.length; i++) {
      if (data[i].ID_Produk === id) {
        sheet.deleteRow(data[i]._rowIndex);
        return successResponse(null, 'Produk berhasil dihapus permanen');
      }
    }
    return errorResponse('Produk tidak ditemukan');
  } catch (e) {
    return errorResponse('Gagal menghapus produk: ' + e.message);
  }
}

/**
 * Menghapus produk (soft delete - set status Nonaktif) — kept for compatibility
 */
function hapusProduk(id) {
  return toggleStatusProduk(id);
}

/**
 * Mengaktifkan kembali produk
 */
function aktifkanProduk(id) {
  try {
    var sheet = getSheet(SHEET_PRODUK);
    var data = getSheetData(SHEET_PRODUK);
    
    for (var i = 0; i < data.length; i++) {
      if (data[i].ID_Produk === id) {
        var rowIndex = data[i]._rowIndex;
        sheet.getRange(rowIndex, 8).setValue('Aktif');
        sheet.getRange(rowIndex, 10).setValue(new Date());
        
        return successResponse(null, 'Produk berhasil diaktifkan');
      }
    }
    
    return errorResponse('Produk tidak ditemukan');
  } catch (e) {
    return errorResponse('Gagal mengaktifkan produk: ' + e.message);
  }
}

/**
 * Mencari produk berdasarkan keyword
 */
function cariProduk(keyword) {
  try {
    var data = getProduk();
    var keywordLower = keyword.toLowerCase();
    
    var result = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i].nama.toLowerCase().indexOf(keywordLower) !== -1 ||
          data[i].deskripsi.toLowerCase().indexOf(keywordLower) !== -1 ||
          data[i].kategori.toLowerCase().indexOf(keywordLower) !== -1) {
        result.push(data[i]);
      }
    }
    
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal mencari produk: ' + e.message);
  }
}

/**
 * Mendapatkan daftar kategori untuk kasir tabs dan dropdown produk.
 * Sumber utama: sheet Kategori (master). Fallback: kolom Kategori di sheet Produk.
 */
function getKategori() {
  try {
    var katSheet = getSheetSafe(SHEET_KATEGORI);
    if (katSheet && katSheet.getLastRow() > 1) {
      var data = katSheet.getRange(2, 1, katSheet.getLastRow() - 1, 1).getValues();
      var result = [];
      for (var i = 0; i < data.length; i++) {
        var nama = String(data[i][0] || '').trim();
        if (nama) result.push(nama);
      }
      return successResponse(result.sort());
    }

    // Fallback: baca dari kolom Kategori produk aktif (backward compat)
    var produkData = getSheetDataSafe(SHEET_PRODUK);
    var katSet = {};
    for (var j = 0; j < produkData.length; j++) {
      if (produkData[j].Status === 'Aktif' && produkData[j].Kategori) {
        katSet[produkData[j].Kategori] = true;
      }
    }
    return successResponse(Object.keys(katSet).sort());
  } catch (e) {
    return errorResponse('Gagal memuat kategori: ' + e.message);
  }
}

/**
 * Mendapatkan semua kategori + jumlah produk per kategori (halaman manajemen).
 * Sumber: sheet Kategori (master) + hitung produk dari sheet Produk.
 */
function getAllKategori() {
  try {
    // Hitung produk per kategori dari sheet Produk
    var produkData = getSheetDataSafe(SHEET_PRODUK);
    var produkCount = {};
    for (var i = 0; i < produkData.length; i++) {
      var k = produkData[i].Kategori;
      if (k) {
        produkCount[k] = (produkCount[k] || 0) + 1;
      }
    }

    // Ambil daftar kategori dari sheet Kategori (master)
    var katList = [];
    var katSheet = getSheetSafe(SHEET_KATEGORI);
    if (katSheet && katSheet.getLastRow() > 1) {
      var rows = katSheet.getRange(2, 1, katSheet.getLastRow() - 1, 1).getValues();
      for (var j = 0; j < rows.length; j++) {
        var nama = String(rows[j][0] || '').trim();
        if (nama) katList.push(nama);
      }
    }

    // Juga tambahkan kategori dari produk yang belum ada di master (backward compat)
    var katSet = {};
    katList.forEach(function(k) { katSet[k] = true; });
    for (var p = 0; p < produkData.length; p++) {
      var pk = produkData[p].Kategori;
      if (pk && !katSet[pk]) {
        katList.push(pk);
        katSet[pk] = true;
      }
    }

    return successResponse({
      kategori: katList.sort(),
      produkCount: produkCount
    });
  } catch (e) {
    return errorResponse('Gagal memuat kategori: ' + e.message);
  }
}

/**
 * Tambah kategori baru ke sheet Kategori (master data).
 * Kategori langsung tersimpan dan bisa dipakai di dropdown produk.
 */
function tambahKategoriProduk(nama) {
  try {
    if (!nama || !nama.trim()) return errorResponse('Nama kategori wajib diisi');
    nama = nama.trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var katSheet = ss.getSheetByName(SHEET_KATEGORI);

    // Auto-create sheet Kategori jika belum ada
    if (!katSheet) {
      katSheet = ss.insertSheet(SHEET_KATEGORI);
      katSheet.appendRow(['Nama_Kategori']);
      formatHeader(katSheet);
      katSheet.setColumnWidth(1, 200);
    }

    // Cek duplikat
    if (katSheet.getLastRow() > 1) {
      var existingData = katSheet.getRange(2, 1, katSheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < existingData.length; i++) {
        if (String(existingData[i][0]).trim().toLowerCase() === nama.toLowerCase()) {
          return errorResponse('Kategori "' + nama + '" sudah ada');
        }
      }
    }

    katSheet.appendRow([nama]);

    // Perbarui data validation di sheet Produk agar kolom Kategori selalu sinkron
    updateValidasiKategoriSheet();

    return successResponse({ nama: nama }, 'Kategori "' + nama + '" berhasil ditambahkan!');
  } catch (e) {
    return errorResponse('Gagal menambahkan kategori: ' + e.message);
  }
}

/**
 * Update data validation kolom Kategori (kolom C) di sheet Produk
 * agar referensi ke sheet Kategori bersifat dinamis.
 * Dipanggil otomatis setiap kali kategori ditambah/diubah/dihapus.
 */
function updateValidasiKategoriSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetProduk = ss.getSheetByName(SHEET_PRODUK);
    var sheetKategori = ss.getSheetByName(SHEET_KATEGORI);
    if (!sheetProduk || !sheetKategori) return;

    var lastRow = Math.max(sheetProduk.getLastRow(), 2);
    var katLastRow = Math.max(sheetKategori.getLastRow(), 2);

    // Buat rule validasi yang merujuk ke range sheet Kategori
    var katRange = sheetKategori.getRange(2, 1, katLastRow - 1, 1);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(katRange, true)  // true = tampilkan dropdown
      .setAllowInvalid(true)               // IZINKAN nilai di luar list (tidak blokir)
      .setHelpText('Pilih dari daftar kategori. Kelola kategori di halaman Manajemen Kategori.')
      .build();

    // Terapkan ke seluruh kolom Kategori (kolom C, baris 2 sampai lastRow)
    sheetProduk.getRange(2, 3, lastRow - 1, 1).setDataValidation(rule);
  } catch (e) {
    Logger.log('updateValidasiKategoriSheet error: ' + e.message);
  }
}

/**
 * Rename kategori:
 * 1. Update nama di sheet Kategori (master)
 * 2. Update kolom Kategori di semua produk yang menggunakan nama lama
 */
function renameKategori(namaLama, namaBaru) {
  try {
    if (!namaLama || !namaBaru) return errorResponse('Nama kategori wajib diisi');
    namaBaru = namaBaru.trim();
    if (namaLama.trim() === namaBaru) return errorResponse('Nama kategori sama');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var katSheet = ss.getSheetByName(SHEET_KATEGORI);

    // Cek duplikat nama baru
    if (katSheet && katSheet.getLastRow() > 1) {
      var katData = katSheet.getRange(2, 1, katSheet.getLastRow() - 1, 1).getValues();
      for (var k = 0; k < katData.length; k++) {
        var existing = String(katData[k][0]).trim().toLowerCase();
        if (existing === namaBaru.toLowerCase() && String(katData[k][0]).trim() !== namaLama) {
          return errorResponse('Kategori "' + namaBaru + '" sudah ada');
        }
      }

      // Update di sheet Kategori
      for (var m = 0; m < katData.length; m++) {
        if (String(katData[m][0]).trim() === namaLama) {
          katSheet.getRange(m + 2, 1).setValue(namaBaru);
          break;
        }
      }
    }

    // Update kolom Kategori di sheet Produk
    var produkData = getSheetDataSafe(SHEET_PRODUK);
    var sheetProduk = getSheet(SHEET_PRODUK);
    var count = 0;
    for (var i = 0; i < produkData.length; i++) {
      if (produkData[i].Kategori === namaLama) {
        sheetProduk.getRange(produkData[i]._rowIndex, 3).setValue(namaBaru);
        count++;
      }
    }

    // Perbarui data validation setelah rename
    updateValidasiKategoriSheet();

    return successResponse(
      { count: count },
      'Kategori diubah: "' + namaLama + '" → "' + namaBaru + '"' + (count > 0 ? ' (' + count + ' produk diperbarui)' : '')
    );
  } catch (e) {
    return errorResponse('Gagal mengubah kategori: ' + e.message);
  }
}

/**
 * Hapus kategori:
 * 1. Hapus dari sheet Kategori (master)
 * 2. Pindahkan produk dengan kategori ini ke "Lainnya"
 */
function hapusKategori(nama) {
  try {
    if (!nama) return errorResponse('Nama kategori wajib');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var katSheet = ss.getSheetByName(SHEET_KATEGORI);

    // Hapus baris dari sheet Kategori
    if (katSheet && katSheet.getLastRow() > 1) {
      var katData = katSheet.getRange(2, 1, katSheet.getLastRow() - 1, 1).getValues();
      for (var k = katData.length - 1; k >= 0; k--) { // dari bawah agar rowIndex tidak geser
        if (String(katData[k][0]).trim() === nama) {
          katSheet.deleteRow(k + 2);
          break;
        }
      }
    }

    // Pindahkan produk ke "Lainnya"
    var sheetProduk = getSheet(SHEET_PRODUK);
    var produkData = getSheetDataSafe(SHEET_PRODUK);
    var count = 0;
    for (var i = 0; i < produkData.length; i++) {
      if (produkData[i].Kategori === nama) {
        sheetProduk.getRange(produkData[i]._rowIndex, 3).setValue('Lainnya');
        count++;
      }
    }

    // Perbarui data validation setelah hapus
    updateValidasiKategoriSheet();

    return successResponse(
      { count: count },
      'Kategori "' + nama + '" dihapus.' + (count > 0 ? ' ' + count + ' produk dipindahkan ke "Lainnya".' : '')
    );
  } catch (e) {
    return errorResponse('Gagal menghapus kategori: ' + e.message);
  }
}

/**
 * MIGRASI: Impor semua kategori unik dari sheet Produk ke sheet Kategori (master).
 * Jalankan SATU KALI dari GAS Editor setelah upload kode ini.
 * Aman dijalankan berulang (tidak duplikat).
 */
function migrasiKategoriDariProduk() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var katSheet = ss.getSheetByName(SHEET_KATEGORI);

    if (!katSheet) {
      katSheet = ss.insertSheet(SHEET_KATEGORI);
      katSheet.appendRow(['Nama_Kategori']);
      formatHeader(katSheet);
      katSheet.setColumnWidth(1, 200);
    }

    // Kumpulkan kategori yang sudah ada di sheet Kategori
    var existing = {};
    if (katSheet.getLastRow() > 1) {
      var existRows = katSheet.getRange(2, 1, katSheet.getLastRow() - 1, 1).getValues();
      for (var e = 0; e < existRows.length; e++) {
        var n = String(existRows[e][0]).trim().toLowerCase();
        if (n) existing[n] = true;
      }
    }

    // Ambil semua kategori unik dari sheet Produk
    var produkData = getSheetDataSafe(SHEET_PRODUK);
    var toAdd = [];
    var seen = {};
    for (var i = 0; i < produkData.length; i++) {
      var kat = String(produkData[i].Kategori || '').trim();
      if (kat && !existing[kat.toLowerCase()] && !seen[kat.toLowerCase()]) {
        toAdd.push([kat]);
        seen[kat.toLowerCase()] = true;
      }
    }

    if (toAdd.length > 0) {
      katSheet.getRange(katSheet.getLastRow() + 1, 1, toAdd.length, 1).setValues(toAdd);
    }

    SpreadsheetApp.getUi().alert(
      '✅ Migrasi Kategori Selesai!\n' +
      '- ' + toAdd.length + ' kategori baru ditambahkan ke sheet "Kategori"\n' +
      '- ' + Object.keys(existing).length + ' kategori sudah ada sebelumnya\n\n' +
      'Total: ' + (toAdd.length + Object.keys(existing).length) + ' kategori di master sheet.'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Gagal migrasi: ' + e.message);
  }
}

