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
 * Mendapatkan daftar kategori unik (untuk kasir tabs) — hanya dari kolom Kategori di Produk
 */
function getKategori() {
  try {
    var data = getSheetDataSafe(SHEET_PRODUK);
    var kategoriSet = {};
    for (var i = 0; i < data.length; i++) {
      if (data[i].Status === 'Aktif' && data[i].Kategori) {
        kategoriSet[data[i].Kategori] = true;
      }
    }
    return successResponse(Object.keys(kategoriSet).sort());
  } catch (e) {
    return errorResponse('Gagal memuat kategori: ' + e.message);
  }
}

/**
 * Mendapatkan semua kategori unik + jumlah produk per kategori (halaman manajemen)
 * Sumber: kolom Kategori di sheet Produk (semua status)
 */
function getAllKategori() {
  try {
    var data = getSheetDataSafe(SHEET_PRODUK);
    var kategoriSet = {};
    var produkCount = {};

    for (var i = 0; i < data.length; i++) {
      var kat = data[i].Kategori;
      if (kat) {
        kategoriSet[kat] = true;
        if (!produkCount[kat]) produkCount[kat] = 0;
        produkCount[kat]++;
      }
    }

    return successResponse({
      kategori: Object.keys(kategoriSet).sort(),
      produkCount: produkCount
    });
  } catch (e) {
    return errorResponse('Gagal memuat kategori: ' + e.message);
  }
}

/**
 * Cek apakah kategori sudah ada di sheet Produk
 * Digunakan untuk validasi sebelum rename agar tidak tumbukan
 */
function tambahKategoriProduk(nama) {
  try {
    if (!nama || !nama.trim()) return errorResponse('Nama kategori wajib diisi');
    nama = nama.trim();

    // Cek apakah sudah digunakan oleh produk manapun
    var data = getSheetDataSafe(SHEET_PRODUK);
    for (var i = 0; i < data.length; i++) {
      if (data[i].Kategori && data[i].Kategori.toLowerCase() === nama.toLowerCase()) {
        return errorResponse('Kategori "' + nama + '" sudah ada di produk');
      }
    }

    // Kategori tidak disimpan terpisah — hanya muncul saat ada produk dengan kategori tsb.
    // Kembalikan sukses dengan petunjuk cara menggunakannya.
    return successResponse(
      { nama: nama },
      'Kategori "' + nama + '" siap digunakan. Tambahkan produk baru dengan kategori ini untuk memunculkannya.'
    );
  } catch (e) {
    return errorResponse('Gagal: ' + e.message);
  }
}

/**
 * Rename kategori — update kolom Kategori di semua produk yang menggunakan nama lama
 */
function renameKategori(namaLama, namaBaru) {
  try {
    if (!namaLama || !namaBaru) return errorResponse('Nama kategori wajib diisi');
    namaBaru = namaBaru.trim();
    if (namaLama === namaBaru) return errorResponse('Nama kategori sama');

    // Cek duplikat nama baru di produk lain
    var data = getSheetDataSafe(SHEET_PRODUK);
    for (var k = 0; k < data.length; k++) {
      if (data[k].Kategori &&
          data[k].Kategori.toLowerCase() === namaBaru.toLowerCase() &&
          data[k].Kategori !== namaLama) {
        return errorResponse('Kategori "' + namaBaru + '" sudah digunakan oleh produk lain');
      }
    }

    // Update kolom Kategori di semua produk yang menggunakan namaLama
    var sheet = getSheet(SHEET_PRODUK);
    var count = 0;
    for (var i = 0; i < data.length; i++) {
      if (data[i].Kategori === namaLama) {
        sheet.getRange(data[i]._rowIndex, 3).setValue(namaBaru);
        count++;
      }
    }

    if (count === 0) return errorResponse('Tidak ada produk dengan kategori "' + namaLama + '"');
    return successResponse(
      { count: count },
      'Kategori diubah: "' + namaLama + '" → "' + namaBaru + '" (' + count + ' produk diperbarui)'
    );
  } catch (e) {
    return errorResponse('Gagal mengubah kategori: ' + e.message);
  }
}

/**
 * Hapus kategori — pindahkan semua produk dengan kategori ini ke "Lainnya"
 */
function hapusKategori(nama) {
  try {
    if (!nama) return errorResponse('Nama kategori wajib');

    var sheet = getSheet(SHEET_PRODUK);
    var data = getSheetDataSafe(SHEET_PRODUK);
    var count = 0;

    for (var i = 0; i < data.length; i++) {
      if (data[i].Kategori === nama) {
        sheet.getRange(data[i]._rowIndex, 3).setValue('Lainnya');
        count++;
      }
    }

    if (count > 0) {
      return successResponse(
        { count: count },
        'Kategori "' + nama + '" dihapus. ' + count + ' produk dipindahkan ke "Lainnya".'
      );
    } else {
      return successResponse(null, 'Kategori "' + nama + '" tidak memiliki produk, tidak ada yang diubah.');
    }
  } catch (e) {
    return errorResponse('Gagal menghapus kategori: ' + e.message);
  }
}
