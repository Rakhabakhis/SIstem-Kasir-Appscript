/**
 * =====================================================
 * BAHAN BAKU SERVICE - Manajemen Bahan Baku & Resep
 * Sistem Kasir (POS) - Google Apps Script
 * =====================================================
 */

// ==================== BAHAN BAKU CRUD ====================

/**
 * Mendapatkan semua bahan baku
 */
function getAllBahanBaku() {
  try {
    var data = getSheetDataSafe(SHEET_BAHAN_BAKU);
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      result.push({
        id: row.ID_Bahan,
        nama: row.Nama_Bahan,
        grup: row.Grup || '',
        stok: Number(row.Stok) || 0,
        satuan: row.Satuan || '',
        stokMinimum: Number(row.Stok_Minimum) || 0,
        keterangan: row.Keterangan || '',
        status: row.Status || 'Aktif',
        tanggalDibuat: row.Tanggal_Dibuat ? formatTanggal(row.Tanggal_Dibuat) : '',
        tanggalDiupdate: row.Tanggal_Diupdate ? formatTanggal(row.Tanggal_Diupdate) : '',
        _rowIndex: row._rowIndex
      });
    }
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal memuat bahan baku: ' + e.message);
  }
}

/**
 * Mendapatkan bahan baku aktif saja (untuk frontend kasir)
 */
function getBahanBakuAktif() {
  try {
    var data = getSheetDataSafe(SHEET_BAHAN_BAKU);
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row.Status === 'Aktif') {
        result.push({
          id: row.ID_Bahan,
          nama: row.Nama_Bahan,
          grup: row.Grup || '',
          stok: Number(row.Stok) || 0,
          satuan: row.Satuan || '',
          stokMinimum: Number(row.Stok_Minimum) || 0
        });
      }
    }
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal memuat bahan baku: ' + e.message);
  }
}

/**
 * Tambah bahan baku baru
 */
function tambahBahanBaku(data) {
  try {
    var errors = validateRequired(data, [
      { key: 'nama', label: 'Nama Bahan' },
      { key: 'satuan', label: 'Satuan' }
    ]);
    if (errors.length > 0) return errorResponse(errors.join(', '));

    var sheet = getSheet(SHEET_BAHAN_BAKU);
    var id = generateId('BHN', 'id_bahan_terakhir');
    var now = getNow();

    sheet.appendRow([
      id,
      data.nama,
      data.grup || '',
      Number(data.stok) || 0,
      data.satuan,
      Number(data.stokMinimum) || 0,
      data.keterangan || '',
      'Aktif',
      now,
      now
    ]);

    _checkAndSendLowStockEmail(false);

    return successResponse({ id: id }, 'Bahan baku "' + data.nama + '" berhasil ditambahkan');
  } catch (e) {
    return errorResponse('Gagal tambah bahan baku: ' + e.message);
  }
}

/**
 * Edit bahan baku
 */
function editBahanBaku(data) {
  try {
    if (!data.id) return errorResponse('ID Bahan Baku tidak valid');
    var sheet = getSheet(SHEET_BAHAN_BAKU);
    var allData = getSheetData(SHEET_BAHAN_BAKU);

    for (var i = 0; i < allData.length; i++) {
      if (allData[i].ID_Bahan === data.id) {
        var row = allData[i]._rowIndex;
        sheet.getRange(row, 2).setValue(data.nama || allData[i].Nama_Bahan);
        sheet.getRange(row, 3).setValue(data.grup !== undefined ? data.grup : allData[i].Grup);
        sheet.getRange(row, 5).setValue(data.satuan || allData[i].Satuan);
        sheet.getRange(row, 6).setValue(Number(data.stokMinimum) || 0);
        sheet.getRange(row, 7).setValue(data.keterangan !== undefined ? data.keterangan : allData[i].Keterangan);
        if (data.status) sheet.getRange(row, 8).setValue(data.status);
        sheet.getRange(row, 10).setValue(getNow());
        _checkAndSendLowStockEmail(false);
        return successResponse(null, 'Bahan baku berhasil diupdate');
      }
    }
    return errorResponse('Bahan baku tidak ditemukan');
  } catch (e) {
    return errorResponse('Gagal edit bahan baku: ' + e.message);
  }
}

/**
 * Hapus (nonaktifkan) bahan baku
 */
function hapusBahanBaku(id) {
  try {
    var sheet = getSheet(SHEET_BAHAN_BAKU);
    var data = getSheetData(SHEET_BAHAN_BAKU);
    for (var i = 0; i < data.length; i++) {
      if (data[i].ID_Bahan === id) {
        sheet.getRange(data[i]._rowIndex, 8).setValue('Nonaktif');
        sheet.getRange(data[i]._rowIndex, 10).setValue(getNow());
        _checkAndSendLowStockEmail(false);
        return successResponse(null, 'Bahan baku dinonaktifkan');
      }
    }
    return errorResponse('Bahan baku tidak ditemukan');
  } catch (e) {
    return errorResponse('Gagal hapus bahan baku: ' + e.message);
  }
}

/**
 * Restok bahan baku (tambah stok masuk)
 */
function restokBahanBaku(id, jumlah, keterangan) {
  try {
    jumlah = Number(jumlah);
    if (!id || jumlah <= 0) return errorResponse('ID dan jumlah restok tidak valid');

    var sheet = getSheet(SHEET_BAHAN_BAKU);
    var data = getSheetData(SHEET_BAHAN_BAKU);

    for (var i = 0; i < data.length; i++) {
      if (data[i].ID_Bahan === id) {
        var stokLama = Number(data[i].Stok) || 0;
        var stokBaru = stokLama + jumlah;
        sheet.getRange(data[i]._rowIndex, 4).setValue(stokBaru);
        sheet.getRange(data[i]._rowIndex, 10).setValue(getNow());

        // Log perubahan
        _logBahanBaku(id, data[i].Nama_Bahan, stokLama, jumlah, stokBaru, 'Masuk', keterangan || 'Restok manual', '', 'Admin');
        _checkAndSendLowStockEmail(false);

        return successResponse({ stokBaru: stokBaru }, 'Restok berhasil. Stok baru: ' + stokBaru + ' ' + data[i].Satuan);
      }
    }
    return errorResponse('Bahan baku tidak ditemukan');
  } catch (e) {
    return errorResponse('Gagal restok: ' + e.message);
  }
}

/**
 * Penyesuaian stok bahan baku (set ke nilai tertentu)
 */
function penyesuaianStokBahanBaku(id, stokBaru, keterangan, namaKasir) {
  try {
    stokBaru = Number(stokBaru);
    if (!id || stokBaru < 0) return errorResponse('ID dan stok baru tidak valid');

    var sheet = getSheet(SHEET_BAHAN_BAKU);
    var data = getSheetData(SHEET_BAHAN_BAKU);

    for (var i = 0; i < data.length; i++) {
      if (data[i].ID_Bahan === id) {
        var stokLama = Number(data[i].Stok) || 0;
        var perubahan = stokBaru - stokLama;
        sheet.getRange(data[i]._rowIndex, 4).setValue(stokBaru);
        sheet.getRange(data[i]._rowIndex, 10).setValue(getNow());

        _logBahanBaku(id, data[i].Nama_Bahan, stokLama, perubahan, stokBaru, 'Penyesuaian', keterangan || 'Penyesuaian stok', '', namaKasir || 'Admin');
        _checkAndSendLowStockEmail(false);

        return successResponse({ stokBaru: stokBaru }, 'Penyesuaian stok berhasil');
      }
    }
    return errorResponse('Bahan baku tidak ditemukan');
  } catch (e) {
    return errorResponse('Gagal penyesuaian stok: ' + e.message);
  }
}

/**
 * Mendapatkan bahan baku dengan stok rendah (di bawah Stok_Minimum)
 */
function getBahanBakuLowStock() {
  try {
    var data = getSheetDataSafe(SHEET_BAHAN_BAKU);
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row.Status === 'Aktif') {
        var stok = Number(row.Stok) || 0;
        var min = Number(row.Stok_Minimum) || 0;
        if (stok <= min) {
          result.push({
            id: row.ID_Bahan,
            nama: row.Nama_Bahan,
            grup: row.Grup || '',
            stok: stok,
            stokMinimum: min,
            satuan: row.Satuan || ''
          });
        }
      }
    }
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal mengambil data stok rendah: ' + e.message);
  }
}

// ==================== RESEP PRODUK ====================

/**
 * Mendapatkan resep berdasarkan ID produk
 */
function getResepByProduk(idProduk) {
  try {
    var data = getSheetDataSafe(SHEET_RESEP_PRODUK);
    var result = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i].ID_Produk === idProduk) {
        result.push({
          id: data[i].ID_Resep,
          idProduk: data[i].ID_Produk,
          namaProduk: data[i].Nama_Produk,
          idBahan: data[i].ID_Bahan || '',
          namaBahan: data[i].Nama_Bahan || '',
          jumlahPerPorsi: Number(data[i].Jumlah_Per_Porsi) || 0,
          satuan: data[i].Satuan || '',
          tipeBahan: data[i].Tipe_Bahan || 'Tetap',
          grupPilihan: data[i].Grup_Pilihan || '',
          _rowIndex: data[i]._rowIndex
        });
      }
    }
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal memuat resep: ' + e.message);
  }
}

/**
 * Mendapatkan semua resep (untuk kebutuhan admin / perhitungan)
 */
function getAllResep() {
  try {
    var data = getSheetDataSafe(SHEET_RESEP_PRODUK);
    var result = [];
    for (var i = 0; i < data.length; i++) {
      result.push({
        id: data[i].ID_Resep,
        idProduk: data[i].ID_Produk,
        namaProduk: data[i].Nama_Produk,
        idBahan: data[i].ID_Bahan || '',
        namaBahan: data[i].Nama_Bahan || '',
        jumlahPerPorsi: Number(data[i].Jumlah_Per_Porsi) || 0,
        satuan: data[i].Satuan || '',
        tipeBahan: data[i].Tipe_Bahan || 'Tetap',
        grupPilihan: data[i].Grup_Pilihan || ''
      });
    }
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal memuat semua resep: ' + e.message);
  }
}

/**
 * Simpan/Update resep produk (replace semua bahan untuk produk tersebut)
 * @param {string} idProduk
 * @param {string} namaProduk
 * @param {Array} resepItems - [{idBahan, namaBahan, jumlahPerPorsi, satuan, tipeBahan, grupPilihan}]
 *   - Tetap: idBahan wajib
 *   - Pilihan: idBahan kosong, grupPilihan wajib
 */
function saveResepProduk(idProduk, namaProduk, resepItems) {
  try {
    if (!idProduk) return errorResponse('ID Produk tidak valid');

    var sheet = getSheet(SHEET_RESEP_PRODUK);
    var allData = getSheetData(SHEET_RESEP_PRODUK);

    // Hapus semua resep lama untuk produk ini (dari bawah ke atas agar index tidak bergeser)
    var rowsToDelete = [];
    for (var i = 0; i < allData.length; i++) {
      if (allData[i].ID_Produk === idProduk) {
        rowsToDelete.push(allData[i]._rowIndex);
      }
    }
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var d = 0; d < rowsToDelete.length; d++) {
      sheet.deleteRow(rowsToDelete[d]);
    }

    // Tambah resep baru
    if (resepItems && resepItems.length > 0) {
      for (var j = 0; j < resepItems.length; j++) {
        var item = resepItems[j];
        var tipe = item.tipeBahan || 'Tetap';
        if (tipe === 'Tetap' && !item.idBahan) continue;           // Tetap wajib idBahan
        if (tipe === 'Pilihan' && !item.grupPilihan) continue;     // Pilihan wajib grupPilihan
        if (!item.jumlahPerPorsi) continue;
        var idResep = generateId('RSP', 'id_resep_terakhir');
        sheet.appendRow([
          idResep,
          idProduk,
          namaProduk || '',
          item.idBahan || '',
          item.namaBahan || '',
          Number(item.jumlahPerPorsi),
          item.satuan || '',
          tipe,
          item.grupPilihan || ''
        ]);
      }
    }

    return successResponse(null, 'Resep produk berhasil disimpan');
  } catch (e) {
    return errorResponse('Gagal simpan resep: ' + e.message);
  }
}

/**
 * Hitung stok tersedia (porsi) berdasarkan bahan baku untuk 1 produk
 * Mendukung tipe Tetap dan Pilihan
 */
function hitungStokDariResep(idProduk) {
  try {
    var result = getStokDariResepSemuaProduk();
    var stok = result.stok[idProduk];
    return stok !== undefined ? stok : null;
  } catch (e) {
    return null;
  }
}

/**
 * Hitung stok tersedia untuk semua produk yang punya resep.
 * - Bahan Tetap: min(floor(stok[idBahan] / qty))
 * - Bahan Pilihan: floor(sum(stok semua bahan dalam Grup) / qty)
 * Returns { stok: {idProduk: porsi}, pilihanPerProduk: {idProduk: [{grupPilihan,jumlahPerPorsi,satuan}]} }
 */
function getStokDariResepSemuaProduk() {
  try {
    var resepData = getSheetDataSafe(SHEET_RESEP_PRODUK);
    var bahanData = getSheetDataSafe(SHEET_BAHAN_BAKU);

    // Map stok dan grup bahan baku
    var bahanMap = {};       // idBahan -> stok
    var grupStokMap = {};    // grupNama -> total stok aktif dalam grup
    for (var i = 0; i < bahanData.length; i++) {
      var bd = bahanData[i];
      bahanMap[bd.ID_Bahan] = Number(bd.Stok) || 0;
      if (bd.Grup && bd.Status === 'Aktif') {
        grupStokMap[bd.Grup] = (grupStokMap[bd.Grup] || 0) + (Number(bd.Stok) || 0);
      }
    }

    // Group resep per produk
    var resepPerProduk = {};
    for (var j = 0; j < resepData.length; j++) {
      var row = resepData[j];
      if (!row.ID_Produk) continue;
      if (!resepPerProduk[row.ID_Produk]) resepPerProduk[row.ID_Produk] = [];
      resepPerProduk[row.ID_Produk].push(row);
    }

    // Hitung porsi per produk
    var stokResult = {};
    var pilihanPerProduk = {}; // {idProduk: [{grupPilihan, jumlahPerPorsi, satuan}]}

    for (var idProduk in resepPerProduk) {
      var bahan = resepPerProduk[idProduk];
      var minPorsi = Infinity;
      var pilihanList = [];

      for (var k = 0; k < bahan.length; k++) {
        var tipe = bahan[k].Tipe_Bahan || 'Tetap';
        var qty = Number(bahan[k].Jumlah_Per_Porsi) || 1;

        if (tipe === 'Pilihan') {
          var grup = bahan[k].Grup_Pilihan || '';
          var stokGrup = grupStokMap[grup] !== undefined ? grupStokMap[grup] : 0;
          var porsi = Math.floor(stokGrup / qty);
          if (porsi < minPorsi) minPorsi = porsi;
          // Catat entry pilihan untuk frontend picker
          var alreadyAdded = false;
          for (var p = 0; p < pilihanList.length; p++) {
            if (pilihanList[p].grupPilihan === grup) { alreadyAdded = true; break; }
          }
          if (!alreadyAdded) {
            pilihanList.push({ grupPilihan: grup, jumlahPerPorsi: qty, satuan: bahan[k].Satuan || '' });
          }
        } else {
          var idBahan = bahan[k].ID_Bahan;
          if (!idBahan) continue;
          var stok = bahanMap[idBahan] !== undefined ? bahanMap[idBahan] : 0;
          var porsiTetap = Math.floor(stok / qty);
          if (porsiTetap < minPorsi) minPorsi = porsiTetap;
        }
      }

      stokResult[idProduk] = minPorsi === Infinity ? 0 : minPorsi;
      if (pilihanList.length > 0) pilihanPerProduk[idProduk] = pilihanList;
    }

    return { stok: stokResult, pilihanPerProduk: pilihanPerProduk };
  } catch (e) {
    Logger.log('getStokDariResepSemuaProduk error: ' + e.message);
    return { stok: {}, pilihanPerProduk: {} };
  }
}

// ==================== DEDUCT BAHAN BAKU (DARI TRANSAKSI) ====================

/**
 * Kurangi stok bahan baku berdasarkan item yang terjual.
 * Mendukung pilihanVarian di cart item untuk bahan bertipe Pilihan.
 * @param {Array} items - [{id, nama, jumlah, pilihanVarian?: [{grupPilihan,idBahan,jumlahPerPorsi}]}]
 * @param {string} idTransaksi
 * @param {string} namaKasir
 */
function kurangiBahanBaku(items, idTransaksi, namaKasir) {
  try {
    var resepData = getSheetDataSafe(SHEET_RESEP_PRODUK);
    if (!resepData || resepData.length === 0) return;

    var bahanSheet = getSheet(SHEET_BAHAN_BAKU);
    var bahanData = getSheetData(SHEET_BAHAN_BAKU);

    var bahanMap = {};
    for (var i = 0; i < bahanData.length; i++) {
      bahanMap[bahanData[i].ID_Bahan] = bahanData[i];
    }

    var kebutuhan = {}; // idBahan -> total qty yang dibutuhkan

    for (var m = 0; m < items.length; m++) {
      var item = items[m];
      var qtyJual = Number(item.jumlah) || 1;

      // Buat peta variant yang dipilih untuk item ini
      var variantMap = {}; // grupPilihan -> idBahan
      if (item.pilihanVarian && item.pilihanVarian.length > 0) {
        for (var v = 0; v < item.pilihanVarian.length; v++) {
          variantMap[item.pilihanVarian[v].grupPilihan] = item.pilihanVarian[v].idBahan;
        }
      }

      for (var r = 0; r < resepData.length; r++) {
        if (resepData[r].ID_Produk !== item.id) continue;
        var tipe = resepData[r].Tipe_Bahan || 'Tetap';
        var qtyPerPorsi = Number(resepData[r].Jumlah_Per_Porsi) || 1;

        if (tipe === 'Pilihan') {
          var grup = resepData[r].Grup_Pilihan || '';
          var idBahanPilih = variantMap[grup];
          if (!idBahanPilih) continue; // Varian tidak dipilih, lewati
          kebutuhan[idBahanPilih] = (kebutuhan[idBahanPilih] || 0) + (qtyPerPorsi * qtyJual);
        } else {
          var idBahanTetap = resepData[r].ID_Bahan;
          if (!idBahanTetap) continue;
          kebutuhan[idBahanTetap] = (kebutuhan[idBahanTetap] || 0) + (qtyPerPorsi * qtyJual);
        }
      }
    }

    // Kurangi stok bahan baku
    for (var idBahan in kebutuhan) {
      var bahan = bahanMap[idBahan];
      if (!bahan) continue;
      var stokLama = Number(bahan.Stok) || 0;
      var dikurangi = kebutuhan[idBahan];
      var stokBaru = Math.max(0, stokLama - dikurangi);

      bahanSheet.getRange(bahan._rowIndex, 4).setValue(stokBaru);
      bahanSheet.getRange(bahan._rowIndex, 10).setValue(getNow());

      _logBahanBaku(
        idBahan, bahan.Nama_Bahan, stokLama, -dikurangi, stokBaru,
        'Keluar', 'Penjualan ' + idTransaksi, idTransaksi, namaKasir || 'Kasir'
      );
    }
    _checkAndSendLowStockEmail(false);
  } catch (e) {
    Logger.log('Warning: kurangiBahanBaku error - ' + e.message);
  }
}

/**
 * Validasi stok bahan baku mencukupi untuk item yang akan dijual.
 * Mendukung pilihanVarian untuk bahan bertipe Pilihan.
 * @returns {string|null} - Pesan error atau null jika OK
 */
function validasiBahanBaku(items) {
  try {
    var resepData = getSheetDataSafe(SHEET_RESEP_PRODUK);
    if (!resepData || resepData.length === 0) return 'Resep produk belum diset. Tidak bisa memproses transaksi.';

    var bahanData = getSheetDataSafe(SHEET_BAHAN_BAKU);
    var bahanMap = {};
    for (var i = 0; i < bahanData.length; i++) {
      bahanMap[bahanData[i].ID_Bahan] = {
        stok: Number(bahanData[i].Stok) || 0,
        nama: bahanData[i].Nama_Bahan,
        satuan: bahanData[i].Satuan || ''
      };
    }

    var resepPerProduk = {};
    for (var x = 0; x < resepData.length; x++) {
      if (!resepPerProduk[resepData[x].ID_Produk]) resepPerProduk[resepData[x].ID_Produk] = true;
    }

    var kebutuhan = {};
    for (var m = 0; m < items.length; m++) {
      var item = items[m];
      var qtyJual = Number(item.jumlah) || 1;
      if (!resepPerProduk[item.id]) {
        return 'Produk "' + (item.nama || item.id) + '" belum memiliki resep bahan baku.';
      }

      var variantMap = {};
      if (item.pilihanVarian && item.pilihanVarian.length > 0) {
        for (var v = 0; v < item.pilihanVarian.length; v++) {
          variantMap[item.pilihanVarian[v].grupPilihan] = item.pilihanVarian[v].idBahan;
        }
      }

      for (var r = 0; r < resepData.length; r++) {
        if (resepData[r].ID_Produk !== item.id) continue;
        var tipe = resepData[r].Tipe_Bahan || 'Tetap';
        var qtyPerPorsi = Number(resepData[r].Jumlah_Per_Porsi) || 1;

        if (tipe === 'Pilihan') {
          var grup = resepData[r].Grup_Pilihan || '';
          var idBahanPilih = variantMap[grup];
          if (!idBahanPilih) continue;
          kebutuhan[idBahanPilih] = (kebutuhan[idBahanPilih] || 0) + (qtyPerPorsi * qtyJual);
        } else {
          var idBahanTetap = resepData[r].ID_Bahan;
          if (!idBahanTetap) continue;
          kebutuhan[idBahanTetap] = (kebutuhan[idBahanTetap] || 0) + (qtyPerPorsi * qtyJual);
        }
      }
    }

    for (var idBahan in kebutuhan) {
      var bahan = bahanMap[idBahan];
      if (!bahan) continue;
      if (bahan.stok < kebutuhan[idBahan]) {
        return 'Stok bahan "' + bahan.nama + '" tidak cukup. Tersedia: ' + bahan.stok + ' ' + bahan.satuan + ', Dibutuhkan: ' + kebutuhan[idBahan] + ' ' + bahan.satuan;
      }
    }

    return null;
  } catch (e) {
    Logger.log('Warning: validasiBahanBaku error - ' + e.message);
    return null;
  }
}

/**
 * Mengembalikan stok bahan baku berdasarkan item transaksi yang dihapus/dibatalkan.
 * Struktur items sama seperti kurangiBahanBaku.
 */
function kembalikanBahanBaku(items, idTransaksi, namaKasir) {
  try {
    var resepData = getSheetDataSafe(SHEET_RESEP_PRODUK);
    if (!resepData || resepData.length === 0) return;

    var bahanSheet = getSheet(SHEET_BAHAN_BAKU);
    var bahanData = getSheetData(SHEET_BAHAN_BAKU);
    var bahanMap = {};
    for (var i = 0; i < bahanData.length; i++) {
      bahanMap[bahanData[i].ID_Bahan] = bahanData[i];
    }

    var kebutuhan = {};
    for (var m = 0; m < items.length; m++) {
      var item = items[m];
      var qtyJual = Number(item.jumlah) || 1;

      var variantMap = {};
      if (item.pilihanVarian && item.pilihanVarian.length > 0) {
        for (var v = 0; v < item.pilihanVarian.length; v++) {
          variantMap[item.pilihanVarian[v].grupPilihan] = item.pilihanVarian[v].idBahan;
        }
      }

      for (var r = 0; r < resepData.length; r++) {
        if (resepData[r].ID_Produk !== item.id) continue;
        var tipe = resepData[r].Tipe_Bahan || 'Tetap';
        var qtyPerPorsi = Number(resepData[r].Jumlah_Per_Porsi) || 1;

        if (tipe === 'Pilihan') {
          var grup = resepData[r].Grup_Pilihan || '';
          var idBahanPilih = variantMap[grup];
          if (!idBahanPilih) continue;
          kebutuhan[idBahanPilih] = (kebutuhan[idBahanPilih] || 0) + (qtyPerPorsi * qtyJual);
        } else {
          var idBahanTetap = resepData[r].ID_Bahan;
          if (!idBahanTetap) continue;
          kebutuhan[idBahanTetap] = (kebutuhan[idBahanTetap] || 0) + (qtyPerPorsi * qtyJual);
        }
      }
    }

    for (var idBahan in kebutuhan) {
      var bahan = bahanMap[idBahan];
      if (!bahan) continue;
      var stokLama = Number(bahan.Stok) || 0;
      var ditambah = kebutuhan[idBahan];
      var stokBaru = stokLama + ditambah;

      bahanSheet.getRange(bahan._rowIndex, 4).setValue(stokBaru);
      bahanSheet.getRange(bahan._rowIndex, 10).setValue(getNow());

      _logBahanBaku(
        idBahan, bahan.Nama_Bahan, stokLama, ditambah, stokBaru,
        'Masuk', 'Pembatalan/Penghapusan transaksi ' + idTransaksi, idTransaksi, namaKasir || 'Sistem'
      );
    }
    _checkAndSendLowStockEmail(false);
  } catch (e) {
    Logger.log('Warning: kembalikanBahanBaku error - ' + e.message);
  }
}

/**
 * Kirim email notifikasi stok rendah secara manual dari halaman Pengaturan.
 * Pastikan sudah menjalankan inisialisasiMailApp() dari editor Apps Script
 * sebelum menggunakan fitur ini.
 */
function kirimTesEmailStokRendah() {
  try {
    var emailList = parseEmailList(getSetting('notif_stok_email_tujuan'));
    if (emailList.length === 0) {
      return errorResponse(
        'Email tujuan notifikasi belum diatur. ' +
        'Isi email tujuan di halaman Pengaturan terlebih dahulu.'
      );
    }

    var namaToko = String(getSetting('nama_toko') || 'Toko');
    var now = getNow();
    var waktuStr = now instanceof Date ? now.toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' }) : String(now);
    var subject = '[POS] Tes Email Notifikasi - ' + namaToko;

    var bodyText = [
      'Ini adalah email tes notifikasi stok rendah dari Sistem Kasir.',
      '',
      'Toko   : ' + namaToko,
      'Waktu  : ' + waktuStr,
      '',
      'Jika email ini diterima, maka konfigurasi notifikasi sudah benar.',
      'Notifikasi otomatis akan dikirim saat ada bahan baku dengan stok di bawah minimum.'
    ].join('\n');

    var bodyHtml = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
      '<div style="background:#0D9488;padding:20px;border-radius:8px 8px 0 0;">' +
      '<h2 style="color:#fff;margin:0;">✅ Tes Email Notifikasi</h2>' +
      '</div>' +
      '<div style="background:#f8fafc;padding:24px;border:1px solid #e2e8f0;border-radius:0 0 8px 8px;">' +
      '<p style="color:#334155;">Ini adalah email tes notifikasi stok rendah dari <strong>Sistem Kasir POS</strong>.</p>' +
      '<table style="width:100%;border-collapse:collapse;margin:16px 0;">' +
      '<tr><td style="padding:8px;color:#64748b;width:80px;">Toko</td>' +
      '<td style="padding:8px;font-weight:600;color:#1e293b;">' + namaToko + '</td></tr>' +
      '<tr style="background:#f1f5f9;"><td style="padding:8px;color:#64748b;">Waktu</td>' +
      '<td style="padding:8px;font-weight:600;color:#1e293b;">' + waktuStr + '</td></tr>' +
      '</table>' +
      '<p style="color:#334155;">Jika email ini diterima, konfigurasi notifikasi sudah benar. ' +
      'Notifikasi otomatis akan dikirim saat ada bahan baku dengan stok di bawah minimum.</p>' +
      '<hr style="border:none;border-top:1px solid #e2e8f0;margin:16px 0;">' +
      '<p style="color:#94a3b8;font-size:12px;">Pesan ini dikirim oleh Sistem Kasir POS</p>' +
      '</div></div>';

    MailApp.sendEmail({
      to: emailList.join(','),
      subject: subject,
      body: bodyText,
      htmlBody: bodyHtml
    });

    return successResponse({
      penerima: emailList,
      waktu: now instanceof Date ? now.toISOString() : String(now)
    }, 'Tes email berhasil dikirim ke: ' + emailList.join(', '));
  } catch (e) {
    var errMsg = e.message || String(e);
    // Deteksi error otorisasi khusus
    if (errMsg.indexOf('script.send_mail') !== -1 || errMsg.indexOf('does not have permission') !== -1 || errMsg.indexOf('tidak memiliki izin') !== -1) {
      return errorResponse(
        '⚠️ Otorisasi MailApp belum diberikan.\n\n' +
        'CARA MEMPERBAIKI:\n' +
        '1. Buka Google Apps Script Editor (Extensions > Apps Script)\n' +
        '2. Pilih fungsi "inisialisasiMailApp" dari dropdown\n' +
        '3. Klik tombol ▶ Run\n' +
        '4. Setujui semua izin yang diminta\n' +
        '5. Coba kirim tes email lagi\n\n' +
        'Detail error: ' + errMsg
      );
    }
    return errorResponse('Gagal kirim email tes: ' + errMsg);
  }
}

/**
 * Cek kondisi stok rendah dan kirim email bila perlu.
 *
 * LOGIKA: per-item tracking (bukan time-based cooldown)
 * - Email dikirim setiap ada bahan baku BARU yang turun di bawah minimum
 * - Tidak ada batasan waktu antar pengiriman
 * - ID bahan yang sudah dinotifikasi disimpan di setting 'notif_stok_email_notified_ids'
 * - Saat bahan direstok di atas minimum, ID-nya otomatis dihapus dari daftar
 * - Cooldown 2 menit hanya untuk mencegah duplicate rapid send (race condition)
 *
 * @param {boolean} forceSend - true untuk bypass tracking dan kirim semua item low stock
 * @returns {{ok:boolean,message:string,data:Object}}
 */
function _checkAndSendLowStockEmail(forceSend) {
  try {
    var aktif = String(getSetting('notif_stok_email_aktif') || 'Tidak') === 'Ya';
    if (!aktif) return { ok: true, message: 'Notifikasi email nonaktif', data: null };

    var emailList = parseEmailList(getSetting('notif_stok_email_tujuan'));
    if (emailList.length === 0) {
      return { ok: false, message: 'Email tujuan notifikasi belum diatur', data: null };
    }

    // Baca stok terbaru langsung dari sheet (selalu fresh setelah operasi stok)
    var lowItems = _getLowStockItemsInternal();
    var currentLowIds = lowItems.map(function(it) { return it.id; });

    // Baca daftar ID yang sudah pernah dinotifikasi
    var notifiedRaw = getSetting('notif_stok_email_notified_ids') || '';
    var notifiedIds = notifiedRaw
      ? notifiedRaw.split(',').map(function(x) { return x.trim(); }).filter(function(x) { return x.length > 0; })
      : [];

    // Hapus ID yang sudah TIDAK low stock lagi (sudah direstok di atas minimum)
    var stillLowNotified = notifiedIds.filter(function(id) {
      return currentLowIds.indexOf(id) !== -1;
    });

    // Cari item BARU yang belum pernah dinotifikasi
    var newLowItems = lowItems.filter(function(it) {
      return stillLowNotified.indexOf(it.id) === -1;
    });

    // Item yang sudah pernah dinotifikasi dan masih low stock
    var prevLowItems = lowItems.filter(function(it) {
      return stillLowNotified.indexOf(it.id) !== -1;
    });

    // Tentukan item mana yang jadi trigger pengiriman
    var itemsToTrigger = forceSend ? lowItems : newLowItems;

    // Update state tracking (hapus yang sudah restok)
    if (!forceSend) {
      setSetting('notif_stok_email_notified_ids', stillLowNotified.join(','));
    }

    if (itemsToTrigger.length === 0) {
      return {
        ok: true,
        message: 'Tidak ada bahan low stock baru (semua sudah dinotifikasi sebelumnya)',
        data: { jumlahLowStock: lowItems.length, jumlahBaru: 0 }
      };
    }

    // Cooldown singkat 2 menit — hanya cegah duplicate send (race condition)
    var minCooldownMs = 2 * 60 * 1000;
    if (!forceSend) {
      var lastSentRaw = getSetting('notif_stok_email_last_sent');
      var lastSent = lastSentRaw ? new Date(lastSentRaw) : null;
      if (lastSent && !isNaN(lastSent.getTime()) && (new Date().getTime() - lastSent.getTime()) < minCooldownMs) {
        return { ok: true, message: 'Cooldown singkat aktif (2 menit)', data: { cooldown: true } };
      }
    }

    var now = new Date();
    var namaToko = String(getSetting('nama_toko') || 'Toko');
    var waktuStr = now.toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
    var adaHabis = itemsToTrigger.some(function(it) { return it.stok <= 0; });

    var subject = '[POS] ' + (adaHabis ? '🔴 Stok HABIS' : '⚠️ Stok Rendah') +
      ' - ' + itemsToTrigger.length + ' bahan baru - ' + namaToko;

    // ===== PLAIN TEXT =====
    var lines = [];
    lines.push((adaHabis ? '🔴 PERINGATAN STOK HABIS' : '⚠️ PERINGATAN STOK RENDAH') + ' - ' + namaToko);
    lines.push('Waktu: ' + waktuStr);
    lines.push('');
    lines.push('BAHAN YANG BARU BERMASALAH (' + itemsToTrigger.length + '):');
    for (var i = 0; i < itemsToTrigger.length; i++) {
      var it = itemsToTrigger[i];
      lines.push('  • [' + (it.stok <= 0 ? 'HABIS' : 'RENDAH') + '] ' + it.nama +
        ' | Stok: ' + it.stok + ' ' + it.satuan + ' | Minimum: ' + it.stokMinimum + ' ' + it.satuan);
    }
    if (prevLowItems.length > 0) {
      lines.push('');
      lines.push('BAHAN SEBELUMNYA YANG BELUM DIRESTOK (' + prevLowItems.length + '):');
      for (var j = 0; j < prevLowItems.length; j++) {
        var pv = prevLowItems[j];
        lines.push('  • ' + pv.nama + ' | Stok: ' + pv.stok + ' ' + pv.satuan);
      }
    }
    lines.push('');
    lines.push('Total bahan bermasalah saat ini: ' + lowItems.length);
    lines.push('Segera lakukan restok dari halaman Bahan Baku di Sistem Kasir.');
    var bodyText = lines.join('\n');

    // ===== HTML =====
    var newRowsHtml = '';
    for (var n = 0; n < itemsToTrigger.length; n++) {
      var ni = itemsToTrigger[n];
      var niHabis = ni.stok <= 0;
      var niColor = niHabis ? '#dc2626' : '#f97316';
      var pct = ni.stokMinimum > 0 ? Math.max(0, Math.min(100, Math.round((ni.stok / ni.stokMinimum) * 100))) : 0;
      newRowsHtml += '<tr style="' + (n % 2 === 0 ? 'background:#fff7ed;' : 'background:#fff;') + '">' +
        '<td style="padding:10px 12px;font-weight:600;color:#1e293b;">' + ni.nama + '</td>' +
        '<td style="padding:10px 12px;text-align:center;">' +
        '<span style="background:' + niColor + ';color:#fff;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;">' +
        (niHabis ? 'HABIS' : 'RENDAH') + '</span></td>' +
        '<td style="padding:10px 12px;text-align:right;">' +
        '<strong style="color:' + niColor + ';font-size:14px;">' + ni.stok + '</strong>' +
        ' <span style="color:#94a3b8;font-size:11px;">' + ni.satuan + '</span>' +
        '<div style="margin-top:4px;height:4px;background:#fee2e2;border-radius:2px;">' +
        '<div style="height:4px;width:' + pct + '%;background:' + niColor + ';border-radius:2px;"></div></div>' +
        '</td>' +
        '<td style="padding:10px 12px;text-align:right;color:#64748b;">' + ni.stokMinimum + ' <span style="font-size:11px;">' + ni.satuan + '</span></td>' +
        '</tr>';
    }

    var prevRowsHtml = '';
    for (var p = 0; p < prevLowItems.length; p++) {
      var pi = prevLowItems[p];
      prevRowsHtml += '<tr style="' + (p % 2 === 0 ? 'background:#f8fafc;' : '') + '">' +
        '<td style="padding:8px 12px;color:#64748b;">' + pi.nama + '</td>' +
        '<td style="padding:8px 12px;text-align:center;">' +
        '<span style="background:' + (pi.stok <= 0 ? '#dc2626' : '#f97316') + ';color:#fff;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:700;">' +
        (pi.stok <= 0 ? 'HABIS' : 'RENDAH') + '</span></td>' +
        '<td style="padding:8px 12px;text-align:right;color:#dc2626;font-weight:600;">' + pi.stok +
        ' <span style="color:#94a3b8;font-size:11px;">' + pi.satuan + '</span></td>' +
        '<td style="padding:8px 12px;text-align:right;color:#64748b;">' + pi.stokMinimum +
        ' <span style="font-size:11px;">' + pi.satuan + '</span></td>' +
        '</tr>';
    }

    var hdrBg = adaHabis ? '#dc2626' : '#b45309';
    var tableHead = '<table style="width:100%;border-collapse:collapse;">' +
      '<thead><tr style="background:#f1f5f9;">' +
      '<th style="padding:9px 12px;text-align:left;color:#475569;font-size:11px;">NAMA BAHAN</th>' +
      '<th style="padding:9px 12px;text-align:center;color:#475569;font-size:11px;">STATUS</th>' +
      '<th style="padding:9px 12px;text-align:right;color:#475569;font-size:11px;">STOK SAAT INI</th>' +
      '<th style="padding:9px 12px;text-align:right;color:#475569;font-size:11px;">MINIMUM</th>' +
      '</tr></thead>';

    var bodyHtml = '<div style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;">' +
      '<div style="background:' + hdrBg + ';padding:20px;border-radius:8px 8px 0 0;">' +
      '<h2 style="color:#fff;margin:0;">' + (adaHabis ? '🔴 Stok Bahan Baku HABIS' : '⚠️ Stok Bahan Baku Rendah') + '</h2>' +
      '<p style="color:rgba(255,255,255,.8);margin:4px 0 0;font-size:13px;">Toko: ' + namaToko + '  |  ' + waktuStr + '</p>' +
      '</div>' +
      '<div style="background:#fff;border:1px solid #e2e8f0;border-top:none;">' +

      // Section bahan BARU
      '<div style="padding:12px 20px;background:#fff7ed;border-bottom:2px solid #f97316;">' +
      '<strong style="color:#c2410c;font-size:13px;">🆕 Bahan yang BARU bermasalah (' + itemsToTrigger.length + ' item)</strong>' +
      '<div style="font-size:12px;color:#9a3412;margin-top:2px;">Pemicu email ini dikirim</div>' +
      '</div>' +
      tableHead + '<tbody>' + newRowsHtml + '</tbody></table>' +

      // Section bahan SEBELUMNYA (jika ada)
      (prevRowsHtml ?
        '<div style="padding:12px 20px;background:#f8fafc;border-top:1px solid #e2e8f0;">' +
        '<strong style="color:#475569;font-size:12px;">📋 Belum direstok dari notifikasi sebelumnya (' + prevLowItems.length + ' item)</strong>' +
        '</div>' +
        tableHead + '<tbody>' + prevRowsHtml + '</tbody></table>' : '') +

      '<div style="padding:14px 20px;border-top:1px solid #e2e8f0;display:flex;justify-content:space-between;align-items:center;">' +
      '<p style="color:#475569;margin:0;font-size:13px;">Segera lakukan <strong>restok</strong> dari halaman Bahan Baku di Sistem Kasir.</p>' +
      '<span style="background:#fef2f2;color:#dc2626;font-weight:700;padding:4px 12px;border-radius:20px;font-size:12px;white-space:nowrap;margin-left:12px;">Total: ' + lowItems.length + ' bahan</span>' +
      '</div></div>' +
      '<div style="padding:10px 20px;"><p style="color:#94a3b8;font-size:11px;margin:0;">Dikirim otomatis oleh Sistem Kasir POS</p></div></div>';

    MailApp.sendEmail({
      to: emailList.join(','),
      subject: subject,
      body: bodyText,
      htmlBody: bodyHtml
    });

    // Update state: tandai item baru sebagai sudah dinotifikasi
    if (!forceSend) {
      var updatedNotified = stillLowNotified.concat(newLowItems.map(function(it) { return it.id; }));
      setSetting('notif_stok_email_notified_ids', updatedNotified.join(','));
    }
    setSetting('notif_stok_email_last_sent', now.toISOString());

    return {
      ok: true,
      message: 'Email dikirim untuk ' + itemsToTrigger.length + ' bahan baru (total low stock: ' + lowItems.length + ')',
      data: { jumlahLowStock: lowItems.length, jumlahBaru: itemsToTrigger.length, penerima: emailList }
    };
  } catch (e) {
    Logger.log('_checkAndSendLowStockEmail error: ' + e.message);
    return { ok: false, message: 'Gagal kirim email notifikasi: ' + e.message, data: null };
  }
}

function _getLowStockItemsInternal() {
  var data = getSheetDataSafe(SHEET_BAHAN_BAKU);
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row.Status !== 'Aktif') continue;
    var stok = Number(row.Stok) || 0;
    var min = Number(row.Stok_Minimum) || 0;
    if (min <= 0) continue;
    if (stok <= min) {
      result.push({
        id: row.ID_Bahan,
        nama: row.Nama_Bahan,
        stok: stok,
        stokMinimum: min,
        satuan: row.Satuan || ''
      });
    }
  }
  return result;
}

// ==================== LOG BAHAN BAKU ====================

/**
 * Internal: Catat perubahan stok bahan baku ke log
 */
function _logBahanBaku(idBahan, namaBahan, stokSebelum, perubahan, stokSesudah, tipe, keterangan, idTransaksi, diupdateOleh) {
  try {
    var sheet = getSheet(SHEET_LOG_BAHAN_BAKU);
    var idLog = generateId('LBB', 'id_log_bahan_terakhir');
    sheet.appendRow([
      idLog,
      getNow(),
      idBahan,
      namaBahan,
      stokSebelum,
      perubahan,
      stokSesudah,
      tipe,
      keterangan || '',
      idTransaksi || '',
      diupdateOleh || ''
    ]);
  } catch (e) {
    Logger.log('Warning: _logBahanBaku error - ' + e.message);
  }
}

/**
 * Mendapatkan log bahan baku (terbaru dulu)
 */
function getLogBahanBaku(limit) {
  try {
    limit = limit || 100;
    var data = getSheetDataSafe(SHEET_LOG_BAHAN_BAKU);
    if (!data || data.length === 0) return successResponse([]);

    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row.ID_Log) continue;
      result.push({
        id: row.ID_Log,
        tanggal: row.Tanggal ? formatTanggal(row.Tanggal) : '',
        idBahan: row.ID_Bahan,
        namaBahan: row.Nama_Bahan,
        stokSebelum: Number(row.Stok_Sebelum) || 0,
        perubahan: Number(row.Perubahan) || 0,
        stokSesudah: Number(row.Stok_Sesudah) || 0,
        tipe: row.Tipe || '',
        keterangan: row.Keterangan || '',
        idTransaksi: row.ID_Transaksi || '',
        diupdateOleh: row.Diupdate_Oleh || ''
      });
    }

    // Sort descending by index (terbaru dulu)
    result.reverse();
    return successResponse(result.slice(0, limit));
  } catch (e) {
    return errorResponse('Gagal memuat log: ' + e.message);
  }
}
