/**
 * =====================================================
 * STOK SERVICE - Manajemen Stok & Logging
 * Sistem Kasir (POS) - Google Apps Script
 * =====================================================
 */

function updateStok(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!data.idProduk) return errorResponse('Pilih produk');
    if (!data.jumlah || Number(data.jumlah) === 0) return errorResponse('Jumlah tidak boleh 0');
    if (!data.tipe) return errorResponse('Pilih tipe perubahan');
    
    var sheet = getSheet(SHEET_PRODUK);
    var produkData = getSheetData(SHEET_PRODUK);
    var namaKasir = data.diupdateOleh || 'Kasir';
    var produk = null;
    for (var i = 0; i < produkData.length; i++) {
      if (produkData[i].ID_Produk === data.idProduk) { produk = produkData[i]; break; }
    }
    if (!produk) return errorResponse('Produk tidak ditemukan');
    
    var stokSebelum = Number(produk.Stok);
    var perubahan = Number(data.jumlah);
    var stokSesudah;
    
    if (data.tipe === 'Masuk') {
      stokSesudah = stokSebelum + Math.abs(perubahan);
      perubahan = Math.abs(perubahan);
    } else if (data.tipe === 'Keluar') {
      perubahan = -Math.abs(perubahan);
      stokSesudah = stokSebelum + perubahan;
      if (stokSesudah < 0) return errorResponse('Stok tidak cukup. Saat ini: ' + stokSebelum);
    } else if (data.tipe === 'Penyesuaian') {
      stokSesudah = Math.abs(perubahan);
      perubahan = stokSesudah - stokSebelum;
    } else {
      return errorResponse('Tipe tidak valid');
    }
    
    sheet.getRange(produk._rowIndex, 5).setValue(stokSesudah);
    sheet.getRange(produk._rowIndex, 10).setValue(new Date());
    logStokChange(data.idProduk, produk.Nama_Produk, stokSebelum, perubahan, stokSesudah, data.tipe, data.keterangan || '-', data.diupdateOleh || namaKasir);
    
    return successResponse({
      idProduk: data.idProduk, nama: produk.Nama_Produk,
      stokSebelum: stokSebelum, perubahan: perubahan, stokSesudah: stokSesudah
    }, 'Stok "' + produk.Nama_Produk + '" diupdate: ' + stokSebelum + ' → ' + stokSesudah);
  } catch (e) {
    return errorResponse('Gagal update stok: ' + e.message);
  } finally { lock.releaseLock(); }
}

function logStokChange(idProduk, namaProduk, stokSebelum, perubahan, stokSesudah, tipe, keterangan, diupdateOleh) {
  try {
    // Cek apakah sheet sudah ada sebelum logging
    var sheetLog = getSheetSafe(SHEET_LOG_STOK);
    if (!sheetLog) {
      Logger.log('LogStok sheet belum ada, skip logging');
      return;
    }
    var idLog = generateId('LOG', 'id_log_terakhir');
    sheetLog.appendRow([idLog, new Date(), idProduk, namaProduk, stokSebelum, perubahan, stokSesudah, tipe, keterangan || '-', diupdateOleh || 'Sistem']);
  } catch (e) { Logger.log('Error logStokChange: ' + e.message); }
}

function getLogStok(filter) {
  try {
    filter = filter || {};
    var data = getSheetDataSafe(SHEET_LOG_STOK);
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i]; var inc = true;
      if (filter.idProduk && row.ID_Produk !== filter.idProduk) inc = false;
      if (filter.tipe && row.Tipe !== filter.tipe) inc = false;
      if (filter.tanggalMulai) {
        var tm = new Date(filter.tanggalMulai); tm.setHours(0,0,0,0);
        if (new Date(row.Tanggal) < tm) inc = false;
      }
      if (filter.tanggalAkhir) {
        var ta = new Date(filter.tanggalAkhir); ta.setHours(23,59,59,999);
        if (new Date(row.Tanggal) > ta) inc = false;
      }
      if (inc) {
        result.push({
          id: row.ID_Log, tanggal: formatTanggal(row.Tanggal), idProduk: row.ID_Produk,
          namaProduk: row.Nama_Produk, stokSebelum: Number(row.Stok_Sebelum),
          perubahan: Number(row.Perubahan), stokSesudah: Number(row.Stok_Sesudah),
          tipe: row.Tipe, keterangan: row.Keterangan, diupdateOleh: row.Diupdate_Oleh
        });
      }
    }
    result.reverse(); // Terbaru di atas
    if (filter.limit && result.length > filter.limit) result = result.slice(0, filter.limit);
    return successResponse(result);
  } catch (e) { return errorResponse('Gagal memuat log stok: ' + e.message); }
}

function getStokRendah(batas) {
  try {
    batas = batas || 10;
    var data = getProduk();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i].stok <= batas) result.push(data[i]);
    }
    result.sort(function(a,b){ return a.stok - b.stok; });
    return successResponse(result);
  } catch (e) { return errorResponse('Gagal memuat stok rendah: ' + e.message); }
}

function getRingkasanStok() {
  try {
    var data = getProduk();
    var ringkasan = {};
    var totalProduk = 0, totalStok = 0, stokRendah = 0;
    for (var i = 0; i < data.length; i++) {
      var kat = data[i].kategori;
      if (!ringkasan[kat]) ringkasan[kat] = { kategori: kat, jumlahProduk: 0, totalStok: 0 };
      ringkasan[kat].jumlahProduk++;
      ringkasan[kat].totalStok += data[i].stok;
      totalProduk++; totalStok += data[i].stok;
      if (data[i].stok <= 10) stokRendah++;
    }
    return successResponse({ perKategori: Object.values(ringkasan), totalProduk: totalProduk, totalStok: totalStok, stokRendah: stokRendah });
  } catch (e) { return errorResponse('Gagal memuat ringkasan: ' + e.message); }
}
