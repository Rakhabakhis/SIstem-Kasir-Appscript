/**
 * =====================================================
 * TRANSAKSI SERVICE - Proses Transaksi & Pembayaran
 * Sistem Kasir (POS) - Google Apps Script
 * =====================================================
 */

/**
 * Memproses transaksi pembayaran
 * @param {Object} transaksiData - Data transaksi dari frontend
 * {
 *   items: [{id, nama, harga, jumlah, catatan}],
 *   namaPelanggan: string,
 *   diskon: number,
 *   jenisPembayaran: string,
 *   jumlahBayar: number,
 *   catatan: string
 * }
 */
function prosesTransaksi(transaksiData) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  
  try {
    // Validasi
    if (!transaksiData.items || transaksiData.items.length === 0) {
      return errorResponse('Tidak ada item dalam pesanan');
    }
    
    if (!transaksiData.jenisPembayaran) {
      return errorResponse('Pilih jenis pembayaran');
    }
    
    // Hitung subtotal
    var subtotal = 0;
    var jumlahItem = 0;
    
    for (var i = 0; i < transaksiData.items.length; i++) {
      var item = transaksiData.items[i];
      subtotal += Number(item.harga) * Number(item.jumlah);
      jumlahItem += Number(item.jumlah);
    }
    
    // Hitung pajak
    var pajakPersen = Number(getSetting('pajak_persen')) || 0;
    var pajakAktif = getSetting('pajak_aktif') === 'Ya';
    var pajak = pajakAktif ? Math.round(subtotal * pajakPersen / 100) : 0;
    
    // Diskon
    var diskon = Number(transaksiData.diskon) || 0;
    
    // Total
    var total = subtotal + pajak - diskon;
    
    // Validasi pembayaran
    var jumlahBayar = Number(transaksiData.jumlahBayar) || 0;
    if (transaksiData.jenisPembayaran === 'Cash' && jumlahBayar < total) {
      return errorResponse('Jumlah pembayaran kurang. Total: ' + formatRupiah(total));
    }
    
    // Jika bukan cash, jumlah bayar = total
    if (transaksiData.jenisPembayaran !== 'Cash') {
      jumlahBayar = total;
    }
    
    var kembalian = jumlahBayar - total;
    
    // Validasi stok
    var produkData = getSheetData(SHEET_PRODUK);
    var produkMap = {};
    for (var j = 0; j < produkData.length; j++) {
      produkMap[produkData[j].ID_Produk] = produkData[j];
    }
    
    for (var k = 0; k < transaksiData.items.length; k++) {
      var itemCheck = transaksiData.items[k];
      var produk = produkMap[itemCheck.id];
      
      if (!produk) {
        return errorResponse('Produk "' + itemCheck.nama + '" tidak ditemukan');
      }
      
      if (Number(produk.Stok) < Number(itemCheck.jumlah)) {
        return errorResponse('Stok "' + itemCheck.nama + '" tidak cukup. Tersedia: ' + produk.Stok);
      }
    }
    
    // Generate ID Transaksi
    var idTransaksi = generateId('TRX', 'id_transaksi_terakhir');
    var now = new Date();
    var namaKasir = transaksiData.namaKasir || 'Kasir';
    
    // Simpan ke sheet Transaksi
    var sheetTransaksi = getSheet(SHEET_TRANSAKSI);
    sheetTransaksi.appendRow([
      idTransaksi,
      now,
      transaksiData.namaPelanggan || '-',
      jumlahItem,
      subtotal,
      pajak,
      diskon,
      total,
      transaksiData.jenisPembayaran,
      jumlahBayar,
      kembalian,
      'Selesai',
      namaKasir,
      transaksiData.catatan || ''
    ]);
    
    // Simpan detail & kurangi stok
    var sheetDetail = getSheet(SHEET_DETAIL_TRANSAKSI);
    var sheetProduk = getSheet(SHEET_PRODUK);
    
    for (var m = 0; m < transaksiData.items.length; m++) {
      var detailItem = transaksiData.items[m];
      var idDetail = generateId('DTL', 'id_detail_terakhir');
      var subtotalItem = Number(detailItem.harga) * Number(detailItem.jumlah);
      
      // Simpan detail transaksi
      sheetDetail.appendRow([
        idDetail,
        idTransaksi,
        detailItem.id,
        detailItem.nama,
        Number(detailItem.harga),
        Number(detailItem.jumlah),
        subtotalItem,
        detailItem.catatan || ''
      ]);
      
      // Kurangi stok
      var produkRow = produkMap[detailItem.id];
      var stokBaru = Number(produkRow.Stok) - Number(detailItem.jumlah);
      sheetProduk.getRange(produkRow._rowIndex, 5).setValue(stokBaru);
      
      // Log stok
      logStokChange(
        detailItem.id,
        detailItem.nama,
        Number(produkRow.Stok),
        -Number(detailItem.jumlah),
        stokBaru,
        'Keluar',
        'Penjualan ' + idTransaksi,
        namaKasir
      );
      
      // Update produkMap for next iteration
      produkMap[detailItem.id].Stok = stokBaru;
    }
    
    return successResponse({
      idTransaksi: idTransaksi,
      total: total,
      jumlahBayar: jumlahBayar,
      kembalian: kembalian,
      jenisPembayaran: transaksiData.jenisPembayaran,
      tanggal: formatTanggal(now),
      items: transaksiData.items,
      pajak: pajak,
      diskon: diskon,
      subtotal: subtotal
    }, 'Transaksi berhasil! ID: ' + idTransaksi);
    
  } catch (e) {
    return errorResponse('Gagal memproses transaksi: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mendapatkan transaksi terakhir
 * @param {number} limit - Jumlah transaksi yang diambil
 */
function getTransaksiTerakhir(limit) {
  try {
    limit = limit || 10;
    var data = getSheetDataSafe(SHEET_TRANSAKSI);
    if (!data || data.length === 0) return successResponse([]);
    
    // Filter baris valid dan sort by tanggal descending
    var valid = data.filter(function(r) { return r.ID_Transaksi && r.Tanggal; });
    valid.sort(function(a, b) {
      var da = a.Tanggal instanceof Date ? a.Tanggal : new Date(a.Tanggal);
      var db = b.Tanggal instanceof Date ? b.Tanggal : new Date(b.Tanggal);
      return db - da;
    });
    
    var result = [];
    var maxItems = Math.min(limit, valid.length);
    
    for (var i = 0; i < maxItems; i++) {
      result.push({
        id: valid[i].ID_Transaksi,
        tanggal: formatTanggal(valid[i].Tanggal),
        pelanggan: valid[i].Nama_Pelanggan || '-',
        jumlahItem: Number(valid[i].Jumlah_Item) || 0,
        total: Number(valid[i].Total) || 0,
        jenisPembayaran: valid[i].Jenis_Pembayaran || '-',
        status: valid[i].Status || '-'
      });
    }
    
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal memuat transaksi: ' + e.message);
  }
}

/**
 * Mendapatkan SEMUA transaksi untuk halaman Riwayat Transaksi
 * Terurut dari terbaru ke terlama
 */
function getSemuaTransaksi() {
  try {
    var data = getSheetDataSafe(SHEET_TRANSAKSI);
    if (!data || data.length === 0) return successResponse([]);

    var valid = data.filter(function(r) { return r.ID_Transaksi; });
    valid.sort(function(a, b) {
      var da = a.Tanggal instanceof Date ? a.Tanggal : new Date(a.Tanggal);
      var db = b.Tanggal instanceof Date ? b.Tanggal : new Date(b.Tanggal);
      return db - da;
    });

    var result = [];
    for (var i = 0; i < valid.length; i++) {
      result.push({
        id:              valid[i].ID_Transaksi,
        tanggal:         formatTanggal(valid[i].Tanggal),
        pelanggan:       valid[i].Nama_Pelanggan || '-',
        jumlahItem:      Number(valid[i].Jumlah_Item) || 0,
        subtotal:        Number(valid[i].Subtotal) || 0,
        pajak:           Number(valid[i].Pajak) || 0,
        total:           Number(valid[i].Total) || 0,
        jenisPembayaran: valid[i].Jenis_Pembayaran || '-',
        jumlahBayar:     Number(valid[i].Jumlah_Bayar) || 0,
        kembalian:       Number(valid[i].Kembalian) || 0,
        status:          valid[i].Status || '-',
        kasir:           valid[i].Kasir || '-'
      });
    }

    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal memuat semua transaksi: ' + e.message);
  }
}

/**
 * Menghapus transaksi secara permanen beserta detail-nya.
 * Jika transaksi berstatus Selesai, stok dikembalikan.
 * @param {string} idTransaksi
 * @param {boolean} kembalikanStok - true = kembalikan stok jika Selesai
 */
function hapusTransaksi(idTransaksi, kembalikanStok) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheetTrx    = getSheet(SHEET_TRANSAKSI);
    var sheetDetail = getSheet(SHEET_DETAIL_TRANSAKSI);
    var trxData     = getSheetDataSafe(SHEET_TRANSAKSI);
    var detailData  = getSheetDataSafe(SHEET_DETAIL_TRANSAKSI);

    // Cari baris transaksi
    var trxRow = null;
    for (var i = 0; i < trxData.length; i++) {
      if (trxData[i].ID_Transaksi === idTransaksi) { trxRow = trxData[i]; break; }
    }
    if (!trxRow) return errorResponse('Transaksi tidak ditemukan');

    // Kembalikan stok jika diminta dan transaksi sudah Selesai
    if (kembalikanStok && trxRow.Status === 'Selesai') {
      var sheetProduk = getSheet(SHEET_PRODUK);
      var produkData  = getSheetDataSafe(SHEET_PRODUK);
      var produkMap   = {};
      for (var p = 0; p < produkData.length; p++) {
        produkMap[produkData[p].ID_Produk] = produkData[p];
      }
      for (var d = 0; d < detailData.length; d++) {
        if (detailData[d].ID_Transaksi === idTransaksi) {
          var prod = produkMap[detailData[d].ID_Produk];
          if (prod) {
            var stokBaru = Number(prod.Stok) + Number(detailData[d].Jumlah);
            sheetProduk.getRange(prod._rowIndex, 5).setValue(stokBaru);
            logStokChange(detailData[d].ID_Produk, detailData[d].Nama_Produk,
              Number(prod.Stok), Number(detailData[d].Jumlah), stokBaru,
              'Masuk', 'Penghapusan transaksi ' + idTransaksi, 'Sistem');
            produkMap[detailData[d].ID_Produk].Stok = stokBaru;
          }
        }
      }
    }

    // Hapus baris detail transaksi (dari bawah ke atas agar rowIndex tidak bergeser)
    var detailRows = [];
    for (var j = 0; j < detailData.length; j++) {
      if (detailData[j].ID_Transaksi === idTransaksi) detailRows.push(detailData[j]._rowIndex);
    }
    detailRows.sort(function(a, b) { return b - a; }); // descending
    for (var r = 0; r < detailRows.length; r++) {
      sheetDetail.deleteRow(detailRows[r]);
    }

    // Hapus baris transaksi
    sheetTrx.deleteRow(trxRow._rowIndex);

    return successResponse(null, 'Transaksi ' + idTransaksi + ' berhasil dihapus');
  } catch (e) {
    return errorResponse('Gagal menghapus transaksi: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mendapatkan detail transaksi berdasarkan ID
 */
function getDetailTransaksi(idTransaksi) {
  try {
    // Get transaksi
    var transaksiData = getSheetData(SHEET_TRANSAKSI);
    var transaksi = null;
    
    for (var i = 0; i < transaksiData.length; i++) {
      if (transaksiData[i].ID_Transaksi === idTransaksi) {
        transaksi = {
          id: transaksiData[i].ID_Transaksi,
          tanggal: formatTanggal(transaksiData[i].Tanggal),
          pelanggan: transaksiData[i].Nama_Pelanggan,
          jumlahItem: Number(transaksiData[i].Jumlah_Item),
          subtotal: Number(transaksiData[i].Subtotal),
          pajak: Number(transaksiData[i].Pajak),
          diskon: Number(transaksiData[i].Diskon),
          total: Number(transaksiData[i].Total),
          jenisPembayaran: transaksiData[i].Jenis_Pembayaran,
          jumlahBayar: Number(transaksiData[i].Jumlah_Bayar),
          kembalian: Number(transaksiData[i].Kembalian),
          status: transaksiData[i].Status,
          kasir: transaksiData[i].Kasir,
          catatan: transaksiData[i].Catatan
        };
        break;
      }
    }
    
    if (!transaksi) {
      return errorResponse('Transaksi tidak ditemukan');
    }
    
    // Get detail items
    var detailData = getSheetData(SHEET_DETAIL_TRANSAKSI);
    var items = [];
    
    for (var j = 0; j < detailData.length; j++) {
      if (detailData[j].ID_Transaksi === idTransaksi) {
        items.push({
          idProduk: detailData[j].ID_Produk,
          nama: detailData[j].Nama_Produk,
          harga: Number(detailData[j].Harga_Satuan),
          jumlah: Number(detailData[j].Jumlah),
          subtotal: Number(detailData[j].Subtotal),
          catatan: detailData[j].Catatan_Item
        });
      }
    }
    
    transaksi.items = items;
    return successResponse(transaksi);
  } catch (e) {
    return errorResponse('Gagal memuat detail transaksi: ' + e.message);
  }
}

/**
 * Membatalkan transaksi dan mengembalikan stok
 */
function batalkanTransaksi(idTransaksi) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    var sheetTransaksi = getSheet(SHEET_TRANSAKSI);
    var transaksiData = getSheetData(SHEET_TRANSAKSI);
    var namaKasir = 'Admin';
    
    // Find transaksi
    var transaksiRow = null;
    for (var i = 0; i < transaksiData.length; i++) {
      if (transaksiData[i].ID_Transaksi === idTransaksi) {
        transaksiRow = transaksiData[i];
        break;
      }
    }
    
    if (!transaksiRow) {
      return errorResponse('Transaksi tidak ditemukan');
    }
    
    if (transaksiRow.Status === 'Dibatalkan') {
      return errorResponse('Transaksi sudah dibatalkan sebelumnya');
    }
    
    // Update status transaksi
    sheetTransaksi.getRange(transaksiRow._rowIndex, 12).setValue('Dibatalkan');
    
    // Kembalikan stok
    var detailData = getSheetData(SHEET_DETAIL_TRANSAKSI);
    var sheetProduk = getSheet(SHEET_PRODUK);
    var produkData = getSheetData(SHEET_PRODUK);
    var produkMap = {};
    
    for (var j = 0; j < produkData.length; j++) {
      produkMap[produkData[j].ID_Produk] = produkData[j];
    }
    
    for (var k = 0; k < detailData.length; k++) {
      if (detailData[k].ID_Transaksi === idTransaksi) {
        var produk = produkMap[detailData[k].ID_Produk];
        if (produk) {
          var stokSebelum = Number(produk.Stok);
          var stokBaru = stokSebelum + Number(detailData[k].Jumlah);
          sheetProduk.getRange(produk._rowIndex, 5).setValue(stokBaru);
          
          // Log pengembalian stok
          logStokChange(
            detailData[k].ID_Produk,
            detailData[k].Nama_Produk,
            stokSebelum,
            Number(detailData[k].Jumlah),
            stokBaru,
            'Masuk',
            'Pembatalan ' + idTransaksi,
            namaKasir
          );
          
          produkMap[detailData[k].ID_Produk].Stok = stokBaru;
        }
      }
    }
    
    return successResponse(null, 'Transaksi ' + idTransaksi + ' berhasil dibatalkan. Stok dikembalikan.');
  } catch (e) {
    return errorResponse('Gagal membatalkan transaksi: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}
