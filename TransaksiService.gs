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
    
    // Validasi stok bahan baku untuk produk yang punya resep
    var errBahan = validasiBahanBaku(transaksiData.items);
    if (errBahan) {
      return errorResponse(errBahan);
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
    
    // Simpan detail transaksi
    var sheetDetail = getSheet(SHEET_DETAIL_TRANSAKSI);
    
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
        buildCatatanItem(detailItem.catatan || '', detailItem.pilihanVarian || [])
      ]);
    }

    // Kurangi stok bahan baku (untuk produk yang punya resep)
    kurangiBahanBaku(transaksiData.items, idTransaksi, namaKasir);
    
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

    // Kembalikan stok bahan baku jika diminta dan transaksi sudah Selesai
    if (kembalikanStok && trxRow.Status === 'Selesai') {
      var itemsRestore = [];
      for (var d = 0; d < detailData.length; d++) {
        if (detailData[d].ID_Transaksi === idTransaksi) {
          var parsed = parseCatatanItem(detailData[d].Catatan_Item);
          itemsRestore.push({
            id: detailData[d].ID_Produk,
            nama: detailData[d].Nama_Produk,
            jumlah: Number(detailData[d].Jumlah) || 0,
            pilihanVarian: parsed.pilihanVarian || []
          });
        }
      }
      kembalikanBahanBaku(itemsRestore, idTransaksi, 'Sistem');
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
        var parsedItem = parseCatatanItem(detailData[j].Catatan_Item);
        items.push({
          idProduk: detailData[j].ID_Produk,
          nama: detailData[j].Nama_Produk,
          harga: Number(detailData[j].Harga_Satuan),
          jumlah: Number(detailData[j].Jumlah),
          subtotal: Number(detailData[j].Subtotal),
          catatan: parsedItem.catatan,
          pilihanVarian: parsedItem.pilihanVarian
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
 * Hanya mengembalikan bahan baku jika transaksi sudah berstatus Selesai
 */
function batalkanTransaksi(idTransaksi) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    var sheetTransaksi = getSheet(SHEET_TRANSAKSI);
    var transaksiData = getSheetData(SHEET_TRANSAKSI);
    var namaKasir = 'Admin';
    
    var transaksiRow = null;
    for (var i = 0; i < transaksiData.length; i++) {
      if (transaksiData[i].ID_Transaksi === idTransaksi) {
        transaksiRow = transaksiData[i];
        break;
      }
    }
    
    if (!transaksiRow) return errorResponse('Transaksi tidak ditemukan');
    if (transaksiRow.Status === 'Dibatalkan') return errorResponse('Transaksi sudah dibatalkan sebelumnya');
    
    var statusSebelumnya = transaksiRow.Status;
    sheetTransaksi.getRange(transaksiRow._rowIndex, 12).setValue('Dibatalkan');
    
    // Kembalikan stok bahan baku hanya jika transaksi sudah selesai (bahan sudah dikurangi)
    if (statusSebelumnya === 'Selesai') {
      var detailData = getSheetData(SHEET_DETAIL_TRANSAKSI);
      var itemsRestore = [];
      for (var k = 0; k < detailData.length; k++) {
        if (detailData[k].ID_Transaksi === idTransaksi) {
          var parsedRestore = parseCatatanItem(detailData[k].Catatan_Item);
          itemsRestore.push({
            id: detailData[k].ID_Produk,
            nama: detailData[k].Nama_Produk,
            jumlah: Number(detailData[k].Jumlah) || 0,
            pilihanVarian: parsedRestore.pilihanVarian || []
          });
        }
      }
      kembalikanBahanBaku(itemsRestore, idTransaksi, namaKasir);
      return successResponse(null, 'Transaksi ' + idTransaksi + ' berhasil dibatalkan. Stok bahan baku dikembalikan.');
    }
    
    return successResponse(null, 'Transaksi ' + idTransaksi + ' berhasil dibatalkan.');
  } catch (e) {
    return errorResponse('Gagal membatalkan transaksi: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Membuat pesanan baru (belum dibayar, belum kurangi stok bahan)
 * @param {Object} data - { items, namaPelanggan, diskon, namaKasir, catatan }
 */
function buatPesanan(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!data.items || data.items.length === 0) return errorResponse('Tidak ada item dalam pesanan');
    if (!data.namaPelanggan || !String(data.namaPelanggan).trim()) return errorResponse('Nama pelanggan wajib diisi');

    var subtotal = 0;
    var jumlahItem = 0;
    for (var i = 0; i < data.items.length; i++) {
      subtotal += Number(data.items[i].harga) * Number(data.items[i].jumlah);
      jumlahItem += Number(data.items[i].jumlah);
    }

    var pajakPersen = Number(getSetting('pajak_persen')) || 0;
    var pajakAktif = getSetting('pajak_aktif') === 'Ya';
    var pajak = pajakAktif ? Math.round(subtotal * pajakPersen / 100) : 0;
    var diskon = Number(data.diskon) || 0;
    var total = subtotal + pajak - diskon;
    var idPesanan = generateId('TRX', 'id_transaksi_terakhir');
    var now = new Date();
    var namaKasir = data.namaKasir || 'Kasir';

    var sheetTrx = getSheet(SHEET_TRANSAKSI);
    sheetTrx.appendRow([
      idPesanan, now, String(data.namaPelanggan).trim(),
      jumlahItem, subtotal, pajak, diskon, total,
      '-', 0, 0, 'Pesanan', namaKasir, data.catatan || ''
    ]);

    var sheetDetail = getSheet(SHEET_DETAIL_TRANSAKSI);
    for (var m = 0; m < data.items.length; m++) {
      var item = data.items[m];
      var idDetail = generateId('DTL', 'id_detail_terakhir');
      var subtotalItem = Number(item.harga) * Number(item.jumlah);
      sheetDetail.appendRow([
        idDetail, idPesanan, item.id, item.nama,
        Number(item.harga), Number(item.jumlah), subtotalItem,
        buildCatatanItem(item.catatan || '', item.pilihanVarian || [])
      ]);
    }

    return successResponse({
      idPesanan: idPesanan,
      namaPelanggan: String(data.namaPelanggan).trim(),
      total: total,
      jumlahItem: jumlahItem,
      tanggal: formatTanggal(now),
      items: data.items
    }, 'Pesanan ' + idPesanan + ' berhasil dibuat!');
  } catch (e) {
    return errorResponse('Gagal membuat pesanan: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mendapatkan semua pesanan yang belum dibayar (Status = 'Pesanan')
 */
function getDaftarPesanan() {
  try {
    var trxData = getSheetDataSafe(SHEET_TRANSAKSI);
    var detailData = getSheetDataSafe(SHEET_DETAIL_TRANSAKSI);

    var pesananList = trxData.filter(function(r) {
      return r.ID_Transaksi && r.Status === 'Pesanan';
    });
    pesananList.sort(function(a, b) {
      var da = a.Tanggal instanceof Date ? a.Tanggal : new Date(a.Tanggal);
      var db = b.Tanggal instanceof Date ? b.Tanggal : new Date(b.Tanggal);
      return db - da;
    });

    var result = [];
    for (var i = 0; i < pesananList.length; i++) {
      var p = pesananList[i];
      var items = [];
      for (var d = 0; d < detailData.length; d++) {
        if (detailData[d].ID_Transaksi === p.ID_Transaksi) {
          var parsed = parseCatatanItem(detailData[d].Catatan_Item);
          items.push({
            idProduk: detailData[d].ID_Produk,
            nama: detailData[d].Nama_Produk,
            harga: Number(detailData[d].Harga_Satuan) || 0,
            jumlah: Number(detailData[d].Jumlah) || 0,
            subtotal: Number(detailData[d].Subtotal) || 0,
            catatan: parsed.catatan || ''
          });
        }
      }
      result.push({
        id: p.ID_Transaksi,
        tanggal: formatTanggal(p.Tanggal),
        namaPelanggan: p.Nama_Pelanggan || '-',
        jumlahItem: Number(p.Jumlah_Item) || 0,
        subtotal: Number(p.Subtotal) || 0,
        pajak: Number(p.Pajak) || 0,
        diskon: Number(p.Diskon) || 0,
        total: Number(p.Total) || 0,
        kasir: p.Kasir || '-',
        items: items
      });
    }
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal memuat daftar pesanan: ' + e.message);
  }
}

/**
 * Menambahkan item baru ke pesanan yang sudah ada (gabungkan)
 * @param {string} idPesanan - ID pesanan target
 * @param {Array} items - Item dari keranjang yang akan ditambahkan
 * @param {string} namaKasir - Nama kasir
 */
function gabungkanPesanan(idPesanan, items, namaKasir) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!items || items.length === 0) return errorResponse('Tidak ada item untuk digabungkan');

    var trxData = getSheetDataSafe(SHEET_TRANSAKSI);
    var trxRow = null;
    for (var i = 0; i < trxData.length; i++) {
      if (trxData[i].ID_Transaksi === idPesanan && trxData[i].Status === 'Pesanan') {
        trxRow = trxData[i]; break;
      }
    }
    if (!trxRow) return errorResponse('Pesanan tidak ditemukan atau sudah diproses');

    var sheetDetail = getSheet(SHEET_DETAIL_TRANSAKSI);
    for (var m = 0; m < items.length; m++) {
      var item = items[m];
      var idDetail = generateId('DTL', 'id_detail_terakhir');
      var subtotalItem = Number(item.harga) * Number(item.jumlah);
      sheetDetail.appendRow([
        idDetail, idPesanan, item.id, item.nama,
        Number(item.harga), Number(item.jumlah), subtotalItem,
        buildCatatanItem(item.catatan || '', item.pilihanVarian || [])
      ]);
    }

    // Hitung ulang total dari semua item yang ada di DetailTransaksi
    var allDetail = getSheetDataSafe(SHEET_DETAIL_TRANSAKSI);
    var totalSubtotal = 0;
    var totalJumlah = 0;
    for (var d = 0; d < allDetail.length; d++) {
      if (allDetail[d].ID_Transaksi === idPesanan) {
        totalSubtotal += Number(allDetail[d].Subtotal) || 0;
        totalJumlah += Number(allDetail[d].Jumlah) || 0;
      }
    }

    var pajakPersen = Number(getSetting('pajak_persen')) || 0;
    var pajakAktif = getSetting('pajak_aktif') === 'Ya';
    var pajak = pajakAktif ? Math.round(totalSubtotal * pajakPersen / 100) : 0;
    var diskon = Number(trxRow.Diskon) || 0;
    var total = totalSubtotal + pajak - diskon;

    var sheetTrx = getSheet(SHEET_TRANSAKSI);
    sheetTrx.getRange(trxRow._rowIndex, 4).setValue(totalJumlah);
    sheetTrx.getRange(trxRow._rowIndex, 5).setValue(totalSubtotal);
    sheetTrx.getRange(trxRow._rowIndex, 6).setValue(pajak);
    sheetTrx.getRange(trxRow._rowIndex, 8).setValue(total);

    return successResponse({
      idPesanan: idPesanan,
      total: total,
      jumlahItem: totalJumlah
    }, 'Item berhasil ditambahkan ke pesanan ' + idPesanan);
  } catch (e) {
    return errorResponse('Gagal menggabungkan pesanan: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Memproses pembayaran untuk pesanan yang sudah ada (Status='Pesanan' -> 'Selesai')
 * Baru di sini bahan baku dikurangi
 * @param {string} idPesanan
 * @param {Object} paymentData - { jenisPembayaran, jumlahBayar, namaKasir }
 */
function prosesPembayaranPesanan(idPesanan, paymentData) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var trxData = getSheetDataSafe(SHEET_TRANSAKSI);
    var trxRow = null;
    for (var i = 0; i < trxData.length; i++) {
      if (trxData[i].ID_Transaksi === idPesanan && trxData[i].Status === 'Pesanan') {
        trxRow = trxData[i]; break;
      }
    }
    if (!trxRow) return errorResponse('Pesanan tidak ditemukan atau sudah diproses');

    var total = Number(trxRow.Total) || 0;
    var jenisPembayaran = paymentData.jenisPembayaran || 'Cash';
    var jumlahBayar = Number(paymentData.jumlahBayar) || 0;

    if (jenisPembayaran === 'Cash' && jumlahBayar < total) {
      return errorResponse('Jumlah pembayaran kurang. Total: ' + formatRupiah(total));
    }
    if (jenisPembayaran !== 'Cash') jumlahBayar = total;
    var kembalian = jumlahBayar - total;
    var namaKasir = paymentData.namaKasir || 'Kasir';

    var detailData = getSheetDataSafe(SHEET_DETAIL_TRANSAKSI);
    var items = [];
    for (var d = 0; d < detailData.length; d++) {
      if (detailData[d].ID_Transaksi === idPesanan) {
        var parsed = parseCatatanItem(detailData[d].Catatan_Item);
        items.push({
          id: detailData[d].ID_Produk,
          nama: detailData[d].Nama_Produk,
          jumlah: Number(detailData[d].Jumlah) || 0,
          harga: Number(detailData[d].Harga_Satuan) || 0,
          pilihanVarian: parsed.pilihanVarian || []
        });
      }
    }

    var errBahan = validasiBahanBaku(items);
    if (errBahan) return errorResponse(errBahan);

    var sheetTrx = getSheet(SHEET_TRANSAKSI);
    sheetTrx.getRange(trxRow._rowIndex, 9).setValue(jenisPembayaran);
    sheetTrx.getRange(trxRow._rowIndex, 10).setValue(jumlahBayar);
    sheetTrx.getRange(trxRow._rowIndex, 11).setValue(kembalian);
    sheetTrx.getRange(trxRow._rowIndex, 12).setValue('Selesai');
    sheetTrx.getRange(trxRow._rowIndex, 13).setValue(namaKasir);

    kurangiBahanBaku(items, idPesanan, namaKasir);

    return successResponse({
      idTransaksi: idPesanan,
      total: total,
      jumlahBayar: jumlahBayar,
      kembalian: kembalian,
      jenisPembayaran: jenisPembayaran,
      subtotal: Number(trxRow.Subtotal) || 0,
      pajak: Number(trxRow.Pajak) || 0,
      diskon: Number(trxRow.Diskon) || 0,
      tanggal: formatTanggal(new Date()),
      items: items
    }, 'Pembayaran pesanan ' + idPesanan + ' berhasil!');
  } catch (e) {
    return errorResponse('Gagal memproses pembayaran: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Membatalkan pesanan pending (tidak mengembalikan stok karena belum pernah dikurangi)
 * @param {string} idPesanan
 */
function batalkanPesanan(idPesanan) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var trxData = getSheetDataSafe(SHEET_TRANSAKSI);
    var trxRow = null;
    for (var i = 0; i < trxData.length; i++) {
      if (trxData[i].ID_Transaksi === idPesanan) { trxRow = trxData[i]; break; }
    }
    if (!trxRow) return errorResponse('Pesanan tidak ditemukan');
    if (trxRow.Status !== 'Pesanan') return errorResponse('Pesanan tidak dalam status aktif');

    var sheetTrx = getSheet(SHEET_TRANSAKSI);
    sheetTrx.getRange(trxRow._rowIndex, 12).setValue('Dibatalkan');

    return successResponse(null, 'Pesanan ' + idPesanan + ' berhasil dibatalkan.');
  } catch (e) {
    return errorResponse('Gagal membatalkan pesanan: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function buildCatatanItem(catatan, pilihanVarian) {
  var cleanCatatan = String(catatan || '').trim();
  var cleanVarian = pilihanVarian || [];
  if (!cleanVarian.length) return cleanCatatan;
  return cleanCatatan + ' [[VARIANT_JSON]]' + JSON.stringify(cleanVarian);
}

function parseCatatanItem(rawCatatan) {
  var raw = String(rawCatatan || '');
  var marker = '[[VARIANT_JSON]]';
  var idx = raw.indexOf(marker);
  if (idx === -1) {
    return { catatan: raw, pilihanVarian: [] };
  }
  var catatan = raw.substring(0, idx).trim();
  var jsonText = raw.substring(idx + marker.length);
  try {
    return {
      catatan: catatan,
      pilihanVarian: JSON.parse(jsonText) || []
    };
  } catch (e) {
    return { catatan: catatan, pilihanVarian: [] };
  }
}
