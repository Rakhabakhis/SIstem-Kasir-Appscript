/**
 * =====================================================
 * LAPORAN SERVICE - Generate Laporan Transaksi
 * Sistem Kasir (POS) - Google Apps Script
 * =====================================================
 */

function getLaporan(filter) {
  try {
    filter = filter || {};
    var data = getSheetDataSafe(SHEET_TRANSAKSI);
    var result = [];
    
    // Parse filter dates sekali saja
    var tmStart = null, tmEnd = null;
    if (filter.tanggalMulai) {
      tmStart = new Date(filter.tanggalMulai);
      tmStart.setHours(0, 0, 0, 0);
    }
    if (filter.tanggalAkhir) {
      tmEnd = new Date(filter.tanggalAkhir);
      tmEnd.setHours(23, 59, 59, 999);
    }
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row.Status === 'Dibatalkan') continue;
      if (!row.ID_Transaksi) continue; // skip baris kosong
      
      // Tanggal dari Sheets bisa berupa Date object atau string
      var tglTransaksi;
      if (row.Tanggal instanceof Date) {
        tglTransaksi = row.Tanggal;
      } else if (row.Tanggal) {
        tglTransaksi = new Date(row.Tanggal);
      } else {
        continue; // skip jika tanggal kosong
      }
      
      if (isNaN(tglTransaksi.getTime())) continue; // skip Invalid Date
      
      var inc = true;
      if (tmStart && tglTransaksi < tmStart) inc = false;
      if (tmEnd && tglTransaksi > tmEnd) inc = false;
      if (filter.jenisPembayaran && row.Jenis_Pembayaran !== filter.jenisPembayaran) inc = false;
      
      if (inc) {
        result.push({
          id: row.ID_Transaksi,
          tanggal: formatTanggal(tglTransaksi),
          tanggalRaw: tglTransaksi,
          pelanggan: row.Nama_Pelanggan || '-',
          jumlahItem: Number(row.Jumlah_Item) || 0,
          subtotal: Number(row.Subtotal) || 0,
          pajak: Number(row.Pajak) || 0,
          diskon: Number(row.Diskon) || 0,
          total: Number(row.Total) || 0,
          jenisPembayaran: row.Jenis_Pembayaran || '-',
          status: row.Status,
          kasir: row.Kasir || '-'
        });
      }
    }
    
    result.sort(function(a, b) { return b.tanggalRaw - a.tanggalRaw; });
    
    // Hitung ringkasan
    var totalTransaksi = result.length;
    var totalPendapatan = 0, totalPajak = 0, totalDiskon = 0;
    var pembayaran = {};
    
    for (var j = 0; j < result.length; j++) {
      totalPendapatan += result[j].total;
      totalPajak += result[j].pajak;
      totalDiskon += result[j].diskon;
      var jp = result[j].jenisPembayaran;
      if (!pembayaran[jp]) pembayaran[jp] = { metode: jp, jumlah: 0, total: 0 };
      pembayaran[jp].jumlah++;
      pembayaran[jp].total += result[j].total;
    }
    
    // Convert tanggalRaw ke string untuk serialisasi
    for (var k = 0; k < result.length; k++) {
      delete result[k].tanggalRaw;
    }
    
    return successResponse({
      transaksi: result,
      ringkasan: {
        totalTransaksi: totalTransaksi,
        totalPendapatan: totalPendapatan,
        totalPajak: totalPajak,
        totalDiskon: totalDiskon,
        rataRata: totalTransaksi > 0 ? Math.round(totalPendapatan / totalTransaksi) : 0
      },
      pembayaran: Object.values(pembayaran)
    });
  } catch (e) {
    return errorResponse('Gagal memuat laporan: ' + e.message);
  }
}

function getLaporanHarian(tanggal) {
  var tgl = tanggal ? new Date(tanggal) : new Date();
  return getLaporan({ tanggalMulai: tgl, tanggalAkhir: tgl });
}

function getLaporanMingguan(tanggalMulai) {
  var mulai = tanggalMulai ? new Date(tanggalMulai) : new Date();
  var dayOfWeek = mulai.getDay();
  mulai.setDate(mulai.getDate() - dayOfWeek);
  var akhir = new Date(mulai);
  akhir.setDate(akhir.getDate() + 6);
  return getLaporan({ tanggalMulai: mulai, tanggalAkhir: akhir });
}

function getLaporanBulanan(bulan, tahun) {
  var now = new Date();
  bulan = (bulan !== undefined && bulan !== null) ? bulan : now.getMonth();
  tahun = tahun || now.getFullYear();
  var mulai = new Date(tahun, bulan, 1);
  var akhir = new Date(tahun, bulan + 1, 0);
  return getLaporan({ tanggalMulai: mulai, tanggalAkhir: akhir });
}

function getProdukTerlaris(filter) {
  try {
    filter = filter || {};
    var detailData = getSheetData(SHEET_DETAIL_TRANSAKSI);
    var transaksiData = getSheetData(SHEET_TRANSAKSI);
    var produkData = getSheetDataSafe(SHEET_PRODUK);
    
    // Build produk kategori map
    var kategoriMap = {};
    for (var p = 0; p < produkData.length; p++) {
      kategoriMap[produkData[p].ID_Produk] = produkData[p].Kategori || '-';
    }
    
    // Build set of valid transaction IDs
    var validTrx = {};
    for (var i = 0; i < transaksiData.length; i++) {
      var row = transaksiData[i];
      if (row.Status === 'Dibatalkan') continue;
      var tgl = new Date(row.Tanggal);
      var inc = true;
      if (filter.tanggalMulai) {
        var tm = new Date(filter.tanggalMulai); tm.setHours(0,0,0,0);
        if (tgl < tm) inc = false;
      }
      if (filter.tanggalAkhir) {
        var ta = new Date(filter.tanggalAkhir); ta.setHours(23,59,59,999);
        if (tgl > ta) inc = false;
      }
      if (inc) validTrx[row.ID_Transaksi] = true;
    }
    
    var produkStats = {};
    for (var j = 0; j < detailData.length; j++) {
      var detail = detailData[j];
      if (!validTrx[detail.ID_Transaksi]) continue;
      var pid = detail.ID_Produk;
      if (!produkStats[pid]) {
        produkStats[pid] = {
          id: pid,
          nama: detail.Nama_Produk,
          kategori: kategoriMap[pid] || '-',
          terjual: 0,
          pendapatan: 0
        };
      }
      produkStats[pid].terjual += Number(detail.Jumlah);
      produkStats[pid].pendapatan += Number(detail.Subtotal);
    }
    
    var result = Object.values(produkStats);
    result.sort(function(a,b){ return b.terjual - a.terjual; });
    
    if (filter.limit) {
      result = result.slice(0, filter.limit);
    }
    
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal memuat produk terlaris: ' + e.message);
  }
}

/**
 * Mendapatkan laporan item terjual dengan filter kategori dan jangka waktu
 * @param {Object} filter - {periode: '7hari'|'1bulan'|'3bulan', kategori: string}
 */
function getLaporanItemTerjual(filter) {
  try {
    filter = filter || {};
    var now = new Date();
    var tanggalAkhir = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var tanggalMulai;
    
    switch (filter.periode) {
      case '7hari':
        tanggalMulai = new Date(tanggalAkhir);
        tanggalMulai.setDate(tanggalMulai.getDate() - 6);
        break;
      case '3bulan':
        tanggalMulai = new Date(tanggalAkhir);
        tanggalMulai.setMonth(tanggalMulai.getMonth() - 3);
        break;
      case '1bulan':
      default:
        tanggalMulai = new Date(tanggalAkhir);
        tanggalMulai.setMonth(tanggalMulai.getMonth() - 1);
        break;
    }
    
    var produkResult = getProdukTerlaris({
      tanggalMulai: tanggalMulai,
      tanggalAkhir: tanggalAkhir
    });
    
    if (!produkResult.success) return produkResult;
    
    var items = produkResult.data;
    
    // Filter by kategori if specified
    if (filter.kategori) {
      items = items.filter(function(item) {
        return item.kategori === filter.kategori;
      });
    }
    
    // Calculate totals
    var totalItem = 0;
    var totalPendapatan = 0;
    for (var i = 0; i < items.length; i++) {
      totalItem += items[i].terjual;
      totalPendapatan += items[i].pendapatan;
    }
    
    // Get unique categories for filter
    var allItems = produkResult.data;
    var kategoriSet = {};
    for (var k = 0; k < allItems.length; k++) {
      if (allItems[k].kategori && allItems[k].kategori !== '-') {
        kategoriSet[allItems[k].kategori] = true;
      }
    }
    
    return successResponse({
      items: items,
      kategoriList: Object.keys(kategoriSet).sort(),
      ringkasan: {
        totalJenisProduk: items.length,
        totalItemTerjual: totalItem,
        totalPendapatan: totalPendapatan
      }
    });
  } catch (e) {
    return errorResponse('Gagal memuat laporan item terjual: ' + e.message);
  }
}

function getGrafikPendapatan(filter) {
  try {
    filter = filter || {};
    var data = getSheetData(SHEET_TRANSAKSI);
    var harian = {};
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row.Status === 'Dibatalkan') continue;
      var tgl = new Date(row.Tanggal);
      var inc = true;
      if (filter.tanggalMulai) {
        var tm = new Date(filter.tanggalMulai); tm.setHours(0,0,0,0);
        if (tgl < tm) inc = false;
      }
      if (filter.tanggalAkhir) {
        var ta = new Date(filter.tanggalAkhir); ta.setHours(23,59,59,999);
        if (tgl > ta) inc = false;
      }
      if (inc) {
        var key = formatTanggalSingkat(tgl);
        if (!harian[key]) harian[key] = { tanggal: key, total: 0, transaksi: 0 };
        harian[key].total += Number(row.Total);
        harian[key].transaksi++;
      }
    }
    
    var result = Object.values(harian);
    result.sort(function(a,b){ 
      var da = a.tanggal.split('/'); var db = b.tanggal.split('/');
      return new Date(da[2],da[1]-1,da[0]) - new Date(db[2],db[1]-1,db[0]);
    });
    
    return successResponse(result);
  } catch (e) {
    return errorResponse('Gagal memuat grafik: ' + e.message);
  }
}
