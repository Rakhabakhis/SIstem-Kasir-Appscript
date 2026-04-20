const ExcelJS = require('exceljs');
const path = require('path');

async function createDatabase() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Sistem Kasir POS';
  workbook.created = new Date();

  // === Style Constants ===
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D9488' } };
  const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Inter' };
  const headerAlignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const headerBorder = {
    top: { style: 'thin', color: { argb: 'FF0F766E' } },
    bottom: { style: 'thin', color: { argb: 'FF0F766E' } },
    left: { style: 'thin', color: { argb: 'FF0F766E' } },
    right: { style: 'thin', color: { argb: 'FF0F766E' } }
  };
  const dataBorder = {
    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
  };
  const dataFont = { size: 10, name: 'Inter' };
  const currencyFormat = '#,##0';

  function styleHeader(sheet) {
    const row = sheet.getRow(1);
    row.height = 28;
    row.eachCell(function(cell) {
      cell.fill = headerFill;
      cell.font = headerFont;
      cell.alignment = headerAlignment;
      cell.border = headerBorder;
    });
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
  }

  function styleDataRow(row) {
    row.eachCell(function(cell) {
      cell.font = dataFont;
      cell.border = dataBorder;
      cell.alignment = { vertical: 'middle' };
    });
  }

  // ============================================================
  // SHEET 1: PRODUK
  // ============================================================
  const sheetProduk = workbook.addWorksheet('Produk', {
    properties: { tabColor: { argb: 'FF0D9488' } }
  });

  sheetProduk.columns = [
    { header: 'ID_Produk', key: 'id', width: 14 },
    { header: 'Nama_Produk', key: 'nama', width: 25 },
    { header: 'Kategori', key: 'kategori', width: 15 },
    { header: 'Harga', key: 'harga', width: 14 },
    { header: 'Stok', key: 'stok', width: 10 },
    { header: 'Deskripsi', key: 'deskripsi', width: 30 },
    { header: 'URL_Gambar', key: 'gambar', width: 45 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Tanggal_Dibuat', key: 'dibuat', width: 18 },
    { header: 'Tanggal_Diupdate', key: 'diupdate', width: 18 }
  ];

  const now = new Date();
  const produkData = [
    { id: 'PRD-0001', nama: 'Classic Beef Burger', kategori: 'Makanan', harga: 35000, stok: 24, deskripsi: 'Daging sapi asli, keju, sayur', gambar: 'https://images.unsplash.com/photo-1568901346375-23c9450c58cd?w=300&h=300&fit=crop', status: 'Aktif', dibuat: now, diupdate: now },
    { id: 'PRD-0002', nama: 'Double Cheese Burger', kategori: 'Makanan', harga: 55000, stok: 18, deskripsi: 'Ekstra keju lumer', gambar: 'https://images.unsplash.com/photo-1553979459-d2229ba7433b?w=300&h=300&fit=crop', status: 'Aktif', dibuat: now, diupdate: now },
    { id: 'PRD-0003', nama: 'French Fries Large', kategori: 'Snack', harga: 20000, stok: 50, deskripsi: 'Kentang goreng renyah', gambar: 'https://images.unsplash.com/photo-1573080496219-bb080dd4f877?w=300&h=300&fit=crop', status: 'Aktif', dibuat: now, diupdate: now },
    { id: 'PRD-0004', nama: 'Iced Caramel Latte', kategori: 'Minuman', harga: 28000, stok: 45, deskripsi: 'Kopi susu karamel dingin', gambar: 'https://images.unsplash.com/photo-1461023058943-07fcbe16d735?w=300&h=300&fit=crop', status: 'Aktif', dibuat: now, diupdate: now },
    { id: 'PRD-0005', nama: 'Iced Lemon Tea', kategori: 'Minuman', harga: 15000, stok: 60, deskripsi: 'Teh lemon segar dingin', gambar: 'https://images.unsplash.com/photo-1556679343-c7306c1976bc?w=300&h=300&fit=crop', status: 'Aktif', dibuat: now, diupdate: now },
    { id: 'PRD-0006', nama: 'Iced Americano', kategori: 'Minuman', harga: 22000, stok: 40, deskripsi: 'Kopi hitam dingin', gambar: 'https://images.unsplash.com/photo-1517701550927-30cf4ba1dba5?w=300&h=300&fit=crop', status: 'Aktif', dibuat: now, diupdate: now },
    { id: 'PRD-0007', nama: 'Chicken Wings (6 pcs)', kategori: 'Makanan', harga: 32000, stok: 30, deskripsi: 'Sayap ayam goreng krispy', gambar: 'https://images.unsplash.com/photo-1567620832903-9fc6debc209f?w=300&h=300&fit=crop', status: 'Aktif', dibuat: now, diupdate: now },
    { id: 'PRD-0008', nama: 'Chocolate Milkshake', kategori: 'Minuman', harga: 25000, stok: 35, deskripsi: 'Milkshake coklat premium', gambar: 'https://images.unsplash.com/photo-1572490122747-3968b75cc699?w=300&h=300&fit=crop', status: 'Aktif', dibuat: now, diupdate: now },
    { id: 'PRD-0009', nama: 'Brownies', kategori: 'Dessert', harga: 18000, stok: 20, deskripsi: 'Brownies coklat lembut', gambar: 'https://images.unsplash.com/photo-1606313564200-e75d5e30476c?w=300&h=300&fit=crop', status: 'Aktif', dibuat: now, diupdate: now },
    { id: 'PRD-0010', nama: 'Paket Hemat 1', kategori: 'Paket Hemat', harga: 45000, stok: 99, deskripsi: 'Burger + Fries + Drink', gambar: 'https://images.unsplash.com/photo-1594212699903-ec8a3eca50f5?w=300&h=300&fit=crop', status: 'Aktif', dibuat: now, diupdate: now }
  ];

  produkData.forEach(d => {
    const row = sheetProduk.addRow(d);
    styleDataRow(row);
  });

  // Format currency column
  sheetProduk.getColumn('harga').numFmt = currencyFormat;
  // Format date columns
  sheetProduk.getColumn('dibuat').numFmt = 'dd/mm/yyyy hh:mm';
  sheetProduk.getColumn('diupdate').numFmt = 'dd/mm/yyyy hh:mm';
  // Data validation for Kategori
  for (let i = 2; i <= 100; i++) {
    sheetProduk.getCell(`C${i}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['"Makanan,Minuman,Snack,Dessert,Paket Hemat"']
    };
    sheetProduk.getCell(`H${i}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['"Aktif,Nonaktif"']
    };
  }
  // Alternate row coloring
  for (let i = 2; i <= produkData.length + 1; i++) {
    if (i % 2 === 0) {
      sheetProduk.getRow(i).eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      });
    }
  }
  styleHeader(sheetProduk);

  // ============================================================
  // SHEET 2: TRANSAKSI
  // ============================================================
  const sheetTransaksi = workbook.addWorksheet('Transaksi', {
    properties: { tabColor: { argb: 'FF3B82F6' } }
  });

  sheetTransaksi.columns = [
    { header: 'ID_Transaksi', key: 'id', width: 14 },
    { header: 'Tanggal', key: 'tanggal', width: 18 },
    { header: 'Nama_Pelanggan', key: 'pelanggan', width: 20 },
    { header: 'Jumlah_Item', key: 'jumlah', width: 14 },
    { header: 'Subtotal', key: 'subtotal', width: 14 },
    { header: 'Pajak', key: 'pajak', width: 12 },
    { header: 'Diskon', key: 'diskon', width: 12 },
    { header: 'Total', key: 'total', width: 14 },
    { header: 'Jenis_Pembayaran', key: 'pembayaran', width: 18 },
    { header: 'Jumlah_Bayar', key: 'bayar', width: 14 },
    { header: 'Kembalian', key: 'kembalian', width: 14 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Kasir', key: 'kasir', width: 15 },
    { header: 'Catatan', key: 'catatan', width: 20 }
  ];

  // Format columns
  ['subtotal','pajak','diskon','total','bayar','kembalian'].forEach(key => {
    sheetTransaksi.getColumn(key).numFmt = currencyFormat;
  });
  sheetTransaksi.getColumn('tanggal').numFmt = 'dd/mm/yyyy hh:mm';
  // Data validation
  for (let i = 2; i <= 500; i++) {
    sheetTransaksi.getCell(`I${i}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['"Cash,QRIS"']
    };
    sheetTransaksi.getCell(`L${i}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['"Selesai,Pending,Dibatalkan"']
    };
  }
  styleHeader(sheetTransaksi);

  // ============================================================
  // SHEET 3: DETAIL TRANSAKSI
  // ============================================================
  const sheetDetail = workbook.addWorksheet('DetailTransaksi', {
    properties: { tabColor: { argb: 'FF8B5CF6' } }
  });

  sheetDetail.columns = [
    { header: 'ID_Detail', key: 'id', width: 14 },
    { header: 'ID_Transaksi', key: 'idTrx', width: 14 },
    { header: 'ID_Produk', key: 'idProduk', width: 14 },
    { header: 'Nama_Produk', key: 'nama', width: 25 },
    { header: 'Harga_Satuan', key: 'harga', width: 14 },
    { header: 'Jumlah', key: 'jumlah', width: 10 },
    { header: 'Subtotal', key: 'subtotal', width: 14 },
    { header: 'Catatan_Item', key: 'catatan', width: 25 }
  ];

  sheetDetail.getColumn('harga').numFmt = currencyFormat;
  sheetDetail.getColumn('subtotal').numFmt = currencyFormat;
  styleHeader(sheetDetail);

  // ============================================================
  // SHEET 4: LOG STOK
  // ============================================================
  const sheetLog = workbook.addWorksheet('LogStok', {
    properties: { tabColor: { argb: 'FFF59E0B' } }
  });

  sheetLog.columns = [
    { header: 'ID_Log', key: 'id', width: 14 },
    { header: 'Tanggal', key: 'tanggal', width: 18 },
    { header: 'ID_Produk', key: 'idProduk', width: 14 },
    { header: 'Nama_Produk', key: 'nama', width: 25 },
    { header: 'Stok_Sebelum', key: 'sebelum', width: 14 },
    { header: 'Perubahan', key: 'perubahan', width: 12 },
    { header: 'Stok_Sesudah', key: 'sesudah', width: 14 },
    { header: 'Tipe', key: 'tipe', width: 14 },
    { header: 'Keterangan', key: 'keterangan', width: 25 },
    { header: 'Diupdate_Oleh', key: 'oleh', width: 16 }
  ];

  sheetLog.getColumn('tanggal').numFmt = 'dd/mm/yyyy hh:mm';
  for (let i = 2; i <= 500; i++) {
    sheetLog.getCell(`H${i}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['"Masuk,Keluar,Penyesuaian"']
    };
  }
  styleHeader(sheetLog);

  // ============================================================
  // SHEET 5: PENGATURAN
  // ============================================================
  const sheetPengaturan = workbook.addWorksheet('Pengaturan', {
    properties: { tabColor: { argb: 'FF64748B' } }
  });

  sheetPengaturan.columns = [
    { header: 'Key', key: 'key', width: 28 },
    { header: 'Value', key: 'value', width: 35 }
  ];

  const pengaturanData = [
    { key: 'nama_toko', value: 'Toko Saya' },
    { key: 'alamat', value: 'Jl. Contoh No. 1' },
    { key: 'telepon', value: '0812-xxxx-xxxx' },
    { key: 'pajak_persen', value: 10 },
    { key: 'pajak_aktif', value: 'Ya' },
    { key: 'id_transaksi_terakhir', value: 0 },
    { key: 'id_produk_terakhir', value: 10 },
    { key: 'id_log_terakhir', value: 0 },
    { key: 'id_detail_terakhir', value: 0 },
    { key: 'id_kasir_terakhir', value: 1 }
  ];

  pengaturanData.forEach(d => {
    const row = sheetPengaturan.addRow(d);
    styleDataRow(row);
    row.getCell('key').font = { ...dataFont, bold: true, color: { argb: 'FF0D9488' } };
  });

  // Alternate rows
  for (let i = 2; i <= pengaturanData.length + 1; i++) {
    if (i % 2 === 0) {
      sheetPengaturan.getRow(i).eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      });
    }
  }
  styleHeader(sheetPengaturan);

  // ============================================================
  // SHEET 6: KASIR
  // ============================================================
  const sheetKasir = workbook.addWorksheet('Kasir', {
    properties: { tabColor: { argb: 'FF22C55E' } }
  });

  sheetKasir.columns = [
    { header: 'ID_Kasir', key: 'id', width: 14 },
    { header: 'Nama', key: 'nama', width: 25 },
    { header: 'PIN', key: 'pin', width: 10 },
    { header: 'Role', key: 'role', width: 12 },
    { header: 'Status', key: 'status', width: 12 }
  ];

  const kasirData = [
    { id: 'KSR-0001', nama: 'Admin', pin: '1234', role: 'Admin', status: 'Aktif' }
  ];

  kasirData.forEach(d => {
    const row = sheetKasir.addRow(d);
    styleDataRow(row);
  });

  for (let i = 2; i <= 50; i++) {
    sheetKasir.getCell(`D${i}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['"Admin,Kasir"']
    };
    sheetKasir.getCell(`E${i}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['"Aktif,Nonaktif"']
    };
  }
  styleHeader(sheetKasir);

  // ============================================================
  // SAVE FILE
  // ============================================================
  const filePath = path.join(__dirname, 'Database_Kasir_Template.xlsx');
  await workbook.xlsx.writeFile(filePath);
  console.log('');
  console.log('✅ File berhasil dibuat: Database_Kasir_Template.xlsx');
  console.log('');
  console.log('📋 Isi file:');
  console.log('   Sheet 1: Produk        (10 produk contoh)');
  console.log('   Sheet 2: Transaksi     (kosong, siap pakai)');
  console.log('   Sheet 3: DetailTransaksi (kosong, siap pakai)');
  console.log('   Sheet 4: LogStok       (kosong, siap pakai)');
  console.log('   Sheet 5: Pengaturan    (10 pengaturan default)');
  console.log('   Sheet 6: Kasir         (1 admin, PIN: 1234)');
  console.log('');
  console.log('📁 Lokasi: ' + filePath);
  console.log('');
  console.log('🚀 Langkah selanjutnya:');
  console.log('   1. Buka Google Sheets');
  console.log('   2. File → Import → Upload → Pilih file ini');
  console.log('   3. Pilih "Replace spreadsheet"');
  console.log('   4. Buka Extensions → Apps Script');
  console.log('   5. Copy semua file .gs dan .html');
  console.log('   6. Deploy sebagai Web App');
}

createDatabase().catch(err => {
  console.error('❌ Error:', err.message);
});
