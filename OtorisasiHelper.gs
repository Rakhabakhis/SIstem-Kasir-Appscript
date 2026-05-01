/**
 * =====================================================
 * OTORISASI HELPER - Manajemen Izin MailApp
 * =====================================================
 * File ini berisi fungsi bantu untuk mengatasi masalah
 * otorisasi MailApp di Google Apps Script.
 *
 * URUTAN PENGGUNAAN:
 *   1. Jalankan "dapatkanUrlOtorisasi" dari GAS Editor
 *   2. Buka URL yang muncul di Log Eksekusi lewat browser
 *   3. Klik Allow / Izinkan
 *   4. Jalankan "inisialisasiMailApp" untuk verifikasi
 * =====================================================
 */

/**
 * STEP 1 — Jalankan fungsi ini pertama kali.
 *
 * Memeriksa status otorisasi dan menampilkan URL yang perlu
 * dibuka di browser untuk memberikan izin MailApp (send email).
 *
 * Fungsi ini TIDAK memerlukan scope apapun, sehingga selalu
 * bisa dijalankan meski otorisasi belum diberikan sama sekali.
 */
function dapatkanUrlOtorisasi() {
  try {
    // ScriptApp.getAuthorizationInfo tidak memerlukan scope apapun
    var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    var status = authInfo.getAuthorizationStatus();

    Logger.log('============================================');
    Logger.log('CEK STATUS OTORISASI MAILAPP');
    Logger.log('============================================');

    if (status === ScriptApp.AuthorizationStatus.REQUIRED) {
      var url = authInfo.getAuthorizationUrl();
      Logger.log('STATUS : ❌ Otorisasi BELUM diberikan');
      Logger.log('');
      Logger.log('LANGKAH SELANJUTNYA:');
      Logger.log('1. Salin URL di bawah ini');
      Logger.log('2. Buka di tab browser baru');
      Logger.log('3. Pilih akun Google Anda');
      Logger.log('4. Klik "Allow" / "Izinkan"');
      Logger.log('5. Kembali ke sini, jalankan "inisialisasiMailApp"');
      Logger.log('');
      Logger.log('URL OTORISASI:');
      Logger.log(url);
      Logger.log('');

      try {
        var ui = SpreadsheetApp.getUi();
        ui.alert(
          '🔐 Otorisasi MailApp Diperlukan',
          'Token OAuth belum mencakup izin "Send Email".\n\n' +
          '✅ URL otorisasi sudah dicetak di Log Eksekusi.\n\n' +
          'Langkah:\n' +
          '1. Buka panel "Log eksekusi" di bawah\n' +
          '2. Salin URL panjang yang muncul\n' +
          '3. Buka URL di tab browser baru\n' +
          '4. Login & klik "Allow" / "Izinkan"\n' +
          '5. Jalankan "inisialisasiMailApp" untuk verifikasi',
          ui.ButtonSet.OK
        );
      } catch (uiErr) {
        // Tidak ada UI (dipanggil dari trigger), cukup log saja
      }

    } else if (status === ScriptApp.AuthorizationStatus.NOT_REQUIRED) {
      Logger.log('STATUS : ✅ Otorisasi SUDAH lengkap');
      Logger.log('');
      Logger.log('Semua scope sudah disetujui. Menjalankan verifikasi MailApp...');
      inisialisasiMailApp();

    } else {
      Logger.log('STATUS : ⚠️ Status tidak dikenal: ' + status);
      Logger.log('Coba revoke akses dari: https://myaccount.google.com/permissions');
      Logger.log('Kemudian jalankan fungsi ini kembali.');
    }

  } catch (e) {
    Logger.log('ERROR di dapatkanUrlOtorisasi: ' + e.message);
    throw e;
  }
}

/**
 * STEP 2 — Jalankan setelah otorisasi berhasil diberikan.
 *
 * Verifikasi MailApp berfungsi dan kirim email tes ke alamat
 * yang sudah diatur di menu Pengaturan.
 *
 * Pastikan sudah:
 * 1. Menjalankan dapatkanUrlOtorisasi() dan membuka URL-nya
 * 2. Mengisi email tujuan di Pengaturan → Email Notifikasi Stok
 */
function inisialisasiMailApp() {
  try {
    // Cek status otorisasi terlebih dahulu
    var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    var status = authInfo.getAuthorizationStatus();

    if (status === ScriptApp.AuthorizationStatus.REQUIRED) {
      var urlBelum = authInfo.getAuthorizationUrl();
      Logger.log('❌ Otorisasi belum diberikan.');
      Logger.log('Jalankan dapatkanUrlOtorisasi() terlebih dahulu.');
      Logger.log('URL: ' + urlBelum);
      try {
        SpreadsheetApp.getUi().alert(
          '❌ Otorisasi belum diberikan!\n\n' +
          'Jalankan fungsi "dapatkanUrlOtorisasi" terlebih\n' +
          'dahulu, buka URL di Log Eksekusi, lalu klik Allow.'
        );
      } catch (uiErr) {}
      return;
    }

    // Otorisasi OK — cek kuota MailApp
    var kuota = MailApp.getRemainingDailyQuota();
    Logger.log('✅ Kuota email harian tersisa: ' + kuota);

    var emailTujuan = parseEmailList(getSetting('notif_stok_email_tujuan'));
    var namaToko = String(getSetting('nama_toko') || 'Toko');
    var waktu = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });

    if (emailTujuan.length === 0) {
      Logger.log('✅ Otorisasi MailApp BERHASIL! Kuota: ' + kuota);
      Logger.log('⚠️  Belum ada email tujuan di Pengaturan.');
      try {
        SpreadsheetApp.getUi().alert(
          '✅ Otorisasi MailApp BERHASIL!\n' +
          'Kuota email tersisa: ' + kuota + '\n\n' +
          '⚠️ Email tes tidak dikirim karena belum ada\n' +
          'email tujuan di menu Pengaturan.\n\n' +
          'Langkah selanjutnya:\n' +
          '1. Buka web app → Pengaturan\n' +
          '2. Aktifkan "Email Notifikasi"\n' +
          '3. Isi email tujuan\n' +
          '4. Simpan & klik "Kirim Tes Email"'
        );
      } catch (uiErr) {}
      return;
    }

    // Kirim email konfirmasi ke email tujuan
    var bodyText = [
      '✅ Otorisasi MailApp Berhasil!',
      '',
      'Toko   : ' + namaToko,
      'Waktu  : ' + waktu,
      'Kuota  : ' + kuota + ' email/hari tersisa',
      '',
      'Fitur notifikasi email stok rendah sekarang aktif.',
      'Notifikasi otomatis akan dikirim saat stok bahan baku',
      'berada di bawah batas minimum yang ditentukan.'
    ].join('\n');

    var bodyHtml = '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;">' +
      '<div style="background:#0D9488;padding:20px;border-radius:8px 8px 0 0;">' +
      '<h2 style="color:#fff;margin:0;">✅ Otorisasi Email Berhasil</h2>' +
      '</div>' +
      '<div style="background:#f8fafc;padding:24px;border:1px solid #e2e8f0;border-radius:0 0 8px 8px;">' +
      '<p style="color:#334155;">Otorisasi <strong>MailApp</strong> untuk Sistem Kasir POS berhasil.</p>' +
      '<table style="width:100%;border-collapse:collapse;margin:16px 0;">' +
      '<tr><td style="padding:8px;color:#64748b;width:70px;">Toko</td>' +
      '<td style="padding:8px;font-weight:600;color:#1e293b;">' + namaToko + '</td></tr>' +
      '<tr style="background:#f1f5f9;"><td style="padding:8px;color:#64748b;">Waktu</td>' +
      '<td style="padding:8px;font-weight:600;color:#1e293b;">' + waktu + '</td></tr>' +
      '<tr><td style="padding:8px;color:#64748b;">Kuota</td>' +
      '<td style="padding:8px;font-weight:600;color:#0D9488;">' + kuota + ' email/hari tersisa</td></tr>' +
      '</table>' +
      '<p style="color:#334155;margin:0;">Notifikasi stok rendah akan dikirim otomatis ke email ini ' +
      'ketika ada bahan baku di bawah batas minimum.</p>' +
      '</div></div>';

    MailApp.sendEmail({
      to: emailTujuan.join(','),
      subject: '[POS] ✅ Otorisasi Email Berhasil - ' + namaToko,
      body: bodyText,
      htmlBody: bodyHtml
    });

    Logger.log('✅ Email konfirmasi berhasil dikirim ke: ' + emailTujuan.join(', '));
    try {
      SpreadsheetApp.getUi().alert(
        '✅ Otorisasi MailApp BERHASIL!\n\n' +
        'Email konfirmasi dikirim ke:\n' +
        emailTujuan.join('\n') + '\n\n' +
        'Fitur notifikasi email stok rendah aktif.\n' +
        'Klik "Kirim Tes Email" di menu Pengaturan kapan saja.'
      );
    } catch (uiErr) {}

  } catch (e) {
    Logger.log('❌ inisialisasiMailApp error: ' + e.message);
    try {
      SpreadsheetApp.getUi().alert(
        '❌ Gagal\n\n' +
        'Detail: ' + e.message + '\n\n' +
        'Coba jalankan "dapatkanUrlOtorisasi" terlebih dahulu.'
      );
    } catch (uiErr) {}
    throw e;
  }
}
