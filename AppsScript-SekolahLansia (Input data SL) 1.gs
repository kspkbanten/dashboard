/**
 * Apps Script — Input Sekolah Lansia (Dashboard KSPK Banten)
 * ---------------------------------------------------------
 * 1. Tempel kode ini di Apps Script spreadsheet
 * 2. Deploy Web app (Execute as Me, Anyone)
 * 3. Authorize Spreadsheet + Google Drive
 * 4. URL /exec → DATA_CONFIG.SEKOLAH_LANSIA_WEBAPP_URL
 *
 * Sheet: SEKOLAH_LANSIA_INPUT
 * Folder SK: SekolahLansia_SK
 */

var SPREADSHEET_ID = '16790nM8tMLqyc9ltO3clW6tbkS7vUyjIL0OpPGUIpKY';
var SHEET_NAME = 'SEKOLAH_LANSIA_INPUT';
var FOLDER_NAME = 'SekolahLansia_SK';

var HEADERS = [
  'Timestamp',
  'Kabupaten/Kota',
  'Kecamatan',
  'Desa/Kelurahan',
  'Nama Sekolah Lansia',
  'Nomor SK',
  'Link SK',
  'Tahun Pembentukan',
  'Jumlah Siswa Saat Ini',
  'Jumlah Siswa Diwisuda',
  'Jumlah Pengurus',
  'Jenjang Kelas',
  'Status Sekolah'
];

function getSs_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/** Pastikan sheet ada dan baris header selalu sesuai HEADERS (termasuk kolom baru). */
function getSheet_() {
  var ss = getSs_();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }
  // Sinkronkan header: tulis ulang baris 1 agar kolom baru (Pengurus, Jenjang) selalu ada
  var lastCol = Math.max(sh.getLastColumn(), HEADERS.length);
  var current = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var needUpdate = current.length < HEADERS.length;
  if (!needUpdate) {
    for (var i = 0; i < HEADERS.length; i++) {
      if (String(current[i] || '').trim() !== HEADERS[i]) {
        needUpdate = true;
        break;
      }
    }
  }
  if (needUpdate) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveSkFile_(skFile) {
  if (!skFile || !skFile.data) {
    return '';
  }
  try {
    var folder = null;
    try {
      var ssFile = DriveApp.getFileById(SPREADSHEET_ID);
      var parents = ssFile.getParents();
      var parent = parents.hasNext() ? parents.next() : null;
      if (parent) {
        var it = parent.getFoldersByName(FOLDER_NAME);
        folder = it.hasNext() ? it.next() : parent.createFolder(FOLDER_NAME);
      }
    } catch (e1) {
      folder = null;
    }
    if (!folder) {
      try {
        var folders = DriveApp.getFoldersByName(FOLDER_NAME);
        folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
      } catch (e2) {
        return 'UPLOAD_SK_GAGAL: tidak ada izin Drive. Authorize ulang deployment dengan akses Google Drive. Detail: ' +
          String(e2 && e2.message ? e2.message : e2);
      }
    }
    var bytes = Utilities.base64Decode(skFile.data);
    var blob = Utilities.newBlob(
      bytes,
      skFile.mimeType || 'application/pdf',
      skFile.name || ('SK_SEKOLAH_LANSIA_' + new Date().getTime() + '.pdf')
    );
    var file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eShare) {}
    return file.getUrl();
  } catch (err) {
    return 'UPLOAD_SK_GAGAL: ' + String(err && err.message ? err.message : err);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      service: 'SekolahLansiaInput',
      time: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var data = JSON.parse(raw);
    if (data.action && data.action !== 'submitSekolahLansia') {
      return json_({ ok: false, message: 'Action tidak dikenal' });
    }
    var required = [
      'kabupaten', 'kecamatan', 'desa', 'namaSekolah', 'nomorSk',
      'tahunPembentukan', 'jumlahSiswa', 'jumlahWisuda',
      'jumlahPengurus', 'jenjangKelas', 'statusSekolah'
    ];
    for (var i = 0; i < required.length; i++) {
      var k = required[i];
      if (data[k] === undefined || data[k] === null || String(data[k]).trim() === '') {
        return json_({ ok: false, message: 'Kolom wajib kosong: ' + k });
      }
    }

    var skUrl = '';
    if (data.skFile && data.skFile.data) {
      skUrl = saveSkFile_(data.skFile);
    }

    // Urutan kolom HARUS sama dengan HEADERS
    var row = [
      new Date(),
      String(data.kabupaten).trim(),
      String(data.kecamatan).trim(),
      String(data.desa).trim(),
      String(data.namaSekolah).trim(),
      String(data.nomorSk).trim(),
      skUrl,
      String(data.tahunPembentukan).trim(),
      Number(data.jumlahSiswa) || 0,
      Number(data.jumlahWisuda) || 0,
      Number(data.jumlahPengurus) || 0,
      String(data.jenjangKelas).trim(),
      String(data.statusSekolah).trim()
    ];
    getSheet_().appendRow(row);

    var msg = 'Data Sekolah Lansia berhasil disimpan.';
    if (skUrl && skUrl.indexOf('UPLOAD_SK_GAGAL') === 0) {
      msg += ' Lampiran SK belum terunggah. Authorize ulang Web App dengan izin Google Drive, lalu kirim ulang jika perlu.';
    }
    return json_({ ok: true, message: msg, skUrl: skUrl });
  } catch (err) {
    return json_({ ok: false, message: String(err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
