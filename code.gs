/**********************************************************************
 * WEB APP PRESENSI & REKAPITULASI SISWA
 * Kelas XI AKL - SMK Negeri 1 Maluku Tengah
 * Tahun Pelajaran 2026/2027
 * Guru Pengampu: Alowisyus Bahy, S.Pd., M.Pd
 *
 * File: code.gs  (Backend API - Google Apps Script)
 * Versi 3 — daftar siswa dinamis (sheet "Siswa"), tambah & hapus siswa
 **********************************************************************/

// ============== KONFIGURASI ==============
const SPREADSHEET_ID  = '1_Y-WhaGPLhakGdeE7F8Hqt-QNXCNIMM75WGmjYTNxXs';
const SHEET_PRESENSI  = 'Presensi';
const SHEET_SISWA     = 'Siswa';        // sheet baru untuk master data siswa

const HEADERS_PRESENSI = [
  'Timestamp', 'Tanggal Presensi', 'Nama Siswa', 'NIS', 'Status', 'Semester', 'Pertemuan Ke-'
];
const HEADERS_SISWA = ['NIS', 'Nama'];  // header sheet Siswa

// Seed awal — hanya dipakai saat sheet Siswa belum ada / masih kosong
const SEED_SISWA = [
  ['5453','Ceska Silaya'],
  ['5454','Christian Mirlau'],
  ['5455','Cindy Grisel De Lima'],
  ['5456','Dhea Nel Lessil'],
  ['5457','Erli Ervyna Marlissa'],
  ['5458','Fidela Latuny'],
  ['5459','Julis P Kainama'],
  ['5460','Maria Defrosa Betaubun'],
  ['5461','Marisa D. G. Lasamahu'],
  ['5462','Melinda Wattimury'],
  ['5463','Melisa Wattimury'],
  ['5464','Nerges Wattimury'],
  ['5465','Nursita R Ibrahim'],
  ['5466','Patresia Agustina Tuasuun'],
  ['5467','Riska Amelia Putri Pattiasina'],
  ['5468','Safira Lefina Ngutra'],
  ['5469','Sarah Chezia Tuapetel'],
  ['5470','Sariati Ode'],
  ['5471','Valen Boby Wattimury'],
  ['5472','Wa Nurbaya'],
  ['5473','Wa Puput'],
  ['5474','Yulia Telussa'],
  ['5475','Yulisa Suriale']
];

// ============== SHEET HELPERS ==============

function getPresensiSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_PRESENSI);
  if (!sh) { sh = ss.insertSheet(SHEET_PRESENSI); }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS_PRESENSI);
    sh.getRange(1, 1, 1, HEADERS_PRESENSI.length)
      .setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Ambil / buat sheet Siswa.
 * Jika baru dibuat, isi dengan data seed supaya tidak kosong.
 */
function getSiswaSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_SISWA);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SISWA);
    sh.appendRow(HEADERS_SISWA);
    sh.getRange(1, 1, 1, HEADERS_SISWA.length)
      .setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    // Seed data awal
    if (SEED_SISWA.length > 0) {
      sh.getRange(2, 1, SEED_SISWA.length, 2).setValues(SEED_SISWA);
    }
  } else if (sh.getLastRow() <= 1) {
    // Ada sheet tapi kosong — isi seed
    if (SEED_SISWA.length > 0) {
      sh.getRange(2, 1, SEED_SISWA.length, 2).setValues(SEED_SISWA);
    }
  }
  return sh;
}

// ============== JSON RESPONSE ==============
function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============== ENTRY POINTS ==============

function doGet(e) {
  const action = e.parameter && e.parameter.action;
  if (!action) return jsonResponse_({ success: true, message: 'API Presensi XI AKL v3 aktif.' });

  if (action === 'getStudentList') return jsonResponse_({ success: true, data: getStudentList() });
  if (action === 'getSummaryData') return jsonResponse_(getSummaryData());

  return jsonResponse_({ success: false, message: 'Action tidak dikenali: ' + action });
}

function doPost(e) {
  try {
    const body    = JSON.parse(e.postData.contents || '{}');
    const action  = body.action;
    const payload = body.payload || {};

    let result;
    switch (action) {
      case 'getStudentList':       result = { success: true, data: getStudentList() }; break;
      case 'submitAttendance':     result = submitBulkAttendance([payload]);           break;
      case 'submitBulkAttendance': result = submitBulkAttendance(payload);             break;
      case 'getSummaryData':       result = getSummaryData();                          break;
      case 'addStudent':           result = addStudent(payload);                       break;
      case 'deleteStudent':        result = deleteStudent(payload);                    break;
      default: result = { success: false, message: 'Action tidak dikenali: ' + action };
    }
    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({ success: false, message: 'Error: ' + err.message });
  }
}

// ============== STUDENT CRUD ==============

/**
 * Baca daftar siswa dari sheet Siswa, urutkan berdasarkan nama.
 */
function getStudentList() {
  const sh      = getSiswaSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const list = [];
  values.forEach(function (row) {
    const nis  = String(row[0]).trim();
    const nama = String(row[1]).trim();
    if (nis && nama) list.push({ nis: nis, nama: nama });
  });
  list.sort(function (a, b) { return a.nama.localeCompare(b.nama, 'id'); });
  return list;
}

/**
 * Tambah satu siswa baru.
 * payload: { nis, nama }
 */
function addStudent(payload) {
  try {
    const nis  = String(payload.nis  || '').trim();
    const nama = String(payload.nama || '').trim();

    if (!nis)  throw new Error('NIS tidak boleh kosong.');
    if (!nama) throw new Error('Nama tidak boleh kosong.');

    const sh      = getSiswaSheet_();
    const lastRow = sh.getLastRow();

    // Cek duplikat NIS
    if (lastRow >= 2) {
      const existing = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (String(existing[i][0]).trim() === nis) {
          throw new Error('NIS ' + nis + ' sudah terdaftar.');
        }
      }
    }

    sh.appendRow([nis, nama]);
    return { success: true, message: nama + ' (NIS: ' + nis + ') berhasil ditambahkan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Hapus siswa berdasarkan NIS.
 * payload: { nis }
 */
function deleteStudent(payload) {
  try {
    const nis = String(payload.nis || '').trim();
    if (!nis) throw new Error('NIS tidak boleh kosong.');

    const sh      = getSiswaSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('Daftar siswa kosong.');

    const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    let deletedRow = -1;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === nis) {
        deletedRow = i + 2; // +2 karena baris 1 adalah header
        break;
      }
    }

    if (deletedRow === -1) throw new Error('Siswa dengan NIS ' + nis + ' tidak ditemukan.');

    sh.deleteRow(deletedRow);
    return { success: true, message: 'Siswa dengan NIS ' + nis + ' berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============== PRESENSI ==============

function submitBulkAttendance(records) {
  try {
    if (!Array.isArray(records) || records.length === 0) {
      throw new Error('Data presensi kosong atau format tidak valid.');
    }
    const sh        = getPresensiSheet_();
    const timestamp = new Date();
    const rows      = [];
    const errors    = [];

    records.forEach(function (data, idx) {
      if (!data || !data.tanggal || !data.nis || !data.nama || !data.status || !data.semester) {
        errors.push('Baris #' + (idx + 1) + ': data tidak lengkap.');
        return;
      }
      rows.push([timestamp, data.tanggal, data.nama, data.nis, data.status, data.semester, data.pertemuan || '']);
    });

    rows.forEach(function (row) { sh.appendRow(row); });

    return {
      success: true, saved: rows.length, failed: errors.length,
      message: rows.length + ' presensi berhasil disimpan.'
        + (errors.length > 0 ? ' ' + errors.length + ' gagal.' : '')
    };
  } catch (err) {
    return { success: false, saved: 0, failed: 0, message: 'Gagal: ' + err.message };
  }
}

// ============== REKAP ==============

function getSummaryData() {
  try {
    const sh      = getPresensiSheet_();
    const lastRow = sh.getLastRow();

    if (lastRow < 2) {
      return { success: true, monthly: {}, semester: { Ganjil: empty_(), Genap: empty_() }, raw: [] };
    }

    const values    = sh.getRange(2, 1, lastRow - 1, HEADERS_PRESENSI.length).getValues();
    const monthly   = {};
    const semester  = { Ganjil: empty_(), Genap: empty_() };
    const monthNames = ['Januari','Februari','Maret','April','Mei','Juni',
                        'Juli','Agustus','September','Oktober','November','Desember'];
    const raw = [];

    values.forEach(function (row) {
      const tanggalRaw = row[1], nama = row[2], nis = row[3],
            status = row[4], semVal = row[5], pertemuan = row[6];
      if (!tanggalRaw || !status) return;

      const d = tanggalRaw instanceof Date ? tanggalRaw : new Date(tanggalRaw);
      if (isNaN(d.getTime())) return;

      const y = d.getFullYear(), m = d.getMonth();
      const key = y + '-' + String(m + 1).padStart(2, '0');
      if (!monthly[key]) monthly[key] = { label: monthNames[m] + ' ' + y, Hadir:0, Sakit:0, Izin:0, Alpa:0 };
      if (monthly[key].hasOwnProperty(status)) monthly[key][status]++;

      if ((semVal === 'Ganjil' || semVal === 'Genap') && semester[semVal].hasOwnProperty(status)) {
        semester[semVal][status]++;
      }

      raw.push({
        tanggal: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        nama: nama, nis: nis, status: status, semester: semVal, pertemuan: pertemuan
      });
    });

    const sorted = {};
    Object.keys(monthly).sort().forEach(function (k) { sorted[k] = monthly[k]; });

    return { success: true, monthly: sorted, semester: semester, raw: raw.reverse() };
  } catch (err) {
    return { success: false, message: 'Gagal rekap: ' + err.message };
  }
}

function empty_() { return { Hadir: 0, Sakit: 0, Izin: 0, Alpa: 0 }; }
