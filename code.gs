/**********************************************************************
 * WEB APP PRESENSI & REKAPITULASI SISWA
 * Kelas XI AKL - SMK Negeri 1 Maluku Tengah
 * Tahun Pelajaran 2026/2027
 * Guru Pengampu: Alowisyus Bahy, S.Pd., M.Pd
 *
 * File: code.gs (Backend API - Google Apps Script)
 *
 * PERUBAHAN v2:
 * - submitAttendance (satu siswa) -> submitBulkAttendance (seluruh kelas sekaligus)
 * - submitAttendance lama tetap dipertahankan untuk kompatibilitas
 **********************************************************************/

// ============== KONFIGURASI ==============
const SPREADSHEET_ID = '1_Y-WhaGPLhakGdeE7F8Hqt-QNXCNIMM75WGmjYTNxXs';
const SHEET_NAME = 'Presensi';

const HEADERS = [
  'Timestamp',
  'Tanggal Presensi',
  'Nama Siswa',
  'NIS',
  'Status',
  'Semester',
  'Pertemuan Ke-'
];

/**
 * Daftar 23 siswa Kelas XI AKL - SMK Negeri 1 Maluku Tengah
 * Tahun Pelajaran 2026/2027 (data resmi).
 */
const DAFTAR_SISWA = [
  { nis: '5453', nama: 'Ceska Silaya' },
  { nis: '5454', nama: 'Christian Mirlau' },
  { nis: '5455', nama: 'Cindy Grisel De Lima' },
  { nis: '5456', nama: 'Dhea Nel Lessil' },
  { nis: '5457', nama: 'Erli Ervyna Marlissa' },
  { nis: '5458', nama: 'Fidela Latuny' },
  { nis: '5459', nama: 'Julis P Kainama' },
  { nis: '5460', nama: 'Maria Defrosa Betaubun' },
  { nis: '5461', nama: 'Marisa D. G. Lasamahu' },
  { nis: '5462', nama: 'Melinda Wattimury' },
  { nis: '5463', nama: 'Melisa Wattimury' },
  { nis: '5464', nama: 'Nerges Wattimury' },
  { nis: '5465', nama: 'Nursita R Ibrahim' },
  { nis: '5466', nama: 'Patresia Agustina Tuasuun' },
  { nis: '5467', nama: 'Riska Amelia Putri Pattiasina' },
  { nis: '5468', nama: 'Safira Lefina Ngutra' },
  { nis: '5469', nama: 'Sarah Chezia Tuapetel' },
  { nis: '5470', nama: 'Sariati Ode' },
  { nis: '5471', nama: 'Valen Boby Wattimury' },
  { nis: '5472', nama: 'Wa Nurbaya' },
  { nis: '5473', nama: 'Wa Puput' },
  { nis: '5474', nama: 'Yulia Telussa' },
  { nis: '5475', nama: 'Yulisa Suriale' }
];

// ============== ENTRY POINT API ==============

function doGet(e) {
  const action = e.parameter && e.parameter.action;

  if (!action) {
    return jsonResponse_({
      success: true,
      message: 'API Presensi XI AKL aktif v2. Gunakan POST dengan field "action".'
    });
  }

  if (action === 'getStudentList') {
    return jsonResponse_({ success: true, data: getStudentList() });
  }
  if (action === 'getSummaryData') {
    return jsonResponse_(getSummaryData());
  }

  return jsonResponse_({ success: false, message: 'Action tidak dikenali untuk GET: ' + action });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    const payload = body.payload || {};

    let result;
    switch (action) {
      case 'getStudentList':
        result = { success: true, data: getStudentList() };
        break;
      case 'submitAttendance':
        // Kompatibilitas lama — bungkus jadi bulk 1 item
        result = submitBulkAttendance([payload]);
        break;
      case 'submitBulkAttendance':
        // NEW: terima array records
        result = submitBulkAttendance(payload);
        break;
      case 'getSummaryData':
        result = getSummaryData();
        break;
      default:
        result = { success: false, message: 'Action tidak dikenali: ' + action };
    }

    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({ success: false, message: 'Gagal memproses request: ' + err.message });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#1a237e')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getStudentList() {
  return DAFTAR_SISWA;
}

/**
 * [LAMA] Simpan satu record presensi.
 * Masih bisa dipanggil internal, tapi dari frontend sudah pakai bulk.
 */
function submitAttendance(data) {
  return submitBulkAttendance([data]);
}

/**
 * [BARU] Simpan banyak record presensi sekaligus (seluruh kelas).
 * @param {Array} records - Array of { tanggal, nama, nis, status, semester, pertemuan }
 * @return {Object} { success, saved, failed, message }
 */
function submitBulkAttendance(records) {
  try {
    if (!Array.isArray(records) || records.length === 0) {
      throw new Error('Data presensi kosong atau format tidak valid.');
    }

    const sheet = getSheet_();
    const timestamp = new Date();
    const rows = [];
    const errors = [];

    records.forEach(function (data, idx) {
      if (!data || !data.tanggal || !data.nis || !data.nama || !data.status || !data.semester) {
        errors.push('Baris #' + (idx + 1) + ' (' + (data && data.nama || 'Unknown') + '): data tidak lengkap.');
        return;
      }
      rows.push([
        timestamp,
        data.tanggal,
        data.nama,
        data.nis,
        data.status,
        data.semester,
        data.pertemuan || ''
      ]);
    });

    if (rows.length > 0) {
      // appendRow batch — lebih efisien
      rows.forEach(function (row) {
        sheet.appendRow(row);
      });
    }

    return {
      success: true,
      saved: rows.length,
      failed: errors.length,
      message: rows.length + ' data presensi berhasil disimpan.' +
        (errors.length > 0 ? ' ' + errors.length + ' gagal: ' + errors.join('; ') : '')
    };
  } catch (err) {
    return {
      success: false,
      saved: 0,
      failed: records ? records.length : 0,
      message: 'Gagal menyimpan presensi: ' + err.message
    };
  }
}

/**
 * Rekap bulanan & semester dari seluruh data sheet.
 */
function getSummaryData() {
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        success: true,
        monthly: {},
        semester: { Ganjil: emptyStatusCount_(), Genap: emptyStatusCount_() },
        raw: []
      };
    }

    const range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
    const values = range.getValues();

    const monthly = {};
    const semester = {
      Ganjil: emptyStatusCount_(),
      Genap: emptyStatusCount_()
    };

    const monthNames = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];

    const raw = [];

    values.forEach(function (row) {
      const tanggalRaw = row[1];
      const nama = row[2];
      const nis = row[3];
      const status = row[4];
      const semesterVal = row[5];
      const pertemuan = row[6];

      if (!tanggalRaw || !status) return;

      let tanggalObj;
      if (tanggalRaw instanceof Date) {
        tanggalObj = tanggalRaw;
      } else {
        tanggalObj = new Date(tanggalRaw);
      }
      if (isNaN(tanggalObj.getTime())) return;

      const year = tanggalObj.getFullYear();
      const monthIdx = tanggalObj.getMonth();
      const monthKey = year + '-' + String(monthIdx + 1).padStart(2, '0');
      const monthLabel = monthNames[monthIdx] + ' ' + year;

      if (!monthly[monthKey]) {
        monthly[monthKey] = { label: monthLabel, Hadir: 0, Sakit: 0, Izin: 0, Alpa: 0 };
      }
      if (monthly[monthKey].hasOwnProperty(status)) {
        monthly[monthKey][status]++;
      }

      if (semesterVal === 'Ganjil' || semesterVal === 'Genap') {
        if (semester[semesterVal].hasOwnProperty(status)) {
          semester[semesterVal][status]++;
        }
      }

      raw.push({
        tanggal: Utilities.formatDate(tanggalObj, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        nama: nama,
        nis: nis,
        status: status,
        semester: semesterVal,
        pertemuan: pertemuan
      });
    });

    const sortedMonthlyKeys = Object.keys(monthly).sort();
    const sortedMonthly = {};
    sortedMonthlyKeys.forEach(function (k) { sortedMonthly[k] = monthly[k]; });

    return {
      success: true,
      monthly: sortedMonthly,
      semester: semester,
      raw: raw.reverse()
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal mengambil data rekapitulasi: ' + err.message
    };
  }
}

function emptyStatusCount_() {
  return { Hadir: 0, Sakit: 0, Izin: 0, Alpa: 0 };
}
