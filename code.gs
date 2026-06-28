/**********************************************************************
 * WEB APP PRESENSI & REKAPITULASI SISWA
 * Kelas XI AKL - SMK Negeri 1 Maluku Tengah
 * Tahun Pelajaran 2026/2027
 * Guru Pengampu: Alowisyus Bahy, S.Pd., M.Pd
 *
 * File: code.gs (Backend API - Google Apps Script)
 *
 * PERUBAHAN PENTING:
 * File ini SEKARANG HANYA BERFUNGSI SEBAGAI API (headless).
 * UI (index.html) tidak lagi disajikan dari sini — UI di-hosting
 * terpisah di GitHub Pages, dan memanggil API ini lewat fetch().
 *
 * Cara kerja:
 *  - Frontend (GitHub Pages) mengirim request POST berisi JSON
 *    dengan field "action" untuk menentukan operasi apa yang
 *    dijalankan (getStudentList / submitAttendance / getSummaryData).
 *  - doPost() membaca action tersebut, memanggil fungsi yang sesuai,
 *    lalu mengembalikan hasil sebagai JSON dengan header CORS supaya
 *    bisa dibaca dari domain GitHub Pages.
 **********************************************************************/

// ============== KONFIGURASI ==============
const SPREADSHEET_ID = '1_Y-WhaGPLhakGdeE7F8Hqt-QNXCNIMM75WGmjYTNxXs';
const SHEET_NAME = 'Presensi'; // Nama sheet tempat data presensi disimpan

// Header kolom pada sheet (urutan harus konsisten)
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

// ============== ENTRY POINT API (HEADLESS) ==============

/**
 * doGet() - dipakai untuk cek API hidup (health check) DAN untuk
 * menerima action via query string (?action=...), supaya frontend
 * yang tidak mengirim POST (misalnya saat debugging via browser)
 * tetap bisa mendapat respon JSON yang masuk akal.
 */
function doGet(e) {
  const action = e.parameter && e.parameter.action;

  if (!action) {
    return jsonResponse_({
      success: true,
      message: 'API Presensi XI AKL aktif. Gunakan POST dengan field "action".'
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

/**
 * doPost() - entry point utama yang dipanggil dari frontend GitHub
 * Pages lewat fetch(). Body request berupa JSON:
 *   { "action": "submitAttendance", "payload": { ... } }
 */
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
        result = submitAttendance(payload);
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

/**
 * Helper: membungkus data jadi JSON response.
 * Catatan: Apps Script Web App secara native MENGIZINKAN akses
 * cross-origin (CORS) untuk response JSON biasa ContentService —
 * tidak perlu header Access-Control-Allow-Origin manual, karena
 * Apps Script tidak mengizinkan set header response secara manual.
 * Yang membuat ini bekerja lintas-domain adalah deployment dengan
 * akses "Anyone" (lihat petunjuk deploy).
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper: Mengambil objek Sheet, membuat sheet & header jika belum ada.
 */
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

/**
 * Mengirim daftar siswa ke frontend (untuk mengisi dropdown).
 */
function getStudentList() {
  return DAFTAR_SISWA;
}

/**
 * Menyimpan data presensi yang dikirim dari form.
 * @param {Object} data - { tanggal, nama, nis, status, semester, pertemuan }
 * @return {Object} status hasil penyimpanan
 */
function submitAttendance(data) {
  try {
    if (!data || !data.tanggal || !data.nis || !data.nama || !data.status || !data.semester) {
      throw new Error('Data tidak lengkap. Pastikan semua field telah diisi.');
    }

    const sheet = getSheet_();
    const timestamp = new Date();

    sheet.appendRow([
      timestamp,
      data.tanggal,
      data.nama,
      data.nis,
      data.status,
      data.semester,
      data.pertemuan || ''
    ]);

    return {
      success: true,
      message: 'Presensi untuk ' + data.nama + ' berhasil disimpan.'
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal menyimpan presensi: ' + err.message
    };
  }
}

/**
 * Mengambil seluruh data presensi dari sheet dan mengolahnya
 * menjadi rekapan bulanan dan rekapan per semester.
 * @return {Object} { monthly: {...}, semester: {...}, raw: [...] }
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

    const monthly = {}; // { "2026-07": { Hadir: n, Sakit: n, Izin: n, Alpa: n } }
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
      const monthIdx = tanggalObj.getMonth(); // 0-11
      const monthKey = year + '-' + String(monthIdx + 1).padStart(2, '0');
      const monthLabel = monthNames[monthIdx] + ' ' + year;

      if (!monthly[monthKey]) {
        monthly[monthKey] = {
          label: monthLabel,
          Hadir: 0,
          Sakit: 0,
          Izin: 0,
          Alpa: 0
        };
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

    // Urutkan monthly berdasarkan key (kronologis)
    const sortedMonthlyKeys = Object.keys(monthly).sort();
    const sortedMonthly = {};
    sortedMonthlyKeys.forEach(function (k) {
      sortedMonthly[k] = monthly[k];
    });

    return {
      success: true,
      monthly: sortedMonthly,
      semester: semester,
      raw: raw.reverse() // data terbaru di atas
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal mengambil data rekapitulasi: ' + err.message
    };
  }
}

/**
 * Helper: objek kosong untuk hitungan status kehadiran.
 */
function emptyStatusCount_() {
  return { Hadir: 0, Sakit: 0, Izin: 0, Alpa: 0 };
}
