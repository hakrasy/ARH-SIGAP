/**
 * ====================================================================
 * SISTEM INFORMASI AKADEMIK WALI KELAS & GURU BAHASA INGGRIS (ARH)
 * Backend Logic - Google Apps Script
 * ====================================================================
 */

const SPREADSHEET_ID = '1-ec2Psh6Kn782scW62b35wYfaiV2QjRlF6cT3hAGPx8';
const DRIVE_FOLDER_ID = '1yzCVVs7LpR1XfqBxuWBsbEvheX2w1grF';

/**
 * Mengembalikan file HTML utama untuk rendering web app.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Indeks')
    .evaluate()
    .setTitle('Sistem Informasi ARH')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    return sheet;
  } catch (err) {
    Logger.log("Error mendapatkan sheet " + sheetName + ": " + err.message);
    return null;
  }
}

function getTimestamp() {
  return Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
}

function initSheetHeaders(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return;
  
  if (sheet.getLastRow() === 0) {
    const headers = {
      'Siswa_WK': ['NIS', 'Nama', 'Jenis Kelamin', 'Agama', 'Tempat Lahir', 'Tanggal Lahir', 'Alamat', 'NISN', 'NIK', 'Nama Ayah', 'Nama Ibu', 'Pekerjaan Ayah', 'Pekerjaan Ibu', 'No. Handphone', 'Timestamp'],
      'Kehadiran_WK': ['Tanggal', 'Nama', 'Status', 'Keterangan', 'Timestamp'],
      'Jurnal_WK': ['Tanggal', 'Nama', 'Catatan', 'Dokumentasi', 'Timestamp'],
      'Siswa_GB': ['NIS', 'Nama', 'Kelas', 'Timestamp'],
      'Kehadiran_GB': ['Tanggal', 'Kelas', 'Nama', 'Status', 'Catatan', 'Dokumentasi', 'Timestamp'],
      'TujuanPembelajaran': ['Semester', 'Kelas', 'Kode TP', 'Tujuan Pembelajaran', 'Elemen', 'Timestamp'],
      'Jadwal': ['Hari', 'Kelas', 'Jam Pelajaran', 'Timestamp'],
      'Nilai_Harian7': ['Semester', 'Kelas', 'NIS', 'Nama', 'Rata-rata', 'Timestamp'],
      'Nilai_Harian8': ['Semester', 'Kelas', 'NIS', 'Nama', 'Rata-rata', 'Timestamp'],
      'Nilai_Harian9': ['Semester', 'Kelas', 'NIS', 'Nama', 'Rata-rata', 'Timestamp'],
      'Jurnal_GB': ['Hari/Tanggal', 'Kelas', 'Jam Ke', 'Materi/Topik', 'Tujuan Pembelajaran', 'Kegiatan Pembelajaran', 'Kehadiran', 'Penilaian', 'Refleksi/Tindak Lanjut', 'Timestamp'],
      'ASA': ['Kelas', 'Nama', 'Ganjil', 'Genap', 'Timestamp']
    };
    if (headers[sheetName]) {
      sheet.appendRow(headers[sheetName]);
    }
  }
}

function uploadImageToDrive(base64Str, fileName) {
  try {
    if (!base64Str) return "";
    let folder;
    try {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    } catch(err) {
      folder = DriveApp.getRootFolder();
    }
    
    const parts = base64Str.match(/^data:(.+);base64,(.*)$/);
    if (parts) {
      const mimeType = parts[1];
      const base64Data = parts[2];
      const decoded = Utilities.base64Decode(base64Data);
      const blob = Utilities.newBlob(decoded, mimeType, fileName);
      const file = folder.createFile(blob);
      
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (sharingError) {
        Logger.log("Gagal mengatur akses file sharing.");
      }
      return file.getUrl();
    }
    return "";
  } catch (e) {
    return "ERROR: " + e.message;
  }
}

function readData(sheetName) {
  try {
    initSheetHeaders(sheetName); 
    const sheet = getSheet(sheetName);
    if (!sheet) throw new Error("Sheet " + sheetName + " tidak ditemukan.");
    
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return []; 
    
    const headers = data[0];
    const rows = data.slice(1);
    return rows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] ? row[index].trim() : "";
      });
      return obj;
    });
  } catch (e) {
    Logger.log("Gagal membaca data dari " + sheetName + ": " + e.message);
    return [];
  }
}

function getDashboardWKData() {
  try {
    const siswaData = readData('Siswa_WK');
    const kehadiranData = readData('Kehadiran_WK');
    const today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    
    const hadirToday = kehadiranData.filter(k => k['Tanggal'] === today && k['Status'] === 'H').length;
    const totalToday = kehadiranData.filter(k => k['Tanggal'] === today).length;
    const persentase = totalToday > 0 ? ((hadirToday / totalToday) * 100).toFixed(1) : 0;

    return {
      status: 'success',
      data: {
        totalSiswa: siswaData.length,
        persentaseKehadiran: persentase,
        kehadiranHistori: kehadiranData 
      }
    };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getSiswaWK() {
  return { status: 'success', data: readData('Siswa_WK') };
}

function addSiswaWK(payload) {
  try {
    const sheet = getSheet('Siswa_WK');
    if (!sheet) throw new Error("Gagal mengakses tabel Siswa_WK.");
    
    const row = [
      payload.nis, payload.nama, payload.jenisKelamin, payload.agama,
      payload.tempatLahir, payload.tanggalLahir, payload.alamat,
      payload.nisn, payload.nik, payload.namaAyah, payload.namaIbu,
      payload.pekerjaanAyah, payload.pekerjaanIbu, payload.noHandphone,
      getTimestamp()
    ];
    sheet.appendRow(row);
    return { status: 'success', message: 'Data Siswa berhasil ditambahkan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getKehadiranWK() {
  return { status: 'success', data: readData('Kehadiran_WK') };
}

function addKehadiranWK(payloadArray) {
  try {
    const sheet = getSheet('Kehadiran_WK');
    if (!sheet) throw new Error("Gagal mengakses tabel Kehadiran_WK.");
    const ts = getTimestamp();
    const rows = payloadArray.map(p => [p.tanggal, p.nama, p.status, p.keterangan, ts]);
    
    if(rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { status: 'success', message: 'Data Kehadiran berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getJurnalWK() {
  return { status: 'success', data: readData('Jurnal_WK') };
}

function addJurnalWK(payload) {
  try {
    initSheetHeaders('Jurnal_WK');
    const sheet = getSheet('Jurnal_WK');
    if (!sheet) throw new Error("Gagal mengakses tabel Jurnal_WK.");
    
    let fileUrl = "";
    if (payload.fileBase64 && payload.fileBase64 !== "") {
      fileUrl = uploadImageToDrive(payload.fileBase64, payload.fileName);
    }
    const row = [
      payload.tanggal, payload.nama, payload.catatan, fileUrl, getTimestamp()
    ];
    sheet.appendRow(row);
    return { status: 'success', message: 'Jurnal Siswa berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getDashboardGBData() {
  try {
    const siswaData = readData('Siswa_GB');
    const jadwalData = readData('Jadwal');
    
    const kelasStats = {};
    siswaData.forEach(s => {
      kelasStats[s['Kelas']] = (kelasStats[s['Kelas']] || 0) + 1;
    });

    const hariIni = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"][new Date().getDay()];
    const jadwalHariIni = jadwalData.filter(j => j['Hari'] === hariIni);

    return {
      status: 'success',
      data: {
        kelasStats: kelasStats,
        jadwalHariIni: jadwalHariIni
      }
    };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getSiswaGB() {
  return { status: 'success', data: readData('Siswa_GB') };
}

function addSiswaGB(payload) {
  try {
    const sheet = getSheet('Siswa_GB');
    if (!sheet) throw new Error("Gagal mengakses tabel Siswa_GB.");
    sheet.appendRow([payload.nis, payload.nama, payload.kelas, getTimestamp()]);
    return { status: 'success', message: 'Data Siswa Kelas berhasil ditambahkan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getKehadiranGB() {
  return { status: 'success', data: readData('Kehadiran_GB') };
}

function addKehadiranGB(payloadArray) {
  try {
    const sheet = getSheet('Kehadiran_GB');
    if (!sheet) throw new Error("Gagal mengakses tabel Kehadiran_GB.");
    const ts = getTimestamp();
    const rows = payloadArray.map(p => [p.tanggal, p.kelas, p.nama, p.status, p.catatan, p.dokumentasi || "", ts]);
    
    if(rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { status: 'success', message: 'Data Kehadiran Kelas berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getTujuanPembelajaran() {
  return { status: 'success', data: readData('TujuanPembelajaran') };
}

function addTujuanPembelajaran(payload) {
  try {
    initSheetHeaders('TujuanPembelajaran');
    const sheet = getSheet('TujuanPembelajaran');
    if (!sheet) throw new Error("Gagal mengakses tabel TujuanPembelajaran.");
    sheet.appendRow([
      payload.semester, 
      payload.kelas, 
      payload.kodeTp, 
      payload.tujuanPembelajaran, 
      payload.elemen, 
      getTimestamp()
    ]);
    return { status: 'success', message: 'Tujuan Pembelajaran berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getJadwal() {
  return { status: 'success', data: readData('Jadwal') };
}

function addJadwal(payload) {
  try {
    const sheet = getSheet('Jadwal');
    if (!sheet) throw new Error("Gagal mengakses tabel Jadwal.");
    sheet.appendRow([payload.hari, payload.kelas, payload.jamPelajaran, getTimestamp()]);
    return { status: 'success', message: 'Jadwal berhasil ditambahkan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * Helper untuk menentukan nama sheet nilai harian berdasarkan nama kelas
 */
function getTargetNilaiSheetName(kelas) {
  const kStr = String(kelas).trim();
  if (kStr.startsWith('7')) return 'Nilai_Harian7';
  if (kStr.startsWith('8')) return 'Nilai_Harian8';
  if (kStr.startsWith('9')) return 'Nilai_Harian9';
  return 'Nilai_Harian7'; // Default fallback
}

/**
 * Membaca data struktur dan isi nilai harian per kelas dan semester
 */
function getNilaiHarianData(payload) {
  try {
    const kelas = payload.kelas;
    const semester = payload.semester;
    const sheetName = getTargetNilaiSheetName(kelas);
    
    initSheetHeaders(sheetName);
    const sheetData = readData(sheetName);
    
    // Ambil daftar TP yang sesuai dengan Semester & Kelas
    const tpData = readData('TujuanPembelajaran');
    const matchedTPs = tpData
      .filter(t => t.Semester.toLowerCase() === semester.toLowerCase() && String(t.Kelas).trim() === String(kelas).trim())
      .map(t => String(t['Kode TP'] || t['KodeTp']).trim())
      .filter((v, i, a) => a.indexOf(v) === i)
      .sort((a, b) => a.localeCompare(b, undefined, {numeric: true}));

    // Ambil daftar siswa untuk kelas tersebut
    const siswaData = readData('Siswa_GB').filter(s => String(s.Kelas).trim() === String(kelas).trim());

    // Filter nilai harian yang tersimpan sesuai semester dan kelas
    const filteredGrades = sheetData.filter(d => 
      String(d.Semester).trim().toLowerCase() === String(semester).trim().toLowerCase() && 
      String(d.Kelas).trim() === String(kelas).trim()
    );

    return {
      status: 'success',
      sheetName: sheetName,
      tps: matchedTPs,
      siswa: siswaData,
      savedGrades: filteredGrades
    };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * Menyimpan seluruh data nilai harian secara berulang/batch untuk satu kelas & semester
 */
function saveNilaiHarianBatch(payload) {
  try {
    const { semester, kelas, items } = payload; // items: array of { nis, nama, tpScores: { [kodeTp]: score } }
    if (!kelas || !semester || !items || !Array.isArray(items)) {
      throw new Error("Data input nilai harian tidak lengkap.");
    }

    const sheetName = getTargetNilaiSheetName(kelas);
    const sheet = getSheet(sheetName);
    if (!sheet) throw new Error("Sheet " + sheetName + " tidak dapat diakses.");

    // Ambil daftar TP terdaftar
    const tpData = readData('TujuanPembelajaran');
    const matchedTPList = tpData
      .filter(t => t.Semester.toLowerCase() === semester.toLowerCase() && String(t.Kelas).trim() === String(kelas).trim())
      .map(t => String(t['Kode TP'] || t['KodeTp']).trim())
      .filter((v, i, a) => a.indexOf(v) === i)
      .sort((a, b) => a.localeCompare(b, undefined, {numeric: true}));

    // Tentukan header baku
    const headers = ['Semester', 'Kelas', 'NIS', 'Nama', ...matchedTPList, 'Rata-rata', 'Timestamp'];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    } else {
      // Periksa header existing
      const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
      if (existingHeaders.join(',') !== headers.join(',')) {
        sheet.clear();
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    }

    // Ambil seluruh data yang ada
    const fullSheetValues = sheet.getDataRange().getDisplayValues();
    const currentHeaders = fullSheetValues[0].map(h => String(h).trim());

    const timestamp = getTimestamp();

    items.forEach(item => {
      let rowIndex = -1;
      for (let i = 1; i < fullSheetValues.length; i++) {
        if (String(fullSheetValues[i][0]).trim().toLowerCase() === String(semester).trim().toLowerCase() && 
            String(fullSheetValues[i][1]).trim() === String(kelas).trim() && 
            (String(fullSheetValues[i][2]).trim() === String(item.nis).trim() || String(fullSheetValues[i][3]).trim() === String(item.nama).trim())) {
          rowIndex = i + 1;
          break;
        }
      }

      let rowData = Array(currentHeaders.length).fill("");
      rowData[0] = semester;
      rowData[1] = kelas;
      rowData[2] = item.nis || "";
      rowData[3] = item.nama || "";

      let sum = 0;
      let count = 0;

      matchedTPList.forEach(tpCode => {
        const colIdx = currentHeaders.indexOf(tpCode);
        if (colIdx !== -1) {
          const val = item.tpScores[tpCode];
          if (val !== undefined && val !== null && val !== "" && !isNaN(val)) {
            rowData[colIdx] = Number(val);
            sum += Number(val);
            count++;
          }
        }
      });

      const avgColIdx = currentHeaders.indexOf('Rata-rata');
      if (avgColIdx !== -1) {
        rowData[avgColIdx] = count > 0 ? (sum / count).toFixed(1) : 0;
      }

      const tsColIdx = currentHeaders.indexOf('Timestamp');
      if (tsColIdx !== -1) {
        rowData[tsColIdx] = timestamp;
      }

      if (rowIndex !== -1) {
        sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
    });

    return { status: 'success', message: `Data nilai harian berhasil disimpan ke sheet '${sheetName}'.` };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getASA() {
  return { status: 'success', data: readData('ASA') };
}

function saveASASingle(payload) {
  try {
    initSheetHeaders('ASA');
    const sheet = getSheet('ASA');
    if (!sheet) throw new Error("Sheet ASA tidak ditemukan.");
    
    const fullSheetValues = sheet.getDataRange().getValues();
    const displayValues = sheet.getDataRange().getDisplayValues();
    
    let rowIndex = -1;
    for (let i = 1; i < displayValues.length; i++) {
      if (String(displayValues[i][0]).trim() === String(payload.kelas).trim() && 
          String(displayValues[i][1]).trim() === String(payload.nama).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    
    let rowData;
    if (rowIndex !== -1) {
      rowData = fullSheetValues[rowIndex - 1];
    } else {
      rowData = [payload.kelas, payload.nama, "", "", ""];
    }
    
    if (payload.semester === 'Ganjil') {
      rowData[2] = Number(payload.nilai);
    } else if (payload.semester === 'Genap') {
      rowData[3] = Number(payload.nilai);
    }
    
    rowData[4] = getTimestamp();
    
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    return { status: 'success', message: 'Nilai Asesmen Sumatif Akhir berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getJurnalGB() {
  return { status: 'success', data: readData('Jurnal_GB') };
}

function addJurnalGB(payload) {
  try {
    const sheet = getSheet('Jurnal_GB');
    if (!sheet) throw new Error("Gagal mengakses tabel Jurnal_GB.");
    const row = [
      payload.hariTanggal, payload.kelas, payload.jamKe, payload.materiTopik,
      payload.tujuanPembelajaran, payload.kegiatanPembelajaran, payload.kehadiran,
      payload.penilaian, payload.refleksiTindakLanjut, getTimestamp()
    ];
    sheet.appendRow(row);
    return { status: 'success', message: 'Jurnal Mengajar berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}
