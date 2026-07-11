/**
 * =========================================================================
 * GOOGLE APPS SCRIPT BACKEND CODE
 * Paste semua kode ini ke dalam file Code.gs di Apps Script
 * =========================================================================
 */

const DB_NAME = "DB_ADMINISTRASI_SEKOLAH_LENGKAP";
const TARGET_SPREADSHEET_ID = "1HFqRPz2sKZkz0V1iw8lLu6-DGyUtz-zh06Q3FalXUaY";
const TARGET_DRIVE_FOLDER_ID = "1BXL9ELJDAGlhaDxINn0uxTkx_pwkMjvo"; // ID Folder Drive Utama Anda

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate().setTitle('Sistem Administrasi Guru')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function openDatabase() {
  try {
    if (TARGET_SPREADSHEET_ID) {
      const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      setupStructure(ss); 
      return ss;
    }
  } catch (e) { 
    console.warn("Spreadsheet target gagal dibuka: " + e.message); 
  }
  return null;
}

function setupStructure(ss) {
  if (!ss) return;
  const struct = [
    { name: "Kelas", headers: ["Nama Kelas"] },
    { name: "Mapel", headers: ["Nama Mapel"] },
    { name: "Users", headers: ["Nama", "Role", "MapelAmpuan", "KelasAmpuan", "Status", "Password", "NIK", "NIP", "GantiPass"] },
    { name: "DataSiswa", headers: ["No", "Nama Siswa", "Kelas", "NISN"] },
    { name: "Absensi", headers: ["Waktu", "Kelas", "Mapel", "PertemuanKe", "Tanggal", "Nama Siswa", "Status", "Nama Guru", "Bulan", "Tahun"] },
    { name: "Nilai", headers: ["Waktu", "JudulPenilaian", "Kategori", "Mapel", "Kelas", "Nama Siswa", "Nilai", "Nama Guru", "Tugas Ke", "Tanggal", "Hari"] },
    { name: "Agenda", headers: ["Hari/Tgl", "Kelas", "Mapel", "Materi", "Status", "Ketuntasan", "JmlHadir", "JmlTdkHadir", "PertemuanKe", "Nama Guru", "Keterangan"] },
    { name: "SiswaBimbingan", headers: ["Nama Siswa", "Kelas", "Guru Wali", "File Data", "Nama File"] },
    { name: "BimbinganWali", headers: ["Tanggal", "Nama Siswa", "Kelas", "Jenis Bimbingan", "Masalah/Topik", "Solusi/Hasil", "Guru Wali", "File Data", "Nama File"] },
    { name: "Jadwal", headers: ["Hari", "Jam", "Kelas", "Mata Pelajaran", "Nama Guru"] },
    { name: "ModulAjar", headers: ["Nama Berkas", "Mata Pelajaran", "Tingkat Kelas", "Tipe Berkas", "Tanggal", "Nama Guru", "File Data", "Catatan Admin"] },
    { name: "StatusModul", headers: ["NIP Guru", "Nama Berkas", "Status", "Link Tautan", "Pengoreksi"] },
    { name: "Pengesahan", headers: ["NIP Guru", "Nama Berkas", "Link Pengesahan"] }, // SHEET BARU UNTUK PENGESAHAN
    { name: "TautanPortal", headers: ["Jenis Portal", "URL Tautan"] }
  ];

  struct.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name); 
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight("bold").setBackground("#1e3a8a").setFontColor("white");
    }
  });
}

/** * FUNGSI INTI UNTUK UPLOAD KE GOOGLE DRIVE */
function uploadFileToDrive(base64Data, mimeType, fileName, subFolderName) {
  try {
    const mainFolder = DriveApp.getFolderById(TARGET_DRIVE_FOLDER_ID);
    let targetFolder = null;
    
    const folders = mainFolder.getFoldersByName(subFolderName);
    if (folders.hasNext()) {
      targetFolder = folders.next();
    } else {
      targetFolder = mainFolder.createFolder(subFolderName);
    }
    
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = targetFolder.createFile(blob);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      console.warn("Peringatan Izin Drive: " + shareErr.message);
    }
    
    return file.getUrl();
  } catch (e) {
    throw new Error("Gagal mengunggah file ke Drive: " + e.message);
  }
}

// -------------------------------------------------------------------------
// DATA PENGGUNA DAN MASTER DATA
// -------------------------------------------------------------------------
function getUsers() {
  const ss = openDatabase(); 
  let result = [];
  
  if(ss) {
      const sheet = ss.getSheetByName("Users");
      if (!sheet) return [];
      
      if (sheet.getLastRow() < 2) {
          sheet.appendRow(["Administrator", "Admin", "Semua", "Semua", "Aktif", "admin123", "admin", "admin", "Sudah"]);
          SpreadsheetApp.flush(); // Eksekusi segera agar siap dibaca
      }
      
      const data = sheet.getDataRange().getValues();
      
      result = data.slice(1).filter(r => {
          const nama = r[0] ? String(r[0]).trim() : "";
          const status = r[4] ? String(r[4]).trim().toLowerCase() : "";
          return nama !== "" && (status === "aktif" || status === ""); 
      }).map(r => {
        let nikRaw = r[6] != null ? r[6].toString().trim() : "";
        let nipRaw = r[7] != null ? r[7].toString().trim() : "";
        if(nikRaw.endsWith(".0")) nikRaw = nikRaw.slice(0, -2);
        if(nipRaw.endsWith(".0")) nipRaw = nipRaw.slice(0, -2);

        return { 
          nama: String(r[0] || "").trim(), 
          role: String(r[1] || "Guru").trim(), 
          mapel: String(r[2] || "").trim(), 
          kelas: String(r[3] || "").trim(), 
          password: String(r[5] || "123456").trim(), 
          nik: nikRaw, 
          nip: nipRaw, 
          gantiPass: String(r[8] || "Belum").trim()
        };
      });
  }

  if (result.length === 0) {
      result.push({
          nama: "Admin (Sistem)", role: "Admin", mapel: "Semua", kelas: "Semua", 
          password: "admin123", nik: "admin", nip: "admin", gantiPass: "Sudah"
      });
  }
  
  return result;
}

function getMasterData() { 
  const ss = openDatabase(); if(!ss) return { mapel: [], kelas: [] };
  const sMapel = ss.getSheetByName("Mapel"); const sKelas = ss.getSheetByName("Kelas"); const sDataSiswa = ss.getSheetByName("DataSiswa");
  
  let mapel = sMapel.getLastRow() > 1 ? sMapel.getRange(2, 1, sMapel.getLastRow()-1, 1).getValues().flat().filter(String) : [];
  let kelas = sKelas.getLastRow() > 1 ? sKelas.getRange(2, 1, sKelas.getLastRow()-1, 1).getValues().flat().filter(String) : [];
  
  if (sDataSiswa && sDataSiswa.getLastRow() > 1) {
      const studentClasses = sDataSiswa.getRange(2, 3, sDataSiswa.getLastRow()-1, 1).getValues().flat().filter(String);
      kelas = [...new Set([...kelas, ...studentClasses])];
  }
  return { mapel: mapel, kelas: kelas };
}

function getSiswaByKelas(k) { 
  const sheet = openDatabase().getSheetByName("DataSiswa"); if(!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
  return data.filter(r => String(r[1]).trim().toUpperCase() === String(k).trim().toUpperCase()).map(r => r[0]).sort();
}

// -------------------------------------------------------------------------
// MODUL AJAR (DENGAN FIX SINKRONISASI / FLUSH & PENGESAHAN)
// -------------------------------------------------------------------------
function tambahModulDrive(rowData, base64Data, mimeType, fileName) {
  const namaGuru = rowData[5];
  const fileUrl = uploadFileToDrive(base64Data, mimeType, fileName, namaGuru);
  rowData[6] = fileUrl; 
  const ss = openDatabase();
  ss.getSheetByName("ModulAjar").appendRow(rowData);
  
  SpreadsheetApp.flush(); 
  
  return "Modul berhasil diunggah ke Google Drive dan Database Anda!";
}

function prosesPerbaikiModulDrive(idx, nip, namaBerkas, base64Data, mimeType, fileName) {
    const ss = openDatabase(); 
    const sModul = ss.getSheetByName("ModulAjar");
    const namaGuru = sModul.getRange(idx + 2, 6).getValue();
    const fileUrl = uploadFileToDrive(base64Data, mimeType, fileName, namaGuru);
    
    sModul.getRange(idx + 2, 4).setValue(mimeType); 
    sModul.getRange(idx + 2, 5).setValue(new Date().toLocaleDateString()); 
    sModul.getRange(idx + 2, 7).setValue(fileUrl);
    
    const sStatus = ss.getSheetByName("StatusModul"); const statusData = sStatus.getDataRange().getValues(); let isFound = false;
    for (let i = 1; i < statusData.length; i++) {
        if (String(statusData[i][0]).trim() == String(nip).trim() && String(statusData[i][1]).trim() == String(namaBerkas).trim()) { 
            sStatus.getRange(i + 1, 3).setValue("Diperbaiki"); isFound = true; break; 
        }
    }
    if (!isFound) { sStatus.appendRow([nip, namaBerkas, "Diperbaiki", "", ""]); } 
    
    SpreadsheetApp.flush();
    return "Berkas perbaikan berhasil diunggah ke Drive Admin!";
}

function getModul(currentUser) {
  const ss = openDatabase();
  const usersData = ss.getSheetByName("Users").getDataRange().getValues();
  let nameToNip = {};
  for(let i=1; i<usersData.length; i++) { 
      nameToNip[String(usersData[i][0]).trim()] = usersData[i][7] ? String(usersData[i][7]).trim() : ""; 
  }
  
  const sModul = ss.getSheetByName("ModulAjar");
  if (!sModul || sModul.getLastRow() < 2) return [];
  
  const maxCol = Math.max(sModul.getLastColumn(), 8);
  const modulData = sModul.getRange(2, 1, sModul.getLastRow() - 1, maxCol).getValues();
  
  // Ambil Data Status Modul (dan Pengoreksi)
  const sStatus = ss.getSheetByName("StatusModul"); let statusMap = {}; 
  if (sStatus && sStatus.getLastRow() > 1) {
      const maxStatCol = Math.max(sStatus.getLastColumn(), 5);
      const statusData = sStatus.getRange(2, 1, sStatus.getLastRow() - 1, maxStatCol).getValues();
      statusData.forEach(r => { 
          let key = String(r[0] || "").trim() + "_" + String(r[1] || "").trim(); 
          statusMap[key] = { status: r[2], link: r[3], pengoreksi: r[4] }; 
      });
  }

  // Ambil Data Pengesahan
  const sPengesahan = ss.getSheetByName("Pengesahan"); let pengesahanMap = {};
  if (sPengesahan && sPengesahan.getLastRow() > 1) {
      const maxPengCol = Math.max(sPengesahan.getLastColumn(), 3);
      const pengesahanData = sPengesahan.getRange(2, 1, sPengesahan.getLastRow() - 1, maxPengCol).getValues();
      pengesahanData.forEach(r => {
          let key = String(r[0] || "").trim() + "_" + String(r[1] || "").trim();
          pengesahanMap[key] = String(r[2] || "").trim();
      });
  }
  
  let result = [];
  for(let i=0; i<modulData.length; i++) {
      let r = modulData[i]; 
      let guruName = String(r[5] || "").trim(); 
      let nipGuru = nameToNip[guruName] || "";
      let uRole = String(currentUser.role).toLowerCase();
      
      if (uRole === 'admin' || guruName === String(currentUser.nama).trim()) {
          let namaBerkas = String(r[0] || "").trim();
          let matchKey = nipGuru + "_" + namaBerkas; 
          let accData = statusMap[matchKey] || { status: 'Menunggu', link: '', pengoreksi: '' };
          let linkPengesahan = pengesahanMap[matchKey] || "";
          
          result.push({ 
             index: i, 
             nama: namaBerkas, 
             mapel: String(r[1] || ""), 
             tingkat: String(r[2] || ""), 
             tipe: String(r[3] || ""), 
             tanggal: r[4] instanceof Date ? Utilities.formatDate(r[4], "GMT+7", "yyyy-MM-dd") : String(r[4] || ""), 
             guru: guruName, 
             nip: nipGuru, 
             fileData: String(r[6] || ""), 
             catatanAdmin: String(r[7] || ""), 
             status: accData.status, 
             linkAcc: accData.link,
             pengoreksi: accData.pengoreksi || "",
             linkPengesahan: linkPengesahan
          });
      }
  } 
  return result;
}

function setModulStatus(idx, nip, namaBerkas, status, pengoreksi, catatan) { 
    const ss = openDatabase(); ss.getSheetByName("ModulAjar").getRange(idx + 2, 8).setValue(catatan); 
    const sStatus = ss.getSheetByName("StatusModul"); 
    const maxStatCol = Math.max(sStatus.getLastColumn(), 5);
    const statusData = sStatus.getRange(1, 1, sStatus.getLastRow(), maxStatCol).getValues(); 
    let isFound = false;
    for (let i = 1; i < statusData.length; i++) {
        if (String(statusData[i][0]).trim() == String(nip).trim() && String(statusData[i][1]).trim() == String(namaBerkas).trim()) {
            sStatus.getRange(i + 1, 3).setValue(status); 
            sStatus.getRange(i + 1, 5).setValue(pengoreksi); 
            isFound = true; break;
        }
    }
    if (!isFound) { sStatus.appendRow([nip, namaBerkas, status, "", pengoreksi]); } 
    
    SpreadsheetApp.flush();
    return "Status dan Catatan berhasil disimpan!"; 
}

function deleteModul(index, currentUser) { 
  const sheet = openDatabase().getSheetByName("ModulAjar"); const data = sheet.getDataRange().getValues(); let count = 0;
  for (let i = 1; i < data.length; i++) { 
      if (String(currentUser.role).toLowerCase() === 'admin' || String(data[i][5]).trim() === String(currentUser.nama).trim()) { 
          if (count === index) { 
              sheet.deleteRow(i + 1); 
              SpreadsheetApp.flush();
              return "Modul dihapus!"; 
          } 
          count++; 
      } 
  }
}

// -------------------------------------------------------------------------
// GURU WALI / KONSELING
// -------------------------------------------------------------------------
function simpanBimbinganWaliServer(tgl, siswa, kls, jns, top, sol, namaGuru, base64Data, mimeType, fileName) {
  let fileUrl = "";
  if (base64Data && base64Data !== "") {
     fileUrl = uploadFileToDrive(base64Data, mimeType, fileName, namaGuru);
  }
  const sheet = openDatabase().getSheetByName("BimbinganWali");
  sheet.appendRow([tgl, siswa, kls, jns, top, sol, namaGuru, fileUrl, fileName]);
  return "Jurnal Bimbingan berhasil tersimpan ke sistem!";
}

function getBimbinganRecords(currentUser) {
  const sheet = openDatabase().getSheetByName("BimbinganWali"); if(!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  return data.filter(r => String(currentUser.role).toLowerCase() === 'admin' || String(r[6]).trim() === String(currentUser.nama).trim()).map(r => ({
    tanggal: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd") : r[0], siswa: r[1], kelas: r[2], jenis: r[3], topik: r[4], solusi: r[5], fileData: r[7], namaFile: r[8]
  }));
}
function getSiswaBimbingan(currentUser) {
  const sheet = openDatabase().getSheetByName("SiswaBimbingan"); if(!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues(); return data.filter(r => String(currentUser.role).toLowerCase() === 'admin' || String(r[2]).trim() === String(currentUser.nama).trim()).map(r => ({ nama: r[0], kelas: r[1] }));
}
function tambahSiswaBimbingan(siswa, kls, guru) { openDatabase().getSheetByName("SiswaBimbingan").appendRow([siswa, kls, guru, "", ""]); return "Didaftarkan!"; }

// -------------------------------------------------------------------------
// ABSENSI & PENILAIAN
// -------------------------------------------------------------------------
function simpanAbsensiServer(kelas, mapel, pertemuan, tanggal, namaGuru, arrSiswa) {
  if (!arrSiswa || !Array.isArray(arrSiswa)) { arrSiswa = []; }
  if (arrSiswa.length === 0) return "Tidak ada data siswa untuk disimpan.";
  const sheet = openDatabase().getSheetByName("Absensi"); const d = new Date(tanggal); let rows = [];
  arrSiswa.forEach(s => rows.push([new Date(), kelas, mapel, pertemuan, tanggal, s.nama, s.status, namaGuru, d.getMonth() + 1, d.getFullYear()]));
  if (rows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 10).setValues(rows); return "Absensi berhasil disimpan!";
}
function getAbsenRecords(kelas, mapel, currentUser) {
  const sheet = openDatabase().getSheetByName("Absensi"); if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  let results = [];
  for (let i = 0; i < data.length; i++) {
    let r = data[i]; let isOwner = (String(currentUser.role).toLowerCase() === 'admin') || (String(r[7]).trim() === String(currentUser.nama).trim());
    if (isOwner && r[1] === kelas && r[2] === mapel) { results.push({ rowIdx: i + 2, kelas: r[1], mapel: r[2], pertemuan: r[3], tanggal: r[4] instanceof Date ? Utilities.formatDate(r[4], "GMT+7", "yyyy-MM-dd") : r[4], nama: r[5], status: r[6] }); }
  } return results;
}
function updateAbsensiStatus(rowIdx, newStatus) {
  if (!rowIdx) throw new Error("ID Baris Tidak Valid.");
  const sheet = openDatabase().getSheetByName("Absensi"); sheet.getRange(rowIdx, 7).setValue(newStatus); return "Status Kehadiran Berhasil Diubah!";
}

function simpanNilaiServer(judulFull, kategori, mapel, kelas, namaGuru, arrSiswa, tk, tgl, hari) {
  if (!arrSiswa || !Array.isArray(arrSiswa)) { arrSiswa = []; }
  if (arrSiswa.length === 0) return "Tidak ada data nilai untuk disimpan.";
  const sheet = openDatabase().getSheetByName("Nilai"); let rows = [];
  arrSiswa.forEach(s => rows.push([new Date(), judulFull, kategori, mapel, kelas, s.nama, parseFloat(s.nilai) || 0, namaGuru, tk, tgl, hari]));
  if (rows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 11).setValues(rows); return "Data nilai berhasil direkam!";
}
function getNilaiRecords(kelas, mapel, currentUser) {
  const sheet = openDatabase().getSheetByName("Nilai"); if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues(); 
  let results = [];
  for (let i = 0; i < data.length; i++) {
    let r = data[i]; let isOwner = (String(currentUser.role).toLowerCase() === 'admin') || (String(r[7]).trim() === String(currentUser.nama).trim());
    if (isOwner && r[4] === kelas && r[3] === mapel) { results.push({ rowIdx: i + 2, judul: r[1], kategori: r[2], mapel: r[3], kelas: r[4], nama: r[5], nilai: r[6], tanggal: r[9] instanceof Date ? Utilities.formatDate(r[9], "GMT+7", "yyyy-MM-dd") : r[9] }); }
  } return results;
}
function updateNilaiRecord(rowIdx, newNilai) {
  if (!rowIdx) throw new Error("ID Baris Tidak Valid.");
  const sheet = openDatabase().getSheetByName("Nilai"); sheet.getRange(rowIdx, 7).setValue(newNilai); return "Poin Nilai Berhasil Dikoreksi!";
}

function getLegerDinamis(kelas, mapel, currentUser, bobot) {
  const sheet = openDatabase().getSheetByName("Nilai"); if (!sheet || sheet.getLastRow() < 2) return { error: "Belum ada data nilai." };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  const filtered = data.filter(r => { const isOwner = (String(currentUser.role).toLowerCase() === 'admin') || (String(r[7]).trim() === String(currentUser.nama).trim()); return isOwner && r[4] === kelas && r[3] === mapel; });
  if (filtered.length === 0) return { error: "Tidak ditemukan rekam nilai untuk kelas & mapel ini." };
  
  let studentsMap = {}; filtered.forEach(r => { studentsMap[r[5]] = true; }); let uniqueStudents = Object.keys(studentsMap).sort(); 
  let rowsOut = [];
  uniqueStudents.forEach((student, idx) => {
    let rowObj = { no: idx + 1, nama: student }; 
    let sumTugas = 0, countTugas = 0, sumUH = 0, countUH = 0, sumUTS = 0, countUTS = 0, sumUS = 0, countUS = 0, sumPoin = 0, countPoin = 0;

    filtered.forEach(r => {
        if (r[5] === student) {
            let val = parseFloat(r[6]) || 0; let kat = r[2] ? String(r[2]).toLowerCase() : "";
            if (kat.includes("tugas")) { sumTugas += val; countTugas++; }
            else if (kat.includes("harian") || kat === "uh") { sumUH += val; countUH++; }
            else if (kat.includes("uts")) { sumUTS += val; countUTS++; }
            else if (kat.includes("us") || kat.includes("uas")) { sumUS += val; countUS++; }
            else if (kat.includes("poin")) { sumPoin += val; countPoin++; }
        }
    });

    rowObj.avg_tugas = countTugas > 0 ? (sumTugas / countTugas).toFixed(1) : 0; rowObj.avg_uh = countUH > 0 ? (sumUH / countUH).toFixed(1) : 0;
    rowObj.avg_uts = countUTS > 0 ? (sumUTS / countUTS).toFixed(1) : 0; rowObj.avg_us = countUS > 0 ? (sumUS / countUS).toFixed(1) : 0; rowObj.avg_poin = countPoin > 0 ? (sumPoin / countPoin).toFixed(1) : 0;
    rowObj.nilai_raport = ((rowObj.avg_tugas * (bobot.tugas / 100)) + (rowObj.avg_uh * (bobot.uh / 100)) + (rowObj.avg_uts * (bobot.uts / 100)) + (rowObj.avg_us * (bobot.us / 100)) + (rowObj.avg_poin * (bobot.poin / 100))).toFixed(2);
    rowsOut.push(rowObj);
  }); 
  return { headers: [], data: rowsOut };
}

function getAllLegerRaw() { 
  const sheet = openDatabase().getSheetByName("Nilai"); 
  return sheet && sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 8).getValues().map(r=>({
    judul: r[1], 
    kategori: r[2], 
    mapel: r[3], 
    kelas: r[4], 
    siswa: r[5], 
    nilai: r[6], 
    guru: r[7]
  })) : []; 
}

function getDetailSiswa(kelas, nama) {
    const ss = openDatabase(); if (!ss) return { nilai: [], absensi: [] };
    let resultNilai = []; const sNilai = ss.getSheetByName("Nilai");
    if (sNilai && sNilai.getLastRow() > 1) { const dataN = sNilai.getRange(2, 1, sNilai.getLastRow() - 1, 10).getValues(); resultNilai = dataN.filter(r => r[3] === kelas && r[4] === nama).map(r => ({ judul: r[1], mapel: r[2], nilai: r[5], guru: r[6], tugasKe: r[7] || "", tanggal: r[8] instanceof Date ? Utilities.formatDate(r[8], "GMT+7", "yyyy-MM-dd") : r[8] })); }
    let resultAbsen = []; const sAbsen = ss.getSheetByName("Absensi");
    if (sAbsen && sAbsen.getLastRow() > 1) { const dataA = sAbsen.getRange(2, 1, sAbsen.getLastRow() - 1, 10).getValues(); resultAbsen = dataA.filter(r => r[1] === kelas && r[5] === nama).map(r => ({ mapel: r[2], pertemuan: r[3], tanggal: r[4] instanceof Date ? Utilities.formatDate(r[4], "GMT+7", "yyyy-MM-dd") : r[4], status: r[6], guru: r[7] })); } 
    return { nilai: resultNilai, absensi: resultAbsen };
}

// -------------------------------------------------------------------------
// AGENDA DAN JADWAL
// -------------------------------------------------------------------------
function tambahAgenda(row) { openDatabase().getSheetByName("Agenda").appendRow(row); return "Jurnal agenda tersimpan!"; }
function getAgendaList(currentUser) {
  const sheet = openDatabase().getSheetByName("Agenda"); if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  return data.filter(r => String(currentUser.role).toLowerCase() === 'admin' || String(r[9]).trim() === String(currentUser.nama).trim()).map(r => ({ tanggal: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd") : r[0], kelas: r[1], mapel: r[2], materi: r[3], status: r[4], tuntas: r[5], hadir: r[6], absen: r[7], pert: r[8], ket: r[10] }));
}
function tambahJadwal(row) { openDatabase().getSheetByName("Jadwal").appendRow(row); return "Jadwal berhasil ditambahkan!"; }
function deleteJadwal(index, currentUser) {
  const sheet = openDatabase().getSheetByName("Jadwal"); const data = sheet.getDataRange().getValues(); let count = 0;
  for (let i = 1; i < data.length; i++) { if (String(currentUser.role).toLowerCase() === 'admin' || String(data[i][4]).trim() === String(currentUser.nama).trim()) { if (count === index) { sheet.deleteRow(i + 1); return "Jadwal dihapus!"; } count++; } }
}
function getJadwal(currentUser) {
  const sheet = openDatabase().getSheetByName("Jadwal"); if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return data.filter(r => String(currentUser.role).toLowerCase() === 'admin' || String(r[4]).trim() === String(currentUser.nama).trim()).map(r => ({ hari: r[0], jam: r[1], kelas: r[2], mapel: r[3], guru: r[4] }));
}

// -------------------------------------------------------------------------
// ADMIN DATABASE
// -------------------------------------------------------------------------
function getAllSiswaDB() { const sheet = openDatabase().getSheetByName("DataSiswa"); if(!sheet || sheet.getLastRow() < 2) return []; const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues(); return data.map((r, i) => ({ index: i, no: r[0], nama: r[1], kelas: r[2], nisn: r[3] || "" })); }
function tambahSiswaDB(nama, kelas, nisn) { const sheet = openDatabase().getSheetByName("DataSiswa"); const nextNo = sheet.getLastRow(); sheet.appendRow([nextNo, nama, kelas, nisn]); return "Siswa berhasil ditambahkan ke Database!"; }
function hapusSiswaDB(index) { const sheet = openDatabase().getSheetByName("DataSiswa"); sheet.deleteRow(index + 2); return "Siswa dihapus dari Database."; }
function tambahGuruDB(nama, nip, pass, mapel, kelas) { const sheet = openDatabase().getSheetByName("Users"); sheet.appendRow([nama, "Guru", mapel, kelas, "Aktif", pass, nip, nip, "Belum"]); return "Akun Guru berhasil dibuat!"; }
function hapusGuruDB(nip) { const sheet = openDatabase().getSheetByName("Users"); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++) { if(data[i][7] == nip || data[i][6] == nip) { sheet.deleteRow(i + 1); return "Akun Guru dihapus."; } } return "Guru tidak ditemukan."; }
function tambahBanyakSiswaDB(dataArray) { const sheet = openDatabase().getSheetByName("DataSiswa"); let nextNo = sheet.getLastRow(); let rowsToInsert = []; for(let i=0; i<dataArray.length; i++) { let s = dataArray[i]; rowsToInsert.push([nextNo + i, s.nama, s.kelas, s.nisn]); } if(rowsToInsert.length > 0) { sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, 4).setValues(rowsToInsert); } return rowsToInsert.length + " data siswa berhasil ditambahkan!"; }
function tambahBanyakGuruDB(dataArray) { const sheet = openDatabase().getSheetByName("Users"); let rowsToInsert = []; for(let i=0; i<dataArray.length; i++) { let g = dataArray[i]; rowsToInsert.push([g.nama, "Guru", g.mapel, g.kelas, "Aktif", g.pass, g.nip, g.nip, "Belum"]); } if(rowsToInsert.length > 0) { sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, 9).setValues(rowsToInsert); } return rowsToInsert.length + " akun guru berhasil ditambahkan!"; }

function getTautanPortal() { 
  const ss = openDatabase(); 
  if(!ss) return { guruWali: "", waliKelas: "", bk: "", kurikulum: "", modulAjar: "" }; 
  let sheet = ss.getSheetByName("TautanPortal"); 
  if(sheet.getLastRow() < 2) { 
    sheet.appendRow(["Administrasi Guru Wali", ""]); 
    sheet.appendRow(["Wali Kelas", ""]); 
    sheet.appendRow(["BK", ""]); 
    sheet.appendRow(["Kurikulum", ""]); 
    sheet.appendRow(["Modul Ajar", ""]); 
  } 
  const data = sheet.getDataRange().getValues(); 
  let result = { guruWali: "", waliKelas: "", bk: "", kurikulum: "", modulAjar: "" }; 
  for(let i = 1; i < data.length; i++) { 
    let jenis = data[i][0].toString().trim(); 
    let link = data[i][1].toString().trim(); 
    if(jenis === "Administrasi Guru Wali") result.guruWali = link; 
    if(jenis === "Wali Kelas") result.waliKelas = link; 
    if(jenis === "BK") result.bk = link; 
    if(jenis === "Kurikulum") result.kurikulum = link; 
    if(jenis === "Modul Ajar") result.modulAjar = link; 
  } 
  return result; 
}

function simpanTautanPortal(guruWali, waliKelas, bk, kurikulum, modulAjar) { 
  const ss = openDatabase(); 
  let sheet = ss.getSheetByName("TautanPortal"); 
  if(!sheet) return "Sheet TautanPortal tidak ditemukan."; 
  sheet.clear(); 
  sheet.appendRow(["Jenis Portal", "URL Tautan"]); 
  sheet.appendRow(["Administrasi Guru Wali", guruWali]); 
  sheet.appendRow(["Wali Kelas", waliKelas]); 
  sheet.appendRow(["BK", bk]); 
  sheet.appendRow(["Kurikulum", kurikulum]); 
  sheet.appendRow(["Modul Ajar", modulAjar]); 
  sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#1e3a8a").setFontColor("white"); 
  return "Tautan Portal Layanan berhasil diperbarui!"; 
}

function updatePassword(nip, newPass) { const ss = openDatabase(); const sheet = ss.getSheetByName("Users"); const data = sheet.getDataRange().getValues(); for(let i = 1; i < data.length; i++) { if(String(data[i][7]).trim() == String(nip).trim() || String(data[i][6]).trim() == String(nip).trim()) { sheet.getRange(i + 1, 6).setValue(newPass); sheet.getRange(i + 1, 9).setValue("Sudah"); return "Kata sandi Anda berhasil diperbarui!"; } } return "Gagal, Akun tidak ditemukan."; }
function editSiswaDB(index, nama, kelas, nisn) { const sheet = openDatabase().getSheetByName("DataSiswa"); sheet.getRange(index + 2, 2).setValue(nama); sheet.getRange(index + 2, 3).setValue(kelas); sheet.getRange(index + 2, 4).setValue(nisn); return "Data Siswa berhasil diperbarui!"; }
function editGuruDB(oldNip, nama, nip, pass, mapel, kelas) { const sheet = openDatabase().getSheetByName("Users"); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++) { if(String(data[i][7]).trim() == String(oldNip).trim() || String(data[i][6]).trim() == String(oldNip).trim()) { sheet.getRange(i + 1, 1).setValue(nama); sheet.getRange(i + 1, 3).setValue(mapel); sheet.getRange(i + 1, 4).setValue(kelas); sheet.getRange(i + 1, 6).setValue(pass); sheet.getRange(i + 1, 7).setValue(nip); sheet.getRange(i + 1, 8).setValue(nip); return "Akun Guru berhasil diperbarui!"; } } return "Guru tidak ditemukan di Database."; }