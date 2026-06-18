// ============================================================
//  PERPUSTAKAAN DIGITAL - Google Apps Script Backend
//  Sheet yang dibutuhkan:
//    1. DataSiswa   → NIS, Nama, Kelas, Password, Role
//    2. Kunjungan   → ID, Tanggal, Jam, NIS, Nama, Kelas, Kegiatan
//    3. BukuDigital → ID, Judul, Penulis, Kategori, Link, Cover
//    4. Admin       → Username, Password
// ============================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ── Nama sheet ──────────────────────────────────────────────
const SHEET = {
  SISWA    : "DataSiswa",
  KUNJUNGAN: "Kunjungan",
  BUKU     : "BukuDigital",
  ADMIN    : "Admin",
  LOG_BACA : "LogBaca",
};

// ── Utilitas tanggal WIB ────────────────────────────────────
function getWIBDate() {
  const now = new Date();
  const wib = new Date(now.getTime() + 7 * 60 * 60 * 1000);
  const y = wib.getUTCFullYear();
  const m = String(wib.getUTCMonth() + 1).padStart(2, "0");
  const d = String(wib.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function getWIBTime() {
  const now = new Date();
  const wib = new Date(now.getTime() + 7 * 60 * 60 * 1000);
  const h = String(wib.getUTCHours()).padStart(2, "0");
  const mi = String(wib.getUTCMinutes()).padStart(2, "0");
  return `${h}:${mi}`;
}

// ── Init: buat sheet jika belum ada ────────────────────────
function initSheets() {
  const headers = {
    [SHEET.SISWA]    : ["NIS", "Nama", "Kelas", "Password", "Role"],
    [SHEET.KUNJUNGAN]: ["ID", "Tanggal", "Jam", "NIS", "Nama", "Kelas", "Kegiatan"],
    [SHEET.BUKU]     : ["ID", "Judul", "Penulis", "Kategori", "Link", "Cover"],
    [SHEET.ADMIN]    : ["Username", "Password"],
    [SHEET.LOG_BACA] : ["ID", "Tanggal", "JamMulai", "JamSelesai", "DurasiMenit", "NIS", "Nama", "Kelas", "IDBuku", "JudulBuku"],
  };

  for (const [name, cols] of Object.entries(headers)) {
    let sh = SS.getSheetByName(name);
    if (!sh) {
      sh = SS.insertSheet(name);
      sh.appendRow(cols);
      sh.getRange(1, 1, 1, cols.length)
        .setFontWeight("bold")
        .setBackground("#1a4d2e")
        .setFontColor("#ffffff");
    }
  }

  // Seed admin default jika kosong
  const adminSh = SS.getSheetByName(SHEET.ADMIN);
  if (adminSh.getLastRow() < 2) {
    adminSh.appendRow(["admin", "admin123"]);
  }
}

// ────────────────────────────────────────────────────────────
//  ROUTER UTAMA
// ────────────────────────────────────────────────────────────
function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // Baca parameter: prioritaskan POST body JSON, fallback ke GET params
  let p = {};
  try {
    if (e.postData && e.postData.contents) {
      p = JSON.parse(e.postData.contents);
    }
  } catch(_) {}
  // Gabung dengan URL params (GET params tetap bisa dipakai)
  const q = e.parameter || {};
  for (const k in q) { if (!p[k]) p[k] = q[k]; }

  const action = p.action || "";
  let result;

  try {
    switch (action) {
      case "login"          : result = loginUser(p);           break;
      case "catatKunjungan" : result = catatKunjungan(p);      break;
      case "getKunjungan"   : result = getKunjungan(p);        break;
      case "getStatistik"   : result = getStatistik(p);        break;
      case "getBuku"        : result = getBuku(p);             break;
      case "tambahBuku"     : result = tambahBuku(p);          break;
      case "hapusBuku"      : result = hapusBuku(p);           break;
      case "getSiswa"       : result = getSiswaList(p);        break;
      case "tambahSiswa"    : result = tambahSiswa(p);         break;
      case "hapusKunjungan"     : result = hapusKunjungan(p);          break;
      case "getPeringkatBulanan": result = getPeringkatBulanan(p);    break;
      case "catatLogBaca"   : result = catatLogBaca(p);        break;
      case "getLogBaca"     : result = getLogBaca(p);          break;
      default                   : result = { ok: false, msg: "Action tidak dikenal: " + action };
    }
  } catch (err) {
    result = { ok: false, msg: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ────────────────────────────────────────────────────────────
//  AUTH
// ────────────────────────────────────────────────────────────
function loginUser({ nis, password, role }) {
  if (role === "admin") {
    const sh   = SS.getSheetByName(SHEET.ADMIN);
    const rows = sh.getDataRange().getValues().slice(1);
    const found = rows.find(r => String(r[0]) === nis && String(r[1]) === password);
    if (found) return { ok: true, role: "admin", nama: "Administrator" };
    return { ok: false, msg: "Username / password admin salah." };
  }

  // Siswa
  const sh   = SS.getSheetByName(SHEET.SISWA);
  const rows = sh.getDataRange().getValues().slice(1);
  const found = rows.find(r => String(r[0]) === nis && String(r[3]) === password);
  if (found) return { ok: true, role: "siswa", nis: found[0], nama: found[1], kelas: found[2] };
  return { ok: false, msg: "NIS / password salah." };
}

// ────────────────────────────────────────────────────────────
//  KUNJUNGAN
// ────────────────────────────────────────────────────────────
function catatKunjungan({ nis, nama, kelas, kegiatan }) {
  if (!nis || !nama) return { ok: false, msg: "Data tidak lengkap." };

  const sh      = SS.getSheetByName(SHEET.KUNJUNGAN);
  const tanggal = getWIBDate();
  const jam     = getWIBTime();

  // Cek duplikat hari ini
  const rows = sh.getDataRange().getValues().slice(1);
  const sudah = rows.find(r => String(r[3]) === nis && String(r[1]) === tanggal);
  if (sudah) return { ok: false, msg: "Kamu sudah mencatat kunjungan hari ini." };

  const id = "KJG" + Date.now();
  sh.appendRow([id, tanggal, jam, nis, nama, kelas || "-", kegiatan || "Baca Buku"]);
  return { ok: true, msg: "Kunjungan berhasil dicatat!", tanggal, jam };
}

function getKunjungan({ tanggal, nis }) {
  const sh   = SS.getSheetByName(SHEET.KUNJUNGAN);
  const rows = sh.getDataRange().getValues().slice(1);
  let data   = rows.map(r => ({
    id      : r[0],
    tanggal : r[1] instanceof Date ? r[1].toISOString().slice(0, 10) : String(r[1]),
    jam     : r[2],
    nis     : r[3],
    nama    : r[4],
    kelas   : r[5],
    kegiatan: r[6],
  }));

  if (tanggal) data = data.filter(d => d.tanggal === tanggal);
  if (nis)     data = data.filter(d => String(d.nis) === nis);

  return { ok: true, data };
}

function hapusKunjungan({ id }) {
  const sh   = SS.getSheetByName(SHEET.KUNJUNGAN);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      sh.deleteRow(i + 1);
      return { ok: true, msg: "Data kunjungan dihapus." };
    }
  }
  return { ok: false, msg: "Data tidak ditemukan." };
}

// ────────────────────────────────────────────────────────────
//  STATISTIK
// ────────────────────────────────────────────────────────────
function getStatistik() {
  const sh   = SS.getSheetByName(SHEET.KUNJUNGAN);
  const rows = sh.getDataRange().getValues().slice(1);
  const today = getWIBDate();

  const kunjunganHariIni = rows.filter(r => {
    const tgl = r[1] instanceof Date ? r[1].toISOString().slice(0, 10) : String(r[1]);
    return tgl === today;
  }).length;

  // Hitung total kunjungan per siswa (top 5)
  const counter = {};
  rows.forEach(r => {
    const key = `${r[3]}|${r[4]}|${r[5]}`;
    counter[key] = (counter[key] || 0) + 1;
  });
  const topPembaca = Object.entries(counter)
    .map(([k, v]) => { const [nis, nama, kelas] = k.split("|"); return { nis, nama, kelas, total: v }; })
    .sort((a, b) => b.total - a.total)
    .slice(0, 5);

  // Kunjungan 7 hari terakhir
  const tujuhHari = [];
  for (let i = 6; i >= 0; i--) {
    const d   = new Date();
    d.setDate(d.getDate() - i);
    const tgl = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
    const cnt = rows.filter(r => {
      const t = r[1] instanceof Date ? r[1].toISOString().slice(0,10) : String(r[1]);
      return t === tgl;
    }).length;
    tujuhHari.push({ tanggal: tgl, total: cnt });
  }

  // Kegiatan terbanyak
  const kegCounter = {};
  rows.forEach(r => { kegCounter[r[6]] = (kegCounter[r[6]] || 0) + 1; });
  const kegiatan = Object.entries(kegCounter).map(([nama, total]) => ({ nama, total }));

  return {
    ok: true,
    kunjunganHariIni,
    totalKunjungan: rows.length,
    topPembaca,
    tujuhHari,
    kegiatan,
  };
}

// ────────────────────────────────────────────────────────────
//  PERINGKAT BULANAN
// ────────────────────────────────────────────────────────────
function getPeringkatBulanan({ bulan, tahun, top }) {
  const sh   = SS.getSheetByName(SHEET.KUNJUNGAN);
  const rows = sh.getDataRange().getValues().slice(1);

  // Gunakan bulan & tahun saat ini jika tidak dikirim
  const now     = new Date();
  const wib     = new Date(now.getTime() + 7 * 60 * 60 * 1000);
  const tBulan  = bulan  ? parseInt(bulan)  : wib.getUTCMonth() + 1;
  const tTahun  = tahun  ? parseInt(tahun)  : wib.getUTCFullYear();
  const topN    = top    ? parseInt(top)    : 10;

  // Filter kunjungan sesuai bulan & tahun
  const filtered = rows.filter(r => {
    const tgl = r[1] instanceof Date ? r[1].toISOString().slice(0, 10) : String(r[1]);
    if (!tgl || tgl.length < 7) return false;
    const [y, m] = tgl.split("-").map(Number);
    return y === tTahun && m === tBulan;
  });

  // Hitung kunjungan per siswa
  const counter = {};
  filtered.forEach(r => {
    const key = `${r[3]}|${r[4]}|${r[5]}`;
    counter[key] = (counter[key] || 0) + 1;
  });

  const peringkat = Object.entries(counter)
    .map(([k, v]) => {
      const [nis, nama, kelas] = k.split("|");
      return { nis, nama, kelas, total: v };
    })
    .sort((a, b) => b.total - a.total)
    .slice(0, topN)
    .map((item, idx) => ({ ...item, peringkat: idx + 1 }));

  const namaBulan = ["Januari","Februari","Maret","April","Mei","Juni",
                     "Juli","Agustus","September","Oktober","November","Desember"];

  return {
    ok: true,
    bulan: tBulan,
    tahun: tTahun,
    labelBulan: `${namaBulan[tBulan - 1]} ${tTahun}`,
    totalKunjungan: filtered.length,
    peringkat,
  };
}


// ────────────────────────────────────────────────────────────
//  LOG BACA BUKU
// ────────────────────────────────────────────────────────────
function catatLogBaca({ nis, nama, kelas, idBuku, judulBuku, jamMulai, jamSelesai, durasiMenit, tanggal }) {
  if (!nis || !idBuku) return { ok: false, msg: "Data tidak lengkap." };

  const sh  = SS.getSheetByName(SHEET.LOG_BACA);
  const tgl = tanggal || getWIBDate();   // gunakan tanggal WIB dari client
  const id  = "LB" + Date.now();

  sh.appendRow([
    id,
    tgl,
    jamMulai     || getWIBTime(),
    jamSelesai   || getWIBTime(),
    durasiMenit  || 0,
    nis,
    nama         || "-",
    kelas        || "-",
    idBuku,
    judulBuku    || "-",
  ]);

  return { ok: true, msg: "Log baca berhasil disimpan.", id };
}

function getLogBaca({ tanggal, nis, idBuku }) {
  const sh   = SS.getSheetByName(SHEET.LOG_BACA);
  const rows = sh.getDataRange().getValues().slice(1);

  let data = rows.map(r => ({
    id          : r[0],
    tanggal     : r[1] instanceof Date ? r[1].toISOString().slice(0, 10) : String(r[1]),
    jamMulai    : r[2],
    jamSelesai  : r[3],
    durasiMenit : r[4],
    nis         : r[5],
    nama        : r[6],
    kelas       : r[7],
    idBuku      : r[8],
    judulBuku   : r[9],
  }));

  if (tanggal) data = data.filter(d => d.tanggal === tanggal);
  if (nis)     data = data.filter(d => String(d.nis) === nis);
  if (idBuku)  data = data.filter(d => String(d.idBuku) === idBuku);

  // Hitung total durasi per buku (untuk statistik)
  const bukuStats = {};
  data.forEach(d => {
    const key = d.idBuku;
    if (!bukuStats[key]) bukuStats[key] = { idBuku: d.idBuku, judulBuku: d.judulBuku, totalMenit: 0, totalSesi: 0 };
    bukuStats[key].totalMenit += Number(d.durasiMenit) || 0;
    bukuStats[key].totalSesi  += 1;
  });

  return {
    ok: true,
    data,
    bukuStats: Object.values(bukuStats).sort((a, b) => b.totalMenit - a.totalMenit),
    totalSesi: data.length,
  };
}

function getBuku({ kategori }) {
  const sh   = SS.getSheetByName(SHEET.BUKU);
  const rows = sh.getDataRange().getValues().slice(1);
  let data   = rows.map(r => ({ id: r[0], judul: r[1], penulis: r[2], kategori: r[3], link: r[4], cover: r[5] }));
  if (kategori && kategori !== "Semua") data = data.filter(d => d.kategori === kategori);
  return { ok: true, data };
}

function tambahBuku({ judul, penulis, kategori, link, cover }) {
  if (!judul || !link) return { ok: false, msg: "Judul dan link wajib diisi." };
  const sh = SS.getSheetByName(SHEET.BUKU);
  const id = "BK" + Date.now();
  sh.appendRow([id, judul, penulis || "-", kategori || "Umum", link, cover || ""]);
  return { ok: true, msg: "Buku berhasil ditambahkan.", id };
}

function hapusBuku({ id }) {
  const sh   = SS.getSheetByName(SHEET.BUKU);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      sh.deleteRow(i + 1);
      return { ok: true, msg: "Buku dihapus." };
    }
  }
  return { ok: false, msg: "Buku tidak ditemukan." };
}

// ────────────────────────────────────────────────────────────
//  DATA SISWA
// ────────────────────────────────────────────────────────────
function getSiswaList() {
  const sh   = SS.getSheetByName(SHEET.SISWA);
  const rows = sh.getDataRange().getValues().slice(1);
  const data = rows.map(r => ({ nis: r[0], nama: r[1], kelas: r[2], role: r[4] }));
  return { ok: true, data };
}

function tambahSiswa({ nis, nama, kelas, password }) {
  if (!nis || !nama || !password) return { ok: false, msg: "Data tidak lengkap." };
  const sh   = SS.getSheetByName(SHEET.SISWA);
  const rows = sh.getDataRange().getValues().slice(1);
  if (rows.find(r => String(r[0]) === nis)) return { ok: false, msg: "NIS sudah terdaftar." };
  sh.appendRow([nis, nama, kelas || "-", password, "siswa"]);
  return { ok: true, msg: "Siswa berhasil didaftarkan." };
}
