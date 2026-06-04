/**
 * NABUNG EMAS - BACKEND ENGINE (Google Apps Script)
 * Author: top8
 * Description: Menangani database spreadsheet secara otomatis, inisialisasi sheet, 
 * dan menyediakan API CRUD aman dengan validasi PIN.
 * * Update: Mengganti DriveApp.searchFiles() dengan iterasi getFiles() yang 100% aman
 * untuk menghindari error 'Invalid argument: q' akibat pembatasan kueri API Google Drive.
 */

// Nama Spreadsheet utama yang digunakan untuk menyimpan seluruh database aplikasi
const SPREADSHEET_NAME = "Nabung Emas DB";

/**
 * Mendapatkan Spreadsheet aktif atau membuat baru jika belum ditemukan.
 * Hanya mencari berkas Google Sheets asli yang tidak berada di dalam Sampah (Trash).
 * @return {Spreadsheet} Spreadsheet dari Google Sheets
 */
function getOrCreateSpreadsheet() {
  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  let validSpreadsheet = null;
  
  while (files.hasNext()) {
    const file = files.next();
    // Memastikan file tidak berada di dalam sampah (trash) 
    // dan memiliki tipe file Google Spreadsheet asli
    if (!file.isTrashed() && file.getMimeType() === "application/vnd.google-apps.spreadsheet") {
      validSpreadsheet = file;
      break; // Temukan berkas aktif pertama yang cocok
    }
  }
  
  if (validSpreadsheet) {
    try {
      return SpreadsheetApp.openById(validSpreadsheet.getId());
    } catch(e) {
      // Jika terjadi kegagalan langka saat membuka, buat yang baru sebagai fallback aman
      const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
      initSheetsStructure(ss);
      return ss;
    }
  } else {
    // Membuat spreadsheet baru jika benar-benar tidak ditemukan di Drive aktif
    const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    initSheetsStructure(ss);
    return ss;
  }
}

/**
 * Inisialisasi struktur sheet dan data default untuk pertama kali dijalankan.
 * @param {Spreadsheet} ss Spreadsheet baru
 */
function initSheetsStructure(ss) {
  // 1. SHEET SETTINGS (Konfigurasi Parameter)
  let sheetSettings = ss.getSheetByName("Settings");
  if (!sheetSettings) {
    sheetSettings = ss.insertSheet("Settings");
    // Hapus sheet bawaan Google Sheets ("Sheet1") jika baru saja dibuat
    const defaultSheet = ss.getSheetByName("Sheet1");
    if (defaultSheet) {
      try { ss.deleteSheet(defaultSheet); } catch(e) {}
    }
  }
  sheetSettings.clear();
  sheetSettings.appendRow(["Key", "Value", "Keterangan"]);
  
  const defaultSettings = [
    ["hargaBeli", "2774000", "Harga jual toko per gram"],
    ["hargaBuyback", "2584000", "Harga buyback toko per gram"],
    ["pph22Percent", "0.45", "Pajak PPh 22 dalam %"],
    ["biayaMaterai", "10000", "Nominal biaya meterai"],
    ["batasMaterai", "10000000", "Batas transaksi wajib meterai"],
    ["pin", "1234", "PIN aplikasi (4 digit)"],
    ["theme", "dark", "Tema default (dark/light)"],
    ["profileLabel", "by top8", "Nama pemilik profil"],
    ["zakatNishab", "85.00", "Nishab zakat emas (gram)"],
    ["zakatRate", "2.5", "Kadar zakat emas pertahun (%)"],
    ["zakatPriceRef", "beli", "Acuan harga zakat (beli/buyback)"]
  ];
  sheetSettings.getRange(2, 1, defaultSettings.length, 3).setValues(defaultSettings);

  // 2. SHEET GOALS (Target Tabungan)
  let sheetGoals = ss.getSheetByName("Goals");
  if (!sheetGoals) {
    sheetGoals = ss.insertSheet("Goals");
  }
  sheetGoals.clear();
  sheetGoals.appendRow(["ID", "Nama Target", "Target Gram", "Warna Tema"]);
  const defaultGoals = [
    ["g1", "Haji", "40.00", "bg-blue-500"],
    ["g2", "Pendidikan Anak", "40.00", "bg-emerald-500"]
  ];
  sheetGoals.getRange(2, 1, defaultGoals.length, 4).setValues(defaultGoals);

  // 3. SHEET TRANSACTIONS (Riwayat Transaksi)
  let sheetTx = ss.getSheetByName("Transactions");
  if (!sheetTx) {
    sheetTx = ss.insertSheet("Transactions");
  }
  sheetTx.clear();
  sheetTx.appendRow(["ID", "Tanggal", "Tipe", "Gram", "Harga Satuan", "Total Rp", "Goal ID", "Catatan"]);
  const defaultTx = [
    ["t1", "2026-05-10", "Beli", "28.00", "1758130", "49227652", "g1", "Modal Haji"],
    ["t2", "2026-05-25", "Beli", "4.00", "2907500", "11630000", "g2", "Pendidikan Anak"]
  ];
  sheetTx.getRange(2, 1, defaultTx.length, 8).setValues(defaultTx);
}

/**
 * Endpoint utama Web App untuk merender UI HTML
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle("Nabung Emas - Premium Dashboard")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Mendapatkan seluruh Settings dalam bentuk JavaScript Map Key-Value
 * @return {object} Key-value settings
 */
function getSettingsMap() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const settings = {};
  data.forEach(row => {
    const key = row[0];
    const val = row[1];
    // Parsing otomatis tipe data angka / string
    if (!isNaN(val) && val.toString().trim() !== "") {
      settings[key] = parseFloat(val);
    } else {
      settings[key] = val;
    }
  });
  return settings;
}

/**
 * Validasi kebenaran PIN yang diinput pengguna dengan PIN tersimpan di spreadsheet.
 * @param {string} enteredPin PIN yang dikirim oleh pengguna
 * @return {boolean} valid / tidak valid
 */
function validatePin(enteredPin) {
  const settings = getSettingsMap();
  const serverPin = settings.pin ? settings.pin.toString().trim() : "1234";
  return enteredPin.toString().trim() === serverPin;
}

/**
 * API: Mengambil seluruh paket data untuk disajikan ke UI Frontend (Read)
 * @return {object} Gabungan Settings, Goals, Transactions, dan Tautan Spreadsheet
 */
function getAppData() {
  try {
    const ss = getOrCreateSpreadsheet();
    const settings = getSettingsMap();
    
    // 1. Ambil Data Goals
    const sheetGoals = ss.getSheetByName("Goals");
    let goals = [];
    if (sheetGoals.getLastRow() > 1) {
      const goalsData = sheetGoals.getRange(2, 1, sheetGoals.getLastRow() - 1, 4).getValues();
      goals = goalsData.map(row => ({
        id: row[0].toString(),
        nama: row[1].toString(),
        targetGram: parseFloat(row[2]) || 0,
        warna: row[3].toString()
      }));
    }

    // 2. Ambil Data Transactions
    const sheetTx = ss.getSheetByName("Transactions");
    let transactions = [];
    if (sheetTx.getLastRow() > 1) {
      const txData = sheetTx.getRange(2, 1, sheetTx.getLastRow() - 1, 8).getValues();
      transactions = txData.map(row => {
        let dateStr = "";
        if (row[1] instanceof Date) {
          dateStr = row[1].toISOString().split('T')[0];
        } else {
          dateStr = new Date(row[1]).toISOString().split('T')[0];
        }
        return {
          id: row[0].toString(),
          tanggal: dateStr,
          tipe: row[2].toString(),
          gram: parseFloat(row[3]) || 0,
          hargaSatuan: parseFloat(row[4]) || 0,
          totalRp: parseFloat(row[5]) || 0,
          goalId: row[6].toString(),
          catatan: row[7].toString()
        };
      });
    }

    return {
      settings: settings,
      goals: goals,
      transactions: transactions,
      spreadsheetUrl: ss.getUrl()
    };
  } catch (e) {
    throw new Error("Gagal memuat data dari Spreadsheet: " + e.message);
  }
}

/**
 * API: Menyimpan pengaturan baru (Update Settings)
 */
function saveSettings(newSettings, pinKeamanan) {
  if (!validatePin(pinKeamanan)) {
    return { success: false, message: "Otorisasi Gagal! PIN salah." };
  }
  
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Settings");
    const lastRow = sheet.getLastRow();
    
    const rangeKeys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let key in newSettings) {
      const value = newSettings[key];
      let found = false;
      
      for (let i = 0; i < rangeKeys.length; i++) {
        if (rangeKeys[i][0] === key) {
          sheet.getRange(i + 2, 2).setValue(value.toString());
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([key, value.toString(), "Parameter Tambahan"]);
      }
    }
    return { success: true, message: "Parameter konfigurasi berhasil diperbarui!" };
  } catch(e) {
    return { success: false, message: "Terjadi kesalahan: " + e.message };
  }
}

/**
 * API: Membuat Target Tabungan Baru (Create Goal)
 */
function addGoal(goal, pinKeamanan) {
  if (!validatePin(pinKeamanan)) {
    return { success: false, message: "Otorisasi Gagal! PIN salah." };
  }
  
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Goals");
    const newId = "g_" + Date.now();
    
    sheet.appendRow([
      newId,
      goal.nama,
      goal.targetGram,
      goal.warna
    ]);
    return { success: true, message: "Target tabungan baru berhasil ditambahkan!" };
  } catch(e) {
    return { success: false, message: "Terjadi kesalahan: " + e.message };
  }
}

/**
 * API: Memperbarui Target Tabungan (Update Goal)
 */
function updateGoal(goal, pinKeamanan) {
  if (!validatePin(pinKeamanan)) {
    return { success: false, message: "Otorisasi Gagal! PIN salah." };
  }
  
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Goals");
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString() === goal.id.toString()) {
        sheet.getRange(i + 2, 2, 1, 3).setValues([[
          goal.nama,
          goal.targetGram,
          goal.warna
        ]]);
        return { success: true, message: "Target tabungan berhasil diperbarui!" };
      }
    }
    return { success: false, message: "Target tidak ditemukan di database." };
  } catch(e) {
    return { success: false, message: "Terjadi kesalahan: " + e.message };
  }
}

/**
 * API: Menghapus Target Tabungan dan Mengalihkan Transaksinya (Delete Goal)
 */
function deleteGoal(id, pinKeamanan) {
  if (!validatePin(pinKeamanan)) {
    return { success: false, message: "Otorisasi Gagal! PIN salah." };
  }
  
  try {
    const ss = getOrCreateSpreadsheet();
    
    // 1. Hapus dari sheet Goals
    const sheetGoals = ss.getSheetByName("Goals");
    const goalsData = sheetGoals.getRange(2, 1, sheetGoals.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < goalsData.length; i++) {
      if (goalsData[i][0].toString() === id.toString()) {
        sheetGoals.deleteRow(i + 2);
        break;
      }
    }
    
    // 2. Alihkan semua transaksi terkait ke "Umum" (kosongkan field goalId)
    const sheetTx = ss.getSheetByName("Transactions");
    if (sheetTx.getLastRow() > 1) {
      const txData = sheetTx.getRange(2, 7, sheetTx.getLastRow() - 1, 1).getValues();
      for (let j = 0; j < txData.length; j++) {
        if (txData[j][0].toString() === id.toString()) {
          sheetTx.getRange(j + 2, 7).setValue("");
        }
      }
    }
    return { success: true, message: "Target berhasil dihapus. Transaksi dialihkan ke Umum." };
  } catch(e) {
    return { success: false, message: "Terjadi kesalahan: " + e.message };
  }
}

/**
 * API: Menambahkan Catatan Transaksi Baru (Create Transaction)
 */
function addTransaction(tx, pinKeamanan) {
  if (!validatePin(pinKeamanan)) {
    return { success: false, message: "Otorisasi Gagal! PIN salah." };
  }
  
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Transactions");
    const newId = "tx_" + Date.now();
    
    // Format Jual: Gramasi dan Nominal di-minus di database untuk operasi sum yang dinamis
    let finalGram = tx.gram;
    let finalTotal = tx.totalRp;
    if (tx.tipe === "Jual") {
      finalGram = -Math.abs(tx.gram);
      finalTotal = -Math.abs(tx.totalRp);
    }
    
    sheet.appendRow([
      newId,
      tx.tanggal,
      tx.tipe,
      finalGram,
      tx.hargaSatuan,
      finalTotal,
      tx.goalId,
      tx.catatan
    ]);
    return { success: true, message: "Catatan transaksi berhasil disimpan!" };
  } catch(e) {
    return { success: false, message: "Terjadi kesalahan: " + e.message };
  }
}

/**
 * API: Memperbarui Detail Transaksi (Update Transaction)
 */
function updateTransaction(tx, pinKeamanan) {
  if (!validatePin(pinKeamanan)) {
    return { success: false, message: "Otorisasi Gagal! PIN salah." };
  }
  
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Transactions");
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    
    let finalGram = tx.gram;
    let finalTotal = tx.totalRp;
    if (tx.tipe === "Jual") {
      finalGram = -Math.abs(tx.gram);
      finalTotal = -Math.abs(tx.totalRp);
    }

    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString() === tx.id.toString()) {
        sheet.getRange(i + 2, 2, 1, 7).setValues([[
          tx.tanggal,
          tx.tipe,
          finalGram,
          tx.hargaSatuan,
          finalTotal,
          tx.goalId,
          tx.catatan
        ]]);
        return { success: true, message: "Transaksi berhasil diperbarui!" };
      }
    }
    return { success: false, message: "Data transaksi tidak ditemukan." };
  } catch(e) {
    return { success: false, message: "Terjadi kesalahan: " + e.message };
  }
}

/**
 * API: Menghapus Catatan Transaksi (Delete Transaction)
 */
function deleteTransaction(id, pinKeamanan) {
  if (!validatePin(pinKeamanan)) {
    return { success: false, message: "Otorisasi Gagal! PIN salah." };
  }
  
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Transactions");
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        sheet.deleteRow(i + 2);
        return { success: true, message: "Transaksi berhasil dihapus dari database!" };
      }
    }
    return { success: false, message: "Transaksi tidak ditemukan." };
  } catch(e) {
    return { success: false, message: "Terjadi kesalahan: " + e.message };
  }
}

/**
 * API: Mengembalikan Aplikasi ke Kondisi Awal (Reset)
 */
function resetAllData() {
  try {
    const ss = getOrCreateSpreadsheet();
    initSheetsStructure(ss);
    return { success: true, message: "Seluruh database aplikasi berhasil direset!" };
  } catch(e) {
    return { success: false, message: "Terjadi kesalahan reset: " + e.message };
  }
}
