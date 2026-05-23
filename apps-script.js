// ─── GOOGLE APPS SCRIPT ─────────────────────────────────────────────────────
// Deploy ini sebagai Web App di Google Apps Script
// Execute as: Me | Who has access: Anyone
// Salin URL deploy ke dalam index.html (variabel APPS_SCRIPT_URL)

const SHEET_NAME = "Nilai Siswa";

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return json({ ok: false, error: "Sheet '" + SHEET_NAME + "' tidak ditemukan." });

    const headers = ["timestamp","nama","kelas","nilai_bab1","pred_bab1","catatan_bab1","nilai_bab2","pred_bab2","catatan_bab2","nilai_bab3","pred_bab3","catatan_bab3","nilai_bab4","pred_bab4","catatan_bab4","rata_rata","pred_akhir","status","guru_penilai"];

    if (action === "delete") {
      const rowIndex = parseInt(e.parameter.rowIndex);
      if (isNaN(rowIndex) || rowIndex < 2) return json({ ok: false, error: "rowIndex tidak valid: " + e.parameter.rowIndex });
      const lastRow = sheet.getLastRow();
      if (rowIndex > lastRow) return json({ ok: false, error: "Baris " + rowIndex + " tidak ada (lastRow=" + lastRow + ")" });
      sheet.deleteRow(rowIndex);
      return json({ ok: true, message: "Data berhasil dihapus!" });
    }

    if (action === "getAll") {
      const data = sheet.getDataRange().getValues();
      const rows = data.slice(1).map((r, i) => {
        const obj = {};
        headers.forEach((h, idx) => obj[h] = r[idx] || "");
        obj.rowIndex = i + 2;
        return obj;
      }).filter(r => r.nama !== "");
      return json({ ok: true, data: rows });
    }

    if (action === "getSiswa") {
      const nama = (e.parameter.nama || "").toLowerCase();
      const data = sheet.getDataRange().getValues();
      const rows = data.slice(1).map(r => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = r[i] || "");
        return obj;
      }).filter(r => r.nama.toLowerCase().includes(nama) && r.nama !== "");
      return json({ ok: true, data: rows });
    }

    return json({ ok: false, error: "Unknown action: " + action });
  } catch(err) {
    return json({ ok: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return json({ ok: false, error: "Sheet '" + SHEET_NAME + "' tidak ditemukan." });

    // DELETE: hapus baris
    if (payload.action === "delete") {
      const rowIndex = parseInt(payload.rowIndex);
      if (isNaN(rowIndex) || rowIndex < 2) return json({ ok: false, error: "rowIndex tidak valid: " + payload.rowIndex });
      const lastRow = sheet.getLastRow();
      if (rowIndex > lastRow) return json({ ok: false, error: "Baris " + rowIndex + " tidak ada (lastRow=" + lastRow + ")" });
      sheet.deleteRow(rowIndex);
      return json({ ok: true, message: "Data berhasil dihapus!" });
    }

    const vals = [payload.nilai_bab1, payload.nilai_bab2, payload.nilai_bab3, payload.nilai_bab4].filter(v => v !== "" && v !== undefined);
    const avg = vals.length ? Math.round(vals.reduce((a,b) => a + parseFloat(b), 0) / vals.length * 10) / 10 : "";
    const getP = v => v >= 90 ? "A" : v >= 75 ? "B" : v >= 60 ? "C" : "D";
    const pred = avg !== "" ? getP(avg) : "";
    const status = avg !== "" ? (avg >= 75 ? "Tuntas" : "Belum Tuntas") : "";

    // UPDATE: edit baris yang sudah ada
    if (payload.action === "update") {
      const rowIndex = parseInt(payload.rowIndex);
      if (isNaN(rowIndex) || rowIndex < 2) return json({ ok: false, error: "rowIndex tidak valid: " + payload.rowIndex });
      sheet.getRange(rowIndex, 2, 1, 18).setValues([[
        payload.nama, payload.kelas,
        payload.nilai_bab1, payload.pred_bab1, payload.catatan_bab1,
        payload.nilai_bab2, payload.pred_bab2, payload.catatan_bab2,
        payload.nilai_bab3, payload.pred_bab3, payload.catatan_bab3,
        payload.nilai_bab4, payload.pred_bab4, payload.catatan_bab4,
        avg, pred, status, payload.guru_penilai || ""
      ]]);
      return json({ ok: true, message: "Nilai berhasil diperbarui!" });
    }

    // INSERT: tambah data baru
    if (sheet.getLastRow() < 1) {
      sheet.appendRow(["Timestamp","Nama Siswa","Kelas","Nilai Bab 1","Predikat Bab 1","Catatan Bab 1","Nilai Bab 2","Predikat Bab 2","Catatan Bab 2","Nilai Bab 3","Predikat Bab 3","Catatan Bab 3","Nilai Bab 4","Predikat Bab 4","Catatan Bab 4","Rata-rata","Predikat Akhir","Status","Guru Penilai"]);
    }
    sheet.appendRow([
      new Date().toLocaleString("id-ID"),
      payload.nama, payload.kelas,
      payload.nilai_bab1, payload.pred_bab1, payload.catatan_bab1,
      payload.nilai_bab2, payload.pred_bab2, payload.catatan_bab2,
      payload.nilai_bab3, payload.pred_bab3, payload.catatan_bab3,
      payload.nilai_bab4, payload.pred_bab4, payload.catatan_bab4,
      avg, pred, status, payload.guru_penilai || ""
    ]);
    return json({ ok: true, message: "Nilai berhasil disimpan!" });

  } catch(err) {
    return json({ ok: false, error: err.toString() });
  }
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.getRange(1,1,1,19).setValues([["Timestamp","Nama Siswa","Kelas","Nilai Bab 1","Predikat Bab 1","Catatan Bab 1","Nilai Bab 2","Predikat Bab 2","Catatan Bab 2","Nilai Bab 3","Predikat Bab 3","Catatan Bab 3","Nilai Bab 4","Predikat Bab 4","Catatan Bab 4","Rata-rata","Predikat Akhir","Status","Guru Penilai"]]);
  sheet.getRange(1,1,1,19).setFontWeight("bold").setBackground("#1565C0").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);
}
