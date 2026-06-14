// ID Spreadsheet: 1baTwpyf77FE8oGLYPcSJogvDCXbU-etrIPlRtI2HRuE
// Nama Sheet: Sheet1

const SPREADSHEET_ID = "1baTwpyf77FE8oGLYPcSJogvDCXbU-etrIPlRtI2HRuE";
const SHEET_NAME = "Sheet1";

function doGet(e) {
  // Helper function to return JSON response
  const output = (content) => {
    return ContentService.createTextOutput(JSON.stringify(content))
      .setMimeType(ContentService.MimeType.JSON);
  };
  
  try {
    const action = e.parameter.action || "read";
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    // If sheet does not exist, create it and set headers
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["timestamp", "nama tamu", "ucapan", "konfirmasi kehadiran", "jumlah tamu"]);
    }
    
    // Auto-repair headers if first row is completely empty
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue().toString().trim() === "") {
      sheet.getRange(1, 1, 1, 5).setValues([["timestamp", "nama tamu", "ucapan", "konfirmasi kehadiran", "jumlah tamu"]]);
    }
    
    if (action === "add") {
      const timestamp = new Date();
      const namaTamu = e.parameter.nama_tamu;
      const ucapan = e.parameter.ucapan;
      const kehadiran = e.parameter.kehadiran;
      const jumlahTamu = e.parameter.jumlah_tamu || 1;
      
      sheet.appendRow([timestamp, namaTamu, ucapan, kehadiran, jumlahTamu]);
      
      return output({ status: "success", message: "Ucapan berhasil dikirim!" });
    } else {
      // Read wishes
      const data = sheet.getDataRange().getValues();
      
      // If no data is present at all
      if (data.length === 0 || (data.length === 1 && data[0].join("").trim() === "")) {
        return output({ status: "success", data: [] });
      }
      
      // Detect if row 1 contains actual data instead of headers (e.g. is a date, a number, or empty)
      let hasHeaders = true;
      const firstVal = data[0][0];
      if (firstVal instanceof Date || (firstVal !== "" && !isNaN(Date.parse(firstVal))) || firstVal === "") {
        hasHeaders = false;
      }
      
      const headers = hasHeaders ? data[0] : ["timestamp", "nama tamu", "ucapan", "konfirmasi kehadiran", "jumlah tamu"];
      const startIndex = hasHeaders ? 1 : 0;
      const rows = [];
      
      for (let i = startIndex; i < data.length; i++) {
        // Skip completely empty rows
        if (data[i].join("").trim() === "") continue;
        
        const row = {};
        headers.forEach((header, index) => {
          // Normalize header key names for JSON, fallback to column index mappings
          const key = (header || "").toString().toLowerCase().trim()
            .replace("nama tamu", "nama_tamu")
            .replace("konfirmasi kehadiran", "kehadiran")
            .replace("jumlah tamu", "jumlah_tamu")
            || ["timestamp", "nama_tamu", "ucapan", "kehadiran", "jumlah_tamu"][index]
            || "column_" + index;
          
          row[key] = data[i][index];
        });
        rows.push(row);
      }
      
      // Sort: latest submissions first
      rows.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
      
      return output({ status: "success", data: rows });
    }
  } catch (error) {
    return output({ status: "error", message: error.toString() });
  }
}
