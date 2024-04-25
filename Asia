function saveAsCSV() {
  try {
    // Mendapatkan spreadsheet aktif
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ganti 'Dashboard Asia-Pacific' dengan nama sheet Anda
    var sheet = spreadsheet.getSheetByName('Dashboard Asia-Pacific'); 
    
    // Memastikan sheet ditemukan
    if (!sheet) {
      throw new Error("Sheet 'Dashboard Asia-Pacific' tidak ditemukan.");
    }
    
    // Mengonversi range data menjadi file CSV
    var csvFile = convertRangeToCsvFile_(sheet);
    
    // Ganti '1QztNNaIC5WL40t6nlTHIRXtx8g9-sl-4?usp=drive_link' dengan ID folder tujuan di Google Drive
    var folderId = '1QztNNaIC5WL40t6nlTHIRXtx8g9-sl-4'; 
    
    // Simpan file CSV di Google Drive
    var folder = DriveApp.getFolderById(folderId);
    var fileName = sheet.getName() + ".csv";
    var file = folder.createFile(fileName, csvFile);
    
    // Menampilkan pesan berhasil
    Logger.log("File CSV berhasil disimpan di Google Drive dengan nama: " + fileName);
  } catch (error) {
    // Menampilkan pesan error
    Logger.log("Terjadi kesalahan: " + error);
  }
}

function convertRangeToCsvFile_(sheet) {
  var data = sheet.getDataRange().getValues();
  var csvFile = "";
  
  for (var row = 0; row < data.length; row++) {
    for (var col = 0; col < data[row].length; col++) {
      if (data[row][col].toString().indexOf(",") != -1) {
        data[row][col] = "\"" + data[row][col] + "\"";
      }
    }
    csvFile += data[row].join(",") + "\r\n";
  }
  
  return csvFile;
}
