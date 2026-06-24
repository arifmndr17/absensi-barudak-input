function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // 1. Dekode data teks menjadi file gambar asli (.jpg)
    var contentType = data.foto.substring(data.foto.indexOf(":")+1, data.foto.indexOf(";"));
    var base64Data = data.foto.substring(data.foto.indexOf(",")+1);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, "Selfie_" + data.nama + "_" + new Date().getTime() + ".jpg");
    
    // ==========================================
    // 2. SIMPAN KE FOLDER DRIVE SPESIFIK
    // ==========================================
    var folderId = "1O-d1F-9G3LZyn-nFgAgeaSD1F8nzLB-s"; // ID Folder Anda
    var folder = DriveApp.getFolderById(folderId);
    var file = folder.createFile(blob); // Simpan foto langsung di dalam folder ini
    // ==========================================
    
    // 3. Atur izin akses agar foto tersebut bisa dirender atau dilihat lewat link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 4. Mengubah link menjadi Direct Link Gambar agar bisa dibaca rumus IMAGE
    var fileId = file.getId();
    var directImageUrl = "https://lh3.googleusercontent.com/d/" + fileId;
    
    // 5. Masukkan data ke baris Sheet dengan rumus IMAGE
    sheet.appendRow([
      new Date(),
      data.nama,
      data.tipe,
      '=IMAGE("' + directImageUrl + '")'
    ]);
    
    // 6. Atur tinggi baris secara otomatis agar gambar terlihat besar
    var lastRow = sheet.getLastRow();
    sheet.setRowHeight(lastRow, 90); 
    sheet.setColumnWidth(4, 120);
    
    return ContentService.createTextOutput(JSON.stringify({"status": "SUCCESS"}))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "ERROR", "message": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function doOptions(e) {
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
}
