function getGalleryData() {
  try {
    var dbId = CONFIG.DB.ACARA;
    if (!dbId || dbId.includes("ID_") || dbId.includes("PASTE_")) {
      return getDemoGallery();
    }
    
    var ss = SpreadsheetApp.openById(dbId);
    var sheet = ss.getSheetByName("Galeri");
    
    if (!sheet) {
      sheet = ss.insertSheet("Galeri");
      sheet.appendRow(["ID_Foto", "Judul", "Kategori", "Link_Foto", "Tanggal_Upload"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f0f0f0");
      sheet.setFrozenRows(1);
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return getDemoGallery();
    
    var gallery = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || String(data[i][0]).trim() === "") continue;
      gallery.push({
        id: String(data[i][0] || ""),
        judul: String(data[i][1] || "Dokumentasi"),
        kategori: String(data[i][2] || "PAUD"),
        linkFoto: String(data[i][3] || ""),
        tanggal: String(data[i][4] || "")
      });
    }
    return gallery.length > 0 ? gallery : getDemoGallery();
  } catch (e) {
    return getDemoGallery();
  }
}

function getDemoGallery() {
  return [
    { id: "GAL-001", judul: "Kegiatan Belajar PAUD Al-Mubarok", kategori: "PAUD", linkFoto: "", tanggal: "2026-09-20" },
    { id: "GAL-002", judul: "Praktek Shalat Santri DTA", kategori: "Santri", linkFoto: "", tanggal: "2026-09-21" }
  ];
}

function uploadGalleryData(form) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.DB.ACARA);
    var sheet = ss.getSheetByName("Galeri");
    if (!sheet) {
      sheet = ss.insertSheet("Galeri");
      sheet.appendRow(["ID_Foto", "Judul", "Kategori", "Link_Foto", "Tanggal_Upload"]);
    }

    var folderName = "Galeri_AlMubarok_Files";
    var folders = DriveApp.getFoldersByName(folderName);
    var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var linkFoto = "";
    if (form.fotoBytes) {
      var fotoBlob = Utilities.newBlob(form.fotoBytes, form.fotoMime, form.fotoName);
      var fotoFile = targetFolder.createFile(fotoBlob);
      fotoFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkFoto = fotoFile.getUrl();
    }

    var idFoto = "GAL-" + new Date().getTime();
    var tanggalStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    sheet.appendRow([
      idFoto, form.judul, form.kategori, linkFoto, tanggalStr
    ]);

    return { success: true, message: "Foto galeri berhasil diunggah!" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteGalleryItem(idFoto) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.DB.ACARA);
    var sheet = ss.getSheetByName("Galeri");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idFoto)) {
        var linkFoto = data[i][3];
        try {
          if (linkFoto) {
            var fotoId = linkFoto.match(/[-\w]{25,}/);
            if (fotoId) DriveApp.getFileById(fotoId[0]).setTrashed(true);
          }
        } catch(err) {}
        
        sheet.deleteRow(i + 1);
        return { success: true, message: "Foto galeri berhasil dihapus!" };
      }
    }
    return { success: false, message: "Foto tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
