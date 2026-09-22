function getPostsData() {
  try {
    var dbId = CONFIG.DB.ACARA;
    if (!dbId || dbId.includes("ID_") || dbId.includes("PASTE_")) {
      return getDemoPosts();
    }
    
    var ss = SpreadsheetApp.openById(dbId);
    var sheet = ss.getSheetByName("Postingan");
    
    if (!sheet) {
      sheet = ss.insertSheet("Postingan");
      sheet.appendRow(["ID_Post", "Judul", "Kategori", "Isi", "Penulis", "Tanggal_Upload", "Link_Gambar", "Status_Publish"]);
      sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#f0f0f0");
      sheet.setFrozenRows(1);
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return getDemoPosts();
    
    var posts = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || String(data[i][0]).trim() === "") continue;
      posts.push({
        id: String(data[i][0] || ""),
        judul: String(data[i][1] || ""),
        kategori: String(data[i][2] || "Umum"),
        isi: String(data[i][3] || ""),
        penulis: String(data[i][4] || "Admin"),
        tanggal: String(data[i][5] || ""),
        linkGambar: String(data[i][6] || "")
      });
    }
    return posts.length > 0 ? posts : getDemoPosts();
  } catch (e) {
    return getDemoPosts();
  }
}

function getDemoPosts() {
  return [
    { 
      id: "POST-001", 
      judul: "Kegiatan Pesantren Kilat Ramadhan", 
      kategori: "Kegiatan", 
      isi: "Alhamdulillah kegiatan pesantren kilat santri DTA Al-Mubarok berjalan dengan lancar dan khidmat.", 
      penulis: "Admin Yayasan", 
      tanggal: "2026-03-25", 
      linkGambar: "" 
    }
  ];
}

function uploadPostData(form) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.DB.ACARA);
    var sheet = ss.getSheetByName("Postingan");
    if (!sheet) {
      sheet = ss.insertSheet("Postingan");
      sheet.appendRow(["ID_Post", "Judul", "Kategori", "Isi", "Penulis", "Tanggal_Upload", "Link_Gambar", "Status_Publish"]);
    }

    var folderName = "Acara_AlMubarok_Files";
    var folders = DriveApp.getFoldersByName(folderName);
    var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var linkGambar = "";
    if (form.fotoBytes) {
      var fotoBlob = Utilities.newBlob(form.fotoBytes, form.fotoMime, form.fotoName);
      var fotoFile = targetFolder.createFile(fotoBlob);
      fotoFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkGambar = fotoFile.getUrl();
    }

    var idPost = "POST-" + new Date().getTime();
    var tanggalStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    sheet.appendRow([
      idPost, form.judul, form.kategori, form.isi, form.penulis || "Admin",
      tanggalStr, linkGambar, "Publish"
    ]);

    return { success: true, message: "Postingan berhasil diterbitkan!" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deletePost(idPost) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.DB.ACARA);
    var sheet = ss.getSheetByName("Postingan");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idPost)) {
        var linkGambar = data[i][6];
        try {
          if (linkGambar) {
            var fotoId = linkGambar.match(/[-\w]{25,}/);
            if (fotoId) DriveApp.getFileById(fotoId[0]).setTrashed(true);
          }
        } catch(err) {}
        
        sheet.deleteRow(i + 1);
        return { success: true, message: "Postingan berhasil dihapus!" };
      }
    }
    return { success: false, message: "Postingan tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
