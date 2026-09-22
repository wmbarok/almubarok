function getBooksData() {
  try {
    var dbId = CONFIG.DB.PUSTAKA;
    if (!dbId || dbId.includes("ID_") || dbId.includes("PASTE_")) {
      return getDemoBooks();
    }
    
    var ss = SpreadsheetApp.openById(dbId);
    var sheet = ss.getSheetByName("Ebooks");
    
    if (!sheet) {
      sheet = ss.insertSheet("Ebooks");
      sheet.appendRow(["ID_Buku", "Judul", "Penulis", "Penerbit", "Tahun_Terbit", "ID_Kategori", "Sinopsis", "Link_PDF", "Link_Cover", "Status_Akses", "Total_Views", "Total_Downloads", "Tanggal_Upload"]);
      sheet.getRange(1, 1, 1, 13).setFontWeight("bold").setBackground("#f0f0f0");
      sheet.setFrozenRows(1);
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return getDemoBooks();
    
    var books = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || String(data[i][0]).trim() === "") continue; 
      books.push({
        id: String(data[i][0] || ""),
        judul: String(data[i][1] || "Tanpa Judul"),
        penulis: String(data[i][2] || "-"),
        kategori: String(data[i][5] || "UMUM"),
        linkPdf: String(data[i][7] || ""),
        linkCover: String(data[i][8] || ""),
        status: String(data[i][9] || "Publik"),
        views: Number(data[i][10]) || 0,
        downloads: Number(data[i][11]) || 0
      });
    }
    return books.length > 0 ? books : getDemoBooks();
  } catch (e) {
    return getDemoBooks();
  }
}

function getDemoBooks() {
  return [
    { id: "BUK-001", judul: "Panduan Shalat DTA (Demo)", penulis: "Tim Yayasan", kategori: "DTA", linkPdf: "", linkCover: "", status: "Publik", views: 12, downloads: 3 },
    { id: "BUK-002", judul: "Mewarnai Huruf Hijaiyah", penulis: "Tim PAUD", kategori: "PAUD", linkPdf: "", linkCover: "", status: "Publik", views: 20, downloads: 7 }
  ];
}

function uploadBookData(form) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.DB.PUSTAKA);
    var sheet = ss.getSheetByName("Ebooks");
    if (!sheet) {
      sheet = ss.insertSheet("Ebooks");
      sheet.appendRow(["ID_Buku", "Judul", "Penulis", "Penerbit", "Tahun_Terbit", "ID_Kategori", "Sinopsis", "Link_PDF", "Link_Cover", "Status_Akses", "Total_Views", "Total_Downloads", "Tanggal_Upload"]);
    }

    var folderName = "Pustaka_AlMubarok_Files";
    var folders = DriveApp.getFoldersByName(folderName);
    var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var linkPdf = "";
    if (form.metodePdf === 'file' && form.pdfBytes) {
      var pdfBlob = Utilities.newBlob(form.pdfBytes, form.pdfMime, form.pdfName);
      var pdfFile = targetFolder.createFile(pdfBlob);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkPdf = pdfFile.getUrl();
    } else if (form.metodePdf === 'link') {
      linkPdf = form.linkPdfUrl || "";
    }

    var linkCover = "";
    if (form.coverBytes) {
      var coverBlob = Utilities.newBlob(form.coverBytes, form.coverMime, form.coverName);
      var coverFile = targetFolder.createFile(coverBlob);
      coverFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkCover = coverFile.getUrl();
    }

    var idBuku = "BUK-" + new Date().getTime();
    sheet.appendRow([
      idBuku, form.judul, form.penulis, form.penerbit || "-", form.tahun || "2026",
      form.kategori, form.sinopsis || "", linkPdf, linkCover, form.status || "Publik", 0, 0, new Date()
    ]);

    return { success: true, message: "E-Book berhasil ditambahkan ke sistem!" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteBook(idBuku) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.DB.PUSTAKA);
    var sheet = ss.getSheetByName("Ebooks");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idBuku)) {
        var linkPdf = data[i][7];
        var linkCover = data[i][8];
        try {
          if (linkPdf && linkPdf.includes("drive.google.com")) {
            var pdfId = linkPdf.match(/[-\w]{25,}/);
            if (pdfId) DriveApp.getFileById(pdfId[0]).setTrashed(true);
          }
          if (linkCover && linkCover.includes("drive.google.com")) {
            var coverId = linkCover.match(/[-\w]{25,}/);
            if (coverId) DriveApp.getFileById(coverId[0]).setTrashed(true);
          }
        } catch(err) {}
        
        sheet.deleteRow(i + 1);
        return { success: true, message: "E-book berhasil dihapus!" };
      }
    }
    return { success: false, message: "Buku tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
