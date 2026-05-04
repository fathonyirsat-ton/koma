/**
 * BACKEND GOOGLE APPS SCRIPT (Code.gs)
 * Fungsi: Sinkronisasi Penuh Database Google Sheets & Upload Google Drive
 */

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ==========================================
  // 1. SETUP SHEET REGISTRASI
  // ==========================================
  var sheetReg = ss.getSheetByName("Registrasi");
  if (!sheetReg) {
    sheetReg = ss.insertSheet("Registrasi");
    var headers = [
      "Timestamp",              // A (0)
      "Nama Lengkap",           // B (1)
      "NISN",                   // C (2)
      "Sekolah Asal",           // D (3)
      "Tingkat Kelas",          // E (4) -> INI ADALAH KOLOM BARU KITA
      "No HP (WhatsApp)",       // F (5)
      "Email",                  // G (6)
      "Password",               // H (7)
      "Bukti Pembayaran (Link)",// I (8)
      "Status",                 // J (9)
      "Role",                   // K (10)
      "Last Login",             // L (11)
      "Last Logout"             // M (12)
    ];
    sheetReg.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3");
    sheetReg.getRange("C:C").setNumberFormat("@");
    sheetReg.getRange("F:F").setNumberFormat("@");
    sheetReg.setFrozenRows(1);
    
    var now = new Date();
    sheetReg.appendRow([now, "Super Admin", "'0000000000", "MGMP Tuban", "-", "'085649842450", "admin@mgmpmatematika.id", "123456", "-", "Active", "Admin", "-", "-"]);
    sheetReg.appendRow([now, "Panitia Verifikasi", "'1111111111", "MGMP Tuban", "-", "'081234567890", "panitia@mgmpmatematika.id", "123456", "-", "Active", "Panitia", "-", "-"]);
    sheetReg.appendRow([now, "Siswa Teladan", "'1234567890", "SMPN 1 Tuban", "VIII", "'089876543210", "siswa@mgmpmatematika.id", "123456", "-", "Active", "Siswa", "-", "-"]);
  }

  // ==========================================
  // 2. SETUP SHEET JADWAL
  // ==========================================
  var sheetJadwal = ss.getSheetByName("Jadwal");
  if (!sheetJadwal) {
    sheetJadwal = ss.insertSheet("Jadwal");
    var headersJadwal = ["ID", "Nama Kegiatan", "Tanggal Mulai", "Tanggal Selesai", "Info Update"];
    sheetJadwal.getRange(1, 1, 1, headersJadwal.length).setValues([headersJadwal]).setFontWeight("bold").setBackground("#cfe2f3");
    
    var defaultTimeline = [
      ["0", "Pendaftaran Peserta", "2026-04-20", "2026-05-08", "Pendaftaran dibuka."],
      ["1", "Technical Meeting", "2026-05-11", "2026-05-11", "Penjelasan teknis lomba."],
      ["2", "Simulasi Uji Coba", "2026-05-13", "2026-05-13", "Uji coba login aplikasi."],
      ["3", "Hari Pelaksanaan", "2026-05-15", "2026-05-15", "Pelaksanaan KoMa 2026."]
    ];
    for (var i = 0; i < defaultTimeline.length; i++) { sheetJadwal.appendRow(defaultTimeline[i]); }
    sheetJadwal.setFrozenRows(1);
  }

  // ==========================================
  // 3. SETUP SHEET LINK APLIKASI
  // ==========================================
  var sheetLink = ss.getSheetByName("Link");
  if (!sheetLink) {
    sheetLink = ss.insertSheet("Link");
    var headersLink = ["Platform", "URL Aplikasi"];
    sheetLink.getRange(1, 1, 1, headersLink.length).setValues([headersLink]).setFontWeight("bold").setBackground("#fff2cc");
    
    sheetLink.appendRow(["Android", "https://drive.google.com/file/d/1QglDIUlIeUY144aaSafsTQ7RWEZzBpNs/view?usp=sharing"]);
    sheetLink.appendRow(["Windows SEB", "https://drive.google.com/file/d/1cC2FtMphb6Igh7M9cc-x-1ebF6d0LcsV/view?usp=sharing"]);
    sheetLink.appendRow(["Mac OS SEB", "https://drive.google.com/file/d/1cho_Ji8tfqLwbB_IFh7NsMLgOARsLgJ3/view?usp=sharing"]);
    sheetLink.appendRow(["iOS SEB", "https://apps.apple.com/us/app/safeexambrowser/id1155002964"]);
    sheetLink.appendRow(["Config SEB", "https://drive.google.com/drive/folders/1kRL_oX86Yw2HYQB4uRDQEV559tDAJlfF"]);
    sheetLink.setFrozenRows(1);
  }
  
  // ==========================================
  // 4. SETUP FOLDER DRIVE
  // ==========================================
  var folderName = "BUKTI_BAYAR_KOMA_2026";
  var folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) { DriveApp.createFolder(folderName); }
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetReg = ss.getSheetByName("Registrasi");
  
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    // --- REGISTER ---
    if (action === 'register') {
      var emailLC = data.email.trim().toLowerCase();
      var exists = sheetReg.getRange("G:G").createTextFinder(emailLC).matchEntireCell(true).findNext();
      if (exists) return createJsonResponse({"result": "error", "message": "Email sudah terdaftar."});
      
      var fileUrl = "Gagal Muat Naik";
      if (data.photoBase64) {
        try {
          var folderName = "BUKTI_BAYAR_KOMA_2026";
          var folders = DriveApp.getFoldersByName(folderName);
          var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
          var fileName = "BUKTI_" + data.nisn + "_" + data.name.toUpperCase().replace(/\s+/g, '_') + ".jpg";
          var decodedFile = Utilities.base64Decode(data.photoBase64);
          var blob = Utilities.newBlob(decodedFile, "image/jpeg", fileName);
          var file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          fileUrl = file.getUrl();
        } catch (err) {}
      }

      var now = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
      // Urutan baru termasuk kelas: Waktu, Nama, NISN, Sekolah, KELAS, WA, Email, Password, Foto, Status, Role, Login, Logout
      sheetReg.appendRow([now, data.name, "'" + data.nisn, data.sekolah, data.kelas, "'" + data.wa, emailLC, data.password, fileUrl, "Menunggu Konfirmasi", "Siswa", now, "-"]);
      return createJsonResponse({"result": "success", "url": fileUrl});
    }

    // --- LOGIN ---
    if (action === 'login') {
      var vals = sheetReg.getDataRange().getDisplayValues();
      var emailLogin = data.email.trim().toLowerCase();
      var passLogin = data.password.toString();
      for (var i = 1; i < vals.length; i++) {
        // Index 6 adalah Email, Index 7 adalah Password
        if (vals[i][6].toLowerCase() === emailLogin && vals[i][7] === passLogin) {
          sheetReg.getRange(i + 1, 12).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
          return createJsonResponse({"result": "success", "user": { 
            name: vals[i][1], 
            nisn: vals[i][2], 
            sekolah: vals[i][3], 
            kelas: vals[i][4], // Kolom Kelas
            email: vals[i][6], 
            password: vals[i][7], 
            status: vals[i][9], 
            role: vals[i][10], 
            rowNum: i + 1 
          }});
        }
      }
      return createJsonResponse({"result": "error", "message": "Email atau Password salah."});
    }

    // --- EDIT PROFILE ---
    if (action === 'updateProfile') {
      var emailLC = data.email.trim().toLowerCase();
      var finder = sheetReg.getRange("G:G").createTextFinder(emailLC).matchEntireCell(true).findNext();
      
      if (!finder) return createJsonResponse({"result": "error", "message": "Akaun tidak dijumpai."});
      var row = finder.getRow();
      
      // Update info
      sheetReg.getRange(row, 2).setValue(data.name);
      sheetReg.getRange(row, 3).setValue("'" + data.nisn);
      sheetReg.getRange(row, 4).setValue(data.sekolah);
      sheetReg.getRange(row, 5).setValue(data.kelas); // Update Kelas
      sheetReg.getRange(row, 8).setValue(data.password);
      
      var newFileUrl = null;
      if (data.photoBase64 && data.photoBase64 !== "") {
        try {
          var folderName = "BUKTI_BAYAR_KOMA_2026";
          var folders = DriveApp.getFoldersByName(folderName);
          var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
          var fileName = "UPDATE_BUKTI_" + data.nisn + "_" + data.name.toUpperCase().replace(/\s+/g, '_') + ".jpg";
          var decodedFile = Utilities.base64Decode(data.photoBase64);
          var blob = Utilities.newBlob(decodedFile, "image/jpeg", fileName);
          var file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          newFileUrl = file.getUrl();
          
          sheetReg.getRange(row, 9).setValue(newFileUrl); // Kolom Bukti
        } catch (err) {}
      }

      var updatedVals = sheetReg.getRange(row, 1, 1, 13).getDisplayValues()[0];
      var updatedUser = { 
        name: updatedVals[1], nisn: updatedVals[2], sekolah: updatedVals[3], kelas: updatedVals[4],
        email: updatedVals[6], password: updatedVals[7], status: updatedVals[9], 
        role: updatedVals[10], rowNum: row 
      };
      
      return createJsonResponse({
        "result": "success", 
        "user": updatedUser, 
        "photoUrl": newFileUrl || updatedVals[8]
      });
    }

    // --- AMBIL DATA ---
    if (action === 'getAllData') {
      var rows = sheetReg.getDataRange().getDisplayValues();
      var list = [];
      for (var i = 1; i < rows.length; i++) {
        // Mengambil Kelas dari Index ke-4
        list.push({ "Nama Lengkap": rows[i][1], "NISN": rows[i][2], "Sekolah Asal": rows[i][3], "Tingkat Kelas": rows[i][4], "Email": rows[i][6], "Status": rows[i][9], "Role": rows[i][10], "rowNum": i + 1 });
      }
      return createJsonResponse({"result": "success", "data": list});
    }

    if (action === 'getBuktiBayar') {
      var rowData = sheetReg.getRange(data.targetRow, 1, 1, 13).getDisplayValues()[0];
      return createJsonResponse({"result": "success", "photoUrl": rowData[8], "lastLogin": rowData[11], "lastLogout": rowData[12]});
    }

    if (action === 'updateStatus') {
      sheetReg.getRange(data.targetRow, 10).setValue(data.newStatus); // Kolom J (10)
      return createJsonResponse({"result": "success"});
    }

    if (action === 'logout_record') {
      var finder = sheetReg.getRange("G:G").createTextFinder(data.email).matchEntireCell(true).findNext();
      if (finder) sheetReg.getRange(finder.getRow(), 13).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      return createJsonResponse({"result": "success"});
    }

    // --- TIMELINE JADWAL ---
    if (action === 'getSchedule') {
      var sheetJ = ss.getSheetByName("Jadwal");
      var rJ = sheetJ.getDataRange().getDisplayValues();
      var s = [];
      for (var j = 1; j < rJ.length; j++) { s.push({ id: rJ[j][0], name: rJ[j][1], start: rJ[j][2], end: rJ[j][3], info: rJ[j][4] }); }
      return createJsonResponse({"result": "success", "data": s});
    }

    if (action === 'updateSchedule') {
      var sheetJadwal = ss.getSheetByName("Jadwal");
      data.updates.forEach(function(it) {
        var rIdx = parseInt(it.id) + 2;
        sheetJadwal.getRange(rIdx, 2, 1, 4).setValues([[it.name, it.start, it.end, it.info]]);
      });
      return createJsonResponse({"result": "success"});
    }

    // --- LINK APLIKASI ---
    if (action === 'getAppLinks') {
      var sheetLink = ss.getSheetByName("Link");
      var rL = sheetLink.getDataRange().getDisplayValues();
      var links = [];
      for (var i = 1; i < rL.length; i++) {
        links.push({ platform: rL[i][0], url: rL[i][1] });
      }
      return createJsonResponse({"result": "success", "data": links});
    }

    if (action === 'updateAppLinks') {
      var sheetLink = ss.getSheetByName("Link");
      var rL = sheetLink.getDataRange().getValues();
      
      data.updates.forEach(function(item) {
        var found = false;
        for (var i = 1; i < rL.length; i++) {
          if (rL[i][0] === item.platform) {
            sheetLink.getRange(i + 1, 2).setValue(item.url);
            found = true;
            break;
          }
        }
        if (!found) {
          sheetLink.appendRow([item.platform, item.url]);
        }
      });
      return createJsonResponse({"result": "success"});
    }

  } catch (err) {
    return createJsonResponse({"result": "error", "message": err.toString()});
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
