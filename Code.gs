function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Rekod Latihan PKDKS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// FUNGSI 1: SIMPAN REKOD BARU ATAU KEMASKINI
// ==========================================
function simpanRekodTab_(dataFormStr, fileName, mimeType, base64Data, sheet, targetRow) {
  try {
    var dataForm = JSON.parse(dataFormStr);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().toLowerCase().trim(); });

    function cariKolom(keywords) {
      for (var i = 0; i < headers.length; i++) {
        for (var k = 0; k < keywords.length; k++) {
          if (headers[i].indexOf(keywords[k].toLowerCase()) > -1) return i + 1;
        }
      } return -1;
    }

    var colTime = cariKolom(["timestamp", "masa"]);
    var colNama = cariKolom(["nama penuh", "nama pegawai"]);
    var colIC = cariKolom(["kad pengenalan", "no. kp"]); 
    var colJawatan = cariKolom(["perjawatan", "jawatan"]);
    var colFasiliti = cariKolom(["tempat bertugas", "jabatan"]);
    var colKursus = cariKolom(["nama kursus", "tajuk kursus"]);
    var colTarikhMula = cariKolom(["tarikh mula", "tarikh kursus", "tarikh"]); 
    var colTarikhTamat = cariKolom(["tarikh tamat", "tarikh akhir"]); 
    var colTempat = cariKolom(["lokasi", "tempat kursus"]); 
    var colAnjuran = cariKolom(["anjur"]);
    var colKos = cariKolom(["kos"]);
    var colJam = cariKolom(["bilangan jam", "jam"]);
    var colCPD = cariKolom(["cpd"]);
    var colHRMIS = cariKolom(["telah diisi dalam hrmis", "hrmis"]); 
    var colSijil = cariKolom(["sijil", "bukti kehadiran"]);
    
    var colUmur = cariKolom(["umur"]);
    var colJantina = cariKolom(["jantina"]);
    var colGred = cariKolom(["skim", "gred"]);
    var colTempohAwam = cariKolom(["sektor awam"]);
    var colTempohJawatan = cariKolom(["di jawatan sekarang", "tempoh berkhidmat di jawatan"]);
    var colTempohKursus = cariKolom(["tempoh kursus", "hari latihan"]);
    var colJenisKursus = cariKolom(["jenis kursus"]);
    var colKategoriKursus = cariKolom(["kategori kursus"]);
    var colDalaman = cariKolom(["dalaman pkd", "anjuran dalaman"]);
    var colTahap1Start = cariKolom(["objektif kursus tercapai"]); 

    var isEdit = dataForm.editRow && dataForm.editRow !== "";
    var resetStatusKeberkesanan = false;
    if (isEdit && dataForm.tahap1 && colTahap1Start > 0) {
      var skorAsal = sheet.getRange(targetRow, colTahap1Start).getValue();
      resetStatusKeberkesanan = skorAsal === "" || skorAsal == null;
    }

    var fileUrl = "";
    if (base64Data && base64Data !== "") {
      if (colSijil < 1) throw new Error('Lajur sijil tidak dijumpai. Fail belum dimuat naik.');
      var folder = folderSijilTahun_(dataForm.tarikhMula);
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, dataForm.ic + "_" + fileName);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
      if (colSijil > 0) sheet.getRange(targetRow, colSijil).setValue(fileUrl);
    }

    if (colTime > 0 && !isEdit) sheet.getRange(targetRow, colTime).setValue(new Date());
    if (colNama > 0) sheet.getRange(targetRow, colNama).setValue(dataForm.nama.toUpperCase());
    if (colIC > 0) sheet.getRange(targetRow, colIC).setValue(dataForm.ic);
    if (colUmur > 0) sheet.getRange(targetRow, colUmur).setValue(dataForm.umur);
    if (colJantina > 0) sheet.getRange(targetRow, colJantina).setValue(dataForm.jantina);
    if (colJawatan > 0) sheet.getRange(targetRow, colJawatan).setValue(dataForm.jawatan);
    if (colGred > 0) sheet.getRange(targetRow, colGred).setValue(dataForm.gred);
    if (colFasiliti > 0) sheet.getRange(targetRow, colFasiliti).setValue(dataForm.fasiliti);
    if (colTempohAwam > 0) sheet.getRange(targetRow, colTempohAwam).setValue(dataForm.tempohAwam);
    if (colTempohJawatan > 0) sheet.getRange(targetRow, colTempohJawatan).setValue(dataForm.tempohJawatan);
    if (colKursus > 0) sheet.getRange(targetRow, colKursus).setValue(dataForm.kursus);
    if (colTarikhMula > 0) sheet.getRange(targetRow, colTarikhMula).setValue(dataForm.tarikhMula);
    if (colTarikhTamat > 0) sheet.getRange(targetRow, colTarikhTamat).setValue(dataForm.tarikhTamat);
    if (colTempohKursus > 0) sheet.getRange(targetRow, colTempohKursus).setValue(dataForm.tempohKursus);
    if (colTempat > 0) sheet.getRange(targetRow, colTempat).setValue(dataForm.tempat);
    if (colJenisKursus > 0) sheet.getRange(targetRow, colJenisKursus).setValue(dataForm.jenisKursus);
    if (colKategoriKursus > 0) sheet.getRange(targetRow, colKategoriKursus).setValue(dataForm.kategoriKursus);
    if (colAnjuran > 0) sheet.getRange(targetRow, colAnjuran).setValue(dataForm.anjuran);
    if (colKos > 0) sheet.getRange(targetRow, colKos).setValue(dataForm.kos);
    if (colJam > 0) sheet.getRange(targetRow, colJam).setValue(dataForm.jam);
    if (colCPD > 0) sheet.getRange(targetRow, colCPD).setValue(dataForm.cpd);
    if (colHRMIS > 0) sheet.getRange(targetRow, colHRMIS).setValue(dataForm.hrmis); 
    if (colDalaman > 0) sheet.getRange(targetRow, colDalaman).setValue(dataForm.anjuran.toUpperCase().indexOf("DALAMAN") > -1 || dataForm.anjuran.toUpperCase().indexOf("PKD") > -1 ? "YA" : "TIDAK");

    if (colTahap1Start > 0) {
      sheet.getRange(targetRow, colTahap1Start, 1, 25).clearContent(); 
      if (dataForm.tahap1) {
        var t1Array = [];
        for (var i = 1; i <= 25; i++) {
          t1Array.push(dataForm.tahap1["S" + i] || "");
        }
        sheet.getRange(targetRow, colTahap1Start, 1, 25).setValues([t1Array]);
      }
    }

    if (resetStatusKeberkesanan) {
      var colStatusKeberkesanan = cariKolom(["status borang"]);
      if (colStatusKeberkesanan > 0) sheet.getRange(targetRow, colStatusKeberkesanan).setValue("Belum");
    }

    return { success: true, message: isEdit ? "Rekod Latihan berjaya dikemaskini!" : "Rekod Latihan berjaya disimpan!", sijilUrl: fileUrl || (colSijil>0 ? String(sheet.getRange(targetRow,colSijil).getValue() || "") : "") };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// FUNGSI TARIK DATA UNTUK EDIT
// ==========================================
function rekodEditTab_(row, ic, sheet) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    function getCol(keywords) {
      for (var i = 0; i < headers.length; i++) {
        for (var k = 0; k < keywords.length; k++) {
          if (headers[i].indexOf(keywords[k].toLowerCase()) > -1) return i; 
        }
      } return -1;
    }

    var colIC = getCol(["kad pengenalan", "no. kp"]); 
    if (colIC === -1) return { success: false, message: "Ralat Sistem: Kolum IC tidak dijumpai." };

    var rekodIC = rowData[colIC] ? rowData[colIC].toString().trim() : "";
    if (icLatihan_(rekodIC) !== icLatihan_(ic)) {
      return { success: false, message: "No. Kad Pengenalan TIDAK sepadan. Akses ditolak." };
    }

    var data = {
      nama: rowData[getCol(["nama penuh", "nama pegawai"])] || "",
      ic: rekodIC,
      umur: rowData[getCol(["umur"])] || "",
      jantina: rowData[getCol(["jantina"])] || "",
      jawatan: rowData[getCol(["perjawatan", "jawatan"])] || "",
      gred: rowData[getCol(["skim", "gred"])] || "",
      fasiliti: rowData[getCol(["tempat bertugas", "jabatan"])] || "",
      tempohAwam: rowData[getCol(["sektor awam"])] || "",
      tempohJawatan: rowData[getCol(["di jawatan sekarang"])] || "",
      kursus: rowData[getCol(["nama kursus", "tajuk kursus", "latihan"])] || "",
      tarikhMula: isoLatihan_(rowData[getCol(["tarikh mula", "tarikh kursus", "tarikh"])]),
      tarikhTamat: isoLatihan_(rowData[getCol(["tarikh tamat", "tarikh akhir"])]),
      tempohKursus: rowData[getCol(["tempoh kursus", "hari latihan"])] || "",
      tempat: rowData[getCol(["lokasi", "tempat kursus"])] || "",
      jenisKursus: rowData[getCol(["jenis kursus"])] || "",
      kategoriKursus: rowData[getCol(["kategori kursus"])] || "",
      anjuran: rowData[getCol(["anjur"])] || "",
      kos: rowData[getCol(["kos"])] || 0,
      jam: rowData[getCol(["bilangan jam", "jam"])] || 0,
      cpd: rowData[getCol(["cpd"])] || 0,
      hrmis: rowData[getCol(["telah diisi dalam hrmis", "hrmis"])] || "" 
    };

    var colTahap1Start = getCol(["objektif kursus tercapai"]);
    if (colTahap1Start > -1 && rowData[colTahap1Start] !== "") {
      data.tahap1 = {};
      for (var i = 0; i < 25; i++) {
        data.tahap1["S" + (i + 1)] = rowData[colTahap1Start + i] || "";
      }
    }

    return { success: true, data: data };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// FUNGSI 2: DASHBOARD - DIOPTIMASI UNTUK LOADING CEPAT
// ==========================================
function dashboardTab_(sheet) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!sheet) return { status: "ERROR", message: "Tab data respon borang tidak dijumpai!" };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: "SUCCESS", senaraiJawatan: [], dataLengkap: {} };

    // OPTIMASI 1: HANYA AMBIL HEADERS SEKALI
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    function cariKolom(keywords) {
      for (var i = 0; i < headers.length; i++) {
        for (var k = 0; k < keywords.length; k++) {
          if (headers[i].indexOf(keywords[k].toLowerCase()) > -1) return i; 
        }
      } return -1;
    }

    var minIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].indexOf("status borang") > -1 || headers[i].indexOf("keberkesanan") > -1) {
        minIdx = i; break;
      }
    }
    if (minIdx === -1) minIdx = 20;

    function cariKolomBelakangSafe(keywords) {
      for (var i = headers.length - 1; i > minIdx; i--) {
        for (var k = 0; k < keywords.length; k++) {
          if (headers[i].indexOf(keywords[k].toLowerCase()) > -1) return i; 
        }
      } return -1;
    }

    // PRE-COMPUTE COLUMN INDICES
    var idxNama = cariKolom(["nama penuh", "nama pegawai"]);
    var idxIC = cariKolom(["kad pengenalan", "no. kp"]);
    var idxJawatan = cariKolom(["perjawatan", "jawatan"]);
    var idxFasiliti = cariKolom(["tempat bertugas", "jabatan"]);
    var idxKursus = cariKolom(["nama kursus", "tajuk kursus", "latihan"]);
    var idxTarikhMula = cariKolom(["tarikh mula", "tarikh kursus", "tarikh"]);
    var idxTarikhTamat = cariKolom(["tarikh tamat", "tarikh akhir"]);
    var idxJenis = cariKolom(["jenis kursus"]);
    var idxKategori = cariKolom(["kategori kursus"]);
    var idxTempat = cariKolom(["lokasi", "tempat kursus"]); 
    var idxAnjuran = cariKolom(["anjur"]);
    var idxJam = cariKolom(["bilangan jam", "jam"]);
    var idxCPD = cariKolom(["cpd"]);
    var idxHRMIS = cariKolom(["telah diisi dalam hrmis", "hrmis"]); 
    var idxSijil = cariKolom(["sijil", "bukti kehadiran"]);
    var idxTahap1Start = cariKolom(["objektif kursus tercapai"]); 
    var idxStatusKeb = cariKolom(["status borang", "keberkesanan"]); 

    var idxTarikhBos = cariKolomBelakangSafe(["tarikh dinilai", "tarikh bahagian e", "tarikh bahagian d", "tarikh penilaian", "tarikh"]); 
    var idxNamaBos = cariKolomBelakangSafe(["1. nama pegawai", "nama pegawai (penyelia)", "nama penyelia", "nama ketua", "nama"]);
    var idxJawatanBos = cariKolomBelakangSafe(["2. jawatan", "jawatan penyelia", "jawatan"]);
    var idxGredBos = cariKolomBelakangSafe(["3. gred", "gred jawatan", "gred"]);
    var idxSkimBos = cariKolomBelakangSafe(["4. skim", "skim penyelia", "skim"]);
    var idxTempohBos = cariKolomBelakangSafe(["5. tempoh", "tempoh menjadi", "tempoh seliaan", "tempoh"]);
    
    var kriteria17 = [
      ["1.1", "tahap pengetahuan pyd"], ["1.2", "mengaplikasi pengetahuan"], ["1.3", "berkongsi pengetahuan"], ["1.4", "kualiti hasil", "hasil kerja", "peningkatan dalam kualiti"], 
      ["2.1", "tahap kemahiran pyd"], ["2.2", "mengaplikasi kemahiran"], ["2.3", "berkongsi kemahiran"], ["2.4", "cekap menjalankan tugas", "berkesan"], 
      ["3.1", "prestasi kerja pyd"], ["3.2", "sumbangan pyd"], ["3.3", "cekap dalam menyelesaikan"], ["3.4", "akauntabiliti pyd"], 
      ["4.1", "memikul tanggungjawab"], ["4.2", "berdisiplin"], ["4.3", "bersikap positif"], ["4.4", "ketahanan mental"], ["4.5", "keyakinan dan komitmen"]
    ];
    var idxKriteria17 = kriteria17.map(function(k) { return cariKolomBelakangSafe(k); });
    var idxUlasanBos = cariKolomBelakangSafe(["cadangan penambahbaikan", "ulasan"]);

    // OPTIMASI 2: BATCH GET ALL DATA SEKALI
    var allData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    
    // OPTIMASI 3: GUNAKAN OBJECT UNTUK CEPAT ACCESS
    var dashboardData = {};
    var jawatanSet = {};
    var latestDetails = {};
    var hariIni = new Date();
    var hariIniTime = hariIni.getTime();

    // PASS 1: Kumpul latest details
    for (var r = 0; r < allData.length; r++) {
      var row = allData[r];
      var rawIC = idxIC > -1 ? row[idxIC].toString().trim().replace(/[^0-9]/g, '') : "";
      if(!rawIC) continue;
      
      latestDetails[rawIC] = {
         nama: idxNama > -1 ? row[idxNama].toString().trim().toUpperCase() : "TIADA NAMA",
         jawatan: idxJawatan > -1 && row[idxJawatan] ? row[idxJawatan].toString().trim() : "TIADA JAWATAN",
         fasiliti: idxFasiliti > -1 && row[idxFasiliti] ? row[idxFasiliti].toString().trim() : "Tiada Rekod Tempat"
      };
    }

    // PASS 2: Process each row
    for (var r = 0; r < allData.length; r++) {
      var row = allData[r];
      var actualRow = r + 2; // Actual sheet row number
      
      var icBorang = idxIC > -1 ? row[idxIC].toString().trim().replace(/[^0-9]/g, '') : "";
      if (!icBorang) continue; 
      
      var latest = latestDetails[icBorang];
      if (!latest) continue;
      
      var jawatanAsli = latest.jawatan;
      var namaAsli = latest.nama;
      var fasilitiAsli = latest.fasiliti;

      // Check if kursus penuh
      var isKursusPenuh = false;
      var tahap1Data = null;
      if (idxTahap1Start > -1 && row[idxTahap1Start] !== "" && row[idxTahap1Start] != null) {
        isKursusPenuh = true;
        tahap1Data = [];
        for (var i = 0; i < 25; i++) {
          tahap1Data.push((row[idxTahap1Start + i] || "-"));
        }
      }
      
      // Date processing
      var tglMula = null;
      var tglMulaRaw = (idxTarikhMula > -1 && row[idxTarikhMula]) ? row[idxTarikhMula] : null;
      if(tglMulaRaw) {
        tglMula = dateLatihan_(tglMulaRaw);
        if(!tglMula || isNaN(tglMula.getTime())) tglMula = null; 
      }

      var tglTamat = null;
      var tglTamatRaw = (idxTarikhTamat > -1 && row[idxTarikhTamat]) ? row[idxTarikhTamat] : null;
      if(tglTamatRaw && tglTamatRaw !== "") {
        tglTamat = dateLatihan_(tglTamatRaw);
        if(!tglTamat || isNaN(tglTamat.getTime())) tglTamat = tglMula; 
      } else {
        tglTamat = tglMula; 
      }

      var tahunKursus = tglMula ? tglMula.getFullYear().toString() : hariIni.getFullYear().toString();
      var jam = (idxJam > -1 && row[idxJam]) ? parseFloat(row[idxJam]) || 0 : 0;
      var cpdValue = (idxCPD > -1 && row[idxCPD]) ? parseFloat(row[idxCPD]) || 0 : 0;
      var statusKeb = (idxStatusKeb > -1 && row[idxStatusKeb]) ? row[idxStatusKeb].toString() : "Menunggu";
      
      var perluPenilaian = false;
      var bezaHari = 0;
      
      if (isKursusPenuh && tglTamat) {
        var tglMatang = new Date(tglTamat);
        tglMatang.setDate(tglMatang.getDate() + 90); 
        bezaHari = Math.floor((hariIniTime - tglTamat.getTime()) / (1000 * 3600 * 24));
        if (hariIni >= tglMatang && statusKeb !== "Selesai") perluPenilaian = true;
      }

      var bahagianCData = null;
      if (isKursusPenuh && statusKeb === "Selesai") {
        var skor17 = [];
        for(var s=0; s<17; s++) {
           skor17.push(idxKriteria17[s] > -1 ? (row[idxKriteria17[s]] || "-") : "-");
        }
        var valTarikhBos = idxTarikhBos > -1 ? row[idxTarikhBos] : null;
        var formattedTarikhBos = "-";
        if (valTarikhBos) {
           formattedTarikhBos = (valTarikhBos instanceof Date) ? Utilities.formatDate(valTarikhBos, "GMT+8", "dd/MM/yyyy") : valTarikhBos.toString();
        }

        bahagianCData = {
          tarikhBos: formattedTarikhBos,
          namaBos: idxNamaBos > -1 ? (row[idxNamaBos] || "-") : "-",
          jawatanBos: idxJawatanBos > -1 ? (row[idxJawatanBos] || "-") : "-",
          gredBos: idxGredBos > -1 ? (row[idxGredBos] || "-") : "-",
          skimBos: idxSkimBos > -1 ? (row[idxSkimBos] || "-") : "-",
          tempohBos: idxTempohBos > -1 ? (row[idxTempohBos] || "-") : "-",
          skor: skor17,
          ulasan: idxUlasanBos > -1 ? (row[idxUlasanBos] || "-") : "-"
        };
      }

      // Initialize if needed
      if (!dashboardData[jawatanAsli]) { 
        dashboardData[jawatanAsli] = {}; 
        jawatanSet[jawatanAsli] = true;
      }
      if (!dashboardData[jawatanAsli][icBorang]) { 
        dashboardData[jawatanAsli][icBorang] = { ic: icBorang, nama: namaAsli, tempatBertugas: fasilitiAsli, courses: [] }; 
      }

      dashboardData[jawatanAsli][icBorang].courses.push({
        row: refLatihan_(sheet, actualRow, row), 
        isKursusPenuh: isKursusPenuh,
        namaKursus: idxKursus > -1 ? (row[idxKursus] || "Tiada Rekod") : "Tiada Rekod", 
        tarikhMula: tglMula ? Utilities.formatDate(tglMula, "GMT+8", "dd/MM/yyyy") : "-", 
        tarikhTamat: tglTamat ? Utilities.formatDate(tglTamat, "GMT+8", "dd/MM/yyyy") : "-", 
        tahun: tahunKursus, 
        jenisKursus: idxJenis > -1 ? (row[idxJenis] || "-") : "-",
        kategoriKursus: idxKategori > -1 ? (row[idxKategori] || "-") : "-",
        tempat: idxTempat > -1 ? (row[idxTempat] || "-") : "-",
        anjuran: idxAnjuran > -1 ? (row[idxAnjuran] || "-") : "-",
        jamKursus: jam, 
        cpd: cpdValue,
        hrmis: idxHRMIS > -1 ? (row[idxHRMIS] || "-") : "-", 
        sijil: idxSijil > -1 ? (row[idxSijil] || "") : "",
        tahap1: tahap1Data,
        perluKeb: perluPenilaian,
        statusKeb: statusKeb,
        hariBerlalu: bezaHari,
        bahagianC: bahagianCData
      });
    }
    
    var jawatanList = Object.keys(jawatanSet).sort();
    return { status: "SUCCESS", senaraiJawatan: jawatanList, dataLengkap: dashboardData };
  } catch (e) { return { status: "ERROR", message: e.toString() }; }
}


// ==========================================
// FUNGSI 3: PADAM REKOD DENGAN IC
// ==========================================
function padamRekodTab_(row, ic, sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().toLowerCase().trim(); });
  
  var colIC = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].indexOf("no. kp") > -1 || headers[i].indexOf("kad pengenalan") > -1) { 
      colIC = i + 1; break; 
    }
  }
  if (colIC === -1) return { success: false, message: "Ralat Sistem: Kolum IC tidak dijumpai." };

  var rekodIC = sheet.getRange(row, colIC).getValue().toString().trim().replace(/[^0-9]/g, '');
  var inputIC = ic.toString().trim().replace(/[^0-9]/g, '');
  
  if (rekodIC === inputIC) {
    sheet.deleteRow(row);
    return { success: true, message: "Rekod anda berjaya dipadam!" };
  } else {
    return { success: false, message: "No. Kad Pengenalan TIDAK sepadan. Rekod gagal dipadam." };
  }
}

// ==========================================
// FUNGSI 4: SIMPAN PENILAIAN DENGAN ZON SELAMAT
// ==========================================
function penilaianTab_(row, dataPenilaianStr, sheet) {
  try {
    var data = JSON.parse(dataPenilaianStr);
    var rowNum = parseInt(row);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    var rowDataToUpdate = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    var isModified = false;

    var minIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].indexOf("status borang") > -1 || headers[i].indexOf("keberkesanan") > -1) {
        minIdx = i; break;
      }
    }
    if (minIdx === -1) minIdx = 20;

    function setFastByHeader(keywords, value) {
      for (var i = 0; i < headers.length; i++) {
        for (var k = 0; k < keywords.length; k++) {
          if (headers[i].indexOf(keywords[k].toLowerCase()) > -1) {
            rowDataToUpdate[i] = value;
            isModified = true;
            return true;
          }
        }
      } return false;
    }

    function setFastByHeaderBelakangSafe(keywords, value) {
      for (var i = headers.length - 1; i > minIdx; i--) {
        for (var k = 0; k < keywords.length; k++) {
          if (headers[i].indexOf(keywords[k].toLowerCase()) > -1) {
            rowDataToUpdate[i] = value;
            isModified = true;
            return true;
          }
        }
      } return false;
    }

    function setFastByHeaderTarikhSafe(keywords, value) {
      for (var i = headers.length - 1; i >= 0; i--) {
        for (var k = 0; k < keywords.length; k++) {
          if (headers[i].indexOf(keywords[k].toLowerCase()) > -1) {
            rowDataToUpdate[i] = value;
            isModified = true;
            return true;
          }
        }
      } return false;
    }
    
    setFastByHeader(["status borang", "keberkesanan"], "Selesai");
    
    var dToday = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    setFastByHeaderTarikhSafe(["tarikh dinilai penyelia", "tarikh dinilai", "tarikh bahagian e", "tarikh penilaian"], dToday); 
    
    setFastByHeaderBelakangSafe(["1. nama pegawai", "nama pegawai (penyelia)", "nama penyelia", "nama ketua", "nama"], data.namaBos);
    setFastByHeaderBelakangSafe(["2. jawatan", "jawatan penyelia", "jawatan"], data.jawatanBos);
    setFastByHeaderBelakangSafe(["3. gred", "gred jawatan", "gred"], data.gredBos);
    setFastByHeaderBelakangSafe(["4. skim", "skim penyelia", "skim"], data.skimBos);
    setFastByHeaderBelakangSafe(["5. tempoh", "tempoh menjadi", "tempoh seliaan", "tempoh"], data.tempohTahun + " Tahun " + data.tempohBulan + " Bulan");
    
    var kriteria17 = [
      ["1.1", "tahap pengetahuan pyd"], ["1.2", "mengaplikasi pengetahuan"], ["1.3", "berkongsi pengetahuan"], ["1.4", "kualiti hasil", "hasil kerja", "peningkatan dalam kualiti"], 
      ["2.1", "tahap kemahiran pyd"], ["2.2", "mengaplikasi kemahiran"], ["2.3", "berkongsi kemahiran"], ["2.4", "cekap menjalankan tugas", "berkesan"], 
      ["3.1", "prestasi kerja pyd"], ["3.2", "sumbangan pyd"], ["3.3", "cekap dalam menyelesaikan"], ["3.4", "akauntabiliti pyd"], 
      ["4.1", "memikul tanggungjawab"], ["4.2", "berdisiplin"], ["4.3", "bersikap positif"], ["4.4", "ketahanan mental"], ["4.5", "keyakinan dan komitmen"]
    ];
    for(var s = 0; s < 17; s++) {
       setFastByHeaderBelakangSafe(kriteria17[s], data.skor[s]);
    }
    
    setFastByHeaderBelakangSafe(["cadangan penambahbaikan", "ulasan"], data.ulasan);
    
    if (isModified) {
       sheet.getRange(rowNum, 1, 1, rowDataToUpdate.length).setValues([rowDataToUpdate]);
    }
    
    return { success: true, message: "Kajian Keberkesanan (Penyelia) berjaya disimpan!" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// FUNGSI 5: SEMAK AKSES ADMIN & PENYELIA
// ==========================================
function checkAdminPenyelia(ic) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("admin n penyelia");
    if (!sheet) return { success: false, message: "Ralat: Sheet 'admin n penyelia' tidak dijumpai. Sila cipta sheet ini dahulu." };

    var data = sheet.getDataRange().getValues();
    var inputIC = ic.toString().trim().replace(/[^0-9]/g, '');

    if(inputIC === "") return { success: false, message: "Sila masukkan No. IC yang sah." };

    for (var r = 1; r < data.length; r++) { 
      var rekodIC = data[r][1] ? data[r][1].toString().trim().replace(/[^0-9]/g, '') : "";
      if (rekodIC === inputIC) {
        return { success: true, message: "Akses dibenarkan." };
      }
    }
    return { success: false, message: "Akses DITOLAK. No. Kad Pengenalan anda tidak berdaftar sebagai Admin/Penyelia." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}


// ==========================================
// FUNGSI 6: LOAD DATA IKUT JAWATAN (PAGINATION)
// ==========================================

function _formatTarikh(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return null;
  var dd = d.getDate();
  var mm = d.getMonth() + 1;
  var yyyy = d.getFullYear();
  return (dd < 10 ? "0" + dd : dd) + "/" + (mm < 10 ? "0" + mm : mm) + "/" + yyyy;
}


function jawatanTab_(sheet) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheet) return { status: "ERROR", message: "Tab tidak dijumpai!" };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { status: "SUCCESS", senaraiJawatan: [], tahunList: [] };

    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });

    function cariKolom(keywords) {
      for (var i = 0; i < headers.length; i++)
        for (var k = 0; k < keywords.length; k++)
          if (headers[i].indexOf(keywords[k].toLowerCase()) > -1) return i;
      return -1;
    }

    var idxJawatan    = cariKolom(["perjawatan", "jawatan"]);
    var idxIC         = cariKolom(["kad pengenalan", "no. kp"]);
    var idxTarikhMula = cariKolom(["tarikh mula", "tarikh kursus", "tarikh"]);

    var jawatanSet = {};
    var jawatanList = [];
    var tahunSet = {};

    for (var r = 1; r < data.length; r++) {
      var ic = idxIC > -1 ? data[r][idxIC].toString().trim().replace(/[^0-9]/g, '') : "";
      if (!ic) continue;

      var jaw = idxJawatan > -1 && data[r][idxJawatan] ? data[r][idxJawatan].toString().trim() : "TIADA JAWATAN";
      if (!jawatanSet[jaw]) { jawatanSet[jaw] = true; jawatanList.push(jaw); }

      var tglRaw = idxTarikhMula > -1 ? data[r][idxTarikhMula] : null;
      if (tglRaw) {
        var tgl = dateLatihan_(tglRaw);
        if (tgl && !isNaN(tgl.getTime())) tahunSet[tgl.getFullYear().toString()] = true;
      }
    }

    jawatanList.sort();
    var tahunList = Object.keys(tahunSet).sort(function(a,b){ return b-a; });

    return { status: "SUCCESS", senaraiJawatan: jawatanList, tahunList: tahunList };
  } catch(e) { return { status: "ERROR", message: e.toString() }; }
}

// ==========================================
// PATCH: tahun kini OPSYENAL. Jika tahun tidak dihantar ("" / null),
// SEMUA tahun akan dikembalikan (perlu untuk tab Kajian Keberkesanan,
// supaya kursus tahun lepas yang baru matang 90 hari tetap kelihatan).
// Client-side (renderStaffList) akan tapis ikut tahun untuk paparan
// "Semak Rekod Latihan" sahaja.
// ==========================================
function dashboardJawatanTab_(jawatan, tahun, sheet) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheet) return { status: "ERROR", message: "Tab tidak dijumpai!" };

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    var hLen = headers.length;

    function cariKolom(keywords) {
      for (var i = 0; i < hLen; i++)
        for (var k = 0; k < keywords.length; k++)
          if (headers[i].indexOf(keywords[k].toLowerCase()) > -1) return i;
      return -1;
    }

    var minIdx = -1;
    for (var i = 0; i < hLen; i++) {
      if (headers[i].indexOf("status borang") > -1 || headers[i].indexOf("keberkesanan") > -1) { minIdx = i; break; }
    }
    if (minIdx === -1) minIdx = 20;

    function cariKolomBelakang(keywords) {
      for (var i = hLen - 1; i > minIdx; i--)
        for (var k = 0; k < keywords.length; k++)
          if (headers[i].indexOf(keywords[k].toLowerCase()) > -1) return i;
      return -1;
    }

    var idxNama        = cariKolom(["nama penuh", "nama pegawai"]);
    var idxIC          = cariKolom(["kad pengenalan", "no. kp"]);
    var idxJawatan     = cariKolom(["perjawatan", "jawatan"]);
    var idxFasiliti    = cariKolom(["tempat bertugas", "jabatan"]);
    var idxKursus      = cariKolom(["nama kursus", "tajuk kursus", "latihan"]);
    var idxTarikhMula  = cariKolom(["tarikh mula", "tarikh kursus", "tarikh"]);
    var idxTarikhTamat = cariKolom(["tarikh tamat", "tarikh akhir"]);
    var idxJenis       = cariKolom(["jenis kursus"]);
    var idxKategori    = cariKolom(["kategori kursus"]);
    var idxTempat      = cariKolom(["lokasi", "tempat kursus"]);
    var idxAnjuran     = cariKolom(["anjur"]);
    var idxJam         = cariKolom(["bilangan jam", "jam"]);
    var idxCPD         = cariKolom(["cpd"]);
    var idxHRMIS       = cariKolom(["telah diisi dalam hrmis", "hrmis"]);
    var idxSijil       = cariKolom(["sijil", "bukti kehadiran"]);
    var idxTahap1Start = cariKolom(["objektif kursus tercapai"]);
    var idxStatusKeb   = cariKolom(["status borang", "keberkesanan"]);
    var idxTarikhBos   = cariKolomBelakang(["tarikh dinilai", "tarikh bahagian e", "tarikh bahagian d", "tarikh penilaian", "tarikh"]);
    var idxNamaBos     = cariKolomBelakang(["1. nama pegawai", "nama pegawai (penyelia)", "nama penyelia", "nama ketua", "nama"]);
    var idxJawatanBos  = cariKolomBelakang(["2. jawatan", "jawatan penyelia", "jawatan"]);
    var idxGredBos     = cariKolomBelakang(["3. gred", "gred jawatan", "gred"]);
    var idxSkimBos     = cariKolomBelakang(["4. skim", "skim penyelia", "skim"]);
    var idxTempohBos   = cariKolomBelakang(["5. tempoh", "tempoh menjadi", "tempoh seliaan", "tempoh"]);
    var idxUlasanBos   = cariKolomBelakang(["cadangan penambahbaikan", "ulasan"]);

    var kriteria17 = [
      ["1.1","tahap pengetahuan pyd"],["1.2","mengaplikasi pengetahuan"],["1.3","berkongsi pengetahuan"],["1.4","kualiti hasil","hasil kerja","peningkatan dalam kualiti"],
      ["2.1","tahap kemahiran pyd"],["2.2","mengaplikasi kemahiran"],["2.3","berkongsi kemahiran"],["2.4","cekap menjalankan tugas","berkesan"],
      ["3.1","prestasi kerja pyd"],["3.2","sumbangan pyd"],["3.3","cekap dalam menyelesaikan"],["3.4","akauntabiliti pyd"],
      ["4.1","memikul tanggungjawab"],["4.2","berdisiplin"],["4.3","bersikap positif"],["4.4","ketahanan mental"],["4.5","keyakinan dan komitmen"]
    ];
    var idxKriteria17 = kriteria17.map(function(k){ return cariKolomBelakang(k); });

    // ══════════════════════════════════════════════════════
    // FILTER: Jawatan (WAJIB). Tahun (OPSYENAL - hanya jika dibekalkan)
    // ══════════════════════════════════════════════════════
    var filteredRows = [];
    var icSet = {};
    for (var r = 1; r < data.length; r++) {
      var ic = idxIC > -1 ? data[r][idxIC].toString().trim().replace(/[^0-9]/g,'') : "";
      if (!ic) continue;
      var jaw = idxJawatan > -1 && data[r][idxJawatan] ? data[r][idxJawatan].toString().trim() : "TIADA JAWATAN";
      if (jaw !== jawatan) continue; // Tiada lagi "ALL" - mesti pilih satu jawatan
      
      // Filter tahun HANYA jika param tahun dibekalkan (bukan string kosong)
      if (tahun && yearPortal_(idxTarikhMula > -1 ? data[r][idxTarikhMula] : null) !== Number(tahun)) continue;
      
      filteredRows.push(r);
      icSet[ic] = jaw;
    }

    // Kumpul latest info per IC
    var latestDetails = {};
    for (var fi = 0; fi < filteredRows.length; fi++) {
      var r = filteredRows[fi];
      var rawIC = data[r][idxIC].toString().trim().replace(/[^0-9]/g,'');
      latestDetails[rawIC] = {
        nama:    idxNama > -1 && data[r][idxNama] ? data[r][idxNama].toString().trim().toUpperCase() : "TIADA NAMA",
        jawatan: jawatan,
        fasiliti: idxFasiliti > -1 && data[r][idxFasiliti] ? data[r][idxFasiliti].toString().trim() : "Tiada Rekod Tempat"
      };
    }

    var hasilData = {};
    var hariIniMs = new Date().getTime();
    var tahunSemasa = new Date().getFullYear().toString();

    for (var fi = 0; fi < filteredRows.length; fi++) {
      var r = filteredRows[fi];
      var row = data[r];
      var icBorang = row[idxIC].toString().trim().replace(/[^0-9]/g,'');
      var latest = latestDetails[icBorang];
      if (!latest) continue;

      var isKursusPenuh = false;
      var tahap1Data = null;
      if (idxTahap1Start > -1 && row[idxTahap1Start] !== "") {
        isKursusPenuh = true;
        tahap1Data = [];
        for (var i = 0; i < 25; i++) tahap1Data.push(row[idxTahap1Start + i] || "-");
      }

      var tglMula = null, tglTamat = null;
      if (idxTarikhMula > -1 && row[idxTarikhMula]) {
        var d = dateLatihan_(row[idxTarikhMula]);
        if (d && !isNaN(d.getTime())) tglMula = d;
      }
      if (idxTarikhTamat > -1 && row[idxTarikhTamat]) {
        var d2 = dateLatihan_(row[idxTarikhTamat]);
        if (d2 && !isNaN(d2.getTime())) tglTamat = d2;
      }
      if (!tglTamat) tglTamat = tglMula;

      var tahunKursus = tglMula ? tglMula.getFullYear().toString() : tahunSemasa;
      var jam = (idxJam > -1 && row[idxJam]) ? parseFloat(row[idxJam]) || 0 : 0;
      var cpdValue = (idxCPD > -1 && row[idxCPD]) ? parseFloat(row[idxCPD]) || 0 : 0;
      var statusKeb = (idxStatusKeb > -1 && row[idxStatusKeb]) ? row[idxStatusKeb].toString() : "Menunggu";

      var perluPenilaian = false, bezaHari = 0;
      if (isKursusPenuh && tglTamat) {
        bezaHari = Math.floor((hariIniMs - tglTamat.getTime()) / 86400000);
        if (bezaHari >= 90 && statusKeb !== "Selesai") perluPenilaian = true;
      }

      var bahagianCData = null;
      if (isKursusPenuh && statusKeb === "Selesai") {
        var skor17 = [];
        for (var s = 0; s < 17; s++) skor17.push(idxKriteria17[s] > -1 ? row[idxKriteria17[s]] : "-");
        var tglBosObj = idxTarikhBos > -1 ? (row[idxTarikhBos] instanceof Date ? row[idxTarikhBos] : new Date(row[idxTarikhBos])) : null;
        var formattedTarikhBos = (tglBosObj && !isNaN(tglBosObj.getTime())) ? _formatTarikh(tglBosObj) : (idxTarikhBos > -1 && row[idxTarikhBos] ? row[idxTarikhBos].toString() : "-");
        bahagianCData = {
          tarikhBos: formattedTarikhBos,
          namaBos:    idxNamaBos    > -1 ? row[idxNamaBos]    : "-",
          jawatanBos: idxJawatanBos > -1 ? row[idxJawatanBos] : "-",
          gredBos:    idxGredBos    > -1 ? row[idxGredBos]    : "-",
          skimBos:    idxSkimBos    > -1 ? row[idxSkimBos]    : "-",
          tempohBos:  idxTempohBos  > -1 ? row[idxTempohBos]  : "-",
          skor: skor17,
          ulasan: idxUlasanBos > -1 ? row[idxUlasanBos] : "-"
        };
      }

      var jawatanAsli = latest.jawatan;
      if (!hasilData[jawatanAsli]) hasilData[jawatanAsli] = {};
      if (!hasilData[jawatanAsli][icBorang]) {
        hasilData[jawatanAsli][icBorang] = { ic: icBorang, nama: latest.nama, tempatBertugas: latest.fasiliti, courses: [] };
      }
      hasilData[jawatanAsli][icBorang].courses.push({
        row: refLatihan_(sheet, r + 1, row), isKursusPenuh: isKursusPenuh,
        namaKursus:     idxKursus   > -1 ? row[idxKursus]   : "Tiada Rekod",
        tarikhMula:     _formatTarikh(tglMula)  || "-",
        tarikhTamat:    _formatTarikh(tglTamat) || "-",
        tahun: tahunKursus,
        jenisKursus:    idxJenis    > -1 ? row[idxJenis]    : "-",
        kategoriKursus: idxKategori > -1 ? row[idxKategori] : "-",
        tempat:         idxTempat   > -1 ? row[idxTempat]   : "-",
        anjuran:        idxAnjuran  > -1 ? row[idxAnjuran]  : "-",
        jamKursus: jam, cpd: cpdValue,
        hrmis:  idxHRMIS > -1 ? row[idxHRMIS] : "-",
        sijil:  idxSijil > -1 ? row[idxSijil] : "",
        tahap1: tahap1Data, perluKeb: perluPenilaian,
        statusKeb: statusKeb, hariBerlalu: bezaHari, bahagianC: bahagianCData
      });
    }

    return { status: "SUCCESS", dataLengkap: hasilData };
  } catch(e) {
    Logger.log("ERROR: " + e.toString());
    return { status: "ERROR", message: e.toString() };
  }
}


function testDashboard() {
  var result = getDashboardData();
  Logger.log("Status: " + result.status);
  Logger.log("Bil jawatan: " + result.senaraiJawatan.length);
  Logger.log("Data size (bytes): " + JSON.stringify(result).length);
}


function testJururawat() {
  var result = getDashboardDataByJawatan("Jururawat (Staff Nurse)", "");
  Logger.log("Status: " + result.status);
  Logger.log("Data size: " + JSON.stringify(result).length);
  var keys = Object.keys(result.dataLengkap);
  Logger.log("Jawatan dalam result: " + keys.join(", "));
  if (keys.length > 0) {
    var firstJaw = keys[0];
    var icList = Object.keys(result.dataLengkap[firstJaw]);
    Logger.log("Bil staf: " + icList.length);
  }
}


// ==========================================
// FUNGSI DASHBOARD LANDING PAGE: % STAF CAPAI 40 JAM IKUT JAWATAN
// ==========================================
// Tambahan graf kategori jam sahaja. Kod pengesahan IC asal dikekalkan.
var PORTAL_V2={year:2026,staff:551};
function icPortal_(value) { return String(value == null ? '' : value).replace(/[-\s]/g,''); }
function colPortal_(headers, terms) {
  // Prefer exact matches, then preferred synonyms in order, before generic ones.
  var h = headers.map(function(x) { return String(x).toLowerCase().trim(); });
  for (var k = 0; k < terms.length; k++) { var exact = h.indexOf(terms[k]); if (exact >= 0) return exact; }
  for (var k = 0; k < terms.length; k++) for (var i = 0; i < h.length; i++) if (h[i].indexOf(terms[k]) >= 0) return i;
  return -1;
}
function yearPortal_(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? null : Number(Utilities.formatDate(value, 'Asia/Kuala_Lumpur', 'yyyy'));
  var s = String(value == null ? '' : value).trim(), m, y, mo, d;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?$/))) { y=+m[1]; mo=+m[2]; d=+m[3]; }
  else if ((m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s.*)?$/))) { y=+m[3]; mo=+m[2]; d=+m[1]; }
  else return null;
  var date = new Date(Date.UTC(y, mo-1, d));
  return date.getUTCFullYear() === y && date.getUTCMonth() === mo-1 && date.getUTCDate() === d ? y : null;
}
function parseJamPortal_(value) {
  var s = String(value == null ? '' : value).trim();
  if (!/^-?\d+(?:[.,]\d+)?$/.test(s)) return null;
  var n = Number(s.replace(',', '.')); return Number.isFinite(n) ? n : null;
}
function tambahJamTepat_(a,b) {
  var x=String(a).replace(',','.').split('.'), y=String(b).replace(',','.').split('.');
  var negativeA=x[0][0]==='-', negativeB=y[0][0]==='-';
  x[0]=x[0].replace('-',''); y[0]=y[0].replace('-','');
  var scale=Math.max((x[1]||'').length,(y[1]||'').length);
  function scaled(v){return (v[0]+(v[1]||'').padEnd(scale,'0')).replace(/^0+(?=\d)/,'');}
  var left=scaled(x),right=scaled(y),width=Math.max(left.length,right.length),sum='',sign=negativeA;
  left=left.padStart(width,'0');right=right.padStart(width,'0');
  if(negativeA===negativeB){
    var carry=0;
    for(var i=width-1;i>=0;i--){var n=Number(left[i])+Number(right[i])+carry;sum=String(n%10)+sum;carry=Math.floor(n/10);}
    if(carry)sum=String(carry)+sum;
  }else{
    if(left<right){var swap=left;left=right;right=swap;sign=negativeB;}
    var borrow=0;
    for(var i=width-1;i>=0;i--){var n=Number(left[i])-Number(right[i])-borrow;borrow=n<0?1:0;if(n<0)n+=10;sum=String(n)+sum;}
  }
  sum=sum.replace(/^0+(?=\d)/,'').padStart(scale+1,'0');
  var value=scale?sum.slice(0,-scale)+'.'+sum.slice(-scale):sum;
  return sign&&Number(value)!==0?'-'+value:value;
}

function aggregatePortal_(data, year, staff) {
  var headers = data[0] || [];
  var dateIdx = colPortal_(headers,['tarikh mula','tarikh mula kursus','tarikh kursus']);
  var icIdx = colPortal_(headers,['no. kad pengenalan','kad pengenalan','no. kp']);
  var jamIdx = colPortal_(headers,['bilangan jam latihan','bilangan jam']);
  var jawIdx = colPortal_(headers,['perjawatan','jawatan']);
  if ([dateIdx,icIdx,jamIdx].some(function(i) { return i < 0; })) throw new Error('Lajur Tarikh Mula, No. Kad Pengenalan dan Bilangan Jam Latihan diperlukan.');
  var negativeRows = [], totals = {}, jawatan = {}, issues = {tarikh:0, ic:0, jam:0, luarKategori:0}, rows = 0;
  for (var r = 1; r < data.length; r++) {
    if (data[r].every(function(v) { return v === '' || v === null; })) continue;
    var yearRow = yearPortal_(data[r][dateIdx]);
    if (yearRow === null) { issues.tarikh++; continue; }
    if (yearRow !== year) continue;
    rows++;
    var ic = icPortal_(data[r][icIdx]), jam = parseJamPortal_(data[r][jamIdx]);
    if (!/^\d{12}$/.test(ic)) { issues.ic++; continue; }
    if (!Object.prototype.hasOwnProperty.call(totals, ic)) totals[ic] = '0';
    if (jam === null) { issues.jam++; continue; }
    if (jam < 0) negativeRows.push(r+1);
    totals[ic] = tambahJamTepat_(totals[ic], String(data[r][jamIdx]).trim());
    jawatan[ic] = jawIdx >= 0 && data[r][jawIdx] ? String(data[r][jawIdx]).trim() : 'TIADA JAWATAN';
  }
  var unique = Object.keys(totals).length;
  if (unique > staff) throw new Error('Bilangan IC unik (' + unique + ') melebihi jumlah organisasi ' + staff + '. Semak data atau jumlah staf.');
  var counts = [staff-unique,0,0,0];
  Object.keys(totals).forEach(function(ic) {
    var n = Number(totals[ic]); totals[ic] = n;
    if (n >= 40) counts[3]++; else if (n >= 20) counts[2]++; else if (n >= 1) counts[1]++; else issues.luarKategori++;
  });
  var messages = [];
  if (issues.tarikh) messages.push(issues.tarikh + ' rekod mempunyai Tarikh Mula kosong/tidak sah');
  if (issues.ic) messages.push(issues.ic + ' rekod ' + year + ' mempunyai IC bukan 12 digit');
  if (issues.jam) messages.push(issues.jam + ' rekod ' + year + ' mempunyai jam kosong/tidak sah');
  if (issues.luarKategori) messages.push(issues.luarKategori + ' staf merekod jumlah kurang 1 jam (di luar kategori 1–19 jam; 0 Jam ditakrif sebagai tiada rekod sahaja)');
  if (messages.length) throw new Error('Graf belum dipaparkan: ' + messages.join('; ') + '. Betulkan data supaya pecahan lengkap.');
  return {warnings:negativeRows.length ? [negativeRows.length + ' rekod jam negatif pada baris ' + negativeRows.join(', ') + '. Nilai asal ini termasuk dalam pengiraan; sila semak pada Sheet.'] : [], tahun:year, jumlahStaf:staff, stafBerrekod:unique, bilRekod:rows, totals:totals, jawatan:jawatan,
    kategori:['0 Jam','1 - 19 Jam','20 - 39 Jam','40 Jam Ke Atas'].map(function(label,i) { return {label:label, bilangan:counts[i], peratus:Math.round(counts[i]/staff*1000)/10}; })};
}
function grafKategoriLatihan_(year, data) {
  try {
    var result=aggregatePortal_(data,year,stafTahunLatihan_(year));
    delete result.totals;delete result.jawatan;result.status='SUCCESS';return result;
  }catch(e){return {status:'ERROR',message:e.message};}
}
function getDashboardKategoriJam(tahun) {
  try {var year=tahunGrafLatihan_(tahun);return grafKategoriLatihan_(year,dataGrafLatihan_(year));}
  catch(e){return {status:'ERROR',message:e.message};}
}

// ==========================================
// TAB LATIHAN MENGIKUT TAHUN
// Jalankan sediakanTabLatihan2027_ SEKALI dari editor sebelum rename tab lama.
// Fungsi berakhir '_' hanya untuk dalaman/editor, bukan google.script.run.
// ==========================================
var PORTAL_STAF_TAHUN = {2027: null}; // Isi bilangan sebenar sebelum guna graf 2027.
var TAHUN_TAB_LAMA = 2026; // Tab lama mungkin mengandungi rekod tahun lebih awal juga.
var SIJIL_ROOT_FOLDER_ID = '1rGGuKOxjtK1mdEII67hBEGnNiNjnBxxc';

// Dipanggil semasa simpan di bawah script lock: elak cipta dua subfolder serentak.
// Tahun sijil ikut Tarikh Mula latihan, bukan tahun upload atau tahun dashboard.
function folderSijilTahun_(tarikhMula) {
  var year = yearPortal_(tarikhMula);
  if (!year || year < 2000 || year > 2099) throw new Error('Tarikh Mula tidak sah untuk menentukan folder sijil.');
  var root;
  try {
    root = DriveApp.getFolderById(SIJIL_ROOT_FOLDER_ID);
    if (root.isTrashed()) throw new Error('Folder dalam Trash.');
  } catch(e) {
    throw new Error('Folder utama sijil tidak dapat diakses. Semak SIJIL_ROOT_FOLDER_ID dan akses akaun yang deploy portal.');
  }
  var folders = root.getFoldersByName(String(year));
  var folder = folders.hasNext() ? folders.next() : null;
  if (folders.hasNext()) throw new Error('Lebih daripada satu subfolder ' + year + ' ditemui dalam folder utama sijil. Semak folder berulang dahulu.');
  if (folder && folder.isTrashed()) throw new Error('Subfolder sijil ' + year + ' berada dalam Trash. Pulihkan atau semak folder dahulu.');
  return folder || root.createFolder(String(year));
}

function kunciRegistriLatihan_() {
  return 'PORTAL_TAB_TAHUN_V1_' + SpreadsheetApp.getActiveSpreadsheet().getId();
}
function senaraiTabLatihan_(validateSchema) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = PropertiesService.getScriptProperties().getProperty(kunciRegistriLatihan_());
  var map = raw ? JSON.parse(raw) : {};
  var found = {};
  Object.keys(map).forEach(function(y) {
    var sh = ss.getSheetById(Number(map[y]));
    if (!sh) throw new Error('Tab berdaftar tahun ' + y + ' tidak dijumpai. Semak tab yang dipadam; jangan cipta pengganti kosong.');
    found[y] = sh;
  });
  function tambah(y, sh) {
    if (!sh) return;
    if (found[y] && found[y].getSheetId() !== sh.getSheetId())
      throw new Error('Dua tab berbeza ditetapkan untuk tahun ' + y + '. Semak nama/pendaftaran tab dahulu.');
    found[y] = sh;
  }
  if (!found[TAHUN_TAB_LAMA]) {
    tambah(TAHUN_TAB_LAMA, ss.getSheetByName('Form responses 1') || ss.getSheetByName('Form Responses 1'));
  }
  ss.getSheets().forEach(function(sh) {
    var m = sh.getName().match(/^Latihan (20\d{2})$/);
    if (m) tambah(Number(m[1]), sh);
  });
  if (!found[TAHUN_TAB_LAMA]) throw new Error('Tab asal tidak ditemui. Kekalkan nama Form responses 1 atau Latihan 2026 sehingga pendaftaran tab selesai.');
  var seen = {};
  return Object.keys(found).sort(function(a,b){return a-b;}).map(function(y) {
    var id = found[y].getSheetId();
    if (seen[id]) throw new Error('Satu tab tidak boleh didaftarkan untuk dua tahun.');
    seen[id] = true;
    if (validateSchema !== false) semakSkemaLatihan_(found[y]);
    return {year:Number(y), sheet:found[y]};
  });
}
function simpanRegistriLatihan_(tabs) {
  var map = {};
  tabs.forEach(function(t){map[t.year] = t.sheet.getSheetId();});
  PropertiesService.getScriptProperties().setProperty(kunciRegistriLatihan_(), JSON.stringify(map));
}
function semakSkemaLatihan_(sh) {
  if (sh.getLastColumn() < 1) throw new Error('Tab ' + sh.getName() + ' tiada tajuk lajur.');
  var h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  [['kad pengenalan','no. kp'],['tarikh mula','tarikh kursus'],['nama kursus','tajuk kursus'],['bilangan jam']].forEach(function(keys) {
    if (colPortal_(h, keys) < 0) throw new Error('Lajur ' + keys[0] + ' tidak ditemui dalam tab ' + sh.getName() + '.');
  });
}
function tabUntukTahunLatihan_(year, create) {
  year = Number(year);
  if (!Number.isInteger(year) || year < 2000 || year > 2099) throw new Error('Tahun latihan tidak sah.');
  var tabs = senaraiTabLatihan_(), targetYear = year <= TAHUN_TAB_LAMA ? TAHUN_TAB_LAMA : year;
  var target = tabs.filter(function(t){return t.year === targetYear;})[0];
  if (target) { if (create) simpanRegistriLatihan_(tabs); return target.sheet; }
  if (!create) return null;
  var template = tabs.filter(function(t){return t.year === TAHUN_TAB_LAMA;})[0].sheet;
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.insertSheet('Latihan ' + targetYear);
  var n = template.getLastColumn();
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n-sh.getMaxColumns());
  var src = template.getRange(1,1,1,n), dst = sh.getRange(1,1,1,n);
  dst.setValues(src.getValues());
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sh.setFrozenRows(1);
  var icCol = colPortal_(src.getValues()[0], ['kad pengenalan','no. kp']) + 1;
  sh.getRange(2,icCol,sh.getMaxRows()-1,1).setNumberFormat('@');
  tabs.push({year:targetYear, sheet:sh});
  simpanRegistriLatihan_(tabs);
  return sh;
}
function sediakanTabLatihan2027_() {
  return denganKunciLatihan_(function() {
    var sh = tabUntukTahunLatihan_(2027, true);
    var msg = 'Tab 2027 sedia: ' + sh.getName() + '. ID semua tab disimpan; tab asal kini boleh dinamakan semula.';
    Logger.log(msg); return msg;
  });
}
function denganKunciLatihan_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); } }
}
function icLatihan_(value) { return String(value == null ? '' : value).replace(/[-\s]/g, ''); }
function dateLatihan_(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  var s = String(value == null ? '' : value).trim(), m;
  if (!yearPortal_(s)) return null;
  if ((m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return new Date(+m[1],+m[2]-1,+m[3],12);
  m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  return new Date(+m[3],+m[2]-1,+m[1],12);
}
function isoLatihan_(value) {
  var d = dateLatihan_(value);
  return d ? Utilities.formatDate(d,'Asia/Kuala_Lumpur','yyyy-MM-dd') : '';
}
function capRekodLatihan_(values) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(values), Utilities.Charset.UTF_8)
    .map(function(b){return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');
}
function refLatihan_(sheet, row, values) {
  return sheet.getSheetId() + ':' + row + ':' + capRekodLatihan_(values);
}
function sasaranRekodLatihan_(ref) {
  var m = String(ref).match(/^(\d+):(\d+):([a-f0-9]{64})$/);
  if (!m) throw new Error('Rujukan rekod lama/tidak sah. Refresh senarai sebelum meneruskan.');
  var tab = senaraiTabLatihan_().filter(function(t){return t.sheet.getSheetId() === Number(m[1]);})[0];
  var row = Number(m[2]);
  if (!tab || row < 2 || row > tab.sheet.getLastRow()) throw new Error('Rekod tidak ditemui. Refresh senarai.');
  var values = tab.sheet.getRange(row,1,1,tab.sheet.getLastColumn()).getValues()[0];
  if (capRekodLatihan_(values) !== m[3]) throw new Error('Rekod/baris telah berubah. Refresh senarai dan cuba semula supaya rekod lain tidak terubah.');
  return {sheet:tab.sheet, year:tab.year, row:row, values:values};
}
function simpanRekodLDP(dataFormStr, fileName, mimeType, base64Data) {
  try {
    return denganKunciLatihan_(function() {
      var form = JSON.parse(dataFormStr), year = yearPortal_(form.tarikhMula);
      if (!year) throw new Error('Tarikh mula tidak sah.');
      if (!String(form.nama || '').trim() || !icLatihan_(form.ic)) throw new Error('Nama dan IC diperlukan.');
      var sh, row;
      if (form.editRow) {
        var ctx = sasaranRekodLatihan_(form.editRow);
        sh=ctx.sheet; row=ctx.row;
        var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
        var idx=colPortal_(h,['kad pengenalan','no. kp']);
        if (icLatihan_(ctx.values[idx]) !== icLatihan_(form.ic)) throw new Error('IC rekod tidak sepadan.');
        if ((year <= TAHUN_TAB_LAMA ? TAHUN_TAB_LAMA : year) !== ctx.year)
          throw new Error('Edit merentas tab tahun belum disokong. Daftar rekod pada tahun yang betul dan padam rekod lama selepas semakan.');
      } else {
        sh=tabUntukTahunLatihan_(year,true); row=sh.getLastRow()+1;
      }
      if (row>sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(),row-sh.getMaxRows());
      return simpanRekodTab_(dataFormStr,fileName,mimeType,base64Data,sh,row);
    });
  } catch(e) { return {success:false,message:e.message}; }
}
function getRekodUntukEdit(ref, ic) {
  try { var ctx=sasaranRekodLatihan_(ref); return rekodEditTab_(ctx.row,ic,ctx.sheet); }
  catch(e) { return {success:false,message:e.message}; }
}
function padamRekodSahkanIC(ref, ic) {
  try { return denganKunciLatihan_(function(){var ctx=sasaranRekodLatihan_(ref); return padamRekodTab_(ctx.row,ic,ctx.sheet);}); }
  catch(e) { return {success:false,message:e.message}; }
}
// Pengakuan staf telah mengisi HRMIS; tidak mengubah jam HRMIS yang disahkan.
function kemasKiniStatusHRMIS(ref, ic, status) {
  try {
    return denganKunciLatihan_(function() {
      var value=String(status == null ? '' : status).trim().toLowerCase();
      if(value!=='ya' && value!=='tidak')throw new Error('Pilih Belum diisi atau Sudah diisi/dihantar.');
      var ctx=sasaranRekodLatihan_(ref);
      var headers=ctx.sheet.getRange(1,1,1,ctx.sheet.getLastColumn()).getValues()[0];
      var icIdx=colPortal_(headers,['kad pengenalan','no. kp']);
      var hrmisIdx=colPortal_(headers,['telah diisi dalam hrmis','hrmis']);
      if(icIdx<0 || hrmisIdx<0)throw new Error('Lajur IC atau status HRMIS tidak ditemui.');
      var inputIC=icLatihan_(ic), recordIC=icLatihan_(ctx.values[icIdx]);
      if(!/^\d+$/.test(inputIC) || !recordIC || inputIC!==recordIC)
        throw new Error('No. Kad Pengenalan tidak sepadan. Status tidak diubah.');
      var canonical=value==='ya' ? 'Ya' : 'Tidak';
      ctx.sheet.getRange(ctx.row,hrmisIdx+1).setValue(canonical);
      var updated=ctx.sheet.getRange(ctx.row,1,1,ctx.sheet.getLastColumn()).getValues()[0];
      return {success:true,hrmis:canonical,recordRef:refLatihan_(ctx.sheet,ctx.row,updated),message:'Status HRMIS berjaya dikemas kini.'};
    });
  }catch(e){return {success:false,message:e.message};}
}
function simpanPenilaianPenyelia(ref, dataPenilaianStr) {
  try { return denganKunciLatihan_(function(){var ctx=sasaranRekodLatihan_(ref); return penilaianTab_(ctx.row,dataPenilaianStr,ctx.sheet);}); }
  catch(e) { return {success:false,message:e.message}; }
}
function gabungDashboardLatihan_(getOne) {
  var data = {};
  senaraiTabLatihan_().forEach(function(tab) {
    var res=getOne(tab.sheet);
    if (!res || res.status !== 'SUCCESS') throw new Error(res ? res.message : 'Tiada respons tab.');
    Object.keys(res.dataLengkap || {}).forEach(function(j) {
      if (!data[j]) data[j]={};
      Object.keys(res.dataLengkap[j]).forEach(function(ic) {
        var staff=res.dataLengkap[j][ic], old=data[j][ic];
        data[j][ic]={ic:staff.ic,nama:staff.nama,tempatBertugas:staff.tempatBertugas,
          courses:(old ? old.courses : []).concat(staff.courses)};
      });
    });
  });
  sertakanJamHRMIS_(data);
  return {status:'SUCCESS',senaraiJawatan:Object.keys(data).sort(),dataLengkap:data};
}
function getDashboardData() {
  try { return gabungDashboardLatihan_(function(sh){return dashboardTab_(sh);}); }
  catch(e) {return {status:'ERROR',message:e.message};}
}
function getDashboardDataByJawatan(jawatan,tahun) {
  try {return gabungDashboardLatihan_(function(sh){return dashboardJawatanTab_(jawatan,tahun,sh);});}
  catch(e) {return {status:'ERROR',message:e.message};}
}
function getDashboardJawatanSahaja() {
  try {
    var jobs={}, years={};
    senaraiTabLatihan_().forEach(function(tab) {
      years[tab.year]=true;
      var res=jawatanTab_(tab.sheet);
      if (res.status !== 'SUCCESS') throw new Error(res.message);
      res.senaraiJawatan.forEach(function(j){jobs[j]=true;});
      res.tahunList.forEach(function(y){years[y]=true;});
    });
    return {status:'SUCCESS',senaraiJawatan:Object.keys(jobs).sort(),tahunList:Object.keys(years).sort(function(a,b){return b-a;})};
  } catch(e) {return {status:'ERROR',message:e.message};}
}
function getPortalTahun() {
  try {
    var years={}, tabs=senaraiTabLatihan_(false);
    tabs.forEach(function(tab) {
      years[String(tab.year)]=true;
      // Tab lama boleh mengandungi rekod sejarah. Baca SATU lajur tarikh sahaja.
      if(tab.year !== TAHUN_TAB_LAMA || tab.sheet.getLastRow()<2)return;
      var sh=tab.sheet, h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
      var d=colPortal_(h,['tarikh mula','tarikh mula kursus','tarikh kursus']);
      if(d<0)throw new Error('Lajur Tarikh Mula/Tarikh Kursus tidak ditemui dalam tab lama.');
      sh.getRange(2,d+1,sh.getLastRow()-1,1).getValues().forEach(function(row) {
        var year=yearPortal_(row[0]); if(year)years[String(year)]=true;
      });
    });
    var list=Object.keys(years).sort(function(a,b){return b-a;});
    var now=Utilities.formatDate(new Date(),'Asia/Kuala_Lumpur','yyyy');
    var preferred=list.indexOf(now)>=0 ? now : String(PORTAL_V2.year);
    if(list.indexOf(preferred)<0)preferred=list[0];
    return {status:'SUCCESS',tahunList:list,tahunPilihan:preferred};
  }catch(e){return {status:'ERROR',message:e.message};}
}
function tahunGrafLatihan_(tahun) {
  var n=Number(tahun || PORTAL_V2.year);
  if (!Number.isInteger(n) || n<2000 || n>2099) throw new Error('Tahun graf tidak sah.');
  return n;
}
function stafTahunLatihan_(year) {
  var n=year===Number(PORTAL_V2.year) ? PORTAL_V2.staff : PORTAL_STAF_TAHUN[year];
  if (typeof n !== 'number' || !Number.isInteger(n) || n<1)
    throw new Error('Jumlah staf tahun '+year+' belum ditetapkan. Isi PORTAL_STAF_TAHUN dalam Code.gs sebelum memaparkan peratus.');
  return n;
}
function dataGrafLatihan_(year) {
  var output=[['No. Kad Pengenalan','Tarikh Mula','Bilangan Jam Latihan','Jawatan']];
  senaraiTabLatihan_(false).forEach(function(tab) {
    var sh=tab.sheet, lastRow=sh.getLastRow();if(lastRow<2)return;
    var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var ic=colPortal_(h,['kad pengenalan','no. kp']), d=colPortal_(h,['tarikh mula','tarikh mula kursus','tarikh kursus']);
    var jam=colPortal_(h,['bilangan jam latihan','bilangan jam']), jaw=colPortal_(h,['perjawatan','jawatan']);
    if(ic<0 || d<0 || jam<0)throw new Error('Lajur IC, tarikh dan jam diperlukan dalam tab '+sh.getName()+'.');
    var cols=[ic,d,jam];if(jaw>=0)cols.push(jaw);
    var first=Math.min.apply(null,cols), last=Math.max.apply(null,cols);
    // Baca blok lajur yang diperlukan sekali; lajur penilaian tidak diperlukan untuk graf.
    var rows=sh.getRange(2,first+1,lastRow-1,last-first+1).getValues();
    for(var i=0;i<rows.length;i++) {
      var row=rows[i], date=row[d-first], recordYear=yearPortal_(date);
      if(recordYear===year || !recordYear)output.push([row[ic-first],date,row[jam-first],jaw>=0 ? row[jaw-first] : 'TIADA JAWATAN']);
    }
  });
  return output;
}
function grafJawatanLatihan_(year, data) {
  try {
    var totals={}, jobs={};
    for(var i=1;i<data.length;i++) {
      if(yearPortal_(data[i][1])!==year)continue;
      var ic=icPortal_(data[i][0]); if(!ic)continue;
      totals[ic]=(totals[ic] || 0)+(parseJamPortal_(data[i][2]) || 0);
      jobs[ic]=String(data[i][3] || 'TIADA JAWATAN').trim();
    }
    var counts={};
    Object.keys(totals).forEach(function(ic){if(totals[ic]>=40)counts[jobs[ic]]=(counts[jobs[ic]] || 0)+1;});
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sh=ss.getSheetByName('Bilangan Perjawatan '+year);
    if(!sh && year===TAHUN_TAB_LAMA)sh=ss.getSheetByName('Bilangan Perjawatan');
    if(!sh)throw new Error('Sediakan tab Bilangan Perjawatan '+year+' untuk bilangan staf mengikut jawatan tahun tersebut.');
    var rows=sh.getDataRange().getValues(), base={};
    for(var i=1;i<rows.length;i++)if(rows[i][0])base[String(rows[i][0]).trim()]=parseInt(rows[i][1]) || 0;
    var chart=Object.keys(base).sort().map(function(j) {
      var capai=counts[j] || 0, total=base[j];
      return {jawatan:j,capai40:capai,jumlahJawatan:total,peratus:total>0 ? Math.round(capai/total*100) : 0};
    });
    return {status:'SUCCESS',tahun:String(year),chartData:chart};
  }catch(e){return {status:'ERROR',message:e.message};}
}

function getDashboardChartLanding(tahun) {
  try {var year=tahunGrafLatihan_(tahun);return grafJawatanLatihan_(year,dataGrafLatihan_(year));}
  catch(e){return {status:'ERROR',message:e.message};}
}
function getPortalDashboard(tahun) {
  try {
    var year=tahunGrafLatihan_(tahun), data=dataGrafLatihan_(year);
    return {status:'SUCCESS',tahun:String(year),jawatan:grafJawatanLatihan_(year,data),kategori:grafKategoriLatihan_(year,data)};
  }catch(e){return {status:'ERROR',message:e.message};}
}

// HRMIS: akses import hanya melalui sesi admin yang disahkan di server.
var HRMIS_HEADERS_=['Tahun','No. Kad Pengenalan','Nama HRMIS','Jawatan HRMIS','Unit HRMIS','Jumlah Minit','Kursus Minit','Sesi Minit','Kendiri Minit','Tarikh Laporan','Sumber','Dikemas Kini'];
function hrmisHash_(value){return capRekodLatihan_([String(value)]);}
function hrmisKey_(part){return 'hrmis:'+SpreadsheetApp.getActiveSpreadsheet().getId()+':'+part;}
function hrmisSecret_(){
  var secret=PropertiesService.getScriptProperties().getProperty('HRMIS_ADMIN_PASSWORD') || '';
  if(secret.length<12)throw new Error('Tetapkan HRMIS_ADMIN_PASSWORD sekurang-kurangnya 12 aksara dalam Project Settings → Script properties dahulu.');
  return secret;
}
function hrmisEqual_(a,b){var diff=a.length^b.length;for(var i=0;i<Math.max(a.length,b.length);i++)diff|=(a.charCodeAt(i)||0)^(b.charCodeAt(i)||0);return diff===0;}
function adminHRMISLogin(password){
  try{return denganKunciLatihan_(function(){
    var cache=CacheService.getScriptCache(),key=hrmisKey_('login-attempts'),attempts=Number(cache.get(key)||0);
    if(attempts>=10)throw new Error('Terlalu banyak cubaan log masuk. Cuba semula selepas 5 minit.');
    var secret=hrmisSecret_();
    if(typeof password!=='string' || password.length>512 || !hrmisEqual_(hrmisHash_(password),hrmisHash_(secret))){cache.put(key,String(attempts+1),300);throw new Error('Kata laluan admin tidak sah.');}
    cache.remove(key);
    var token=Utilities.getUuid()+Utilities.getUuid();
    cache.put(hrmisKey_('session:'+hrmisHash_(token)),hrmisHash_(secret),1800);
    return {success:true,token:token,message:'Log masuk admin berjaya. Sesi sah sehingga 30 minit.'};
  });}catch(e){return {success:false,message:e.message};}
}
function semakAdminHRMIS_(token){
  if(typeof token!=='string' || token.length>150 || token.length<32)throw new Error('Log masuk admin diperlukan.');
  var credential=CacheService.getScriptCache().get(hrmisKey_('session:'+hrmisHash_(token)));
  if(!credential || !hrmisEqual_(credential,hrmisHash_(hrmisSecret_())))throw new Error('Sesi admin tamat atau tidak sah. Log masuk semula.');
  return hrmisHash_(token);
}
function adminHRMISLogout(token){if(typeof token==='string'&&token.length<=150)CacheService.getScriptCache().remove(hrmisKey_('session:'+hrmisHash_(token)));return {success:true};}
function teksHRMIS_(value){return String(value==null?'':value).normalize('NFKC').replace(/\s+/g,' ').trim();}
function normalHRMIS_(value){return teksHRMIS_(value).toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();}
function minitHRMIS_(value){
  if(typeof value==='number')return Number.isFinite(value)&&value>=0?Math.round(value*60):null;
  var s=teksHRMIS_(value).toLowerCase(),m;
  if((m=s.match(/^(\d+)\s*jam\s*(\d+)\s*minit$/)))return +m[2]<60?+m[1]*60+(+m[2]):null;
  if((m=s.match(/^(\d+)\s*jam$/)))return +m[1]*60;
  if((m=s.match(/^(\d+):(\d{2})$/)))return +m[2]<60?+m[1]*60+(+m[2]):null;
  if(/^\d+(?:[.,]\d+)?$/.test(s))return Math.round(Number(s.replace(',','.'))*60);
  return null;
}
function isoTarikhHRMIS_(value){
  if(value instanceof Date)return Utilities.formatDate(value,'Asia/Kuala_Lumpur','yyyy-MM-dd');
  var s=teksHRMIS_(value),m;if((m=s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/)))s=m[3]+'-'+m[2]+'-'+m[1];
  if(!/^\d{4}-\d{2}-\d{2}$/.test(s) || !yearPortal_(s))throw new Error('Tarikh laporan tidak sah.');return s;
}
function tajukHRMIS_(row){
  var h=row.map(normalHRMIS_),identity=h.findIndex(function(x){return /^(no|nombor) (kp|kad pengenalan|ic)( nama)?$/.test(x);});
  var total=h.indexOf('jumlah jam');
  if(identity<0 || total<0)return null;
  return {identity:identity,name:h.indexOf('nama'),job:h.indexOf('jawatan'),total:total,
    course:h.indexOf('kursus latihan'),session:h.indexOf('sesi pembelajaran'),self:h.indexOf('pembelajaran kendiri')};
}
function parseLaporanHRMIS_(books,year,reportDate){
  var map={},issues=[],duplicates=0,inputRows=0,periods=[],sources=[];
  books.forEach(function(book){
    if(!book || !Array.isArray(book.rows) || book.rows.length>15000)throw new Error('Struktur/baris laporan tidak sah.');
    var source=teksHRMIS_(book.sourceName).slice(0,180)+' / '+teksHRMIS_(book.sheetName).slice(0,80),head=null,unit='',count=0;
    sources.push(source);
    book.rows.forEach(function(row,index){
      if(!Array.isArray(row) || row.length>100)throw new Error('Laporan mempunyai format lajur yang tidak disokong.');
      var joined=row.map(teksHRMIS_).join(' '),period=joined.match(/BERMULA\s+DARI\s+(\d{2}\/\d{2}\/\d{4})\s+HINGGA\s+(\d{2}\/\d{2}\/\d{4})/i);
      if(period){var start=isoTarikhHRMIS_(period[1]),end=isoTarikhHRMIS_(period[2]);periods.push({start:start,end:end});if(start!==year+'-01-01'||end!==reportDate)issues.push(source+': tempoh laporan mesti 01/01/'+year+' hingga tarikh laporan yang dipilih.');}
      var found=tajukHRMIS_(row);if(found){head=found;return;}
      if(!head){if(/^(KLINIK|UNIT|PEJABAT|BAHAGIAN|MAKMAL)\b/i.test(teksHRMIS_(row[1])))unit=teksHRMIS_(row[1]);return;}
      var identity=teksHRMIS_(row[head.identity]),total=teksHRMIS_(row[head.total]);
      var match=identity.match(/^([\d\s-]+)\s*\/\s*(.+)$/),rawIC=match?match[1]:identity,ic=rawIC.replace(/[-\s]/g,'');
      if(!/^\d{12}$/.test(ic)){
        if(total && identity && !/^(jumlah|bil|no\b|nama\b|muka surat)/i.test(identity))issues.push(source+' baris '+(index+1)+': IC bukan 12 digit; rekod tidak diimport.');
        if(/^(KLINIK|UNIT|PEJABAT|BAHAGIAN|MAKMAL)\b/i.test(teksHRMIS_(row[1])))unit=teksHRMIS_(row[1]);return;
      }
      count++;inputRows++;
      var minutes=minitHRMIS_(row[head.total]),name=match?match[2]:head.name>=0?teksHRMIS_(row[head.name]):'';
      var parts=[head.course,head.session,head.self].map(function(i){return i>=0?minitHRMIS_(row[i]):null;});
      if(minutes==null || minutes>527040){issues.push(source+' baris '+(index+1)+': jumlah jam tidak sah.');return;}
      if(parts.every(function(n){return n!=null;}) && parts.reduce(function(a,b){return a+b;},0)!==minutes){issues.push(source+' baris '+(index+1)+': jumlah tiga kategori tidak sepadan dengan Jumlah Jam.');return;}
      if(!name){issues.push(source+' baris '+(index+1)+': nama staf kosong.');return;}
      var rec={ic:ic,name:name,job:head.job>=0?teksHRMIS_(row[head.job]):'',unit:unit,minutes:minutes,parts:parts,source:source};
      if(map[ic]){
        duplicates++;
        if(map[ic].minutes!==minutes || map[ic].parts.some(function(v,i){return v!=null && parts[i]!=null && v!==parts[i];}))issues.push('IC '+ic+': jumlah jam berbeza antara laporan. Pilih satu laporan yang betul.');
        if(map[ic].source.indexOf(source)<0)map[ic].source+='; '+source;
      }else map[ic]=rec;
    });
    if(!count)issues.push(source+': tiada rekod IC/jam yang dapat dikenal pasti.');
  });
  return {records:Object.keys(map).sort().map(function(ic){return map[ic];}),issues:Array.from(new Set(issues)),duplicates:duplicates,inputRows:inputRows,periods:periods,sources:sources};
}
function bacaJadualHRMIS_(){
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HRMIS_JAM'),rows=[],map={};
  if(sheet && sheet.getLastRow()>0){
    rows=sheet.getRange(1,1,sheet.getLastRow(),Math.max(sheet.getLastColumn(),HRMIS_HEADERS_.length)).getValues();
    if(rows[0].length!==HRMIS_HEADERS_.length || rows[0].some(function(v,i){return teksHRMIS_(v)!==HRMIS_HEADERS_[i];}))throw new Error('Struktur tab HRMIS_JAM tidak sepadan. Semak tajuk lajur sebelum import.');
  }else rows=[HRMIS_HEADERS_.slice()];
  for(var i=1;i<rows.length;i++){
    if(rows[i].every(function(v){return v===''||v==null;}))continue;
    var y=Number(rows[i][0]),ic=icPortal_(rows[i][1]),n=Number(rows[i][5]);
    if(!Number.isInteger(y)||!/^\d{12}$/.test(ic)||!Number.isInteger(n)||n<0)throw new Error('Data HRMIS_JAM tidak sah pada baris '+(i+1)+'.');
    var key=y+'|'+ic;if(map[key])throw new Error('IC/tahun berulang dalam HRMIS_JAM pada baris '+(i+1)+'.');
    map[key]={row:i,values:rows[i]};
  }
  return {sheet:sheet,rows:rows,map:map,hash:capRekodLatihan_([!!sheet,rows])};
}
function indeksStafHRMIS_(){
  var map={};senaraiTabLatihan_(false).forEach(function(tab){
    var sh=tab.sheet,n=sh.getLastRow();if(n<2)return;
    var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],ic=colPortal_(h,['kad pengenalan','no. kp']),name=colPortal_(h,['nama penuh','nama pegawai']);
    if(ic<0||name<0)return;var first=Math.min(ic,name),last=Math.max(ic,name);
    sh.getRange(2,first+1,n-1,last-first+1).getValues().forEach(function(r){var id=icPortal_(r[ic-first]);if(/^\d{12}$/.test(id))map[id]=teksHRMIS_(r[name-first]);});
  });return map;
}
function booksPautanHRMIS_(links){
  if(!Array.isArray(links)||links.length>4)throw new Error('Maksimum empat pautan Google Sheet.');
  var books=[],seen={};links.forEach(function(url){
    var m=String(url).trim().match(/^https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)(?:[\/?#].*)?$/);
    if(!m)throw new Error('Gunakan pautan Google Sheet. Untuk fail XLSX, gunakan pilihan upload.');if(seen[m[1]])return;seen[m[1]]=true;
    var ss;try{ss=SpreadsheetApp.openById(m[1]);}catch(e){throw new Error('Google Sheet tidak boleh dibaca. Pastikan akaun yang deploy portal mempunyai akses.');}
    var sheets=ss.getSheets();if(sheets.length>30)throw new Error('Terlalu banyak tab dalam Google Sheet laporan.');
    sheets.forEach(function(sh){if(sh.getLastRow()>15000||sh.getLastColumn()>100)throw new Error('Laporan terlalu besar.');if(sh.getLastRow())books.push({sourceName:ss.getName(),sheetName:sh.getName(),rows:sh.getDataRange().getDisplayValues()});});
  });return books;
}
function simpanPreviewHRMIS_(owner,plan){
  var id=Utilities.getUuid(),text=JSON.stringify(plan),chunks=[],cache=CacheService.getScriptCache();
  for(var i=0;i<text.length;i+=20000)chunks.push(text.slice(i,i+20000));
  if(chunks.length>40)throw new Error('Pratonton terlalu besar. Import laporan secara berasingan.');
  var entries={};chunks.forEach(function(s,i){entries[hrmisKey_('preview:'+id+':'+i)]=s;});
  cache.putAll(entries,600);cache.put(hrmisKey_('preview:'+id),JSON.stringify({owner:owner,chunks:chunks.length}),600);return id;
}
function adminHRMISPreview(token,payloadJson){
  try{
    var owner=semakAdminHRMIS_(token);
    if(typeof payloadJson!=='string'||payloadJson.length>3000000)throw new Error('Saiz input laporan melebihi had.');
    var input=JSON.parse(payloadJson),year=Number(input.year);
    var books=input.books||[];if(!Array.isArray(books)||books.length>30)throw new Error('Terlalu banyak tab/fail.');
    books=books.concat(booksPautanHRMIS_(input.links||[]));if(!books.length)throw new Error('Pilih fail XLSX atau masukkan pautan Google Sheet.');
    var detected={};books.forEach(function(book){if(book&&Array.isArray(book.rows))book.rows.slice(0,100).forEach(function(row){if(!Array.isArray(row))return;var m=row.map(teksHRMIS_).join(' ').match(/HINGGA\s+(\d{2}\/\d{2}\/\d{4})/i);if(m)detected[isoTarikhHRMIS_(m[1])]=true;});});
    var dates=Object.keys(detected),date=isoTarikhHRMIS_(input.reportDate || (dates.length===1?dates[0]:''));
    if(!Number.isInteger(year)||year<2000||year>2099||Number(date.slice(0,4))!==year)throw new Error('Tahun dan tarikh laporan mesti sepadan.');
    var parsed=parseLaporanHRMIS_(books,year,date),table=bacaJadualHRMIS_(),staff=indeksStafHRMIS_(),counts={baru:0,kemasKini:0,sama:0,padan:0,tiadaPortal:0},buckets=[0,0,0,0];
    var preview=parsed.records.map(function(r){
      var old=table.map[year+'|'+r.ic],oldDate=old?isoTarikhHRMIS_(old.values[9]):null;
      if(oldDate && oldDate>date)parsed.issues.push('IC '+r.ic+': laporan '+date+' lebih lama daripada rekod '+oldDate+'.');
      var change=!old?'Baru':Number(old.values[5])===r.minutes&&oldDate===date?'Sama':'Kemas kini';
      if(!old)counts.baru++;else if(change==='Sama')counts.sama++;else counts.kemasKini++;
      if(staff[r.ic])counts.padan++;else counts.tiadaPortal++;
      buckets[r.minutes===0?0:r.minutes<1200?1:r.minutes<2400?2:3]++;
      return {ic:r.ic,name:r.name,job:r.job,unit:r.unit,minutes:r.minutes,oldMinutes:old?Number(old.values[5]):null,oldDate:oldDate,matched:!!staff[r.ic],portalName:staff[r.ic]||'',change:change};
    });
    if(!parsed.records.length)parsed.issues.push('Tiada staf untuk diimport.');
    var issues=Array.from(new Set(parsed.issues)),plan={year:year,date:date,records:parsed.records,tableHash:table.hash,counts:counts,buckets:buckets,duplicates:parsed.duplicates};
    var id=issues.length?'':simpanPreviewHRMIS_(owner,plan);
    return {success:true,canConfirm:!issues.length,previewId:id,year:year,reportDate:date,count:preview.length,duplicates:parsed.duplicates,inputRows:parsed.inputRows,counts:counts,buckets:buckets,issues:issues,records:preview};
  }catch(e){return {success:false,message:e.message};}
}
function teksSelHRMIS_(s){s=teksHRMIS_(s).slice(0,2000);return /^[=+\-@]/.test(s)?"'"+s:s;}
function adminHRMISConfirm(token,id){
  try{return denganKunciLatihan_(function(){
    var owner=semakAdminHRMIS_(token),cache=CacheService.getScriptCache();
    if(typeof id!=='string'||!/^[-a-zA-Z0-9]{10,80}$/.test(id))throw new Error('Pratonton tidak sah.');
    var key=hrmisKey_('preview:'+id),raw=cache.get(key);if(!raw)throw new Error('Pratonton tamat. Klik Semak import semula.');
    var meta=JSON.parse(raw);if(meta.owner!==owner)throw new Error('Pratonton bukan milik sesi admin ini.');if(meta.done)return meta.result;
    var parts=[];for(var i=0;i<meta.chunks;i++){var part=cache.get(key+':'+i);if(part==null)throw new Error('Pratonton tamat. Klik Semak import semula.');parts.push(part);}
    var plan=JSON.parse(parts.join('')),table=bacaJadualHRMIS_();if(table.hash!==plan.tableHash)throw new Error('HRMIS_JAM berubah selepas pratonton. Semak import semula sebelum mengesahkan.');
    var rows=table.rows.map(function(r){return r.slice();}),now=new Date().toISOString();
    plan.records.forEach(function(r){
      var row=[plan.year,r.ic,teksSelHRMIS_(r.name),teksSelHRMIS_(r.job),teksSelHRMIS_(r.unit),r.minutes,r.parts[0]==null?'':r.parts[0],r.parts[1]==null?'':r.parts[1],r.parts[2]==null?'':r.parts[2],plan.date,teksSelHRMIS_(r.source),now];
      var old=table.map[plan.year+'|'+r.ic];if(old)rows[old.row]=row;else rows.push(row);
    });
    var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=table.sheet||ss.insertSheet('HRMIS_JAM');
    if(sh.getMaxColumns()<HRMIS_HEADERS_.length)sh.insertColumnsAfter(sh.getMaxColumns(),HRMIS_HEADERS_.length-sh.getMaxColumns());
    if(sh.getMaxRows()<rows.length)sh.insertRowsAfter(sh.getMaxRows(),rows.length-sh.getMaxRows());
    sh.getRange(2,2,Math.max(1,rows.length-1),1).setNumberFormat('@');sh.getRange(2,10,Math.max(1,rows.length-1),1).setNumberFormat('@');
    sh.getRange(1,1,rows.length,HRMIS_HEADERS_.length).setValues(rows);sh.setFrozenRows(1);SpreadsheetApp.flush();
    var result={success:true,count:plan.records.length,year:plan.year,reportDate:plan.date,message:plan.records.length+' staf disimpan dalam HRMIS_JAM. Jam diganti mengikut IC/tahun, bukan ditambah.'};
    cache.put(key,JSON.stringify({owner:owner,done:true,result:result}),600);return result;
  });}catch(e){return {success:false,message:e.message};}
}
// Public dashboards receive only the summaries belonging to staff already in that dashboard.
function sertakanJamHRMIS_(data){
  var table;try{table=bacaJadualHRMIS_();}catch(e){return;}
  Object.keys(data).forEach(function(job){Object.keys(data[job]).forEach(function(ic){
    var staff=data[job][ic];staff.hrmisByYear={};Object.keys(table.map).forEach(function(key){
      var item=table.map[key].values;if(icPortal_(item[1])===ic)staff.hrmisByYear[String(item[0])]={minutes:Number(item[5]),reportDate:isoTarikhHRMIS_(item[9])};
    });
  });});
}
