/**
 * Aplikasi Instrumen Bestari Penilaian PdP v2026.09
 * Dibangunkan oleh: Ustaz Zaimuddin Hassan (Guru Al-Quran & Bahasa Arab, SMK Padang Pak Amat, Pasir Puteh, Kelantan)
 * CC BY-NC-SA 2026 SolusiBestariGuru (t.me/solusibestariguru)
 *
 * Google Apps Script Server-Side Controller
 * Container Spreadsheet: Terikat dengan Google Sheet yang mengandungi tab:
 * 1. TETAPAN (Maklumat Sekolah & Konfigurasi URL Deployment)
 * 2. GURU (Lajur: Emel, Nama Guru, Jantina, Opsyen, Pencerap, Jawatan)
 * 3. PENGISIAN
 */

// ============================================================================
// 1. MENU PENTADBIR DALAM GOOGLE SHEET (TETAPAN Instrumen Bestari Penilaian PdP)
// ============================================================================

/**
 * Menu tersuai dipaparkan secara automatik dalam Google Sheet container.
 */
function onOpen(e) {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("⚙️ Instrumen Bestari PPdP")
      .addItem(
        "🚀 Setup Instrumen Bestari PPdP",
        "setupInstrumenBestari",
      )
      .addSeparator()
      .addItem(
        "📋 Buka Lembaran Tetapan (Maklumat Sekolah B5:B9)",
        "bukaLembaranTetapan",
      )
      .addItem(
        "🌐 Semak / Kemas Kini URL Web App",
        "konfigurasiUrlDeploymentNatif",
      )
      .addToUi();
  } catch (err) {
    // Abaikan jika dijalankan di luar persekitaran UI spreadsheet
  }
}

/**
 * 0. FUNGSI SETUP UTAMA:
 * Membaca sel yang mengandungi URL deployment dan maklumat sekolah
 * daripada tab TETAPAN (URL di B4, Butiran Sekolah di B5:B9)
 * lalu menyimpannya ke dalam ScriptProperties & DocumentProperties.
 */
function setupInstrumenBestari() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = getSpreadsheet_();
    if (!ss) throw new Error("Spreadsheet container tidak dapat dicapai.");

    // A. Pastikan atau baca tab TETAPAN untuk URL & Parameter
    var sheetTetapan = ss.getSheetByName("TETAPAN");
    if (!sheetTetapan || sheetTetapan.getLastRow() < 3) {
      sheetTetapan = sediakanLembaranTetapan_(ss);
    }

    var webAppUrl = "";

    if (sheetTetapan) {
      // 1. Baca nilai terus dari sel B4
      var b4Val = (sheetTetapan.getRange("B4").getValue() || "").toString().trim();
      if (b4Val && b4Val !== "-") {
        webAppUrl = b4Val;
      }

      // 2. Sandaran: Padankan berdasarkan lajur A (kunci parameter) jika kedudukan baris berbeza
      if (!webAppUrl) {
        var dataTetapan = sheetTetapan.getDataRange().getValues();
        for (var i = 0; i < dataTetapan.length; i++) {
          var k = dataTetapan[i][0];
          var v = (dataTetapan[i][1] || "").toString().trim();
          if (normalizeConfigKey_(k) === "WEBAPP_URL") {
            if (v && v !== "-") {
              webAppUrl = v;
              break;
            }
          }
        }
      }

      // 3. Jika sel B4 kosong tetapi pernah disimpan dalam ScriptProperties, pulihkan ke sel B4
      if (!webAppUrl) {
        try {
          var spOld = PropertiesService.getScriptProperties();
          var savedOld = spOld ? (spOld.getProperty("WEBAPP_URL") || "").trim() : "";
          if (savedOld && savedOld !== "-") {
            webAppUrl = savedOld;
            sheetTetapan.getRange("B4").setValue(webAppUrl);
          }
        } catch (eOld) {}
      }
    }

    // B. Baca maklumat sekolah daripada tab TETAPAN (Owner mengisi di range B5:B9)
    // Susunan baris:
    // B5 = KOD_SEKOLAH
    // B6 = NAMA_SEKOLAH
    // B7 = NEGERI_SEKOLAH
    // B8 = PPD_SEKOLAH
    // B9 = EMEL_SEKOLAH
    var sekolahObj = {
      kodSekolah: "",
      namaSekolah: "",
      negeri: "",
      ppd: "",
      emelSekolah: "",
    };

    if (sheetTetapan) {
      // 1. Baca nilai terus dari range B5:B9
      if (sheetTetapan.getLastRow() >= 5) {
        var dataSekolah = sheetTetapan.getRange("B5:B9").getValues();
        sekolahObj.kodSekolah = (dataSekolah[0][0] || "").toString().trim();
        sekolahObj.namaSekolah = (dataSekolah[1] && dataSekolah[1][0] ? dataSekolah[1][0] : "").toString().trim();
        sekolahObj.negeri = (dataSekolah[2] && dataSekolah[2][0] ? dataSekolah[2][0] : "").toString().trim();
        sekolahObj.ppd = (dataSekolah[3] && dataSekolah[3][0] ? dataSekolah[3][0] : "").toString().trim();
        sekolahObj.emelSekolah = (dataSekolah[4] && dataSekolah[4][0] ? dataSekolah[4][0] : "").toString().trim();
      }

      // 2. Sokongan sandaran: padankan berdasarkan lajur A (kunci parameter) jika kedudukan baris berbeza
      if (!sekolahObj.kodSekolah && !sekolahObj.namaSekolah) {
        var dataTetapanAll = sheetTetapan.getDataRange().getValues();
        for (var r = 0; r < dataTetapanAll.length; r++) {
          var kunci = normalizeConfigKey_(dataTetapanAll[r][0]);
          var nilai = (dataTetapanAll[r][1] || "").toString().trim();
          if (kunci === "KOD_SEKOLAH") sekolahObj.kodSekolah = nilai;
          else if (kunci === "NAMA_SEKOLAH") sekolahObj.namaSekolah = nilai;
          else if (kunci === "NEGERI_SEKOLAH") sekolahObj.negeri = nilai;
          else if (kunci === "PPD_SEKOLAH") sekolahObj.ppd = nilai;
          else if (kunci === "EMEL_SEKOLAH") sekolahObj.emelSekolah = nilai;
        }
      }
    }

    // C. Simpan ke dalam Tab TETAPAN, ScriptProperties dan DocumentProperties
    var propsToSave = {
      WEBAPP_URL: webAppUrl,
      KOD_SEKOLAH: sekolahObj.kodSekolah,
      NAMA_SEKOLAH: sekolahObj.namaSekolah,
      NEGERI_SEKOLAH: sekolahObj.negeri,
      PPD_SEKOLAH: sekolahObj.ppd,
      EMEL_SEKOLAH: sekolahObj.emelSekolah,
      SEKOLAH_JSON: JSON.stringify(sekolahObj),
    };

    setConfigValues_(propsToSave);

    // Pastikan WEBAPP_URL disimpan secara eksplisit dan teguh ke ScriptProperties
    try {
      var spExplicit = PropertiesService.getScriptProperties();
      if (spExplicit) {
        if (webAppUrl) {
          spExplicit.setProperty("WEBAPP_URL", webAppUrl);
        } else {
          spExplicit.deleteProperty("WEBAPP_URL");
        }
      }
    } catch (eSpExplicit) {}

    // D. Padam sebarang kunci legasi daripada Properties
    hapusKunciKonfigurasiLegasi_();

    // D. Segerakkan senarai pencerap yang sah daripada tab GURU ke Properties untuk log masuk ultra-pantas
    var mapPencerap = segerakPencerapKeProperties_(ss);
    var bilPencerap = Object.keys(mapPencerap || {}).length;

    // E. Segerakkan semua teks literal dari tab-tab Spreadsheet ke ScriptProperties
    var hasilLiteral = segerakLiteralTeksKeProperties_(ss);

    // F. Padam trigger onEdit sedia ada & pasangkan semula trigger onEdit (Installable Trigger)
    var triggerStatus = "Belum dipasang";
    try {
      var allTriggers = ScriptApp.getProjectTriggers();
      for (var t = 0; t < allTriggers.length; t++) {
        if (allTriggers[t].getHandlerFunction() === "onEdit") {
          ScriptApp.deleteTrigger(allTriggers[t]);
        }
      }
      ScriptApp.newTrigger("onEdit").forSpreadsheet(ss).onEdit().create();
      triggerStatus = "✅ Berjaya dipasang (Installable Trigger)";
    } catch (errTrig) {
      if (
        errTrig.message &&
        errTrig.message.indexOf("script.scriptapp") !== -1
      ) {
        triggerStatus =
          "⚠️ Kebenaran Diperlukan: Sila sahkan tetingkap kebenaran (Authorization Required) Google atau jalankan semula Setup.";
      } else {
        triggerStatus = "⚠️ Gagal dipasang: " + errTrig.message;
      }
    }

    // Kosongkan cache
    try {
      CacheService.getScriptCache().remove("PUBLIC_FORM_DATA");
      CacheService.getScriptCache().remove("ISPPK_LITERAL_TEXTS");
    } catch (e) {}

    var mesej =
      "✅ Setup Instrumen Bestari PPdP Berjaya Diselaraskan!\n\n" +
      "Seni Bina: 1 Deployment Sahaja (Owner Relay - DELIMa Domain)\n\n" +
      "Data berikut telah dibaca dari tab TETAPAN (B5:B9) & tab GURU, lalu disegerakkan ke Properties:\n\n" +
      "• Kod Sekolah : " +
      (sekolahObj.kodSekolah || "(Kosong)") +
      "\n" +
      "• Nama Sekolah : " +
      (sekolahObj.namaSekolah || "(Kosong)") +
      "\n" +
      "• Negeri : " +
      (sekolahObj.negeri || "(Kosong)") +
      "\n" +
      "• PPD : " +
      (sekolahObj.ppd || "(Kosong)") +
      "\n" +
      "• Emel Sekolah : " +
      (sekolahObj.emelSekolah || "(Kosong)") +
      "\n\n" +
      "• Pencerap Berdaftar : " +
      bilPencerap +
      " orang (Disimpan ke Properties)\n" +
      "• Teks Literal Rubrik & Pilihan : Disegerakkan ke ScriptProperties (" +
      hasilLiteral.bilTabDisegerak +
      " tab)\n" +
      "• URL Web App (B4) : " +
      (webAppUrl ? webAppUrl : "⚠️ KOSONG (Sila salin dari Apps Script UI: Deploy > Manage deployments dan tampal ke sel B4)") +
      "\n\n" +
      "• Trigger onEdit : " +
      triggerStatus;

    if (!webAppUrl) {
      mesej +=
        "\n\n⚠️ PERINGATAN PENTING:\n" +
        "Sel B4 tab TETAPAN masih kosong. Sila salin Web app URL dari Apps Script UI (Deploy > Manage deployments), tampal ke sel B4, dan jalankan Setup Instrumen Bestari sekali lagi.";
    }

    ui.alert("Setup Selesai", mesej, ui.ButtonSet.OK);
    return {
      success: true,
      sekolah: sekolahObj,
      urls: { webAppUrl: webAppUrl, ownerUrl: webAppUrl },
      literalTexts: hasilLiteral,
      trigger: triggerStatus,
    };
  } catch (err) {
    ui.alert(
      "Ralat Setup",
      "Gagal menjalankan Setup: " + err.message,
      ui.ButtonSet.OK,
    );
    return { success: false, error: err.message };
  }
}


// ============================================================================
// 1.1 PENGURUSAN TEKS LITERAL & PENYELARASAN TAB GOOGLE SHEET KE SCRIPTPROPERTIES
// ============================================================================

/**
 * Membaca tab-tab teks literal daripada Spreadsheet container dan menyimpannya
 * ke dalam ScriptProperties menggunakan seni bina storan bersegmen (segmented storage)
 * bagi mengelakkan had saiz 9KB per entri Google Apps Script.
 * 
 * Tab-tab yang disokong:
 * 1. RUBRIK_GURU (Item rubrik domain 1, 2, 3 beserta skala 1-5)
 * 2. SKALA_GURU (Tahap penguasaan guru: Kesedaran, Asas, Pertengahan, Lanjutan, Pakar)
 * 3. ITEM_MURID (10 item instrumen murid & label ringkas carta)
 * 4. SKALA_MURID (Huraian skala peratusan murid 1-5)
 * 5. TAHAP_KBAT (Piawaian 5 tahap pencapaian KBAT, julat skor, ringkasan & warna)
 * 6. INFO_INSTRUMEN (Tujuan, Objektif, Aspek yang dicerap)
 * 7. BILANGAN_MURID (Pilihan dropdown bilangan murid hadir)
 * 8. TAHUN_TINGKATAN (Pilihan dropdown tahun / tingkatan)
 * 9. SENARAI_NEGERI (Senarai 16 negeri & Wilayah Persekutuan)
 */
function segerakLiteralTeksKeProperties_(ss) {
  ss = ss || getSpreadsheet_();
  var props = PropertiesService.getScriptProperties();
  var bilTabDisegerak = 0;
  var errors = [];

  // 1. RUBRIK_GURU
  try {
    var sheetRubrik = ss ? ss.getSheetByName("RUBRIK_GURU") : null;
    var domainMap = {
      perancangan: { domain: "1. PERANCANGAN (Skor Maksimum: 20)", domainKey: "perancangan", items: [] },
      pelaksanaan: { domain: "2. PELAKSANAAN (Skor Maksimum: 25)", domainKey: "pelaksanaan", items: [] },
      refleksi: { domain: "3. REFLEKSI (Skor Maksimum: 5)", domainKey: "refleksi", items: [] }
    };
    var adaDataRubrik = false;

    if (sheetRubrik && sheetRubrik.getLastRow() >= 2) {
      var rubrikData = sheetRubrik.getDataRange().getValues();
      for (var r = 1; r < rubrikData.length; r++) {
        var row = rubrikData[r];
        var dKey = (row[0] || "").toString().trim().toLowerCase();
        var dNama = (row[1] || "").toString().trim();
        var kod = (row[2] || "").toString().trim();
        var tajuk = (row[3] || "").toString().trim();
        var aspekRaw = (row[4] || "").toString().trim();
        var sk1 = (row[5] || "").toString().trim();
        var sk2 = (row[6] || "").toString().trim();
        var sk3 = (row[7] || "").toString().trim();
        var sk4 = (row[8] || "").toString().trim();
        var sk5 = (row[9] || "").toString().trim();

        if (kod && tajuk) {
          if (!domainMap[dKey]) {
            domainMap[dKey] = {
              domain: dNama || dKey.toUpperCase(),
              domainKey: dKey,
              items: []
            };
          } else if (dNama) {
            domainMap[dKey].domain = dNama;
          }

          var aspekArr = aspekRaw
            ? aspekRaw.split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean)
            : [];

          domainMap[dKey].items.push({
            kod: kod,
            tajuk: tajuk,
            aspek: aspekArr,
            skala: { 1: sk1, 2: sk2, 3: sk3, 4: sk4, 5: sk5 }
          });
          adaDataRubrik = true;
        }
      }
    }

    if (adaDataRubrik) {
      // Simpan bersegmen mengikut domain (saiz setiap segmen < 5.2 KB, mematuhi had 9KB GAS)
      props.setProperty("ISPPK_RUBRIK_PERANCANGAN", JSON.stringify(domainMap.perancangan || {}));
      props.setProperty("ISPPK_RUBRIK_PELAKSANAAN", JSON.stringify(domainMap.pelaksanaan || {}));
      props.setProperty("ISPPK_RUBRIK_REFLEKSI", JSON.stringify(domainMap.refleksi || {}));
      bilTabDisegerak++;
    }
  } catch (eRubrik) {
    errors.push("RUBRIK_GURU: " + eRubrik.message);
  }

  // 2. SKALA_GURU
  try {
    var sheetSkalaGuru = ss ? ss.getSheetByName("SKALA_GURU") : null;
    var listSG = {};
    var adaDataSG = false;
    if (sheetSkalaGuru && sheetSkalaGuru.getLastRow() >= 2) {
      var dataSG = sheetSkalaGuru.getDataRange().getValues();
      for (var r = 1; r < dataSG.length; r++) {
        var row = dataSG[r];
        var skor = Number(row[0]) || r;
        var label = (row[1] || "").toString().trim();
        var labelCaps = (row[2] || label.toUpperCase()).toString().trim();
        var ket = (row[3] || "").toString().trim();
        if (skor && label) {
          listSG[skor] = { skor: skor, label: label, label_huruf_besar: labelCaps, keterangan: ket };
          adaDataSG = true;
        }
      }
    }
    if (adaDataSG) {
      props.setProperty("ISPPK_SKALA_GURU", JSON.stringify(listSG));
      bilTabDisegerak++;
    }
  } catch (eSG) {
    errors.push("SKALA_GURU: " + eSG.message);
  }

  // 3. ITEM_MURID
  try {
    var sheetMurid = ss ? ss.getSheetByName("ITEM_MURID") : null;
    var listMurid = [];
    if (sheetMurid && sheetMurid.getLastRow() >= 2) {
      var dataMurid = sheetMurid.getDataRange().getValues();
      for (var r = 1; r < dataMurid.length; r++) {
        var row = dataMurid[r];
        var no = Number(row[0]) || r;
        var kod = (row[1] || "M" + no).toString().trim();
        var lbl = (row[2] || "Item " + no).toString().trim();
        var txt = (row[3] || "").toString().trim();
        if (txt) {
          listMurid.push({ no: no, kod: kod, label_pendek: lbl, text: txt });
        }
      }
    }
    if (listMurid.length > 0) {
      props.setProperty("ISPPK_ITEM_MURID", JSON.stringify(listMurid));
      bilTabDisegerak++;
    }
  } catch (eMurid) {
    errors.push("ITEM_MURID: " + eMurid.message);
  }

  // 4. SKALA_MURID
  try {
    var sheetSkalaMurid = ss ? ss.getSheetByName("SKALA_MURID") : null;
    var mapSM = {};
    var adaDataSM = false;
    if (sheetSkalaMurid && sheetSkalaMurid.getLastRow() >= 2) {
      var dataSM = sheetSkalaMurid.getDataRange().getValues();
      for (var r = 1; r < dataSM.length; r++) {
        var row = dataSM[r];
        var sNum = Number(row[0]) || r;
        var desc = (row[2] || row[1] || "").toString().trim();
        if (desc) {
          mapSM[sNum] = desc;
          adaDataSM = true;
        }
      }
    }
    if (adaDataSM) {
      props.setProperty("ISPPK_SKALA_MURID", JSON.stringify(mapSM));
      bilTabDisegerak++;
    }
  } catch (eSM) {
    errors.push("SKALA_MURID: " + eSM.message);
  }

  // 5. TAHAP_KBAT
  try {
    var sheetTahap = ss ? ss.getSheetByName("TAHAP_KBAT") : null;
    var listTahap = [];
    if (sheetTahap && sheetTahap.getLastRow() >= 2) {
      var dataTahap = sheetTahap.getDataRange().getValues();
      for (var r = 1; r < dataTahap.length; r++) {
        var row = dataTahap[r];
        var tId = Number(row[0]);
        var sMin = Number(row[1]) || 0;
        var sMax = Number(row[2]) || 0;
        var jLbl = (row[3] || "").toString().trim();
        var ringkasan = (row[4] || "").toString().trim();
        var lencana = (row[5] || ringkasan).toString().trim();
        var warna = (row[6] || "slate").toString().trim();
        var desc = (row[7] || "").toString().trim();
        if (ringkasan) {
          listTahap.push({
            tahap_id: tId,
            skor_min: sMin,
            skor_max: sMax,
            julat_label: jLbl,
            ringkasan: ringkasan,
            status_lencana: lencana,
            warna_tema: warna,
            penerangan: desc,
          });
        }
      }
    }
    if (listTahap.length > 0) {
      props.setProperty("ISPPK_TAHAP_KBAT", JSON.stringify(listTahap));
      bilTabDisegerak++;
    }
  } catch (eTahap) {
    errors.push("TAHAP_KBAT: " + eTahap.message);
  }

  // 6. INFO_INSTRUMEN
  try {
    var sheetInfo = ss ? ss.getSheetByName("INFO_INSTRUMEN") : null;
    var listInfo = [];
    if (sheetInfo && sheetInfo.getLastRow() >= 2) {
      var dataInfo = sheetInfo.getDataRange().getValues();
      for (var r = 1; r < dataInfo.length; r++) {
        var row = dataInfo[r];
        var kat = (row[0] || "").toString().trim();
        var urut = Number(row[1]) || r;
        var tajuk = (row[2] || "").toString().trim();
        var ket = (row[3] || "").toString().trim();
        if (kat && ket) {
          listInfo.push({ kategori: kat, urutan: urut, tajuk: tajuk, keterangan: ket });
        }
      }
    }
    if (listInfo.length > 0) {
      props.setProperty("ISPPK_INFO_INSTRUMEN", JSON.stringify(listInfo));
      bilTabDisegerak++;
    }
  } catch (eInfo) {
    errors.push("INFO_INSTRUMEN: " + eInfo.message);
  }

  // 7. BILANGAN_MURID
  try {
    var sheetBM = ss ? ss.getSheetByName("BILANGAN_MURID") : null;
    var listBM = [];
    if (sheetBM && sheetBM.getLastRow() >= 2) {
      var dataBM = sheetBM.getDataRange().getValues();
      for (var r = 1; r < dataBM.length; r++) {
        var row = dataBM[r];
        var bId = Number(row[0]) || r;
        var julat = (row[1] || "").toString().trim();
        var minM = Number(row[2]) || 0;
        var maxM = Number(row[3]) || 0;
        if (julat) listBM.push({ id: bId, julat: julat, min: minM, max: maxM });
      }
    }
    if (listBM.length > 0) {
      props.setProperty("ISPPK_BILANGAN_MURID", JSON.stringify(listBM));
      bilTabDisegerak++;
    }
  } catch (eBM) {
    errors.push("BILANGAN_MURID: " + eBM.message);
  }

  // 8. TAHUN_TINGKATAN
  try {
    var sheetTT = ss ? ss.getSheetByName("TAHUN_TINGKATAN") : null;
    var listTT = [];
    if (sheetTT && sheetTT.getLastRow() >= 2) {
      var dataTT = sheetTT.getDataRange().getValues();
      for (var r = 1; r < dataTT.length; r++) {
        var row = dataTT[r];
        var tId = Number(row[0]) || r;
        var kat = (row[1] || "").toString().trim();
        var nama = (row[2] || "").toString().trim();
        if (nama) listTT.push({ id: tId, kategori: kat, nama_pilihan: nama });
      }
    }
    if (listTT.length > 0) {
      props.setProperty("ISPPK_TAHUN_TINGKATAN", JSON.stringify(listTT));
      bilTabDisegerak++;
    }
  } catch (eTT) {
    errors.push("TAHUN_TINGKATAN: " + eTT.message);
  }

  // 9. SENARAI_NEGERI
  try {
    var sheetNegeri = ss ? ss.getSheetByName("SENARAI_NEGERI") : null;
    var listNegeri = [];
    if (sheetNegeri && sheetNegeri.getLastRow() >= 2) {
      var dataNegeri = sheetNegeri.getDataRange().getValues();
      for (var r = 1; r < dataNegeri.length; r++) {
        var nNama = (dataNegeri[r][1] || dataNegeri[r][0] || "").toString().trim();
        if (nNama && nNama.toLowerCase() !== "nama_negeri" && nNama.toLowerCase() !== "kod") {
          listNegeri.push(nNama);
        }
      }
    }
    if (listNegeri.length > 0) {
      props.setProperty("ISPPK_SENARAI_NEGERI", JSON.stringify(listNegeri));
      bilTabDisegerak++;
    }
  } catch (eNegeri) {
    errors.push("SENARAI_NEGERI: " + eNegeri.message);
  }

  if (bilTabDisegerak > 0) {
    props.setProperty("ISPPK_TEXTS_TIMESTAMP", Date.now().toString());
  }

  // Kosongkan CacheService
  try {
    CacheService.getScriptCache().remove("ISPPK_LITERAL_TEXTS");
  } catch (eCache) {}

  return {
    success: true,
    bilTabDisegerak: bilTabDisegerak,
    errors: errors,
  };
}

/**
 * Mengambil set lengkap teks literal daripada ScriptProperties (dengan sokongan Cache).
 * Nilai lalai TIDAK lagi disimpan di Code.gs; sebarang kunci yang tiada dalam ScriptProperties
 * akan dikembalikan sebagai null agar antaramuka Index.html menggunakan pemalar lalai (DEFAULT_*) miliknya.
 */
function dapatkanLiteralTeksDariProperties_() {
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get("ISPPK_LITERAL_TEXTS");
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  var props = PropertiesService.getScriptProperties();

  // 1. Rubrik Guru (gabungkan 3 domain)
  var rubrikGuru = null;
  try {
    var d1 = JSON.parse(props.getProperty("ISPPK_RUBRIK_PERANCANGAN") || "null");
    var d2 = JSON.parse(props.getProperty("ISPPK_RUBRIK_PELAKSANAAN") || "null");
    var d3 = JSON.parse(props.getProperty("ISPPK_RUBRIK_REFLEKSI") || "null");
    if (d1 && d1.items && d1.items.length && d2 && d2.items && d2.items.length && d3 && d3.items && d3.items.length) {
      rubrikGuru = [d1, d2, d3];
    }
  } catch (e) {}

  // 2. Skala Guru
  var skalaGuru = null;
  try {
    skalaGuru = JSON.parse(props.getProperty("ISPPK_SKALA_GURU") || "null");
  } catch (e) {}

  // 3. Item Murid
  var itemMurid = null;
  try {
    itemMurid = JSON.parse(props.getProperty("ISPPK_ITEM_MURID") || "null");
  } catch (e) {}

  // 4. Skala Murid
  var skalaMurid = null;
  try {
    skalaMurid = JSON.parse(props.getProperty("ISPPK_SKALA_MURID") || "null");
  } catch (e) {}

  // 5. Tahap KBAT
  var tahapKbat = null;
  try {
    tahapKbat = JSON.parse(props.getProperty("ISPPK_TAHAP_KBAT") || "null");
  } catch (e) {}

  // 6. Info Instrumen
  var infoInstrumen = null;
  try {
    infoInstrumen = JSON.parse(props.getProperty("ISPPK_INFO_INSTRUMEN") || "null");
  } catch (e) {}

  // 7. Bilangan Murid
  var bilMurid = null;
  try {
    bilMurid = JSON.parse(props.getProperty("ISPPK_BILANGAN_MURID") || "null");
  } catch (e) {}

  // 8. Tahun / Tingkatan
  var tahunTingkatan = null;
  try {
    tahunTingkatan = JSON.parse(props.getProperty("ISPPK_TAHUN_TINGKATAN") || "null");
  } catch (e) {}

  // 9. Senarai Negeri
  var senaraiNegeri = null;
  try {
    senaraiNegeri = JSON.parse(props.getProperty("ISPPK_SENARAI_NEGERI") || "null");
  } catch (e) {}

  var result = {
    rubrikGuru: rubrikGuru,
    skalaGuru: skalaGuru,
    itemMurid: itemMurid,
    skalaMurid: skalaMurid,
    tahapKbat: tahapKbat,
    infoInstrumen: infoInstrumen,
    bilanganMurid: bilMurid,
    tahunTingkatan: tahunTingkatan,
    senaraiNegeri: senaraiNegeri,
  };

  try {
    cache.put("ISPPK_LITERAL_TEXTS", JSON.stringify(result), 21600); // 6 jam
  } catch (e) {}

  return result;
}


/**
 * Mengambil maklumat sekolah daripada ScriptProperties (atau fallback membaca tab TETAPAN jika belum diset).
 */
function getMaklumatSekolahDariProperties_() {
  var sekolah = {
    kodSekolah: "",
    namaSekolah: "",
    negeri: "",
    ppd: "",
    emelSekolah: "",
  };

  try {
    // 1. Cuba dapatkan daripada Properties / Cache / Tab TETAPAN melalui getConfigValue_
    var jsonStr = getConfigValue_("SEKOLAH_JSON", "");
    if (jsonStr) {
      try {
        var parsed = JSON.parse(jsonStr);
        if (
          parsed &&
          typeof parsed === "object" &&
          (parsed.kodSekolah || parsed.namaSekolah)
        ) {
          return parsed;
        }
      } catch (e) {}
    }

    var kod = getConfigValue_("KOD_SEKOLAH", "");
    var nama = getConfigValue_("NAMA_SEKOLAH", "");
    var negeri = getConfigValue_("NEGERI_SEKOLAH", "");
    var ppd = getConfigValue_("PPD_SEKOLAH", "");
    var emel = getConfigValue_("EMEL_SEKOLAH", "");

    if (kod || nama) {
      return {
        kodSekolah: kod,
        namaSekolah: nama,
        negeri: negeri,
        ppd: ppd,
        emelSekolah: emel,
      };
    }

    // 2. Fallback baca terus daripada tab 'TETAPAN' (range B5:B9) dan auto-simpan ke Properties
    var ss = getSpreadsheet_();
    if (ss) {
      var sheetTetapan = ss.getSheetByName("TETAPAN");
      if (sheetTetapan && sheetTetapan.getLastRow() >= 5) {
        var vals = sheetTetapan.getRange("B5:B9").getValues();
        sekolah = {
          kodSekolah: (vals[0][0] || "").toString().trim(),
          namaSekolah: (vals[1] && vals[1][0] ? vals[1][0] : "").toString().trim(),
          negeri: (vals[2] && vals[2][0] ? vals[2][0] : "").toString().trim(),
          ppd: (vals[3] && vals[3][0] ? vals[3][0] : "").toString().trim(),
          emelSekolah: (vals[4] && vals[4][0] ? vals[4][0] : "").toString().trim(),
        };

        // Fallback jika kedudukan baris berbeza: cari mengikut kunci lajur A
        if (!sekolah.kodSekolah && !sekolah.namaSekolah) {
          var allData = sheetTetapan.getDataRange().getValues();
          for (var r = 0; r < allData.length; r++) {
            var k = normalizeConfigKey_(allData[r][0]);
            var v = (allData[r][1] || "").toString().trim();
            if (k === "KOD_SEKOLAH") sekolah.kodSekolah = v;
            else if (k === "NAMA_SEKOLAH") sekolah.namaSekolah = v;
            else if (k === "NEGERI_SEKOLAH") sekolah.negeri = v;
            else if (k === "PPD_SEKOLAH") sekolah.ppd = v;
            else if (k === "EMEL_SEKOLAH") sekolah.emelSekolah = v;
          }
        }

        if (sekolah.kodSekolah || sekolah.namaSekolah) {
          try {
            setConfigValues_({
              KOD_SEKOLAH: sekolah.kodSekolah,
              NAMA_SEKOLAH: sekolah.namaSekolah,
              NEGERI_SEKOLAH: sekolah.negeri,
              PPD_SEKOLAH: sekolah.ppd,
              EMEL_SEKOLAH: sekolah.emelSekolah,
              SEKOLAH_JSON: JSON.stringify(sekolah),
            });
          } catch (eSave) {}
        }
      }
    }
  } catch (err) {}

  return sekolah;
}

/**
 * 1. KAEDAH NATIF (TANPA google.script.run):
 * Memaparkan prompt sistem Google Sheets secara langsung.
 * Berjalan sepenuhnya pada pelayan (server-side synchronous), 100% bebas daripada ralat PERMISSION_DENIED multi-account.
 */
function konfigurasiUrlDeploymentNatif() {
  var ui = SpreadsheetApp.getUi();
  var urls = getDeploymentUrls();
  var activeUrl = urls.ownerUrl || urls.webAppUrl || getConfigValue_("WEBAPP_URL", "") || "";

  var prompt = ui.prompt(
    "🌐 Semak / Kemas Kini URL Web App",
    "Aplikasi menggunakan 1 Deployment Tunggal (Owner Relay - DELIMa Domain).\n\n" +
      "URL Web App Semasa (daripada sel B4 TETAPAN / ScriptProperties):\n" +
      (activeUrl ? activeUrl : "(Belum ditetapkan di sel B4)") +
      "\n\n" +
      "Salin URL Web App dari Apps Script UI (Deploy > Manage deployments) dan tampal di bawah (atau terus ke sel B4 tab TETAPAN):",
    ui.ButtonSet.OK_CANCEL,
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  var inputUrl = prompt.getResponseText().trim();
  if (inputUrl) {
    var res = simpanDeploymentUrls(inputUrl);
    if (res.success) {
      ui.alert(
        "✅ Berjaya Disimpan!",
        "URL Deployment Web App telah berjaya disimpan ke dalam Google Sheet (Tab TETAPAN & Properties):\n\n" +
          inputUrl,
        ui.ButtonSet.OK,
      );
    } else {
      ui.alert(
        "Ralat",
        "Gagal menyimpan URL: " + (res.error || "Ralat tidak diketahui"),
        ui.ButtonSet.OK,
      );
    }
  }
}

/**
 * 2. KAEDAH LEMBARAN KERJA:
 * Membuka tab "TETAPAN" dalam Google Sheet dan mengarahkan fokus ke sel B5 (Maklumat Sekolah).
 */
function bukaLembaranTetapan() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName("TETAPAN");
  if (!sheet) {
    sheet = sediakanLembaranTetapan_(ss);
  }
  ss.setActiveSheet(sheet);
  sheet.getRange("B5").activate();
  SpreadsheetApp.getUi().alert(
    "📋 Lembaran Tetapan Dibuka",
    'Anda kini berada di tab "TETAPAN".\n\n' +
      "Sila semak atau lengkapkan maklumat asas sekolah dalam sel berwarna kuning (B5:B9):\n\n" +
      "• Sel B5 : KOD_SEKOLAH\n" +
      "• Sel B6 : NAMA_SEKOLAH\n" +
      "• Sel B7 : NEGERI_SEKOLAH\n" +
      "• Sel B8 : PPD_SEKOLAH\n" +
      "• Sel B9 : EMEL_SEKOLAH\n\n",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/**
 * Menjana dan memformatkan tab TETAPAN dalam Google Sheet sekiranya belum wujud.
 */
function sediakanLembaranTetapan_(ss) {
  ss = ss || getSpreadsheet_();
  var sheet = ss.getSheetByName("TETAPAN");
  if (!sheet) {
    sheet = ss.insertSheet("TETAPAN");
  }

  sheet.clear();

  var rows = [
    [
      "⚙️ PANDUAN & KONFIGURASI APLIKASI INSTRUMEN BESTARI PENILAIAN PdP",
      "",
      "",
    ],
    [
      "💡 Isikan maklumat asas sekolah dalam sel kuning (B5:B9). URL Web App (B4) dikesan secara automatik oleh sistem.",
      "",
      "",
    ],
    [
      "KUNCI PARAMETER",
      "NILAI KONFIGURASI",
      "PENERANGAN & STATUS PENGGUNAAN",
    ],
    [
      "URL Web App",
      "",
      `Salin URL Web App daripada Apps Script UI (Deploy > Manage deployments) dan tampal di sel B4.
Deployment: 
- Execute as "Me (g-0000000@moe-dl.edu.my)", 
- Who has access "Anyone within Ministry Of Education Malaysia"`,
    ],
    ["KOD SEKOLAH", "", "Kod rasmi sekolah (cth: DEA4295)"],
    ["NAMA SEKOLAH", "", "Nama penuh sekolah (cth: SMK PADANG PAK AMAT)"],
    ["NEGERI SEKOLAH", "", "Negeri sekolah (cth: KELANTAN)"],
    ["PPD SEKOLAH", "", "Pejabat Pendidikan Daerah (cth: PPD Pasir Puteh)"],
    ["EMEL SEKOLAH", "", "Emel rasmi sekolah"],
    [
      "SEKOLAH_JSON",
      "",
      "Penyelarasan automatik JSON maklumat sekolah (Dijana secara automatik oleh Setup)",
    ],
  ];

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);

  // Gaya visual (Styling)
  sheet
    .getRange("A1:C1")
    .merge()
    .setBackground("#0369a1")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(12)
    .setHorizontalAlignment("left");
  sheet
    .getRange("A2:C2")
    .merge()
    .setBackground("#f0fdf4")
    .setFontColor("#166534")
    .setFontStyle("italic")
    .setFontSize(10);
  sheet
    .getRange("A3:C3")
    .setBackground("#1e293b")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(10);

  // Sel input rujukan URL B4 (kuning lembut seperti B5:B9 untuk menandakan input pengguna)
  sheet
    .getRange("B4")
    .setBackground("#fef9c3")
    .setFontFamily("monospace")
    .setFontSize(10)
    .setWrap(true);

  // Sel input kuning lembut untuk maklumat asas sekolah (B5:B9)
  sheet
    .getRange("B5:B9")
    .setBackground("#fef9c3")
    .setFontFamily("monospace")
    .setFontSize(10)
    .setWrap(true);

  sheet.getRange("A4:A10").setFontWeight("bold").setFontFamily("monospace");
  sheet.getRange("C4:C10").setFontSize(9).setFontColor("#475569");

  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 500);
  sheet.setColumnWidth(3, 400);
  sheet.setFrozenRows(3);

  // Masukkan nilai sedia ada jika telah disimpan sebelum ini
  var currentWebapp = getConfigValue_("WEBAPP_URL", "");
  var currentKod = getConfigValue_("KOD_SEKOLAH", "");
  var currentNama = getConfigValue_("NAMA_SEKOLAH", "");
  var currentNegeri = getConfigValue_("NEGERI_SEKOLAH", "");
  var currentPpd = getConfigValue_("PPD_SEKOLAH", "");
  var currentEmel = getConfigValue_("EMEL_SEKOLAH", "");

  if (currentWebapp) sheet.getRange("B4").setValue(currentWebapp);
  if (currentKod) sheet.getRange("B5").setValue(currentKod);
  if (currentNama) sheet.getRange("B6").setValue(currentNama);
  if (currentNegeri) sheet.getRange("B7").setValue(currentNegeri);
  if (currentPpd) sheet.getRange("B8").setValue(currentPpd);
  if (currentEmel) sheet.getRange("B9").setValue(currentEmel);

  return sheet;
}

/**
 * Pengesan perubahan (onEdit) sekiranya pentadbir mengemas kini tab TETAPAN secara langsung.
 * Berjalan melalui installable trigger dengan kebenaran penuh (AuthMode.FULL).
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    // Jika dicetuskan sebagai simple trigger (had kebenaran tanpa akses Properties),
    // abaikan supaya installable trigger menjalankannya dengan kebenaran penuh (AuthMode.FULL)
    if (e.authMode && e.authMode !== ScriptApp.AuthMode.FULL) {
      return;
    }

    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    if (sheetName === "TETAPAN") {
      var row = e.range.getRow();
      var col = e.range.getColumn();
      if (col === 2 && row >= 4 && row <= 9) {
        var key = sheet.getRange(row, 1).getValue().toString().trim();
        var val = sheet.getRange(row, 2).getValue().toString().trim();
        if (key) {
          var propKey = normalizeConfigKey_(key);
          var obj = {};
          obj[propKey] = val;

          // Jika baris 5-9 (maklumat sekolah) disunting, kemaskini SEKOLAH_JSON serentak
          if (row >= 5 && row <= 9) {
            try {
              var vals = sheet.getRange("B5:B9").getValues();
              var sek = {
                kodSekolah: (vals[0][0] || "").toString().trim(),
                namaSekolah: (vals[1] && vals[1][0] ? vals[1][0] : "").toString().trim(),
                negeri: (vals[2] && vals[2][0] ? vals[2][0] : "").toString().trim(),
                ppd: (vals[3] && vals[3][0] ? vals[3][0] : "").toString().trim(),
                emelSekolah: (vals[4] && vals[4][0] ? vals[4][0] : "").toString().trim(),
              };
              obj["SEKOLAH_JSON"] = JSON.stringify(sek);
              try {
                sheet.getRange("B10").setValue(obj["SEKOLAH_JSON"]);
              } catch (eCell) {}
            } catch (eSek) {}
          }

          try {
            PropertiesService.getScriptProperties().setProperties(obj);
          } catch (err1) {}
          try {
            PropertiesService.getDocumentProperties().setProperties(obj);
          } catch (err2) {}
          try {
            CacheService.getScriptCache().remove("PUBLIC_FORM_DATA");
          } catch (err3) {}
        }
      }
    } else if (sheetName === "GURU") {
      try {
        segerakPencerapKeProperties_(sheet.getParent());
      } catch (errGuru) {}
      try {
        CacheService.getScriptCache().remove("PUBLIC_FORM_DATA");
      } catch (errCache) {}
    }
  } catch (err) {}
}

/**
 * Peta susunan baris tetap bagi Tab Sheet 'TETAPAN'.
 * B4: WEBAPP_URL
 * B5-B9: Butiran Maklumat Asas Sekolah
 * B10: SEKOLAH_JSON
 */
var CONFIG_KEY_ROW_MAP_ = {
  WEBAPP_URL: 4,
  KOD_SEKOLAH: 5,
  NAMA_SEKOLAH: 6,
  NEGERI_SEKOLAH: 7,
  PPD_SEKOLAH: 8,
  EMEL_SEKOLAH: 9,
  SEKOLAH_JSON: 10,
};

/**
 * Menormalkan kunci konfigurasi ke format seragam huruf besar bergaris bawah.
 * Contoh: "URL Web App" -> "WEBAPP_URL", "KOD SEKOLAH" -> "KOD_SEKOLAH"
 */
function normalizeConfigKey_(k) {
  if (!k) return "";
  var s = k.toString().trim().toUpperCase().replace(/[\s\-_]+/g, "_");
  if (s === "URL_WEB_APP" || s === "WEB_APP_URL" || s === "URL" || s === "WEBAPPURL") {
    return "WEBAPP_URL";
  }
  return s;
}

/**
 * Mengambil nilai konfigurasi daripada sistem storan hibrid (3 lapisan berperingkat):
 * 1. Tab Sheet 'TETAPAN' dalam Google Sheet (Kebal sepenuhnya daripada ralat Google multi-account)
 * 2. Script Properties (PropertiesService.getScriptProperties)
 * 3. Document Properties (PropertiesService.getDocumentProperties)
 */
function getConfigValue_(key, defaultValue) {
  defaultValue =
    defaultValue === undefined || defaultValue === null ? "" : defaultValue;
  if (!key) return defaultValue;

  var normKey = normalizeConfigKey_(key);

  // Lapisan 1: Baca terus dari Tab 'TETAPAN' dalam Google Sheet
  try {
    var ss = getSpreadsheet_();
    if (ss) {
      var sheetTetapan = ss.getSheetByName("TETAPAN");
      if (sheetTetapan && sheetTetapan.getLastRow() >= 4) {
        var dataTetapan = sheetTetapan.getDataRange().getValues();
        // 1.1 Padanan dinamik mengikut label Lajur A
        for (var i = 0; i < dataTetapan.length; i++) {
          var colA = (dataTetapan[i][0] || "").toString().trim();
          if (colA && normalizeConfigKey_(colA) === normKey) {
            var valT = dataTetapan[i][1];
            if (
              valT !== null &&
              valT !== undefined &&
              valT.toString().trim() !== "" &&
              valT.toString().trim() !== "-"
            ) {
              return valT.toString().trim();
            }
          }
        }

        // 1.2 Fallback kedudukan baris tetap jika label diubahsuai
        var fixedRow = CONFIG_KEY_ROW_MAP_[normKey];
        if (fixedRow && fixedRow <= dataTetapan.length) {
          var valFixed = dataTetapan[fixedRow - 1][1];
          if (
            valFixed !== null &&
            valFixed !== undefined &&
            valFixed.toString().trim() !== "" &&
            valFixed.toString().trim() !== "-"
          ) {
            return valFixed.toString().trim();
          }
        }
      }
    }
  } catch (e) {}

  // Lapisan 2: Script Properties
  try {
    var sp = PropertiesService.getScriptProperties();
    if (sp) {
      var val = sp.getProperty(normKey) || sp.getProperty(key);
      if (val !== null && val !== undefined && val !== "") {
        return val;
      }
    }
  } catch (e) {
    // Abaikan sekatan sesi akaun Google
  }

  // Lapisan 3: Document Properties
  try {
    var dp = PropertiesService.getDocumentProperties();
    if (dp) {
      var valDoc = dp.getProperty(normKey) || dp.getProperty(key);
      if (valDoc !== null && valDoc !== undefined && valDoc !== "") {
        return valDoc;
      }
    }
  } catch (e) {}

  return defaultValue;
}

/**
 * Menyimpan pasangan kunci-nilai ke dalam sistem storan hibrid:
 * - Mengemas kini terus ke sel Tab Sheet 'TETAPAN' yang tepat (tanpa appendRow).
 * - Cuba menyegerak serentak ke Document Properties & Script Properties.
 */
function setConfigValues_(obj) {
  var savedToSheet = false;
  var savedToProps = false;
  var errors = [];
  if (!obj || typeof obj !== "object") {
    return { success: false, error: "Objek konfigurasi tidak sah." };
  }

  // Lapisan 1: Simpan ke Tab Sheet 'TETAPAN' dalam Google Sheet
  try {
    var ss = getSpreadsheet_();
    if (ss) {
      var sheetTetapan = ss.getSheetByName("TETAPAN");
      if (!sheetTetapan || sheetTetapan.getLastRow() < 3) {
        sheetTetapan = sediakanLembaranTetapan_(ss);
      }
      if (sheetTetapan) {
        var data = sheetTetapan.getDataRange().getValues();
        var dynamicKeyMap = {};
        for (var r = 0; r < data.length; r++) {
          var colA = (data[r][0] || "").toString().trim();
          if (colA) {
            dynamicKeyMap[normalizeConfigKey_(colA)] = r + 1;
          }
        }

        for (var k in obj) {
          var val =
            obj[k] !== undefined && obj[k] !== null
              ? obj[k].toString().trim()
              : "";
          var normK = normalizeConfigKey_(k);
          var targetRow = dynamicKeyMap[normK] || CONFIG_KEY_ROW_MAP_[normK];

          if (targetRow) {
            sheetTetapan.getRange(targetRow, 2).setValue(val);
          }
        }
        savedToSheet = true;
      }
    }
  } catch (errSheet) {
    errors.push("Sheet TETAPAN: " + errSheet.message);
  }

  // Lapisan 2: Document Properties
  try {
    var dp = PropertiesService.getDocumentProperties();
    if (dp) {
      dp.setProperties(obj);
      savedToProps = true;
    }
  } catch (errDp) {
    errors.push("DocumentProperties: " + errDp.message);
  }

  // Lapisan 3: Script Properties
  try {
    var sp = PropertiesService.getScriptProperties();
    if (sp) {
      sp.setProperties(obj);
      savedToProps = true;
    }
  } catch (errSp) {
    errors.push("ScriptProperties: " + errSp.message);
  }

  if (savedToSheet || savedToProps) {
    return {
      success: true,
      savedToSheet: savedToSheet,
      savedToProps: savedToProps,
      message: "Berjaya disimpan!",
    };
  }

  return {
    success: false,
    error: errors.join("; ") || "Gagal menyimpan ke mana-mana lapisan storan.",
  };
}

/**
 * Mengambil URL deployment daripada sistem storan hibrid.
 * Diperolehi daripada tab TETAPAN sel B4 atau ScriptProperties.
 */
function getDeploymentUrls() {
  var webAppUrl = getConfigValue_("WEBAPP_URL", "");

  // Segerakkan semula ke ScriptProperties jika belum wujud (self-healing cache)
  if (webAppUrl) {
    try {
      var sp = PropertiesService.getScriptProperties();
      if (sp && !sp.getProperty("WEBAPP_URL")) {
        sp.setProperty("WEBAPP_URL", webAppUrl);
      }
    } catch (e) {}
  }

  return {
    webAppUrl: webAppUrl,
    ownerUrl: webAppUrl,
  };
}

/**
 * Menyimpan URL deployment ke dalam sistem storan hibrid (1 Deployment).
 */
function simpanDeploymentUrls(webAppUrl) {
  try {
    var cleanUrl = (webAppUrl || "").trim();

    // Pastikan sebarang kunci konfigurasi legasi dibersihkan
    hapusKunciKonfigurasiLegasi_();

    return setConfigValues_({
      WEBAPP_URL: cleanUrl,
    });
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Memadam sebarang kunci konfigurasi legasi dua-deployment daripada Properties.
 */
function hapusKunciKonfigurasiLegasi_() {
  var legacyKeys = ["AUTH_SECRET", "USER_ACCESS_WEBAPP_URL", "OWNER_WEBAPP_URL"];
  try {
    var sp = PropertiesService.getScriptProperties();
    if (sp) {
      legacyKeys.forEach(function (k) {
        sp.deleteProperty(k);
      });
    }
  } catch (e) {}
  try {
    var dp = PropertiesService.getDocumentProperties();
    if (dp) {
      legacyKeys.forEach(function (k) {
        dp.deleteProperty(k);
      });
    }
  } catch (e) {}
}

// ============================================================================
// 2. CONTROLLER WEB APP (doGet & doPost)
// ============================================================================

/**
 * Memaparkan Web App (Seni Bina 1 Deployment Tunggal - Owner Relay):
 * - Mengesan emel DELIMa pengguna secara langsung melalui Session.getActiveUser().getEmail().
 * - Jika pengguna ialah Pencerap yang sah, sistem menyediakan status authData.isPencerap = true.
 * - Membuka aplikasi utama dalam Mod Kendiri secara lalai untuk semua pengguna.
 */
function doGet(e) {
  e = e || { parameter: {} };
  var activeUserEmail = Session.getActiveUser().getEmail();
  var urls = getDeploymentUrls();
  var ownerUrl = urls.ownerUrl || urls.webAppUrl || "";

  // --------------------------------------------------------------------------
  // LALUAN A: Route Modul Hantar ke Google Forms (Khusus Pencerap Berdaftar)
  // --------------------------------------------------------------------------
  if (e.parameter.page === "hantarGF" || e.parameter.route === "hantarGF") {
    var gfAuthEmail = activeUserEmail || e.parameter.authEmail;
    var isValidObserver = false;
    var observerInfo = null;

    if (gfAuthEmail) {
      var semakanGf = semakStatusPencerap_(gfAuthEmail);
      if (semakanGf.isPencerap) {
        isValidObserver = true;
        observerInfo = {
          email: gfAuthEmail,
          nama: semakanGf.nama,
          jawatan: semakanGf.jawatan,
          isPencerap: true,
        };
      }
    }

    if (!isValidObserver) {
      // Paparan Penafian Akses Keselamatan jika bukan pencerap sah
      var deniedHtml = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
          <title>Akses Ditolak - Modul Google Forms Aplikasi Instrumen Bestari Penilaian PdP</title>
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <style>
            body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; display: flex; align-items: center; justify-content: center; height: 100vh; background: #0f172a; color: #f8fafc; margin: 0; padding: 20px; }
            .card { background: #1e293b; padding: 2.5rem; border-radius: 1.25rem; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.5); max-width: 480px; text-align: center; border: 1px solid #ef4444; }
            .icon { font-size: 3rem; margin-bottom: 1rem; }
            h3 { margin: 0 0 0.5rem; font-size: 1.3rem; font-weight: 700; color: #f87171; }
            p { margin: 0 0 1.5rem; font-size: 0.9rem; color: #94a3b8; line-height: 1.6; }
            a { display: inline-block; background: #3b82f6; color: white; padding: 0.6rem 1.2rem; border-radius: 0.5rem; text-decoration: none; font-weight: 500; font-size: 0.9rem; transition: background 0.2s; }
            a:hover { background: #2563eb; }
          </style>
        </head>
        <body>
          <div class="card">
            <div class="icon">🔒</div>
            <h3>Akses Terhad: Mod Pencerap Sahaja</h3>
            <p>Modul integrasi Google Forms ini hanya boleh diakses oleh Pencerap Berdaftar yang telah disahkan melalui log masuk DELIMa.</p>
            <a href="${ownerUrl}?mod=pencerap">Kembali ke Aplikasi Instrumen Bestari Penilaian PdP</a>
          </div>
        </body>
        </html>
      `;
      return HtmlService.createHtmlOutput(deniedHtml)
        .setTitle("Akses Ditolak - Aplikasi Instrumen Bestari Penilaian PdP")
        .addMetaTag("viewport", "width=device-width, initial-scale=1");
    }

    // Akses Sah: Render modul IBPPdP_HantarGF_ISPPK secara langsung tanpa createTemplateFromFile
    var gfHtmlContent = HtmlService.createHtmlOutputFromFile("IBPPdP_HantarGF_ISPPK").getContent();
    var pubData = getPublicFormData();
    var props = PropertiesService.getScriptProperties();
    var savedUrl = (props.getProperty("ISPPK_GF_TARGET_URL") || "").trim();
    var savedMappingRaw = props.getProperty("ISPPK_GF_MAPPING") || "{}";
    var savedBundleRaw = props.getProperty("ISPPK_GF_MAPPING_BUNDLE") || "";
    var savedItemsRaw = props.getProperty("ISPPK_GF_ITEMS") || "[]";
    var savedPagesRaw = props.getProperty("ISPPK_GF_PAGES") || "[]";
    var savedMapping = {};
    var savedItems = [];
    var savedPages = [];
    try { savedMapping = JSON.parse(savedMappingRaw); } catch(e) {}
    if (savedBundleRaw) {
      try {
        var bundle = JSON.parse(savedBundleRaw);
        if (bundle && bundle.url) {
          if (!savedUrl || bundle.url.trim() === savedUrl) {
            savedUrl = bundle.url.trim();
            savedMapping = bundle.mapping || {};
          } else {
            savedMapping = {};
          }
        }
      } catch(e) {}
    }
    var currentYear = (pubData && pubData.currentYear) || dapatkanTahunSemasa_();
    var submittedKey = "ISPPK_GF_SUBMITTED_" + currentYear;
    var savedSubmittedRaw = props.getProperty(submittedKey);
    if (!savedSubmittedRaw) {
      savedSubmittedRaw = props.getProperty("ISPPK_GF_SUBMITTED") || "{}";
    }
    var savedSubmittedRecords = {};
    try { savedItems = JSON.parse(savedItemsRaw); } catch(e) {}
    try { savedPages = JSON.parse(savedPagesRaw); } catch(e) {}
    try { savedSubmittedRecords = JSON.parse(savedSubmittedRaw); } catch(e) {}

    var gfServerData = JSON.stringify({
      authData: observerInfo,
      sekolah: pubData.sekolah,
      currentYear: currentYear,
      savedFormUrl: savedUrl,
      savedMapping: savedMapping,
      savedFormItems: savedItems,
      savedFormPages: savedPages,
      savedSubmittedRecords: savedSubmittedRecords,
      webAppUrl: ownerUrl,
      ownerUrl: ownerUrl
    });

    gfHtmlContent = gfHtmlContent.replace(
      /<\?!=[\s\S]*?\?>/,
      function () { return gfServerData; }
    );

    return HtmlService.createHtmlOutput(gfHtmlContent)
      .setTitle("Penghantaran Data ke Google Forms - ISPPK KBAT " + dapatkanTahunSemasa_())
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // --------------------------------------------------------------------------
  // LALUAN B: Dijalankan di bawah Deployment Utama (Owner Relay)
  // (Menggunakan createHtmlOutputFromFile + replace untuk mengelakkan pemotongan/korupsi
  //  kod JavaScript & templat literal oleh enjin createTemplateFromFile)
  // --------------------------------------------------------------------------
  var htmlContent = HtmlService.createHtmlOutputFromFile("Index").getContent();

  // Pengesahan status pengguna DELIMa secara langsung
  var semakan = semakStatusPencerap_(activeUserEmail);
  var authData = {
    isLoggedIn: !!activeUserEmail,
    email: activeUserEmail || "",
    nama: semakan.nama || "",
    jawatan: semakan.jawatan || "",
    isPencerap: semakan.isPencerap === true,
    accessDenied: false,
    errorMessage: "",
  };

  // Parameter ID Rekod jika dibuka daripada pautan emel
  var initialIdRekod = e.parameter.idRekod || "";
  var initialMod = "guru"; // Sentiasa mula dalam Mod Guru (Kendiri) secara lalai untuk semua pengguna

  // Ambil maklumat sekolah (dari Properties) dan data borang awam
  var publicFormData = getPublicFormData();

  var serverDataJson = JSON.stringify({
    authData: authData,
    initialIdRekod: initialIdRekod,
    initialMod: initialMod,
    currentYear: publicFormData.currentYear || dapatkanTahunSemasa_(),
    currentDate: publicFormData.currentDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+8", "yyyy-MM-dd"),
    sekolah: publicFormData.sekolah,
    senaraiPencerap: publicFormData.senaraiPencerap,
    senaraiGuruTersedia: publicFormData.senaraiGuruTersedia,
    senaraiSemuaGuru:
      publicFormData.senaraiSemuaGuru || publicFormData.senaraiGuruTersedia,
    literalTexts: publicFormData.literalTexts || dapatkanLiteralTeksDariProperties_(),
    urls: {
      webAppUrl: ownerUrl,
      ownerUrl: ownerUrl,
    },
  });

  htmlContent = htmlContent.replace(
    /<\?!=[\s\S]*?\?>/,
    function () { return serverDataJson; }
  );

  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle("Instrumen Bestari Penilaian PdP")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Pengendali POST untuk geganti atau integrasi API luaran.
 */
function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    var action = data.action || "";
    var result = { success: false, error: "Aksi tidak sah" };

    if (action === "getInitialData") {
      result = getInitialData();
    } else if (action === "getPublicFormData") {
      result = getPublicFormData();
    } else if (action === "getLiteralTexts") {
      result = { success: true, literalTexts: dapatkanLiteralTeksDariProperties_() };
    } else if (action === "getRekodPengisianPencerap") {
      result = getRekodPengisianPencerap(data.namaPencerap);
    } else if (action === "padamRekodPenilaian") {
      result = padamRekodPenilaian(data.idRekod);
    } else if (action === "simpanPenilaian") {
      result = simpanPenilaian(data.payload);
    } else if (action === "hantarEmelNotifikasi") {
      result = hantarEmelNotifikasiGuru_(
        data.recipientEmail,
        data.payload,
        data.idRekod,
      );
    } else if (action === "cariRekod") {
      result = cariRekod(data.kataKunci);
    } else if (action === "semakPencerap") {
      result = semakStatusPencerap_(data.email);
    } else if (action === "semakAuthPenggunaAktif") {
      result = semakAuthPenggunaAktif(data.email || data.fallbackEmail);
    } else if (action === "getDataAnalisis") {
      result = getDataAnalisis();
    } else if (action === "dapatkanElemenBorangGF") {
      result = dapatkanElemenBorangGF(data.url);
    } else if (action === "simpanPemetaanGF") {
      result = simpanPemetaanGF(data.urlBorang, data.pemetaan, data.items, data.pages);
    } else if (action === "dapatkanPemetaanGF") {
      result = dapatkanPemetaanGF();
    } else if (action === "getRekodGuruSkorMuktamad") {
      result = getRekodGuruSkorMuktamad();
    } else if (action === "dapatkanUrlModulHantarGF") {
      result = dapatkanUrlModulHantarGF(data.email);
    } else if (action === "simpanStatusHantaranGF") {
      result = simpanStatusHantaranGF(data.submittedData, data.tahun);
    } else if (action === "getStatusHantaranGF") {
      result = getStatusHantaranGF(data.tahun);
    }

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
      ContentService.MimeType.JSON,
    );
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.message }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================================
// 3. PENGURUSAN DATA GOOGLE SHEET (CONTAINER)
// ============================================================================

/**
 * Mendapatkan Spreadsheet Aktif (Container) dengan fallback kepada ID spreadsheet.
 */
function getSpreadsheet_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (err) {}
  var SPREADSHEET_ID = "19OJepezsW3ugbhyhJ8cLgSxqigOizInrkq1oIwGerX0";
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Mendapatkan sheet GURU (dengan sokongan migrasi automatik daripada BUTIRAN GURU).
 */
function getSheetGuru_(ss) {
  var sheet = ss.getSheetByName("GURU");
  if (!sheet) {
    // Semak sekiranya nama lama masih wujud
    var oldSheet = ss.getSheetByName("BUTIRAN GURU");
    if (oldSheet) {
      oldSheet.setName("GURU");
      sheet = oldSheet;
    } else {
      sheet = ss.insertSheet("GURU");
      var headers = ["Emel", "Nama Guru", "Jawatan", "Jantina", "Opsyen", "Pencerap", "Aktif"];
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#1e293b"); // slate-800
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
    }
  }
  // Pastikan Kolum G (Lajur 7) mempunyai header 'Aktif' sekiranya helaian telah wujud
  if (sheet) {
    if (sheet.getLastColumn() < 7 || !sheet.getRange(1, 7).getValue()) {
      sheet.getRange(1, 7).setValue("Aktif");
      sheet.getRange(1, 7).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff").setHorizontalAlignment("center");
    }
  }
  return sheet;
}

/**
 * Mendapatkan atau membina tab sheet 'PENGISIAN'.
 * Menyediakan 35 tajuk lajur standard sekiranya tab belum wujud.
 */
function getSheetPengisian_(ss) {
  if (!ss) ss = getSpreadsheet_();
  var sheet = ss.getSheetByName("PENGISIAN");
  if (!sheet) {
    sheet = ss.insertSheet("PENGISIAN");
    var headers = [
      "Timestamp",
      "ID Rekod",
      "Mod Penilaian",
      "Nama Guru Dicerap",
      "E-mel Guru (DELIMa)",
      "Jantina Guru",
      "Opsyen Guru",
      "Nama Pencerap",
      "Jawatan Pencerap",
      "Tarikh Pencerapan",
      "Masa Pencerapan",
      "Mata Pelajaran Dicerap",
      "Tajuk / Topik",
      "Bilangan Murid Hadir",
      "Tahun / Tingkatan",
      "Nama Kelas / Aliran",
      "Skor K - Perancangan",
      "Skor M - Perancangan",
      "Skor K - Pelaksanaan",
      "Skor M - Pelaksanaan",
      "Skor K - Refleksi",
      "Skor M - Refleksi",
      "Skor K - Murid",
      "Skor M - Murid",
      "Jumlah Skor K",
      "Peratus K (%)",
      "Tahap KBAT K",
      "Jumlah Skor M",
      "Peratus M (%)",
      "Tahap KBAT M",
      "Refleksi Kendiri Guru 1",
      "Refleksi Kendiri Guru 2",
      "Rumusan Pencerap",
      "JSON Payload"
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1e293b"); // slate-800
    headerRange.setFontColor("#ffffff");
    headerRange.setHorizontalAlignment("center");
  }
  return sheet;
}

/**
 * Mengekstrak tahun 4 digit daripada objek Date atau rentetan tarikh pelbagai format.
 */
function ekstrakTahunDaripadaTarikh_(tarikhRaw) {
  if (!tarikhRaw) return "";
  if (tarikhRaw instanceof Date) {
    return Utilities.formatDate(tarikhRaw, Session.getScriptTimeZone() || "GMT+8", "yyyy");
  }
  var s = tarikhRaw.toString().trim();
  var match = s.match(/\b(20\d{2})\b/);
  if (match) return match[1];
  return "";
}

/**
 * Mendapatkan tahun semasa daripada pelayan Google Apps Script (GMT+8).
 */
function dapatkanTahunSemasa_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+8", "yyyy");
}

/**
 * Menghurai baris data guru daripada tab 'GURU'.
 * Menyokong susunan standard: [Emel, Nama Guru, Jawatan, Jantina, Opsyen, Pencerap, Aktif]
 * serta secara pintar mengendalikan baris teranjak / tanpa jawatan di mana Lajur C ialah Jantina.
 */
function parseBarisGuru_(row) {
  var col0 = (row[0] || "").toString().trim(); // Emel
  var col1 = (row[1] || "").toString().trim(); // Nama Guru
  var col2 = (row[2] || "").toString().trim();
  var col3 = (row[3] || "").toString().trim();
  var col4 = (row[4] || "").toString().trim();
  var col5 = (row[5] || "").toString().trim();
  var col6 = row[6] !== undefined && row[6] !== null ? row[6].toString().trim() : ""; // Aktif (Kolum G)

  var emel = col0;
  var nama = col1;
  var jawatan = "";
  var jantina = "Lelaki";
  var opsyen = "";
  var pencerap = "tidak";
  var aktif = "ya";

  var c2Lower = col2.toLowerCase();
  var c3Lower = col3.toLowerCase();
  var c4Lower = col4.toLowerCase();
  var c5Lower = col5.toLowerCase();
  var c6Lower = col6.toLowerCase();

  // Kes A: Jika Lajur C (col2) ialah 'lelaki' atau 'perempuan', baris ini teranjak (Jawatan tiada):
  // Format teranjak: [Emel, Nama, Jantina, Opsyen, Pencerap, Jawatan, Aktif]
  if (c2Lower === "lelaki" || c2Lower === "perempuan") {
    jantina = c2Lower === "perempuan" ? "Perempuan" : "Lelaki";
    opsyen = col3;
    if (c4Lower === "ya" || c5Lower === "ya") {
      pencerap = "ya";
    }
    jawatan = (c5Lower !== "ya" && c5Lower !== "tidak" && col5) ? col5 : "";
  } else {
    // Kes B: Mengikut susunan rasmi header tab GURU:
    // Format standard: [Emel, Nama, Jawatan, Jantina, Opsyen, Pencerap, Aktif]
    jawatan = col2;

    if (c3Lower === "perempuan") {
      jantina = "Perempuan";
    } else if (c3Lower === "lelaki") {
      jantina = "Lelaki";
    }

    opsyen = col4;

    // Pencerap terletak di Lajur F (col5)
    if (c5Lower === "ya" || c4Lower === "ya") {
      pencerap = "ya";
      if (c4Lower === "ya" && !col5) {
        opsyen = "";
      }
    }
  }

  // Jika guru adalah pencerap tetapi jawatan kosong, tetapkan lalai kepada 'Pencerap'
  if (pencerap === "ya" && !jawatan) {
    jawatan = "Pencerap";
  }

  // Pengendalian status Aktif (Kolum G / col6).
  // Lalai adalah 'ya' jika kosong; hanya 'tidak' jika eksplisit 'tidak' atau 'false'.
  if (c6Lower === "tidak" || c6Lower === "false" || row[6] === false) {
    aktif = "tidak";
  }

  return {
    emel: emel,
    nama: nama,
    jawatan: jawatan,
    jantina: jantina,
    opsyen: opsyen,
    pencerap: pencerap,
    aktif: aktif,
  };
}

/**
 * Menyegerak senarai pencerap yang sah (Pencerap === 'ya' dan Aktif !== 'tidak') daripada tab GURU ke ScriptProperties & DocumentProperties.
 * Membolehkan pengesahan identiti pencerap dilakukan dalam milisaat tanpa perlu mengakses helaian Google Sheets.
 */
function segerakPencerapKeProperties_(ss) {
  try {
    if (!ss) ss = getSpreadsheet_();
    if (!ss) return {};
    var sheet = getSheetGuru_(ss);
    if (!sheet || sheet.getLastRow() < 2) {
      var kosong = {};
      try {
        PropertiesService.getScriptProperties().setProperty("SENARAI_PENCERAP_PROPS", JSON.stringify(kosong));
      } catch (e1) {}
      try {
        PropertiesService.getDocumentProperties().setProperty("SENARAI_PENCERAP_PROPS", JSON.stringify(kosong));
      } catch (e2) {}
      return kosong;
    }

    var numRows = sheet.getLastRow() - 1;
    var numCols = Math.max(sheet.getLastColumn(), 7);
    var data = sheet.getRange(2, 1, numRows, numCols).getValues();
    var mapPencerap = {};
    for (var i = 0; i < data.length; i++) {
      var guru = parseBarisGuru_(data[i]);
      var isAktif = guru.aktif !== "tidak" && guru.aktif !== false;
      if (guru.emel && guru.pencerap === "ya" && isAktif) {
        mapPencerap[guru.emel.toLowerCase()] = {
          nama: guru.nama,
          jawatan: guru.jawatan || "Pencerap",
          emel: guru.emel.toLowerCase(),
        };
      }
    }

    var jsonStr = JSON.stringify(mapPencerap);
    try {
      PropertiesService.getScriptProperties().setProperty("SENARAI_PENCERAP_PROPS", jsonStr);
    } catch (e1) {}
    try {
      PropertiesService.getDocumentProperties().setProperty("SENARAI_PENCERAP_PROPS", jsonStr);
    } catch (e2) {}

    return mapPencerap;
  } catch (err) {
    return {};
  }
}

/**
 * Menyemak status pencerap berdasarkan emel pengguna.
 * Keutamaan 1: Membaca secara ultra-pantas dari ScriptProperties / DocumentProperties (milisaat).
 * Keutamaan 2: Sandaran (fallback) ke tab GURU jika data properties belum wujud atau rekod baharu.
 */
function semakStatusPencerap_(email) {
  if (!email) return { isPencerap: false };
  var cleanEmail = email.toLowerCase().trim();

  // 1. Semakan Pantas dari ScriptProperties / DocumentProperties
  try {
    var rawProps = PropertiesService.getScriptProperties().getProperty("SENARAI_PENCERAP_PROPS");
    if (!rawProps) {
      rawProps = PropertiesService.getDocumentProperties().getProperty("SENARAI_PENCERAP_PROPS");
    }
    if (rawProps) {
      var mapPencerap = JSON.parse(rawProps);
      if (mapPencerap && typeof mapPencerap === "object") {
        if (mapPencerap[cleanEmail]) {
          var p = mapPencerap[cleanEmail];
          return {
            isPencerap: true,
            nama: p.nama,
            jawatan: p.jawatan,
            emel: cleanEmail,
          };
        }
      }
    }
  } catch (errProp) {}

  // 2. Sandaran (Fallback): Baca terus dari sheet GURU dan segerak semula Properties
  try {
    var ss = getSpreadsheet_();
    var sheet = getSheetGuru_(ss);
    if (!sheet || sheet.getLastRow() < 2) return { isPencerap: false };

    var numRows = sheet.getLastRow() - 1;
    var numCols = Math.max(sheet.getLastColumn(), 7);
    var data = sheet.getRange(2, 1, numRows, numCols).getValues();
    var foundResult = null;

    for (var i = 0; i < data.length; i++) {
      var guru = parseBarisGuru_(data[i]);
      if (guru.emel.toLowerCase() === cleanEmail) {
        var isAktif = guru.aktif !== "tidak" && guru.aktif !== false;
        if (guru.pencerap === "ya" && isAktif) {
          foundResult = {
            isPencerap: true,
            nama: guru.nama,
            jawatan: guru.jawatan,
            emel: cleanEmail,
          };
        }
        break;
      }
    }

    // Segerak properties di latar belakang supaya semakan berikutnya pantas
    segerakPencerapKeProperties_(ss);

    if (foundResult) {
      return foundResult;
    }
  } catch (errSheet) {}

  return { isPencerap: false };
}

/**
 * Menyemak identiti pengguna aktif DELIMa (Session.getActiveUser().getEmail())
 * dan mengembalikan data autentikasi serta kelayakan status pencerap terkini.
 * @param {string} [fallbackEmail] Emel sandaran jika Session.getActiveUser().getEmail() kosong
 * @return {Object} Status autentikasi dan butiran pencerap
 */
function semakAuthPenggunaAktif(fallbackEmail) {
  try {
    var activeEmail = Session.getActiveUser().getEmail();
    if (!activeEmail && fallbackEmail) {
      activeEmail = (fallbackEmail || "").toString().trim();
    }
    var semakan = semakStatusPencerap_(activeEmail);
    return {
      success: true,
      email: activeEmail || "",
      isLoggedIn: !!activeEmail,
      nama: semakan.nama || "",
      jawatan: semakan.jawatan || "",
      isPencerap: semakan.isPencerap === true,
      accessDenied: !!activeEmail && !semakan.isPencerap,
      errorMessage: "",
    };
  } catch (err) {
    return {
      success: false,
      error: err.message,
      email: "",
      isLoggedIn: false,
      nama: "",
      jawatan: "",
      isPencerap: false,
      accessDenied: false,
      errorMessage: err.message,
    };
  }
}

/**
 * Mengambil data awal Google Sheet (HANYA dipanggil apabila pengguna log masuk atau dibenarkan).
 */
function getInitialData() {
  try {
    var ss = getSpreadsheet_();

    // 1. BUTIRAN SEKOLAH (dari Properties dengan fallback tab TETAPAN)
    var sekolah = getMaklumatSekolahDariProperties_();
    var currentYear = dapatkanTahunSemasa_();
    var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+8", "yyyy-MM-dd");

    // 2. Kenal pasti guru yang telah mempunyai ID Rekod dalam PENGISIAN bagi TAHUN SEMASA
    var hantarMap = getGuruSudahHantarMap_(ss, currentYear);

    // 3. GURU & PENCERAP DARI SHEET 'GURU'
    // Susunan Lajur: A=Emel, B=Nama Guru, C=Jawatan, D=Jantina, E=Opsyen, F=Pencerap, G=Aktif
    var sheetGuru = getSheetGuru_(ss);
    var senaraiGuru = [];
    var senaraiGuruTersedia = [];
    var senaraiPencerap = [];

    if (sheetGuru && sheetGuru.getLastRow() >= 2) {
      var numRows = sheetGuru.getLastRow() - 1;
      var numCols = Math.max(sheetGuru.getLastColumn(), 7);
      var dataGuru = sheetGuru.getRange(2, 1, numRows, numCols).getValues();

      for (var i = 0; i < dataGuru.length; i++) {
        var guru = parseBarisGuru_(dataGuru[i]);

        if (guru.nama !== "") {
          var namaLower = guru.nama.toLowerCase();
          var emelLower = guru.emel.toLowerCase();
          var isAktif = guru.aktif !== "tidak" && guru.aktif !== false;
          var guruObj = {
            id: i + 1,
            emel: guru.emel,
            nama: guru.nama,
            jantina: guru.jantina,
            opsyen: guru.opsyen,
            pencerap: guru.pencerap,
            jawatan: guru.jawatan,
            aktif: guru.aktif || "ya",
          };
          senaraiGuru.push(guruObj);

          // Tapis pencerap secara dinamik (Hanya Pencerap yang AKTIF)
          if (guru.pencerap === "ya" && isAktif) {
            senaraiPencerap.push({
              id: senaraiPencerap.length + 1,
              emel: guru.emel,
              nama: guru.nama,
              jawatan: guru.jawatan || "Pencerap",
            });
          }

          // Tapis guru yang AKTIF dan BELUM mempunyai ID Rekod pada tahun semasa untuk Mod Kendiri
          if (isAktif) {
            var sudahHantar =
              hantarMap.nama[namaLower] ||
              (emelLower && hantarMap.emel[emelLower]);
            if (!sudahHantar) {
              senaraiGuruTersedia.push(guruObj);
            }
          }
        }
      }
    }

    return {
      success: true,
      currentYear: currentYear,
      currentDate: currentDate,
      sekolah: sekolah,
      senaraiGuru: senaraiGuru,
      senaraiSemuaGuru: senaraiGuru,
      senaraiGuruTersedia: senaraiGuruTersedia,
      senaraiPencerap: senaraiPencerap,
    };
  } catch (error) {
    return {
      success: false,
      error: error.message,
    };
  }
}

/**
 * Fungsi dalaman: Mengumpul nama dan emel guru yang telah mempunyai rekod pencerapan bagi tahun semasa dalam 'PENGISIAN'.
 */
function getGuruSudahHantarMap_(ss, targetYear) {
  var mapNama = {};
  var mapEmel = {};
  var tahunSemasa = targetYear || dapatkanTahunSemasa_();
  try {
    var sheetPengisian = getSheetPengisian_(ss);
    if (sheetPengisian && sheetPengisian.getLastRow() >= 2) {
      var numRows = sheetPengisian.getLastRow() - 1;
      var numCols = sheetPengisian.getLastColumn();
      var dataPengisian = sheetPengisian
        .getRange(2, 1, numRows, numCols)
        .getValues();

      for (var r = 0; r < dataPengisian.length; r++) {
        var rowP = dataPengisian[r];
        var idR = (rowP[1] || "").toString().trim(); // Kolum B: ID Rekod
        var n = (rowP[3] || "").toString().trim().toLowerCase(); // Kolum D: Nama Guru Dicerap
        var em = (rowP[4] || "").toString().trim().toLowerCase(); // Kolum E: E-mel Guru (DELIMa)
        var tarikhRaw = rowP[9]; // Kolum J: Tarikh Pencerapan
        var thnPencerapan = ekstrakTahunDaripadaTarikh_(tarikhRaw);

        // Hanya kira sebagai sudah hantar jika pencerapan berlaku pada tahun semasa
        if (thnPencerapan === tahunSemasa) {
          if (idR && n) {
            mapNama[n] = true;
          }
          if (idR && em) {
            mapEmel[em] = true;
          }

          // Semak juga emel dalam payload JSON lajur terakhir
          var jsonP = rowP[numCols - 1];
          if (jsonP && typeof jsonP === "string" && jsonP.startsWith("{")) {
            try {
              var parsedP = JSON.parse(jsonP);
              if (parsedP.guru && parsedP.guru.emel) {
                var emPayload = parsedP.guru.emel.toString().trim().toLowerCase();
                if (emPayload && idR) {
                  mapEmel[emPayload] = true;
                }
              }
            } catch (e) {}
          }
        }
      }
    }
  } catch (err) {}
  return { nama: mapNama, emel: mapEmel };
}

/**
 * Mengambil data borang awam untuk Mod Guru Kendiri:
 * - Maklumat sekolah (daripada ScriptProperties)
 * - Tahun semasa rasmi pelayan (currentYear & currentDate)
 * - Senarai pencerap sah dan aktif (Pencerap === 'ya' dan Aktif !== 'tidak' dalam sheet GURU)
 * - Senarai guru aktif yang BELUM mempunyai ID Rekod bagi tahun semasa (ditapis daripada PENGISIAN)
 */
function getPublicFormData() {
  try {
    var ss = getSpreadsheet_();
    var sekolah = getMaklumatSekolahDariProperties_();
    var currentYear = dapatkanTahunSemasa_();
    var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+8", "yyyy-MM-dd");

    // 1. Kenal pasti nama & emel guru yang telah pun membuat pengisian bagi TAHUN SEMASA
    var hantarMap = getGuruSudahHantarMap_(ss, currentYear);

    // 2. Baca sheet GURU
    var sheetGuru = getSheetGuru_(ss);
    var senaraiSemuaGuru = [];
    var senaraiGuruTersedia = [];
    var senaraiPencerap = [];

    if (sheetGuru && sheetGuru.getLastRow() >= 2) {
      var numRows = sheetGuru.getLastRow() - 1;
      var numCols = Math.max(sheetGuru.getLastColumn(), 7);
      var dataGuru = sheetGuru.getRange(2, 1, numRows, numCols).getValues();

      for (var i = 0; i < dataGuru.length; i++) {
        var guru = parseBarisGuru_(dataGuru[i]);

        if (guru.nama !== "") {
          var namaLower = guru.nama.toLowerCase();
          var emelLower = guru.emel.toLowerCase();
          var isAktif = guru.aktif !== "tidak" && guru.aktif !== false;
          var guruObj = {
            id: i + 1,
            emel: guru.emel,
            nama: guru.nama,
            jantina: guru.jantina,
            opsyen: guru.opsyen,
            pencerap: guru.pencerap,
            jawatan: guru.jawatan,
            aktif: guru.aktif || "ya",
          };
          senaraiSemuaGuru.push(guruObj);

          // Tapis pencerap (Hanya Pencerap yang AKTIF)
          if (guru.pencerap === "ya" && isAktif) {
            senaraiPencerap.push({
              id: senaraiPencerap.length + 1,
              emel: guru.emel,
              nama: guru.nama,
              jawatan: guru.jawatan || "Pencerap",
            });
          }

          // Tapis guru yang AKTIF dan BELUM membuat pengisian bagi tahun semasa
          if (isAktif) {
            var sudahHantar =
              hantarMap.nama[namaLower] ||
              (emelLower && hantarMap.emel[emelLower]);
            if (!sudahHantar) {
              senaraiGuruTersedia.push(guruObj);
            }
          }
        }
      }
    }

    return {
      success: true,
      currentYear: currentYear,
      currentDate: currentDate,
      sekolah: sekolah,
      senaraiPencerap: senaraiPencerap,
      senaraiGuruTersedia: senaraiGuruTersedia,
      senaraiSemuaGuru: senaraiSemuaGuru,
      senaraiGuru: senaraiSemuaGuru,
      literalTexts: dapatkanLiteralTeksDariProperties_(),
    };
  } catch (err) {
    return {
      success: false,
      error: err.message,
    };
  }
}

/**
 * Mengambil semua rekod pengisian daripada sheet 'PENGISIAN'
 * yang mempunyai 'Nama Pencerap' padan dengan nama guru pencerap yang sedang log masuk
 * dan berada dalam TAHUN SEMASA sahaja.
 */
function getRekodPengisianPencerap(namaPencerap) {
  try {
    var ss = getSpreadsheet_();
    var sheet = getSheetPengisian_(ss);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, senaraiRekod: [] };
    }

    var lastRow = sheet.getLastRow();
    var numCols = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var targetPencerap = (namaPencerap || "").toString().toLowerCase().trim();
    var senaraiRekod = [];
    var tahunSemasa = dapatkanTahunSemasa_();

    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      var idRekod = row[1] ? row[1].toString().trim() : "";
      var rowNamaGuru = row[3] ? row[3].toString().trim() : "";
      var rowNamaPencerap = row[7] ? row[7].toString().trim() : ""; // Kolum 8 (index 7): Nama Pencerap
      var tarikhRaw = row[9]; // Kolum J: Tarikh Pencerapan
      var thnRekod = ekstrakTahunDaripadaTarikh_(tarikhRaw);

      // Hanya paparkan rekod pencerapan bagi tahun semasa untuk pencerap
      if (thnRekod && thnRekod !== tahunSemasa) {
        continue;
      }

      // Padanan nama pencerap (boleh jadi format 'Nama' atau 'Nama (Jawatan)')
      var match = false;
      if (targetPencerap) {
        var cleanRowPencerap = rowNamaPencerap.toLowerCase();
        if (
          cleanRowPencerap === targetPencerap ||
          cleanRowPencerap.indexOf(targetPencerap) !== -1 ||
          targetPencerap.indexOf(cleanRowPencerap) !== -1
        ) {
          match = true;
        }
      }

      var jsonStr = row[numCols - 1];
      var parsed = null;
      if (jsonStr && typeof jsonStr === "string" && jsonStr.startsWith("{")) {
        try {
          parsed = JSON.parse(jsonStr);
          if (parsed && parsed.pencerap && parsed.pencerap.nama) {
            var pNama = parsed.pencerap.nama.toLowerCase().trim();
            if (
              pNama === targetPencerap ||
              pNama.indexOf(targetPencerap) !== -1 ||
              targetPencerap.indexOf(pNama) !== -1
            ) {
              match = true;
            }
          }
        } catch (e) {}
      }

      if (match) {
        if (parsed) {
          parsed.idRekod = idRekod;
          senaraiRekod.push(parsed);
        } else {
          senaraiRekod.push({
            idRekod: idRekod,
            guru: {
              nama: rowNamaGuru,
              emel: row[4] || cariEmelGuru_(rowNamaGuru),
              jantina: row[5] || "",
              opsyen: row[6] || "",
            },
            pencerap: {
              nama: rowNamaPencerap,
              jawatan: row[8] || "",
              tarikh: row[9] || "",
              masa: row[10] || "",
            },
            pdp: {
              mataPelajaran: row[11] || "",
              tajuk: row[12] || "",
              bilMurid: row[13] || "",
              tingkatan: row[14] || "",
              namaKelas: row[15] || "",
            },
          });
        }
      }
    }

    return {
      success: true,
      senaraiRekod: senaraiRekod,
    };
  } catch (err) {
    return {
      success: false,
      error: err.message,
    };
  }
}

/**
 * Memadamkan rekod penilaian pencerapan daripada sheet 'PENGISIAN'.
 */
function padamRekodPenilaian(idRekod) {
  try {
    if (!idRekod) throw new Error("ID Rekod diperlukan.");
    var ss = getSpreadsheet_();
    var sheet = getSheetPengisian_(ss);
    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error("Tiada rekod dalam pangkalan data.");
    }

    var lastRow = sheet.getLastRow();
    var idValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    var targetRow = -1;

    for (var r = 0; r < idValues.length; r++) {
      if (
        idValues[r][0] &&
        idValues[r][0].toString().trim() === idRekod.toString().trim()
      ) {
        targetRow = r + 2;
        break;
      }
    }

    if (targetRow > 0) {
      sheet.deleteRow(targetRow);
      try {
        CacheService.getScriptCache().remove("PUBLIC_FORM_DATA");
        CacheService.getScriptCache().remove("ISPPK_DATA_ANALISIS");
      } catch (e) {}
      return {
        success: true,
        idRekod: idRekod,
        message: "Rekod pencerapan berjaya dipadam.",
      };
    } else {
      throw new Error("Rekod dengan ID " + idRekod + " tidak dijumpai.");
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Mengambil data agregat dan statistik daripada tab 'PENGISIAN'.
 * CIRI KESELAMATAN & PRIVASI:
 * - Medan 'Nama Guru Dicerap' dan 'E-mel Guru (DELIMa)' TIDAK DIKEMBALIKAN sama sekali.
 * - Setiap pencerapan dilabelkan secara anonim (cth: 'Pencerapan #1', 'Pencerapan #2' dsb).
 * - Dilengkapi dengan Script Cache untuk akses sepantas kilat.
 */
function getDataAnalisis() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("ISPPK_DATA_ANALISIS");
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (eCache) {}
    }

    var ss = getSpreadsheet_();
    var sheet = getSheetPengisian_(ss);
    if (!sheet || sheet.getLastRow() < 2) {
      return {
        success: true,
        jumlahRekod: 0,
        rekodDinilaiM: 0,
        senaraiRekod: [],
        ringkasan: {
          purataPeratusK: 0,
          purataPeratusM: 0,
          jurangPersepsi: 0,
          kadarCemerlangBaik: 0,
          bilanganPencerap: 0,
          bilanganSubjek: 0
        },
        tahapKDist: { Cemerlang: 0, Baik: 0, Sederhana: 0, Minimum: 0, "Tidak Memenuhi": 0 },
        tahapMDist: { Cemerlang: 0, Baik: 0, Sederhana: 0, Minimum: 0, "Tidak Memenuhi": 0 },
        domainStats: {
          perancangan: { k: 0, m: 0 },
          pelaksanaan: { k: 0, m: 0 },
          refleksi: { k: 0, m: 0 },
          murid: { k: 0, m: 0 }
        },
        statistikPencerap: [],
        statistikSubjek: [],
        statistikTingkatan: [],
        statistikJantina: { Lelaki: 0, Perempuan: 0, Lain: 0 }
      };
    }

    var lastRow = sheet.getLastRow();
    var numCols = sheet.getLastColumn();
    var range = sheet.getRange(2, 1, lastRow - 1, numCols);
    var data = range.getValues();
    var displayData = range.getDisplayValues();

    var senaraiRekod = [];
    var totalPeratusK = 0;
    var totalPeratusM = 0;
    var countValidM = 0;
    var countCemerlangBaikM = 0;

    var sumDomain = {
      kPerancangan: 0, mPerancangan: 0,
      kPelaksanaan: 0, mPelaksanaan: 0,
      kRefleksi: 0, mRefleksi: 0,
      kMurid: 0, mMurid: 0
    };

    var tahapKDist = { Cemerlang: 0, Baik: 0, Sederhana: 0, Minimum: 0, "Tidak Memenuhi": 0 };
    var tahapMDist = { Cemerlang: 0, Baik: 0, Sederhana: 0, Minimum: 0, "Tidak Memenuhi": 0 };

    var pencerapMap = {};
    var subjekMap = {};
    var tingkatanMap = {};
    var jantinaMap = { Lelaki: 0, Perempuan: 0, Lain: 0 };

    var totalBilMurid = 0;
    var totalM_Lelaki = 0, countM_Lelaki = 0;
    var totalM_Perempuan = 0, countM_Perempuan = 0;
    var countLelaki = 0, countPerempuan = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var idRekod = (row[1] || "").toString().trim() || ("ISPPK-" + (i + 1));
      var modPenilaian = (row[2] || "").toString().trim();
      var jantina = (row[5] || "").toString().trim() || "Lelaki";
      var opsyen = (row[6] || "").toString().trim() || "-";

      var rawPencerap = (row[7] || "").toString().trim() || "Pentadbir";
      var pencerapNama = rawPencerap.replace(/\s*\([^)]*\)/g, "").trim();
      var extractedJawatan = "";
      var mJawatan = rawPencerap.match(/\(([^)]+)\)/);
      if (mJawatan) extractedJawatan = mJawatan[1].trim();
      var pencerapJawatan = (row[8] || "").toString().trim() || extractedJawatan || "Pencerap";

      var parsedObj = null;
      if (row[33]) {
        try { parsedObj = JSON.parse(row[33]); } catch (e) {}
      }

      var tarikhRaw = row[9];
      var tarikh = "";
      if (parsedObj && parsedObj.pencerap && parsedObj.pencerap.tarikh) {
        tarikh = String(parsedObj.pencerap.tarikh).trim();
      } else if (tarikhRaw instanceof Date) {
        tarikh = Utilities.formatDate(tarikhRaw, Session.getScriptTimeZone() || "GMT+8", "yyyy-MM-dd");
      } else {
        tarikh = (tarikhRaw || "").toString().trim();
      }
      var thn = ekstrakTahunDaripadaTarikh_(tarikhRaw);

      var masa = "";
      if (parsedObj && parsedObj.pencerap && parsedObj.pencerap.masa) {
        masa = String(parsedObj.pencerap.masa).trim();
      } else if (displayData && displayData[i] && displayData[i][10]) {
        masa = String(displayData[i][10]).trim();
      } else if (row[10] instanceof Date) {
        var hM = row[10].getHours();
        var mM = row[10].getMinutes();
        masa = (hM < 10 ? "0" + hM : "" + hM) + ":" + (mM < 10 ? "0" + mM : "" + mM);
      } else {
        masa = (row[10] || "").toString().trim();
      }
      if (/^\d{1,2}:\d{2}:\d{2}$/.test(masa)) masa = masa.substring(0, 5);
      if (/^\d:\d{2}$/.test(masa)) masa = "0" + masa;

      var subjek = (row[11] || "").toString().trim() || "Tidak Dinyatakan";
      var tajuk = (row[12] || "").toString().trim();
      var bilMurid = Number(row[13]) || 0;
      var tingkatan = (row[14] || "").toString().trim() || "Lain-lain";
      var kelas = (row[15] || "").toString().trim();

      var skPerancangan = Number(row[16]) || 0;
      var smPerancangan = Number(row[17]) || 0;
      var skPelaksanaan = Number(row[18]) || 0;
      var smPelaksanaan = Number(row[19]) || 0;
      var skRefleksi = Number(row[20]) || 0;
      var smRefleksi = Number(row[21]) || 0;
      var skMurid = Number(row[22]) || 0;
      var smMurid = Number(row[23]) || 0;

      var jumlahSkorK = Number(row[24]) || 0;
      var peratusK = parseFloat(row[25]) || 0;
      var tahapK = (row[26] || "").toString().trim();

      var jumlahSkorM = Number(row[27]) || 0;
      var peratusM = parseFloat(row[28]) || 0;
      var tahapM = (row[29] || "").toString().trim();

      totalPeratusK += peratusK;
      sumDomain.kPerancangan += skPerancangan;
      sumDomain.kPelaksanaan += skPelaksanaan;
      sumDomain.kRefleksi += skRefleksi;
      sumDomain.kMurid += skMurid;

      // Pengiraan Tahap Kendiri (K) berasaskan peratusan sebenar mengikut piawaian 5 tahap
      if (peratusK > 80) tahapKDist.Cemerlang++;
      else if (peratusK > 60) tahapKDist.Baik++;
      else if (peratusK > 50) tahapKDist.Sederhana++;
      else if (peratusK > 40) tahapKDist.Minimum++;
      else tahapKDist["Tidak Memenuhi"]++;

      // Pengiraan Tahap Muktamad (M) berasaskan peratusan sebenar mengikut piawaian 5 tahap
      if (peratusM > 0 || jumlahSkorM > 0) {
        countValidM++;
        totalPeratusM += peratusM;
        sumDomain.mPerancangan += smPerancangan;
        sumDomain.mPelaksanaan += smPelaksanaan;
        sumDomain.mRefleksi += smRefleksi;
        sumDomain.mMurid += smMurid;

        if (peratusM > 80) {
          tahapMDist.Cemerlang++;
          countCemerlangBaikM++;
        } else if (peratusM > 60) {
          tahapMDist.Baik++;
          countCemerlangBaikM++;
        } else if (peratusM > 50) {
          tahapMDist.Sederhana++;
        } else if (peratusM > 40) {
          tahapMDist.Minimum++;
        } else {
          tahapMDist["Tidak Memenuhi"]++;
        }
      }

      // Agregat Jantina & Saiz Kelas
      totalBilMurid += bilMurid;
      if (jantina.toLowerCase().indexOf("lelaki") !== -1) {
        countLelaki++;
        if (peratusM > 0) {
          totalM_Lelaki += peratusM;
          countM_Lelaki++;
        }
      } else {
        countPerempuan++;
        if (peratusM > 0) {
          totalM_Perempuan += peratusM;
          countM_Perempuan++;
        }
      }

      // Agregat Pencerap
      if (!pencerapMap[pencerapNama]) {
        pencerapMap[pencerapNama] = {
          nama: pencerapNama,
          jawatan: pencerapJawatan,
          bilanganPencerapan: 0,
          totalM: 0,
          purataM: 0
        };
      }
      pencerapMap[pencerapNama].bilanganPencerapan++;
      if (peratusM > 0) pencerapMap[pencerapNama].totalM += peratusM;

      // Agregat Subjek
      if (!subjekMap[subjek]) {
        subjekMap[subjek] = { subjek: subjek, bilangan: 0, totalM: 0, totalK: 0 };
      }
      subjekMap[subjek].bilangan++;
      subjekMap[subjek].totalK += peratusK;
      if (peratusM > 0) subjekMap[subjek].totalM += peratusM;

      // Agregat Tingkatan
      if (!tingkatanMap[tingkatan]) {
        tingkatanMap[tingkatan] = { tingkatan: tingkatan, bilangan: 0, totalM: 0 };
      }
      tingkatanMap[tingkatan].bilangan++;
      if (peratusM > 0) tingkatanMap[tingkatan].totalM += peratusM;

      // Ekstrak perincian 10 item murid daripada JSON payload jika wujud
      var itemMuridPayload = null;
      if (parsedObj && parsedObj.skorMurid) {
        itemMuridPayload = parsedObj.skorMurid;
      }

      // Rekod tertapis tanpa maklumat nama & emel guru!
      senaraiRekod.push({
        no: i + 1,
        labelAnonim: "Pencerapan #" + (i + 1),
        idRekod: idRekod,
        tarikh: tarikh,
        tahun: thn || dapatkanTahunSemasa_(),
        masa: masa,
        pencerap: pencerapNama,
        jawatan: pencerapJawatan,
        jantina: jantina,
        opsyen: opsyen,
        subjek: subjek,
        tingkatan: tingkatan,
        kelas: kelas,
        bilMurid: bilMurid,
        skorMurid: itemMuridPayload,
        skorK: {
          perancangan: skPerancangan,
          pelaksanaan: skPelaksanaan,
          refleksi: skRefleksi,
          murid: skMurid,
          jumlah: jumlahSkorK,
          peratus: peratusK,
          tahap: peratusK > 80 ? "Cemerlang" : peratusK > 60 ? "Baik" : peratusK > 50 ? "Sederhana" : peratusK > 40 ? "Minimum" : "Tidak Memenuhi"
        },
        skorM: {
          perancangan: smPerancangan,
          pelaksanaan: smPelaksanaan,
          refleksi: smRefleksi,
          murid: smMurid,
          jumlah: jumlahSkorM,
          peratus: peratusM,
          tahap: peratusM > 80 ? "Cemerlang" : peratusM > 60 ? "Baik" : peratusM > 50 ? "Sederhana" : peratusM > 40 ? "Minimum" : peratusM > 0 ? "Tidak Memenuhi" : "-"
        },
        jurang: peratusM > 0 ? Math.round((peratusM - peratusK) * 10) / 10 : 0
      });
    }

    var N = data.length;
    var purataPeratusK = N > 0 ? Math.round((totalPeratusK / N) * 10) / 10 : 0;
    var purataPeratusM = countValidM > 0 ? Math.round((totalPeratusM / countValidM) * 10) / 10 : 0;
    var jurangPersepsi = countValidM > 0 ? Math.round((purataPeratusM - purataPeratusK) * 10) / 10 : 0;
    var kadarCemerlangBaik = countValidM > 0 ? Math.round((countCemerlangBaikM / countValidM) * 100) : 0;

    var domainStats = {
      perancangan: {
        k: N > 0 ? Math.round((sumDomain.kPerancangan / N) * 10) / 10 : 0,
        m: countValidM > 0 ? Math.round((sumDomain.mPerancangan / countValidM) * 10) / 10 : 0
      },
      pelaksanaan: {
        k: N > 0 ? Math.round((sumDomain.kPelaksanaan / N) * 10) / 10 : 0,
        m: countValidM > 0 ? Math.round((sumDomain.mPelaksanaan / countValidM) * 10) / 10 : 0
      },
      refleksi: {
        k: N > 0 ? Math.round((sumDomain.kRefleksi / N) * 10) / 10 : 0,
        m: countValidM > 0 ? Math.round((sumDomain.mRefleksi / countValidM) * 10) / 10 : 0
      },
      murid: {
        k: N > 0 ? Math.round((sumDomain.kMurid / N) * 10) / 10 : 0,
        m: countValidM > 0 ? Math.round((sumDomain.mMurid / countValidM) * 10) / 10 : 0
      }
    };

    var statistikPencerap = [];
    for (var p in pencerapMap) {
      var itemP = pencerapMap[p];
      itemP.purataM = itemP.bilanganPencerapan > 0 && itemP.totalM > 0 ? Math.round((itemP.totalM / itemP.bilanganPencerapan) * 10) / 10 : 0;
      delete itemP.totalM;
      statistikPencerap.push(itemP);
    }
    statistikPencerap.sort(function (a, b) { return b.bilanganPencerapan - a.bilanganPencerapan; });

    var statistikSubjek = [];
    for (var s in subjekMap) {
      var itemS = subjekMap[s];
      itemS.purataK = itemS.bilangan > 0 ? Math.round((itemS.totalK / itemS.bilangan) * 10) / 10 : 0;
      itemS.purataM = itemS.bilangan > 0 && itemS.totalM > 0 ? Math.round((itemS.totalM / itemS.bilangan) * 10) / 10 : 0;
      delete itemS.totalK;
      delete itemS.totalM;
      statistikSubjek.push(itemS);
    }
    statistikSubjek.sort(function (a, b) { return b.bilangan - a.bilangan; });

    var statistikTingkatan = [];
    for (var t in tingkatanMap) {
      var itemT = tingkatanMap[t];
      itemT.purataM = itemT.bilangan > 0 && itemT.totalM > 0 ? Math.round((itemT.totalM / itemT.bilangan) * 10) / 10 : 0;
      delete itemT.totalM;
      statistikTingkatan.push(itemT);
    }
    statistikTingkatan.sort(function (a, b) { return b.bilangan - a.bilangan; });

    // Kumpul semua tahun unik yang wujud dalam pangkalan data
    var tahunSemasa = dapatkanTahunSemasa_();
    var tahunSet = {};
    tahunSet[tahunSemasa] = true;
    for (var k = 0; k < data.length; k++) {
      var tK = ekstrakTahunDaripadaTarikh_(data[k][9]);
      if (tK) tahunSet[tK] = true;
    }
    var senaraiTahun = Object.keys(tahunSet).sort(function (a, b) {
      return Number(b) - Number(a);
    });

    var output = {
      success: true,
      tahunSemasa: tahunSemasa,
      senaraiTahun: senaraiTahun,
      jumlahRekod: N,
      rekodDinilaiM: countValidM,
      senaraiRekod: senaraiRekod,
      ringkasan: {
        purataPeratusK: purataPeratusK,
        purataPeratusM: purataPeratusM,
        jurangPersepsi: jurangPersepsi,
        kadarCemerlangBaik: kadarCemerlangBaik,
        bilanganPencerap: statistikPencerap.length,
        bilanganSubjek: statistikSubjek.length
      },
      tahapKDist: tahapKDist,
      tahapMDist: tahapMDist,
      domainStats: domainStats,
      statistikPencerap: statistikPencerap,
      statistikSubjek: statistikSubjek,
      statistikTingkatan: statistikTingkatan,
      statistikJantina: {
        Lelaki: {
          bilangan: countLelaki,
          purataM: countM_Lelaki > 0 ? Math.round((totalM_Lelaki / countM_Lelaki) * 10) / 10 : 0
        },
        Perempuan: {
          bilangan: countPerempuan,
          purataM: countM_Perempuan > 0 ? Math.round((totalM_Perempuan / countM_Perempuan) * 10) / 10 : 0
        }
      },
      statistikKelas: {
        totalMurid: totalBilMurid,
        purataMurid: N > 0 ? Math.round((totalBilMurid / N) * 10) / 10 : 0
      }
    };
    try {
      cache.put("ISPPK_DATA_ANALISIS", JSON.stringify(output), 600); // 10 minit
    } catch (ePut) {}

    return output;
  } catch (err) {
    return {
      success: false,
      error: err.message
    };
  }
}

// ============================================================================
// 4. OPERASI CRUD GURU (KHAS UNTUK PENCERAP YANG LOG MASUK)
// ============================================================================

/**
 * Menambah rekod guru baharu ke dalam sheet GURU.
 */
function tambahGuru(guruData) {
  try {
    var ss = getSpreadsheet_();
    var sheet = getSheetGuru_(ss);
    if (!sheet) throw new Error("Sheet 'GURU' tidak dijumpai.");

    var emel = (guruData.emel || "").trim();
    var nama = (guruData.nama || "").trim();
    var jantina = (guruData.jantina || "Lelaki").trim();
    var opsyen = (guruData.opsyen || "").trim();
    var pencerap =
      (guruData.pencerap || "tidak").toLowerCase().trim() === "ya"
        ? "ya"
        : "tidak";
    var jawatan = (guruData.jawatan || "").trim();
    var aktif =
      (guruData.aktif || "ya").toString().toLowerCase().trim() === "tidak" || guruData.aktif === false
        ? "tidak"
        : "ya";

    if (!nama) throw new Error("Nama guru diperlukan.");

    // Susunan Kolum: Emel (A), Nama Guru (B), Jawatan (C), Jantina (D), Opsyen (E), Pencerap (F), Aktif (G)
    sheet.appendRow([emel, nama, jawatan, jantina, opsyen, pencerap, aktif]);
    segerakPencerapKeProperties_(ss);
    try {
      CacheService.getScriptCache().remove("PUBLIC_FORM_DATA");
    } catch (e) {}

    return getInitialData();
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Mengemaskini rekod guru sedia ada dalam sheet GURU.
 */
function kemaskiniGuru(guruData, indexRow) {
  try {
    var ss = getSpreadsheet_();
    var sheet = getSheetGuru_(ss);
    if (!sheet) throw new Error("Sheet 'GURU' tidak dijumpai.");

    var rowNumber = Number(indexRow);
    if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
      throw new Error("Nombor baris guru tidak sah.");
    }

    var emel = (guruData.emel || "").trim();
    var nama = (guruData.nama || "").trim();
    var jantina = (guruData.jantina || "Lelaki").trim();
    var opsyen = (guruData.opsyen || "").trim();
    var pencerap =
      (guruData.pencerap || "tidak").toLowerCase().trim() === "ya"
        ? "ya"
        : "tidak";
    var jawatan = (guruData.jawatan || "").trim();
    var aktif =
      (guruData.aktif || "ya").toString().toLowerCase().trim() === "tidak" || guruData.aktif === false
        ? "tidak"
        : "ya";

    // Susunan Kolum: Emel (A), Nama Guru (B), Jawatan (C), Jantina (D), Opsyen (E), Pencerap (F), Aktif (G)
    sheet
      .getRange(rowNumber, 1, 1, 7)
      .setValues([[emel, nama, jawatan, jantina, opsyen, pencerap, aktif]]);
    segerakPencerapKeProperties_(ss);
    try {
      CacheService.getScriptCache().remove("PUBLIC_FORM_DATA");
    } catch (e) {}

    return getInitialData();
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Memadam rekod guru daripada sheet GURU.
 */
function padamGuru(indexRow) {
  try {
    var ss = getSpreadsheet_();
    var sheet = getSheetGuru_(ss);
    if (!sheet) throw new Error("Sheet 'GURU' tidak dijumpai.");

    var rowNumber = Number(indexRow);
    if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
      throw new Error("Nombor baris guru tidak sah.");
    }

    sheet.deleteRow(rowNumber);
    segerakPencerapKeProperties_(ss);
    try {
      CacheService.getScriptCache().remove("PUBLIC_FORM_DATA");
    } catch (e) {}

    return getInitialData();
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================================================
// 5. PENYIMPANAN PENILAIAN & PENGHANTARAN EMEL URL ID REKOD
// ============================================================================

/**
 * Menyimpan data penilaian ke dalam sheet 'PENGISIAN'
 * dan menghantar emel notifikasi beserta URL ID Rekod bagi pengisian kendiri guru.
 */
function simpanPenilaian(payload) {
  try {
    var ss = getSpreadsheet_();
    var sheet = getSheetPengisian_(ss);

    // Semakan Tahun Semasa: Pengisian dan penyuntingan hanya dibenarkan bagi tahun semasa
    var thnSemasa = dapatkanTahunSemasa_();
    var tarikhInput = payload && payload.pencerap && payload.pencerap.tarikh ? payload.pencerap.tarikh : "";
    var thnRekod = ekstrakTahunDaripadaTarikh_(tarikhInput);
    if (thnRekod && thnRekod !== thnSemasa) {
      throw new Error("Pengisian dan penyuntingan maklumat hanya dibenarkan bagi tahun semasa (" + thnSemasa + ") sahaja. Tarikh pencerapan adalah " + (payload.pencerap.tarikh || "-") + ".");
    }

    // 1. Kemas kini E-mel Guru (DELIMa) dan Opsyen dalam tabsheet 'GURU' jika berubah atau ditetapkan semula
    var guruSheetUpdate = { dikemaskini: false };
    if (payload && payload.guru && payload.guru.nama) {
      guruSheetUpdate = kemaskiniEmelDanOpsyenGuru_(ss, payload.guru);
    }

    var now = new Date();
    var timestampStr = Utilities.formatDate(
      now,
      Session.getScriptTimeZone() || "Asia/Kuala_Lumpur",
      "yyyy-MM-dd HH:mm:ss",
    );

    // Jana atau guna ID Rekod sedia ada
    var idRekod =
      payload.idRekod ||
      "ISPPK-" +
        Utilities.formatDate(now, "GMT+8", "yyyyMMdd") +
        "-" +
        Math.floor(1000 + Math.random() * 9000);
    payload.idRekod = idRekod;

    var rowData = [
      timestampStr,
      idRekod,
      payload.mod || "Guru",
      payload.guru ? payload.guru.nama : "",
      payload.guru ? payload.guru.emel : "",
      payload.guru ? payload.guru.jantina : "",
      payload.guru ? payload.guru.opsyen : "",
      payload.pencerap ? payload.pencerap.nama : "",
      payload.pencerap ? payload.pencerap.jawatan : "",
      payload.pencerap ? payload.pencerap.tarikh : "",
      payload.pencerap ? payload.pencerap.masa : "",
      payload.pdp ? payload.pdp.mataPelajaran : "",
      payload.pdp ? payload.pdp.tajuk : "",
      payload.pdp ? payload.pdp.bilMurid : "",
      payload.pdp ? payload.pdp.tingkatan : "",
      payload.pdp ? payload.pdp.namaKelas : "",
      // Skor Domain Guru
      payload.skorK_Perancangan || 0,
      payload.skorM_Perancangan || 0,
      payload.skorK_Pelaksanaan || 0,
      payload.skorM_Pelaksanaan || 0,
      payload.skorK_Refleksi || 0,
      payload.skorM_Refleksi || 0,
      // Skor Murid
      payload.skorK_Murid || 0,
      payload.skorM_Murid || 0,
      // Jumlah & Peratus
      payload.jumlahSkorK || 0,
      payload.peratusK || 0,
      payload.tahapK || "",
      payload.jumlahSkorM || 0,
      payload.peratusM || 0,
      payload.tahapM || "",
      // Refleksi Kualitatif
      payload.refleksiGuru1 || "",
      payload.refleksiGuru2 || "",
      payload.rumusanPencerap || "",
      JSON.stringify(payload),
    ];

    // Semak sama ada ID Rekod sudah wujud untuk dikemaskini
    var lastRow = sheet.getLastRow();
    var existingRow = -1;
    if (lastRow >= 2) {
      var checkRange = sheet.getRange(2, 2, lastRow - 1, 9).getValues(); // Dari Kolum B (ID Rekod) hingga Kolum J (Tarikh Pencerapan, index 8)
      for (var r = 0; r < checkRange.length; r++) {
        if (
          checkRange[r][0] &&
          checkRange[r][0].toString().trim() === idRekod.toString().trim()
        ) {
          var existingDate = checkRange[r][8]; // Kolum J: Tarikh Pencerapan
          var thnExisting = ekstrakTahunDaripadaTarikh_(existingDate);
          if (thnExisting && thnExisting !== thnSemasa) {
            throw new Error("Rekod pencerapan arkib bagi tahun " + thnExisting + " tidak boleh disunting atau dikemaskini.");
          }
          existingRow = r + 2;
          break;
        }
      }
    }

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    // Kosongkan cache pangkalan data awam agar senarai guru dan analisis sentiasa segar
    try {
      CacheService.getScriptCache().remove("PUBLIC_FORM_DATA");
      CacheService.getScriptCache().remove("ISPPK_DATA_ANALISIS");
    } catch (e) {}

    // ------------------------------------------------------------------------
    // PENGHANTARAN EMEL KEPADA GURU (Jika disimpan oleh pengguna tidak log masuk / mod kendiri)
    // ------------------------------------------------------------------------
    var emelStatus = { dihantar: false, penerima: "", ralat: null };
    var isKendiri =
      !payload.isPencerap && (payload.mod === "Guru" || payload.mod === "guru");

    if (isKendiri) {
      try {
        // Keutamaan 1: Emel yang diisi/disahkan oleh guru dalam borang
        var recipientEmail = "";
        if (
          payload.guru &&
          payload.guru.emel &&
          payload.guru.emel.indexOf("@") !== -1
        ) {
          recipientEmail = payload.guru.emel.trim();
        }
        // Keutamaan 2: Carian nama guru dalam tab GURU jika tiada dalam borang
        if (!recipientEmail && payload.guru && payload.guru.nama) {
          recipientEmail = cariEmelGuru_(payload.guru.nama);
        }

        if (recipientEmail) {
          var resEmel = hantarEmelNotifikasiGuru_(recipientEmail, payload, idRekod);
          if (resEmel && resEmel.success === false) {
            emelStatus.ralat = resEmel.error || "Ralat penghantaran emel.";
          } else {
            emelStatus.dihantar = true;
            emelStatus.penerima = recipientEmail;
          }
        } else {
          emelStatus.ralat = "Alamat e-mel guru tidak sah atau tidak diisi.";
        }
      } catch (eMailErr) {
        emelStatus.ralat = eMailErr.message;
      }
    }

    return {
      success: true,
      idRekod: idRekod,
      timestamp: timestampStr,
      action: existingRow > 0 ? "dikemaskini" : "disimpan",
      emelStatus: emelStatus,
      guruSheetUpdate: guruSheetUpdate,
    };
  } catch (error) {
    return {
      success: false,
      error: error.message,
    };
  }
}

/**
 * Mengemaskini E-mel Guru (DELIMa) dan Opsyen dalam tabsheet 'GURU'
 * sekiranya nilai tersebut berubah atau ditetapkan semula oleh guru semasa simpanan.
 */
function kemaskiniEmelDanOpsyenGuru_(ss, guruInput) {
  try {
    if (!guruInput || !guruInput.nama) return { dikemaskini: false };
    var sheet = getSheetGuru_(ss);
    if (!sheet || sheet.getLastRow() < 2) return { dikemaskini: false };

    var targetNama = guruInput.nama.toString().trim().toLowerCase();
    var newEmel = (guruInput.emel || "").toString().trim();
    var newOpsyen = (guruInput.opsyen || "").toString().trim();

    var numRows = sheet.getLastRow() - 1;
    var numCols = Math.max(sheet.getLastColumn(), 6);
    var data = sheet.getRange(2, 1, numRows, numCols).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var parsed = parseBarisGuru_(row);
      if (parsed.nama && parsed.nama.toLowerCase() === targetNama) {
        var rowNum = i + 2;
        var emelChanged = newEmel !== "" && newEmel !== parsed.emel;
        var opsyenChanged = newOpsyen !== "" && newOpsyen !== parsed.opsyen;

        if (emelChanged || opsyenChanged) {
          var c2Lower = (row[2] || "").toString().trim().toLowerCase();
          var isShifted = c2Lower === "lelaki" || c2Lower === "perempuan";

          if (emelChanged) {
            sheet.getRange(rowNum, 1).setValue(newEmel);
          }

          if (opsyenChanged) {
            if (isShifted) {
              // Format teranjak: [Emel (1), Nama (2), Jantina (3), Opsyen (4), Pencerap (5), Jawatan (6)]
              sheet.getRange(rowNum, 4).setValue(newOpsyen);
            } else {
              // Format standard: [Emel (1), Nama (2), Jawatan (3), Jantina (4), Opsyen (5), Pencerap (6)]
              var c4Lower = (row[4] || "").toString().trim().toLowerCase();
              var c5Val = (row[5] || "").toString().trim();
              if (c4Lower === "ya" && !c5Val) {
                sheet.getRange(rowNum, 6).setValue("ya");
              }
              sheet.getRange(rowNum, 5).setValue(newOpsyen);
            }
          }

          if (parsed.pencerap === "ya" && emelChanged) {
            segerakPencerapKeProperties_(ss);
          }

          return {
            dikemaskini: true,
            emel: emelChanged ? newEmel : parsed.emel,
            opsyen: opsyenChanged ? newOpsyen : parsed.opsyen,
          };
        }
        return { dikemaskini: false };
      }
    }
    return { dikemaskini: false };
  } catch (err) {
    return { dikemaskini: false, ralat: err.message };
  }
}

/**
 * Mencari emel guru dalam sheet GURU berdasarkan nama guru.
 */
function cariEmelGuru_(namaGuru) {
  if (!namaGuru) return "";
  var ss = getSpreadsheet_();
  var sheet = getSheetGuru_(ss);
  if (!sheet || sheet.getLastRow() < 2) return "";

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var cleanTarget = namaGuru.toLowerCase().trim();

  for (var i = 0; i < data.length; i++) {
    var email = (data[i][0] || "").toString().trim();
    var name = (data[i][1] || "").toString().toLowerCase().trim();
    if (name === cleanTarget && email.indexOf("@") !== -1) {
      return email;
    }
  }
  return "";
}

/**
 * Menghantar emel pengesahan dan pautan URL ID Rekod kepada guru.
 * Menggunakan akaun Owner Relay dan menetapkan noReply: true.
 */
function hantarEmelNotifikasiGuru_(recipientEmail, payload, idRekod) {
  try {
    var urls = getDeploymentUrls();
    var baseUrl = urls.webAppUrl || urls.ownerUrl || getConfigValue_("WEBAPP_URL", "") || "";
    if (!baseUrl) {
      console.warn("⚠️ Peringatan: WEBAPP_URL belum dikonfigurasi di tab TETAPAN B4. Pautan semakan dalam emel mungkin tidak lengkap.");
    }
    var editUrl =
      baseUrl +
      (baseUrl.indexOf("?") === -1 ? "?" : "&") +
      "idRekod=" +
      encodeURIComponent(idRekod);
    var namaGuru = payload.guru ? payload.guru.nama : "Guru";
    var tahunSemasa = new Date().getFullYear();
    var subjek = `[Instrumen Bestari Penilaian PdP] Salinan Pengisian Penilaian Kendiri - ${namaGuru} (ID: ${idRekod})`;

    var bodyHtml = `
      <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
        <div style="background: linear-gradient(135deg, #0369a1, #1e1b4b); color: #ffffff; padding: 24px; text-align: center;">
          <span style="display: inline-block; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 9999px;">Kementerian Pendidikan Malaysia</span>
          <h2 style="margin: 12px 0 4px; font-size: 20px; font-weight: 800;">Instrumen ISPPK KBAT ${tahunSemasa}</h2>
          <p style="margin: 0; font-size: 13px; color: #bae6fd;">Pengesahan Penilaian Kendiri Guru</p>
        </div>

        <div style="padding: 24px; color: #334155; line-height: 1.6; font-size: 14px;">
          <p>Salam sejahtera <b>${namaGuru}</b>,</p>
          <p>Pengisian penilaian kendiri anda bagi instrumen pembudayaan KBAT dalam PdP telah berjaya direkodkan ke dalam pangkalan data sekolah.</p>

          <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; margin: 16px 0;">
            <table style="width: 100%; font-size: 13px; border-collapse: collapse;">
              <tr><td style="padding: 4px 0; color: #64748b; width: 140px;">ID Rekod:</td><td style="padding: 4px 0; font-weight: bold; font-family: monospace; color: #0284c7;">${idRekod}</td></tr>
              <tr><td style="padding: 4px 0; color: #64748b;">Mata Pelajaran:</td><td style="padding: 4px 0; font-weight: 600;">${payload.pdp ? payload.pdp.mataPelajaran || "-" : "-"}</td></tr>
              <tr><td style="padding: 4px 0; color: #64748b;">Skor Kendiri:</td><td style="padding: 4px 0; font-weight: 600;">${payload.jumlahSkorK || 0} / 100 (${payload.peratusK || 0}%)</td></tr>
              <tr><td style="padding: 4px 0; color: #64748b;">Tahap Pembudayaan:</td><td style="padding: 4px 0; font-weight: 600; color: #0369a1;">${payload.tahapK || "-"}</td></tr>
            </table>
          </div>

          <div style="background: #eff6ff; border-left: 4px solid #0284c7; padding: 14px 16px; border-radius: 0 8px 8px 0; margin: 20px 0;">
            <b style="color: #1e40af; display: block; margin-bottom: 4px;">Pindaan / Semakan Semula Pengisian:</b>
            <p style="margin: 0; font-size: 13px; color: #1e3a8a;">
              Sekiranya anda ingin mengubah semula pengisian atau menyemak skor kendiri ini, sila klik butang di bawah atau gunakan pautan yang dikepilkan:
            </p>
            <div style="text-align: center; margin-top: 14px;">
              <a href="${editUrl}" target="_blank" style="display: inline-block; background: #0284c7; color: #ffffff; text-decoration: none; padding: 10px 22px; font-size: 13px; font-weight: bold; border-radius: 8px; box-shadow: 0 2px 4px rgba(2,132,199,0.3);">
                ✏️ Buka & Ubah Semula Pengisian
              </a>
            </div>
          </div>

          <p style="font-size: 12px; color: #94a3b8; word-break: break-all; margin-top: 16px;">
            Pautan terus: <br><a href="${editUrl}" style="color: #0284c7;">${editUrl}</a>
          </p>
        </div>

        <div style="background: #f1f5f9; padding: 16px; text-align: center; font-size: 11px; color: #64748b; border-top: 1px solid #e2e8f0;">
          Emel ini dijana secara automatik oleh Aplikasi Instrumen Bestari Penilaian PdP. Sila jangan balas emel ini.
        </div>
      </div>
    `;

    var mailOptions = {
      to: recipientEmail,
      subject: subjek,
      htmlBody: bodyHtml,
      name: "Instrumen Bestari Penilaian PdP",
      noReply: true,
    };

    // Percubaan 1: Hantar dengan noReply: true
    try {
      MailApp.sendEmail(mailOptions);
      return { success: true };
    } catch (errNoReply) {
      // Percubaan 2: Jika akaun Gmail pengguna biasa tidak menyokong noReply: true
      try {
        delete mailOptions.noReply;
        MailApp.sendEmail(mailOptions);
        return { success: true };
      } catch (errRetry) {
        return { success: false, error: errRetry.message };
      }
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================================================
// 6. CARIAN DAN PEMUATAN REKOD
// ============================================================================

/**
 * Mencari rekod terdahulu berdasarkan ID rekod, nama guru, atau emel.
 */
function cariRekod(kataKunci) {
  try {
    var ss = getSpreadsheet_();
    var sheet = getSheetPengisian_(ss);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, rekod: null };
    }

    var lastRow = sheet.getLastRow();
    var numCols = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var found = null;
    var target = (kataKunci || "").toString().toLowerCase().trim();
    if (!target) {
      return { success: true, rekod: null };
    }

    // Jika kataKunci ialah emel, cari nama guru yang berkaitan dalam sheet GURU
    var namaDariEmel = "";
    var sheetGuru = getSheetGuru_(ss);
    if (sheetGuru && sheetGuru.getLastRow() >= 2) {
      var dataGuru = sheetGuru
        .getRange(2, 1, sheetGuru.getLastRow() - 1, 2)
        .getValues();
      for (var g = 0; g < dataGuru.length; g++) {
        var gEmel = (dataGuru[g][0] || "").toString().toLowerCase().trim();
        if (gEmel === target) {
          namaDariEmel = (dataGuru[g][1] || "").toString().toLowerCase().trim();
          break;
        }
      }
    }

    // Cari dari rekod paling terkini (bawah ke atas)
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      var idRekod = row[1] ? row[1].toString().trim() : "";
      var namaGuru = row[3] ? row[3].toString().trim() : "";
      var emelGuru = row[4] ? row[4].toString().trim() : "";
      var jsonStr = row[numCols - 1];

      var match = false;
      if (idRekod && idRekod.toLowerCase() === target) match = true;
      if (namaGuru && namaGuru.toLowerCase() === target) match = true;
      if (emelGuru && emelGuru.toLowerCase() === target) match = true;
      if (namaDariEmel && namaGuru && namaGuru.toLowerCase() === namaDariEmel)
        match = true;

      // Semak emel dalam payload JSON lajur terakhir
      if (!match && jsonStr && typeof jsonStr === "string") {
        if (jsonStr.toLowerCase().indexOf(target) !== -1) {
          try {
            var parsedTemp = JSON.parse(jsonStr);
            if (
              parsedTemp.guru &&
              parsedTemp.guru.emel &&
              parsedTemp.guru.emel.toString().toLowerCase().trim() === target
            ) {
              match = true;
            }
          } catch (e) {}
        }
      }

      if (
        match &&
        jsonStr &&
        typeof jsonStr === "string" &&
        jsonStr.startsWith("{")
      ) {
        try {
          found = JSON.parse(jsonStr);
          found.idRekod = idRekod;
          break;
        } catch (e) {}
      }
    }

    return {
      success: true,
      rekod: found,
    };
  } catch (error) {
    return {
      success: false,
      error: error.message,
    };
  }
}


/**
 * Fungsi Pembantu Piawai: Penentuan Ringkasan Tahap KBAT
 * @param {number} p - Peratusan (0 - 100)
 * @param {number} [total] - Jumlah skor
 * @return {string} Ringkasan tahap (Cemerlang, Baik, Sederhana, Minimum, Tidak Memenuhi)
 */
function getTahapRingkasan_(p, total) {
  if (total === 0) return "-";
  var num = Number(p) || 0;
  try {
    var lt = dapatkanLiteralTeksDariProperties_();
    if (lt && lt.tahapKbat && lt.tahapKbat.length > 0) {
      for (var i = 0; i < lt.tahapKbat.length; i++) {
        var t = lt.tahapKbat[i];
        if (t.skor_min > 0 && num <= t.skor_max) return t.ringkasan;
      }
      return lt.tahapKbat[lt.tahapKbat.length - 1].ringkasan;
    }
  } catch (e) {}
  if (total === 0) return "-";
  var num = Number(p) || 0;
  if (num <= 40) return "Tidak Memenuhi";
  if (num <= 50) return "Minimum";
  if (num <= 60) return "Sederhana";
  if (num <= 80) return "Baik";
  return "Cemerlang";
}

// ============================================================================
// 10. INTEGRASI GOOGLE FORMS JPN (ISPPK KBAT 2026)
// ============================================================================

/**
 * Menjana pautan selamat ke modul IBPPdP_HantarGF_ISPPK bagi pencerap berdaftar.
 * @param {string} email - Emel DELIMa pencerap
 * @return {Object} Status dan URL sasaran
 */
function dapatkanUrlModulHantarGF(email) {
  try {
    var userEmail = email || Session.getActiveUser().getEmail();
    if (!userEmail) throw new Error("Emel pencerap diperlukan.");
    var semakan = semakStatusPencerap_(userEmail);
    if (!semakan.isPencerap) {
      return { success: false, error: "Akses Ditolak: Emel ini (" + userEmail + ") bukan pencerap berdaftar." };
    }
    var urls = getDeploymentUrls();
    var ownerUrl = urls.ownerUrl || urls.webAppUrl || getConfigValue_("WEBAPP_URL", "") || "";
    if (!ownerUrl) {
      return {
        success: false,
        error: "URL Web App belum ditetapkan. Sila salin URL Web App dari Apps Script UI (Deploy > Manage deployments), tampal ke sel B4 tab TETAPAN dan jalankan Setup Instrumen Bestari.",
      };
    }
    var targetUrl =
      ownerUrl +
      (ownerUrl.indexOf("?") === -1 ? "?" : "&") +
      "page=hantarGF";

    return { success: true, url: targetUrl };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Memastikan sama ada URL sepadan dengan URL Google Forms yang sah.
 * @param {string} url
 * @return {boolean}
 */
function isValidFormsUrl(url) {
  if (!url) return false;
  var u = url.trim();
  // Terima format forms.gle, goo.gl, atau docs.google.com/forms/d/e/... (termasuk sokongan multi-login /u/0/ dan pelbagai format)
  var isShort = /^(https?:\/\/)?(forms\.gle|goo\.gl)\/[a-zA-Z0-9_-]+/i.test(u);
  var isDocs = /^https:\/\/docs\.google\.com\/forms\/(?:u\/\d+\/)?d\/(?:e\/[^\/?#]+|[^\/?#]+)(?:\/(?:viewform|formResponse|edit))?(?:\?.*)?(?:#.*)?$/i.test(u);
  return isShort || isDocs;
}

/**
 * Meleraikan pautan pendek Google Forms (forms.gle / goo.gl) kepada URL sebenar.
 * Mengikuti header Location sehingga 6 lelaran redirect.
 * @param {string} url
 * @return {string} Real URL
 */
function resolveShortFormsUrl(url) {
  try {
    var u = (url || "").trim();
    if (!u) return "";
    if (/^https:\/\/docs\.google\.com\/forms\/d\//i.test(u)) {
      return u;
    }
    var parsed = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: false });
    var headers = parsed.getAllHeaders ? parsed.getAllHeaders() : parsed.getHeaders();
    var loc = headers["Location"] || headers["location"];
    var current = u;
    var attempts = 0;
    while (loc && attempts < 6) {
      if (loc.startsWith("/")) {
        var m = current.match(/^(https?:\/\/[^/]+)/);
        loc = (m ? m[1] : "") + loc;
      }
      current = loc;
      if (/^https:\/\/docs\.google\.com\/forms\/d\//i.test(current)) {
        return current;
      }
      var resp = UrlFetchApp.fetch(current, { muteHttpExceptions: true, followRedirects: false });
      var h = resp.getAllHeaders ? resp.getAllHeaders() : resp.getHeaders();
      loc = h["Location"] || h["location"];
      attempts++;
    }
    return current;
  } catch (err) {
    console.warn("resolveShortFormsUrl err", err);
    return url;
  }
}

/**
 * Mengambil kandungan HTML halaman borang Google Forms.
 * @param {string} url
 * @return {string}
 */
function fetchUrlContent(url) {
  var resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    }
  });
  var code = resp.getResponseCode();
  if (code >= 400) {
    throw new Error("Gagal mengambil borang Google (HTTP " + code + ")");
  }
  return resp.getContentText();
}

/**
 * Mengekstrak dan menghuraikan pembolehubah FB_PUBLIC_LOAD_DATA_ daripada HTML borang.
 * @param {string} html
 * @return {Array|null}
 */
function extractFBPublicLoadData(html) {
  var re = /FB_PUBLIC_LOAD_DATA_\s*=\s*(\[[\s\S]*?\])\s*;/m;
  var m = html.match(re);
  if (!m) return null;
  var rhs = m[1];
  try {
    return JSON.parse(rhs);
  } catch (e) {
    try {
      var fn = new Function("return " + rhs + ";");
      return fn();
    } catch (err) {
      console.error("Gagal menghurai FB_PUBLIC_LOAD_DATA_:", err);
      return null;
    }
  }
}

/**
 * Membersihkan kunci objek agar serasi dengan Caja / Google Apps Script RPC.
 * Peraturan Caja melarang keras sebarang kunci objek pengguna (User Properties)
 * yang bermula atau berakhir dengan '__' (dunder), contohnya '__other_option__'.
 * @param {string} key
 * @return {string}
 */
function sanitizeGasPropertyKey_(key) {
  if (!key) return "pilihan";
  var str = String(key).trim();
  if (str === "__other_option__") return "Lain-lain";
  // Buang awalan atau akhiran dunder '__'
  str = str.replace(/^__+/, "").replace(/__+$/, "");
  return str || "pilihan";
}

/**
 * Menapis dan membersihkan struktur data (Objek & Array) secara rekursif sebelum
 * dipulangkan kepada pelayar melalui google.script.run bagi mengelakkan ralat:
 * "User properties cannot end with '__'. Failed on property: ..."
 * @param {*} data
 * @param {Set} [seen]
 * @return {*}
 */
function sanitizeForGasRpc_(data, seen) {
  if (data === null || data === undefined) return data;
  if (typeof data !== "object") return data;

  if (!seen) seen = new Set();
  if (seen.has(data)) return data; // Cegah rujukan berpusing (circular reference)
  seen.add(data);

  if (Array.isArray(data)) {
    return data.map(function (item) {
      return sanitizeForGasRpc_(item, seen);
    });
  }

  var cleanObj = {};
  for (var k in data) {
    if (Object.prototype.hasOwnProperty.call(data, k)) {
      var safeKey = sanitizeGasPropertyKey_(k);
      cleanObj[safeKey] = sanitizeForGasRpc_(data[k], seen);
    }
  }
  return cleanObj;
}

/**
 * Memproses data mentah FB_PUBLIC_LOAD_DATA_ kepada senarai elemen soalan tersusun.
 * Mengesan seksyen borang (item[3] === 8 / 6) untuk memperincikan label soalan mengikut seksyen
 * seperti 'KOD DAN NAMA SEKOLAH: PPD BACHOK' atau 'KOD DAN NAMA SEKOLAH: PPD GUA MUSANG'.
 * @param {Array} formData
 * @return {Object}
 */
function processFBData(formData) {
  var itemsResult = [];
  var pagesResult = [];
  var currentSection = "Maklumat Am";
  var currentSectionId = "PAGE_0";
  var currentPageIndex = 0;

  pagesResult.push({
    pageIndex: 0,
    sectionId: "PAGE_0",
    title: "Maklumat Am",
    description: "",
    navDestination: null,
    isSubmitPage: false
  });

  try {
    var items = formData && formData[1] && formData[1][1];
    if (!Array.isArray(items)) {
      return { items: [], pages: pagesResult, lastPageIndex: 0, totalPages: 1 };
    }

    items.forEach(function (item) {
      if (!item) return;

      // item[3] === 8 adalah 'Section Header' / 'Page Break'
      // item[3] === 6 adalah 'Title & Description' block
      if (item[3] === 8) {
        currentPageIndex++;
        currentSectionId = String(item[0] || ("PAGE_" + currentPageIndex));
        var secTitleBreak = String(item[1] || "").replace(/\s{2,}/g, " ").trim();
        var secDescBreak = String(item[2] || "").trim();
        var navTarget = item[5] !== undefined && item[5] !== null ? String(item[5]) : null; // "-2" bermaksud SUBMIT_FORM
        if (secTitleBreak) {
          currentSection = secTitleBreak;
        }
        pagesResult.push({
          pageIndex: currentPageIndex,
          sectionId: currentSectionId,
          title: secTitleBreak || ("Halaman " + (currentPageIndex + 1)),
          description: secDescBreak,
          navDestination: navTarget,
          isSubmitPage: navTarget === "-2" || navTarget === -2
        });
      } else if (item[3] === 6) {
        var secTitleInfo = String(item[1] || "").replace(/\s{2,}/g, " ").trim();
        if (secTitleInfo) {
          currentSection = secTitleInfo;
        }
      }

      // Soalan sebenar berada dalam item[4]
      if (Array.isArray(item[4]) && item[4].length > 0) {
        var labelItem = String(item[1] || "").replace(/\s{2,}/g, " ").trim();
        var arrayItem = item[4][0];
        if (!arrayItem) return;

        // Label berserta konteks seksyen jika soalan ialah 'KOD DAN NAMA SEKOLAH'
        var fullLabel = labelItem;
        var upperLabel = labelItem.toUpperCase();
        if (upperLabel.includes("KOD") && upperLabel.includes("SEKOLAH")) {
          if (currentSection) {
            fullLabel = labelItem.replace(/:\s*$/, "") + ": " + currentSection;
          }
        }

        var isRequired = item[4][2] === 1;

        if (Array.isArray(arrayItem) && arrayItem.length === 8) {
          // Jenis Tarikh (Date Item)
          var idTarikh = String(arrayItem[0]);
          itemsResult.push({
            id: "entry." + idTarikh,
            rawId: idTarikh,
            label: fullLabel,
            baseLabel: labelItem,
            section: currentSection,
            sectionId: currentSectionId,
            pageIndex: currentPageIndex,
            type: "Tarikh",
            isDate: true,
            required: isRequired,
            options: [],
            choiceTargets: {},
            hasBranching: false
          });
        } else {
          // Jenis Teks, Dropdown, Radio, Checkbox
          var idItem = String(arrayItem[0]);
          var senaraiPilihan = [];
          var choiceTargets = {};
          var hasOtherOption = false;

          if (Array.isArray(arrayItem[1])) {
            arrayItem[1].forEach(function (p) {
              try {
                var rawOpt = (p[0] !== undefined && p[0] !== null) ? String(p[0]).trim() : "";
                var isOther = (p[4] === 1) || (rawOpt === "__other_option__");
                if (isOther) hasOtherOption = true;

                // Tentukan label pilihan yang mesra pengguna dan selamat daripada dunder '__'
                var optLabel = "";
                if (isOther) {
                  optLabel = (rawOpt && rawOpt !== "__other_option__") ? rawOpt : "Lain-lain";
                } else {
                  optLabel = rawOpt;
                }

                if (optLabel === "__other_option__") {
                  optLabel = "Lain-lain";
                }

                if (optLabel) {
                  senaraiPilihan.push(optLabel);
                  if (p[2] !== undefined && p[2] !== null) {
                    var safeTargetKey = sanitizeGasPropertyKey_(optLabel);
                    choiceTargets[safeTargetKey] = String(p[2]);
                  }
                }
              } catch (e) {
                var fallbackOpt = p[0] ? String(p[0]).trim() : "";
                if (fallbackOpt === "__other_option__") fallbackOpt = "Lain-lain";
                if (fallbackOpt) {
                  senaraiPilihan.push(fallbackOpt);
                  if (p[2] !== undefined && p[2] !== null) {
                    var safeFallbackKey = sanitizeGasPropertyKey_(fallbackOpt);
                    choiceTargets[safeFallbackKey] = String(p[2]);
                  }
                }
              }
            });
          }

          var jenisSoalan = "Teks";
          if (senaraiPilihan.length > 0) {
            if (item[3] === 3) jenisSoalan = "Dropdown";
            else if (item[3] === 4) jenisSoalan = "Kotak Semak";
            else jenisSoalan = "Pilihan Tunggal";
          }

          var hasBranching = Object.keys(choiceTargets).length > 0;

          itemsResult.push({
            id: "entry." + idItem,
            rawId: idItem,
            label: fullLabel,
            baseLabel: labelItem,
            section: currentSection,
            sectionId: currentSectionId,
            pageIndex: currentPageIndex,
            type: jenisSoalan,
            isDate: false,
            required: isRequired,
            options: senaraiPilihan,
            choiceTargets: choiceTargets,
            hasBranching: hasBranching,
            hasOther: hasOtherOption
          });
        }
      }
    });
  } catch (err) {
    console.error("processFBData error", err);
  }

  // Tandakan halaman terakhir sebagai halaman submit lalai jika belum ditandakan
  if (pagesResult[currentPageIndex]) {
    pagesResult[currentPageIndex].isSubmitPage = true;
  }

  return {
    items: itemsResult,
    pages: pagesResult,
    lastPageIndex: currentPageIndex,
    totalPages: currentPageIndex + 1
  };
}

/**
 * API Utama: Mengambil dan memproses elemen soalan daripada Google Form URL.
 * @param {string} url
 * @return {Object}
 */
function dapatkanElemenBorangGF(url) {
  try {
    if (!url) throw new Error("Sila masukkan URL Google Form.");
    var cleanUrl = url.trim();
    if (!isValidFormsUrl(cleanUrl)) {
      throw new Error("URL tidak sah. Sila pastikan pautan Google Forms rasmi atau pautan forms.gle.");
    }

    var realUrl = resolveShortFormsUrl(cleanUrl);
    realUrl = normalizeToViewformUrl_(realUrl);

    var rawHtml = fetchUrlContent(realUrl);
    var fbData = extractFBPublicLoadData(rawHtml);
    if (!fbData) {
      throw new Error("Tidak dapat mengekstrak data soalan (FB_PUBLIC_LOAD_DATA_). Pastikan borang dibuka kepada umum atau penerima.");
    }

    var formTitle = (fbData[1] && fbData[1][8]) || (fbData[8] && fbData[8][0]) || "Borang Google Forms ISPPK";
    var processed = processFBData(fbData);
    if (!processed.items || processed.items.length === 0) {
      throw new Error("Tiada elemen soalan ditemui dalam borang ini.");
    }

    var resultObj = {
      success: true,
      url: realUrl,
      title: formTitle,
      items: processed.items,
      pages: processed.pages,
      lastPageIndex: processed.lastPageIndex,
      totalPages: processed.totalPages
    };

    return sanitizeForGasRpc_(resultObj);
  } catch (err) {
    return sanitizeForGasRpc_({
      success: false,
      error: err.message
    });
  }
}

/**
 * Menyimpan konfigurasi pemetaan serta struktur borang ke ScriptProperties.
 * @param {string} urlBorang
 * @param {Object} pemetaanData
 * @param {Array} itemsData
 * @param {Array} pagesData
 * @return {Object}
 */
function simpanPemetaanGF(urlBorang, pemetaanData, itemsData, pagesData) {
  try {
    if (!urlBorang) throw new Error("URL borang diperlukan.");
    if (!pemetaanData || typeof pemetaanData !== "object") {
      throw new Error("Data pemetaan tidak sah.");
    }

    var cleanUrl = urlBorang.trim();
    var safeItems = itemsData && Array.isArray(itemsData) ? sanitizeForGasRpc_(itemsData) : itemsData;
    var safeMapping = pemetaanData ? sanitizeForGasRpc_(pemetaanData) : pemetaanData;
    var safePages = pagesData && Array.isArray(pagesData) ? sanitizeForGasRpc_(pagesData) : pagesData;

    var props = PropertiesService.getScriptProperties();
    props.setProperty("ISPPK_GF_TARGET_URL", cleanUrl);
    props.setProperty("ISPPK_GF_MAPPING", JSON.stringify(safeMapping));
    props.setProperty("ISPPK_GF_MAPPING_BUNDLE", JSON.stringify({
      url: cleanUrl,
      mapping: safeMapping,
      savedAt: new Date().toISOString()
    }));

    if (safeItems) {
      try {
        props.setProperty("ISPPK_GF_ITEMS", JSON.stringify(safeItems));
      } catch (itemQuotaErr) {
        // Abaikan jika senarai pilihan sekolah terlalu besar melebihi had 9KB satu kunci ScriptProperties
      }
    }
    if (safePages) {
      try {
        props.setProperty("ISPPK_GF_PAGES", JSON.stringify(safePages));
      } catch (pageQuotaErr) {}
    }

    return {
      success: true,
      urlBorang: cleanUrl,
      message: "Pemetaan elemen soalan bersama URL Google Form berjaya disimpan."
    };
  } catch (err) {
    return {
      success: false,
      error: err.message
    };
  }
}

/**
 * Mengambil konfigurasi pemetaan yang disimpan bersama URL Google Form.
 * @return {Object}
 */
function dapatkanPemetaanGF() {
  try {
    var props = PropertiesService.getScriptProperties();
    var url = (props.getProperty("ISPPK_GF_TARGET_URL") || "").trim();
    var mappingRaw = props.getProperty("ISPPK_GF_MAPPING") || "{}";
    var bundleRaw = props.getProperty("ISPPK_GF_MAPPING_BUNDLE") || "";
    var itemsRaw = props.getProperty("ISPPK_GF_ITEMS") || "[]";
    var pagesRaw = props.getProperty("ISPPK_GF_PAGES") || "[]";

    var mapping = {};
    var items = [];
    var pages = [];
    try { mapping = JSON.parse(mappingRaw); } catch (e) {}
    if (bundleRaw) {
      try {
        var bundle = JSON.parse(bundleRaw);
        if (bundle && bundle.url) {
          if (!url || bundle.url.trim() === url) {
            url = bundle.url.trim();
            mapping = bundle.mapping || {};
          } else {
            mapping = {};
          }
        }
      } catch (e) {}
    }
    try { items = JSON.parse(itemsRaw); } catch (e) {}
    try { pages = JSON.parse(pagesRaw); } catch (e) {}

    return sanitizeForGasRpc_({
      success: true,
      urlBorang: url,
      pemetaan: mapping,
      items: items,
      pages: pages
    });
  } catch (err) {
    return sanitizeForGasRpc_({
      success: false,
      error: err.message
    });
  }
}

/**
 * Menyimpan rekod status penghantaran dan tarikh/masa ke ScriptProperties mengikut tahun.
 * Menggunakan kunci berasaskan tahun bagi mengelakkan had saiz simpanan Properties (9KB/kunci).
 * @param {Object} submittedData
 * @param {string|number} [tahun]
 * @return {Object}
 */
function simpanStatusHantaranGF(submittedData, tahun) {
  try {
    var safeData = (submittedData && typeof submittedData === "object") ? sanitizeForGasRpc_(submittedData) : {};
    var thn = (tahun !== undefined && tahun !== null && String(tahun).trim() !== "")
      ? String(tahun).trim()
      : dapatkanTahunSemasa_();
    var propKey = "ISPPK_GF_SUBMITTED_" + thn;
    var props = PropertiesService.getScriptProperties();
    props.setProperty(propKey, JSON.stringify(safeData));
    return { success: true, key: propKey, tahun: thn };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Mendapatkan rekod status penghantaran daripada ScriptProperties mengikut tahun.
 * Menyokong fallback kepada kunci legasi sekiranya kunci bertahun belum wujud.
 * @param {string|number} [tahun]
 * @return {Object}
 */
function getStatusHantaranGF(tahun) {
  try {
    var thn = (tahun !== undefined && tahun !== null && String(tahun).trim() !== "")
      ? String(tahun).trim()
      : dapatkanTahunSemasa_();
    var propKey = "ISPPK_GF_SUBMITTED_" + thn;
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(propKey) || props.getProperty("ISPPK_GF_SUBMITTED") || "{}";
    var records = {};
    try { records = JSON.parse(raw); } catch (e) {}
    return { success: true, records: records, tahun: thn };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Membaca rekod guru daripada tab 'PENGISIAN' yang LENGKAP SEMUA SKOR MUKTAMAD (M).
 * Hanya rekod yang telah dinilai skor M Perancangan, Pelaksanaan, Refleksi, Murid,
 * serta mempunyai Tahap KBAT M akan dikembalikan.
 * @return {Object}
 */
function getRekodGuruSkorMuktamad() {
  try {
    var ss = getSpreadsheet_();
    var sheet = getSheetPengisian_(ss);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, senaraiGuru: [] };
    }

    var lastRow = sheet.getLastRow();
    var numCols = sheet.getLastColumn();
    var range = sheet.getRange(2, 1, lastRow - 1, numCols);
    var data = range.getValues();
    var displayData = range.getDisplayValues();
    var pubData = getPublicFormData();
    var maklumatSekolah = pubData.sekolah || {};

    var senaraiGuru = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var idRekod = (row[1] || "").toString().trim();
      var namaGuru = (row[3] || "").toString().trim();
      if (!idRekod || !namaGuru) continue;

      // Ambil metadata payload JSON terlebih dahulu jika wujud
      var payloadObj = null;
      var jsonCol = row[33];
      if (jsonCol && typeof jsonCol === "string" && jsonCol.startsWith("{")) {
        try { payloadObj = JSON.parse(jsonCol); } catch (e) {}
      }

      var skorM_Perancangan = Number(row[17]) || 0;
      var skorM_Pelaksanaan = Number(row[19]) || 0;
      var skorM_Refleksi = Number(row[21]) || 0;
      var skorM_Murid = Number(row[23]) || 0;
      var jumlahSkorM = Number(row[27]) || 0;
      var peratusM = Number(row[28]) || 0;
      var tahapM = (row[29] || "").toString().trim();
      if (!tahapM && peratusM > 0) {
        tahapM = getTahapRingkasan_(peratusM, jumlahSkorM);
      }

      // Semak kelayakan skor muktamad lengkap
      var isLengkapM =
        skorM_Perancangan > 0 &&
        skorM_Pelaksanaan > 0 &&
        skorM_Refleksi > 0 &&
        skorM_Murid > 0 &&
        tahapM !== "";

      if (!isLengkapM) continue;

      // Ambil tarikh & masa
      var tarikhRaw = row[9];
      var tarikhStr = "";
      var tahunStr = "";
      var bulanStr = "";
      var hariStr = "";
      var namaBulanStr = "";

      // Keutamaan 1: Tarikh daripada payload JSON asal
      if (payloadObj && payloadObj.pencerap && payloadObj.pencerap.tarikh) {
        tarikhStr = String(payloadObj.pencerap.tarikh).trim();
      }

      if (!tarikhStr) {
        if (tarikhRaw instanceof Date) {
          tarikhStr = Utilities.formatDate(tarikhRaw, "GMT+8", "yyyy-MM-dd");
        } else if (typeof tarikhRaw === "string" && tarikhRaw.trim() !== "") {
          tarikhStr = tarikhRaw.trim();
        }
      }

      if (tarikhStr) {
        var parts = tarikhStr.split(/[-/]/);
        if (parts.length === 3) {
          if (parts[0].length === 4) {
            tahunStr = parts[0];
            bulanStr = parts[1];
            hariStr = parts[2];
          } else {
            hariStr = parts[0];
            bulanStr = parts[1];
            tahunStr = parts[2];
          }
        }
      }

      if (!namaBulanStr && bulanStr) {
        var mIdx = parseInt(bulanStr, 10) - 1;
        var senaraiNamaBulan = [
          "Januari", "Februari", "Mac", "April", "Mei", "Jun",
          "Julai", "Ogos", "September", "Oktober", "November", "Disember"
        ];
        if (mIdx >= 0 && mIdx < 12) {
          namaBulanStr = senaraiNamaBulan[mIdx];
        }
      }

      // Tapis: Penghantaran ke Google Forms HANYA bagi rekod pencerapan dalam tahun semasa
      var tahunSemasa = dapatkanTahunSemasa_();
      if (tahunStr && tahunStr !== tahunSemasa) continue;

      // Ambil masa pencerapan secara tepat & selamat:
      // Keutamaan 1: Daripada payload JSON (sentiasa format asal bersih yang disimpan cth: '10:10')
      // Keutamaan 2: Daripada getDisplayValues() cell (teks paparan sheet cth: '10:10')
      // Keutamaan 3: Daripada masaRaw (mengekstrak jam dan minit secara terus bagi mengelakkan anjakan zon masa 1899)
      var masaStr = "";
      if (payloadObj && payloadObj.pencerap && payloadObj.pencerap.masa) {
        masaStr = String(payloadObj.pencerap.masa).trim();
      }
      if (!masaStr && displayData && displayData[i] && displayData[i][10]) {
        masaStr = String(displayData[i][10]).trim();
      }
      if (!masaStr) {
        var masaRaw = row[10];
        if (masaRaw instanceof Date) {
          var h = masaRaw.getHours();
          var m = masaRaw.getMinutes();
          masaStr = (h < 10 ? "0" + h : "" + h) + ":" + (m < 10 ? "0" + m : "" + m);
        } else {
          masaStr = (masaRaw || "").toString().trim();
        }
      }
      // Bersihkan jika ada saat (cth: '10:10:00' -> '10:10')
      if (/^\d{1,2}:\d{2}:\d{2}$/.test(masaStr)) {
        masaStr = masaStr.substring(0, 5);
      }
      // Pastikan format standard 2 digit jam (cth: '9:30' -> '09:30')
      if (/^\d:\d{2}$/.test(masaStr)) {
        masaStr = "0" + masaStr;
      }

      var rekod = {
        idRekod: idRekod,
        timestamp: (row[0] || "").toString(),
        mod: (row[2] || "").toString(),
        sekolah: {
          kodSekolah: (payloadObj && payloadObj.sekolah && payloadObj.sekolah.kodSekolah) || maklumatSekolah.kodSekolah || "",
          namaSekolah: (payloadObj && payloadObj.sekolah && payloadObj.sekolah.namaSekolah) || maklumatSekolah.namaSekolah || "",
          negeri: (payloadObj && payloadObj.sekolah && payloadObj.sekolah.negeri) || maklumatSekolah.negeri || "",
          ppd: (payloadObj && payloadObj.sekolah && payloadObj.sekolah.ppd) || maklumatSekolah.ppd || "",
          emelSekolah: (payloadObj && payloadObj.sekolah && payloadObj.sekolah.emelSekolah) || maklumatSekolah.emelSekolah || ""
        },
        guru: {
          nama: namaGuru,
          emel: (row[4] || "").toString().trim(),
          jantina: (row[5] || "").toString().trim(),
          opsyen: (row[6] || "").toString().trim()
        },
        pencerap: {
          nama: (row[7] || "").toString().trim(),
          jawatan: (row[8] || "").toString().trim(),
          email: (payloadObj && payloadObj.pencerap && (payloadObj.pencerap.email || payloadObj.pencerap.emel)) || "",
          tarikh: tarikhStr,
          tarikh_tahun: tahunStr,
          tarikh_bulan: bulanStr,
          tarikh_hari: hariStr,
          tarikh_nama_bulan: namaBulanStr,
          masa: masaStr
        },
        pdp: {
          mataPelajaran: (row[11] || "").toString().trim(),
          tajuk: (row[12] || "").toString().trim(),
          bilMurid: (row[13] || "").toString().trim(),
          tingkatan: (row[14] || "").toString().trim(),
          namaKelas: (row[15] || "").toString().trim()
        },
        skorM_Perancangan: skorM_Perancangan,
        skorM_Pelaksanaan: skorM_Pelaksanaan,
        skorM_Refleksi: skorM_Refleksi,
        skorM_Murid: skorM_Murid,
        jumlahSkorM: jumlahSkorM,
        peratusM: peratusM,
        tahapM: tahapM,
        tahapM_ringkas: getTahapRingkasan_(peratusM, jumlahSkorM),
        refleksiGuru1: (row[30] || "").toString().trim(),
        refleksiGuru2: (row[31] || "").toString().trim(),
        rumusanPencerap: (row[32] || "").toString().trim()
      };

      senaraiGuru.push(rekod);
    }

    return {
      success: true,
      senaraiGuru: senaraiGuru
    };
  } catch (err) {
    return {
      success: false,
      error: err.message
    };
  }
}

/**
 * Menormalkan URL Google Forms kepada format /viewform
 * Menyokong URL daripada profil pelbagai akaun Google (/u/0/) dan pelbagai format pautan borang
 * @param {string} url
 * @return {string}
 */
function normalizeToViewformUrl_(url) {
  if (!url) return "";
  var clean = url.trim();
  var m = clean.match(/^(https:\/\/docs\.google\.com\/forms\/(?:u\/\d+\/)?d\/(?:e\/[^\/?#]+|[^\/?#]+))/i);
  if (m && m[1]) {
    return m[1] + "/viewform";
  }
  return clean.replace(/\/(?:formResponse|edit).*$/, "/viewform");
}
