
/**
 * ============================================================================
 * TEST PROGRAMI: "Test" Sayfasında Dünden Kalan Eksik 8. Adımın Tamamlanması
 * G11 (Prediction) ve G12 (Reality) Hücrelerinin Doğrulanması
 * ============================================================================
 */

/**
 * ============================================================================
 * SIFIRLAMALI VE DOĞAL AKIŞLI TEST SUITE
 * ============================================================================
 * Mantık: Bugünü temizler, sayfayı geçmiş duruma çeker ve EnhancedDashboard()'ı tetikler.
 */

function testRunMissingStep8OnTestSheet() {
  Logger.clear();
  Logger.log("=================================================");
  Logger.log("   DOĞAL AKIŞLI DASHBOARD TESTİ BAŞLATILIYOR     ");
  Logger.log("=================================================\n");

  var passedCount = 0;
  var totalTests = 0;

  function assert(condition, testName, detail) {
    totalTests++;
    if (condition) {
      passedCount++;
      Logger.log("✅ [PASSED] " + testName + (detail ? " -> " + detail : ""));
    } else {
      Logger.log("❌ [FAILED] " + testName + " -> " + (detail || "Beklenen değer bulunamadı!"));
    }
  }

  try {
    if (typeof Trns === 'undefined' || !Trns.trs) {
      Logger.log("❌ HATA: Trns ortamı bulunamadı.");
      return false;
    }

    // 1. ADIM: SEKME KONTROLÜ
    assert(Trns.Name === "Test", "Sekme Doğrulaması: Hedef sayfa 'Test' mi?", "Mevcut: " + Trns.Name);

    // 2. ADIM: BUGÜNÜN PANOSUNU / SATIRINI TEMİZLEME (Sanki dün akşamdayız)
    Logger.log("   • Bugüne ait geçici dashboard verileri temizleniyor...");
    var targetRow = Trns.targetRow || (Trns.emptyRow - 12);
    if (targetRow > 0) {
      var gridRange = Trns.trs.getRange(targetRow, 1, 10, 10);
      gridRange.clearContent();
      gridRange.clearFormat();
    }

    // 3. ADIM: ANA DASHBOARD MOTORUNU ÇAĞIRMA
    Logger.log("   • EnhancedDashboard() fonksiyonu tetikleniyor...");
    var dashResult = EnhancedDashboard();
    assert(dashResult === true, "Dashboard Çalıştırma: EnhancedDashboard() 'true' döndürdü mü?");

    // 4. ADIM: G11 VE G12 HÜCRE SONUÇLARINI DOĞRULAMA
    if (dashResult) {
      Logger.log("\n-------------------------------------------------");
      Logger.log("   G11 VE G12 HÜCRE YAZIM DOĞRULAMASI");
      Logger.log("-------------------------------------------------");

      var g7DateVal  = Trns.trs.getRange("G7").getValue();
      var g11Price   = Number(Trns.trs.getRange("G11").getValue());
      var g12Ratio   = Number(Trns.trs.getRange("G12").getValue());

      assert(g7DateVal !== "", "G7 Kontrolü: Tahmin Tarihi yazıldı mı?", "Tarih: " + g7DateVal);
      assert(!isNaN(g11Price) && g11Price > 0, "G11 Kontrolü: Hedef Tahmin Fiyatı (Low/High) yazıldı mı?", "G11: " + g11Price);
      assert(!isNaN(g12Ratio) && g12Ratio >= 0, "G12 Kontrolü: En Düşükten En Yükseğe Genlik Oranı yazıldı mı?", "G12 Ratio: " + (g12Ratio * 100).toFixed(2) + "%");
    }

  } catch (err) {
    assert(false, "Sıfırlamalı Test Hasılası", err.toString());
  }

  Logger.log("\n=================================================");
  Logger.log("  TEST SONUÇLARI: " + passedCount + " / " + totalTests + " BAŞARILI");
  if (passedCount === totalTests) {
    Logger.log("  SONUÇ: SAYFA BİR ÖNCEKİ GÜNE ÇEKİLİP EnhancedDashboard() TAM MÜKEMMELLİKLE TEST EDİLDİ!");
  } else {
    Logger.log("  SONUÇ: BAZI ADIMLAR BAŞARISIZ OLDU. Logları inceleyin.");
  }
  Logger.log("=================================================");
}
/**
 * ============================================================================
 * MARKOV WINNING OHL & FORWARD PREDICTION ENGINE (G11/G12 GÜNCEL MODÜL)
 * ============================================================================
 * G Sütunu: GELECEK TAHMİNLERİ VE ÖLÇÜM
 * (G7: Date, G8: Price, G9: %, G10: Days, G11: Prediction, G12: Reality)
 */

// 1. ADIM: SAYFA VE VERİ DOĞRULAMA OBJERİ
Trns.step1_Validator = {
  validate: function () {
    if (!Trns.trs || !Trns.predictions || !Trns.metricsData) {
      Logger.log("[Step 1 HATA] Temel Trns verileri veya tahmin objeleri eksik.");
      return false;
    }
    return true;
  }
};

// 2. ADIM: KAZANAN MODELİ (WINNING MODEL) TESPİT ETME OBJSİ
Trns.step2_WinningModelSelector = {
  getWinningPrediction: function () {
    var winningModelKey = Trns.metricsData.winningModel.name.toLowerCase();
    if (winningModelKey.indexOf("cpla") !== -1) return Trns.predictions.cpla;
    if (winningModelKey.indexOf("cnvx") !== -1) return Trns.predictions.cnvx;
    if (winningModelKey.indexOf("lg-rti") !== -1) return Trns.predictions.lgrti;
    return Trns.predictions.mrkv;
  }
};

// 3. ADIM: RELATİF SÜTUN İNDEKSİ HESAPLAYICI OBJE
Trns.step3_RelColCalculator = {
  getColIndex: function (dayOffset) {
    var baseCol = 11; // K Sütunu (İndeks 11)
    var offset = dayOffset || 0; // Bugün = 0 (K Sütunu), Yarın = 1 (J Sütunu)
    return Math.max(1, baseCol - offset);
  }
};

// 4. ADIM: WINNING OHL DEĞERLERİNİ RELATİF SÜTUNA (3-5 SATIRLAR) YAZICI OBJE
Trns.step4_OhlcWriter = {
  write: function (winPred, colIndex) {
    var ohlcGrid = [
      [winPred.open], // Satır 3: Open
      [winPred.high], // Satır 4: High
      [winPred.low]   // Satır 5: Low
    ];
    var targetRange = Trns.trs.getRange(3, colIndex, 3, 1);
    targetRange.setValues(ohlcGrid);
    targetRange.setNumberFormat("#,##0.00");
    return targetRange.getA1Notation();
  }
};

// 5. ADIM: I SÜTUNUNDAN SON DALGALANMA VE TARİH VERİLERİNİ OKUYUCU OBJE
Trns.step5_FluctuationReader = {
  read: function () {
    return {
      lastCycle: {
        date:  Trns.trs.getRange("I7").getValue(),
        price: Number(Trns.trs.getRange("I8").getValue()) || 0,
        pct:   Number(Trns.trs.getRange("I9").getValue()) || 0,
        days:  Number(Trns.trs.getRange("I10").getValue()) || 1
      }
    };
  }
};

// 6. ADIM: MARKOV TABANLI GELECEK TAHMİNİ (G7-G10) HESAPLAYICI OBJE
Trns.step6_MarkovCalculator = {
  calculate: function (lastCycleData) {
    var markovRes = Trns.markovProcessor.processTransitions(22);
    if (!markovRes) return null;

    var expRatios = markovRes.expLogRatios;
    var probVec = markovRes.probVector;

    // A. Yüzdesel Genlik Tahmini (%)
    var predictedPct = Math.exp(expRatios.r_oc) - 1;

    // B. Hedef Fiyat Tahmini (MRKV Modeli Üzerinden)
    var modelResults = Trns.predictionEvaluator.evaluateModels(lastCycleData.price, probVec, expRatios);
    var predictedPrice = modelResults.mrkv.close;

    // C. Tahmini Süre (Gün) Hesabı
    var bullWeight = probVec[1] + probVec[3];
    var bearWeight = probVec[0] + probVec[2];
    var momentumFactor = Math.abs(bullWeight - bearWeight) + 0.5;
    var predictedDays = Math.max(1, Math.round(lastCycleData.days * (1 / momentumFactor)));

    // D. Bugünün / Tahminin Tarihi (G7)
    var todayDate = Trns.dateFormatted || new Date();

    return {
      todayDate: todayDate,
      predictedPrice: predictedPrice,
      predictedPct: predictedPct,
      predictedDays: predictedDays
    };
  }
};

// 7. ADIM: GERÇEKLEŞME ORANI (REALITY) VE SAPMA ÖLÇÜCÜ OBJE
Trns.step7_RealityEvaluator = {
  evaluate: function (yesterdayPct, predictedPct) {
    if (!predictedPct || predictedPct === 0) return 0;
    return yesterdayPct / predictedPct;
  }
};

// 8. ADIM: ORKESTRATÖR VE SÜRÜCÜ MOTOR OBJE (SABİT İSİMLİ)
Trns.winningOhlcAndFluctuationEngine = {
  version: "v6.0.0",

  run: function (dayOffset) {
    try {
      Logger.log("=== 8 Adımlı Winning OHL & Forward Prediction Engine (G Sütunu G11/G12) Başlatılıyor ===");

      // Adım 1: Doğrulama
      if (!Trns.step1_Validator.validate()) return false;

      // Adım 2: Kazanan Model Tahminini Al
      var winPred = Trns.step2_WinningModelSelector.getWinningPrediction();

      // Adım 3: Dynamic Relatif Sütun Hesapla
      var colIndex = Trns.step3_RelColCalculator.getColIndex(dayOffset);

      // Adım 4: Win OHL Değerlerini Dinamik Sütuna (3-5 Satırlar) Yaz
      var writtenRange = Trns.step4_OhlcWriter.write(winPred, colIndex);

      // Adım 5: Son Dalgalanma Verilerini (I Sütunundan) Oku
      var fData = Trns.step5_FluctuationReader.read();

      // Adım 6: Markov Gelecek Tahminlerini Hesapla
      var mPred = Trns.step6_MarkovCalculator.calculate(fData.lastCycle);
      if (!mPred) return false;

      // Adım 7: G SÜTUNUNA GELECEK TAHMİNLERİNİ YAZMA (G7:G10)
      Trns.trs.getRange("G7").setValue(mPred.todayDate);                                  // Date
      Trns.trs.getRange("G8").setValue(mPred.predictedPrice).setNumberFormat("#,##0.00"); // Price
      Trns.trs.getRange("G9").setValue(mPred.predictedPct).setNumberFormat("0.00%");      // %
      Trns.trs.getRange("G10").setValue(mPred.predictedDays).setNumberFormat("0");        // Duration (Gün)

      // Adım 8: G11 (PREDICTION) VE G12 (REALITY) HÜCRELERİNİ GÜNCELLEME
      var yesterdayPct = Number(Trns.trs.getRange("H9").getValue()) || 0;
      var realityRatio = Trns.step7_RealityEvaluator.evaluate(yesterdayPct, mPred.predictedPct);

      Trns.trs.getRange("G11").setValue(mPred.predictedPrice).setNumberFormat("#,##0.00"); // G11 -> Prediction
      Trns.trs.getRange("G12").setValue(realityRatio).setNumberFormat("0.0000");     // G12 -> Reality

      SpreadsheetApp.flush();

      Logger.log("✅ [G SÜTUNU TAHMİN VE REALITY YAZIMI TAMAMLANDI]");
      Logger.log("   • G7  (Date)        : " + mPred.todayDate);
      Logger.log("   • G11 (Prediction)  : " + mPred.predictedPrice);
      Logger.log("   • G12 (Reality)     : " + realityRatio.toFixed(4));
      Logger.log("   • OHL Aralığı       : " + writtenRange);

      return true;

    } catch (err) {
      Logger.log("[winningOhlcAndFluctuationEngine HATA] " + err.toString());
      return false;
    }
  }
};
/**
 * MARKOV GLOBAL TEST (Nihai Sürüm)
 * @param {boolean} useDynamicData - true ise Trns.trs("Test") üzerinden, false ise Hard-Code veriden okur.
 */
function test() {
  Logger.log("==========================================");
  Logger.log(">>> 1. AŞAMA: SABİT (HARD-CODE) TEST <<<");
  Logger.log("==========================================");
  runMarkovGlobalTest(false);
  Logger.log("\n==========================================");
  Logger.log(">>> 2. AŞAMA: DİNAMİK (TRNS) TEST <<<");
  Logger.log("==========================================");
  runMarkovGlobalTest(true);
}

function runMarkovGlobalTest(useDynamicData = true) {
  Logger.log("=== MARKOV GLOBAL TEST BAŞLATILDI ===");
  Logger.log("Veri Kaynağı: " + (useDynamicData ? "Dinamik (Trns.trs(\"Test\"))" : "Sabit (Hard Data)"));

  // 1. Veriyi Çek
  var ohlcvData = getGlobalOHLCVData(useDynamicData);

  // 2. Tarihleri Kronolojik Sıraya Al (Eskiden Yeniye)
  // Eğer gelen veri azalan sırada ise (Örn: 04.09 -> 01.09), ters çevirir.
  var isDescending = new Date(ohlcvData[0].date) > new Date(ohlcvData[ohlcvData.length - 1].date);
  if (isDescending) {
    ohlcvData.reverse();
  }

  // 3. OHLCV Barlarından Durumları (States) Hesapla
  var states = ohlcvData.map(function(bar) {
    return calculateStateFromOHLCV(bar);
  });

  // 4. stateCount Hesaplama (Max Durum + 1)
  var maxState = Math.max.apply(null, states);
  var stateCount = maxState + 1;

  Logger.log("İşlenen Bar Sayısı: " + ohlcvData.length);
  Logger.log("Sıralı Tarih Aralığı: " + ohlcvData[0].date + " -> " + ohlcvData[ohlcvData.length - 1].date);
  Logger.log("Hesaplanan Durumlar: " + JSON.stringify(states));
  Logger.log("Matris Boyutu (stateCount): " + stateCount);

  // 5. Geçiş Matrisini Hesapla
  var transitionMatrix = calculateGlobalTransitionMatrix(states, stateCount);

  // 6. Matrisi Logla
  Logger.log("\n--- Markov Geçiş Matrisi ---");
  for (var i = 0; i < stateCount; i++) {
    Logger.log("Durum " + i + " -> " + JSON.stringify(transitionMatrix[i]));
  }

  // 7. Satır Toplamı Doğrulaması (Assertion)
  var isValid = verifyMatrixRowSums(transitionMatrix);
  Logger.log("Test Sonucu: " + (isValid ? "BAŞARILI" : "BAŞARISIZ"));
  Logger.log("=== MARKOV GLOBAL TEST TAMAMLANDI ===");
}

/**
 * DATA FETCHING (Veri Getirici)
 */
function getGlobalOHLCVData(isDynamic) {
  var MIN_DATA_SIZE = isDynamic ? 4 : 4;

  if (isDynamic) {
    // DINAMIK DATA: Trns.trs("Test") ve A + nextcell() üzerinden okur
    var sheet = Trns.trs("Test");
    var startRow = nextcell();
    
    // Col A: Date, B: Open, C: High, D: Low, E: Close, F: Volume
    var rawValues = sheet.getRange(startRow, 1, MIN_DATA_SIZE, 6).getValues();
    
    return rawValues.map(function(row) {
      return {
        date: row[0],
        open: Number(row[1]),
        high: Number(row[2]),
        low: Number(row[3]),
        close: Number(row[4]),
        volume: Number(row[5])
      };
    });
  } else {
    // HARD-CODE DATA: Test sayfasının ilk 4 satır verisi
    return [
      { date: "04.09.2026", open: 359.70, high: 360.16, low: 353.70, close: 357.90, volume: 32863265 },
      { date: "03.09.2026", open: 351.74, high: 359.40, low: 342.33, close: 357.16, volume: 60242596 },
      { date: "02.09.2026", open: 369.68, high: 371.09, low: 364.65, close: 367.24, volume: 38874579 },
      { date: "01.09.2026", open: 364.25, high: 371.40, low: 362.00, close: 369.68, volume: 19288938 }
    ];
  }
}

/**
 * OHLCV BARDAN STATE HESAPLAMA
 */
function calculateStateFromOHLCV(bar) {
  if (!bar.open || bar.open === 0) return 1;
  var changePercent = ((bar.close - bar.open) / bar.open) * 100;
  
  if (changePercent < -0.5) return 0;      // Düşüş
  if (changePercent > 0.5)  return 2;      // Yükseliş
  return 1;                                 // Yatay
}

/**
 * GEÇİŞ MATRİSİ HESAPLAYICI
 */
function calculateGlobalTransitionMatrix(statesArray, statesCount) {
  var counts = [];
  var matrix = [];
  
  for (var i = 0; i < statesCount; i++) {
    counts[i] = new Array(statesCount).fill(0);
    matrix[i] = new Array(statesCount).fill(0);
  }

  for (var t = 0; t < statesArray.length - 1; t++) {
    var fromState = statesArray[t];
    var toState = statesArray[t + 1];
    counts[fromState][toState]++;
  }

  for (var r = 0; r < statesCount; r++) {
    var rowSum = counts[r].reduce(function(a, b) { return a + b; }, 0);
    for (var c = 0; c < statesCount; c++) {
      matrix[r][c] = rowSum > 0 ? Number((counts[r][c] / rowSum).toFixed(4)) : (1 / statesCount);
    }
  }

  return matrix;
}

/**
 * MATRİS DOĞRULAMA (Assertion)
 */
function verifyMatrixRowSums(matrix) {
  for (var i = 0; i < matrix.length; i++) {
    var sum = matrix[i].reduce(function(a, b) { return a + b; }, 0);
    if (Math.abs(sum - 1.0) > 0.001) return false;
  }
  return true;
}