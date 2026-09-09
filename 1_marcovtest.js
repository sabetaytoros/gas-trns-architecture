/**
 * MARKOV TEST VE SÜRÜCÜ MODÜLÜ - SÜRÜM v4.5
 */

/**
 * SÜRÜCÜ FONKSİYON (Driver Function)
 * Tüm veri akışı ve doğrulama testlerini sırasıyla çalıştırır.
 */
function drvTestEngine() {
  Logger.clear();
  Logger.log("=================================================");
  Logger.log("      MARKOV TEST SÜRÜCÜSÜ ÇALIŞTIRILIYOR       ");
  Logger.log("=================================================\n");

  // 1. Veri Akış ve Parametre Geçiş Testi
  var pipelinePassed = testMarkovDataPipeline();

  if (!pipelinePassed) {
    Logger.log("\n❌ KRİTİK HATA: Veri akış testi başarısız oldu. Doğrulama testi iptal edildi.");
    return;
  }

  // 2. Backtest ve Forward Prediction Hizalama Testi
  Logger.log("\n-------------------------------------------------");
  runManualVsActualValidationTest(4);
}

/**
 * 1. VERİ AKIŞI VE PARAMETRE TESTİ
 * OHLCV okumasını ve calculateMACD parametre aktarımını doğrular.
 */
function testMarkovDataPipeline() {
  Logger.log(">>> [TEST 1] Veri Akışı ve MACD Parametre Kontrolü...");

  var backwardSteps = 4;
  const LOOKBACK_RATIO = 1.5;

  var emptyRow = nextcell();
  Logger.log("   • Boş Satır (emptyRow): " + emptyRow);

  if (!emptyRow) {
    Logger.log("   ❌ HATA: nextcell() geçerli satır döndürmedi!");
    return false;
  }

  var baseStartRow = emptyRow + backwardSteps;
  var totalRowsToRead = Math.ceil(backwardSteps * LOOKBACK_RATIO);
  var startReadRow = Math.max(1, baseStartRow - totalRowsToRead + 1);
  var readRowCount = (baseStartRow - startReadRow) + 1;

  var rawValues = Trns.trs.getRange(startReadRow, 1, readRowCount, 6).getValues();
  Logger.log("   • Okunan Ham Satır Sayısı: " + (rawValues ? rawValues.length : 0));

  if (!rawValues || rawValues.length === 0) {
    Logger.log("   ❌ HATA: A:F aralığından veri okunamadı!");
    return false;
  }

  var chronologicalData = [];
  for (var i = 0; i < rawValues.length; i++) {
    var r = rawValues[i];
    var actualRow = startReadRow + i;

    if (r[0] && r[4] !== "" && !isNaN(Number(r[4]))) {
      chronologicalData.push({
        sheetRow: actualRow,
        date: r[0],
        open: Number(r[1]),
        high: Number(r[2]),
        low: Number(r[3]),
        close: Number(r[4]),
        volume: isNaN(Number(r[5])) ? 0 : Number(r[5])
      });
    }
  }

  Logger.log("   • İşlenen Geçerli Bar Sayısı: " + chronologicalData.length);

  if (chronologicalData.length === 0) {
    Logger.log("   ❌ HATA: chronologicalData boş kaldı!");
    return false;
  }

  try {
    var macdResults = calculateMACD(chronologicalData, 12, 26, 9);
    Logger.log("   ✅ MACD Başarıyla Hesaplandı. Üretilen Dizi Boyutu: " + (macdResults ? macdResults.length : 0));
    return true;
  } catch (err) {
    Logger.log("   ❌ MACD HESAPLAMA HATASI: " + err.toString());
    return false;
  }
}

/**
 * 2. BACKTEST VE FORWARD HİZALAMA TESTİ
 * Gelecek tahminlerinin doğru satırdan (emptyRow - 2) başlayıp yukarı kaydığını doğrular.
 */
function runManualVsActualValidationTest(testSteps = 4) {
  Logger.log(">>> [TEST 2] Tablo Döküm ve Hizalama Doğrulaması...");

  var engineMetrics = runFullMarkovEngineAscending10Days(testSteps);

  if (!engineMetrics) {
    Logger.log("   ❌ HATA: Motor çalıştırılamadı.");
    return;
  }

  var errorCount = 0;
  var startCol = engineMetrics.startCol; // 7 (G Sütunu)
  var forwardStartRow = engineMetrics.forwardStartRow; // emptyRow - 2 (126. Satır)

  var actualDate = Trns.trs.getRange(forwardStartRow, startCol).getValue();
  var actualOpen = Trns.trs.getRange(forwardStartRow, startCol + 1).getValue();

  Logger.log("   • Forward Start Row: " + forwardStartRow + " | Sütun: " + startCol);

  if (String(actualDate).trim() === "") {
    errorCount++;
    Logger.log("   ❌ TARİH HATASI -> G" + forwardStartRow + " hücresine tarih basılmadı!");
  } else {
    Logger.log("   ✅ Tarih Doğrulandı: " + actualDate);
  }

  if (actualOpen === "" || actualOpen === 0) {
    errorCount++;
    Logger.log("   ❌ FİYAT HATASI -> H" + forwardStartRow + " hücresine açılış fiyatı yazılmadı!");
  } else {
    Logger.log("   ✅ Açılış Fiyatı Doğrulandı: " + actualOpen);
  }

  Logger.log("\n=================================================");
  if (errorCount === 0) {
    Logger.log("   SONUÇ: TÜM TESTLER BAŞARIYLA GEÇTİ (PASSED)");
  } else {
    Logger.log("   SONUÇ: TEST BAŞARISIZ (FAILED)");
  }
  Logger.log("=================================================");
}