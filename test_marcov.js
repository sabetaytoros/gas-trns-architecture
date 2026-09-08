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