/**
 * ============================================================================
 * MARKOV CHAIN & MACD FINANCIAL FORECASTING ENGINE
 * Sürüm: v2.2.0
 * Tarih: 07.09.26
 * Açıklama: Fonksiyon çağrısı parantezleri kaldırıldı. Doğrudan Trns.trs nesnesi kullanıldı.
 * ============================================================================
 */

/**
 * @function runFullMarkovEngineAscending10Days
 * @description Ana yürütücü fonksiyon. Önce eski çıktıları temizler, geçmiş verileri okur,
 *              en iyi modeli seçer, durumları hesaplar ve çıktı alanını günceller.
 */
function runFullMarkovEngineAscending10Days() {
  var startRowT0 = nextcell();

  if (!startRowT0 || startRowT0 < 2) return;

  // --------------------------------------------------------------------------
  // 1. ÖNCEKİ ÇIKTILARI TEMİZLEME (G:Q Sütunları, Header dahil t+10 -> t0)
  // --------------------------------------------------------------------------
  var headerRow = startRowT0 - 11;
  var startCleanRow = Math.max(1, headerRow);
  var numRowsToClean = startRowT0 - startCleanRow + 1;

  if (numRowsToClean > 0) {
    // 7. Sütun (G) 'den itibaren 11 sütunluk (G:Q) çıktı alanını temizler
    Trns.trs.getRange(startCleanRow, 7, numRowsToClean, 11).clearContent();
  }

  // --------------------------------------------------------------------------
  // 2. VERİ OKUMA VE HAZIRLIK
  // --------------------------------------------------------------------------
  var TOTAL_BARS = 145;
  var startReadRow = Math.max(1, startRowT0 - TOTAL_BARS + 1);
  var numRowsToRead = startRowT0 - startReadRow + 1;

  // A:F Sütunlarından OHLCV Verilerini Oku
  var rawValues = Trns.trs.getRange(startReadRow, 1, numRowsToRead, 6).getValues();

  var fullData = rawValues.map(function(row, index) {
    return {
      rowIndex: startReadRow + index,
      date: row[0],
      open: Number(row[1]),
      high: Number(row[2]),
      low: Number(row[3]),
      close: Number(row[4]),
      volume: Number(row[5])
    };
  });

  // Tarih sıralamasını kronolojik yap (Eski -> Yeni)
  if (new Date(fullData[0].date) > new Date(fullData[fullData.length - 1].date)) {
    fullData.reverse();
  }

  // MACD ve 4-Durum Hesaplamaları
  var macdResults = calculateMACD(fullData, 12, 26, 9);
  var allStates = fullData.map(function(bar, idx) {
    return calculate4StatesFromOHLCV(bar, macdResults[idx]);
  });

  var idxT0 = fullData.length - 1;
  var idxT10 = fullData.length - 11;

  // 32 Bar (Kısa) vs 100 Bar (Uzun) Model Seçimi ve Backtest
  var backtest32 = runMarkovBacktest(allStates, idxT10, 32, 10);
  var backtest100 = runMarkovBacktest(allStates, idxT10, 100, 10);
  var actualLast10 = allStates.slice(idxT10 + 1, idxT0 + 1);

  var score32 = calculateAccuracy(backtest32.predictions, actualLast10);
  var score100 = calculateAccuracy(backtest100.predictions, actualLast10);

  var winningModel = (score32 >= score100) ? "32 Bar (Kısa)" : "100 Bar (Uzun)";
  var winningWindowSize = (score32 >= score100) ? 32 : 100;
  var winningScore = Math.max(score32, score100);

  // Geçiş Matrisi ve Başlangıç Vektörü
  var transitionMatrix = build4x4TransitionMatrix(allStates, idxT0, winningWindowSize);
  var initialVector = getInitialStateVector(allStates[idxT0]);

  // Dynamic Header Düzenlemesi (t+10 üzerindeki satır)
  if (headerRow >= 1) {
    setupInPlaceHeaders(headerRow);
  }

  // --------------------------------------------------------------------------
  // 3. ÇIKTI YAZDIRMA DÖNGÜSÜ (t0 -> t+10)
  // --------------------------------------------------------------------------
  for (var step = 0; step <= 10; step++) {
    var targetRow = startRowT0 - step;
    if (targetRow <= 1) break;

    var stepVector = multiplyVectorMatrix(initialVector, matrixPower(transitionMatrix, step));
    var predictedState = getArgMax(stepVector);

    var stepLabel = (step === 0) ? "t_0 (Bugün)" : "t+" + step;
    var forecastDate = getBusinessDayOffset(fullData[idxT0].date, step);
    var macdSignalText = macdResults[idxT0] ? macdResults[idxT0].signalText : "N/A";
    var actionText = get4StateActionAdvice(predictedState, stepVector);

    var baseStateCalc = "State " + allStates[idxT0];

    var outputValues = [[
      stepLabel,
      forecastDate,
      baseStateCalc,
      winningModel + " (%" + (winningScore * 100).toFixed(0) + ")",
      macdSignalText,
      (stepVector[0] * 100).toFixed(1) + "%",
      (stepVector[1] * 100).toFixed(1) + "%",
      (stepVector[2] * 100).toFixed(1) + "%",
      (stepVector[3] * 100).toFixed(1) + "%",
      "State " + predictedState,
      actionText
    ]];

    Trns.trs.getRange(targetRow, 7, 1, 11).setValues(outputValues);
  }
}

/**
 * @function setupInPlaceHeaders
 * @description Çıktı sütunları için başlık formatlaması.
 */
function setupInPlaceHeaders(headerRow) {
  var headers = [
    ["Etiket", "Tarih", "Hesaplanan Durum", "Model (İsabet)", "MACD Sinyal", "Prob S0", "Prob S1", "Prob S2", "Prob S3", "Tahmin Durum", "Aksiyon Tavsiyesi"]
  ];
  var range = Trns.trs.getRange(headerRow, 7, 1, 11);
  range.setValues(headers);
  range.setFontWeight("bold");
  range.setBackground("#202124");
  range.setFontColor("#ffffff");
}

/**
 * @function calculateMACD
 * @description EMA (12, 26) ve Signal (9) değerlerini hesaplar.
 */
function calculateMACD(dataArray, shortPeriod, longPeriod, signalPeriod) {
  var closes = dataArray.map(function(d) { return d.close; });
  var emaShort = calculateEMA(closes, shortPeriod);
  var emaLong = calculateEMA(closes, longPeriod);
  
  var macdLine = [];
  for (var i = 0; i < closes.length; i++) {
    macdLine.push(emaShort[i] - emaLong[i]);
  }
  
  var signalLine = calculateEMA(macdLine, signalPeriod);
  
  return macdLine.map(function(mVal, idx) {
    var sVal = signalLine[idx];
    var hist = mVal - sVal;
    var signalText = (mVal > sVal) ? "BULLISH" : "BEARISH";
    return { macd: mVal, signal: sVal, histogram: hist, signalText: signalText };
  });
}

/**
 * @function calculateEMA
 * @description Üstel Hareketli Ortalama (EMA) hesaplar.
 */
function calculateEMA(values, period) {
  var k = 2 / (period + 1);
  var emaArray = [];
  var sum = 0;
  
  for (var i = 0; i < values.length; i++) {
    if (i < period - 1) {
      sum += values[i];
      emaArray.push(values[i]);
    } else if (i === period - 1) {
      sum += values[i];
      emaArray.push(sum / period);
    } else {
      var prevEma = emaArray[i - 1];
      emaArray.push((values[i] * k) + (prevEma * (1 - k)));
    }
  }
  return emaArray;
}

/**
 * @function calculate4StatesFromOHLCV
 * @description Fiyat değişimi ve MACD bileşiminden 4 durumlu sınıflandırma yapar.
 */
function calculate4StatesFromOHLCV(bar, macdObj) {
  var isUp = bar.close >= bar.open;
  var isBullishMACD = macdObj ? (macdObj.macd >= macdObj.signal) : true;

  if (!isUp && !isBullishMACD) return 0;
  if (!isUp && isBullishMACD)  return 1;
  if (isUp && !isBullishMACD)  return 2;
  return 3;
}

/**
 * @function build4x4TransitionMatrix
 * @description Belirtilen pencere genişliğinde 4x4 Markov Geçiş Olasılık Matrisini oluşturur.
 */
function build4x4TransitionMatrix(states, endIdx, windowSize) {
  var matrix = [
    [0, 0, 0, 0],
    [0, 0, 0, 0],
    [0, 0, 0, 0],
    [0, 0, 0, 0]
  ];
  var startIdx = Math.max(0, endIdx - windowSize);

  for (var i = startIdx; i < endIdx; i++) {
    var currentState = states[i];
    var nextState = states[i + 1];
    if (currentState !== undefined && nextState !== undefined) {
      matrix[currentState][nextState]++;
    }
  }

  for (var r = 0; r < 4; r++) {
    var rowSum = matrix[r].reduce(function(a, b) { return a + b; }, 0);
    if (rowSum > 0) {
      for (var c = 0; c < 4; c++) {
        matrix[r][c] /= rowSum;
      }
    } else {
      for (var c = 0; c < 4; c++) {
        matrix[r][c] = 0.25;
      }
    }
  }
  return matrix;
}

/**
 * @function runMarkovBacktest
 * @description Test periyodunda Markov tahmin başarısını ölçer.
 */
function runMarkovBacktest(states, testStartIdx, windowSize, steps) {
  var transMatrix = build4x4TransitionMatrix(states, testStartIdx, windowSize);
  var initVec = getInitialStateVector(states[testStartIdx]);
  var predictions = [];

  for (var h = 1; h <= steps; h++) {
    var hVec = multiplyVectorMatrix(initVec, matrixPower(transMatrix, h));
    predictions.push(getArgMax(hVec));
  }
  return { predictions: predictions };
}

/**
 * @function calculateAccuracy
 * @description Tahmin edilen durumlar ile gerçekleşen durumları karşılaştırarak isabet oranını döner.
 */
function calculateAccuracy(predictions, actuals) {
  if (!predictions || !actuals || predictions.length === 0) return 0;
  var matches = 0;
  var count = Math.min(predictions.length, actuals.length);
  for (var i = 0; i < count; i++) {
    if (predictions[i] === actuals[i]) matches++;
  }
  return matches / count;
}

/**
 * @function getInitialStateVector
 * @description Geçerli durumu birim vektöre (one-hot) çevirir.
 */
function getInitialStateVector(state) {
  var vec = [0, 0, 0, 0];
  if (state >= 0 && state <= 3) {
    vec[state] = 1;
  } else {
    vec[0] = 1;
  }
  return vec;
}

/**
 * @function multiplyVectorMatrix
 * @description Vektör ile matris çarpımı.
 */
function multiplyVectorMatrix(vec, mat) {
  var result = [0, 0, 0, 0];
  for (var c = 0; c < 4; c++) {
    for (var r = 0; r < 4; r++) {
      result[c] += vec[r] * mat[r][c];
    }
  }
  return result;
}

/**
 * @function matrixPower
 * @description Matrisin n. kuvvetini alır.
 */
function matrixPower(mat, power) {
  if (power === 0) {
    return [
      [1, 0, 0, 0],
      [0, 1, 0, 0],
      [0, 0, 1, 0],
      [0, 0, 0, 1]
    ];
  }
  var result = mat;
  for (var p = 1; p < power; p++) {
    result = multiplyMatrices(result, mat);
  }
  return result;
}

/**
 * @function multiplyMatrices
 * @description İki 4x4 matrisi çarpar.
 */
function multiplyMatrices(a, b) {
  var res = [
    [0, 0, 0, 0],
    [0, 0, 0, 0],
    [0, 0, 0, 0],
    [0, 0, 0, 0]
  ];
  for (var r = 0; r < 4; r++) {
    for (var c = 0; c < 4; c++) {
      for (var k = 0; k < 4; k++) {
        res[r][c] += a[r][k] * b[k][c];
      }
    }
  }
  return res;
}

/**
 * @function getArgMax
 * @description Dizideki en yüksek değerin indeksini döner.
 */
function getArgMax(arr) {
  var maxVal = -1;
  var maxIdx = 0;
  for (var i = 0; i < arr.length; i++) {
    if (arr[i] > maxVal) {
      maxVal = arr[i];
      maxIdx = i;
    }
  }
  return maxIdx;
}

/**
 * @function getBusinessDayOffset
 * @description Hafta sonlarını atlayarak iş günü tarihini DD.MM.YY formatında hesaplar.
 */
function getBusinessDayOffset(baseDate, offsetDays) {
  var d = new Date(baseDate);
  var added = 0;
  while (added < offsetDays) {
    d.setDate(d.getDate() + 1);
    var dayOfWeek = d.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) { // 0 = Pazar, 6 = Cumartesi
      added++;
    }
  }
  return formatDateCustom(d);
}

/**
 * @function formatDateCustom
 * @description Tarihi DD.MM.YY formatında stringe çevirir.
 */
function formatDateCustom(dateObj) {
  var d = new Date(dateObj);
  var day = ("0" + d.getDate()).slice(-2);
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  var year = d.getFullYear().toString().slice(-2);
  return day + "." + month + "." + year;
}

/**
 * @function get4StateActionAdvice
 * @description Olasılık vektörüne göre stratejik işlem tavsiyesi üretir.
 */
function get4StateActionAdvice(state, probVec) {
  var bullProb = probVec[1] + probVec[3];
  if (state === 3 && bullProb > 0.60) return "GÜÇLÜ AL / POZİSYON KORU";
  if (state === 3 || state === 2) return "KADEMELİ ALIM / TUT";
  if (state === 1) return "İZLE / BOĞA DÖNÜŞ SİNYALİ";
  return "SAT / NAKİTTE KAL";
}