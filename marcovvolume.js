/**
 * ============================================================================
 * MARKOV BACKTEST ENGINE - KESİNLEŞMİŞ BARDAN BAŞLAMA (T-1)
 * ============================================================================
 * Açıklama:
 *  - Bugünü (Satır 127) eksik/tamamlanmamış kabul edip atlar.
 *  - Analizi ve projeksiyonu strictly tamamlanmış son gün olan 128. satırdan (nextcell + 2) başlatır.
 * ============================================================================
 */
function runMarkovEngineFromCompletedBar() {
  var nextEmptyRow = nextcell(); // Örn: 126. satır

  if (!nextEmptyRow) return;

  // BUGÜNÜN (EKSİK BARIN) ATLANMASI:
  // nextEmptyRow = 126 ise -> 127 (Bugün/Eksik) -> 128 (Tamamlanmış Son Gün / T0)
  var startRowT0 = nextEmptyRow + 2; // Satır 128

  var maxForecastSteps = 4; // Forward için 4 adım
  var maxBackSteps = 4;     // Backtest için 4 adım (Üst sınır = 4)

  // 1. TEMİZLİK
  var startCleanRow = Math.max(1, startRowT0 - maxForecastSteps);
  Trns.trs.getRange(startCleanRow, 7, (startRowT0 + maxBackSteps) - startCleanRow + 1, 11).clearContent();

  // 2. VERİ OKUMA
  var rawValues = Trns.trs.getRange(1, 1, startRowT0 + 10, 6).getValues();

  var chronologicalData = [];
  var sheetRowToChronIdx = {};

  for (var i = 0; i < rawValues.length; i++) {
    var r = rawValues[i];
    if (r[0] && r[4] !== "" && !isNaN(Number(r[4]))) {
      chronologicalData.push({
        sheetRow: i + 1,
        date: r[0],
        open: Number(r[1]),
        high: Number(r[2]),
        low: Number(r[3]),
        close: Number(r[4]),
        volume: Number(r[5])
      });
    }
  }

  // Kronolojik Hizalama (Eski -> Yeni)
  if (chronologicalData.length > 1 && new Date(chronologicalData[0].date) > new Date(chronologicalData[chronologicalData.length - 1].date)) {
    chronologicalData.reverse();
  }

  for (var c = 0; c < chronologicalData.length; c++) {
    sheetRowToChronIdx[chronologicalData[c].sheetRow] = c;
  }

  // MACD ve Durum Analizi
  var macdResults = calculateMACD(chronologicalData, 12, 26, 9);
  var allStates = chronologicalData.map(function(bar, idx) {
    return calculate4StatesFromOHLCV(bar, macdResults[idx]);
  });

  // --------------------------------------------------------------------------
  // 3. TAMAMLATILMIŞ BARLARDAN BACKTEST (128, 129, 130, 131)
  // --------------------------------------------------------------------------
  for (var offset = 0; offset < maxBackSteps; offset++) {
    var currentRow = startRowT0 + offset; // 128. satırdan başlar
    var chronIdx = sheetRowToChronIdx[currentRow];

    if (chronIdx === undefined || chronIdx < 1) continue;

    var actualBar = chronologicalData[chronIdx];
    var prevBar = chronologicalData[chronIdx - 1]; // Tahmin için bir önceki gün (t-1)

    var fDate = formatDateCustom(actualBar.date);
    var initStateVec = getInitialStateVector(allStates[chronIdx - 1]);

    var mat = build4x4TransitionMatrix(allStates, chronIdx - 1, 4);
    var vec = multiplyVectorMatrix(initStateVec, mat);
    var bullWeight = vec[1] + vec[3]; 
    var changePct = (bullWeight - 0.5) * 0.04;

    // FİYAT PROJEKSİYONLARI
    var modelClose = Number((prevBar.close * (1 + changePct)).toFixed(2));
    var modelOpen = Number(prevBar.close.toFixed(2));
    var modelHigh = Number((Math.max(modelOpen, modelClose) * 1.003).toFixed(2));
    var modelLow = Number((Math.min(modelOpen, modelClose) * 0.997).toFixed(2));

    var deviationVal = Number((((actualBar.close - modelClose) / actualBar.close) * 100).toFixed(2));

    // 22G Back (G:K)
    Trns.trs.getRange(currentRow, 7).setValue(fDate);
    Trns.trs.getRange(currentRow, 8, 1, 4).setValues([[modelOpen, modelHigh, modelLow, modelClose]]);
    Trns.trs.getRange(currentRow, 12).setValue(deviationVal);

    // 100G Back (M:Q)
    Trns.trs.getRange(currentRow, 13).setValue(fDate);
    Trns.trs.getRange(currentRow, 14, 1, 4).setValues([[modelOpen, modelHigh, modelLow, modelClose]]);
  }

  // --------------------------------------------------------------------------
  // 4. FORWARD PROJEKSİYON (128. Satırdan İleriye Doğru Tahmin)
  // --------------------------------------------------------------------------
  var idxT0 = sheetRowToChronIdx[startRowT0]; // Satır 128 (Tamamlanmış son bar)
  if (idxT0 !== undefined) {
    var baseBar = chronologicalData[idxT0];
    var initialVectorT0 = getInitialStateVector(allStates[idxT0]);
    var transMatrixT0 = build4x4TransitionMatrix(allStates, idxT0, 4);

    for (var step = 1; step <= maxForecastSteps; step++) {
      var targetRow = startRowT0 - step; // 127, 126, 125... satırlarına yazar
      var forecastDate = getBusinessDayOffsetForward(baseBar.date, step);

      var v = multiplyVectorMatrix(initialVectorT0, matrixPower(transMatrixT0, step));
      var bw = v[1] + v[3];
      var fFactor = 1 + ((bw - 0.5) * 0.03 * step);

      var fClose = Number((baseBar.close * fFactor).toFixed(2));
      var fOpen = Number(baseBar.close.toFixed(2));
      var fHigh = Number((Math.max(fOpen, fClose) * (1 + 0.003 * step)).toFixed(2));
      var fLow = Number((Math.min(fOpen, fClose) * (1 - 0.003 * step)).toFixed(2));

      // Forward Yazdırma
      Trns.trs.getRange(targetRow, 7).setValue(forecastDate);
      Trns.trs.getRange(targetRow, 8, 1, 4).setValues([[fOpen, fHigh, fLow, fClose]]);

      Trns.trs.getRange(targetRow, 13).setValue(forecastDate);
      Trns.trs.getRange(targetRow, 14, 1, 4).setValues([[fOpen, fHigh, fLow, fClose]]);
    }
  }
}

/**
 * MACD Hesaplama Fonksiyonu
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
    return { macd: mVal, signal: sVal, histogram: mVal - sVal };
  });
}

/**
 * EMA (Üstel Hareketli Ortalama) Hesaplama
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
 * OHLCV ve MACD Sinyalinden 4'lü Durum (State) Belirleme
 */
function calculate4StatesFromOHLCV(bar, macdObj) {
  var isUp = bar.close >= bar.open;
  var isBullishMACD = macdObj ? (macdObj.macd >= macdObj.signal) : true;

  if (!isUp && !isBullishMACD) return 0; // Ayı Mum + Ayı MACD
  if (!isUp && isBullishMACD)  return 1; // Ayı Mum + Boğa MACD
  if (isUp && !isBullishMACD)  return 2; // Boğa Mum + Ayı MACD
  return 3;                             // Boğa Mum + Boğa MACD
}

/**
 * Geçiş Matrisi Oluşturucu
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
 * Başlangıç Durum Vektörü
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
 * Vektör-Matris Çarpımı
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
 * Matris Kuvveti Alma
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
 * Matris-Matris Çarpımı
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
 * Tarih Formatlayıcı (GG.AA.YY)
 */
function formatDateCustom(dateVal) {
  if (!dateVal) return "";
  var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return "";

  var day = ("0" + d.getDate()).slice(-2);
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  var year = d.getFullYear().toString().slice(-2);
  return day + "." + month + "." + year;
}

/**
 * İş Günü İleri Tarih Atlama
 */
function getBusinessDayOffsetForward(baseDate, offsetDays) {
  var d = (baseDate instanceof Date) ? new Date(baseDate.getTime()) : new Date(baseDate);
  if (isNaN(d.getTime())) return "";

  var added = 0;
  while (added < offsetDays) {
    d.setDate(d.getDate() + 1);
    var dayOfWeek = d.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      added++;
    }
  }
  return formatDateCustom(d);
}