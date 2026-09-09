/**
 * MARKOV ENGINE - SÜRÜM v5.3 (Hassas Hizalama & Indeks Sabitleme)
 */

function drvEngine() {
  var ts = 4;
  return runFullMarkovEngineAscending10Days(ts);
}

function runFullMarkovEngineAscending10Days(backwardSteps = 22) {
  var VERSION = "MARKOV ENGINE - SÜRÜM v5.3";
  Trns.trs.getRange('A116').setValue(VERSION);

  var START_COL = 7; // G Sütunu
  var emptyRow = nextcell();
  if (!emptyRow) return null;

  // 1. HAM VERİYİ OKUMA VE HİZALAMA
  var lastRowInSheet = Trns.trs.getLastRow();
  if (lastRowInSheet < 2) return null;

  var rawValues = Trns.trs.getRange(2, 1, lastRowInSheet - 1, 6).getValues();
  var chronologicalData = [];

  for (var i = 0; i < rawValues.length; i++) {
    var r = rawValues[i];
    var actualRow = 2 + i;

    // Sadece geçerli sayısal kapanış fiyatı olan satırları al
    if (r[0] && r[4] !== "" && !isNaN(Number(r[4]))) {
      chronologicalData.push({
        sheetRow: actualRow, // Satır numarasını bellekte çiviliyoruz
        date: r[0],
        open: Number(r[1]),
        high: Number(r[2]),
        low: Number(r[3]),
        close: Number(r[4]),
        volume: isNaN(Number(r[5])) ? 0 : Number(r[5])
      });
    }
  }

  // Eskiden yeniye sıralama (Düz Kronoloji)
  chronologicalData.sort(function(a, b) {
    return new Date(a.date) - new Date(b.date);
  });

  var totalBars = chronologicalData.length;
  if (totalBars === 0) return null;

  // 2. MACD VE DURUM HESAPLAMA
  var macdResults = (totalBars >= 17) ? calculateMACD(chronologicalData, 12, 26, 9) : null;
  var allStates = chronologicalData.map(function(bar, idx) {
    return (totalBars >= 17 && macdResults) 
      ? calculate4StatesFromOHLCV(bar, macdResults[idx])
      : ((bar.close >= bar.open) ? 1 : 0);
  });

  // 3. BACKTEST (TAM SATIR HİZALAMASI İLE)
  var lastIdx = totalBars - 1;

  for (var step = 0; step < backwardSteps; step++) {
    var chronIdx = lastIdx - step;
    if (chronIdx < 0) break;

    var baseBar = chronologicalData[chronIdx];
    
    // YAZILACAK HEDEF SATIR: Birebir kaynağın bir üst satırı
    var targetRow = baseBar.sheetRow - 1; 

    var initStateVec = getInitialStateVector(allStates[chronIdx]);
    var mat = build4x4TransitionMatrix(allStates, chronIdx, Math.min(chronIdx, 22));
    var vec = multiplyVectorMatrix(initStateVec, mat);
    var bullWeight = vec[1] + vec[3]; 
    var changePct = (bullWeight - 0.5) * 0.04;

    var modelOpen = Number(baseBar.close.toFixed(2));
    var modelClose = Number((baseBar.close * (1 + changePct)).toFixed(2));
    var modelHigh = Number((Math.max(modelOpen, modelClose) * 1.003).toFixed(2));
    var modelLow = Number((Math.min(modelOpen, modelClose) * 0.997).toFixed(2));

    var deviationVal = 0;
    var fDate = formatDateCustom(baseBar.date);

    // Gerçekleşen sonraki bar var ise tarih ve sapma eşleştirmesi
    if (chronIdx + 1 < totalBars) {
      var nextActualBar = chronologicalData[chronIdx + 1];
      fDate = formatDateCustom(nextActualBar.date);
      deviationVal = Number((((nextActualBar.close - modelClose) / nextActualBar.close)).toFixed(3));
    }

    // Doğrudan kaynak satırına paralel yazım
    Trns.trs.getRange(targetRow, START_COL).setValue(fDate);
    Trns.trs.getRange(targetRow, START_COL + 1, 1, 4).setValues([[modelOpen, modelHigh, modelLow, modelClose]]);
    Trns.trs.getRange(targetRow, START_COL + 5).setValue(deviationVal);
  }

  // 4. FORWARD PREDICTION (EN SON GEÇERLİ BARDAN BAŞLATMA)
  var lastValidBar = chronologicalData[lastIdx];
  var currentBaseClose = lastValidBar.close;
  var currentState = allStates[lastIdx];
  var baseDate = new Date(lastValidBar.date);

  // Tahminlerin yazılacağı başlangıç satırı: En son geçerli verinin tam üzeri
  var currentTargetRow = lastValidBar.sheetRow - 1;

  for (var f = 1; f <= 10; f++) {
    var initStateVecF = getInitialStateVector(currentState);
    var matF = build4x4TransitionMatrix(allStates, lastIdx, Math.min(lastIdx, 22));
    var vecF = multiplyVectorMatrix(initStateVecF, matF);
    var bullWeightF = vecF[1] + vecF[3]; 
    var changePctF = (bullWeightF - 0.5) * 0.04;

    var fOpen = Number(currentBaseClose.toFixed(2));
    var fClose = Number((currentBaseClose * (1 + changePctF)).toFixed(2));
    var fHigh = Number((Math.max(fOpen, fClose) * 1.003).toFixed(2));
    var fLow = Number((Math.min(fOpen, fClose) * 0.997).toFixed(2));
    var fDev = Number(changePctF.toFixed(3));

    baseDate.setDate(baseDate.getDate() + 1);
    if (baseDate.getDay() === 6) baseDate.setDate(baseDate.getDate() + 2);
    if (baseDate.getDay() === 0) baseDate.setDate(baseDate.getDate() + 1);
    var fDateStr = formatDateCustom(baseDate);

    Trns.trs.getRange(currentTargetRow, START_COL).setValue(fDateStr);
    Trns.trs.getRange(currentTargetRow, START_COL + 1, 1, 5).setValues([[fOpen, fHigh, fLow, fClose, fDev]]);

    currentBaseClose = fClose;
    currentState = (fClose >= fOpen) ? 1 : 0; 
    currentTargetRow--; 
  }

  return { emptyRow: emptyRow, totalBarsProcessed: totalBars };
}


function calculateMACD(data, fastPeriod = 12, slowPeriod = 26, signalPeriod = 9) {
  // GÜVENLİK KONTROLÜ: Veri yoksa veya boşsa çökme, null döndür
  if (!data || !Array.isArray(data) || data.length === 0) {
    Logger.log("⚠️ calculateMACD Hatası: Gelen veri kümesi boş veya tanımlı değil!");
    return [];
  }

  // Kapanış fiyatlarını güvenli çekme
  var closes = data.map(function(bar) { 
    return (bar && bar.close !== undefined) ? bar.close : 0; 
  });

  if (closes.length < slowPeriod) {
    Logger.log("⚠️ calculateMACD Hatası: Yeterli bar sayısı yok (" + closes.length + " / " + slowPeriod + ")");
    return new Array(data.length).fill({ macd: 0, signal: 0, histogram: 0 });
  }

  // --- Var olan EMA ve MACD hesaplama mantığınız burada aynen devam eder ---
  var emaFast = calculateEMA(closes, fastPeriod);
  var emaSlow = calculateEMA(closes, slowPeriod);

  var macdLine = [];
  for (var i = 0; i < closes.length; i++) {
    macdLine.push(emaFast[i] - emaSlow[i]);
  }

  var signalLine = calculateEMA(macdLine, signalPeriod);
  var macdResults = [];

  for (var j = 0; j < closes.length; j++) {
    var hist = macdLine[j] - signalLine[j];
    macdResults.push({
      macd: macdLine[j],
      signal: signalLine[j],
      histogram: hist
    });
  }

  return macdResults;
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