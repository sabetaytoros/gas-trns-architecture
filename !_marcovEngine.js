/**
 * MARKOV ENGINE - SÜRÜM v5.9.8.2
 * - Dinamik Sürüm Yazımı (nextcell() - 11)
 * - Otomatik Forward (G-L) ve Backtest (N-S) Başlıkları
 */

function drvEngine() {
  var ts = 22; // Tam 22 Adet Backtest Adımı
  return runMarkovEngine(ts);
}

function parseDateCustom(dateStr) {
  if (!dateStr) return new Date(0);
  if (dateStr instanceof Date) return dateStr;
  var parts = String(dateStr).trim().split('.');
  if (parts.length === 3) {
    var day = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10) - 1;
    var year = parseInt(parts[2], 10);
    if (year < 100) year += 2000;
    return new Date(year, month, day);
  }
  return new Date(dateStr);
}

function runMarkovEngine(backwardSteps = 22) {
  var VERSION = "MARKOV ENGINE - SÜRÜM v5.9.8.2";

  var G_START_COL = 7;  // G Sütunu (Forward 11 Adım Tahmin)
  var N_START_COL = 14; // N Sütunu (Backtest 22 Adım Tahmin)
  
  var emptyRow = nextcell(); // Örn: Satır 127
  if (!emptyRow) return null;

  // 0. DİNAMİK SÜRÜM BİLGİSİ YAZIMI (nextcell() - 11)
  var versionRow = emptyRow - 10; // Örn: 127 - 11 = Satır 116
  Trns.trs.getRange(versionRow, 1).setValue(VERSION);

  // 1. TEMİZLEME İŞLEMLERİ (Eski Verileri Temizler)
  Trns.trs.getRange(emptyRow - 9, G_START_COL, 10, 6).clearContent(); // Forward Tahminler (G-L)
  Trns.trs.getRange(emptyRow + 1, N_START_COL, backwardSteps, 6).clearContent(); // Backtest (N-S)

  // 2. DİNAMİK BAŞLIKLAR
  var headers = [['Date', 'Open', 'High', 'Low', 'Close', 'Dev']];
  
  // Forward Başlıkları (G - L Sütunları, emptyRow - 11) -> Örn: Satır 116
  Trns.trs.getRange(versionRow, G_START_COL, 1, 6).setValues(headers);

  // Backtest Başlıkları (N - S Sütunları, emptyRow) -> Örn: Satır 127
  Trns.trs.getRange(emptyRow, N_START_COL, 1, 6).setValues(headers);


  // 3. HAM VERİ OKUMA (22 ADIM İÇİN NEXCELL()+1 -> NEXCELL()+23)
  var startReadRow = emptyRow + 1; // Satır 128
  var totalReadRows = backwardSteps + 1; // 23 Satır
  
  var rawValues = Trns.trs.getRange(startReadRow, 1, totalReadRows, 6).getValues();
  
  var chronologicalData = [];
  var rowMap = {}; 

  for (var i = 0; i < rawValues.length; i++) {
    var r = rawValues[i];
    var actualRow = startReadRow + i;
    var rawClose = Number(r[4]);

    if (r[0] && r[4] !== "" && !isNaN(rawClose) && rawClose > 0) {
      var rawVol = Number(r[5]);
      var cleanVolume = (isNaN(rawVol) || r[5] === "#N/A") ? 0 : rawVol;

      var barObj = {
        sheetRow: actualRow,
        parsedDate: parseDateCustom(r[0]),
        rawDateStr: r[0],
        open: Number(r[1]) || rawClose,
        high: Number(r[2]) || rawClose,
        low: Number(r[3]) || rawClose,
        close: rawClose,
        volume: cleanVolume
      };

      chronologicalData.push(barObj);
      rowMap[actualRow] = barObj;
    }
  }

  // Kronolojik Sıralama (Eskiden Yeniye)
  chronologicalData.sort(function(a, b) {
    return a.parsedDate - b.parsedDate;
  });

  var totalBars = chronologicalData.length;
  if (totalBars === 0) return null;

  // 4. MACD VE DURUM HESAPLAMA
  var macdResults = (totalBars >= 17) ? calculateMACD(chronologicalData, 12, 26, 9) : null;
  var allStates = chronologicalData.map(function(bar, idx) {
    return (macdResults && macdResults[idx]) 
      ? calculate4StatesFromOHLCV(bar, macdResults[idx])
      : ((bar.close >= bar.open) ? 1 : 0);
  });

  // 5. BACKTEST - 22 ADIM (N-S SÜTUNLARINA YAZILIR)
  for (var step = 0; step < backwardSteps; step++) {
    var targetRow = (emptyRow + 1) + step; // 128, 129... 149
    var baseBar = rowMap[targetRow];
    var prevBar = rowMap[targetRow + 1];

    if (!baseBar || !prevBar) continue;

    var chronIdx = chronologicalData.indexOf(baseBar);
    if (chronIdx <= 0) continue;

    var initStateVec = getInitialStateVector(allStates[chronIdx - 1]);
    var mat = build4x4TransitionMatrix(allStates, chronIdx - 1, Math.min(chronIdx - 1, 22));
    var vec = multiplyVectorMatrix(initStateVec, mat);
    
    var bullWeight = vec[1] + vec[3]; 
    var bearWeight = vec[0] + vec[2];

    var changePct = (bullWeight - 0.5) * 0.04;
    var highChangePct = Math.max(0, (bullWeight - 0.5) * 0.02);
    var lowChangePct  = Math.max(0, (bearWeight - 0.5) * 0.02);

    var prevClose = prevBar.close; 
    
    var btPred = {
      open: Number((prevClose * (1 + (changePct * 0.1))).toFixed(2)),
      close: 0,
      high: 0,
      low: 0,
      dev: "",
      dateStr: formatDateCustom(baseBar.parsedDate)
    };
    
    btPred.close = Number((btPred.open * (1 + changePct)).toFixed(2));
    
    var maxBody = Math.max(btPred.open, btPred.close);
    var minBody = Math.min(btPred.open, btPred.close);

    btPred.high = Number((maxBody * (1 + highChangePct + 0.003)).toFixed(2));
    btPred.low  = Number((minBody * (1 - lowChangePct - 0.003)).toFixed(2));

    if (baseBar.close > 0) {
      btPred.dev = Number((((baseBar.close - btPred.close) / baseBar.close)).toFixed(3));
    }

    var writeValues = [[btPred.open, btPred.high, btPred.low, btPred.close, btPred.dev]];

    // N Sütununa Tarih, O-S Sütunlarına Tahmin Yazılır
    Trns.trs.getRange(targetRow, N_START_COL).setValue(btPred.dateStr);
    Trns.trs.getRange(targetRow, N_START_COL + 1, 1, 5).setValues(writeValues);
  }

  // 6. FORWARD PREDICTION (126 -> 117 G-L SÜTUNLARINA YAZILIR)
  var anchorRow = emptyRow - 1; // Satır 126
  var baseBarFwd = chronologicalData[totalBars - 1]; 
  var currentBaseClose = Number(baseBarFwd.close);
  var currentState = allStates[totalBars - 1];
  var baseDate = new Date(baseBarFwd.parsedDate);

  var currentTargetRow = anchorRow + 1; 
   
  for (var f = 1; f <= 10; f++) {
    var initStateVecF = getInitialStateVector(currentState);
    var matF = build4x4TransitionMatrix(
      allStates, 
      totalBars - 1, 
      Math.min(totalBars - 1, 22)
    );
    var vecF = multiplyVectorMatrix(initStateVecF, matF);
    
    var bullWeightF = vecF[1] + vecF[3]; 
    var bearWeightF = vecF[0] + vecF[2];

    var changePctF = (bullWeightF - 0.5) * 0.04;
    var highChangePctF = Math.max(0, (bullWeightF - 0.5) * 0.02);
    var lowChangePctF  = Math.max(0, (bearWeightF - 0.5) * 0.02);

    var pred = {
      open: Number((currentBaseClose * (1 + (changePctF * 0.05))).toFixed(2)),
      close: 0,
      high: 0,
      low: 0,
      dev: "", 
      dateStr: ""
    };

    pred.close = Number((pred.open * (1 + changePctF)).toFixed(2));

    var maxBodyF = Math.max(pred.open, pred.close);
    var minBodyF = Math.min(pred.open, pred.close);

    pred.high = Number((maxBodyF * (1 + highChangePctF + 0.003)).toFixed(2));
    pred.low  = Number((minBodyF * (1 - lowChangePctF - 0.003)).toFixed(2));

    baseDate.setDate(baseDate.getDate() + 1);
    if (baseDate.getDay() === 6) baseDate.setDate(baseDate.getDate() + 2);
    if (baseDate.getDay() === 0) baseDate.setDate(baseDate.getDate() + 1);
    pred.dateStr = formatDateCustom(baseDate);

    var fwdWriteValues = [[pred.open, pred.high, pred.low, pred.close, pred.dev]];

    // G Sütununa Tarih, H-L Sütunlarına Forward Tahmin Yazılır
    Trns.trs.getRange(currentTargetRow, G_START_COL).setValue(pred.dateStr);
    Trns.trs.getRange(currentTargetRow, G_START_COL + 1, 1, 5).setValues(fwdWriteValues);
    currentBaseClose = pred.close;
    currentState = (pred.close >= pred.open) ? 1 : 0; 
    currentTargetRow--; 
  }

  return { emptyRow: emptyRow, totalBarsProcessed: totalBars };
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