/**
 * 1_MarcovConEngine
 * MARKOV CONVEX OPTIMIZATION ENGINE - SÜRÜM v6.3.0
 * - Yöntem 2: Markov Chain + Post-Processing Convex Optimization (POCS)
 * - Ham Markov tahminleri matematiksel kısıt kümesine (H >= max(O,C) ve L <= min(O,C)) projekte edilir.
 * - 22 Günlük Backtest WMAPE (%) ve MAE (%) Ölçümü Entegre Edilmiştir.
 */

function drvEngineConvexOpt() {
  var ts = 22; // 22 Adet Backtest Adımı
  return runMarkovEngineConvexOpt(ts);
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

function runMarkovEngineConvexOpt(backwardSteps = 22) {
  var VERSION = "MARKOV CONVEX OPT ENGINE - SÜRÜM v6.3.0";

  var G_START_COL = 7;  // G Sütunu (Forward 10 Adım Tahmin)
  var N_START_COL = 14; // N Sütunu (Backtest 22 Adım Tahmin)
  
  var emptyRow = nextcell(); // Örn: Satır 127
  if (!emptyRow) return null;

  // 0. DİNAMİK SÜRÜM BİLGİSİ VE BAŞLIKLAR
  var versionRow = emptyRow - 11;
  Trns.trs.getRange(versionRow, 1).setValue(VERSION);

  var headers = [['Date', 'Open', 'High', 'Low', 'Close', 'Dev']];
  
  Trns.trs.getRange(versionRow, G_START_COL, 1, 6).setValues(headers);
  Trns.trs.getRange(emptyRow - 1, N_START_COL, 1, 6).setValues(headers);

  // 1. TEMİZLEME İŞLEMLERİ
  Trns.trs.getRange(emptyRow - 10, G_START_COL, 11, 6).clearContent(); 
  Trns.trs.getRange(emptyRow, N_START_COL, backwardSteps, 6).clearContent(); 

  // 2. HAM VERİ OKUMA
  var startReadRow = emptyRow;
  var totalReadRows = backwardSteps + 1; 
  
  var rawValues = Trns.trs.getRange(startReadRow, 1, totalReadRows, 6).getValues();
  
  var chronologicalData = [];
  var rowMap = {}; 

  for (var i = 0; i < rawValues.length; i++) {
    var r = rawValues[i];
    var actualRow = startReadRow + i;
    var rawClose = Number(r[4]);

    if (r[0] && r[4] !== "" && !isNaN(rawClose) && rawClose > 0) {
      var rawVol = Number(r[5]);
      var cleanVolume = (isNaN(rawVol) || r[5] === "#N/A" || rawVol <= 0) ? 1 : rawVol;

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

  chronologicalData.sort(function(a, b) {
    return a.parsedDate - b.parsedDate;
  });

  var totalBars = chronologicalData.length;
  if (totalBars === 0) return null;

  // 3. MACD VE DURUM HESAPLAMA
  var macdResults = (totalBars >= 17) ? calculateMACD(chronologicalData, 12, 26, 9) : null;
  var allStates = chronologicalData.map(function(bar, idx) {
    return (macdResults && macdResults[idx]) 
      ? calculate4StatesFromOHLCV(bar, macdResults[idx])
      : ((bar.close >= bar.open) ? 1 : 0);
  });

  // HATA ÖLÇÜM DEĞİŞKENLERİ
  var sumAbsoluteError = 0; 
  var sumActualClose = 0;   
  var sumPercentageError = 0; 
  var validBacktestCount = 0;

  // 4. BACKTEST - 22 ADIM (BUGÜNDEN BAŞLAR)
  for (var step = 0; step < backwardSteps; step++) {
    var targetRow = emptyRow + step;
    
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
    
    // HAM MARKOV TAHMİNİ (Unconstrained)
    var rawPredOpen = prevClose * (1 + (changePct * 0.1));
    var rawPredClose = rawPredOpen * (1 + changePct);
    var rawPredHigh = Math.max(rawPredOpen, rawPredClose) * (1 + highChangePct + 0.003);
    var rawPredLow  = Math.min(rawPredOpen, rawPredClose) * (1 - lowChangePct - 0.003);

    // POST-PROCESSING CONVEX OPTIMIZATION (POCS DÜZELTMESİ)
    var optimizedOHLC = projectToConvexOHLC(rawPredOpen, rawPredHigh, rawPredLow, rawPredClose);

    var btPred = {
      open: Number(optimizedOHLC.open.toFixed(2)),
      close: Number(optimizedOHLC.close.toFixed(2)),
      high: Number(optimizedOHLC.high.toFixed(2)),
      low: Number(optimizedOHLC.low.toFixed(2)),
      dev: "",
      dateStr: formatDateCustom(baseBar.parsedDate)
    };

    if (baseBar.close > 0) {
      var devVal = (baseBar.close - btPred.close) / baseBar.close;
      btPred.dev = Number(devVal.toFixed(3));

      var absDiff = Math.abs(baseBar.close - btPred.close);
      sumAbsoluteError += absDiff;
      sumActualClose += baseBar.close;
      sumPercentageError += Math.abs(devVal);
      validBacktestCount++;
    }

    var writeValues = [[btPred.open, btPred.high, btPred.low, btPred.close, btPred.dev]];

    Trns.trs.getRange(targetRow, N_START_COL).setValue(btPred.dateStr);
    Trns.trs.getRange(targetRow, N_START_COL + 1, 1, 5).setValues(writeValues);
  }

  // 22 GÜNLÜK HATA ÖLÇÜMLERİNİ YAZDIR
  if (validBacktestCount > 0 && sumActualClose > 0) {
    var wmapePct = (sumAbsoluteError / sumActualClose) * 100;
    var meanDevPct = (sumPercentageError / validBacktestCount) * 100;

    var metricRow = emptyRow - 2;
    Trns.trs.getRange(metricRow, N_START_COL).setValue("22D WMAPE:");
    Trns.trs.getRange(metricRow, N_START_COL + 1).setValue(wmapePct.toFixed(2) + "%");
    Trns.trs.getRange(metricRow, N_START_COL + 2).setValue("22D MAE:");
    Trns.trs.getRange(metricRow, N_START_COL + 3).setValue(meanDevPct.toFixed(2) + "%");
  }

  // 5. FORWARD PREDICTION (10 ADIM)
  var currentTargetRow = emptyRow;
  var baseBarFwd = chronologicalData[totalBars - 1]; 
  var currentBaseClose = Number(baseBarFwd.close);
  var currentState = allStates[totalBars - 1];
  var baseDate = new Date(baseBarFwd.parsedDate);

  for (var f = 1; f <= 11; f++) {
    var initStateVecF = getInitialStateVector(currentState);
    var matF = build4x4TransitionMatrix(allStates, totalBars - 1, Math.min(totalBars - 1, 22));
    var vecF = multiplyVectorMatrix(initStateVecF, matF);
    
    var bullWeightF = vecF[1] + vecF[3]; 
    var bearWeightF = vecF[0] + vecF[2];

    var changePctF = (bullWeightF - 0.5) * 0.04;
    var highChangePctF = Math.max(0, (bullWeightF - 0.5) * 0.02);
    var lowChangePctF  = Math.max(0, (bearWeightF - 0.5) * 0.02);

    var rawFwdOpen = currentBaseClose * (1 + (changePctF * 0.05));
    var rawFwdClose = rawFwdOpen * (1 + changePctF);
    var rawFwdHigh = Math.max(rawFwdOpen, rawFwdClose) * (1 + highChangePctF + 0.003);
    var rawFwdLow  = Math.min(rawFwdOpen, rawFwdClose) * (1 - lowChangePctF - 0.003);

    // CONVEX PROJEKSİYON
    var optFwd = projectToConvexOHLC(rawFwdOpen, rawFwdHigh, rawFwdLow, rawFwdClose);

    if (baseDate.getDay() === 6) baseDate.setDate(baseDate.getDate() + 2);
    if (baseDate.getDay() === 0) baseDate.setDate(baseDate.getDate() + 1);

    var fwdWriteValues = [[
      Number(optFwd.open.toFixed(2)),
      Number(optFwd.high.toFixed(2)),
      Number(optFwd.low.toFixed(2)),
      Number(optFwd.close.toFixed(2)),
      ""
    ]];

    Trns.trs.getRange(currentTargetRow, G_START_COL).setValue(formatDateCustom(baseDate));
    Trns.trs.getRange(currentTargetRow, G_START_COL + 1, 1, 5).setValues(fwdWriteValues);

    currentBaseClose = optFwd.close;
    currentState = (optFwd.close >= optFwd.open) ? 3 : 0; 
    currentTargetRow--; 
    baseDate.setDate(baseDate.getDate() + 1);
  }

  // 6. DEVIATION FORMÜLÜ (L127)
  var devFormula = "=(E" + emptyRow + "-K" + emptyRow + ")/E" + emptyRow;
  Trns.trs.getRange(emptyRow, G_START_COL + 5).setFormula(devFormula);

  return { emptyRow: emptyRow, totalBarsProcessed: totalBars };
}

/******************************************************************************
 * POST-PROCESSING CONVEX OPTIMIZATION (PROJECTION ONTO CONVEX SETS)         *
 ******************************************************************************/

/**
 * Ham Markov tahminlerini kısıt kümesine projekte eden Dışbükey İyileştirici
 */
function projectToConvexOHLC(open, high, low, close) {
  var optOpen = open;
  var optClose = close;

  var maxBody = Math.max(optOpen, optClose);
  var minBody = Math.min(optOpen, optClose);

  // Kısıt 1: High >= max(Open, Close)
  var optHigh = Math.max(high, maxBody * 1.001); // En az %0.1 üst fitil marjı

  // Kısıt 2: Low <= min(Open, Close)
  var optLow = Math.min(low, minBody * 0.999);  // En az %0.1 alt fitil marjı

  // Kısıt 3: High >= Low mantıksal kontrolü
  if (optHigh < optLow) {
    optHigh = maxBody * 1.002;
    optLow = minBody * 0.998;
  }

  return {
    open: optOpen,
    high: optHigh,
    low: optLow,
    close: optClose
  };
}

/********************************0HLCV / MACD Yardımcıları********************/

function calculateMACD(data, fastPeriod, slowPeriod, signalPeriod) {
  fastPeriod = fastPeriod || 12; slowPeriod = slowPeriod || 26; signalPeriod = signalPeriod || 9;
  var closes = data.map(function(bar) { return bar.close; });
  var fastEMA = calculateEMA(closes, fastPeriod);
  var slowEMA = calculateEMA(closes, slowPeriod);
  var macdLine = [];
  for (var i = 0; i < closes.length; i++) {
    macdLine.push((fastEMA[i] !== null && slowEMA[i] !== null) ? fastEMA[i] - slowEMA[i] : null);
  }
  var validIndices = [], validValues = [];
  for (var j = 0; j < macdLine.length; j++) {
    if (macdLine[j] !== null) { validIndices.push(j); validValues.push(macdLine[j]); }
  }
  var validSignal = calculateEMA(validValues, signalPeriod);
  var signalLine = new Array(closes.length).fill(null);
  for (var k = 0; k < validIndices.length; k++) { signalLine[validIndices[k]] = validSignal[k]; }
  var results = [];
  for (var m = 0; m < closes.length; m++) { results.push({ macd: macdLine[m], signal: signalLine[m] }); }
  return results;
}

function calculateEMA(values, period) {
  var ema = new Array(values.length).fill(null);
  if (values.length < period) return ema;
  var sum = 0;
  for (var i = 0; i < period; i++) sum += values[i];
  ema[period - 1] = sum / period;
  var mult = 2 / (period + 1);
  for (var j = period; j < values.length; j++) {
    ema[j] = (values[j] - ema[j - 1]) * mult + ema[j - 1];
  }
  return ema;
}

function calculate4StatesFromOHLCV(bar, macdObj) {
  var isUp = bar.close >= bar.open;
  var isBullishMACD = macdObj ? (macdObj.macd >= macdObj.signal) : true;
  if (!isUp && !isBullishMACD) return 0;
  if (!isUp && isBullishMACD)  return 1;
  if (isUp && !isBullishMACD)  return 2;
  return 3;                             
}

function build4x4TransitionMatrix(states, endIdx, windowSize) {
  var matrix = [[0,0,0,0],[0,0,0,0],[0,0,0,0],[0,0,0,0]];
  var startIdx = Math.max(0, endIdx - windowSize);
  for (var i = startIdx; i < endIdx; i++) {
    if (states[i] !== undefined && states[i + 1] !== undefined) matrix[states[i]][states[i + 1]]++;
  }
  for (var r = 0; r < 4; r++) {
    var rowSum = matrix[r].reduce(function(a, b) { return a + b; }, 0);
    for (var c = 0; c < 4; c++) matrix[r][c] = rowSum > 0 ? matrix[r][c] / rowSum : 0.25;
  }
  return matrix;
}

function getInitialStateVector(state) {
  var vec = [0, 0, 0, 0];
  if (state >= 0 && state <= 3) vec[state] = 1; else vec[0] = 1;
  return vec;
}

function multiplyVectorMatrix(vec, mat) {
  var result = [0, 0, 0, 0];
  for (var c = 0; c < 4; c++) {
    for (var r = 0; r < 4; r++) result[c] += vec[r] * mat[r][c];
  }
  return result;
}

function formatDateCustom(dateVal) {
  if (!dateVal) return "";
  var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return "";
  var day = ("0" + d.getDate()).slice(-2);
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  var year = d.getFullYear().toString().slice(-2);
  return day + "." + month + "." + year;
}