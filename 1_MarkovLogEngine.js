/**
 * 1_MarkovLogEngine.gs
 * MARKOV LOG-RATIO ENGINE - SÜRÜM v6.2.0 (Volume-Weighted Log-Ratio)
 * - Hacim Ağırlıklı Log-Ratio Ortalamaları (Volume-Weighted Log-Ratios) Entegrasyonu
 * - 22 Günlük Backtest WMAPE (%) ve MAE (%) Ölçüm Entegrasyonu
 * - Matematiksel OHLC Mantık Garantisi (High >= max(O,C) ve Low <= min(O,C))
 */

function drvEngineLogRatio() {
  var ts = 22; // 22 Adet Backtest Adımı
  return runMarkovEngineLogRatio(ts);
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

function runMarkovEngineLogRatio(backwardSteps = 22) {
  var VERSION = "MARKOV LOG-RATIO ENGINE - SÜRÜM v6.2.0 (Vol-Weighted)";

  var G_START_COL = 7;  // G Sütunu (Forward 10 Adım Tahmin)
  var N_START_COL = 14; // N Sütunu (Backtest 22 Adım Tahmin)
  
  var emptyRow = nextcell(); // Örn: Satır 127
  if (!emptyRow) return null;

  // 0. DİNAMİK SÜRÜM BİLGİSİ VE FORWARD / BACKTEST BAŞLIKLARI
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
      var cleanVolume = (isNaN(rawVol) || r[5] === "#N/A" || rawVol <= 0) ? 1 : rawVol; // Hacim sıfırsa 1 alarak bölme hatasını önle

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

      // Log-Ratio Dönüşümleri
      var maxBody = Math.max(barObj.open, barObj.close);
      var minBody = Math.min(barObj.open, barObj.close);

      barObj.r_oc = Math.log(barObj.close / barObj.open);
      barObj.r_ho = Math.max(0, Math.log(barObj.high / maxBody));
      barObj.r_ol = Math.max(0, Math.log(minBody / barObj.low));

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

  // Hacim Ağırlıklı Log-Ratio İstatistikleri Hesaplama
  var stateStats = computeVolumeWeightedStateLogRatioStats(chronologicalData, allStates);

  // HATA ÖLÇÜM DEĞİŞKENLERİ (22 Günlük)
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

    var expRatios = computeExpectedLogRatios(vec, stateStats);

    var prevClose = prevBar.close;
    var predOpen = prevClose;
    var predClose = predOpen * Math.exp(expRatios.r_oc);

    var predMaxBody = Math.max(predOpen, predClose);
    var predMinBody = Math.min(predOpen, predClose);

    var predHigh = predMaxBody * Math.exp(expRatios.r_ho);
    var predLow  = predMinBody * Math.exp(-expRatios.r_ol);

    var btPred = {
      open: Number(predOpen.toFixed(2)),
      close: Number(predClose.toFixed(2)),
      high: Number(predHigh.toFixed(2)),
      low: Number(predLow.toFixed(2)),
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
    var matF = build4x4TransitionMatrix(
      allStates, 
      totalBars - 1, 
      Math.min(totalBars - 1, 22)
    );
    var vecF = multiplyVectorMatrix(initStateVecF, matF);
    
    var expRatiosF = computeExpectedLogRatios(vecF, stateStats);

    var fwdOpen = currentBaseClose;
    var fwdClose = fwdOpen * Math.exp(expRatiosF.r_oc);

    var fwdMaxBody = Math.max(fwdOpen, fwdClose);
    var fwdMinBody = Math.min(fwdOpen, fwdClose);

    var fwdHigh = fwdMaxBody * Math.exp(expRatiosF.r_ho);
    var fwdLow  = fwdMinBody * Math.exp(-expRatiosF.r_ol);

    if (baseDate.getDay() === 6) baseDate.setDate(baseDate.getDate() + 2);
    if (baseDate.getDay() === 0) baseDate.setDate(baseDate.getDate() + 1);

    var fwdWriteValues = [[
      Number(fwdOpen.toFixed(2)),
      Number(fwdHigh.toFixed(2)),
      Number(fwdLow.toFixed(2)),
      Number(fwdClose.toFixed(2)),
      ""
    ]];

    Trns.trs.getRange(currentTargetRow, G_START_COL).setValue(formatDateCustom(baseDate));
    Trns.trs.getRange(currentTargetRow, G_START_COL + 1, 1, 5).setValues(fwdWriteValues);

    currentBaseClose = fwdClose;
    currentState = (fwdClose >= fwdOpen) ? 3 : 0; 
    currentTargetRow--; 
    baseDate.setDate(baseDate.getDate() + 1);
  }

  // 6. DEVIATION FORMÜLÜ (L127)
  var devFormula = "=(E" + emptyRow + "-K" + emptyRow + ")/E" + emptyRow;
  Trns.trs.getRange(emptyRow, G_START_COL + 5).setFormula(devFormula);

  return { emptyRow: emptyRow, totalBarsProcessed: totalBars };
}

/******************************************************************************
 * HACİM AĞIRLIKLI İSTATİSTİK FONKSİYONU                                     *
 ******************************************************************************/

/**
 * Durumların Log-Oranlarını İşlem Hacimleri ile Ağırlıklandırarak Ortalama Hesaplar
 */
function computeVolumeWeightedStateLogRatioStats(bars, states) {
  var stats = {
    0: { sumVolR_oc: 0, sumVolR_ho: 0, sumVolR_ol: 0, totalVol: 0 },
    1: { sumVolR_oc: 0, sumVolR_ho: 0, sumVolR_ol: 0, totalVol: 0 },
    2: { sumVolR_oc: 0, sumVolR_ho: 0, sumVolR_ol: 0, totalVol: 0 },
    3: { sumVolR_oc: 0, sumVolR_ho: 0, sumVolR_ol: 0, totalVol: 0 }
  };

  for (var i = 0; i < bars.length; i++) {
    var st = states[i];
    var vol = bars[i].volume || 1;

    if (stats[st]) {
      stats[st].sumVolR_oc += bars[i].r_oc * vol;
      stats[st].sumVolR_ho += bars[i].r_ho * vol;
      stats[st].sumVolR_ol += bars[i].r_ol * vol;
      stats[st].totalVol += vol;
    }
  }

  var avgStats = {};
  for (var s = 0; s < 4; s++) {
    var totVol = stats[s].totalVol;
    if (totVol > 0) {
      avgStats[s] = {
        r_oc: stats[s].sumVolR_oc / totVol,
        r_ho: Math.max(0.001, stats[s].sumVolR_ho / totVol),
        r_ol: Math.max(0.001, stats[s].sumVolR_ol / totVol)
      };
    } else {
      avgStats[s] = { r_oc: 0, r_ho: 0.002, r_ol: 0.002 };
    }
  }
  return avgStats;
}

function computeExpectedLogRatios(probVec, stateStats) {
  var exp_r_oc = 0;
  var exp_r_ho = 0;
  var exp_r_ol = 0;

  for (var s = 0; s < 4; s++) {
    var p = probVec[s];
    exp_r_oc += p * stateStats[s].r_oc;
    exp_r_ho += p * stateStats[s].r_ho;
    exp_r_ol += p * stateStats[s].r_ol;
  }

  return {
    r_oc: exp_r_oc,
    r_ho: Math.max(0, exp_r_ho),
    r_ol: Math.max(0, exp_r_ol)
  };
}

/********************************0HLCV / MACD Yardımcıları********************/

function calculateMACD(data, fastPeriod, slowPeriod, signalPeriod) {
  fastPeriod = fastPeriod || 12;
  slowPeriod = slowPeriod || 26;
  signalPeriod = signalPeriod || 9;

  var closes = data.map(function(bar) { return bar.close; });
  var fastEMA = calculateEMA(closes, fastPeriod);
  var slowEMA = calculateEMA(closes, slowPeriod);

  var macdLine = [];
  for (var i = 0; i < closes.length; i++) {
    if (fastEMA[i] !== null && slowEMA[i] !== null) {
      macdLine.push(fastEMA[i] - slowEMA[i]);
    } else {
      macdLine.push(null);
    }
  }

  var validMacdIndices = [];
  var validMacdValues = [];
  for (var j = 0; j < macdLine.length; j++) {
    if (macdLine[j] !== null) {
      validMacdIndices.push(j);
      validMacdValues.push(macdLine[j]);
    }
  }

  var validSignal = calculateEMA(validMacdValues, signalPeriod);
  var signalLine = new Array(closes.length).fill(null);
  for (var k = 0; k < validMacdIndices.length; k++) {
    signalLine[validMacdIndices[k]] = validSignal[k];
  }

  var results = [];
  for (var m = 0; m < closes.length; m++) {
    results.push({
      macd: macdLine[m],
      signal: signalLine[m]
    });
  }
  return results;
}

function calculateEMA(values, period) {
  var ema = new Array(values.length).fill(null);
  if (values.length < period) return ema;

  var sum = 0;
  for (var i = 0; i < period; i++) {
    sum += values[i];
  }
  var sma = sum / period;
  ema[period - 1] = sma;

  var multiplier = 2 / (period + 1);
  for (var j = period; j < values.length; j++) {
    ema[j] = (values[j] - ema[j - 1]) * multiplier + ema[j - 1];
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

function getInitialStateVector(state) {
  var vec = [0, 0, 0, 0];
  if (state >= 0 && state <= 3) {
    vec[state] = 1;
  } else {
    vec[0] = 1;
  }
  return vec;
}

function multiplyVectorMatrix(vec, mat) {
  var result = [0, 0, 0, 0];
  for (var c = 0; c < 4; c++) {
    for (var r = 0; r < 4; r++) {
      result[c] += vec[r] * mat[r][c];
    }
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
