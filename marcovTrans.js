function testmarcovChain() {
  Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  Trns.trs = Trns.sht.getActiveSheet();
  Trns.Name = Trns.trs.getSheetName()  
  Logger.log('Name %s',Trns.Name) 
 // setOrderRow()
  processOHLCWithDynamicMarkov()
}
function prsOHLCWithDynamicMarkov() {
  let sr = nextcell() // 110
  src = 'J'+ sr
  Trns.trs.getRange(src).activate();
  src = Trns.trs.getSelection()
    .getNextDataRange(SpreadsheetApp.Direction.DOWN).getA1Notation()
  Logger.log(' whole src %s',src)
  Logger.log(Trns.trs.getRange(src).getDisplayValues())
  dateData =  GetDataArray(Trns.trs, src); // 
  lng = dateData.length
  dateData.push(Trns.trs.getRange('A'+sr).getValue())
  Logger.log('src %s dateData lenght %s dateData %s ', src, lng, dateData)
  src = 'K' + sr+':N' + (sr+lng-1)
  Logger.log('src %s',src)
  const ohlcData = Trns.trs.getRange(src).getValues();
  Logger.log(' ohlcData %s', ohlcData)

/**/
  // 1. Derive discrete states from OHLC data
  let states = [];
  for (let i = 0; i < ohlcData.length; i++) {
    if (i === 0) {
      states.push(1); // Assume stable for first row
    } else {
      const prevClose = ohlcData[i - 1][3];
      const currClose = ohlcData[i][3];
      if (currClose > prevClose * 1.005) states.push(2);  // Up
      else if (currClose < prevClose * 0.995) states.push(0); // Down
      else states.push(1); // Stable
    }
  }

  // 2. Build transition matrix dynamically from states
  let transitionMatrix = buildTransitionMatrix(states, 3);

  // 3. Predict next states and OHLC values
  let outRows = [];
  for (let i = 0; i < ohlcData.length; i++) {
    let nextState = markovNextState(states[i], transitionMatrix);
    let prediction = predictNextOHLC(
      {
        Open: ohlcData[i][0],
        High: ohlcData[i][1],
        Low: ohlcData[i][2],
        Close: ohlcData[i][3],
      },
      nextState,
      transitionMatrix
    );

    // Use next date from the dataset as next date
    // let nextDate = dateData[i + 1][0];

    outRows.push([
      dateData[i+1],
      nextState,
      prediction.ohlc.Open,
      prediction.ohlc.High,
      prediction.ohlc.Low,
      prediction.ohlc.Close,
    ]);
  }

  // 4. Write predictions back to sheet (F: next date, G: state, H-K: OHLC)
  Trns.trs.getRange(sr+1, 16, outRows.length, 6).setValues(outRows);
  Logger.log("Dynamic Markov OHLC prediction done.");
}

// Utility: Build transition matrix from states sequence
function buildTransitionMatrix(states, numStates) {
  let counts = [];
  for (let i = 0; i < numStates; i++) {
    counts[i] = new Array(numStates).fill(0);
  }
  for (let i = 0; i < states.length - 1; i++) {
    counts[states[i]][states[i + 1]]++;
  }
  let matrix = [];
  for (let i = 0; i < numStates; i++) {
    let total = counts[i].reduce((a, b) => a + b, 0);
    if (total === 0) matrix[i] = new Array(numStates).fill(1 / numStates);
    else matrix[i] = counts[i].map(c => c / total);
  }
  return matrix;
}

// Existing Markov next state function
function markovNextState(currentStateIndex, transitionMatrix) {
  let probabilities = transitionMatrix[currentStateIndex];
  let rand = Math.random();
  let cumulative = 0;
  for (let i = 0; i < probabilities.length; i++) {
    cumulative += probabilities[i];
    if (rand < cumulative) return i;
  }
  return probabilities.length - 1;
}

// Existing predictNextOHLC (make sure this includes the toFixed fix)
function predictNextOHLC(previousOHLC, nextState, transitionMatrix) {
  let open = previousOHLC.Open;
  let high = previousOHLC.High;
  let low = previousOHLC.Low;
  let close = previousOHLC.Close;

  let newOHLC = {};

  switch (nextState) {
    case 0: // Down
      newOHLC.Open = open * 0.99;
      newOHLC.High = high;
      newOHLC.Low = low * 0.98;
      newOHLC.Close = close * 0.985;
      break;
    case 1: // Stable
      newOHLC.Open = open * (1 + (Math.random() - 0.5) * 0.002);
      newOHLC.High = high * (1 + (Math.random() - 0.5) * 0.002);
      newOHLC.Low = low * (1 + (Math.random() - 0.5) * 0.002);
      newOHLC.Close = close * (1 + (Math.random() - 0.5) * 0.002);
      break;
    case 2: // Up
      newOHLC.Open = open * 1.01;
      newOHLC.High = high * 1.02;
      newOHLC.Low = low;
      newOHLC.Close = close * 1.015;
      break;
  }

  // Safe rounding to 2 decimals
  for (let key in newOHLC) {
    let num = Number(newOHLC[key]);
    newOHLC[key] = parseFloat(num.toFixed(2));
  }

  return { state: nextState, ohlc: newOHLC };
}
function weightedAverageErrorCorrection() {
 // var ss = SpreadsheetApp.getActiveSpreadsheet();
 // var sheet = ss.getSheetByName("TahminVeGerceklesen");
 // var outputSheet = ss.getSheetByName("DuzeltmeSonuclari");
  
  var lastRow = sheet.getLastRow();
  var daysToConsider = 30;
  if (lastRow < daysToConsider + 1) {
    Logger.log("Yeterli veri yok.");
    return;
  }
  
  // Son 30 günün tahmin ve gerçekleşen OHLC değerlerini al (Open, High, Low, Close için sütun B-E ve F-I)
  var predicted = sheet.getRange(lastRow - daysToConsider + 1, 2, daysToConsider, 4).getValues();
  var actual = sheet.getRange(lastRow - daysToConsider + 1, 6, daysToConsider, 4).getValues();
  
  // Ağırlıklar (son güne daha fazla ağırlık vermek için ters sıra, örn: 1, 2, 3, ..., 30)
  var weights = [];
  for (var i = 1; i <= daysToConsider; i++) {
    weights.push(i);
  }
  
  var totalWeight = weights.reduce((a, b) => a + b, 0);
  
  // Her OHLC için ağırlıklı hata toplamı
  var weightedErrorSum = [0, 0, 0, 0];
  
  for (var i = 0; i < daysToConsider; i++) {
    for (var j = 0; j < 4; j++) {
      var error = actual[i][j] - predicted[i][j];
      weightedErrorSum[j] += error * weights[i];
    }
  }
  
  // Ağırlıklı ortalama hata
  var weightedAvgError = weightedErrorSum.map(function(sum) {
    return sum / totalWeight;
  });
  
  // Bugünkü tahmin değerleri
  var todayPredicted = sheet.getRange(lastRow, 2, 1, 4).getValues()[0];
  
  var correctionFactor = 1; // Ağırlıklı ortalama hatanın tamamını veya bir kısmını uygulayabilirsiniz
  
  // Düzeltme uygulama
  var correctedToday = [];
  for (var k = 0; k < 4; k++) {
    correctedToday.push(todayPredicted[k] + correctionFactor * weightedAvgError[k]);
  }
  
  // Düzeltmeyi sonuç sayfasına ekle
  outputSheet.appendRow([new Date(), ...correctedToday]);
  
  Logger.log("Bugünkü düzeltme tamamlandı: " + JSON.stringify(correctedToday));
}






/*
Usage
Run processOHLCWithDynamicMarkov()

It will dynamically build and update the transition matrix based on your current data states.

Predict the next day's state and OHLC values accordingly.

Outputs predictions in columns F-K in the OHLC_Data sheet.

This makes your model adaptive to observed market behavior each time it runs.

Let me know if you want me to further help with backtesting, output formatting, or smoother adaptive techniques!

Related
Show a runnable Apps Script that replaces -, $, and empty cells with 0 in a sheet
Use createTextFinder with regex to replace multiple characters in a range
Convert a sheet range to a 2D array and replace values then write back
Detect last row/column dynamically and apply replacements only to data area
Provide a minimal test sheet and Apps Script to validate the replacements
*/
