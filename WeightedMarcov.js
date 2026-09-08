calltst = 0
function testmarcovChain() {
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  Trns.trs = Trns.sht.getActiveSheet();
  Trns.Name = Trns.trs.getSheetName()  
  Logger.log('Name %s',Trns.Name) 
 // setOrderRow()

  processOHLCWithDynamicMarkov()
}
function processOHLCWithDynamicMarkov() {
  let sr = nextcell()  // Date sonraki
  let nr = sr + 1
  let ser = 'A' + (sr+20)  
  let src = 'A'+ nr
  let s ='= GOOGLEFINANCE($B$1,"all",'+ ser + ','+src+')'
  src = "J"+ (sr-1)
  Trns.trs.getRange(src).setFormula(s)
  src = 'J'+ sr
  Trns.trs.getRange(src).activate();
  src = Trns.trs.getSelection()
    .getNextDataRange(SpreadsheetApp.Direction.DOWN).getA1Notation()
//  Logger.log(' whole src %s',src)
//  Logger.log(Trns.trs.getRange(src).getDisplayValues())
  dateData =  GetDataArray(Trns.trs, src); // 
  lng = dateData.length
  dateData.push(Trns.trs.getRange('A'+sr).getValue())
//  Logger.log('src %s dateData lenght %s dateData %s ', src, lng, dateData)
  src = 'K' + sr+':N' + (sr+lng-1)
//  Logger.log('src %s',src)
  const ohlcData = Trns.trs.getRange(src).getValues();
//  Logger.log(' ohlcData %s', ohlcData)

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
  let predicted = []
  let n = ohlcData.length - 1
  const Political_Event = 'FALSE'
  for (let i = 0; i < ohlcData.length; i++) {
    let nextState = (i == n && Political_Event == 'TRUE') ? 2
                  : markovNextState(states[i], transitionMatrix);
    //Logger.log('i %s n %s nextState %s',i, n, nextState)

    let prediction = predictNxtOHLC(
      {
      Open: ohlcData[i][0],
      High: ohlcData[i][1],
      Low: ohlcData[i][2],
      Close: ohlcData[i][3],
    },
      nextState,
      transitionMatrix
    );
//Logger.log('prediction %s',prediction)
    // Use next date from the dataset as next date

    predicted.push(prediction.ohlc)//
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
  ++sr

  Trns.trs.getRange(sr, 16, outRows.length, 6).setValues(outRows)
  Trns.trs.getRange(sr,18,outRows.length,4).setNumberFormat('#,##0.00')
  Logger.log("Dynamic Markov OHLC prediction done.");
  weightedAverageErrorCorrection(ohlcData, sr)
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

// Existing predictNxtOHLC (make sure this includes the toFixed fix)
function predictNxtOHLC(previousOHLC, nextState, transitionMatrix) {
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



function weightedAverageErrorCorrection(actual,sr) {
  //Logger.log('actual %s \n', actual)
  var daysToConsider = actual.length;

  src = 'R' + sr+':U' + (sr+daysToConsider-1)
  //Logger.log('src %s',src)
  const predicted = Trns.trs.getRange(src).getValues();
  //Logger.log(' predicted %s', predicted)


  // Ağırlıklar (son güne daha fazla ağırlık vermek için ters sıra, örn: 1, 2, 3, ..., 30)
  var weights = [];
  for (var i = 1; i <= daysToConsider; i++) {
    weights.push(i);
  }
  
  var totalWeight = weights.reduce((a, b) => a + b, 0);
  
  // Her OHLC için ağırlıklı hata toplamı
  var weightedErrorSum = [0, 0, 0, 0];
  // Oransal hata dizileri
  var pctErrors = [[], [], [], []]; // [Open[], High[], Low[], Close[]]

  for (var i = 0; i < daysToConsider; i++) {
    for (var j = 0; j < 4; j++) {
      error = actual[i][j] - predicted[i][j];           
      weightedErrorSum[j] += error * weights[i]; 
      pv = error / predicted[i][j]
      pctErrors[j].push(pv);
      //Logger.log('j %s err %s wErrSum[j] %s pv %s ', j, error,weightedErrorSum[j], pv)
    }
  }
  // Ağırlıklı ortalama hata
  var weightedAvgError = weightedErrorSum.map(function(sum) {
    return sum / totalWeight;
  });

  //Logger.log('totalWeight %s weightedAvgError %s', totalWeight, weightedAvgError)

  // Bugünkü tahmin değerleri
  var todayPredicted = predicted.pop() //sheet.getRange(lastRow, 2, 1, 4).getValues()[0];
  // Logger.log('todayPredicted %s', todayPredicted)
  var correctionFactor = 1; // Ağırlıklı ortalama hatanın tamamını veya bir kısmını uygulayabilirsiniz
  
  // Düzeltme uygulama
  var correctedToday = [];
  for (var k = 0; k < 4; k++) {
    correctedToday.push(todayPredicted[k] + correctionFactor * weightedAvgError[k]);
  }
  //Logger.log('correctedToday %s', correctedToday)

  sr += daysToConsider+1
  src = 'R' + sr+':U' + sr
//  Logger.log('src %s',src)
  Trns.trs.getRange(src).setValues([correctedToday]).setNumberFormat('#,##0.00')
  Trns.trs.getRange('P'+sr).setValue("Weighted Corrected")
    // Ortalama oransal hata
  function mean(arr) {
    return arr.reduce((a, b) => a + b, 0) / arr.length;
  }
  var correctionRatios = pctErrors.map(mean); // Her kalem için ortalama düzeltme 
   // Bugünkü tahmin değerleri
  //var todayPredicted = sheet.getRange(lastRow, 2, 1, 4).getValues()[0];
  var Relcorrected = todayPredicted.map(function(val, idx) {
    return val * (1 + correctionRatios[idx]);
  });
  ++sr
  Trns.trs.getRange('P'+sr).setValue("Relatif Corrected")
  src = 'R' + sr+':U' + sr
  Trns.trs.getRange(src).setValues([Relcorrected]).setNumberFormat('#,##0.00');
}



