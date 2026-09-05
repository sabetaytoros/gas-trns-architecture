/**
 * Master Execution Pipeline using Trns.nextcell and Singleton RAM caching.
 */
function RUN_MARKOV_PIPELINE() {
  // Hard reset Singleton properties
  delete Trns._tab;
  delete Trns._stateCache;

  const ctx = Trns.getPipelineContext(90);
  if (!ctx) return;

  const DASHBOARD_START_COL = 9;  // Column I
  const SUMMARY_METRIC_COL = "M";  // Column M
  const HEADER_OFFSET = 1;

  // --- MARKOV CALCULATIONS IN RAM ---
  const rawMatrixWithHeaders = MARKOV_TRANSITION_MATRIX(ctx.stateHistory, ctx.effectiveWindow);
  const pureMatrixValues = rawMatrixWithHeaders.slice(HEADER_OFFSET).map(row => row.slice(HEADER_OFFSET));

  const verifyReport = VERIFY_TRANSITION_MATRIX(pureMatrixValues);
  const overallStatusRow = verifyReport[verifyReport.length - 1];

  if (overallStatusRow[1] !== "MATRIX VALID (Stochastic)") {
    console.log(`Matrix verification failed for ${Trns.tab.getName()}. Aborting.`);
    return;
  }

  const durationTestResults = TEST_STATE_DURATIONS(ctx.stateHistory, ctx.effectiveWindow);
  const accuracyWinRate = BACKTEST_MARKOV_ACCURACY(ctx.stateHistory, ctx.effectiveWindow);

  // --- CHART PAYLOAD & DYNAMIC BOUNDS ---
  const chartPayload = BUILD_CHART_PAYLOAD();

  // --- SINGLE WRITE RENDERING ---
  Trns.tab.getRange(1, DASHBOARD_START_COL, durationTestResults.length, durationTestResults[0].length)
          .setValues(durationTestResults);

  Trns.tab.getRange(`${SUMMARY_METRIC_COL}1`).setValue(`${ctx.effectiveWindow}-Day Win Rate (W)`);
  Trns.tab.getRange(`${SUMMARY_METRIC_COL}2`).setValue(accuracyWinRate).setNumberFormat("0.00%");

  UPDATE_PRICE_CHART(chartPayload);

  SpreadsheetApp.flush();
  console.log(`Pipeline successfully executed for ${Trns.tab.getName()} using Trns.nextcell!`);
}
  
  
  /**
 * Main Pipeline - Handles Markov Analytics and Chained Dynamic Charting.
 * Gracefully processes ANY available dataset length (>= 2 rows).
 */



function drvRUN_MARKOV_PIPELINE() {
  // --- CONFIGURATION CONSTANTS ---
  const TARGET_WINDOW_SIZE = 90;
  const ABSOLUTE_MIN_TRANSITION_ROWS = 2; // Mathematical minimum to observe 1 transition
  
  const STATE_COLUMN_LETTER = "G";
  const DASHBOARD_START_COL = 9;  // Column I
  const SUMMARY_METRIC_COL = "M";  // Column M
  const HEADER_OFFSET = 1;
  Trns.trs = Trns.sht.getActiveSheet();
  Trns.Name = Trns.trs.getSheetName()  

  delete nextcell.cache;
  // --- 1. DYNAMIC BOUNDS VIA SINGLETONS ---
  const lastRow = Trns.trs.getLastRow();
  const DATA_START_ROW = nextcell()+1; // Singleton call
  Logger.log('Name %s DATA_START_ROW %s',Trns.Name, DATA_START_ROW)
  const availableDataRows = (lastRow - DATA_START_ROW) + 1;

  // Only exit if there are fewer than 2 rows (cannot compute a transition)
  if (availableDataRows < ABSOLUTE_MIN_TRANSITION_ROWS) {
    Logger.log(`Skipping sheet ${Trns.trs.getName()}: Insufficient rows to form a transition (Available: ${availableDataRows}).`);
    return;
  }

  // Automatically adapt window size to whatever length is available (up to 90)
  const effectiveWindowSize = Math.min(availableDataRows, TARGET_WINDOW_SIZE);
  const targetStartRow = (lastRow - effectiveWindowSize) + 1;

  // --- 2. MARKOV ANALYTICS (IN-MEMORY) ---
  const stateRangeString = `${STATE_COLUMN_LETTER}${targetStartRow}:${STATE_COLUMN_LETTER}${lastRow}`;
    Logger.log('stateRangeString %s ', stateRangeString)
    Logger.log('rng %s ', Trns.trs.getRange(stateRangeString).getA1Notation())
  const stateHistory = Trns.trs.getRange(stateRangeString).getValues();

  const rawMatrixWithHeaders = MARKOV_TRANSITION_MATRIX(stateHistory, effectiveWindowSize);
  const pureMatrixValues = rawMatrixWithHeaders.slice(HEADER_OFFSET).map(row => row.slice(HEADER_OFFSET));
  
  const verifyReport = VERIFY_TRANSITION_MATRIX(pureMatrixValues);
  const overallStatusRow = verifyReport[verifyReport.length - 1];

  if (overallStatusRow[1] !== "MATRIX VALID (Stochastic)") {
    Logger.log(`Matrix verification failed for ${Trns.trs.getName()}. Aborting execution.`);
    return;
  }

  const durationTestResults = TEST_STATE_DURATIONS(stateHistory, effectiveWindowSize);
  const accuracyWinRate = BACKTEST_MARKOV_ACCURACY(stateHistory, effectiveWindowSize);

  // --- 3. CHART PAYLOAD GENERATION (IN-MEMORY BOUNDS) ---
  const chartPayload = BUILD_CHART_PAYLOAD(targetStartRow, lastRow);

  // --- 4. CHAINED SINGLE-WRITE RENDERING ---
  // Output Dashboard Tables
  Trns.trs.getRange(1, DASHBOARD_START_COL, durationTestResults.length, durationTestResults[0].length)
          .setValues(durationTestResults);

  // Dynamic Label reflects exact number of days evaluated
  Trns.trs.getRange(`${SUMMARY_METRIC_COL}1`).setValue(`${effectiveWindowSize}-Day Win Rate (W)`);
  Trns.trs.getRange(`${SUMMARY_METRIC_COL}2`).setValue(accuracyWinRate).setNumberFormat("0.00%");

  // Update Embedded Chart
  UPDATE_PRICE_CHART(chartPayload);

  SpreadsheetApp.flush();
  Logger.log(`Pipeline executed successfully for ${Trns.trs.getName()} using all available data (${effectiveWindowSize} days).`);
}
/**
 * Calculates the 4x4 Markov Transition Probability Matrix.
 * 
 * @param {Range|String} startCell The starting top-left cell (e.g., A2 or "A2")
 * @param {Number} numDays Number of historical days/rows to process (e.g., 100)
 * @return {Array[]} 4x4 Probability Transition Matrix (%)
 * @customfunction
 */
function MARKOV_TRANSITION_MATRIX(startCell, numDays) {
  Logger.log('MARKOV_TRANSITION_MATRIX startCell %s, numDays %s', startCell, numDays)
  const { opens, highs, lows, closes, volumes } = extractOHLCVData(startCell, numDays);
  const dataLength = closes.length;

  if (dataLength < 25) {
    return [["Error: Minimum 25 valid OHLCV rows required."]];
  }

  const stateSequence = [];

  // Iterate backwards from the past to index 0 (Today)
  for (let i = dataLength - 2; i >= 0; i--) {
    const todayAvg = (opens[i] + highs[i] + lows[i] + closes[i]) / 4;
    const prevAvg = (opens[i + 1] + highs[i + 1] + lows[i + 1] + closes[i + 1]) / 4;

    const logReturn = Math.log(todayAvg / prevAvg);
    const gapReturn = Math.log(opens[i] / closes[i + 1]);

    let volumeSum = 0;
    const period = Math.min(20, dataLength - 1 - i);
    for (let k = 1; k <= period; k++) {
      volumeSum += volumes[i + k];
    }
    const avgVolume = volumeSum / period;
    const isHighVolume = volumes[i] > avgVolume;

    let state = 1;
    if ((logReturn > 0.003 || gapReturn > 0.005) && isHighVolume) {
      state = 0;
    } else if ((logReturn < -0.01 || gapReturn < -0.01) && isHighVolume) {
      state = 3;
    } else if (logReturn < -0.003) {
      state = 2;
    } else {
      state = 1;
    }

    stateSequence.push(state);
  }

  const countMatrix = [
    [0, 0, 0, 0],
    [0, 0, 0, 0],
    [0, 0, 0, 0],
    [0, 0, 0, 0]
  ];

  for (let t = 0; t < stateSequence.length - 1; t++) {
    countMatrix[stateSequence[t]][stateSequence[t + 1]]++;
  }

  const result = [
    ["Today \\ Tomorrow", "S0 (Bull)", "S1 (Neutral)", "S2 (Pullback)", "S3 (Bear)"]
  ];
  const labels = ["S0 (Bull)", "S1 (Neutral)", "S2 (Pullback)", "S3 (Bear)"];

  for (let r = 0; r < 4; r++) {
    const rowSum = countMatrix[r].reduce((a, b) => a + b, 0);
    const rowOutput = [labels[r]];

    for (let c = 0; c < 4; c++) {
      rowOutput.push(rowSum > 0 ? (countMatrix[r][c] / rowSum) : 0);
    }
    result.push(rowOutput);
  }

  return result;
}