/**
 * newgunson.gs
 */

const ADMIN_EMAIL = Session.getEffectiveUser().getEmail();
const MAX_EXECUTION_TIME_MS = 4 * 60 * 1000; // 4 minutes

const JOB_INITIALIZED_KEY = 'JOB_INITIALIZED';
const REMAINING_TABS_KEY = 'REMAINING_TABS_KEY';
const TRIGGER_FUNCTION_NAME = 'startBatchProcess';
function drvStart() {
Logger.log('Start')

  userProp.deleteProperty(JOB_INITIALIZED_KEY);
  //BatchType = flgBt.FHAFTASONU
  startBatchProcess()
}

function startBatchProcess() {
  let remainingTabNames = [];
  const isInitialized = userProp.getProperty(JOB_INITIALIZED_KEY);
  if (!isInitialized) {
    Logger.log("=== STARTING NEW BATCH JOB ===");
    // --- YOUR INITIALIZATION FUNCTION HERE ---
    runInitialSetup();
    // Fetch tab names (excluding temp or unwanted tabs if needed)
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const tabNames = sheets
      .map(sheet => sheet.getName())
      .filter(name => !exclude.includes(name));
    Logger.log('tabNames %s', tabNames)
    // Save initial state to Script Properties
    userProp.setProperty(REMAINING_TABS_KEY, JSON.stringify(tabNames));
    userProp.setProperty(JOB_INITIALIZED_KEY, 'TRUE'); // Lock initialization
    Logger.log("Initialized job with " + tabNames.length + " total tabs.");
  } else Trns.Holiday = userProp.getProperty('IS_HOLIDAY')
  Logger.log('Trns.Holiday %s', Trns.Holiday)
  // 2. ADIM: SONRAKİ TETİKLENMELERDE HAFIZADAKİ LİSTEYİ OKU
  const savedRemaining = userProp.getProperty(REMAINING_TABS_KEY);
  if (savedRemaining) {
    remainingTabNames = JSON.parse(savedRemaining);
  }
  ensure15MinTriggerExists_();
  processQueue();
}
/*********************************************************************************** */
function processQueue() {
  const startTime = Date.now();
  let remainingTabNames = JSON.parse(userProp.getProperty(REMAINING_TABS_KEY) || '[]');
  while (remainingTabNames.length > 0) {
    const elapsedTime = Date.now() - startTime;
    // 1. TIMEOUT CHECK: Save state and schedule continuation
    if (elapsedTime >= MAX_EXECUTION_TIME_MS) {
      Logger.log("Approaching limit (" + Math.round(elapsedTime / 1000) + "s elapsed). Scheduling continuation...");
      userProp.setProperty(REMAINING_TABS_KEY, JSON.stringify(remainingTabNames));
      scheduleContinuation();
      return;
    }
    // Get the next tab name from queue
    const currentTabName = remainingTabNames[0];
    Logger.log("Processing tab: " + currentTabName);
    // 2. ISOLATED TAB EXECUTION WITH TRY-CATCH
    try {
      processSingleTab(currentTabName);
      // İşlenen sekme diziden çıkarılır
      remainingTabNames.shift();
      if (remainingTabNames.length === 0) {
        // İşlenecek sekme yoksa süreci temizle ve kapat
        if (remainingTabNames.length === 0) {
          Logger.log("İşlenecek sekme kalmadı. finalizeBatchJob. cagriliyor ");
          finalizeBatchJob()
          return;
        } else {
          userProp.setProperty(REMAINING_TABS_KEY, JSON.stringify(remainingTabNames));
          Logger.log("Successfully completed: " + currentTabName);
        }
      }
    } catch (tabError) {
      // Catch failure for THIS specific tab so the queue can keep moving
      Logger.log("ERROR on tab [" + currentTabName + "]: " + tabError.stack);
    } finally {
      // 3. ALWAYS UPDATE STATE: Ensures queue progresses even on failure
      userProp.setProperty(REMAINING_TABS_KEY, JSON.stringify(remainingTabNames));
    }
  }

  // 4. BATCH COMPLETE: Wrap up 
}
/**
 * HEAVY PER-TAB LOGIC (Safely called inside try-catch block)
 */
function processSingleTab(tabName) {
  Trns.trs(tabName);
  // Custom check: throw error if tab was deleted mid-run
  if (!Trns.trs) {
    throw new Error("Sheet tab '" + tabName + "' no longer exists.");
  }
  if (BatchType != flgBt.FGUNSONU) {
    newHaftaSonu()
  } else { 
    if (canbeProcessed()) {
      if (Trns.Holiday)
        processHoliDay()
      else {
        CommonProcess()
      }
    }
  }
}  
function newHaftaSonu() {
  Trns.trs.getRange('BA14:BZ14').activate()
  Trns.trs.getRange('BA14:BZ14').copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  Trns.trs.getRange('BA16:BY16').activate()
  Trns.trs.getRange('BA16:BY16').copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  Trns.trs.getRange('G13:BT22').moveTo(Trns.trs.getRange('L13'));
  Trns.trs.getRange('G2:BT6').moveTo(Trns.trs.getRange('L2'));
  generateHaftalikTarih()
}
/**
 * FINALIZE JOB: Clean up storage & send a summary report if any tabs failed.
 */
function finalizeBatchJob() {
  if (!userProp.getProperty('UPDATE_KEBIR_PAGE')) {
    userProp.deleteProperty(JOB_INITIALIZED_KEY);
    Logger.log("Batch processing complete.");
    userProp.deleteProperty(REMAINING_TABS_KEY);
    deleteContinuationTriggers(TRIGGER_FUNCTION_NAME);
    Logger.log("All tabs processed with 0 errors.");
    userProp.deleteProperty('JOB_INITIALIZED');
    updateKebirPage();
  }
}
/**
 * CONTINUATION TRIGGER: Schedules next run safely in 1 minute.
 */
function scheduleContinuation() {
  try {
    deleteContinuationTriggers(TRIGGER_FUNCTION_NAME);
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
      .timeBased()
      .after(60 * 1000)
      .create();
    Logger.log("Continuation trigger active for next run.");
  } catch (err) {
    Logger.log("CRITICAL: Failed to create continuation trigger - " + err.message);
    sendAlertEmail("Trigger Creation Failure", "Could not set continuation trigger: " + err.message);
  }
}
/**
 * HOOK 1: Function executed ONCE before tab processing begins.
 */
function runInitialSetup() {
  Logger.log("Running initial setup... (e.g., clearing master logs, fetching API auth, locking sheets)");
  userProp.deleteProperty('UPDATE_KEBIR_PAGE');
  userProp.deleteProperty('IS_HOLIDAY')
  if (BatchType == flgBt.FGUNSONU) {
    enSureGunSonu()
    getSecNamesfromOrderSheet()
    orderSnapShot()
    isHoliday()
    userProp.setProperty('IS_HOLIDAY', Trns.Holiday ? 'TRUE' :'FALSE')
  }
}
/**
 * HELPER: Deletes project triggers for startBatchProcess
 */
function deleteContinuationTriggers(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  }
  if (count > 0) {
    Logger.log(`${count} adet eski trigger temizlendi.`);
  }
}
/**
 * 15 dakikalık emniyet zamanlayıcısının varlığını kontrol eder, yoksa kurar.
 */
function ensure15MinTriggerExists_() {
  const exists = ScriptApp.getProjectTriggers().some(
    t => t.getHandlerFunction() === TRIGGER_FUNCTION_NAME
  );
  if (exists) 
	deleteContinuationTriggers(TRIGGER_FUNCTION_NAME);
  ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
      .timeBased()
      .after(15 * 60 * 1000)
      .create();
  Logger.log("15 dakikalık emniyet zamanlayıcısı kuruldu.");
}
/**
 * HELPER: Sends email alerts for failures.
 */
function sendAlertEmail(subject, body) {
  try {
    MailApp.sendEmail(ADMIN_EMAIL, "[Apps Script Alert] " + subject, body);
  } catch (e) {
    Logger.log("Failed to send email alert: " + e.message);
  }
}

/**
 * UTILITY: Logs the current contents of REMAINING_TABS 
 */
function inspectScriptProperties() {
  const remaining = userProp.getProperty(REMAINING_TABS_KEY);
  Logger.log("--- REMAINING TABS ---");
  Logger.log(remaining ? JSON.parse(remaining) : "No remaining tabs found (Property is empty/deleted)");
}
/**
 * Updates a chart's vertical axis bounds using a Low column for min
 * and a High column for max.
 * 
 * @param {string} lowCol Letter of the column containing Low values (e.g., "D").
 * @param {string} highCol Letter of the column containing High values (e.g., "C").
 * @param {number} chartIndex Index of the chart on the sheet (default: 0 for first chart).
 * @param {number} paddingPercent Percentage buffer to add/subtract (default: 0.05 for 5%).
 */
function updateChartBoundsFromLowHighColumns(startRow = 0, chartIndex = 0, lowCol = "E", highCol = "D", paddingPercent = 0.05) {
  const charts = Trns.trs.getCharts();
  if (charts.length <= chartIndex) {
    Logger.log("No chart found at index " + chartIndex);
    return;
  }
  const se = startRow + 90;
  const lastRow = Trns.trs.getLastRow() > se ? se : Trns.trs.getLastRow();

  if (lastRow - startRow < 2) {
    Logger.log("Not enough data rows in sheet.");
    return;
  }
  // 1. Extract Low values from the Low column
  const lowValues = Trns.sht.getRange(`${lowCol}${startRow}:${lowCol}${lastRow}`)
    .getValues()
    .flat()
    .filter(val => typeof val === 'number' && !isNaN(val));

  // 2. Extract High values from the High column
  const highValues = Trns.sht.getRange(`${highCol}${startRow}:${highCol}${lastRow}`)
    .getValues()
    .flat()
    .filter(val => typeof val === 'number' && !isNaN(val));

  if (lowValues.length === 0 || highValues.length === 0) {
    Logger.log("Missing numeric data in specified Low/High columns.");
    return;
  }

  // Calculate absolute minimum from Low column and absolute maximum from High column
  const minLow = Math.min(...lowValues);
  const maxHigh = Math.max(...highValues);

  // Apply padding percentage
  const spread = maxHigh - minLow;
  const buffer = spread === 0 ? Math.abs(minLow) * paddingPercent || 1 : spread * paddingPercent;

  const minBound = Math.floor(minLow - buffer);
  const maxBound = Math.ceil(maxHigh + buffer);

  // Rebuild the chart with updated vertical axis boundaries
  const targetChart = charts[chartIndex];
  const updatedChart = targetChart.modify()
    .setOption('vAxis.minValue', minBound)
    .setOption('vAxis.maxValue', maxBound)
    .build();

  Trns.trs.updateChart(updatedChart);
  Logger.log(` Rows (Start ${startRow} End ${lastRow}) Chart [${chartIndex}] updated -> Min: ${minBound} (from Col ${lowCol}), Max: ${maxBound} (from Col ${highCol})`);
}

/**
 * Example execution wrapper
 */
function runLowHighUpdate() {
  // Pass Low column "E", High column "D", Chart index 0, 5% padding
  updateChartBoundsFromLowHighColumns(102);
}
