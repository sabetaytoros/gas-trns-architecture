// Global scope - 0_Global.gs
const Trns = {
  _sht: null,
  _cch: {},

  // 1. WorkBook
  get sht() {
    if (!this._sht) {
      this._sht = SpreadsheetApp.getActiveSpreadsheet();
      console.log("Connected to Work Book: " + this._sht.getName());
    }
    return this._sht;
  },

  // 2. Sekmeler Dizisi (.map, .filter vb. için)
  get shts() {
    return this.sht.getSheets();
  },

  // --- ORİJİNAL GLOBAL DEĞİŞKENLER (Aynen korundu) ---
  oName: [],
  sName: [],
  tupple: {},
  
  ssId: "1zINSfs5yQCysx1bIn0LllK9eIQEkjE2VmnCTx2RH5XI",
  wurl: "https://script.google.com/macros/s/AKfycby5uvkdlQ0YAUEb52UFNQdcNDTI41qbH08q7CO5Ds_f/dev"
};
/**
 * Dynamic Tab Pointer & Proxy Wrapper
 */
(() => {
  let activeSheet = null;

  function trs(tnm) {
    if (tnm === 'active' || !tnm) {
      activeSheet = Trns.sht.getActiveSheet();
      Trns._cch[activeSheet.getName()] = activeSheet;
    } else if (tnm) {
      if (!Trns._cch[tnm]) {
        const s = Trns.sht.getSheetByName(tnm);
        if (!s) throw new Error("Tab not found: " + tnm);
        Trns._cch[tnm] = s;
      }
      activeSheet = Trns._cch[tnm];
    }

    if (!activeSheet) {
      throw new Error("No active tab selected!");
    }

    return activeSheet;
  }

  // 1. Dynamic Proxy Object Binding for Trns.trs
  Object.defineProperty(Trns, 'trs', {
    get: function() {
      if (!activeSheet) {
        activeSheet = Trns.sht.getActiveSheet();
      }

      const fn = function(tnm) { return trs(tnm); };

      return new Proxy(fn, {
        get(target, prop) {
          if (prop in target) return target[prop];
          const val = activeSheet[prop];
          return typeof val === 'function' ? val.bind(activeSheet) : val;
        }
      });
    },
    configurable: true
  });

  // 2. Backward-Compatibility Getter for Trns.Name
  Object.defineProperty(Trns, 'Name', {
    get: function() {
      if (!activeSheet) {
        activeSheet = Trns.sht.getActiveSheet();
      }
      return activeSheet.getName();
    },
    configurable: true
  });
})();


function nextcell(r = 24) {
  if (typeof nextcell.cache === 'undefined') {
    nextcell.cache = singleNext(r)
  }
  return nextcell.cache;
}


const userProp = PropertiesService.getUserProperties()

const GLOBAL_WEEKEND = isWeekend();
/**
 * The main function that runs every morning, but skips Sundays and Mondays.
 */


/*****************************************************************************/
const flgBt = { FGUNSONU: 0, FHAFTASONU: 1}
BatchType = flgBt.FGUNSONU
function triggerFunction() {
  // Skip execution on weekends
  if (GLOBAL_WEEKEND) return 
  // Place your weekday logic here:
  checkSesionFinish() // GunSonu Baslatir
  startBatchProcess()
  Logger.log("Weekday 08:00 AM trigger executed successfully.");
}
function triggerSaturdayNightFunction() {
  // Place your weekday logic here:
  checkSesionFinish() // GunSonu Baslatir
  Logger.log("Weekday 08:00 AM trigger executed successfully.");
  // hafta sonu işlemleri baslatılır
  BatchType = flgBt.FHAFTASONU
  Logger.log('Batch type %s', BatchType)
  if (Trns.sht.getRange('Kebir!G1').getValue() == 7)
    startBatchProcess() // Hafta sonu işlemleri yapılır
}
/**
 * Run this function ONCE to set up the daily 8:00 AM trigger.
 */
function createDailyTrigger() {
  // Clear any existing triggers for 'mainFunction' to prevent duplicates
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'triggerFunction') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }

  // Create a daily trigger (it will run every day, but logic will filter weekends)
  ScriptApp.newTrigger('triggerFunction')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(0)
    .create();
}
function createSaturdayTrigger() {
  // Clear any existing triggers for 'mainFunction' to prevent duplicates
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'triggerSaturdayNightFunction') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }

  // Create a daily trigger (it will run every day, but logic will filter weekends)
  ScriptApp.newTrigger('triggerSaturdayNightFunction')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SATURDAY)
      .atHour(0) // 0 represents midnight 
      .create();

  Logger.log('Trigger successfully scheduled for Saturday midnight.');
}
function trimAllSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let totalCellsReduced = 0;

  sheets.forEach(sheet => {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const maxRows = sheet.getMaxRows();
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const maxCols = sheet.getMaxColumns();

    // Delete extra rows
    if (maxRows > lastRow) {
      sheet.deleteRows(lastRow + 1, maxRows - lastRow);
      totalCellsReduced += (maxRows - lastRow) * maxCols;
    }

    // Delete extra columns
    if (maxCols > lastCol) {
      sheet.deleteColumns(lastCol + 1, maxCols - lastCol);
      totalCellsReduced += (maxCols - lastCol) * lastRow;
    }
  });

  Logger.log("Trimmed blank cells. Reduced total count by: " + totalCellsReduced);
}