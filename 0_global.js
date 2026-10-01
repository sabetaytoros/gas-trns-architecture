// --- Global Scope (0_Global.gs) ---
const Trns = {
  _sht: null,
  _cch: {},
  _kbr: null,

  // 1. WorkBook
  get sht() {
    if (!this._sht) {
      this._sht = SpreadsheetApp.getActiveSpreadsheet();
      this._kbr = Trns.sht.getRange('Kebir!G1').getValue();
      console.log("Connected to Work Book: " + this._sht.getName());
    }
    return this._sht;
  },
  get kbr() {
    return this._kbr;
  },
  // 2. Sekmeler Dizisi
  get shts() {
    return this.sht.getSheets();
  },

  // --- SINGLETON MARKOV STATE CONTROLLER (v5.9.8 & v6.0) ---
  markovState: (function() {
    let instance = null;

    function createInstance() {
      return {
        gamma: 0.35, // v6.0 Newton-Raphson Sönümleme Katsayısı
        matrix: [
          [0, 0, 0, 0],
          [0, 0, 0, 0],
          [0, 0, 0, 0],
          [0, 0, 0, 0]
        ],
        P: [],
        V0: [0, 0, 0, 0],
        error: {     // v6.0 5D Hata Vektör Bitişikleri
          e_Close: 0,
          e_Vol: 0,
          multiplier: 1.0
        }
      };
    }

    return {
      // Singleton Örneğini Getir (Yoksa Oluştur)
      getInstance: function() {
        if (!instance) {
          instance = createInstance();
        }
        return instance;
      },
      // Tab Değişiminde Tekil Örneği ve Cache'i Sıfırla
      resetInstance: function() {
        instance = createInstance();
        delete nextcell.cache; // nextcell önbelleğini temizle
      }
    };
  })(),

  // --- ORİJİNAL GLOBAL DEĞİŞKENLER ---
  oName: [],
  sName: [],
  tupple: {},
  
  ssId: "1zINSfs5yQCysx1bIn0LllK9eIQEkjE2VmnCTx2RH5XI",
  wurl: "https://script.google.com/macros/s/AKfycby5uvkdlQ0YAUEb52UFNQdcNDTI41qbH08q7CO5Ds_f/dev"
};
// --- 0_Global.gs İçine Eklenebilecek MACD ve Grafik Sabitleri ---
Trns.CONFIG = {
  // MACD Gösterge Parametreleri
  FAST_EMA: 12,
  SLOW_EMA: 26,
  SIGNAL_EMA: 9,

  // Markov Motoru ve Veri Boyutu Parametreleri
  WINDOW_SIZE: 22,                       // Markov pencere boyutu (1 işlem ayı)
  TOTAL_BARS: 46,                        // EMA26 ısınma payı + 22 bar Markov matrisi
  MIN_REQUIRED_BARS: 22                  // İşlem yapmak için gereken minimum bar sayısı
};
// Google Sheets Candlestick için zorunlu sütun sırası: Low, Open, Close, High
// Bize gereken veri sütun aralıkları (örnek varsayılanlar):
Trns.INDICATOR_COLS = {
  DATE: 1,      // Sütun A: Tarih
  LOW: 2,       // Sütun B: Low
  OPEN: 3,      // Sütun C: Open
  CLOSE: 4,     // Sütun D: Close
  HIGH: 5,      // Sütun E: High
  MACD: 7,      // Sütun G: MACD Line
  SIGNAL: 8,    // Sütun H: Signal Line
  HISTOGRAM: 9  // Sütun I: MACD Histogram
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

    // TAB DEĞİŞİMİNDE AUTOMATIC SINGLETON SIFIRLAMA
    if (Trns.markovState && typeof Trns.markovState.resetInstance === 'function') 
      Trns.markovState.resetInstance();

    return activeSheet;
  }

  // Dynamic Proxy Object Binding for Trns.trs
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

  // Backward-Compatibility Getter for Trns.Name
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

// Dynamic emptyRow Getter
Object.defineProperty(Trns, 'emptyRow', {
  get: function() {
    return nextcell();
  },
  configurable: true
});

function nextcell(r = 24) {
  if (typeof nextcell.cache === 'undefined') {
    nextcell.cache = singleNext(r)
    Trns.empty = nextcell.cache - 1; // Boş satırın bir önceki satırı (Örn: 127 - 1 = 126)
    Trns.emptyRow = nextcell.cache;  // Doğrudan boş satır numarası (Örn: 127)
  }
  return nextcell.cache;
}
// RowContext Singleton Tanımı (const olarak global düzeyde)
const RowContext = (function () {
  var instance;

  function createInstance() {
    return {
      // Varsayılan satır indeksleri / bağlam değerleri
      startRow: 0,
      endRow: 0,
      currentRow: 0,
      reset: function () {
        this.startRow = 0;
        this.endRow = 0;
        this.currentRow = 0;
      }
    };
  }

  return {
    getInstance: function () {
      if (!instance) {
        instance = createInstance();
      }
      return instance;
    }
  };
})();
/*****************************************************************************/
const userProp = PropertiesService.getUserProperties()
const GLOBAL_WEEKEND = isWeekend();
/*****************************************************************************/
const flgBt = { FGUNSONU: 0, FHAFTASONU: 1}
BatchType = flgBt.FGUNSONU
/** 
 * The main function that runs every morning, but skips Sundays and Mondays.
*/
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