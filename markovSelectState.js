// markovSelectState.gs
/**
 * Özel konfigürasyon üreten fabrika fonksiyonu
 */
Trns.createCustomConfig = function(overrideParams) {
  const base = Object.assign({}, Trns.CONFIG, overrideParams);
  return Object.assign(base, {
    get REQUIRED_DATA_BARS() { return this.MARKOV_BACKWARD_DAYS + this.MACD_WARMUP_BARS; },
    get TOTAL_DISPLAY_ROWS() { return this.REQUIRED_DATA_BARS + this.CHART_DISPLAY_BARS - 1; },
    get MAX_FETCH_ROWS() { return this.TOTAL_DISPLAY_ROWS + 1; }
  });
};

/**
 * ============================================================================
 * SÜRÜM: MARKOV ENHANCED DASHBOARD v4.1.4.0
 * DOSYA: markovSelectState.gs
 * MİMARİ: Top-Down / Tam Modüler Akış Yöneticisi & GC Memory Optimization
 * ============================================================================
 */
const UNIVERSAL_BORDER_COLOR = "#5f6368";


/* ============================================================================
    1. BÖLÜM: ANA SÜRÜCÜ (DRIVER) VE AKTİF SEKME KONTROLÜ
 * ============================================================================ */
function EnhancedDashboard() {
  return Trns.newdashboardEngine.run();
}

/**
 * ANA DASHBOARD MOTORU (Modüler Akış Yöneticisi)
 */
const ftest = true;
Trns.newdashboardEngine = (function () {
  return {
    run: function (config) {
      const cfg = config || Trns.CONFIG || {};
      try {
        if (!activeTab()) {
          Logger.log("[Trns.dashboardEngine] activeTab doğrulaması başarısız.");
          return false;
        }

        const fetchRows = cfg.MAX_FETCH_ROWS || 76;
        Trns.chronologicalData = Trns.dataExtractor.getChronologicalBars(fetchRows);
        if (!Trns.chronologicalData || Trns.chronologicalData.length === 0) return false;

        if (!Trns.sheetArchiver.prepareAndArchive()) return false;

        const backwardDays = cfg.MARKOV_BACKWARD_DAYS || 22;
        var markovRes = Trns.markovProcessor.processTransitions(backwardDays);
        if (!markovRes) return false;

        if (!Trns.priceExtractor.getPrices(markovRes.probVector, markovRes.expLogRatios)) return false;

        Trns.metricsData = Trns.metricsCalculator.evaluateAllMetrics(Trns.predictions, Trns.priceData.actual);

        Trns.grid = Trns.dashboardGridBuilder.buildGrid(Trns.predictions, Trns.metricsData);
        Trns.dashboardRenderer.render();
        // =========================================================================
        // 🎯 MARKOV GÜNLÜK TAHMİN BİLEŞENİ (G7:G12)
        // =========================================================================
        if (ftest && typeof calculateAndWriteMarkovPredictions === 'function') {
          calculateAndWriteMarkovPredictions();
        }
        return true;
      } catch (err) {
        Logger.log("[Trns.dashboardEngine HATA] " + err.toString());
        return false;
      } finally {
        Trns.cleanup();
      }
    }
  };
})();

function activeTab() {

  if (typeof Trns === 'undefined' || !Trns.Name) return false;
  if (Trns.Name === "Test" && typeof ftest !== 'undefined' && ftest === true) {
    return true;
  }
  if (typeof exclude !== 'undefined' && Array.isArray(exclude) && exclude.indexOf(Trns.Name) !== -1) {
    Logger.log("[activeTab] Geçersiz veya korumalı sekme: " + Trns.Name + ". İşlem iptal edildi.");
    return false;
  }
  return true;
}

Trns.cleanup = function () {
  Trns.chronologicalData = null;
  Trns.priceData         = null;
  Trns.predictions       = null;
  Trns.metricsData       = null;
  Trns.grid              = null;
  Logger.log("[Trns.cleanup] Geçici state verileri bellekten temizlendi.");
};

/* ============================================================================
    2. BÖLÜM: MODÜLER BİLEŞENLER
 * ============================================================================ */

// MODÜL 1: VERİ ÇEKME VE MACD HESAPLAMA MOTORU
Trns.dataExtractor = {
  version: "v1.2.3",
  getChronologicalBars: function (barCount) {
    var cfg = Trns.CONFIG || { FAST_EMA: 12, SLOW_EMA: 26, SIGNAL_EMA: 9 };
    var count = barCount || cfg.MAX_FETCH_ROWS || 76;

    if (!Trns.emptyRow || !Trns.trs) {
      Logger.log("[dataExtractor HATA] Trns.emptyRow veya Trns.trs bulunamadı.");
      return null;
    }

    var startRow = Trns.emptyRow + 1;
    var rawValues = Trns.trs.getRange(startRow, 1, count, 6).getValues();
    var chronologicalData = [];

    for (var i = 0; i < rawValues.length; i++) {
      var r = rawValues[i];
      var c = Number(r[4]);
      if (r[0] && !isNaN(c) && c > 0) {
        var o = Number(r[1]) || c;
        var h = Number(r[2]) || c;
        var l = Number(r[3]) || c;
        var maxB = Math.max(o, c);
        var minB = Math.min(o, c);

        chronologicalData.push({
          date: r[0],
          open: o, high: h, low: l, close: c,
          r_oc: Math.log(c / o),
          r_ho: Math.max(0, Math.log(h / maxB)),
          r_ol: Math.max(0, Math.log(minB / l)),
          macd: 0, signal: 0, hist: 0
        });
      }
    }

    var totalBars = chronologicalData.length;
    if (totalBars === 0) return null;

    var kFast   = 2 / (cfg.FAST_EMA + 1);
    var kSlow   = 2 / (cfg.SLOW_EMA + 1);
    var kSignal = 2 / (cfg.SIGNAL_EMA + 1);

    var emaFast = chronologicalData[totalBars - 1].close;
    var emaSlow = chronologicalData[totalBars - 1].close;
    var signalEma = 0;

    var macdOutputGrid = new Array(totalBars);

    for (var j = totalBars - 1; j >= 0; j--) {
      var closePrice = chronologicalData[j].close;
      emaFast = closePrice * kFast + emaFast * (1 - kFast);
      emaSlow = closePrice * kSlow + emaSlow * (1 - kSlow);

      var macdVal = emaFast - emaSlow;
      signalEma = (j === totalBars - 1) ? macdVal : (macdVal * kSignal + signalEma * (1 - kSignal));
      var histVal = macdVal - signalEma;

      chronologicalData[j].macd = macdVal;
      chronologicalData[j].signal = signalEma;
      chronologicalData[j].hist = histVal;

      macdOutputGrid[j] = [macdVal, signalEma, histVal];
    }

    Trns.trs.getRange(startRow, Trns.INDICATOR_COLS.MACD, totalBars, 3)
      .setValues(macdOutputGrid).setNumberFormat("0.0000");

    return chronologicalData;
  }
};

// MODÜL 2: SAYFA VE ARŞİV YÖNETİCİSİ
Trns.sheetArchiver = {
  version: "v1.1.0",
  prepareAndArchive: function () {
    if (!Trns.emptyRow || !Trns.trs) return false;

    var actualDateStr = Trns.trs.getRange(Trns.emptyRow, 1).getValue();
    Trns.dateFormatted = (typeof formatDateCustom === 'function') 
      ? formatDateCustom(actualDateStr) 
      : actualDateStr;

    Trns.outputRows = 10;
    Trns.dashboardCols = 10;
    Trns.targetRow = Trns.emptyRow - (Trns.outputRows + 2);
    Trns.targetCol = 1;

    var isTestSheet = (Trns.Name === "Test");
    var currentDashboardDate = Trns.trs.getRange(Trns.targetRow + 2, 1).getValue();

    if (currentDashboardDate && currentDashboardDate.toString() === Trns.dateFormatted.toString()) {
      if (isTestSheet) {
        Logger.log("[sheetArchiver] Test sekmesinde bugünün bloğu temizlenip yeniden hazırlanıyor...");
        var targetRange = Trns.trs.getRange(Trns.targetRow, Trns.targetCol, Trns.outputRows, Trns.dashboardCols);
        targetRange.clearContent();
        targetRange.clearFormat();
        return true;
      }
      Logger.log("[sheetArchiver] Bugünün verileri zaten yazılmış, işlem atlandı.");
      return false;
    }

    var currentDashboardRange = Trns.trs.getRange(Trns.targetRow, Trns.targetCol, Trns.outputRows, Trns.dashboardCols);

    if (Trns.trs.getRange(Trns.targetRow + 2, 1).getValue() !== "") {
      var copyTargetRange = Trns.trs.getRange(Trns.targetRow, Trns.targetCol + Trns.dashboardCols, Trns.outputRows, Trns.dashboardCols);
      currentDashboardRange.copyTo(copyTargetRange);
      currentDashboardRange.copyTo(copyTargetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    }

    return true;
  }
};

// MODÜL 3: MARKOV İŞLEMCİSİ VE ANALİZ MOTORU
Trns.markovProcessor = {
  version: "v1.0.1",
  processTransitions: function (windowSize) {
    if (!Trns.chronologicalData || Trns.chronologicalData.length === 0) {
      Logger.log("[markovProcessor HATA] ChronologicalData bulunamadı.");
      return null;
    }

    var winSize = windowSize || (Trns.CONFIG ? Trns.CONFIG.MARKOV_BACKWARD_DAYS : 22);
    var totalBars = Trns.chronologicalData.length;

    var allStates = Trns.chronologicalData.map(function (bar) {
      return Trns.markovState.calculate4States(bar);
    });

    var currentState = allStates[0] !== undefined ? allStates[0] : 0;
    var initStateVec = Trns.markovState.getInitialStateVector(currentState);

    var mat = Trns.markovState.buildTransitionMatrix(allStates, totalBars - 1, Math.min(totalBars - 1, winSize));
    var probVector = Trns.markovState.multiplyVectorMatrix(initStateVec, mat);

    var stateLogStats = Trns.markovState.computeStateLogRatioStats(Trns.chronologicalData, allStates);
    var expLogRatios = Trns.markovState.computeExpectedLogRatios(probVector, stateLogStats);

    return {
      allStates: allStates,
      probVector: probVector,
      expLogRatios: expLogRatios
    };
  }
};

// MODÜL 4: FİYAT ÇEKME VE TAHMİN HAZIRLAMA MOTORU
Trns.priceExtractor = {
  version: "v1.1.0",
  getPrices: function (probVector, expLogRatios) {
    if (!Trns.emptyRow || !Trns.trs) return null;

    var actualOpen  = Number(Trns.trs.getRange(Trns.emptyRow, 2).getValue());
    var actualHigh  = Number(Trns.trs.getRange(Trns.emptyRow, 3).getValue());
    var actualLow   = Number(Trns.trs.getRange(Trns.emptyRow, 4).getValue());
    var actualClose = Number(Trns.trs.getRange(Trns.emptyRow, 5).getValue());

    if (!actualClose || actualClose <= 0) return null;

    var prevClose = Number(Trns.trs.getRange(Trns.emptyRow + 1, 5).getValue()) || actualOpen;

    Trns.priceData = {
      actual: {
        open: actualOpen,
        high: actualHigh,
        low: actualLow,
        close: actualClose
      },
      prevClose: prevClose
    };

    if (probVector && expLogRatios && Trns.predictionEvaluator) {
      Trns.predictions = Trns.predictionEvaluator.evaluateModels(
        prevClose, 
        probVector, 
        expLogRatios
      );
    }

    return Trns.priceData;
  }
};

// MODÜL 5: TAHMİN MODELLERİ MOTORU (OBJE / MAP YAPISI)
Trns.predictionEvaluator = {
  version: "v2.1.0",

  models: {
    lgrti: {
      id: "lgrti",
      name: "LG-RTI v6.2",
      colStart: 7,
      calculate: function (prevClose, probVector, expLogRatios) {
        var open = prevClose;
        var close = open * Math.exp(expLogRatios.r_oc);
        var maxVal = Math.max(open, close);
        var minVal = Math.min(open, close);
        return {
          open: open,
          close: close,
          high: maxVal * Math.exp(expLogRatios.r_ho),
          low: minVal * Math.exp(-expLogRatios.r_ol)
        };
      }
    },
    cpla: {
      id: "cpla",
      name: "CPLA v6.4",
      colStart: 3,
      calculate: function (prevClose, probVector, expLogRatios, lgrtiClose) {
        var open = prevClose * (1 + (expLogRatios.r_oc * 0.1));
        var close = prevClose * Math.exp(expLogRatios.r_oc * 0.95);
        return {
          open: open,
          close: close,
          high: Math.max(prevClose, lgrtiClose) * Math.exp(expLogRatios.r_ho * 1.05),
          low: Math.min(prevClose, lgrtiClose) * Math.exp(-expLogRatios.r_ol * 0.95)
        };
      }
    },
    cnvx: {
      id: "cnvx",
      name: "CNVX v6.3",
      colStart: 5,
      calculate: function (prevClose, probVector, expLogRatios) {
        var open = prevClose * (1 + (probVector[0] - probVector[3]) * 0.001);
        var close = prevClose * (1 + (probVector[1] + probVector[3] - 0.5) * 0.03);
        var maxVal = Math.max(open, close);
        var minVal = Math.min(open, close);
        return {
          open: open,
          close: close,
          high: maxVal * (1 + expLogRatios.r_ho * 0.98),
          low: minVal * (1 - expLogRatios.r_ol * 0.98)
        };
      }
    },
    mrkv: {
      id: "mrkv",
      name: "MRKV v5.9",
      colStart: 9,
      calculate: function (prevClose, probVector, expLogRatios) {
        var bullWeight = probVector[1] + probVector[3];
        var bearWeight = probVector[0] + probVector[2];
        var changePct = (bullWeight - 0.5) * 0.04;
        var open = Number((prevClose * (1 + (changePct * 0.1))).toFixed(2));
        var close = Number((open * (1 + changePct)).toFixed(2));
        var maxVal = Math.max(open, close);
        var minVal = Math.min(open, close);
        return {
          open: open,
          close: close,
          high: Number((maxVal * (1 + Math.max(0, (bullWeight - 0.5) * 0.02) + 0.003)).toFixed(2)),
          low: Number((minVal * (1 - Math.max(0, (bearWeight - 0.5) * 0.02) - 0.003)).toFixed(2))
        };
      }
    }
  },

  evaluateModels: function (prevClose, probVector, expLogRatios) {
    var predictions = {};
    
    predictions.lgrti = this.models.lgrti.calculate(prevClose, probVector, expLogRatios);
    predictions.cpla  = this.models.cpla.calculate(prevClose, probVector, expLogRatios, predictions.lgrti.close);
    predictions.cnvx  = this.models.cnvx.calculate(prevClose, probVector, expLogRatios);
    predictions.mrkv  = this.models.mrkv.calculate(prevClose, probVector, expLogRatios);

    return predictions;
  }
};

// MODÜL 6: METRİK VE MODEL KAZANAN HESAPLAYICI
Trns.metricsCalculator = {
  version: "v2.1.0",

  calcMetrics: function (predObj, actObj) {
    var relErrors = [
      Math.abs((predObj.open - actObj.open) / actObj.open),
      Math.abs((predObj.high - actObj.high) / actObj.high),
      Math.abs((predObj.low - actObj.low) / actObj.low),
      Math.abs((predObj.close - actObj.close) / actObj.close)
    ];
    var absErrors = [
      Math.abs(predObj.open - actObj.open),
      Math.abs(predObj.high - actObj.high),
      Math.abs(predObj.low - actObj.low),
      Math.abs(predObj.close - actObj.close)
    ];
    var actSums = actObj.open + actObj.high + actObj.low + actObj.close;
    
    return {
      mae: Number((relErrors.reduce(function(a, b) { return a + b; }, 0) / 4).toFixed(4)),
      wmape: actSums > 0 ? Number((absErrors.reduce(function(a, b) { return a + b; }, 0) / actSums).toFixed(4)) : 0
    };
  },

  evaluateAllMetrics: function (predictions, actObj) {
    var self = this;
    var metricsResult = {};
    var modelsList = [];
    var modelsObj = Trns.predictionEvaluator.models;

    Object.keys(modelsObj).forEach(function (key) {
      var model = modelsObj[key];
      var pred = predictions[key];
      var mData = self.calcMetrics(pred, actObj);
      metricsResult[key] = mData;

      modelsList.push({
        id: model.id,
        name: model.name,
        wmape: mData.wmape,
        colStart: model.colStart
      });
    });

    modelsList.sort(function(a, b) { return a.wmape - b.wmape; });
    metricsResult.winningModel = modelsList[0];

    return metricsResult;
  }
};

// MODÜL 7: GRID / TABLO MATRİSİ HAZIRLAYICI
Trns.dashboardGridBuilder = {
  version: "v2.2.0",

  buildGrid: function (predictions, metricsData) {
    var winId = metricsData.winningModel.id;
    var dikmtr = transposeToColumn(predictions[winId]);
    Trns.trs.getRange(3, Trns.kbr, 3, 1).setValues(dikmtr.slice(0, 3)).setNumberFormat("0.00");
    Trns.trs.getRange(17, Trns.kbr, 1, 1).setValue(predictions[winId].close).setNumberFormat("0.00");

    var VERSION = "MARKOV DASHBOARD v4.1.1.0";
    var winCol = metricsData.winningModel.colStart;
    var winningRow = ["", "", "", "", "", "", "", "", "", ""];
    winningRow[winCol - 1] = "Winning";

    var rO = Trns.targetRow + 3, rH = Trns.targetRow + 4, rL = Trns.targetRow + 5, rC = Trns.targetRow + 6;
    var fActO = "=B" + Trns.emptyRow, fActH = "=C" + Trns.emptyRow, fActL = "=D" + Trns.emptyRow, fActC = "=E" + Trns.emptyRow;
    var fDev = function(col, row) { return "=IF(B" + row + "=0,0,(" + col + row + "-B" + row + ")/B" + row + ")"; };

    var headerRow = [Trns.dateFormatted, "Actual"];
    var openRow   = ["open", fActO];
    var highRow   = ["High", fActH];
    var lowRow    = ["Low", fActL];
    var closeRow  = ["Close", fActC];
    var wmapeRow  = ["22D Backtest WMAPE", ""];
    var maeRow    = ["22D Backtest MAE", ""];
    var bodyRow   = ["Tahmini Gövde Tipi", ""];

    var modelsObj = Trns.predictionEvaluator.models;

    Object.keys(modelsObj).forEach(function (key, idx) {
      var model = modelsObj[key];
      var p = predictions[key];
      var m = metricsData[key];
      
      var valColLetter = String.fromCharCode(67 + (idx * 2));

      headerRow.push(model.name, "dev");
      openRow.push(p.open.toFixed(2), fDev(valColLetter, rO));
      highRow.push(p.high.toFixed(2), fDev(valColLetter, rH));
      lowRow.push(p.low.toFixed(2), fDev(valColLetter, rL));
      closeRow.push(p.close.toFixed(2), fDev(valColLetter, rC));
      
      wmapeRow.push(m.wmape, "");
      maeRow.push(m.mae, "");
      bodyRow.push((p.close >= p.open ? "boga" : "ayi"), "");
    });

    return [
      [VERSION, "", "", "", "", "", "", "", "", ""],
      winningRow,
      headerRow,
      openRow,
      highRow,
      lowRow,
      closeRow,
      wmapeRow,
      maeRow,
      bodyRow
    ];
  }
};

// MODÜL 8: GÖRSEL SAYFA RENDERER & BİÇİMLENDİRİCİ
Trns.dashboardRenderer = {
  version: "v1.1.2",
  render: function () {
    if (!Trns.grid || !Trns.metricsData || !Trns.metricsData.winningModel) {
      Logger.log("[dashboardRenderer HATA] Render için gerekli Trns.grid veya Trns.metricsData verileri eksik.");
      return;
    }

    var winningModel = Trns.metricsData.winningModel;
    var range = Trns.trs.getRange(Trns.targetRow, Trns.targetCol, Trns.outputRows, Trns.dashboardCols);

    range.clearContent();
    range.clearFormat();
    range.clearNote();
    range.setBorder(false, false, false, false, false, false);

    range.setValues(Trns.grid);
    range.setHorizontalAlignment("right");
    Trns.trs.getRange(Trns.targetRow, 1, Trns.outputRows, 1).setHorizontalAlignment("left");

    Trns.sheetFormatter.setNumberFormats([4, 6, 8, 10], "0.0000", Trns.targetRow + 3, 4);
    Trns.sheetFormatter.setNumberFormats([3, 5, 7, 9], "0.0000", Trns.targetRow + 7, 2);

    var winRange = Trns.trs.getRange(Trns.targetRow + 1, winningModel.colStart, 9, 2);
    winRange.setBorder(true, true, true, true, false, false, UNIVERSAL_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);

    Trns.chartEngine.renderCharts();
  }
};

/* ============================================================================
 * GRAFİK MOTORU (Trns.chartEngine)
 * ============================================================================ */
Trns.chartEngine = (function () {
  return {
    version: "v3.0.7.0",
    renderCharts: function () {
      try {
        if (!Trns.trs) return;

        var existingCharts = Trns.trs.getCharts();
        for (var i = 0; i < existingCharts.length; i++) {
          Trns.trs.removeChart(existingCharts[i]);
        }

        var startRow = Trns.emptyRow + 1; 
        var numRows = (Trns.CONFIG && Trns.CONFIG.TOTAL_DISPLAY_ROWS) ? Trns.CONFIG.TOTAL_DISPLAY_ROWS : 75;

        var colDate  = 1;
        var colOpen  = 2;
        var colHigh  = 3;
        var colLow   = 4;
        var colClose = 5;

        var colMacd  = 7;
        var colSig   = 8;
        var colHist  = 9;

        Trns.trs.getRange(startRow, colDate, numRows, 1).setNumberFormat("dd.mm.yy");

        var candleChartTargetRow = Math.max(1, startRow - numRows - 10);
        var chartTargetCol = 1; 

        var chronologicalData = Trns.chronologicalData;
        var minLow = Infinity;
        var maxHigh = -Infinity;

        if (chronologicalData && chronologicalData.length > 0) {
          var count = Math.min(chronologicalData.length, numRows);
          for (var k = 0; k < count; k++) {
            var b = chronologicalData[k];
            if (b && typeof b.low === 'number' && b.low < minLow)     minLow = b.low;
            if (b && typeof b.high === 'number' && b.high > maxHigh)  maxHigh = b.high;
          }
        }

        var hasValidBounds = (minLow !== Infinity && maxHigh !== -Infinity && minLow < maxHigh);
        var yAxisOptions = {};
        if (hasValidBounds) {
          var margin = (maxHigh - minLow) * 0.01;
          yAxisOptions = {
            viewWindow: {
              min: Math.floor(minLow - margin),
              max: Math.ceil(maxHigh + margin)
            }
          };
        }

        var CHART_WIDTH = 1150;
        var CHART_HEIGHT = 380;

        var candleChartBuilder = Trns.trs.newChart()
          .setChartType(Charts.ChartType.CANDLESTICK)
          .addRange(Trns.trs.getRange(startRow, colDate, numRows, 1))
          .addRange(Trns.trs.getRange(startRow, colLow, numRows, 1))
          .addRange(Trns.trs.getRange(startRow, colOpen, numRows, 1))
          .addRange(Trns.trs.getRange(startRow, colClose, numRows, 1))
          .addRange(Trns.trs.getRange(startRow, colHigh, numRows, 1))
          .setPosition(candleChartTargetRow, chartTargetCol, 0, 0)
          .setOption('title', 'OHLC Fiyat Hareketi')
          .setOption('legend', { position: 'none' })
          .setOption('useFirstColumnAsDomain', true)
          .setOption('hAxis', {
            direction: -1,
            format: 'dd.mm.yy',
            slantedText: true,
            slantedTextAngle: 90
          })
          .setOption('width', CHART_WIDTH)
          .setOption('height', CHART_HEIGHT);

        if (hasValidBounds) {
          candleChartBuilder.setOption('vAxis', yAxisOptions);
        }

        Trns.trs.insertChart(candleChartBuilder.build());

        var macdChartTargetRow = candleChartTargetRow + 20;

        var macdChart = Trns.trs.newChart()
          .setChartType(Charts.ChartType.COMBO)
          .addRange(Trns.trs.getRange(startRow, colDate, numRows, 1))
          .addRange(Trns.trs.getRange(startRow, colMacd, numRows, 1))
          .addRange(Trns.trs.getRange(startRow, colSig, numRows, 1))
          .addRange(Trns.trs.getRange(startRow, colHist, numRows, 1))
          .setPosition(macdChartTargetRow, chartTargetCol, 0, 0)
          .setOption('title', 'MACD (12, 26, 9) Göstergesi')
          .setOption('useFirstColumnAsDomain', true)
          .setOption('hAxis', {
            direction: -1,
            format: 'dd.mm.yy',
            slantedText: true,
            slantedTextAngle: 90
          })
          .setOption('series', {
            0: { type: 'line', color: '#1a73e8' },
            1: { type: 'line', color: '#d93025' },
            2: { type: 'bars', color: '#34a853' }
          })
          .setOption('width', CHART_WIDTH)
          .setOption('height', CHART_HEIGHT)
          .build();

        Trns.trs.insertChart(macdChart);

      } catch (err) {
        Logger.log("[chartEngine HATA] " + err.toString());
      }
    }
  };
})();

Trns.sheetFormatter = {
  version: "v1.0.0",
  setNumberFormats: function (colStarts, format, startRow, height) {
    colStarts.forEach(function(c) {
      Trns.trs.getRange(startRow, c, height, 1).setNumberFormat(format);
    });
  }
};

/* ============================================================================
    4. BÖLÜM: MARKOV STATE MOTORU
 * ============================================================================ */
Trns.markovState = {
  version: "v3.0.0.4",
  calculate4States: function (bar) {
    var isUp = bar.close >= bar.open;
    var isBullishMACD = bar.macd >= bar.signal;
    if (!isUp && !isBullishMACD) return 0;
    if (!isUp && isBullishMACD)  return 1;
    if (isUp && !isBullishMACD)  return 2;
    return 3;
  },
  getInitialStateVector: function (state) {
    var vec = [0, 0, 0, 0];
    if (state >= 0 && state <= 3) vec[state] = 1;
    else vec[0] = 1;
    return vec;
  },
  buildTransitionMatrix: function (states, endIdx, windowSize) {
    var targetMatrix = [
      [0, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0]
    ];
    var startIdx = Math.max(0, endIdx - windowSize);

    for (var i = startIdx; i < endIdx; i++) {
      var st = states[i], nst = states[i + 1];
      if (st !== undefined && nst !== undefined) {
        targetMatrix[st][nst]++;
      }
    }

    var smoothedMatrix = [
      [0, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0]
    ];

    for (var r = 0; r < 4; r++) {
      var rowSum = targetMatrix[r].reduce(function(a, b) { return a + b; }, 0);
      for (var c = 0; c < 4; c++) {
        smoothedMatrix[r][c] = (targetMatrix[r][c] + 1) / (rowSum + 4);
      }
    }

    return smoothedMatrix;
  },
  multiplyVectorMatrix: function (vec, mat) {
    var result = [0, 0, 0, 0];
    for (var c = 0; c < 4; c++) {
      for (var r = 0; r < 4; r++) result[c] += vec[r] * mat[r][c];
    }
    return result;
  },
  computeStateLogRatioStats: function (bars, states) {
    var stats = {
      0: { r_oc: [], r_ho: [], r_ol: [] },
      1: { r_oc: [], r_ho: [], r_ol: [] },
      2: { r_oc: [], r_ho: [], r_ol: [] },
      3: { r_oc: [], r_ho: [], r_ol: [] }
    };

    for (var i = 0; i < bars.length; i++) {
      var st = states[i];
      if (stats[st]) {
        stats[st].r_oc.push(bars[i].r_oc);
        stats[st].r_ho.push(bars[i].r_ho);
        stats[st].r_ol.push(bars[i].r_ol);
      }
    }

    var avgStats = {};
    for (var s = 0; s < 4; s++) {
      var oc = stats[s].r_oc, ho = stats[s].r_ho, ol = stats[s].r_ol;
      avgStats[s] = {
        r_oc: oc.length > 0 ? (oc.reduce(function(a,b){ return a+b; }, 0) / oc.length) : 0,
        r_ho: ho.length > 0 ? Math.max(0.001, ho.reduce(function(a,b){ return a+b; }, 0) / ho.length) : 0.002,
        r_ol: ol.length > 0 ? Math.max(0.001, ol.reduce(function(a,b){ return a+b; }, 0) / ol.length) : 0.002
      };
    }
    return avgStats;
  },
  computeExpectedLogRatios: function (probVec, stateStats) {
    var exp_r_oc = 0, exp_r_ho = 0, exp_r_ol = 0;
    for (var s = 0; s < 4; s++) {
      var p = probVec[s];
      exp_r_oc += p * stateStats[s].r_oc;
      exp_r_ho += p * stateStats[s].r_ho;
      exp_r_ol += p * stateStats[s].r_ol;
    }
    return { r_oc: exp_r_oc, r_ho: Math.max(0, exp_r_ho), r_ol: Math.max(0, exp_r_ol) };
  }
};