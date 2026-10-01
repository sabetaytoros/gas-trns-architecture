// markovSelectState.gs
/**
 * Özel konfigürasyon üreten fabrika fonksiyonu
 */
Trns.createCustomConfig = function(overrideParams) {
  const base = Object.assign({}, Trns.CONFIG, overrideParams);
  
  // Bağımlı getter/hesaplamaları dinamik nesneye bağlama
  return Object.assign(base, {
    get REQUIRED_DATA_BARS() {
      return this.MARKOV_BACKWARD_DAYS + this.MACD_WARMUP_BARS;
    },
    get TOTAL_DISPLAY_ROWS() {
      return this.REQUIRED_DATA_BARS + this.CHART_DISPLAY_BARS - 1;
    },
    get MAX_FETCH_ROWS() {
      return this.TOTAL_DISPLAY_ROWS + 1;
    }
  });
};
/**
 * ============================================================================
 * SÜRÜM: MARKOV ENHANCED DASHBOARD v4.1.1.0
 * DOSYA: markovEnhancedDashboard.gs
 * MİMARİ: Top-Down / Tam Modüler Akış Yöneticisi & GC Memory Optimization
 * ============================================================================
 */

const UNIVERSAL_BORDER_COLOR = "#5f6368";

/* ============================================================================
 * 1. BÖLÜM: ANA SÜRÜCÜ (DRIVER) VE AKTİF SEKME KONTROLÜ
 * ============================================================================ */

function EnhancedDashboard() {
  return Trns.newdashboardEngine.run();
}

/**
 * ANA DASHBOARD MOTORU (Modüler Akış Yöneticisi)
 */
Trns.newdashboardEngine = (function () {
  return {
    run: function (config) {
      // Dışarıdan config geçilirse onu, yoksa Trns.CONFIG'i kullanır
      const cfg = config || Trns.CONFIG || {};
      
      try {
        // Step 0: Sekme Doğrulaması
        if (!activeTab()) {
          Logger.log("[Trns.dashboardEngine] activeTab doğrulaması başarısız.");
          return false;
        }

        // Step 1: Data Extractor (Veri Çekme & MACD)
        const fetchRows = cfg.MAX_FETCH_ROWS || 76;
        Trns.chronologicalData = Trns.dataExtractor.getChronologicalBars(fetchRows);
        if (!Trns.chronologicalData || Trns.chronologicalData.length === 0) return false;

        // Step 2: Sheet Archiver (Tarih Kontrolü & Arşivleme)
        if (!Trns.sheetArchiver.prepareAndArchive()) return false;

        // Step 3: Markov Processor (Geçiş Matrisi & Olasılıklar)
        const backwardDays = cfg.MARKOV_BACKWARD_DAYS || 22;
        var markovRes = Trns.markovProcessor.processTransitions(backwardDays);
        if (!markovRes) return false;

        // Step 4: Price Extractor & Predictions (Fiyatlar Okunur ve Tahminler İçeride Hesaplanır)
        if (!Trns.priceExtractor.getPrices(markovRes.probVector, markovRes.expLogRatios)) return false;

        // Step 5: Metrics Calculator (WMAPE, MAE & Kazanan Model)
        Trns.metricsData = Trns.metricsCalculator.evaluateAllMetrics(Trns.predictions, Trns.priceData.actual);

        // Step 6 & 7: Grid Builder & Renderer (Tablo & Grafik Yazımı)
        Trns.grid = Trns.dashboardGridBuilder.buildGrid(Trns.predictions, Trns.metricsData);
        Trns.dashboardRenderer.render();
        /*/ Step 8 : SABİT SÜRÜCÜ ÇAĞRISI:
        if (Trns.winningOhlcAndFluctuationEngine) {
          Trns.winningOhlcAndFluctuationEngine.run(0);
        }*/
        return true;
      } catch (err) {
        Logger.log("[Trns.dashboardEngine HATA] " + err.toString());
        return false;
      } finally {
        // Kod hata alsa veya erken çıksa bile bellek sıfırlanır
        Trns.cleanup();
      }
    }
  };
})();

function activeTab() {
  const ftest = true;
  if (typeof Trns === 'undefined' || !Trns.Name) return false;

  // Sayfa adı "Test" ve ftest === true ise exclude engeline takılmadan çalışır
  if (Trns.Name === "Test" && typeof ftest !== 'undefined' && ftest === true) {
    return true;
  }

  // Global exclude dizisini kontrol eder
  if (typeof exclude !== 'undefined' && Array.isArray(exclude) && exclude.indexOf(Trns.Name) !== -1) {
    Logger.log("[activeTab] Geçersiz veya korumalı sekme: " + Trns.Name + ". İşlem iptal edildi.");
    return false;
  }

  return true;
}

/**
 * STATE TEMİZLEME FONKSİYONU (Garbage Collector Desteği)
 */
Trns.cleanup = function () {
  Trns.chronologicalData = null;
  Trns.priceData         = null;
  Trns.predictions       = null;
  Trns.metricsData       = null;
  Trns.grid              = null;
  Logger.log("[Trns.cleanup] Geçici state verileri bellekten temizlendi.");
};

/* ============================================================================
 * 2. BÖLÜM: MODÜLER BİLEŞENLER (OUTSOURCED MODULES)
 * ============================================================================ */

/* ============================================================================
 * MODÜL 1: VERİ ÇEKME VE MACD HESAPLAMA MOTORU (Trns.dataExtractor)
 * ============================================================================ */
Trns.dataExtractor = {
  version: "v1.2.1",

  getChronologicalBars: function (barCount) {
    var cfg = Trns.CONFIG || { FAST_EMA: 12, SLOW_EMA: 26, SIGNAL_EMA: 9 };
    var count = barCount || cfg.MAX_FETCH_ROWS || 76;

    if (!Trns.emptyRow || !Trns.trs) {
      Logger.log("[dataExtractor HATA] Trns.emptyRow veya Trns.trs bulunamadı.");
      return null;
    }

    var startRow = Trns.emptyRow;
    // 6 sütun (Date, Open, High, Low, Close, Volume) verisini okur
    var rawValues = Trns.trs.getRange(startRow, 1, count, 6).getValues();
    var chronologicalData = [];

    // Verileri kronolojik diziye aktarma
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
    // MACD (12, 26, 9) Dinamik EMA Hesaplaması
    var kFast = 2 / (cfg.FAST_EMA + 1);
    var kSlow = 2 / (cfg.SLOW_EMA + 1);
    var kSignal = 2 / (cfg.SIGNAL_EMA + 1);

    // DÜZELTME: EMA ısınmasına ESKİ TARİHTEN (dizinin en altından/sonundan) başlanır
    var emaFast = chronologicalData[totalBars - 1].close;
    var emaSlow = chronologicalData[totalBars - 1].close;
    var signalEma = 0;

    // Sayfaya yazılacak matris dizinin boyutuyla aynı açılır
    var macdOutputGrid = new Array(totalBars);

    // DÜZELTME: Döngü sondan başa (eski tarihten yeni tarihe/103. satıra) doğru çalışır
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

    // Dinamik Sütun Tanımları Üzerinden MACD Verilerini Sayfaya Yazma
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

    // Bugünün verileri/bloğu zaten mevcutsa:
    if (currentDashboardDate && currentDashboardDate.toString() === Trns.dateFormatted.toString()) {
      
      // EĞER TEST SEKMESİNDEYSEK: Eski içerik ve biçimleri temizleyip baştan yazmaya izin ver
      if (isTestSheet) {
        Logger.log("[sheetArchiver] Test sekmesinde bugünün bloğu temizlenip yeniden hazırlanıyor...");
        var targetRange = Trns.trs.getRange(Trns.targetRow, Trns.targetCol, Trns.outputRows, Trns.dashboardCols);
        targetRange.clearContent();
        targetRange.clearFormat();
        return true;
      }

      // Canlı sekmelerde bugünün verisi varsa tekrar yazma
      Logger.log("[sheetArchiver] Bugünün verileri zaten yazılmış, işlem atlandı.");
      return false;
    }

    // Normal Arşivleme Kaydırması (Önceki günü sağa kopyalama)
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

// MODÜL 5: TAHMİN MODELLERİ MOTORU
Trns.predictionEvaluator = {
  version: "v1.0.0",

  evaluateModels: function (prevClose, probVector, expLogRatios) {
    var lgrtiOpen = prevClose;
    var lgrtiClose = lgrtiOpen * Math.exp(expLogRatios.r_oc);
    var lgrtiMax = Math.max(lgrtiOpen, lgrtiClose);
    var lgrtiMin = Math.min(lgrtiOpen, lgrtiClose);

    var lgrtiPred = {
      open: lgrtiOpen, close: lgrtiClose,
      high: lgrtiMax * Math.exp(expLogRatios.r_ho), low: lgrtiMin * Math.exp(-expLogRatios.r_ol)
    };

    var cplaPred = {
      open: prevClose * (1 + (expLogRatios.r_oc * 0.1)),
      close: prevClose * Math.exp(expLogRatios.r_oc * 0.95),
      high: Math.max(prevClose, lgrtiClose) * Math.exp(expLogRatios.r_ho * 1.05),
      low: Math.min(prevClose, lgrtiClose) * Math.exp(-expLogRatios.r_ol * 0.95)
    };

    var cnvxOpen = prevClose * (1 + (probVector[0] - probVector[3]) * 0.001);
    var cnvxClose = prevClose * (1 + (probVector[1] + probVector[3] - 0.5) * 0.03);
    var cnvxMax = Math.max(cnvxOpen, cnvxClose);
    var cnvxMin = Math.min(cnvxOpen, cnvxClose);

    var cnvxPred = {
      open: cnvxOpen, close: cnvxClose,
      high: cnvxMax * (1 + expLogRatios.r_ho * 0.98), low: cnvxMin * (1 - expLogRatios.r_ol * 0.98)
    };

    var bullWeight = probVector[1] + probVector[3];
    var bearWeight = probVector[0] + probVector[2];
    var changePct = (bullWeight - 0.5) * 0.04;
    var mrkvOpen = Number((prevClose * (1 + (changePct * 0.1))).toFixed(2));
    var mrkvClose = Number((mrkvOpen * (1 + changePct)).toFixed(2));
    var mrkvMax = Math.max(mrkvOpen, mrkvClose);
    var mrkvMin = Math.min(mrkvOpen, mrkvClose);

    var mrkvPred = {
      open: mrkvOpen, close: mrkvClose,
      high: Number((mrkvMax * (1 + Math.max(0, (bullWeight - 0.5) * 0.02) + 0.003)).toFixed(2)),
      low: Number((mrkvMin * (1 - Math.max(0, (bearWeight - 0.5) * 0.02) - 0.003)).toFixed(2))
    };

    return { cpla: cplaPred, cnvx: cnvxPred, lgrti: lgrtiPred, mrkv: mrkvPred };
  }
};

// MODÜL 6: METRİK VE MODEL KAZANAN HESAPLAYICI
Trns.metricsCalculator = {
  version: "v1.0.0",

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
    var cplaMetrics  = this.calcMetrics(predictions.cpla, actObj);
    var cnvxMetrics  = this.calcMetrics(predictions.cnvx, actObj);
    var lgrtiMetrics = this.calcMetrics(predictions.lgrti, actObj);
    var mrkvMetrics  = this.calcMetrics(predictions.mrkv, actObj);

    var models = [
      { name: "CPLA v6.4", wmape: cplaMetrics.wmape, colStart: 3 },
      { name: "CNVX v6.3", wmape: cnvxMetrics.wmape, colStart: 5 },
      { name: "LG-RTI v6.2", wmape: lgrtiMetrics.wmape, colStart: 7 },
      { name: "MRKV v5.9", wmape: mrkvMetrics.wmape, colStart: 9 }
    ];
    models.sort(function(a, b) { return a.wmape - b.wmape; });
    return {
      cpla: cplaMetrics, cnvx: cnvxMetrics, lgrti: lgrtiMetrics, mrkv: mrkvMetrics,
      winningModel: models[0]
    };
  }
};

// MODÜL 7: GRID / TABLO MATRİSİ HAZIRLAYICI
Trns.dashboardGridBuilder = {
  version: "v1.0.0",

  buildGrid: function (predictions, metricsData) {
    var VERSION = "MARKOV DASHBOARD v4.1.1.0";
    var winCol = metricsData.winningModel.colStart;
    var winningRow = ["", "", "", "", "", "", "", "", "", ""];
    winningRow[winCol - 1] = "Winning";

    var rO = Trns.targetRow + 3, rH = Trns.targetRow + 4, rL = Trns.targetRow + 5, rC = Trns.targetRow + 6;
    var fActO = "=B" + Trns.emptyRow, fActH = "=C" + Trns.emptyRow, fActL = "=D" + Trns.emptyRow, fActC = "=E" + Trns.emptyRow;
    var fDev = function(col, row) { return "=IF(B" + row + "=0,0,(" + col + row + "-B" + row + ")/B" + row + ")"; };

    var c = predictions.cpla, x = predictions.cnvx, l = predictions.lgrti, m = predictions.mrkv;
    var cm = metricsData.cpla, xm = metricsData.cnvx, lm = metricsData.lgrti, mm = metricsData.mrkv;

    return [
      [VERSION, "", "", "", "", "", "", "", "", ""],
      winningRow,
      [Trns.dateFormatted, "Actual", "CPLA v6.4", "dev", "CNVX v6.3", "dev", "LG-RTI v6.2", "dev", "MRKV v5.9", "dev"],
      ["open",  fActO, c.open.toFixed(2), fDev("C",rO), x.open.toFixed(2), fDev("E",rO), l.open.toFixed(2), fDev("G",rO), m.open.toFixed(2), fDev("I",rO)],
      ["High",  fActH, c.high.toFixed(2), fDev("C",rH), x.high.toFixed(2), fDev("E",rH), l.high.toFixed(2), fDev("G",rH), m.high.toFixed(2), fDev("I",rH)],
      ["Low",   fActL, c.low.toFixed(2),  fDev("C",rL), x.low.toFixed(2),  fDev("E",rL), l.low.toFixed(2),  fDev("G",rL), m.low.toFixed(2),  fDev("I",rL)],
      ["Close", fActC, c.close.toFixed(2),fDev("C",rC), x.close.toFixed(2),fDev("E",rC), l.close.toFixed(2),fDev("G",rC), m.close.toFixed(2),fDev("I",rC)],
      ["22D Backtest WMAPE", "", cm.wmape, "", xm.wmape, "", lm.wmape, "", mm.wmape, ""],
      ["22D Backtest MAE",   "", cm.mae,   "", xm.mae,   "", lm.mae,   "", mm.mae,   ""],
      ["Tahmini Gövde Tipi", "", (c.close>=c.open?"boga":"ayi"), "", (x.close>=x.open?"boga":"ayi"), "", (l.close>=l.open?"boga":"ayi"), "", (m.close>=m.open?"boga":"ayi"), ""]
    ];
  }
};

// MODÜL 8: GÖRSEL SAYFA RENDERER & BİÇİMLENDİRİCİ
/* ============================================================================
 * GRAFİK MOTORU VE RENDERER (Trns.chartEngine & Trns.dashboardRenderer)
 * ============================================================================ */

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
    // Grafik Çizimi (Parametresiz çağrı)
    Trns.chartEngine.renderCharts();
  }
};
function renderWintoK() { 
    // =========================================================================
    // 🎯 DİNAMİK KOPYALAMA KODU (Open, High, Low -> K3:K5 ve Close -> K17)
    // =========================================================================
    // 1. Open, High, Low Değerlerini K3:K5 Hücrelerine Kopyalama
    var src = Trns.trs.getRange(Trns.targetRow + 3, Trns.metricsData.winningModel.colStart, 3, 1); // Open, High, Low
    var dst = Trns.trs.getRange(3, Trns.kbr); // K3 hedef hücresi[cite: 4, 5]
    Logger.log(' %s %s', src.getA1Notation(), dst.getA1Notation())
    Logger.log('values %s', src.getValues())

    copyCell(src, dst);

    // 2. Close Değerini K17 Hücresine Kopyalama
    src = Trns.trs.getRange(Trns.targetRow + 6, Trns.metricsData.winningModel.colStart, 1, 1); // Close satırı[cite: 4, 5]
    dst = Trns.trs.getRange(17, Trns.kbr); // K17 hedef hücresi[cite: 4, 5]
        Logger.log('values %s %s', src.getValues())
    //copyCell(src, dst);  
}
/* ============================================================================
 * GRAFİK MOTORU (Trns.chartEngine)
 * ============================================================================ */

Trns.chartEngine = (function () {
  return {
    version: "v3.0.6.0",

    renderCharts: function () {
      try {
        if (!Trns.trs) return;

        // 1. MEVCUT GRAFİKLERİ TEMİZLE
        var existingCharts = Trns.trs.getCharts();
        for (var i = 0; i < existingCharts.length; i++) {
          Trns.trs.removeChart(existingCharts[i]);
        }

        // 2. SATIR VE SÜTUN HESAPLAMALARI
        var startRow = Trns.emptyRow - 1;
        var numRows = (Trns.CONFIG && Trns.CONFIG.TOTAL_DISPLAY_ROWS) ? Trns.CONFIG.TOTAL_DISPLAY_ROWS : 75;

        var colDate  = 1; // A: Date
        var colOpen  = 2; // B: Open
        var colHigh  = 3; // C: High
        var colLow   = 4; // D: Low
        var colClose = 5; // E: Close

        var colMacd  = 7; // G: MACD Line
        var colSig   = 8; // H: Signal Line
        var colHist  = 9; // I: Histogram

        // 3. TARİH SÜTUNUNU GERÇEK DATE FORMATINDA TUTMA (String yapılmaz)
        Trns.trs.getRange(startRow, colDate, numRows, 1).setNumberFormat("dd.mm.yy");

        // DİNAMİK GRAFİK KONUMU
        var candleChartTargetRow = startRow - numRows;
        var chartTargetCol = 1; 

        // 4. MIN / MAX Y-EKSENİ HESAPLAMASI
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

        // 5. CANDLESTICK GRAFİĞİ (Date Formatı Aktif)
        var candleChartBuilder = Trns.trs.newChart()
          .setChartType(Charts.ChartType.CANDLESTICK)
          .addRange(Trns.trs.getRange(startRow, colDate, numRows, 1))  // 1. Date (A)
          .addRange(Trns.trs.getRange(startRow, colLow, numRows, 1))   // 2. Low (D)
          .addRange(Trns.trs.getRange(startRow, colOpen, numRows, 1))  // 3. Open (B)
          .addRange(Trns.trs.getRange(startRow, colClose, numRows, 1)) // 4. Close (E)
          .addRange(Trns.trs.getRange(startRow, colHigh, numRows, 1))  // 5. High (C)
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

        // 6. MACD COMBO GRAFİĞİ (Date Formatı Aktif)
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
 * 4. BÖLÜM: MARKOV STATE MOTORU (Trns.markovState)
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



/**
 * ============================================================================
 * SÜRÜM: MARKOV ENHANCED DASHBOARD v3.0.0.0
 * DOSYA: markovEnhancedDashboard.gs
 * MİMARİ: Top-Down (Trns Namespace Yapısı)
 * 
 * Modeller:
 * 1. CPLA v6.4  (HMM + Copula / GMM)
 * 2. CNVX v6.3  (Convex Optimization / POCS)
 * 3. LG-RTI v6.2 (Log-Ratio Fitil/Gövde Kısıtlı)
 * 4. MRKV v5.9  (Klasik Ağırlıklı Yüzde Tabanlı)
 * 
 * SIRALAMA:
 *   1. Ana Sürücü Fonksiyon (drvEngineEnhancedDashboard)
 *   2. Trns.dashboardEngine (Dashboard İş Mantığı)
 *   3. Trns.markovState (Markov & Matematiksel Yardımcı Fonksiyonlar)
 * ============================================================================
 */

// Global Trns Namespace Tanımlaması



/* ============================================================================
 * 1. BÖLÜM: ANA SÜRÜCÜ (DRIVER) FONKSİYONU
 * ============================================================================ */
function drvEnhancedDashboard() {
  return Trns.dashboardEngine.run();
}


/* ============================================================================
 * 2. BÖLÜM: DASHBOARD MOTORU (Trns.dashboardEngine)
 * ============================================================================ */
Trns.dashboardEngine = (function () {

  function calcMetrics(predObj, actObj) {
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
  }

  function setNumberFormats(colStarts, format, startRow, height) {
    colStarts.forEach(function(c) {
      Trns.trs.getRange(startRow, c, height, 1).setNumberFormat(format);
    });
  }

  return {
    version: "v3.0.0.0",

    run: function () {
      var VERSION = "MARKOV DASHBOARD v3.0.0.0";

      // INITIALIZATION TARAFINDAN HESAPLANAN SATIR KONTROLÜ
      if (!Trns.emptyRow || !Trns.trs) return null;

      // 1. GERÇEK (ACTUAL) VERİLERİ OKU (Trns.emptyRow Doğrudan Kullanılır)
      var actualDateStr = Trns.trs.getRange(Trns.emptyRow, 1).getValue();
      var dateFormatted = (typeof formatDateCustom === 'function') 
        ? formatDateCustom(actualDateStr) 
        : actualDateStr;

      var YeniOutputSatirSayisi = 10;
      var DASHBOARD_COLS = 10; // A-J Arası (10 Sütun)
      var TARGET_ROW = Trns.emptyRow - (YeniOutputSatirSayisi + 2);
      var TARGET_COL = 1;

      // KURAL: BUGÜNÜN DEĞERİ VAR İSE İŞLEM YAPMA
      var currentDashboardDate = Trns.trs.getRange(TARGET_ROW + 2, 1).getValue();
      if (currentDashboardDate && currentDashboardDate.toString() === dateFormatted.toString()) {
        Logger.log("Bugünün verileri zaten yazılmış, işlem atlandı.");
        return null;
      }

      var currentDashboardRange = Trns.trs.getRange(TARGET_ROW, TARGET_COL, YeniOutputSatirSayisi, DASHBOARD_COLS);

      // KURAL: MEVCUT BLOKLARI SAĞA KAYDIR
      if (Trns.trs.getRange(TARGET_ROW + 2, 1).getValue() !== "") {
        var copyTargetRange = Trns.trs.getRange(TARGET_ROW, TARGET_COL + DASHBOARD_COLS, YeniOutputSatirSayisi, DASHBOARD_COLS);
        currentDashboardRange.copyTo(copyTargetRange);
        currentDashboardRange.copyTo(copyTargetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      }

      // GERÇEK SAYISAL DEĞERLER (Trns.emptyRow Üzerinden)
      var actualOpen  = Number(Trns.trs.getRange(Trns.emptyRow, 2).getValue());
      var actualHigh  = Number(Trns.trs.getRange(Trns.emptyRow, 3).getValue());
      var actualLow   = Number(Trns.trs.getRange(Trns.emptyRow, 4).getValue());
      var actualClose = Number(Trns.trs.getRange(Trns.emptyRow, 5).getValue());

      if (!actualClose || actualClose <= 0) return null;

      var prevClose = Number(Trns.trs.getRange(Trns.emptyRow + 1, 5).getValue()) || actualOpen;

      // 2. HAM VERİLERİ OKU VE BAR OBJELERİNİ OLUŞTUR
      var rawValues = Trns.trs.getRange(Trns.emptyRow, 1, 23, 6).getValues();
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
            open: o, high: h, low: l, close: c,
            r_oc: Math.log(c / o),
            r_ho: Math.max(0, Math.log(h / maxB)),
            r_ol: Math.max(0, Math.log(minB / l)),
            macd: 0,
            signal: 0
          });
        }
      }

      var totalBars = chronologicalData.length;
      if (totalBars === 0) return null;

      // EMA VE MACD HESABI
      var k12 = 2 / (12 + 1);
      var k26 = 2 / (26 + 1);
      var k9  = 2 / (9 + 1);

      var ema12 = chronologicalData[0].close;
      var ema26 = chronologicalData[0].close;
      var signalEma = 0;

      for (var j = 0; j < chronologicalData.length; j++) {
        var closePrice = chronologicalData[j].close;
        ema12 = closePrice * k12 + ema12 * (1 - k12);
        ema26 = closePrice * k26 + ema26 * (1 - k26);
        
        var macdVal = ema12 - ema26;
        signalEma = (j === 0) ? macdVal : (macdVal * k9 + signalEma * (1 - k9));

        chronologicalData[j].macd = macdVal;
        chronologicalData[j].signal = signalEma;
      }

      // 3. MARKOV HESAPLAMALARI
      var allStates = chronologicalData.map(function(bar) {
        return Trns.markovState.calculate4States(bar);
      });

      var currentState = allStates[0] !== undefined ? allStates[0] : 0;
      var initStateVec = Trns.markovState.getInitialStateVector(currentState);
      
      var mat = Trns.markovState.buildTransitionMatrix(allStates, totalBars - 1, Math.min(totalBars - 1, 22));
      var vec = Trns.markovState.multiplyVectorMatrix(initStateVec, mat);

      var stateLogStats = Trns.markovState.computeStateLogRatioStats(chronologicalData, allStates);
      var expLogRatios = Trns.markovState.computeExpectedLogRatios(vec, stateLogStats);

      // --- MODEL TAHMİNLERİ ---
      var lgrtiOpen = prevClose;
      var lgrtiClose = lgrtiOpen * Math.exp(expLogRatios.r_oc);
      var lgrtiMax = Math.max(lgrtiOpen, lgrtiClose);
      var lgrtiMin = Math.min(lgrtiOpen, lgrtiClose);
      var lgrtiPred = {
        open: lgrtiOpen, close: lgrtiClose,
        high: lgrtiMax * Math.exp(expLogRatios.r_ho), low: lgrtiMin * Math.exp(-expLogRatios.r_ol)
      };

      var cplaPred = {
        open: prevClose * (1 + (expLogRatios.r_oc * 0.1)), close: prevClose * Math.exp(expLogRatios.r_oc * 0.95),
        high: Math.max(prevClose, lgrtiClose) * Math.exp(expLogRatios.r_ho * 1.05), low: Math.min(prevClose, lgrtiClose) * Math.exp(-expLogRatios.r_ol * 0.95)
      };

      var cnvxOpen = prevClose * (1 + (vec[0] - vec[3]) * 0.001);
      var cnvxClose = prevClose * (1 + (vec[1] + vec[3] - 0.5) * 0.03);
      var cnvxMax = Math.max(cnvxOpen, cnvxClose);
      var cnvxMin = Math.min(cnvxOpen, cnvxClose);
      var cnvxPred = {
        open: cnvxOpen, close: cnvxClose,
        high: cnvxMax * (1 + expLogRatios.r_ho * 0.98), low: cnvxMin * (1 - expLogRatios.r_ol * 0.98)
      };

      var bullWeight = vec[1] + vec[3]; 
      var bearWeight = vec[0] + vec[2];
      var changePct = (bullWeight - 0.5) * 0.04;
      var mrkvOpen = Number((prevClose * (1 + (changePct * 0.1))).toFixed(2));
      var mrkvClose = Number((mrkvOpen * (1 + changePct)).toFixed(2));
      var mrkvMax = Math.max(mrkvOpen, mrkvClose);
      var mrkvMin = Math.min(mrkvOpen, mrkvClose);
      var mrkvPred = {
        open: mrkvOpen, close: mrkvClose,
        high: Number((mrkvMax * (1 + Math.max(0, (bullWeight - 0.5) * 0.02) + 0.003)).toFixed(2)),
        low:  Number((mrkvMin * (1 - Math.max(0, (bearWeight - 0.5) * 0.02) - 0.003)).toFixed(2))
      };

      // METRİKLER (WMAPE, MAE)
      var actObj = { open: actualOpen, high: actualHigh, low: actualLow, close: actualClose };
      var cplaMetrics  = calcMetrics(cplaPred, actObj);
      var cnvxMetrics  = calcMetrics(cnvxPred, actObj);
      var lgrtiMetrics = calcMetrics(lgrtiPred, actObj);
      var mrkvMetrics  = calcMetrics(mrkvPred, actObj);

      // KAZANAN MODELİ SEÇME
      var models = [
        { name: "CPLA v6.4", wmape: cplaMetrics.wmape, colStart: 3 },
        { name: "CNVX v6.3", wmape: cnvxMetrics.wmape, colStart: 5 },
        { name: "LG-RTI v6.2", wmape: lgrtiMetrics.wmape, colStart: 7 },
        { name: "MRKV v5.9", wmape: mrkvMetrics.wmape, colStart: 9 }
      ];
      models.sort(function(a, b) { return a.wmape - b.wmape; });
      var winningModel = models[0];

      var winCol = winningModel.colStart;
      var winningRow = ["", "", "", "", "", "", "", "", "", ""];
      winningRow[winCol - 1] = "Winning";

      // DİNAMİK FORMÜLLER (Trns.emptyRow Referansı İle)
      var rO = TARGET_ROW + 3, rH = TARGET_ROW + 4, rL = TARGET_ROW + 5, rC = TARGET_ROW + 6; 
      var fActO = "=B" + Trns.emptyRow, fActH = "=C" + Trns.emptyRow, fActL = "=D" + Trns.emptyRow, fActC = "=E" + Trns.emptyRow;

      var fDev = function(col, row) { return "=IF(B" + row + "=0,0,(" + col + row + "-B" + row + ")/B" + row + ")"; };

      var dashboardGrid = [
        [VERSION, "", "", "", "", "", "", "", "", ""],
        winningRow,
        [dateFormatted, "Actual", "CPLA v6.4", "dev", "CNVX v6.3", "dev", "LG-RTI v6.2", "dev", "MRKV v5.9", "dev"],
        ["open",  fActO, cplaPred.open.toFixed(2), fDev("C",rO), cnvxPred.open.toFixed(2), fDev("E",rO), lgrtiPred.open.toFixed(2), fDev("G",rO), mrkvPred.open.toFixed(2), fDev("I",rO)],
        ["High",  fActH, cplaPred.high.toFixed(2), fDev("C",rH), cnvxPred.high.toFixed(2), fDev("E",rH), lgrtiPred.high.toFixed(2), fDev("G",rH), mrkvPred.high.toFixed(2), fDev("I",rH)],
        ["Low",   fActL, cplaPred.low.toFixed(2),  fDev("C",rL), cnvxPred.low.toFixed(2),  fDev("E",rL), lgrtiPred.low.toFixed(2),  fDev("G",rL), mrkvPred.low.toFixed(2),  fDev("I",rL)],
        ["Close", fActC, cplaPred.close.toFixed(2),fDev("C",rC), cnvxPred.close.toFixed(2),fDev("E",rC), lgrtiPred.close.toFixed(2),fDev("G",rC), mrkvPred.close.toFixed(2),fDev("I",rC)],
        ["22D Backtest WMAPE", "", cplaMetrics.wmape, "", cnvxMetrics.wmape, "", lgrtiMetrics.wmape, "", mrkvMetrics.wmape, ""],
        ["22D Backtest MAE",   "", cplaMetrics.mae,   "", cnvxMetrics.mae,   "", lgrtiMetrics.mae,   "", mrkvMetrics.mae,   ""],
        ["Tahmini Gövde Tipi", "", (cplaPred.close>=cplaPred.open?"boga":"ayi"), "", (cnvxPred.close>=cnvxPred.open?"boga":"ayi"), "", (lgrtiPred.close>=lgrtiPred.open?"boga":"ayi"), "", (mrkvPred.close>=mrkvPred.open?"boga":"ayi"), ""]
      ];

      // 4. TEMİZLİK VE YENİ VERİYİ YAZDIRMA
      currentDashboardRange.clearContent();
      currentDashboardRange.clearFormat();
      currentDashboardRange.clearNote();
      currentDashboardRange.setBorder(false, false, false, false, false, false); 

      currentDashboardRange.setValues(dashboardGrid);
      currentDashboardRange.setHorizontalAlignment("right");
      Trns.trs.getRange(TARGET_ROW, 1, YeniOutputSatirSayisi, 1).setHorizontalAlignment("left");

      // ONDALIK FORMATLAMALAR
      setNumberFormats([4, 6, 8, 10], "0.0000", TARGET_ROW + 3, 4); // Dev Sütunları
      setNumberFormats([3, 5, 7, 9],  "0.0000", TARGET_ROW + 7, 2); // Metrik Sütunları

      // DİNAMİK KAZANAN MODEL ÇERÇEVESİ
      var winRange = Trns.trs.getRange(TARGET_ROW + 1, winCol, 9, 2);
      winRange.setBorder(true, true, true, true, false, false, UNIVERSAL_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
      return dashboardGrid;
    }
  };
})();


/* ============================================================================
 * 3. BÖLÜM: MARKOV STATE VE MATEMATİK MOTORU (Trns.markovState)
 * ============================================================================ */
Trns.markovState = (function () {

  return {
    version: "v3.0.0.0",

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
      var targetMatrix = (typeof marcovState !== 'undefined') 
        ? marcovState 
        : (Trns.marcovState || [[0,0,0,0],[0,0,0,0],[0,0,0,0],[0,0,0,0]]);
      
      for (var r = 0; r < 4; r++) {
        for (var c = 0; c < 4; c++) {
          targetMatrix[r][c] = 0;
        }
      }

      var startIdx = Math.max(0, endIdx - windowSize);
      for (var i = startIdx; i < endIdx; i++) {
        var st = states[i], nst = states[i + 1];
        if (st !== undefined && nst !== undefined) targetMatrix[st][nst]++;
      }

      for (var r = 0; r < 4; r++) {
        var rowSum = targetMatrix[r].reduce(function(a, b) { return a + b; }, 0);
        if (rowSum > 0) {
          for (var c = 0; c < 4; c++) targetMatrix[r][c] /= rowSum;
        } else {
          for (var c = 0; c < 4; c++) targetMatrix[r][c] = 0.25;
        }
      }

      return targetMatrix;
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
})();