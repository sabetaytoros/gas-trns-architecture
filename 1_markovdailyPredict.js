/**
 * ============================================================================
 * DOSYA: 1_markovdailyPredict.gs
 * TANIM: Markov öngörüsü ile kırılma, trend devam analizi ve Binary Search tarih tespiti.
 * ============================================================================
 */

/**
 * 1. Adım: I7 ve H7 hücrelerinden tarihleri okuyan, previousObj.indx için Binary Search 
 * yapan, currentObj.indx değerini son eleman olarak sabitleyen ve kronoData'yı 
 * içeren tuple nesnesini döndüren fonksiyon.
 * @returns {Object} { previousObj, currentObj, kronoData }
 */
function getPreviousAndCurrentObjects() {
  var previousObj = {};
  var currentObj = {};
  var kronoData = Trns.chronologicalData || [];
  
  // Tarihleri doğrudan nesne özelliklerine string olarak alıyoruz
  previousObj.date = String(Trns.trs.getRange("I7").getValue()).trim();
  currentObj.date = String(Trns.trs.getRange("H7").getValue()).trim();

  // currentObj.date güncel veriyi temsil ettiği için index doğrudan son eleman olarak atanır
  currentObj.indx = kronoData.length > 0 ? kronoData.length - 1 : -1;

  // Hard-code içermeyen sade eşleşme kontrolü
  var isDateMatch = function(dVal, target) {
    if (!dVal || !target) return false;
    var dStr = String(dVal).trim();
    return dStr.indexOf(target) !== -1;
  };

  // Binary Search ile Previous (I7) indeksini bulma (previousObj.indx güncellenir)
  previousObj.indx = -1;
  var low = 0;
  var high = kronoData.length - 1;

  while (low <= high) {
    var mid = Math.floor((low + high) / 2);
    var dStr = kronoData[mid].date.trim();

    if (isDateMatch(dStr, previousObj.date)) {
      previousObj.indx = mid;
      break;
    }

    if (dStr > previousObj.date) {
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }

  if (previousObj.indx === -1) {
    for (var i = 0; i < kronoData.length; i++) {
      if (isDateMatch(kronoData[i].date, previousObj.date)) {
        previousObj.indx = i;
        break;
      }
    }
  }

  // Kalan özellikleri adım adım nesnelere ekliyoruz
  previousObj.price = Number(Trns.trs.getRange("I9").getValue());
  previousObj.changePct = Number(Trns.trs.getRange("I10").getValue());
  previousObj.direction = Number(Trns.trs.getRange("I10").getValue()) >= 0 ? 1 : -1;

  currentObj.open = Number(Trns.trs.getRange(3, Trns.kbr).getValue());
  currentObj.close = Number(Trns.trs.getRange(17, Trns.kbr).getValue());
  currentObj.reversalDate = null;
  currentObj.reversalValue = 0;
  currentObj.continuationDate = null;
  currentObj.continuationValue = 0;

  return { 
    previousObj: previousObj, 
    currentObj: currentObj, 
    kronoData: kronoData 
  };
}

/**
 * 2. Adım (Parça 1): ChronologicalData verisinin doğrulanması ve konsola loglanması
 */
function logAndValidateChronologicalData(kronoData) {
  if (!kronoData || kronoData.length === 0) {
    Logger.log("[logAndValidateChronologicalData] ⚠️ Uyarı: chronologicalData boş, işlem atlanıyor.");
    return false;
  }
  
  Logger.log("=================================================");
  Logger.log(" 📊 CHRONOLOGICAL DATA İÇERİĞİ (Toplam Bar: " + kronoData.length + ")");
  Logger.log("=================================================");
  for (var idx = 0; idx < Math.min(kronoData.length, 5); idx++) {
    Logger.log("Bar[" + idx + "] -> " + JSON.stringify(kronoData[idx]));
  }
  Logger.log("=================================================");
  return true;
}

/**
 * 2. Adım (Parça 2): Forward test sırasında yükseliş/düşüş trendi tespiti 
 * ve G11/G12 hücrelerine son tarih ile değerin (yükselişte high, düşüşte low) atanması
 * @param {Object} tuple - previousObj, currentObj ve kronoData içeren ana tuple yapısı
 */
function calculateForwardContinuation(tuple) {
  var kronoData = tuple.kronoData || [];
  var currentObj = tuple.currentObj;

  if (kronoData.length > 0) {
    var probVector = [0.25, 0.25, 0.25, 0.25];
    if (Trns.markovProcessor && Trns.chronologicalData) {
      var markovRes = Trns.markovProcessor.processTransitions();
      if (markovRes && markovRes.probVector) probVector = markovRes.probVector;
    }

    var bullWeight = probVector[1] + probVector[3];
    var bearWeight = probVector[0] + probVector[2];

    // Yükseliş Trendi -> High değerine yönlendirilir
    if (currentObj.close >= currentObj.open || bullWeight >= bearWeight) {
      currentObj.continuationValue = Number((currentObj.high * (1 + (bullWeight - 0.5) * 0.05)).toFixed(2));
      currentObj.continuationDate = currentObj.date; // G11 için son tarih
      Logger.log("📈 Yükseliş Trendi Tespit Edildi -> Tarih (G11): " + currentObj.continuationDate + ", Değer/High (G12): " + currentObj.continuationValue);
    } 
    // Düşüş Trendi -> Low değerine yönlendirilir
    else {
      currentObj.continuationValue = Number((currentObj.low * (1 - (bearWeight - 0.5) * 0.05)).toFixed(2));
      currentObj.continuationDate = currentObj.date; // G11 için son tarih
      Logger.log("📉 Düşüş Trendi Tespit Edildi -> Tarih (G11): " + currentObj.continuationDate + ", Değer/Low (G12): " + currentObj.continuationValue);
    }
  }
}

/**
 * 2. Adım (Parça 3): Hesaplanan sonuçların G sütununa (G7:G8, G11:G12) yazılması
 */
function writeResultsToSheet(currentObj) {
  if (currentObj.reversalDate) {
    Trns.trs.getRange("G7").setValue(currentObj.reversalDate);
  }
  if (currentObj.reversalValue) {
    Trns.trs.getRange("G8").setValue(currentObj.reversalValue).setNumberFormat("0.00");
  }

  if (currentObj.continuationDate) {
    Trns.trs.getRange("G11").setValue(currentObj.continuationDate);
  }
  if (currentObj.continuationValue) {
    Trns.trs.getRange("G12").setValue(currentObj.continuationValue).setNumberFormat("0.00");
  }
}

/**
 * 3. Adım, 4. Adım ve 5. Adım: Forward Test, Kırılma Kontrolü ve 10 Günlük Eşik Mantığı
 * @param {Object} tuple - previousObj, currentObj ve kronoData içeren tuple yapısı
 */
function processForwardTestAndBreakCheck(tuple) {
  var kronoData = tuple.kronoData || [];
  if (!kronoData || kronoData.length === 0) {
    Logger.log("[processForwardTestAndBreakCheck] ⚠️ Uyarı: kronoData bulunamadı.");
    return;
  }

  var probVector = [0.25, 0.25, 0.25, 0.25];
  var expLogRatios = { r_oc: 0.002, r_ho: 0.005, r_ol: 0.003 };

  if (Trns.markovProcessor) {
    var markovRes = Trns.markovProcessor.processTransitions();
    if (markovRes) {
      if (markovRes.probVector) probVector = markovRes.probVector;
      if (markovRes.expLogRatios) expLogRatios = markovRes.expLogRatios;
    }
  }

  Logger.log("=======================================================================");
  Logger.log(" 📊 FORWARD TEST VE 10 GÜNLÜK KIRILMA KONTROL ANALİZİ");
  Logger.log("=======================================================================");

  var maxDays = 10;
  var breakOccurred = false;
  var baseClose = tuple.currentObj.close;

  // İleriye dönük (Forward) simülasyon ve kırılma tespiti (Maksimum 10 gün)
  for (var day = 1; day <= Math.min(maxDays, kronoData.length); day++) {
    var bar = kronoData[kronoData.length - day]; // İleriye dönük bar taraması
    if (!bar) continue;

    var prevClose = bar.close;
    var evaluatedPredictions = Trns.predictionEvaluator.evaluateModels(prevClose, probVector, expLogRatios);
    
    // Yön değişimi / Kırılma Kontrolü (Önceki yön ile güncel bar yön uyuşmazlığı)
    var barChange = bar.close - bar.open;
    var barDir = barChange >= 0 ? 1 : -1;

    if (barDir !== tuple.previousObj.direction) {
      breakOccurred = true;
      Logger.log("⚡ [4. Adım] " + day + ". günde kırılma (break) oluştu! Tarih: " + bar.date);
      // 10 günden önce kırılma oluştuğu için kalan günler için hesaplama yapılmaz, döngü sonlandırılır.
      break;
    }

    // 5. Adım: Eğer 10 gün boyunca hiç kırılma oluşmadıysa, tam 10. günde G7, G8, G9 hücrelerine yazma
    if (!breakOccurred && day === maxDays) {
      var targetDate = bar.date;
      var targetValue = bar.close;
      var changePct = ((targetValue - baseClose) / baseClose) * 100;

      if (Trns.trs) {
        Trns.trs.getRange("G7").setValue(targetDate);
        Trns.trs.getRange("G8").setValue(targetValue).setNumberFormat("0.00");
        Trns.trs.getRange("G9").setValue(changePct / 100).setNumberFormat("0.00%");
      }

      Logger.log("✅ [5. Adım] 10 gün boyunca kırılma olmadı! 10. Gün Değerleri Yazıldı -> Tarih (G7): " + targetDate + ", Değer (G8): " + targetValue + ", Artış Yüzdesi (G9): %" + changePct.toFixed(2));
    }
  }

  if (breakOccurred) {
    Logger.log("ℹ️ 10 günden önce kırılma tespit edildiği için G7/G8/G9 10 günlük eşik güncellemeleri atlandı.");
  }
}

/**
 * 2. Adımın Ana Mantığını ve Hesaplamalarını Yürüten Fonksiyon
 * @param {Object} tuple - Previous, Current objelerini ve kronoData'yı içeren tuple yapısı
 * @returns {boolean} İşlem başarılı ise true
 */
function processMarkovCalculationsAndWrites(tuple) {
  var kronoData = tuple.kronoData || [];
  if (!logAndValidateChronologicalData(kronoData)) {
    return false;
  }
  
  // Forward Test sırasında trend devam mantığıyla G11/G12 ataması
  calculateForwardContinuation(tuple);

  // G Sütununa Yazma İşlemleri (G7:G8 ve G11:G12)
  writeResultsToSheet(tuple.currentObj);

  // İzleme Konsolu Logları
  Logger.log("=================================================");
  Logger.log(" 🔍 DEĞİŞİM VE DURUM İZLEME (TUPLE)");
  Logger.log("=================================================");
  Logger.log("Previous Objesi : " + JSON.stringify(tuple.previousObj));
  Logger.log("Current Objesi  : " + JSON.stringify(tuple.currentObj));
  Logger.log("=================================================");

  return true;
}

/**
 * Ana Markov hesaplama ve G sütununa yazma sarmalayıcı fonksiyonu
 */
function calculateAndWriteMarkovPredictions() {
  return newCalculateWriteMarkov();
}

/**
 * Modüler Yeni Markov Hesaplama ve Yazma Fonksiyonu
 */
function newCalculateWriteMarkov() {
  // 1. Adım: Tuple referansı üzerinden Previous, Current objelerini ve kronoData'yı alma
  var tuple = getPreviousAndCurrentObjects();

  // 2. Adım: Hesaplama, doğrulama ve yazma işlemlerini yürüten modüler fonksiyon çağrısı
  var success = processMarkovCalculationsAndWrites(tuple);
  if (!success) {
    return false;
  }

  // 3. Adım & 4. Adım & 5. Adım: Forward Test, 10 günlük kırılma eşiği ve G7/G8/G9 atama mantığı
  processForwardTestAndBreakCheck(tuple);
  
  // Garbage Collection: return true öncesi tuple referansının bellekten silinmesi
  tuple = null;

  return true;
}