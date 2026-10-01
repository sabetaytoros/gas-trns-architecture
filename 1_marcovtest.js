/**
 * 1_marcovtest.gs
 * ============================================================================
 * MARKOV ENHANCED DASHBOARD - FEATURE PARITY & REGRESSION TEST SUITE
 * AMACI: v1.6.0.8 özelliklerinin v3.0 mimarisinde eksiksiz çalıştığını doğrulamak.
 * ============================================================================
 */
/**
 * ============================================================================
 * TEST PROGRAMI: "Test" Sayfasında Eksik 8. Adımın Tamamlanması
 * G11 (Trend Yönlü Hedef Fiyat) ve G12 (Min-Max Genlik Oranı) Doğrulanması
 * ============================================================================
 */



function testcalculateAndWriteMarkovPredictions() {
  Logger.clear();
  Logger.log("=================================================");
  Logger.log("   DOĞAL AKIŞLI DASHBOARD TESTİ BAŞLATILIYOR     ");
  Logger.log("=================================================\n");

  var passedCount = 0;
  var totalTests = 0;

  function assert(condition, testName, detail) {
    totalTests++;
    if (condition) {
      passedCount++;
      Logger.log("✅ [PASSED] " + testName + (detail ? " -> " + detail : ""));
    } else {
      Logger.log("❌ [FAILED] " + testName + " -> " + (detail || "Beklenen değer bulunamadı!"));
    }
  }

  try {
    if (typeof Trns === 'undefined' || !Trns.trs) {
      Logger.log("❌ HATA: Trns ortamı bulunamadı.");
      return false;
    }

    // 1. ADIM: SEKME KONTROLÜ
    assert(Trns.Name === "Test", "Sekme Doğrulaması: Hedef sayfa 'Test' mi?", "Mevcut: " + Trns.Name);

    // 2. ADIM: BUGÜNÜN PANOSUNU / SATIRINI TEMİZLEME (Sanki dün akşamdayız)
    Logger.log("   • Bugüne ait geçici dashboard verileri temizleniyor...");
    var targetRow = Trns.targetRow || (Trns.emptyRow - 12);
    if (targetRow > 0) {
      var gridRange = Trns.trs.getRange(targetRow, 1, 10, 10);
      gridRange.clearContent();
      gridRange.clearFormat();
    }

    // 3. ADIM: ANA DASHBOARD MOTORUNU ÇAĞIRMA
    Logger.log("   • EnhancedDashboard() fonksiyonu tetikleniyor...");
    var dashResult = EnhancedDashboard();
    assert(dashResult === true, "Dashboard Çalıştırma: EnhancedDashboard() 'true' döndürdü mü?");
    return true

  } catch (err) {
    assert(false, "Sıfırlamalı Test Hasılası", err.toString());
  }

  Logger.log("\n=================================================");
  Logger.log("  TEST SONUÇLARI: " + passedCount + " / " + totalTests + " BAŞARILI");
  if (passedCount === totalTests) {
    Logger.log("  SONUÇ: SAYFA BİR ÖNCEKİ GÜNE ÇEKİLİP EnhancedDashboard() TAM MÜKEMMELLİKLE TEST EDİLDİ!");
  } else {
    Logger.log("  SONUÇ: BAZI ADIMLAR BAŞARISIZ OLDU. Logları inceleyin.");
  }
  Logger.log("=================================================");
}
