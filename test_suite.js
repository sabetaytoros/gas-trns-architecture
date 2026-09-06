/**
 * GEÇİCİ TEST ALANI (SANDBOX)
 * Yeni geliştirilen modüller ve anlık denemeler burada test edilir.
 */
function runSandboxTests() {
  Logger.log("=== GEÇİCİ TEST BAŞLATILDI ===");
  
  try {
    // GELİŞTİRME AŞAMASINDAKİ TEST
    testYeniGelistirme();
    
    Logger.log("🟢 [SANDBOX PASSED] Yeni kod testi geçti!");
  } catch (error) {
    Logger.log("🔴 [SANDBOX FAILED] Hata: " + error.message);
  }
}

function testYeniGelistirme() {
  // Örnek: Marcov veya yeni eklenen bir fonksiyon denemesi
  var girdi = "YUKSEK";
  var sonuc = MarcovGlobal.predictNextState(girdi);
  
  if (!sonuc) {
    throw new Error("Tahmin çıktısı alınamadı!");
  }
  Logger.log("Geçici Test Çıktısı: " + JSON.stringify(sonuc));
}
