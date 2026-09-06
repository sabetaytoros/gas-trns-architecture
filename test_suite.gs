/**
 * GENEL TEST PAKETİ (REGRESSION & INTEGRATION SUITE)
 * Yeni bir versiyon canlıya alınmadan önce TÜM testler buradan geçer.
 */
function runFullTestSuite() {
  Logger.log("==================================================");
  Logger.log("=== TAM SİSTEM TESTİ BAŞLATILDI (TÜM MODÜLLER) ===");
  Logger.log("==================================================");
  
  var basariliTestSayisi = 0;
  var toplamTestSayisi = 0;
  
  var testler = [
    { ad: "Marcov Matris Olasılık Toplamları", fn: testMarcovOlasiliklar },
    { ad: "Marcov Kare Matris Doğrulaması", fn: testMarcovMatrisBoyut },
    { ad: "Marcov Tahmin Mekanizması", fn: testMarcovTahmin },
    { ad: "Geçersiz Veri Yönetimi", fn: testGecersizVeri }
  ];
  
  for (var i = 0; i < testler.length; i++) {
    toplamTestSayisi++;
    try {
      testler[i].fn();
      Logger.log("🟢 [PASS] " + testler[i].ad);
      basariliTestSayisi++;
    } catch (e) {
      Logger.log("🔴 [FAIL] " + testler[i].ad + " -> " + e.message);
    }
  }
  
  Logger.log("--------------------------------------------------");
  Logger.log("SONUÇ: " + basariliTestSayisi + " / " + toplamTestSayisi + " test başarıyla tamamlandı.");
  
  if (basariliTestSayisi !== toplamTestSayisi) {
    throw new Error("Sürüm testi BAŞARISIZ! Bazı testler patladı, push yapmayın.");
  }
}

// GENEL HAVA VE MARCOV TESTLERİ
function testMarcovOlasiliklar() {
  var matrix = MarcovGlobal.getTransitionMatrix();
  for (var i = 0; i < matrix.length; i++) {
    var toplam = matrix[i].reduce(function(a, b) { return a + b; }, 0);
    if (Math.abs(toplam - 1.0) > 0.0001 && toplam !== 0) {
      throw new Error("Satır " + i + " olasılık toplamı 1 değil (" + toplam + ")");
    }
  }
}

function testMarcovMatrisBoyut() {
  var matrix = MarcovGlobal.getTransitionMatrix();
  var len = matrix.length;
  for (var i = 0; i < len; i++) {
    if (matrix[i].length !== len) throw new Error("Matris kare değil!");
  }
}

function testMarcovTahmin() {
  var res = MarcovGlobal.predictNextState("YUKSEK");
  if (!res) throw new Error("Tahmin üretilemedi!");
}

function testGecersizVeri() {
  var res = MarcovGlobal.predictNextState(null);
  if (res !== null && res !== "UNKNOWN") throw new Error("Null girdi hatalı ele alındı!");
}