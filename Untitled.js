/**
 * Fiyat serisinden Günlük Logaritmik Getirileri (Daily Log Returns) hesaplar.
 * 
 * @param {Range} priceRange Fiyat dizisi (Örn: A2:A31 - En güncel fiyat en üstte olmalı)
 * @return {Array[]} Logaritmik getiri dizisi
 * @customfunction
 */
function DAILY_LOG_RETURNS(priceRange) {
  // Veriyi düzleştir ve boş/geçersiz hücreleri temizle
  const prices = priceRange.flat().filter(v => typeof v === 'number' && !isNaN(v) && v > 0);

  if (prices.length < 2) {
    return [["Hata: En az 2 fiyat gerekli"]];
  }

  const logReturns = [];
  // Fiyatlar en yeni üstte olacak şekilde sıralıysa: ln(Bugün / Dün) = ln(prices[i] / prices[i+1])
  for (let i = 0; i < prices.length - 1; i++) {
    const logReturn = Math.log(prices[i] / prices[i + 1]);
    logReturns.push([logReturn]);
  }

  return logReturns;
}


/**
 * Verilen serinin Örneklem Varyansını (Sample Variance - N-1) hesaplar.
 * 
 * @param {Range} dataRange Getiri veya veri serisi (Örn: Log Getiri sütunu)
 * @return {number} Örneklem varyansı
 * @customfunction
 */
function SAMPLE_VARIANCE(dataRange) {
  const values = dataRange.flat().filter(v => typeof v === 'number' && !isNaN(v));
  const n = values.length;

  if (n < 2) {
    return "Hata: Örneklem varyansı için en az 2 veri noktası gerekir (N > 1)";
  }

  // 1. Ortalama (Mean) Hesaplama
  const mean = values.reduce((sum, val) => sum + val, 0) / n;

  // 2. Kareler Farkı Toplamı (Sum of Squared Differences)
  const sumOfSquaredDiffs = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0);

  // 3. Örneklem Varyansı (N - 1 bölmesi)
  const sampleVariance = sumOfSquaredDiffs / (n - 1);

  return sampleVariance;
}