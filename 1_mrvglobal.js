/**
 * MARKOV ZİNCİRİ GLOBAL DEĞİŞKENİ (mrv)
 */
const mrv = {
  // Durum kümesi (States: UP, FLAT, DOWN)
  sts: ['UP', 'FLAT', 'DOWN'],
  
  // Geçiş Matrisi (Transition Matrix - P)
  pMtx: [
    [0.333, 0.333, 0.334],
    [0.333, 0.333, 0.334],
    [0.333, 0.333, 0.334]
  ],

  // Mevcut Olasılık Vektörü (Current Vector)
  curVec: [1, 0, 0],

  // Adım sayısı
  stps: 1
};

/**
 * MARKOV ZİNCİRİ HESAPLAMA MOTORU (MrvEng)
 */
const MrvEng = {

  /**
   * Olasılık vektörünü geçiş matrisi ile bir adım ileri taşır (v * P)
   */
  nxtVec(vec, mtx) {
    const len = vec.length;
    const res = new Array(len).fill(0);
    
    for (let j = 0; j < len; j++) {
      for (let i = 0; i < len; i++) {
        res[j] += vec[i] * mtx[i][j];
      }
    }
    return res;
  },

  /**
   * n-adım sonraki durumu tahmin eder
   */
  prdct(stps = 1) {
    let vec = [...mrv.curVec];
    const mtx = mrv.pMtx;

    for (let k = 0; k < stps; k++) {
      vec = this.nxtVec(vec, mtx);
    }
    
    return vec;
  },

  /**
   * Geçmiş veri dizisinden Geçiş Matrisini (Transition Matrix) hesaplar
   */
  fitMatrix(sequence) {
    const sts = mrv.sts;
    const n = sts.length;
    
    const counts = Array.from({ length: n }, () => new Array(n).fill(0));
    const totals = new Array(n).fill(0);

    for (let t = 0; t < sequence.length - 1; t++) {
      const fromIdx = sts.indexOf(sequence[t]);
      const toIdx = sts.indexOf(sequence[t + 1]);

      if (fromIdx !== -1 && toIdx !== -1) {
        counts[fromIdx][toIdx]++;
        totals[fromIdx]++;
      }
    }

    const pMtx = Array.from({ length: n }, () => new Array(n).fill(0));
    for (let i = 0; i < n; i++) {
      for (let j = 0; j < n; j++) {
        pMtx[i][j] = totals[i] > 0 ? (counts[i][j] / totals[i]) : (1 / n);
      }
    }

    mrv.pMtx = pMtx;
    return pMtx;
  }
};
