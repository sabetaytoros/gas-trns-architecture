/****************************************************************************/
// observeChange.gs 
/*****************************************************************************/
const ms = 1000
const mn = 60 * ms
const MAX_RUNNING_TIME = 320000
const ntest = true // release de true ya cekilir
const B3 = ntest ? 'B3' : 'Y24' // Current
const F5 = ntest ? 'F5' : 'Z24' // High
const F4 = ntest ? 'F4' : 'AA24' // Low
const F3 = ntest ? 'F3' : 'AB24' // Open
const cas = 3 // orderSheet starting No 
var orO = {}
/****************************************************************************/
const ctB = 4   // Security Sheet Buy column 'D'
const ctS = 5   // Security Sheet Sell column 'E'
const coB = 6   // Order Sheet Buy column 'F'
const coS = 7   // Order Sheet Sell column 'G'
const coH = 10  // High
const coL = 11  // Low 
const fntclr = '#ffffff' // Siyah #ffffff 000000
const bckclrY = '#00ff00' // Buy background Yesil
const bckclrK = '#f50000' // Sell background Kirmizi

function newtest() {
  // prepareNewDaySecurities()
  Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  Trns.trs = Trns.sht.getActiveSheet()
  Logger.log('Name %s', Trns.trs.getName())
  pollValues()
  return
  Trns.ord = Trns.sht.getSheets()[Orders]
  orderSheetNewDay()
  return
  Trns.Name = Trns.trs.getSheetName()
  Logger.log('Name %s', Trns.Name)
  getSecNamesfromOrderSheet()
  var charts = Trns.trs.getCharts();

  if (charts.length > 0) {
    var chart = charts[0]; // Select first chart

    var modifiedChart = chart.modify()
      .setChartType(Charts.ChartType.CANDLESTICK)
      .setOption('candlestick', {
        'hollowIsRising': true,
        'fallingColor': {
          'stroke': 'red',
          'fill': 'red'
        },
        'risingColor': {
          'stroke': 'green',
          'fill': 'green'
        }
      })
      .build();

    Trns.trs.updateChart(modifiedChart);
  }
}

/****************************************************************************/
function drvTestfunction() {
  Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  setPropertyActiveSheet(cas)
  prcOpenUrl()
}
/****************************************************************************/
function testsimulateBracketOrder() {
  Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
   Trns.trs = Trns.sht.getActiveSheet() 
   profitPercent = 1.9  
   stopLossPercent = 0.75

  // Convert percentage inputs to decimal multipliers
  parentPrice = Trns.trs.getRange('F3').getValue()
  for(i = 0; i <12; ++i ) {
  var profitTarget = parentPrice * (1 + profitPercent / 100);
  var stopLossTarget = parentPrice * (1 - stopLossPercent / 100);
  Logger.log('parent   profit stoploss %s %s %s', parentPrice.toFixed(2), profitTarget.toFixed(2),stopLossTarget.toFixed(2)) 
    parentPrice *= 0.985 
   }
  //simulateBracketOrder(Trns.trs.getRange('F3').getValue(), 1.9, 0.75);
}
/****************************************************************************/
function generateGunluk() {
  Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  Trns.trs = Trns.sht.getActiveSheet()
  Trns.gnl = Trns.sht.getSheets()[Gunluk]; // Gunluk Sheet
  Trns.Name = Trns.trs.getSheetName()
  if (Trns.trs.getRange('A25').isBlank()) return
  Trns.trs.getRange('A24').activate()
  Trns.trs.getSelection()
    .getNextDataRange(SpreadsheetApp.Direction.DOWN).activate()
  // Logger.log('Name %s r %s rw %s', Trns.Name, Trns.trs.getActiveRange().getA1Notation(), Trns.trs.getActiveRange().getLastRow())
  /*
    let lrw = Trns.trs.getActiveRange().getLastRow()
    for (r = lrw; r > 23; --r) {
   //   Logger.log('r %s', r)
      if (!Trns.trs.getRange(r - 1, 9).isBlank()) return
      createBuyGunluk(r, true)
      createSellGunluk(--r, true)
    }*/
}

const rhg = 4
const rlw = 5
const fhg = 'F' + rhg
/****************************************************************************/
function setCellHighColor(c) {
  try {
    Trns.tenHgh = Trns.trs.getRange('F8').getValue()
    cvl = Trns.trs.getRange(fhg).getValue()
    // Logger.log('tenHgh %s v %s',Trns.tenHgh,cvl)
    if (cvl >= Trns.tenHgh || !eps(Trns.tenHgh, cvl)) {
      let rnhg = Trns.trs.getRange(rhg, c).getA1Notation()
      if (Trns.trs.getRange(rnhg).getFontColor() != fntclr) {
        Trns.trs.getRange(rnhg).setFontColor(fntclr).setBackground(bckclrK)
        Trns.trs.getRange('F8').setFontColor(fntclr).setBackground(bckclrK)
        setOrderRow()
        Trns.ord.getRange(Trns.or, coS).setValue(Trns.tenHgh)
          .setFontColor(fntclr).setBackground(bckclrK)
        Trns.ord.getRange(Trns.or, coH).setFontColor(fntclr).setBackground(bckclrK)
      }
    }
  }
  catch (e) {
    Logger.log('catch Hig or %s', Trns.or)
  }
}
const rnlw = 'F' + rlw
/****************************************************************************/
function setCellLowColor(c) {
  try {
    Trns.tenLow = Trns.trs.getRange('F7').getValue()
    v = Trns.trs.getRange(rnlw).getValue()
    var flg = eps(Trns.tenLow, v)
    if (v <= Trns.tenLow || !flg) {
      let rng = Trns.trs.getRange(rlw, c).getA1Notation()
      if (Trns.trs.getRange(rng).getFontColor() != fntclr) {
        Trns.trs.getRange(rng).setFontColor(fntclr).setBackground(bckclrY)
        Trns.trs.getRange('F7').setFontColor(fntclr).setBackground(bckclrY)
        setOrderRow()
        //        Logger.log('Sonra or %s', Trns.or)
        Trns.ord.getRange(Trns.or, coB).setValue(Trns.tenLow)
          .setFontColor(fntclr).setBackground(bckclrY)
        Trns.ord.getRange(Trns.or, coL)
          .setFontColor(fntclr).setBackground(bckclrY)
      }
    }
  }
  catch (e) {
    Logger.log('catch  or %s', Trns.or)
  }
}
/************************************************************************/
function testdailyChartValues() {
  Trns.sht = SpreadsheetApp.openById(Trns.ssId)
  setOrderRow()
  Trns.trs = Trns.sht.getSheets()[4]
  Trns.Name = Trns.trs.getName();
  //  Logger.log('Name %s', Trns.Name)
  const cl = 29
  Trns.trs.getRange(26, cl).activate()
  Trns.trs.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate()
  rng = Trns.trs.getActiveRange().getA1Notation()
  var Splt = rng.split(':')
  var tp = fromA1Notation(Splt[1])
  //  Logger.log('Name %s rng %s val %s tr %s', Trns.Name, Trns.trs.getRange(26, cl).getA1Notation(), Trns.trs.getRange(26, cl).getValue(), tp.row)
  let s = tp.row
  let f = 25 //s-1
  //  Logger.log('s %s f %s', s, f)

  /*  for (im = s; im > f; --im ) { 
      let rng = Trns.trs.getRange(im, cl).getA1Notation()
      Logger.log('ı %s  rng %s v %s', im, rng, Trns.trs.getRange(im, cl).getValue())     
      copyFormatRange(rng,'Y24')    
      let c = Trns.trs.getRange('Y24').getValue()   
      let h = Trns.trs.getRange('Z24').getValue()
      let l = Trns.trs.getRange('AA24').getValue()
      let o = Trns.trs.getRange('AB24').getValue()
      if(c > h) copyFormatRange('Y24', 'Z24')
      if(l > c  || Trns.trs.getRange('AA24').isBlank()) copyFormatRange('Y24', 'AA24')
      Logger.log('c %s h %s l %s o %s ', c, h, l, o)
      pollValues()
  
    }  */
  pollValues()
}
/***************************************************************************/
function testdecideSellorBuy() {
  Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  Trns.trs = Trns.sht.getActiveSheet()
  Trns.Name = Trns.trs.getSheetName()
  //  Logger.log(' Name %s', Trns.Name)
  getSecNamesfromOrderSheet()
  decideSellorBuy()
}
/*******************************************************************/
function getTestChangeValues() {
  try {
    var sTime = new Date()
    sTime = new Date(sTime.getTime() + MAX_RUNNING_TIME)
    var i = 0;
    var lock = LockService.getScriptLock()
    lock.releaseLock();
    Logger.log('getTestChangeValues haslock %s', lock.hasLock())
    if (lock.tryLock(5000)) {
      Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
      getSecNamesfromOrderSheet(true)
      pollingSecurities(sTime)
      lock.releaseLock();
    }
  }
  catch (e) {
    Logger.log(' Fail to return catch edildi message %s Stack %s', e.message, e.stack)
    lock.releaseLock();
  }
  if (!isCurrentTimeinSession())
    setPropertyActiveSheet(cas)
}
/***************************************************************************/
function restoreMaxMinColors() {

  Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  let n = 26
  while (n < 27) {
    ++n
    Trns.trs = Trns.sht.getSheets()[n]
    Trns.Name = Trns.trs.getName()
    var r = 9
    var c = 8 // h9
    var rd = 7
    Trns.trs.getRange(r, c).activate()
    Trns.trs.getSelection()
      .getNextDataRange(SpreadsheetApp.Direction.NEXT).activate()
    let d = Trns.trs.getActiveRange().getA1Notation()
    Logger.log(' Name %s ac %s', Trns.Name, d)
    Trns.trs.getRange(d).getDisplayValues()
      .forEach(
        function (row) {
          row.forEach(
            function (cell) {
              //    Logger.log(' %s , ',cell)
              let clr = cell > 0 ? Trns.fclrs.Max : Trns.fclrs.Min
              Trns.trs.getRange(r, c, 3, 1).setFontColor(clr)
              let cld = cell > 0 ? Trns.fclrs.High : Trns.fclrs.Low
              Trns.trs.getRange(rd, c, 2, 1).setFontColor(cld)
              c++
            });
        }
      )
  }
}
// Asagidaki function tamir icin kullanilmistir.
function restoreGunSonu() {
  Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  let n = 26
  while (n < 27) {
    n++
    Trns.trs = Trns.sht.getSheets()[n]
    Trns.Name = Trns.trs.getName()
    var c = Trns.trs.getRange('Kebir!G1').getValue()
    var r = Trns.trs.getRange('Kebir!I1').getValue()
    Trns.trs.getRange(r, c).setFormula('=B3')
    copyCell('B3', Trns.trs.getRange(r, c + 1).getA1Notation())
    let cr = Trns.trs.getRange(r, c).getA1Notation()
    let nr = Trns.trs.getRange(r, c + 1).getA1Notation()
    let sFrm = '=(' + cr + '-' + nr + ')/' + nr

    Trns.trs.getRange(r - 1, c).setFormula(sFrm)

    let s = Trns.trs.getRange(r - 1, c + 1, 4, 1).getA1Notation()
    let d = Trns.trs.getRange(r - 1, c).getA1Notation()

    formatRange(s, d)
    Trns.trs.getRange(r - 1, c).activate()
    Trns.trs.getSelection()
      .getNextDataRange(SpreadsheetApp.Direction.NEXT).activate()

    d = Trns.trs.getActiveRange().getA1Notation()
    r--
    Trns.trs.getRange(r, c).activate()
    cr = Trns.trs.getActiveRange().getA1Notation()
    Trns.trs.getActiveRange().autoFill(Trns.trs.getRange(d)
      , SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    let o = r
    --r
    Trns.trs.getRange(d).getDisplayValues()
      .forEach(
        function (row) {
          row.forEach(
            function (cell) {
              //    Logger.log(' %s , ',cell)
              let cl = cell > 0 ? Trns.fclrs.High : Trns.fclrs.Low
              Trns.trs.getRange(r, c, 3, 1).setFontColor(cl)
              s = Trns.trs.getRange(o, c).getA1Notation()
              copyCell(s, s)
              c++
            });
        }
      )
    s = Trns.trs.getRange(2, 6, 5, 1).getA1Notation()
    copyCell(s, 'L2')
    //    Logger.log('Name %s src', Trns.Name, s)

  }
}
function createHighLowtoOpen(c, rc) {
  let r = 2;
  let o = Trns.trs.getRange(3, c).getA1Notation()
  let l = Trns.trs.getRange(4, c).getA1Notation()
  let h = Trns.trs.getRange(5, c).getA1Notation()
  var sl = '=(' + l + '-' + o + ')/ ' + o
  var sh = '=(' + h + '-' + o + ')/ ' + o
  Trns.trs.getRange(2, c).setFormula(sl)
  Trns.trs.getRange(6, c).setFormula(sh)
  copyCell(o, Trns.trs.getRange(rc, 2).getA1Notation())
  copyCell(h, Trns.trs.getRange(rc, 3).getA1Notation())
  copyCell(l, Trns.trs.getRange(rc, 4).getA1Notation())
  //  Logger.log('c %s val %s', Trns.trs.getRange(15, c).getA1Notation(),
  //    Trns.trs.getRange(15, c).getValue())
  copyCell(Trns.trs.getRange(15, c).getA1Notation(),
    Trns.trs.getRange(rc, 5).getA1Notation())
  //  Logger.log('Name %s o %s l %s h %s sl %s sh %s', Trns.Name, o, l, h, sl, sh)

  Trns.trs.getRange(2, c).activate()
  Trns.trs.getRange(2, c).copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  Trns.trs.getRange(6, c).activate()
  Trns.trs.getRange(6, c).copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}
/****************************************************************************/
function displayOrderProfitRatio(c) {
  //  Logger.log('displayOrderProfitRatio dSale %s Cost %s', orO.dSale, orO.oCost)
  orO.Tax = orO.dSale > orO.oCost ? (orO.dSale - orO.oCost) * 0.20 : 0
  // Logger.log(' Tax %s', orO.Tax)
  orO.oProfit = orO.dSale - orO.oCost - orO.Tax
  orO.oRatio = orO.oProfit / orO.oCost
  orO.oProfit *= orO.oQnty
  //  Logger.log(' Profit %s Ratio %s', orO.oProfit, orO.oRatio)
  Trns.trs.getRange(19, c).setValue(orO.oProfit)
  Trns.trs.getRange(20, c).setValue(orO.oRatio)
}
/****************************************************************************/
function getTrnsLastRow() {
  var dr = 24
  var lr = Trns.trs.getLastRow()
  while (dr < lr)
    if (Trns.trs.getRange(dr, 1).getValue() != "Date") dr++
    else break
  return dr
}
/****************************************************************************/
function oComputePartialCosts() {
  var lRw = getTrnsLastRow()
  var qnty = orO.oQnty
  var rw = 25
  orO.oCost = 0
  for (rw = 25; rw < lRw; rw++) {
    var pQ = Trns.trs.getRange(rw, 21).getValue()
    //    Logger.log('pq %s ', pQ)
    if (pQ > 0) {
      let f = qnty - pQ
      if (qnty > pQ || (f > -0.001 && f < 0.001)) {
        orO.oCost += Trns.trs.getRange(rw, 14).getValue()
        qnty -= pQ
        //        Logger.log(' qnty %s rw %s pRmnQ %s oCost %s', qnty, rw, pQ, orO.oCost)
        if (qnty < 0.001) {
          //         Logger.log(' qnty == 0 %s return ediyor', qnty < 0.001)
          return orO.oCost / orO.oQnty
        }
      } else {
        //        Logger.log(' parsiyel calculate qnty %s p %s', qnty, Trns.trs.getRange(rw, 16).getValue())
        orO.oCost += qnty * Trns.trs.getRange(rw, 16).getValue()
        //        Logger.log(' rw %s pRmnQ %s oCost %s', rw, pQ, orO.oCost)
        return orO.oCost / orO.oQnty
      }
    }
  }
}

/**************************************************************************/
function wiOrderSell(c) {
  //  Logger.log('wiOrderSell %s', c)
  orO.tQnty = Trns.trs.getRange('I24').getValue()
  if (orO.tQnty > 0) {
    orO.oQnty = Trns.trs.getRange(8, c).getValue()
    orO.oPrice = Trns.trs.getRange(9, c).getValue()
    orO.dSale = orO.oPrice * (1 - Trns.trs.getRange('Gunluk!$I$2').getValue())

    //    Logger.log('wiOrderSell  dSale %s', orO.dSale)
    let f = orO.tQnty - orO.oQnty
    //    Logger.log(' tQnty %s oQnty %s oPrice %s f %s', orO.tQnty, orO.oQnty, orO.oPrice, f)

    if (f == 0)
      orO.oCost = Trns.trs.getRange('H24').getValue()
    else
      orO.oCost = oComputePartialCosts()
    //  Logger.log('f %s oCost %s', f, orO.oCost)
    displayOrderProfitRatio(c)
  } else {
    Trns.trs.getRange(8, c).clear({ contentsOnly: true, skipFilteredRows: true })
    Trns.trs.getRange(19, c, 2, 1)
      .clear({ contentsOnly: true, skipFilteredRows: true })
  }
}
/***************************************************************************/
function wiOrderBuy(c) {
  //  Logger.log('wiOrderBuy %s',c)
  orO.oQnty = Trns.trs.getRange(8, c).getValue()
  if (orO.oQnty <= 0) return
  orO.tQnty = Trns.trs.getRange('I24').getValue()
  orO.oPrice = Trns.trs.getRange(9, c).getValue()
  orO.oCost = orO.oPrice * (1 + Trns.trs.getRange('Gunluk!$I$2').getValue())
  //  Logger.log(' tQnty %s oQntTrnsy %s oPrice %s', orO.tQnty, orO.oQnty, orO.oPrice)
  if (orO.tQnty < 0)
    orO.dSale = Trns.trs.getRange('O24').getValue()
  else {
    orO.cPrice = Trns.trs.getRange('B3').getValue()
    orO.dSale = orO.cPrice * (1 - Trns.trs.getRange('Gunluk!$I$2').getValue())
    //    Logger.log('oCost %s cPrice %s dSale %s', orO.oCost, orO.cPrice, orO.dSale)
  }
  displayOrderProfitRatio(c)
}
const cp = 26// Crnt to Prv 23 26 degistirildi
const csv = 17 // Yield Q
/*************************************************************************/
function orderSnapShot(No) {
  let filter = Trns.ord.getFilter()
  let rf = filter.getRange()
  let fr = rf.getRow()
  let lr = rf.getLastRow()
  let rs = lr - fr
  Trns.sName = []
  Logger.log('orderSnapShot rf %s fr %s lr %s cp %s rs %s', rf.getA1Notation(), fr, lr, cp, rs)
  Trns.ord.getRange(2, cp + 2, rs + 1, 1)
    .insertCells(SpreadsheetApp.Dimension.COLUMNS)
  //Trns.ord.getRange(2, cp + 2).setValue(new Date()).setNumberFormat('hh":"mm')
  // Security test
  Trns.ord.getRange(fr + 1, csv).activate();
  Trns.ord.getSelection()
    .getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  if (Trns.ord.getActiveRange().getLastRow() == fr + 1) { // Securities yok
    Trns.ord.getRange(2, 1).activate()
    Trns.ord.getActiveRange()
      .offset(1, 0, rs, rf.getLastColumn())
      .sort([{ column: cp, ascending: false }]);
  } else {// Securities var ise
    Logger.log('securities var ')
    Trns.ord.getRange(2, 1).activate()
    Trns.ord.getActiveRange()
      .offset(1, 0, rs, rf.getLastColumn())
      .sort([{ column: csv, ascending: true }, { column: cp, ascending: false }])
  }
  let src = Trns.ord.getRange(3, cp, rs, 1).getA1Notation()
  let dst = Trns.ord.getRange(3, cp).offset(0, 2).getA1Notation()
  Logger.log('src %s dst %s', src, dst)
  Trns.ord.getRange(src).activate()
  Trns.ord.getActiveRange().copyTo(Trns.ord.getRange(dst)
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  Trns.ord.getActiveRange().copyTo(Trns.ord.getRange(dst)
    , SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  getSecNamesfromOrderSheet(true)
  Trns.ord.activate()
  Logger.log('orderSnapShot finished No %s', No)
}
const rSB = 28 // Starting book keeping row
const cSB = 27 //  Starting book keeping column Current baslangıcı
/****************************************************************************/


/*********************************************************************** */
function sFormatRange(s, src, dst) {
  //  Logger.log('src %s dst %s', src, Trns.trs.getActiveRange().getA1Notation())
  s.getRange(dst).activate()
  s.getRange(src).copyTo(s.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  //  Logger.log('copy format range bitti')
}
const cU = 21
/****************************************************************************/
function orderSheetNewDay() {
  Trns.trs = Trns.sht.getSheetByName('Orders')
  let rf = Trns.trs.getFilter().getRange()
  let fr = rf.getRow() + 1
  let lr = rf.getLastRow()
  for (i = 0; i < 2; ++i) {
    let c = i == 0 ? coB : coL
    Trns.trs.getRange(fr, c, lr - fr + 1, 1)
      .setBackground(fntclr).setFontColor(bckclrY)
    c = i == 0 ? coS : coH
    Trns.trs.getRange(fr, c, lr - fr + 1, 1)
      .setBackground(fntclr).setFontColor(bckclrK)
  }
  let lc = rf.getLastColumn() - cSB - 1
  if (lc > 0) {
    Trns.trs.getRange(fr, cSB, 1, lc).activate();
    Trns.trs.deleteColumns(Trns.trs.getActiveRange().getColumn()
      , Trns.trs.getActiveRange().getNumColumns());
  }
  setProp('TESTCYCLE_NO', '0')
  setPropertyActiveSheet(cas)  
  Logger.log('assign Actual Yields bitti')
}
/****************************************************************************/
function assignActualYields() {
  let c = Trns.ord.getRange('Kebir!G1').getValue()
  Trns.Name = Trns.trs.getName();
  setOrderRow()
  // Logger.log('assignAc Name %s or %s c %s', Trns.Name,Trns.or, c)
  let sf = '=\'' + Trns.Name + '\'!' + Trns.trs.getRange(16, c).getA1Notation()
  // Logger.log('cu %s sf %s',cu, sf)
  Trns.ord.getRange(Trns.or, cu).setFormula(sf)
}

/****************************************************************************/
function moveRange(src, dst) {
  Trns.trs.getRange(dst).activate()
  //  Logger.log('src %s dst %s active range s', src , dst, Trns.trs.getActiveRange().////getA1Notation()) 
  Trns.trs.getRange(src).moveTo(Trns.trs.getActiveRange())
}
/********************************************************************************/
function presentTime() {
  Trns.pd = new Date()
  Trns.trs.getRange('W26').setValue(Trns.pd).setNumberFormat('hh":"mm');
}
/********************************************************************************/
function initPolling() {
  //  Logger.log(' initPolling started ')
  src = 'W24:AA24'
  Trns.CHL = []
  Trns.CHL = GetDataArray(Trns.trs, src)
  for (i = 2; i < 5; i++) {
    if (Trns.CHL[i] == "#N/A")
      return
  }
  r = 26
  for (i = 0; i < 2; ++i, ++r) {
    dst = 'W' + r
    copyFormatRange(src, dst)
  }
}
/*********************************************************************************/
  function setTabColor() {
    n = Trns.sName.indexof(Trns.Name)
    //  Logger.log(' setTabColor  Tab No %s', n)
  }
/****************************************************************************/
  function copyCandleValues() {
    Logger.log('copyCandleValues')
    presentTime()
    let dst = 'W27:AA27'
    Trns.trs.getRange(dst).insertCells(SpreadsheetApp.Dimension.ROWS);
    let src = 'W26:AA26'
    copyFormatRange(src, dst)
    if (eps(Trns.CHL[0], Trns.pCHL[0])) 
      setCellHighColor(crc)
    if (eps(Trns.CHL[1], Trns.pCHL[1]))
      setCellLowColor(crc)
  }
/****************************************************************************/
function pollValues() {
  try {
    Trns.StockQnty = Trns.trs.getRange('I24').getValue()
    // if (Trns.StockQnty != 0) {
    setOrderRow()
    Logger.log(' pollValues started Name %s or %s', Trns.trs.getName(), Trns.or)
    let vFark = false
    if (Trns.trs.getRange('W26').isBlank()) {
      initPolling()
      testupdateLastTransaction()
      decideSellorBuy()
    } else {
      vFark = getCurrentValues()
      Logger.log('getCurrentValues vFark %s ', vFark)
      if (vFark) {
        copyCandleValues()
        testupdateLastTransaction()
        decideSellorBuy()
      }
    }

  }
  catch (e) {
    Logger.log(' poll Values exception message %s Stack', e.Message, e.stack)
  }
}

  /*********************************************************************************/
  function isnotAvailable() {
    if (Trns.trs.getRange('B3').getValue() == "#N/A") {
      Trns.trs.getRange('W26:AB26')
        .clear({ contentsOnly: true, skipFilteredRows: true })
      return true
    }
    return false
  }

  /*****************************************************************************/
  function getCurrentValues() {
    Trns.CHL = []
    Trns.chg = []
    Trns.pCHL = []    
    src = 'Y27:AA27'
    Trns.pCHL = GetDataArray(Trns.trs, src)
     Logger.log('pchl %s', Trns.pCHL)
    Trns.CHL = GetDataArray(Trns.trs, 'Y26:AA26')
     Logger.log('CHL %s', Trns.CHL)
    for (i = 0; i < 3; i++) {
      //Logger.log('i %s crn %s prv %s', i, Trns.CHL[i],Trns.pCHL[i])
      if (eps(Trns.CHL[i], Trns.pCHL[i])) return true
    }
    return false
  }

  /****************************************************************************/
  function setOrderRow() {
    if (Trns.Name == null) return
    if (Trns.ord == null)
      Trns.ord = Trns.sht.getSheets()[Orders]
    if (Trns.oName.length == 0)
      createoNameArray()
    if (Trns.oName.length > 0) {
      if (exclude.indexOf(Trns.trs.getName()) > -1) return
    }
    Trns.or = Trns.oName.indexOf(Trns.Name) + 3
    let f = Trns.ord.getRange(Trns.or, oc).getValue() != Trns.Name ? true : false
    if (f) {
      Logger.log('orow %s Name %s orderName %s', Trns.or, Trns.Name, Trns.ord.getRange(Trns.or, oc).getValue())
      createoNameArray()
    }
  }
  /***************************************************************************/
  function createoNameArray() {
    Trns.oName = []
    var b = true
    let or = 3
    while (b) {
      var v = Trns.ord.getRange(or, oc).getValue()
      if (v == "") break;
      Trns.oName.push(v)
      //    createtoNotifyOrders(or)
      or++
    }
    //  Logger.log('Order names %s', Trns.oName)
    if (Trns.trs != null)
      setOrderRow()
  }
  const oc = 5
  /***************************************************************************/
  function getSecNamesfromOrderSheet(flg = false) {
    if (Trns.ord == null)
      Trns.ord = Trns.sht.getSheets()[Orders]
    if (Trns.oName.length == 0 || flg)
      createoNameArray()
    if (Trns.oName.length > 0 && Trns.Name != null)
      setOrderRow()
  }
  /****************************************************************************/
  function setDate(r = 24) {
    var Rng = Trns.trs.getRange(Trns.trs.getRange('Kebir!H1').getValue()).activate()
      .offset(-2, 0).getA1Notation()
    Trns.trs.getRange('A' + r).activate();
    Trns.trs.getRange(Rng).copyTo(Trns.trs.getActiveRange()
      , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }

  /****************************************************************************/
  function updateLastTransaction() {
    //  if (Trns.StockQnty > 0) setDate()
    let clr = Trns.trs.getRange('B5').getValue() > 0 ? Trns.fclrs.High : Trns.fclrs.Low
    Trns.trs.getRange('B3').setFontColor(clr)
    Trns.trs.getRange('B5').setFontColor(clr)
    computeSituation()
  }
  /****************************************************************************/
  function getRange(r) {
    return Trns.trs.getRange(r).getValue()
  }
  /****************************************************************************/
  function elapsedTime(eTime) {
    var crt = new Date()
    // Logger.log(' Test testElapsedTime crt %s eTime %s kt %s ', crt, eTime, eTime - crt)
    return eTime - crt < 10000 ? flgEnum.FTIMEOUT : flgEnum.FCONT
  }
  /****************************************************************************/
  function updateDayLastTime() {
    t = new Date()
    Trns.trs.getRange('V25').setValue(t).setNumberFormat('hh":"mm');
    //Logger.log('frk false time %s B3 %s CHL[0] %s',t, Trns.trs.getRange('B3').getValue(), Trns.CHL[0])
  }
  /****************************************************************************/
  function dayStart() {
    Logger.log(' day Start Name %s', Trns.Name)
    setOrderRow()
    let c = Trns.sht.getRange('Kebir!G1').getValue() + 1
    let opn = Trns.trs.getRange('F3').getValue()
    let pop = Trns.trs.getRange(3, c).getValue()
    let dlo = (opn - pop) / pop
    //  Logger.log(' c %s opn %s popn %s dlo %s or %s ', c, opn, pop, dlo, Trns.or)
    let r = Trns.ord.getRange(Trns.or, 1).getA1Notation()
    Trns.ord.getRange(r).setValue(dlo)
    /*
      Trns.ord.getRange(Trns.or, 1).setValue(price)
        .setFontColor(fontColor)
        .setBackground(background);
      //Logger.log(' or %s fntColor %s back %s', Trns.or, fontColor, background)*/

  }
  /****************************************************************************/
  function setSheetNamesArray() {
    if (Trns.sName.length > 0) return
    Trns.sht.getSheets().forEach(
      function (e) {
        Trns.sName.push(e.getName());
      }
    );
  }


  /****************************************************************************/
  function pollingSecurities(eTime) {
    var alls = Trns.sht.getSheets();
    setSheetNamesArray()
    let lr = Trns.ord.getFilter().getRange().getLastRow() + 1
    let s = getStartingPoint()
    info = {}
    for (; ;) {
      Trns.Name = Trns.ord.getRange(s, oc).getValue()
      Logger.log('Name from order Sheet %s', Trns.Name)
      let i = Trns.sName.indexOf(Trns.Name)
      Trns.trs = alls[i]
      if (exclude.indexOf(Trns.Name) == -1) {
        Trns.trs.activate()
        Trns.trs.getRange('X21:Z21')
          .clear({ contentsOnly: true, skipFilteredRows: true })
        if (elapsedTime(eTime) == flgEnum.FTIMEOUT) {
          setPropertyActiveSheet(s)
          Logger.log('time out starting point %s ', getStartingPoint())
          info.scrty = Trns.Name
          info.flg = flgEnum.FTIMEOUT
          return info
        } else
          pollValues()
      }
      ++s
      if (s + 1 > lr) 
        s = cas    
      setPropertyActiveSheet(s)
    }
  }
  /***************************************************************************/
  function updateSnapShotCyle() {

    var No = parseInt(logProp('TESTCYCLE_NO'))
    if (!parseInt(No % 6)) {
      Logger.log('updateSnapShotCyle No %s  orderSnapShot cagriliyor', No)
      orderSnapShot(No)
    }
    ++No
    setProp('TESTCYCLE_NO', No.toString())
    return No
  }

  const eTrnd = { FSELL: 1, FBUY: 2, FEND: 3 }
  /*******************************************************************************/
  function isSecurityProfiting() {
    Trns.Yield = Trns.trs.getRange('F24').getValue()
    if (Trns.Yield > 0.0125) return false
    Trns.prvPrice = Trns.trs.getRange('B4').getValue()
    if (Trns.crPrice > Trns.prvPrice) return false
    //  Logger.log('isSecurityProfiting Yield %s', Trns.Yield)
    return true
  }
  /******************************************************************************
  function updateOrderSheetOrderPrice(c, price, fontColor, background) {
    Logger.log(' updateOrderSheetOrderPrice')
    setOrderRow()
    Trns.ord.getRange(Trns.or, c).setValue(price)
      .setFontColor(fontColor)
      .setBackground(background);
    Logger.log(' or %s fntColor %s back %s', Trns.or, fontColor, background)
  }*/

  /****************************************************************************/
  //     Trns.crPrice = 151.30 Bu degerler test icin
  //     prValue = 152.04  Bu degerler test icin
  //      Trns.StockQnty =0   Bu degerler test icin 
  function testUpdateBuyOrder(prValue) {
    //  Logger.log(' testUpdateBuyOrder() icinde prValue %s Name %s', prValue, Trns.trs.getName())
    let orValue = Trns.trs.getRange(9, ctB).getValue()


    let sta = prValue > Trns.crPrice && Trns.prValue < orValue // 152.04 > 151.30 
    let stb = orValue > Trns.crPrice || Trns.tenLow > Trns.crPrice

    //  Logger.log('crPrice %s sta %s orValue %s stb %s tenLow %s', Trns.crPrice, sta, orValue, stb, Trns.tenLow)
    if (sta || stb) {
      //   Logger.log('sta veya stb ')
      let v = sta ? Trns.prValue : Trns.crPrice
      //   updateOrderSheetOrderPrice(coB, v  , fntclr, bckclrY)
      //  Logger.log(' qnty - 0 f %s v %s fc %s bc %s', prValue > Trns.crPrice, v, fntclr, bckclrY)
      wiOrderBuy(ctB)
      Trns.trs.getRange(9, ctB).setValue(v)
        .setFontColor(fntclr)
        .setBackground(bckclrY)
    }
  }
  /*******************************************************************************/
  function testUpdateSellOrder(prValue) {
    //Logger.log(' testUpdateSellOrder() icinde prValue %s', prValue)
    let orValue = Trns.trs.getRange(9, ctS).getValue()

    //     Trns.crPrice = 151.30 Bu degerler test icin
    //     prValue = 152.04  Bu degerler test icin
    //     Trns.StockQnty = 0   Bu degerler test icin 
    let sta = prValue < Trns.crPrice && Trns.prValue > orValue // 152.04 > 151.30 
    let stb = orValue < Trns.crPrice || isSecurityProfiting()
    //Logger.log('crPrice %s sta %s orValue %s stb %s tenHgh %s ', Trns.crPrice, sta, orValue, stb, Trns.tenHgh)
    if (sta || stb) {
      let v = sta ? Trns.prValue : Trns.crPrice
      //   updateOrderSheetOrderPrice(coS, v  , fntclr, bckclrK)
      //    Logger.log('update Order Sht donus %s', prValue > Trns.crPrice)
      //    Logger.log(' qnty  %s v %s fc %s bc %s', prValue > Trns.crPrice, v, fntclr, bckclrK)
      wiOrderSell(ctS)
      Trns.trs.getRange(9, ctS).activate();
      Trns.trs.getRange(9, ctS).setValue(v)
        .setFontColor(fntclr)
        .setBackground(bckclrK)
    }
  }
  /****************************************************************************/
  function decideSellorBuy() {
    Logger.log('decideSellorBuy started')
    //  Logger.log('D3 %s ', Trns.trs.getRange('D3').getValue().toString())
    var cr = Trns.trs.getRange('D3').getFormula().substring(1)
    if (cr != '#') {
      var rr = Trns.trs.getRange(cr).offset(1, 1).getA1Notation()
      var rt = parseFloat(Trns.trs.getRange(rr).getValue())
      var art = parseFloat(Trns.trs.getRange(cr).offset(4, 0).getValue())
      //Logger.log('Content of D3 ratio %s Trend ratio %s reality %s', rr, rt, art)
      Trns.typ = rt > 0 ? 'S' : 'B'
      Trns.orPrice = []
      Trns.crPrice = Trns.trs.getRange('B3').getValue()// prediction value      
      prValue = Trns.trs.getRange(cr).offset(3, 0).getValue()
      Trns.StockQnty = Trns.trs.getRange('I24').getValue()
      Trns.Clsd = Trns.trs.getRange('L24').getValue()
      if (eps(Trns.Clsd, Trns.crPrice))
        updateLastTransaction()
      // Logger.log(" Type %s qnty %s prValue %s  crPrice %s", Trns.typ, Trns.StockQnty, prValue, Trns.crPrice)
      if (Trns.StockQnty > 0) {
        //     Logger.log('testUpdateSellOrder() cagriliyor')
        testUpdateSellOrder(prValue)
      } else {
        testUpdateBuyOrder(prValue)
        //     Logger.log('testUpdateBuyOrder() bitti')
      }
    }
  }
  /*******************************************************************************/
  function updateOrder(i) {
    var c = Trns.typ == "B" ? 2 : 3
    Trns.orPrice[i] = parseFloat(Trns.trs.getRange(9, c).getValue())
    var flg = (Trns.typ == "S") ? Trns.crPrice > Trns.orPrice[i]
      : Trns.crPrice < Trns.orPrice[i]
    //Logger.log('updateOrder flg  %s', flg)
    if (flg) {
      Trns.orPrice[i] = Trns.crPrice
      Trns.orPrice[i] *= ((i == 0) ? 1
        : Trns.type == 'S' ? 1.015 : 0.985)
      Trns.trs.getRange(9, c).setValue(Trns.orPrice[i])
      formatRange('B3', Trns.trs.getRange(9, c).getA1Notation())
      updateOrderSheet(Trns.trs.getRange(9, c).getA1Notation())
      //  Logger.log("Order Price Updated %s ", Trns.orPrice[i])
    }
  }


  /*******************************************************************************/
  function format(number, decimals = 2, decimalSeparator = '.'
    , thousandsSeparator = ',') {
    if (number == NaN || number == null) number = 0
    const roundedNumber = number.toFixed(decimals);
    let integerPart = '', fractionalPart = '';
    if (decimals == 0) {
      integerPart = roundedNumber;
      decimalSeparator = '';
    } else {
      let numberParts = roundedNumber.split('.');
      integerPart = numberParts[0];
      fractionalPart = numberParts[1];
    }
    integerPart = integerPart.replace(/(\d)(?=(\d{3})+(?!\d))/g, `$1${thousandsSeparator}`);
    return `${integerPart}${decimalSeparator}${fractionalPart}`;
  }
  /*****************************************************************************
  function cRatio(f, s) { return (f == s) ? 0 : (f - s) / s }
  /*******************************************************************************/
  function eps(l, r) {
    var sl = l.toString()
    var nl = sl.replace(',', '')
    var sr = r.toString()
    var nr = sr.replace(',', '')

    //  Logger.log('sl %s nl %s nr %s f %s', sl, nl,nr, nl - nr)
    var f = abs(nl - nr);
    //  Logger.log('sl %s r %s l-r %s f %s flg %s', sl, r, l-r, f, f>0.005)
    return f > 0.0005
  }

  /*******************************************************************************/
  function sprintf() {
    var args = arguments,
      string = args[0],
      i = 1;

    return string.replace(/%((%)|s|d)/g, function (m) {
      // m is the matched format, e.g. %s, %d
      var val = null;
      if (m[2]) {
        val = m[2];
      } else {
        val = args[i];
        // A switch statement so that the formatter can be extended. Default is %s
        switch (m) {
          case '%d':
            val = parseFloat(val);
            if (isNaN(val)) {
              val = 0;
            }
            break;
        }
        i++;
      }
      return val;
    });
  }
/********************************************************************************/
function createHourlyCandlestick() {
  // Open the spreadsheet and select the relevant sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Assuming your minute-level data is in columns A to E (Time, Open, High, Low, Close)
  var dataRange = sheet.getRange("A2:E"); // Adjust as necessary
  var data = dataRange.getValues();
  
  // Initialize an array to hold hourly candlestick data
  var hourlyCandles = [];
  var previousClose = null;

  // Loop through the minute-level data
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var time = new Date(row[0]); // Assuming first column is Time
    var openValue = (previousClose !== null) ? previousClose : row[1]; // Open from previous close or current open
    var highValue = row[2];
    var lowValue = row[3];
    var closeValue = row[4];

    // Check if we're at a new hour
    if (i === 0 || time.getHours() !== new Date(data[i - 1][0]).getHours()) {
      if (hourlyCandles.length > 0) {
        // Check if the last close is unchanged
        if (hourlyCandles[hourlyCandles.length - 1][3] === closeValue) {
          continue; // Do nothing if the last close is unchanged
        }
      }

      // Create a new hourly entry
      hourlyCandles.push([time, openValue, highValue, lowValue, closeValue]);
      previousClose = closeValue; // Update previous close for next hour's open
    } else {
      // Update high and low values for the current hour
      hourlyCandles[hourlyCandles.length - 1][2] = Math.max(hourlyCandles[hourlyCandles.length - 1][2], highValue);
      hourlyCandles[hourlyCandles.length - 1][3] = Math.min(hourlyCandles[hourlyCandles.length - 1][3], lowValue);
      hourlyCandles[hourlyCandles.length - 1][4] = closeValue; // Update the close value to the latest one
    }
  }

  // Write the hourly candlestick data back to the sheet starting from column G
  var outputRange = sheet.getRange(2, 7, hourlyCandles.length, hourlyCandles[0].length); // Adjust starting column as necessary
  outputRange.setValues(hourlyCandles);
  
  // Optionally create a candlestick chart with this data
  createChart(sheet);
}

function createChart(sheet) {
  var chartRange = sheet.getRange("G2:K"); // Adjust based on where your output is
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.CANDlestick)
    .addRange(chartRange)
    .setPosition(5, 7, 0, 0) // Adjust position as needed
    .setOption('title', 'Hourly Candlestick Chart')
    .build();
  
  sheet.insertChart(chart);
}
simulatePrice = []
/**********************************************************/
function simulateBracketOrder(parentPrice, profitPercent, stopLossPercent) {
  // Convert percentage inputs to decimal multipliers
  var profitTarget = parentPrice * (1 + profitPercent / 100);
  var stopLossTarget = parentPrice * (1 - stopLossPercent / 100);

  // Display order levels
  Logger.log("Parent Order Price: " + parentPrice);
  Logger.log("Profit Limit Price: " + profitTarget.toFixed(2));
  Logger.log("Stop Loss Price: " + stopLossTarget.toFixed(2));

  // Simulate price movement with random steps around the parent price
/*  var prices = [];
  for (var i = 0; i < 10; i++) {
    var simulatedPrice = parentPrice * (1 + (Math.random() - 0.5) / 20); // ±2.5% randomness
    prices.push(simulatedPrice);
  }*/
 // Logger.log("Simulated Prices: " + prices.map(p => p.toFixed(2)).join(", "));

  // Determine which condition gets triggered first
  var hit = "None";
  for (var j = 0; j < prices.length; j++) {
    if (prices[j] >= profitTarget) {
      hit = "Profit Target Reached at " + prices[j].toFixed(2);
      simulateBracketOrder(profitTarget, profitPercent, stopLossPercent) 
      break;
    }
    if (prices[j] <= stopLossTarget) {
      hit = "Stop Loss Triggered at " + prices[j].toFixed(2);
      simulateBracketOrder(stopLossTarget, profitPercent, stopLossPercent) 
      break;
    }
  }

  if (hit === "None") {
    Logger.log("No order executed within simulation.");
  } else {
    Logger.log(hit);
  }
}