/*****************************************************************************/
// Transaction.gs 
/*****************************************************************************/
const Orders = 0;
const Kebir = 2
const Gunluk = 3
const ibkrFee = 2.5
var Ara = false

// Array holding the names of the sheets to exclude from the execution
const exclude = ["Kebir", "Orders", "Gunluk", "Test", "QBTS+","AR(1)Sml"];
const trnType = { tUpdNrm: 0, tUpdShort: 1, tTrnsNrm: 2, tTrnsShort: 3 }
const trState = { Buy: 0, Sell: 1, Info : 2, fBuy : 3 }
const rtax =  columnLetterToNumber('S')
/*****************************************************************************/
Trns.fclrs = {};
Trns.fclrs.Buy = '#1c4587'
Trns.fclrs.Sell = '#cc0000'

Trns.fclrs.High = '#274e13'        // Green
Trns.fclrs.Low =  '#7f6000'//'#ff00ff'  // Magenta

Trns.fclrs.Max = '#660000'
Trns.fclrs.Min = '#073763'

Trns.fclrs.Profit = '#274e13' // Transaction
Trns.fclrs.Loss = '#1c4587'
const coM = 13


/*****************************************************************************/
function testComputeSituation() {
  Logger.log('tax number %s ', rtax)
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  Trns.ord = Trns.sht.getSheets()[Orders]; // Order Sheet
  Logger.log(' Ord Sheet Name %s ', Trns.ord.getName())
  Trns.trs = Trns.sht.getActiveSheet()
  Trns.Name = Trns.trs.getSheetName()
  Logger.log("Name %s", Trns.Name)

  setOrderRow()

  // Logger.log('Name %s Ord Name %s', Trns.Name, Trns.ord.getName())   
  computeSituation()
//displayTotLine22(countURow())

}
let monitoredcell = ['B9','C9','D9','E9']

/******************  Start of Transaction ************************************/
function transaction() {
  try {
    //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
    Trns.trs = Trns.sht.getActiveSheet();  // Transaction sheet
    Trns.Name = Trns.trs.getName()
    Logger.log('Sec Name %s', Trns.trs.getName())   
    Trns.cc = Trns.trs.getCurrentCell().getA1Notation()
    if (Trns.cc != 'B9' && Trns.cc != 'C9' && Trns.cc != 'D9' && Trns.cc != 'E9') {
      SpreadsheetApp.getUi().alert(
        Trns.Name,
        'Transaction Cell (Buy veya Sell) B9/C9/D9/E9 secildikten sonra Transaction komutu girilir  ',
        SpreadsheetApp.getUi().ButtonSet.OK);
      return
    }
    Trns.Qnty = Trns.trs.getRange(Trns.cc).offset(-1, 0).getValue()
    if (Trns.Qnty <= 0) {
       SpreadsheetApp.getUi().alert(
        Trns.Name, 'Sell or Buy Quantity must be greater then zero ', SpreadsheetApp.getUi().ButtonSet.OK);
      return
    }
    /*Trns.cc = 'c9' */
    Trns.gnl = Trns.sht.getSheets()[Gunluk]; // Gunluk Sheet
    setTransactionType();
    if (transactionCanbeCreated()) {
      if (Trns.StockQnty == 0)
        Trns.trs.getRange('L24')
          .clear({ contentsOnly: true, skipFilteredRows: true });
      if (Trns.Clr == Trns.fclrs.Sell)
        createSellTransaction()
      else createBuyTransaction()   
      if (Trns.or != 2) {
        revertOrder()
      }
    }
  }
  catch (e) {
     Logger.log("Error: " + e.message);
  }
}
// n = nRmn
// tr = Transaction Row
/*****************************************************************************/
function generateProfitRatio(n,tr) {
  var c = Trns.trs.getRange('Kebir!G1').getValue()
  let rw = 17 + n
  var lRng = GetUnvRange(Trns, rw, c) // Last Row and column 18  
  if (Trns.trs.getRange(lRng).isBlank() == false)
    shiftRightProfitRatio(rw, c)
  Trns.trs.getRange(lRng).setValue(Trns.Price).setFontColor(Trns.Clr)
  formatRange('B3', lRng)
  var nRng = lRng + ":BZ" + (n > 1 ? rw - 1 : rw)
  var rRng = Trns.trs.getRange(nRng)
    .getNextDataCell(SpreadsheetApp.Direction.NEXT).getA1Notation()
  if (!Trns.trs.getRange(rRng).isBlank()) {
    var sFrm = '=(' + lRng + ' - $' + rRng + ') / $' + rRng
    // Update  Range
    var uRng = GetUnvRange(Trns, 22, Trns.trs.getRange('Kebir!G1').getValue())//19
    var clr = Trns.trs.getRange(uRng).getValue() > 0
      ? Trns.fclrs.Profit : Trns.fclrs.Loss
    if (Trns.typ == 'B')
      Trns.trs.getRange(uRng).setFormula(sFrm)
        .setFontColor(clr).setNumberFormat('#,##0.0000')
    else {
      copyFormatRange('F25', uRng)
      Trns.trs.getRange(uRng).setFontColor(clr)
    }
  }
  var cRng = GetUnvRange(Trns, 15, Trns.trs.getRange('Kebir!G1').getValue())
  sFrm = '=(' + cRng + ' - $' + lRng + ') / $' + lRng
  uRng = GetUnvRange(Trns, rw - 1, Trns.trs.getRange('Kebir!G1').getValue())//17
  Trns.trs.getRange(uRng).setFormula(sFrm).setFontColor(Trns.Clr).setNumberFormat('#,##0.0000')
}
/*****************************************************************************/
function setClsdPrice(rw) {
  Trns.ClsdPrice = (Trns.trs.getRange(Trns.cc).getValue() != 'N/A')
                 ? Trns.trs.getRange(Trns.cc).getValue()
                 : Trns.trs.getRange('B4').getValue()  
  Qnty = Trns.trs.getRange( "U"+ rw).getValue()  
  Trns.DealerSale = Trns.ClsdPrice - (ibkrFee / Qnty)
  //Logger.log('ClsdPrice %s DealerSale %s', Trns.ClsdPrice, Trns.DealerSale )
  
}
/***************************  Buy or Sell Transaction type   ****************/
function setTransactionType() {
  getSecNamesfromOrderSheet()
  if (Trns.or == -1)
    throw (' Not a Valid Securities to do Transaction')
  Trns.typ = (Trns.cc == 'B9' || Trns.cc == 'D9') ? 'B' : 'S'
  Trns.Clr = (Trns.typ == 'B') ? Trns.fclrs.Buy : Trns.fclrs.Sell
}
/****************************************************************************/
function transactionCanbeCreated() {
  Trns.StockQnty = Trns.trs.getRange('I24').getValue()
  Logger.log('Name %s StockQ %s',Trns.Name, Trns.StockQnty)
  var s = ''
  if (Trns.Clr == Trns.fclrs.Sell) {
    s = ' Warning SELL TRANSACTION \n'
    if (Trns.StockQnty == 0 || Trns.Qnty > Trns.StockQnty) {
       s += sprintf(' Stocked Quantity is %s  or less than Sales Quantity %s.\n Do you want to Perform Short Sale operation ? ', Trns.StockQnty, Trns.Qnty)
    }
  }  else {
    s = 'Warning BUY TRANSACTION \n'
    if (Trns.StockQnty != 0) {
    s += sprintf(' Stocked Quantity is %s total Quantiy will be  %s.\n Do you want to continue anyway ', Trns.StockQnty, Trns.Qnty + Trns.StockQnty)
    }    
  }
  let ui = SpreadsheetApp.getUi();
  let r = ui.alert(Trns.Name, s, ui.ButtonSet.YES_NO);
  //Logger.log('respose %s', r)
  return r == ui.Button.YES
}

/*********************** Prepare Order Row for the next Order ****************/
function revertOrder() {
  //  LoggerAlert(sprintf('RevertOrder '))
  var Clr = Trns.typ == 'B' ? Trns.fclrs.Sell : Trns.fclrs.Buy
  for (i = 0; i < 2; i++) {
    var crp = Trns.typ == 'B' ? (i == 0 ? 1.035 : 1.025)
      : (i == 0 ? 0.965 : 0.975)
    var v = Trns.Price * crp
    var c = Trns.typ == 'B' ? (i == 0 ? 3 : 5)
      : (i == 0 ? 2 : 4)
    //    LoggerAlert(sprintf(' c %s crp %s v %s', c, crp, v))
    Trns.trs.getRange(9, c).setValue(v)
      .setNumberFormat('#,##0.00')
      .setFontColor(Clr);
  }
}
/*****************************************************************************/
function displayGunlukDate(rw) {
  Logger.log('rw %s',rw)
  Trns.gnl.insertRowsBefore('4', 1)
  Trns.trs.getRange(rw, 1).copyTo(Trns.gnl.getRange('B4')
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  Trns.gnl.getRange('B4')
    .setNumberFormat('dd"."mm"."yyyy')
    .setFontColor(Trns.Clr)
  Trns.gnl.getRange('C4')
    .setValue(Trns.trs.getRange('A1').getValue())
    .setFontColor(Trns.Clr)
}
/*****************************************************************************/
function createSellGunluk(rw, f = false) {
  if (f) {
    Trns.Profit = Trns.trs.getRange(rw,5).getValue()
    Trns.Yield = Trns.trs.getRange(rw,6).getValue()
  }
  displayGunlukDate(rw)
  var TtlVal = Trns.gnl.getRange('F5').getValue()
  Trns.trs.getRange(rw, 3).copyTo(Trns.gnl.getRange('D4'),
    SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  Trns.gnl.getRange('D4').setNumberFormat('#,##0.00')
  TtlVal += Trns.gnl.getRange('D4').getValue()
  Trns.gnl.getRange('F4').setValue(TtlVal).setNumberFormat('#,##0.00')
  Trns.gnl.getRange('G4').setValue(Trns.Profit).setNumberFormat('#,##0.00')
  Trns.gnl.getRange('H4').setValue(Trns.Yield).setNumberFormat('#0.0000')  
}
/*****************************************************************************/
function createBuyGunluk(re, f = false) {
  var TtlVal = Trns.gnl.getRange('F4').getValue()
   Logger.log('TtlVal %s re %s', TtlVal, re)
  if (!f && Trns.trs.getRange('I24').getValue() == 0) ++re
  let bRw = !f ? re + 1 : re
  Logger.log('bRw %s',bRw)
  displayGunlukDate(bRw)
  Trns.trs.getRange(bRw, 4).copyTo(Trns.gnl.getRange('E4'),
    SpreadsheetApp.CopyPasteType.PASTE_VALUES.false)
  Trns.gnl.getRange('E4').setNumberFormat('#,##0.00')
  TtlVal -= Trns.gnl.getRange('E4').getValue()
  Trns.gnl.getRange('F4').setValue(TtlVal).setNumberFormat('#,##0.00')
  if (Trns.trnType == trnType.tTrnsShort) {
    displayGunlukDate(re)
    Trns.trs.getRange(re, 3).copyTo(Trns.gnl.getRange('D4'),
      SpreadsheetApp.CopyPasteType.PASTE_VALUES.false)
    Trns.gnl.getRange('D4').setNumberFormat('#,##0.00')
    TtlVal += Trns.gnl.getRange('D4').getValue()
    Trns.gnl.getRange('F4').setValue(TtlVal).setNumberFormat('#,##0.00')
  }
}

const L12 = 12
const N14 = 14
const O15 = 15
const P16 = 16

/*****************************************************************************/
function displayCell(r, v) {
  Trns.trs.getRange(r).setValue(v)
  formatRange('B3', r)
}
/*****************************************************************************/
function displayCellFormat(r,v,f='#,##0.00') { 
  Trns.trs.getRange(r).setValue(v).setNumberFormat(f)
}
/*****************************************************************************/
function countURow() {
  let q = abs(getRange('I24'))
  let r = 24
  let n = 0
  Trns.uCntLRw = 24
  while (q > 0) {
    if (!Trns.trs.getRange('U' + r).isBlank()) {
      s = getRange('U' + r)
      let rq = parseInt(s)
      q -= rq
      ++n
      Trns.uCntLRw = r
    }
    ++r
    if (Trns.trs.getRange('A' + r).isBlank()) return n
  }
  return n
}

const rO = 9
/****************************************************************************/
function computeSituation(src = 'B3') {   
  // Logger.log(' Caller Name %s',computeSituation.caller.name);
  Logger.log('computeSituation started')
  try {
    Trns.StockQnty = getRange('I24')
    if (Trns.StockQnty == 0) return
    if (logProp('GUNSONU_PRC',false) == flgEnum.FCONT)
      Trns.cc = 'B3'
    else {
      Trns.cc = Trns.trs.getActiveRange();
      if (Trns.cc != null) {
        Trns.cc = Trns.cc.getA1Notation()
        if (!(Trns.cc == 'B9' || Trns.cc == 'C9' || Trns.cc == 'D9' || Trns.cc == 'E9'))
          Trns.cc = src
      } else Trns.cc = src
    }
    //Logger.log('cc %s', Trns.cc)
    let nRmn = 0
    Trns.StockQnty = getRange('I24')
    if (Trns.StockQnty != 0) {
      Trns.state = trState.Info
      nRmn = countURow()
      initTotals()
      if (Trns.StockQnty > 0) {
        evaluateBuyState(nRmn)
      } else
        evaluateShortState(nRmn, Trns.cc)
    }
    displayTotLine22(nRmn)
    //if (Trns.state != trState.Info) 

  }
  catch (e) {
    Logger.log('Compute Situation message %s Stack %s', e.message, e.stack)
  }
}
const sr = 24

/**************************************************************************/
function generateSum(name, clmn, sr, lr, dst) {
  s ='=sum('+name+'!' + (clmn+sr) +':'+name+ '!' + (clmn+lr)+ ')'
  // Logger.log('dst %s s %s',dst, s)
  Trns.kbr.getRange(dst).setValue(s)
}
/**************************************************************************/
function evaluateShortState(nRmn,src) {
  let Qnty = Trns.state == trState.Buy ? -Trns.Qnty : Trns.StockQnty 
  let rs = Trns.uCntLRw // Starting Row
  let c = nRmn
  initTotals()
  Trns.trnType = trnType.tTrnsShort
  Trns.Price = Trns.state == trState.Buy ?getRange(Trns.cc) :  getRange(src) 
  if (Trns.Price == 'N/A') Trns.Price = getRange('')
  while (c > 0 && rs > 23) { 
    let urng ='U'+ rs
    let iuQ = parseInt(getRange(urng))   
    if (iuQ > 0) {
      let f = Qnty + iuQ
      if (f <= 0) { //-18400+9200
        displayShortState(rs,iuQ)
        Qnty += iuQ
        if (Qnty == 0) {
          if (Trns.state == trState.Buy) {
            Trns.state = trState.fBuy
            Trns.bRw = rs - 1
            if (c > 1) {
              Qnty = updateShortBuyStockQnty(c - 1, rs - 2)
              // Finished Buy state 
              initTotals()
            }
            --nRmn
          }
          if (Qnty == 0) return nRmn
        }
      }
      --c
    }
    rs -= 2
  }
  return c
}
/**************************************************************************/
function displayShortState(rw, q) {
  let re = rw - 1
  Trns.DealerSale = getRange('O' + re)
Logger.log('DealerSale %s Item Price  %s', Trns.DealerSale, Trns.Price)
  displayCost(rw, q)
  displayComputeCommon(rw, q)  // D,H,Q,R,S,G
  if (Trns.state == trState.Buy) {
    moveRange('U' + rw, 'J' + rw)
    moveRange('B' + re, 'C' + re)
    Trns.trs.getRange('I' + re)
      .clear({ contentsOnly: true, skipFilteredRows: true });
  } else if (Trns.state == trState.Info) 
    setDate(rw)
}
/**************************************************************************/
function updateShortBuyStockQnty(c,r) {
  let qnty = 0
  for(i = 0 ; i < c; ++i) {
    qnty -= getRange('U'+r)
    let m = r - 1
    displayCellFormat('I'+m, qnty, '#,##0')
    r -= 2
  }
  return qnty
}
/****************************************************************************/
function createSellTransaction() {
  LoggerAlert(sprintf('Sell TransAction \nColor %s Name %s Order Row %s cc %s'
    , Trns.Clr, Trns.trs.getName(), Trns.or, Trns.cc))
  Trns.spCost = false
  Trns.state = trState.Sell
  let f = Trns.StockQnty - Trns.Qnty
  if (f < 0)
    selltoOpen(-f)
  else {
    nRmn = newsellStocs()
    displayTotLine22(nRmn)
  }
}
/*****************************************************************************/
function selltoOpen(f) {  
  let nRmn = 0
  Logger.log(' Sell to open Stock Qnty %s Qnty %s', Trns.StockQnty, Trns.Qnty)
  if (Trns.StockQnty > 0) {
    Trns.Qnty = Trns.StockQnty
    newsellStocs()
    Trns.Qnty = f
  }

  Trns.trnType = trnType.tTrnsShort
  if (Trns.StockQnty < 0) {
    insertLine()
    nRmn = countURow()
  }
  shortSellLine()
  ++nRmn
  if (nRmn == 0) {
    insertLine()
    displayCellFormat('I24', 0, '#,##0')
  }
  computeSituation()
  Logger.log(' selltoOpen finished Trns.Qnty %s', Trns.Qnty)
}
/*****************************************************************************/
function displayCellOran(r,v) {
  Trns.trs.getRange(r).setValue(v).setNumberFormat('#0.0000')
}
/*****************************************************************************/
Trns.computeProfit = function (Value,cost = Trns.TotalCost,rw = 24) {
  Trns.Profit = Value - cost
  displayCellFormat('E'+rw,Trns.Profit)
  Trns.Yield = Trns.Profit / cost
  displayCellOran('F'+rw, Trns.Yield)
}

/*****************************************************************************/
function displayTotLine22(n) {  
  setOrderRow()
  Logger.log('displayTotLine22 Name %s or %s', Trns.Name, Trns.or)
  try {  
    srw = 24
    lrw = srw + n*2
    if (n > 0) {    
      sp = '=sum(B' + srw + ':B' + lrw + ')+sum(S'+srw +':S'+lrw+')'
      displayCellFormat('B22',sp) 
      sp = '=sum(D' + srw + ':D' + lrw + ')'
      displayCellFormat('D22',sp) 
      sp = '=B22 - D22'
      displayCellFormat('E22',sp) 
      sp = '= E22 / D22'
      displayCellFormat('F22',sp) 
    }
      srw = lrw
      Trns.trs.getRange( 'A24').activate()
      lrw = Trns.trs.getSelection()
            .getNextDataRange(SpreadsheetApp.Direction.DOWN).getLastRow() 
      sp = '=sum(C' + srw + ':C' + lrw + ')+sum(S'+srw +':S'+lrw+')'    
      displayCellFormat('C21',sp) 
      sp = '=sum(D' + srw + ':D' + lrw + ')'
      displayCellFormat('D21',sp) 
      sp = '= C21 - D21'
      displayCellFormat('E21',sp) 
      sp = '= E21 / D21'
      displayCellFormat('F21',sp) 
/* 
      let s = '=\'' + Trns.Name + '\'!I24'
      Trns.ord.getRange('N' + Trns.or).setFormula(s)
      Trns.ord.getRange('N' + Trns.or)
        .clear({ contentsOnly: true, skipFilteredRows: true });*/

  } 
  catch (e) {
    Logger.log('catch displayLine 21 %s Msg %s', Trns.or, e)
  }
}

/*****************************************************************************/
function shortSellLine() { 
  insertLine()     
  let rw = 25 
  shortBuyLine(rw)
  setClsdPrice(Trns.cc)
  displayCell('O'+24, Trns.DealerSale)
  Trns.StockQnty -= Trns.Qnty
  displayCellFormat('I24',Trns.StockQnty, '#,##0')
  displayCellFormat('K24',Trns.Qnty,'#,##0')
  displayCell('L24', Trns.ClsdPrice)
  displayCell('O24', Trns.DealerSale)  
  displayComputeCommon(rw,Trns.Qnty)  // D,H,N, Q,R,S,G
}

/*****************************************************************************/
function shortBuyLine(rw) {
  if (Trns.state == trState.Sell) { 
    setDate()    
    displayCellFormat('U'+rw,Trns.Qnty,'#,##0')
    Trns.Price = getRange(Trns.cc)
  } 
  displayCost(rw) //// M,P ,H,N,D
}
/*****************************************************************************/
function newsellStocs() {
  setClsdPrice(Trns.cc)
  Trns.Price = Trns.ClsdPrice
  Trns.trntype = trnType.tTrnsNrm
  nRmn = countURow()
  nRmn -= evaluateSellState(nRmn)
  Trns.StockQnty -= Trns.Qnty
  if (nRmn > 0) {
    updateStockQnty(nRmn)
    Trns.state = trState.Info
    evaluateBuyState(nRmn)
  } else {
    insertLine()
  }
  return nRmn
}
/****************************************************************************/
function createBuyTransaction() {
 Logger.log('CreateBuyTransaction Name %s', Trns.trs.getName())
  Trns.state = trState.Buy
  Trns.Qnty = Trns.trs.getRange(Trns.cc).offset(-1, 0).getValue()
  Trns.trnType = Trns.StockQnty < 0 ? trnType.tTrnsShort : trnType.tTrnsNrm
  Trns.Price = Trns.trs.getRange(Trns.cc).getValue()
/*   
  c = 3
  for (i = 0; i < 2; i++, c += 2)
    Trns.trs.getRange(8, c).setValue(Trns.StockQnty)
*/   
  let rw = 24
  if (Trns.trnType == trnType.tTrnsShort) {
    Trns.rmnQ = Trns.Qnty + Trns.StockQnty
    if (Trns.rmnQ  > 0) Trns.Qnty = -Trns.StockQnty
    let nRmn = countURow()
    Logger.log('nRmn %s',nRmn)
    evaluateShortState(nRmn) 
    if (Trns.rmnQ == 0) {
      insertLine(rw) 
      displayTotLine22(Trns.rmnQ)
      return 
    }
    if (Trns.rmnQ <= 0) return
    Trns.StockQnty = 0
    Trns.Qnty = Trns.rmnQ
    insertLine(rw)
  } 
  setClsdPrice(Trns.cc)
  let nRmn = 0
  nRmn = countURow()
 // if (Trns.StockQnty > 0) {   
  //  evaluateBuyState(nRmn) 
  //  rw = Trns.uCntLRw+1
  //  insertLine(rw)
 // }
Logger.log('StockQnty evvel %s',Trns.StockQnty)
  Trns.StockQnty += Trns.Qnty
  Trns.state = trState.Buy 
  Trns.trnType = trnType.tTrnsNrm  
Logger.log('StockQnty sonra %s',Trns.StockQnty)
  rw = Trns.uCntLRw 
  if (nRmn > 0) {  
    ++rw
    insertLine(rw)
  }
  computeBuyItem(rw)
  updateStockQnty(nRmn)  
  computeSituation()    
}
/*****************************************************************************/
function  updateStockQnty(nRmn) {
  rs = 24
  let q = Trns.StockQnty
  for (i=0; i < nRmn; ++i,++rs) {
    displayCellFormat('I'+rs,q,'#,##0')
    ++rs 
    q -= getRange('J'+rs)
  }
}
/*****************************************************************************/
function displayCost(rw=24, qnty = Trns.Qnty) { // M,P,H,N,D
  displayCell('M'+rw,Trns.Price)
  Trns.DealerCost = Trns.Price + (ibkrFee / Trns.Qnty)
  displayCell('P'+rw,Trns.DealerCost)
  displayCell('H'+rw,Trns.DealerCost)
  Trns.Cost = qnty * Trns.DealerCost
  displayCellFormat('N'+rw,Trns.Cost)
  displayCellFormat('D'+rw,Trns.Cost)    
}
/*****************************************************************************/
function computeBuyItem(rw) {
  setDate(rw)
  displayCellFormat('J'+rw,Trns.Qnty,'#,##0')
  displayCellFormat('U'+rw,Trns.Qnty,'#,##0')  
  displayCost(rw) // // M,P,H,N,D
  insertLine(rw)
  calculateBuyItem(rw+1)
}
/*****************************************************************************/
function initTotals() { 
  Trns.TotalCost = 0
  Trns.TotalRvnue = 0
  Trns.TotalTax = 0
  Trns.TotalValue = 0
}
/*****************************************************************************/
function calculateSellItem(rw) { 
  let re = rw - 1  // 25
  setDate(re)
  commonItemValue(rw,re, 'C') 
  moveRange('U' + rw, 'K' + re)
  moveRange('B' + re, 'C'+ re)
  Trns.trs.getRangeList([ 'I' + re])
    .clear({ contentsOnly: true, skipFilteredRows: true });
  if (Trns.state == trState.Sell)
    createSellGunluk(re)
}
/*****************************************************************************/
function commonItemValue(rw,re) {
  setClsdPrice(rw)
  displayCell('L'+re, Trns.ClsdPrice)
  displayCell('O'+re, Trns.DealerSale)
  displayComputeCommon(rw,getRange('U'+rw))
  if (Trns.state != trState.Buy)
    setDate()
}
/*****************************************************************************/
function displayComputeCommon(rw, qnty, sr = 'B'){ // D,H,Q,R,S,G
  let re = rw - 1
  Trns.Cost = getRange('D'+rw)
  Trns.TotalCost += Trns.Cost  
  let ucost = getRange('H'+rw)
  let utax = 0.0
  let rvnue = Trns.DealerSale > ucost ? Trns.DealerSale - ucost : 0.0
  if (rvnue > 0.0) {
    //Logger.log('rv %s qnty %s', rvnue,qnty)
    Trns.Rvnue = rvnue * qnty
    displayCellFormat('Q' + re, rvnue)
    utax = rvnue > 0 ? rvnue * 0.25 : 0
    displayCellFormat('R' + re, utax)
    let ttax = utax * qnty
    displayCellFormat('S' + re, ttax)
    Trns.TotalTax += ttax
  } else { 
    rng = 'Q'+re +':S'+re
    Logger.log('rng %s',rng)
    Trns.trs.getRange(rng).clear({ contentsOnly: true, skipFilteredRows: true });
  }
  let uSale = Trns.DealerSale - utax
  displayCell('G' + re, uSale)
  let iv = uSale * qnty
  displayCellFormat(sr + re, iv)
  Trns.TotalValue += iv
  Trns.computeProfit(iv, Trns.Cost, re)
}
/*****************************************************************************/
function calculateBuyItem(rw, nRmn = 1) {
  let re = rw - 1
  commonItemValue(rw, re, 'B')
  if (Trns.state == trState.Buy) {
    displayCellFormat('I' + re, Trns.Qnty, '#,##0')
    createBuyGunluk(re)
  }
}
/****************************************************************************/
function insertLine(rw=24) {
  if (rw ==24) {
    if (Trns.trs.getRange(24,1).isBlank()) setDate(rw)
  }
  let s = 'A'+rw+':U'+rw
  Trns.trs.getRange(s).activate();
  Logger.log('Active Range %s', Trns.trs.getActiveRange().getA1Notation())
  try {
    Trns.trs.getRange(s).insertCells(SpreadsheetApp.Dimension.ROWS);
  }
  catch(e) {
    Trns.trs.getRange( 'A'+rw).activate()
    Trns.trs.getSelection()
      .getNextDataRange(SpreadsheetApp.Direction.DOWN).activate()
    let s = 'A'+rw+':U'+Trns.trs.getActiveRange().getLastRow()  
    moveRange(s,'A'+(rw+1))
  }
  setDate(rw)  
  formatRange('A'+(rw+1),'A'+rw)
} 
/****************************************************************************/
function evaluateBuyState(nRmn) {
  let Qnty = Trns.state == trState.Sell ? Trns.Qnty :  Trns.StockQnty  
  //Logger.log('evaluate buy state %s Info %s ', Trns.state, trState.Info)
  let rs = 25 // Starting Row
  let c = nRmn
  while (c > 0) {
    let urng = 'U' + rs
    let rmnQ = parseInt(getRange(urng))
    //Logger.log(' rmnQ %s,', rmnQ)
    let f = Qnty - rmnQ
    if (f >= 0) {
      calculateBuyItem(rs, nRmn)
      Qnty -= rmnQ
      // Logger.log('Qnty > rmnQ val %s', getRange('P' + rs))
      if (Qnty == 0)
        return 
    }
    c--
    rs += 2
    if (Trns.trs.getRange('A' + rs).isBlank()) return
  }
}
/*****************************************************************************/
function evaluateSellState(nRmn) {
  let Qnty =  Trns.Qnty
  //Logger.log('evaluate buy state %s Info %s ', Trns.state, trState.Info)
  let rs = 23+nRmn*2 // Starting Row
  let c = nRmn
  let si = 0
  while (c > 0) {
    let urng = 'U' + rs
    let rmnQ = parseInt(getRange(urng))
    //Logger.log(' rmnQ %s,', rmnQ)
    let f = Qnty - rmnQ
    if (f >= 0) {
      calculateSellItem(rs, nRmn)
      Qnty -= rmnQ
      ++si
      if (Qnty == 0) return si
    }
    c--
    rs -= 2
    if (Trns.trs.getRange('A' + rs).isBlank()) break
  }
  return si
}


/*******************************************************************************/
function createChart(rng, min, max) {
  var achts = Trns.trs.getCharts();
  for (var i in achts)
    Trns.trs.removeChart(achts[i]);
 
  var chart = Trns.trs.newChart()
    .setChartType(Charts.ChartType.CANDLESTICK)
    .setOption('candlestick', {
        fallingColor: { fill: 'red', stroke: 'red' },
        risingColor: { fill: 'green', stroke: 'green' }
    })
    .addRange(Trns.trs.getRange(rng))
    /*
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('bubble.stroke', '#000000')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('isStacked', 'false')
    .setOption('title', 'Open, High, Low and Close')
    .setOption('annotations.domain.textStyle.color', '#8d9a9c')
    .setOption('textStyle.color', '#1a3438')
    .setOption('legend.textStyle.color', '#31484c')
    .setOption('titleTextStyle.color', '#839194')
    .setXAxisTitle('Date')
    .setOption('hAxis.textStyle.color', '#1a3438')
    .setRange(min, max)
    .setOption('vAxes.0.textStyle.color', '#1a3438')
    .setOption('height', 533)    
    .setOption('width', 1064)
    .setPosition(27, 1, 14, 20)
    .setOption('annotations.total.textStyle.color', '#8d9a9c')*/
    .build();
  Trns.trs.insertChart(chart);
};
/*******************************************************************************/
function getDateRow() {
  var dr = 24
  var lr = Trns.trs.getLastRow()
  while(dr < lr) 
    if (Trns.trs.getRange(dr, 1).getValue() != "Date") dr++
    else break  
  if (Trns.trs.getRange(dr,1).getValue() == "Date") return dr+1
  var mr = Trns.trs.getMaxRows()
  if (mr - dr < 3) {
    while(dr > 24)
      if (Trns.trs.getRange(dr,1).isBlank()) dr-- 
      else break;  
    dr += 2
    var fr = mr > dr ? mr - dr : 0
    //Logger.log(' dr %s Fark row %s', dr, fr)
    if (fr < 3) 
      Trns.trs.insertRowsAfter(lr,3-fr)    
    //Logger.log('dr %s mr %s',dr, mr)     
  }
  var Header = [["Date", "Open", "High", "Low", "Close"]]
  Trns.trs.getRange(dr,1,1,5).setValues(Header)
  return dr+1
}
/*******************************************************************************/
function insertCrntDayValues(rw) {
  //Logger.log('insertCrnt %s r %s',rw, Trns.trs.getRange(rw,1).getA1Notation())
  var c = Trns.trs.getRange('Kebir!G1').getValue()
  //Logger.log(' date %s', Trns.trs.getRange(13,c).getValue())
  Trns.trs.getRange(rw,1)
    .setValue(Trns.trs.getRange(13,c).getValue())
  //Logger.log('Date %s',Trns.trs.getRange(rw,1).getValue())
  Trns.trs.getRange(rw,2).setFormula('=F3')
  Trns.trs.getRange(rw,3).setFormula('=F5')
  Trns.trs.getRange(rw,4).setFormula('=F4')
  Trns.trs.getRange(rw,5).setFormula('=B3')
}
/*******************************************************************************/
function copyFormatValues() {
/*  var sr = getDateRow()+1
  //Logger.log('CopyFormatValues\n Name %s sr %s ', Trns.trs.getName(), sr)
  var rRng = Trns.sco.getRange(3, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN)
  var or = rRng.getRow()
  var rcs = Trns.sco.getRange(3, 1, rRng.getRow() - 2, 5)
  //Logger.log('or %s rcs %s', or, rcs.getA1Notation())
  var cs = Trns.sco.getRange(rcs.getA1Notation()).getValues()
  var ncs = cs.map(function (row) {
    return row.map(function (cell, i) {
      return (i > 0) ? cell / 100 : cell
    })    
  })*/

  var rng = Trns.trs.getRange(sr-1, 1, ncs.length, ncs[0].length).getA1Notation()
  var rtp = or + sr 
  var mr = Trns.trs.getMaxRows()
  if (rtp > mr) {
    var ir = rtp - mr
    //Logger.log('rtp  %s mr %s  ir ', rtp, mr, ir)
    Trns.trs.insertRowsAfter(mr, ir)
    mr = Trns.trs.getMaxRows()
    //Logger.log('after %s ', mr)
  } else Trns.trs.getRange(rng).clear({ contentsOnly: true, skipFilteredRows: true });
  //Logger.log('Range rng %s', rng)
  Trns.trs.getRange(sr, 1, ncs.length, ncs[0].length).setValues(ncs)
  insertCrntDayValues(sr-1)
  formatRange('H8', Trns.trs.getRange(sr, 2, ncs.length, ncs[0].length).getA1Notation())
  Trns.trs.getRange(sr, 1, ncs.length, ncs[0].length).activate()
    .sort({ column: 1, ascending: false });
  var mrng = Trns.trs.getRange(sr, 4, ncs.length, 1).getA1Notation()
  var sFrm = '=MIN(' + mrng + ') '
  Trns.trs.getRange(sr, ncs[0].length + 1).setFormula(sFrm) // 'F1
  var min = parseInt(Trns.trs.getRange(sr, ncs[0].length + 1).getValue() * 20) * 0.05

  mrng = Trns.trs.getRange(sr, 3, ncs.length, 1).getA1Notation()
  sFrm = '=MAX(' + mrng + ') '
  Trns.trs.getRange(sr, ncs[0].length + 2).setFormula(sFrm)
  var max = parseInt(Trns.trs.getRange(sr, ncs[0].length + 2).getValue() * 20) * 0.05
  //Logger.log('min %s max %s', min, max)
  createChart(rng, min, max)
}
/*******************************************************************************/
function createRangeForChart() {
  //Logger.log('Name %s', Trns.sco.getName())
  var sFrm = '=googlefinance(\"' + Trns.trs.getRange('B1').getValue()
    + '\", \"all\" , \"' + Trns.sco.getRange('C1').getDisplayValue() + '\", \"'
    + Trns.sco.getRange('D1').getDisplayValue() + '\")'
  //Logger.log('sFrm %s ', sFrm)
  Trns.sco.getRange('A2').setFormula(sFrm)
}

/******************************************************************************/
function drvTstCs() {
  const rst = 101
  const rmn = rst+2

  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  Trns.trs = Trns.sht.getActiveSheet();
  Logger.log('drvTestfunction() ssName %s sName %s' , Trns.sht.getName(), Trns.trs.getName())
  var rng = Trns.trs.getRange(rst,2,62,5).getA1Notation()
  Logger.log('rng %s ', rng)
/*  var rtp = or + sr 
  var mr = Trns.trs.getMaxRows()
  if (rtp > mr) {
    var ir = rtp - mr
    //Logger.log('rtp  %s mr %s  ir ', rtp, mr, ir)
    Trns.trs.insertRowsAfter(mr, ir)
    mr = Trns.trs.getMaxRows()
    //Logger.log('after %s ', mr)
  } else Trns.trs.getRange(rng).clear({ contentsOnly: true, skipFilteredRows: true });
  //Logger.log('Range rng %s', rng)
  Trns.trs.getRange(sr, 1, ncs.length, ncs[0].length).setValues(ncs)
  insertCrntDayValues(sr-1)
  formatRange('H8', Trns.trs.getRange(sr, 2, ncs.length, ncs[0].length).getA1Notation())
  Trns.trs.getRange(sr, 1, ncs.length, ncs[0].length).activate()
    .sort({ column: 1, ascending: false });
  var mrng = Trns.trs.getRange(sr, 4, ncs.length, 1).getA1Notation()*/
  a ='D'+rmn
  b = 'F'+rmn
  var sFrm = '=MIN('+a+':D168)'
  Trns.trs.getRange(b).setFormula(sFrm) // 'F1
  min = Math.round(Trns.trs.getRange(b).getDisplayValue() / 5) * 5;
  a ='C'+rmn
  b ='G'+rmn
  sFrm = '=MAX('+a+':c168)'
  Trns.trs.getRange(b).setFormula(sFrm)
  max = Math.round(Trns.trs.getRange(b).getDisplayValue() / 5) * 5;
  Logger.log('min %s max %s', min, max)
  createCandlestickChart(rng,min, max)
}
function createCandlestickChart(rng,minV, maxV) {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Remove existing charts
  var charts = sheet.getCharts();
  for (var i in charts) {
    sheet.removeChart(charts[i]);
  }
  
  // Create a new chart, starting with LineChart
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.CANDLESTICK)
    .addRange(sheet.getRange(rng))
    .setPosition(27, 1, 14, 20)
    .setOption('title', 'Open, High, Low and Close')
    .setOption('height', 533)
    .setOption('width', 1064)
    .setOption('vAxis.viewWindow.min', minV)
    .setOption('vAxis.viewWindow.max', maxV) 				
    .setOption('series.0.data', [
      {sourceColumn: 1, role: 'Date'},  // Date column
      {sourceColumn: 2, role: 'Open'},    // Open price      
      {sourceColumn: 3, role: 'High'},     // High price      
      {sourceColumn: 4, role: 'Low'},     // Low price
      {sourceColumn: 5, role: 'Close'},   // Close price
    ])
    .build();
  
  sheet.insertChart(chart);

  // Modify the chart to set candlestick-specific options

  chart = sheet.getCharts()[0].modify()
    .setOption('candlestick.fallingColor.fill', 'red')
    .setOption('candlestick.fallingColor.stroke', 'red')
    .setOption('candlestick.risingColor.fill', 'green')
    .setOption('candlestick.risingColor.stroke', 'green')
    .setOption('hAxis.title', 'Date')
    .build();
  
  sheet.updateChart(chart);
}