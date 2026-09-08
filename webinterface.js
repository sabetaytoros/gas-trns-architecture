/**************************************************************************/
//  webInterface.gs 
/**************************************************************************/

// son version https://script.google.com/macros/s/AKfycbwY9azvFBy4IxOfJiiB1nA8-X3es5brZr6Wzy0kiJFjHzLhxrvj4_KzX1pq0CVTNth7tg/exec
// cagrilan https://script.google.com/macros/s/AKfycby5uvkdlQ0YAUEb52UFNQdcNDTI41qbH08q7CO5Ds_f/dev
function Test() {
  //testSingleton()
  //openMissingUrl() 
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet() 
  checkSesionFinish() // GunSonu Baslatir
  //checkStartaNewDay() // Yeni bir gun baslatir
  //includeIBKRPageCont()
  //setPropertyActiveSheet(cas)
  Logger.log('crc %s', crc)

}
/***************************************************************/
function checkStartaNewDay() {  // UpdateKebirPage set Eder
  checkifnoSet('POLL_VALUES', flgEnum.FPLVSTRT)
  checkifnoSet('CLOSE_WIN', flgEnum.FSTRT)
  checkifnoSet('DAY_START', flgEnum.FSTRT)
  checkifnoSet('SESION_TRIG',flgEnum.FSTRT) 
  checkifnoSet('IDLE.STATE', flgEnum.FIDLE) 
  checkifnoSet('GUNSONU_PRC',flgEnum.FEND)
}

/***************************************************************/
function checkSesionFinish() {
    PropertiesService.getScriptProperties().deleteProperty('REMAINING_TABS');
    PropertiesService.getScriptProperties().deleteProperty('JOB_INITIALIZED');    
    checkifnoSet('SESION_TRIG', flgEnum.FEND)
    checkifnoSet('DAY_START', flgEnum.FSTRT)
    checkifnoSet('GUNSONU_PRC', flgEnum.FSTRT)
    checkifnoSet('POLL_VALUES', flgEnum.FPLVSTRT) 
    checkifnoSet('IDLE.STATE', flgEnum.FEND)
    checkifnoSet('GUN_SONUUPDATE', flgEnum.FSTRT)

}
/***************************************************************/
function checkifnoSet(pr, vl) {
  if (logProp(pr,false) != vl) { 
    setProp(pr,vl)
    return true
  }
  return false
}
/***************************************************************/
function setProp(p, s, f = true) {
  userProp.setProperty(p, s)
  return logProp(p,f)
}
/***************************************************************/
function logProp(s, f = true) {
  let en = userProp.getProperty(s)
  if (en == null) { 
    Logger.log(' prop null caller name %s', logProp.caller.name)
    return true
  }
  let fN = flgName[parseInt(en)]
  if (f) Logger.log('en %s s = %s flgName %s', en, s, fN)
  return en
}
/***************************************************************/
// Prepare the initial info object
const initialInfo = {
  status: "",
  timestamp: "",
  message: ""
};
function doGet(e) {
  return HtmlService.createTemplateFromFile("pageIBKR")
    .evaluate()
    .setTitle('Borsa IBKR')
//    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}
/***************************************************************/
function newdoGet(e) {
  // Use the prepared initialInfo object
  const template = HtmlService.createTemplateFromFile('pageIBKR');
  initialInfo.timestamp = new Date().toISOString()
  template.info = initialInfo;
  // Simulate some preparation time
  Utilities.sleep(3000);
  // Update the info object with any additional data
  initialInfo.message = "Universal Process started ";
  return template.evaluate()
                 .setTitle('Borsa IBKR')
                 .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/****************************************************************/
// Called by client when window closes
function setWindowClosed() {
  CacheService.getScriptCache().remove(CACHE_KEY);
  Logger.log('Cache Service open Window setClosed')
}
/****************************************************************/
function includeCont(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

/****************************************************************/
function includeIBKRPageCont() { // pageIBKR.html tarafindan cagriliyor  

  let filename = ""
  Logger.log('includeIBKRPageCont')
  if (logProp('MISSING_DAYS') == 'true') {
    filename = "pageEksik-js"
  } else {
    sgnl =  !isCurrentTimeinSession(true) &&
            logProp('CLOSE_WIN') == flgEnum.FSTRT && 
            logProp('IDLE.STATE') == flgEnum.FIDLE 
    filename = sgnl ? "CloseWindow" : "pageUniv"
    if (sgnl) setProp('CLOSE_WIN', flgEnum.FEND)
  }
  Logger.log('includePage filename %s', filename)
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}  
// Create the info object
  const infoObject = {
    status: "",               // eventCount > 90 ? "end" : "running",
    timestamp: "" ,               // new Date().toISOString(),
    message: ""     // Processed ${eventCount} events in ${processingTime.toFixed(2)} seconds.
  };
/****************************************************************/
function createUniInfo(flg,  msg) {
  // Simulate processing or fetch real data
  const eventCount = Math.floor(Math.random() * 100);
  const processingTime = Math.random() * 5;
  infoObject.status = flgName[flg]
  infoObject.timestamp = new Date().toISOString()
  infoObject.message = msg
  // Log the object (optional, for debugging)
  console.log('Returning info object:', JSON.stringify(infoObject, null, 2));
  return infoObject;
};
/****************************************************************/
function processUniEvent() {
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet() 
  const ssFlg = isCurrentTimeinSession();
  let info = {}
  if (ssFlg) {
    if (logProp('DAY_START', false) == flgEnum.FSTRT) 
      prepareNewDay()
    if (logProp('POLL_VALUES',false)  == flgEnum.FPLVSTRT) {
      return getTestChangeValues() // infoObject return eder
    }
  } else  {
    if (logProp('IDLE.STATE') == flgEnum.FIDLE) {
      return createUniInfo( flgEnum.FIDLE,  ' State IDLE ')
    }
  }
  if (logProp('POLL_VALUES',false) == flgEnum.FPLVEND) {
    if (logProp('GUNSONU_PRC',false) == flgEnum.FCONT)
      return processSecurityGunsonu(flgEnum.FCONT)
  }
  Logger.log(info)
  return infoObject
}

function isWeekend() {
  const today = new Date();
  const dayOfWeek = today.getDay();
  // Local constant - only visible inside this function
  const IS_WEEKEND = (dayOfWeek === 0 || dayOfWeek === 1);
  return IS_WEEKEND;
}
/****************************************************************/
const startTime = '17:45'; // 4:30 PM
const endTime = '23:40';   // 11:30 PYNM
function isCurrentTimeinSession(rflg = false) {
  flg = !GLOBAL_WEEKEND
  if (flg) {
    const now = new Date();
    const currentTime = now.getHours() * 60 + now.getMinutes();
    // Logger.log('now %s', now)
    // Logger.log('crnt T %s h %s',currentTime, now.getHours(), now.getMinutes())
    // Convert start and end times to minutes since midnight
    const [startHour, startMinute] = startTime.split(':').map(Number);
    const [endHour, endMinute] = endTime.split(':').map(Number);
    const startMinutes = startHour * 60 + startMinute;
    // Logger.log(' sh %s sm %s sms %s ', startHour, startMinute, startMinutes)
    const endMinutes = endHour * 60 + endMinute;
    // Logger.log(' eh %s em %s ems %s', endHour, endMinute, endMinutes )
    // Handle cases where end time is on the next day
    // Logger.log('0 secenek %s',(endMinutes < startMinutes))
    // Logger.log(' 1. secenek %s ', currentTime >= startMinutes || currentTime < endMinutes)
    // Logger.log(' 2. secenek %s', currentTime >= startMinutes && currentTime < endMinutes)
    flg = (endMinutes < startMinutes)
      ? currentTime >= startMinutes || currentTime < endMinutes
      : currentTime >= startMinutes && currentTime < endMinutes;
  }
  if (!rflg) 
    checkSesTrigred(flg)
  return flg
}

function checkSesTrigred(flg) {
  if (flg) {
    if (logProp('IDLE.STATE') == flgEnum.FIDLE) {
      setProp('IDLE.STATE', flgEnum.FEND, false)
      setProp('SESION_TRIG', flgEnum.FEND, false)
      setProp('DAY_START', flgEnum.FSTRT, false)
      setProp('GUNSONU_PRC', flgEnum.FSTRT, false)
      setProp('GUN_SONUUPDATE',flgEnum.FSTRT, false)
      setPropertyActiveSheet(cas)
    }
  } else {      

    if (logProp('IDLE.STATE',false) == flgEnum.FIDLE) {
      if (logProp('DAY_START',false) == flgEnum.FSTRT) 
        prepareNewDay()
      orderSheetNewDay()
      if (logProp('GUNSONU_PRC',false) == flgEnum.FEND)
        return
    }
    if (logProp('POLL_VALUES',false) == flgEnum.FPLVSTRT) {
      setProp('SESION_TRIG', flgEnum.FSTRT)
      setStartingActiveSheet()
      gunSonuIslemleri()
    }
  }
}
function prepareNewDay() {
  Logger.log(' prepareNewDay basladi ')
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()   
  Trns.ord = Trns.sht.getSheets()[Orders]
  Trns.sht.getRange('Kebir!K1').clear({ contentsOnly: true, skipFilteredRows: true })
  Trns.sht.getSheets().forEach(function (s) {
    s.activate()
    Trns.trs = s
    var i = exclude.indexOf(s.getName());
    if (i == -1 && !s.getRange('E1').isBlank()) {
      //      Logger.log('Name %s', s.getName())
      s.getRange('E1').clear({ contentsOnly: true, skipFilteredRows: true })
      if (s.getRange(F5).getFontColor() == fntclr) {
        sFormatRange(s, 'AA26', F4)
        sFormatRange(s, 'AA26', 'F7')
      }
      if (s.getRange(F4).getFontColor() == fntclr) {
        sFormatRange(s, 'Z26', F5)
        sFormatRange(s, 'Z26', 'F8')
      }
      src = 'W27'
      if (s.getRange(src).getValue() != '') {
        s.getRange(src).activate()
        var r = s.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).getA1Notation()
        //Logger.log('r %s rSB %s', r,rSB)
        var Splt = r.split(':')
        var str = Splt[1]
        var tuple = fromA1Notation(str)
        src +=':AA'+tuple.row 
        // Logger.log('%s %s',tuple, src)
        s.getRange(src).clear({ contentsOnly: true, skipFilteredRows: true })
        //        s.getActiveRange.setBackground('#ffffff');
        assignActualYields()
      }
    }
  })
  setProp('DAY_START',flgEnum.FDAYINIT)
  setPropertyActiveSheet(cas)
  createUniInfo(flgEnum.FDAYINIT,' DAY_START PrepareNewDay Bitti ' )
  Logger.log('prepareNewDay Bitti')
}
/**************************************************************************/
function prcOpenUrl() {
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  var userProperties = PropertiesService.getUserProperties();
  setProp('MISSING_DAYS', 'false');
  openUrl()
}
/**************************************************************************/
function openMissingUrl() {
  // Create a new UrlFetchApp object.
  Logger.log('openMissingUrl icinde')
  startMissingDay()
  // Check the response status code.
  var urlFetchApp = UrlFetchApp;
  var response = null;
  try {
    response = urlFetchApp.fetch(getUrl());
  } catch (e) {
    Logger.log(' // The URL does not exist. catch e %s', e.Message )
  }
  if (response.getResponseCode() === 200) {
    // The request was successful.
    openUrl()
  } else {
    Logger.log('The request failed'+ e.message)
    // The request failed.
  }
}
/****************************************************************/
function setPropActiveSheet() {
  setWebProperty('ACTIVE_SHEET', Trns.trs.getName());
}
/****************************************************************/
/*** Open a URL in a new tab.*/
function openUrl() {
  openUrl
  var html = HtmlService.createHtmlOutput('<html><script>'
    + 'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
    + 'var a = document.createElement("a"); a.href="' + Trns.wurl + '"; a.target="_blank";'
    + 'if(document.createEvent){'
    + '  var event=document.createEvent("MouseEvents");'
    + '  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'
    + '  event.initEvent("click",true,true); a.dispatchEvent(event);'
    + '}else{ a.click() }'
    + 'close();'
    + '</script>'
    // Offer URL as clickable link in case above code fails.
    + '<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="' + Trns.wurl + '" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
    + '<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
    + '</html>')
    .setWidth(90).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening ...");
}

/**************************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Menu')
    .addItem('Openurl', 'prcOpenUrl')
    .addItem('Transaction', 'transaction')
    .addItem('ComputeSituation', 'testComputeSituation')
    .addItem('OpenMissingDaysurl', 'openMissingUrl')
    .addToUi();
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  setProp('SPREAD_SHEETID', Trns.sht.getSheetId(),false);  
  setSheetNamesArray()
}

/****************************************************************/
function getPropOpenUrl() {

  return logProp('URL_OPEN') == true;
}
/****************************************************************/
function setWebProperty(flg) { //true, false
  setProp('URL_OPEN', flg);
}
/**
 * Sekme değiştiğinde PropertiesService kullanarak durumu kontrol eder
 * ve imleci yalnızca ilk geçişte A1 hücresine çeker.
 */
function onSelectionChange(e) {
  if (!e || !e.range) return;

  const currentSheet = e.range.getSheet();
  const currentSheetName = currentSheet.getName();

  // 1. Trns pointer'ını aktif sekmeye bağla
  Trns.trs('active');

  // 2. Kalıcı hafızayı çağır (Sadece bu tabloya özel alan)
  const props = PropertiesService.getDocumentProperties();
  const lastSheet = props.getProperty("CHANGE_SHEET");

  // 3. Geçilen sekme ile hafızadaki sekme eşit değilse:
  if (lastSheet !== currentSheetName) {
    // Yeni sekme adını kalıcı olarak yaz
    props.setProperty("CHANGE_SHEET", currentSheetName);
    
    // İmleç A1'de değilse A1 hücresine çek
    if (e.range.getA1Notation() !== "A1") {
      const cellA1 = currentSheet.getRange("A1");
      currentSheet.setCurrentCell(cellA1);
      cellA1.activate();
    }
  }
}
/*
function onSelectionChange(e) {
  //et sht = SpreadsheetApp.getActiveSpreadsheet()
  // Get current sheet name and compare to previously saved sheet
  let trs = Trns.sht.getActiveSheet();
  let Name = trs.getSheetName();
  let changeName = logProp('CHANGE_SHEET');
  let pv = logProp('POLL_VALUES')
  if (changeName != Name) {
    var i = exclude.indexOf(Name)
    Logger.log('i %s POLL_VALUES %s', i, pv)
    setProp('CHANGE_SHEET', Name);
    if (i == -1) {
      if (changeName == 'Orders') onExitOrderSheet(sht, Name)
      let r = (pv == 'true') ? 'W25:AB25' : 'G7'
      Logger.log('r %s', r)
      trs.getRange(r).activate()
    }
  }
}
*/
/*****************************************************************************/
function testupdateLastTransaction() {
  Trns.StockQnty = Trns.trs.getRange('I24').getValue()
  if (Trns.StockQnty != 0) {
    Trns.Crnt = Trns.trs.getRange('B3').getValue()
    Trns.Last = Trns.trs.getRange(Trns.StockQnty > 0 ? 'L24' : 'M24').getValue()
    if (eps(Trns.Crnt, Trns.Last)) {
      Trns.tenLow = Trns.trs.getRange('F14').getValue()
      Trns.tenHgh = Trns.trs.getRange('F15').getValue()
      updateLastTransaction()
    }
  }
}

const cNames = ['B9', 'C9', 'D9', 'E9']
const bc = 6 //  Order sheet buy column no (F)
const sc = 7 //  Order sheet sell column no (G)
const dc = 13 // Order sheet date column no (M) 
const nc = 5 // Order sheet name column no (E)
/*********************************************************************************/
function procOrderSheet(e) {
  // 'syncSecurities'
  var c = e.range.getColumn()
  if (c == bc || c == sc) {
    var clr = (c == bc ? Trns.fclrs.Buy : Trns.fclrs.Sell)
    var vc = e.range.getA1Notation()
    Trns.ord.getRange(vc).setFontColor(clr)
    var r = e.range.getRow()
//    Trns.ord.getRange(r, dc).setValue(new Date())
    Trns.Name = Trns.ord.getRange(r, nc).getValue()
    updateSecurityPrice(vc, clr == Trns.fclrs.Buy ? 2 : 3, clr)
  }
}
/*********************************************************************************/
function onEdit(e) {
  var cn = e.range.getA1Notation();
  Trns.sht = e.source;
  Trns.trs = Trns.sht.getActiveSheet()
  Trns.ord = Trns.sht.getSheets()[Orders]
  Trns.Name = Trns.trs.getName()
  var i = exclude.indexOf(Trns.Name)
  LoggerAlert('onEdit i %s', i)
  if (Trns.Name == 'Orders') procOrderSheet(e)
  else if (i == -1) {
    LoggerAlert('cn %s index %s', cn, cNames.indexOf(cn))
    if (cNames.indexOf(cn) > -1) {
      secSheetUpdateOrder(e)
    }
  }
}
/*****************************************************************************/
function onExitOrderSheet(sht, Name) {
  Trns.sht = sht
  Logger.log('Name %s', Name)
  getSecNamesfromOrderSheet(true)
  Trns.or = Trns.oName.indexOf(Name) + 3
  Logger.log('Change name %s or %s', Name, Trns.or)
}

/*********************************************************************************/
function secSheetUpdateOrder(e) {
  var ce = e.range.getA1Notation()
  //Logger.log(' secSheetUpdateOrder() ce %s', ce)
  //  cNames['B9', 'D9','C9', 'E9']
  if (cNames.indexOf(ce) > -1) {
    getSecNamesfromOrderSheet()
    setOrderRow()
    //Logger.log(' Order Name %s Sht Name %s Sira No %s', Trns.ord.getName(), Trns.Name, Trns.or)
    var clr = (ce == 'B9' || ce == 'D9') ? Trns.fclrs.Buy : Trns.fclrs.Sell
    //Logger.log('ce %s  Price %s', ce, Trns.trs.getRange(ce).getValue())
    var oc = clr == Trns.fclrs.Buy ? bc : sc
    Trns.trs.getRange(ce).setFontColor(clr).setBackground('#ffffff')
    Trns.ord.getRange(Trns.or, oc).setValue(Trns.trs.getRange(ce).getValue())
      .setFontColor(clr).setBackground('#ffffff')
    Trns.ord.getRange(Trns.or, dc).setValue(new Date())
    var c = e.range.getColumn()
    if (clr == Trns.fclrs.Sell)
      wiOrderSell(c)
    else wiOrderBuy(c)
  }
}
/*****************************************************************************/
function drvGnrl() {
  Trns.sht = SpreadsheetApp.getActive()
  Trns.ord = Trns.sht.getSheets()[Orders]
  Trns.trs = Trns.sht.getActiveSheet()
  Logger.log('Order Sheet Name %s', Trns.sht.getName())
  setOrderStuff()
}
const ro = 9 // Security order row
/*****************************************************************************/
function updateSecurityPrice(vc, c, clr) {
  Trns.trs = Trns.sht.getSheetByName(Trns.Name)
  var a1 = Trns.trs.getRange(ro, c).getA1Notation()
  CopyCellDif(Trns.ord, vc, Trns.trs, a1)
  formatRange('B3', a1)
  Trns.trs.getRange(a1).setFontColor(clr)
}
/*****************************************************************************/
function GetDataArray(sht, rng) {
  var Data = []
  var i = 0;
  LoggerAlert(sprintf(' Sht Name %s\n rng %s,', sht.getName(), rng))
  sht.getRange(rng).getDisplayValues()
    .forEach(
      function (row) {
        row.forEach(
          function (cell) {
            LoggerAlert(sprintf(cell));
            Data[i++] = cell
          });
      }
    );
    //Logger.log('Data')
  return Data
}
/*****************************************************************************/
function findCellValue(sht, r, sv) {
  var i = 0;
  // Logger.log('fcv func name %s rng %s, sval %s', sht.getSheetName(), r, sv)
  o = {i,i, i}
  rng = sht.getRange(r)
  const values = rng.getDisplayValues()
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      //  Logger.log('v %s',values[i][j])
      if (values[i][j] === sv || !eps(values[i][j],sv)) {
        r = i + rng.getRow()
        c = j + rng.getColumn()
        g = sht.getRange(r, c).getA1Notation()
        o = { r, c, g }
        // Logger.log('o %s',o)
        return o
      }
    }
  }
  return o
}
/*****************************************************************************/
function copyCell(src, dst) {
  Trns.trs.getRange(src)
    .copyTo(Trns.trs.getRange(dst), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  Trns.trs.getRange(src)
    .copyTo(Trns.trs.getRange(dst), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

/*****************************************************************************/
function copyCellNormal(src, dst) {
  Trns.trs.getRange(src)
    .copyTo(Trns.trs.getRange(dst),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}
/*****************************************************************************/
function CopyCellDif(srcSheet, rsrc, dstSheet, rdst) {
  srcSheet.getRange(rsrc).copyTo(dstSheet.getRange(rdst),
      SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  srcSheet.getRange(rsrc).copyTo(dstSheet.getRange(rdst),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false)
}
/*****************************************************************************/
const fromA1Notation = (cell) => {
  const [, columnName, row] = cell.toUpperCase().match(/([A-Z]+)([0-9]+)/);
  const characters = "Z".charCodeAt() - "A".charCodeAt() + 1;
  let column = 0;
  columnName.split("").forEach((char) => {
    column *= characters;
    column += char.charCodeAt() - "A".charCodeAt() + 1;
  });
  return { row, column };
};
/*****************************************************************************/
function setSecurityCyle() {//setOrderStuff
  Ara = true
  var s = Trns.trs.getRange('D3').getFormula()
  //Logger.log('setSecurityCyle \n Name %s s %s', Trns.trs.getName(), s)
  var Splt = s.split('=')
  if (Splt.length > 1) {
    var str = Splt[1]
    var tuple = fromA1Notation(str)
    var cr = Trns.trs.getRange(tuple.row, tuple.column)
      .offset(1, 1).getA1Notation()
    var s = '=D3*(1+' + cr + ')'
    //Logger.log(' tuple %s cr %s frm %s', tuple, cr, s)
    Trns.trs.getRange('F9').setFormula(s)
    //Logger.log('get Formula %s', Trns.trs.getRange('f9').getFormula())
    s = '=B13 / ' + cr
    // Crnt  Realtiy Ratio
    Trns.trs.getRange('B15').setFormula(s)
      .setNumberFormat('#,##0.0000')
    var gr = 9
    var sr = 18
    var c = 2
    //  Order Reality Ratio    
    for (var i = 0; i < 4; i++, c++) {
      s = '=(' + Trns.trs.getRange(gr, c).getA1Notation() + ' /  D3 - 1) / ' + cr
      var Clr = i % 2 ? Trns.fclrs.Buy : Trns.fclrs.Sell
      Trns.trs.getRange(sr, c).setFormula(s)
        .setNumberFormat('#,##0.0000').setFontColor(Clr)
    }
    s = '=D3*(1+' + cr + '* B18)'
    Trns.trs.getRange('F10').setFormula(s).setNumberFormat('#,##0.00')
    s = '=D3*(1+' + cr + '* F12)'
    Trns.trs.getRange('F11').setFormula(s).setNumberFormat('#,##0.00')
    Trns.trs.getRange('F12').setValue(0.333)
  }
}
/*****************************************************************************/
function securtiyAnalys() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  Trns.sht = SpreadsheetApp.getActive()
  Trns.trs = Trns.sht.getActiveSheet();  // Transaction sheet
  Trns.Name = Trns.trs.getName()

  var response = ui.alert(
    Trns.Name,
    'Message \n\nMultiplemessagesss = 0.5444\nMultiplemessagesss = 0.5444\nMultiplemessagesss = 0.5444\nMultiplemessagesss = 0.5444\nMultiplemessagesss = 0.5444  ',
    ui.ButtonSet.OK_CANCEL);

  // Process the user's response.

  if (response == ui.Button.OK) {
    // User clicked "OK".
    if (ui.alert(' Buy or Sell Order will be given.  ', ui.ButtonSet.OK_CANCEL) == ui.Button.OK) {
      ui.alert(' Ordered\n')
    }
    ;
  } else if (response == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('User clicked the Cancel');
  } else if (response == ui.Button.CLOSE) {
    //
    ui.alert(' User clicked X in the title bar.');
  }
}
/*****************************************************************************/
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
/*****************************************************************************/
function LoggerAlert(ws) {
  if (Ara)
    Logger.log(ws)
}
/*****************************************************************************/
function GetUnvRange(Trns, Rw, Cl) {
  Trns.ord.getRange('P1').setValue(Rw)
  Trns.ord.getRange('Q1').setValue(Cl)
  //LoggerAlert(sprintf('Row %s Clmn %s returnv %s', Rw, Cl, Trns.ord.getRange('R1').getValue()))
  return Trns.ord.getRange('R1').getValue()
}
/********************************************************************************/
function AbsGetUnvRange(Trns, Rw, Cl) {
  Trns.ord.getRange('P1').setValue(Rw)
  Trns.ord.getRange('Q1').setValue(Cl)
  return Trns.ord.getRange('S1').getValue()
}


function getUrl() {
  return Trns.wurl;
}

/**
 * Global object A
 */
const A = {
  // Internal variable to store the instance
  _bInstance: null,

  /**
   * Getter for property b.
   * Logic: If the instance doesn't exist, create it. Otherwise, return the existing one.
   */
  get b() {
    if (!this._bInstance) {
      console.log("Initializing Singleton Instance for 'b'...");
      
      // Define your singleton object here
      this._bInstance = {
        timestamp: new Date().getTime(),
        id: Math.random().toString(36).substring(7),
        greet: function() {
          return "Hello from Singleton B!";
        }
      };
      
      // Optional: Prevent further modifications to the instance
      Object.freeze(this._bInstance);
    }
    return this._bInstance;
  }
};

/**
 * Test function to verify the Singleton behavior
 */
function testSingleton() {
  const instance1 = A.b;
  const instance2 = A.b;

  console.log("Instance 1 ID:", instance1.id);
  console.log("Instance 2 ID:", instance2.id);

  if (instance1 === instance2) {
    console.log("Success: Both instances are identical.");
  } else {
    console.error("Failure: Instances are different.");
  }
}

