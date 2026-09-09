/*****************************************************************************/
// gunSonu.gs
/*****************************************************************************/
const flgEnum = { FSTRT: 0, FCONT: 1, FTIMEOUT: 2, FEND: 3, FHLDY: 4, FGNSON: 5, FERGNSON: 6, FIDLE : 7, FPLVSTRT : 8, FPLVEND : 9, FDAYINIT : 10}
const flgName = ["FSTRT", "FCONT", "FTIMEOUT", "FEND","FHLDY","FGNSON", "FERGNSON", "FIDLE", "FPLVSTRT","FPLVEND","FDAYINIT"];
const GSMAX_RUNNING_TIME = 5.5 * 60 * 1000
const EPS_TIME = 60000
function testHoliday(){
  Trns.trs()
  Trns.Holiday = isHoliday()
  Logger.log('Trns.Holiday '+ Trns.Holiday)
  // processHoliDay() 

}

/**
 * Trns.trs mimarisini önce Aktif Sayfa, sonra 'Kebir' sayfası için test eder.
 */
function test_NameProperty() {
  Trns.trs('Kebir');

  // Legacy property access:
  console.log("Trns.Name:", Trns.Name); // "Kebir"

  // Native Sheet method access via Proxy:
  console.log("Trns.trs.getName():", Trns.trs.getName()); // "Kebir"
}

function test_Trns_Active_And_Hardcoded() {
  console.log("=== TRNS.TRS İKİLİ TEST BAŞLIYOR ===");

  // --- 1. AŞAMA: Aktif Sayfayı Otomatik Yakalama ---
  // Parantez içi boş bırakıldığında açık olan aktif sekmeyi alır
  Trns.trs(); 
  
  const activeName = Trns.trs.getName();
  const activeLastRow = Trns.trs.getLastRow();
  
  console.log("1. Aktif Sayfa Yakalandı:");
  console.log("   - Sayfa Adı: " + activeName);
  console.log("   - Son Satır: " + activeLastRow);

  // --- 2. AŞAMA: Hard-coded 'Kebir' Sayfasına Geçiş ---
  // İsmen 'Kebir' sekmesini seçer ve hedefi günceller
  Trns.trs('Kebir');
  
  const kebirName = Trns.trs.getName();
  const kebirCrc = Trns.trs.getRange('G1').getValue();
  
  console.log("2. 'Kebir' Sayfasına Geçildi:");
  console.log("   - Sayfa Adı: " + kebirName);
  console.log("   - G1 Değeri: " + kebirCrc);

  console.log('ORCL in adi '+Trns.trs('ORCL').getName())
  console.log("=== TÜM TESTLER BAŞARIYLA TAMAMLANDI ===");
}

/**
 * Trns yapısını ve .map() metodunu test eden fonksiyon
 */
function test_Trns_Architecture() {
  console.log("=== TEST BAŞLIYOR ===");

  // 1. TEST: Workbook (Çalışma Kitabı) Bağlantısı
  const shtName = Trns.sht.getName();
  console.log("1. Workbook Adı (Trns.sht): " + shtName);

  // 2. TEST: .map() kullanımı (Tüm sekme adlarını dizi olarak çekme)
  const allSheetNames = Trns.shts.map(sheet => sheet.getName());
  console.log("2. Tüm Sekmeler (.map sonucu):", allSheetNames);

  // 3. TEST: .filter() ve .map() birlikte (Örn: Sadece belirli sekmeleri alma)
  const activeSheets = Trns.shts
    .map(s => s.getName())
    .filter(name => !name.startsWith("Test")); // Test ile başlamayan sekmeler
  console.log("3. Filtrelenmiş Sekmeler:", activeSheets);

  // 4. TEST: Kebir!G1 Hücresini Okuma (Eski 'sht' mantığı)
  const kebirValue = Trns.sht.getRange('Kebir!G1').getValue();
  console.log("4. Kebir!G1 Değeri : " + kebirValue);

  console.log("=== TEST BAŞARIYLA TAMAMLANDI ===");
}
/*****************************************************************************/
function drvFunction() {
  const scriptCache = CacheService.getScriptCache()
  scriptCache.remove('procSec')
  
 setPropertyActiveSheet(9)

  // processSecurityGunsonu(flgEnum.FCONT)

 // gunSonuIslemleri()
 // updateKebirPage()
 // Trns.trs = Trns.sht.getActiveSheet(); 
 // Trns.Name = Trns.trs.getSheetName() 
 // updateChartBoundsFromLowHighColumns(nextcell()+1)
}

function testnChart() {
  Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  //Trns.bindTab(Trns.sht.getActiveSheet()) 
  Trns.Name = Trns.trs.getSheetName()
  Logger.log('Name %s', Trns.Name)
  updateRange( nextcell()-1)
}

function tesChart() {
  Trns.trs = Trns.sht.getActiveSheet();  
  Trns.Name = Trns.trs.getSheetName()
  //Trns.bindTab(Trns.sht.getActiveSheet()) 
  setChart()
  sr = nextcell() - 1
  Trns.trs.getRange(sr, 1).activate();
  Trns.trs.getSelection()
      .getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  rng = Trns.trs.getRange(sr, 1, Trns.trs.getActiveRange().getLastRow() -sr +1, 5).getA1Notation() 
  Logger.log('rng %s LastRow %s', rng, Trns.trs.getActiveRange().getLastRow())

//  createChart(rng,8.00, 13)
}
/*****************************M***********************************************/

function testperiod() {
  Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  Logger.log('Id = %s', Trns.sht.getName())
  //Trns.bindTab(Trns.sht.getActiveSheet()) 
  Trns.trs = Trns.sht.getActiveSheet()
  Trns.sName = Trns.trs.getName()  


  Trns.ord = Trns.sht.getSheets()[Orders]; // Order Sheet 
  Logger.log('shtName %s ordNAME %s', Trns.sName, Trns.ord.getName())
  CommonProcess()
  return
  //newPeriod( Trns.trs.getRange('Kebir!G1').getValue() )

    ac = columnLetterToNumber('Fk')
    Logger.log('ac %s ', ac)

   // ac = 156
      cn = '$' + columnToLetter(ac) +  '$15'

  Logger.log('cn %s',cn )

  for(i=ac-1; i > columnLetterToNumber('J'); --i ){
    s = '= (' + columnToLetter(i)+ '15 - '+ cn +') / ' + cn    
    Logger.log('i %s s %s',i, s)
    Trns.trs.getRange(16,i).setValue(s)
    if (newPeriod(i)) 
      cn = '$' + columnToLetter(c+1) + '$15'
   }
  
 // processLowHigh()
}


/****************************************************************************/
function CommonProcess() {
  delete nextcell.cache;
  setOrderRow()
  processLowHigh();
  newPeriod(Trns.trs.getRange('Kebir!G1').getValue() )
  processDateBlock(Trns);
  Trns.trs.getRange('E1').setValue(new Date())
  setCycle()
  setChart()
  testupdateLastTransaction()
  defineQuarters()
  //processOHLCWithDynamicMarkov()
  Logger.log('CommonProcess bitti')
}
/****************************************************************************/
function setPropertyActiveSheet(No) {
  var sNo = No.toString()
  userProp.setProperty('ACTIVE_SHEETNO', sNo)
  Logger.log('ACTIVE_SHEETNo %s',parseInt(userProp.getProperty('ACTIVE_SHEETNO')) )
}

/****************************************************************************/
function setStartingActiveSheet(shtNo = 4) {
  if (logProp('GUNSONU_PRC',false) == flgEnum.FCONT) {
    Logger.log(' Buraya Niye buraya geldik' )  
    Logger.log('Caller Name %s', setStartingActiveSheet.caller.name);
    return 
  }  
  Trns.trs = Trns.sht.getSheets()[shtNo]
  //Trns.bindTab(Trns.sht.getSheets()[shtNo])
  setPropertyActiveSheet(shtNo)
  logProp('GUNSONU_PRC', flgEnum.FSTRT,false)
  var info = {}
  info.flg = logProp('GUNSONU_PRC')
  Logger.log(' info.flg  %s ', info.flg)
}
/****************************************************************************/
function gunSonuIslemleri() {
  try {
    //Trns.sht = SpreadsheetApp.openById(Trns.ssId)
    info = {}
    info.flg = logProp('GUNSONU_PRC',false)
    if (info.flg == flgEnum.FSTRT) {
      setProp('POLL_VALUES', flgEnum.FPLVEND, false);
      Trns.pollValues = false
      enSureGunSonu()
      testHaftaSonu(Trns.sht)
      var datum = new Date()
      var d = datum.getDay()
      Logger.log(' d %s ', d)  // && d != 7
      if (d != 7 ) { 
        Trns.Holiday = isHoliday()
        if (Trns.Holiday) {
          info.flg = flgEnum.FHLDY
          Logger.log(' Eger pH bitmez is HTML in pHolidayi calistirmasi gerek')
        } else {
          getSecNamesfromOrderSheet()
          orderSnapShot()    
        }          
        setProp('GUNSONU_PRC', flgEnum.FCONT,false)   
        createUniInfo(flgEnum.FCONT, " Gunsonu Islemleri tamamlandi " )   
        return processSecurityGunsonu(flgEnum.FCONT)
      }
    }
    return info
  }
  catch (e) {
    info = {}
    info.flg = flgEnum.FERGNSON
    info.scrty = "Gun Sonu procedurunde Hata " + e.message + e.stack
    Logger.log('GunSonu err message %s Stack %s', e.message, e.stack);
    return info
  }
}
function enSureGunSonu() {
  Trns.sht.getSheets().forEach(function (s) {
    s.activate()
    var i = exclude.indexOf(s.getName());
    if (i == -1) {
      if (!s.getRange('E1').isBlank())
        //      Logger.log('Name %s', s.getName())
        s.getRange('E1').clear({ contentsOnly: true, skipFilteredRows: true })
    }
  })
}
/*****************************************************************************/
function processHoliDay() {
  if (Trns.trs.getRange('E1').isBlank()) {
    Logger.log('Name %s', Trns.trs.getName())
    var c = Trns.trs.getRange('Kebir!G1').getValue()    
    Logger.log('c '+c)
    srng = Trns.trs.getRange(13,c-1).getA1Notation()
    copyCell(srng, 'E1')
    Trns.trs.getRange(13, c).deleteCells(SpreadsheetApp.Dimension.COLUMNS);
    Trns.trs.getRange(14, c - 1, 7, 1)
      .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
    Trns.trs.getRange(2, c - 1, 5, 1)
      .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
    delete nextcell.cache;
    copyCell(srng,('A'+nextcell()))
  }
}
/*****************************************************************************/
function testHaftaSonu(ss) {
  if (ss.getRange('Kebir!G1').getValue() == 7)
    haftaSonu()
}
/*****************************************************************************/
function haftaSonu() {
  //Trns.sht = SpreadsheetApp.getActive();
  var allsheets = Trns.sht.getSheets();
  for (var s in allsheets) {
   Trns.trs = allsheets[s];
   //Trns.bindTab(allsheets[s])
   Trns.sht.setActiveSheet(Trns.trs);
    var Name = Trns.trs.getName();
    var i = exclude.indexOf(Name);
    if (i == -1) {
      Trns.trs.getRange('BA14:BZ14').activate()
      Trns.trs.getRange('BA14:BZ14').copyTo(Trns.trs.getActiveRange()
        , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      Trns.trs.getRange('BA16:BY16').activate()
      Trns.trs.getRange('BA16:BY16').copyTo(Trns.trs.getActiveRange()
        , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      Trns.trs.getRange('G13:BT22').moveTo(Trns.trs.getRange('L13'));
      Trns.trs.getRange('G2:BT6').moveTo(Trns.trs.getRange('L2'));
      generateHaftalikTarih()
    }
  };
  Trns.sht.setActiveSheet(Trns.sht.getSheets()[Kebir]);
  var Name = Trns.sht.getName();
  Logger.log('Hafta Sonu Name %s', Name)
  Trns.sht.getRange('G1').activate();
  Trns.sht.getCurrentCell().setValue('12');
}
/****************************************************************************/
function generateHaftalikTarih() {
  var c = 12
  var r = 13
  datum = new Date(Trns.trs.getRange(r, c).getValue())
  //  Logger.log(Trns.trs.getRange(r, c).getA1Notation())
  for (i = 0; i < 5; i++) {
    var n = i == 0 ? 3 : 1
    datum.setDate(datum.getDate() + n);
    c--;
    Trns.trs.getRange(r, c).setValue(datum).setNumberFormat('dd"."mm"."yyyy');
  }
}
const DAYTIME = 24 * 3600 * 1000

/****************************************************************************/
function isHoliday() {
  if (BatchType == flgBt.FHAFTASONU) return false
  var c = Trns.trs.getRange('Kebir!G1').getValue()  
  for (s = 4; s < 7; s++) {
    Trns.trs = Trns.sht.getSheets()[s];
    v = abs(Trns.trs.getRange(14, c).getValue())
    if (v > 0.000) return false
  }    
  return true
}
/****************************************************************************/
function getStartingPoint() {
  var sNo = userProp.getProperty('ACTIVE_SHEETNO')
  Logger.log(' getStartingPoint %s', sNo)
  return parseInt(sNo)
}
const et = 120000
/****************************************************************************/
function newtestElapsedTime(eTime, scriptCache) {
  ct = new Date()
  // Logger.log(' ct %s endTime %s kalan %s', ct, eTime, eTime - ct)
  // Logger.log(' return %s', ct > eTime ? flgEnum.FTIMEOUT : flgEnum.FCONT)
  if (eTime - ct < 120000) scriptCache.remove('procSec')
  return eTime - ct < 60000 ? flgEnum.FTIMEOUT : flgEnum.FCONT
}
/*****************************************************************************/
function canbeProcessed() {
  if (Trns.trs.getRange('E1').isBlank()) {
    Trns.Name = Trns.trs.getSheetName()
    Logger.log('Process edilecek')
    return true
  }
  Logger.log('Process edildi')
  return false
}
const abs = x => Number(x.toString().replace('-', ''))
/*****************************************************************************/
function releaseRange(src, dst = src) {
  // Logger.log('src %s dst %s', src, dst)
  Trns.trs.getRange(dst).activate();// formulden kurtariliyor
  Trns.trs.getRange(src).copyTo(Trns.trs.getActiveRange(),
    SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}
/*****************************************************************************/
function formatRange(src, dst) {
  Trns.trs.getRange(dst).activate()
  Trns.trs.getRange(src).copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}
/***************************************************************************/
function copyFormatRange(src, dst) {
//  Logger.log('src %s dst %s', src, Trns.trs.getActiveRange().getA1Notation())
  Trns.trs.getRange(dst).activate()
  Trns.trs.getRange(src).copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  Trns.trs.getRange(src).copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  //  Logger.log('copy format range bitti')
}
/***************************************************************************/
function setCycle() {
  if (!(Trns.trs.getRange('I11').getFormula().substring(0, 1) === '=')) {
    var sFrm = "i8*(1+J9)"
    Trns.trs.getRange('I11').setFormula(sFrm)
    formatRange('B3', "I11")
    sFrm = 'H9/J9'
    Trns.trs.getRange('I12').setFormula(sFrm)
    formatRange('B5', 'I12')
    Trns.trs.getRange('H11:H12')
      .clear({ contentsOnly: true, skipFilteredRows: true });
  }
  if (Trns.trs.getRange("I12").isBlank()) {
    //  Logger.log('setCyle Name %s', Trns.trs.getName())
    let r = abs(parseFloat(Trns.trs.getRange('H9').getValue()))
    //  Logger.log(' r %s', r)
    if (r > 0.03) {
      let r = abs(parseFloat(Trns.trs.getRange('I9').getValue()))
      if (r > 0.03) {
        var sFrm = 'I8*(1+J9)'
        Trns.trs.getRange('I11').setFormula(sFrm)
        formatRange('B3', 'I11')
        sFrm = 'H9/J9'
        Trns.trs.getRange('I12').setFormula(sFrm)
        releaseRange('J10:J12', 'J10')
        formatRange('J7:K12', 'H7')
        Trns.trs.getRange('D3').setFormula('=I8')
        formatRange('I8', 'D3')
      }
    }
  }
  if (Trns.trs.getRange('J11').getFormula().substring(0, 1) === '=') {
    if (abs(parseFloat(Trns.trs.getRange('H9').getValue())) > 0.03) {
      releaseRange('J10:J12', 'J10')
      formatRange('L7:L12', 'J7')
    }
  }
  var sFrm = "i8*(1+J9)"
  if (Trns.trs.getRange('D3').getFormula() != '=I8') {
    var c = 9
    for (; ; c++) {
      if (abs(parseFloat(Trns.trs.getRange(9, c).getValue())) > 0.03) break
    }
    Trns.trs.getRange('D3')
      .setFormula('=' + Trns.trs.getRange(8, c).getA1Notation())
    setSecurityCyle()
  }
  Logger.log('setCyle bitti')
}
/*****************************************************************************/
function paintAllBackground() {
  Trns.trs.getRange(1, 1, Trns.trs.getMaxRows(), Trns.trs.getMaxColumns()).activate();
  Trns.trs.getActiveRangeList().setBackground('#ffffff');
};
/***************************************************************************/
function singleNext(r) {
  //Logger.log(' function icinde')
  let pr = r
  for (; ;) {
    Trns.trs.getRange(r, 1).activate();
    //  Logger.log('Range %s a1 val %s', Trns.trs.getActiveRange().getA1Notation(), 
    //Trns.trs.getRange(r, 1).getValue())
    if (Trns.trs.getRange(r, 1).getValue() == 'Date') break
    Trns.trs.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    r = Trns.trs.getActiveRange().getLastRow()
    if (pr == r) break
    pr = r
  }
  return r + 1
};

/************************************************************************************************/
function setChart() {
  var sr = nextcell()
  src = Trns.trs.getRange(13, Trns.trs.getRange('Kebir!G1').getValue() - 1)
    .getA1Notation()
  let sd = new Date(Trns.trs.getRange(src).getValue())
  let cd = new Date(Trns.trs.getRange(sr, 1).getValue())
  //Logger.log('sd %s cd %s ', sd, cd)
  if (!isequalDate(sd, cd)) {
    let sr1 = sr + 1
    let s = 'A' + sr1 + ':F' + sr1
    //Logger.log('setChart()\n rng %s s %s  ', Trns.trs.getRange(s).getA1Notation(), s)
    const rng = Trns.trs.getRange(s);
    const expandedRange = rng.getMergedRanges();
    if (expandedRange.length > 0) {  // Unmerge all intersected merged blocks completely
      expandedRange.forEach(mergedRange => mergedRange.breakApart());
    } else {// Fallback if it was a standard non-intersecting merge
      rng.breakApart();
    }
    Trns.trs.getRange(s).insertCells(SpreadsheetApp.Dimension.ROWS);
    copyFormatRange(Trns.trs.getRange(sr, 1, 1, 6).getA1Notation()
      , Trns.trs.getRange(sr1, 1).getA1Notation())
    src = Trns.trs.getRange(13, Trns.trs.getRange('Kebir!G1').getValue() - 1).getA1Notation()
    dst = Trns.trs.getRange(sr, 1).getA1Notation()
   // Logger.log("src %s dst %s",src,dst)
    copyCell(src, dst)
/*     s ='A' + sr
    s = '= TEXT(' + s +',"dd,mm,yy")'
    src = 'B' +sr
    Trns.trs.getRange(src).setValue(s)
    updateRange(sr-1)
    updateChartBoundsFromLowHighColumns(sr-1)*/
  }
  Logger.log('setChart Bitti ')
}
function updateRange(sr) {
    ser = sr+90
    const lastRow = Trns.trs.getLastRow();
    if (lastRow < ser) ser  = lastRow
    s = 'B' + sr + ':F' + ser
    // Get all charts on the sheet
    var charts = Trns.trs.getCharts();
    // Select the first chart (change index if needed)
    var chart = charts[0];
    // Define the new data range 
    var newRange = Trns.trs.getRange(s);
    // Modify the chart's data range
    var newChart = chart.modify()
      .clearRanges()
      .addRange(newRange)
      .build();

    // Update the chart on the sheet
    Trns.trs.updateChart(newChart);
}
/****************************************************************************/
function processSecurityGunsonu(flg) {

  const scriptCache = CacheService.getScriptCache()
  //Trns.sht = SpreadsheetApp.openById(Trns.ssId)
  if (logProp('GUNSONU_PRC') == flgEnum.FEND) {
    return  createUniInfo(flgEnum.FEND, ' ProcessSecurityGunSonu finished')
  }
  if (scriptCache.get('procSec')) {
    Logger.log('scriptCache exist')
    scriptCache.remove('procSec')
  }

  //Ara = true
  try {
    info ={}
    scriptCache.put('procSec', 'exist', 400000)
    Logger.log('ScritpCache init Edildi %s ', scriptCache.get('procSec'))
    var sTime = new Date()
    let eTime = new Date(sTime.getTime() + GSMAX_RUNNING_TIME)
    Logger.log('Baslangic sTime %s eTime %s', sTime, eTime)
    var lock = LockService.getScriptLock()
    // Logger.log(' has lock %s', lock.haslock)
    // flg = true terminate false continue  
    infoObject.status = "lock Active " 
    if (!lock.tryLock(60000)) return infoObject
    Trns.ord = Trns.sht.getSheets()[Orders]; // Order Sheet 
    getSecNamesfromOrderSheet()
    var alls = Trns.sht.getSheets()
    infoObject.status = " Contine"
    for (var s = getStartingPoint(); s < alls.length; s++) {
      Trns.trs = alls[s];
      //Trns.bindTab(alls[s])
      //  Logger.log(" Sheet Name %s ", Trns.trs.getName())
      var i = exclude.indexOf(Trns.trs.getName());
      if (i == -1) {
        Trns.sht.setActiveSheet(Trns.trs)
        delete nextcell.cache;
        infoObject.message = Trns.trs.getSheetName()
        Logger.log(" info scrty name %s Active sheet name %s", info.scrty, Trns.sht.getActiveSheet().getName())
        if (newtestElapsedTime(eTime, scriptCache) == flgEnum.FTIMEOUT) {
          Logger.log(' Time Out %s', flgEnum.FTIMEOUT)
          infoObject.status = flgEnum.FTIMEOUT
          SpreadsheetApp.flush()
          lock.releaseLock()
          scriptCache.remove('procSec')
          Logger.log('ScriptCache remove edildi %s', scriptCache.get('procSec'))
          return info
        }
        if (canbeProcessed()) {
          if (Trns.Holiday) {
            processHoliDay()
            continue
          } else
            CommonProcess()
        }
        if (s+1 <= alls.length)
          setPropertyActiveSheet(s + 1)
      }
      //  Logger.log("info flg %s Name %s", info.flg, info.scrty)
    }
    if (checkifnoSet('GUN_SONUUPDATE', flgEnum.FEND) || checkMissingDay()) 
      updateKebirPage();

    scriptCache.remove('procSec')
    return info  
  }

  
  catch (e) {
    scriptCache.remove('procSec')
    lock.releaseLock()
    info.flg = flgEnum.FCONT
    info.scrty += 'processSecurityGunsonu Exception ' + e.message + ' Stack ; '+ e.stack
    Logger.log('processSecurityGunsonu Exception : %s\n Stack : %s', e.message, e.stack)
    s = getStartingPoint() + 1
    Logger.log('starting point %s info %s', s, info)
    if (s < Trns.sht.getSheets())
      setPropertyActiveSheet(s)

  }
  return info
}

function checkMissingDay() {
  Logger.log('GUN_SONUUPDATE %s FEND val %s', userProp.getProperty('GUN_SONUUPDATE'), flgEnum.FEND)
  return true
}
function updateTrh(c) {
  src = Trns.trs.getRange(13, c).getA1Notation()
  // Logger.log('Tarih update edilir')
  copyCell(src,'H7')
  updateDayDif()
}
function newPeriod(crc) {// previos column
  pdt = Trns.trs.getRange('H7').getDisplayValue() // Period date
  pva = Trns.trs.getRange('H8').getValue() // Period value
  prt = Trns.trs.getRange('H9').getValue() // Period ratio
  pvc = crc + 1 // Current column
  // Logger.log('Privious coluupdateTrhmn %s Current Column %s', pvc, crc)
  // Logger.log('Period Date %s value %s Ratio %s ', pdt, pva, prt)
  lgd = Trns.trs.getRange(13, crc).getDisplayValue() // Crnt gun date
  lgo = Trns.trs.getRange(14, crc).getValue()       // Crnt Gun oran
  lva = Trns.trs.getRange(15, crc).getValue()
  lto = Trns.trs.getRange(16, crc).getValue()        // Crnt Topla
  orn = (lva - pva) / pva
  if (prt > 0) {
    if (lgo > 0) { // Last Gunluk oran pozitif
      updateTrh(crc)
      if (lva > pva) {
        src = Trns.trs.getRange(15,crc,2,1).getA1Notation()
        copyFormatRange(src,'H8')
        // Logger.log('Periodun valuesu ve toplam orani update edilir')
      } else {
        // Logger.log(' last value  period val dan kucuk bir sey yapilmaz')
      }
    } else {
      if (orn < -0.03) {
        // Logger.log('Yeni oran %sYeni Period eklenir ', orn)
        insertNewPeriod(crc,pva) 
        return     
      } else {
        // Logger.log('Yeni oran 0.03 ten kucuk sadece Periodun Tarihi update edilir ')
        updateTrh(crc)
      }
    }
  }  else { // Period ratio negatif
    if (lgo < 0) { //lst gunluk oran negatif
      updateTrh(crc) // Current column
      if (lva < pva) { // Last value Periodun valuesundan kucuk
        src = Trns.trs.getRange(15,crc,2,1).getA1Notation()
        copyFormatRange(src,'H8')
        // Logger.log('Periodun valuesu ve toplam orani update edilir')
      } else {
          // Logger.log('orn 0.03 kucuk sadece Periodun Tarihi degisir')
      }
    } else {
      if (orn > 0.03) {
        // Logger.log('Yeni oran %sYeni Period eklenir ', orn)
        insertNewPeriod(crc,pva) 
        return     
      } else {
        // Logger.log('Yeni oran 0.03 ten buyuk sadece Periodun Tarihi update edilir ')
        updateTrh(crc)
      }      
    }
  } 
  return false
}
function columnLetterToNumber(columnLetter) {
  columnLetter = columnLetter.toUpperCase();
  let column = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, columnLetter.length - i - 1);
  }
  return column;
}
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
function insertNewPeriod(c,pva) {
 //  Logger.log('insert new Period c = %s pva %s',c, pva)
  let sc = c
  for(;;) {
    //Logger.log(' v = %s',Trns.trs.getRange(15,c).getDisplayValue())
    if(Trns.trs.getRange(15,c).getDisplayValue() == pva) break;
    ++c 
    //Logger.log('for icinde c %s',c)
  }
  updateTrh(c)
  cn = '$' + columnToLetter(c) + '$15'
  // Logger.log('c %s',cn)
  while (c > sc) {
    --c
    s = '= (' + columnToLetter(c)+ '15 - '+ cn +') / ' + cn
    // Logger.log('s %s',s)
    Trns.trs.getRange(16,c).setValue(s)
  }
  Trns.trs.getRange('H7:H12').insertCells(SpreadsheetApp.Dimension.COLUMNS);
  // Logger.log('sc %s',c)
  updateTrh(c)
  src = Trns.trs.getRange(15,sc,2,1).getA1Notation()
  copyCell(src,'H8')
  updateDayDif()
}

/****************************************************************************/
function fnRatio(a, b) {
  var r = (a - b) / b
  if (r < 0) r *= - 1
  return r
}
/*****************************************************************************/
function getRatio(s) {
  var rtio = Trns.trs.getRange(s).getValue()
  if (rtio < 0) rtio *= -1
  return rtio
}
/*********************--******************************************************/
function updateDailyRatio() {
  updateTarih()
  var src = Trns.vrng.offset(1, 1).getA1Notation()
  //  LoggerAlert(sprintf('updateDailyRatio src %s dsrc %s ' , src, Trns.vrng.offset(1, 0).getA1Notation()))
  var dst = src + ':' + Trns.vrng.offset(1, 0).getA1Notation()
  Trns.trs.getRange(src).autoFill(Trns.trs.getRange(dst)
    , SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}
/*****************************************************************************/
function canbeWorth() {
  var dcr = fnRatio(Trns.nval, Trns.trs.getRange('H8').getValue())
  var cr = getRatio('H9')
  //  LoggerAlert(sprintf(' Canbe worth dcr %s cr %s', dcr, cr))
  if (dcr < 0.015) {
    var ncr = fnRatio(Trns.nval, Trns.trs.getRange('H8').getValue())
    //    LoggerAlert(sprintf('ncr %s cr %s', ncr, cr))
    if (cr > ncr) updateDailyRatio()
    else updateH()
    return false
  } else {
    // Eger bir onceki  oran % 3.5 kucuk ise devam saglaniyor
    //   LoggerAlert(sprintf(' Previous Oran %s pv %s', getRatio('H9'), getRatio('I9')))
    var pr = getRatio('H9')
    if (pr < 0.035) {
      pv = getRatio('I9')
      if (dcr > pv) {
        var or = getRatio('J9')
        if (or < 0.035)
          shiftLeft('J7:J12', 'H7')
        else shiftLeft('I7:I12', 'H7')
        updateH()
        return false
      }
    }
  }
  return true;
}

/*********************--******************************************************/
function predictionBuySell() {
  //Ara = true
  var data = Trns.trs.getRange('H8:J8').getValues();
  Trns.nval = Trns.vrng.getValue()
  Logger.log('predictionBuySell Trn.nval %s',Trns.nval)
  if (testSimple(data)) {
    //    LoggerAlert(sprintf(' Test Simple true UpdateH calisacak'))
    updateH()
    return
  }
  if (Trns.nval > data[0][0]) {
    //  LoggerAlert(sprintf('0 nval > data[0][0]  %s  %s', Trns.nval, data[0][0]))
    if (Trns.nval > data[0][1]) {
      //   LoggerAlert(sprintf('1 nval > data[0][1]  %s  %s', Trns.nval, data[0][1]))
      if (data[0][0] < data[0][1]) {
        restoreMinMax(data);
      } else {
        //LoggerAlert(sprintf('3 data[0][0] > data[0][1]  %s  %s', data[0][0], data[0][1]))
        updateH()
      }
    } else {
      //LoggerAlert(sprintf('4 else data[0][0] < nval < data[0][1]  %s  %s  %s', data[0][0], Trns.nval, data[0][1]))
      moveNext()
    }
  } else {
    //  LoggerAlert(sprintf('5  nval < data[0][0]  %s  %s', Trns.nval, data[0][1]))
    if (Trns.nval > data[0][1]) moveNext()
    else {
      if (data[0][0] > data[0][1]) {
        restoreMinMax(data);
      } else {
        //  LoggerAlert(sprintf('5A  UpdateH or Date %s'))
        var r = fnRatio(Trns.nval, Trns.trs.getRange('I8').getValue())
        var Orn = Trns.trs.getRange('H9').getValue()
        if (Orn < 0) Orn *= -1;
        if (r < 0.015) {
          var nr = fnRatio(Trns.nval, Trns.trs.getRange('H8').getValue())
          if (Orn > nr) {
            // LoggerAlert(sprintf('5A Canbe worth %s ', nr))
            updateDailyRatio()
            return
          }
        }
        updateH()
      }
    }
  }
};
/*************************************************************************** */
function restoreMinMax(data) {
  var dcr = fnRatio(Trns.nval, Trns.trs.getRange('H8').getValue())
  var cr = getRatio('H9')
  // LoggerAlert(sprintf('Restore Min Max dcr %s  cr %s', dcr, cr))
  if (cr > 0.035) moveNext()
  else {
    var flg = false
    if (data[0][0] < data[0][1] && Trns.nval < data[0][1]) flg = true
    if (data[0][0] > data[0][1] && Trns.nval > data[0][1]) flg = true
    if (flg) {
      shiftLeft('I7:I12', 'H7')
      updateH()
    } else moveNext()
  }
}
/************************************************************************** */
function updateDayDif() {
  //Updating Day Differece
  var lDate = new Date(Trns.trs.getRange('H7').getValue())
  var rDate = new Date(Trns.trs.getRange('I7').getValue());
  var diffInDays = Math.floor((lDate.getTime() - rDate.getTime())
    / (DAYTIME));
 // Logger.log('lD %s rD %s diff %s',lDate,rDate,diffInDays)
  Trns.trs.getRange('H10').setValue(diffInDays)
}
/*****************************************************************************/
function updateRatio() {
  var Ratio = (Trns.trs.getRange('H8').getValue()
    - Trns.trs.getRange('I8').getValue())
    / Trns.trs.getRange('I8').getValue()
  Trns.trs.getRange('H9').setValue(Ratio)
    .setFontColor(Trns.aClr)
    .setNumberFormat('#,##0.0000')
  Ratio = A1Ratio('H8', 'I8')
  var src = Trns.vrng.offset(1, 0).getA1Notation()
  //  LoggerAlert(sprintf('Ratio %s  src %s', Ratio, src))
}
/*****************************************************************************/
function restoreRatio() {
  var i = Trns.trs.getRange('H10').getValue()
  var Date = Trns.trs.getRange('I7').getDisplayValue()
  //  LoggerAlert(sprintf('restoreRatio \n Date %s vrng %s', Date, Trns.vrng.getA1Notation()))
  var r = Trns.vrng.offset(-2, 0).getA1Notation()
    + ':' + Trns.vrng.offset(-2, i).getA1Notation()
  var DateA = Trns.trs.getRange(r).getDisplayValues()
  var e = DateA[0].indexOf(Date)
  //  LoggerAlert(sprintf('DateA %s Date %s end %s', DateA, Date, e))
  var srng = Trns.vrng.offset(1, 0).getA1Notation()
  var erng = Trns.vrng.offset(1, e - 1).getA1Notation()
  var s = srng + ':' + erng
  Trns.trs.getRange(erng).activate();
  Trns.trs.getActiveRange().autoFill(Trns.trs.getRange(s),
    SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  //  LoggerAlert(sprintf('restoreRatio i %s r %s  srng %s erng %s s %s', i, r, erng, s))
}
/*****************************************************************************/
function restoreReality() {
  Trns.trs.getRange('H12')
    .clear({ contentsOnly: true, skipFilteredRows: true });
  c = 9
  cn = 10
  var sFrm = 'H9/J9'
  Trns.trs.getRange('I12').setFormula(sFrm)
  var rRng = Trns.trs.getRange('I12')
    .getNextDataCell(SpreadsheetApp.Direction.NEXT).getA1Notation()
  Trns.trs.getRange('I12').copyTo(Trns.trs.getRange(rRng))
  Trns.trs.getRange(rRng).setNumberFormat('##,0.0000')
}
/*****************************************************************************/
function shiftLeft(Rng, dRng) {
  Trns.trs.getRange(Rng).activate();
  Trns.trs.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  var Rng = Trns.trs.getActiveRange().getA1Notation()
  Trns.trs.getRange(dRng).activate();
  Trns.trs.getRange(Rng).moveTo(Trns.trs.getActiveRange());
  data = Trns.trs.getRange('H8:J8').getValues()
  setColor()
  updateDayDif()
  updateRatio();
  // LoggerAlert(sprintf('Shift Left data %s vRng %s ', data, Trns.vrng.getA1Notation()))
  restoreRatio()
  restoreReality()
}
/*****************************************************************************/
function updateTarih() {
  Trns.trs.getRange('H7').activate();
  Logger.log('updateTarih srct %s', Trns.vrng.getA1Notation())
  Trns.vrng.offset(-2, 0).copyTo(Trns.trs.getActiveRange(),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  setColor()
  Trns.trs.getRange('H7').setFontColor(Trns.Clr)
  Logger.log('H7 color %s', Trns.Clr)
  updateDayDif()
}
/*****************************************************************************/
function setColor() {
  var data = Trns.trs.getRange('H8:J8').getValues()
  Trns.Clr = (data[0][0] > data[0][1]) ? Trns.fclrs.High : Trns.fclrs.Low
  Trns.aClr = (data[0][0] > data[0][1]) ? Trns.fclrs.Min : Trns.fclrs.Max
}
/******************************************************************************/
function updateH() {
  //  LoggerAlert(sprintf('UpdateH'))
  updateTarih()
  // Updating Value
  Trns.trs.getRange('H8').activate();
  Trns.vrng.copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  data = Trns.trs.getRange('H8:J8').getValues();
  setColor()
  formatRange('B3', 'H8') //.setNumberFormat('#,##0.00')   
  Trns.trs.getRange('H8').setFontColor(Trns.Clr)
  return
  updateRatio()
  updateDailyRatio()
  //  LoggerAlert(sprintf('6 Value %s RwCl %s', Trns.vrng.getValue(), Trns.vrng.getA1Notation()))
  var pr = getRatio('I9')
  if (pr < 0.035) {
    var cr = getRatio('H9')
    if (cr > pr) {
      shiftLeft('J7:J12', 'H7')
      updateH()
    }
  }
}
/******************************************************************************/
 function A1Ratio(ra, rb) {
  var a = Trns.trs.getRange(ra).getValue()
  var b = Trns.trs.getRange(rb).getValue()
  var r = (a - b) / b
  if (r < 0) r *= - 1
  return r
}
/*****************************************************************************/
function checkMaxMax() {
  // MAX / MIN artis yuzdesi
  var y = Trns.trs.getRange(Trns.vrng.offset(-1, 0).getA1Notation()).getValue()
  //  LoggerAlert(sprintf('CheckMaxMax y %s ', y))
  //  LoggerAlert(sprintf('H1 val %s', Trns.trs.getRange('H1').getValue()))
  if (y < Trns.trs.getRange('H1').getValue()) {
    copyCell(Trns.vrng.offset(-2, 0).getA1Notation(), 'G1')
    copyCell(Trns.vrng.offset(-1, 0).getA1Notation(), 'H1')
  }
  if (y > Trns.trs.getRange('K1').getValue()) {
    copyCell(Trns.vrng.offset(-2, 0).getA1Notation(), 'J1')
    copyCell(Trns.vrng.offset(-1, 0).getA1Notation(), 'K1')
  }
  // MAX/MIN degeri
  v = Trns.vrng.getValue()
  //  LoggerAlert(sprintf(' val %s ', v))
  if (v < Trns.trs.getRange('O1').getValue()) {
    copyCell(Trns.vrng.offset(-2, 0).getA1Notation(), 'N1')
    Trns.trs.getRange('O1').setValue(v)
    formatRange('B3', 'O1') //.setNumberFormat('#,##0.00')
  }
  if (v > Trns.trs.getRange('R1').getValue()) {
    copyCell(Trns.vrng.offset(-2, 0).getA1Notation(), 'Q1')
    Trns.trs.getRange('R1').setValue(v)
    formatRange('B3', 'R1')
  }
}

/*****************************************************************************/
function processDateBlock(Trns) {
  // Day Cell Adress
  var Rng = Trns.trs.getRange('Kebir!H1').getValue()
  Trns.trs.getRange(Rng).activate()
  Trns.trs.getActiveRange().copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  // Logger.log('rng %s', Trns.trs.getActiveRange().getA1Notation())
  formatRange('B3', Trns.trs.getActiveRange().getA1Notation())
  var cell = Trns.trs.getCurrentCell();
  Trns.trs.getCurrentCell().offset(0, -1).activate();
  Trns.trs.getCurrentCell().setFormula('=B3');
  // Current  Previos Cell
  var cv = cell.getValue();
  Trns.trs.getRange('B4').activate();
  var pv = Trns.trs.getCurrentCell().getValue();
  Trns.trs.getRange('B3').copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  Trns.trs.getCurrentCell().offset(-1, 0, 2, 1).activate();
  let clr = (cv > pv) ? Trns.fclrs.High : Trns.fclrs.Low      // Green
  Trns.trs.getActiveRangeList().setFontColor(clr);
  // Ratio Cell      
  cell.offset(-1, -1).activate();
  //  Logger.log('cell %s', Trns.trs.getActiveRange().getA1Notation())
  Trns.trs.getCurrentCell().offset(0, 1).copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  Trns.trs.getCurrentCell().offset(2, 0).activate();
  Trns.trs.getCurrentCell().offset(0, 1).copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  Trns.trs.getCurrentCell().offset(1, 0).activate();
  // if (!Trns.trs.getRange('A24').isBlank())
  //   Trns.transactionRatios()
  Logger.log('Process Date Block Bitti')
}

/*****************************************************************************/
function transactionRatios () {
  var RTrns = AbsGetUnvRange(Trns, 18, Trns.trs.getRange('Kebir!G1').getValue())
  //  Logger.log('transactionRatios RTrns %s', Trns.trs.getRange(RTrns).getA1Notation())
  if (Trns.trs.getRange(RTrns).isBlank())
    Trns.trs.getCurrentCell().offset(0, 1).copyTo(Trns.trs.getActiveRange()
      , SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  else {
    /*****************Transaction Ratio */
    var c = Trns.trs.getRange('Kebir!G1').getValue()
    Rng = Trns.trs.getRange(15, c - 1).getA1Notation()
    var sFrm = '=(' + Rng + '-' + RTrns + ')/' + RTrns
    Trns.trs.getCurrentCell().setFormula(sFrm);
    Trns.trs.getCurrentCell().setNumberFormat('#,##0.0000')
    var Clr = Trns.trs.getRange(RTrns).getFontColor()
    Trns.trs.getCurrentCell().setFontColor(Clr)
  }
  // Logger.log('Transaction Ratios bitti')
}
/*****************************************************************************/
function updateSonucSatiri() {
  // Sonuc Satiri Update ediliyor
  Trns.trs.getRange('B21').activate();
  Trns.trs.getRange('B24').copyTo(Trns.trs.getActiveRange()
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}
/*****************************************************************************/
function updateKebirPage() {
  Trns.kbr = Trns.sht.getSheets()[Kebir];
  Trns.sht.setActiveSheet(Trns.kbr);
  var cv = BatchType == flgBt.FGUNSONU ? parseInt(Trns.kbr.getRange('G1').getValue()) - 1 : 12
  Trns.kbr.getRange('G1').setValue(cv);
  setProp("GUNSONU_PRC",flgEnum.FEND,false)
  setProp('POLL_VALUES', flgEnum.FPLVSTRT,false)  
  setProp('IDLE.STATE', flgEnum.FIDLE,false)
  setProp('SESION_TRIG', flgEnum.FSTRT,false)
  setProp('DAY_START',flgEnum.FSTRT,false)
  setProp('CLOSE_WIN', flgEnum.FSTRT,false)
  setProp('UPDATE_KEBIR_PAGE',flgEnum.FEND, true)
  Logger.log('updateKebirPage procSecGunSonu bitti  %s', updateKebirPage.caller.name);
}
/*****************************************************************************/
function checkColor(r, c, v, t) { // row, col,val, {Min or maX}
  try {
  //  Logger.log('r %s c %s v %s t %s', r,c,v,t )
    src = Trns.trs.getRange(r, c, 1, 10).getA1Notation()
  //  Logger.log('src %s v %s',src, v)
    let o = findCellValue(Trns.trs, src, v)
   
    
    if (o.r != null && o.c != null) {
      bckc = Trns.trs.getRange(o.r, o.c).getBackgroundColor()
      clr = (t == "M") ? bckclrY : bckclrK
      rng = Trns.trs.getRange(o.r, o.c).getA1Notation()
    //  Logger.log(' rng %s obj %s bckc %s clr %s', rng, o, bckc, clr)
      if (bckc != clr)
        Trns.trs.getRange(o.r, o.c).setFontColor(fntclr).setBackground(clr)
    } else {
       Logger.log('CHECK COLOR Error r %s c %s v %s t %s o %s', r,c,v,t, o)
    }
  }
  catch(e) {
    Logger.log('Excp checkColor err message %s Stack %s', e.message, e.stack);
  }
}
//        s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).activate();
//        s.getActiveRangeListc().setBackground('#ffffff');

/*****************************************************************************/
function processLowHigh() {
  let c = Trns.trs.getRange('Kebir!G1').getValue()
  let m = c - 1
  let dst = Trns.trs.getRange(2, m).getA1Notation()
  //Logger.log('c %s m %s dst %s', c, m, dst)
  if (Trns.trs.getRange(dst).isBlank()) {
    copyCellNormal('F2:F6', dst)
    dst = Trns.trs.getRange(2, c).getA1Notation()
    copyFormatRange('F2:F6', dst)
    dst = Trns.trs.getRange(3, c).getA1Notation()
    let sFrm = '=(F3 -' + dst + ')/' + dst
    Trns.trs.getRange('D4').setFormula(sFrm) // Previos Day open dif
    dst = Trns.trs.getRange(rlw, m, 1, 10).getA1Notation()
    sFrm = '=min(' + dst + ')'
    checkColor(rlw,c, Trns.trs.getRange('F7').getDisplayValue(), "M")
    Trns.trs.getRange('F7').setFormula(sFrm) // Son on gunun Lowu
    sFrm = '=Average(' + dst + ')'
    Trns.trs.getRange('F16').setFormula(sFrm)  // Lowlarin ortalamasi  
    dst = Trns.trs.getRange(rhg, c - 1, 1, 10).getA1Notation()
    sFrm = 'max(' + dst + ')'
    Trns.trs.getRange('F8').setFormula(sFrm)
    sFrm = '=Average(' + dst + ')'
    checkColor(rhg,c,Trns.trs.getRange('F8').getDisplayValue(), "X")
    Trns.trs.getRange('F17').setFormula(sFrm)
    formatRange('B3', 'F7:F8')
    formatRange('B3', 'F16:F17')
    // verify today Loww and High
    setCellHighColor(c)
    setCellLowColor(c)
  }
  Logger.log('processLowHigh finished')
}
/*****************************************************************************/
function processDateColor(ss, as) {
  ss.setActiveSheet(as);
  var cv = as.getRange('B3').getValue()
  var pv = as.getRange('B4').getValue()
  var Clr = (cv < pv) ? Trns.fclrs.Low : Trns.fclrs.High
  as.getRange(as.getRange('Kebir!H1').getValue()).activate();
  as.getCurrentCell().offset(-2, 0, 3, 1).activate();
  as.getActiveRangeList().setFontColor(Clr);
  as.getRange('B3').setFontColor(Clr)
}
