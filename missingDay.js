

// missingDay
// Copy tabinda F1'e  Baslangic tarihi 09.11.2023 G1 tabina da bitis tarihi	12.11.2023 gibi yazilir.
// Custom Menuden missingOpenUrl calistirilir.
// Son olarak restoreGoogleFinance() functioni cagrilir.

const cu = 21
const cVal = 'Copy!A10'

const msdState = { stOHLC: 0, stGNSN : 1, stEND : 2 }

function testing() {
  restoreGoogleFinance()
  return
  CacheService.getScriptCache().remove('procSec')  
  /*    
  userProp.setProperty('MISSING_DAYS', 'true')
  setPropertyActiveSheet(11) 
  userProp.setProperty('MISDAY_PRC', msdState.stGNSN);  

  userProp.setProperty('MISDAY_PRC',msdState.stOHLC)
  userProp.setProperty('OHLC_STATE', flgEnum.FCONT)  


  userProp.setProperty('MISSING_DAYS', 'false')
  //checkaNewMissing()
  //setStartingActiveSheet()
  //usrProperties.setProperty('MISDAY_PRC', )
  //startMissingDay()  
  */
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  procMisDay()
}

function startMissingDay() {
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('MISSING_DAYS', 'true');
  
  clearE1()
  setPropertyActiveSheet(4) // 
  userProperties.setProperty('MISDAY_PRC', msdState.stOHLC)
  userProperties.setProperty('OHLC_STATE', flgEnum.FCONT)
}
function checkaNewMissing() {
  var info = {}
  let s = Trns.sht.getRange(cVal).getValue()
  Logger.log('A10 val %s', s)
  if (s == '') {
    restoreGoogleFinance()
    info.flg = flgEnum.FEND
    info.scrty = "<br> flg = FEND interval end signaled "
    return info
  } else { 
    Trns.sht.getRange(cVal).copyTo(Trns.sht.getRange('Copy!F1')
    , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)

    Trns.trs.getRange(cVal).copyTo(Trns.sht.getRange('Copy!F1')
    , SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    startMissingDay()
    info.flg = flgEnum.FCONT
    info.scrty = "<br> A new day started "
  }
  Logger.log(' starting value %s', Trns.sht.getRange('Copy!F1').getValue())
  return info
}
function misHoliday() {
  v = Trns.sht.getSheets()[1].getRange('Copy!A9').getValue()
  crDate = new Date(v)
  Logger.log('v %s crDate %s day %s', v, crDate, crDate.getDay())
  var c = Trns.sht.getRange('Kebir!G1').getValue()
  Trns.trs = Trns.sht.getSheets()[4]
  v = Trns.trs.getRange(13, c).getValue()
  pgDate = new Date(v)
  Logger.log('v %s pgDate %s day %s', v, pgDate, pgDate.getDay())
  flg = crDate.getDay() == pgDate.getDay()
  Logger.log('flg %s', flg)
  if (flg) return false
  gunSonuIslemleri()
  processSecurityGunsonu(flgEnum.FCONT)
  return true
}
function procMisDay() {
  var info = {}
  var usrProperties = PropertiesService.getUserProperties()
  flg = usrProperties.getProperty('MISSING_DAYS') 
  if (flg == 'false') {
    flg = getStockPriceWithErrorHandling()
    if (flg === flgEnum.FCONT) {
      info = checkaNewMissing()
      if (info.flg == flgEnum.END) return
    }
  }
  let smsd = usrProperties.getProperty('MISDAY_PRC')
  Logger.log('MISDAY_PRC %s', smsd)
  if (smsd == msdState.stOHLC) {
      info.flg = processOHLC()
      if (info.flg == flgEnum.FCONT) return  
      userProp.setProperty('MISDAY_PRC',msdState.stGNSN)
  }
  if (smsd == msdState.stGNSN) {
      info.flg = processSecurityGunsonu(flgEnum.FCONT)
      if (info.flg == flgEnum.FCONT) return
      usrProperties.setProperty('MISDAY_PRC', msdState.stEND);
      smsd = msdState.stEND
  }
  if (smsd == msdState.stEND) {
    info = checkaNewMissing()
    if (info.flg == flgEnum.FEND) {
      usrProperties.setProperty('MISSING_DAYS', 'false')
      Logger.log('Page eksike signal end yollaniyior ')
      return flgEnum.FEND
    }
  }
}

function processOHLC() {
  info = {}
  var usrProperties = PropertiesService.getUserProperties()
  if (usrProperties.getProperty('OHLC_STATE') == flgEnum.FCONT) {
    const scriptCache = CacheService.getScriptCache()
    if (scriptCache.get('procSec')) {
      //  Logger.log('scriptCache exist')
      scriptCache.remove('procSec')
      info.flg = flgEnum.FCONT
      info.scrty = ' scriptCache exsist remove edildi'
      return info
    }
    scriptCache.put('procSec', 'exist', 400000)
    Logger.log('ScritpCache init Edildi %s ', scriptCache.get('procSec'))
    var sTime = new Date()
    let eTime = new Date(sTime.getTime() + GSMAX_RUNNING_TIME)
    Trns.sco = Trns.sht.getSheets()[1]
    var alls = Trns.sht.getSheets()
    Logger.log('starting index %s', getStartingPoint())
    info.flg = flgEnum.FCONT
    for (var s = getStartingPoint(); s < alls.length; s++) {
      Trns.trs = alls[s];
      var i = exclude.indexOf(Trns.trs.getName());
      if (i == -1) {
        Trns.sht.setActiveSheet(Trns.trs)
        info.scrty = Trns.trs.getSheetName()
        Logger.log(" info scrty name %s Active sheet name %s", info.scrty, Trns.sht.getActiveSheet().getName())
        if (newtestElapsedTime(eTime, scriptCache) == flgEnum.FTIMEOUT) {
          Logger.log('Line 82 Time Out %s', flgEnum.FTIMEOUT)
          scriptCache.remove('procSec')
          info.flg = flgEnum.FCONT
          SpreadsheetApp.flush()
          scriptCache.remove('procSec')
          //    Logger.log('Line 87 ScriptCache remove edildi %s', scriptCache.get('procSec'))
          info.scrty = "Time out info.flg" + info.flg + ' ' + Trns.trs.getName() + 'wil be executed'
          return info
        }
        setOHLCValues()
        setPropertyActiveSheet(s + 1)
      }
    }
    usrProperties.setProperty('OHLC_STATE' , flgEnum.FEND)
  }    
  info.flg = usrProperties.getProperty('MISDAY_PRC')
  checkSesionFinish()
  Logger.log('checkSesion da prop GUN_SONUUPDATE val %s', userProp.getProperty('GUN_SONUUPDATE'))
  setStartingActiveSheet()
  info = gunSonuIslemleri()
  usrProperties.setProperty('MISDAY_PRC', msdState.stGNSN);

  return info
}
const dcl = ['F3', 'F4', 'F5', 'B3']
function setOHLCValues() {
  Trns.Name = Trns.trs.getName()
  //Logger.log('Name %s', Trns.Name)
  Trns.trs.getRange('E1').clear({ contentsOnly: true, skipFilteredRows: true });
  //Logger.log('Name %s sName %s', Trns.Name, Trns.sco.getName())
  // Create a Google Finance formula to get the historical data.
  var sFrm = Trns.Name + '!B1'
  sFrm = '\"' + Trns.trs.getRange(sFrm).getValue() + '\"'
  sFrm += ',Copy!B1, Copy!F1, Copy!G1'
  sFrm = 'GoogleFinance(' + sFrm + ')'
  //Logger.log('sFrm %s', sFrm)
  // Evaluate the formula and get the results.
  Trns.sco.getRange('Copy!A8').setFormula(sFrm);
  let src = 'Copy!' + Trns.sco.getRange(9, 2,1,4).getA1Notation()
    //Logger.log('src %s ', src)
  scoA = GetDataArray(Trns.sco,src)
  //Logger.log('scoA %s',src,scoA)
  for (i = 0; i <4; ++i) {
    Trns.trs.getRange(dcl[i]).setValue(scoA[i])
    //Logger.log('scOA[i] %s v %s', scoA[i], Trns.trs.getRange(dcl[i]).getValue())
  }
  formatEksikCells()
}

function formatEksikCells() {
  formatRange('B4', 'B3')
  formatRange('b3', 'F3:F5')
  let c = Trns.trs.getRange('Kebir!G1').getValue()
  let dst = Trns.trs.getRange(2, c, 5, 1).getA1Notation()
  //Logger.log('dst %s',dst)
  copyFormatRange('F2:F6', dst)
  dst = Trns.trs.getRange(15, c).getA1Notation()
  //Logger.log('b3 icin dst %s',dst)
  formatRange('B3', dst)
}
function clearE1() {
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  Trns.sht.getSheets().forEach(function (s) {
    Trns.trs = s
    Trns.trs.activate()
    var i = exclude.indexOf(Trns.trs.getName());
    if (i == -1) {
      Trns.trs.getRange('E1').clear({ contentsOnly: true, skipFilteredRows: true });
    }
  })
  Trns.sht.getRange('Kebir!K1').clear({ contentsOnly: true, skipFilteredRows: true })
  Trns.sht.getRange('Kebir!G2').clear({ contentsOnly: true, skipFilteredRows: true })
}
/* ********************************************** */
function getStockPriceWithErrorHandling() {
  Trns.sco = Trns.sht.getSheets()[1]
  Trns.trs = Trns.sht.getSheets()[4]
  Trns.Name = Trns.trs.getName()
  //Logger.log('Name %s', Trns.Name)
  //Logger.log('Name %s sName %s', Trns.Name, Trns.sco.getName())
  // Create a Google Finance formula to get the historical data.
  var sFrm = Trns.Name + '!B1'
  sFrm = '\"' + Trns.trs.getRange(sFrm).getValue() + '\"'
  sFrm += ',Copy!B1, Copy!F1, Copy!G1'
  sFrm = 'GoogleFinance(' + sFrm + ')'
  //Logger.log('sFrm %s', sFrm)
  // Evaluate the formula and get the results.
  sFrm = '=IFERROR(' + sFrm + ', "ERROR")'
  Logger.log('srm %s ',sFrm)
  const cell = Trns.sco.getRange('Copy!A8').setFormula(sFrm);
  // Evaluate the formula and get the results.
  // Set the GOOGLEFINANCE formula with IFERROR wrapper
  // cell.setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}"), "ERROR")`);
  // Force calculation to complete
  SpreadsheetApp.flush();
  // Get the value
  const value = cell.getValue();
  // Check if it's an error or empty
  if (value === "ERROR" || value === "" || value === null) {
    Logger.log(`Failed to retrieve data for `);
    return flgEnum.FEND; 
  }
  // Check if the cell contains an error type
  if (typeof value === 'string' && value.startsWith('#')) {
    Logger.log(`Error returned: ${value} for ${ticker}`);
    return  flgEnum.FEND;
  }
  if (misHoliday()) return flgEnum.END
  Logger.log(`Successfully retrieved: `);
  return flgEnum.FCONT;
}
function restoreGoogleFinance() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('MISSING_DAYS', 'false');
  Trns.trs = Trns.sht.getSheets()[Kebir]
  Trns.trs.getRange('K1').clear({ contentsOnly: true, skipFilteredRows: true })
  Trns.sht.getSheets().forEach(function (s) {
    Trns.trs = s
    Trns.trs.activate()
    var i = exclude.indexOf(Trns.trs.getName());
    if (i == -1) {
      Trns.Name = Trns.trs.getName()
      Trns.trs.getRange('E1').clear({ contentsOnly: true, skipFilteredRows: true });
      // Current
      Trns.trs.getRange('B3').setFormula('=GoogleFinance(B1,Orders!C1)')
      // Open
      Trns.trs.getRange('F3').setFormula('=GoogleFinance($B$1,Orders!$U$1)')
      // Low 
      Trns.trs.getRange('F4').setFormula('=GoogleFinance($B$1,Orders!$D$1)')
      // High
      Trns.trs.getRange('F5').setFormula('=GoogleFinance($B$1,Orders!$E$1)')
      formatEksikCells()
    }
  })
  setStartingActiveSheet()
};


function setOHLCValues() {
  Trns.Name = Trns.trs.getName()
  //Logger.log('Name %s', Trns.Name)
  Trns.trs.getRange('E1').clear({ contentsOnly: true, skipFilteredRows: true });
  //Logger.log('Name %s sName %s', Trns.Name, Trns.sco.getName())
  // Create a Google Finance formula to get the historical data.
  var sFrm = Trns.Name + '!B1'
  sFrm = '\"' + Trns.trs.getRange(sFrm).getValue() + '\"'
  sFrm += ',Copy!B1, Copy!F1, Copy!G1'
  sFrm = 'GoogleFinance(' + sFrm + ')'
  //Logger.log('sFrm %s', sFrm)
  // Evaluate the formula and get the results.
  Trns.sco.getRange('Copy!A8').setFormula(sFrm);
  let src = 'Copy!' + Trns.sco.getRange(9, 2,1,4).getA1Notation()
    //Logger.log('src %s ', src)
  scoA = GetDataArray(Trns.sco,src)
  //Logger.log('scoA %s',src,scoA)
  for (i = 0; i <4; ++i) {
    Trns.trs.getRange(dcl[i]).setValue(scoA[i])
    //Logger.log('scOA[i] %s v %s', scoA[i], Trns.trs.getRange(dcl[i]).getValue())
  }
  formatEksikCells()
}

/**
 * Global object A
 */
const Trn = {
  // Internal variable to store the instance
  _sht: null,

  /**
   * Getter for property b.
   * Logic: If the instance doesn't exist, create it. Otherwise, return the existing one.
   */
  get sht() {
    if (!this._sht) {
      console.log("Initializing Singleton Instance for 'sht'...");
      
      // Define your singleton object here
      this._sht = SpreadsheetApp.getActiveSheet()
      /*{
        timestamp: new Date().getTime(),
        id: Math.random().toString(36).substring(7),
        greet: function() {
          return "Hello from Singleton B!";
        }
      };*/
      
      // Optional: Prevent further modifications to the instance
      Object.freeze(this._sht);
    }
    return this._sht;
  }
};


