/****************************************************************************/
// mailto
/****************************************************************************/
var oNtfy = {}
const fEnm = { fCoB: 6, fCoS: 7, fCoH: 10, fCoL: 11 }
const cOpn = 12 // Open
const cCrnt = 8 // Crnt
const cPrv = 9  // Previous
const messagetype = { 6: "    buy = ", 7: "   sell = ", 10: "   high = ", 11: "    low = " }
/****************************************************************************/
// fCtB  Security Sheet Buy column 'D'
// fCtS  Security Sheet Sell column 'E'
// fCoB  Order Sheet Buy column 'F'
// fCoS  Order Sheet Sell column 'G'
// fCoH  High
// FCoL  Low 
var obj = fEnm
let or = 0

function testmailfromSnapShot() {
 // var userProperties = PropertiesService.getUserProperties();
  Trns.sht = SpreadsheetApp.getActive()
  Trns.trs = Trns.sht.getActiveSheet()
  getSecNamesfromOrderSheet()
  // Logger.log('Name %s Ord Name %s', Trns.Name, Trns.ord.getName()) 
  mailtome(5)
  //addDiftoOpen(5)
}
function putOkey() { 
  Trns.sht = SpreadsheetApp.getActive()
  Trns.sht.getSheets().forEach(function (s) {
    s.activate()
    var i = exclude.indexOf(s.getName());
    if (i == -1) 
      s.getRange('Y1').setValue('OKEY')       
  })
}



 
function mailtome(No) {
  Logger.log(' %s %s', No, parseInt(No % 5))
  if (parseInt(No % 15)) return
  or = 3
  let msgBody = ""
  Trns.oName.forEach(function (entry) {
    //      Logger.log('No %s Name %s or %s', No, entry, or)
    msgBody += addObjMesages(or,entry,No)
    ++or
  })
  //    Logger.log('mailfromSnapShot bitti') 
  if (msgBody != "") {
    let eMailName = Session.getActiveUser().getEmail();
    let sbjctLine = "Snap Shot Results No : " + No
    Logger.log(' %s', msgBody)
    Logger.log('eMailName %s sjctLine %s', eMailName, sbjctLine)
    MailApp.sendEmail(eMailName, sbjctLine, msgBody)
  }
}

function addObjMesages(or, sName,No) {
  ix = or - 3
  let Msg = ""
  try {
    let cv = true
    for (var key in obj) {
      let bc = Trns.ord.getRange(or, obj[key]).getBackground() != fntclr
      if (bc) {
        //       Logger.log('key %s obj[key] %s bc %s',key,obj[key],bc)
        Msg += addtypeMessages(or, obj[key])
        if (key == fEnm.fCoH) {
          let v = Trns.ord.getRange(or, obj[key]).getValue()
          if (v > 2.0) Msg += sprintf('Dikkat High %2 den buyuk %s', v)
          if (v < 0.001) Msg += sprintf('Dikkat High Sifir %s', v)
        }
        if (key == fEnm.fCoL) {
          let v = Trns.ord.getRange(or, obj[key]).getValue()
          if (v < 0.02) Msg += sprintf('Dikkat Low %2 den kucuk %s', v)
          if (v > 0.02) Msg += sprintf('Dikkat Low Sifir %s', v)
        }
      }
      if (cv) {
        cv = false
        let v = Trns.ord.getRange(or, cp).getValue()
        if (abs(v) > 0.02) {
          Msg += sprintf('Dikkat PrvCrnt %s = ', v)
          Msg += (v > 0.02) ? ' %2 den buyuk \n' : ' % 2 den kucuk \n'
        }
      }
    }
    Msg += testTenDaysMinMax(sName)
    if (Msg != "") {
      Msg = sprintf("\n %s \n   Current Value  = %s \n", Trns.oName[ix],
      Trns.ord.getRange(or, 8).getValue()) + Msg  

    }
  }
  
  catch (e) {
    Logger.log('mail from snapshot hata')
  }
  return Msg
}
function testTenDaysMinMax(sName) {
  s = ''
  sht = Trns.sht.getSheetByName(sName)
  for (i = 0; i<2; ++i) { 
    sv = (i == 0) ? 'F7' : 'F8'
    let c = sht.getRange(sv).getBackgroundColor()
    if ((i == 0 && c == bckclrY) || (i == 1 && c == bckclrK)){
      v = sht.getRange(sv).getValue()
      s += i == 0 ? sprintf(' On gunluk Low Value %s  \n', v)
                  : sprintf ('On gunluk High Value %s \n', v)
      vl = sht.getRange('D3').getValue()
      vr = sht.getRange('J9').getValue()
      cv = (v /  vl - 1) / vr
      s += sprintf('       Crnt Value %s \n', sht.getRange('B3').getValue())
      s += sprintf('         Reality Ratio %s \n',cv)
    }
  }
  let y =  sht.getRange('Y1').getValue()
  if (y == 'OKEY') s += addDiftoOpen(sht)
  return s
}
function addtypeMessages(or, key) {
  let s = sprintf('\t %s %s \n',
    messagetype[key], Trns.ord.getRange(or, key).getValue())
  //  Logger.log('sprintf %s',s)
  return s
}

function addDiftoOpen(sht) {
  let opn = sht.getRange('F3').getValue()
  let prv = sht.getRange('B3').getValue()
  let opnprv = (prv - opn) / opn
  s = ''
  if (abs(opnprv) > 0.02) {
    s += sprintf('Dikkat PrvOpen %s Prv %s Crnt %s = ', v, prv, opn)
    s += (opnprv > 0.02) ? ' %2 den buyuk \n' : ' % 2 den kucuk \n'
  }
  sht.getRange('Y1').setValue('NOTOK')
  return s
}

function createBackup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const folderName = "Google Sheets Backups";
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy_HH-mm');
  const backupName = `${spreadsheet.getName()} Backup ${timestamp}`;

  // Create or locate the backup folder
  let folder = DriveApp.getFoldersByName(folderName).hasNext() 
    ? DriveApp.getFoldersByName(folderName).next() 
    : DriveApp.createFolder(folderName);

  // Create a full spreadsheet backup
  const backupFile = DriveApp.getFileById(spreadsheet.getId()).makeCopy(backupName, folder);
  Logger.log(`Backup created: ${backupFile.getUrl()}`);
}

