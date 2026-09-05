/**
 * ******************    minmax.gs     *********************
 * */
adrw = []   // array of date starting rows
amnO = []  // array of Min Quarter Objects
amxO = []   // array of Max Quarter Objets
stack = [] // merge amno & amxO
finalStack = []
ratioStack = []
function testdefineQuarters() {
  //Trns.sht = SpreadsheetApp.getActiveSpreadsheet()
  //Trns.trs = Trns.sht.getActiveSheet();
  //Trns.bindTab(Trns.sht.getActiveSheet())
  //Trns.Name = Trns.trs.getSheetName()
  setOrderRow()
  // Logger.log('Name %s Order Name %s row %s',Trns.Name, Trns.ord.getName(), Trns.or) 
  defineQuarters()
  /*  Trns.oName.forEach(
      function (sName) {
        if (include.indexOf(sName) > -1) {
          Trns.or = Trns.oName.indexOf(sName) + 3
          //Trns.trs = Trns.sht.getSheetByName(sName)
          Trns.bindTab(Trns.sht.getSheetByName(sName))
          //Trns.Name = Trns.trs.getSheetName()
          // Logger.log('sName %s rName %s-', sName, Trns.Name)
          Trns.trs.activate()
          defineQuarters()
        }
      });*/
}
function isequalDate(sdt, chd) {
  // Logger.log('start %s check %s eq %s', sdt, chd, sdt == chd)
  if (chd != null) {
    if (sdt.getDay() == chd.getDay()) {
      if (sdt.getMonth() == chd.getMonth())
        return true
    }
  }
  return false
}


/*****************************************************************************/
function defineQuarters() {
  try {
    const lrow = Trns.trs.getLastRow()
    //Logger.log('lrow %s ', lrow)
    srow = nextcell() + 1
    // Logger.log('srow %s',srow)
    const fqsr = srow
    qrr = srow - 3 // Quarter starting displlay row
    //s = 'A'+(qrr-10)
    // Check if it is done
    sdt = new Date(Trns.trs.getRange('A' + srow).getValue()) // Starting date
    srng = 'C' + qrr
    chd = Trns.trs.getRange(srng).isBlank() ? null
      : new Date(Trns.trs.getRange(srng).getValue()) // Chekck if it is done
    //Logger.log('srng %s sdt %s chd %s',srng, sdt,chd)
    if (isequalDate(sdt, chd))
      return
    let gr = qrr - 9
    s = 'A' + qrr + ':N' + gr
    Trns.trs.getRange(s).clear({ contentsOnly: true, skipFilteredRows: true })
    copyCell('A' + srow, srng)
    Trns.trs.getRange(qrr, 2).setValue(srow)

    dsr = srow - 10 // Display sheet row
    //Logger.log('Name %s srow %s srng %s dsr %s', Trns.trs.getName(), srow,s, dsr)  
    const dksg = 90 * 24 * 60 * 60 * 1000
    // Logger.log('sdt %s', sdt)
    for (i = 0; i < 4; ++i, --qrr) {
      edt = new Date(sdt.getTime() - dksg);
      //Logger.log('i %s start date %s end date %s srow %s', i, sdt, edt, srow)
      if (srow + 90 >= lrow - 1) {
        s = 'A' + qrr + ':O' + gr
        Trns.trs.getRange(s).clear({ contentsOnly: true, skipFilteredRows: true })
        return
      }
      adrw[i] = getQuarterEnd(srow, edt)
      sdt = edt
      //Logger.log("srow %s adrw[i] %s dsr %s s %s", srow, adrw[i], dsr, s)
      //Logger.log(' srow+45 %s lrow %s', srow + 45, lrow)

      getMinMaxValues(i, srow)
      setglobalRatios(i, adrw[i], gr)
      srow = adrw[i] + 1
      /*
       Logger.log('new srow %s qrr %s', srow, qrr)
//      Logger.log('quarter No %s Min Values\n %s', i, amnO[i])
//      Logger.log('quarter No %s Max Values\n %s', i, amxO[i]) */
      copyCell('A' + srow, 'C' + qrr)
      Trns.trs.getRange(qrr, 2).setValue(srow)
      displaytoSht(i, dsr)
      dsr += 2
    }
    generatefinalStack()
    generateRatios(fqsr)
    updateOrderSheet()
    Logger.log('defineQuarters bitti ')
  }
  catch (e) {
    Logger.log('Define Quarters Error occurred  msg\n %s \n stack %s', e.message, e.stack)
  }
  return true
}
const ss = ['Y', 'X', 'W', 'V']
/*****************************************************************************/
function updateOrderSheet() {
  Trns.or = Trns.oName.indexOf(Trns.Name) + 3
  for (i = 0; i < 4; i++) {
    o = ratioStack[i] 
    sdst = ss[i] + Trns.or
    if (o.v != null)
      CopyCellDif(Trns.trs, o.v, Trns.ord, sdst)
  }
}
/*****************************************************************************/
function setglobalRatios(i, srow, gr) {
  sp = 'E' + srow
  s = '=(B3 - ' + sp + ') / ' + sp
  Trns.trs.getRange(gr, i + 2).setValue(s)
}
/*****************************************************************************/
function generateRatio(first, last) {
  s = '=(' + first + '- ' + last + ')/' + last
  return s
}
/*****************************************************************************/
function pushRatioStack(v, d, b = true) {
  obj = {
    v, d
  }
  if (b)
    ratioStack.push(obj)
  else ratioStack.unshift(obj);
  //Logger.log('rastioStackpush %s', obj)
}
/*****************************************************************************/
function todaysRatio(firstPop, fqsr) {
  // Logger.log('todaysRatio firstpop %s', firstPop)
  if (firstPop.row == fqsr) {
    for (i = 0; i < 2; ++i)
      pushRatioStack(null, firstPop.date)
    return
  }
  dr = 'A' + fqsr
  dv = Trns.trs.getRange(dr).getValue()
  fs = (firstPop.type == 'Min' ? 'F' : 'I') + firstPop.rno
  fv = Trns.trs.getRange(fs).getValue()
  //Logger.log('dv %s fv %s ',  dv, fv)
  if (firstPop.type == 'Max') {
    rs = 'D' + fqsr + ':D' + (firstPop.row - 1) // 'C'  'D'
    fs = '=MIN(' + rs + ')'
    rno = firstPop.rno - 1
    sl = 'E' + rno
    Trns.trs.getRange(sl).setValue(fs)
    v = Trns.trs.getRange(sl).getValue()
    // Logger.log('sl %s v %s',sl, v)
    let o = findCellValue(Trns.trs, rs, v)
    if (o == null) {
      //Logger.log('o %s v %s fs %s sl %s',o, v,rs,sl) 
      //Logger.log('todays ratio o null gelmemesi gerekir')
    }
    //Logger.log('o %s v %s fs %s sl %s',o, v,rs,sl) 
    copyCell('A' + o.r, 'F' + rno)
    ns = 'H' + firstPop.rno
    fs = '=(' + sl + ' - ' + ns + ')/' + ns
    rg = 'L' + rno
    Trns.trs.getRange(rg).setValue(fs)
    pushRatioStack(rg, Trns.trs.getRange(o.r, 1).getValue())
    // Logger.log('o.r %s fqsr %s', o.r, fqsr)
    if (o.r == fqsr) {
      pushRatioStack(null, firstPop.date, false)
      return
    }
    orno = rno
    --rno
    copyCell('D' + fqsr, 'H' + rno) //
    copyCell('A' + fqsr, 'I' + rno)
    ns = 'E' + orno
    fs = '=(H' + rno + ' - ' + ns + ')/' + ns
    rg = 'K' + rno
    Trns.trs.getRange(rg).setValue(fs)
    pushRatioStack(rg, Trns.trs.getRange(o.r, 1).getValue(), false)
  } else { // Min
    rs = 'C' + fqsr + ':C' + (firstPop.row - 1)
    fs = '=MAX(' + rs + ')'
    rno = firstPop.rno - 1
    sl = 'H' + rno
    Trns.trs.getRange(sl).setValue(fs)
    v = Trns.trs.getRange(sl).getValue()
    // Logger.log('ra %s v %s o %s',rs,v, o)
    o = findCellValue(Trns.trs, rs, v)
    copyCell('A' + o.r, 'I' + rno)
    ns = 'E' + firstPop.rno
    fs = '=(' + sl + ' - ' + ns + ')/' + ns
    rg = 'K' + rno
    Trns.trs.getRange(rg).setValue(fs)
    pushRatioStack(rg, Trns.trs.getRange(o.r, 1).getValue())
    if (o.r == fqsr) {
      pushRatioStack(null, firstPop.date, false)
      return
    }
    orno = rno
    --rno
    copyCell('D' + fqsr, 'E' + rno)
    copyCell('A' + fqsr, 'F' + rno)
    ns = 'H' + orno
    fs = '=(E' + rno + ' - ' + ns + ')/' + ns
    rg = 'L' + rno
    Trns.trs.getRange(rg).setValue(fs)
    pushRatioStack(rg, Trns.trs.getRange(fqsr, 1).getValue(), false)
  }

}
/*****************************************************************************/
function generateRatios(fqsr) {
  firstPop = finalStack.pop()
  ratioStack = []
  // Logger.log('todays once RatioStack %s',ratioStack)
  todaysRatio(firstPop, fqsr)
  // Logger.log('todays sonra RatioStack %s',ratioStack)
  for (m = 0; ; ++m) {
    if (finalStack.length == 0) break;
    // Logger.log('firstPop %s m %s length %s', firstPop, m, finalStack.length)
    lastPop = finalStack.pop()
    s = generateRatio((firstPop.type == 'Min' ? 'E' : 'H') + firstPop.rno,
      (lastPop.type == 'Min' ? 'E' : 'H') + lastPop.rno)

    rs = firstPop.type == 'Min' ? 'L' : 'K'
    rs += firstPop.rno
    // Logger.log('v %s s %s rs %s', s,rs)
    Trns.trs.getRange(rs).setValue(s)
    pushRatioStack(rs, firstPop.date)
    s = generatePassDay((firstPop.type == 'Min' ? 'F' : 'I') + firstPop.rno,
      (firstPop.type == 'Min' ? 'I' : 'F') + lastPop.rno)
    rs = 'M' + firstPop.rno
    Trns.trs.getRange(rs).setValue(s)
    //  Logger.log('Obj %s',obj)
    firstPop = lastPop
  }
}
/*****************************************************************************/
function generatePassDay(first, last) {
  return '=' + first + '-' + last
}
/*****************************************************************************/
function generatefinalStack() {
  i = 0;
  firstPop = stack.pop()
  finalStack.push(firstPop)
  for (; ;) {
    if (stack.length == 0) break
    lastPop = stack.pop()
    if (firstPop.type != lastPop.type) {
      finalStack.push(lastPop)
      firstPop = lastPop
    } else { // typelar esit
      // Logger.log('firstPop %s\nlastPop %s',firstPop,lastPop)
      if ((firstPop.type == 'Min' && firstPop.val > lastPop.val) ||
        (firstPop.type == 'Max' && lastPop.val > firstPop.val)) {
        finalStack.pop()
        finalStack.push(lastPop)
        firstPop = lastPop
      }
    }
  }
  /*  for(i = 0; i<finalStack.length; ++i) {
      Logger.log('i %s final[i] %s', i, finalStack[i])
    }*/
}
/*****************************************************************************/
function displaytoSht(i, dsr) {
  Trns.trs.getRange('J' + dsr).setValue(i + 1)
  mnc = 6
  mxc = 8
  xr = amxO[i].date > amnO[i].date ? dsr : dsr + 1
  amxO[i].xr = xr
  src = 'A' + amxO[i].row
  dst = 'I' + xr
  copyCell(src, dst)
  src = 'C' + amxO[i].row
  dst = 'H' + xr
  copyCell(src, dst)
  mr = amxO[i].date > amnO[i].date ? dsr + 1 : dsr
  amnO[i].mr = mr
  //Logger.log('displaytoSheet() i %s xr %s mr %s', i, amxO[i].xr, amnO[i].mr)
  src = 'A' + amnO[i].row
  dst = 'F' + mr
  copyCell(src, dst)
  src = 'D' + amnO[i].row
  dst = 'E' + mr
  copyCell(src, dst)
  if (amxO[i].date > amnO[i].date) {
    pushamxO(i)
    pushamnO(i)
  } else {
    pushamnO(i)
    pushamxO(i)
  }
  /*****************************************************************************/
  function pushamxO(i) {
    Obj = {
      type: 'Max',
      cell: amxO[i].cell,
      row: amxO[i].row,
      date: amxO[i].date,
      val: amxO[i].val,
      rno: amxO[i].xr
    }
    stack.push(Obj)
  }
  /*****************************************************************************/
  function pushamnO(i) {
    Obj = {
      type: 'Min',
      cell: amnO[i].cell,
      row: amnO[i].row,
      date: amnO[i].date,
      val: amnO[i].val,
      rno: amnO[i].mr
    }
    stack.push(Obj)
  }
}
/*****************************************************************************/
function getMinMaxValues(i, srow) {
  //Logger.log('i %s adrw[i] - srow + 1 %s  adrw[i] %s srow %s ', i, adrw[i] - srow + 1,  adrw[i], srow )
  rw = adrw[i] - srow + 1
  if (rw <= 0) rw = 1
  const rng = Trns.trs.getRange(srow, 4, rw, 2)
  //Logger.log('getMinMaxValues i %s srow %s rng %s',i, srow, rng.getA1Notation())
  const values = rng.getValues();
  let minValue = Infinity;
  let minRow = -1;
  let minCol = 5
  let maxValue = -999999999
  maxRow = -1
  maxCol = 4
  //Logger.log('getMinMaxValues()')
  //Logger.log('rng %s values.length %s ',rng.getA1Notation(), values.length)
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const currentValue = values[i][j];
      if (j == 1) {
        if (typeof currentValue === 'number' && currentValue < minValue) {
          minValue = currentValue;
          minRow = i;
          //Logger.log('minValue %s minRow %s', minValue, i)
        }
      } else {
        if (typeof currentValue === 'number' && currentValue > maxValue) {
          maxValue = currentValue;
          maxRow = i;
          // Logger.log('maxValue %s maxRow %s', maxValue, i)
        }
      }
    }
  }
  if (minValue === Infinity) {
    amnO[i] = {}
  } else {
    const resultCell = Trns.trs.getRange(srow + minRow, minCol);
    amnO[i] = {
      minValue: minValue,
      cell: resultCell.getA1Notation(),
      row: resultCell.getRow(),
      column: resultCell.getColumn(),
      date: Trns.trs.getRange(resultCell.getRow(), 1).getValue(),
      val: Trns.trs.getRange(resultCell.getRow(), minCol).getValue(),
      mr: 0
    }
  }

  if (maxValue === 999999999) {
    amxO[i] = {}
  } else {
    crow = srow + maxRow
    //Logger.log('srow %s maxRow %s srow+maxRow %s', srow, maxRow, crow)
    const resultCell = Trns.trs.getRange(crow, maxCol);
    amxO[i] = {
      maxValue: maxValue,
      cell: resultCell.getA1Notation(),
      row: resultCell.getRow(),
      column: resultCell.getColumn(),
      date: Trns.trs.getRange(resultCell.getRow(), 1).getValue(),
      val: Trns.trs.getRange(resultCell.getRow(), maxCol).getValue(),
      xr: 0
    };
  }
}

/*****************************************************************************/
function getQuarterEnd(rw, edt) {
  rw += 45
  const rng = Trns.trs.getRange(rw, 1, 45, 1)
  // Logger.log(' Sht Name %s row %s rng %s ', Trns.trs.getName(),rw,rng.getA1Notation())
  const values = rng.getValues();
  fm = 0  // first month
  mc = 0  // end month
  erw = 0
  values.forEach(
    function (row) {
      row.forEach(
        function (cell) {
          if (erw == 0) {
            var dt = new Date(cell)
            // Logger.log('row %s date %s', rw, dt)
            if (dt < edt)
              erw = rw
          }
        }
      );
      ++rw
    }
  );
  return erw
}



