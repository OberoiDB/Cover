function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('COVER!')
      .addItem('1. Poll', 'pollAllTeachers')
      .addItem('2. Load Requests and Assignments', 'loadRequests')
      .addItem('3. Save Assignments', 'saveAssignments')
      .addItem('3. Poll Selected Teacher', 'pollSelectedTeacher')
      .addItem('4. Prepare Plan Preview', 'previewPlan')
      .addItem('5. Publish Plan', 'publishPlan')
      .addItem('6. Send Daily Email', 'sendDailyEmail')
      .addItem('7. Notify Cover Teachers', 'notifyCover')
      .addSeparator()
      .addItem('Lock and Archive Plans', 'archivePlan')
      .addItem('Load Teacher Data', 'loadTeacherData')
      .addItem('Update Teacher Sheets', 'updateTeachers')
      .addToUi();  
}

function onEdit(e){
  if (e.range.getSheet().getName()!="Planner") { return; }
  if ((e.range.getRow()<4) && e.range.getColumn()==3) {
    if (e.range.getRow()==1) {
      loadDayData(); 
    } else {
      sortData();
    }
  }  
}

function lastNonBlank(nbSheet, nbColumn, nbStartRow) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var pSheet = ss.getSheetByName("Planner");

  var lr = nbSheet.getLastRow();

  var ir = nbStartRow;
  
  var isNotBlank = !nbSheet.getRange(ir,nbColumn).isBlank();
  
  while (isNotBlank) {
    ir+=1;
    if (ir>lr) {
      isNotBlank=false;
    } else {
      isNotBlank = !nbSheet.getRange(ir,nbColumn).isBlank();
    }
  }
  ir -= 1;
  
  return ir;
}

function loadSubmitted() {
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Wait While We Update Your Planner ... ',3);
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var pSheet = ss.getSheetByName("Planner");
  var sSheet = ss.getSheetByName("Cover Submissions");
  var pDate = pSheet.getRange(1, 3).getValue();
  pSheet.getRange(6, 16, pSheet.getLastRow()-5, 9).clearDataValidations();
  pSheet.getRange(6, 16, pSheet.getLastRow()-5, 9).clear();

  var sLR = lastNonBlank(sSheet, 1, 2);
  if (sLR<2) { 
    SpreadsheetApp.getActiveSpreadsheet().toast('', 'No Submissions!',3);
    return; 
  }
  var sData = sSheet.getRange(2, 1, sLR-1, 11).getValues();
  var rows = new Array();

  for(var si = 0; si < sLR-1; si++) {
    var sStatus = sData[si][0];
    var sDate   = sData[si][4];
    
    if ((sStatus=="Submitted"||sStatus=="Assigned"||sStatus=="Planned") && sDate==pDate) {
      var sCover  = sData[si][1];
      var sTeach  = sData[si][2];
      var sDept   = sData[si][3];
      var sPeriod = sData[si][5];
      var sBlock  = sData[si][6];
      var sClass  = sData[si][7];
      var sReqs   = sData[si][8];
      var sSubmit = sData[si][9];
      var sInstr  = sData[si][10];
      if (sStatus=="Submitted") sCover="";
      Logger.log(sCover);
      rows.push([[sStatus],[sCover],[sTeach],[sDept],[sPeriod],[sBlock],[sClass],[sReqs],[sSubmit+sInstr]]);
    }
  }
  
  if (rows.length==0) { 
    SpreadsheetApp.getActiveSpreadsheet().toast('', 'No Submissions for Date!',3);
    return; 
  }
  
  var nr = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('TeacherList');
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(nr).build();
  var formatRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('PRequestsFormat');

  pSheet.getRange(6, 16, rows.length, 9).setValues(rows);
  pSheet.getRange(6,17,rows.length,1).setDataValidation(rule);
  formatRange.copyFormatToRange(pSheet, 16, 24, 6, 5+rows.length);
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Planner Ready!',3);
}

function AssignTeachers() {
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Wait While We Update Your Assignments ... ',3);
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var pSheet = ss.getSheetByName("Planner");
  var sSheet = ss.getSheetByName("Cover Submissions");
  var pDate = pSheet.getRange(1, 3).getValue();

  var pLR = lastNonBlank(pSheet, 16, 6);
  var pData = pSheet.getRange(6, 16, pLR-5, 6).getValues();
    Logger.log(pLR);
  
  var sLR = lastNonBlank(sSheet, 1, 2);
  var sData = sSheet.getRange(2, 1, sLR-1, 7).getValues();
    Logger.log(sLR);

  for(var pi = 0; pi < pLR-5; pi++) {
    //Status	AssignedTeacher	RequestingTeacher	Dept	Period	Block
    var pName  = pData[pi][2];
    Logger.log("p:"+pName);
    var pBlock = pData[pi][5];
    for(var si = 0; si < sLR-1; si++) {
      var sName  = sData[si][2];
      Logger.log("s:"+sName);
      var sDate  = sData[si][4];
      var sBlock = sData[si][6];
      //Status	AssignedTeacher	RequestingTeacher	Dept	Date	Period	Block
      if (sName==pName && sDate==pDate && sBlock==pBlock) {
        Logger.log("FF");
        var pCover = pData[pi][1];
        sData[si][1]=pCover;
        if (pCover=="") {
          sData[si][0]="Submitted";
        } else {
          sData[si][0]="Assigned";
        }
        si=sLR
      }
    }  
  }  
  
  sSheet.getRange(2, 1, sLR-1, 7).setValues(sData);
  pSheet.getRange(6, 16, pSheet.getLastRow()-5, 9).clearDataValidations();
  pSheet.getRange(6, 16, pSheet.getLastRow()-5, 9).clear();
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Assignments Ready!',3);
}

function submitPlan() {
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Wait While We Submit The Plan ... ',3);
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var lSheet = ss.getSheetByName("Planner");
  var sSheet = ss.getSheetByName("Cover Submissions");
  var rSheet = ss.getSheetByName("Rotation");
  var toSS    = SpreadsheetApp.openById("1D_ucmFfQEPNWUakthyyJLGM799HgrNqT4h854x54oj8");
  var toPSheet = toSS.getSheetByName("Daily Cover Plan");

  if (toPSheet.getMaxRows()>5) {
    toPSheet.deleteRows(4, toPSheet.getMaxRows()-4);
  }

  var sLR = lastNonBlank(sSheet, 1, 2);

  var pDate = lSheet.getRange(1, 3).getValue();
  var crDayNumber = Number(pDate.substr(1, 1));
  var crDayOfWeek = pDate.substr(3,3);
  
  var p1 = rSheet.getRange(2+crDayNumber, 2).getValue();
  var p2 = rSheet.getRange(2+crDayNumber, 3).getValue();
  var p3 = rSheet.getRange(2+crDayNumber, 4).getValue();
  var p4 = rSheet.getRange(2+crDayNumber, 5).getValue();

  var dateOut = "Cover for:  " + pDate + " (hr-" + p1 + "-" + p2 + "-" + p3 + "-" + p4 + ")";
  if (crDayOfWeek=="Wed") {
    dateOut = "Cover for:  " + pDate + " (" + p1 + "-" + p2 + "-" + p3 + "-HR-" + p4 + ")";
  }

  var sStatuses = sSheet.getRange(2, 1, sLR-1, 1).getValues();
  var sData     = sSheet.getRange(2, 2, sLR-1, 10).getValues();
  var rows = new Array();
  var lastName = "";

  for(var si = 0; si < sLR-1; si++) {
    
    var sStatus = sStatuses[si][0];
    var sCover  = sData[si][0];
    var sDate   = sData[si][3];

    if (sDate==pDate) {
      if (sStatus=="Assigned" || sStatus=="Planned") {
        var rowData = new Array();
        var sTeach  = sData[si][1];
        var sDept   = sData[si][2];
        var sPeriod = sData[si][4];
        var sBlock  = sData[si][5];
        var sClass  = sData[si][6];
        var sInstr  = sData[si][9];
        
        if (sTeach==lastName) {
          rowData.push("");
        } else {
          if (lastName!="") {
            rows.push([[],[],[],[],[],[]]);
          }
          rowData.push(sTeach + " (" + sDept + ")");
          lastName = sTeach;
        }
        
        rowData.push(sCover);
        rowData.push(sPeriod + " - " + sBlock);
        rowData.push(sClass);
        rowData.push("link");
        rowData.push(sInstr);
        rows.push(rowData);
        sStatuses[si][0]="Planned";
      }
    }
  }
  sSheet.getRange(2, 1, sLR-1, 1).setValues(sStatuses);

  rows.push([[],[],[],[],[],[]]);
  
  toPSheet.getRange(1, 1).setValue(dateOut);
  toPSheet.insertRowsAfter(3, rows.length);
  toPSheet.getRange(4, 1, rows.length, 6).setValues(rows);
  
  var formatRange = toSS.getRangeByName('SPlanFormat');
  formatRange.copyFormatToRange(toPSheet, 1, 6, 4, 3+rows.length);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Plan Submitted!',3);
}

function sendEmailAll() {
}

function sendEmailCover() {
}

function archivePlan() {
}

function loadTeacherData() {
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Loading Teacher Data ... ',3);
  var toSS    = SpreadsheetApp.getActiveSpreadsheet(); 
  var toSheet = toSS.getSheetByName("Teachers");
  
  var frSS    = SpreadsheetApp.openById("1ZtdBIwzDyUCSg2lIge3nioUVvWlV_jRGRUknVzhdpkw");
  var frSheet = frSS.getSheetByName("Teacher timetable");

  var frLR = lastNonBlank(frSheet, 1, 2);
  var toLR = lastNonBlank(toSheet, 1, 2);

  if (toLR>1) { toSheet.getRange(2, 1, toLR-1, 14).clear(); }
  
  for(var i = 2; i < frLR+1; i++) {
    toSheet.getRange(i, 1, 1, 14).setValues(frSheet.getRange(i, 1, 1, 14).getValues());
  }
  toSheet.getRange(2, 1, frLR-1, 14).sort([{column: 1, ascending: true}]);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Loading Teacher Data Complete! ',5);
}

function moveLinks() {
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Updating Links ... ',3);

  var frSS    = SpreadsheetApp.getActiveSpreadsheet(); 
  var frTSheet = frSS.getSheetByName("Teachers");
  var frDSheet = frSS.getSheetByName("Cover Data");
  
  var toSS    = SpreadsheetApp.openById("1D_ucmFfQEPNWUakthyyJLGM799HgrNqT4h854x54oj8");
  var toWSheet = toSS.getSheetByName("Work");
  var toLSheet = toSS.getSheetByName("Teacher Request Links");
  
  toLSheet.getRange(2, 1, toLSheet.getMaxRows()-1, toLSheet.getMaxColumns()).clearContent();
  toWSheet.clearContents();

  var tLR = lastNonBlank(frTSheet, 1, 2);
  var dLR = lastNonBlank(frDSheet, 1, 2);

  var tNames = frTSheet.getRange(2, 1, tLR-1, 1).getValues();
  var dNames = frDSheet.getRange(2, 1, dLR-1, 1).getValues();
  var dLinks = frDSheet.getRange(2, 6, dLR-1, 1).getFormulas();
  var tDepts = frTSheet.getRange(2, 2, tLR-1, 1).getValues();

  var tLinks = new Array();
  
  var dl = 0;
  for(var ti = 0; ti < tLR-1; ti++) {
    for(var di = 0; di < dLR-1; di++) {
      if (tNames[ti][0] == dNames[di][0]) { 
        dl += 1;
        tLinks.push(dLinks[di]); 
      }
    }
  }
  Logger.log(dl);
  Logger.log(tLR);

  toWSheet.getRange(2, 1, tLR-1, 1).setValues(tDepts);
  toWSheet.getRange(2, 2, tLR-1, 1).setValues(tLinks);
  
  toWSheet.getRange(2, 1, tLR-1, 2).sort([{column: 1, ascending: true},{column: 2, ascending: true}]);
  
  toWSheet.getRange(1, 1, 1, 2).copyFormatToRange(toWSheet, 1, 2, 2, tLR);
  
  var colI = 0;
  var rowI = 2;
  var dept = "";
  
  var ri=2;
  while (ri < tLR+1) {
    if (dept!= toWSheet.getRange(ri, 1, 1, 1).getValue()) {
      colI += 1;
      rowI = 2;
      dept= toWSheet.getRange(ri, 1, 1, 1).getValue();
      toLSheet.getRange(1,colI).setValue(dept);
    }
    toLSheet.getRange(rowI,colI).setFormula(toWSheet.getRange(ri, 2, 1, 1).getFormula());
    rowI += 1;
    ri+=1;
  }
}

function updateTeachers() {
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Wait While We Update The Teacher Spreadsheets ... ',3);
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var tSheet = ss.getSheetByName("Teachers");
  var dSheet = ss.getSheetByName("Cover Data");
  var dtSheet = ss.getSheetByName("Dates");
  var dtValues = ["Select Date"];
  
  var i=2;
  var dt = new Date();
  var dtDate = dtSheet.getRange(i, 3).getValue();
  
  while (dtDate<dt) {
    i+=1;
    dtDate = dtSheet.getRange(i, 3).getValue();
  }

  i-=1;
  
  for(var dti = i; dti < i+21; dti++) {
    dtValues.push(dtSheet.getRange(dti, 6).getValue());
  }  
  
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(dtValues, true);

  var tLR = lastNonBlank(tSheet, 1, 2);
  var dLR = lastNonBlank(dSheet, 1, 2);
  
  var ti=2;
  
  while (ti<tLR + 1) {
    
    var tData = tSheet.getRange(ti, 1, 1, 14).getValues();
    
    var tName = tData[0][0];
    var tDept = tData[0][1];
    var tAloc = tData[0][2];
    var tRoom = tData[0][3];
    var tBudy = tData[0][4];
    var tBlkA = tData[0][5];
    var tBlkB = tData[0][6];
    var tBlkC = tData[0][7];
    var tBlkD = tData[0][8];
    var tBlkE = tData[0][9];
    var tBlkF = tData[0][10];
    var tBlkG = tData[0][11];
    var tBlkH = tData[0][12];
    var tMail = tData[0][13];

    var dFound = 0;

    var dNames = dSheet.getRange(1, 1, dLR, 1).getValues();
    var di=2;

    while (di<dLR + 1) {
      
      if (dNames[di-1][0]==tName) {
        dFound=di;
      }
      di+=1;
    }
    if (dFound==0) {
      dSheet.insertRowsAfter(dLR, 1);
      dLR+=1;
      dFound = dLR;
      dSheet.getRange(dFound, 1).setValue(tName);
      dSheet.getRange(dFound, 2).setValue(tMail);
      dSheet.getRange(dFound, 4).setValue(0);
      dSheet.getRange(dFound, 5).setValue(0);
    }
    
    var tCRID = dSheet.getRange(dFound, 3).getValue();
    var isNew = false;
    if (tCRID=="") {
      var templateSheet = DriveApp.getFileById("1GBTmDyWDqlr4_uwrN3e1rarBbxG8ce8_4HGDDI-JJaw");
      var newSheet = templateSheet.makeCopy(tName+" CR Form");
      tCRID = newSheet.getId();
      dSheet.getRange(dFound, 3).setValue(tCRID);
      dSheet.getRange(dFound, 6).setFormula('=HYPERLINK("https://docs.google.com/spreadsheets/d/'+tCRID+'/edit#gid=0","'+tName+'")');
      isNew=true;
    }
    dSheet.getRange(dFound, 7).setValue(new Date());
    
    var cs = SpreadsheetApp.openById(tCRID);
    if (isNew) {
      cs.addEditor(tMail);
      cs.addEditor("sandy.vannooten@oberoi-is.org");
    }
    var ccSheet = cs.getSheetByName("Cover Request");
    ccSheet.getRange(1, 4).setValue(tName);
    ccSheet.getRange(2, 4).setValue(tDept);
    ccSheet.getRange(1, 6).setValue(tMail);
    ccSheet.getRange(2, 6).setValue(tRoom);
    ccSheet.getRange(3, 6).setValue(tBudy);
    ccSheet.getRange(5, 4).setDataValidation(rule);

    var ctSheet = cs.getSheetByName("Timetable");
    ctSheet.getRange(2, 3).setValue(tBlkA);
    ctSheet.getRange(3, 3).setValue(tBlkB);
    ctSheet.getRange(4, 3).setValue(tBlkC);
    ctSheet.getRange(5, 3).setValue(tBlkD);
    ctSheet.getRange(6, 3).setValue(tBlkE);
    ctSheet.getRange(7, 3).setValue(tBlkF);
    ctSheet.getRange(8, 3).setValue(tBlkG);
    ctSheet.getRange(9, 3).setValue(tBlkH);

    var crSheet = cs.getSheetByName("Rotation");
    crSheet.getRange(12, 2).setValue("dennis.blum@oberoi-is.org");
    crSheet.getRange(13, 2).setValue("sandy.vannooten@oberoi-is.org" + "," + tMail);    
    //crSheet.getRange(13, 2).setValue("kdceci@blumhome.com");    
    if (isNew) { crSheet.getRange(14, 2).setValue(new Date()); }  

    SpreadsheetApp.flush();
    dSheet.getRange(dFound, 7).setValue(new Date());
    
    ti+=1;
  }
  moveLinks();
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Teacher Spreadsheets Updated!',3);
}

function pollTeacher(teacherName) {
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Polling ... ',120);
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var cSheet = ss.getSheetByName("Cover Data");
  var pSheet = ss.getSheetByName("Planner");
  var nowDate = new Date();
  var pDate = pSheet.getRange(1, 3).getValue();
  
  var cLR = lastNonBlank(cSheet, 1, 2);
  var cNames = cSheet.getRange(2, 1, cLR-1, 1).getValues();
  var cCRIDs = cSheet.getRange(2, 3, cLR-1, 1).getValues();
    
  for(var ci = 0; ci < cLR-1; ci++) {
    if (teacherName=="" || cNames[ci][0]==teacherName) {
      var cs = SpreadsheetApp.openById(cCRIDs[ci][0]);
      loadToServer(cs);
      cSheet.getRange(ci+2, 8).setValue(new Date());
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Polling Complete! ',5);
}

function pollAllTeachers() {
  pollTeacher("");
}

function pollSelectedTeacher() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var pSheet = ss.getSheetByName("Planner");
  pollTeacher(pSheet.getRange(3, 17).getValue());
}

function loadToServer(tSS) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet = tSS.getSheetByName("Cover Request");
  var sSheet = tSS.getSheetByName("Cover Submission");
  var rSheet = tSS.getSheetByName("Rotation");
  
  var tSheet = ss.getSheetByName("Cover Submissions");
  var mTeacher = sheet.getRange(1, 4).getValue();
  var mDept = sheet.getRange(2, 4).getValue();

  var i=2;
  var rowNum = tSheet.getLastRow();    
  Logger.log(rowNum);
  
  while (i<sSheet.getLastRow()+1) {
    if (sSheet.getRange(i, 1).getValue()=="Pending") {
      var mDate         =  sSheet.getRange(i, 3).getValue();
      var mPeriod       =  sSheet.getRange(i, 4).getValue();
      var mBlock        =  sSheet.getRange(i, 5).getValue();
      var mClass        =  sSheet.getRange(i, 6).getValue();
      var mReqs         =  sSheet.getRange(i, 7).getValue();
      var mCover        =  sSheet.getRange(i, 8).getValue();
      var mInstructions =  sSheet.getRange(i, 9).getValue();

      tSheet.insertRowAfter(rowNum);
      rowNum  += 1;
      tSheet.getRange(rowNum, 1).setValue("Pending");
      tSheet.getRange(rowNum, 2).setValue("");
      tSheet.getRange(rowNum, 3).setValue(mTeacher);
      tSheet.getRange(rowNum, 4).setValue(mDept);
      tSheet.getRange(rowNum, 5).setValue(mDate);
      tSheet.getRange(rowNum, 6).setValue(mPeriod);
      tSheet.getRange(rowNum, 7).setValue(mBlock);
      tSheet.getRange(rowNum, 8).setValue(mClass);
      tSheet.getRange(rowNum, 9).setValue(mReqs);
      tSheet.getRange(rowNum, 10).setValue(mCover);
      tSheet.getRange(rowNum, 11).setValue(mInstructions);
      tSheet.getRange(rowNum, 1).setValue("Submitted");
      sSheet.getRange(     i, 1).setValue("Submitted");
    }
    i+=1;
  }
}

function loadDayData() {
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Wait While We Update Your Planner ... ',20);
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var tSheet = ss.getSheetByName("Teachers");
  var cSheet = ss.getSheetByName("Cover Data");
  var pSheet = ss.getSheetByName("Planner");
  var rSheet = ss.getSheetByName("Rotation");

  pSheet.getRange(6, 16, pSheet.getLastRow()-5, 9).clearDataValidations();
  pSheet.getRange(6, 16, pSheet.getLastRow()-5, 9).clear();

  var crDayText = pSheet.getRange(1, 3).getValue();
  var crDayNumber = Number(crDayText.substr(1, 1));
  var crDayOfWeek = crDayText.substr(3,3);
  var crSDayNumber = Number(crDayText.replace("#"," ").substr(12,6))

  var pHRColName = 'HR';
  var p01ColName = rSheet.getRange(2+crDayNumber, 2).getValue();
  var p02ColName = rSheet.getRange(2+crDayNumber, 3).getValue();
  var p03ColName = rSheet.getRange(2+crDayNumber, 4).getValue();
  var p04ColName = rSheet.getRange(2+crDayNumber, 5).getValue();
  var pHRColNum  = rSheet.getRange(2+crDayNumber, 6).getValue();
  var p01ColNum  = rSheet.getRange(2+crDayNumber, 7).getValue();
  var p02ColNum  = rSheet.getRange(2+crDayNumber, 8).getValue();
  var p03ColNum  = rSheet.getRange(2+crDayNumber, 9).getValue();
  var p04ColNum  = rSheet.getRange(2+crDayNumber, 10).getValue();
  
  var tLR = lastNonBlank(tSheet, 1, 2);
  var tData = tSheet.getRange(2, 1, tLR-1, 14).getValues();
  var cLR = lastNonBlank(cSheet, 1, 2);
  var cData = cSheet.getRange(2, 1, cLR-1, 14).getValues();
  var oData = new Array();
    
  for(var ti = 0; ti < tLR-1; ti++) {
    var tName = tData[ti][0];
    var tDept = tData[ti][1];
    var tAloc = tData[ti][2];
    var tDays = 0;
    var tLast = 0;

    for(var ci = 0; ci < cLR-1; ci++) {
      if (tData[ti][0] == cData[ci][0]) { 
        tDays = cData[ci][3];
        tLast = cData[ci][4];
        ci = cLR;
      }
    }

    var tSinc = crSDayNumber - tLast;
    var tRank = tSinc * tAloc;
    var tNote = "";

    
    var tpHRText = tData[ti][pHRColNum-1];
    var tp01Text = tData[ti][p01ColNum-1]; 
    var tp02Text = tData[ti][p02ColNum-1]; 
    var tp03Text = tData[ti][p03ColNum-1]; 
    var tp04Text = tData[ti][p04ColNum-1]; 
//    Logger.log(tpHRText);

    var r = new Array();
    
    r = planRow(pHRColName, tName, tDept, tAloc, tDays, tSinc, tRank, tpHRText, tNote, 13);
    if (r.length>0) { oData.push(r); }
    r = planRow(p01ColName, tName, tDept, tAloc, tDays, tSinc, tRank, tp01Text, tNote, 10);
    if (r.length>0) { oData.push(r); }
    r = planRow(p02ColName, tName, tDept, tAloc, tDays, tSinc, tRank, tp02Text, tNote, 11);
    if (r.length>0) { oData.push(r); }
    r = planRow(p03ColName, tName, tDept, tAloc, tDays, tSinc, tRank, tp03Text, tNote, 12);
    if (r.length>0) { oData.push(r); }
    r = planRow(p04ColName, tName, tDept, tAloc, tDays, tSinc, tRank, tp04Text, tNote, 14);
    if (r.length>0) { oData.push(r); }
  }
  Logger.log(oData.length);
  
  pSheet.getRange(6, 1, pSheet.getLastRow()-5, 14).clear();
  if (oData.length+5 > pSheet.getMaxRows()) { pSheet.insertRowsAfter(pSheet.getLastRow(), oData.length+10 - pSheet.getMaxRows()); }
  pSheet.getRange(6, 1, oData.length, 14).setValues(oData);
                                             
  var formatRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('TeachersFormat');
  formatRange.copyFormatToRange(pSheet, 1, 9, 6, oData.length+5);

  updatePeriodList();
  sortData();
  loadSubmitted();
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Planner Ready!',3);
}

function planRow(rBlock, rName, rDept, rAloc, rDays, rSinc, rRank, rText, rNote, sortCol) {  
  var rowData = new Array();
  if (rText!="") {
    if (rBlock=="HR") { return rowData }
    if (rText=="No Cover") { return rowData }
    var cn = rText.substr(0,2);
    if (cn=="06" || cn=="07" || cn=="08" || cn=="09" || cn=="10" || cn=="11" || cn=="12") { return rowData }
  }
  rowData.push(rBlock);
  rowData.push(rName);
  rowData.push(rDept);
  rowData.push(rAloc);
  rowData.push(rDays);
  rowData.push(rSinc);
  rowData.push(rRank);
  rowData.push(rText);
  rowData.push(rNote);

  rowData.push(2);
  rowData.push(2);
  rowData.push(2);
  rowData.push(2);
  rowData.push(2);
  rowData[sortCol-1]=1;
  Logger.log(rName+rBlock+rText);
  return rowData;
}

function updatePeriodList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var pSheet = ss.getSheetByName("Planner");
  var rSheet = ss.getSheetByName("Rotation");
  var crDayText = pSheet.getRange(1, 3).getValue();
  var crDayNumber = Number(crDayText.substr(1, 1));
  var crDayOfWeek = crDayText.substr(3,3);

  var p1 = 'Period 1 (Block ' + rSheet.getRange(2+crDayNumber, 2).getValue()+')';
  var p2 = 'Period 2 (Block ' + rSheet.getRange(2+crDayNumber, 3).getValue()+')';
  var p3 = 'Period 3 (Block ' + rSheet.getRange(2+crDayNumber, 4).getValue()+')';
  var p4 = 'Period 4 (Block ' + rSheet.getRange(2+crDayNumber, 5).getValue()+')';
  var defaultValue = 'Homeroom';
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Homeroom', p1, p2, p3, p4]).build();
  
  if (crDayOfWeek=="Wed") {
    rule = SpreadsheetApp.newDataValidation().requireValueInList([p1, p2, p3, 'Homeroom', p4]).build();
    defaultValue = p1;
  }

  pSheet.getRange(2,3).setDataValidation(rule);
  pSheet.getRange(2,3).setValue(defaultValue);

}

function sortData() {
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Sorting Data!',10);
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var pSheet = ss.getSheetByName("Planner");
  var pVal = pSheet.getRange(2, 3).getValue();
  var sortVal = pSheet.getRange(3, 3).getValue();
  if (pVal=="Homeroom") {
    pCol = 4;
  } else {
    pCol = Number(pVal.substr(7,1));
    if (pCol==4) pCol=5;
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('', pVal,3);
  if (sortVal=="Block / Rank") {
    pSheet.getRange(6, 1, pSheet.getLastRow()-5, 14).sort([{column: 9+pCol, ascending: true}, {column: 7, ascending: false}]);
  } else {
    if (sortVal=="Block / Name") {
      pSheet.getRange(6, 1, pSheet.getLastRow()-5, 14).sort([{column: 9+pCol, ascending: true}, {column: 2, ascending: true}]);
    } else {
      pSheet.getRange(6, 1, pSheet.getLastRow()-5, 14).sort([{column: 9+pCol, ascending: true}, {column: 3, ascending: true}, {column: 7, ascending: false}]);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Planner Ready!',3);
  return;  
}
