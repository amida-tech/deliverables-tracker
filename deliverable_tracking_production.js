function bulkAdd() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var numRow = ss.getLastRow();
  var employeeInfo = getEmployeeValues(ss);
  var deliverableValues = getDeliverableValues(ss);
  var checkRange = sheet.getRange(12, 1, numRow);
  var concatRange = sheet.getRange(12, 29, numRow);
  var concatValues = concatRange.getValues();
  var oneChecked = false; // Variable that checks if at least one box has been checked
  for (var row = 0; row < deliverableValues.length; row++) {
    if (deliverableValues[row][0]) {
      oneChecked = true;
      var randomNum = Math.floor(100000000 + Math.random() * 900000000); // Creates 9-Digit Unique ID
      var randRange = sheet.getRange(row + 12, 26); // Range to add Unique ID
      var added = false,
        incomplete = false;
      if (randRange.getValue()) {
        added = true;
        Browser.msgBox('Deliverable has already been added.');
        break;
      }
      for (var i = 0; i < deliverableValues[row].length; i++) {
        if (deliverableValues[row][i] == '' && i != 23 && i != 25) {
          incomplete = true;
          break;
        }
      }
      if (incomplete) {
        Browser.msgBox('Row is not complete.');
        break;
      }
      randRange.setValue(randomNum); // Adds Unique ID
      deliverableValues = getDeliverableValues(ss);
      var deliverableRowData = getDeliverableRowData(
        deliverableValues,
        row,
        employeeInfo
      ); //TODO: Add space after comma

      addToMaster(deliverableRowData, employeeInfo, concatValues, row, ss); // Adds all info to Master Sheet

      var programManagerEvent = addProgramManagerEvent(
        employeeInfo,
        deliverableRowData
      );
      var editorEvent = addEditorEvent(employeeInfo, deliverableRowData);

      if (deliverableRowData.tier == 'Tier 1') {
        var cLevelEvent = addCLevelEvent(employeeInfo, deliverableRowData);
        var ceoEvent = addCeoEvent(employeeInfo, deliverableRowData);
        var customerEvent = addCustomerEvent(deliverableRowData);
        var recipientList = getRecipientList(3, employeeInfo); // TODO: Comment 3 is yes/no column for tier 1 email
        sendTier1AddEmail(recipientList, employeeInfo, deliverableRowData);
      } else if (deliverableRowData.tier == 'Tier 2') {
        var cLevelEvent = addCLevelEvent(employeeInfo, deliverableRowData);
        var customerEvent = addCustomerEvent(deliverableRowData);
        var recipientList = getRecipientList(4, employeeInfo); // TODO: Comment 4 is yes/no column for tier 2 email
        sendTier2AddEmail(recipientList, employeeInfo, deliverableRowData);
      } else if (deliverableRowData.tier == 'Tier 3') {
        var customerEvent = addCustomerEvent(deliverableRowData);
        var recipientList = getRecipientList(5, employeeInfo); // TODO: Comment 5 is yes/no column for tier 3 email
        sendTier3AddEmail(recipientList, employeeInfo, deliverableRowData);
      }

      // Adding PMs and Program Manager to Prog Man Event
      addProgramManagerEventGuests(programManagerEvent, employeeInfo);
      addEditorEventGuests(editorEvent, employeeInfo);
      addCustomerEventGuests(customerEvent, employeeInfo);
      if (
        deliverableRowData.tier == 'Tier 1' ||
        deliverableRowData.tier == 'Tier 2'
      ) {
        addCLevelEventGuests(cLevelEvent, employeeInfo);
      }
      if (deliverableRowData.tier == 'Tier 1') {
        addCeoEventGuests(ceoEvent, employeeInfo);
      }
      Browser.msgBox('Row ' + (row + 12) + ' was added succesfully.');
      break;
    }
  }
  checkRange.setValue(false); // Resets all checboxes to unchecked
  removeEmptyRows(); // Temp fix to google sheets automatically adding rows
  for (var i = 0; i < 1; i++) {
    addRow();
  }
}

function bulkEdit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var masterSheet = ss.getSheetByName('Master');
  var employeeInfo = getEmployeeValues(ss);
  var numRow = ss.getLastRow();
  var checkRange = sheet.getRange(12, 1, numRow);
  var deliverableValues = getDeliverableValues(ss);
  var idValues = getIdValues(ss);
  var masterValues = getMasterValues(ss);
  var employeeInfo = getEmployeeValues(ss);
  var changed = false;
  var concatRange = sheet.getRange(12, 29, numRow);
  var concatValues = concatRange.getValues();
  var oneChecked = false;
  var masterChange;
  for (var row = 0; row < deliverableValues.length; row++) {
    var incomplete = false;
    if (deliverableValues[row][0]) {
      
      for (var i = 0; i < deliverableValues[row].length; i++) {
        if (deliverableValues[row][i] == '' && i != 23 && i != 25 && i != 18) {
          incomplete = true;
          break;
        }
      }
      if (incomplete) {
        continue;
      }
      if (!checkIfChanged(row)) {
      Browser.msgBox('Nothing was changed. Please try again.');
      return;
    }
      masterChange = deliverableValues[row][25];
    }
  }

  for (var row = masterValues.length - 1; row >= 0; row--) {
    if (masterChange == masterValues[row][27]) {
      masterSheet.deleteRow(row + 3);
    }
  }
  for (var row = 0; row < deliverableValues.length; row++) {
    var deliverableRowData = getDeliverableRowData(
      deliverableValues,
      row,
      employeeInfo
    );
    if (deliverableValues[row][0]) {
      oneChecked = true;
      for (var i = 0; i < deliverableValues[row].length; i++) {
        if (deliverableValues[row][i] == '' && i != 23 && i != 25 && i != 18) {
          incomplete = true;
          break;
        }
      }
      if (incomplete) {
        continue;
      }
      var update = updateDeliverables(
        row,
        deliverableRowData,
        deliverableValues,
        idValues,
        employeeInfo
      );
      if (deliverableRowData.tier == 'Tier 1') {
        var recipientList = getRecipientList(3, employeeInfo); // TODO: Comment 3 is yes/no column for tier 1 email
        sendTier1UpdateEmail(recipientList, employeeInfo, deliverableRowData);
      }
      if (deliverableRowData.tier == 'Tier 2') {
        var recipientList = getRecipientList(4, employeeInfo); // TODO: Comment 4 is yes/no column for tier 2 email
        sendTier2UpdateEmail(recipientList, employeeInfo, deliverableRowData);
      }
      if (deliverableRowData.tier == 'Tier 3') {
        var recipientList = getRecipientList(5, employeeInfo); // TODO: Comment 5 is yes/no column for tier 3 email
        sendTier3UpdateEmail(recipientList, employeeInfo, deliverableRowData);
      }

      addToMaster(deliverableRowData, employeeInfo, concatValues, row, ss);
      Browser.msgBox('Row' + (row + 12) + ' was successfully updated.');
      break;
    }
  }
  displayNoneCheckedMessage(oneChecked);
  checkRange.setValue(false);
  removeEmptyRows(); // Resets all checkboxes to empty
}
function bulkDelete() {
  var ui = SpreadsheetApp.getUi(); // Spreadsheet UI for buttons
  if (sureYouWantToDelete(ui) == 'NO') {
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var masterSheet = ss.getSheetByName('Master');
  var employeeInfo = getEmployeeValues(ss);
  var numRow = ss.getLastRow();
  var checkRange = sheet.getRange(12, 1, numRow);
  var deliverableValues = getDeliverableValues(ss);
  var idValues = getIdValues(ss);
  var masterValues = getMasterValues(ss);
  var oneChecked = false;
  var deleted;
  for (var row = deliverableValues.length - 1; row >= 0; row--) {
    var deliverableRowData = getDeliverableRowData(
      deliverableValues,
      row,
      employeeInfo
    );
    if (deliverableValues[row][0]) {
      oneChecked = true;
      deleteEvents(idValues, deliverableRowData, row);
      var deleted = deliverableRowData.id;
      if (deliverableRowData.tier == 'Tier 1') {
        var recipientList = getRecipientList(3, employeeInfo);
        sendTier1DeleteEmail(recipientList, employeeInfo, deliverableRowData);
      } else if (deliverableRowData.tier == 'Tier 2') {
        var recipientList = getRecipientList(4, employeeInfo);
        sendTier2DeleteEmail(recipientList, employeeInfo, deliverableRowData);
      } else if (deliverableRowData.tier == 'Tier 3') {
        var recipientList = getRecipientList(5, employeeInfo);
        sendTier3DeleteEmail(recipientList, employeeInfo, deliverableRowData);
      }
      Browser.msgBox('Row was successfully deleted.');
      deleteFromMaster(masterValues, deleted);
      break;
    }
  }
  finishFunction(checkRange, oneChecked);
}
function bulkArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi(); // Spreadsheet UI for buttons
  var employeeInfo = getEmployeeValues(ss);
  var deliverableValues = getDeliverableValues(ss);
  var masterValues = getMasterValues(ss);
  var numRow = ss.getLastRow();
  var checkRange = sheet.getRange(12, 1, numRow);
  var oneChecked = false;
  var archived;
  for (var row = deliverableValues.length - 1; row >= 0; row--) {
    var deliverableRowData = getDeliverableRowData(
      deliverableValues,
      row,
      employeeInfo
    );
    if (sureYouWantToArchive(ui, deliverableRowData) == 'NO') {
      return;
    }
    if (deliverableValues[row][0]) {
      oneChecked = true;
      archived = archiveDeliverable(deliverableRowData, row, employeeInfo);
      var recipientList = getRecipientList(6, employeeInfo);
      sendArchiveEmail(recipientList, employeeInfo, deliverableRowData);
      Browser.msgBox('Row ' + row + ' was successfully archived.');
      break;
    }
  }
  deleteFromMaster(masterValues, archived);
  finishFunction(checkRange, oneChecked);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//HELPER FUNCTIONS BELOW////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


var ss = SpreadsheetApp.getActive(); // Active SpreadSheet



function checkIfAlreadyAdded(alreadyAdded, idField, row) {
  if (idField.getValue()) {
    added = true;
    alreadyAdded.push(row + 12);
  }
  return {
    alreadyAdded: alreadyAdded,
    added: added
  };
}
// TODO: Add comment explaining why this function exists
function addRow() {
  var sheet = ss.getActiveSheet(),
    numRow = sheet.getLastRow();
  var numCol = sheet.getLastColumn(),
    range = ss.getSheetByName('Row Format').getRange(1, 1, 1, numCol);
  sheet.insertRowsAfter(numRow, 1);
  range.copyTo(sheet.getRange(numRow + 1, 1, 1, numCol), {
    contentsOnly: false
  });
}

// Add link to onOpen() Google documentation
// TODO: Add comment explaining why this function exists
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Sheet Functions');
  var addProject = menu.addItem('Add', 'bulkAdd');
  menu.addSeparator();
  var editProject = menu.addItem('Update', 'bulkEdit');
  menu.addSeparator();
  var deleteProject = menu.addItem('Delete', 'bulkDelete');
  menu.addSeparator();
  var archiveProject = menu.addItem('Archive', 'bulkArchive');
  menu.addSeparator();
  var addRows = menu.addItem('New Row', 'addRow');
  addProject.addToUi();
  editProject.addToUi();
  deleteProject.addToUi();
  archiveProject.addToUi();
  addRows.addToUi();
}

//TODO: Change function name to remove formatted rows
//TODO: Change 12,8 to variables
function removeEmptyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = 0;
  var sheet = ss.getActiveSheet();
  var maxRows = sheet.getMaxRows();
  var range = sheet.getRange(12, 8, sheet.getLastRow());
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0].length) {
      lastRow = i + 11; // Change 11 to var
      break;
    }
  }
  if (maxRows - lastRow != 0) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
}

// TODO: make a function approximately like this
// TODO: In comment, say idValues is a 2D array. Meh, maybe.
function getIdValues(ss) {
  var idSheet = ss.getSheetByName('Event IDs');
  var idNumCol = idSheet.getLastColumn();
  var idNumRow = idSheet.getLastRow();
  var idRange = idSheet.getRange(1, 1, idNumRow, idNumCol); // Range of Event IDs
  var idValues = idRange.getValues();

  return idValues;
}

function getMasterValues(ss) {
  var masterSheet = ss.getSheetByName('Master');
  var masterNumCol = masterSheet.getLastColumn();
  var masterNumRow = masterSheet.getLastRow();
  var masterRange = masterSheet.getRange(3, 1, masterNumRow, masterNumCol);
  var masterValues = masterRange.getValues();

  return masterValues;
}
//TODO:COMMENT
function getEmployeeValues(ss) {
  var sheet = ss.getActiveSheet();
  var projNameRange = sheet.getRange(8, 3); //TODO: Comment what 8 and 3 are
  var projName = projNameRange.getValue();
  var firstRow = 2,
    firstCol = 23,
    lastRow = 10,
    lastCol = 33;
  var employeeRange = sheet.getRange(firstRow, firstCol, lastRow, lastCol); // Range of employee names/positions
  var employeeValues = employeeRange.getValues();
  var editorName = employeeValues[5][1],
    pmName = employeeValues[6][1],
    backupPmName = employeeValues[7][1],
    progManName = employeeValues[8][1],
    ceoName = employeeValues[1][1],
    cooName = employeeValues[2][1];
  return {
    editorName: editorName,
    pmName: pmName,
    backupPmName: backupPmName,
    progManName: progManName,
    ceoName: ceoName,
    cooName: cooName,
    employeeValues: employeeValues,
    projName: projName
  };
}

//TODO:COMMENT
function getDeliverableValues(ss) {
  var sheet = ss.getActiveSheet();
  var firstRow = 12,
    firstCol = 1,
    lastRow = ss.getLastRow(),
    lastCol = 27;
  var deliverablesRange = sheet.getRange(firstRow, firstCol, lastRow, lastCol); // Range of actual deliverables
  var deliverablesValues = deliverablesRange.getValues();

  return deliverablesValues;
}

//TODO:COMMENT
function getDeliverableRowData(deliverableValues, rowNumber, employeeInfo) {
  var delivName = deliverableValues[rowNumber][2],
    delivType = deliverableValues[rowNumber][3],
    pages = deliverableValues[rowNumber][4],
    tier = deliverableValues[rowNumber][5],
    dbProgMan = deliverableValues[rowNumber][6],
    dayProgMan = deliverableValues[rowNumber][7],
    dateProgMan = deliverableValues[rowNumber][8],
    timeProgMan = deliverableValues[rowNumber][9],
    dbEditor = deliverableValues[rowNumber][10],
    dayEditor = deliverableValues[rowNumber][11],
    dateEditor = deliverableValues[rowNumber][12],
    timeEditor = deliverableValues[rowNumber][13],
    dbThird = deliverableValues[rowNumber][14],
    dayThird = deliverableValues[rowNumber][15],
    dateThird = deliverableValues[rowNumber][16],
    timeThird = deliverableValues[rowNumber][17],
    dbCeo = deliverableValues[rowNumber][18],
    dayCeo = deliverableValues[rowNumber][19],
    dateCeo = deliverableValues[rowNumber][20],
    timeCeo = deliverableValues[rowNumber][21],
    dateCustomer = deliverableValues[rowNumber][22],
    notes = deliverableValues[rowNumber][23],
    status = deliverableValues[rowNumber][24],
    id = deliverableValues[rowNumber][25];
  var progManDate = new Date(deliverableValues[rowNumber][8]);
  var progManDesc =
    'Due to ' + employeeInfo.progManName + ': ' + delivName + progManDate;
  var editorDate = new Date(deliverableValues[rowNumber][12]);
  var editorDesc =
    'Due to ' + employeeInfo.editorName + ': ' + delivName + editorDate;
  var cLevelDate = new Date(deliverableValues[rowNumber][16]);
  var cLevelDesc =
    'Due to ' + employeeInfo.cooName + ': ' + delivName + cLevelDate;
  var ceoDate = new Date(deliverableValues[rowNumber][20]);
  var ceoDesc = 'Due to ' + employeeInfo.ceoName + ': ' + delivName + ceoDate;
  var delivDate = new Date(deliverableValues[rowNumber][22]);
  var notifyPeterDate = new Date(deliverableValues[rowNumber][26]);
  return {
    delivName: delivName,
    delivType: delivType,
    pages: pages,
    tier: tier,
    dbProgMan: dbProgMan,
    dayProgMan: dayProgMan,
    dateProgMan: dateProgMan,
    timeProgMan: timeProgMan,
    dbEditor: dbEditor,
    dayEditor: dayEditor,
    dateEditor: dateEditor,
    timeEditor: timeEditor,
    dbThird: dbThird,
    dayThird: dayThird,
    dateThird: dateThird,
    timeThird: timeThird,
    dbCeo: dbCeo,
    dayCeo: dayCeo,
    dateCeo: dateCeo,
    timeCeo: timeCeo,
    dateCustomer: dateCustomer,
    notes: notes,
    status: status,
    id: id,
    progManDate: progManDate,
    progManDesc: progManDesc,
    editorDate: editorDate,
    editorDesc: editorDesc,
    cLevelDate: cLevelDate,
    cLevelDesc: cLevelDesc,
    ceoDate: ceoDate,
    ceoDesc: ceoDesc,
    delivDate: delivDate,
    notifyPeterDate: notifyPeterDate
  };
}

//TODO:COMMENT
function addToMaster(
  deliverableRowData,
  employeeInfo,
  concatValues,
  rowNumber,
  ss
) {
  var masterSheet = ss.getSheetByName('Master');
  masterSheet.appendRow([
    employeeInfo.projName,
    employeeInfo.pmName,
    employeeInfo.backupPmName,
    employeeInfo.progManName,
    deliverableRowData.delivName,
    deliverableRowData.delivType,
    deliverableRowData.pages,
    deliverableRowData.tier,
    deliverableRowData.dbProgMan,
    deliverableRowData.dayProgMan,
    deliverableRowData.dateProgMan,
    deliverableRowData.timeProgMan,
    deliverableRowData.dbEditor,
    deliverableRowData.dayEditor,
    deliverableRowData.dateEditor,
    deliverableRowData.timeEditor,
    deliverableRowData.dbThird,
    deliverableRowData.dayThird,
    deliverableRowData.dateThird,
    deliverableRowData.timeThird,
    deliverableRowData.dbCeo,
    deliverableRowData.dayCeo,
    deliverableRowData.dateCeo,
    deliverableRowData.timeCeo,
    deliverableRowData.dateCustomer,
    deliverableRowData.notes,
    deliverableRowData.status,
    deliverableRowData.id,
    deliverableRowData.notifyPeterDate,
    7,
    concatValues[rowNumber][0]
  ]);
}

function getRecipientList(col, employeeInfo) {
  var recipientList = '';
  for (var i = 0; i < employeeInfo.employeeValues.length; i++) {
    if (employeeInfo.employeeValues[i][col] == 'Yes') {
      recipientList = recipientList + employeeInfo.employeeValues[i][2];
      if (i < employeeInfo.employeeValues.length - 1) {
        recipientList = recipientList + ',';
      }
    }
  }
  return recipientList;
}

function createCalendarEvent(title, date, cal, name) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var event = cal.createAllDayEvent('Due to ' + name + ': ' + title, date);
  event.setDescription(
    'Link to Project/Deliverable Tracker\n' + sheet.getRange('C6').getValue()
  );

  return event;
}

function addProgramManagerEvent(employeeInfo, deliverableRowData) {
var deliverablesGoogleCalendarUrl =
  ss.getActiveSheet().getRange('Y2').getValue();
// TODO: Replace with this variable throughout the code.
var deliverablesGoogleCalendar = CalendarApp.getCalendarById(
  deliverablesGoogleCalendarUrl
);
  var ss = SpreadsheetApp.getActive();
  var programManagerEvent = createCalendarEvent(
    deliverableRowData.delivName,
    deliverableRowData.progManDate,
    deliverablesGoogleCalendar,
    employeeInfo.progManName
  );
  ss.getSheetByName('Event IDs').appendRow([
    deliverableRowData.id + ' FIRST ' + deliverableRowData.progManDesc,
    programManagerEvent.getId()
  ]);

  return programManagerEvent;
}

function addEditorEvent(employeeInfo, deliverableRowData) {
  var ss = SpreadsheetApp.getActive();
  var deliverablesGoogleCalendarUrl =
  ss.getActiveSheet().getRange('Y2').getValue();
// TODO: Replace with this variable throughout the code.
var deliverablesGoogleCalendar = CalendarApp.getCalendarById(
  deliverablesGoogleCalendarUrl
);
  var editorEvent = createCalendarEvent(
    deliverableRowData.delivName,
    deliverableRowData.editorDate,
    deliverablesGoogleCalendar,
    employeeInfo.editorName
  );
  ss.getSheetByName('Event IDs').appendRow([
    deliverableRowData.id + ' SECOND ' + deliverableRowData.editorDesc,
    editorEvent.getId()
  ]);

  return editorEvent;
}

function addCLevelEvent(employeeInfo, deliverableRowData) {
  var ss = SpreadsheetApp.getActive();
  var deliverablesGoogleCalendarUrl =
  ss.getActiveSheet().getRange('Y2').getValue();
// TODO: Replace with this variable throughout the code.
var deliverablesGoogleCalendar = CalendarApp.getCalendarById(
  deliverablesGoogleCalendarUrl
);
  var cLevelEvent = createCalendarEvent(
    deliverableRowData.delivName,
    deliverableRowData.cLevelDate,
    deliverablesGoogleCalendar,
    employeeInfo.cooName
  );
  ss.getSheetByName('Event IDs').appendRow([
    deliverableRowData.id + ' THIRD ' + deliverableRowData.cLevelDesc,
    cLevelEvent.getId()
  ]);

  return cLevelEvent;
}

function addCeoEvent(employeeInfo, deliverableRowData) {
  var ss = SpreadsheetApp.getActive();
  var deliverablesGoogleCalendarUrl =
  ss.getActiveSheet().getRange('Y2').getValue();
// TODO: Replace with this variable throughout the code.
var deliverablesGoogleCalendar = CalendarApp.getCalendarById(
  deliverablesGoogleCalendarUrl
);
  var ceoEvent = createCalendarEvent(
    deliverableRowData.delivName,
    deliverableRowData.ceoDate,
    deliverablesGoogleCalendar,
    employeeInfo.ceoName
  );
  ss.getSheetByName('Event IDs').appendRow([
    deliverableRowData.id + ' FOURTH ' + deliverableRowData.ceoDesc,
    ceoEvent.getId()
  ]);

  return ceoEvent;
}

function addCustomerEvent(deliverableRowData, ss) {
  var ss = SpreadsheetApp.getActive();
  var deliverablesGoogleCalendarUrl =
  ss.getActiveSheet().getRange('Y2').getValue();
// TODO: Replace with this variable throughout the code.
var deliverablesGoogleCalendar = CalendarApp.getCalendarById(
  deliverablesGoogleCalendarUrl
);
  var customerEvent = createCalendarEvent(
    deliverableRowData.delivName,
    deliverableRowData.delivDate,
    deliverablesGoogleCalendar,
    'Customer'
  );
  ss.getSheetByName('Event IDs').appendRow([
    deliverableRowData.id + ' Customer ',
    customerEvent.getId()
  ]);

  return customerEvent;
}

function sendTier1AddEmail(recipientList, employeeInfo, deliverableRowData) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  MailApp.sendEmail(
    recipientList,
    employeeInfo.projName +
      ' ' +
      deliverableRowData.delivName +
      ' Deliverable Timeline Due - ' +
      deliverableRowData.delivDate.toDateString().substring(4, 10),
    'A new deliverable has been added to the Amida Deliverables Tracker\n\nProject: ' +
      employeeInfo.projName +
      '\nDeliverable: ' +
      deliverableRowData.delivName +
      '\nDue to Customer: ' +
      deliverableRowData.delivDate.toDateString() +
      '\nDeliverable Type: ' +
      deliverableRowData.delivType +
      '\nEst. Pages: ' +
      deliverableRowData.pages +
      '\n\nPM: ' +
      employeeInfo.pmName +
      '\nBackup PM: ' +
      employeeInfo.backupPmName +
      '\nProgram Manager: ' +
      employeeInfo.progManName +
      '\n\nReview Timeline:\nThis is a ' +
      deliverableRowData.tier +
      ' Deliverable\nDue to ' +
      employeeInfo.progManName +
      ': ' +
      deliverableRowData.progManDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.editorName +
      ': ' +
      deliverableRowData.editorDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.cooName +
      ': ' +
      deliverableRowData.cLevelDate.toDateString() +
      ', 5pm ET\nDue to ' +
      employeeInfo.ceoName +
      ': ' +
      deliverableRowData.ceoDate.toDateString() +
      ', 5pm ET (Notify on ' +
      deliverableRowData.notifyPeterDate.toDateString() +
      ')\nStatus: ' +
      deliverableRowData.status +
      '\nNotes: ' +
      deliverableRowData.notes +
      "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
      employeeInfo.pmName +
      '.\n- Any timeline changes should be made on the ' +
      employeeInfo.projName +
      ' Tab of the Deliverables Tracker ' +
      sheet.getRange('C6').getValue()
  );
}

function sendTier2AddEmail(recipientList, employeeInfo, deliverableRowData) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  MailApp.sendEmail(
    recipientList,
    employeeInfo.projName +
      ' ' +
      deliverableRowData.delivName +
      ' Deliverable Timeline Due - ' +
      deliverableRowData.delivDate.toDateString().substring(4, 10),
    'A new deliverable has been added to the Amida Deliverables Tracker\n\nProject: ' +
      employeeInfo.projName +
      '\nDeliverable: ' +
      deliverableRowData.delivName +
      '\nDue to Customer: ' +
      deliverableRowData.delivDate.toDateString() +
      '\nDeliverable Type: ' +
      deliverableRowData.delivType +
      '\nEst. Pages: ' +
      deliverableRowData.pages +
      '\n\nPM: ' +
      employeeInfo.pmName +
      '\nBackup PM: ' +
      employeeInfo.backupPmName +
      '\nProgram Manager: ' +
      employeeInfo.progManName +
      '\n\nReview Timeline:\nThis is a ' +
      deliverableRowData.tier +
      ' Deliverable\nDue to ' +
      employeeInfo.progManName +
      ': ' +
      deliverableRowData.progManDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.editorName +
      ': ' +
      deliverableRowData.editorDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.cooName +
      ': ' +
      deliverableRowData.cLevelDate.toDateString() +
      ', 5pm ET\nDue to ' +
      employeeInfo.ceoName +
      ': N/A\nStatus: ' +
      deliverableRowData.status +
      '\nNotes: ' +
      deliverableRowData.notes +
      "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
      employeeInfo.pmName +
      '.\n- Any timeline changes should be made on the ' +
      employeeInfo.projName +
      ' Tab of the Deliverables Tracker ' +
      sheet.getRange('C6').getValue()
  );
}

function sendTier3AddEmail(recipientList, employeeInfo, deliverableRowData) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  MailApp.sendEmail(
    recipientList,
    employeeInfo.projName +
      ' ' +
      deliverableRowData.delivName +
      ' Deliverable Timeline Due - ' +
      deliverableRowData.delivDate.toDateString().substring(4, 10),
    'A new deliverable has been added to the Amida Deliverables Tracker\n\nProject: ' +
      employeeInfo.projName +
      '\nDeliverable: ' +
      deliverableRowData.delivName +
      '\nDue to Customer: ' +
      deliverableRowData.delivDate.toDateString() +
      '\nDeliverable Type: ' +
      deliverableRowData.delivType +
      '\nEst. Pages: ' +
      deliverableRowData.pages +
      '\n\nPM: ' +
      employeeInfo.pmName +
      '\nBackup PM: ' +
      employeeInfo.backupPmName +
      '\nProgram Manager: ' +
      employeeInfo.progManName +
      '\n\nReview Timeline:\nThis is a ' +
      deliverableRowData.tier +
      ' Deliverable\n\nDue to ' +
      employeeInfo.progManName +
      ': ' +
      deliverableRowData.progManDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.editorName +
      ': ' +
      deliverableRowData.editorDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.cooName +
      ': N/A\nDue to ' +
      employeeInfo.ceoName +
      ': N/A\nStatus: ' +
      deliverableRowData.status +
      '\nNotes: ' +
      deliverableRowData.notes +
      "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
      employeeInfo.pmName +
      '.\n- Any timeline changes should be made on the ' +
      employeeInfo.projName +
      ' Tab of the Deliverables Tracker ' +
      sheet.getRange('C6').getValue()
  );
}

function addProgramManagerEventGuests(progManEvent, employeeInfo) {
  progManEvent.addGuest(employeeInfo.employeeValues[6][2]);
  progManEvent.addGuest(employeeInfo.employeeValues[7][2]);
  progManEvent.addGuest(employeeInfo.employeeValues[8][2]);
}

function addEditorEventGuests(editorEvent, employeeInfo) {
  for (var i = 1; i < employeeInfo.employeeValues.length; i++) {
    if (employeeInfo.employeeValues[i][7] == 'Yes' && (i != 9 || i != 8)) {
      editorEvent.addGuest(employeeInfo.employeeValues[i][2]);
    }
  }
}

function addCLevelEventGuests(cLevelEvent, employeeInfo) {
  for (var i = 1; i < employeeInfo.employeeValues.length; i++) {
    if (employeeInfo.employeeValues[i][8] == 'Yes') {
      cLevelEvent.addGuest(employeeInfo.employeeValues[i][2]);
    }
  }
}

function addCeoEventGuests(ceoEvent, employeeInfo) {
  for (var i = 1; i < employeeInfo.employeeValues.length; i++) {
    if (employeeInfo.employeeValues[i][9] == 'Yes') {
      ceoEvent.addGuest(employeeInfo.employeeValues[i][2]);
    }
  }
}

function addCustomerEventGuests(customerEvent, employeeInfo) {
  for (var i = 1; i < employeeInfo.employeeValues.length; i++) {
    if (employeeInfo.employeeValues[i][10] == 'Yes') {
      customerEvent.addGuest(employeeInfo.employeeValues[i][2]);
    }
  }
}

function displayAddSuccessMessage(addSuccess) {
  if (addSuccess.length == 1) {
    Browser.msgBox(
      'This row was successfully added \n' + JSON.stringify(addSuccess)
    );
  } else if (addSuccess.length > 1) {
    Browser.msgBox(
      'These rows were successfully added \n' + JSON.stringify(addSuccess)
    );
  }
}

function displayAlreadyAddedMessage(alreadyAdded) {
  if (alreadyAdded.length == 1) {
    Browser.msgBox(
      'This row was already added previously\n' + JSON.stringify(alreadyAdded)
    );
  } else if (alreadyAdded.length > 1) {
    Browser.msgBox(
      'These rows were already added previously\n' +
        JSON.stringify(alreadyAdded)
    );
  }
}

function displayNotFilledMessage(notFilled) {
  if (notFilled.length == 1) {
    Browser.msgBox('This row was not complete\n' + JSON.stringify(notFilled));
  } else if (notFilled.length > 1) {
    Browser.msgBox(
      'These rows were not complete\n' + JSON.stringify(notFilled)
    );
  }
}

function displayNoneCheckedMessage(oneChecked) {
  if (!oneChecked) {
    Browser.msgBox('No rows were checked. Please try again.');
    return;
  }
}

function updateDeliverables(
  row,
  deliverableRowData,
  deliverableValues,
  idValues,
  employeeInfo
) {
  var idSheet = ss.getSheetByName('Event IDs');
  var existing = false;
  for (var i = 0; i < idValues.length; i++) {
    if (idValues[i][0].substring(0, 9) == deliverableRowData.id) {
      if (idValues[i][0].indexOf('FIRST') !== -1) {
        var tempEvent = deliverablesGoogleCalendar.getEventById(idValues[i][1]);
        var oldDate = tempEvent.getAllDayStartDate();
        tempEvent.setTitle(
          'Due to ' +
            employeeInfo.progManName +
            ': ' +
            deliverableRowData.delivName
        );
        existing = true;
        if (
          oldDate.getTime() != new Date(deliverableValues[row][8]).getTime()
        ) {
          tempEvent.setAllDayDate(new Date(deliverableValues[row][8]));
          idSheet.getRange(i + 1, 1).setValue('EVENT CHANGED');
          idSheet.getRange(i + 1, 2).setValue('EVENT CHANGED');
          idSheet.appendRow([
            deliverableRowData.id + ' FIRST ' + deliverableRowData.progManDesc,
            tempEvent.getId()
          ]);
        }
      }
      if (idValues[i][0].indexOf('SECOND') !== -1) {
        var tempEvent = deliverablesGoogleCalendar.getEventById(idValues[i][1]);
        tempEvent.setTitle(
          'Due to ' +
            employeeInfo.editorName +
            ': ' +
            deliverableRowData.delivName
        );
        var oldDate = tempEvent.getAllDayStartDate();
        existing = true;
        if (
          oldDate.getTime() != new Date(deliverableValues[row][12]).getTime()
        ) {
          tempEvent.setAllDayDate(new Date(deliverableValues[row][12]));
          idSheet.getRange(i + 1, 1).setValue('EVENT CHANGED');
          idSheet.getRange(i + 1, 2).setValue('EVENT CHANGED');
          idSheet.appendRow([
            deliverableRowData.id + ' SECOND ' + deliverableRowData.editorDesc,
            tempEvent.getId()
          ]);
        }
      }
      if (idValues[i][0].indexOf('THIRD') !== -1) {
        var tempEvent = deliverablesGoogleCalendar.getEventById(idValues[i][1]);
        var oldDate = tempEvent.getAllDayStartDate();
        tempEvent.setTitle(
          'Due to ' + employeeInfo.cooName + ': ' + deliverableRowData.delivName
        );
        if (
          oldDate.getTime() != new Date(deliverableValues[row][16]).getTime()
        ) {
          tempEvent.setAllDayDate(new Date(deliverableValues[row][16]));
          idSheet.getRange(i + 1, 1).setValue('EVENT CHANGED');
          idSheet.getRange(i + 1, 2).setValue('EVENT CHANGED');
          idSheet.appendRow([
            deliverableRowData.id + ' THIRD ' + deliverableRowData.cLevelDesc,
            tempEvent.getId()
          ]);
        }
      }
      if (idValues[i][0].indexOf('FOURTH') !== -1) {
        var tempEvent = deliverablesGoogleCalendar.getEventById(idValues[i][1]);
        var oldDate = tempEvent.getAllDayStartDate();
        tempEvent.setTitle(
          'Due to ' + employeeInfo.ceoName + ': ' + deliverableRowData.delivName
        );
        if (
          oldDate.getTime() != new Date(deliverableValues[row][20]).getTime()
        ) {
          tempEvent.setAllDayDate(new Date(deliverableValues[row][20]));
          idSheet.getRange(i + 1, 1).setValue('EVENT CHANGED');
          idSheet.getRange(i + 1, 2).setValue('EVENT CHANGED');
          idSheet.appendRow([
            deliverableRowData.id + ' FOURTH ' + deliverableRowData.ceoDesc,
            tempEvent.getId()
          ]);
        }
      }
      if (idValues[i][0].indexOf('Customer') !== -1) {
        var tempEvent = deliverablesGoogleCalendar.getEventById(idValues[i][1]);
        var oldDate = tempEvent.getAllDayStartDate();
        tempEvent.setTitle('Due to Customer: ' + deliverableRowData.delivName);
        if (
          oldDate.getTime() != new Date(deliverableValues[row][22]).getTime()
        ) {
          tempEvent.setAllDayDate(new Date(deliverableValues[row][22]));
          idSheet.getRange(i + 1, 1).setValue('EVENT CHANGED');
          idSheet.getRange(i + 1, 2).setValue('EVENT CHANGED');
          idSheet.appendRow([
            deliverableRowData.id + ' Customer ',
            tempEvent.getId()
          ]);
        }
      }
    }
  }
  return {
    existing: existing
  };
}

function sendTier1UpdateEmail(recipientList, employeeInfo, deliverableRowData) {
  var sheet = ss.getActiveSheet();
  MailApp.sendEmail(
    recipientList,
    'Updated Deliverable Timeline: ' +
      employeeInfo.projName +
      ' ' +
      deliverableRowData.delivName +
      ' Due - ' +
      deliverableRowData.delivDate.toDateString().substring(4, 10),
    'Updates have been made to the following deliverable in the Amida Deliverables Tracker\nProject: ' +
      employeeInfo.projName +
      '\nDeliverable: ' +
      deliverableRowData.delivName +
      '\nDue to Customer: ' +
      deliverableRowData.delivDate.toDateString() +
      '\nDeliverable Type: ' +
      deliverableRowData.delivType +
      '\nEst. Pages: ' +
      deliverableRowData.pages +
      '\n\nPM: ' +
      employeeInfo.pmName +
      '\nBackup PM: ' +
      employeeInfo.backupPmName +
      '\nProgram Manager: ' +
      employeeInfo.progManName +
      '\n\nReview Timeline:\nThis is a ' +
      deliverableRowData.tier +
      ' Deliverable\nDue to ' +
      employeeInfo.progManName +
      ': ' +
      deliverableRowData.progManDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.editorName +
      ': ' +
      deliverableRowData.editorDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.cooName +
      ': ' +
      deliverableRowData.cLevelDate.toDateString() +
      ', 5pm ET\nDue to ' +
      employeeInfo.ceoName +
      ': ' +
      deliverableRowData.ceoDate.toDateString() +
      ', 5pm ET (Notify on ' +
      deliverableRowData.notifyPeterDate.toDateString() +
      ')\nStatus: ' +
      deliverableRowData.status +
      '\nNotes: ' +
      deliverableRowData.notes +
      "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
      employeeInfo.pmName +
      '.\n- Any timeline changes should be made on the ' +
      employeeInfo.projName +
      ' Tab of the Deliverables Tracker ' +
      sheet.getRange('C6').getValue()
  );
}

function sendTier2UpdateEmail(recipientList, employeeInfo, deliverableRowData) {
  var sheet = ss.getActiveSheet();
  MailApp.sendEmail(
    recipientList,
    'Updated Deliverable Timeline: ' +
      employeeInfo.projName +
      ' ' +
      deliverableRowData.delivName +
      ' Due - ' +
      deliverableRowData.delivDate.toDateString().substring(4, 10),
    'Updates have been made to the following deliverable in the Amida Deliverables Tracker\nProject: ' +
      employeeInfo.projName +
      '\nDeliverable: ' +
      deliverableRowData.delivName +
      '\nDue to Customer: ' +
      deliverableRowData.delivDate.toDateString() +
      '\nDeliverable Type: ' +
      deliverableRowData.delivType +
      '\nEst. Pages: ' +
      deliverableRowData.pages +
      '\n\nPM: ' +
      employeeInfo.pmName +
      '\nBackup PM: ' +
      employeeInfo.backupPmName +
      '\nProgram Manager: ' +
      employeeInfo.progManName +
      '\n\nReview Timeline:\nThis is a ' +
      deliverableRowData.tier +
      ' Deliverable\n\nDue to ' +
      employeeInfo.progManName +
      ': ' +
      deliverableRowData.progManDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.editorName +
      ': ' +
      deliverableRowData.editorDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.cooName +
      ': ' +
      deliverableRowData.cLevelDate.toDateString() +
      ', 5pm ET\nDue to ' +
      employeeInfo.ceoName +
      ': N/A\nStatus: ' +
      deliverableRowData.status +
      '\nNotes: ' +
      deliverableRowData.notes +
      "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
      employeeInfo.progManName +
      '.\n- Any timeline changes should be made on the ' +
      employeeInfo.projName +
      ' Tab of the Deliverables Tracker ' +
      sheet.getRange('C6').getValue()
  );
}

function sendTier3UpdateEmail(recipientList, employeeInfo, deliverableRowData) {
  var sheet = ss.getActiveSheet();
  MailApp.sendEmail(
    recipientList,
    'Updated Deliverable Timeline: ' +
      employeeInfo.projName +
      ' ' +
      deliverableRowData.delivName +
      ' Due - ' +
      deliverableRowData.delivDate.toDateString().substring(4, 10),
    'Updates have been made to the following deliverable in the Amida Deliverables Tracker\nProject: ' +
      employeeInfo.projName +
      '\nDeliverable: ' +
      deliverableRowData.delivName +
      '\nDue to Customer: ' +
      deliverableRowData.delivDate.toDateString() +
      '\nDeliverable Type: ' +
      deliverableRowData.delivType +
      '\nEst. Pages: ' +
      deliverableRowData.pages +
      '\n\nPM: ' +
      employeeInfo.pmName +
      '\nBackup PM: ' +
      employeeInfo.backupPmName +
      '\nProgram Manager: ' +
      employeeInfo.progManName +
      '\n\nReview Timeline:\nThis is a ' +
      deliverableRowData.tier +
      ' Deliverable\nDue to ' +
      employeeInfo.progManName +
      ': ' +
      deliverableRowData.progManDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.editorName +
      ': ' +
      deliverableRowData.editorDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.cooName +
      ': N/A\nDue to ' +
      employeeInfo.ceoName +
      ': N/A\nStatus: ' +
      deliverableRowData.status +
      '\nNotes: ' +
      deliverableRowData.notes +
      "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
      employeeInfo.progManName +
      '.\n- Any timeline changes should be made on the ' +
      employeeInfo.projName +
      ' Tab of the Deliverables Tracker ' +
      sheet.getRange('C6').getValue()
  );
}

function deleteEvents(idValues, deliverableRowData, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cal = deliverablesGoogleCalendar;
  var idSheet = ss.getSheetByName('Event IDs');
  var sheet = ss.getActiveSheet();
  for (var i = 0; i < idSheet.getLastRow(); i++) {
    if (idValues[i][0].substring(0, 9) == deliverableRowData.id) {
      try {
        cal.getEventById(idValues[i][1]).deleteEvent(); // Function to time.
      } catch (e) {
        var deleted = deliverableRowData.id;
      }
      idSheet.getRange(i + 1, 1).setValue('EVENT DELETED');
      idSheet.getRange(i + 1, 2).setValue('EVENT DELETED');
    }
  }
  sheet.deleteRow(row + 12);
}

function sendTier1DeleteEmail(recipientList, employeeInfo, deliverableRowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  MailApp.sendEmail(
    recipientList,
    ' Deleted Deliverable: ' +
      employeeInfo.projName +
      ' ' +
      deliverableRowData.delivName +
      ' ' +
      deliverableRowData.delivDate.toDateString().substring(4, 10),
    'The following deliverable has been deleted from the Amida Deliverables Tracker and removed from the calendar. If this deletion was a mistake, you will need to re-enter the data in a new row and add it once again.\n\nProject: ' +
      employeeInfo.projName +
      '\nDeliverable: ' +
      deliverableRowData.delivName +
      '\nDue to Customer: ' +
      deliverableRowData.delivDate.toDateString() +
      '\nDeliverable Type: ' +
      deliverableRowData.delivType +
      '\nEst. Pages: ' +
      deliverableRowData.pages +
      '\n\nPM: ' +
      employeeInfo.pmName +
      '\nBackup PM: ' +
      employeeInfo.backupPmName +
      '\nProgram Manager: ' +
      employeeInfo.progManName +
      '\n\nReview Timeline:\nThis is a ' +
      deliverableRowData.tier +
      ' Deliverable\nDue to ' +
      employeeInfo.progManName +
      ': ' +
      deliverableRowData.progManDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.editorName +
      ': ' +
      deliverableRowData.editorDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.cooName +
      ': ' +
      deliverableRowData.cLevelDate.toDateString() +
      ', 5pm ET\nDue to ' +
      employeeInfo.ceoName +
      ': ' +
      deliverableRowData.ceoDate.toDateString() +
      ', 5pm ET (Notify on ' +
      deliverableRowData.notifyPeterDate.toDateString() +
      ')\nStatus: ' +
      deliverableRowData.status +
      '\nNotes: ' +
      deliverableRowData.notes +
      "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
      employeeInfo.pmName +
      '.\n- Any timeline changes should be made on the ' +
      employeeInfo.projName +
      ' Tab of the Deliverables Tracker ' +
      sheet.getRange('C6').getValue()
  );
}

function sendTier2DeleteEmail(recipientList, employeeInfo, deliverableRowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  MailApp.sendEmail(
    recipientList,
    ' Deleted Deliverable: ' +
      employeeInfo.projName +
      ' ' +
      deliverableRowData.delivName +
      ' ' +
      deliverableRowData.delivDate.toDateString().substring(4, 10),
    'The following deliverable has been deleted from the Amida Deliverables Tracker and removed from the calendar. If this deletion was a mistake, you will need to re-enter the data in a new row and add it once again.\n\nProject: ' +
      employeeInfo.projName +
      '\nDeliverable: ' +
      deliverableRowData.delivName +
      '\nDue to Customer: ' +
      deliverableRowData.delivDate.toDateString() +
      '\nDeliverable Type: ' +
      deliverableRowData.delivType +
      '\nEst. Pages: ' +
      deliverableRowData.pages +
      '\n\nPM: ' +
      employeeInfo.pmName +
      '\nBackup PM: ' +
      employeeInfo.backupPmName +
      '\nProgram Manager: ' +
      employeeInfo.progManName +
      '\n\nReview Timeline:\nThis is a ' +
      deliverableRowData.tier +
      ' Deliverable\nDue to ' +
      employeeInfo.progManName +
      ': ' +
      deliverableRowData.progManDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.editorName +
      ': ' +
      deliverableRowData.editorDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.cooName +
      ': ' +
      deliverableRowData.cLevelDate.toDateString() +
      ', 5pm ET\nDue to ' +
      employeeInfo.ceoName +
      ': N/A\nStatus: ' +
      deliverableRowData.status +
      '\nNotes: ' +
      deliverableRowData.notes +
      "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
      employeeInfo.pmName +
      '.\n- Any timeline changes should be made on the ' +
      employeeInfo.projName +
      ' Tab of the Deliverables Tracker ' +
      sheet.getRange('C6').getValue()
  );
}

function sendTier3DeleteEmail(recipientList, employeeInfo, deliverableRowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  MailApp.sendEmail(
    recipientList,
    ' Deleted Deliverable: ' +
      employeeInfo.projName +
      ' ' +
      deliverableRowData.delivName +
      ' ' +
      deliverableRowData.delivDate.toDateString().substring(4, 10),
    'The following deliverable has been deleted from the Amida Deliverables Tracker and removed from the calendar. If this deletion was a mistake, you will need to re-enter the data in a new row and add it once again.\n\nProject: ' +
      employeeInfo.projName +
      '\nDeliverable: ' +
      deliverableRowData.delivName +
      '\nDue to Customer: ' +
      deliverableRowData.delivDate.toDateString() +
      '\nDeliverable Type: ' +
      deliverableRowData.delivType +
      '\nEst. Pages: ' +
      deliverableRowData.pages +
      '\n\nPM: ' +
      employeeInfo.pmName +
      '\nBackup PM: ' +
      employeeInfo.backupPmName +
      '\nProgram Manager: ' +
      employeeInfo.progManName +
      '\n\nReview Timeline:\nThis is a ' +
      deliverableRowData.tier +
      ' Deliverable\nDue to ' +
      employeeInfo.progManName +
      ': ' +
      deliverableRowData.progManDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.editorName +
      ': ' +
      deliverableRowData.editorDate.toDateString() +
      ', 12pm ET\nDue to ' +
      employeeInfo.cooName +
      ': N/A\nDue to ' +
      employeeInfo.ceoName +
      ': N/A\nStatus: ' +
      deliverableRowData.status +
      '\nNotes: ' +
      deliverableRowData.notes +
      "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
      employeeInfo.pmName +
      '.\n- Any timeline changes should be made on the ' +
      employeeInfo.projName +
      ' Tab of the Deliverables Tracker ' +
      sheet.getRange('C6').getValue()
  );
}

function deleteFromMaster(masterValues, deleted) {
  Logger.log(deleted);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName('Master');
  for (var row = masterValues.length - 1; row >= 0; row--) {
    if (deleted == masterValues[row][27]) {
      masterSheet.deleteRow(row + 3);
      break;
    }
  }
}

function archiveDeliverable(deliverableRowData, row, employeeInfo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var archiveSheet = ss.getSheetByName('Archive');
  var archived = deliverableRowData.id;
  archiveSheet.appendRow([
    employeeInfo.projName,
    employeeInfo.pmName,
    employeeInfo.backUpPmName,
    employeeInfo.progManName,
    deliverableRowData.delivName,
    deliverableRowData.delivDate,
    deliverableRowData.notes,
    deliverableRowData.id
  ]);
  sheet.deleteRow(row + 12);
  return archived;
}

function sureYouWantToArchive(ui, deliverableRowData) {
  if (
    deliverableRowData.status != 'Complete' &&
    deliverableRowData.status != ''
  ) {
    Logger.log(deliverableRowData);
    var response = ui.alert(
      'The deliverable has not been set to complete. Are you sure you want to archive?',
      ui.ButtonSet.YES_NO
    );
    if (response == ui.Button.NO) {
      return 'NO';
    }
    return 'YES';
  }
}

function sureYouWantToDelete(ui) {
  var response = ui.alert(
    'Are you sure you want to delete this deliverable?',
    ui.ButtonSet.YES_NO
  );
  if (response == ui.Button.NO) {
    return 'NO';
  }
  return 'YES';
}
function sendArchiveEmail(recipientList, employeeInfo, deliverableRowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  MailApp.sendEmail(
    recipientList,
    ' Archived Deliverable: ' +
      employeeInfo.projName +
      ' ' +
      deliverableRowData.delivName +
      ' ' +
      deliverableRowData.delivDate.toDateString().substring(4, 10),
    'The following deliverable has been marked as COMPLETE, removed from the ' +
      employeeInfo.projName +
      ' Project sheet and Master sheet, and moved to the Archive sheet.\n\nProject: ' +
      employeeInfo.projName +
      '\nDeliverable: ' +
      deliverableRowData.delivName +
      '\nDue to Customer: ' +
      deliverableRowData.delivDate.toDateString() +
      '\nDeliverable Type: ' +
      deliverableRowData.delivType +
      '\nEst. Pages: ' +
      deliverableRowData.pages +
      '\n\nPM: ' +
      employeeInfo.pmName +
      '\nBackup PM: ' +
      employeeInfo.backupPmName +
      '\nProgram Manager: ' +
      employeeInfo.progManName +
      '\n\nStatus: ' +
      deliverableRowData.status +
      '\nNotes: ' +
      deliverableRowData.notes +
      "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
      employeeInfo.pmName +
      '.\n- Any timeline changes should be made on the ' +
      employeeInfo.projName +
      ' Tab of the Deliverables Tracker ' +
      sheet.getRange('C6').getValue()
  );
}

function checkIfChanged(row) {
  var changedValues = false;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var numRow = ss.getLastRow();
  var savedRange = sheet.getRange(12, 2, numRow);
  var savedValues = savedRange.getValues();
  if (savedValues[row][0] == '✗') {
      Logger.log(savedValues[row][0]);
    changedValues = true;
  }
  return changedValues;
}

function finishFunction(checkRange, oneChecked) {
  if (!oneChecked) {
    Browser.msgBox('No rows were checked. Please try again.');
    return;
  }
  checkRange.setValue(false);
  removeEmptyRows();
  addRow();
}

