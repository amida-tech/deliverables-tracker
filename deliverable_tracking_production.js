// TODO: Add comment explaining what this does
function addRow() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet(),
    numRow = sheet.getLastRow();
  var numCol = sheet.getLastColumn(),
    range = ss.getSheetByName("Row Format").getRange(1, 1, 1, numCol);
  sheet.insertRowsAfter(numRow, 1);
  range.copyTo(sheet.getRange(numRow + 1, 1, 1, numCol), {
    contentsOnly: false
  });
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Sheet Functions");
  var addProject = menu.addItem("Add", "bulkAdd");
  menu.addSeparator();
  var editProject = menu.addItem("Update", "bulkEdit");
  menu.addSeparator();
  var deleteProject = menu.addItem("Delete", "bulkDelete");
  menu.addSeparator();
  var archiveProject = menu.addItem("Archive", "bulkArchive");
  menu.addSeparator();
  var addRows = menu.addItem("New Row", "addRow");
  addProject.addToUi();
  editProject.addToUi();
  deleteProject.addToUi();
  archiveProject.addToUi();
  addRows.addToUi();
}
function removeEmptyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = 0;
  var sheet = ss.getActiveSheet();
  var maxRows = sheet.getMaxRows();
  var range = sheet.getRange(12, 8, sheet.getLastRow());
  var values = range.getValues();
  console.log(typeof values[0][0]);
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0].length) {
      lastRow = i + 11;
      break;
    }
  }
  console.log(lastRow);
  if (maxRows - lastRow != 0) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
}

function bulkAdd() {
  var count = 0;
  var ui = SpreadsheetApp.getUi(); // Spreadsheet UI for buttons
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
    var calName = sheet.getRange("Y2").getValue();
  var cal = CalendarApp.getCalendarById(
    calName
  );
  var listRange = sheet.getRange(2, 23, 10, 33); // Range of employee names/positions
  var listValues = listRange.getValues();
  var editorName = listValues[5][1],
    pmName = listValues[6][1],
    backupPmName = listValues[7][1],
    progManName = listValues[8][1],
    ceoName = listValues[1][1],
    cooName = listValues[2][1];
  var numRow = ss.getLastRow();
  var checkRange = sheet.getRange(12, 1, numRow);
  var range = sheet.getRange(12, 1, numRow, 27); // Range of actual deliverables
  var values = range.getValues();
  var idSheet = ss.getSheetByName("Event IDs");
  var idNumCol = idSheet.getLastColumn();
  var idNumRow = idSheet.getLastRow();
  var idRange = idSheet.getRange(1, 1, 1000, 1000); // Range of Event IDs
  var idValues = idRange.getValues();
  var masterSheet = ss.getSheetByName("Master");
  var masterNumCol = masterSheet.getLastColumn();
  var masterNumRow = masterSheet.getLastRow();
  var masterRange = masterSheet.getRange(1, 1, masterNumRow, masterNumCol);
  var projNameRange = sheet.getRange(8, 3);
  var projName = projNameRange.getValue();
  var concatRange = sheet.getRange(12, 29, numRow);
  var concatValues = concatRange.getValues();
  var oneChecked = false;
  var alreadyAdded = [],
    addSuccess = [],
    notFilled = []; // Arrays to store any projects that are already added or don't have enough info

  for (var row = 0; row < values.length; row++) {
    if (values[row][0]) {
      // Using sleep function to fix Google error of adding too many new events in a certain amount of time
      if (count > 3) {
        Utilities.sleep(15000);
        count = 0;
      }
      oneChecked = true;
      var randomNum = Math.floor(100000000 + Math.random() * 900000000); // Creates 9-Digit Unique ID
      var randRange = sheet.getRange(row + 12, 26); // Range to add Unique ID
      var added = false,
        incomplete = false;
      for (var i = 0; i < idNumRow; i++) {
        if (idValues[i][0].substring(0, 9) == id) {
          alreadyAdded.push(row + 12);
          added = true;
          break;
        }
      }
      if (added) {
        continue;
      }
      for (var i = 0; i < values[row].length; i++) {
        if (values[row][i] == "" && i != 23 && i != 25) {
          notFilled.push(row + 12);
          incomplete = true;
          break;
        }
      }
      if (incomplete) {
        continue;
      }
      var delivName = values[row][2],
        delivType = values[row][3],
        pages = values[row][4],
        tier = values[row][5],
        dbProgMan = values[row][6],
        dayProgMan = values[row][7],
        dateProgMan = values[row][8],
        timeProgMan = values[row][9],
        dbEditor = values[row][10],
        dayEditor = values[row][11],
        dateEditor = values[row][12],
        timeEditor = values[row][13],
        dbThird = values[row][14],
        dayThird = values[row][15],
        dateThird = values[row][16],
        timeThird = values[row][17],
        dbCeo = values[row][18],
        dayCeo = values[row][19],
        dateCeo = values[row][20],
        timeCeo = values[row][21],
        dateCustomer = values[row][22],
        notes = values[row][23],
        status = values[row][24],
        id = values[row][25];
      randRange.setValue(randomNum);
      addSuccess.push(row + 12);
      var progManDate = new Date(values[row][8]);
      var progManDesc =
        "Due to " + progManName + ": " + delivName + progManDate;
      var editorDate = new Date(values[row][12]);
      var editorDesc = "Due to " + editorName + ": " + delivName + editorDate;
      var cLevelDate = new Date(values[row][16]);
      var cLevelDesc = "Due to " + cooName + ": " + delivName + cLevelDate;
      var ceoDate = new Date(values[row][20]);
      var ceoDesc = "Due to " + ceoName + ": " + delivName + ceoDate;
      var delivDate = new Date(values[row][22]);
      var notifyPeterDate = new Date(values[row][26]);
      var progManEvent, editorEvent, cLevelEvent, ceoEvent, delivEvent;
      masterSheet.appendRow([
        projName,
        pmName,
        backupPmName,
        progManName,
        delivName,
        delivType,
        pages,
        tier,
        dbProgMan,
        dayProgMan,
        dateProgMan,
        timeProgMan,
        dbEditor,
        dayEditor,
        dateEditor,
        timeEditor,
        dbThird,
        dayThird,
        dateThird,
        timeThird,
        dbCeo,
        dayCeo,
        dateCeo,
        timeCeo,
        dateCustomer,
        notes,
        status,
        randomNum,
        notifyPeterDate,
        7,
        concatValues[row][0]
      ]);
      //  Browser.msgBox("COUNT: " + count);
      progManEvent = cal.createAllDayEvent(
        "Due to " + progManName + ": " + delivName,
        progManDate
      );
      progManEvent.setDescription(
        "Link to Project/Deliverable Tracker\n" +
          sheet.getRange("C6").getValue()
      );
      count++;
      ss.getSheetByName("Event IDs").appendRow([
        randomNum + " FIRST " + progManDesc,
        progManEvent.getId()
      ]);

      editorEvent = cal.createAllDayEvent(
        "Due to " + editorName + ": " + delivName,
        editorDate
      );
      editorEvent.setDescription(
        "Link to Project/Deliverable Tracker\n" +
          sheet.getRange("C6").getValue()
      );
      count++;
      ss.getSheetByName("Event IDs").appendRow([
        randomNum + " SECOND " + editorDesc,
        editorEvent.getId()
      ]);
      if (tier == "Tier 1") {
        cLevelEvent = cal.createAllDayEvent(
          "Due to " + cooName + ": " + delivName,
          cLevelDate
        );
        cLevelEvent.setDescription(
          "Link to Project/Deliverable Tracker\n" +
            sheet.getRange("C6").getValue()
        );
        count++;

        ss.getSheetByName("Event IDs").appendRow([
          randomNum + " THIRD" + cLevelDesc,
          cLevelEvent.getId()
        ]);
        ceoEvent = cal.createAllDayEvent(
          "Due to " + ceoName + ": " + delivName,
          ceoDate
        );
        ceoEvent.setDescription(
          "Link to Project/Deliverable Tracker\n" +
            sheet.getRange("C6").getValue()
        );
        count++;

        ss.getSheetByName("Event IDs").appendRow([
          randomNum + " FOURTH" + ceoDesc,
          ceoEvent.getId()
        ]);
        delivEvent = cal.createAllDayEvent(
          "Due to Customer: " + delivName,
          new Date(values[row][22])
        );
        delivEvent.setDescription(
          "Link to Project/Deliverable Tracker\n" +
            sheet.getRange("C6").getValue()
        );
        count++;

        ss.getSheetByName("Event IDs").appendRow([
          randomNum + " Due to Customer: ",
          delivEvent.getId()
        ]);
        var recipientList = "";
        for (var i = 0; i < listValues.length; i++) {
          if (listValues[i][3] == "Yes") {
            recipientList = recipientList + listValues[i][2];
            if (i < listValues.length - 1) {
              recipientList = recipientList + ",";
            }
          }
        }
        MailApp.sendEmail(
          recipientList,
          projName +
            " " +
            delivName +
            " Deliverable Timeline Due - " +
            delivDate.toDateString().substring(4, 10),
          "A new deliverable has been added to the Amida Deliverables Tracker\n\nProject: " +
            projName +
            "\nDeliverable: " +
            delivName +
            "\nDue to Customer: " +
            delivDate.toDateString() +
            "\nDeliverable Type: " +
            delivType +
            "\nEst. Pages: " +
            pages +
            "\n\nPM: " +
            listValues[6][1] +
            "\nBackup PM: " +
            listValues[7][1] +
            "\nProgram Manager: " +
            listValues[8][1] +
            "\n\nReview Timeline:\nThis is a " +
            tier +
            " Deliverable\nDue to " +
            editorName +
            ": " +
            editorDate.toDateString() +
            ", 12pm ET\nDue to " +
            cooName +
            ": " +
            cLevelDate.toDateString() +
            ", 5pm ET\nDue to " +
            ceoName +
            ": " +
            ceoDate.toDateString() +
            ", 5pm ET (Notify on " +
            notifyPeterDate.toDateString() +
            ")\nStatus: " +
            status +
            "\nNotes: " +
            notes +
            "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
            listValues[7][2] +
            ".\n- Any timeline changes should be made on the " +
            projName +
            " Tab of the Deliverables Tracker " +
            sheet.getRange("C6").getValue()
        );
      } else if (tier == "Tier 2") {
        cLevelEvent = cal.createAllDayEvent(
          "Due to " + cooName + ": " + delivName,
          cLevelDate
        );
        ss.getSheetByName("Event IDs").appendRow([
          randomNum + " " + cLevelDesc,
          cLevelEvent.getId()
        ]);
        delivEvent = cal.createAllDayEvent(
          "Due to Customer: " + delivName,
          new Date(values[row][22])
        );
        ss.getSheetByName("Event IDs").appendRow([
          randomNum + " Due to Customer: ",
          delivEvent.getId()
        ]);
        var recipientList = "";
        for (var i = 0; i < listValues.length; i++) {
          if (listValues[i][4] == "Yes") {
            recipientList = recipientList + listValues[i][2];
            if (i < listValues.length - 1) {
              recipientList = recipientList + ",";
            }
          }
        }
        MailApp.sendEmail(
          recipientList,
          projName +
            " " +
            delivName +
            " Deliverable Timeline Due - " +
            delivDate.toDateString().substring(4, 10),
          "A new deliverable has been added to the Amida Deliverables Tracker\n\nProject: " +
            projName +
            "\nDeliverable: " +
            delivName +
            "\nDue to Customer: " +
            delivDate.toDateString() +
            "\nDeliverable Type: " +
            delivType +
            "\nEst. Pages: " +
            pages +
            "\n\nPM: " +
            listValues[6][1] +
            "\nBackup PM: " +
            listValues[7][1] +
            "\nProgram Manager: " +
            listValues[8][1] +
            "\n\nReview Timeline:\nThis is a " +
            tier +
            " Deliverable\n\nDue to " +
            editorName +
            ": " +
            editorDate.toDateString() +
            ", 12pm ET\nDue to " +
            cooName +
            ": " +
            cLevelDate.toDateString() +
            ", 5pm ET\nDue to " +
            ceoName +
            ": N/A\nStatus: " +
            status +
            "\nNotes: " +
            notes +
            "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
            listValues[6][1] +
            ".\n- Any timeline changes should be made on the " +
            projName +
            " Tab of the Deliverables Tracker " +
            sheet.getRange("C6").getValue()
        );
      } else if (tier == "Tier 3") {
        delivEvent = cal.createAllDayEvent(
          "Due to Customer: " + delivName,
          new Date(values[row][22])
        );
        console.log("CREATED:" + delivEvent);
        ss.getSheetByName("Event IDs").appendRow([
          randomNum + " Due to Customer: ",
          delivEvent.getId()
        ]);
        var recipientList = "";
        for (var i = 0; i < listValues.length; i++) {
          if (listValues[i][5] == "Yes") {
            recipientList = recipientList + listValues[i][2];
            if (i < listValues.length - 1) {
              recipientList = recipientList + ",";
            }
          }
        }
        MailApp.sendEmail(
          recipientList,
          projName +
            " " +
            delivName +
            " Deliverable Timeline Due - " +
            delivDate.toDateString().substring(4, 10),
          "A new deliverable has been added to the Amida Deliverables Tracker\n\nProject: " +
            projName +
            "\nDeliverable: " +
            delivName +
            "\nDue to Customer: " +
            delivDate.toDateString() +
            "\nDeliverable Type: " +
            delivType +
            "\nEst. Pages: " +
            pages +
            "\n\nPM: " +
            listValues[6][1] +
            "\nBackup PM: " +
            listValues[7][1] +
            "\nProgram Manager: " +
            listValues[8][1] +
            "\n\nReview Timeline:\nThis is a " +
            tier +
            " Deliverable\n\nDue to " +
            editorName +
            ": " +
            editorDate.toDateString() +
            ", 12pm ET\nDue to " +
            cooName +
            ": N/A\nDue to " +
            ceoName +
            ": N/A\nStatus: " +
            status +
            "\nNotes: " +
            notes +
            "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
            listValues[6][1] +
            ".\n- Any timeline changes should be made on the " +
            projName +
            " Tab of the Deliverables Tracker " +
            sheet.getRange("C6").getValue()
        );
      }

      // Adding PMs and Program Manager to Prog Man Event
      progManEvent.addGuest(listValues[6][2]);
      progManEvent.addGuest(listValues[7][2]);
      progManEvent.addGuest(listValues[8][2]);

      for (var i = 1; i < listValues.length; i++) {
        if (listValues[i][7] == "Yes" && (i != 9 || i != 8)) {
          editorEvent.addGuest(listValues[i][2]);
        }
      }

      if (tier == "Tier 1" || tier == "Tier 2") {
        for (var i = 1; i < listValues.length; i++) {
          if (listValues[i][8] == "Yes") {
            cLevelEvent.addGuest(listValues[i][2]);
          }
        }
      }

      if (tier == "Tier 1") {
        for (var i = 1; i < listValues.length; i++) {
          if (listValues[i][9] == "Yes") {
            ceoEvent.addGuest(listValues[i][2]);
          }
        }
      }

      for (var i = 1; i < listValues.length; i++) {
        if (listValues[i][10] == "Yes") {
          console.log("List Val:" + listValues[i][9]);
          console.log("deliv Event:" + delivEvent);
          delivEvent.addGuest(listValues[i][2]);
        }
      }
    }
  }
  if (!oneChecked) {
    Browser.msgBox("No rows were checked. Please try again.");
    return;
  }

  checkRange.setValue(false); // Resets all checkboxes to empty

  if (addSuccess.length == 1) {
    Browser.msgBox(
      "This row was successfully added \n" + JSON.stringify(addSuccess)
    );
  } else if (addSuccess.length > 1) {
    Browser.msgBox(
      "These rows were successfully added \n" + JSON.stringify(addSuccess)
    );
  }

  if (alreadyAdded.length == 1) {
    Browser.msgBox(
      "This row was already added previously\n" + JSON.stringify(alreadyAdded)
    );
  } else if (alreadyAdded.length > 1) {
    Browser.msgBox(
      "These rows were already added previously\n" +
        JSON.stringify(alreadyAdded)
    );
  }

  if (notFilled.length == 1) {
    Browser.msgBox("This row was not complete\n" + JSON.stringify(notFilled));
  } else if (notFilled.length > 1) {
    Browser.msgBox(
      "These rows were not complete\n" + JSON.stringify(notFilled)
    );
  }

  removeEmptyRows(); // Temp fix to google sheets automatically adding rows

  for (var i = 0; i < 3; i++) {
    addRow();
  }
}

function bulkEdit() {
  var ui = SpreadsheetApp.getUi(); // Spreadsheet UI for buttons
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
    var calName = sheet.getRange("Y2").getValue();
  var cal = CalendarApp.getCalendarById(
    calName
  );
  var listRange = sheet.getRange(2, 23, 10, 33); // Range of employee names/positions
  var listValues = listRange.getValues();
  var editorName = listValues[5][1],
    pmName = listValues[6][1],
    backupPmName = listValues[7][1],
    progManName = listValues[8][1],
    ceoName = listValues[1][1],
    cooName = listValues[2][1];
  var numRow = ss.getLastRow();
  var checkRange = sheet.getRange(12, 1, numRow);
  var range = sheet.getRange(12, 1, numRow, 27); // Range of actual deliverables
  var values = range.getValues();
  var idSheet = ss.getSheetByName("Event IDs");
  var idNumCol = idSheet.getLastColumn();
  var idNumRow = idSheet.getLastRow();
  var idRange = idSheet.getRange(1, 1, 1000, 1000); // Range of Event IDs
  var idValues = idRange.getValues();
  var masterSheet = ss.getSheetByName("Master");
  var masterNumCol = masterSheet.getLastColumn();
  var masterNumRow = masterSheet.getLastRow();
  var masterRange = masterSheet.getRange(3, 1, masterNumRow, masterNumCol);
  var masterValues = masterRange.getValues();
  var changed = false;
  var projNameRange = sheet.getRange(8, 3);
  var projName = projNameRange.getValue();
  var concatRange = sheet.getRange(12, 29, numRow);
  var concatValues = concatRange.getValues();
  var oneChecked = false;
  var masterChange = [],
    notFilled = [];
  for (var row = 0; row < values.length; row++) {
    var incomplete = false;
    if (values[row][0]) {
      for (var i = 0; i < values[row].length; i++) {
        if (values[row][i] == "" && i != 23 && i != 25) {
          notFilled.push(row + 12);
          incomplete = true;
          break;
        }
      }
      if (incomplete) {
        continue;
      }
      masterChange.push(values[row][25]);
    }
  }

  for (var row = masterValues.length - 1; row >= 0; row--) {
    if (masterChange.indexOf(masterValues[row][27]) !== -1) {
      masterSheet.deleteRow(row + 3);
    }
  }
  for (var row = 0; row < values.length; row++) {
    var delivName = values[row][2],
      delivType = values[row][3],
      pages = values[row][4],
      tier = values[row][5],
      dbProgMan = values[row][6],
      dayProgMan = values[row][7],
      dateProgMan = values[row][8],
      timeProgMan = values[row][9],
      dbEditor = values[row][10],
      dayEditor = values[row][11],
      dateEditor = values[row][12],
      timeEditor = values[row][13],
      dbThird = values[row][14],
      dayThird = values[row][15],
      dateThird = values[row][16],
      timeThird = values[row][17],
      dbCeo = values[row][18],
      dayCeo = values[row][19],
      dateCeo = values[row][20],
      timeCeo = values[row][21],
      dateCustomer = values[row][22],
      notes = values[row][23],
      status = values[row][24],
      id = values[row][25];
    var progManDate = new Date(values[row][8]);
    var progManDesc = "Due to " + progManName + ": " + delivName + progManDate;
    var editorDate = new Date(values[row][12]);
    var editorDesc = "Due to " + editorName + ": " + delivName + editorDate;
    var cLevelDate = new Date(values[row][16]);
    var cLevelDesc = "Due to " + cooName + ": " + delivName + cLevelDate;
    var ceoDate = new Date(values[row][20]);
    var ceoDesc = "Due to " + ceoName + ": " + delivName + ceoDate;
    var delivDate = new Date(values[row][22]);
    var notifyPeterDate = new Date(values[row][26]);
    var existing = false;
    var changed = false;
    var masterChange = [];
    if (values[row][0]) {
      oneChecked = true;
      for (var i = 0; i < values[row].length; i++) {
        if (values[row][i] == "" && i != 23 && i != 25) {
          notFilled.push(row + 12);
          incomplete = true;
          break;
        }
      }
      if (incomplete) {
        continue;
      }
      for (var i = 0; i < idValues.length; i++) {
        if (idValues[i][0].substring(0, 9) == id) {
          if (idValues[i][0].indexOf("FIRST") !== -1) {
            var tempEvent = cal.getEventById(idValues[i][1]);
            var oldDate = tempEvent.getAllDayStartDate();
            existing = true;
            if (oldDate.getTime() != new Date(values[row][8]).getTime()) {
              changed = true;
              tempEvent.setAllDayDate(new Date(values[row][8]));
              idSheet.getRange(i + 1, 1).setValue("EVENT CHANGED");
              idSheet.getRange(i + 1, 2).setValue("EVENT CHANGED");
              idSheet.appendRow([id + " " + progManDesc, tempEvent.getId()]);
            }
          }
          if (idValues[i][0].indexOf("SECOND") !== -1) {
            var tempEvent = cal.getEventById(idValues[i][1]);
            var oldDate = tempEvent.getAllDayStartDate();
            existing = true;
            if (oldDate.getTime() != new Date(values[row][12]).getTime()) {
              changed = true;
              tempEvent.setAllDayDate(new Date(values[row][12]));
              idSheet.getRange(i + 1, 1).setValue("EVENT CHANGED");
              idSheet.getRange(i + 1, 2).setValue("EVENT CHANGED");
              idSheet.appendRow([id + " " + editorDesc, tempEvent.getId()]);
            }
          }
          if (idValues[i][0].indexOf("THIRD") !== -1) {
            var tempEvent = cal.getEventById(idValues[i][1]);
            var oldDate = tempEvent.getAllDayStartDate();
            if (oldDate.getTime() != new Date(values[row][16]).getTime()) {
              changed = true;
              tempEvent.setAllDayDate(new Date(values[row][16]));
              idSheet.getRange(i + 1, 1).setValue("EVENT CHANGED");
              idSheet.getRange(i + 1, 2).setValue("EVENT CHANGED");
              idSheet.appendRow([id + " " + cLevelDesc, tempEvent.getId()]);
            }
          }
          if (idValues[i][0].indexOf("FOURTH") !== -1) {
            var tempEvent = cal.getEventById(idValues[i][1]);
            var oldDate = tempEvent.getAllDayStartDate();
            if (oldDate.getTime() != new Date(values[row][20]).getTime()) {
              changed = true;
              tempEvent.setAllDayDate(new Date(values[row][20]));
              idSheet.getRange(i + 1, 1).setValue("EVENT CHANGED");
              idSheet.getRange(i + 1, 2).setValue("EVENT CHANGED");
              idSheet.appendRow([id + " " + ceoDesc, tempEvent.getId()]);
            }
          }
          if (idValues[i][0].indexOf("Customer") !== -1) {
            var tempEvent = cal.getEventById(idValues[i][1]);
            var oldDate = tempEvent.getAllDayStartDate();
            if (oldDate.getTime() != new Date(values[row][22]).getTime()) {
              changed = true;
              tempEvent.setAllDayDate(new Date(values[row][22]));
              idSheet.getRange(i + 1, 1).setValue("EVENT CHANGED");
              idSheet.getRange(i + 1, 2).setValue("EVENT CHANGED");
              idSheet.appendRow([
                id + " " + "Due to Customer: " + delivDate,
                tempEvent.getId()
              ]);
            }
          }
        }
      }
      if (changed) {
        if (tier == "Tier 1") {
          var recipientList = "";
          for (var i = 0; i < listValues.length; i++) {
            if (listValues[i][3] == "Yes") {
              recipientList = recipientList + listValues[i][2];
              if (i < listValues.length - 1) {
                recipientList = recipientList + ",";
              }
            }
          }
          MailApp.sendEmail(
            recipientList,
            "Updated Deliverable Timeline: " +
              projName +
              " " +
              delivName +
              " Due - " +
              delivDate.toDateString().substring(4, 10),
            "Updates have been made to the following deliverable in the Amida Deliverables Tracker\nProject: " +
              projName +
              "\nDeliverable: " +
              delivName +
              "\nDue to Customer: " +
              delivDate.toDateString() +
              "\nDeliverable Type: " +
              delivType +
              "\nEst. Pages: " +
              pages +
              "\n\nPM: " +
              listValues[6][1] +
              "\nBackup PM: " +
              listValues[7][1] +
              "\nProgram Manager: " +
              listValues[8][1] +
              "\n\nReview Timeline:\nThis is a " +
              tier +
              " Deliverable\nDue to " +
              editorName +
              ": " +
              editorDate.toDateString() +
              ", 12pm ET\nDue to " +
              cooName +
              ": " +
              cLevelDate.toDateString() +
              ", 5pm ET\nDue to " +
              ceoName +
              ": " +
              ceoDate.toDateString() +
              ", 5pm ET (Notify on " +
              notifyPeterDate.toDateString() +
              ")\nStatus: " +
              status +
              "\nNotes: " +
              notes +
              "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
              listValues[6][1] +
              ".\n- Any timeline changes should be made on the " +
              projName +
              " Tab of the Deliverables Tracker " +
              sheet.getRange("C6").getValue()
          );
        }
        if (tier == "Tier 2") {
          var recipientList = "";
          for (var i = 0; i < listValues.length; i++) {
            if (listValues[i][4] == "Yes") {
              Logger.log("Yeah Boi");
              recipientList = recipientList + listValues[i][2];
              if (i < listValues.length - 1) {
                recipientList = recipientList + ",";
              }
            }
          }
          MailApp.sendEmail(
            recipientList,
            "Updated Deliverable Timeline: " +
              projName +
              " " +
              delivName +
              " Due - " +
              delivDate.toDateString().substring(4, 10),
            "Updates have been made to the following deliverable in the Amida Deliverables Tracker\nProject: " +
              projName +
              "\nDeliverable: " +
              delivName +
              "\nDue to Customer: " +
              delivDate.toDateString() +
              "\nDeliverable Type: " +
              delivType +
              "\nEst. Pages: " +
              pages +
              "\n\nPM: " +
              listValues[6][1] +
              "\nBackup PM: " +
              listValues[7][1] +
              "\nProgram Manager: " +
              listValues[8][1] +
              "\n\nReview Timeline:\nThis is a " +
              tier +
              " Deliverable\n\nDue to " +
              editorName +
              ": " +
              editorDate.toDateString() +
              ", 12pm ET\nDue to " +
              cooName +
              ": " +
              cLevelDate.toDateString() +
              ", 5pm ET\nDue to " +
              ceoName +
              ": N/A\nStatus: " +
              status +
              "\nNotes: " +
              notes +
              "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
              listValues[6][1] +
              ".\n- Any timeline changes should be made on the " +
              projName +
              " Tab of the Deliverables Tracker " +
              sheet.getRange("C6").getValue()
          );
        }
        if (tier == "Tier 3") {
          var recipientList = "";
          for (var i = 0; i < listValues.length; i++) {
            if (listValues[i][5] == "Yes") {
              recipientList = recipientList + listValues[i][2];
              if (i < listValues.length - 1) {
                recipientList = recipientList + ",";
              }
            }
          }
          MailApp.sendEmail(
            recipientList,
            "Updated Deliverable Timeline: " +
              projName +
              " " +
              delivName +
              " Due - " +
              delivDate.toDateString().substring(4, 10),
            "Updates have been made to the following deliverable in the Amida Deliverables Tracker\nProject: " +
              projName +
              "\nDeliverable: " +
              delivName +
              "\nDue to Customer: " +
              delivDate.toDateString() +
              "\nDeliverable Type: " +
              delivType +
              "\nEst. Pages: " +
              pages +
              "\n\nPM: " +
              listValues[6][1] +
              "\nBackup PM: " +
              listValues[7][1] +
              "\nProgram Manager: " +
              listValues[8][1] +
              "\n\nReview Timeline:\nThis is a " +
              tier +
              " Deliverable\n\nDue to " +
              editorName +
              ": " +
              editorDate.toDateString() +
              ", 12pm ET\nDue to " +
              cooName +
              ": N/A\nDue to " +
              ceoName +
              ": N/A\nStatus: " +
              status +
              "\nNotes: " +
              notes +
              "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
              listValues[6][1] +
              ".\n- Any timeline changes should be made on the " +
              projName +
              " Tab of the Deliverables Tracker " +
              sheet.getRange("C6").getValue()
          );
        }
      }
      masterSheet.appendRow([
        projName,
        pmName,
        backupPmName,
        progManName,
        delivName,
        delivType,
        pages,
        tier,
        dbProgMan,
        dayProgMan,
        dateProgMan,
        timeProgMan,
        dbEditor,
        dayEditor,
        dateEditor,
        timeEditor,
        dbThird,
        dayThird,
        dateThird,
        timeThird,
        dbCeo,
        dayCeo,
        dateCeo,
        timeCeo,
        dateCustomer,
        notes,
        status,
        id,
        notifyPeterDate,
        7,
        concatValues[row][0]
      ]);
    }
  }
  if (!oneChecked) {
    Browser.msgBox("No rows were checked. Please try again.");
    return;
  }
  checkRange.setValue(false);
  removeEmptyRows();
  for (var i = 0; i < 5; i++) {
    addRow();
  } // Resets all checkboxes to empty
}
function bulkDelete() {

  var ui = SpreadsheetApp.getUi(); // Spreadsheet UI for buttons
  var response = ui.alert(
    "Are you sure you want to delete?",
    ui.ButtonSet.YES_NO
  );
  if (response == ui.Button.NO) {
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
    var calName = sheet.getRange("Y2").getValue();
  var cal = CalendarApp.getCalendarById(
    calName
  );
  var listRange = sheet.getRange(2, 23, 10, 33); // Range of employee names/positions
  var listValues = listRange.getValues();
  var editorName = listValues[5][1],
    progManName = listValues[8][1],
    ceoName = listValues[1][1],
    cooName = listValues[2][1];
  var numRow = ss.getLastRow();
  var checkRange = sheet.getRange(12, 1, numRow);
  var range = sheet.getRange(12, 1, numRow, 27); // Range of actual deliverables
  var values = range.getValues();
  var idSheet = ss.getSheetByName("Event IDs");
  var idNumCol = idSheet.getLastColumn();
  var idNumRow = idSheet.getLastRow();
  var idRange = idSheet.getRange(1, 1, 1000, 1000); // Range of Event IDs
  var idValues = idRange.getValues();
  var masterSheet = ss.getSheetByName("Master");
  var masterNumCol = masterSheet.getLastColumn();
  var masterNumRow = masterSheet.getLastRow();
  var masterRange = masterSheet.getRange(4, 1, masterNumRow, masterNumCol);
  var masterValues = masterRange.getValues();
  var projNameRange = sheet.getRange(8, 3);
  var projName = projNameRange.getValue();
  var changed = false;
  var deleted = [];
  for (var row = values.length - 1; row >= 0; row--) {
    var delivName = values[row][2],
      delivType = values[row][3],
      pages = values[row][4],
      tier = values[row][5],
      notes = values[row][23],
      status = values[row][24],
      id = values[row][25];
    var progManDate = new Date(values[row][8]);
    var progManDesc = "Due to " + progManName + ": " + delivName + progManDate;
    var editorDate = new Date(values[row][12]);
    var editorDesc = "Due to " + editorName + ": " + delivName + editorDate;
    var cLevelDate = new Date(values[row][16]);
    var cLevelDesc = "Due to " + cooName + ": " + delivName + cLevelDate;
    var ceoDate = new Date(values[row][20]);
    var ceoDesc = "Due to " + ceoName + ": " + delivName + ceoDate;
    var delivDate = new Date(values[row][22]);
    var notifyPeterDate = new Date(values[row][26]);
    var oneChecked = false;
    if (values[row][0]) {
      oneChecked = true;
      for (var i = 0; i < idNumRow; i++) {
        if (idValues[i][0].substring(0, 9) == id) {
          try {
            cal.getEventById(idValues[i][1]).deleteEvent(); // Function to time.
          } catch (e) {
            deleted.push(id);
          }
          idSheet.getRange(i + 1, 1).setValue("EVENT DELETED");
          idSheet.getRange(i + 1, 2).setValue("EVENT DELETED");
        }
      }
      sheet.deleteRow(row + 12);
      if (tier == "Tier 1") {
        var recipientList = "";
        for (var i = 0; i < listValues.length; i++) {
          if (listValues[i][3] == "Yes") {
            recipientList = recipientList + listValues[i][2];
            if (i < listValues.length - 1) {
              recipientList = recipientList + ",";
            }
          }
        }
        MailApp.sendEmail(
          recipientList,
          " Deleted Deliverable: " +
            projName +
            " " +
            delivName +
            " " +
            delivDate.toDateString().substring(4, 10),
          "The following deliverable has been deleted from the Amida Deliverables Tracker and removed from the calendar. If this deletion was a mistake, you will need to re-enter the data in a new row and add it once again.\n\nProject: " +
            projName +
            "\nDeliverable: " +
            delivName +
            "\nDue to Customer: " +
            delivDate.toDateString() +
            "\nDeliverable Type: " +
            delivType +
            "\nEst. Pages: " +
            pages +
            "\n\nPM: " +
            listValues[6][1] +
            "\nBackup PM: " +
            listValues[7][1] +
            "\nProgram Manager: " +
            listValues[8][1] +
            "\n\nReview Timeline:\nThis is a " +
            tier +
            " Deliverable\nDue to " +
            editorName +
            ": " +
            editorDate.toDateString() +
            ", 12pm ET\nDue to " +
            cooName +
            ": " +
            cLevelDate.toDateString() +
            ", 5pm ET\nDue to " +
            ceoName +
            ": " +
            ceoDate.toDateString() +
            ", 5pm ET (Notify on " +
            notifyPeterDate.toDateString() +
            ")\nStatus: " +
            status +
            "\nNotes: " +
            notes +
            "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
            listValues[7][2] +
            ".\n- Any timeline changes should be made on the " +
            projName +
            " Tab of the Deliverables Tracker " +
            sheet.getRange("C6").getValue()
        );
      } else if (tier == "Tier 2") {
        var recipientList = "";
        for (var i = 0; i < listValues.length; i++) {
          if (listValues[i][4] == "Yes") {
            recipientList = recipientList + listValues[i][2];
            if (i < listValues.length - 1) {
              recipientList = recipientList + ",";
            }
          }
        }
        MailApp.sendEmail(
          recipientList,
          " Deleted Deliverable: " +
            projName +
            " " +
            delivName +
            " " +
            delivDate.toDateString().substring(4, 9),
          "The following deliverable has been deleted from the Amida Deliverables Tracker and removed from the calendar. If this deletion was a mistake, you will need to re-enter the data in a new row and add it once again.\n\nProject: " +
            projName +
            "\nDeliverable: " +
            delivName +
            "\nDue to Customer: " +
            delivDate.toDateString() +
            "\nDeliverable Type: " +
            delivType +
            "\nEst. Pages: " +
            pages +
            "\n\nPM: " +
            listValues[6][1] +
            "\nBackup PM: " +
            listValues[7][1] +
            "\nProgram Manager: " +
            listValues[8][1] +
            "\n\nReview Timeline:\nThis is a " +
            tier +
            " Deliverable\n\nDue to " +
            editorName +
            ": " +
            editorDate.toDateString() +
            ", 12pm ET\nDue to " +
            cooName +
            ": " +
            cLevelDate.toDateString() +
            ", 5pm ET\nDue to " +
            ceoName +
            ": N/A\nStatus: " +
            status +
            "\nNotes: " +
            notes +
            "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
            listValues[6][1] +
            ".\n- Any timeline changes should be made on the " +
            projName +
            " Tab of the Deliverables Tracker " +
            sheet.getRange("C6").getValue()
        );
      } else if (tier == "Tier 3") {
        var recipientList = "";
        for (var i = 0; i < listValues.length; i++) {
          if (listValues[i][5] == "Yes") {
            recipientList = recipientList + listValues[i][2];
            if (i < listValues.length - 1) {
              recipientList = recipientList + ",";
            }
          }
        }
        MailApp.sendEmail(
          recipientList,
          " Deleted Deliverable: " +
            projName +
            " " +
            delivName +
            " " +
            delivDate.toDateString().substring(4, 9),
          "The following deliverable has been deleted from the Amida Deliverables Tracker and removed from the calendar. If this deletion was a mistake, you will need to re-enter the data in a new row and add it once again.\n\nProject: " +
            projName +
            "\nDeliverable: " +
            delivName +
            "\nDue to Customer: " +
            delivDate.toDateString() +
            "\nDeliverable Type: " +
            delivType +
            "\nEst. Pages: " +
            pages +
            "\n\nPM: " +
            listValues[6][1] +
            "\nBackup PM: " +
            listValues[7][1] +
            "\nProgram Manager: " +
            listValues[8][1] +
            "\n\nReview Timeline:\nThis is a " +
            tier +
            " Deliverable\n\nDue to " +
            editorName +
            ": " +
            editorDate.toDateString() +
            ", 12pm ET\nDue to " +
            cooName +
            ": N/A\nDue to " +
            ceoName +
            ": N/A\nStatus: " +
            status +
            "\nNotes: " +
            notes +
            "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
            listValues[6][1] +
            ".\n- Any timeline changes should be made on the " +
            projName +
            " Tab of the Deliverables Tracker " +
            sheet.getRange("C6").getValue()
        );
      }

      // addRow();
    }
    // removeEmptyRows(numRow);
  }

  if (!oneChecked) {
    Browser.msgBox("No rows were checked. Please try again.");
    return;
  }
  // Browser.msgBox(JSON.stringify(deleted));
  // Browser.msgBox(deleted[0]);
  // Browser.msgBox(masterValues[0][13]);
  for (var row = masterValues.length - 1; row >= 0; row--) {
    if (deleted.indexOf(masterValues[row][27]) !== -1) {
      masterSheet.deleteRow(row + 4);
    }
  }

  checkRange.setValue(false);
  removeEmptyRows();
  for (var i = 0; i < 3; i++) {
    addRow();
  } // Resets all checkboxes to empty
}
function bulkArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var archiveSheet = ss.getSheetByName("Archive");
  var sheet = ss.getActiveSheet();
  var listRange = sheet.getRange(2, 23, 10, 33); // Range of employee names/positions
  var listValues = listRange.getValues();
  var editorName = listValues[5][1],
    progManName = listValues[8][1],
    ceoName = listValues[1][1],
    cooName = listValues[2][1];
  var numRow = ss.getLastRow();
  var checkRange = sheet.getRange(12, 1, numRow);
  var range = sheet.getRange(12, 1, numRow, 27); // Range of actual deliverables
  var values = range.getValues();
  var idSheet = ss.getSheetByName("Event IDs");
  var idNumCol = idSheet.getLastColumn();
  var idNumRow = idSheet.getLastRow();
  var idRange = idSheet.getRange(1, 1, 1000, 1000); // Range of Event IDs
  var idValues = idRange.getValues();
  var masterSheet = ss.getSheetByName("Master");
  var masterNumCol = masterSheet.getLastColumn();
  var masterNumRow = masterSheet.getLastRow();
  var masterRange = masterSheet.getRange(4, 1, masterNumRow, masterNumCol);
  var masterValues = masterRange.getValues();
  var projNameRange = sheet.getRange(8, 3);
  var projName = projNameRange.getValue();
  var oneChecked = false;
  var archived = [];
  for (var row = values.length - 1; row >= 0; row--) {
    var progManDate = new Date(values[row][8]);
    var progManDesc = "Due to " + progManName + ": " + delivName + progManDate;
    var editorDate = new Date(values[row][12]);
    var editorDesc = "Due to " + editorName + ": " + delivName + editorDate;
    var cLevelDate = new Date(values[row][16]);
    var cLevelDesc = "Due to " + cooName + ": " + delivName + cLevelDate;
    var ceoDate = new Date(values[row][20]);
    var ceoDesc = "Due to " + ceoName + ": " + delivName + ceoDate;
    var delivDate = new Date(values[row][22]);
    var notifyPeterDate = new Date(values[row][26]);
    var delivName = values[row][2],
      delivType = values[row][3],
      pages = values[row][4],
      tier = values[row][5],
      notes = values[row][23],
      status = values[row][24],
      id = values[row][25];
    if (values[row][0]) {
      Logger.log("CHECKED");
      archived.push(id);
      oneChecked = true;
      archiveSheet.appendRow([
        projName,
        listValues[6][1],
        listValues[7][1],
        progManName,
        delivName,
        delivDate,
        notes,
        id
      ]);
      sheet.deleteRow(row + 12);
      var recipientList = "";
      for (var i = 0; i < listValues.length; i++) {
        if (listValues[i][6] == "Yes") {
          recipientList = recipientList + listValues[i][2];
          if (i < listValues.length - 1) {
            recipientList = recipientList + ",";
          }
        }
      }
      MailApp.sendEmail(
        recipientList,
        " Archived Deliverable: " +
          projName +
          " " +
          delivName +
          " " +
          delivDate.toDateString().substring(4, 9),
        "The following deliverable has been marked as COMPLETE, removed from the " +
          projName +
          " Project sheet and Master sheet, and moved to the Archive sheet.\n\nProject: " +
          projName +
          "\nDeliverable: " +
          delivName +
          "\nDue to Customer: " +
          delivDate.toDateString() +
          "\nDeliverable Type: " +
          delivType +
          "\nEst. Pages: " +
          pages +
          "\n\nPM: " +
          listValues[6][1] +
          "\nBackup PM: " +
          listValues[7][1] +
          "\nProgram Manager: " +
          listValues[8][1] +
          "\n\nStatus: " +
          status +
          "\nNotes: " +
          notes +
          "\n\n===Reminders===\n- Always make sure that the PM and Backup PM are cc'd on all deliverables-related correspondence, and have access to the latest versions of all deliverable drafts.\n- If you have any questions, or need to request changes to the review timeline, please contact " +
          listValues[7][2] +
          ".\n- Any timeline changes should be made on the " +
          projName +
          " Tab of the Deliverables Tracker " +
          sheet.getRange("C6").getValue()
      );
    }
  }
  if (!oneChecked) {
    Browser.msgBox("No rows were checked. Please try again.");
    return;
  }

  for (var row = masterValues.length - 1; row >= 0; row--) {
    if (archived.indexOf(masterValues[row][27]) !== -1) {
      masterSheet.deleteRow(row + 4);
    }
  }

  for (var i = 0; i < 3; i++) {
    addRow();
  }

  checkRange.setValue(false);
  removeEmptyRows();

  Browser.msgBox("Successfully Archived!");
}

