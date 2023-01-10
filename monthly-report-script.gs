// Displays buttons on opening file
function onOpen() {
  showButtons();
}

// If Month cell is edited, option to change the year is presented
function onEdit(e){
  if(e.range.getA1Notation() === 'A5'){
      changeYear();
  } 
}

// Functionality for using button to switch between Monthly Report view and Instructions view
function switchToFromInstructions() {
  var sheet = SpreadsheetApp.getActiveSheet().getName();
  var moveTo;
  if(sheet == "Monthly Report"){
    moveTo = SpreadsheetApp.getActive().getSheetByName("Instructions");
  } else {
    moveTo = SpreadsheetApp.getActive().getSheetByName("Monthly Report");
  }
  moveTo.activate();
}

// Focus on cell A1 (Ensuring update suceeds in any case)
function updateFocus(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Monthly Report')
  sheet.setActiveRange(sheet.getRange(1, 1));
}

// Calculate public holidays and work out what days the ALT is likely to have been at school based on days of the week
function updateReport(){
  updateFocus();
  hideRows(false);
  japaneseHolidays();
  updateDaysAtSchool();
  hideRows(true);
}

// Show the buttons for updating, hiding buttons and switching to instructions view
function showButtons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Monthly Report')
  const drawings = sheet.getDrawings();
  var row = 1
  var col = 14
  // Drawing 0 = Update button
  // Drawing 1 = Hide button
  // Drawing 2 = Instructions button
  for(var i = 0 ; i<drawings.length;i++){
    const button = drawings[i];
    // Changing parameters for each buttom
    if(i==1){
      row = 5;
    } else if(i==2){
      row = 1;
      col = 17;
    }
    button.setPosition(row, col, 0, 0);
    button.setHeight(80);
    button.setWidth(250);
  }
}

// Hiding the buttons (as one pixel on the monthly report - try and find it! I'll give you a cookie if you can!)
function hideButtons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Monthly Report');

  const drawings = sheet.getDrawings();
  for(var i = 0 ; i<drawings.length;i++){
    const button = drawings[i];
    button.setHeight(1);
    button.setWidth(1);
    button.setPosition(48, 1, 0, 0);
  }
}

// Subscribes user to the Japanese Holidays calendar and checks the events in case of a public holiday.
// This function fills in the public holidays automatically on the Monthly Report view
function japaneseHolidays() {   
  var mrSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Monthly Report");
  var instructionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Instructions");
  var daysAtSchoolSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Days at School");
  var days = daysAtSchoolSheet.getRange("A1:A31").getValues();
  var pasteToCells = daysAtSchoolSheet.getRange("E1:E31");

  var publicHolidays = [];

  var year = instructionsSheet.getRange("C2").getValue();
  var month = mrSheet.getRange("A5").getValue();

  // The id of the publically accessible Japanese Holidays calendar to subscribe the user to
  var id = "ja.japanese#holiday@group.v.calendar.google.com";

  CalendarApp.getCalendarById("")

  var cal = CalendarApp.getCalendarById(id);
  if(cal == null){
    CalendarApp.subscribeToCalendar(id)
  }
  cal = CalendarApp.getCalendarById(id); 

  // For all days in the month, if a public holiday exists in the publically accessible Japanese Holidays calendar,
  // record the name of the holiday
  for(var i=0; i < days.length; i++){
    var date = new Date(year,month-1,days[i]);
    Logger.log(date);
    var holidays = cal.getEventsForDay(date);
    if(holidays.length>0){
      // is a public holiday
      var holidayName = holidays[0].getTitle();
      publicHolidays.push([holidayName]);
      Logger.log([holidays[0].getTitle()]);
    } else {
      publicHolidays.push([""]);
    }
  }
  // Set the values in the monthly report sheet
  pasteToCells.setValues(publicHolidays);
}

// Calculates days the ALT is likely to be at school and copies them to the Monthly Report sheet
function updateDaysAtSchool() {
  var mrSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Monthly Report");
  var daysAtSchoolSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Days at School");

  // copy public holidays
  var cellsToCopyHolidays = daysAtSchoolSheet.getRange("E1:E31").getValues();
  var pasteToCellsHolidays = mrSheet.getRange("L15:L45");
  pasteToCellsHolidays.setValues(cellsToCopyHolidays);

  // copy school name
  var cellsToCopy = daysAtSchoolSheet.getRange("B1:C31").getValues();
  var pasteToCells = mrSheet.getRange("B15:C45");
  pasteToCells.setValues(cellsToCopy);
}

// Change the year of the monthly report programatically (in Cell C2 in Instructions sheet)
function changeYear() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Instructions");
  var year = sheet.getRange("C2");


  var result = ui.alert(
      'Change year?',
      'The current year set is ' + year.getValue() + ".\n Do you want to change it?",
      ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
      var yearSelect = ui.prompt(
        'Change year',
        'Insert the year',
        ui.ButtonSet.OK_CANCEL);

      var okCancel = yearSelect.getSelectedButton();
      var text = parseInt(yearSelect.getResponseText());

      if(okCancel == ui.Button.OK){
        // User clicked "YES".
        if(text <= 2100 && text >= 2021){
          year.setValue(text);
          var result = ui.alert(
            'Change year',
            'The year has been changed to ' + year.getValue() + ".",
            ui.ButtonSet.OK);
        } else {
          ui.alert("Invalid Year.")
          changeYear();
        }
      }
  } 
}

// Hides 29th, 30th and 31st rows for the relevant months in Monthly Report view.
function hideRows(hide){
  var mrSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Monthly Report");
  var daysAtSchoolSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Days at School");
  var twentyNinthRow = 43; // rows for the 29th, 30th and 31st
  var thirtiethRow = 44;
  var thirtyFirstRow = 45;

  var twentyNinthCheck = parseInt(daysAtSchoolSheet.getRange("D29").getValue());
  var thirtiethCheck = parseInt(daysAtSchoolSheet.getRange("D30").getValue());
  var thirtyFirstCheck = parseInt(daysAtSchoolSheet.getRange("D31").getValue());

  var twentyNineCells = mrSheet.getRange("D43:G43");
  var thirtyCells = mrSheet.getRange("D44:G44");
  var thirtyOneCells = mrSheet.getRange("D45:G45");

  if(hide) {
    if(twentyNinthCheck != 29){
      mrSheet.hideRows(twentyNinthRow);
      twentyNineCells.setValue(0);
    }
    if(thirtiethCheck != 30){
      mrSheet.hideRows(thirtiethRow);
      thirtyCells.setValue(0);
    }
    if(thirtyFirstCheck != 31){
      mrSheet.hideRows(thirtyFirstRow);
      thirtyOneCells.setValue(0);
    }
  } else {
    mrSheet.unhideRow(mrSheet.getRange("A" + twentyNinthRow.toString()));
    mrSheet.unhideRow(mrSheet.getRange("A" + thirtiethRow.toString()));
    mrSheet.unhideRow(mrSheet.getRange("A" + thirtyFirstRow.toString()));
    twentyNineCells.setValue("");
    thirtyCells.setValue("");
    thirtyOneCells.setValue("");
  }

}
