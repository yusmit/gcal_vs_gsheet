var calendarId = "YOURID@group.calendar.google.com";
var sheet = SpreadsheetApp.getActiveSheet();
var calendar = CalendarApp.getCalendarById(calendarId);

function onOpen() {
  "use strict";
  var menuEntries = [
  {
    name: "importFromCalendar",
    functionName: 'importFromCalendar'
  },
  {
    name: "exportToCalendar",
    functionName: 'exportToCalendar'
  },
  {
    name: "removeEvents",
    functionName: 'removeEvents'
  },  
  {
    name: "Clear",
    functionName: 'clearCalendar'
  },
  ], activeSheet;
  
  activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  activeSheet.addMenu('Macros', menuEntries);
}

function clearCalendar() { 
  // Set filters
  var startDate = new Date("01/01/2022 00:00:00");
  var endDate = new Date("12/31/2022 23:59:59");
  Logger.log(endDate);
  var searchText = "";
  var events = calendar.getEvents(startDate, endDate, {search: searchText});
  Logger.log(events);
  // var events = (searchText == '') ? calendar.getEvents(startDate, endDate) : calendar.getEvents(startDate, endDate, {search: searchText});
  // Display events 
  for (var i=0; i<events.length; i++) {
    var ev = events[i];
    ev.deleteEvent();
    Logger.log(ev);
  }
}
function removeEvents() {
  const rowStart = 2;
  const colStart = 1;
  var numRows = sheet.getLastRow();
  var numColumns = sheet.getLastColumn();
  //Извлекаем данные таблицы в указанных диапазонах
  var range = sheet.getRange(rowStart, colStart, numRows, numColumns)
  var events = sheet.getRange(rowStart, colStart, numRows, numColumns).getValues();
  for (var x=0; x<events.length; x++) {
  var evt = events[x];
    if (evt[5] != "свободно"){
      Logger.log(evt[4]+"  "+evt[5] + "events removed");
      calendar.getEventById(evt[4]).deleteEvent();
    }
  }
}


function exportToCalendar() {
  //Индексы первой строки и первого столбца в таблице с данными
  const rowStart = 2;
  const colStart = 1;
  var numRows = sheet.getLastRow();
  var numColumns = sheet.getLastColumn();
  //Извлекаем данные таблицы в указанных диапазонах
  var range = sheet.getRange(rowStart, colStart, numRows, numColumns)
  var events = sheet.getRange(rowStart, colStart, numRows, numColumns).getValues();
  var check_data = events[0];
  for (var x=0; x<events.length; x++) {
    var evt = events[x];
    var startTime = new Date(evt[1]);
    var endTime = new Date(evt[2]);
    var title = evt[0];
    var descr = evt[5];
    calendar.createEvent(title, startTime, endTime, {description: descr});
  }
}

function importFromCalendar() { 
  //var calendarId = sheet.getRange('B1').getValue().toString(); 
  //Индексы первой строки и первого столбца в таблице с данными
  const rowStart = 1;
  const colStart = 1;
  var startDate = new Date();
  // endDate = startDate + 2 month
  var endDate = new Date();
  endDate.setMonth(startDate.getMonth()+2);
  var searchText = '';
  // Print header
  var header = [["Заголовок", "Дата и время начала", "Время окончания", "День недели","Id", "Описание"]];
  len_header = header[0].length;
  var range = sheet.getRange(rowStart, colStart, 1, len_header);
  range.setValues(header);
  range.setFontWeight("bold")
  // Clear worksheet
  var numRows = sheet.getLastRow();
  var numColumns = sheet.getLastColumn();
  var range_to_clr = sheet.getRange(rowStart+1, colStart, numRows, len_header);
  range_to_clr.clear();
  // Get events based on filters
  var events = calendar.getEvents(startDate, endDate, {search: searchText});
  var check_events = [];
  // Display events 
  for (var i=0; i<events.length; i++) {
    var row = i+rowStart+1;
    var details = [[events[i].getTitle(), events[i].getStartTime(), events[i].getEndTime(), events[i].getStartTime().getDay(), events[i].getId(), events[i].getDescription()]];
    len_details=details[0].length;
    range = sheet.getRange(row,1,1,len_details);
    range.setValues(details);
  }
  // Set the data validation for cell H2 to require a value from startTime, with a dropdown menu.
  // var cell = sheet.getRange(rowStart+1, 2, events.length, 1);
  //   cell.setNumberFormat('dd.mm hh:mm');
  //   cell = sheet.getRange(rowStart+1, 3, events.length, 1);
  //   cell.setNumberFormat('hh:mm');
    cell = sheet.getRange(rowStart+1, 4, events.length, 1);
    cell.setNumberFormat('DDD');

  // var cell = sheet.getRange('H2:H6');
  // var range = sheet.getRange(rowStart+1, 2, events.length, 1);
  // var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  // cell.setDataValidation(rule);

}
