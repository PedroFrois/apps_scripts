function copyNextWeekToCurrent(spreadsheet) {
  spreadsheet.getSheetByName('Semana Atual').activate();
  spreadsheet.getRange('C4').activate();
  spreadsheet.getRange('\'Semana Seguinte\'!C4:I36').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
};

function deleteNextWeekEvents(spreadsheet){
  spreadsheet.getSheetByName('Semana Seguinte').activate();
  spreadsheet.getRange('C6:I33').activate();
  spreadsheet.getActiveRangeList().clearContent();
};

function prepareSpreadsheet(spreadsheet) {
  copyNextWeekToCurrent(spreadsheet);
  deleteNextWeekEvents(spreadsheet);
};

function getCalendarIds(){
  var response = Calendar.CalendarList.list().items;
  return response;
};

function getWeekEvents(calendarId){
  var minInMiliSec = 60*1000;
  var dayInMiliSec = 24*60*minInMiliSec;
  
  var now = new Date();
  now.setTime(now.getTime() + dayInMiliSec*7);
  var nextWeek = new Date();
  nextWeek.setTime(now.getTime() + dayInMiliSec*7);
  
  var optionalArgs = {
    timeMin: now.toISOString(),
    timeMax: nextWeek.toISOString(),
    showDeleted: false,
    singleEvents: true
  };
  return Calendar.Events.list(calendarId, optionalArgs).items; 
};

function getColumn(date) {
  switch(date.getDay()){
    case 1://seg
      return 'C';
      break;
    case 2://ter
      return 'D';
      break;
    case 3://qua
      return 'E';
      break;
    case 4://qui
      return 'F';
      break;
    case 5://sex
      return 'G';
      break;
    case 6://sab
      return 'H';
      break;
    case 0://dom
      return 'I';
      break;
  }
};

function getRow(date) {
  var hour = date.getHours() - 6;
  var minute = Math.round(date.getMinutes()/30);
  var row = 4 + minute;
  while(hour > 0){
    row += 2;
    hour--;
  }
  return row;
};

function fillEventsOfTheWeek() {
  var spreadsheet = SpreadsheetApp.getActive();
  var nextWeekSheet = spreadsheet.getSheetByName('Semana Seguinte');
  //prepare spreadsheet
  prepareSpreadsheet(spreadsheet);
  //get all calendar IDs
  var calendarIds = getCalendarIds();
  Logger.log('Calendar IDs: %s', calendarIds);
  if (calendarIds.length > 0) {
    for (i = 0; i < calendarIds.length; i++) {
      var calendarId = calendarIds[i];
      //from each calendar id
      //get events of the week  
      Logger.log('Calendar: %s', calendarId);
      Logger.log('CalendarId: %s', calendarId.id);
      var weekEvents = getWeekEvents(calendarId.id);
      if (weekEvents.length > 0) {
        for (j = 0; j < weekEvents.length; j++) {
          var event = weekEvents[j];
          var eventName = event.summary;
          Logger.log('Event name: ', eventName);
          Logger.log('Event date: %s', event.start.dateTime);
          if(typeof event.start.dateTime === 'undefined'){
            continue;
          }
          var startDate = new Date(event.start.dateTime);
          startDate.setTime(startDate.getTime());
          Logger.log('Date created: %s', startDate.toISOString());
          var endDate = new Date(event.end.dateTime);
          endDate.setTime(endDate.getTime());
          Logger.log('EndDate created: %s', endDate.toISOString());
          
          var column = getColumn(startDate);
          Logger.log('Coluna: %s', column);
          
          var startRow = getRow(startDate);
          Logger.log('Start row: %s', startRow);
          var endRow = getRow(endDate) - 1;
          Logger.log('End row: %s', endRow);
          
          
          var range = String(column) + String(startRow) + ':' + String(column) + String(endRow);
          Logger.log('Range: %s', range);
          

          nextWeekSheet.getRange(range).setValue(eventName);        
        }
      }
    }
  }    
};

function test(){ 
  //var spreadsheet = SpreadsheetApp.getActive();
  //copyNextWeekToCurrent(spreadsheet);
  //deleteNextWeekEvents(spreadsheet);  
};
