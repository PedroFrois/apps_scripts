//global
var ROWS_TABLE = 34;
var TABLE_STARTING_HOUR = 6;
var SHEET_NAMES = ["Last Week","Current Week","Next Week"];

//add menu
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    //{name: 'Prepare sheet...', functionName: 'prepareSheet_'},
    {name: 'Get events...', functionName: 'main'},
    {name: 'Preencher com x...', functionName: 'xFill'}
  ];
  spreadsheet.addMenu('Scripts', menuItems);
}

//util
function getColumn(d){
  return Math.abs(d.getDay() == 0 ? 6:(d.getDay()-1));
}

function getRow(d){
  var rowHour = d.getHours() - TABLE_STARTING_HOUR;
  var rowMinute = Math.round(d.getMinutes()/30);
  var row = rowHour*2 + rowMinute + 1;//header
  return row;
}

function getMonday(d) {
  d = new Date(d);
  var day = d.getDay(),
      diff = d.getDate() - day + (day == 0 ? -6:1); // adjust when day is sunday
  d = new Date(d.setDate(diff));
  d.setHours(0);
  d.setMinutes(0);
  return d;
}

function getFinalDayOfWeek(d){
  d = new Date(d);
  var diff = d.getDate() + 7; // adjust when day is sunday
  d = new Date(d.setDate(diff));
  d.setHours(0);
  d.setMinutes(0);
  return d;
}

//clean table (empty spaces between 2 filled spaces get filled with x)
function cleanTable(rawTable){
  for(j = 0; j < rawTable[0].length; j++){
    for(i = 1; i < rawTable.length - 1; i++){
      if(rawTable[i][j] == "" && rawTable[i-1][j] != "" && rawTable[i+1][j] != ""){
        rawTable[i][j] = "Junk space";
      }
    }
  }
  return rawTable;
}

//get events of week
function getEvents(initDate, finalDate, calendarIds){
  var events = [];
  var optionalArgs = {
    timeMin: initDate.toISOString(),
    timeMax: finalDate.toISOString(),
    showDeleted: false,
    singleEvents: true,
    orderBy: "startTime"
  };
  
  for(i = 0; i < calendarIds.length; i++){
    var eventsAux = Calendar.Events.list(calendarIds[i], optionalArgs).items;
    for(j = 0; j < eventsAux.length; j++){
      var eventAux = eventsAux[j];
      try{
        if(eventAux.summary != 'undefined' && eventAux.start.dateTime != 'undefined' && eventAux.end.dateTime != 'undefined'){
          Logger.log("Evento incluido", j);
          events.push(eventAux);
        }
      }catch(ex){
        Logger.log(ex);
      }
    }
  } 
  return events;
}

function getCalendarIds(){
  Logger.log('Getting Calendar Ids');  
  var response = Calendar.CalendarList.list().items;
  var ids = [];
  for(i = 0; i < response.length; i++){
      ids.push(response[i].id);
  }
  return ids;
}

//fill table with events
function prepareDataToInsert(mondayDate, weekEvents){
  var header = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];
  for(i = 0; i < header.length; i++){
    var d = (new Date(mondayDate));
    d.setDate(mondayDate.getDate() + i);
    Logger.log("D data:",d);
    var day = d.getDate();
    var month = d.getMonth();
    header[i] = header[i].concat(" ", day, "/", month);
  }
  var rows = [header];
  for(i = 1; i < ROWS_TABLE; i++){
    row = ["","","","","","",""];
    rows.push(row);
  }
  
  Logger.log("Week ",weekEvents);

  for (i = 0; i < weekEvents.length; i++) {
    try{
      var event = weekEvents[i];
      Logger.log("Event", event);
      
      var eventName = event.summary;
      Logger.log('Event name: ', eventName);
      Logger.log('Event date: %s', event.start.dateTime);
      
      var startDateTime = new Date(event.start.dateTime);
      Logger.log('Date created: %s', startDateTime.toISOString());
      var endDateTime = new Date(event.end.dateTime);
      Logger.log('EndDate created: %s', endDateTime.toISOString());
      
      var column = getColumn(startDateTime);
      Logger.log('Coluna: %s', column);
      
      var startRow = getRow(startDateTime);
      Logger.log('Start row: %s', startRow);
      var endRow = getRow(endDateTime);
      Logger.log('End row: %s', endRow);
      
      for(k = startRow; k < endRow; k++){
        rows[k][column] = eventName;
      } 
    }catch(ex){
      Logger.log(ex);
    }
  }
  return rows;     
}

//insert table in spreadsheet
function insertTable(spreadsheetId, table, tableRange) {
  var request = {
    'valueInputOption': 'USER_ENTERED',
    'data': [
      {
        'range': tableRange, //'Sheet1!B1:D2'
        'majorDimension': 'ROWS',
        'values': table
      }
    ]
  };
  var response = Sheets.Spreadsheets.Values.batchUpdate(request, spreadsheetId);
  Logger.log(response);
}

//Call everything
function main(){
  Logger.log('Starting Main');

  var mySpreadsheet = SpreadsheetApp.getActive();
  
  var mondayDate = getMonday(new Date());
  Logger.log('Monday', mondayDate);  
  
  var calendarIds = getCalendarIds();
  Logger.log('Calendar IDs: %s', calendarIds);

  for(var i = 0; i < SHEET_NAMES.length; i++){
    var initDate = new Date(mondayDate);
    switch(i){
      case 0:
        initDate.setDate(initDate.getDate()-7)
      break;
      case 1:
        //data correta
      break;
      case 2:
        initDate.setDate(initDate.getDate()+7)
      break;
    }

    var finalDate = getFinalDayOfWeek(initDate);
    Logger.log('Final', finalDate);  
  
    var weekEvents = getEvents(initDate, finalDate, calendarIds);
    Logger.log('Events from %s to %s', initDate, finalDate, weekEvents);

    Logger.log('Preparing data');  
    var rawTableToInsert = prepareDataToInsert(initDate, weekEvents);

    Logger.log('Cleaning');  
    var tableToInsert = cleanTable(rawTableToInsert);

    var range = "";
    range = range.concat("'",SHEET_NAMES[i],"'!C3:I36");

    Logger.log('Inserting in ', range);  
    insertTable(mySpreadsheet.getId(), tableToInsert, range);
  }
}

/////////////////////////////////////////////////////////////////////////////////////////////////////

//fill with x all the non usable times for classes
function xFill() {
  for(var i = 0; i < SHEET_NAMES.length; i++){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('C4').activate();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEET_NAMES[i]), true);
    spreadsheet.getCurrentCell().setValue('x');
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C4:C9'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    spreadsheet.getRange('C4:C5').activate();
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C4:I5'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    spreadsheet.getRange('E5').activate();
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('E5:E9'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    spreadsheet.getRange('C34').activate();
    spreadsheet.getCurrentCell().setValue('x');
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C34:I34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    spreadsheet.getRange('I14').activate();
    spreadsheet.getCurrentCell().setValue('x');
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('I14:I19'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    spreadsheet.getRange('C34:I34').activate();
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C34:I36'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    spreadsheet.getRange('C17').activate();
    spreadsheet.getCurrentCell().setValue('x');
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C17:C22'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    spreadsheet.getRange('C17:C22').activate();
    spreadsheet.setCurrentCell(spreadsheet.getRange('C22'));
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C17:F22'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    spreadsheet.getRange('C17:F22').activate();
    spreadsheet.setCurrentCell(spreadsheet.getRange('C22'));
  }
};
