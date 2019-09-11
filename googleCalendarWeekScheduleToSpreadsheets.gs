//Global
var ROWS_TABLE = 34;
var TABLE_STARTING_HOUR = 6;
var SHEET_NAMES = ["Last Week","Current Week","Next Week"];

//prepare spreadsheet
function prepare(mySheet){
  mySheet.deleteColumns(1,20);
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

function getColumn(d){
  return Math.abs(d.getDay() == 0 ? 6:(d.getDay()-1));
}

function getRow(d){
  var rowHour = d.getHours() - TABLE_STARTING_HOUR;
  var rowMinute = Math.round(d.getMinutes()/30);
  var row = rowHour*2 + rowMinute + 1;//header
  return row;
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
//insert table

/**
 * Write to multiple, disjoint data ranges.
 * @param {string} spreadsheetId The spreadsheet ID to write to.
 */
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

//format table


//Call everything
function main(){
  Logger.log('Starting Main');

  var mySpreadsheet = SpreadsheetApp.getActive();
  
  var mondayDate = getMonday(new Date());
  Logger.log('Monday', mondayDate);  
  
  var calendarIds = getCalendarIds();
  Logger.log('Calendar IDs: %s', calendarIds);

  for(i = 0; i < SHEET_NAMES.length; i++){
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
    range.concat("'",SHEET_NAMES[i],"'!C3:I36");

    Logger.log('Inserting in ', range);  
    insertTable(mySpreadsheet.getId(), tableToInsert, range);
  }
}


function test(){ 
  
  var mondayDate = getMonday(new Date());
  Logger.log('Monday',mondayDate);  
  var finalDate = getFinalDayOfWeek(mondayDate);
  Logger.log('Final', finalDate);
                
  //var calendarIds = getCalendarIds();
  //Logger.log('Calendar IDs: %s', calendarIds);

  //var events = getEvents(mondayDate, finalDate, calendarIds);
  //Logger.log('Events from %s to %s', mondayDate, finalDate, weekEvents);

  Logger.log('Preparing data');  
  var rawTableToInsert = prepareDataToInsert(mondayDate, [[1]]);

  Logger.log('Cleaning');  
  var tableToInsert = cleanTable(rawTableToInsert);

  Logger.log('Inserting');  
  insertTable(mySpreadsheet.getId(), tableToInsert, range);
}
