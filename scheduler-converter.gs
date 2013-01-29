function run() {
  var gRow = 8;
  var gCol = 1;
  var pRow = gRow
  var pCol = gCol + 2;
  var events = [];
  
  var calName = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  pCol = SpreadsheetApp.getActiveSheet().getActiveCell().getColumn();
  //var value = SpreadsheetApp.getActiveSheet().getRange(8, 1, 6, 5).getValues();
  
  //var row = SpreadsheetApp.getActiveSheet().getRange(9, 1).getNumRows();
  var lastRow = SpreadsheetApp.getActiveSheet().getLastRow() + 1;
  
  //var value = SpreadsheetApp.getActiveSheet().getRange(8, 1, lastRow - 6, 5).getValues();
  //Browser.msgBox(value);
  
  var graphic = SpreadsheetApp.getActiveSheet().getRange(gRow, gCol, lastRow - gRow, gCol + 1).getValues();
  var predmets = SpreadsheetApp.getActiveSheet().getRange(pRow, pCol, lastRow - pRow, pCol + 2).getValues();
  var curDay;
  for (var row = 0; row < lastRow - gRow; row++) {
    if (graphic[row][gCol-1] != '' && graphic[row][gCol-1] != undefined)
      curDay = graphic[row][gCol-1];
    
    var discipline = predmets[row][0];
    var teacher = predmets[row][1];
    var room = predmets[row][2];
    if (discipline != '' && discipline != undefined &&
        teacher != '' && teacher != undefined &&
        room != '' && room != undefined) {
      events.push(createEvent(curDay, graphic[row][gCol], discipline, teacher, room));
    }
  }

  for (var i = 0; i < events.length; i++)
    writeEvent(calName, events[i]);
}

function getCurDay(sCurDay) {
  var data = sCurDay.substring(sCurDay.indexOf(' ')+1, sCurDay.length);
  var day = parseInt((data[0] == ' ') ? data[1] : data.substring(0, 2));
  var month = data.substring(3,5);
  month = parseInt(month[0] == '0' ? month[1] : month) - 1;
  var year = new Date();
  year = year.getFullYear();
  
  return new Date(year, month, day);
}

function getStartEndTime(sTime) {
  var result = [];
  result.push(getTime(sTime.substring(0, 5)));
  result.push(getTime(sTime.substring(6, 11)));
  return result;
}

function getTime(sTime) {
  var hour = sTime.substring(0, 2);
  // hour = parseInt(hour[0] == '0' ? hour[1] : hour);
  // Какие-то проблемы с часовыми поясами. "-1" - временное решение.
  hour = parseInt(hour[0] == '0' ? hour[1] : hour) - 1;
  var minute = sTime.substring(3,5);
  minute = parseInt(minute[0] == '0' ? minute[1] : minute);
  
  return [hour, minute];
}

function createEvent(curDay, time, discipline, teacher, room) {
  result = [];
  result.push(getCurDay(curDay));
  result.push(getStartEndTime(time));
  result.push(discipline);
  result.push(teacher);
  result.push(room);
  
  return result;
}

function writeEvent(calName, event) {
  var cal;
  var allCals = CalendarApp.getAllCalendars();
  for (var i = 0; i < allCals.length; i++) {
    var tmpCal = allCals[i];
    if (tmpCal.getName() == calName) {
      cal = tmpCal;
      break;
    }
  }
  
  if (cal == undefined)
    cal = CalendarApp.createCalendar(calName);
  
  var year = event[0].getFullYear();
  var month = event[0].getMonth();
  var day = event[0].getDate();
  var startDate = new Date(year, month, day, event[1][0][0], event[1][0][1]);
  var finishDate = new Date(year, month, day, event[1][1][0], event[1][1][1]);
  cal.createEvent(event[2], startDate, finishDate, 
                  {location:event[4], description:event[3]});
}

function clearCalendar() {
  var calName = 'Scheduler'
  var cal;
  var allCals = CalendarApp.getAllCalendars();
  for (var i = 0; i < allCals.length; i++) {
    var tmpCal = allCals[i];
    if (tmpCal.getName() == calName) {
      cal = tmpCal;
      break;
    }
  }
  
  var startTime;
  var endTime;
  var today = new Date();
  var month = today.getMonth();
  
  if (month < 3) {
    startTime = new Date(today.getYear(), 0, 1);
    endTime = new Date(today.getYear(), 2, 1);
  } else {
    startTime = new Date(today.getYear(), 4, 1);
    endTime = new Date(today.getYear(), 7, 1);
  }
  
  var events = cal.getEvents(startTime, endTime);
  for (var i = 0; i < events.length; i++)
    events[0].deleteEvent();
}  
