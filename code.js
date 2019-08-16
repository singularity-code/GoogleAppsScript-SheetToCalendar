function create_event() {
  var calendar = CalendarApp.getCalendarById(''); // Place your calendar ID
  //var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  // To test with active sheet
  var sheet = SpreadsheetApp.openById("").getActiveSheet(); // Place your sheet ID connected with Form
  var formId = ''; // Place your Form ID
  var data = get_sheetData();
  var names = sheet.getRange("B2:B").getValues();
  var eventIds = sheet.getRange("R2:R").getValues();
  var status = sheet.getRange("Q2:Q").getValues();
  var targets = [];
    
  //Logger.log(data);
  if(data) {
    for(var k = 0; k < eventIds.length; k++) {
      if(eventIds[k].toString() != 'Y' && status[k] == 'Complete') {
        targets.push(k);
      }
    }
 
    Logger.log(targets);
    for(var j = 0; j < targets.length; j++) {
      var result;
      var index = targets[j];
      result = calendar.createEvent(data[index].name, new Date(data[index].startTime), new Date(data[index].endTime), data[index].description);
      sheet.getRange(data[index].rowNo, 11).setValue(result.getId());
      sheet.getRange(data[index].rowNo, 18).setValue("Y");
    }
  }
}
function get_sheetData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("Q1:Q");
  var value = range.getValues();
  var result = [];
  //Logger.log(value)
  //Logger.log(value.length);
  for(var i = 1; i < value.length; i++) {
    if(value[i] != '') {
      //Logger.log(value[i].toString());
      //Logger.log(sheet.getRange(i + 1, 1, 1, 17).getValues());
      var resultRow = sheet.getRange(i + 1, 1, 1, 17).getValues();
      var obj = {};
      obj.name = resultRow[0][1].toString() + " " + resultRow[0][4].toString();
      obj.startTime = resultRow[0][2].toString();
      obj.endTime= resultRow[0][8].toString();
      obj.leaveType = resultRow[0][4].toString();
      obj.description = {
        sendInvites: false,
        description: resultRow[0][5].toString(),
        guests: 'chris.lee@koi.edu.au'
      };
      obj.rowNo = i + 1;
      result.push(obj);
    }
  }
  return result;
}
