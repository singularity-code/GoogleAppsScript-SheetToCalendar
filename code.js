function create_event() {
  var calendar = CalendarApp.getCalendarById(''); // Place your calendar ID
  //var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  // To test with active sheet
  var sheet = SpreadsheetApp.openById("").getActiveSheet(); // Place your sheet ID connected with Form
  var formId = ''; // Place your Form ID
 
  // The ranges are depends on the structure of the sheet
  var data = get_sheetData();
  var names = sheet.getRange("B2:B").getValues();
  var status = sheet.getRange("Y2:Y").getValues();
  var eventIds = sheet.getRange("S2:S").getValues();
  var eventResult = sheet.getRange("Z2:Z").getValues();
  var targets = [];
    
  if(data) {
    for(var k = 0; k < status.length; k++) {
      if(eventResult[k].toString() != 'Y' && status[k] == 'Complete') {
        targets.push(k);
      }
    }
 
    for(var j = 0; j < targets.length; j++) {
      var result;
      var index = targets[j];
      result = calendar.createEvent(data[index].name, new Date(data[index].startTime), new Date(data[index].endTime), data[index].description);
      if(result) {
        sheet.getRange(data[index].rowNo, 19).setValue(result.getId());
        sheet.getRange(data[index].rowNo, 26).setValue("Y");
      }
    }
  }
}

// To convert DateTime format to PlanText on the sheet, otherwise calendar data will created in wrong format
function setDateToStringOnSheet() {
  var sheet = SpreadsheetApp.openById("1LkVbEzPjHWXB-_LY2INGudMMTLkJXnhsjRlfPm6cldc").getActiveSheet();
  var date1 = sheet.getRange("C2:C");
  var date2 = sheet.getRange("I2:I");
  
  date1.setNumberFormat("@");
  date2.setNumberFormat("@");
}

function get_sheetData() {
  var sheet = SpreadsheetApp.openById("1LkVbEzPjHWXB-_LY2INGudMMTLkJXnhsjRlfPm6cldc").getActiveSheet();
  //var range = sheet.getRange(2,17,50,1);
  var range = sheet.getRange("Y1:Y");
  var value = range.getValues();
  
  var result = [];

  for(var i = 1; i < value.length; i++) {
    if(value[i] != '') {
      var resultRow = sheet.getRange(i + 1, 1, 1, 17).getValues();
      var obj = {};
      obj.name = resultRow[0][1].toString() + " " + resultRow[0][4].toString();
      obj.startTime = resultRow[0][2].toString();
      obj.endTime= resultRow[0][8].toString();
      obj.leaveType = resultRow[0][4].toString();
      obj.description = {
        sendInvites: false,
        description: resultRow[0][5].toString(),
        guests: ''
      };
      obj.rowNo = i + 1;
      result.push(obj);
    }
  }
  return result;
}