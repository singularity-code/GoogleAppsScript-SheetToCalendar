function create_event() {
    var calendar = CalendarApp.getCalendarById(''); // Your Calendar ID
    var formId = ''; // Your Form ID
    var data = get_sheetData();
    
    if(data) {
      for(var j = 0; j < data.length; j++) {
        calendar.createEvent(data[j].name, new Date(data[j].startTime), new Date(data[j].endTime), data[j].description);
      }
    }
  }

function get_sheetData() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getRange("Q1:Q");
    var value = range.getValues();
    var result = [];
    for(var i = 1; i < value.length; i++) {
      if(value[i] != '' && value[i] == 'Complete') {
        var resultRow = sheet.getRange(i + 1, 1, 1, 17).getValues(); // 17 because my one row ending have 17 columns
        var obj = {};

        //This indexes are depends on your sheet data structure
        obj.name = resultRow[0][1].toString();
        obj.startTime = resultRow[0][2].toString();
        obj.endTime= resultRow[0][8].toString();
        obj.leaveType = resultRow[0][4].toString();
        obj.description = {
          sendInvites: false,
          description: resultRow[0][5].toString(),
          guests: 'chris.lee@koi.edu.au'
        };
        result.push(obj);
        //Logger.log(result);  
      }
    }
    return result;
  }
