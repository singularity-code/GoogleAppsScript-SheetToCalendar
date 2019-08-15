function create_event() {
    var calendar = CalendarApp.getCalendarById(''); // Your Calendar ID
    var formId = ''; // Your Form ID
    var data = get_sheetData();
    
    var data = get_sheetData();
  
    //Logger.log(data);
    if(data) {
      var names = sheet.getRange("B2:B").getValues();
      var eventIds = sheet.getRange("K2:K").getValues();
      var status = sheet.getRange("Q1:Q").getValues();
      
      for(var k = 0; k < eventIds.length; k++) {
        if(eventIds[k].toString() == '' && status[k] == 'Complete') {
          Logger.log("Create!");
          for(var j = 0; j < data.length; j++) {
            var event = calendar.getEvents(new Date(data[j].startTime), new Date(data[j].endTime));
            var result;
            result = calendar.createEvent(data[j].name, new Date(data[j].startTime), new Date(data[j].endTime), data[j].description);
            // Add event Id to sheet column
            sheet.getRange(data[j].rowNo, 11).setValue(result.getId());
          }
        }
      }
    }
  }
  
  function get_sheetData() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    //var range = sheet.getRange(2,17,50,1);
    var range = sheet.getRange("Q1:Q");
    var value = range.getValues();
    var result = [];
    //Logger.log(value)
    //Logger.log(value.length);
    for(var i = 1; i < value.length; i++) {
      if(value[i] != '' && value[i] == 'Complete') {
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
        //Logger.log(result);  
      }
    }
    return result;
  }
