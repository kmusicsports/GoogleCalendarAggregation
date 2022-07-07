function myFunction() {
  
  var rowNum = 1;
  var columNum = 3;
  var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  var range = sheet.getRange(rowNum, columNum, 6, 1);
  var values = range.getValues();

  var startDate = new Date(values[0][0], values[1][0] - 1, values[2][0]);
  var endDate = new Date(values[3][0], values[4][0] - 1, values[5][0]);

  var calendars = CalendarApp.getAllCalendars();
  rowNum = 8;
  for (var i = 1; i < calendars.length; i++) {

    var calendar = calendars[i];
    var calendarName = calendar.getName();

    var firstChar = calendarName.substring(0, 1);
    if((firstChar != "D" && firstChar != "H") || calendarName == "D_誕生日") continue;

    var cell = sheet.getRange(rowNum, 1);
    cell.setValue(calendarName);
    cell.setBackground(calendar.getColor());

    Logger.log("[" + calendarName + "]");
    var sum = 0;
    var date = new Date(startDate.getTime());
    while(date.getTime() <= endDate.getTime()) {
      var schedules = calendar.getEventsForDay(date);
      for (var j = 0; j < schedules.length; j++) {
        sum += schedules[j].getEndTime() - schedules[j].getStartTime();
        Logger.log("- " + schedules[j].getTitle())
      }

      date.setDate(date.getDate() + 1);
    }
    cell = sheet.getRange(rowNum, 2);
    var sumHours = sum / 3600000;
    cell.setValue(sumHours);
    rowNum++;
  }
}
