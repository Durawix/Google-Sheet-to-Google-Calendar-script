function ScheduleScrims() {
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var calendarId = spreadsheet.getRange("F4").getValue();
  var eventCal = CalendarApp.getCalendarById(calendarId);

  var range = spreadsheet.getRange("A6:H37");
  var values = range.getValues();

  for (var x = 0; x < values.length; x++) {
    var row = values[x];
    var startTime = new Date(row[0]); // Convert the text value to a Date object
    var subject = row[4];
    var description = row[5];
    var durationHours = parseFloat(row[7]); // Duration in hours, assuming it's in the 7th column
    var location = row[6]; // Location value from the 6th column

    var events = eventCal.getEventsForDay(startTime);
    for (var i = 0; i < events.length; i++) {
      var ev = events[i];
      if (ev.getTitle() === subject && ev.getStartTime().getTime() === startTime.getTime()) {
        ev.deleteEvent();
        break; // Exit the loop after deleting the event
      }
    }

    if (subject !== "" && durationHours > 0) {
      var durationMinutes = Math.round(durationHours * 60); // Convert duration from hours to minutes
      var endTime = new Date(startTime.getTime() + durationMinutes * 60 * 1000);
      var event = eventCal.createEvent(subject, startTime, endTime).setDescription(description);

      if (location !== "") {
        event.setLocation(location);
      }
    }
  }
}

function onOpen() { //it adds button for using sync
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar')
  .addItem('Update Calendar', 'ScheduleScrims')
  .addToUi();
}
