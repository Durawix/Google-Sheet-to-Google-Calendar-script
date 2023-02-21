function ScheduleScrims()
{
var spreadsheet = SpreadsheetApp.getActiveSheet();
var calendarId = spreadsheet.getRange("G4").getValue(); 
/*z tej komórki bierzemy info o ID*/
 if (!calendarId) {
    Logger.log("Invalid calendar ID");
    return;
  }
var eventCal = CalendarApp.getCalendarById(calendarId);
if (!eventCal) {
    Logger.log("Invalid calendar object"); /* it shows if something is wrong with id of calendar*/
    return;
  }

var row = spreadsheet.getRange("A6:H44").getValues();
/*określamy zakres z którego bierzemy info o wartościach w komórkach*/
/** [13/05/2022 18:00:00	Scrim vs BeKind	3 maps Villa kafe Theme]
    [14/05/2022 18:00:00	Scrim vs BB	3 maps Villa kafe Theme]
    [15/05/2022 18:00:00	Scrim vs G2	3 maps Villa kafe Theme]
    [16/05/2022 18:00:00	Scrim vs Rogue	3 maps Villa kafe Theme]
    [17/05/2022 18:00:00	Scrim vs NaVi	3 maps Villa kafe Theme] **/

/** make it that everyone can use it **/

for (var x=0; x<row.length; x++) { /**x rośnie o 1 dopóki x jest mniejszy niż liczba wierszy */
  var shift = row[x]; /** shift to nazwa zmiennej a row.length (wiersz) to wartości z danych wierszy */

  var startTime = shift[0]; /** */
  // var start_hour = shift[3]
  var subject = shift[5];
  var description = shift[6];
  var durationInHours = shift[7];
//Logger.log('durationInHours: ' + durationInHours)
;
var endTime = new Date(startTime.getTime() + (durationInHours * 60 *60*1000)); // ustawiamy koniec wydarzenia na godzinę później od początku
//DLA SPRAWDZENIA Logger.log('endTime: ' + endTime); // wyświetlamy wartość zmiennej w konsoli deweloperskiej

var events = eventCal.getEventsForDay(startTime); /** weź zdarzenia w dniu startTime */
  for(var i=0; i<events.length;i++) //loop through all events
  {
    var ev = events[i];
    ev.deleteEvent(); /** usuwanie zdarzeń z dnia */
  }
if (subject!=""){
    


  eventCal.createEvent(subject, startTime, endTime).setDescription(description); /** zdarzenia wrzuca do kalendarza i dodaje do nich opis*/

}
 

}
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar')
  .addItem('Update scrim Calendar', 'ScheduleScrims')
  .addToUi();
}
