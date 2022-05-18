function ScheduleScrims()
{
var spreadsheet = SpreadsheetApp.getActiveSheet();
var calendarId = spreadsheet.getRange("C4").getValue(); 
/*z tej komórki bierzemy info o ID*/
var eventCal = CalendarApp.getCalendarById(calendarId);


var row = spreadsheet.getRange("A6:C37").getValues();
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
  var subject = shift[1];
  var description = shift[2];
var events = eventCal.getEventsForDay(startTime); /** weź zdarzenia w dniu startTime */
  for(var i=0; i<events.length;i++) //loop through all events
  {
    var ev = events[i];
    ev.deleteEvent(); /** usuwanie zdarzeń z dnia */
  }
if (subject!=""){
    


  eventCal.createEvent(subject, startTime, startTime).setDescription(description); /** zdarzenia wrzuca do kalendarza i dodaje do nich opis*/

}
 

}
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar')
  .addItem('Update scrim Calendar', 'ScheduleScrims')
  .addToUi();
}
