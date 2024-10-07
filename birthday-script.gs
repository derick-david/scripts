// Define the calendar name
var calendarName = "JY Birthdays";

// Define the columns to use for names and birthdays
var nameColumn = "B";
var birthdayColumn = "E";

// Function to create birthday events
function createBirthdayEvents() {
  // Open the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the data from the specified columns (assuming the first row is the header)
  var names = sheet.getRange(nameColumn + "2:" + nameColumn).getValues();
  var birthdays = sheet.getRange(birthdayColumn + "2:" + birthdayColumn).getValues();
  
  // Get or create the calendar
  var calendars = CalendarApp.getCalendarsByName(calendarName);
  var calendar;
  
  if (calendars.length > 0) {
    calendar = calendars[0];
  } else {
    calendar = CalendarApp.createCalendar(calendarName);
  }
  
  // Loop through each row
  for (var i = 0; i < birthdays.length; i++) {
    var name = names[i][0];
    var birthday = birthdays[i][0];
    
    // Check if the cells are not empty
    if (name && birthday) {
      // Strip the year from the birthday
      var date = new Date(birthday);
      var month = date.getMonth();
      var day = date.getDate();
      
      // Create a new date object for the current year
      var eventDate = new Date(new Date().getFullYear(), month, day);
      
      // Check if an event already exists
      var events = calendar.getEventsForDay(eventDate);
      var eventExists = events.some(function(event) {
        return event.getTitle() === name + "'s Birthday";
      });
      
      // If no event exists, create a new recurring event
      if (!eventExists) {
        calendar.createAllDayEventSeries(
          name + "'s Birthday",
          eventDate,
          CalendarApp.newRecurrence().addYearlyRule()
        );
      } else {
        // Log if a duplicate event is found
        Logger.log("Duplicate event found for: " + name + " on " + eventDate);
      }
    }
  }
}
