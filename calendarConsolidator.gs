/*
SYNOPSIS
--------
This script creates a new calendar, temporarily
subscribes to Google calendars given by ID, and 
consolidates 2 years of events (the last and 
next year) into the new calendar before unsubscribing.

As an example, the script is currently 
set up to consolidate several publicly 
available sports calendars into a newly
created Google calendar named "Sports".

Be aware of Google Calendar use limits: 
https://support.google.com/a/answer/2905486?hl=en

LICENSE: MIT (c) 2020 Jordan Bradford

GITHUB: jrdnbradford
*/


var consolidationCalendarName = "Sports";

function consolidateCalendarsById(){
    var consolidationCalendarIds = [
        "mlb_2_%42oston+%52ed+%53ox#sports@group.v.calendar.google.com", // Boston Red Sox
        "nfl_17_%4eew+%45ngland+%50atriots#sports@group.v.calendar.google.com", // NE Patriots
        "ncaaf_68_%47eorgia+%42ulldogs#sports@group.v.calendar.google.com" // GA Bulldogs
    ];
  
    // This new calendar will hold all events of interest
    var consolidationCalendar = CalendarApp.createCalendar(consolidationCalendarName);
    consolidationCalendar
        .setDescription("Created with Google Apps Script.\nGitHub: jrdnbradford")
        .setSelected(true);
  
    var today = new Date();
    var year = today.getFullYear();
    var month = today.getMonth();
    var day = today.getDate();
    var yearAgo = new Date(year - 1, month, day);
    var yearFromNow = new Date(year + 1, month, day);

    // Temporarily subscribe to each calendar to copy their respective events
    consolidationCalendarIds.forEach(function(id){
        var calendarToConsolidate = CalendarApp.subscribeToCalendar(id, {selected: false});
        var events = calendarToConsolidate.getEvents(yearAgo, yearFromNow);
        events.forEach(function(event){
            var title = event.getTitle();
            var startTime = event.getStartTime();
            var endTime = event.getEndTime();  
            var description = event.getDescription();
            var location = event.getLocation();
            consolidationCalendar.createEvent(title, startTime, endTime, {
                description: description,
                location: location
            });
            // Prevents Google error
            Utilities.sleep(100);
        });
        // Unsubscribe after copying events to consolidation calendar
        calendarToConsolidate.unsubscribeFromCalendar();      
    });   
}