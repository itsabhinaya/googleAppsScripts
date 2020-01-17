//Add a calendar invite
//Example sheet on how to set up https://docs.google.com/spreadsheets/d/1-7Ud1CIE2Wi4QeqFL2VF1X3LtJO8u_Nx5bhgVuRti78/edit?usp=sharing

function addEventToCalendar(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Calendar_Invites"); //Change the sheet name to what ever you have.
  var index = 2;
  
  var data2 = sheet.getDataRange().getValues();
  for (var i = 0, len = data2.length; i < len; i++) {
    if (data2[i][0] =='') break; 
  }
  var lastRow = i;


  for(;index <= lastRow; index++){
    var title = sheet.getRange(index, 1).getValue();
    var startTime = sheet.getRange(index, 2).getValue();
    var endTime = sheet.getRange(index, 3).getValue();
    var guests = sheet.getRange(index, 4).getValue();
    var location = sheet.getRange(index, 5).getValue();
    var description = sheet.getRange(index, 6).getValue();

    var sendInvites = true;
    
    var calendar = CalendarApp.getCalendarById("YOUR_EMAIL_HERE").createEvent(title, new Date(startTime), new Date(endTime),{description: description,location:location, guests: guests, sendInvites: sendInvites});
    
    var currentDate = new Date();
    sheet.getRange(index, 7).setValue("Sent at:"+ currentDate);
    
  }
}
