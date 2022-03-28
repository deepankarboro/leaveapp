var calendarId = "Put calendar ID here"; //Data Team Calendar
var sheet = SpreadsheetApp.getActive().getSheetByName("FY 2021-22");
var formTimeStampId = 1; //timestamp of the form submission
var userId = 2; //autologs the emailID based on google account
var eventId = 3;
var leaveTypeId = 4; //Leave Type
var reasLeaveId = 5; //Reason for Leave
var startDtId = 6; //Start date and time
var endDtId = 7; //End date and time

//Activity for deleting rows from between the spreadsheet. Had taken this from https://www.googlecloudcommunity.com/gc/Tips-Tricks/How-to-delete-blank-rows-in-Google-Sheets/m-p/383137. Kudos to Greg for this.
function deleteBlankRows() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  // Get sheets
  var sheets = SS.getSheets();
  // Loop through sheets. Delete blank rows in each sheet.
  for (var s=0; s < sheets.length; s++) {
    var currentSheet = sheets[s];
    var sheetName = sheets[s].getName();
    var searchDataRange = currentSheet.getRange(1,1,currentSheet.getMaxRows(),currentSheet.getMaxColumns()); // get the ENTIRE sheet. not just where the data is.
    var searchValues = searchDataRange.getValues();
    var numRows = searchValues.length;
    var numCols = searchDataRange.getNumColumns();
    var rowsToDel = [];
    var delRow = -1;
    var prevDelRow = -2;
    var rowClear = false;
    
    // Loop through Rows in this sheet
    for (var r=0; r < numRows; r++) {
      
      // Loop through columns in this row
      for (var c=0; c < numCols; c++) {
        if (searchValues[r][c].toString().trim() === "") {
          rowClear = true;
        } else {
          rowClear = false;
          break;
        }
      }
      
      // If row is clear, add it to rowsToDel
      if (rowClear) {
        if (prevDelRow === r-1) {
          rowsToDel[delRow][1] = parseInt(rowsToDel[delRow][1]) + 1;
        } else {
          rowsToDel.push([[r+1],[1]]);
          delRow += 1;
        }
        prevDelRow = r;
      }
    }
    Logger.log("numRows: " + numRows);
    Logger.log("rowsToDel.length: " + rowsToDel.length);
    
    // Delete blank rows in this sheet, if we have rows to delete.
    if (rowsToDel.length>0) {
      // We need to make sure we don't delete all rows in the sheet. Sheets must have at least one row.
      if (numRows === rowsToDel[0][1]) {
        // This means the number of rows in the sheet (numRows) equals the number of rows to be deleted in the first set of rows to delete (rowsToDel[0][1]).
        // Delete all but the first row.
        if (numRows > 1) {
          currentSheet.deleteRows(2,numRows-1);
        }
      } else {
        // Go through each set of rows to delete them.
        var rowsToDeleteLen = rowsToDel.length;  
        for (var rowDel = rowsToDeleteLen-1; rowDel >= 0; rowDel--) {
          currentSheet.deleteRows(rowsToDel[rowDel][0],rowsToDel[rowDel][1]);
        }
      }
    }
  }
}


//Update Calendar Event IDs through this function, if event id not available on calendar.
function updateEventIDs() // Read from the Marketing calendar and update column 3(individual event ID) to google sheet
{
  var now = new Date();
  var data = sheet.getDataRange().getValues();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
  });
  Logger.log(events.items.length);
  Logger.log(data.length);
  for (var i = 0; i<events.items.length; i++)
  {
    var event = events.items[i];
    Logger.log("event desc- "+event.description);
    for (var j = 0; j<data.length; j++)
    {
      var desc = data[j][3] + "\n" + data[j][4] + "\n" + "Added by: " + data[j][1] + "\n" + "Submitted on: " + data[j][0];
      if(event.description == desc)
      {
        sheet.getRange(j+1,3).setValue(event.id);
        Logger.log("Updating for "+data[j][1]+"with event id: "+event.id+"\nEvents Details Below:\n"+event.description);
      }
    }
  }
}
//Get trigger value from "On Change Trigger"
function onChangeTrigger(e)
{  
  if(e.changeType=="INSERT_ROW")
    addNewEvent()
  else if(e.changeType=="EDIT")
  {
    Logger.log("Edit Triggered");
    forModificationDeletion();
  }  
  else
    Logger.log("No new events found");
}

// Modifying or deleting existing calendar event
function forModificationDeletion()
{
  var now = new Date();
  //var sheet = SpreadsheetApp.getActiveSheet();
  var sheet = SpreadsheetApp.getActive().getSheetByName("FY 2021-22");
  var rows = sheet.getDataRange();
  var lr = rows.getLastRow();
  var searchflag = 0;
  Logger.log("last row = "+lr);
  var data = sheet.getDataRange().getValues();
  var eventcounter=0;
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
  });
  Logger.log("Total Rows "+data.length);
  Logger.log("Total Future Calendar Events "+events.items.length);
  for(var i=lr; i>=2; i--)
  {
    var startDt = sheet.getRange(i,startDtId,1,1).getValue();
    var endDt = sheet.getRange(i,endDtId,1,1).getValue();
    var eventID = sheet.getRange(i,eventId,1,1).getValue();
    var emailID = sheet.getRange(i,userId,1,1).getValue();
    var timeStamp = sheet.getRange(i,formTimeStampId,1,1).getValue();
    var leaveType = sheet.getRange(i,leaveTypeId,1,1).getValue();
    var reasLeave = sheet.getRange(i,reasLeaveId,1,1).getValue();
    var subOn = "Submitted on: "+timeStamp;
    var desc = leaveType+"\n"+reasLeave+"\n"+"Added by: "+ emailID+"\n"+subOn;
    var title = emailID+" - "+leaveType;
    //Logger.log("Title: "+title+"\nDescription: "+desc+"\nStart Date: "+startDt+"\nEnd Date: "+endDt);
    for(var j = 0; j<events.items.length; j++)
    {
      //Checking for change in event details
      if(eventID==events.items[j].id)
      {
        var stdateformatted = Utilities.formatDate(new Date(events.items[j].start.dateTime), 'Asia/Kolkata', 'MMMM dd, yyyy HH:mm:ss Z');
        var startDtformatted = Utilities.formatDate(startDt, 'Asia/Kolkata', 'MMMM dd, yyyy HH:mm:ss Z');
        var enddateformatted = Utilities.formatDate(new Date(events.items[j].end.dateTime), 'Asia/Kolkata', 'MMMM dd, yyyy HH:mm:ss Z');
        var endDtformatted = Utilities.formatDate(endDt, 'Asia/Kolkata', 'MMMM dd, yyyy HH:mm:ss Z');
        if(desc!=events.items[j].description || startDtformatted!=stdateformatted || endDtformatted!=enddateformatted)
        {
          Logger.log("Calendar updated for Event ID: "+eventID);
          Logger.log("Original Event Details:"+"\nStart Date: "+events.items[j].start.dateTime+"\nEnd Date: "+events.items[j].end.dateTime+"\nEvent Details: "+events.items[j].description);
          updateEvent(eventID,title,desc,startDt,endDt);
        }
        else
          Logger.log("No event changes for Event ID:"+eventID);
        eventcounter=eventcounter+1;
      }
    }
  }
  Logger.log("Event Counter is "+eventcounter);
  var delevntctrs = events.items.length-eventcounter;
  if(events.items.length == eventcounter)
    Logger.log("No Event Deleted");
  else
  {
    Logger.log("Total Events Deleted: "+delevntctrs);
    //Checking for deleted events on gsheets which is still on calendar.
    for(var i = 0; i<events.items.length; i++)
    {
      eventcounter = 1;
      for(var j=lr; j>=2; j--)
      {
        var eventID = sheet.getRange(j,eventId,1,1).getValue();
        if(eventID==events.items[i].id)
          eventcounter=0;
      }
      if(eventcounter == 1)
      {
        Calendar.Events.remove(calendarId,events.items[i].id);
        Logger.log("Event with event ID: "+events.items[i].id+" is removed from calendar");
      }
    }
  }  
  deleteBlankRows(); //Delete Blank Rows if any  
}

//Used for deleting all events from the Calendar.
function deleteAllCalendarEvents() 
{
  var now = new Date();
  var start = new Date(now.getTime() - (112 * 24 * 60 * 60 * 1000)); //Check the first start date from google sheets and change the formula accordingly.
  var end = new Date(now.getTime() + (300 * 24 * 60 * 60 * 1000)); //Check the last end date from google sheets and change the formula accordingly.
  Logger.log("Start Date Time: "+start);
  Logger.log("End Date Time: "+end);
  var events = CalendarApp.getCalendarById(calendarId).getEvents(start, end);
  Logger.log(events.length);
  var counter = 0;
  for(var i = 0; i<events.length; i++)
  {
    events[i].deleteEvent();
    Logger.log("Deleted event with description:\n"+events[i].getDescription());
    counter++;
  }
  Logger.log("Deleted Events "+counter);
  /*for (var i = 0; i<events.items.length; i++) //This part is only to be used if Calendar API service is used for calling calendar services.
  {
    var event = events.items[i];
    Logger.log("Removing event with Event ID: "+event.id+"\n"+"Event Description: "+event.description+"\nRow Processed: "+i);
    Calendar.Events.remove(calendarId,event.id);
  }*/
}

//Add All Evets from Google Sheets to Calendar based on the selected Calendar ID
function addingAllEventsfromGsheets()
{
  var sheet = SpreadsheetApp.getActive().getSheetByName("FY 2021-22");
  var rows = sheet.getDataRange();
  var counter = 0;
  var lr = rows.getLastRow();
  Logger.log("Current Last Row is "+lr);
  var counter = 0;
  for(var i=2; i<=lr; i++)
  {
    var startDt = sheet.getRange(i,startDtId,1,1).getValue();
    var endDt = sheet.getRange(i,endDtId,1,1).getValue();
    var emailID = sheet.getRange(i,userId,1,1).getValue();
    var timeStamp = sheet.getRange(i,formTimeStampId,1,1).getValue();
    var leaveType = sheet.getRange(i,leaveTypeId,1,1).getValue();
    var reasLeave = sheet.getRange(i,reasLeaveId,1,1).getValue();
    var subOn = "Submitted on: "+timeStamp;
    var desc = leaveType+"\n"+reasLeave+"\n"+"Added by: "+ emailID+"\n"+subOn;
    var title = emailID+" - "+leaveType;
    Logger.log("Start Date: "+startDt+"\nEnd Date: "+endDt+"\nTitle: "+title+"\nDescription: "+desc);
    var eventid = createEvent(calendarId,startDt,endDt,title,desc);
    sheet.getRange(i,3).setValue(eventid);
    Logger.log("Created event with Event ID: "+eventid+" Row processed "+i);
    counter=counter+1;
  }
  Logger.log("Total Events Added: "+counter);
}

//Get last row values from google sheets and call calendar function for adding a new calendar event
function addNewEvent() {
  deleteBlankRows(); //Delete blank rows if any.
  var sheet = SpreadsheetApp.getActive().getSheetByName("FY 2021-22");
  var rows = sheet.getDataRange();
  var lr = rows.getLastRow();
  Logger.log("Current Last Row is "+lr);
  var startDt = sheet.getRange(lr,startDtId,1,1).getValue();
  var endDt = sheet.getRange(lr,endDtId,1,1).getValue();
  var emailID = sheet.getRange(lr,userId,1,1).getValue();
  var timeStamp = sheet.getRange(lr,formTimeStampId,1,1).getValue();
  var leaveType = sheet.getRange(lr,leaveTypeId,1,1).getValue();
  var reasLeave = sheet.getRange(lr,reasLeaveId,1,1).getValue();
  var subOn = "Submitted on: "+timeStamp;
  var desc = leaveType+"\n"+reasLeave+"\n"+"Added by: "+ emailID+"\n"+subOn;
  var title = emailID+" - "+leaveType;
  var maildomainchecker= "";
  var empName = "";
  for(var i=(emailID.length-14);i<emailID.length;i++)
    maildomainchecker=maildomainchecker+emailID[i];
  for(var j=0;j<(emailID.length-14);j++)
    empName = empName+emailID[j];
  if(startDt!="" && endDt!="" && emailID!="" && timeStamp!="" && leaveType!="" && maildomainchecker == "@<email domain here>")
  {
      Logger.log("<Company Employee>");
      var subject = "New Leave Created on "+timeStamp;
      var messageBody ="\n\nA new leave has been created. Please find details below for the same.\n"+"LeaveType: "+leaveType+"\nReason for Leave: "+reasLeave+"\nLeave Start Date & Time: "+startDt+"\nLeave End Date & Time: "+endDt;
      var signature = "\nRegards,\nDeepankar";
      var message = "Dear "+empName+","+messageBody+signature;
      if(startDt<endDt)
      {
        var eventid = createEvent(calendarId,startDt,endDt,title,desc);
        MailApp.sendEmail(emailID,subject,message);
        sheet.getRange(lr,3).setValue(eventid);
        Logger.log("New event created for "+emailID+"\nEvent Details Below:\n"+"Event ID: "+eventid+"\n"+desc); 
      }
      else
      {
        var subject = "Error with dates "+timeStamp;
        var messageBody ="\n\nLeave can't be created since start date can't be greater than end date.\nCreate Leave again.";
        var signature = "\nRegards,\nDeepankar";
        var message = "Dear "+empName+","+messageBody+signature;
        MailApp.sendEmail(emailID,subject,message);
        sheet.deleteRow(lr);
        Logger.log("Start Date is greater than End Date.\nEntry deleted for row "+lr+" useremail: "+emailID);
      }    
  }
  else
  {
    if(maildomainchecker != "@<email domain here>")
    {
        Logger.log("Not Company Employee");
        sheet.deleteRow(lr);
        Logger.log("Entry deleted for row "+lr+" useremail: "+emailID);
    }
    else
    {
        Logger.log("All required details aren't filled");
        var subject = "Recent leave entry deleted!!";
        var messageBody = "\n\nAll required fields for leave details aren't filled."+"\nThe existing entry has been deleted"+"\nPlease fill up a new leave request for all required details.";
        var signature = "\nRegards,\nDeepankar";
        var message = "Dear "+empName+","+messageBody+signature;
        MailApp.sendEmail(emailID,subject,message);
        sheet.deleteRow(lr);
        Logger.log("Entry deleted for row "+lr+" useremail: "+emailID);
    }
  }
}

//Creating New Calendar Event
function createEvent(calendarId,startDt,endDt,title,desc)
{
  var start = new Date(startDt);
  var end = new Date(endDt);
  Logger.log(start);
  Logger.log(end);
  var event = {
    summary: title,
    description: desc,
    start: {
      dateTime: start.toISOString()
    },
    end: {
      dateTime: end.toISOString()
    },
  };
  event = Calendar.Events.insert(event, calendarId);
  Logger.log('New event created with Event ID: ' + event.id);
  return event.id;
}

//Updating already existing Calendar Event
function updateEvent(eventID,title,desc,startDt,endDt)
{
  var start = new Date(startDt);
  var end = new Date(endDt);
  var event = {
    summary: title,
    description: desc,
    start: {
      dateTime: start.toISOString()
    },
    end: {
      dateTime: end.toISOString()
    },
  };
  Calendar.Events.update(event,calendarId,eventID);
  Logger.log("Modified Event Details: \n"+"Start Date: "+start+"\nEnd Date: "+end+"\nEvent Details: "+event.description);
}
