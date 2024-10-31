

// Purge messages automatically after how many days?
var DELETE_AFTER_DAYS = 30

// Maximum number of message threads to process per run. 
var PAGE_SIZE = 150

/**
 * Create a trigger that executes the purge function every day.
 * Execute this function to install the script.
 */
function setPurgeTrigger() {
  ScriptApp
    .newTrigger('purge')
    .timeBased()
    .everyDays(1)
    .create()
}

/**
 * Create a trigger that executes the purgeMore function two minutes from now
 */
function setPurgeMoreTrigger(){
  ScriptApp.newTrigger('purgeMore')
  .timeBased()
  .at(new Date((new Date()).getTime() + 1000 * 60 * 2))
  .create()
}

/**
 * Deletes all triggers that call the purgeMore function.
 */
function removePurgeMoreTriggers(){
  var triggers = ScriptApp.getProjectTriggers()
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i]
    if(trigger.getHandlerFunction() === 'purgeMore'){
      ScriptApp.deleteTrigger(trigger)
    }
  }
}

function removeAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

/**
 * Wrapper for the purge function
 */
function purgeMore() {
  purge();
}

/**
 * Deletes any emails from the inbox that are more than 7 days old
 * and not starred or marked as important.
 */
function purge() {
  removePurgeMoreTriggers();

  try {
    var search = 'in:inbox -in:starred -in:important older_than:' + DELETE_AFTER_DAYS + 'd';    
    var threads = GmailApp.search(search, 0, PAGE_SIZE);

    if (threads.length === PAGE_SIZE) {
      Logger.log('PAGE_SIZE exceeded. Setting a trigger to call the purgeMore function in 2 minutes.');
      setPurgeMoreTrigger();
    }

    Logger.log('Processing ' + threads.length + ' threads...');
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - DELETE_AFTER_DAYS);

    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      if (thread.getLastMessageDate() < cutoff) {
        thread.moveToTrash();
      }
    }
  } catch (e) {
    Logger.log("Error during purge: " + e.message);
  }
}



/**
 * Fetches the oldest 50 emails in the inbox that are not starred or important,
 * and logs the results in a Google Sheet.
 */
function logOldEmails() {
  // Search query for non-starred, non-important emails in the inbox
  var searchQuery = 'in:inbox -in:starred -in:important older_than:1d';
  var threads = GmailApp.search(searchQuery, 0, 50);
  
  // Create a new Google Sheet to display the results
  var sheet = SpreadsheetApp.create("Old Emails Log");
  var sheetName = sheet.getSheets()[0];
  
  // Add headers to the spreadsheet
  sheetName.appendRow(["Date", "Sender", "Subject"]);
  
  // Loop through each thread and add the details to the spreadsheet
  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    var oldestMessage = messages[0]; // Assumes messages are sorted by date in the thread
    
    // Extract details of the oldest message in the thread
    var date = oldestMessage.getDate();
    var sender = oldestMessage.getFrom();
    var subject = oldestMessage.getSubject();
    
    // Append the email details to the sheet
    sheetName.appendRow([date, sender, subject]);
  });

  Logger.log("Oldest 50 non-starred, non-important emails have been logged to the spreadsheet.");
}





