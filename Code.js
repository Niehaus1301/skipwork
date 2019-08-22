// Config vars
TEMPLATE_DOC_ID = "1syTKKGw9IQc-7uSxCjHfLz0xFTXkt0v8zDub_zd7y84"
FOLDER_ID = "1_ySK_tp-ogWuMCOwWI9t_tOQcKlgL4x2"
EMAIL_RECIPENT = "niehaus.1301@gmail.com"


// Main function which triggers on form submit
function handler() {

  // Get todays date
  var today = getDate()

  // Get form results
  var data = getFormResults()

  // Evaluate form results and generate text
  if (data[0] === "Yes") {
    var title = today
    var text = "am " + today
    var plain = today
  } else {
    if (data[2] === "") {
      var title = data[1]
      var text = "am " + data[1]
      var plain = data[1]
    } else {
      var title = data[1] + " - " + data[2]
      var text = "vom " + data[1] + " bis zum " + data[2]
      var plain = data[1] + " bis zum " + data[2]
    }
  }

  if (!(data[3] === "")) {
    var text = text + " aufgrund von " + data[3]
  }

  // Copy template and open copy
  var template = DriveApp.getFileById(TEMPLATE_DOC_ID);
  var destFolder = DriveApp.getFolderById(FOLDER_ID); 
  var copy = template.makeCopy(getDate(), destFolder).getId();
  var doc = DocumentApp.openById(copy)
  var body = DocumentApp.openById(copy).getBody()

  // Replace placeholders with data
  body.replaceText('{{date}}', today);
  body.replaceText('{{timereason}}', text);

  // Send confirmation email
  sendMail(doc.getUrl(), plain)
}

function getFormResults() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lrow = sh.getLastRow();

  var singleDay = sh.getRange(lrow, 2).getValue()
  var startDate = sh.getRange(lrow, 3).getDisplayValue()
  var endDate = sh.getRange(lrow, 4).getDisplayValue()
  var reason = sh.getRange(lrow, 5).getValue()
  
  return [singleDay, startDate, endDate, reason]
}


function getDate() {
  // Create new timestamp
  var d = new Date();
  var dd = String(d.getDate());
  var mm = String(d.getMonth() + 1)
  var yyyy = String(d.getFullYear());

  // Return in correct formatting
  return dd + '/' + mm + '/' + yyyy;
}

function sendMail(docURL, plain) {
  // Specify message
  var recipent = EMAIL_RECIPENT
  var subject = "Krankmeldung vom " + plain
  var message = "Die folgende Krankmeldung ist zum ausdrucken bereit:\n\n" + docURL 
  
  // Submit email
  MailApp.sendEmail(recipent, subject, message);
}