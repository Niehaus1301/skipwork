TEMPLATE_DOC_ID = "1syTKKGw9IQc-7uSxCjHfLz0xFTXkt0v8zDub_zd7y84"
FOLDER_ID = "1_ySK_tp-ogWuMCOwWI9t_tOQcKlgL4x2"
EMAIL_RECIPENT = "niehaus.1301@gmail.com"


function handler() {
  var template = DriveApp.getFileById(TEMPLATE_DOC_ID);
  var destFolder = DriveApp.getFolderById(FOLDER_ID); 
  var copy = template.makeCopy(getDate(), destFolder).getId();
  var doc = DocumentApp.openById(copy)
  var body = DocumentApp.openById(copy).getBody()

  var data = getFormData()

  if (data[0] === "Ja") {
    var title = getDate()
    var text = "am " + getDate()
  } else {
    if (data[2] === "") {
      var title = data[1]
      var text = "am " + data[1]
    } else {
      var title = data[1] - data[2]
      var text = "vom " + data[1] + " bis zum " + data[2]
    }
  }

  if (!(data[3] === "")) {
    var text = text + " aufgrund von " + data[3]
  }

  body.replaceText('{{timereason}}', text);
  body.replaceText('{{date}}', getDate());

  sendMail(doc.getUrl())
}

function getFormData() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lrow = sh.getLastRow();

  var singleDay = sh.getRange(lrow, 2).getValue()
  var startDate = sh.getRange(lrow, 3).getDisplayValue()
  var endDate = sh.getRange(lrow, 4).getDisplayValue()
  var reason = sh.getRange(lrow, 5).getValue()
  
  return [singleDay, startDate, endDate, reason]
}


function getDate() {
  var d = new Date();
  var dd = String(d.getDate());
  var mm = String(d.getMonth() + 1)
  var yyyy = String(d.getFullYear());

  return dd + '/' + mm + '/' + yyyy;
}

function sendMail(docURL) {
  var recipent = EMAIL_RECIPENT
  var subject = "Krankmeldung vom " + getDate()
  var message = "Die folgende Krankmeldung ist zum ausdrucken bereit:\n\n" + docURL 
  MailApp.sendEmail(recipent, subject, message);
}