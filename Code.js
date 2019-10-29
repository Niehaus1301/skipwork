// Config vars
TEMPLATE_DOC_ID = "1syTKKGw9IQc-7uSxCjHfLz0xFTXkt0v8zDub_zd7y84"
FOLDER_ID = "1_ySK_tp-ogWuMCOwWI9t_tOQcKlgL4x2"
EMAIL_RECIPENT = "niehaus.1301@gmail.com"
CLIENT_ID = '372488105238-v3il767jkqppgk61948mdtuffidl4oi7.apps.googleusercontent.com'
CLIENT_SECRET = 'W1q4_2Ui6TORy_s9s8yPOqHv'

function showURL() {
  var cpService = getCloudPrintService();
  if (!cpService.hasAccess()) {
    Logger.log(cpService.getAuthorizationUrl());
  }
}

function getCloudPrintService() {
  return OAuth2.createService('print')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('https://www.googleapis.com/auth/cloudprint')
    .setParam('login_hint', Session.getActiveUser().getEmail())
    .setParam('access_type', 'offline')
    .setParam('approval_prompt', 'force');
}

function authCallback(request) {
  var isAuthorized = getCloudPrintService().handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('You can now use Google Cloud Print from Apps Script.');
  } else {
    return HtmlService.createHtmlOutput('Cloud Print Error: Access Denied');
  }
}

function getPrinterList() {

  var response = UrlFetchApp.fetch('https://www.google.com/cloudprint/search', {
    headers: {
      Authorization: 'Bearer ' + getCloudPrintService().getAccessToken()
    },
    muteHttpExceptions: true
  }).getContentText();

  var printers = JSON.parse(response).printers;

  for (var p in printers) {
    Logger.log("%s %s %s", printers[p].id, printers[p].name, printers[p].description);
  }
}


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
  var copy = template.makeCopy(title, destFolder).getId();
  var doc = DocumentApp.openById(copy)
  var body = DocumentApp.openById(copy).getBody()

  // Replace placeholders with data
  body.replaceText('{{date}}', today);
  body.replaceText('{{timereason}}', text);

  // Send confirmation email
  sendMail(doc.getUrl(), plain)
  
  // Print doc
  // printGoogleDocument(copy, '1581be5e-30d3-65a8-fb9b-fd9d675e3097', title)
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

function printGoogleDocument(docID, printerID, docName) {

  var ticket = {
    version: "1.0",
    print: {
      color: {
        type: "STANDARD_COLOR",
        vendor_id: "Color"
      },
      duplex: {
        type: "NO_DUPLEX"
      }
    }
  };

  var payload = {
    "printerid" : printerID,
    "title"     : docName,
    "content"   : DriveApp.getFileById(docID).getBlob(),
    "contentType": "application/pdf",
    "ticket"    : JSON.stringify(ticket)
  };

  var response = UrlFetchApp.fetch('https://www.google.com/cloudprint/submit', {
    method: "POST",
    payload: payload,
    headers: {
      Authorization: 'Bearer ' + getCloudPrintService().getAccessToken()
    },
    "muteHttpExceptions": true
  });

  response = JSON.parse(response);

  if (response.success) {
    Logger.log("%s", response.message);
  } else {
    Logger.log("Error Code: %s %s", response.errorCode, response.message);
  }
}