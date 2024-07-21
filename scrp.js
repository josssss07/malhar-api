var ID_LENGTH = 20;

// Generates a random UID
function generateUID () {
  var ALPHABET = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
  var rtn = '';
  for (var i = 0; i < ID_LENGTH; i++) {
    rtn += ALPHABET.charAt(Math.floor(Math.random() * ALPHABET.length));
  }
  return rtn;
}

function autoResponder(e) {
  console.log(e.range.rowStart) 

  // Get Active Sheet
  var ss = SpreadsheetApp.getActiveSheet();

  var sheetname = ss.getName();

  var responseId = generateUID()

  // Append response ID and Attendance status to row
  var formrow = e.range.rowStart
  ss.getRange(formrow, 5).setValue(responseId);
  ss.getRange(formrow, 6).setValue("FALSE");
  console.log(e.namedValues)

  var name = e.namedValues.Name
  var email = e.namedValues['Email Address'][0]
  console.log(e.namedValues['Email Address'][0])
  var adminEmail = "valerontoscano@gmail.com"

  MailApp.sendEmail({
    to: email,
    subject: "Entry pass for " + sheetname + " event",
    htmlBody: '<html><body><div>Thank you for registering for the event. Here is the QR code that will grant you entry for the event.<br />Please be present at the venue at least 15 minutes&nbsp;prior to the start of the event.<br /><br /><table width="100%"><tr><td align="center"><img src="cid:qrcode" style="width:256px" align="middle"></td></tr></table><br/><br/>Regards,<br />ECC Team</span></div></body></html>',
  });
}