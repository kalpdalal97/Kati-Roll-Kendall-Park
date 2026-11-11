// Google Apps Script: logs each sign-up as a new row in the sign-up spreadsheet.
// Works whether or not the script was started from the sheet's Extensions menu.
var SHEET_ID = '1f5dTP5_0OaD9Bf2U-zNYzrZNGSwH16IxslAEKVwnB14';

function doGet() {
  return ContentService.createTextOutput('Kati Roll sign-up logger is running.');
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.botcheck) return ContentService.createTextOutput('ignored');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('Sign Ups') || ss.insertSheet('Sign Ups');
    if (sh.getLastRow() === 0) {
      sh.appendRow(['#', 'Signed up (ET)', 'First name', 'Last name', 'Phone', 'Email', 'Location']);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, 7).setFontWeight('bold');
      sh.getRange('E:E').setNumberFormat('@'); // keep phone numbers as text
    }
    var n = sh.getLastRow(); // header is row 1, so this is the next entry number
    var when = Utilities.formatDate(new Date(), 'America/New_York', 'MMM d, yyyy h:mm a');
    sh.appendRow([n, when, p.first_name || '', p.last_name || '', p.phone || '', p.email || '', p.location || '']);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

// Optional: run this once from the editor to test that the script can reach the sheet.
function testRow() {
  doPost({ parameter: { first_name: 'Test', last_name: 'Row', phone: '5555555555', email: 'test@example.com', location: 'Kendall Park' } });
}
