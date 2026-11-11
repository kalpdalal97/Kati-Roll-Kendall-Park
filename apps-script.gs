// Google Apps Script: logs each sign-up as a new row in this spreadsheet.
// Setup steps are in README.md.
function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.botcheck) return ContentService.createTextOutput('ignored');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
