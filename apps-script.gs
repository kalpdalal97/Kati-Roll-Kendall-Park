// Google Apps Script: logs each sign-up as a new row in the sign-up spreadsheet.
// Works whether or not the script was started from the sheet's Extensions menu.
// A sign-up with a phone number that is already in the sheet is skipped (first one is kept).
var SHEET_ID = '1f5dTP5_0OaD9Bf2U-zNYzrZNGSwH16IxslAEKVwnB14';
var HEADERS = ['#', 'Signed up (ET)', 'First name', 'Last name', 'Phone', 'Email', 'Location'];

// Reduces a phone number to digits only, ignoring a leading US country code,
// so "(732) 555-1234", "732-555-1234" and "+1 732 555 1234" all match.
function normPhone_(v) {
  var d = String(v || '').replace(/\D/g, '');
  if (d.length === 11 && d.charAt(0) === '1') d = d.substring(1);
  return d;
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Sign Ups') || ss.insertSheet('Sign Ups');
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.getRange('E:E').setNumberFormat('@'); // keep phone numbers as text
  }
  return sh;
}

function doGet() {
  return ContentService.createTextOutput('Kati Roll sign-up logger is running.');
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.botcheck) return ContentService.createTextOutput('ignored');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet_();
    var phone = normPhone_(p.phone);
    var last = sh.getLastRow();
    if (phone && last > 1) {
      var existing = sh.getRange(2, 5, last - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (normPhone_(existing[i][0]) === phone) {
          return ContentService.createTextOutput('duplicate'); // keep the first sign-up only
        }
      }
    }
    var n = last; // header is row 1, so this is the next entry number
    var when = Utilities.formatDate(new Date(), 'America/New_York', 'MMM d, yyyy h:mm a');
    sh.appendRow([n, when, p.first_name || '', p.last_name || '', p.phone || '', p.email || '', p.location || '']);
    sh.getRange(sh.getLastRow(), 5).setNumberFormat('@').setValue(p.phone || '');
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

// Run once from the editor to clean duplicates that are already in the sheet.
// Keeps the first row for each phone number, deletes later ones, and renumbers column A.
function removeDuplicates() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 3) return;
  var seen = {};
  for (var r = 2; r <= sh.getLastRow(); ) {
    var ph = normPhone_(sh.getRange(r, 5).getValue());
    if (ph && seen[ph]) { sh.deleteRow(r); }
    else { if (ph) seen[ph] = true; r++; }
  }
  var rows = sh.getLastRow() - 1;
  if (rows > 0) {
    var nums = [];
    for (var k = 1; k <= rows; k++) nums.push([k]);
    sh.getRange(2, 1, rows, 1).setValues(nums);
  }
}

// Optional: run once from the editor to test that the script can reach the sheet.
function testRow() {
  doPost({ parameter: { first_name: 'Test', last_name: 'Row', phone: '5555555555', email: 'test@example.com', location: 'Kendall Park' } });
}
