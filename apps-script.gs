// Google Apps Script: logs each valid sign-up as a new row in the sign-up spreadsheet.
// - One free roll per person: a sign-up whose phone number OR email is already in the sheet is skipped (first one is kept).
// - Obviously fake phone numbers and throwaway email addresses are rejected.
// - Skipped and rejected sign-ups are written to a "Rejected" tab so nothing disappears silently.
// - Staff tick the "Redeemed" box when a customer collects the free roll.
var SHEET_ID = '1f5dTP5_0OaD9Bf2U-zNYzrZNGSwH16IxslAEKVwnB14';
var HEADERS = ['#', 'Signed up (ET)', 'First name', 'Last name', 'Phone', 'Email', 'Location', 'Redeemed'];
var REJ_HEADERS = ['Time (ET)', 'Reason', 'First name', 'Last name', 'Phone', 'Email'];
var DISPOSABLE = ['mailinator.com','guerrillamail.com','guerrillamail.net','sharklasers.com','10minutemail.com',
  'tempmail.com','temp-mail.org','yopmail.com','trashmail.com','getnada.com','dispostable.com','maildrop.cc',
  'throwawaymail.com','fakeinbox.com','mohmal.com','tempail.com','emailondeck.com','mintemail.com','spamgourmet.com'];

// ---------- helpers ----------

// Digits only, ignoring a leading US country code: "(732) 555-1234" and "+1 732 555 1234" match.
function normPhone_(v) {
  var d = String(v || '').replace(/\D/g, '');
  if (d.length === 11 && d.charAt(0) === '1') d = d.substring(1);
  return d;
}

// Lowercase; Gmail ignores dots and +tags; other providers ignore +tags.
function normEmail_(v) {
  var e = String(v || '').trim().toLowerCase();
  var at = e.lastIndexOf('@');
  if (at < 1) return e;
  var local = e.substring(0, at).split('+')[0];
  var domain = e.substring(at + 1);
  if (domain === 'googlemail.com') domain = 'gmail.com';
  if (domain === 'gmail.com') local = local.replace(/\./g, '');
  return local + '@' + domain;
}

// Returns a reason if the phone number cannot be real, otherwise ''.
function badPhoneReason_(d) {
  if (d.length !== 10) return 'Phone is not a 10-digit US number';
  if (/^[01]/.test(d)) return 'Invalid area code';
  if (/^\d{3}[01]/.test(d)) return 'Invalid exchange';
  if (/^(\d)\1{9}$/.test(d)) return 'Repeated digits';
  if ('01234567890123456789'.indexOf(d) >= 0 || '98765432109876543210'.indexOf(d) >= 0) return 'Sequential digits';
  if (/^(\d{2})\1{4}$/.test(d) || /^(\d{5})\1$/.test(d)) return 'Repeating pattern';
  if (d.substring(3, 6) === '555' && d.substring(6, 8) === '01') return 'Fictional 555-01xx number';
  return '';
}

function badEmailReason_(e) {
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]{2,}$/.test(e)) return 'Email looks invalid';
  var domain = e.substring(e.lastIndexOf('@') + 1).toLowerCase();
  if (DISPOSABLE.indexOf(domain) >= 0) return 'Throwaway email address';
  return '';
}

function now_() {
  return Utilities.formatDate(new Date(), 'America/New_York', 'MMM d, yyyy h:mm a');
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Sign Ups') || ss.insertSheet('Sign Ups');
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.getRange('E:E').setNumberFormat('@'); // keep phone numbers as text
  } else if (sh.getRange(1, 8).getValue() === '') {
    sh.getRange(1, 8).setValue('Redeemed').setFontWeight('bold'); // sheets created before this column existed
  }
  return sh;
}

function getRejectedSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Rejected') || ss.insertSheet('Rejected');
  if (sh.getLastRow() === 0) {
    sh.appendRow(REJ_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, REJ_HEADERS.length).setFontWeight('bold');
    sh.getRange('E:E').setNumberFormat('@');
  }
  return sh;
}

function logRejected_(reason, p) {
  var rs = getRejectedSheet_();
  rs.appendRow([now_(), reason, p.first_name || '', p.last_name || '', '', p.email || '']);
  rs.getRange(rs.getLastRow(), 5).setNumberFormat('@').setValue(p.phone || '');
}

// ---------- web app ----------

function doGet() {
  return ContentService.createTextOutput('Kati Roll sign-up logger is running.');
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.botcheck) return ContentService.createTextOutput('ignored');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var phone = normPhone_(p.phone);
    var email = normEmail_(p.email);

    var reason = badPhoneReason_(phone) || badEmailReason_(email);
    if (reason) { logRejected_(reason, p); return ContentService.createTextOutput('rejected'); }

    var sh = getSheet_();
    var last = sh.getLastRow();
    if (last > 1) {
      var rows = sh.getRange(2, 5, last - 1, 2).getValues(); // phone, email columns
      for (var i = 0; i < rows.length; i++) {
        if (normPhone_(rows[i][0]) === phone) { logRejected_('Duplicate phone number', p); return ContentService.createTextOutput('duplicate'); }
        if (normEmail_(rows[i][1]) === email) { logRejected_('Duplicate email', p); return ContentService.createTextOutput('duplicate'); }
      }
    }
    sh.appendRow([last, now_(), p.first_name || '', p.last_name || '', '', p.email || '', p.location || '']);
    var r = sh.getLastRow();
    sh.getRange(r, 5).setNumberFormat('@').setValue(p.phone || '');
    sh.getRange(r, 8).insertCheckboxes();
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

// ---------- run manually from the editor ----------

// Cleans duplicates and fake entries already in the sheet. Keeps the first row for each phone/email,
// moves the rest to the Rejected tab, and renumbers column A.
function cleanUpSheet() {
  var sh = getSheet_();
  var seenPhone = {}, seenEmail = {};
  for (var r = 2; r <= sh.getLastRow(); ) {
    var v = sh.getRange(r, 1, 1, 8).getValues()[0];
    var phone = normPhone_(v[4]), email = normEmail_(v[5]);
    var reason = badPhoneReason_(phone) || badEmailReason_(email);
    if (!reason && seenPhone[phone]) reason = 'Duplicate phone number';
    if (!reason && seenEmail[email]) reason = 'Duplicate email';
    if (reason) {
      logRejected_(reason + ' (cleanup)', { first_name: v[2], last_name: v[3], phone: v[4], email: v[5] });
      sh.deleteRow(r);
    } else {
      seenPhone[phone] = true; seenEmail[email] = true; r++;
    }
  }
  var n = sh.getLastRow() - 1;
  if (n > 0) {
    var nums = [];
    for (var k = 1; k <= n; k++) nums.push([k]);
    sh.getRange(2, 1, n, 1).setValues(nums);
    sh.getRange(2, 8, n, 1).insertCheckboxes();
  }
}

// Sends a valid-looking test sign-up. Delete the row afterwards.
function testRow() {
  doPost({ parameter: { first_name: 'Test', last_name: 'Row', phone: '7325559876', email: 'test.row@example.com', location: 'Kendall Park' } });
}
