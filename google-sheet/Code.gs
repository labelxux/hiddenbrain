// The Hidden Brain — Google Sheet collector
// Paste into Extensions → Apps Script of your Google Sheet, then Deploy → New deployment → Web app.

const SHEET_NAME = 'Submissions';
const SEND_EMAIL = true;               // email the blueprint to the user
const NOTIFY_ME = '';                  // optional: your email to get a copy of every signup
const FROM_NAME = 'The Hidden Brain';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const email = String(data.email || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json({ ok: false, error: 'invalid email' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) sh.appendRow(['Date', 'Email', 'Page', 'Blueprint']);
    sh.appendRow([new Date(), email, data.page || '', String(data.blueprint || '').slice(0, 45000)]);

    if (SEND_EMAIL) {
      MailApp.sendEmail({
        to: email,
        name: FROM_NAME,
        subject: 'Your AI Work Brain Blueprint',
        body: 'Here is the blueprint you built.\n\n' + (data.blueprint || '') + '\n\n—\n' + (data.page ? 'Come back any time: ' + data.page : ''),
      });
    }
    if (NOTIFY_ME) MailApp.sendEmail(NOTIFY_ME, 'New Hidden Brain signup: ' + email, data.blueprint || '');

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
