/**
 * SD Christmas Tree Removal - booking backend (Google Apps Script)
 *
 * What it does:
 *  - Saves every booking to a Google Sheet called "Tree Pickup Bookings"
 *  - Puts every booking on a Google Calendar called "Tree Pickups" (with the address)
 *  - Emails you each time someone books
 *  - Tells the website which times are already taken so nobody double-books
 */

const CALENDAR_NAME = 'Tree Pickups';
const SHEET_NAME = 'Tree Pickup Bookings';
const TIMEZONE = 'America/Los_Angeles';
const MINUTES_PER_PICKUP = 90;

// The days the website offers, matched to real dates
const DAYS = {
  'Sat, Dec 26': '2026-12-26',
  'Sun, Dec 27': '2026-12-27',
  'Mon, Dec 28': '2026-12-28',
  'Tue, Dec 29': '2026-12-29'
};
const TIMES = ['8:00 AM', '9:30 AM', '11:00 AM', '12:30 PM', '2:00 PM', '3:30 PM'];

const HEADERS = ['Booked at', 'Date', 'Time', 'Service', 'Big tree', 'Needle cleanup',
                 'Total', 'Name', 'Phone', 'Email', 'Address', 'Notes'];

// The website asks: "which times are taken?"
function doGet() {
  return json_({ taken: takenSlots_() });
}

// The website sends a new booking
function doPost(e) {
  const p = e.parameter || {};
  if (p.website) return json_({ ok: true }); // spam bot trap

  if (!DAYS[p.date] || TIMES.indexOf(p.time) === -1 || !p.name || !p.phone || !p.address) {
    return json_({ ok: false, error: 'missing' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const key = p.date + ' ' + p.time;
    if (takenSlots_().indexOf(key) !== -1) {
      return json_({ ok: false, error: 'taken', taken: takenSlots_() });
    }

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(), p.date, p.time, p.service || '', p.bigTree || 'No', p.needles || 'No',
      p.total || '', p.name, p.phone, p.email || '', p.address, p.notes || ''
    ]);

    const start = Utilities.parseDate(DAYS[p.date] + ' ' + to24h_(p.time), TIMEZONE, 'yyyy-MM-dd HH:mm');
    const end = new Date(start.getTime() + MINUTES_PER_PICKUP * 60000);
    const details = [
      'Service: ' + (p.service || ''),
      'Big tree: ' + (p.bigTree || 'No'),
      'Needle cleanup: ' + (p.needles || 'No'),
      'Total: ' + (p.total || ''),
      'Phone: ' + p.phone,
      'Email: ' + (p.email || ''),
      'Notes: ' + (p.notes || '')
    ].join('\n');
    getCalendar_().createEvent('🎄 ' + p.name + ' · ' + (p.total || ''), start, end,
      { location: p.address, description: details });

    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      '🎄 New pickup: ' + p.date + ' at ' + p.time,
      'Name: ' + p.name + '\nAddress: ' + p.address + '\n' + details);

    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Run this once by hand to set everything up and give permissions
function setup() {
  getSheet_();
  getCalendar_();
  Logger.log('All set! Sheet: ' + getSheet_().getParent().getUrl());
}

function takenSlots_() {
  const values = getSheet_().getDataRange().getDisplayValues().slice(1);
  return values.filter(r => r[1] && r[2]).map(r => r[1] + ' ' + r[2]);
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    props.setProperty('SHEET_ID', ss.getId());
    const sh = ss.getSheets()[0];
    sh.getRange('A:L').setNumberFormat('@'); // keep everything as plain text
    sh.getRange('A1:L1').setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return ss.getSheets()[0];
}

function getCalendar_() {
  const found = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  return found.length ? found[0] : CalendarApp.createCalendar(CALENDAR_NAME, { color: CalendarApp.Color.GREEN });
}

function to24h_(t) {
  const m = t.match(/(\d+):(\d+)\s*(AM|PM)/);
  let h = Number(m[1]);
  if (m[3] === 'PM' && h !== 12) h += 12;
  if (m[3] === 'AM' && h === 12) h = 0;
  return (h < 10 ? '0' : '') + h + ':' + m[2];
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
