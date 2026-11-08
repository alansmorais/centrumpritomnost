/**
 * Centrum Přítomnost — booking backend
 * Install in an Apps Script project owned by centrumpritomnost@gmail.com.
 * Run setupCentrum() once before deploying the web app.
 */

const CONFIG = {
  SPREADSHEET_ID: '1h1uEVdEO_8-4c6YwMMuHA3t9SvqY3Nqofc8nSPiPY8k',
  ADMIN_EMAIL: 'centrumpritomnost@gmail.com',
  CALENDAR_ID: 'centrumpritomnost@gmail.com',
  TIME_ZONE: 'Europe/Prague',
  REQUEST_SHEET: 'Requests',
  SETTINGS_SHEET: 'Settings',
  HOLD_PREFIX: '[PENDING] '
};

const REQUEST_HEADERS = [
  'Request ID', 'Created At', 'Status', 'Category', 'Service',
  'Date', 'Start Time', 'Duration (minutes)', 'Customer Name',
  'Customer Email', 'Customer Phone', 'Message', 'Calendar Event ID'
];

function setupCentrum() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let requests = ss.getSheetByName(CONFIG.REQUEST_SHEET);
  if (!requests) requests = ss.insertSheet(CONFIG.REQUEST_SHEET);
  if (requests.getLastRow() === 0) {
    requests.getRange(1, 1, 1, REQUEST_HEADERS.length).setValues([REQUEST_HEADERS]);
    requests.setFrozenRows(1);
    requests.getRange(1, 1, 1, REQUEST_HEADERS.length).setFontWeight('bold');
    requests.autoResizeColumns(1, REQUEST_HEADERS.length);
  }

  let settings = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
  if (!settings) settings = ss.insertSheet(CONFIG.SETTINGS_SHEET);
  if (settings.getLastRow() === 0) {
    settings.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
    settings.getRange(2, 1, 4, 2).setValues([
      ['Admin email', CONFIG.ADMIN_EMAIL],
      ['Calendar ID', CONFIG.CALENDAR_ID],
      ['Time zone', CONFIG.TIME_ZONE],
      ['Booking policy', 'Requests require manual approval. Pending calendar events temporarily block the requested time.']
    ]);
    settings.setFrozenRows(1);
    settings.getRange(1, 1, 1, 2).setFontWeight('bold');
    settings.autoResizeColumns(1, 2);
  }

  // Confirm that the owner-authorized account can access its primary calendar.
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) throw new Error('Primary Google Calendar is not accessible to the account running setupCentrum().');
  return 'Setup completed. Spreadsheet and primary calendar are accessible.';
}

function doGet(e) {
  try {
    const action = getQueryParam_(e, 'action');
    if (action === 'availability') {
      const dateString = getQueryParam_(e, 'date');
      return jsonOutput_(getAvailability_(dateString));
    }
    return jsonOutput_({ ok: true, service: 'Centrum Přítomnost booking backend', status: 'ready' });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message || 'Unable to process request.' });
  }
}

function getQueryParam_(e, name) {
  if (e && e.parameter && e.parameter[name] != null && e.parameter[name] !== '') {
    return String(e.parameter[name]).trim();
  }
  if (e && e.parameters && e.parameters[name] && e.parameters[name].length) {
    return String(e.parameters[name][0]).trim();
  }
  const query = e && e.queryString ? String(e.queryString) : '';
  for (const part of query.split('&')) {
    if (!part) continue;
    const pieces = part.split('=');
    const key = decodeURIComponent((pieces.shift() || '').replace(/\+/g, ' '));
    if (key === name) {
      return decodeURIComponent(pieces.join('=').replace(/\+/g, ' ')).trim();
    }
  }
  return '';
}

function doPost(e) {
  try {
    const data = (e && e.parameter) ? e.parameter : {};
    const result = createRequest_(data);
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message || 'Unable to process request.' });
  }
}

function getAvailability_(dateString) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateString || '')) {
    return { ok: false, error: 'Invalid date. Use YYYY-MM-DD.' };
  }
  const start = new Date(dateString + 'T00:00:00');
  const end = new Date(dateString + 'T23:59:59');
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) throw new Error('Calendar is not accessible.');
  const events = calendar.getEvents(start, end).map(event => ({
    start: Utilities.formatDate(event.getStartTime(), CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ss"),
    end: Utilities.formatDate(event.getEndTime(), CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ss"),
    title: event.getTitle()
  }));
  return { ok: true, date: dateString, busy: events, timeZone: CONFIG.TIME_ZONE };
}

function createRequest_(data) {
  const required = ['category', 'service', 'date', 'startTime', 'duration', 'name', 'email', 'phone'];
  required.forEach(key => {
    if (!String(data[key] || '').trim()) throw new Error('Please complete all required fields.');
  });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data.date)) throw new Error('Invalid date.');
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(data.startTime)) throw new Error('Invalid start time.');
  const duration = Number(data.duration);
  if (!Number.isInteger(duration) || duration < 30 || duration > 720) throw new Error('Invalid duration.');
  const email = String(data.email).trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Invalid email address.');

  const start = new Date(data.date + 'T' + data.startTime + ':00');
  const end = new Date(start.getTime() + duration * 60000);
  if (start <= new Date()) throw new Error('Please choose a future time.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    if (!calendar) throw new Error('Calendar is not accessible.');
    const conflicts = calendar.getEvents(start, end);
    if (conflicts.length) throw new Error('This time is no longer available. Please choose another time.');

    const id = Utilities.getUuid();
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.REQUEST_SHEET);
    if (!sheet) throw new Error('Run setupCentrum() before accepting requests.');

    const title = CONFIG.HOLD_PREFIX + data.category + ' — ' + data.service + ' — ' + data.name;
    const description = [
      'Status: PENDING — manual approval required',
      'Request ID: ' + id,
      'Category: ' + data.category,
      'Service: ' + data.service,
      'Customer: ' + data.name,
      'Email: ' + email,
      'Phone: ' + data.phone,
      'Message: ' + (data.message || '')
    ].join('\n');

    const event = calendar.createEvent(title, start, end, { description: description });
    sheet.appendRow([
      id, new Date(), 'PENDING', data.category, data.service,
      data.date, data.startTime, duration, data.name, email,
      data.phone, data.message || '', event.getId()
    ]);

    const formattedStart = Utilities.formatDate(start, CONFIG.TIME_ZONE, 'dd.MM.yyyy HH:mm');
    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: 'New booking request — Centrum Přítomnost',
      body: 'A new booking request is waiting for your review.\n\n' +
        'Category: ' + data.category + '\nService: ' + data.service +
        '\nDate and time: ' + formattedStart + '\nDuration: ' + duration + ' minutes' +
        '\nCustomer: ' + data.name + '\nEmail: ' + email +
        '\nPhone: ' + data.phone + '\nMessage: ' + (data.message || '') +
        '\n\nThe calendar slot is temporarily blocked as PENDING. Contact the customer to confirm or propose another time.'
    });
    MailApp.sendEmail({
      to: email,
      subject: 'We received your request — Centrum Přítomnost',
      body: 'Hello ' + data.name + ',\n\nWe have received your booking request for ' +
        data.service + ' on ' + formattedStart + '.\n\nThis is not yet a confirmed reservation. ' +
        'We will contact you after reviewing availability.\n\nCentrum Přítomnost'
    });
    return { ok: true, requestId: id, status: 'PENDING' };
  } finally {
    lock.releaseLock();
  }
}

// Manual one-time test. Run this function from the Apps Script editor only when ready.
function runBookingTestOnce() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('CENTRUM_BOOKING_TEST_DONE') === 'yes') {
    throw new Error('The booking test has already been submitted. Remove the test event/request manually before retrying.');
  }
  const result = createRequest_({
    category: 'Rental of space',
    service: 'Test booking',
    date: '2026-11-09',
    startTime: '10:00',
    duration: '30',
    name: 'Alan Piquê Moraes',
    email: 'Alanpkmorais@gmail.com',
    phone: '12345678901',
    message: 'One-time technical test booking. Please disregard after verification.'
  });
  if (result && result.ok) props.setProperty('CENTRUM_BOOKING_TEST_DONE', 'yes');
  return result;
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
