/*
  Day Drooler RSVPs: a Google Apps Script attached to the band's RSVP Sheet.

  - doPost:      the site sends an RSVP here; it becomes a row, and a confirmation email goes out
  - doGet:       the site asks for "going" counts per show (counts only, never names)
  - dailyEmails: runs every morning; sends day-before reminders and retries confirmations
                 that hit Gmail's daily sending limit
  - setup:       run once by hand; makes the RSVPs tab and the daily trigger

  Setup steps are in README.md next to this file.
*/

/* ===== EMAIL WORDING: edit freely =====
   These fill in per person and show:
   {first} first name   {venue} venue   {city} city   {day} e.g. Sunday, December 13   {time} e.g. 8 PM
   After editing: Deploy → Manage deployments → pencil → New version → Deploy.            */
const COPY = {
  confirmationSubject: "You're on the list: Day Drooler at {venue}",
  confirmationLead:    "You're on the list for Day Drooler at {venue}.",
  confirmationClose:   "We'll send a reminder the day before. See you there.",

  reminderSubject:     "Tomorrow: Day Drooler at {venue}",
  reminderLead:        "Day Drooler plays {venue} tomorrow.",
  reminderClose:       "See you there.",

  merch:               "CDs and tees at every show.",
  signoff:             "Day Drooler",

  calendarTitle:       "Day Drooler, live at {venue}",   // the event name in "Add to calendar"
};
/* ====================================== */

const TAB = "RSVPs";
const FROM_NAME = "Day Drooler";
const SITE = "https://daydrooler.com";
const TZ = "America/New_York";
const REMINDER_HOUR = 10; // daily run, local time
const SHOW_HOURS = 3;     // length of the calendar event when the time is known

// New columns go at the end so a Sheet made by an older version keeps lining up
const HEADERS = ["Submitted", "Show", "Show date", "Venue", "City", "Time",
                 "Name", "Email", "Email list", "Confirmation", "Reminder", "Address"];
const COL = Object.fromEntries(HEADERS.map((h, i) => [h, i]));

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TAB) || ss.insertSheet(TAB);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  sh.setFrozenRows(1);
  // Keep "2026-12-13" and "8 PM" as text; otherwise Sheets turns them into date/time values
  for (const c of ["Show date", "Time"]) sh.getRange(2, COL[c] + 1, sh.getMaxRows() - 1, 1).setNumberFormat("@");

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "dailyEmails")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("dailyEmails").timeBased().everyDays(1).atHour(REMINDER_HOUR).inTimezone(TZ).create();
}

function doGet() {
  const counts = {};
  for (const r of rows_()) counts[r[COL["Show"]]] = (counts[r[COL["Show"]]] || 0) + 1;
  return json_({ ok: true, counts });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "bad request" }); }
  if (d.website) return json_({ ok: true }); // hidden field only bots fill in

  const name = clean_(d.name, 80);
  const email = clean_(d.email, 120).toLowerCase();
  const show = {
    key: clean_(d.show, 120), date: clean_(d.date, 10), venue: clean_(d.venue, 80),
    city: clean_(d.city, 80), time: clean_(d.time, 20), address: clean_(d.address, 160),
  };
  if (!name) return json_({ ok: false, error: "Add your name" });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json_({ ok: false, error: "That email doesn't look right" });
  if (!show.key || !/^\d{4}-\d{2}-\d{2}$/.test(show.date)) return json_({ ok: false, error: "bad request" });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let rowNum, already = false;
  try {
    const sh = sheet_();
    const data = rows_();
    const i = data.findIndex(r => r[COL["Show"]] === show.key && r[COL["Email"]] === email);
    const listed = d.list ? "yes" : "no";
    if (i >= 0) {
      // Same person RSVPing twice: update instead of double-counting
      rowNum = i + 2; already = true;
      sh.getRange(rowNum, COL["Name"] + 1).setValue(name);
      sh.getRange(rowNum, COL["Email list"] + 1).setValue(listed);
    } else {
      rowNum = sh.getLastRow() + 1;
      const row = HEADERS.map(() => "");
      row[COL["Submitted"]] = new Date();
      row[COL["Show"]] = show.key; row[COL["Show date"]] = show.date;
      row[COL["Venue"]] = show.venue; row[COL["City"]] = show.city; row[COL["Time"]] = show.time;
      row[COL["Address"]] = show.address;
      row[COL["Name"]] = name; row[COL["Email"]] = email; row[COL["Email list"]] = listed;
      for (const c of ["Show date", "Time"]) sh.getRange(rowNum, COL[c] + 1).setNumberFormat("@");
      sh.getRange(rowNum, 1, 1, HEADERS.length).setValues([row]);
    }
  } finally {
    lock.releaseLock();
  }

  if (!already) trySend_(rowNum, "Confirmation", confirmation_);
  const count = rows_().filter(r => r[COL["Show"]] === show.key).length;
  return json_({ ok: true, already, count });
}

function dailyEmails() {
  const tomorrow = Utilities.formatDate(new Date(Date.now() + 864e5), TZ, "yyyy-MM-dd");
  const today = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
  rows_().forEach((r, i) => {
    const showDate = String(r[COL["Show date"]]);
    if (r[COL["Confirmation"]] !== "sent" && showDate > today) trySend_(i + 2, "Confirmation", confirmation_);
    if (r[COL["Reminder"]] !== "sent" && showDate === tomorrow) trySend_(i + 2, "Reminder", reminder_);
  });
}

// Sends one email for a row and records "sent" or why not in that row's column.
// Goes through band@'s own Gmail (GmailApp): Gmail silently dropped mail sent via MailApp from this new domain.
// Workspace accounts can send about 1,500 a day; anything over waits for tomorrow's run.
function trySend_(rowNum, column, build) {
  const sh = sheet_();
  const r = sh.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0];
  const cell = sh.getRange(rowNum, COL[column] + 1);
  try {
    const m = build(r);
    GmailApp.sendEmail(r[COL["Email"]], m.subject, m.text, { htmlBody: m.html, name: FROM_NAME });
    cell.setValue("sent");
  } catch (err) {
    const limit = /too many times|limit/i.test(String(err));
    cell.setValue(limit ? "waiting (daily limit)" : "failed: " + String(err).slice(0, 80));
  }
}

function confirmation_(r) {
  const s = showOf_(r);
  return letter_(r, s, COPY.confirmationSubject, COPY.confirmationLead, COPY.confirmationClose);
}

function reminder_(r) {
  const s = showOf_(r);
  return letter_(r, s, COPY.reminderSubject, COPY.reminderLead, COPY.reminderClose);
}

function letter_(r, s, subjectT, leadT, closeT) {
  const first = String(r[COL["Name"]]).trim().split(/\s+/)[0];
  const fill = t => t.replace(/\{(\w+)\}/g, (all, k) =>
    ({ first, venue: s.venue, city: s.city, day: s.pretty, time: s.time })[k] ?? all);
  const subject = fill(subjectT), lead = fill(leadT), close = fill(closeT);
  const when = s.pretty + (s.time ? ", " + s.time : "");
  const where = s.address || s.city;
  const map = "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(s.venue + ", " + where);
  const gcal = googleCalendarLink_(s);

  const text = [
    "Hi " + first + ",", "", lead, "",
    s.venue, where, when, "",
    "Map: " + map,
    "Add to calendar: " + gcal, "",
    COPY.merch, close, "", COPY.signoff, SITE,
  ].join("\n");

  const p = 'style="margin:0 0 14px"';
  const html =
    '<div style="font-family:Courier New,monospace;font-size:15px;line-height:1.5;color:#111">' +
    "<p " + p + ">Hi " + esc_(first) + ",</p>" +
    "<p " + p + ">" + esc_(lead) + "</p>" +
    '<p style="margin:0 0 14px;font-size:17px"><b>' + esc_(s.venue) + "</b><br>" +
      '<a href="' + map + '" style="color:#111">' + esc_(where) + "</a><br>" + esc_(when) + "</p>" +
    "<p " + p + '><a href="' + gcal + '">Add to calendar</a></p>' +
    "<p " + p + ">" + esc_(COPY.merch) + "<br>" + esc_(close) + "</p>" +
    '<p style="margin:0">' + esc_(COPY.signoff) + '<br><a href="' + SITE + '">daydrooler.com</a></p></div>';
  return { subject, html, text };
}

function showOf_(r) {
  const date = cellText_(r[COL["Show date"]], "yyyy-MM-dd");
  const [y, m, d] = date.split("-").map(Number);
  const pretty = Utilities.formatDate(new Date(Date.UTC(y, m - 1, d, 17)), TZ, "EEEE, MMMM d");
  const time = cellText_(r[COL["Time"]], "h:mm a").replace(":00 ", " ");
  return { date, pretty, venue: String(r[COL["Venue"]]), city: String(r[COL["City"]]),
           time, address: String(r[COL["Address"]] || "") };
}

// Rows saved before a column was set to text may hold a Sheets date/time instead of the text the site sent
function cellText_(v, pattern) {
  if (Object.prototype.toString.call(v) !== "[object Date]") return String(v);
  return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), pattern);
}

// "8 PM" / "10:30pm" -> {h: 20, m: 0}; anything else -> null (the event becomes all-day)
function parseTime_(time) {
  const t = /^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)$/i.exec(String(time).trim());
  if (!t) return null;
  return { h: Number(t[1]) % 12 + (t[3].toUpperCase() === "PM" ? 12 : 0), m: Number(t[2] || 0) };
}

// Start and end as real moments in time (handles EST/EDT), or null when there's no time
function showTimes_(s) {
  const t = parseTime_(s.time);
  if (!t) return null;
  const start = Utilities.parseDate(s.date + " " + pad_(t.h) + ":" + pad_(t.m), TZ, "yyyy-MM-dd HH:mm");
  return { start, end: new Date(start.getTime() + SHOW_HOURS * 36e5) };
}

// Google Calendar "add event" link
function googleCalendarLink_(s) {
  const times = showTimes_(s);
  const dates = times
    ? Utilities.formatDate(times.start, "UTC", "yyyyMMdd'T'HHmmss'Z'") + "/" + Utilities.formatDate(times.end, "UTC", "yyyyMMdd'T'HHmmss'Z'")
    : s.date.replace(/-/g, "") + "/" + nextDay_(s.date.replace(/-/g, ""));
  return "https://calendar.google.com/calendar/render?action=TEMPLATE" +
    "&text=" + encodeURIComponent(COPY.calendarTitle.replace(/\{venue\}/g, s.venue)) +
    "&dates=" + dates +
    "&location=" + encodeURIComponent(s.venue + ", " + (s.address || s.city)) +
    "&details=" + encodeURIComponent(SITE);
}

function nextDay_(yyyymmdd) {
  const d = new Date(Date.UTC(+yyyymmdd.slice(0, 4), +yyyymmdd.slice(4, 6) - 1, +yyyymmdd.slice(6, 8) + 1));
  return d.toISOString().slice(0, 10).replace(/-/g, "");
}

function pad_(n) { return String(n).padStart(2, "0"); }

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB);
  if (!sh) throw new Error('No "' + TAB + '" tab. Run setup() first.');
  return sh;
}

function rows_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getValues() : [];
}

function clean_(v, max) { return String(v == null ? "" : v).trim().slice(0, max); }
function esc_(v) { return String(v).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
