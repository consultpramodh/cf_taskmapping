/**
 * Contractor Weekly Check — Google Apps Script V8
 * Reads Gmail, Calendar and Striven; writes to one tracker tab ("Contractor Tracker").
 * Searches Kawartha first, then Steve. Steps 1-2 are read-only; only step 3
 * ("Create missing contractor tasks") writes to Striven, after you confirm.
 * Each tracker row is an EMAIL, not a confirmed job. Mixed-job emails and PDFs need manual review.
 * Results go to ONE tab, "Contractor Tracker": each run adds only emails not
 * already listed. Paste into a spreadsheet-bound Apps Script project, reload the spreadsheet,
 * and use the "Contractor check" menu (or run contractorWeeklyCheck).
 */
const CWC = Object.freeze({
  timezone: 'America/Toronto',
  sheetName: 'Contractor Tracker', // single tab reused on every run
  // Link for Striven task IDs/titles in the tracker. Open any task in Striven, copy
  // the address bar URL and replace the task number with {id}. Leave '' for no links.
  strivenTaskUrl: 'https://classicfireplace.striven.com/Tasks/TaskInfo.aspx?TaskID={id}',
  createStatus: 'Done',   // contractor emailed that the work is complete, so create the task as Done
  // Optional fixed IDs; when null they are looked up from existing Striven tasks.
  // e.g. {'Custom Work': {typeId: 123, vendorId: 45}}
  strivenIds: {},
  // Task type 111 from Tasks.aspx?TaskTypeIDs=111, preferred when several types match.
  // Each contractor's exampleTaskId (19671 Custom Work, 19677 Wood Installation) is the
  // authoritative source for that contractor's task type and subcontractor vendor.
  strivenHints: {taskIds: [], typeIds: [111]},
  // Updating an existing task: PATCH v2/tasks/{id} (confirmed from Striven's docs).
  // It has no Status field. Marking Done: POST v2/tasks/{id}/mark-done (Striven docs), whose
  // body is the task type's "mark done" custom fields [{Id, Value, ValueText}]; Custom Work
  // has none (task 19671: markDoneCustomFields null), so an empty list is sent.
  strivenStatusUpdate: {method: 'post', path: 'v2/tasks/{id}/mark-done', body: []},
  lookbackDays: 7,            // rolling period ending at execution time
  calendarWindowDays: 30,
  pageSize: 100,
  maxThreadsPerContractor: 1000,
  maxRuntimeMs: 300000,       // Apps Script hard limit is 360 s; leave room to write
  headers: ['Contractor','Email date (Toronto)','Subject','From','Email link','Photos found','PDFs found','Other files','Attachment names','SO numbers found','Address hints','Calendar candidates','Calendar dates (Toronto)','Review status','Notes'],
  contractors: [
    {name:'Kawartha / Jamie',
     calendarId:'classicfireplace.ca_2pbtbkjl1g2sos1fspgfq029h0@group.calendar.google.com',
     query:'{"Kawartha Home and Hearth" khhltd@gmail.com} -in:trash -in:spam',
     identity:/kawartha home and hearth|khhltd@gmail\.com/i,
     taskType:'Wood Fireplace Installation', taskTypeMatch:/wood\s*(fireplace\s*)?install/i, exampleTaskId:19676, alsoExampleIds:[19677],
     // Structure copied from reference task 19676: same type, assigned to the same pool
     // ("To Be Assigned"). Title includes street and city, like Steve's.
     assignTo: null, titleAddressParts: 2},   // Striven task this contractor's work belongs on
    {name:'Steve Jensen',
     calendarId:'steve.jensen83@gmail.com',
     query:'{from:steve.jensen83@gmail.com "Steve Jensen"} -in:trash -in:spam',
     identity:/steve\s+jensen|steve\.jensen83@gmail\.com/i,
     taskType:'Custom Work', taskTypeMatch:/custom\s*work/i, exampleTaskId:19671}
  ]
});

/**
 * Contractor check menu. There is no onOpen here, because the project already has
 * one. Call this from your existing onOpen, either as its own menu:
 *     cwcAddContractorMenu_();
 * or as a submenu inside a menu you already build:
 *     ui.createMenu('Tools').addItem(...).addSubMenu(cwcContractorMenu_(ui)).addToUi();
 */
function cwcContractorMenu_(ui) {
  return (ui || SpreadsheetApp.getUi()).createMenu('Contractor check')
    .addItem('1. Run weekly check', 'contractorWeeklyCheck')
    .addItem('2. Verify Striven tasks', 'contractorAuditStrivenCandidates')
    .addItem('3. Create missing contractor tasks', 'contractorCreateMissingTasks')
    .addItem('4. Complete contractor tasks (Done, dates, requested by, files)', 'contractorCompleteTasks')
    .addSeparator()
    .addItem('Run automation now (Steve + Jamie)', 'contractorAutoRun')
    .addItem('Check Striven setup only (no changes)', 'contractorPreflightReport')
    .addItem('Jamie: analyse Wood Fireplace Installation tasks (no changes)', 'contractorJamieReport')
    .addItem('Jamie: fix Requested By and mark all existing tasks Done', 'contractorJamieCloseExisting')
    .addItem('Start schedule (Steve + Jamie)', 'contractorInstallSchedule')
    .addItem('Stop schedule', 'contractorRemoveSchedule');
}
function cwcAddContractorMenu_() { cwcContractorMenu_().addToUi(); }

/**
 * Run guard for THIS script only. LockService.getScriptLock() is shared by every
 * file in the Apps Script project, so another script holding it made this fail
 * with "already running". The marker lives in CacheService (not Script
 * Properties, which are kept for real settings) and expires after 7 minutes.
 */
const CWC_LEASE_KEY = 'CWC_RUNNING_UNTIL';
let cwcDepth_ = 0;     // lets the scheduled run call each step while holding the guard once
let cwcOnly_ = null;  // when set (scheduled run), steps only handle these contractor names
function cwcAcquire_() {
  if (cwcDepth_++ > 0) return;
  const cache = CacheService.getScriptCache();
  const until = Number(cache.get(CWC_LEASE_KEY) || 0);
  if (until > Date.now()) {
    cwcDepth_ = 0;
    throw new Error('A contractor check is already running (or stopped mid-run). It clears itself at ' +
      Utilities.formatDate(new Date(until), CWC.timezone, 'h:mm:ss a') + ', or run cwcResetRunGuard to clear it now.');
  }
  cache.put(CWC_LEASE_KEY, String(Date.now() + 7 * 60000), 7 * 60);
}
function cwcRelease_() {
  if (--cwcDepth_ > 0) return;
  cwcDepth_ = 0;
  CacheService.getScriptCache().remove(CWC_LEASE_KEY);
  PropertiesService.getScriptProperties().deleteProperty(CWC_LEASE_KEY);  // remove the old-version marker
}
/** Run manually only if you are sure no contractor check is in progress. */
function cwcResetRunGuard() { cwcRelease_(); Logger.log('Contractor check run guard cleared.'); }

/** Main entry point; safe to run repeatedly. */
function contractorWeeklyCheck() {
  cwcAcquire_();
  try {
    const started = new Date();
    const from = new Date(started.getTime() - CWC.lookbackDays * 86400000);
    const book = SpreadsheetApp.getActiveSpreadsheet();
    if (!book) throw new Error('Run from a spreadsheet-bound Apps Script project.');
    // One permanent tracker tab. New emails are appended; emails already on the
    // tab are left alone so their Striven verification results are kept.
    let sheet = book.getSheetByName(CWC.sheetName);
    if (!sheet) {
      sheet = book.insertSheet(CWC.sheetName);
      // Plain text for date and SO columns so Sheets does not turn "01234" into 1234.
      sheet.getRange('B:B').setNumberFormat('@');
      sheet.getRange('J:J').setNumberFormat('@');
      sheet.getRange(1, 1, 1, CWC.headers.length).setValues([CWC.headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    } else if (sheet.getRange(1, 1, 1, CWC.headers.length).getDisplayValues()[0].join('|') !== CWC.headers.join('|')) {
      throw new Error('Tab "' + CWC.sheetName + '" exists but its headers are not the tracker headers.');
    }
    const known = new Set();
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getDisplayValues()
        .forEach(r => { const id = String(r[0]).split('/').pop(); if (id) known.add(id); });
    }
    let added = 0;

    const totals = [];
    const list = CWC.contractors.filter(c => !cwcOnly_ || cwcOnly_.indexOf(c.name) >= 0);
    for (let i = 0; i < list.length; i++) {
      const contractor = list[i];
      // Split the remaining time fairly so the first contractor cannot starve the next.
      const remaining = CWC.maxRuntimeMs - (Date.now() - started.getTime());
      const deadline = Date.now() + Math.max(0, remaining / (list.length - i));
      let rows;
      try {
        const result = cwcCheckContractor_(contractor, from, started, deadline);
        rows = result.rows;
        // Summary rows (no matches / incomplete) go to the log and the header note,
        // not into the tracker, so the tab only ever holds real emails.
        const fresh = rows.filter(r => r[4] && !known.has(String(r[4]).split('/').pop()));
        const notes = rows.filter(r => !r[4]).map(r => r[13] + ': ' + r[14]);
        totals.push({contractor: contractor.name, emailsFound: result.count, newRows: fresh.length, status: result.status, notes: notes});
        rows = fresh;
      } catch (err) {
        const msg = String(err && err.message || err);
        rows = [];
        totals.push({contractor: contractor.name, status: 'ERROR / INCOMPLETE', error: msg});
      }
      cwcWriteRows_(sheet, rows);   // write per contractor so partial results survive
      added += rows.length;
    }
    if (added) sheet.autoResizeColumns(1, CWC.headers.length);
    sheet.getRange(1, 1).setNote('Last run ' + cwcFormat_(started) + ' (emails since ' + cwcFormat_(from) + '):\n' +
      totals.map(t => t.contractor + ': ' + t.status + (t.error ? ' - ' + t.error : ', ' + t.newRows + ' new of ' + t.emailsFound)).join('\n'));
    Logger.log(JSON.stringify({tracker: sheet.getName(), url: book.getUrl() + '#gid=' + sheet.getSheetId(),
      from: from.toISOString(), to: started.toISOString(), newRows: added, totals: totals}));
    if (totals.some(t => t.error)) throw new Error('Some contractors failed: ' + totals.filter(t => t.error).map(t => t.contractor + ' - ' + t.error).join('; '));
  } finally { cwcRelease_(); }
}

function cwcCheckContractor_(contractor, from, to, deadline) {
  const rows = [];
  // A missing/unshared calendar should not throw away the email results.
  const span = CWC.calendarWindowDays * 86400000;
  const cal = cwcLoadEvents_(contractor.calendarId, new Date(from.getTime() - span), new Date(to.getTime() + span));
  const events = cal.events, calendarNote = cal.note;
  // Epoch seconds are exact and timezone-independent (yyyy/MM/dd is read in Pacific time).
  const query = contractor.query + ' after:' + Math.floor(from.getTime() / 1000) + ' before:' + Math.ceil(to.getTime() / 1000 + 60);
  const seen = Object.create(null);
  let count = 0, offset = 0, incomplete = false, hitLimit = false;

  while (true) {
    if (offset >= CWC.maxThreadsPerContractor) { hitLimit = true; break; }
    if (Date.now() > deadline) { incomplete = true; break; }
    const threads = GmailApp.search(query, offset, CWC.pageSize);
    if (!threads.length) break;
    const messagesByThread = GmailApp.getMessagesForThreads(threads);   // one call instead of N
    for (const messages of messagesByThread) {
      for (const message of messages) {
        if (Date.now() > deadline) { incomplete = true; break; }
        const date = message.getDate(), id = message.getId();
        if (seen[id] || date < from || date > to) continue;
        seen[id] = true;
        // Only keep emails SENT by the contractor (From / Reply-To; group mail shows
        // as "'Steve Jensen' via Everyone"). Internal reports that merely mention
        // the contractor in the body are skipped.
        const sender = [message.getFrom(), message.getReplyTo()].join(' ');
        if (!contractor.identity.test(sender)) continue;
        const body = message.getPlainBody() || '';
        const subject = message.getSubject() || '';

        // iPhone/Apple Mail marks every attachment "inline", so inline parts must be
        // included; tiny inline images (logos, signature icons) are dropped.
        const attachments = message.getAttachments({includeInlineImages: true, includeAttachments: true})
          .filter(a => !(/^image\//i.test(a.getContentType()) && a.getSize() < 15000));
        const names = attachments.map(a => a.getName() || '(unnamed)');
        const isPdf = a => /\.pdf$/i.test(a.getName() || '') || a.getContentType() === 'application/pdf';
        const pdfs = attachments.filter(isPdf).length;
        const photos = attachments.filter(a => !isPdf(a) && (/^image\//i.test(a.getContentType()) || /\.(jpe?g|png|heic|webp)$/i.test(a.getName() || ''))).length;
        const other = attachments.length - photos - pdfs;
        const orderNumbers = cwcOrders_([subject, body, names.join(' ')].join('\n'));
        const addresses = cwcAddresses_(subject + '\n' + names.join('\n'));
        const matched = events.filter(e => cwcMatches_(e, orderNumbers, addresses));
        const candidates = matched.slice(0, 10).map(e => e.title + ' | ' + e.loc);
        const dates = matched.slice(0, 10).map(e => e.when);
        const receipt = /payment receipt|statement from/i.test(subject);
        const notes = [];
        if (receipt) notes.push('Receipt/statement; normally exclude from job entry.');
        if (pdfs) notes.push('PDF contents were not extracted; inspect invoice and exact work wording.');
        if (addresses.length > 1) notes.push('Multiple address hints: may contain multiple jobs.');
        if (matched.length > 10) notes.push('More than 10 calendar candidates; inspect calendar.');
        if (calendarNote) notes.push(calendarNote);
        notes.push('Run "Verify Striven tasks" to check the task.');
        rows.push([contractor.name, cwcFormat_(date), subject, message.getFrom(),
          'https://mail.google.com/mail/u/0/#all/' + id, photos, pdfs, other, names.join('\n'),
          orderNumbers.join(', '), addresses.join('\n'), candidates.join('\n'), dates.join('\n'),
          receipt ? 'EXCLUDE / REVIEW' : 'REVIEW REQUIRED', notes.join(' ')]);
        count++;
      }
      if (incomplete) break;
    }
    offset += threads.length;
    if (incomplete || threads.length < CWC.pageSize) break;
  }
  if (hitLimit) rows.push(cwcMarker_(contractor.name, 'INCOMPLETE', 'Stopped at ' + CWC.maxThreadsPerContractor + ' threads; some emails may be missing.'));
  if (incomplete) rows.push(cwcMarker_(contractor.name, 'INCOMPLETE', 'Stopped at the runtime limit; some emails may be missing. Re-run or shorten lookbackDays.'));
  if (!count && !incomplete && !hitLimit) rows.push(cwcMarker_(contractor.name, 'NO MATCHING EMAILS', 'Search completed for this configured query; other aliases may not be covered. ' + calendarNote));
  return {rows: rows, count: count, status: (incomplete || hitLimit) ? 'INCOMPLETE' : 'SEARCH COMPLETE'};
}

/** SO numbers: "SO 12345", "SO#12345", "SO-12345", "Sales Order #12345".
 *  "SO" must be upper case so ordinary text like "so 1500 dollars" is not matched. */
function cwcOrders_(text) {
  const re = /\b(?:SO|S\.O\.|[Ss]ales\s*[Oo]rder)\s*(?:#|no\.?|:|-)?\s*(\d{4,8})\b/g;
  return cwcUnique_([...String(text).matchAll(re)].map(m => m[1]));
}
function cwcMatches_(e, orders, addresses) {
  if (orders.length) {
    const evOrders = cwcOrders_([e.title, e.desc, e.loc].join(' '));
    if (orders.some(n => evOrders.indexOf(n) >= 0)) return true;
  }
  // Match on whole words in the location or title ("14 warwick" must not match "114 warwick").
  const hay = ' ' + cwcNormalize_(e.loc + ' ' + e.title) + ' ';
  return addresses.some(a => { const n = cwcNormalize_(a); return !!n && hay.indexOf(' ' + n + ' ') >= 0; });
}
function cwcAddresses_(text) {
  const re = /\b\d{1,5}\s+[A-Za-z0-9][A-Za-z0-9 .'-]{1,60}?\s+(?:road|rd|street|st|drive|dr|lane|ln|avenue|ave|grove|court|ct|crescent|cres|boulevard|blvd|line|trail|trl|way|place|pl|circle|cir|parkway|pkwy|terrace|terr|gate|heights|hts)\b/gi;
  // Short form without a street type, e.g. "14 Warwick, Ajax".
  const shortRe = /\b\d{1,5}\s+[A-Z][A-Za-z'-]+(?:\s+[A-Z][A-Za-z'-]+){0,2}(?=\s*,)/g;
  const full = (text.match(re) || []).map(s => s.replace(/\s+/g, ' ').trim());
  const short = (text.match(shortRe) || []).map(s => s.replace(/\s+/g, ' ').trim())
    .filter(sh => !full.some(f => cwcNormalize_(f).indexOf(cwcNormalize_(sh)) === 0));
  return cwcUnique_(full.concat(short));
}
function cwcNormalize_(s) {
  return String(s).toLowerCase().replace(/\broad\b/g,'rd').replace(/\bstreet\b/g,'st').replace(/\bdrive\b/g,'dr').replace(/\blane\b/g,'ln').replace(/\bavenue\b/g,'ave').replace(/\bcourt\b/g,'ct').replace(/\bcrescent\b/g,'cres').replace(/\bboulevard\b/g,'blvd').replace(/\btrail\b/g,'trl').replace(/[^a-z0-9]+/g,' ').trim();
}
/**
 * Loads events as {title, desc, loc, when}. Tries CalendarApp first (calendars in
 * the user's list), then reads the calendar directly through the Calendar API,
 * which works for PUBLIC calendars such as Steve's without subscribing to them.
 */
function cwcLoadEvents_(calendarId, start, end) {
  const calendar = CalendarApp.getCalendarById(calendarId);
  if (calendar) {
    return {note: '', events: calendar.getEvents(start, end).map(ev => ({
      title: ev.getTitle() || '', desc: ev.getDescription() || '', loc: ev.getLocation() || '', when: cwcEventTime_(ev),
      creator: String((ev.getCreators() || [])[0] || '').toLowerCase()}))};
  }
  try {
    const items = [];
    let pageToken = '';
    do {
      const url = 'https://www.googleapis.com/calendar/v3/calendars/' + encodeURIComponent(calendarId) + '/events' +
        '?singleEvents=true&orderBy=startTime&maxResults=2500' +
        '&timeMin=' + encodeURIComponent(start.toISOString()) + '&timeMax=' + encodeURIComponent(end.toISOString()) +
        (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
      const res = UrlFetchApp.fetch(url, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true});
      if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
      const body = JSON.parse(res.getContentText());
      (body.items || []).forEach(it => items.push(it));
      pageToken = body.nextPageToken || '';
    } while (pageToken);
    return {note: '', events: items.filter(it => it.status !== 'cancelled').map(it => ({
      title: it.summary || '', desc: it.description || '', loc: it.location || '', when: cwcApiEventTime_(it),
      creator: String((it.creator && it.creator.email) || (it.organizer && it.organizer.email) || '').toLowerCase()}))};
  } catch (err) {
    Logger.log('Calendar ' + calendarId + ' not readable: ' + (err && err.message || err));
    return {events: [], note: 'Calendar not readable (see execution log); calendar matching skipped.'};
  }
}
function cwcApiEventTime_(it) {
  const s = it.start || {}, e = it.end || {};
  if (s.date) return 'All-day: ' + s.date + ' to ' + (e.date || '') + ' (end exclusive)';
  return (s.dateTime ? cwcFormat_(new Date(s.dateTime)) : '') + ' to ' + (e.dateTime ? cwcFormat_(new Date(e.dateTime)) : '');
}
function cwcEventTime_(event) {
  if (event.isAllDayEvent()) return 'All-day: ' + Utilities.formatDate(event.getAllDayStartDate(), CWC.timezone, 'yyyy-MM-dd') + ' to ' + Utilities.formatDate(event.getAllDayEndDate(), CWC.timezone, 'yyyy-MM-dd') + ' (end exclusive)';
  return cwcFormat_(event.getStartTime()) + ' to ' + cwcFormat_(event.getEndTime());
}
function cwcFormat_(date) { return Utilities.formatDate(date, CWC.timezone, 'yyyy-MM-dd h:mm a'); }
function cwcUnique_(items) { return [...new Set(items.filter(Boolean))]; }
function cwcSafe_(v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v; }
function cwcMarker_(name, status, note) { return [name,'','','','','','','','','','','','',status,note]; }
/** One setValues call per batch instead of one appendRow per email. */
function cwcWriteRows_(sheet, rows) {
  if (!rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, CWC.headers.length)
    .setValues(rows.map(r => r.map(cwcSafe_)));
}

/**
 * STRIVEN TASK VERIFICATION — READ ONLY
 * Run contractorWeeklyCheck first, then contractorAuditStrivenCandidates
 * (menu "Contractor check > 2"). For every job email it answers:
 *   1. Was a Striven task already created for this job?  (by SO number from the
 *      email or the matched calendar event, by the number at the start of the
 *      calendar title, and by street address)
 *   2. Was that task updated after the contractor's email?  (status, last-updated
 *      and completed dates compared with the email date)
 * Credentials: Script properties CLIENT_ID + CLIENT_SECRET. Token is cached in
 * CacheService. Nothing is created or changed in Striven. Safe to re-run.
 * If the verdict column says field names were not recognised, run
 * cwcStrivenProbe once and send the execution log (it logs field NAMES only).
 */
const CWC_AUDIT_HEADERS = ['Striven task IDs','Striven task titles','Striven status / type','Striven dates','Task created?','Task updated after email?','Striven verdict','Striven attachments'];
const CWC_VERDICT = 6;   // index of 'Striven verdict' within CWC_AUDIT_HEADERS

function contractorAuditStrivenCandidates() {
  cwcAcquire_();
  try {
    const sh = cwcFindTracker_(SpreadsheetApp.getActiveSpreadsheet());
    const n = sh.getLastRow() - 1;
    if (n <= 0) return;
    const H = CWC_AUDIT_HEADERS, width = sh.getLastColumn();
    const existing = sh.getRange(1, 1, 1, width).getValues()[0];
    let start = existing.indexOf(H[0]) + 1;
    if (!start) { start = width + 1; sh.getRange(1, start, 1, H.length).setValues([H]).setFontWeight('bold'); }
    else {
      const cur = existing.slice(start - 1, start - 1 + H.length);
      if (cur.some((v, j) => v && v !== H[j])) throw new Error('Striven columns differ from expected layout.');
      sh.getRange(1, start, 1, H.length).setValues([H]).setFontWeight('bold');   // adds new columns to older tabs
    }

    const rows = sh.getRange(2, 1, n, 15).getDisplayValues();
    const prev = sh.getRange(2, start, n, H.length).getDisplayValues();
    const out = prev.map(r => r.slice());
    const ctx = {auth: {token: null}, search: Object.create(null), task: Object.create(null)};
    const deadline = Date.now() + 240000;
    let stoppedAt = -1;
    for (let i = 0; i < rows.length; i++) {
      const done = String(prev[i][CWC_VERDICT] || '');
      if (/^(OK|SKIPPED)/.test(done)) continue;   // final; ACTION/NOT FOUND/REVIEW rows are re-checked each run
      if (cwcOnly_ && cwcOnly_.indexOf(rows[i][0]) < 0) continue;
      if (Date.now() > deadline) { stoppedAt = i; break; }
      const r = rows[i];
      if (!r[4]) { out[i] = cwcAuditRow_([], '', '', 'SKIPPED: summary row'); continue; }
      if (/^EXCLUDE/.test(r[13])) { out[i] = cwcAuditRow_([], '', '', 'SKIPPED: receipt/statement'); continue; }
      try { out[i] = cwcVerifyRow_(r, ctx); }
      catch (e) { out[i] = cwcAuditRow_([], '', '', 'ERROR: ' + String(e && e.message || e)); }
    }
    if (stoppedAt >= 0) out[stoppedAt][CWC_VERDICT] = 'INCOMPLETE: run again to continue';
    sh.getRange(2, start, n, H.length).setValues(out.map(r => r.map(cwcSafe_)));
    cwcLinkTasks_(sh.getRange(2, start, n, 2));   // task ID + title columns
    sh.autoResizeColumns(start, H.length);
    Logger.log('Striven verification ' + (stoppedAt >= 0 ? 'paused; run again to continue.' : 'finished.') + ' No Striven changes were made.');
  } finally { cwcRelease_(); }
}

/** Verifies one tracker row; returns the 7 audit cells. */
function cwcVerifyRow_(r, ctx) {
  const calText = r[11] || '';                                  // "Calendar candidates"
  const sos = cwcUnique_(String(r[9] || '').split(/[,\s]+/).filter(x => /^\d{4,8}$/.test(x))
    .concat(cwcOrders_(calText)));
  const addresses = cwcUnique_(String(r[10] || '').split('\n').map(s => s.trim()).filter(Boolean));
  // Number at the start of a calendar title ("62746 Dunslow - SO585675 - ...") may be a task/job number.
  const leadIds = cwcUnique_(calText.split('\n').map(l => (l.match(/^\s*(\d{4,7})\b/) || [])[1]));
  const emailDate = r[1] ? Utilities.parseDate(r[1], CWC.timezone, 'yyyy-MM-dd h:mm a') : null;
  if (!sos.length && !addresses.length && !leadIds.length) return cwcAuditRow_([], '', '', 'REVIEW: no SO number, address or calendar match to search with');

  // Collect candidate task IDs from title searches and direct ID lookups.
  const found = Object.create(null);
  const shortAddr = addresses.map(a => a.replace(/\s+(road|rd|street|st|drive|dr|lane|ln|avenue|ave|grove|court|ct|crescent|cres|boulevard|blvd|line|trail|trl|way|place|pl|circle|cir|parkway|pkwy|terrace|terr|gate|heights|hts)\.?$/i, ''));
  for (const term of cwcUnique_(sos.concat(shortAddr))) {
    for (const t of cwcStrivenSearch_(term, ctx)) found[t.id] = t;
  }
  for (const id of leadIds) {
    if (found[id]) continue;
    const t = cwcStrivenTask_(id, ctx);
    if (t) found[id] = {id: id, title: cwcTitle_(t)};
  }
  // Keep only tasks that really belong to this job (SO number or address in the task).
  const keys = sos.map(s => s).concat(addresses.map(cwcNormalize_), shortAddr.map(cwcNormalize_));
  const tasks = [];
  for (const id of Object.keys(found).slice(0, 8)) {
    const detail = cwcStrivenTask_(id, ctx) || {};
    const blob = cwcNormalize_(found[id].title + ' ' + JSON.stringify(detail));
    const rel = keys.some(k => k && (' ' + blob + ' ').indexOf(' ' + cwcNormalize_(k) + ' ') >= 0 || (/^\d+$/.test(k) && blob.indexOf(k) >= 0));
    if (rel) { const t = cwcTaskSummary_(id, found[id].title, detail); t.detail = detail; tasks.push(t); }
  }
  // Each contractor's work belongs on a specific task type (Steve: Custom Work,
  // Jamie: Wood Installation). Other tasks for the job (e.g. the install) are shown
  // for reference but do not count.
  const contractor = CWC.contractors.filter(c => c.name === r[0])[0] || {};
  const exType = contractor.exampleTaskId ? String(cwcPick_(cwcStrivenTask_(String(contractor.exampleTaskId), ctx) || {}, ['Type']) || '') : '';
  const typeOk = t => !contractor.taskTypeMatch || (exType && t.type === exType) || contractor.taskTypeMatch.test(t.type);
  if (contractor.taskTypeMatch && tasks.length && !tasks.some(typeOk)) {
    // The required task may not carry the SO/address in its title: also search by the
    // customer's name and keep same-customer tasks of the required type.
    const custIds = cwcUnique_(tasks.map(t => String(cwcPick_(t.detail, ['Customer.Id']) || '')));
    const custNames = cwcUnique_(tasks.map(t => String(cwcPick_(t.detail, ['Customer.Name']) || '')));
    for (const name of custNames) {
      for (const hit of cwcStrivenSearch_(name, ctx)) {
        if (tasks.some(t => t.id === hit.id)) continue;
        const d = cwcStrivenTask_(hit.id, ctx) || {};
        const t = cwcTaskSummary_(hit.id, hit.title, d); t.detail = d;
        if (custIds.indexOf(String(cwcPick_(d, ['Customer.Id']) || '')) >= 0 && typeOk(t)) tasks.push(t);
      }
    }
  }
  const others = tasks.filter(t => !typeOk(t));
  const req = tasks.filter(typeOk);
  const otherNote = others.length ? ' (Other tasks for this job: ' + others.map(t => t.id + ' ' + t.type + ' / ' + t.status).join('; ') + ')' : '';
  const label = contractor.taskType || 'job';
  if (!req.length) {
    return cwcAuditRow_(others, 'NO ' + label + ' task', 'n/a',
      (tasks.length ? 'ACTION: create a ' + label + ' task' : 'NOT FOUND: no Striven task matches') +
      ' for SO ' + (sos.join(', ') || '-') + ' / ' + (addresses.join(', ') || '-') + '.' + otherNote);
  }
  tasks.length = 0; req.forEach(t => tasks.push(t));
  // Attachments on the task, compared with the files in the contractor's email.
  const after = d => d && emailDate && d.getTime() >= emailDate.getTime() - 60000;
  const emailFiles = String(r[8] || '').split('\n').map(x => x.trim()).filter(Boolean);
  for (const t of tasks.slice(0, 3)) t.att = cwcCompareAttachments_(cwcStrivenAttachments_(t.id, t.detail, ctx), emailFiles, after);

  // Was each task updated after the email? (new attachments count as an update)
  const verdicts = tasks.map(t => {
    if (!t.recognised) return 'UNKNOWN (field names not recognised; run cwcStrivenProbe)';
    if (t.att && t.att.newer) t.updated = t.updated && after(t.updated) ? t.updated : t.att.newestDate;
    if (t.isDone && (after(t.completed) || after(t.updated))) return 'YES: completed after email';
    if (t.isDone) return 'NO: already "' + t.status + '" since ' + (t.updated ? cwcFormat_(t.updated) : '?') + ', not touched after the email';
    if (after(t.updated)) return 'YES: edited after email, but status is "' + t.status + '"';
    return 'NO: status "' + t.status + '", not updated since email';
  });
  let verdict;
  if (tasks.length > 1) verdict = 'REVIEW: ' + tasks.length + ' ' + label + ' tasks match this job; confirm the right one.';
  else if (tasks[0].att && tasks[0].att.known && tasks[0].isDone && tasks[0].att.allPresent)
    verdict = 'OK: task is "' + tasks[0].status + '" and all ' + emailFiles.length + ' email files are on it.';
  else if (tasks[0].att && tasks[0].att.known && tasks[0].isDone)
    verdict = 'ACTION: task is "' + tasks[0].status + '" but ' + tasks[0].att.missing.length + ' of ' + emailFiles.length + ' email files are not on it: ' + tasks[0].att.missing.join(', ');
  else if (tasks[0].att && tasks[0].att.known && tasks[0].att.allPresent)
    verdict = 'ACTION: files are attached but task is still "' + tasks[0].status + '". Mark it complete.';
  else if (/^YES: completed/.test(verdicts[0])) verdict = 'CHECK: task completed after the email; attachments could not be read (run cwcStrivenProbe).';
  else if (tasks[0].att && !tasks[0].att.known && tasks[0].isDone) verdict = 'CHECK: task is "' + tasks[0].status + '" and not edited since the email; attachment names could not be read, so check the photos/invoice in Striven.';
  else if (/^UNKNOWN/.test(verdicts[0])) verdict = 'TASK EXISTS; update status unknown. Run cwcStrivenProbe and send the log.';
  else if (tasks[0].isDone) verdict = 'ACTION: task is already marked "' + tasks[0].status + '" but was not updated after the contractor\'s email. Add the email\'s photos/invoice to the task (or confirm they were added earlier).';
  else verdict = 'ACTION: task exists but is still "' + tasks[0].status + '". Mark it complete and add the photos/invoice.';
  return cwcAuditRow_(tasks, 'YES: ' + label + (tasks.length > 1 ? ' (' + tasks.length + ')' : ''), verdicts.join('\n'), verdict + otherNote);
}

function cwcAuditRow_(tasks, created, updated, verdict) {
  return [tasks.map(t => t.id).join('\n'), tasks.map(t => t.title).join('\n'),
    tasks.map(t => [t.status, t.type].filter(Boolean).join(' / ')).join('\n'),
    tasks.map(t => ['updated ' + (t.updated ? cwcFormat_(t.updated) : '?'), t.completed ? 'completed ' + cwcFormat_(t.completed) : ''].filter(Boolean).join(', ')).join('\n'),
    created, updated, verdict,
    tasks.map(t => t.att ? t.att.summary : '').filter(Boolean).join('\n')];
}

/**
 * Lists a task's attachments as [{name, date}]. Uses an attachment list inside
 * the task details if there is one, otherwise GETs the task's attachments
 * endpoint. The working endpoint is discovered once and remembered for 6 hours.
 * Returns null if Striven exposes no readable attachment list.
 */
const CWC_ATTACH_PATHS = ['v2/tasks/{id}/attachments', 'v1/tasks/{id}/attachments'];  // confirmed by cwcStrivenProbe
function cwcStrivenAttachments_(id, detail, ctx) {
  // Striven: GET v2/tasks/{id}/attachments -> data[{id, fileName, originalFileName, filePath, uploadedBy, dateCreated}]
  const toList = arr => arr.map(a => ({
    name: String(cwcPick_(a, ['OriginalFileName', 'FileName', 'Name', 'DocumentName', 'Title']) || ''),
    alt: String(cwcPick_(a, ['FileName']) || ''),
    by: String(cwcPick_(a, ['UploadedBy']) || ''),
    date: cwcDate_(cwcPick_(a, ['DateCreated', 'CreatedDate', 'DateUploaded', 'UploadedDate', 'DateAdded', 'CreatedOn', 'DateModified']))
  }));
  if (detail && typeof detail === 'object') {
    const key = Object.keys(detail).filter(k => /attach|document|file/i.test(k) && Array.isArray(detail[k]))[0];
    if (key && detail[key].length) return toList(detail[key]);
  }
  const cache = CacheService.getScriptCache();
  if (ctx.attPath === undefined) ctx.attPath = cache.get('cwcAttachPath');
  if (ctx.attPath === 'NONE') return null;
  const paths = ctx.attPath ? [ctx.attPath] : CWC_ATTACH_PATHS;
  for (const p of paths) {
    const res = cwcStriven_('get', p.replace('{id}', encodeURIComponent(id)), null, ctx);
    if (res.code === 200) {
      const list = cwcList_(res.json);
      if (Array.isArray(list)) {
        if (!ctx.attPath) { ctx.attPath = p; cache.put('cwcAttachPath', p, 21600); }
        return toList(list);
      }
    } else if (res.code !== 404 && res.code !== 405 && res.code !== 400) {
      throw new Error('Striven attachments for task ' + id + ' HTTP ' + res.code);
    }
  }
  if (!ctx.attPath) { ctx.attPath = 'NONE'; cache.put('cwcAttachPath', 'NONE', 3600); }
  return null;
}
/** Compares Striven attachments with the email's files (by file name, or uploaded after the email). */
function cwcCompareAttachments_(list, emailFiles, after) {
  if (!list) return {known: false, summary: 'Attachments not readable via API (run cwcStrivenProbe)'};
  // Don't report files as missing when we simply couldn't read the attachment names.
  if (list.length && !list.some(a => a.name || a.alt))
    return {known: false, summary: list.length + ' attachments on task, but file names not recognised (run cwcStrivenProbe)'};
  const norm = n => String(n).toLowerCase().replace(/\.[a-z0-9]+$/, '').replace(/[^a-z0-9]+/g, '');
  const names = list.map(a => norm(a.name)).concat(list.map(a => norm(a.alt)));
  const newerList = list.filter(a => after(a.date));
  const missing = emailFiles.filter(f => names.indexOf(norm(f)) < 0);
  // Striven may rename uploads; if enough files were added after the email, accept that.
  const allPresent = !missing.length || newerList.length >= emailFiles.length;
  const ext = a => ((a.name + ' ' + a.alt).match(/\.([a-z0-9]{2,5})\b/ig) || []).join(' ').toLowerCase();
  const photos = list.filter(a => /\.(jpe?g|png|heic|webp|gif)\b/.test(ext(a))).length;
  const pdfs = list.filter(a => /\.pdf\b/.test(ext(a))).length;
  const newestDate = newerList.map(a => a.date).sort((x, y) => y - x)[0] || null;
  return {known: true, allPresent: allPresent, missing: allPresent ? [] : missing, newer: newerList.length, newestDate: newestDate,
    summary: list.length + ' on task (' + photos + ' photos, ' + pdfs + ' PDFs), ' + newerList.length + ' added after email; ' +
      (allPresent ? 'all email files present' : 'missing: ' + missing.join(', ')) +
      '\nOn task: ' + list.map(a => (a.name || a.alt) + (a.date ? ' (' + Utilities.formatDate(a.date, CWC.timezone, 'yyyy-MM-dd') + (a.by ? ', ' + a.by : '') + ')' : '')).join('; ')};
}

/** Reads a value by trying several possible Striven field names ("A.B" = nested). */
function cwcPick_(o, paths) {
  for (const p of paths) {
    let v = o;
    for (const k of p.split('.')) {
      if (v == null) break;
      v = v[k] !== undefined ? v[k] : v[k.charAt(0).toLowerCase() + k.slice(1)];
    }
    if (v && typeof v === 'object' && !(v instanceof Date)) v = v.Name || v.name || v.Value || v.value || '';
    if (v !== undefined && v !== null && v !== '') return v;
  }
  return '';
}
function cwcDate_(v) { if (!v) return null; const d = new Date(v); return isNaN(d.getTime()) ? null : d; }
function cwcTitle_(t) { return String(cwcPick_(t, ['Title', 'TaskTitle', 'Name', 'Subject']) || ''); }
function cwcTaskSummary_(id, title, t) {
  const status = String(cwcPick_(t, ['Status', 'TaskStatus', 'StatusName', 'Stage']) || '');
  const type = String(cwcPick_(t, ['TaskType', 'Type', 'TaskTypeName']) || '');
  const updated = cwcDate_(cwcPick_(t, ['LastUpdatedDate', 'DateLastUpdated', 'LastModifiedDate', 'DateModified', 'ModifiedDate', 'UpdatedDate', 'DateUpdated', 'LastUpdated']));
  const completed = cwcDate_(cwcPick_(t, ['DateCompleted', 'CompletedDate', 'CompletionDate', 'DateClosed', 'ClosedDate']));
  return {id: id, title: title || cwcTitle_(t), status: status || '?', type: type,
    updated: updated, completed: completed, recognised: !!(status || updated || completed),
    isDone: /complete|closed|done|finished|resolved/i.test(status) || !!completed};
}

/** Low-level Striven call with token renewal on 401 and one retry on 429. */
function cwcStriven_(method, path, payload, ctx) {
  if (!ctx.auth.token) ctx.auth.token = cwcStrivenToken_(false);
  const call = () => UrlFetchApp.fetch('https://api.striven.com/' + path, {
    method: method, contentType: 'application/json', muteHttpExceptions: true,
    headers: {Authorization: 'Bearer ' + ctx.auth.token},
    payload: payload ? JSON.stringify(payload) : undefined
  });
  let res = call(), code = res.getResponseCode();
  if (code === 401) { ctx.auth.token = cwcStrivenToken_(true); res = call(); code = res.getResponseCode(); }
  if (code === 429) { Utilities.sleep(2000); res = call(); code = res.getResponseCode(); }
  let json = null; try { json = JSON.parse(res.getContentText()); } catch (e) {}
  return {code: code, json: json};
}
/** Title search; returns [{id,title}]. */
function cwcStrivenSearch_(term, ctx) {
  if (ctx.search[term]) return ctx.search[term];
  const all = [];
  for (let page = 0; page < 5; page++) {
    const res = cwcStriven_('post', 'v2/tasks/search', {Title: term, PageIndex: page, PageSize: 100}, ctx);
    if (res.code !== 200) throw new Error('Striven task search HTTP ' + res.code + (res.code === 401 ? ' (check CLIENT_ID/CLIENT_SECRET)' : ''));
    const data = cwcList_(res.json);
    if (!Array.isArray(data)) throw new Error('Unexpected Striven search response (keys: ' + Object.keys(res.json || {}).join(', ') + '); run cwcStrivenProbe.');
    for (const t of data) {
      const id = t.Id !== undefined ? t.Id : t.id;
      if (id != null) all.push({id: String(id), title: cwcTitle_(t)});
    }
    const total = res.json.TotalCount !== undefined ? res.json.TotalCount : res.json.totalCount;
    if (data.length < 100 || (total != null && (page + 1) * 100 >= Number(total))) break;
  }
  return (ctx.search[term] = all);
}
/** Finds the result list whatever Striven calls it (Data, Records, Items, Results...). */
function cwcList_(json) {
  if (Array.isArray(json)) return json;
  if (!json || typeof json !== 'object') return null;
  // Striven sends {"totalCount":0,"data":null} when nothing matches.
  for (const k of ['Data', 'data']) if (k in json && json[k] == null) return [];
  for (const k of ['Data', 'data', 'Records', 'records', 'Items', 'items', 'Results', 'results', 'Tasks', 'tasks']) {
    if (Array.isArray(json[k])) return json[k];
    if (json[k] && typeof json[k] === 'object') { const inner = cwcList_(json[k]); if (inner) return inner; }
  }
  const firstArray = Object.keys(json).map(k => json[k]).filter(Array.isArray)[0];
  return firstArray || null;
}
/** Task detail by ID (v2, falling back to v1); null if not found. */
function cwcStrivenTask_(id, ctx) {
  if (id in ctx.task) return ctx.task[id];
  let t = null;
  for (const path of ['v2/tasks/' + id, 'v1/tasks/' + id]) {
    const res = cwcStriven_('get', path, null, ctx);
    if (res.code === 200 && res.json && typeof res.json === 'object') { t = res.json.Data || res.json.data || res.json; break; }
    if (res.code !== 404 && res.code !== 400 && res.code !== 405) throw new Error('Striven task ' + id + ' HTTP ' + res.code);
  }
  return (ctx.task[id] = t);
}

/**
 * Diagnostic: logs the STRUCTURE (field names and types, no values) of a Striven
 * task search result and one task's details, so field mapping can be confirmed.
 * Status/type names and dates are logged because they are not customer data.
 */
function cwcStrivenProbe() {
  // 1. Endpoints Striven documents for attachments/documents/files (from its public help page).
  try {
    const html = UrlFetchApp.fetch('https://api.striven.com/Help', {muteHttpExceptions: true}).getContentText();
    const eps = cwcUnique_((html.match(/(GET|POST|PUT|DELETE)\s+v\d\/[^<\s"]+/g) || []).filter(e => /attach|document|file/i.test(e)));
    Logger.log('Documented attachment endpoints: ' + (eps.join(' | ') || '(none found on help page)'));
  } catch (e) { Logger.log('Help page not readable: ' + e); }
  CacheService.getScriptCache().remove('cwcAttachPath');
  const ctx = {auth: {token: null}, search: Object.create(null), task: Object.create(null)};
  const shape = (o, depth) => {
    if (Array.isArray(o)) return o.length ? [shape(o[0], depth)] : [];
    if (o && typeof o === 'object') {
      if (depth <= 0) return '{…}';
      const r = {}; Object.keys(o).forEach(k => r[k] = shape(o[k], depth - 1)); return r;
    }
    return o === null ? 'null' : typeof o;
  };
  const term = '585675';
  const s = cwcStriven_('post', 'v2/tasks/search', {Title: term, PageIndex: 0, PageSize: 5}, ctx);
  Logger.log('SEARCH "' + term + '" HTTP ' + s.code + ' shape: ' + JSON.stringify(shape(s.json, 4)));
  const data = cwcList_(s.json) || [];
  const ids = data.map(t => String(t.Id !== undefined ? t.Id : t.id)).concat(['62746']);
  for (const id of ids.slice(0, 2)) {
    for (const path of ['v2/tasks/' + id, 'v1/tasks/' + id]) {
      const d = cwcStriven_('get', path, null, ctx);
      Logger.log('GET ' + path + ' HTTP ' + d.code + ' shape: ' + JSON.stringify(shape(d.json, 3)));
      if (d.code === 200 && d.json) {
        const t = d.json.Data || d.json.data || d.json, sum = cwcTaskSummary_(id, '', t);
        Logger.log('Parsed: status=' + sum.status + ' type=' + sum.type + ' updated=' + sum.updated + ' completed=' + sum.completed + ' recognised=' + sum.recognised);
        for (const ap of CWC_ATTACH_PATHS) {
          const a = cwcStriven_('get', ap.replace('{id}', id), null, ctx);
          Logger.log('GET ' + ap.replace('{id}', id) + ' HTTP ' + a.code + ' shape: ' + JSON.stringify(shape(a.json, 3)));
        }
      }
    }
  }
}

/** Turns each task ID and title (one per line) into a link to that task in Striven. */
function cwcLinkTasks_(range) {
  if (!CWC.strivenTaskUrl) return;
  const rich = range.getDisplayValues().map(row => {
    const ids = row[0].split('\n').map(x => x.trim());
    return [cwcLinkLines_(row[0], ids), cwcLinkLines_(row[1], ids)];
  });
  range.setRichTextValues(rich);
}
function cwcLinkLines_(text, ids) {
  const b = SpreadsheetApp.newRichTextValue().setText(text);
  let pos = 0;
  text.split('\n').forEach((line, i) => {
    if (line && /^\d+$/.test(ids[i] || '')) b.setLinkUrl(pos, pos + line.length, CWC.strivenTaskUrl.replace('{id}', ids[i]));
    pos += line.length + 1;
  });
  return b.build();
}

function cwcFindTracker_(book) {
  const isTracker = s => s.getLastColumn() >= 15 && s.getRange(1, 1).getValue() === 'Contractor' && s.getRange(1, 10).getValue() === 'SO numbers found';
  const main = book.getSheetByName(CWC.sheetName);
  if (main && isTracker(main)) return main;
  const active = book.getActiveSheet();
  if (active && isTracker(active)) return active;
  // Older versions made timestamped tabs; fall back to the newest one.
  const trackers = book.getSheets().filter(s => /^Contractors \d{4}-\d{2}-\d{2} \d{6}/.test(s.getName()) && isTracker(s))
    .sort((a, b) => a.getName() < b.getName() ? 1 : -1);
  if (!trackers.length) throw new Error('No tracker tab found. Run contractorWeeklyCheck first.');
  return trackers[0];
}

/** Returns a Striven bearer token, renewing it from client credentials when configured. */
function cwcStrivenToken_(forceRefresh) {
  const props = PropertiesService.getScriptProperties();
  const first = names => names.map(n => props.getProperty(n)).filter(Boolean)[0] || '';
  const id = first(['CWC_STRIVEN_CLIENT_ID', 'STRIVEN_CLIENT_ID', 'CLIENT_ID']);
  const secret = first(['CWC_STRIVEN_CLIENT_SECRET', 'STRIVEN_CLIENT_SECRET', 'CLIENT_SECRET']);
  if (!id || !secret) {
    const manual = first(['CWC_STRIVEN_TOKEN']);
    if (!manual) throw new Error('Set script properties CLIENT_ID and CLIENT_SECRET (Striven API credentials).');
    return manual;
  }
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) { const cached = cache.get('cwcStrivenToken'); if (cached) return cached; }
  const res = UrlFetchApp.fetch('https://api.striven.com/accesstoken', {
    method: 'post', muteHttpExceptions: true,
    headers: {Authorization: 'Basic ' + Utilities.base64Encode(id + ':' + secret)},
    payload: {grant_type: 'client_credentials', ClientId: id}
  });
  if (res.getResponseCode() !== 200) throw new Error('Striven token request HTTP ' + res.getResponseCode() + '; check client ID/secret.');
  let body; try { body = JSON.parse(res.getContentText()); } catch (e) { throw new Error('Invalid Striven token response.'); }
  const token = body.access_token || body.AccessToken;
  if (!token) throw new Error('Striven token response had no access_token.');
  const ttl = Math.min(21600, Math.max(60, Number(body.expires_in || 3600) - 300));
  cache.put('cwcStrivenToken', token, ttl);
  return token;
}

/**
 * CREATE MISSING CONTRACTOR TASKS — the only part of this script that writes to Striven.
 * Run from the menu after "Verify Striven tasks". For each tracker row whose verdict is
 * "ACTION: create a <type> task", it re-checks Striven, then shows every task it would
 * create and writes nothing unless you click YES. Each new task copies customer,
 * location, Sales Order, division and priority from the job's existing task (e.g. the
 * Install task), uses the same title, is assigned to the contractor as subcontractor,
 * and is read back afterwards to confirm what Striven stored. Questionnaire fields are
 * left blank. Photos/invoice are not uploaded; the verify step lists them as missing.
 */
function contractorCreateMissingTasks() {
  cwcAcquire_();
  try {
    const sh = cwcFindTracker_(SpreadsheetApp.getActiveSpreadsheet());
    const n = sh.getLastRow() - 1;
    if (n <= 0) return;
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const start = head.indexOf(CWC_AUDIT_HEADERS[0]) + 1;
    if (!start) throw new Error('Run "Verify Striven tasks" first.');
    const rows = sh.getRange(2, 1, n, 15).getDisplayValues();
    const audit = sh.getRange(2, start, n, CWC_AUDIT_HEADERS.length).getDisplayValues();
    const ctx = {auth: {token: null}, search: Object.create(null), task: Object.create(null), v1: Object.create(null)};

    // 1. Build a plan, re-checking Striven so nothing is created twice.
    const plan = [], skipped = [];
    for (let i = 0; i < n; i++) {
      // Rows needing a task, plus rows where an earlier create attempt failed (re-checked below).
      if (!/^(ACTION: create a |ERROR creating task)/.test(audit[i][CWC_VERDICT])) continue;
      if (cwcOnly_ && cwcOnly_.indexOf(rows[i][0]) < 0) continue;
      if (!cwcIsCompletionEmail_(rows[i])) { skipped.push(rows[i][2] + ': not a job-completion email (no "completed" / invoice); left for review'); continue; }
      const r = rows[i];
      const contractor = CWC.contractors.filter(c => c.name === r[0])[0];
      if (!contractor || !contractor.taskType) { skipped.push(r[2] + ': no task type configured'); continue; }
      const fresh = cwcVerifyRow_(r, ctx);
      if (!/^ACTION: create a /.test(fresh[CWC_VERDICT])) { skipped.push(r[2] + ': ' + fresh[CWC_VERDICT]); continue; }
      const templateId = String(audit[i][0] || fresh[0]).split('\n')[0].trim();
      if (!/^\d+$/.test(templateId)) { skipped.push(r[2] + ': no existing task for this job to copy from'); continue; }
      try {
        const ids = cwcCreateIds_(contractor, ctx);
        const tpl = cwcStrivenV1Task_(templateId, ctx);
        if (!tpl) throw new Error('could not read task ' + templateId);
        // Start/end: 9 AM on the calendar event's first day to 5 PM on its last day.
        const span = cwcEventDays_(r[12]);
        const startDate = span.start, endDate = span.end;
        const email = cwcEmailDetails_(r);
        const requested = cwcRequestedByFromCalendar_(contractor, r, ctx);   // throws if not resolvable
        let body, api;
        if (ids.vendorId) {   // subcontractor vendor: v1 field names (as returned by GET v1/tasks/{id})
          api = 'v1';
          body = {
            taskName: tpl.taskName, taskTypeId: ids.typeId, accountId: tpl.accountId, locationId: tpl.locationId,
            orderId: tpl.orderId || null, divisionId: tpl.divisionId || null, priorityId: tpl.priorityId || null,
            useSubcontractor: true, assignedVendorId: ids.vendorId,
            desiredStartDate: startDate + 'T09:00:00', desiredEndDate: endDate + 'T17:00:00', taskDescription: email.description
          };
        } else {              // regular assignee: v2 field names (as returned by GET v2/tasks/{id})
          api = 'v2';
          const t2 = cwcStrivenTask_(templateId, ctx) || {};
          const ref = o => (o && o.id) ? {id: o.id} : null;
          body = {
            title: cwcContractorTitle_(t2, tpl, r), type: {id: ids.typeId}, customer: ref(t2.customer), location: ref(t2.location),
            requestedBy: {id: requested.id, type: 'employee'},
            salesOrder: ref(t2.salesOrder), priority: ids.priorityId ? {id: ids.priorityId} : ref(t2.priority), useSubContractor: false,
            assignments: ids.assignments.map(a => ({id: a.id, type: a.type})),
            startDateTime: startDate + 'T09:00:00', dueDateTime: endDate + 'T17:00:00', description: email.description
          };
        }   // status is set to Done only at the end, after details and files
        plan.push({row: i, contractor: contractor, templateId: templateId, ids: ids, body: body, api: api, email: r[2], files: email.files, text: email.text, requested: requested,
          title: body.taskName || body.title, start: startDate + ' 9:00 AM to ' + endDate + ' 5:00 PM',
          who: ids.vendorId ? 'subcontractor ' + ids.vendorName + ' (#' + ids.vendorId + ')' : 'assigned to ' + ids.assignments.map(a => a.name + ' (' + (a.type || '') + ' #' + a.id + ')').join(', ')});
      } catch (e) { skipped.push(r[2] + ': ' + (e && e.message || e)); }
    }
    if (!plan.length) {
      Logger.log('Nothing to create. ' + (skipped.length ? 'Skipped:\n• ' + skipped.join('\n• ') : 'No rows need a new task. Run "Verify Striven tasks" first.'));
      return;
    }

    // 2. Ask before writing anything.
    const text = plan.map(p => '• ' + p.contractor.taskType + ' task "' + p.title + '"\n   type ' + p.ids.typeName + ' (#' + p.ids.typeId + '), ' + p.who + ', status ' + (p.ids.statusName || 'Striven default') +
      ', start ' + p.start + '\n   customer, location and Sales Order copied from task ' + p.templateId +
      ', requested by ' + p.requested.name + ' (calendar event created by ' + p.requested.email + ')' +
      '\n   description: ' + (p.text ? '"' + p.text.slice(0, 200) + '"' : '(empty)') + '\n   attach ' + p.files.length + ' file(s): ' + p.files.map(f => f.getName()).join(', ') + '\n   for email: ' + p.email).join('\n\n') +
      (skipped.length ? '\n\nSkipped:\n• ' + skipped.join('\n• ') : '');
    Logger.log('Creating ' + plan.length + ' Striven task(s):\n' + text);

    // 3. Create, read back, and record the result on the tracker.
    const results = [];
    for (const p of plan) {
      let msg;
      try {
        const res = cwcStriven_('post', p.api + '/tasks', p.body, ctx);
        const newId = res.json && (typeof res.json === 'number' ? res.json : (res.json.id || res.json.Id || res.json.taskId || (res.json.data && res.json.data.id)));
        if (res.code !== 200 && res.code !== 201) throw new Error('HTTP ' + res.code + ' ' + JSON.stringify(res.json || '').slice(0, 200));
        if (!newId) throw new Error('Striven did not return the new task ID; check Striven before retrying.');
        const problems = [];
        if (p.api === 'v1') {
          const back = cwcStrivenV1Task_(String(newId), ctx, true) || {};
          if (back.taskTypeId !== p.body.taskTypeId) problems.push('type');
          if (back.accountId !== p.body.accountId) problems.push('customer');
          if (p.body.orderId && back.orderId !== p.body.orderId) problems.push('sales order');
          if (back.assignedVendorId !== p.body.assignedVendorId) problems.push('vendor');
        } else {
          delete ctx.task[String(newId)];
          const back = cwcStrivenTask_(String(newId), ctx) || {};
          const idOf = o => o && o.id;
          if (idOf(back.type) !== p.body.type.id) problems.push('type');
          if (p.body.customer && idOf(back.customer) !== p.body.customer.id) problems.push('customer');
          if (p.body.location && idOf(back.location) !== p.body.location.id) problems.push('location');
          if (p.body.salesOrder && idOf(back.salesOrder) !== p.body.salesOrder.id) problems.push('sales order');
          if (p.body.requestedBy && idOf(back.requestedBy) !== p.body.requestedBy.id) problems.push('requested by');
          const got = (back.assignments || []).map(a => a.id);
          if (p.body.assignments.some(a => got.indexOf(a.id) < 0)) problems.push('assignee');
        }
        // Attach the email's photos/invoice, then confirm the status.
        const up = p.files.map(f => { try { cwcStrivenUpload_(String(newId), f, ctx); return null; } catch (e) { return f.getName() + ' (' + (e && e.message || e) + ')'; } }).filter(Boolean);
        // Mark Done last, only if the details read back correctly and every file attached.
        delete ctx.task[String(newId)];
        if (!up.length && !problems.length && p.ids.statusId) {
          const now = cwcStrivenTask_(String(newId), ctx) || {};
          if (!now.status || now.status.id !== p.ids.statusId) {
            try { cwcStrivenUpdateTask_(now, {status: {id: p.ids.statusId}}, ctx); } catch (e) { problems.push(String(e && e.message || e)); }
            delete ctx.task[String(newId)];
          }
        }
        const st = String(cwcPick_(cwcStrivenTask_(String(newId), ctx) || {}, ['Status']) || '');
        if (p.ids.statusName && st.toLowerCase() !== p.ids.statusName.toLowerCase()) problems.push('status is "' + st + '"');
        msg = 'CREATED: ' + p.contractor.taskType + ' task ' + newId + ' (' + (st || '?') + '), ' + (p.files.length - up.length) + '/' + p.files.length + ' files attached' +
          (up.length ? '; NOT attached: ' + up.join(', ') : '') + (problems.length ? '; check in Striven: ' + problems.join(', ') : '; read back OK') + '. Run Verify to confirm.';
        sh.getRange(p.row + 2, start).setValue(String(newId));
      } catch (e) {
        msg = 'ERROR creating task: ' + (e && e.message || e);
      }
      sh.getRange(p.row + 2, start + CWC_VERDICT).setValue(cwcSafe_(msg));
      results.push(p.email + ' → ' + msg);
    }
    cwcLinkTasks_(sh.getRange(2, start, n, 2));
    Logger.log(results.join('\n'));
  } finally { cwcRelease_(); }
}

/** v1 task details (flat ids: taskTypeId, accountId, locationId, orderId, assignedVendorId...). */
function cwcStrivenV1Task_(id, ctx, fresh) {
  if (!fresh && ctx.v1[id]) return ctx.v1[id];
  const res = cwcStriven_('get', 'v1/tasks/' + id, null, ctx);
  return (ctx.v1[id] = res.code === 200 ? res.json : null);
}

/**
 * Finds the task type ID, the contractor's vendor ID and the create-status ID by
 * reading existing Striven tasks (type, status and assigned vendor are on every task).
 * Results are cached for 6 hours. Fixed IDs in CWC.strivenIds take precedence.
 */
function cwcCreateIds_(contractor, ctx) {
  const fixed = CWC.strivenIds[contractor.taskType] || {};
  const cache = CacheService.getScriptCache(), key = 'cwcIds8_' + contractor.taskType.replace(/\W/g, '');
  const cached = cache.get(key);
  if (cached && !fixed.typeId) return JSON.parse(cached);
  const ids = {typeId: fixed.typeId || null, typeName: contractor.taskType, vendorId: fixed.vendorId || null, vendorName: '', statusId: null, statusName: ''};
  const statusRe = new RegExp('^' + CWC.createStatus + '$', 'i');
  const typeTasks = [];
  // The contractor's example task defines the type and vendor exactly.
  if (contractor.exampleTaskId && !ids.typeId) {
    const ex = String(contractor.exampleTaskId);
    const d = cwcStrivenTask_(ex, ctx) || {}, v = cwcStrivenV1Task_(ex, ctx) || {};
    if (d.type && d.type.id) { ids.typeId = d.type.id; ids.typeName = d.type.name; typeTasks.push(ex); }
    if (d.status && statusRe.test(d.status.name || '')) { ids.statusId = d.status.id; ids.statusName = d.status.name; }
    if (d.priority && d.priority.id) ids.priorityId = d.priority.id;
    if (d.requestedBy && d.requestedBy.id) ids.requestedBy = {id: d.requestedBy.id, name: d.requestedBy.name, type: d.requestedBy.type};
    if (!ids.vendorId && v.assignedVendorId) { ids.vendorId = v.assignedVendorId; ids.vendorName = v.assignedVendorName || ''; }
    // Contractor assigned as a regular assignee (not a subcontractor vendor): copy those assignments.
    const asg = (d.assignments || []).filter(a => a && a.id);
    if (!ids.vendorId && asg.length) ids.assignments = asg.map(a => ({id: a.id, name: a.name, type: a.type}));
    if (contractor.assignTo) { ids.vendorId = null; ids.assignments = [{id: contractor.assignTo.id, name: contractor.assignTo.name || '', type: contractor.assignTo.type}]; }
  }
  // Other known example tasks: used only if their type matches this contractor.
  for (const hid of (CWC.strivenHints && CWC.strivenHints.taskIds) || []) {
    const d = cwcStrivenTask_(String(hid), ctx) || {};
    const type = d.type || {};
    if (!contractor.taskTypeMatch.test(type.name || '')) continue;
    if (!ids.typeId) ids.typeId = type.id;
    if (type.id === ids.typeId) { ids.typeName = type.name; typeTasks.push(hid); }
  }
  for (let page = 0; page < 10 && (!ids.typeId || !ids.statusId || typeTasks.length < 15); page++) {
    const res = cwcStriven_('post', 'v2/tasks/search', {PageIndex: page, PageSize: 100}, ctx);
    const data = res.code === 200 ? (cwcList_(res.json) || []) : [];
    for (const t of data) {
      const type = t.type || {}, status = t.status || {};
      if (!ids.statusId && statusRe.test(status.name || '')) { ids.statusId = status.id; ids.statusName = status.name; }
      const hinted = ((CWC.strivenHints && CWC.strivenHints.typeIds) || []).indexOf(type.id) >= 0;
      if (contractor.taskTypeMatch.test(type.name || '')) {
        if (!ids.typeId || (hinted && typeTasks.length === 0)) { ids.typeId = type.id; }
        if (type.id === ids.typeId) { ids.typeName = type.name; typeTasks.push(t.id); }
      }
    }
    if (data.length < 100) break;
  }
  if (!ids.typeId) throw new Error('no existing "' + contractor.taskType + '" task found to learn its type ID; set CWC.strivenIds');
  // Vendor: the subcontractor on existing tasks of this type whose name matches the contractor.
  if (!ids.vendorId && !ids.assignments) {
    for (const id of typeTasks.slice(0, 15)) {
      const v = cwcStrivenV1Task_(String(id), ctx) || {};
      if (v.assignedVendorId && contractor.identity.test(String(v.assignedVendorName || ''))) { ids.vendorId = v.assignedVendorId; ids.vendorName = v.assignedVendorName; break; }
    }
  }
  if (!ids.vendorId && !(ids.assignments && ids.assignments.length)) throw new Error('example task ' + (contractor.exampleTaskId || '') + ' has no subcontractor vendor or assignee for ' + contractor.name + '; set CWC.strivenIds');
  if (ids.vendorId && !ids.vendorName) ids.vendorName = contractor.name;
  cache.put(key, JSON.stringify(ids), 21600);
  return ids;
}

/** Text and files from the contractor email on a tracker row (message ID is in the Email link). */
function cwcEmailDetails_(r) {
  const id = String(r[4] || '').split('/').pop();
  const msg = id ? GmailApp.getMessageById(id) : null;
  if (!msg) throw new Error('email not found in Gmail');
  let body = String(msg.getPlainBody() || '').replace(/\r/g, '');
  // Drop quoted replies/forward history so only the contractor's own text is kept.
  body = body.split(/\n(?:On .{5,200}wrote:|-{2,}\s*Original Message|From: .+\nSent: )/)[0].replace(/\n{3,}/g, '\n\n').trim();
  if (body.length > 3500) body = body.slice(0, 3500) + '…';
  const files = msg.getAttachments({includeInlineImages: true, includeAttachments: true})
    .filter(a => !(/^image\//i.test(a.getContentType()) && a.getSize() < 15000));
  // Same style as task 19671: the contractor's notes as an HTML bullet list.
  const esc = t => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const points = body.split(/\n\s*\n|\n(?=\s*[-•*]\s)/).map(x => x.replace(/^\s*[-•*]\s*/, '').replace(/\s*\n\s*/g, ' ').trim())
    .filter(x => x && !/^(thanks|thank you|regards|cheers|sent from my)/i.test(x) && !/^steve\s*(jensen)?$/i.test(x));
  const description = points.length ? '<ul>\n' + points.map(x => '<li>' + esc(x) + '</li>').join('\n') + '</ul>' : '';
  return {description: description, files: files, text: points.join(' | ')};
}

/** Uploads one file to a task (multipart form). The working endpoint is remembered. */
function cwcStrivenUpload_(taskId, blob, ctx) {
  const paths = ctx.upPath ? [ctx.upPath] : ['v2/tasks/{id}/attachments', 'v1/tasks/{id}/attachments'];
  let last = '';
  for (const p of paths) {
    if (!ctx.auth.token) ctx.auth.token = cwcStrivenToken_(false);
    const res = UrlFetchApp.fetch('https://api.striven.com/' + p.replace('{id}', taskId), {
      method: 'post', muteHttpExceptions: true, headers: {Authorization: 'Bearer ' + ctx.auth.token},
      payload: {file: blob.copyBlob().setName(blob.getName())}
    });
    const code = res.getResponseCode();
    if (code === 200 || code === 201) { ctx.upPath = p; return; }
    last = 'HTTP ' + code + ' ' + res.getContentText().slice(0, 150);
    if (code !== 404 && code !== 405) break;
  }
  throw new Error(last);
}

/**
 * Logs everything the Striven API returns for one task (v1 + v2 details, custom
 * fields, assignments, attachments, notes) so an example task can be copied exactly.
 * Read-only. Asks for the task ID when run from the menu; defaults to 19671.
 */
function cwcAuditStrivenTask(taskId) {
  let id = taskId ? String(taskId) : '';
  if (!id) id = '19671';   // run from the editor with another ID: cwcAuditStrivenTask(19677)
  if (!/^\d+$/.test(id)) throw new Error('Task ID must be a number.');
  const ctx = {auth: {token: null}, search: Object.create(null), task: Object.create(null), v1: Object.create(null)};
  const paths = ['v2/tasks/' + id, 'v1/tasks/' + id, 'v2/tasks/' + id + '/attachments',
    'v2/tasks/' + id + '/notes', 'v1/tasks/' + id + '/notes', 'v2/tasks/' + id + '/comments', 'v1/tasks/' + id + '/notelogs'];
  const out = [];
  for (const p of paths) {
    const res = cwcStriven_('get', p, null, ctx);
    if (res.code === 404) { out.push('GET ' + p + ' → 404'); continue; }
    out.push('GET ' + p + ' → HTTP ' + res.code + '\n' + JSON.stringify(res.json, null, 1));
  }
  // Logger truncates long entries, so log in chunks.
  const text = out.join('\n\n');
  for (let i = 0; i < text.length; i += 7000) Logger.log(text.slice(i, i + 7000));
}

/** Title in the Custom Work format of task 19671: "SO#11568 - Allison Newton - 45 Lynndale Rd, Scarborough". */
function cwcContractorTitle_(t2, v1, r) {
  const so = (t2.salesOrder && t2.salesOrder.number) || (cwcOrders_(String(r[11] || '')) || [])[0] || '';
  const name = (t2.customer && t2.customer.name) || v1.accountName || '';
  // v1 locationName: "Primary Location - 45 Lynndale Rd, Scarborough, Toronto, ON M1N 1B9, CANADA"
  const loc = String(v1.locationName || '').replace(/^[^-]*-\s*/, '').split(',').map(x => x.trim()).filter(Boolean);
  const contractor = CWC.contractors.filter(c => c.name === r[0])[0] || {};
  const addr = loc.length ? loc.slice(0, contractor.titleAddressParts || 2).join(', ') : String(r[10] || '').split('\n')[0];
  return [so ? 'SO#' + so : '', name, addr].filter(Boolean).join(' - ');
}

/**
 * Requested By = the employee who created the job's calendar event (same rule as the
 * Pre-Inspection workflow: creator email must exactly match a Striven employee email).
 */
function cwcRequestedByFromCalendar_(contractor, r, ctx) {
  const title = String(r[11] || '').split('\n')[0].split(' | ')[0].trim();
  if (!title) throw new Error('no calendar event matched this email, so Requested By cannot be set');
  const d = Utilities.parseDate(r[1], CWC.timezone, 'yyyy-MM-dd h:mm a');
  const span = CWC.calendarWindowDays * 86400000;
  const ev = cwcLoadEvents_(contractor.calendarId, new Date(d.getTime() - span), new Date(d.getTime() + span))
    .events.filter(e => e.title.trim() === title)[0];
  if (!ev) throw new Error('calendar event "' + title + '" not found again');
  if (!ev.creator) throw new Error('calendar event "' + title + '" has no creator email');
  if (!ctx.employees) {
    const res = cwcStriven_('get', 'v1/employees', null, ctx);
    if (res.code !== 200) throw new Error('Striven employee list HTTP ' + res.code);
    ctx.employees = (cwcList_(res.json) || []).map(e => ({
      id: Number(e.Id || e.id || e.EmployeeId || e.employeeId || 0),
      name: String(e.Name || e.name || e.EmployeeName || e.employeeName || ''),
      email: String(e.Email || e.email || e.EmailAddress || e.emailAddress || '').trim().toLowerCase()
    })).filter(e => e.id > 0 && e.email);
  }
  const emp = ctx.employees.filter(e => e.email === ev.creator)[0];
  if (!emp) throw new Error('calendar event creator ' + ev.creator + ' is not a Striven employee email; Requested By unresolved');
  return {id: emp.id, name: emp.name, email: ev.creator};
}

/**
 * First and last day of the matched calendar event, from the tracker's
 * "Calendar dates" cell (first line), e.g.
 *   "All-day: 2026-11-04 to 2026-11-14 (end exclusive)" -> 2026-11-04 .. 2026-11-13
 *   "2026-11-04 9:00 AM to 2026-11-05 3:00 PM"          -> 2026-11-04 .. 2026-11-05
 */
function cwcEventDays_(cell) {
  const line = String(cell || '').split('\n')[0];
  const days = line.match(/\d{4}-\d{2}-\d{2}/g) || [];
  if (!days.length) throw new Error('no calendar event dates on this row, so start/end cannot be set');
  let end = days[1] || days[0];
  if (/^All-day/.test(line) && days[1]) {   // all-day end date is exclusive
    const d = Utilities.parseDate(days[1], CWC.timezone, 'yyyy-MM-dd');
    end = Utilities.formatDate(new Date(d.getTime() - 12 * 3600000), CWC.timezone, 'yyyy-MM-dd');
    if (end < days[0]) end = days[0];
  }
  return {start: days[0], end: end};
}

/**
 * COMPLETE CONTRACTOR TASKS — writes to Striven, after confirmation.
 * For each tracker row whose contractor task (Custom Work / Wood Installation) exists
 * but is not yet verified OK, brings the task in line with the rules:
 *   status Done; Requested By = calendar event creator; start 9 AM on the event's first
 *   day, end 5 PM on its last day; any email files not yet on the task are attached.
 * Shows every change first; then updates, reads back and reports what Striven stored.
 */
function contractorCompleteTasks() {
  cwcAcquire_();
  try {
    const sh = cwcFindTracker_(SpreadsheetApp.getActiveSpreadsheet());
    const n = sh.getLastRow() - 1;
    if (n <= 0) return;
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const start = head.indexOf(CWC_AUDIT_HEADERS[0]) + 1;
    if (!start) throw new Error('Run "Verify Striven tasks" first.');
    const rows = sh.getRange(2, 1, n, 15).getDisplayValues();
    const audit = sh.getRange(2, start, n, CWC_AUDIT_HEADERS.length).getDisplayValues();
    const ctx = {auth: {token: null}, search: Object.create(null), task: Object.create(null), v1: Object.create(null)};
    const plan = [], skipped = [];
    for (let i = 0; i < n; i++) {
      const created = String(audit[i][4] || ''), verdict = String(audit[i][CWC_VERDICT] || '');
      if (!/^YES/.test(created) && !/^CREATED/.test(verdict)) continue;
      if (/^(OK|SKIPPED)/.test(verdict)) continue;
      if (cwcOnly_ && cwcOnly_.indexOf(rows[i][0]) < 0) continue;
      if (!cwcIsCompletionEmail_(rows[i])) { skipped.push(rows[i][2] + ': not a job-completion email; not marking Done'); continue; }
      const r = rows[i];
      const contractor = CWC.contractors.filter(c => c.name === r[0])[0];
      const taskId = String(audit[i][0] || '').split('\n')[0].trim();
      if (!contractor || !/^\d+$/.test(taskId)) continue;
      try {
        delete ctx.task[taskId];
        const t = cwcStrivenTask_(taskId, ctx);
        if (!t) throw new Error('task ' + taskId + ' not readable');
        const ids = cwcCreateIds_(contractor, ctx);
        if (!(t.type && t.type.id === ids.typeId)) { skipped.push(r[2] + ': task ' + taskId + ' is not a ' + contractor.taskType + ' task'); continue; }
        const days = cwcEventDays_(r[12]);
        const requested = cwcRequestedByFromCalendar_(contractor, r, ctx);
        const want = {statusId: ids.statusId, start: days.start + 'T09:00:00', due: days.end + 'T17:00:00', requested: requested};
        const changes = [];
        if (!t.status || t.status.id !== want.statusId) changes.push('status ' + (t.status && t.status.name) + ' → ' + ids.statusName);
        if (String(t.startDateTime || '').slice(0, 19) !== want.start) changes.push('start ' + String(t.startDateTime || '').slice(0, 16) + ' → ' + want.start.slice(0, 16));
        if (String(t.dueDateTime || '').slice(0, 19) !== want.due) changes.push('end ' + String(t.dueDateTime || '').slice(0, 16) + ' → ' + want.due.slice(0, 16));
        if (!t.requestedBy || t.requestedBy.id !== requested.id) changes.push('requested by ' + (t.requestedBy && t.requestedBy.name) + ' → ' + requested.name + ' (' + requested.email + ')');
        // Email files not on the task yet.
        const email = cwcEmailDetails_(r);
        const onTask = (cwcStrivenAttachments_(taskId, null, ctx) || []);
        const norm = x => String(x).toLowerCase().replace(/\.[a-z0-9]+$/, '').replace(/[^a-z0-9]+/g, '');
        const have = onTask.map(a => norm(a.name)).concat(onTask.map(a => norm(a.alt)));
        const missing = email.files.filter(f => have.indexOf(norm(f.getName())) < 0);
        if (missing.length) changes.push('attach ' + missing.map(f => f.getName()).join(', '));
        if (!changes.length) { skipped.push(r[2] + ': task ' + taskId + ' already matches'); continue; }
        if (!String(t.description || '').trim() && email.description) changes.splice(changes.length - (missing.length ? 1 : 0), 0, 'description ← email text');
        // Done is applied last, after details and files.
        const doneIdx = changes.findIndex(c => /^status /.test(c));
        if (doneIdx >= 0) changes.push(changes.splice(doneIdx, 1)[0] + ' (last step)');
        plan.push({row: i, taskId: taskId, task: t, want: want, missing: missing, changes: changes, email: r[2], description: email.description});
      } catch (e) { skipped.push(r[2] + ': ' + (e && e.message || e)); }
    }
    if (!plan.length) { Logger.log('Nothing to change. ' + (skipped.length ? '\n• ' + skipped.join('\n• ') : 'No contractor tasks need changes.')); return; }
    const text = plan.map(p => '• Task ' + p.taskId + ' "' + p.task.title + '"\n   ' + p.changes.join('\n   ')).join('\n\n') + (skipped.length ? '\n\nSkipped:\n• ' + skipped.join('\n• ') : '');
    Logger.log('Updating ' + plan.length + ' Striven task(s):\n' + text);

    const results = [];
    for (const p of plan) {
      let msg;
      try {
        // 1. Details (dates, requested by, description) with the current status.
        const details = {startDateTime: p.want.start, dueDateTime: p.want.due, requestedBy: {id: p.want.requested.id, type: 'employee'}};
        if (!String(p.task.description || '').trim() && p.description) details.description = p.description;
        cwcStrivenUpdateTask_(p.task, details, ctx);
        // 2. Files from the email.
        const failed = p.missing.map(f => { try { cwcStrivenUpload_(p.taskId, f, ctx); return null; } catch (e) { return f.getName(); } }).filter(Boolean);
        // 3. Mark Done last, only when everything above worked.
        let doneNote = '';
        if (!failed.length) {
          delete ctx.task[p.taskId];
          const now = cwcStrivenTask_(p.taskId, ctx) || p.task;
          if (!now.status || now.status.id !== p.want.statusId) {
            try { cwcStrivenUpdateTask_(now, {status: {id: p.want.statusId}}, ctx); } catch (e) { doneNote = String(e && e.message || e); }
          }
        }
        delete ctx.task[p.taskId];
        const back = cwcStrivenTask_(p.taskId, ctx) || {};
        const problems = [];
        if (!back.status || back.status.id !== p.want.statusId) problems.push('status is ' + (back.status && back.status.name) + (doneNote ? ' (' + doneNote + ')' : ''));
        if (String(back.startDateTime || '').slice(0, 19) !== p.want.start) problems.push('start is ' + back.startDateTime);
        if (String(back.dueDateTime || '').slice(0, 19) !== p.want.due) problems.push('end is ' + back.dueDateTime);
        if (!back.requestedBy || back.requestedBy.id !== p.want.requested.id) problems.push('requested by is ' + (back.requestedBy && back.requestedBy.name));
        if (failed.length) problems.push('not attached: ' + failed.join(', '));
        msg = (problems.length ? 'PARTLY UPDATED: task ' + p.taskId + '; check: ' + problems.join('; ') : 'UPDATED: task ' + p.taskId + ' is ' + back.status.name + ', dates/requested by/files set (read back OK)') + '. Run Verify to confirm.';
      } catch (e) { msg = 'ERROR updating task ' + p.taskId + ': ' + (e && e.message || e); }
      sh.getRange(p.row + 2, start + CWC_VERDICT).setValue(cwcSafe_(msg));
      results.push(p.email + ' → ' + msg);
    }
    Logger.log(results.join('\n'));
  } finally { cwcRelease_(); }
}

/**
 * Updates an existing task with PATCH v2/tasks/{id} (Striven docs), sending only the
 * changed fields. That endpoint has no Status field, so a status change goes through
 * CWC.strivenStatusUpdate, which stays blocked until Striven's status endpoint is confirmed.
 */
function cwcStrivenUpdateTask_(t, changes, ctx) {
  const body = {};
  const map = {title: 'Title', startDateTime: 'StartDateTime', dueDateTime: 'DueDateTime', description: 'Description',
    priority: 'Priority', customer: 'Customer', location: 'Location', salesOrder: 'SalesOrder', useSubContractor: 'UseSubContractor'};
  Object.keys(changes).forEach(k => {
    if (k === 'status') return;
    if (k === 'requestedBy') body.RequestedBy = {Id: changes.requestedBy.id, Type: changes.requestedBy.type || 'employee'};
    else if (map[k]) body[map[k]] = (changes[k] && typeof changes[k] === 'object' && changes[k].id) ? {Id: changes[k].id} : changes[k];
    else throw new Error('field "' + k + '" cannot be changed with PATCH v2/tasks/{id}');
  });
  if (Object.keys(body).length) {
    const res = cwcStriven_('patch', 'v2/tasks/' + t.id, body, ctx);
    if (res.code !== 200 && res.code !== 204) throw new Error('PATCH v2/tasks/' + t.id + ' HTTP ' + res.code + ' ' + JSON.stringify(res.json || '').slice(0, 200));
  }
  if (changes.status) {
    const su = CWC.strivenStatusUpdate;
    if (!su) throw new Error('details updated, but Striven\'s status (mark Done) endpoint is not confirmed yet; task left as "' + (t.status && t.status.name) + '"');
    const res = cwcStriven_(su.method, su.path.replace('{id}', t.id), changes.markDoneFields || su.body, ctx);
    if (res.code !== 200 && res.code !== 204) throw new Error(su.method.toUpperCase() + ' ' + su.path + ' HTTP ' + res.code + ' ' + JSON.stringify(res.json || '').slice(0, 200));
  }
}

/**
 * Read-only: lists every task-related endpoint documented on Striven's API help page
 * (https://api.striven.com/Help), so the correct UPDATE / status endpoint can be set.
 */
function cwcListStrivenTaskEndpoints() {
  const get = u => { const r = UrlFetchApp.fetch(u, {muteHttpExceptions: true, followRedirects: true}); return {code: r.getResponseCode(), text: r.getContentText()}; };
  const plain = h => h.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/g, ' ').replace(/<[^>]+>/g, ' ').replace(/&quot;/g, '"').replace(/\s+/g, ' ').trim();
  // 1. Index pages / API descriptions that might list v2 endpoints.
  for (const u of ['https://api.striven.com/Help', 'https://api.striven.com/Help/Index/v2', 'https://api.striven.com/Help?version=v2',
                   'https://api.striven.com/swagger/docs/v2', 'https://api.striven.com/swagger/docs/v1', 'https://api.striven.com/swagger/v2/swagger.json']) {
    const r = get(u);
    const links = cwcUnique_((r.text.match(/(?:Help\/Api\/)?(?:GET|POST|PUT|PATCH|DELETE)-v\d+-[Tt]asks[A-Za-z0-9_\-]*/g) || []).map(x => x.replace('Help/Api/', '')));
    const paths = cwcUnique_((r.text.match(/"\/?v\d+\/tasks[^"]*"\s*:\s*\{\s*"(?:get|post|put|patch|delete)"/gi) || []));
    Logger.log(u + ' → HTTP ' + r.code + '; task links: ' + (links.join(' | ') || '-') + (paths.length ? '; swagger: ' + paths.join(' | ') : ''));
  }
  // 2. Help pages for the likely update / status endpoints (ASP.NET help-page naming).
  const pages = ['POST-v2-tasks', 'PUT-v2-tasks-id', 'POST-v2-tasks-id', 'PATCH-v2-tasks-id', 'PUT-v2-tasks',
    'PUT-v2-tasks-id-status', 'POST-v2-tasks-id-status', 'POST-v2-tasks-id-mark-done', 'POST-v2-tasks-id-markdone',
    'POST-v2-tasks-id-attachments', 'POST-v1-Tasks', 'PUT-v1-Tasks-id', 'POST-v1-Tasks-id', 'POST-v1-Tasks-id-Status',
    'PUT-v1-Tasks-id-Status', 'POST-v1-Tasks-id-Attachments', 'GET-v2-tasks-id', 'GET-v1-Tasks-id'];
  for (const pg of pages) {
    const r = get('https://api.striven.com/Help/Api/' + pg);
    const t = plain(r.text);
    if (r.code !== 200 || /not found|error/i.test(t.slice(0, 200)) || t.length < 200) { Logger.log(pg + ' → none (HTTP ' + r.code + ')'); continue; }
    const i = Math.max(0, t.indexOf(pg.split('-')[0] + ' v'));
    Logger.log(pg + ' → ' + t.slice(i, i + 1500));
  }
}

/* ===================== SCHEDULED END-TO-END RUN ===================== */
const CWC_SCHEDULE = Object.freeze({
  handler: 'contractorAutoRun',
  contractors: ['Steve Jensen', 'Kawartha / Jamie'],   // who the scheduled run handles
  everyHours: 2,                   // how often it runs (1, 2, 4, 6, 8 or 12)
  startHour: 7, endHour: 20,       // only between these hours, Toronto time
  notifyEveryRun: false            // false: email only when something changed or failed
});

/**
 * Steve Jensen email → Striven Custom Work task, end to end, no prompts:
 *  1. add new Steve emails to the tracker   2. verify Striven
 *  3. create missing Custom Work tasks      4. complete them (details, files, Done last)
 *  5. verify again so each row shows the final result.
 * Runs from the time trigger installed by contractorInstallSchedule.
 */
function contractorAutoRunSteve(e) { return contractorAutoRun(e); }   // old trigger name
function contractorAutoRun(e) {
  const scheduled = !!(e && e.triggerUid);   // menu/editor runs: always run and always email
  const hour = Number(Utilities.formatDate(new Date(), CWC.timezone, 'H'));
  if (scheduled && (hour < CWC_SCHEDULE.startHour || hour >= CWC_SCHEDULE.endHour)) return;   // outside working hours
  cwcOnly_ = CWC_SCHEDULE.contractors.slice();
  cwcPreflightResult_ = {};
  cwcAcquire_();
  const log = [], started = new Date();
  let before = {}, failed = null;
  try { before = cwcSnapshot_(); } catch (e) {}
  try {
    const step = (name, fn) => {
      try { fn(); log.push(name + ': ok'); }
      catch (e) { log.push(name + ': ' + (e && e.message || e)); throw e; }
    };
    step('1 weekly check', contractorWeeklyCheck);
    step('2 verify', contractorAuditStrivenCandidates);
    // Pre-flight: confirm every Striven endpoint and the model task for each contractor.
    // Contractors that fail are still checked (steps 1, 2, 5) but nothing is written for them.
    const all = cwcOnly_.slice(), ready = [];
    for (const name of all) {
      const c = CWC.contractors.filter(x => x.name === name)[0];
      const pf = cwcPreflight_(c);
      cwcPreflightResult_[name] = pf;
      log.push('preflight ' + name + ': ' + (pf.ok ? 'ok' : 'BLOCKED – ' + pf.problems.join('; ')));
      if (pf.ok) ready.push(name);
    }
    cwcOnly_ = ready;
    if (ready.length) {
      step('3 create', contractorCreateMissingTasks);
      step('4 complete', contractorCompleteTasks);
    }
    cwcOnly_ = all;
    step('5 verify', contractorAuditStrivenCandidates);
  } catch (e) {
    failed = e;
  } finally {
    Logger.log('Scheduled run (' + cwcOnly_.join(', ') + '): ' + log.join(' | '));
    try { cwcNotifyRun_(started, before, log, failed, !scheduled); } catch (e) { Logger.log('Notification failed: ' + e); }
    cwcRelease_();
    cwcOnly_ = null;
  }
  if (failed) throw failed;   // still marks the run as failed in Executions
}

/** Tracker rows for the scheduled contractors, keyed by email link. */
function cwcSnapshot_() {
  const book = SpreadsheetApp.getActiveSpreadsheet(), sh = book.getSheetByName(CWC.sheetName);
  if (!sh || sh.getLastRow() < 2) return {};
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const start = head.indexOf(CWC_AUDIT_HEADERS[0]) + 1;
  const n = sh.getLastRow() - 1;
  const rows = sh.getRange(2, 1, n, 15).getDisplayValues();
  const audit = start ? sh.getRange(2, start, n, CWC_AUDIT_HEADERS.length).getDisplayValues() : rows.map(() => []);
  const out = {};
  rows.forEach((r, i) => {
    if (!r[4] || (cwcOnly_ && cwcOnly_.indexOf(r[0]) < 0)) return;
    out[r[4]] = {contractor: r[0], date: r[1], subject: r[2], link: r[4], taskIds: String(audit[i][0] || '').split('\n')[0],
      status: audit[i][2] || '', verdict: audit[i][CWC_VERDICT] || '', files: audit[i][7] || ''};
  });
  return out;
}

/** Emails a summary of this run in the TMV3 notification format: new emails, changed results, errors. */
function cwcNotifyRun_(started, before, log, failed, always) {
  const after = cwcSnapshot_();
  const changed = Object.keys(after).filter(k => !before[k] || before[k].verdict !== after[k].verdict);
  const open = Object.keys(after).filter(k => !/^(OK|SKIPPED)/.test(after[k].verdict));
  const blocked = Object.keys(cwcPreflightResult_).filter(k => !cwcPreflightResult_[k].ok);
  if (!failed && !changed.length && !always && !CWC_SCHEDULE.notifyEveryRun) return;
  const attention = !!failed || blocked.length > 0 || open.length > 0;
  const item = k => { const a = after[k];
    return {lines: [a.contractor + ' – ' + a.subject, 'Email: ' + a.date, 'Striven task: ' + (a.taskIds || 'none') + (a.status ? ' (' + a.status + ')' : ''),
      a.verdict + (before[k] ? '' : ' (new email)')],
      links: [['Open email', a.link]].concat(/^\d+$/.test(a.taskIds) ? [['Open Striven task', CWC.strivenTaskUrl.replace('{id}', a.taskIds)]] : [])}; };
  const blocks = [];
  if (failed) blocks.push({lines: ['The run stopped with an error:', String(failed && failed.message || failed)]});
  if (blocked.length) blocks.push({lines: ['Nothing was written to Striven for ' + blocked.join(', ') + ':'].concat(blocked.map(b => cwcPreflightResult_[b].problems.join('; ')))});
  blocks.push({lines: ['Run at ' + cwcFormat_(started) + ' for ' + CWC_SCHEDULE.contractors.join(', ') + '.', 'Steps: ' + log.join(' | ')]});
  if (changed.length) { blocks.push({lines: ['Changed in this run: ' + changed.length]}); changed.forEach(k => blocks.push(item(k))); }
  const stillOpen = open.filter(k => changed.indexOf(k) < 0);
  if (stillOpen.length) { blocks.push({lines: ['Still needing attention: ' + stillOpen.length]}); stillOpen.forEach(k => blocks.push(item(k))); }
  if (!changed.length && !stillOpen.length) blocks.push({lines: ['No new contractor emails and nothing needing attention.']});
  blocks.push({links: [['Open Contractor Tracker', SpreadsheetApp.getActiveSpreadsheet().getUrl()]]});
  cwcSendNotice_(
    'Contractor automation: ' + (attention ? 'action required' : 'run complete'),
    attention ? (failed ? 'Contractor automation stopped and requires review.' : 'Contractor automation requires review.')
              : 'Contractor automation completed: ' + changed.length + ' updated.',
    attention, blocks);
}

/**
 * Sends one notification in the same format as the four TMV3 verticals
 * (Pre-Inspection, Install, Delivery, Service): automated-notice line, "Hi <name>,",
 * bold red headline when action is required, short paragraphs, "Open …" links,
 * sender name "Classic Fireplace & BBQ Store", plain-text body alongside the HTML.
 * Recipient: TMV3.OPERATIONS.verticalFailureAlertEmail when that config is in the
 * project, otherwise the account the script runs as.
 */
function cwcSendNotice_(subject, headline, attention, blocks) {
  if (MailApp.getRemainingDailyQuota() < 1) { Logger.log('Mail quota used up; notice not sent: ' + subject); return; }
  const to = (typeof TMV3 !== 'undefined' && TMV3.OPERATIONS && TMV3.OPERATIONS.verticalFailureAlertEmail) || Session.getEffectiveUser().getEmail();
  const first = (typeof tmv3_preInspectionCreatorFirstName_ === 'function') ? tmv3_preInspectionCreatorFirstName_(to)
    : to.split('@')[0].split(/[._-]/)[0].replace(/^./, ch => ch.toUpperCase());
  const esc = (typeof tmv3_preInspectionHtmlEscape_ === 'function') ? tmv3_preInspectionHtmlEscape_
    : t => String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const notice = 'This is an automated email notification from Classic Fireplace & BBQ Store.';
  const html = '<p><em>This is an automated email notification from Classic Fireplace &amp; BBQ Store.</em></p>' +
    '<p>Hi ' + esc(first) + ',</p>' +
    '<p style="' + (attention ? 'color:#B42318;' : '') + 'font-weight:700;">' + esc(headline) + '</p>' +
    blocks.map(b => (b.lines && b.lines.length ? '<p>' + b.lines.map(esc).join('<br>') + '</p>' : '') +
      (b.links && b.links.length ? '<p>' + b.links.map(l => '<a href="' + esc(l[1]) + '">' + esc(l[0]) + '</a>').join(' &nbsp;|&nbsp; ') + '</p>' : '')).join('');
  const text = notice + '\n\nHi ' + first + ',\n\n' + headline + '\n\n' +
    blocks.map(b => (b.lines || []).join('\n') + ((b.links || []).length ? '\n' + b.links.map(l => l[0] + ': ' + l[1]).join('\n') : '')).join('\n\n');
  MailApp.sendEmail({to: to, subject: subject, body: text, htmlBody: html, name: 'Classic Fireplace & BBQ Store'});
}

/** Run once: installs the schedule (replaces any earlier copy of it). */
function contractorInstallSchedule() {
  contractorRemoveSchedule();
  ScriptApp.newTrigger(CWC_SCHEDULE.handler).timeBased().everyHours(CWC_SCHEDULE.everyHours).create();
  Logger.log('Scheduled ' + CWC_SCHEDULE.handler + ' every ' + CWC_SCHEDULE.everyHours + ' h (' +
    CWC_SCHEDULE.startHour + ':00–' + CWC_SCHEDULE.endHour + ':00 Toronto) for ' + CWC_SCHEDULE.contractors.join(', ') + '.');
}
/** Stops the schedule. */
function contractorRemoveSchedule() {
  ScriptApp.getProjectTriggers().filter(t => ['contractorAutoRun', 'contractorAutoRunSteve'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/* ===================== PRE-FLIGHT (Striven set-up checks) ===================== */
let cwcPreflightResult_ = {};

/**
 * Read-only checks, per contractor, of everything the automation relies on:
 *  - auth token (POST /accesstoken)              - GET v2/tasks/{id} and v1/tasks/{id} on the model task
 *  - model task type matches the contractor       - other example tasks have the same type
 *  - model task has an assignee (who new tasks are assigned to)
 *  - "mark done" custom fields: none required, or the automation cannot mark Done
 *  - Done status id is known                      - GET v2/tasks/{id}/attachments works
 *  - POST v2/tasks/search works                   - GET v1/employees works (Requested By lookup)
 *  - the contractor's calendar is readable
 * Returns {ok, problems[], notes[]} and never writes to Striven.
 */
function cwcPreflight_(c) {
  const problems = [], notes = [];
  const ctx = {auth: {token: null}, search: Object.create(null), task: Object.create(null), v1: Object.create(null)};
  if (!c) return {ok: false, problems: ['contractor not configured'], notes: notes};
  try {
    try { ctx.auth.token = cwcStrivenToken_(false); } catch (e) { problems.push('Striven login: ' + e.message); throw e; }
    const id = String(c.exampleTaskId || '');
    const t2 = id && cwcStrivenTask_(id, ctx), t1 = id && cwcStrivenV1Task_(id, ctx);
    if (!t2 || !t1) problems.push('model task ' + id + ' not readable');
    else {
      const typeName = (t2.type && t2.type.name) || '';
      notes.push('model ' + id + ': "' + t2.title + '", type ' + typeName + ' (#' + (t2.type && t2.type.id) + '), status ' + (t2.status && t2.status.name));
      if (!c.taskTypeMatch.test(typeName)) problems.push('model task ' + id + ' is "' + typeName + '", not ' + c.taskType);
      for (const other of c.alsoExampleIds || []) {
        const o = cwcStrivenTask_(String(other), ctx);
        if (o && o.type && t2.type && o.type.id !== t2.type.id) problems.push('example ' + other + ' is type "' + o.type.name + '" but model ' + id + ' is "' + typeName + '"');
      }
      const asg = (t2.assignments || []).filter(a => a && a.id);
      if (!asg.length && !t1.assignedVendorId) problems.push('model task ' + id + ' has no assignee or subcontractor');
      else notes.push('assigned to ' + (asg.map(a => a.name + ' (' + a.type + ')').join(', ') || t1.assignedVendorName + ' (subcontractor)'));
      if (c.assignTo) notes.push('new tasks assigned to ' + (c.assignTo.name || c.assignTo.id) + ' (' + c.assignTo.type + ') from settings');
      else if (asg.length && asg.every(a => /pool/i.test(a.type))) notes.push('new tasks go to the "' + asg[0].name + '" pool, as on the reference task');
      const md = (t2.markDoneCustomFields || []).filter(f => f && (f.id || f.Id));
      if (md.length) {
        const req = md.filter(f => f.isRequired !== false);
        if (req.length) problems.push('marking Done needs custom fields: ' + req.map(f => f.name || f.Name || f.id).join(', ') + ' (tell Claude what to fill in)');
        else notes.push('optional mark-done fields left blank: ' + md.map(f => f.name).join(', '));
      }
      if (/^done$/i.test(CWC.createStatus) && !(t2.status && /^done$/i.test(t2.status.name)))
        notes.push('model task is not Done; Done status id taken from other tasks');
      const att = cwcStriven_('get', 'v2/tasks/' + id + '/attachments', null, ctx);
      if (att.code !== 200) problems.push('attachments endpoint HTTP ' + att.code);
      else notes.push('model has ' + ((cwcList_(att.json) || []).length) + ' attachments');
    }
    const sr = cwcStriven_('post', 'v2/tasks/search', {PageIndex: 0, PageSize: 1}, ctx);
    if (sr.code !== 200) problems.push('task search HTTP ' + sr.code);
    const em = cwcStriven_('get', 'v1/employees', null, ctx);
    if (em.code !== 200) problems.push('employee list HTTP ' + em.code + ' (needed for Requested By)');
    try {
      const ids = cwcCreateIds_(c, ctx);
      if (!ids.statusId) problems.push('status "' + CWC.createStatus + '" id not found');
    } catch (e) { problems.push(e.message); }
    const cal = cwcLoadEvents_(c.calendarId, new Date(Date.now() - 86400000), new Date());
    if (cal.note) problems.push('calendar: ' + cal.note);
  } catch (e) {
    if (!problems.length) problems.push(String(e && e.message || e));
  }
  return {ok: !problems.length, problems: problems, notes: notes};
}

/** Menu: runs the pre-flight for every scheduled contractor and logs/emails the result. Read-only. */
function contractorPreflightReport() {
  const lines = CWC_SCHEDULE.contractors.map(name => {
    const pf = cwcPreflight_(CWC.contractors.filter(x => x.name === name)[0]);
    return name + ': ' + (pf.ok ? 'READY' : 'BLOCKED') + '\n  ' + pf.problems.concat(pf.notes).join('\n  ');
  });
  Logger.log(lines.join('\n\n'));
  const notReady = lines.filter(l => /: BLOCKED/.test(l.split('\n')[0])).length;
  cwcSendNotice_('Contractor automation: ' + (notReady ? 'action required' : 'set-up check'),
    notReady ? 'Contractor automation set-up requires review.' : 'Contractor automation set-up check passed.',
    notReady > 0, lines.map(l => ({lines: l.split('\n').map(x => x.trim())})));
}

/** A contractor email counts as "job completed" if it says so or carries an invoice PDF. */
function cwcIsCompletionEmail_(r) {
  if (/^EXCLUDE/.test(r[13])) return false;                       // receipts / statements
  return /\b(complet\w*|finished|done|invoice)\b/i.test(r[2]) || Number(r[6]) > 0;
}

/**
 * Read-only evidence report for Jamie / Kawartha. Scans existing Striven tasks of the
 * Wood Fireplace Installation type (from model task 19676) and summarises how they are
 * really set up: who they are assigned to (employee / pool / subcontractor), who requests
 * them, title format, whether marking Done needs custom fields, attachments, and any task
 * that mentions Kawartha or Jamie. Logs and emails the result. Writes nothing.
 */
function contractorJamieReport() {
  const c = CWC.contractors.filter(x => x.name === 'Kawartha / Jamie')[0];
  const ctx = {auth: {token: null}, search: Object.create(null), task: Object.create(null), v1: Object.create(null)};
  const model = cwcStrivenTask_(String(c.exampleTaskId), ctx);
  if (!model || !model.type) throw new Error('model task ' + c.exampleTaskId + ' not readable');
  const typeId = model.type.id, lines = [];
  lines.push('Type: ' + model.type.name + ' (#' + typeId + '), model task ' + c.exampleTaskId);

  // 1. Find tasks of this type (search results carry type and status).
  const found = [];
  for (let page = 0; page < 15; page++) {
    const res = cwcStriven_('post', 'v2/tasks/search', {PageIndex: page, PageSize: 100}, ctx);
    const data = res.code === 200 ? (cwcList_(res.json) || []) : [];
    data.filter(t => t.type && t.type.id === typeId).forEach(t => found.push(t));
    if (data.length < 100) break;
  }
  const byStatus = {};
  found.forEach(t => { const k = (t.status && t.status.name) || '?'; byStatus[k] = (byStatus[k] || 0) + 1; });
  lines.push('Tasks of this type found: ' + found.length + ' — ' + Object.keys(byStatus).map(k => k + ' ' + byStatus[k]).join(', '));

  // 2. Details for the most recent Done ones and a few open ones.
  const recent = found.slice().sort((a, b) => String(b.lastUpdatedDate).localeCompare(String(a.lastUpdatedDate)));
  const sample = recent.filter(t => /done/i.test(t.status && t.status.name)).slice(0, 15)
    .concat(recent.filter(t => !/done/i.test(t.status && t.status.name)).slice(0, 5));
  const tally = (m, k) => { m[k] = (m[k] || 0) + 1; };
  const assignees = {}, vendors = {}, requesters = {}, markDone = {}, titles = [];
  const kawartha = [];
  for (const t of sample) {
    const d = cwcStrivenTask_(String(t.id), ctx) || {}, v = cwcStrivenV1Task_(String(t.id), ctx) || {};
    const done = /done/i.test(d.status && d.status.name) ? 'Done' : 'Open';
    (d.assignments || []).forEach(a => tally(assignees, done + ': ' + a.name + ' (' + a.type + ')'));
    if (!(d.assignments || []).length) tally(assignees, done + ': (none)');
    if (v.assignedVendorId) tally(vendors, done + ': ' + v.assignedVendorName + ' (#' + v.assignedVendorId + ')');
    if (d.requestedBy) requesters[t.id + ' [' + done + ']'] = d.requestedBy.name;
    (d.markDoneCustomFields || []).forEach(f => tally(markDone, (f.name || f.id) + (f.isRequired ? ' [required]' : '')));
    if (titles.length < 6) titles.push(t.id + ' [' + done + '] ' + d.title);
    if (/kawartha|jamie|khh/i.test(JSON.stringify(d) + JSON.stringify(v))) kawartha.push(t.id + ' [' + done + '] ' + d.title);
  }
  const show = (label, m) => lines.push(label + ': ' + (Object.keys(m).sort((a, b) => m[b] - m[a]).map(k => k + ' ×' + m[k]).join('; ') || 'none'));
  lines.push('Sampled ' + sample.length + ' tasks (up to 15 most recent Done + 5 open):');
  show('  Assigned to', assignees);
  show('  Subcontractor vendor', vendors);
  lines.push('  Requested by (one per task): ' + (Object.keys(requesters).map(k => k + ' ' + requesters[k]).join('; ') || 'none'));
  show('  Mark-done custom fields', markDone);
  lines.push('  Example titles:\n    ' + titles.join('\n    '));
  lines.push('  Tasks mentioning Kawartha/Jamie: ' + (kawartha.join('; ') || 'none in sample'));

  // 3. Jamie's rows on the tracker.
  const snap = (() => { const keep = cwcOnly_; cwcOnly_ = ['Kawartha / Jamie']; try { return cwcSnapshot_(); } finally { cwcOnly_ = keep; } })();
  const rows = Object.keys(snap).map(k => snap[k]);
  lines.push('Tracker rows for Kawartha / Jamie: ' + rows.length + (rows.length ? '\n    ' + rows.map(r => r.date + ' ' + r.subject + ' → ' + (r.verdict || '(not verified)')).join('\n    ') : ''));

  const text = lines.join('\n');
  Logger.log(text);
  cwcSendNotice_('Contractor automation: Wood Fireplace Installation analysis', 'Wood Fireplace Installation tasks (Jamie / Kawartha): analysis.',
    false, [{lines: text.split('\n').map(x => x.trim())}]);
}

/**
 * One-off for Jamie / Kawartha: every existing Wood Fireplace Installation task
 * (type of model task 19676) gets
 *   1. Requested By = the employee who created the job's calendar event on the Kawartha
 *      calendar (event matched by the task's SO number, else its street address), then
 *   2. status Done (POST v2/tasks/{id}/mark-done) as the last step.
 * Tasks whose calendar event or creator cannot be resolved are still marked Done, with
 * Requested By left as is and the reason reported. Each task is read back afterwards.
 * Results are logged and emailed in the standard notification format.
 */
function contractorJamieCloseExisting() {
  const c = CWC.contractors.filter(x => x.name === 'Kawartha / Jamie')[0];
  cwcAcquire_();
  try {
    const ctx = {auth: {token: null}, search: Object.create(null), task: Object.create(null), v1: Object.create(null)};
    const model = cwcStrivenTask_(String(c.exampleTaskId), ctx);
    if (!model || !model.type) throw new Error('model task ' + c.exampleTaskId + ' not readable');
    const typeId = model.type.id;
    const ids = cwcCreateIds_(c, ctx);
    if (!ids.statusId) throw new Error('Done status id not found');
    // All tasks of this type.
    const tasks = [];
    for (let page = 0; page < 15; page++) {
      const res = cwcStriven_('post', 'v2/tasks/search', {PageIndex: page, PageSize: 100}, ctx);
      const data = res.code === 200 ? (cwcList_(res.json) || []) : [];
      data.filter(t => t.type && t.type.id === typeId).forEach(t => tasks.push(t.id));
      if (data.length < 100) break;
    }
    // Employee directory for Requested By.
    const emRes = cwcStriven_('get', 'v1/employees', null, ctx);
    const employees = (cwcList_(emRes.json) || []).map(e => ({id: Number(e.Id || e.id || 0), name: String(e.Name || e.name || ''),
      email: String(e.Email || e.email || e.EmailAddress || e.emailAddress || '').trim().toLowerCase()})).filter(e => e.id && e.email);
    const results = [];
    for (const id of tasks) {
      const t = cwcStrivenTask_(String(id), ctx) || {};
      const v1 = cwcStrivenV1Task_(String(id), ctx) || {};
      const lines = ['Task ' + id + ': ' + t.title, 'Status before: ' + (t.status && t.status.name) + '; Requested By before: ' + (t.requestedBy && t.requestedBy.name)];
      try {
        // 1. Requested By from the calendar event creator.
        const so = (t.salesOrder && t.salesOrder.number) || '';
        const street = String(v1.locationName || '').replace(/^[^-]*-\s*/, '').split(',')[0].trim();
        const centre = t.startDateTime ? new Date(t.startDateTime) : new Date();
        const span = 120 * 86400000;
        const evs = cwcLoadEvents_(c.calendarId, new Date(centre.getTime() - span), new Date(centre.getTime() + span)).events;
        const ev = evs.filter(e => so && cwcOrders_(e.title + ' ' + e.desc).indexOf(so) >= 0)[0] ||
                   evs.filter(e => street && cwcNormalize_(e.title + ' ' + e.loc).indexOf(cwcNormalize_(street)) >= 0)[0];
        const emp = ev && ev.creator && employees.filter(e => e.email === ev.creator)[0];
        if (!ev) lines.push('Requested By unchanged: no calendar event found for SO ' + (so || '-') + ' / ' + (street || '-'));
        else if (!emp) lines.push('Requested By unchanged: event "' + ev.title + '" creator ' + (ev.creator || '(none)') + ' is not a Striven employee');
        else if (t.requestedBy && t.requestedBy.id === emp.id) lines.push('Requested By already ' + emp.name + ' (created event "' + ev.title + '")');
        else { cwcStrivenUpdateTask_(t, {requestedBy: {id: emp.id, type: 'employee'}}, ctx); lines.push('Requested By → ' + emp.name + ' (created event "' + ev.title + '")'); }
        // 2. Done, last.
        delete ctx.task[String(id)];
        const now = cwcStrivenTask_(String(id), ctx) || t;
        if (!now.status || now.status.id !== ids.statusId) cwcStrivenUpdateTask_(now, {status: {id: ids.statusId}}, ctx);
        delete ctx.task[String(id)];
        const back = cwcStrivenTask_(String(id), ctx) || {};
        lines.push('Read back: status ' + (back.status && back.status.name) + ', Requested By ' + (back.requestedBy && back.requestedBy.name) +
          ((back.status && back.status.id === ids.statusId) ? '' : '  ⚠ NOT Done'));
      } catch (e) { lines.push('ERROR: ' + (e && e.message || e)); }
      results.push({lines: lines, links: [['Open Striven task', CWC.strivenTaskUrl.replace('{id}', id)]]});
    }
    Logger.log(results.map(r => r.lines.join('\n')).join('\n\n'));
    const bad = results.filter(r => r.lines.some(l => /ERROR|NOT Done|unchanged/.test(l))).length;
    cwcSendNotice_('Contractor automation: ' + (bad ? 'action required' : 'run complete'),
      bad ? 'Wood Fireplace Installation tasks: ' + bad + ' of ' + results.length + ' need review.'
          : 'Wood Fireplace Installation tasks: all ' + results.length + ' are Done with Requested By from the calendar.',
      bad > 0, results);
  } finally { cwcRelease_(); }
}
