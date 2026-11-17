/**
 * Contractor Weekly Check — Google Apps Script V8
 * Read-only toward Gmail, Calendar and Striven; writes to one tracker tab ("Contractor Tracker").
 * Searches Kawartha first, then Steve. It does not create/update tasks, upload
 * files, send messages, or install triggers. Each tracker row is an EMAIL, not
 * a confirmed job. Mixed-job emails and PDFs require manual review.
 * Results go to ONE tab, "Contractor Tracker": each run adds only emails not
 * already listed. Paste into a spreadsheet-bound Apps Script project, reload the spreadsheet,
 * and use the "Contractor check" menu (or run contractorWeeklyCheck).
 */
const CWC = Object.freeze({
  timezone: 'America/Toronto',
  sheetName: 'Contractor Tracker', // single tab reused on every run
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
     identity:/kawartha home and hearth|khhltd@gmail\.com/i},
    {name:'Steve Jensen',
     calendarId:'steve.jensen83@gmail.com',
     query:'{from:steve.jensen83@gmail.com "Steve Jensen"} -in:trash -in:spam',
     identity:/steve\s+jensen|steve\.jensen83@gmail\.com/i}
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
    .addItem('Striven field check (diagnostic)', 'cwcStrivenProbe');
}
function cwcAddContractorMenu_() { cwcContractorMenu_().addToUi(); }

/**
 * Run guard for THIS script only. LockService.getScriptLock() is shared by every
 * file in the Apps Script project, so another script holding it made this fail
 * with "already running". The marker lives in CacheService (not Script
 * Properties, which are kept for real settings) and expires after 7 minutes.
 */
const CWC_LEASE_KEY = 'CWC_RUNNING_UNTIL';
function cwcAcquire_() {
  const cache = CacheService.getScriptCache();
  const until = Number(cache.get(CWC_LEASE_KEY) || 0);
  if (until > Date.now()) {
    throw new Error('A contractor check is already running (or stopped mid-run). It clears itself at ' +
      Utilities.formatDate(new Date(until), CWC.timezone, 'h:mm:ss a') + ', or run cwcResetRunGuard to clear it now.');
  }
  cache.put(CWC_LEASE_KEY, String(Date.now() + 7 * 60000), 7 * 60);
}
function cwcRelease_() {
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
    const list = CWC.contractors;
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
      title: ev.getTitle() || '', desc: ev.getDescription() || '', loc: ev.getLocation() || '', when: cwcEventTime_(ev)}))};
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
      title: it.summary || '', desc: it.description || '', loc: it.location || '', when: cwcApiEventTime_(it)}))};
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
const CWC_AUDIT_HEADERS = ['Striven task IDs','Striven task titles','Striven status / type','Striven dates','Task created?','Task updated after email?','Striven verdict'];

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
    else if (existing.slice(start - 1, start - 1 + H.length).join('|') !== H.join('|')) throw new Error('Striven columns differ from expected layout.');

    const rows = sh.getRange(2, 1, n, 15).getDisplayValues();
    const prev = sh.getRange(2, start, n, H.length).getDisplayValues();
    const out = prev.map(r => r.slice());
    const ctx = {auth: {token: null}, search: Object.create(null), task: Object.create(null)};
    const deadline = Date.now() + 240000;
    let stoppedAt = -1;
    for (let i = 0; i < rows.length; i++) {
      const done = String(prev[i][H.length - 1] || '');
      if (/^(OK|SKIPPED)/.test(done)) continue;   // final; ACTION/NOT FOUND/REVIEW rows are re-checked each run
      if (Date.now() > deadline) { stoppedAt = i; break; }
      const r = rows[i];
      if (!r[4]) { out[i] = cwcAuditRow_([], '', '', 'SKIPPED: summary row'); continue; }
      if (/^EXCLUDE/.test(r[13])) { out[i] = cwcAuditRow_([], '', '', 'SKIPPED: receipt/statement'); continue; }
      try { out[i] = cwcVerifyRow_(r, ctx); }
      catch (e) { out[i] = cwcAuditRow_([], '', '', 'ERROR: ' + String(e && e.message || e)); }
    }
    if (stoppedAt >= 0) out[stoppedAt][H.length - 1] = 'INCOMPLETE: run again to continue';
    sh.getRange(2, start, n, H.length).setValues(out.map(r => r.map(cwcSafe_)));
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
    if (rel) tasks.push(cwcTaskSummary_(id, found[id].title, detail));
  }
  if (!tasks.length) {
    return cwcAuditRow_([], 'NO', 'n/a', 'NOT FOUND: no Striven task matches SO ' + (sos.join(', ') || '-') + ' / ' + (addresses.join(', ') || '-') + '. Check Striven before creating one.');
  }
  // Was each task updated after the email?
  const verdicts = tasks.map(t => {
    if (!t.recognised) return 'UNKNOWN (field names not recognised; run cwcStrivenProbe)';
    const after = d => d && emailDate && d.getTime() >= emailDate.getTime() - 60000;
    if (t.isDone && (after(t.completed) || after(t.updated))) return 'YES: completed after email';
    if (t.isDone) return 'COMPLETED, but before this email (check it is the right task)';
    if (after(t.updated)) return 'YES: edited after email, but status is "' + t.status + '"';
    return 'NO: status "' + t.status + '", not updated since email';
  });
  let verdict;
  if (tasks.length > 1) verdict = 'REVIEW: ' + tasks.length + ' tasks match this job; confirm the right one.';
  else if (/^YES: completed/.test(verdicts[0])) verdict = 'OK: task exists and was completed after the email. Confirm photos/invoice are attached.';
  else if (/^UNKNOWN/.test(verdicts[0])) verdict = 'TASK EXISTS; update status unknown. Run cwcStrivenProbe and send the log.';
  else verdict = 'ACTION: task exists but needs updating (status, photos, invoice).';
  return cwcAuditRow_(tasks, 'YES (' + tasks.length + ')', verdicts.join('\n'), verdict);
}

function cwcAuditRow_(tasks, created, updated, verdict) {
  return [tasks.map(t => t.id).join('\n'), tasks.map(t => t.title).join('\n'),
    tasks.map(t => [t.status, t.type].filter(Boolean).join(' / ')).join('\n'),
    tasks.map(t => ['updated ' + (t.updated ? cwcFormat_(t.updated) : '?'), t.completed ? 'completed ' + cwcFormat_(t.completed) : ''].filter(Boolean).join(', ')).join('\n'),
    created, updated, verdict];
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
      }
    }
  }
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

