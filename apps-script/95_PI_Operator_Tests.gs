// Google Apps Script — PreInspection operator tools. Keep ONLY ONE copy of this file in the project.
// Never edit 00_Configuration or 30_PreInspection to test. Run from the editor (no arguments):
//
//   piTest_EndToEndCheck()            READ-ONLY: checks the whole live setup, PASS / WARN / FAIL per item
//   piTest_SimulateSlot()             READ-ONLY: emulates the 08:00/13:00/17:00 trigger for all events
//   piTest_PreviewNotices()           READ-ONLY: who would get a missing-details email now
//   piTest_SendTestEmails()           sends next-slot missing-details emails to Pramodh only, TEST subject
//   piTest_SendQuestionTest()         missing-details email with YES/NO buttons, to Pramodh only, TEST subject
//   piTest_ConfirmationAnswers()      READ-ONLY: recorded YES/NO answers
//   piTest_LastState()                READ-ONLY: last run result and pending queue
//   piTest_PropertyInventory()        READ-ONLY: Script Properties as KEEP / JUNK / OTHER
//   piTest_WRITE_CleanupProperties()  deletes JUNK properties only
//   piTest_CheckEvent()               READ-ONLY: why one event (settings.eventId) has or lacks Task links
//   piTest_WRITE_FixEvent()           WRITES: reconciles that one event now (same safety checks as the slots)
//   piTest_HeldRecords()              READ-ONLY: unfinished writes + duplicate-Task guards, checked in Striven
//   piTest_WRITE_ClearResolvedHeld()  deletes held records that are past or proven resolved in Striven

// Settings live inside a function (no top-level variables), so this file can never
// clash with another file's declarations.
function piTest_settings_() {
  return {
    simulate: {
      slot: '13:00',      // label only; all three slots run the same code
      fromStart: true     // true: evaluate every event; false: continue the real saved sweep position
    },
    emailTo: 'pramodh@classicfireplace.ca',  // test emails go ONLY here
    // One event for piTest_CheckEvent / piTest_WRITE_FixEvent (Calendar event ID, with or without @google.com).
    eventId: '1c5o71kgartm6q3j1ooj43d12d@google.com'   // 63267 James & Nadia Nickelo, Nov 13
  };
}

/** READ-ONLY emulation of a scheduled slot. Nothing is written to Striven, Calendar,
 * email, the audit log or cache sheets; PreInspection Script Properties are restored. */
function piTest_SimulateSlot() {
  const cfg = piTest_settings_().simulate || {};
  const props = PropertiesService.getScriptProperties();
  const isPiKey = function(k) { return /^TMV3_(?:PI_|PREINSPECTION_)/.test(k); };
  const before = props.getProperties();
  const titles = {};
  const would = {executions:[],emails:[],emailRecords:[],alerts:[],audit:0,cacheWrites:0};
  const names = ['tmv3_executeFreshStep7Selection_','tmv3_piSendMissingDetailsAfterRelease_','tmv3_preInspectionFailureAlert_',
    'tmv3_audit_','tmv3_upsertTaskCacheRow_','tmv3_preInspectionOptionalOrderLocationEvidence_','tmv3_piCalendarReviewRefresh_',
    'tmv3_preInspectionRefreshOperatorView_','tmv3_refreshAutomationSheetsAfter_'];
  const saved = {gate:TMV3_PI_CLEANED_RELEASE_APPROVED};
  names.forEach(function(n) { saved[n] = globalThis[n]; });
  let result = null, error = '';
  try {
    TMV3_PI_CLEANED_RELEASE_APPROVED = true;
    if (cfg.fromStart !== false) { props.deleteProperty('TMV3_PI_SWEEP_STATE'); props.deleteProperty('TMV3_PI_CURSOR'); }
    tmv3_executeFreshStep7Selection_ = function(selection) {
      const c = selection && selection.contract || {};
      would.executions.push({eventId:c.eventId,plan:c.plan,actions:c.actions,taskId:c.taskId || ''});
      return {status:'SIMULATED_NO_CHANGE'};
    };
    tmv3_piSendMissingDetailsAfterRelease_ = function(record) {
      const row = piTest_noticeRow_(record);
      if (row.missing.length && !would.emails.some(function(e) { return e.title === row.title && e.start === row.start; })) {
        would.emails.push(row);
        would.emailRecords.push(record);
      }
      return {status:'SIMULATED_NOT_SENT',sent:0};
    };
    tmv3_preInspectionFailureAlert_ = function(eventId, err) { would.alerts.push({eventId:eventId,error:String(err).slice(0, 200)}); return {status:'SIMULATED'}; };
    tmv3_audit_ = function() { would.audit++; };
    tmv3_upsertTaskCacheRow_ = function() { would.cacheWrites++; };
    tmv3_preInspectionOptionalOrderLocationEvidence_ = function() { return {status:'SIMULATED_SKIPPED_CACHE_REPAIR'}; };
    tmv3_piCalendarReviewRefresh_ = function() { return {status:'PI_REVIEW_UNCHANGED',simulated:true}; };
    tmv3_preInspectionRefreshOperatorView_ = function() { return {status:'PI_OPERATOR_VIEW_REFRESHED',simulated:true}; };
    tmv3_refreshAutomationSheetsAfter_ = function() { return {status:'SIMULATED_SKIPPED'}; };
    try { tmv3_preInspectionIntakeRecords_().forEach(function(r) { titles[r.eventId] = r.title + ' | ' + r.start; }); } catch (e) {}
    result = tmv3_runPreInspectionDailySlot_(cfg.slot || 'SIM');
  } catch (e) {
    error = String(e && e.message || e);
  } finally {
    TMV3_PI_CLEANED_RELEASE_APPROVED = saved.gate;
    names.forEach(function(n) { globalThis[n] = saved[n]; });
    const after = props.getProperties();
    Object.keys(after).forEach(function(k) { if (isPiKey(k) && !(k in before)) props.deleteProperty(k); });
    Object.keys(before).forEach(function(k) { if (isPiKey(k) && after[k] !== before[k]) props.setProperty(k, before[k]); });
  }
  const perEvent = (result && result.results || []).map(function(one) {
    const exec = would.executions.filter(function(x) { return x.eventId === one.eventId; })[0];
    return {event:titles[one.eventId] || one.eventId,eventId:one.eventId,
      outcome:exec ? ((exec.actions || []).some(function(a) { return a !== 'VERIFY_CALENDAR_LINKS'; }) ? 'WOULD_WRITE' : 'LINK_CHECK_ONLY') :
        one.error ? 'BLOCKED' : one.blocker ? 'NEEDS_REVIEW' : (one.status || one.plan || one.intake || ''),
      plan:one.plan || '',actions:exec ? exec.actions : [],taskId:one.taskId || '',blocker:one.blocker || '',
      guard:one.error || '',recreate:one.recreateAuthorized || '',orderCheck:one.orderStatusBeforeCreate && one.orderStatusBeforeCreate.status || ''};
  });
  const out = {status:error ? 'SIMULATION_ERROR' : 'SIMULATED_NO_WRITES',error:error,slot:cfg.slot,fromStart:cfg.fromStart !== false,
    slotStatus:result && result.status,sweep:result && result.sweep,batches:result && result.batchStatuses,
    wouldWriteTasks:would.executions.filter(function(x) { return (x.actions || []).some(function(a) { return a !== 'VERIFY_CALENDAR_LINKS'; }); }).length,
    wouldSendEmails:would.emails.filter(function(x) { return x.wouldSendNow; }).length,
    wouldAlert:would.alerts.length,offShared:result && result.offShared && {checked:result.offShared.checked,candidates:result.offShared.candidates,
      reviewNoEmail:result.offShared.reviewNoEmail || []},
    events:perEvent,emails:would.emails,emailRecords:would.emailRecords,alerts:would.alerts,propertiesRestored:true};
  Logger.log('PI_SIM_SUMMARY=' + JSON.stringify({status:out.status,slotStatus:out.slotStatus,events:perEvent.length,
    wouldWriteTasks:out.wouldWriteTasks,wouldSendEmails:out.wouldSendEmails,wouldAlert:out.wouldAlert,sweep:out.sweep,error:error}));
  perEvent.forEach(function(e) { Logger.log('PI_SIM_EVENT=' + JSON.stringify(e)); });
  would.emails.forEach(function(e) { Logger.log('PI_SIM_EMAIL=' + JSON.stringify(e)); });
  if (would.alerts.length) Logger.log('PI_SIM_ALERTS=' + JSON.stringify(would.alerts));
  if (out.offShared) Logger.log('PI_SIM_OFF_SHARED=' + JSON.stringify(out.offShared));
  return out;
}

/** Email preview row shared by the notice preview and the slot simulation. */
function piTest_noticeRow_(record) {
  const a = tmv3_preInspectionMissingDetailAssessment_(record);
  let state = null;
  const raw = PropertiesService.getScriptProperties().getProperty('TMV3_PI_NOTIFICATION_' + tmv3_hash_(record.eventId));
  if (raw && raw !== 'SENT') { try { state = JSON.parse(raw); } catch (e) {} }
  const d = a.shouldNotify ? tmv3_preInspectionNotificationDecision_(record, raw === 'SENT' ? {status:'SENT',lastSentAt:new Date().toISOString()} : state) : {status:'NO_MISSING_DETAILS',send:false};
  const blocked = !a.shouldNotify ? '' : !tmv3_normEmail_(a.creatorEmail) ? 'CREATOR_UNVERIFIED' :
    tmv3_piCreatorEmailExcluded_(a.creatorEmail) ? 'CREATOR_EXCLUDED' : '';
  return {title:record.title,start:record.start,creator:a.creatorEmail,onShared:!record.piOffSharedCandidate,
    missing:a.missing.map(function(m) { return m.message; }),wouldSendNow:!!d.send && !blocked,
    decision:blocked || d.status,eligibleAt:d.eligibleAt || ''};
}

function piTest_LastState() {
  const props = PropertiesService.getScriptProperties();
  const out = {lastResult:piTest_json_(tmv3_piEphemeral_().get('TMV3_PI_LAST_RESULT') || props.getProperty('TMV3_PI_LAST_RESULT')),
    pending:piTest_json_(props.getProperty('TMV3_PI_PENDING_EVENTS')),
    releaseGateInCode:TMV3_PI_CLEANED_RELEASE_APPROVED};
  Logger.log('PI_TEST_LAST_STATE=' + JSON.stringify(out));
  return out;
}

function piTest_norm_(id) { return String(id || '').trim().toLowerCase().replace(/@google\.com$/, ''); }

function piTest_json_(raw) { try { return raw ? JSON.parse(raw) : null; } catch (e) { return raw; } }

/** READ-ONLY: who would get a missing-details email right now, and why. Sends nothing. */
function piTest_PreviewNotices() {
  const saved = tmv3_piSendMissingDetailsAfterRelease_;
  const rows = [];
  const previewOne = piTest_noticeRow_;
  try {
    tmv3_piSendMissingDetailsAfterRelease_ = function(record) { rows.push(previewOne(record)); return {status:'PREVIEW_ONLY',sent:0}; };
    const off = tmv3_piNotifyOffSharedCandidates_({maxLiveChecks:12});
    tmv3_preInspectionIntakeRecords_().forEach(function(r) {
      if (tmv3_preInspectionMissingDetailAssessment_(r).shouldNotify) rows.push(previewOne(r));
    });
    const out = {status:'PREVIEW_ONLY_NO_EMAILS',offSharedChecked:off.checked,offSharedCandidates:off.candidates,
      skippedLiveCap:off.skippedLiveCap,sourceErrors:off.sourceErrors,reviewNoEmail:off.reviewNoEmail || [],wouldSendNow:rows.filter(function(r) { return r.wouldSendNow; }).length,rows:rows};
    Logger.log('PI_NOTICE_PREVIEW=' + JSON.stringify(out));
    return out;
  } finally { tmv3_piSendMissingDetailsAfterRelease_ = saved; }
}


/**
 * TEST SEND: every missing-details email the next slot would produce, rendered with the
 * approved template and sent ONLY to 'pramodh@classicfireplace.ca' with "TEST " added to the subject.
 * No CC/BCC. Does not touch the real notification state, so live sends are unaffected.
 */
function piTest_SendTestEmails() {
  if (piTest_settings_().emailTo !== 'pramodh@classicfireplace.ca') throw new Error('PI_TEST_RECIPIENT_CHANGED');
  const sim = piTest_SimulateSlot();
  if (sim.status !== 'SIMULATED_NO_WRITES') throw new Error('PI_TEST_SIMULATION_FAILED: ' + sim.error);
  const records = (sim.emailRecords || []).filter(function(r) {
    const a = tmv3_preInspectionMissingDetailAssessment_(r);
    return a.shouldNotify && tmv3_normEmail_(a.creatorEmail) && !tmv3_piCreatorEmailExcluded_(a.creatorEmail);
  });
  if (MailApp.getRemainingDailyQuota() < records.length + 1) throw new Error('PI_TEST_EMAIL_QUOTA_LOW');
  const sent = [];
  records.forEach(function(record) {
    const email = tmv3_preInspectionMissingDetailEmailPreview_(Object.assign({}, record, {notificationFollowupNumber:0}));
    if (!email) return;
    const assessment = tmv3_preInspectionMissingDetailAssessment_(record);
    const message = {to:'pramodh@classicfireplace.ca',subject:'TEST ' + email.subject,
      body:'Please update the missing Pre-Inspection details: ' + assessment.missing.map(function(x) { return x.message; }).join('; '),
      htmlBody:email.htmlBody,name:'Classic Fireplace & BBQ Store'};
    if (message.to !== 'pramodh@classicfireplace.ca' || message.cc || message.bcc) throw new Error('PI_TEST_RECIPIENT_GUARD');
    MailApp.sendEmail(message);
    sent.push({subject:message.subject,wouldGoTo:tmv3_piCreatorRecipients_(email.recipient).join(','),wouldCc:email.cc});
  });
  const out = {status:'TEST_EMAILS_SENT',to:'pramodh@classicfireplace.ca',count:sent.length,emails:sent};
  Logger.log('PI_TEST_EMAILS=' + JSON.stringify(out));
  return out;
}

/** TEST: the missing-details email WITH the YES/NO box, for every booking on Stephen's calendar that looks like a
 * pre-inspection but is missing CF Preinspects, to Pramodh only.
 * Clicking a TEST button records a test answer and never changes the appointment's classification. */
function piTest_SendQuestionTest() {
  if (piTest_settings_().emailTo !== 'pramodh@classicfireplace.ca') throw new Error('PI_TEST_RECIPIENT_CHANGED');
  const saved = tmv3_piSendMissingDetailsAfterRelease_;
  let off;
  try {
    tmv3_piSendMissingDetailsAfterRelease_ = function() { return {status:'SUPPRESSED_FOR_QUESTION_TEST',sent:0}; };
    off = tmv3_piNotifyOffSharedCandidates_({maxLiveChecks:12});
  } finally { tmv3_piSendMissingDetailsAfterRelease_ = saved; }
  // Same rule as live: looks like a pre-inspection, CF Preinspects not a guest, not yet answered.
  const sent = (off.candidateRecords || []).filter(function(r) { return !tmv3_piVisitConfirmation_(r.eventId); }).map(function(record) {
    const r = Object.assign({}, record, {piOffSharedCandidate:true, notificationFollowupNumber:0});
    const q = tmv3_piAddVisitQuestion_(tmv3_preInspectionMissingDetailEmailPreview_(r), r, true);
    if (!q.questionAdded) throw new Error('PI_TEST_QUESTION_SECTION_NOT_ADDED');
    const message = {to:'pramodh@classicfireplace.ca',subject:'TEST ' + q.subject,
      body:'Please update the missing Pre-Inspection details.' + (q.bodySuffix || ''),htmlBody:q.htmlBody,name:'Classic Fireplace & BBQ Store'};
    if (message.to !== 'pramodh@classicfireplace.ca' || message.cc || message.bcc) throw new Error('PI_TEST_RECIPIENT_GUARD');
    MailApp.sendEmail(message);
    return {subject:message.subject,wouldGoTo:tmv3_piCreatorRecipients_(q.recipient).join(','),wouldCc:q.cc};
  });
  const out = {status:'TEST_QUESTIONS_SENT',to:'pramodh@classicfireplace.ca',count:sent.length,emails:sent};
  Logger.log('PI_TEST_QUESTIONS=' + JSON.stringify(out));
  return out;
}

/** READ-ONLY: answers recorded from TEST question emails (and any live answers). */
function piTest_ConfirmationAnswers() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const rows = Object.keys(props).filter(function(k) { return /^TMV3_PI_CONFIRM_(?:TEST_)?[0-9a-f]+$/.test(k); })
    .map(function(k) { return Object.assign({kind:k.indexOf('_TEST_') > 0 ? 'TEST' : 'LIVE'}, piTest_json_(props[k])); });
  Logger.log('PI_CONFIRM_ANSWERS=' + JSON.stringify(rows));
  return rows;
}

/** Classify a Script Property key: KEEP (needed state), JUNK (safe to delete) or OTHER (not PreInspection: never touched). */
function piTest_propertyClass_(key, value, todayIso, upcoming) {
  const json = piTest_json_(value);
  // Records tied to one appointment: useful only while that appointment is still ahead.
  // upcoming = hashes of every event from yesterday on (CF Preinspects + Stephen); null = calendars unreadable, keep all.
  const eventKey = /^TMV3_PI_(?:NOTIFICATION|QUESTION|ORIGIN|ASSIGNMENT_OVERRIDE)_([0-9a-f]{64})$/.exec(key);
  if (eventKey) {
    if (upcoming && !upcoming[eventKey[1]]) return {cls:'JUNK',why:'appointment is in the past'};
    return {cls:'KEEP',why:/NOTIFICATION|QUESTION/.test(key) ? 'stops repeat emails (appointment ahead)' : 'needed for an upcoming appointment'};
  }
  const unfinished = json && ['INTENT','UNCERTAIN','SENDING'].indexOf(json.state) >= 0;
  if (/^TMV3_(?:TASK|LOCATION)_CREATE_GUARD_/.test(key) && json && json.eventId) {
    if (unfinished) return {cls:'KEEP',why:'unfinished write: check Striven for this appointment first'};
    if (upcoming && !upcoming.ids[json.eventId]) return {cls:'JUNK',why:'duplicate-create guard for a past appointment'};
    return {cls:'KEEP',why:'stops a duplicate Task/Location for an upcoming appointment'};
  }
  const vEvent = /^TMV3_VERTICAL_EVENT_([0-9a-f]{64})$/.exec(key);
  if (vEvent) {
    if (unfinished) return {cls:'KEEP',why:'unfinished write: check Striven for this appointment first'};
    if (json && json.state === 'VERIFIED') return {cls:'JUNK',why:'log of a finished write (TM Audit has it)'};
    if (upcoming && !upcoming.byVertical[vEvent[1]]) return {cls:'JUNK',why:'write record for a past appointment'};
    return {cls:'KEEP',why:'write record for an upcoming appointment'};
  }
  if (/^TMV3_VERTICAL_ALERT_/.test(key)) return value === 'SENT' ? {cls:'JUNK',why:'failure-alert "sent today" flag'} : {cls:'KEEP',why:'alert being sent'};
  if (/^TMV3_VERTICAL_(?:LAST|CURSOR)_/.test(key) || key === 'TMV3_STANDARD_SCHEDULE')
    return {cls:'JUNK',why:'old status/cursor no code reads'};
  if (/^TMV3_VERTICAL_RESULT_/.test(key) || /^TMV3_SHEET_REFRESH_LAST_(?:SUCCESS|ERROR)$/.test(key))
    return {cls:'JUNK',why:'status log (now kept in cache; TM Audit has it)'};
  if (key === 'TMV3_STRIVEN_API_BUDGET_EXHAUSTED_DATE' || key === 'TMV3_STRIVEN_API_BUDGET_EXHAUSTED_REASON') {
    const d = String((piTest_json_(props_exhaustedDate_()) || '')).slice(0, 10);
    return d === todayIso ? {cls:'KEEP',why:'Striven daily limit hit today: pauses calls'} : {cls:'JUNK',why:'old "Striven limit reached" note from a past day'};
  }
  if (/^TMV3_SD_ORDER_BACKFILL_20261020_/.test(key))
    return typeof TMV3_SD_ORDER_BACKFILL_RETIRED !== 'undefined' && TMV3_SD_ORDER_BACKFILL_RETIRED === true ?
      {cls:'JUNK',why:'one-time Oct 20 order backfill (retired)'} : {cls:'KEEP',why:'Oct 20 order backfill still running (retire it in 10_Shared_Core first)'};
  const sentOn = /^TMV3_INSTALL_SO_REMINDER_/.test(key) && /^(\d{4}-\d{2}-\d{2})T/.exec(String(value || ''));
  if (sentOn) return sentOn[1] < todayIso ? {cls:'JUNK',why:'"already reminded today" flag from a past day'} : {cls:'KEEP',why:'stops a second reminder today'};
  if (/^TMV3_OWNER_BUDGET_RESET_/.test(key)) return {cls:'JUNK',why:'one-off budget reset note (no code reads it)'};
  if (key === 'TMV3_INSTALL_VIEW_LAST_SUCCESS') return {cls:'JUNK',why:'status note no code reads (now kept in cache)'};
  const junkExact = ['TMV3_PI_LAST_RESULT','TMV3_PI_BUDGET_PAUSE_DATE','TMV3_PI_RETRY_SWEEP_START','TMV3_PI_DISCOVERY',
    'TMV3_PI_LAST_PENDING_RECONCILIATION','TMV3_PI_VERIFY_CURSOR'];
  if (junkExact.indexOf(key) >= 0) return {cls:'JUNK',why:'cache/diagnostic now kept in CacheService or unused'};
  if (/^TMV3_PI_(?:DONE|FAILURE|OFFSHARED_INTENT|INFERRED)_/.test(key)) return {cls:'JUNK',why:'short-lived cache'};
  if (/^TMV3_PI_CONFIRM_TEST_/.test(key)) return {cls:'JUNK',why:'test answer'};
  const dated = /^TMV3_PI_(?:ELIGIBLE|TODAY)_CURSOR_(\d{4}-\d{2}-\d{2})$/.exec(key);
  if (dated) return dated[1] < todayIso ? {cls:'JUNK',why:'cursor for a past day'} : {cls:'KEEP',why:'PreInspection run state'};
  // Finished one-off fixes (Oct 2026). No code reads these keys any more.
  if (/^TMV3_(?:SD_REQUESTED_BY_2026|OPEN_REFRESH_2026|SD_CACHE_REPAIR_PRE_|JIM_[A-Z_]*PATCH_)/.test(key))
    return {cls:'JUNK',why:'finished one-off fix marker (no code reads it)'};
  const range = /^TMV3_VERTICAL_SWEEP_[A-Za-z]+_RANGE_\d{4}-\d{2}-\d{2}_(\d{4}-\d{2}-\d{2})(?:_|$)/.exec(key);
  if (range && range[1] < todayIso) return {cls:'JUNK',why:'repair sweep for a past week'};
  if (/^TMV3_(?:PI_|PREINSPECTION_)/.test(key)) return {cls:'KEEP',why:'PreInspection run state'};
  return {cls:'OTHER',why:'shared V3 / other vertical — not touched'};
}

/** Every appointment from yesterday on, on every workflow calendar (PreInspection incl. Stephen,
 * Install, Delivery, Service technicians). Returns null if any calendar can't be read, so nothing
 * appointment-linked is deleted on a bad read. */
function piTest_upcomingEventHashes_() {
  const from = new Date(); from.setDate(from.getDate() - 1); from.setHours(0, 0, 0, 0);
  const to = new Date(from); to.setDate(to.getDate() + 400);
  const out = {ids:{}, byVertical:{}};
  const verticals = ['PreInspection', 'Install', 'Delivery', 'Service'];
  for (let v = 0; v < verticals.length; v++) {
    const name = verticals[v], cfg = TMV3.VERTICALS[name];
    const ids = tmv3_verticalCalendars_(name, cfg).map(function(c) { return c.calendarId; })
      .concat(name === 'PreInspection' ? (cfg.secondaryCalendarIds || []) : []);
    for (let i = 0; i < ids.length; i++) {
      let cal = null;
      try { cal = CalendarApp.getCalendarById(ids[i]); } catch (e) {}
      if (!cal) return null;
      cal.getEvents(from, to).forEach(function(ev) {
        const id = ev.getId();
        out[tmv3_hash_(id)] = true; out.ids[id] = true;
        verticals.forEach(function(vv) { out.byVertical[tmv3_hash_(vv + '|' + id)] = true; });
      });
    }
  }
  return out;
}

function props_exhaustedDate_() {
  try { return PropertiesService.getScriptProperties().getProperty('TMV3_STRIVEN_API_BUDGET_EXHAUSTED_DATE') || ''; } catch (e) { return ''; }
}

/** Group label: collapses hashes, IDs and dates so similar keys show as one line. */
function piTest_propertyGroup_(k) {
  return k.replace(/_[0-9a-f]{16,}/ig, '_*').replace(/_\d{4}-\d{2}-\d{2}/g, '_<date>').replace(/_\d{8}_\d+$/, '_<date>_*')
    .replace(/_\d{4,}$/, '_*').replace(/(_<date>)+/g, '_<date>');
}

/** READ-ONLY: every Script Property grouped by prefix with KEEP / JUNK / OTHER. */
function piTest_PropertyInventory() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const today = Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'yyyy-MM-dd');
  const upcoming = piTest_upcomingEventHashes_();
  if (!upcoming) Logger.log('PI_PROPS_NOTE=calendars unreadable: appointment-linked records kept');
  const groups = {};
  Object.keys(props).forEach(function(k) {
    const c = piTest_propertyClass_(k, props[k], today, upcoming);
    const prefix = piTest_propertyGroup_(k);
    const g = groups[prefix] || (groups[prefix] = {prefix:prefix,count:0,KEEP:0,JUNK:0,OTHER:0,why:{}});
    g.count++; g[c.cls]++; g.why[c.why] = true;
  });
  const list = Object.keys(groups).map(function(k) { const g = groups[k]; g.why = Object.keys(g.why).join('; '); return g; })
    .sort(function(a, b) { return b.count - a.count; });
  const totals = list.reduce(function(t, g) { t.total += g.count; t.KEEP += g.KEEP; t.JUNK += g.JUNK; t.OTHER += g.OTHER; return t; },
    {total:0,KEEP:0,JUNK:0,OTHER:0});
  Logger.log('PI_PROPS_TOTALS=' + JSON.stringify(totals));
  const backfill = {};
  Object.keys(props).forEach(function(k) {
    if (!/^TMV3_SD_ORDER_BACKFILL_20261020_\d+$/.test(k)) return;
    const st = (piTest_json_(props[k]) || {}).state || 'UNKNOWN'; backfill[st] = (backfill[st] || 0) + 1;
  });
  if (Object.keys(backfill).length) Logger.log('PI_PROPS_BACKFILL_STATES=' + JSON.stringify(backfill));
  const unfinished = list.filter(function(g) { return /unfinished write/.test(g.why); });
  if (unfinished.length) Logger.log('PI_PROPS_UNFINISHED=' + JSON.stringify(unfinished.map(function(g) { return g.prefix; })));
  list.forEach(function(g) { Logger.log('PI_PROPS_GROUP=' + JSON.stringify(g)); });
  return {totals:totals,groups:list};
}

/** WRITES (deletes): removes only PreInspection keys classified JUNK. Never touches OTHER or KEEP. */
function piTest_WRITE_CleanupProperties() {
  const store = PropertiesService.getScriptProperties();
  const props = store.getProperties();
  const today = Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'yyyy-MM-dd');
  const upcoming = piTest_upcomingEventHashes_();
  const deleted = {};
  Object.keys(props).forEach(function(k) {
    const c = piTest_propertyClass_(k, props[k], today, upcoming);
    if (c.cls !== 'JUNK') return;
    store.deleteProperty(k);
    const prefix = piTest_propertyGroup_(k);
    deleted[prefix] = (deleted[prefix] || 0) + 1;
  });
  const out = {status:'CLEANED',deleted:deleted,deletedTotal:Object.keys(deleted).reduce(function(n, k) { return n + deleted[k]; }, 0),
    remaining:Object.keys(store.getProperties()).length};
  Logger.log('PI_PROPS_CLEANUP=' + JSON.stringify(out));
  return out;
}


/** Map each held record (unfinished write or duplicate-Task guard) to its appointment and check
 * Striven fresh: RESOLVED when the Task exists, is linked and needs no further action; PAST when the
 * appointment is no longer on any calendar; OPEN otherwise (left alone). Read-only. */
function piTest_heldRecords_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const upcoming = piTest_upcomingEventHashes_();
  if (!upcoming) throw new Error('PI_TEST_CALENDARS_UNREADABLE_NOTHING_CHECKED');
  const reverse = {};
  Object.keys(upcoming.ids).forEach(function(id) {
    ['PreInspection', 'Install', 'Delivery', 'Service'].forEach(function(v) { reverse[tmv3_hash_(v + '|' + id)] = {vertical:v, eventId:id}; });
  });
  const rows = [];
  Object.keys(props).forEach(function(key) {
    const json = piTest_json_(props[key]) || {};
    let target = null, kind = '';
    const ev = /^TMV3_VERTICAL_EVENT_([0-9a-f]{64})$/.exec(key);
    if (ev && ['INTENT','UNCERTAIN'].indexOf(json.state) >= 0) { kind = 'UNFINISHED_WRITE'; target = reverse[ev[1]] || null; }
    else if (/^TMV3_TASK_CREATE_GUARD_/.test(key) && json.eventId) {
      kind = 'TASK_GUARD'; target = upcoming.ids[json.eventId] ? {vertical:json.vertical, eventId:json.eventId} : null;
    } else return;
    const row = {key:key, kind:kind, state:json.state || '', error:String(json.error || '').slice(0, 160),
      vertical:target && target.vertical || json.vertical || '', eventId:target && target.eventId || json.eventId || ''};
    if (!target) { row.verdict = 'PAST'; rows.push(row); return; }
    try {
      const plans = tmv3_step7FreshPlansForEvent_(target.vertical, target.eventId);
      row.plans = plans.map(function(p) { return {taskId:p.taskId || '', plan:p.plan, actions:p.actions || [], blocker:p.blocker || ''}; });
      const done = plans.length > 0 && plans.every(function(p) {
        return p.taskId && !p.blocker && (p.actions || []).every(function(a) { return a === 'VERIFY_CALENDAR_LINKS'; });
      });
      const guardTask = kind === 'TASK_GUARD' && Number(json.taskId || 0);
      row.verdict = done && (!guardTask || plans.some(function(p) { return Number(p.taskId) === guardTask; })) ? 'RESOLVED' : 'OPEN';
    } catch (e) { row.verdict = 'OPEN'; row.checkError = String(e && e.message || e).slice(0, 160); }
    rows.push(row);
  });
  return rows;
}

/** READ-ONLY: list unfinished writes and duplicate-Task guards with a verdict for each. */
function piTest_HeldRecords() {
  const rows = piTest_heldRecords_();
  rows.forEach(function(r) { Logger.log('PI_HELD=' + JSON.stringify(r)); });
  const count = {};
  rows.forEach(function(r) { count[r.verdict] = (count[r.verdict] || 0) + 1; });
  Logger.log('PI_HELD_SUMMARY=' + JSON.stringify(count));
  return rows;
}

/** WRITES (deletes): removes held records whose appointment is past, or whose Task Striven shows
 * exists, is linked and needs nothing more. OPEN records are left for a person to check. */
function piTest_WRITE_ClearResolvedHeld() {
  const store = PropertiesService.getScriptProperties();
  const rows = piTest_heldRecords_();
  const cleared = [];
  rows.forEach(function(r) {
    if (r.verdict !== 'RESOLVED' && r.verdict !== 'PAST') return;
    store.deleteProperty(r.key);
    cleared.push({kind:r.kind, verdict:r.verdict, vertical:r.vertical, eventId:r.eventId});
  });
  const out = {status:'CLEARED', cleared:cleared.length, open:rows.filter(function(r) { return r.verdict === 'OPEN'; }).length,
    remaining:Object.keys(store.getProperties()).length, details:cleared};
  Logger.log('PI_HELD_CLEARED=' + JSON.stringify(out));
  return out;
}

/** READ-ONLY end-to-end check of the live setup. Each line is PASS / WARN / FAIL with what to do.
 * Writes nothing to Striven, Calendar, email or sheets (the slot simulation restores its state). */
function piTest_EndToEndCheck() {
  const rows = [];
  const add = function(area, status, detail, fix) { rows.push({area:area, status:status, detail:detail, fix:fix || ''}); };
  const safe = function(area, fn) { try { fn(); } catch (e) { add(area, 'FAIL', String(e && e.message || e).slice(0, 200), 'See detail'); } };
  const src = function(fn) { try { return String(fn); } catch (e) { return ''; } };

  // 1. Code version: the files pasted into the project are the current ones.
  safe('Code: 30_PreInspection', function() {
    const ok = typeof tmv3_piAddVisitQuestion_ === 'function' && typeof tmv3_piCreatorRecipients_ === 'function' &&
      src(tmv3_piSendMissingDetailsAfterRelease_).indexOf('piOffSharedCandidate === true') >= 0;
    add('Code: 30_PreInspection', ok ? 'PASS' : 'FAIL', ok ? 'current version' : 'older version in the project', 'Paste the latest 30_PreInspection');
  });
  safe('Code: 10_Shared_Core', function() {
    const ok = typeof tmv3_statusNotePut_ === 'function' && typeof TMV3_SD_ORDER_BACKFILL_RETIRED !== 'undefined' &&
      src(tmv3_operationNormalizeCalendarTitle_).indexOf('PROPAGATED_FROM_ORGANIZER_VERIFIED') >= 0;
    add('Code: 10_Shared_Core', ok ? 'PASS' : 'FAIL', ok ? 'current version' : 'older version in the project', 'Paste the latest 10_Shared_Core');
  });
  safe('Code: 40_Installation', function() {
    const ok = src(tmv3_installStandardDailyTriggers).indexOf('tmv3_statusNotePut_') >= 0;
    add('Code: 40_Installation', ok ? 'PASS' : 'WARN', ok ? 'current version' : 'still writes status notes to Script Properties', 'Paste the latest 40_Installation');
  });
  safe('Code: no old operator file', function() {
    const old = typeof piTest_DryRun === 'function' || typeof piTest_WRITE_Reconcile === 'function';
    add('Code: no old operator file', old ? 'FAIL' : 'PASS', old ? 'an old copy of the operator tests is still in the project' : 'one operator file',
      'Delete the old operator block (search every file for "operator test runner")');
  });
  safe('Email template', function() {
    tmv3_piAssertLockedStandard_();
    add('Email template', 'PASS', 'approved template unchanged (fingerprint ' + TMV3_HARD_RULES.PreInspection.notificationRendererFingerprint + ')');
  });

  // 2. Switches.
  safe('Switches', function() {
    const on = function(v) { return v === true ? 'ON' : 'OFF'; };
    const live = TMV3.MODE === 'PRODUCTION_WRITE' && TMV3_PI_CLEANED_RELEASE_APPROVED === true;
    add('Switch: automation writes', live ? 'PASS' : 'FAIL', 'mode ' + TMV3.MODE + ', release gate ' + on(TMV3_PI_CLEANED_RELEASE_APPROVED),
      'Set TMV3_PI_CLEANED_RELEASE_APPROVED = true in 30');
    const mail = TMV3.OPERATIONS.preInspectionMissingDetailEmailsEnabled === true && TMV3_PI_CREATOR_EMAILS_ENABLED === true;
    add('Switch: missing-details emails', mail ? 'PASS' : 'FAIL', 'config ' + on(TMV3.OPERATIONS.preInspectionMissingDetailEmailsEnabled) +
      ', 30 ' + on(TMV3_PI_CREATOR_EMAILS_ENABLED), 'Turn both on');
    add('Switch: emails for bookings missing CF Preinspects', TMV3_PI_OFF_SHARED_NOTICES_ENABLED === true ? 'PASS' : 'WARN', on(TMV3_PI_OFF_SHARED_NOTICES_ENABLED));
    add('Switch: YES/NO box', TMV3_PI_VISIT_QUESTIONS_ENABLED === true ? 'PASS' : 'WARN', on(TMV3_PI_VISIT_QUESTIONS_ENABLED),
      TMV3_PI_VISIT_QUESTIONS_ENABLED === true ? '' : 'Off by choice until approved: set TMV3_PI_VISIT_QUESTIONS_ENABLED = true in 30');
    add('Switch: Steve + Colleen paired', tmv3_piCreatorRecipients_('steve@classicfireplace.ca').length === 2 ? 'PASS' : 'FAIL',
      tmv3_piCreatorRecipients_('steve@classicfireplace.ca').join(', '));
  });

  // 3. Triggers: exactly the standard daily schedule, no duplicates, nothing obsolete.
  safe('Triggers', function() {
    const schedule = tmv3_standardDailySchedule_();
    const expected = [];
    schedule.verticals.forEach(function(v) {
      const prefix = v === 'PreInspection' ? 'preInspection' : v.toLowerCase();
      schedule.hours.forEach(function(h) { expected.push('tmv3_' + prefix + 'Check' + ('0' + h).slice(-2) + '00'); });
    });
    const have = {};
    ScriptApp.getProjectTriggers().forEach(function(t) { const h = t.getHandlerFunction(); have[h] = (have[h] || 0) + 1; });
    const missing = expected.filter(function(h) { return !have[h]; });
    const dupes = Object.keys(have).filter(function(h) { return have[h] > 1 && h !== 'tmv3_scheduledSheetRefresh'; });
    // A pending one-off follow-on run (unfinished sweep) is expected, not extra.
    const extra = Object.keys(have).filter(function(h) { return expected.indexOf(h) < 0 && h !== 'tmv3_scheduledSheetRefresh' &&
      !(h === 'tmv3_preInspectionContinueEligibleSweep' && have[h] === 1 && TMV3_PI_CONTINUE_SWEEP_UNTIL_DONE === true); });
    const pi = ['tmv3_preInspectionCheck0800', 'tmv3_preInspectionCheck1300', 'tmv3_preInspectionCheck1700'].filter(function(h) { return !have[h]; });
    add('Triggers: PreInspection 08:00 / 13:00 / 17:00', pi.length ? 'FAIL' : 'PASS', pi.length ? 'missing ' + pi.join(', ') : 'all three installed',
      'Run tmv3_installStandardDailyTriggers');
    add('Triggers: all workflows', missing.length || dupes.length ? 'WARN' : 'PASS',
      'expected ' + expected.length + ', missing [' + missing.join(', ') + '], duplicated [' + dupes.join(', ') + ']', 'Run tmv3_installStandardDailyTriggers');
    add('Triggers: nothing extra', extra.length ? 'WARN' : 'PASS', extra.length ? 'extra: ' + extra.join(', ') : 'none', 'Delete the extra triggers listed');
    add('Triggers: sheet refresh', (have.tmv3_scheduledSheetRefresh || 0) === 3 ? 'PASS' : 'WARN', (have.tmv3_scheduledSheetRefresh || 0) + ' of 3');
  });

  // 4. Calendars readable.
  safe('Calendars', function() {
    const bad = [];
    let n = 0;
    ['PreInspection', 'Install', 'Delivery', 'Service'].forEach(function(v) {
      const cfg = TMV3.VERTICALS[v];
      const ids = tmv3_verticalCalendars_(v, cfg).map(function(c) { return c.calendarId; })
        .concat(v === 'PreInspection' ? (cfg.secondaryCalendarIds || []) : []);
      ids.forEach(function(id) { n++; let cal = null; try { cal = CalendarApp.getCalendarById(id); } catch (e) {} if (!cal) bad.push(v + ': ' + id); });
    });
    add('Calendars readable', bad.length ? 'FAIL' : 'PASS', bad.length ? bad.join('; ') : n + ' calendars', 'Share the listed calendars with the script owner');
  });

  // 5. YES/NO web app.
  safe('Web app', function() {
    const p = PropertiesService.getScriptProperties();
    const url = p.getProperty('TMV3_PI_CONFIRM_WEBAPP_URL') || '';
    add('Web app URL', /^https:\/\/script\.google\.com\/.+\/exec$/.test(url) ? 'PASS' : 'FAIL', url ? url.replace(/\/s\/[^/]+\//, '/s/…/') : 'not set',
      'Deploy → Web app, then set TMV3_PI_CONFIRM_WEBAPP_URL to the /exec URL');
    add('Web app signing key', p.getProperty('TMV3_PI_CONFIRM_SECRET') ? 'PASS' : 'WARN', p.getProperty('TMV3_PI_CONFIRM_SECRET') ? 'present' : 'created on first send');
    const answers = Object.keys(p.getProperties()).filter(function(k) { return /^TMV3_PI_CONFIRM_TEST_/.test(k); }).length;
    add('Web app click test', answers ? 'PASS' : 'WARN', answers ? answers + ' test answer(s) recorded' : 'no test click recorded', 'Run piTest_SendQuestionTest and click a button');
  });

  // 6. Striven connection and budget.
  safe('Striven', function() {
    const b = tmv3_strivenApiBudgetStatus_();
    add('Striven API budget', b.exhausted ? 'FAIL' : (b.limitMode === 'ENFORCED' && b.remaining < 200) || b.count > 0.8 * b.planDailyLimit ? 'WARN' : 'PASS', 'used ' + b.count + ' today' + (b.limitMode === 'ENFORCED' ? ', remaining ' + b.remaining : ' (tracking only; Striven plan limit ' + b.planDailyLimit + '/day)'),
      b.exhausted ? 'Daily limit reached; resumes after reset' : '');
    const pr = PropertiesService.getScriptProperties();
    add('Striven credentials', pr.getProperty('CLIENT_ID') && pr.getProperty('CLIENT_SECRET') ? 'PASS' : 'FAIL', 'CLIENT_ID / CLIENT_SECRET', 'Restore the Striven credentials');
  });

  // 7. Script Properties stay clean.
  safe('Script Properties', function() {
    const props = PropertiesService.getScriptProperties().getProperties();
    const today = Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'yyyy-MM-dd');
    const upcoming = piTest_upcomingEventHashes_();
    const junk = Object.keys(props).filter(function(k) { return piTest_propertyClass_(k, props[k], today, upcoming).cls === 'JUNK'; }).length;
    add('Script Properties', junk ? 'WARN' : 'PASS', Object.keys(props).length + ' total, ' + junk + ' junk', 'Run piTest_WRITE_CleanupProperties');
  });

  // 8. Dry run of a scheduled slot (no writes, no emails).
  safe('Slot simulation', function() {
    const sim = piTest_SimulateSlot();
    const blocked = (sim.events || []).filter(function(e) { return e.outcome === 'BLOCKED'; }).length;
    add('Slot simulation', sim.status !== 'SIMULATED_NO_WRITES' ? 'FAIL' : blocked ? 'WARN' : 'PASS',
      (sim.events || []).length + ' events, would write ' + sim.wouldWriteTasks + ' Task(s), send ' + sim.wouldSendEmails + ' email(s), alert ' +
      sim.wouldAlert + ', blocked ' + blocked + (sim.error ? ', error ' + sim.error : ''), blocked ? 'See PI_SIM_EVENT lines for the blocked events' : '');
  });

  const count = {PASS:0, WARN:0, FAIL:0};
  rows.forEach(function(r) { count[r.status]++; Logger.log('PI_E2E ' + r.status + ' | ' + r.area + ' | ' + r.detail + (r.status !== 'PASS' && r.fix ? ' | FIX: ' + r.fix : '')); });
  Logger.log('PI_E2E_SUMMARY=' + JSON.stringify(count));
  return {summary:count, rows:rows};
}

function piTest_eventId_() {
  const id = String(piTest_settings_().eventId || '').trim();
  if (!/^[a-z0-9_]+(@google\.com)?$/i.test(id)) throw new Error('PI_TEST_EVENT_ID_INVALID: ' + id);
  return /@google\.com$/i.test(id) ? id : id + '@google.com';
}

/** READ-ONLY: is the event on CF Preinspects, is it in the queue, and what would the automation do. */
function piTest_CheckEvent() {
  const id = piTest_eventId_();
  const cfg = TMV3.VERTICALS.PreInspection;
  const out = {eventId:id};
  const shared = CalendarApp.getCalendarById(cfg.primaryCalendarId);
  const ev = shared && shared.getEventById(id);
  out.onCfPreinspects = !!ev;
  if (ev) {
    out.title = ev.getTitle(); out.start = ev.getStartTime().toISOString();
    out.hasTaskLink = /TaskID=\d+/.test(ev.getDescription() || '');
  } else {
    const stephen = CalendarApp.getCalendarById((cfg.secondaryCalendarIds || [])[0]);
    const s2 = stephen && stephen.getEventById(id);
    if (s2) { out.title = s2.getTitle(); out.start = s2.getStartTime().toISOString(); }
    out.why = 'CF Preinspects is not a guest: automation only emails the booker; no Task or links until it is added.';
  }
  let state = null; try { state = JSON.parse(PropertiesService.getScriptProperties().getProperty('TMV3_PI_SWEEP_STATE') || 'null'); } catch (e) {}
  out.sweep = state ? {status:state.status, remaining:state.remaining, total:state.total} : null;
  out.inPendingQueue = (piTest_json_(PropertiesService.getScriptProperties().getProperty('TMV3_PI_PENDING_EVENTS')) || []).indexOf(id) >= 0;
  if (ev) {
    try { out.dryRun = tmv3_preInspectionDryRun(id); } catch (e) { out.dryRunError = String(e && e.message || e).slice(0, 300); }
  }
  Logger.log('PI_TEST_CHECK_EVENT=' + JSON.stringify(out));
  return out;
}

/** WRITES: reconcile only settings.eventId now — create/link/update its Task with the slot's safety checks.
 * Missing-details emails and failure alerts are not sent from this run. */
function piTest_WRITE_FixEvent() {
  const id = piTest_eventId_();
  const saved = {send:tmv3_preInspectionSendMissingDetails_, alert:tmv3_preInspectionFailureAlert_};
  let r;
  try {
    tmv3_preInspectionSendMissingDetails_ = function() { return {status:'SUPPRESSED_BY_OPERATOR_FIX',sent:0}; };
    tmv3_preInspectionFailureAlert_ = function() { return {status:'SUPPRESSED_BY_OPERATOR_FIX',sent:0}; };
    r = tmv3_preInspectionAutomationCycle_({force:true, eventId:id, reason:'OPERATOR_FIX_EVENT'});
  } finally {
    tmv3_preInspectionSendMissingDetails_ = saved.send; tmv3_preInspectionFailureAlert_ = saved.alert;
  }
  const one = (r.results || [])[0] || {};
  const out = {eventId:id, status:r.status, writes:r.writes, plan:one.plan || '', taskId:one.taskId || '',
    blocker:one.blocker || '', error:one.error || r.error || '', recreate:one.recreateAuthorized || ''};
  if (!r.results || !r.results.length) out.note = 'Event not found on CF Preinspects (or before today): nothing reconciled.';
  try { tmv3_refreshAutomationSheetsAfter_(Date.now() - 600000); } catch (ignored) {}
  Logger.log('PI_TEST_FIX_EVENT=' + JSON.stringify(out));
  return out;
}
