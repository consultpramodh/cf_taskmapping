// Google Apps Script — PreInspection operator tools. Keep ONLY ONE copy of this file in the project.
// Never edit 00_Configuration or 30_PreInspection to test. Run from the editor (no arguments):
//
//   piTest_SimulateSlot()             READ-ONLY: emulates the 08:00/13:00/17:00 trigger for all events
//   piTest_PreviewNotices()           READ-ONLY: who would get a missing-details email now
//   piTest_SendTestEmails()           sends next-slot missing-details emails to Pramodh only, TEST subject
//   piTest_SendQuestionTest()         sends YES/NO pre-inspection questions to Pramodh only, TEST subject
//   piTest_ConfirmationAnswers()      READ-ONLY: recorded YES/NO answers
//   piTest_LastState()                READ-ONLY: last run result and pending queue
//   piTest_PropertyInventory()        READ-ONLY: Script Properties as KEEP / JUNK / OTHER
//   piTest_WRITE_CleanupProperties()  deletes PreInspection JUNK properties only

// Settings live inside a function (no top-level variables), so this file can never
// clash with another file's declarations.
function piTest_settings_() {
  return {
    simulate: {
      slot: '13:00',      // label only; all three slots run the same code
      fromStart: true     // true: evaluate every event; false: continue the real saved sweep position
    },
    emailTo: 'pramodh@classicfireplace.ca'  // test emails go ONLY here
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
    sent.push({subject:message.subject,wouldGoTo:email.recipient,wouldCc:email.cc});
  });
  const out = {status:'TEST_EMAILS_SENT',to:'pramodh@classicfireplace.ca',count:sent.length,emails:sent};
  Logger.log('PI_TEST_EMAILS=' + JSON.stringify(out));
  return out;
}

/** TEST: send the YES/NO question for every unclear Stephen-calendar booking to Pramodh only.
 * Clicking a TEST button records a test answer and never changes the appointment's classification. */
function piTest_SendQuestionTest() {
  if (piTest_settings_().emailTo !== 'pramodh@classicfireplace.ca') throw new Error('PI_TEST_RECIPIENT_CHANGED');
  const saved = tmv3_piSendMissingDetailsAfterRelease_;
  let off;
  try {
    tmv3_piSendMissingDetailsAfterRelease_ = function() { return {status:'SUPPRESSED_FOR_QUESTION_TEST',sent:0}; };
    off = tmv3_piNotifyOffSharedCandidates_({maxLiveChecks:12});
  } finally { tmv3_piSendMissingDetailsAfterRelease_ = saved; }
  const sent = (off.reviewRecords || []).map(function(record) {
    const q = tmv3_piVisitQuestionEmail_(record, true);
    const message = {to:'pramodh@classicfireplace.ca',subject:'TEST ' + q.subject,body:q.body,htmlBody:q.htmlBody,name:'Classic Fireplace & BBQ Store'};
    if (message.to !== 'pramodh@classicfireplace.ca' || message.cc || message.bcc) throw new Error('PI_TEST_RECIPIENT_GUARD');
    MailApp.sendEmail(message);
    return {subject:message.subject,wouldGoTo:q.recipient};
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
    if (upcoming && !upcoming.byVertical[vEvent[1]]) return {cls:'JUNK',why:'write record for a past appointment'};
    return {cls:'KEEP',why:'write record for an upcoming appointment'};
  }
  if (/^TMV3_VERTICAL_ALERT_/.test(key)) return value === 'SENT' ? {cls:'JUNK',why:'failure-alert "sent today" flag'} : {cls:'KEEP',why:'alert being sent'};
  if (/^TMV3_VERTICAL_(?:LAST|CURSOR)_/.test(key) || key === 'TMV3_STANDARD_SCHEDULE')
    return {cls:'JUNK',why:'old status/cursor no code reads'};
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

