// MIGRATION STAGING ONLY. Source: reconstructed V3 3.12.21. NOT APPROVED FOR DEPLOYMENT.

// Original: 30_Calendar; Function: tmv3_extractLegacyInstallOrderNumber_
function tmv3_extractLegacyInstallOrderNumber_(text) {
  const clean = tmv3_stripPhonesForOrderParsing_(text);
  if (!clean) return '';

  let match = clean.match(
    /\b(?:SO|S\/O|Sales\s*Order)\s*[#:\-]?\s*(\d{4,8})(?!\d)/i
  );
  if (match) return match[1];

  match = clean.match(/#\s*(\d{4,8})(?!\d)/);
  return match ? match[1] : '';
}

// Original: 30_Calendar; Function: tmv3_installLegacyCustomerName_
function tmv3_installLegacyCustomerName_(title) {
  return tmv3_clean_(title)
    .replace(/\b(?:SO|S\/O|Sales\s*Order)\s*[#:\-]?\s*\d{4,8}\b/ig, ' ')
    .replace(/#\s*\d{4,8}\b/g, ' ')
    .replace(/\b(?:install|installation|fireplace|bbq|delivery)\b/ig, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Original: 32_Step1_Calendar_View; Function: tmv3_installStep1CalendarLiveSync
function tmv3_installStep1CalendarLiveSync() {
  const handler = 'tmv3_calendarEventUpdated';
  const expected = tmv3_step1CalendarIds_();

  // Stage-1 trigger upgrade rule:
  // remove all V3 managed triggers first so an older version-pinned trigger
  // cannot repopulate the four operator sheets with stale logic.
  const managed = typeof tmv3_managedTriggerHandlers_ === 'function'
    ? tmv3_managedTriggerHandlers_()
    : [handler];

  const removed = [];
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (managed.indexOf(trigger.getHandlerFunction()) === -1) return;

    removed.push({
      handler: trigger.getHandlerFunction(),
      source: String(trigger.getTriggerSource()),
      sourceId: tmv3_safeCalendar_(function() {
        return trigger.getTriggerSourceId();
      }, '')
    });

    ScriptApp.deleteTrigger(trigger);
  });

  const created = [];

  expected.forEach(function(calendarId) {
    ScriptApp.newTrigger(handler)
      .forUserCalendar(calendarId)
      .onEventUpdated()
      .create();

    created.push(calendarId);
  });

  const fallbackTrigger = ScriptApp.newTrigger('tmv3_calendarReconciliationFallback')
    .timeBased()
    .everyHours(2)
    .create();

  const initial =
    typeof tmv3_calendarStageRefresh_ === 'function'
      ? tmv3_calendarStageRefresh_('LIVE_SYNC_INSTALL_REFRESHED')
      : tmv3_step1CalendarRun('LIVE_SYNC_INSTALL_REFRESHED');

  const result = {
    status: initial.status === 'PASS' ? 'INSTALLED' : 'INSTALLED_WITH_REVIEW',
    mode: 'ONE_WAY_CALENDAR_TO_V3_SHEETS',
    executionStage: tmv3_executionStage_(),
    handler: handler,
    calendarsExpected: expected.length,
    removedManagedTriggers: removed,
    createdCalendarTriggers: created,
    reconciliationFallbackCreated: !!fallbackTrigger,
    calendarWritesPerformed: false,
    initialSync: initial,
    triggerStatus: tmv3_step1CalendarLiveSyncStatus()
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Original: 46_Step7_Reconciliation; Function: tmv3_step7InstallReconciliationRun
function tmv3_step7InstallReconciliationRun(reason) {
  return tmv3_step7ReconciliationRun(
    reason || 'INSTALL_STEP7_VERIFY',
    true,
    'Install'
  );
}

// Original: 58_Operations; Function: tmv3_installReminderAlreadySentToday_
function tmv3_installReminderAlreadySentToday_(eventId) {
  const key = 'TMV3_INSTALL_SO_REMINDER_' + tmv3_hash_([
    eventId,
    tmv3_date_(new Date())
  ].join('|'));
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

// Original: 58_Operations; Function: tmv3_markInstallReminderSent_
function tmv3_markInstallReminderSent_(eventId) {
  const key = 'TMV3_INSTALL_SO_REMINDER_' + tmv3_hash_([
    eventId,
    tmv3_date_(new Date())
  ].join('|'));
  PropertiesService.getScriptProperties().setProperty(key, tmv3_iso_(new Date()));
}

// Original: 58_Operations; Function: tmv3_sendInstallMissingSoReminders_
function tmv3_sendInstallMissingSoReminders_(scope) {
  const policy = tmv3_operationPolicy_();
  if (policy.installRemindersEnabled !== true) {
    return { status: 'REMINDERS_GATED', checked: 0, sent: 0 };
  }

  const today = new Date();
  if (today.getDay() === 0 || today.getDay() === 6) {
    return { status: 'WEEKEND', checked: 0, sent: 0 };
  }

  const cfg = TMV3.VERTICALS.Install;
  const cal = CalendarApp.getCalendarById(cfg.calendarIds[0]);
  if (!cal) throw new Error('Install Calendar is unavailable.');

  const start = new Date();
  start.setHours(0,0,0,0);
  const end = new Date();
  end.setDate(end.getDate() + 7);
  end.setHours(23,59,59,999);

  let checked = 0;
  let sent = 0;
  let skipped = 0;
  let errors = 0;

  cal.getEvents(start, end).forEach(function(event) {
    const record = tmv3_calendarEvent_(
      'Install',
      cfg,
      { calendarId: cfg.calendarIds[0] },
      event
    );
    if (!record) return;
    checked++;

    if (record.orderNumber || record.existingOrderId) {
      skipped++;
      return;
    }

    const businessDays = tmv3_businessDaysUntil_(record.start);
    if (businessDays === null || businessDays < 0 || businessDays > 3) {
      skipped++;
      return;
    }

    if (tmv3_installReminderAlreadySentToday_(record.eventId)) {
      skipped++;
      return;
    }

    const creator=tmv3_preInspectionCreatorEmail_(record);
    const recipients=creator?[creator]:[];
    if (!recipients.length) {
      skipped++;
      tmv3_audit_('Install',record.eventId,'','MISSING_SO_REMINDER','ATTENTION','No valid recipient found.');
      return;
    }

    const when = tmv3_date_(record.start) + ' ' + tmv3_time_(record.start) +
      '–' + tmv3_time_(record.end);
    const subject = tmv3_brandSubject_('Install', 'ACTION REQUIRED: Sales Order # missing from install event');
    const textBody =
      'ACTION REQUIRED\n\n' +
      'The install calendar event below is missing the Sales Order #.\n\n' +
      'Event: ' + (record.title || '(no title)') + '\n' +
      'When: ' + when + '\n' +
      'Where: ' + (record.location || 'n/a') + '\n\n' +
      'Next step: Add the correct Sales Order # to the Calendar Event.';

    if(MailApp.getRemainingDailyQuota()<1){errors++;return;}
    const reminderKey='TMV3_INSTALL_SO_REMINDER_'+tmv3_hash_([record.eventId,tmv3_date_(new Date())].join('|'));
    // A timeout is an uncertain send; the durable intent prevents same-day replay.
    PropertiesService.getScriptProperties().setProperty(reminderKey,'SENDING');
    try {
      const greeting=tmv3_preInspectionCreatorFirstName_(creator);
      const esc=tmv3_preInspectionHtmlEscape_;
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: subject,
        body:'This is an automated email notification from Classic Fireplace & BBQ Store.\n\nHi '+greeting+',\n\n'+textBody,
        htmlBody:tmv3_brandNotification_('Install',{htmlBody:'x'}).htmlBody.replace(/x$/,'')+'<p><em>This is an automated email notification from Classic Fireplace &amp; BBQ Store.</em></p>'+
          '<p>Hi '+esc(greeting)+',</p><p style="color:#B42318;font-weight:700;">Sales Order Number missing</p>'+
          '<p>'+esc(record.title || '(no title)')+'<br>'+esc(when)+'<br>'+esc(record.location || '')+'</p>'+
          '<p>Please add the correct Sales Order number to the appointment.</p>'+
          '<p><a href="'+esc(tmv3_calendarEventUrl_(record.eventId,record.calendarId || cfg.calendarIds[0]))+'">Open calendar appointment</a></p>',
        name:'Classic Fireplace & BBQ Store'
      });
      tmv3_markInstallReminderSent_(record.eventId);
      sent++;
      tmv3_audit_('Install',record.eventId,'','MISSING_SO_REMINDER','PASS','Sent to ' + recipients.join(','));
    } catch (err) {
      errors++;
      tmv3_audit_('Install',record.eventId,'','MISSING_SO_REMINDER','ATTENTION',String(err && err.message || err));
    }
  });

  return {
    status: errors ? 'COMPLETE_WITH_ERRORS' : 'COMPLETE',
    scope: scope,
    checked: checked,
    sent: sent,
    skipped: skipped,
    errors: errors
  };
}

// Install notification — same standard format and dynamic behaviour as the PreInspection
// missing-details email (operator rule 2026-11-13). One email per install listing ONLY what is
// wrong right now: Sales Order # missing, Sales Order not Approved (Quoted / Pending Approval /
// Incomplete), Sales Order not found in Striven, or job-site address missing. First email and
// follow-ups use the same timing rules as PreInspection (tmv3_preInspectionNotificationDecision_).
// While ON, it replaces the older plain "Sales Order # missing" reminder.
var TMV3_INSTALL_NOTICE_ENABLED = false;   // ON after the TEST email is approved
const TMV3_INSTALL_SO_NOT_APPROVED = Object.freeze({'18':'incomplete','19':'quoted','20':'pending approval'});
const TMV3_INSTALL_NOTICE_LOOKAHEAD_DAYS = 31;
// One-off special projects: booked without a Sales Order on purpose (operator 2026-11-13).
// Their Install notice never asks for an SO #.
const TMV3_INSTALL_SPECIAL_PROJECTS = Object.freeze([/^\s*Times\s*-\s*Chatsworth\b/i]);

/** Fresh Striven status of the SO on an Install event: {status:'FOUND', orderId, number, statusId, statusName} or a reason. */
function tmv3_installFreshOrderStatus_(record) {
  const number = String(tmv3_clean_(record && record.orderNumber));
  if (!/^\d{4,8}$/.test(number)) return {status:'NO_VALID_SO_NUMBER',number:number};
  const query = {PageIndex:0,PageSize:2,Number:number};
  const cid = Number(record.customerNumber || 0);
  if (Number.isSafeInteger(cid) && cid > 0) query.CustomerId = cid;
  const response = tmv3_fetchJson_(TMV3.API_BASE + '/v1/sales-orders/search',
    {method:'post',contentType:'application/json',payload:JSON.stringify(query)});
  const rows = response && (response.Data || response.data);
  const count = Number(response && (response.TotalCount !== undefined ? response.TotalCount : response.totalCount));
  if (!Array.isArray(rows) || count !== 1 || rows.length !== 1 ||
      String(tmv3_first_(rows[0], ['Number','OrderNumber'])) !== number) return {status:'SO_NOT_FOUND_OR_NOT_UNIQUE',number:number};
  const orderId = Number(tmv3_first_(rows[0], ['Id','id']));
  const order = tmv3_fetchJson_(TMV3.API_BASE + '/v1/sales-orders/' + orderId, {method:'get'}) || {};
  if (Number(tmv3_first_(order, ['Id','id'])) !== orderId) return {status:'SO_READ_MISMATCH',number:number};
  const st = tmv3_first_(order, ['Status','status']) || {};
  return {status:'FOUND',orderId:orderId,number:number,
    statusId:String(tmv3_first_(st, ['Id','id']) || ''),statusName:tmv3_norm_(tmv3_first_(st, ['Name','name']) || '')};
}

/** What is wrong with one install event right now (dynamic list, like the PreInspection assessment). */
function tmv3_installNoticeAssessment_(record) {
  const missing = [];
  let order = null;
  // Sales Orders are 6 digits. A 5-digit number (e.g. "#59298") is a customer #, not an SO.
  // Also read "SO585380-…" / "SO 585380" written straight after other text in the title or notes.
  let so = String(tmv3_clean_(record.orderNumber));
  if (!/^\d{6}$/.test(so)) {
    const m = /SO\s*#?\s*(\d{6})(?!\d)/.exec(String(record.title || '') + ' ' + String(record.descriptionClean || record.rawDescription || ''));
    so = m ? m[1] : '';
  }
  record = Object.assign({}, record, {orderNumber:so});
  const special = TMV3_INSTALL_SPECIAL_PROJECTS.some(function(rx) { return rx.test(String(record.title || '')); });
  if (!so && special) {
    // Special project: no Sales Order by design.
  } else if (!so) {
    missing.push({code:'SALES_ORDER',field:'Sales Order',message:'Sales Order # missing',now:'Sales Order # missing',should:'Sales Order # of the Approved order'});
  } else {
    order = tmv3_installFreshOrderStatus_(record);
    if (order.status === 'FOUND') {
      const label = TMV3_INSTALL_SO_NOT_APPROVED[order.statusId];
      // ID and name must agree, so a renamed status never sends an email.
      if (label && label === order.statusName) {
        const shown = label.replace(/\b\w/g, function(c) { return c.toUpperCase(); });
        missing.push({code:'SO_NOT_APPROVED',field:'Sales Order',message:'Sales Order #' + order.number + ' is ' + shown + ', not Approved',
          now:'SO #' + order.number + ' — ' + shown,should:'SO #' + order.number + ' — Approved'});
      }
    } else if (order.status === 'SO_NOT_FOUND_OR_NOT_UNIQUE' || order.status === 'NO_VALID_SO_NUMBER') {
      missing.push({code:'SO_NOT_FOUND',field:'Sales Order',message:'Sales Order #' + (order.number || '') + ' not found in Striven',
        now:'SO #' + (order.number || '?') + ' — not found',should:'The correct Sales Order #'});
    }
  }
  const streetInNotes = /\b\d{1,6}\s+[A-Za-z][\w'.-]*(?:\s+[A-Za-z][\w'.-]*){0,3}\s+(?:St|Street|Ave|Avenue|Rd|Road|Dr|Drive|Cres|Crescent|Blvd|Boulevard|Ct|Court|Ln|Lane|Way|Pl|Place|Trail|Terrace|Circle|Cir|Pkwy|Parkway|Gate|Hwy|Highway|Line|Sideroad)\b/i
    .test(String(record.descriptionClean || record.rawDescription || ''));
  if (!tmv3_clean_(record.location) && !streetInNotes) {
    missing.push({code:'LOCATION',field:'Location',message:'Job-site address missing',now:'Location missing',should:'Full Job-Site Address'});
  }
  const creatorEmail = tmv3_preInspectionCreatorEmail_(record);
  return {missing:missing,order:order,creatorEmail:creatorEmail,shouldNotify:missing.length > 0};
}

/** The Install email in the standard notification format (same layout as PreInspection). */
function tmv3_installNoticeEmail_(record, assessment, followupNumber) {
  const esc = tmv3_preInspectionHtmlEscape_;
  const firstName = tmv3_preInspectionCreatorFirstName_(assessment.creatorEmail);
  const when = Utilities.formatDate(new Date(record.start), TMV3_TIMEZONE, 'MMM d, yyyy') + ' · ' +
    Utilities.formatDate(new Date(record.start), TMV3_TIMEZONE, 'h:mm a') + ' – ' + Utilities.formatDate(new Date(record.end), TMV3_TIMEZONE, 'h:mm a');
  const followupLabel = followupNumber ? 'FOLLOW-UP #' + followupNumber : '';
  const items = assessment.missing.map(function(m) {
    return '<div style="color:#B42318;font-weight:700;">❌ ' + esc(m.message) + '</div>'; }).join('');
  const rows = assessment.missing.map(function(m) {
    return '<tr><td valign="top" style="padding:12px;border:1px solid #dddddd;font-weight:700;">' + esc(m.field) +
      '</td><td valign="top" style="padding:12px;border:1px solid #dddddd;"><strong style="color:#B42318;">' + esc(m.now) +
      '</strong></td><td valign="top" style="padding:12px;border:1px solid #dddddd;color:#256b35;font-weight:700;">' + esc(m.should) + '</td></tr>';
  }).join('');
  const order = assessment.order;
  const soLink = order && order.status === 'FOUND' ? '<div style="font-size:14px;margin-bottom:14px;"><a href="' +
    esc('https://classicfireplace.striven.com/next/crm#/sales-orders/' + order.orderId) + '">Open Sales Order #' + esc(order.number) + '</a></div>' : '';
  const html =
    '<!doctype html><html lang="en"><body style="margin:0;padding:0;background:#f4f3f1;font-family:Arial,Helvetica,sans-serif;color:#222;">' +
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0"><tr><td align="center" style="padding:20px 12px;">' +
    '<table role="presentation" width="760" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:760px;background:#ffffff;">' +
    '<tr><td style="background:#202020;border-top:4px solid #b42b2f;padding:18px 24px;">' +
      '<div style="font-size:18px;font-weight:700;color:#ffffff;">' + (followupLabel ? followupLabel + ' — ' : '') + 'Install Update Required</div>' +
      '<div style="font-size:12px;color:#aaaaaa;margin-top:3px;">Classic Fireplace &amp; BBQ Store</div>' +
    '</td></tr>' +
    '<tr><td style="padding:20px 24px 12px;">' +
      '<div style="font-size:12px;color:#666;font-style:italic;margin-bottom:16px;">This is an automated email notification from Classic Fireplace &amp; BBQ Store.</div>' +
      '<div style="font-size:14px;line-height:21px;">Hi <strong>' + esc(firstName) + '</strong>,</div>' +
      '<div style="margin-top:12px;padding:12px;background:#f6f5f3;border:1px solid #e2e0dc;"><strong>Install requiring attention</strong><br>' +
        esc(record.title || '(Untitled event)') + '<br><strong>' + esc(when) + '</strong><br>' + esc(record.location || 'Location missing') + '</div>' +
      '<div style="font-size:14px;line-height:21px;margin-top:8px;">This install needs the following before it can go ahead. Please update the items highlighted below.</div>' +
      '<div style="margin-top:12px;background:#fff5f5;border-left:4px solid #b42b2f;padding:11px 13px;font-size:13px;line-height:22px;color:#8f2024;">' + items + '</div>' +
    '</td></tr>' +
    '<tr><td style="padding:0 24px 14px;"><table width="100%" cellspacing="0" cellpadding="0" border="0" style="background:#f6f5f3;border:1px solid #e2e0dc;"><tr><td style="padding:12px 14px;">' +
      '<div style="font-size:15px;font-weight:700;">' + esc(record.title || '(Untitled event)') + '</div>' +
      '<div style="font-size:12px;color:#666;margin-top:4px;">' + esc(when) + '</div>' +
      (record.location ? '<div style="font-size:12px;color:#666;margin-top:4px;">' + esc(record.location) + '</div>' : '') +
    '</td></tr></table></td></tr>' +
    '<tr><td style="padding:0 24px 20px;"><table width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;border:1px solid #dddddd;font-size:13px;">' +
      '<tr><td width="18%" style="background:#f3f3f3;padding:10px 12px;border-right:1px solid #dddddd;border-bottom:1px solid #dddddd;"></td>' +
        '<td width="41%" style="background:#fff1f1;color:#9b1c1f;font-weight:700;padding:10px 12px;border-right:1px solid #dddddd;border-bottom:1px solid #dddddd;">What it is now</td>' +
        '<td width="41%" style="background:#f1f8f1;color:#28632f;font-weight:700;padding:10px 12px;border-bottom:1px solid #dddddd;">What it should be</td></tr>' +
      rows +
    '</table></td></tr>' +
    '<tr><td style="padding:0 24px 22px;">' + soLink +
      '<div style="font-size:14px;margin-bottom:14px;"><a href="' +
      esc(tmv3_calendarEventUrl_(record.eventId, record.calendarId || TMV3.VERTICALS.Install.calendarIds[0])) + '">Open calendar appointment</a></div>' +
      '<div style="font-size:14px;margin-bottom:14px;">Please get the Sales Order approved in Striven (or update the calendar event) before the install.</div>' +
      '<div style="font-size:14px;">Thank you.</div></td></tr>' +
    '<tr><td style="background:#202020;padding:12px 24px;text-align:center;"><div style="font-size:12px;font-weight:700;color:#ffffff;">Classic Fireplace &amp; BBQ Store</div></td></tr>' +
    '</table></td></tr></table></body></html>';
  return tmv3_brandNotification_('Install', {
    recipient:assessment.creatorEmail,
    cc:tmv3_preInspectionNotificationCc_(),
    subject:(followupLabel ? followupLabel + ' | ' : '') + 'INSTALL DETAILS REQUIRED: ' + (record.calendarCustomerName || record.title || 'Install') +
      ' | ' + when + ' | ' + assessment.missing.map(function(m) { return m.code === 'SO_NOT_APPROVED' ? 'SO not Approved' : m.code === 'SO_NOT_FOUND' ? 'SO not found' : m.field; })
        .filter(function(x, i, a) { return a.indexOf(x) === i; }).join(', '),
    body:'Install update required: ' + assessment.missing.map(function(m) { return m.message; }).join('; '),
    htmlBody:html
  });
}

/** Install events in the lookahead window, each with its assessment (Striven read only when an SO # is present). */
function tmv3_installNoticeRecords_() {
  const cfg = TMV3.VERTICALS.Install;
  const cal = CalendarApp.getCalendarById(cfg.calendarIds[0]);
  if (!cal) throw new Error('Install Calendar is unavailable.');
  const start = new Date(); start.setHours(0,0,0,0);
  const end = new Date(start); end.setDate(end.getDate() + TMV3_INSTALL_NOTICE_LOOKAHEAD_DAYS);
  const out = [];
  cal.getEvents(start, end).forEach(function(event) {
    if (event.isAllDayEvent && event.isAllDayEvent()) return;
    let record = null;
    try { record = tmv3_calendarEvent_('Install', cfg, {calendarId:cfg.calendarIds[0]}, event); } catch (e) { return; }
    if (!record || record.status === 'cancelled') return;
    let assessment;
    try { assessment = tmv3_installNoticeAssessment_(record); }
    catch (err) { out.push({record:record,error:String(err && err.message || err).slice(0, 160)}); return; }
    out.push({record:record,assessment:assessment});
  });
  return out;
}

/** Send: same first-email / follow-up timing as PreInspection; Steve + Colleen paired like PreInspection. */
function tmv3_sendInstallNotices_(scope) {
  if (TMV3_INSTALL_NOTICE_ENABLED !== true) return {status:'INSTALL_NOTICES_OFF',sent:0};
  if (TMV3.MODE !== 'PRODUCTION_WRITE' || TMV3.OPERATIONS.installRemindersEnabled !== true) return {status:'REMINDERS_GATED',sent:0};
  const props = PropertiesService.getScriptProperties();
  const out = {status:'COMPLETE',scope:scope,checked:0,needAttention:0,sent:0,errors:0,events:[]};
  tmv3_installNoticeRecords_().forEach(function(x) {
    out.checked++;
    if (x.error) { out.events.push({eventId:x.record.eventId,status:'SO_READ_FAILED',error:x.error}); return; }
    if (!x.assessment.shouldNotify) return;
    out.needAttention++;
    const record = x.record, key = 'TMV3_INSTALL_NOTICE_' + tmv3_hash_(record.eventId);
    let state = null;
    try { state = JSON.parse(props.getProperty(key) || 'null'); } catch (e) { out.events.push({eventId:record.eventId,status:'UNCERTAIN_SEND_NO_RETRY'}); return; }
    const decision = tmv3_preInspectionNotificationDecision_(record, state);
    if (!decision.send) { out.events.push({eventId:record.eventId,status:decision.status}); return; }
    const email = tmv3_installNoticeEmail_(record, x.assessment, decision.followupNumber);
    if (!email.recipient) { out.events.push({eventId:record.eventId,status:'CREATOR_UNVERIFIED'}); return; }
    if (MailApp.getRemainingDailyQuota() < 3) { out.errors++; return; }
    const to = tmv3_piCreatorRecipients_(email.recipient);
    const cc = (email.cc || []).filter(function(c) { return to.indexOf(String(c).toLowerCase()) < 0; });
    const next = {status:'SENDING',sentCount:Number(state && state.sentCount || 0) + 1,lastSentAt:new Date().toISOString()};
    props.setProperty(key, JSON.stringify(next));
    try {
      MailApp.sendEmail({to:to.join(','),cc:cc.join(','),subject:email.subject,body:email.body,htmlBody:email.htmlBody,name:'Classic Fireplace & BBQ Store'});
      next.status = 'SENT'; props.setProperty(key, JSON.stringify(next));
      out.sent++;
      out.events.push({eventId:record.eventId,status:'SENT',to:to.join(','),followup:decision.followupNumber});
      tmv3_audit_('Install', record.eventId, '', 'INSTALL_NOTICE_EMAIL', 'SENT',
        JSON.stringify({to:to,followupNumber:decision.followupNumber,missing:x.assessment.missing.map(function(m) { return m.code; })}));
    } catch (err) {
      out.errors++;
      tmv3_audit_('Install', record.eventId, '', 'INSTALL_NOTICE_EMAIL', 'ATTENTION', String(err && err.message || err));
    }
  });
  if (out.errors) out.status = 'COMPLETE_WITH_ERRORS';
  return out;
}

// Original: 58_Operations; Function: tmv3_sendInstallMissingSoRemindersNow
function tmv3_sendInstallMissingSoRemindersNow() {
  return tmv3_sendInstallMissingSoReminders_('MANUAL');
}

// Original: 60_Automation; Function: tmv3_installTriggers
function tmv3_installTriggers() { return tmv3_installStandardDailyTriggers(); }

// Original: 60_Automation; Function: tmv3_installReminderCheck
function tmv3_installReminderCheck() {
  if (TMV3.MODE!=='PRODUCTION_WRITE' || TMV3.OPERATIONS.installRemindersEnabled!==true) {
    return {status:'REMINDERS_GATED',sent:0};
  }
  // The standard Install notice (dynamic, all issues in one email) replaces the plain missing-SO reminder.
  if (TMV3_INSTALL_NOTICE_ENABLED === true) {
    try { return tmv3_sendInstallNotices_('AUTO'); }
    catch (err) { return {status:'ATTENTION',error:String(err && err.message || err).slice(0,200)}; }
  }
  return tmv3_sendInstallMissingSoReminders_('AUTO');
}

// Original: 60_Automation; Function: tmv3_runInstallAutomation
function tmv3_runInstallAutomation() {const r=tmv3_runFullVerticalSweep_('Install');Logger.log(JSON.stringify(r));return r;}

// Original: 60_Automation; Function: tmv3_installVerticalProductionTrigger
function tmv3_installVerticalProductionTrigger() { return tmv3_installStandardDailyTriggers(); }

// Original: 60_Automation; Function: tmv3_installAutomationSheetRefreshTrigger
function tmv3_installAutomationSheetRefreshTrigger() {
  const handler='tmv3_scheduledSheetRefresh';
  const existing=ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()===handler;});
  if(ScriptApp.getProjectTriggers().length-existing.length+3>20)throw Error('SHEET_REFRESH_TRIGGER_CAPACITY');
  existing.forEach(function(t){ScriptApp.deleteTrigger(t);});
  [8,13,17].forEach(function(hour){ScriptApp.newTrigger(handler).timeBased().atHour(hour).nearMinute(0).everyDays(1).inTimezone('America/Toronto').create();});
  const triggers=tmv3_listTriggers();
  if(triggers.filter(function(t){return t.handler===handler;}).length!==3)throw Error('SHEET_REFRESH_TRIGGER_READBACK_FAILED');
  return {status:'SHEET_REFRESH_TRIGGER_VERIFIED',triggers:triggers};
}

// Original: 60_Automation; Function: tmv3_installCheck0800
function tmv3_installCheck0800() { return tmv3_standardVerticalSlot_('Install','08:00'); }

// Original: 60_Automation; Function: tmv3_installCheck1300
function tmv3_installCheck1300() { return tmv3_standardVerticalSlot_('Install','13:00'); }

// Original: 60_Automation; Function: tmv3_installCheck1700
function tmv3_installCheck1700() { return tmv3_standardVerticalSlot_('Install','17:00'); }

// Original: 60_Automation; Function: tmv3_installStandardDailyTriggers
function tmv3_installStandardDailyTriggers() {
  const schedule=tmv3_standardDailySchedule_();
  const handlers=[];
  schedule.verticals.forEach(function(v){
    const prefix=v==='PreInspection'?'preInspection':v.toLowerCase();
    schedule.hours.forEach(function(hour){handlers.push({handler:'tmv3_'+prefix+'Check'+('0'+hour).slice(-2)+'00',hour:hour});});
  });
  const before=ScriptApp.getProjectTriggers();
  const obsolete=['tmv3_dailySourceRefresh','tmv3_scheduledSheetRefresh',
    'tmv3_preInspectionContinueEligibleSweep','tmv3_installSweepContinue',
    'tmv3_deliverySweepContinue','tmv3_serviceSweepContinue',
    'tmv3_allVerticalAutomationTick','tmv3_scheduledOperations',
    'tmv3_calendarEventUpdated','tmv3_calendarReconciliationFallback',
    'tmv3_installReminderCheck','tmv3_installPublishView',
    'tmv3_serviceDeliveryOrderBackfill'];
  const removable=before.filter(function(t){return obsolete.indexOf(t.getHandlerFunction())>=0;});
  const retained=before.filter(function(t){return removable.indexOf(t)<0;});
  const missing=handlers.filter(function(spec){return !retained.some(function(t){return t.getHandlerFunction()===spec.handler;});});
  if(retained.length+missing.length+3>20)throw Error('STANDARD_SCHEDULE_CAPACITY_NO_CHANGE');
  removable.forEach(function(t){ScriptApp.deleteTrigger(t);});
  missing.forEach(function(spec){
    ScriptApp.newTrigger(spec.handler).timeBased().atHour(spec.hour).nearMinute(0).everyDays(1).inTimezone(schedule.timezone).create();
  });
  schedule.hours.forEach(function(hour){ScriptApp.newTrigger('tmv3_scheduledSheetRefresh').timeBased().atHour(hour).nearMinute(0).everyDays(1).inTimezone(schedule.timezone).create();});
  const triggers=tmv3_listTriggers();
  const expected=handlers.map(function(x){return x.handler;});
  if(expected.some(function(h){return triggers.filter(function(t){return t.handler===h;}).length!==1;}) ||
    triggers.filter(function(t){return t.handler==='tmv3_scheduledSheetRefresh';}).length!==3 ||
    triggers.some(function(t){return obsolete.indexOf(t.handler)>=0 && t.handler!=='tmv3_scheduledSheetRefresh';}))throw Error('STANDARD_SCHEDULE_READBACK_FAILED');
  const result={status:'STANDARD_DAILY_TRIGGERS_VERIFIED',schedule:schedule,triggers:triggers,businessTriggerCount:12,maintenanceTriggerCount:3};
  tmv3_statusNotePut_('TMV3_STANDARD_SCHEDULE',JSON.stringify(result));
  return result;
}

// Original: 60_Automation; Function: tmv3_installSweepContinue
function tmv3_installSweepContinue(){
  // Follow-on run: continues an unfinished Install sweep a minute after a slot ran out of time.
  tmv3_verticalClearFollowOn_('Install');
  if(!TMV3_VERTICAL_FOLLOW_ON.Install)return {status:'RETIRED_OUTSIDE_STANDARD_SLOTS',writes:0};
  const result=tmv3_runFullVerticalSweep_('Install',{followOn:true});
  tmv3_audit_('Install','','','STANDARD_DAILY_SLOT',result.status,JSON.stringify({slot:'FOLLOW_ON',result:result}));
  return result;
}

// Original: 60_Automation; Function: tmv3_installResolveContactFresh_
function tmv3_installResolveContactFresh_(record,customerId,preferredContactId,plan) {
  const cid=tmv3_clean_(customerId);
  if(!/^[1-9]\d*$/.test(cid))throw Error('INSTALL_CONTACT_CUSTOMER_REQUIRED');
  const existing=plan && plan.actualRequestedByType==='contact' ? plan.actualRequestedById : '';
  const known=tmv3_clean_(preferredContactId || existing);
  if(known) {
    const proof=tmv3_getContactById_(known,cid);
    if(!proof || proof.__ownershipVerified!==true)throw Error('INSTALL_CONTACT_OWNERSHIP_UNPROVEN');
    return known;
  }
  const authored=typeof tmv3_identityDescriptionForRecord_==='function'?tmv3_identityDescriptionForRecord_(record):'';
  const phone=tmv3_phone10_(record.phone || '');
  const emails=tmv3_extractEmails_(authored);
  if(!phone && !emails.length)throw Error('INSTALL_CONTACT_EVIDENCE_MISSING');
  const contacts=tmv3_getCustomerContacts_(cid,{forceFresh:true,requireComplete:true});
  const owned=contacts.filter(function(c){return tmv3_clean_(c['Customer ID'])===cid;});
  const phoneIds=[],emailIds=[];
  owned.forEach(function(c){
    const id=tmv3_clean_(c['Contact ID']);
    if(!/^[1-9]\d*$/.test(id))return;
    const phones=(c.Phones||[c.Phone||'']).map(tmv3_phone10_);
    const addresses=(c.Emails||[c.Email||'']).map(tmv3_normEmail_);
    if(phone && phones.indexOf(phone)>=0 && phoneIds.indexOf(id)<0)phoneIds.push(id);
    if(emails.some(function(e){return addresses.indexOf(tmv3_normEmail_(e))>=0;}) && emailIds.indexOf(id)<0)emailIds.push(id);
  });
  const matches=phoneIds.length && emailIds.length ? phoneIds.filter(function(id){return emailIds.indexOf(id)>=0;}) : (phoneIds.length?phoneIds:emailIds);
  if(matches.length!==1)throw Error(matches.length?'INSTALL_CONTACT_AMBIGUOUS':'INSTALL_CONTACT_NO_CORROBORATED_MATCH');
  const proof=tmv3_getContactById_(matches[0],cid);
  if(!proof || proof.__ownershipVerified!==true)throw Error('INSTALL_CONTACT_OWNERSHIP_UNPROVEN');
  return matches[0];
}

// Original: 60_Automation; Function: tmv3_installRefreshView_
function tmv3_installRefreshView_(reason) {
  const cfg=TMV3.VERTICALS.Install,start=new Date(),end=new Date();
  start.setDate(start.getDate()-Number(cfg.lookbackDays||0));start.setHours(0,0,0,0);
  end.setDate(end.getDate()+Number(cfg.lookaheadDays||365));end.setHours(23,59,59,999);
  const input=[];
  tmv3_verticalCalendars_('Install',cfg).forEach(function(calCfg){
    const calendar=CalendarApp.getCalendarById(calCfg.calendarId);if(!calendar)throw Error('INSTALL_CALENDAR_UNAVAILABLE');
    const resolved=Object.assign({},calCfg,{calendarName:calendar.getName()});
    calendar.getEvents(start,end).forEach(function(event){const record=tmv3_calendarEvent_('Install',cfg,resolved,event,{applyEligibility:false});if(record)input.push(record);});
  });
  const records=tmv3_mergeLogicalCalendarRecords_(input);
  const stage2=tmv3_step2CalendarRecords_(records);
  stage2.forEach(function(r){if(!tmv3_installCreatorAllowed_(r))r.step2=tmv3_step2Decision_('SKIP','INSTALL_CREATOR_UNVERIFIED','Internal Calendar creator could not be verified.',false);});
  const stage3=tmv3_step3BusinessAnchorRecords_(stage2,tmv3_step3AnchorIndex_());
  const stage4=tmv3_step4IdentityRecords_(stage3,tmv3_step4IdentityIndex_());
  const stage5=tmv3_step5TaskRecords_(stage4,tmv3_step5TaskIndex_());
  const decisions=tmv3_step6DecisionRecords_(stage5).sort(function(a,b){return new Date(a.start)-new Date(b.start);});
  const prior=tmv3_sheet_(cfg.sheet).getLastRow();
  tmv3_replaceRows_(cfg.sheet,TMV3_STEP1_HEADERS.slice(),decisions.map(tmv3_step6OperatorRow_));
  tmv3_step1FormatOperatorView_(cfg.sheet,'Install',decisions,prior);
  const count=tmv3_rows_(cfg.sheet).filter(function(r){return !!tmv3_clean_(r['Event ID']);}).length;
  if(count!==decisions.length)throw Error('INSTALL_VIEW_ROW_READBACK_FAILED');
  const result={status:'INSTALL_VIEW_REFRESHED',at:new Date().toISOString(),reason:reason,rows:count,strivenWrites:0,calendarWrites:0};
  tmv3_statusNotePut_('TMV3_INSTALL_VIEW_LAST_SUCCESS',JSON.stringify(result));return result;
}

// Original: 60_Automation; Function: tmv3_installSchedulePublication_
function tmv3_installSchedulePublication_() {
  return {status:'DEFERRED_TO_NEXT_STANDARD_SLOT'};
}

// Original: 60_Automation; Function: tmv3_installPublishView
function tmv3_installPublishView() {
  return {status:'RETIRED_OUTSIDE_STANDARD_SLOTS',writes:0};
}

// Original: 60_Automation; Function: tmv3_installReadiness
function tmv3_installReadiness() {
  const triggers=tmv3_listTriggers(),handlers=['tmv3_installCheck0800','tmv3_installCheck1300','tmv3_installCheck1700'];
  const rows=tmv3_rows_(TMV3.VERTICALS.Install.sheet),today=tmv3_step6LocalDay_(new Date());
  const upcoming=rows.filter(function(r){return tmv3_clean_(r['Event ID']) && tmv3_step6LocalDay_(r['Calendar Start'])>=today;});
  const issues={};upcoming.forEach(function(r){const status=tmv3_clean_(r['Mapping Status']);issues[status]=(issues[status]||0)+1;});
  let calendarAccessible=false,calendarError='';
  try{calendarAccessible=!!CalendarApp.getCalendarById(TMV3.VERTICALS.Install.calendarIds[0]);}catch(e){calendarError=String(e);}
  const policy=TMV3.MODE==='PRODUCTION_WRITE' && TMV3.OPERATIONS.verticalProductionEnabled===true && (TMV3.OPERATIONS.verticalProductionVerticals||[]).indexOf('Install')>=0;
  const approved=TMV3_HARD_RULES.Install.requiredSalesOrderStatus==='Approved';
  const exact=handlers.every(function(h){return triggers.filter(function(t){return t.handler===h;}).length===1;});
  return {version:TMV3.VERSION,status:approved&&exact&&calendarAccessible&&policy?'CONFIGURED_PENDING_QUEUE_VERIFICATION':'CONFIGURATION_REVIEW',
    readyForTriggerInstallation:approved&&calendarAccessible&&policy,calendarAccessible:calendarAccessible,calendarError:calendarError,productionPolicyEnabled:policy,
    approvedSalesOrderRequired:approved,triggerHours:[8,13,17],timezone:TMV3_TIMEZONE,
    installedInstallTriggers:triggers.filter(function(t){return handlers.indexOf(t.handler)>=0;}),
    installOnlyWorker:true,requestedByRequired:true,contactFallback:'CUSTOMER_OWNED_EXACT_PHONE_OR_EMAIL',
    upcomingAppointments:upcoming.length,cachedMapping:issues,
    queue:JSON.parse(PropertiesService.getScriptProperties().getProperty('TMV3_VERTICAL_SWEEP_Install')||'null'),
    budget:tmv3_strivenApiBudgetStatus_(),businessWrites:0,triggerWrites:0};
}

// Original: 60_Automation; Function: tmv3_installCreatorAllowed_
function tmv3_installCreatorAllowed_(record){return /@classicfireplace\.ca$/i.test(tmv3_normEmail_(record && record.creator || ''));}