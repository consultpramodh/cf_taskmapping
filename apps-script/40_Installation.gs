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
    const subject = 'ACTION REQUIRED: Sales Order # missing from install event';
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
        htmlBody:'<p><em>This is an automated email notification from Classic Fireplace &amp; BBQ Store.</em></p>'+
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
function tmv3_installSweepContinue(){return {status:'RETIRED_OUTSIDE_STANDARD_SLOTS',writes:0};}

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