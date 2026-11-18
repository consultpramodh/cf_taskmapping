// MIGRATION STAGING ONLY. Source: reconstructed V3 3.12.21. NOT APPROVED FOR DEPLOYMENT.

// Original: 60_Automation; Function: tmv3_removeTriggers
function tmv3_removeTriggers() {
  const managed = tmv3_managedTriggerHandlers_();
  let deleted = 0;

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (managed.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  });

  tmv3_audit_(
    'SYSTEM','','','REMOVE_MANAGED_TRIGGERS','PASS',
    'Deleted ' + deleted + ' managed trigger(s).'
  );
  return tmv3_listTriggers();
}

// Original: 60_Automation; Function: tmv3_listTriggers
function tmv3_listTriggers() {
  return ScriptApp.getProjectTriggers().map(function(t) {
    return {
      handler: t.getHandlerFunction(),
      source: String(t.getTriggerSource()),
      eventType: String(t.getEventType()),
      sourceId: tmv3_safeCalendar_(function() { return t.getTriggerSourceId(); }, '')
    };
  });
}

// Original: 60_Automation; Function: tmv3_logTriggers
function tmv3_logTriggers() {
  const result = tmv3_listTriggers();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Original: 60_Automation; Function: tmv3_allVerticalAutomationTick
function tmv3_allVerticalAutomationTick() {
  const hour=Number(Utilities.formatDate(new Date(),TMV3_TIMEZONE,'H'));
  if(hour<8 || hour>18) return {status:'OUTSIDE_BUSINESS_HOURS',writes:0};
  const props=PropertiesService.getScriptProperties(),all=['Service','Delivery','Install'];
  const start=Number(props.getProperty('TMV3_VERTICAL_ROTATION')||0)%all.length;
  props.setProperty('TMV3_VERTICAL_ROTATION',String((start+1)%all.length));
  const result={status:'ACTIVE',results:[]},began=Date.now();
  for(let i=0;i<all.length;i++){
    if(Date.now()-began>210000)break;
    try{result.results.push(tmv3_verticalProductionCycle_(all[(start+i)%all.length]));}
    catch(err){result.results.push({vertical:all[(start+i)%all.length],status:'ATTENTION',error:String(err)});}
  }
  result.sheetRefresh=tmv3_refreshAutomationSheetsAfter_(began);
  Logger.log(JSON.stringify(result));return result;
}

// Original: 60_Automation; Function: tmv3_scheduledSheetRefresh
function tmv3_scheduledSheetRefresh() {
  const hour=Number(Utilities.formatDate(new Date(),TMV3_TIMEZONE,'H'));
  if([8,13,17].indexOf(hour)<0)return {status:'OUTSIDE_STANDARD_SLOTS'};
  const result=tmv3_refreshAutomationSheets();
  // Daily missing-links digest (operator only) after the morning slots have run.
  if(hour===13){try{result.linksDigest=tmv3_sendMissingLinksDigest_(false);}catch(err){result.linksDigest={status:'ATTENTION',error:String(err&&err.message||err).slice(0,200)};}}
  return result;
}

// Original: 60_Automation; Function: tmv3_standardDailySchedule_
function tmv3_standardDailySchedule_() {
  return {hours:[8,13,17],timezone:'America/Toronto',verticals:['PreInspection','Install','Delivery','Service']};
}

// Original: 60_Automation; Function: tmv3_standardVerticalSlot_
function tmv3_standardVerticalSlot_(vertical,slot) {
  const result=tmv3_runFullVerticalSweep_(vertical);
  if(vertical==='Install') {
    result.reminders=tmv3_installReminderCheck();
  }
  tmv3_audit_(vertical,'','','STANDARD_DAILY_SLOT',result.status,tmv3_verticalRunSummary_(slot,result));
  return result;
}

/** One TM Audit row per run: counts by outcome, plus only the installs that changed or need attention. */
function tmv3_verticalRunSummary_(slot,result) {
  const counts={},items=[];
  (result.results||[]).forEach(function(r){
    counts[r.status]=(counts[r.status]||0)+1;
    if(['ALREADY_ALIGNED_READ_ONLY','FULFILLED_NO_RECREATE','SKIPPED_CREATOR_UNVERIFIED'].indexOf(r.status)>=0)return;
    if(r.status==='VERIFIED_CONVERGENCE' && !(r.executions||[]).some(function(e){return (e.actions||[]).some(function(a){return a!=='VERIFY_CALENDAR_LINKS';});}))return;
    items.push({eventId:r.eventId,status:r.status,error:String(r.error||'').slice(0,160)||undefined,
      changes:r.executions?r.executions.map(function(e){return (e.actions||[]).filter(function(a){return a!=='VERIFY_CALENDAR_LINKS';}).join('+');}).join(','):undefined});
  });
  const out={slot:slot,status:result.status,checked:Number(result.scanned||0),writes:Number(result.writes||0),
    pending:Number(result.pending||0),counts:counts,items:items,reminders:result.reminders};
  let text=JSON.stringify(out);
  // A Sheets cell holds 50,000 characters.
  while(text.length>45000 && out.items.length){out.items.length=Math.floor(out.items.length/2);out.itemsTruncated=true;text=JSON.stringify(out);}
  return text.length>49000?text.slice(0,49000):text;
}
// The big Striven report refresh (Customers, Locations, Orders, Tasks...) runs on its own, before each
// slot window, so no business slot holds the script lock while it runs. Operator 2026-11-18: the Install
// 08:00 slot timed out at 30 minutes on 2026-11-16 and 2026-11-18 inside this refresh.
const TMV3_SOURCE_REFRESH_HOURS = Object.freeze([5, 11, 15]);
// A slot leaves the refresh to the dedicated run when that run started within this window.
const TMV3_SOURCE_REFRESH_SLOT_WINDOW_MS = 5 * 3600000;

/** Trigger handler: refresh the Striven source caches (5:00, 11:00, 15:00 Toronto). One audit row. */
function tmv3_sourceRefreshSlot() {
  const started = Date.now(), props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(120000)) {
    tmv3_audit_('SYSTEM','','','SOURCE_REFRESH','LOCK_BUSY','');
    return {status:'LOCK_BUSY'};
  }
  let result;
  try {
    // Recorded before the refresh, so a timed-out refresh still keeps the slots from retrying it inline.
    props.setProperty('TMV3_SOURCE_REFRESH_SLOT_ATTEMPT_AT', String(started));
    const r = tmv3_refreshOperationalSourcesIfNeeded_(120, false);
    result = {status:r.status,master:!!(r.master && r.master.refreshed),transactions:!!(r.transactions && r.transactions.refreshed)};
  } catch (err) {
    result = {status:'ATTENTION',error:String(err && err.message || err).slice(0, 300)};
  } finally { lock.releaseLock(); }
  result.seconds = Math.round((Date.now() - started) / 1000);
  tmv3_audit_('SYSTEM','','','SOURCE_REFRESH',result.status,JSON.stringify(result));
  return result;
}

/** Used by the business slots instead of refreshing inline. Falls back to the inline refresh only when
 * the dedicated run has not started in the last 5 hours (e.g. its triggers are not installed). */
function tmv3_slotSourceRefresh_() {
  const at = Number(PropertiesService.getScriptProperties().getProperty('TMV3_SOURCE_REFRESH_SLOT_ATTEMPT_AT') || 0);
  if (at && Date.now() - at < TMV3_SOURCE_REFRESH_SLOT_WINDOW_MS)
    return {status:'LEFT_TO_SOURCE_REFRESH_RUN',lastRunStarted:new Date(at).toISOString()};
  return tmv3_refreshOperationalSourcesIfNeeded_(120, false);
}
