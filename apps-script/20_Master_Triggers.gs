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

// =====================================================================
// ORDERED SLOT RUN (2026-11-20). Before this, PreInspection, Install, Delivery and Service each had their
// own 08:00/13:00/17:00 trigger. Apps Script fires "nearMinute" triggers anywhere in a wide window, so on
// 2026-11-20 08:00 Install ran 07:24, Service 08:02, PreInspection 08:04 and Delivery 08:39. PreInspection
// started while Service held the script lock: its Calendar review gave up (PI_REVIEW_LOCK_BUSY) and lock
// waits used most of its 3-minute batch window, so only 3 events were checked and 4 were left for 13:00.
// Now ONE trigger per slot starts PreInspection, and each workflow starts the next only when it has
// finished, always in tmv3_standardDailySchedule_().verticals order, each in its own execution.
// =====================================================================
const TMV3_SLOT_CHAIN_KEY = 'TMV3_SLOT_CHAIN';
const TMV3_SLOT_CHAIN_NEXT_HANDLER = 'tmv3_slotChainNext';
// Gap between one workflow finishing and the next one starting.
const TMV3_SLOT_CHAIN_GAP_MS = 60000;
// A step still marked RUNNING after this long was killed by the Apps Script time limit: move on.
const TMV3_SLOT_CHAIN_STEP_TIMEOUT_MS = 35 * 60000;

/** Trigger handlers installed by tmv3_installStandardDailyTriggers (one per slot). */
function tmv3_standardSlot0800() { return tmv3_slotChainStart_('08:00'); }
function tmv3_standardSlot1300() { return tmv3_slotChainStart_('13:00'); }
function tmv3_standardSlot1700() { return tmv3_slotChainStart_('17:00'); }

/** The business trigger handlers the standard schedule installs: one per slot hour. */
function tmv3_standardSlotHandlers_() {
  return tmv3_standardDailySchedule_().hours.map(function(hour) {
    return {handler:'tmv3_standardSlot' + ('0' + hour).slice(-2) + '00', hour:hour};
  });
}

/** Per-workflow slot handlers from before 2026-11-20. Still callable by hand; no longer installed. */
function tmv3_legacyVerticalSlotHandlers_() {
  const schedule = tmv3_standardDailySchedule_(), out = [];
  schedule.verticals.forEach(function(v) {
    const prefix = v === 'PreInspection' ? 'preInspection' : v.toLowerCase();
    schedule.hours.forEach(function(h) { out.push('tmv3_' + prefix + 'Check' + ('0' + h).slice(-2) + '00'); });
  });
  return out;
}

function tmv3_slotChainRead_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(TMV3_SLOT_CHAIN_KEY) || 'null'); }
  catch (ignored) { return null; }
}

function tmv3_slotChainWrite_(state) {
  PropertiesService.getScriptProperties().setProperty(TMV3_SLOT_CHAIN_KEY, JSON.stringify(state));
}

function tmv3_slotChainClearTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === TMV3_SLOT_CHAIN_NEXT_HANDLER) ScriptApp.deleteTrigger(t);
  });
}

/** Exactly one pending one-off chain trigger at a time. */
function tmv3_slotChainSchedule_(delayMs) {
  tmv3_slotChainClearTriggers_();
  ScriptApp.newTrigger(TMV3_SLOT_CHAIN_NEXT_HANDLER).timeBased().after(Math.max(60000, delayMs)).create();
}

/** Runs one workflow's standard slot exactly as its own trigger used to. */
function tmv3_slotChainRunVertical_(vertical, slot) {
  if (vertical === 'PreInspection') return tmv3_runPreInspectionDailySlot_(slot);
  return tmv3_standardVerticalSlot_(vertical, slot);
}

/** Slot trigger: start a new ordered run with the first workflow (PreInspection). */
function tmv3_slotChainStart_(slot) {
  const prior = tmv3_slotChainRead_();
  if (prior && prior.status === 'RUNNING') {
    // The previous slot never finished its queue. Record which workflows it skipped; the new slot
    // starts again from PreInspection and each workflow's own sweep carries on where it stopped.
    tmv3_audit_('SYSTEM', '', '', 'SLOT_CHAIN', 'ABANDONED', JSON.stringify({slot:prior.slot, day:prior.day,
      skipped:prior.order.slice(prior.index), current:prior.step || null}));
  }
  const state = {slot:slot, day:Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'yyyy-MM-dd'),
    order:tmv3_standardDailySchedule_().verticals.slice(), index:0, status:'RUNNING',
    startedAt:new Date().toISOString(), steps:[]};
  tmv3_slotChainWrite_(state);
  return tmv3_slotChainRunCurrent_(state);
}

/** One-off trigger: run the next workflow, or move past a step the time limit killed. */
function tmv3_slotChainNext() {
  const state = tmv3_slotChainRead_();
  if (!state || state.status !== 'RUNNING') { tmv3_slotChainClearTriggers_(); return {status:'NO_ACTIVE_SLOT_CHAIN'}; }
  const step = state.step;
  if (step && step.status === 'RUNNING') {
    const age = Date.now() - new Date(step.startedAt).getTime();
    if (age < TMV3_SLOT_CHAIN_STEP_TIMEOUT_MS) {
      // Still inside its time limit (should not happen): check again when the limit has passed.
      tmv3_slotChainSchedule_(TMV3_SLOT_CHAIN_STEP_TIMEOUT_MS - age + 60000);
      return {status:'SLOT_CHAIN_STEP_STILL_RUNNING', vertical:step.vertical};
    }
    step.status = 'TIMED_OUT'; step.finishedAt = new Date().toISOString();
    tmv3_audit_('SYSTEM', '', '', 'SLOT_CHAIN', 'STEP_TIMED_OUT', JSON.stringify({slot:state.slot, vertical:step.vertical, startedAt:step.startedAt}));
    state.steps.push(step); state.step = null; state.index++;
    tmv3_slotChainWrite_(state);
  }
  return tmv3_slotChainRunCurrent_(state);
}

function tmv3_slotChainRunCurrent_(state) {
  if (state.index >= state.order.length) {
    state.status = 'COMPLETE'; state.finishedAt = new Date().toISOString(); state.step = null;
    tmv3_slotChainWrite_(state); tmv3_slotChainClearTriggers_();
    tmv3_audit_('SYSTEM', '', '', 'SLOT_CHAIN', 'COMPLETE', JSON.stringify({slot:state.slot, day:state.day, steps:state.steps}));
    return {status:'SLOT_CHAIN_COMPLETE', slot:state.slot, steps:state.steps};
  }
  const vertical = state.order[state.index];
  state.step = {vertical:vertical, status:'RUNNING', startedAt:new Date().toISOString()};
  tmv3_slotChainWrite_(state);
  // Safety net first: if this execution is killed by the time limit, the one-off trigger moves the chain on.
  tmv3_slotChainSchedule_(TMV3_SLOT_CHAIN_STEP_TIMEOUT_MS);
  let result;
  try { result = tmv3_slotChainRunVertical_(vertical, state.slot) || {}; }
  catch (err) {
    result = {status:'ATTENTION', error:String(err && err.message || err).slice(0, 300)};
    try { tmv3_audit_(vertical, '', '', 'STANDARD_DAILY_SLOT', 'ATTENTION', JSON.stringify({slot:state.slot, error:result.error})); } catch (ignored) {}
  }
  // A newer slot may have taken over while this step ran; never overwrite its state.
  const latest = tmv3_slotChainRead_();
  if (!latest || latest.startedAt !== state.startedAt) return {status:'SLOT_CHAIN_SUPERSEDED', vertical:vertical, result:result};
  state.step.status = 'DONE'; state.step.result = String(result.status || 'UNKNOWN');
  state.step.finishedAt = new Date().toISOString();
  state.steps.push(state.step); state.step = null; state.index++;
  tmv3_slotChainWrite_(state);
  if (state.index >= state.order.length) return tmv3_slotChainRunCurrent_(state);
  tmv3_slotChainSchedule_(TMV3_SLOT_CHAIN_GAP_MS);
  return {status:'SLOT_CHAIN_STEP_DONE', vertical:vertical, next:state.order[state.index], result:result};
}

/** READ-ONLY: where the current or last ordered slot run is. */
function tmv3_slotChainStatus() {
  const out = {state:tmv3_slotChainRead_(),
    pendingChainTriggers:ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === TMV3_SLOT_CHAIN_NEXT_HANDLER; }).length};
  Logger.log(JSON.stringify(out)); return out;
}
