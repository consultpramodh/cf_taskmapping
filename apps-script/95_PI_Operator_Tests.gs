// Google Apps Script — PreInspection operator test runner.
// Edit ONLY the settings below. Never edit 00_Configuration or 30_PreInspection to test.
//
// Run from the editor (no arguments needed):
//   piTest_DryRun()     read-only stage report for each event below
//   piTest_Plans()      read-only fresh Striven plan for each event below
//   piTest_Reconcile()  WRITES: one targeted reconciliation per event below
//   piTest_LastState()  read-only: last cycle result and pending queue
//
// piTest_Reconcile opens the release gate ONLY inside this execution, for ONLY
// these events, with creator emails and failure alerts suppressed, then closes
// it again (even on error). Scheduled triggers run in their own executions and
// still see the gate closed.

const PI_TEST = {
  // Google Calendar event IDs on CF Preinspects (with @google.com).
  events: [
    '5mafnqht16o2hrqv4232rofmm3@google.com', // Terry Beettam 53480
    '0aa25j196lol5imushv6mrpmv1@google.com'  // Jakob Pizzera 62883
  ],
  // Event IDs whose operator intake hold (00_Configuration) is lifted for this run only.
  liftHoldFor: [
    '5mafnqht16o2hrqv4232rofmm3@google.com'
  ],
  // Must be true to let piTest_Reconcile write to Striven / Calendar.
  allowWrites: false
};

function piTest_DryRun() {
  return piTest_each_('DRY_RUN', function(id) { return tmv3_preInspectionDryRun(id); });
}

function piTest_Plans() {
  return piTest_each_('PLANS', function(id) {
    return tmv3_step7FreshPlansForEvent_('PreInspection', id).map(tmv3_preInspectionCompactPlan_);
  });
}

function piTest_LastState() {
  const props = PropertiesService.getScriptProperties();
  const out = {lastResult:piTest_json_(props.getProperty('TMV3_PI_LAST_RESULT')),
    pending:piTest_json_(props.getProperty('TMV3_PI_PENDING_EVENTS')),
    releaseGateInCode:TMV3_PI_CLEANED_RELEASE_APPROVED};
  Logger.log('PI_TEST_LAST_STATE=' + JSON.stringify(out));
  return out;
}

function piTest_Reconcile() {
  if (PI_TEST.allowWrites !== true) {
    const refused = {status:'WRITES_NOT_ALLOWED_SET_PI_TEST_allowWrites_TRUE',writes:0};
    Logger.log('PI_TEST_RECONCILE=' + JSON.stringify(refused));
    return refused;
  }
  const ids = piTest_ids_();
  const saved = {gate:TMV3_PI_CLEANED_RELEASE_APPROVED,send:tmv3_preInspectionSendMissingDetails_,
    alert:tmv3_preInspectionFailureAlert_,hold:tmv3_preInspectionApprovalHold_};
  const lifted = PI_TEST.liftHoldFor.map(piTest_norm_);
  const results = [];
  try {
    TMV3_PI_CLEANED_RELEASE_APPROVED = true;
    tmv3_preInspectionSendMissingDetails_ = function() { return {status:'SUPPRESSED_BY_OPERATOR_TEST',sent:0}; };
    tmv3_preInspectionFailureAlert_ = function() { return {status:'SUPPRESSED_BY_OPERATOR_TEST',sent:0}; };
    tmv3_preInspectionApprovalHold_ = function(eventId) {
      return lifted.indexOf(piTest_norm_(eventId)) >= 0 ? null : saved.hold(eventId);
    };
    for (let i = 0; i < ids.length; i++) {
      let r;
      try { r = tmv3_reconcilePreInspectionEvent(ids[i]); }
      catch (e) { r = {status:'ERROR',error:String(e && e.message || e)}; }
      const one = (r.results || [])[0] || {};
      const summary = {eventId:ids[i],status:r.status,processed:r.processed,writes:r.writes,
        eventStatus:one.status || one.plan || one.intake || '',plan:one.plan || '',blocker:one.blocker || '',
        taskId:one.taskId || '',recreate:one.recreateAuthorized || '',error:one.error || r.error || '',
        afterPlans:one.afterPlans || []};
      results.push(summary);
      Logger.log('PI_TEST_RECONCILE=' + JSON.stringify(summary));
      if (r.status !== 'COMPLETE') break; // stop on the first event that did not converge
    }
  } finally {
    TMV3_PI_CLEANED_RELEASE_APPROVED = saved.gate;
    tmv3_preInspectionSendMissingDetails_ = saved.send;
    tmv3_preInspectionFailureAlert_ = saved.alert;
    tmv3_preInspectionApprovalHold_ = saved.hold;
  }
  return {status:results.every(function(x) { return x.status === 'COMPLETE'; }) && results.length === ids.length ?
    'ALL_TARGETS_COMPLETE' : 'STOPPED_REVIEW_LOG',results:results};
}

function piTest_each_(label, fn) {
  return piTest_ids_().map(function(id) {
    let out;
    try { out = fn(id); } catch (e) { out = {status:'ERROR',error:String(e && e.message || e)}; }
    Logger.log('PI_TEST_' + label + ' ' + id + ' = ' + JSON.stringify(out));
    return {eventId:id,result:out};
  });
}

function piTest_ids_() {
  const ids = (PI_TEST.events || []).map(function(x) { return String(x || '').trim(); });
  if (!ids.length) throw new Error('PI_TEST_NO_EVENTS');
  ids.forEach(function(id) { if (!/^[a-z0-9_]+@google\.com$/i.test(id)) throw new Error('PI_TEST_EVENT_ID_INVALID: ' + id); });
  return ids;
}

function piTest_norm_(id) { return String(id || '').trim().toLowerCase().replace(/@google\.com$/, ''); }

function piTest_json_(raw) { try { return raw ? JSON.parse(raw) : null; } catch (e) { return raw; } }
