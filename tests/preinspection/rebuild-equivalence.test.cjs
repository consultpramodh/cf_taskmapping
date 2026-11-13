'use strict';
// Differential test: live baseline 30_PreInspection (2026-11-11 capture) vs the
// rebuilt module, both loaded with the same shared files and stubs. Outputs,
// thrown messages and every stubbed side effect must match, except for the
// documented, intentional guard additions which are asserted separately.
//
//   node tests/preinspection/rebuild-equivalence.test.cjs
//   SHARED_DIR=/path/to/live/apps-script node tests/preinspection/rebuild-equivalence.test.cjs
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const h = require('./harness.cjs');

const ROOT = path.resolve(__dirname, '..', '..');
const SHARED = process.env.SHARED_DIR || path.join(ROOT, 'apps-script');
const BASELINE = fs.readFileSync(path.join(__dirname, 'baseline', '30_PreInspection.live-2026-11-11.gs'), 'utf8');
const REBUILT = fs.readFileSync(path.join(ROOT, 'apps-script', '30_PreInspection.gs'), 'utf8');
const REGRESSION = fs.readFileSync(path.join(__dirname, 'baseline', '90_Regression_Tests.live-2026-11-11.gs'), 'utf8');
const NOW = Date.parse('2026-11-10T14:05:00Z'); // Tuesday 09:05 Toronto

let checks = 0;
const failures = [];
function check(name, fn) {
  checks++;
  try { fn(); } catch (e) { failures.push(name + ': ' + String(e && e.message || e).slice(0, 600)); }
}
const BASELINE_RULES = t => t.replace(/^var TMV3_PI_REQUESTED_BY_CREATOR_FIRST = true;/m, 'var TMV3_PI_REQUESTED_BY_CREATOR_FIRST = false;').replace(/^var TMV3_PI_CONTINUE_SWEEP_UNTIL_DONE = true;/m, 'var TMV3_PI_CONTINUE_SWEEP_UNTIL_DONE = false;').replace(/^var TMV3_PI_ONE_WRITE_PER_BATCH = false;/m, 'var TMV3_PI_ONE_WRITE_PER_BATCH = true;').replace(/^var TMV3_PI_GRACE_NEVER_PAST_APPOINTMENT = true;/m, 'var TMV3_PI_GRACE_NEVER_PAST_APPOINTMENT = false;').replace(/^var TMV3_PI_NOTIFY_MISSING_SALES_ORDER = true;/m, 'var TMV3_PI_NOTIFY_MISSING_SALES_ORDER = false;')
  .replace(/^var TMV3_PI_OFF_SHARED_NOTICES_ENABLED = true;/m, 'var TMV3_PI_OFF_SHARED_NOTICES_ENABLED = false;')
  .replace(/^var TMV3_PI_EPHEMERAL_STATE_IN_CACHE = true;/m, 'var TMV3_PI_EPHEMERAL_STATE_IN_CACHE = false;');
const CLOSE = (f, t) => f === '30_PreInspection.gs' ? BASELINE_RULES(t.replace(/^var TMV3_PI_CLEANED_RELEASE_APPROVED = true;/m, 'var TMV3_PI_CLEANED_RELEASE_APPROVED = false;')) : t;
function load(source, opts) {
  opts = Object.assign({transform: CLOSE}, opts || {});
  // The V3 regression suite was removed from live on 2026-11-12; it stays a test fixture here.
  return h.load(SHARED, Object.assign({fixedNow: NOW, overrides: {'30_PreInspection.gs': source, '90_Regression_Tests.gs': REGRESSION}}, opts));
}
function outcome(fn) {
  try { return {ok: JSON.parse(JSON.stringify(fn() === undefined ? null : fn()))}; }
  catch (e) { return {error: String(e && e.message || e)}; }
}
function same(name, run) {
  check(name, () => {
    const a = load(BASELINE), b = load(REBUILT);
    const x = run(a), y = run(b);
    assert.deepEqual(y, x);
  });
}
// Run fn(ctx, arg) for each arg in both contexts; compare normalized outcome.
function sameOver(name, fnName, argsList, setup) {
  check(name, () => {
    const a = load(BASELINE), b = load(REBUILT);
    if (setup) { setup(a); setup(b); }
    argsList.forEach((args, i) => {
      const x = outcome(() => a.ctx[fnName].apply(null, clone(args)));
      const y = outcome(() => b.ctx[fnName].apply(null, clone(args)));
      assert.deepEqual(y, x, fnName + ' case #' + i + ' ' + JSON.stringify(args).slice(0, 200));
    });
  });
}
function clone(v) { return v === undefined ? v : JSON.parse(JSON.stringify(v)); }

// ---------------------------------------------------------------- fixtures
const PRIMARY = 'c_3088a3989f3eb809957ed5c40137a7111a0ac97f68c29b40c144028cb14320dc@group.calendar.google.com';
const STEPHEN_CAL = 'classicfireplace.ca_c20qcqfhvjbv784asn9pvuiaf4@group.calendar.google.com';
const titles = ['12345 - Sample Customer - (416) 555-0100', 'C#12345 Sample Customer 416-555-0100', 'Preinspect - Jane Doe',
  'Team Meeting', 'do not book', 'Newmarket', 'Travel to Barrie', '586828 - Pat Lee (905) 555-1234',
  '', 'Service - Furnace repair 905 555 4444', 'Pre-Inspection 54321 Bob Smith', 'Install shut off and tags',
  'Out of office', 'SO#5853/67 Sample', '62883 - Jakob - Pizzera - (647) 555-0111'];
const descriptions = ['', 'Pre-inspection <br> Remove propane fireplace', 'SO 586828 notes', 'Sales Order: 123456\nNotes: bring ladder',
  '<p>Line one</p><p>Line two</p><b>bold</b>', 'Not a pre-inspection, just a call', 'pre-inspection already completed',
  'Measure for trim return visit', 'Order # 77777 and SO 88888', 'Task #19694 – Jeff Acorn',
  '12345 - Sample Customer - (416) 555-0100<br><br>Notes<br>-----Striven Links-----<br><a href="https://x.example/a">A</a>'];
function record(i) {
  const t = titles[i % titles.length], d = descriptions[(i * 7) % descriptions.length];
  const roles = [['PRIMARY_SHARED'], ['SECONDARY_STEPHEN'], [''], ['PRIMARY_SHARED', 'SECONDARY_STEPHEN']][i % 4];
  return {vertical: 'PreInspection', eventId: 'evt' + i + '@google.com', title: t, rawDescription: d, descriptionClean: d,
    calendarRole: roles[0], sourceCalendarRoles: roles, calendarId: roles[0] === 'PRIMARY_SHARED' ? PRIMARY : STEPHEN_CAL,
    customerNumber: /\b(\d{5})\b/.exec(t) ? /\b(\d{5})\b/.exec(t)[1] : '', calendarCustomerName: i % 3 ? 'Sample Customer' : '',
    phone: i % 5 ? '(416) 555-0100' : '', location: i % 4 ? '10 Example St, Toronto, ON' : '', orderNumber: i % 6 === 0 ? '586828' : '',
    isAllDay: i === 13, status: i === 9 ? 'cancelled' : 'confirmed',
    creator: ['adam@classicfireplace.ca', 'stephen@classicfireplace.ca', '', 'a@classicfireplace.ca,b@classicfireplace.ca', 'x@gmail.com'][i % 5],
    sourceCreators: [], guests: 'stephen@classicfireplace.ca,' + PRIMARY,
    organizerCalendarId: ['stephen@classicfireplace.ca', 'jay@classicfireplace.ca', 'karen@classicfireplace.ca', STEPHEN_CAL, '', 'z@x.com'][i % 6],
    start: new Date(NOW + (i - 3) * 86400000 + 3600000).toISOString(), end: new Date(NOW + (i - 3) * 86400000 + 7200000).toISOString(),
    calendarCreatedAt: new Date(NOW - (i % 9) * 86400000).toISOString(), existingTaskId: i % 7 === 0 ? '19694' : '',
    fingerprint: 'fp' + i};
}
const RECORDS = Array.from({length: 40}, (_, i) => record(i));
const CUSTOMERS = [
  {'Customer ID': '100', 'Customer Number': '12345', Name: 'Sample Customer', 'Primary Phone': '4165550100'},
  {'Customer ID': '200', 'Customer Number': '54321', Name: 'Bob Smith', 'Primary Phone': '9055550000'},
  {'Customer ID': '300', 'Customer Number': '62883', Name: 'Jakob Pizzera', 'Primary Phone': '6475550111'}];
function withRows(env) {
  env.ctx.tmv3_rows_ = function(name) { return name === env.ctx.TMV3_SHEET_CUSTOMERS_NAME ? CUSTOMERS : CUSTOMERS; };
  env.ctx.tmv3_audit_ = function() { env.log.push({kind: 'AUDIT', args: Array.from(arguments).slice(3, 5)}); };
}

// ---------------------------------------------------------------- 1 interface
check('INTERFACE_PRESERVED_FOR_CALLERS', () => {
  const defs = s => new Set((s.match(/^function\s+([A-Za-z0-9_$]+)/mg) || []).map(x => x.replace(/^function\s+/, '')));
  const oldDefs = defs(BASELINE), newDefs = defs(REBUILT);
  const others = h.ORDER.filter(f => f !== '30_PreInspection.gs' && fs.existsSync(path.join(SHARED, f)))
    .map(f => fs.readFileSync(path.join(SHARED, f), 'utf8')).join('\n');
  const external = [...oldDefs].filter(n => new RegExp('\\b' + n + '\\b').test(others));
  const handlers = ['tmv3_preInspectionCheck0800', 'tmv3_preInspectionCheck1300', 'tmv3_preInspectionCheck1700',
    'tmv3_preInspectionContinueEligibleSweep', 'tmv3_preInspectionScheduledCycle', 'tmv3_runPreInspectionAutomation',
    'tmv3_previewPreInspectionMissingDetailEmails'];
  const missing = external.concat(handlers).filter(n => !newDefs.has(n));
  assert.deepEqual(missing, [], 'functions still called elsewhere are missing');
  const removed = [...oldDefs].filter(n => !newDefs.has(n)).sort();
  const approved = ['tmv3_importTerryQuietCalendarIntake', 'tmv3_piNov11IdentityReadback_', 'tmv3_piNov11OrderReadOnly_',
    'tmv3_piNov11ReadOnlyOne_', 'tmv3_piNov11TerryJakobReadOnly', 'tmv3_piNov11TerryJakobRunIfReleased',
    'tmv3_piOriginal_installPreInspectionAutomationTrigger', 'tmv3_piOriginal_preInspectionAutomaticRecovery_',
    'tmv3_piOriginal_preInspectionClearContinuationTrigger_', 'tmv3_piOriginal_preInspectionOptionalOrderLocationEvidence_',
    'tmv3_piOriginal_reconcilePreInspectionEvent', 'tmv3_piOriginal_reconcilePreInspectionPendingReview',
    'tmv3_piOriginal_runPreInspectionAutomation', 'tmv3_piOriginal_runPreInspectionDailySlot_',
    'tmv3_preInspectionLastWeekAudit', 'tmv3_preInspectionOct21Preview', 'tmv3_preInspectionOct21RefreshCache',
    'tmv3_preInspectionOct21RunNext', 'tmv3_preInspectionRepairNov9Stephen', 'tmv3_preInspectionRepairThreeAddedByOthersNov9',
    'tmv3_preInspectionStoreDiscovery_', 'tmv3_preInspectionTodayEligibleQueue_', 'tmv3_preInspectionTodayLocationResponseShape',
    'tmv3_preInspectionTodayRefreshTaskCacheSchedules', 'tmv3_preInspectionTodayRunEligibleBatch', 'tmv3_preInspectionTodayRunNext',
    'tmv3_preInspectionTodayRunUnlinked', 'tmv3_previewTerryQuietCalendarIntake', 'tmv3_restoreRecordedPreInspectionLegacyTitles',
    'tmv3_terryQuietCalendarIntake_'].sort();
  assert.deepEqual(removed, approved, 'unexpected removals');
});

check('EVERY_CALLED_TMV3_FUNCTION_RESOLVES', () => {
  const b = load(REBUILT);
  // Ignore names inside regex literals such as /^tmv3_preInspectionCheck(0800|1300|1700)$/.
  const called = new Set((REBUILT.replace(/\/\^tmv3_[^/]*\//g, '').match(/\btmv3_[A-Za-z0-9_]+(?=\s*\()/g) || []));
  const unresolved = [...called].filter(n => typeof b.ctx[n] !== 'function');
  assert.deepEqual(unresolved, []);
});

check('FROZEN_RENDERER_FINGERPRINT', () => {
  const b = load(REBUILT);
  assert.equal(b.ctx.tmv3_piRendererFingerprint_(), 'b77a1a07');
  assert.equal(b.ctx.tmv3_piRendererFingerprint_(), load(BASELINE).ctx.tmv3_piRendererFingerprint_());
});

check('RELEASED_WITH_CREATOR_EMAILS_OFF', () => {
  assert.match(REBUILT, /^var TMV3_PI_CLEANED_RELEASE_APPROVED = true;$/m);
  assert.match(REBUILT, /^var TMV3_PI_CREATOR_EMAILS_ENABLED = true;$/m);
});

// ---------------------------------------------------------------- 2 pure functions
const one = RECORDS.map(r => [r]);
sameOver('LEGACY_ORDER_NUMBER', 'tmv3_extractLegacyPreInspectionOrderNumber_', descriptions.concat(titles).map(x => [x]));
sameOver('CALENDAR_CUSTOMER_NAME', 'tmv3_preInspectionCalendarCustomerName_', titles.flatMap(t => [[t, ''], [t, '12345'], [t, '586828']]));
sameOver('NON_CUSTOMER', 'tmv3_step2PreInspectionNonCustomer_', one);
sameOver('FORMAT_WARNINGS', 'tmv3_step2PreInspectionFormatWarnings_', one);
sameOver('STEP2_CLASSIFY', 'tmv3_step2ClassifyPreInspection_', one);
sameOver('DISCOVERY_DECISION', 'tmv3_preInspectionDiscoveryDecision_', RECORDS.map(r => [r, {secondaryIgnoredCreatorEmails: ['stephen@classicfireplace.ca']}]));
sameOver('HAS_SO_REFERENCE', 'tmv3_preInspectionHasSalesOrderReference_', one);
sameOver('CREATOR_EMAIL', 'tmv3_preInspectionCreatorEmail_', one);
sameOver('CREATOR_FIRST_NAME', 'tmv3_preInspectionCreatorFirstName_', [['adam@classicfireplace.ca'], ['mary-jo.smith@x.ca'], [''], ['_x@y']]);
sameOver('NOTIFICATION_CUSTOMER_NUMBER', 'tmv3_preInspectionNotificationCustomerNumber_', RECORDS.map(r => [r, CUSTOMERS]));
sameOver('MISSING_DETAIL_ASSESSMENT', 'tmv3_preInspectionMissingDetailAssessment_', one, withRows);
sameOver('EMAIL_PREVIEW_HTML_SUBJECT', 'tmv3_preInspectionMissingDetailEmailPreview_',
  RECORDS.map((r, i) => [Object.assign({}, r, {notificationFollowupNumber: i % 3})]), withRows);
sameOver('PREVIEW_FROM_RECORDS', 'tmv3_preInspectionMissingDetailPreviewFromRecords_', [[RECORDS]], withRows);
sameOver('DISPLAY_GUESTS', 'tmv3_preInspectionDisplayGuests_', one);
sameOver('HTML_ESCAPE', 'tmv3_preInspectionHtmlEscape_', [['<a href="x">&\'</a>'], [null], [undefined], [0]]);
const nows = [NOW, NOW + 86400000 * 2, NOW + 86400000 * 4, NOW - 3600000 * 30, Date.parse('2026-11-14T15:00:00Z')];
const states = [null, {status: 'SENT', sentCount: 1, lastSentAt: new Date(NOW - 90000000).toISOString()},
  {status: 'SENT', sentCount: 2, lastSentAt: new Date(NOW - 3600000).toISOString()}, {status: 'SENDING'}, {status: 'SENT', lastSentAt: 'bad'}];
sameOver('NOTIFICATION_DECISION', 'tmv3_preInspectionNotificationDecision_',
  RECORDS.flatMap(r => states.flatMap(s => nows.map(n => [r, s, n]))));
const links = [{key: 'c', url: 'https://classicfireplace.striven.com/next/crm#/sales-orders?accountId=100', label: 'View Sales Orders – Sample (#12345)'},
  {key: 't', url: 'https://classicfireplace.striven.com/Tasks/TaskInfo.aspx?TaskID=10', label: 'Task #10 – Sample - 10 Example St, Toronto - (416) 555-0100'}];
const descs = descriptions.concat(['plain text\nline two\n\nparagraph', '<!-- PREINSPECT_STRIVEN_TASK_LINK_START -->old<!-- PREINSPECT_STRIVEN_TASK_LINK_END -->kept']);
sameOver('AUTHORED_MARKUP', 'tmv3_preInspectionAuthoredMarkup_', descs.map(d => [d]));
sameOver('MANAGED_DESCRIPTION', 'tmv3_preInspectionClickableManagedDescription_',
  descs.flatMap(d => [[d, {links}], [d, {links: [links[0]]}], [d, {links: [links[0], links[0]]}], [d, {links: []}]]));
sameOver('LINK_PRESENT', 'tmv3_preInspectionClickableLinkPresent_', descs.flatMap(d => links.map(l => [d, l])));
sameOver('ENSURE_LEGACY_TITLE', 'tmv3_preInspectionEnsureLegacyTitle_', descs.flatMap(d => titles.slice(0, 4).map(t => [d, t])));
sameOver('TITLE_VARIANTS', 'tmv3_preInspectionTitleVariants_', [[{customerNumber: '1', customerName: 'A', phone: 'P', currentTitle: 'X'}], [{}], [null]]);
sameOver('STRIP_PRESERVED_TITLE', 'tmv3_stripPreInspectionPreservedTitle_', descs.map(d => [d, {customerNumber: '12345', customerName: 'Sample Customer', phone: '(416) 555-0100', currentTitle: titles[0]}]));
sameOver('NORMALIZE_WRAPS', 'tmv3_normalizePreInspectionAuthoredWraps_', descs.map(d => [d]));
sameOver('NOTES_TEXT', 'tmv3_preInspectionCalendarNotesText_', descs.map(d => [d, {}]));
sameOver('ASSERT_LINK_DESCRIPTION', 'tmv3_assertPreInspectionLinkDescription_', descs.map(d => [d, {links}]));
sameOver('CALENDAR_TITLE_ASSERT', 'tmv3_piAssertCalendarTitle_', titles.map(t => [t]));
sameOver('TASK_TITLE_ASSERT', 'tmv3_piAssertTaskTitle_', titles.concat(['A - 1 St, Town - (416) 555-0100']).map(t => [t]));
sameOver('OWNER_EMPLOYEE', 'tmv3_piVerifiedOwnerEmployee_', one);
sameOver('APPROVAL_HOLD', 'tmv3_preInspectionApprovalHold_', [['5mafnqht16o2hrqv4232rofmm3@google.com'], ['5MAFNQHT16O2HRQV4232ROFMM3'], ['x'], ['']]);
sameOver('CREATOR_EXCLUDED', 'tmv3_piCreatorEmailExcluded_', [['Jay@classicfireplace.ca'], ['adam@classicfireplace.ca'], ['']]);
const reqBys = [undefined, null, {Type: 'employee', Id: 15}, {type: 'Employee', id: '1'}, {Type: 'contact', Id: 1}, {Type: 'employee', Id: 0}, {Type: 'employee', Id: '1.5'}];
sameOver('REQUESTED_BY_SHAPE', 'tmv3_piAssertRequestedByEmployeeShape_', reqBys.map(r => [r]));
sameOver('REQUESTED_BY_PLAN', 'tmv3_piAssertRequestedByPlan_', RECORDS.slice(0, 12).flatMap(r =>
  [{expectedRequestedByType: 'employee', expectedRequestedById: '15'}, {expectedRequestedByType: 'employee', expectedRequestedById: 1},
    {expectedRequestedByType: 'contact', expectedRequestedById: '15'}, {}].map(p => [p, r])));
const createPayloads = [
  {Type: {Id: 105}, Description: '', RequestedBy: {Type: 'employee', Id: 15}},
  {Type: {Id: 105}, Description: 'x', RequestedBy: {Type: 'employee', Id: 15}},
  {Type: {Id: 104}, Description: ''}, {Type: {id: 105}, Description: '', SalesOrder: {Id: 1}, RequestedBy: {Type: 'employee', Id: 1}},
  {Type: {Id: 105}, Description: '', InfoCustomFields: [], RequestedBy: {Type: 'employee', Id: 1}}, {}, null];
sameOver('ENFORCE_CREATE_PAYLOAD', 'tmv3_enforcePreInspectionCreatePayloadHardRules_', createPayloads.map(p => [p]));
sameOver('ASSERT_CREATE_PAYLOAD_TOP_LEVEL', 'tmv3_assertPreInspectionCreatePayloadHardRules_', createPayloads.map(p => [p]));
sameOver('SAFE_TASK_PATCH_TOP_LEVEL', 'tmv3_piAssertSafeTaskPatch_', [[{Id: 1, Title: 'A - 1 St, Town - (416) 555-0100'}], [{Description: ''}],
  [{requestedBy: {}}], [{RequestedBy: {Type: 'contact', Id: 3}}], [{Location: {Id: 3}}], [{Title: 'bad'}], [{Type: {Id: 105}}]]);
sameOver('PARITY_EQUAL', 'tmv3_preInspectionCalendarParityEqual_', [[{title: 'a'}, {title: 'a'}], [{title: 'a'}, {title: 'b'}], [null, undefined],
  [{requiredStephenPresent: true}, {requiredStephenPresent: false}]]);
sameOver('GUEST_SYNC_DECISION', 'tmv3_preInspectionGuestSyncDecision_', [[[], 's', ''], [[], 's', 'o']]);
sameOver('CREATE_POLICY', 'tmv3_preInspectionCreatePolicy', [[]]);
sameOver('REQUESTED_BY_CONTRACT_RULE', 'tmv3_piRequestedByContractFromRule_', [[{requestedByType: 'contact'}], [{}], [null], [{requestedByType: ' Employee '}]]);
sameOver('LOCKED_STANDARD', 'tmv3_piAssertLockedStandard_', [[]]);
sameOver('ENSURE_MIRROR', 'tmv3_ensurePreInspectionMirror_', one.slice(0, 3));
sameOver('REJECTED_ADDRESS_GUARD', 'tmv3_preInspectionRejectedAddressGuard_', [
  [{state: 'UNCERTAIN', error: 'HTTP 400 Address object is required'}, {owned: [], matches: []}],
  [{state: 'UNCERTAIN', addressSchema: 'nested-address-v1', error: 'HTTP 400 country_Invalid country code Canada could not be read'}, {owned: [], matches: []}],
  [{state: 'UNCERTAIN', error: 'HTTP 500'}, {owned: [], matches: []}], [null, null], [{state: 'DONE'}, {owned: [], matches: []}]]);
const tasks = [
  {'Task ID': '10', 'Task Type ID': 105, 'Customer ID': '100', 'Location ID': '200', Status: 'Open', Name: 'Sample Customer - 10 Example St - (416) 555-0100'},
  {'Task ID': '11', 'Task Type ID': 105, 'Customer ID': '100', 'Location ID': '999', Status: 'Open', Name: 'Other'},
  {'Task ID': '12', 'Task Type ID': 105, 'Customer ID': '100', 'Location ID': '200', Status: 'Completed', Name: 'Sample Customer'},
  {'Task ID': '13', 'Task Type ID': 105, 'Customer ID': '100', 'Location ID': '200', Status: 'On Hold', Name: 'Sample Customer'},
  {'Task ID': '19694', 'Task Type ID': 105, 'Customer ID': '100', 'Location ID': '', Status: 'Open', Name: 'x'},
  {'Task ID': '14', 'Task Type ID': 104, 'Customer ID': '100', Status: 'Open'}];
function timed(t, r, offsetMin) { return Object.assign({}, t, {Start: new Date(Date.parse(r.start) + offsetMin * 60000).toISOString(), Due: r.end}); }
const taskCases = [];
RECORDS.slice(0, 14).forEach(r => {
  [[tasks[0]], [tasks[1]], [tasks[0], tasks[1]], [tasks[2]], [tasks[3]], [tasks[4]], [tasks[5]], [], [tasks[0], tasks[2]]].forEach(set => {
    [0, 90, 1440].forEach(off => taskCases.push([r, CUSTOMERS[0], {'Location ID': '200'}, set.map(t => timed(t, r, off))]));
  });
  taskCases.push([r, {'Customer ID': 'x'}, null, []]);
});
sameOver('TASK_DECISION', 'tmv3_preInspectionTaskDecision_', taskCases);
sameOver('STEP5_RESOLVE', 'tmv3_step5ResolvePreInspection_', taskCases.map(c => [Object.assign({}, c[0], {step4: {customer: c[1], location: c[2]}}),
  {tasksByCustomer: {'100': c[3]}, taskById: Object.fromEntries(c[3].map(t => [t['Task ID'], t]))}]));
sameOver('CANONICAL_CACHE_TASK', 'tmv3_preInspectionCanonicalCacheTask_', [
  [tasks[0], {taskId: 10, startDateTime: '2026-11-11T13:00:00-05:00', dueDateTime: '2026-11-11T14:00:00-05:00'}],
  [tasks[0], {taskId: 10, startDateTime: '2026-11-11T15:00:00-05:00', dueDateTime: '2026-11-11T14:00:00-05:00'}],
  [tasks[0], null], [tasks[5], null], [tasks[0], {taskId: 11, startDateTime: 'x', dueDateTime: 'y'}]]);
const plans = [
  {vertical: 'PreInspection', disposition: 'FULFILLED_NO_RECREATE', plan: 'VERIFY_ONLY_HISTORY_NO_WRITE', readStatus: 'FRESH_TASK_GET', taskId: '12', actions: [],
    taskStatus: 'Completed', customerCheck: 'MATCH', expectedCustomerId: '100', actualCustomerId: '100', locationCheck: 'MATCH',
    calendarStart: RECORDS[3].start, taskStart: RECORDS[3].start, taskDue: RECORDS[3].end, existingTaskId: '12'},
  {plan: 'NO_CHANGE', readStatus: 'FRESH_TASK_GET', taskId: '10', actions: [], customerCheck: 'MATCH', locationCheck: 'MATCH', requestedByCheck: 'MATCH',
    startCheck: 'MATCH', endCheck: 'MATCH', assignmentCheck: 'MATCH', calendarTitleCheck: 'MATCH', taskNameCheck: 'MATCH'},
  {plan: 'NO_CHANGE', readStatus: 'CACHE', taskId: '10', actions: []}, {plan: 'NO_CHANGE', readStatus: 'FRESH_TASK_GET', taskId: '10', actions: [], customerCheck: 'MATCH'},
  null, {}];
sameOver('HISTORICAL_COVERAGE', 'tmv3_preInspectionHistoricalCoverageVerified_', plans.map(p => [p]));
sameOver('COMPLETION_VERIFIED', 'tmv3_preInspectionCompletionVerified_', plans.map(p => [[p], {}]).concat([[[], {}], [null, {}], [plans.slice(0, 2), {}]]));
check('SELF_TEST_ORIGINAL_31_UNCHANGED_PLUS_5_NEW', () => {
  const x = load(BASELINE).ctx.tmv3_runPreInspectionCleanedSelfTest(), y = load(REBUILT).ctx.tmv3_runPreInspectionCleanedSelfTest();
  assert.equal(x.checks, 31); assert.equal(y.checks, 36); assert.equal(y.status, 'SYNTHETIC_SELF_TEST_PASS');
  assert.deepEqual(JSON.parse(JSON.stringify(y.results.slice(0, 31))), JSON.parse(JSON.stringify(x.results)));
});

// Step 3 anchor with synthetic refs.
check('STEP3_ANCHOR', () => {
  const refs = {customerByNumber: {'12345': CUSTOMERS[0], '54321': CUSTOMERS[1]}, customerById: {'100': CUSTOMERS[0], '200': CUSTOMERS[1], '300': CUSTOMERS[2]},
    ordersByNumber: {'586828': [{'Order ID': '9', 'Order Number': '586828', 'Customer ID': '300', 'Order Type': 'SALES_ORDER', Status: 'Quoted'}],
      '123456': [{'Order ID': '7', 'Order Number': '123456', 'Customer ID': '100', 'Order Type': 'SALES_ORDER', Status: 'Quoted'}]}};
  const a = load(BASELINE), b = load(REBUILT);
  RECORDS.forEach((r, i) => {
    const x = outcome(() => a.ctx.tmv3_step3ResolvePreInspectionAnchor_(clone(r), refs));
    const y = outcome(() => b.ctx.tmv3_step3ResolvePreInspectionAnchor_(clone(r), refs));
    assert.deepEqual(y, x, 'record ' + i);
  });
});

// ---------------------------------------------------------------- 3 Striven read adapters (stubbed HTTP)
function scriptedFetch(env, responses) {
  env.ctx.tmv3_fetchJson_ = function(url, opts) {
    env.log.push({kind: 'FETCH', url: url, payload: opts && opts.payload});
    for (const r of responses) if (r.match.test(url)) return clone(r.body);
    throw new Error('UNSCRIPTED_FETCH ' + url);
  };
}
const orderCases = [
  {status: {Id: 19, Name: 'Quoted'}}, {status: {Id: 22, Name: 'Approved'}}, {status: {Id: 25, Name: 'In Progress'}},
  {status: {Id: 18, Name: 'Incomplete'}}, {status: {Id: 20, Name: 'Pending Approval'}}, {status: {Id: 30, Name: 'Closed'}},
  {status: {Id: 19, Name: 'Quoted'}, customer: 999}, {status: {Id: 19, Name: 'Quoted'}, total: 2}];
check('ORDER_BEFORE_CREATE_MATCHES_BASELINE_FOR_CONSISTENT_STATUS', () => {
  orderCases.forEach((c, i) => {
    const run = src => {
      const env = load(src);
      scriptedFetch(env, [
        {match: /sales-orders\/search/, body: {Data: c.total === 2 ? [{Id: 5, Number: '586828'}, {Id: 6, Number: '586828'}] : [{Id: 5, Number: '586828'}], TotalCount: c.total || 1}},
        {match: /sales-orders\/5$/, body: {Id: 5, OrderNumber: '586828', Customer: {Id: c.customer || 100}, Status: c.status}}]);
      return [outcome(() => env.ctx.tmv3_preInspectionOrderBeforeCreate_(Object.assign({}, RECORDS[0]), {expectedCustomerId: 100})), env.log];
    };
    const base = run(BASELINE);
    assert.ok(!/UNSCRIPTED/.test(String(base[0].error || '')), 'fixture not exercised: ' + base[0].error);
    assert.deepEqual(run(REBUILT), base, 'case ' + i);
  });
  ['', 'abc', '12'].forEach(n => {
    const x = outcome(() => load(BASELINE).ctx.tmv3_preInspectionOrderBeforeCreate_(Object.assign({}, RECORDS[0], {orderNumber: n}), {expectedCustomerId: 100}));
    const y = outcome(() => load(REBUILT).ctx.tmv3_preInspectionOrderBeforeCreate_(Object.assign({}, RECORDS[0], {orderNumber: n}), {expectedCustomerId: 100}));
    assert.deepEqual(y, x);
  });
});
check('ORDER_BEFORE_CREATE_NEW_GUARD_STATUS_ID_MUST_MATCH_NAME', () => {
  const env = load(REBUILT);
  scriptedFetch(env, [{match: /search/, body: {Data: [{Id: 5, Number: '586828'}], TotalCount: 1}},
    {match: /sales-orders\/5$/, body: {Id: 5, OrderNumber: '586828', Customer: {Id: 100}, Status: {Id: 22, Name: 'Quoted'}}}]);
  assert.match(outcome(() => env.ctx.tmv3_preInspectionOrderBeforeCreate_(RECORDS[0], {expectedCustomerId: 100})).error, /STATUS_ID_NAME_MISMATCH/);
  const base = load(BASELINE);
  scriptedFetch(base, [{match: /search/, body: {Data: [{Id: 5, Number: '586828'}], TotalCount: 1}},
    {match: /sales-orders\/5$/, body: {Id: 5, OrderNumber: '586828', Customer: {Id: 100}, Status: {Id: 22, Name: 'Quoted'}}}]);
  assert.equal(outcome(() => base.ctx.tmv3_preInspectionOrderBeforeCreate_(RECORDS[0], {expectedCustomerId: 100})).ok.status, 'SO_CORROBORATION_ONLY_NOT_ATTACHED');
});
check('CUSTOMER_LOCATIONS_FRESH', () => {
  const bodies = [{TotalCount: 0, Data: null}, {TotalCount: 1, Data: [{Id: 200, Customer: {Id: 100}, Address: {Address1: '10 Example St', City: 'Toronto'}}]},
    {TotalCount: 2, Data: [{Id: 200, Customer: {Id: 100}, Address: {Address1: 'a'}}]}, {TotalCount: 1, Data: [{Id: 200, Customer: {Id: 7}, Address: {Address1: 'a'}}]},
    {TotalCount: 2, Data: [{Id: 200, Customer: {Id: 100}, Address: {Address1: 'a'}}, {Id: 200, Customer: {Id: 100}, Address: {Address1: 'a'}}]}, null];
  bodies.forEach((body, i) => {
    const run = src => { const env = load(src); scriptedFetch(env, [{match: /locations/, body}]); return outcome(() => env.ctx.tmv3_preInspectionCustomerLocationsFresh_(100)); };
    assert.deepEqual(run(REBUILT), run(BASELINE), 'case ' + i);
  });
});
check('TASK_ID_SEARCH_AND_HISTORY', () => {
  const page = (n, start) => ({Data: Array.from({length: n}, (_, i) => ({Id: start + i})), TotalCount: n});
  const scenarios = [[page(3, 1)], [page(100, 1), page(2, 101)], [page(100, 1), page(100, 1)], [page(100, 1), page(100, 101), page(100, 201)], [{}]];
  scenarios.forEach((pages, si) => {
    const run = src => {
      const env = load(src); let k = 0;
      env.ctx.tmv3_fetchJson_ = () => clone(pages[Math.min(k++, pages.length - 1)]);
      env.ctx.tmv3_getTaskById_ = id => ({'Task ID': String(id), 'Task Type ID': 105, 'Customer ID': '100'});
      return [outcome(() => env.ctx.tmv3_searchPreInspectionTaskIdsForCustomer_(CUSTOMERS[0])),
        (k = 0, outcome(() => env.ctx.tmv3_searchPreInspectionTasks_(CUSTOMERS[0])))];
    };
    assert.deepEqual(run(REBUILT), run(BASELINE), 'scenario ' + si);
  });
});
check('STEP7_HYDRATE_LINKED_TASK', () => {
  const run = src => {
    const env = load(src);
    env.ctx.tmv3_step7FreshTask_ = id => ({ok: true, task: {'Task ID': id, 'Task Type ID': 105, 'Customer ID': '100', 'Order ID': '9'}});
    env.ctx.tmv3_upsertTaskCacheRow_ = t => env.log.push({kind: 'UPSERT', id: t['Task ID']});
    const refs = {taskById: {}, tasks: [{'Task ID': '19694'}], tasksByCustomer: {'100': [{'Task ID': '19694'}]}, tasksByOrder: {'9': [{'Task ID': '19694'}]}};
    const recs = [Object.assign({}, RECORDS[0], {step4: {disposition: 'VERIFIED', customer: CUSTOMERS[0]}}),
      Object.assign({}, RECORDS[7], {existingTaskId: 'x1', step4: {disposition: 'VERIFIED', customer: CUSTOMERS[0]}})];
    return [outcome(() => env.ctx.tmv3_step7HydrateLinkedPreInspectionTasks_(recs.slice(0, 1), refs, {})), refs, env.log,
      outcome(() => env.ctx.tmv3_step7HydrateLinkedPreInspectionTasks_(recs.slice(1), refs, {}))];
  };
  assert.deepEqual(run(REBUILT), run(BASELINE));
});
check('VERIFY_AUTOMATION_DIAGNOSTIC', () => {
  const rows = RECORDS.slice(0, 8).map((r, i) => ({'Vertical': 'PreInspection', 'Calendar Role': i % 3 ? 'PRIMARY_SHARED' : 'SECONDARY',
    'Calendar Start': r.start.replace('T', ' ').slice(0, 19), 'Event ID': r.eventId, 'Existing Task Link': i % 2 ? 'https://x/TaskInfo.aspx?TaskID=1' + i : '',
    'Task': 'Task #1' + (i % 4 ? i : 9), 'Mapping Status': ['STEP 6 — MATCH_EXISTING', 'REVIEW', 'STEP 6 — CREATE_TASK', ''][i % 4],
    'Data Checklist': i % 2 ? 'Customer ✅\nAssignee ⏳' : 'Location ⚠ x', 'Customer # (Calendar)': '12345', Fingerprint: 'f' + i}));
  const run = src => {
    const env = load(src); env.ctx.tmv3_rows_ = () => clone(rows);
    return [outcome(() => env.ctx.tmv3_verifyPreInspectionAutomation()), outcome(() => env.ctx.tmv3_verifyPreInspectionAutomation()),
      outcome(() => env.ctx.tmv3_verifyPreInspectionAutomation(RECORDS[4].eventId)), env.props.store];
  };
  const x = run(BASELINE), y = run(REBUILT);
  x.forEach(o => o && o.ok && delete o.ok.elapsedMs); y.forEach(o => o && o.ok && delete o.ok.elapsedMs);
  assert.deepEqual(y, x);
});

// ---------------------------------------------------------------- 4 Calendar review sheet (read-only sources)
check('PI_CALENDAR_REVIEW_SHEET', () => {
  const mk = (id, cal, title, desc, creators, dayOffset, loc) => ({id, cal, title, desc, creators, start: NOW + dayOffset * 86400000, loc});
  const evs = [mk('a1', 'CF Preinspects', '12345 - Sample - (416) 555-0100', 'Task #10', ['adam@classicfireplace.ca'], 1, '10 Example'),
    mk('a1', 'Stephen', '12345 - Sample - (416) 555-0100', 'Task #10', ['adam@classicfireplace.ca'], 1, '10 Example'),
    mk('a2', 'CF Preinspects', 'Jane', 'x', ['adam@classicfireplace.ca'], 2, ''), mk('a2', 'Stephen', 'Jane changed', 'x', ['adam@classicfireplace.ca'], 2, ''),
    mk('b1', 'Jay', 'Measure fireplace insert', '', ['jay@classicfireplace.ca'], 3, '1 Rd'), mk('b2', 'Karen', 'Team meeting', '', ['karen@classicfireplace.ca'], 3, ''),
    mk('b3', 'Karen', 'Service call', '', ['karen@classicfireplace.ca'], 4, ''), mk('b4', 'Jay', 'Call 905-555-1234', '', ['x@y.com'], 5, ''),
    mk('b5', 'Stephen', 'Lunch', '', ['stephen@classicfireplace.ca'], 6, ''), mk('b6', 'Jay', 'Delivery', '', ['service@classicfireplace.ca'], 6, ''),
    mk('b7', 'Karen', 'Random', '', ['q@z.com'], 7, '')];
  const run = (src, denyKaren) => {
    const env = load(src);
    const ids = {'CF Preinspects': PRIMARY, Stephen: STEPHEN_CAL, Jay: 'jay@classicfireplace.ca', Karen: 'karen@classicfireplace.ca'};
    env.ctx.CalendarApp = {getCalendarById: id => {
      const name = Object.keys(ids).find(k => ids[k] === id);
      if (denyKaren && name === 'Karen') throw new Error('denied');
      return {getEvents: () => evs.filter(e => e.cal === name).map(e => ({
        getId: () => e.id + '@google.com', getStartTime: () => new env.ctx.Date(e.start), getEndTime: () => new env.ctx.Date(e.start + 3600000),
        getTitle: () => e.title, getLocation: () => e.loc, getDescription: () => e.desc, getCreators: () => e.creators, isAllDayEvent: () => false}))};
    }};
    const cells = {}; const writes = [];
    const old = [['DONE', '2026-10-01T09:00:00-04:00', '', 'Old', 'Karen', '', '', '', 'zz', '', '', '', '', '', '', '', '', '', 'keep', 'note']];
    const sheet = {getLastRow: () => 3 + old.length, getMaxRows: () => 1000, insertRowsAfter() {},
      getRange: (r, c, nr, nc) => ({getValues: () => clone(old), getDisplayValue: () => cells[r + ':' + c] || '',
        getDisplayValues: () => Array.from({length: nr || 1}, () => Array(nc || 1).fill('')),
        setValue: v => { cells[r + ':' + c] = v; writes.push(['setValue', r, c, v]); },
        setValues: v => { if (r === 3) cells['3:1'] = v[0][0]; writes.push(['setValues', r, c, v]); },
        clearContent: () => writes.push(['clear', r, nr])})};
    env.ctx.SpreadsheetApp = {openById: () => ({getSheetByName: () => sheet}), flush() {}};
    return clone([outcome(() => env.ctx.tmv3_piCalendarReviewRefresh_()), writes]);
  };
  assert.deepEqual(run(REBUILT, false), run(BASELINE, false));
  assert.deepEqual(run(REBUILT, true), run(BASELINE, true));
});

// ---------------------------------------------------------------- 5 gate closed: no side effects at any entry point
check('GATE_CLOSED_ENTRY_POINTS_MATCH_AND_DO_NOTHING', () => {
  const entries = [['tmv3_preInspectionAutomationCycle_', [{eventId: 'e@google.com'}]], ['tmv3_runPreInspectionAutomation', []],
    ['tmv3_reconcilePreInspectionEvent', ['e@google.com']], ['tmv3_runPreInspectionDailySlot_', ['08:00']],
    ['tmv3_preInspectionCheck0800', []], ['tmv3_preInspectionCheck1300', []], ['tmv3_preInspectionCheck1700', []],
    ['tmv3_installPreInspectionAutomationTrigger', []], ['tmv3_preInspectionClearContinuationTrigger_', []],
    ['tmv3_preInspectionAutomaticRecovery_', ['x']], ['tmv3_reconcilePreInspectionPendingReview', []],
    ['tmv3_preInspectionOptionalOrderLocationEvidence_', [RECORDS[0]]], ['tmv3_preInspectionSendMissingDetails_', [RECORDS[0]]],
    ['tmv3_preInspectionFailureAlert_', ['e', 'err']], ['tmv3_preInspectionContinueEligibleSweep', []], ['tmv3_preInspectionScheduledCycle', []]];
  entries.forEach(([fn, args]) => {
    const run = src => { const env = load(src); return [outcome(() => env.ctx[fn].apply(null, clone(args))), env.log, env.props.store]; };
    const x = run(BASELINE), y = run(REBUILT);
    assert.deepEqual(y, x, fn);
    assert.deepEqual(y[1], [], fn + ' had side effects');
  });
});

// ---------------------------------------------------------------- 6 gate OPEN: full cycle scenarios with stubbed V3 core
const OPEN = (f, t) => f === '30_PreInspection.gs' ? BASELINE_RULES(t.replace(/^var TMV3_PI_CLEANED_RELEASE_APPROVED = (?:false|true);/m, 'var TMV3_PI_CLEANED_RELEASE_APPROVED = true;')
  .replace(/^var TMV3_PI_CREATOR_EMAILS_ENABLED = false;/m, 'var TMV3_PI_CREATOR_EMAILS_ENABLED = true;')) : t;
function cycleEnv(src, scenario) {
  return cycleEnvOn(load(src, {transform: OPEN}), scenario);
}
function cycleEnvOn(env, scenario) {
  const c = env.ctx, trace = env.log;
  const note = (k, extra) => trace.push(Object.assign({kind: k}, extra || {}));
  const recs = scenario.records.map(clone);
  c.tmv3_operationWritesEnabled_ = () => true;
  c.tmv3_verticalCalendars_ = () => [{calendarId: PRIMARY, calendarName: 'CF Preinspects', role: 'PRIMARY_SHARED'}];
  c.CalendarApp = {getCalendarById: id => id === PRIMARY ? {getEvents: () => recs.map(r => ({__r: r}))} : null};
  c.tmv3_calendarEvent_ = (v, cfg, src2, ev) => clone(ev.__r);
  c.tmv3_findCalendarEventById_ = (cal, id) => { const r = recs.concat(scenario.offline || []).find(x => x.eventId === id); return r && cal === r.calendarId ? {event: {__r: r}} : null; };
  c.tmv3_strivenApiBudgetStatus_ = () => clone(scenario.budget || {exhausted: false, remaining: 500, count: 1, date: '2026-11-10', softLimit: 900});
  c.tmv3_refreshOperationalSourcesIfNeeded_ = () => note('REFRESH');
  c.tmv3_listTriggers = () => [];
  c.tmv3_audit_ = function() { note('AUDIT', {args: Array.from(arguments).slice(1, 5)}); };
  c.tmv3_rows_ = () => clone(CUSTOMERS);
  const planCalls = {};
  c.tmv3_step7FreshPlansForEvent_ = (v, id) => {
    planCalls[id] = (planCalls[id] || 0) + 1; note('PLAN', {id, n: planCalls[id]});
    const seq = scenario.plans[id] || [[{plan: 'NO_CHANGE', blocker: 'X', actions: []}]];
    const p = seq[Math.min(planCalls[id] - 1, seq.length - 1)];
    if (p === 'THROW_BUDGET') throw new Error('TMV3_STRIVEN_API_DAILY_GUARD reached');
    if (p === 'THROW') throw new Error('PLAN_READ_FAILED');
    return clone(p);
  };
  c.tmv3_executeFreshStep7Selection_ = (sel, scope) => { note('EXECUTE', {plan: sel.previousPlan, scope}); return clone(scenario.execution || {status: 'VERIFIED'}); };
  c.tmv3_findFreshEventRecord_ = (v, id) => clone(recs.concat(scenario.offline || []).find(r => r.eventId === id));
  c.tmv3_step4IdentityIndex_ = () => ({customerById: {'100': CUSTOMERS[0]}});
  c.tmv3_fetchJson_ = (url, o) => {
    note('FETCH', {url});
    if (/tasks\/search/.test(url)) return {Data: [{Id: 10}], TotalCount: 1};
    if (/sales-orders\/search/.test(url)) return {Data: [{Id: 5, Number: '586828'}], TotalCount: 1};
    if (/sales-orders\/5$/.test(url)) return {Id: 5, OrderNumber: '586828', Customer: {Id: 100}, Status: {Id: 19, Name: 'Quoted'}};
    throw new Error('UNSCRIPTED ' + url);
  };
  c.tmv3_getTaskById_ = id => ({'Task ID': String(id), 'Task Type ID': 105, 'Customer ID': '100'});
  c.tmv3_upsertTaskCacheRow_ = t => note('UPSERT', {id: t['Task ID']});
  c.tmv3_refreshAutomationSheetsAfter_ = () => ({status: 'STUB_SHEETS'});
  c.tmv3_piCalendarReviewRefresh_ = () => clone(scenario.review || {status: 'PI_REVIEW_UNCHANGED'});
  c.tmv3_preInspectionRefreshOperatorView_ = () => ({status: 'PI_OPERATOR_VIEW_REFRESHED'});
  Object.entries(scenario.props || {}).forEach(([k, v]) => env.props.setProperty(k, typeof v === 'string' ? v : JSON.stringify(v)));
  return env;
}
function rec(i, extra) {
  return Object.assign(record(i), {calendarRole: 'PRIMARY_SHARED', sourceCalendarRoles: ['PRIMARY_SHARED'], calendarId: PRIMARY,
    status: 'confirmed', isAllDay: false, title: '12345 - Sample Customer - (416) 555-0100', organizerCalendarId: 'stephen@classicfireplace.ca',
    creator: 'adam@classicfireplace.ca', start: new Date(NOW + (i + 1) * 86400000).toISOString(),
    end: new Date(NOW + (i + 1) * 86400000 + 3600000).toISOString()}, extra || {});
}
const CONVERGED = {plan: 'NO_CHANGE', readStatus: 'FRESH_TASK_GET', taskId: '10', actions: [], customerCheck: 'MATCH', locationCheck: 'MATCH',
  requestedByCheck: 'MATCH', startCheck: 'MATCH', endCheck: 'MATCH', assignmentCheck: 'MATCH', calendarTitleCheck: 'MATCH', taskNameCheck: 'MATCH'};
const PATCH = {plan: 'PATCH_TASK', actions: ['PATCH_TASK_RELATIONSHIPS'], taskId: '10', expectedRequestedByType: 'employee', expectedRequestedById: '15'};
const VERIFY_LINKS = {plan: 'LINKS', actions: ['VERIFY_CALENDAR_LINKS'], taskId: '10'};
const SCENARIOS = {
  converged: {records: [rec(0), rec(1)], plans: {'evt0@google.com': [[CONVERGED]], 'evt1@google.com': [[CONVERGED]]}},
  patchThenConverge: {records: [rec(0)], plans: {'evt0@google.com': [[PATCH], [CONVERGED]]}},
  twoWritesDefer: {records: [rec(0), rec(1), rec(2)], plans: {'evt0@google.com': [[PATCH], [CONVERGED]], 'evt1@google.com': [[PATCH]], 'evt2@google.com': [[CONVERGED]]}},
  createPath: {records: [rec(0, {orderNumber: '586828'})], plans: {'evt0@google.com': [[{plan: 'CREATE', actions: ['CREATE_TASK'], expectedCustomerId: '100',
    expectedRequestedByType: 'employee', expectedRequestedById: '15'}], [{plan: 'CREATE', actions: ['CREATE_TASK'], expectedCustomerId: '100',
    expectedRequestedByType: 'employee', expectedRequestedById: '15'}], [CONVERGED]]}},
  recreateBlocked: {records: [rec(0)], plans: {'evt0@google.com': [[{plan: 'RECREATE', actions: ['RECREATE_TASK'], expectedCustomerId: '100'}]]}},
  wrongRequester: {records: [rec(0, {organizerCalendarId: 'karen@classicfireplace.ca'})], plans: {'evt0@google.com': [[PATCH]]}},
  multiplePlans: {records: [rec(0)], plans: {'evt0@google.com': [[CONVERGED, CONVERGED]]}},
  linksOnly: {records: [rec(0)], plans: {'evt0@google.com': [[VERIFY_LINKS], [CONVERGED]]}},
  executionUnverified: {records: [rec(0)], execution: {status: 'FAILED'}, plans: {'evt0@google.com': [[PATCH]]}},
  budgetErrorMidBatch: {records: [rec(0), rec(1), rec(2)], plans: {'evt0@google.com': [[CONVERGED]], 'evt1@google.com': ['THROW_BUDGET']}},
  readFailure: {records: [rec(0), rec(1)], plans: {'evt0@google.com': ['THROW'], 'evt1@google.com': [[CONVERGED]]}},
  budgetLow: {records: [rec(0)], budget: {exhausted: false, remaining: 10, count: 990, date: '2026-11-10', softLimit: 900}, plans: {}},
  hold: {records: [rec(0, {eventId: '5mafnqht16o2hrqv4232rofmm3@google.com'}), rec(1)], plans: {'evt1@google.com': [[CONVERGED]]}},
  doneCache: {records: [rec(0)], props: {['TMV3_PI_DONE_' + 'X']: 'unused'}, plans: {'evt0@google.com': [[CONVERGED]]}},
  missingDetailsEmail: {records: [rec(0, {location: '', calendarCreatedAt: new Date(NOW - 86400000).toISOString()}),
    rec(5, {location: '', creator: 'stephen@classicfireplace.ca'}), rec(1, {title: 'Jane - (416) 555-0100', customerNumber: ''})],
  plans: {'evt0@google.com': [[CONVERGED]], 'evt5@google.com': [[CONVERGED]], 'evt1@google.com': [[CONVERGED]]}},
  uncertainPriorSend: {records: [rec(0, {location: ''})], props: {}, plans: {'evt0@google.com': [[CONVERGED]]}},
  pendingRecovery: {records: [rec(0)], offline: [rec(9, {calendarRole: 'SECONDARY_STEPHEN', calendarId: STEPHEN_CAL, title: 'Lunch'})],
    props: {TMV3_PI_PENDING_EVENTS: ['evt9@google.com', 'evt0@google.com']}, plans: {'evt0@google.com': [[CONVERGED]]}},
  historicalFulfilled: {records: [rec(3)], plans: {'evt3@google.com': [[Object.assign({}, plans[0], {calendarStart: rec(3).start, taskStart: rec(3).start, taskDue: rec(3).end})]]}},
  blockerOrderEvidence: {records: [rec(0)], plans: {'evt0@google.com': [[{plan: 'REVIEW', blocker: 'AMBIGUOUS', actions: []}]]}}
};
const RUNS = [
  ['tmv3_preInspectionAutomationCycle_', [{force: true, reason: 'T'}]],
  ['tmv3_preInspectionAutomationCycle_', [{force: true, eventId: 'evt0@google.com'}]],
  ['tmv3_preInspectionAutomationCycle_', [{force: true, sweepStart: true}]],
  ['tmv3_preInspectionAutomationCycle_', [{force: true, sweepContinue: true}]],
  ['tmv3_preInspectionAutomationCycle_', [{}]],
  ['tmv3_runPreInspectionDailySlot_', ['08:00']],
  ['tmv3_reconcilePreInspectionPendingReview', []],
  ['tmv3_runPreInspectionAutomation', []]
];
// Intentional change (operator rule 2026-11-11): no open Task -> create a new one.
const RECREATE_CHANGED = new Set(['recreateBlocked']);
check('NEW_RULE_NO_OPEN_TASK_RECREATES', () => {
  const env = cycleEnv(REBUILT, Object.assign({}, SCENARIOS.recreateBlocked, {plans: {'evt0@google.com': [
    [{plan: 'RECREATE', actions: ['RECREATE_TASK'], expectedCustomerId: '100', expectedRequestedByType: 'employee', expectedRequestedById: '15'}],
    [{plan: 'RECREATE', actions: ['RECREATE_TASK'], expectedCustomerId: '100', expectedRequestedByType: 'employee', expectedRequestedById: '15'}],
    [CONVERGED]]}}));
  const r = env.ctx.tmv3_preInspectionAutomationCycle_({force: true});
  assert.equal(r.writes, 1); assert.equal(r.status, 'COMPLETE');
  assert.equal(r.results[0].recreateAuthorized, 'NO_OPEN_TASK_NEW_VISIT_OPERATOR_RULE_2026_11_11');
  assert.equal(env.log.filter(e => e.kind === 'EXECUTE').length, 1);
  assert.ok(env.log.some(e => e.kind === 'FETCH' && /tasks\/search/.test(e.url)), 'fresh history must be read before recreate');
  const base = cycleEnv(BASELINE, SCENARIOS.recreateBlocked);
  assert.match(base.ctx.tmv3_preInspectionAutomationCycle_({force: true}).results[0].error, /RECREATE_DISABLED/);
  // Same-day completed Task in fresh history: no new Task.
  const sameDay = cycleEnv(REBUILT, Object.assign({}, SCENARIOS.recreateBlocked, {plans: {'evt0@google.com': [
    [{plan: 'RECREATE', actions: ['RECREATE_TASK'], expectedCustomerId: '100', expectedRequestedByType: 'employee', expectedRequestedById: '15'}]]}}));
  const ev = SCENARIOS.recreateBlocked.records[0];
  sameDay.ctx.tmv3_getTaskById_ = id => ({'Task ID': String(id), 'Task Type ID': 105, 'Customer ID': '100', Status: 'Completed', Start: ev.start});
  const sd = sameDay.ctx.tmv3_preInspectionAutomationCycle_({force: true});
  assert.match(sd.results[0].error, /PI_SAME_DAY_COMPLETED_TASK_NO_RECREATE:10/); assert.equal(sd.writes, 0);
  assert.equal(sameDay.log.filter(e => e.kind === 'EXECUTE').length, 0);
  const undated = cycleEnv(REBUILT, Object.assign({}, SCENARIOS.recreateBlocked, {plans: {'evt0@google.com': [
    [{plan: 'RECREATE', actions: ['RECREATE_TASK'], expectedCustomerId: '100', expectedRequestedByType: 'employee', expectedRequestedById: '15'}]]}}));
  undated.ctx.tmv3_getTaskById_ = id => ({'Task ID': String(id), 'Task Type ID': 105, 'Customer ID': '100', Status: 'Completed'});
  assert.match(undated.ctx.tmv3_preInspectionAutomationCycle_({force: true}).results[0].error, /UNDATED_REVIEW_NO_RECREATE/);
  // Older completed Task on another day: new Task allowed.
  const older = cycleEnv(REBUILT, Object.assign({}, SCENARIOS.recreateBlocked, {plans: {'evt0@google.com': [
    [{plan: 'RECREATE', actions: ['RECREATE_TASK'], expectedCustomerId: '100', expectedRequestedByType: 'employee', expectedRequestedById: '15'}],
    [{plan: 'RECREATE', actions: ['RECREATE_TASK'], expectedCustomerId: '100', expectedRequestedByType: 'employee', expectedRequestedById: '15'}], [CONVERGED]]}}));
  older.ctx.tmv3_getTaskById_ = id => ({'Task ID': String(id), 'Task Type ID': 105, 'Customer ID': '100', Status: 'Completed', Start: '2026-09-01T13:00:00-04:00'});
  assert.equal(older.ctx.tmv3_preInspectionAutomationCycle_({force: true}).writes, 1);
  const blocked = cycleEnv(REBUILT, Object.assign({}, SCENARIOS.recreateBlocked, {plans: {'evt0@google.com': [
    [{plan: 'RECREATE', disposition: 'FULFILLED_NO_RECREATE', actions: ['RECREATE_TASK'], expectedCustomerId: '100', expectedRequestedByType: 'employee', expectedRequestedById: '15'}]]}}));
  assert.match(blocked.ctx.tmv3_preInspectionAutomationCycle_({force: true}).results[0].error, /SAME_VISIT_FULFILLED/);
});
Object.entries(SCENARIOS).filter(([name]) => !RECREATE_CHANGED.has(name)).forEach(([name, scenario]) => {
  RUNS.forEach(([fn, args]) => {
    check('CYCLE_' + name + '_' + fn + '_' + JSON.stringify(args), () => {
      const run = src => {
        const env = cycleEnv(src, scenario);
        if (name === 'doneCache') env.props.setProperty('TMV3_PI_DONE_' + env.ctx.tmv3_hash_('evt0@google.com'), JSON.stringify({fingerprint: 'fp0', at: NOW - 60000}));
        if (name === 'uncertainPriorSend') env.props.setProperty('TMV3_PI_NOTIFICATION_' + env.ctx.tmv3_hash_('evt0@google.com'), JSON.stringify({status: 'SENDING'}));
        const first = outcome(() => env.ctx[fn].apply(null, clone(args)));
        const second = outcome(() => env.ctx[fn].apply(null, clone(args)));
        return {first, second, trace: env.log, props: env.props.store};
      };
      const x = run(BASELINE), y = run(REBUILT);
      assert.deepEqual(y, x);
    });
  });
});

// ---------------------------------------------------------------- 7 V3 regression entry points: identical outcomes
check('V3_REGRESSION_SUITE_OUTCOMES_UNCHANGED', () => {
  const src = REGRESSION;
  const names = (src.match(/^function (tmv3_[A-Za-z0-9_]*)\s*\(\)/mg) || []).map(s => s.replace(/^function /, '').replace(/\s*\(\)$/, ''))
    .filter(n => !/^tmv3_(update|compare|old|v3|regressionKey)/.test(n));
  assert.ok(names.length >= 15);
  const a = load(BASELINE), b = load(REBUILT);
  names.forEach(n => {
    const x = outcome(() => a.ctx[n]()), y = outcome(() => b.ctx[n]());
    const norm = o => JSON.parse(JSON.stringify(o).replace(/"(elapsedMs|durationMs|revision)":("[^"]*"|\d+)/g, '"$1":0').replace(/\d{8}-\d{6}/g, 'TS'));
    assert.deepEqual(norm(y), norm(x), n);
  });
});

// ---------------------------------------------------------------- 8 intentional additions
check('NEW_DEEP_PAYLOAD_GUARD_CREATE_AND_PATCH', () => {
  const b = load(REBUILT), a = load(BASELINE);
  const nested = {Type: {Id: 105}, Description: '', RequestedBy: {Type: 'employee', Id: 15}, Location: {Id: 1, CustomFields: [{Id: 854}]}};
  assert.match(outcome(() => b.ctx.tmv3_assertPreInspectionCreatePayloadHardRules_(clone(nested))).error, /PI_PAYLOAD_REJECT_PROTECTED_FIELDS_NO_WRITE/);
  assert.equal(outcome(() => a.ctx.tmv3_assertPreInspectionCreatePayloadHardRules_(clone(nested))).ok, true);
  const patch = {Id: 1, Location: {Id: 3, Extra: [{fieldId: 852}]}};
  assert.match(outcome(() => b.ctx.tmv3_piAssertSafeTaskPatch_(clone(patch))).error, /PROTECTED_FIELDS/);
  const realCreate = {Title: 'Sample - 1 St, Town - (416) 555-0100', Type: {Id: 105, Name: 'Pre Inspection'}, Description: '', Budget: 0,
    StartDateTime: '11/11/2026 1:00 PM', DueDateTime: '11/11/2026 2:00 PM', UseSubContractor: false, Priority: {Id: 2, Name: 'Normal'},
    Customer: {Id: 100, Number: '12345', Name: 'Sample'}, Location: {Id: 200, Name: '1 St'}, RequestedBy: {Id: 15, Name: 'Stephen Foley', Type: 'employee'}};
  assert.equal(outcome(() => b.ctx.tmv3_assertPreInspectionCreatePayloadHardRules_(clone(realCreate))).ok, true);
  assert.equal(outcome(() => b.ctx.tmv3_piAssertSafeTaskPatch_({Id: 10, Location: {Id: 200}, RequestedBy: {Id: 15, Type: 'employee'}})).ok, true);
});
check('NEW_DRY_RUN_IS_PURE', () => {
  const b = load(REBUILT); withRows(b);
  const before = JSON.stringify(b.props.store);
  const r = rec(0, {location: '', calendarCreatedAt: new Date(NOW - 86400000).toISOString()});
  const out = b.ctx.tmv3_piDryRunRecord_(r, NOW);
  assert.equal(out.sideEffects, 0);
  const n = out.stages.find(s => s.stage === 'NOTIFICATION_DECISION').result;
  assert.equal(n.send, false); assert.equal(n.wouldSendIfReleased, true); assert.match(n.preview.subject, /MISSING DETAILS/);
  assert.equal(JSON.stringify(b.props.store), before);
  assert.deepEqual(b.log.filter(e => e.kind !== 'AUDIT'), []);
  const held = b.ctx.tmv3_piDryRunRecord_(rec(0, {eventId: '5mafnqht16o2hrqv4232rofmm3@google.com', location: ''}), NOW);
  assert.equal(held.stages.find(s => s.stage === 'NOTIFICATION_DECISION').result.status, 'OPERATOR_APPROVAL_HOLD_NO_SEND');
});

check('NEW_CREATOR_EMAILS_OFF_AND_LINK_ONLY_PLAN_CONVERGED', () => {
  const asShipped = (f, x) => f === '30_PreInspection.gs' ? BASELINE_RULES(x.replace(/^var TMV3_PI_CREATOR_EMAILS_ENABLED = true;/m, 'var TMV3_PI_CREATOR_EMAILS_ENABLED = false;')) : x; // emails switched OFF
  const sc = SCENARIOS.missingDetailsEmail;
  const run = (fn) => { const env = load(REBUILT, {transform: asShipped}); Object.assign(env, cycleEnvOn(env, sc)); return [env.ctx.tmv3_preInspectionAutomationCycle_({force: true}), env.log]; };
  const [r, log] = run();
  assert.equal(log.filter(e => e.kind === 'MAIL').length, 0); assert.equal(r.notifications, 0);
  const b = load(REBUILT);
  const linkOnly = Object.assign({}, CONVERGED, {actions: ['VERIFY_CALENDAR_LINKS']});
  assert.equal(b.ctx.tmv3_preInspectionCompletionVerified_([linkOnly], {}), true);
  assert.equal(b.ctx.tmv3_preInspectionCompletionVerified_([Object.assign({}, linkOnly, {actions: ['VERIFY_CALENDAR_LINKS', 'PATCH_TASK_RELATIONSHIPS']})], {}), false);
});

check('NEW_COMPLETED_TASK_NO_ALERT_AND_PLAIN_ALERT', () => {
  const sc = {records: [rec(0)], plans: {'evt0@google.com': [[PATCH]]}};
  const env = cycleEnv(REBUILT, sc);
  env.ctx.tmv3_executeFreshStep7Selection_ = () => { throw new Error('Fresh Task state is no longer OPEN.'); };
  env.ctx.tmv3_getTaskById_ = id => ({'Task ID': String(id), 'Task Type ID': 105, 'Customer ID': '100', Status: 'Completed'});
  const r = env.ctx.tmv3_preInspectionAutomationCycle_({force: true});
  assert.equal(r.results[0].status, 'TASK_COMPLETED_NO_ACTION'); assert.equal(r.status, 'COMPLETE');
  const alerts = l => l.filter(e => e.kind === 'MAIL' && /Automation: action required/.test(e.subject)).length;
  assert.equal(alerts(env.log), 0);
  const env2 = cycleEnv(REBUILT, sc);
  env2.ctx.tmv3_executeFreshStep7Selection_ = () => { throw new Error('Fresh Task state is no longer OPEN.'); };
  env2.ctx.tmv3_getTaskById_ = id => ({'Task ID': String(id), 'Task Type ID': 105, 'Customer ID': '100', Status: 'Cancelled'});
  const r2 = env2.ctx.tmv3_preInspectionAutomationCycle_({force: true});
  assert.equal(r2.results[0].status, 'ATTENTION'); assert.equal(alerts(env2.log), 1);
  assert.match(env2.ctx.tmv3_piAlertExplain_('Fresh Task state is no longer OPEN.').what, /no longer Open/);
  assert.match(env2.ctx.tmv3_piAlertExplain_('PI_REQUESTED_BY_MUST_MATCH_VERIFIED_EVENT_OWNER_NO_WRITE').what, /Requested By/);
  assert.match(env2.ctx.tmv3_piAlertExplain_('weird').what, /changed nothing/);
});

check('NEW_OFF_SHARED_REP_CALENDAR_NOTICES', () => {
  const env = load(REBUILT, {transform: (f, x) => f === '30_PreInspection.gs' ? x : x}); // as shipped: rules ON
  const c = env.ctx, sent = [];
  const STEPHEN = 'classicfireplace.ca_c20qcqfhvjbv784asn9pvuiaf4@group.calendar.google.com';
  const mk = (id, title, desc, creator, loc) => ({getId: () => id + '@google.com', isAllDayEvent: () => false, __r: Object.assign(rec(0), {
    eventId: id + '@google.com', title, rawDescription: desc, descriptionClean: desc, creator, location: loc,
    calendarRole: 'SECONDARY_OWNER', sourceCalendarRoles: ['SECONDARY_OWNER'], calendarId: STEPHEN, organizerCalendarId: STEPHEN,
    calendarCreatedAt: new Date(NOW - 5 * 86400000).toISOString(), fingerprint: 'f-' + id, orderNumber: '', customerNumber: ''})});
  const repEvents = [mk('rep1', 'Pre-inspection Jane Doe 416-555-0199', '', 'adam@classicfireplace.ca', '1 Main St'),
    mk('rep2', 'Service call furnace', '', 'adam@classicfireplace.ca', '2 Main St'),
    mk('rep3', 'Team meeting', '', 'adam@classicfireplace.ca', ''),
    mk('shared1', 'Pre-inspection Bob', '', 'adam@classicfireplace.ca', '3 Main St')];
  c.CalendarApp = {getCalendarById: id => id === PRIMARY ? {getEvents: () => [{getId: () => 'shared1@google.com'}]} :
    id === STEPHEN ? {getEvents: () => repEvents} : null};
  c.tmv3_calendarEvent_ = (v, cfg, src, ev) => JSON.parse(JSON.stringify(ev.__r));
  c.tmv3_audit_ = () => {};
  c.tmv3_rows_ = () => CUSTOMERS;
  c.tmv3_piSendMissingDetailsAfterRelease_ = r => { sent.push(c.tmv3_preInspectionMissingDetailAssessment_(r)); return {status: 'SENT', sent: 1}; };
  const out = c.tmv3_piNotifyOffSharedCandidates_({maxLiveChecks: 0});
  assert.equal(out.candidates, 1); assert.equal(out.notified, 1);
  assert.equal(sent[0].eventId, 'rep1@google.com');
  // Rep-booked customer visit without the word pre-inspection is a candidate; service/meeting/inspector-own are not.
  const cfgPI = {secondaryIgnoredCreatorEmails: ['stephen@classicfireplace.ca', 'jay@classicfireplace.ca', 'trevor@classicfireplace.ca', 'pramodh@classicfireplace.ca', 'thang@classicfireplace.ca']};
  const visit = r => c.tmv3_piRepBookedIntent_(Object.assign(rec(0), {calendarRole: 'SECONDARY_OWNER', customerNumber: '', orderNumber: ''}, r), cfgPI);
  assert.equal(visit({title: 'John Barber (416) 881-3225', creator: 'adam@classicfireplace.ca', location: '5 Wheatley Crescent', rawDescription: 'Remove gas unit', descriptionClean: 'Remove gas unit'}).status, 'EXPLICIT_PREINSPECTION');
  assert.equal(visit({title: 'Jessica Kirsch', creator: 'steve@classicfireplace.ca', location: '64 Alcorn Ave', phone: '', rawDescription: 'Wood Insert.', descriptionClean: 'Wood Insert.'}).status, 'EXPLICIT_PREINSPECTION');
  assert.equal(visit({title: 'Al Konstantinou 2 fireplace service', creator: 'warren@classicfireplace.ca', location: '179 Lawson Rd'}).status, 'IGNORE');
  assert.equal(visit({title: 'Chatsworth - Layout meeting', creator: 'adam@classicfireplace.ca', location: ''}).status, 'IGNORE');
  assert.equal(visit({title: 'Ann Marie paananen 416)266-9473', creator: 'stephen@classicfireplace.ca', location: '26 Cathedral Bluffs Dr'}).status, 'IGNORE');
  assert.equal(visit({title: 'Vince raise units install power vent head', creator: 'adam@classicfireplace.ca', location: '9 Drumsnab Rd'}).status, 'IGNORE');
  const d = x => ({rawDescription: x, descriptionClean: x});
  assert.equal(visit(Object.assign({title: 'Ann Lee 416-555-0101', creator: 'steve@classicfireplace.ca', location: '1 A St'}, d('Fireplace not working, pilot will not light'))).reason, 'SERVICE_CALL');
  assert.equal(visit(Object.assign({title: 'Bo Kim 416-555-0102', creator: 'steve@classicfireplace.ca', location: '1 A St'}, d('Remove gas unit in preparation for demo by contractor'))).reason, 'INSTALLATION');
  assert.equal(visit(Object.assign({title: 'Cy Doe 416-555-0103', creator: 'steve@classicfireplace.ca', location: '1 A St'}, d('Valor G3 or G3.5'))).status, 'EXPLICIT_PREINSPECTION');
  assert.equal(visit(Object.assign({title: 'Di Roe 416-555-0104', creator: 'steve@classicfireplace.ca', location: '1 A St'}, d(''))).reason, 'VISIT_TYPE_UNCLEAR_NO_EMAIL');
  const codes = sent[0].missing.map(m => m.code);
  assert.ok(codes.includes('SHARED_CALENDAR') && codes.includes('CUSTOMER_NUMBER') && codes.includes('SALES_ORDER'), codes.join(','));
  assert.equal(out.sourceErrors.length, 0, 'only Stephen is scanned');
  // SO missing is notifiable on CF Preinspects events too.
  const a = c.tmv3_preInspectionMissingDetailAssessment_(Object.assign(rec(0), {title: '12345 - Sample Customer - (416) 555-0100', rawDescription: '', descriptionClean: '', orderNumber: '', location: '1 Main St, Toronto'}));
  assert.deepEqual(JSON.parse(JSON.stringify(a.missing.map(m => m.code))), ['SALES_ORDER']);
  // Email template unchanged; the new row has its own 'should be' text.
  assert.equal(c.tmv3_piRendererFingerprint_(), 'b77a1a07');
  const html = c.tmv3_preInspectionComparisonRowsHtml_({title: 'x'}, {missing: [{code: 'SHARED_CALENDAR', field: 'Guests', message: 'CF Preinspects calendar not added as a guest'}]});
  assert.match(html, /Add the CF Preinspects calendar as a guest/);
});

// ---------------------------------------------------------------- report
if (failures.length) {
  console.error('FAIL ' + failures.length + '/' + checks + '\n' + failures.join('\n'));
  process.exit(1);
}
check('NEW_ALL_ELIGIBLE_WRITES_IN_ONE_BATCH', () => {
  // Operator rule 2026-11-12: no one-write-per-batch limit; every eligible event in the batch is corrected.
  const sc = {records: [rec(0), rec(1), rec(2)], plans: {'evt0@google.com': [[PATCH], [CONVERGED]],
    'evt1@google.com': [[PATCH], [CONVERGED]], 'evt2@google.com': [[CONVERGED]]}};
  const live = cycleEnvOn(load(REBUILT, {transform: (f, t) => f === '30_PreInspection.gs' ? t.replace(/^var TMV3_PI_CLEANED_RELEASE_APPROVED = (?:false|true);/m, 'var TMV3_PI_CLEANED_RELEASE_APPROVED = true;') : t}), sc);
  const r = live.ctx.tmv3_preInspectionAutomationCycle_({force: true});
  assert.equal(r.writes, 2, JSON.stringify(r)); assert.equal(r.processed, 3); assert.ok(!r.deferredWriteBudget);
  assert.equal(live.log.filter(e => e.kind === 'EXECUTE').length, 2);
  // The old limit still exists behind its switch (used by the equivalence baseline).
  const capped = cycleEnv(REBUILT, sc).ctx.tmv3_preInspectionAutomationCycle_({force: true});
  assert.equal(capped.writes, 1); assert.equal(capped.deferredWriteBudget, true);
});

console.log('PASS: ' + checks + ' equivalence and safety checks (shared files: ' + path.relative(ROOT, SHARED) + ')');
