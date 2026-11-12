'use strict';
// As shipped, a PreInspection run writes only durable state to Script Properties
// (caches go to CacheService), and the cleanup never touches shared/other keys.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const h = require('./harness.cjs');
const ROOT = path.resolve(__dirname, '..', '..');
const env = h.load(path.join(ROOT, 'apps-script'), {fixedNow: Date.parse('2026-11-12T15:00:00Z')});
vm.runInContext(fs.readFileSync(path.join(ROOT, 'apps-script', '95_PI_Operator_Tests.gs'), 'utf8'), env.ctx);
const c = env.ctx;

// Ephemeral writes land in CacheService, not Script Properties.
const store = c.tmv3_piEphemeral_();
store.set('TMV3_PI_DONE_abc', '{"at":1}', 3600);
store.set('TMV3_PI_LAST_RESULT', '{}');
assert.equal(store.get('TMV3_PI_DONE_abc'), '{"at":1}');
assert.equal(env.props.getProperty('TMV3_PI_DONE_abc'), null);
assert.equal(env.props.getProperty('TMV3_PI_LAST_RESULT'), null);

// Failure-alert dedupe uses the cache too.
c.TMV3_PI_CLEANED_RELEASE_APPROVED = true;
c.tmv3_audit_ = () => {};
c.tmv3_findFreshEventRecord_ = () => null;
c.tmv3_preInspectionFailureAlert_('e@google.com', 'X');
assert.ok(!Object.keys(env.props.store).some(k => /^TMV3_PI_FAILURE_/.test(k)), 'no failure keys in Script Properties');
assert.equal(c.tmv3_preInspectionFailureAlert_('e@google.com', 'X').status, 'ALREADY_ALERTED');

// Cleanup: junk removed, needed state and other verticals kept.
const seed = {
  TMV3_PI_DONE_1111111111111111: '{}', TMV3_PI_FAILURE_2222222222222222: 'SENT', TMV3_PI_LAST_RESULT: '{}',
  TMV3_PI_RETRY_SWEEP_START: '1', TMV3_PI_CONFIRM_TEST_3333333333333333: '{"answer":"YES"}',
  TMV3_PI_NOTIFICATION_4444444444444444444444444444444444444444444444444444444444444444: '{"status":"SENT","lastSentAt":"2026-08-01T12:00:00Z"}',
  TMV3_PI_NOTIFICATION_5555555555555555555555555555555555555555555555555555555555555555: '{"status":"SENT","lastSentAt":"2026-11-11T12:00:00Z"}',
  TMV3_PI_PENDING_EVENTS: '[]', TMV3_PI_SWEEP_STATE: '{}', TMV3_PI_CONFIRM_SECRET: 'x', TMV3_PI_CONFIRM_WEBAPP_URL: 'https://script.google.com/x',
  TMV3_PI_CONFIRM_6666666666666666: '{"answer":"NO"}', TMV3_DURABLE_WRITE_GUARD_abc: '{"state":"DONE"}', TMV3_SERVICE_CURSOR: '3',
  TMV3_SD_REQUESTED_BY_20261021_19275: 'DONE', TMV3_OPEN_REFRESH_20261021_TASK_18908: 'DONE', 'TMV3_PI_ELIGIBLE_CURSOR_2026-10-19': '3',
  'TMV3_PI_ELIGIBLE_CURSOR_2026-11-12': '1', TMV3_SD_ORDER_BACKFILL_20261020_19223: '{"state":"DONE"}',
  'TMV3_VERTICAL_SWEEP_Service_RANGE_2026-10-26_2026-11-02_IDS_0': '[]', TMV3_VERTICAL_SWEEP_Service: '{}',
  ['TMV3_PI_ORIGIN_' + 'a'.repeat(64)]: '{}', ['TMV3_PI_ORIGIN_' + 'b'.repeat(64)]: '{}',
  TMV3_INSTALL_SO_REMINDER_x1: '2026-11-03T12:43:02-05:00', TMV3_INSTALL_SO_REMINDER_x2: '2026-11-12T08:00:00-05:00',
  TMV3_INSTALL_SO_REMINDER_x3: 'SENDING', TMV3_OWNER_BUDGET_RESET_20261021: '{}', TMV3_INSTALL_VIEW_LAST_SUCCESS: '{}',
  TMV3_LAST_MASTER_SOURCE_REFRESH: '2026-11-12T13:01:48Z',
  TMV3_LOCATION_CREATE_GUARD_g1: JSON.stringify({vertical:'PreInspection', eventId:'past@google.com'}),
  TMV3_LOCATION_CREATE_GUARD_g2: JSON.stringify({vertical:'PreInspection', eventId:'soon@google.com'})};
// Upcoming appointments: notification 5…, origin b…, and the 'soon' event.
c.piTest_upcomingEventHashes_ = () => ({['5'.repeat(64)]: true, ['b'.repeat(64)]: true,
  ids: {'soon@google.com': true, 'inst-soon@google.com': true},
  byVertical: {[c.tmv3_hash_('Install|inst-soon@google.com')]: true}});
Object.assign(seed, {
  ['TMV3_VERTICAL_EVENT_' + c.tmv3_hash_('Install|inst-past@google.com')]: '{"state":"VERIFIED"}',
  ['TMV3_VERTICAL_EVENT_' + c.tmv3_hash_('Install|inst-soon@google.com')]: '{"state":"VERIFIED"}',
  ['TMV3_VERTICAL_EVENT_' + c.tmv3_hash_('Service|svc-past@google.com')]: '{"state":"UNCERTAIN"}',
  TMV3_TASK_CREATE_GUARD_t1: JSON.stringify({vertical:'Install', eventId:'inst-past@google.com', state:'VERIFIED'}),
  TMV3_TASK_CREATE_GUARD_t2: JSON.stringify({vertical:'Install', eventId:'inst-past@google.com', state:'INTENT'}),
  TMV3_VERTICAL_ALERT_a1: 'SENT', TMV3_VERTICAL_LAST_Service: '{}', TMV3_SD_ORDER_BACKFILL_20261020_19223: '{"state":"VERIFIED"}'});
Object.entries(seed).forEach(([k, v]) => env.props.setProperty(k, v));
Object.entries(seed).forEach(([k, v]) => env.props.setProperty(k, v));
const inv = c.piTest_PropertyInventory();
assert.ok(inv.totals.JUNK >= 6);
const out = c.piTest_WRITE_CleanupProperties();
const left = Object.keys(env.props.store).sort();
['TMV3_PI_NOTIFICATION_5555555555555555555555555555555555555555555555555555555555555555', 'TMV3_PI_PENDING_EVENTS', 'TMV3_PI_SWEEP_STATE', 'TMV3_PI_CONFIRM_SECRET',
  'TMV3_PI_CONFIRM_WEBAPP_URL', 'TMV3_PI_CONFIRM_6666666666666666', 'TMV3_DURABLE_WRITE_GUARD_abc', 'TMV3_SERVICE_CURSOR',
  'TMV3_PI_ELIGIBLE_CURSOR_2026-11-12', 'TMV3_VERTICAL_SWEEP_Service',
  'TMV3_VERTICAL_EVENT_' + c.tmv3_hash_('Service|svc-past@google.com'),
  'TMV3_TASK_CREATE_GUARD_t2',
  'TMV3_PI_ORIGIN_' + 'b'.repeat(64), 'TMV3_INSTALL_SO_REMINDER_x2', 'TMV3_INSTALL_SO_REMINDER_x3',
  'TMV3_LAST_MASTER_SOURCE_REFRESH', 'TMV3_LOCATION_CREATE_GUARD_g2']
  .forEach(k => assert.ok(left.includes(k), 'kept ' + k));
['TMV3_PI_DONE_1111111111111111', 'TMV3_PI_FAILURE_2222222222222222', 'TMV3_PI_LAST_RESULT', 'TMV3_PI_RETRY_SWEEP_START',
  'TMV3_PI_CONFIRM_TEST_3333333333333333', 'TMV3_PI_NOTIFICATION_4444444444444444444444444444444444444444444444444444444444444444',
  'TMV3_SD_REQUESTED_BY_20261021_19275', 'TMV3_OPEN_REFRESH_20261021_TASK_18908', 'TMV3_PI_ELIGIBLE_CURSOR_2026-10-19',
  'TMV3_VERTICAL_SWEEP_Service_RANGE_2026-10-26_2026-11-02_IDS_0',
  'TMV3_PI_ORIGIN_' + 'a'.repeat(64), 'TMV3_INSTALL_SO_REMINDER_x1', 'TMV3_OWNER_BUDGET_RESET_20261021',
  'TMV3_INSTALL_VIEW_LAST_SUCCESS', 'TMV3_LOCATION_CREATE_GUARD_g1', 'TMV3_SD_ORDER_BACKFILL_20261020_19223',
  'TMV3_VERTICAL_EVENT_' + c.tmv3_hash_('Install|inst-past@google.com'), 'TMV3_TASK_CREATE_GUARD_t1',
  'TMV3_VERTICAL_ALERT_a1', 'TMV3_VERTICAL_LAST_Service',
  'TMV3_VERTICAL_EVENT_' + c.tmv3_hash_('Install|inst-soon@google.com')]
  .forEach(k => assert.ok(!left.includes(k), 'deleted ' + k));
assert.equal(out.deletedTotal, 21);
// Calendars unreadable: nothing appointment-linked is deleted.
c.piTest_upcomingEventHashes_ = () => null;
env.props.setProperty('TMV3_PI_ORIGIN_' + 'c'.repeat(64), '{}');
c.piTest_WRITE_CleanupProperties();
assert.ok(env.props.getProperty('TMV3_PI_ORIGIN_' + 'c'.repeat(64)) !== null, 'kept when calendars unreadable');
console.log('PASS: caches kept out of Script Properties; cleanup deletes only PreInspection junk');
